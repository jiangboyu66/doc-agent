/**
 * Word 原生编辑引擎（DocxDocument）
 *
 * 直接在 docx 的 OOXML 上工作，不经过任何格式转换：
 *   - 读：把每个 <w:p> 映射成"渲染文本 + 片段(segment)表"，片段记录每段文字来自哪个 <w:r>/<w:t>；
 *   - 改：把"把 A 改成 B"拆成最小修改块，只拆分/替换受影响的 run，新文字继承原 run 的格式；
 *   - 写：只重新序列化被改动的节点，ZIP 中未改动的部件按原压缩字节拷贝。
 * 结果：用户没要求改的每一个字、每一个格式、页眉页脚、图片、分节、域、书签都保持原样。
 *
 * 支持 Word 修订模式（Track Changes）：打开后所有修改以 <w:ins>/<w:del>/<w:rPrChange>/<w:pPrChange>
 * 记录，用户可以在 Word 里逐条接受或拒绝——把最终决定权交还给用户。
 */

import {
  XDocument, XElement, XNode, parseXml, parseFragment, serializeNode, serializeDocument,
  removeNode, insertBefore, insertAfter, replaceNode, appendChild, markDirty, encodeText, encodeAttr, innerText,
} from "./xml.js";
import { diffArrays } from "diff";
import { ZipPackage } from "./zip.js";
import {
  NS, REL_TYPES, CT_COMMENTS, RPR_ORDER, PPR_ORDER, TRPR_ORDER, StyleSheet, parseRels, relsPathFor,
  resolveTarget, orderedInsertIndex, ptToTwips, REL_EXTRA, CT, SECTPR_ORDER, TCPR_ORDER, TBLPR_ORDER, STYLE_ORDER,
} from "./ooxml.js";
import {
  runContent, imageInfo, pictureXml, shapesGroupXml, chartXml, chartWorkbook, chartInlineXml, abstractNumXml,
  equationParagraphXml, emu, twip, alternateContentXml, type ShapeSpec, type ChartSpec,
} from "./build.js";
import { findAll, computeHunks, clip, type Hunk } from "../textMatch.js";
import { symbolChar } from "./symbolFont.js";
import {
  DocError, type DocumentAdapter, type DocumentCapabilities, type BlockInfo, type SearchHit, type DocumentSummary,
  type EditOptions, type EditResult, type RunFormat, type ParagraphFormat, type NewBlock, type FidelityReport,
  type OutlineItem,
} from "../types.js";

// ---------------------------------------------------------------------------
// 类型
// ---------------------------------------------------------------------------

type SegKind = "text" | "special" | "atom" | "zw";

interface Seg {
  kind: SegKind;
  start: number;
  end: number;
  text: string;
  run: XElement | null;
  node: XElement;
  /** 域标记、批注引用等结构性零宽标记：修改范围不能跨越 */
  hard?: boolean;
}

interface ParaEntry {
  ref: string;
  part: string;
  el: XElement;
  ordinal: number;
  paraId?: string;
  location: string;
  container: "body" | "table" | "textbox" | "header" | "footer" | "footnote" | "endnote";
}

const PART_LABEL = (part: string): string => {
  const f = part.split("/").pop()!.replace(/\.xml$/, "");
  if (f.startsWith("header")) return `页眉(${f})`;
  if (f.startsWith("footer")) return `页脚(${f})`;
  if (f === "footnotes") return "脚注";
  if (f === "endnotes") return "尾注";
  return "正文";
};

const TRANSPARENT_CONTAINERS = new Set(["hyperlink", "ins", "moveTo", "smartTag", "customXml", "fldSimple", "dir", "bdo"]);
const CLEANUP_CONTAINERS = new Set(["hyperlink", "ins", "moveTo", "smartTag", "customXml", "dir", "bdo"]);
const REVISION_ID_ELEMENTS = new Set([
  "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange",
  "tcPrChange", "numberingChange", "commentRangeStart", "commentRangeEnd", "commentReference", "comment",
  "bookmarkStart", "bookmarkEnd", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
]);

// ---------------------------------------------------------------------------

export class DocxDocument implements DocumentAdapter {
  readonly format = "docx" as const;
  readonly capabilities: DocumentCapabilities = {
    replaceText: true, formatText: true, paragraphProps: true, insertBlocks: true, deleteBlocks: true,
    tables: true, comments: true, trackChanges: true, styles: true,
    media: true, layout: true, lists: true, notes: true, links: true, styleEdit: true, revisions: true, tableEdit: true, move: true, equations: true,
  };
  private _structureVersion = 0;
  /** 结构版本号；外部（会话重载 / 干跑副本）设置时需要重建引用表 */
  get structureVersion(): number { return this._structureVersion; }
  set structureVersion(v: number) {
    if (v === this._structureVersion) return;
    this._structureVersion = v;
    if (this.paras) this.reindex();
  }

  private pkg!: ZipPackage;
  private parts = new Map<string, XDocument>();
  private prefixes = new Map<string, string>();
  mainPart!: string;
  private contentParts: string[] = [];
  private styles!: StyleSheet;
  private stylesPart: string | null = null;
  private paras: ParaEntry[] = [];
  private byRef = new Map<string, ParaEntry>();
  private nextRevisionId = -1;

  constructor(buf: Buffer) {
    this.load(buf);
  }

  /** 整体替换文档内容（排版类工具在文档副本上完成全局调整后写回）；结构版本递增，旧引用失效 */
  reload(buf: Buffer): void {
    this.parts.clear();
    this.prefixes.clear();
    this.snapshots.clear();
    this.paras = [];
    this.load(buf);
    this._structureVersion++;
    this.reindex();
  }

  private load(buf: Buffer): void {
    this.pkg = new ZipPackage(buf);
    const rootRels = parseRels(this.optText("_rels/.rels"));
    const main = rootRels.find((r) => r.type === REL_TYPES.officeDocument);
    this.mainPart = main ? resolveTarget("", main.target) : "word/document.xml";
    if (!this.pkg.has(this.mainPart)) throw new DocError("不是有效的 Word 文档：缺少主文档部件", "invalid");

    const rels = parseRels(this.optText(relsPathFor(this.mainPart)));
    const pick = (type: string) =>
      rels.filter((r) => r.type === type && !r.external).map((r) => resolveTarget(this.mainPart, r.target)).filter((p) => this.pkg.has(p));
    const natural = (a: string, b: string) => a.localeCompare(b, undefined, { numeric: true });
    this.contentParts = [
      this.mainPart,
      ...pick(REL_TYPES.header).sort(natural),
      ...pick(REL_TYPES.footer).sort(natural),
      ...pick(REL_TYPES.footnotes),
      ...pick(REL_TYPES.endnotes),
    ];
    const stylesPart = pick(REL_TYPES.styles)[0];
    this.stylesPart = stylesPart ?? null;
    this.styles = new StyleSheet(stylesPart ? this.optText(stylesPart) : null);
    this.reindex();
  }

  // -------------------------------------------------------------------------
  // 部件访问
  // -------------------------------------------------------------------------

  private optText(name: string): string | null {
    return this.pkg.has(name) ? this.decode(name, this.pkg.read(name)) : null;
  }

  private decode(name: string, buf: Buffer): string {
    if (buf[0] === 0xff && buf[1] === 0xfe) throw new DocError(`部件 ${name} 使用 UTF-16 编码，暂不支持编辑`, "unsupported");
    return buf.toString("utf8");
  }

  private part(name: string): XDocument {
    let d = this.parts.get(name);
    if (!d) {
      d = parseXml(this.decode(name, this.pkg.read(name)));
      this.parts.set(name, d);
      const root = d.root;
      let prefix = "w";
      for (const [k, v] of root.attrs) {
        if (k.startsWith("xmlns:") && v === NS.w) prefix = k.slice(6);
      }
      this.prefixes.set(name, prefix);
    }
    return d;
  }

  private w(part: string): string {
    this.part(part);
    return this.prefixes.get(part)!;
  }

  private src(part: string): string {
    return this.part(part).source;
  }

  // -------------------------------------------------------------------------
  // 段落索引
  // -------------------------------------------------------------------------

  /** 各结构版本下的段落顺序快照：旧的 P 序号引用只要段落还在，就能映射到它现在的位置 */
  private snapshots = new Map<number, XElement[]>();
  private indexVersion = 0;

  private reindex(): void {
    if (this.paras?.length) {
      this.snapshots.set(this.indexVersion, this.paras.map((p) => p.el));
      if (this.snapshots.size > 60) this.snapshots.delete(this.snapshots.keys().next().value!);
    }
    this.indexVersion = this._structureVersion;
    this.paras = [];
    let tableCounter = 0;
    for (const part of this.contentParts) {
      const doc = this.part(part);
      const w = this.w(part);
      const label = PART_LABEL(part);
      const partKind: ParaEntry["container"] = part === this.mainPart ? "body"
        : label.startsWith("页眉") ? "header" : label.startsWith("页脚") ? "footer"
        : label === "脚注" ? "footnote" : "endnote";

      const visit = (node: XElement, loc: string, container: ParaEntry["container"]) => {
        for (const c of node.children) {
          if (c.type !== "el") continue;
          if (c.local === "AlternateContent") {
            const choice = c.elements().find((x) => x.local === "Choice");
            if (choice) visit(choice, loc, container);
            continue;
          }
          if (c.prefix !== w) {
            visit(c, loc, container);
            continue;
          }
          if (c.local === "p") {
            this.paras.push({ ref: "", part, el: c, ordinal: 0, paraId: c.getAttr("w14:paraId"), location: loc, container });
            visit(c, loc, container); // 段落内可能嵌有文本框里的段落
            continue;
          }
          if (c.local === "txbxContent") {
            visit(c, `${loc} › 文本框`, "textbox");
            continue;
          }
          if (c.local === "tbl") {
            const t = ++tableCounter;
            let r = 0;
            for (const tr of c.childrenNamed(`${w}:tr`)) {
              r++;
              let col = 0;
              for (const tc of tr.childrenNamed(`${w}:tc`)) {
                col++;
                visit(tc, `${loc === "正文" ? "" : loc + " › "}表格${t} 第${r}行第${col}列`, "table");
              }
            }
            continue;
          }
          if (c.local === "pPr" || c.local === "rPr" || c.local === "sectPr" || c.local === "del" || c.local === "moveFrom") continue;
          visit(c, loc, container);
        }
      };

      const root = doc.root;
      if (part === this.mainPart) {
        const body = root.child(`${w}:body`);
        if (body) visit(body, "正文", "body");
      } else if (partKind === "footnote" || partKind === "endnote") {
        for (const note of root.elements()) {
          const type = note.getAttr(`${w}:type`);
          if (type && type !== "normal") continue;
          const id = note.getAttr(`${w}:id`);
          visit(note, `${partKind === "footnote" ? "脚注" : "尾注"}${id}`, partKind);
        }
      } else {
        visit(root, label, partKind);
      }
    }

    const idCount = new Map<string, number>();
    for (const p of this.paras) if (p.paraId) idCount.set(p.paraId, (idCount.get(p.paraId) ?? 0) + 1);
    this.byRef.clear();
    this.paras.forEach((p, i) => {
      p.ordinal = i + 1;
      p.ref = p.paraId && idCount.get(p.paraId) === 1 ? `#${p.paraId}` : `P${i + 1}@v${this.structureVersion}`;
      this.byRef.set(p.ref, p);
    });
  }

  /** 把旧版本的 P 序号引用换算成当前引用（干跑副本没有历史快照，执行前先统一换算） */
  translateRef(ref: string): string {
    if (!/^P\d+@v\d+$/.test(ref.trim())) return ref;
    try {
      return this.resolve(ref).ref;
    } catch {
      return ref;
    }
  }

  private resolve(ref: string): ParaEntry {
    const r = ref.trim();
    const direct = this.byRef.get(r) ?? this.byRef.get(`#${r}`);
    if (direct) return direct;
    const m = /^P(\d+)(?:@v(\d+))?$/.exec(r);
    if (m) {
      if (m[2] !== undefined && Number(m[2]) !== this.structureVersion) {
        const snap = this.snapshots.get(Number(m[2]));
        const el = snap?.[Number(m[1]) - 1];
        if (el) {
          const now = this.paras.find((x) => x.el === el);
          if (now) return now;
          throw new DocError(`段落 ${r} 已被删除（文档结构在之后发生了变化）。请重新调用 doc_read 获取最新引用。`, "stale_ref");
        }
        throw new DocError(
          `段落引用 ${r} 已过期：之后文档插入/删除过段落（当前结构版本 v${this.structureVersion}）。请重新调用 doc_read 或 doc_search 获取最新引用。`,
          "stale_ref"
        );
      }
      const p = this.paras[Number(m[1]) - 1];
      if (p) return p;
    }
    throw new DocError(`找不到段落引用 ${r}。段落引用必须来自 doc_read / doc_search / doc_outline 的返回结果。`, "not_found");
  }

  // -------------------------------------------------------------------------
  // 段落 → 渲染文本 + 片段
  // -------------------------------------------------------------------------

  private segments(p: ParaEntry): { text: string; segs: Seg[] } {
    const w = this.w(p.part);
    const segs: Seg[] = [];
    let pos = 0;
    const push = (kind: SegKind, text: string, run: XElement | null, node: XElement, hard = false) => {
      segs.push({ kind, start: pos, end: pos + text.length, text, run, node, hard });
      pos += text.length;
    };

    const describeObject = (el: XElement): string => {
      for (const d of el.descendants()) if (d.local === "txbxContent") return "⟨文本框⟩";
      if (el.local === "object") return "⟨嵌入对象⟩";
      for (const d of el.descendants()) if (d.local === "chart") return "⟨图表⟩";
      return "⟨图片⟩";
    };

    const walkRun = (run: XElement) => {
      for (const c of run.children) {
        if (c.type !== "el") continue;
        if (c.prefix !== w) {
          if (c.local === "AlternateContent") push("atom", describeObject(c), run, c);
          continue;
        }
        switch (c.local) {
          case "t": {
            let s = "";
            for (const t of c.children) if (t.type === "text") s += t.text;
            if (s) push("text", s, run, c);
            else push("zw", "", run, c);
            break;
          }
          case "tab": case "ptab": push("special", "\t", run, c); break;
          case "br": {
            const type = c.getAttr(`${w}:type`);
            if (type === "page") push("atom", "⟨分页符⟩", run, c);
            else if (type === "column") push("atom", "⟨分栏符⟩", run, c);
            else push("special", "\n", run, c);
            break;
          }
          case "cr": push("special", "\n", run, c); break;
          case "noBreakHyphen": push("special", "‑", run, c); break;
          case "softHyphen": push("zw", "", run, c); break;
          case "sym": {
            const ch = symbolChar(c.getAttr(`${w}:font`), c.getAttr(`${w}:char`));
            push("atom", ch ? `⟨符号${ch}⟩` : "⟨符号⟩", run, c);
            break;
          }
          case "drawing": case "pict": case "object": push("atom", describeObject(c), run, c); break;
          case "footnoteReference": push("atom", `⟨脚注${c.getAttr(`${w}:id`) ?? ""}⟩`, run, c); break;
          case "endnoteReference": push("atom", `⟨尾注${c.getAttr(`${w}:id`) ?? ""}⟩`, run, c); break;
          case "footnoteRef": case "endnoteRef": push("atom", "⟨注号⟩", run, c); break;
          case "separator": case "continuationSeparator": push("atom", "⟨分隔线⟩", run, c); break;
          case "fldChar": case "commentReference": push("zw", "", run, c, true); break;
          default: break; // rPr / instrText / delText / lastRenderedPageBreak 等不参与渲染
        }
      }
    };

    const walk = (node: XElement) => {
      for (const c of node.children) {
        if (c.type !== "el") continue;
        if (c.prefix !== w) {
          if (c.local === "oMath" || c.local === "oMathPara") push("atom", "⟨公式⟩", null, c);
          else if (c.local === "AlternateContent") {
            const choice = c.elements().find((x) => x.local === "Choice");
            if (choice) walk(choice);
          }
          continue;
        }
        if (c.local === "r") walkRun(c);
        else if (TRANSPARENT_CONTAINERS.has(c.local)) walk(c);
        else if (c.local === "sdt") {
          const content = c.child(`${w}:sdtContent`);
          if (content) walk(content);
        }
      }
    };
    walk(p.el);
    return { text: segs.map((s) => s.text).join(""), segs };
  }

  private textOf(p: ParaEntry): string {
    return this.segments(p).text;
  }

  // -------------------------------------------------------------------------
  // 读接口
  // -------------------------------------------------------------------------

  private blockInfo(p: ParaEntry): BlockInfo {
    const w = this.w(p.part);
    const pPr = p.el.child(`${w}:pPr`);
    const styleId = pPr?.child(`${w}:pStyle`)?.getAttr(`${w}:val`) ?? this.styles.defaultParagraphStyle()?.id;
    const style = this.styles.get(styleId);
    const ownOl = pPr?.child(`${w}:outlineLvl`)?.getAttr(`${w}:val`);
    const level = ownOl !== undefined && Number(ownOl) < 9 ? Number(ownOl) + 1 : this.styles.headingLevel(styleId);
    const flags: string[] = [];
    if (pPr?.child(`${w}:numPr`) || style?.hasNumbering) flags.push("自动编号");
    if (pPr?.child(`${w}:sectPr`)) flags.push("分节符");
    let hasIns = false, hasDel = false;
    for (const d of p.el.descendants()) {
      if (d.prefix !== w) continue;
      if (d.local === "ins") hasIns = true;
      if (d.local === "del") hasDel = true;
    }
    if (hasIns || hasDel) flags.push("含修订");
    let kind: BlockInfo["kind"] = "paragraph";
    if (this.styles.isTitle(styleId)) kind = "title";
    else if (level !== undefined) kind = "heading";
    else if (p.container === "table") kind = "table-cell";
    else if (flags.includes("自动编号")) kind = "list";
    return {
      ref: p.ref,
      location: p.location,
      kind,
      level,
      style: style ? style.name : styleId,
      text: this.textOf(p),
      flags: flags.length ? flags : undefined,
    };
  }

  listBlocks(opts?: { part?: string }): BlockInfo[] {
    return this.paras.filter((p) => !opts?.part || p.location.startsWith(opts.part)).map((p) => this.blockInfo(p));
  }

  getBlock(ref: string): BlockInfo {
    return this.blockInfo(this.resolve(ref));
  }

  search(query: string, opts: { regex?: boolean; caseSensitive?: boolean }): SearchHit[] {
    const hits: SearchHit[] = [];
    for (const p of this.paras) {
      const text = this.textOf(p);
      if (opts.regex) {
        let re: RegExp;
        try {
          re = new RegExp(query, opts.caseSensitive ? "g" : "gi");
        } catch (e: any) {
          throw new DocError(`正则表达式无效：${e.message}`, "invalid");
        }
        for (const m of text.matchAll(re)) {
          hits.push({ ref: p.ref, location: p.location, text, index: m.index ?? 0, match: m[0] });
        }
      } else {
        for (const m of findAll(text, query, opts.caseSensitive ?? false)) {
          hits.push({ ref: p.ref, location: p.location, text, index: m.start, match: text.slice(m.start, m.end) });
        }
      }
    }
    return hits;
  }

  listStyles() {
    const usage = new Map<string, number>();
    for (const p of this.paras) {
      const w = this.w(p.part);
      const id = p.el.child(`${w}:pPr`)?.child(`${w}:pStyle`)?.getAttr(`${w}:val`) ?? this.styles.defaultParagraphStyle()?.id;
      if (id) usage.set(id, (usage.get(id) ?? 0) + 1);
      for (const d of p.el.descendants()) {
        if (d.local === "rStyle") {
          const v = d.getAttr(`${w}:val`);
          if (v) usage.set(v, (usage.get(v) ?? 0) + 1);
        }
      }
    }
    return [...this.styles.byId.values()]
      .filter((s) => s.type === "paragraph" || s.type === "character")
      .map((s) => ({ id: s.id, name: s.name, type: s.type, inUse: usage.get(s.id) ?? 0 }))
      .sort((a, b) => b.inUse - a.inUse || a.name.localeCompare(b.name));
  }

  describeFormat(ref: string): string {
    const p = this.resolve(ref);
    const w = this.w(p.part);
    const info = this.blockInfo(p);
    const pPr = p.el.child(`${w}:pPr`);
    const lines: string[] = [];
    lines.push(`段落 ${p.ref}（${p.location}）`);
    lines.push(`样式：${info.style ?? "默认"}${info.level ? `（标题 ${info.level} 级）` : ""}${info.flags ? `  标记：${info.flags.join("、")}` : ""}`);
    if (pPr) {
      const attrs = (el: XElement | undefined) => (el ? el.attrs.map(([k, v]) => `${k.replace(`${w}:`, "")}=${v}`).join(" ") : "");
      const jc = pPr.child(`${w}:jc`)?.getAttr(`${w}:val`);
      if (jc) lines.push(`对齐：${jc}`);
      const sp = pPr.child(`${w}:spacing`);
      if (sp) lines.push(`间距（twips，1pt=20）：${attrs(sp)}`);
      const ind = pPr.child(`${w}:ind`);
      if (ind) lines.push(`缩进（twips）：${attrs(ind)}`);
    }
    lines.push("run 列表（每个 run 是一段格式统一的文字）：");
    const { segs } = this.segments(p);
    const runs: XElement[] = [];
    for (const s of segs) if (s.run && !runs.includes(s.run)) runs.push(s.run);
    runs.forEach((r, i) => {
      const txt = segs.filter((s) => s.run === r).map((s) => s.text).join("");
      lines.push(`  [${i + 1}] "${clip(txt, 60)}"  ${this.describeRPr(r.child(`${w}:rPr`), w) || "（继承样式格式）"}`);
    });
    return lines.join("\n");
  }

  private describeRPr(rPr: XElement | undefined, w: string): string {
    if (!rPr) return "";
    const out: string[] = [];
    for (const c of rPr.elements()) {
      const v = c.getAttr(`${w}:val`);
      const off = v === "0" || v === "false" || v === "none";
      switch (c.local) {
        case "rStyle": out.push(`字符样式=${this.styles.get(v)?.name ?? v}`); break;
        case "b": if (!off) out.push("加粗"); break;
        case "i": if (!off) out.push("斜体"); break;
        case "u": if (!off) out.push(`下划线(${v ?? "single"})`); break;
        case "strike": if (!off) out.push("删除线"); break;
        case "caps": if (!off) out.push("全大写"); break;
        case "smallCaps": if (!off) out.push("小型大写"); break;
        case "sz": out.push(`字号=${Number(v) / 2}pt`); break;
        case "color": out.push(`颜色=${v}`); break;
        case "highlight": out.push(`高亮=${v}`); break;
        case "vertAlign": out.push(v === "superscript" ? "上标" : v === "subscript" ? "下标" : ""); break;
        case "rFonts": {
          const f = c.getAttr(`${w}:ascii`) ?? c.getAttr(`${w}:eastAsia`) ?? c.getAttr(`${w}:hAnsi`);
          if (f) out.push(`字体=${f}`);
          break;
        }
        case "rPrChange": out.push("含格式修订"); break;
      }
    }
    return out.filter(Boolean).join("，");
  }

  summary(): DocumentSummary {
    const headings: OutlineItem[] = [];
    let chars = 0;
    for (const p of this.paras) {
      const info = this.blockInfo(p);
      chars += info.text.replace(/\s/g, "").length;
      if ((info.kind === "heading" || info.kind === "title") && info.text.trim()) {
        headings.push({ ref: info.ref, level: info.kind === "title" ? 0 : info.level ?? 1, text: info.text.trim(), location: info.location });
      }
    }
    const main = this.part(this.mainPart);
    const w = this.w(this.mainPart);
    let tables = 0, images = 0, sections = 0, tracked = 0;
    for (const d of main.root.descendants()) {
      if (d.prefix === w) {
        if (d.local === "tbl") tables++;
        else if (d.local === "sectPr") sections++;
        else if (d.local === "ins" || d.local === "del") tracked++;
        else if (d.local === "drawing" || d.local === "pict") images++;
      }
    }
    let comments = 0;
    const commentsPart = this.commentsPartName();
    if (commentsPart && this.pkg.has(commentsPart)) comments = this.part(commentsPart).root.elements().length;
    return {
      format: "docx",
      blockCount: this.paras.length,
      charCount: chars,
      headings,
      parts: this.contentParts.map((p) => `${p}（${PART_LABEL(p)}）`),
      tables,
      images,
      sections,
      trackedChanges: tracked,
      comments,
      template: this.detectTemplate(),
      notes: [],
    };
  }

  private detectTemplate(): DocumentSummary["template"] {
    const names = new Set([...this.styles.byId.values()].map((s) => s.name));
    const evidence: string[] = [];
    const app = this.optText("docProps/app.xml");
    const tpl = app ? /<Template>([^<]*)<\/Template>/.exec(app)?.[1] : undefined;
    if (tpl) evidence.push(`docProps/app.xml Template=${tpl}`);
    const ieeeMarks = ["Paper Title", "H1_List (No Space)", "H2_First", "REF Txt", "PARA", "Abstract", "AU Bios"].filter((n) => names.has(n));
    if (ieeeMarks.length >= 3 || (tpl && /ieee|access|trans_jour/i.test(tpl))) {
      evidence.push(`IEEE 模板样式：${ieeeMarks.join("、")}`);
      return { id: "ieee", name: "IEEE 期刊/会议论文模板", evidence };
    }
    if (tpl && !/^normal(\.dotm)?$/i.test(tpl)) return { id: "custom", name: `模板 ${tpl}`, evidence };
    return undefined;
  }

  // -------------------------------------------------------------------------
  // 底层编辑原语
  // -------------------------------------------------------------------------

  private revId(): number {
    if (this.nextRevisionId < 0) {
      let max = 0;
      for (const part of [...this.contentParts, this.commentsPartName()].filter((x): x is string => !!x && this.pkg.has(x))) {
        const w = this.w(part);
        for (const d of this.part(part).root.descendants()) {
          if (d.prefix === w && REVISION_ID_ELEMENTS.has(d.local)) {
            const id = Number(d.getAttr(`${w}:id`));
            if (Number.isFinite(id) && id > max) max = id;
          }
        }
      }
      this.nextRevisionId = max + 1;
    }
    return this.nextRevisionId++;
  }

  private revAttrs(w: string, o: EditOptions): string {
    return `${w}:id="${this.revId()}" ${w}:author="${encodeAttr(o.author)}" ${w}:date="${o.date}"`;
  }

  private attrsXml(el: XElement): string {
    return el.attrs.map(([k, v]) => ` ${k}="${v}"`).join("");
  }

  private textElementXml(name: string, text: string): string {
    const needsPreserve = /^\s|\s$/.test(text);
    return `<${name}${needsPreserve ? ' xml:space="preserve"' : ""}>${encodeText(text)}</${name}>`;
  }

  /** 把纯文本转换成 run 的子元素：\t → <w:tab/>，\n → <w:br/> */
  private runChildrenXml(w: string, text: string, textTag = "t"): string {
    let out = "";
    for (const piece of text.split(/(\t|\n)/)) {
      if (piece === "") continue;
      if (piece === "\t") out += `<${w}:tab/>`;
      else if (piece === "\n") out += `<${w}:br/>`;
      else out += this.textElementXml(`${w}:${textTag}`, piece);
    }
    return out;
  }

  /** 取 run 的 rPr 片段，去掉修订历史 */
  private rPrXmlOf(run: XElement | null, part: string): string {
    if (!run) return "";
    const w = this.w(part);
    const rPr = run.child(`${w}:rPr`);
    if (!rPr) return "";
    const kids = rPr.children.filter((c) => !(c.type === "el" && c.local === "rPrChange"));
    if (!kids.length) return "";
    return `<${rPr.name}>${kids.map((k) => serializeNode(k, this.src(part))).join("")}</${rPr.name}>`;
  }

  /** 段落标记的 rPr（用于空段落中插入文字时的格式来源），去掉修订标记 */
  private paraMarkRPrXml(p: ParaEntry): string {
    const w = this.w(p.part);
    const rPr = p.el.child(`${w}:pPr`)?.child(`${w}:rPr`);
    if (!rPr) return "";
    const skip = new Set(["ins", "del", "moveFrom", "moveTo", "rPrChange"]);
    const kids = rPr.children.filter((c) => !(c.type === "el" && skip.has(c.local)));
    if (!kids.length) return "";
    return `<${w}:rPr>${kids.map((k) => serializeNode(k, this.src(p.part))).join("")}</${w}:rPr>`;
  }

  private styleIdOf(p: ParaEntry): string | undefined {
    const w = this.w(p.part);
    return p.el.child(`${w}:pPr`)?.child(`${w}:pStyle`)?.getAttr(`${w}:val`) ?? this.styles.defaultParagraphStyle()?.id;
  }

  /** 找离 near 最近、同部件同容器类型、有文字的指定样式段落（优先向前找） */
  private findStyleTemplate(styleId: string, near: ParaEntry): ParaEntry | undefined {
    const pool = this.paras.filter(
      (x) => x.part === near.part && x.container === near.container && this.styleIdOf(x) === styleId && this.textOf(x).trim()
    );
    const before = pool.filter((x) => x.ordinal < near.ordinal);
    return before[before.length - 1] ?? pool.find((x) => x.ordinal > near.ordinal);
  }

  /** 段落中覆盖文字最多的那种 run 格式——代表这个段落"正常"的文字格式 */
  private dominantRPrXml(p: ParaEntry): string {
    const weight = new Map<string, number>();
    for (const s of this.segments(p).segs) {
      if (s.kind !== "text" || !s.run) continue;
      const key = this.rPrXmlOf(s.run, p.part);
      weight.set(key, (weight.get(key) ?? 0) + s.text.length);
    }
    let best: string | null = null;
    let bestW = -1;
    for (const [k, v] of weight) if (v > bestW) [best, bestW] = [k, v];
    return best ?? this.paraMarkRPrXml(p);
  }

  private buildRunXml(w: string, rPrXml: string, text: string): string {
    return `<${w}:r>${rPrXml}${this.runChildrenXml(w, text)}</${w}:r>`;
  }

  /** 在渲染文本位置 pos 处把 run 一分为二，使 pos 落在 run 边界上 */
  private splitAt(p: ParaEntry, pos: number): void {
    const { segs } = this.segments(p);
    const seg = segs.find((s) => s.start < pos && pos < s.end);
    if (!seg) return;
    if (seg.kind !== "text" || !seg.run) {
      throw new DocError(`位置落在不可拆分的对象 ${seg.text} 内部`, "protected");
    }
    const w = this.w(p.part);
    const src = this.src(p.part);
    const run = seg.run;
    const off = pos - seg.start;
    const rPr = run.child(`${w}:rPr`);
    const rPrXml = rPr ? serializeNode(rPr, src) : "";
    const open = `<${run.name}${this.attrsXml(run)}>`;
    const close = `</${run.name}>`;
    let left = "", right = "";
    let before = true;
    for (const c of run.children) {
      if (c === rPr) continue;
      if (c === seg.node) {
        const full = seg.text;
        left += this.textElementXml(seg.node.name, full.slice(0, off));
        right += this.textElementXml(seg.node.name, full.slice(off));
        before = false;
        continue;
      }
      const xml = serializeNode(c, src);
      if (before) left += xml;
      else right += xml;
    }
    replaceNode(run, parseFragment(open + rPrXml + left + close + open + rPrXml + right + close));
  }

  /** 找到 run 在段落这一层的"锚点"：跨过修订容器，停在超链接等语义容器内 */
  private anchorOf(run: XElement, pEl: XElement): XNode {
    let n: XElement = run;
    while (n.parent && n.parent !== pEl && n.parent.type === "el" && ["ins", "moveTo"].includes(n.parent.local)) {
      n = n.parent;
    }
    return n;
  }

  private removeRunAndCleanup(run: XElement, pEl: XElement): void {
    let parent = run.parent as XElement;
    removeNode(run);
    while (parent !== pEl && CLEANUP_CONTAINERS.has(parent.local) && parent.elements().length === 0) {
      const up = parent.parent as XElement;
      removeNode(parent);
      parent = up;
    }
  }

  /** 修订模式下删除一个 run：把 w:t 改为 w:delText 并包进 <w:del> */
  private trackDeleteRun(run: XElement, p: ParaEntry, o: EditOptions): XNode | null {
    const w = this.w(p.part);
    const src = this.src(p.part);
    const parent = run.parent as XElement;
    if (parent.local === "ins" && parent.getAttr(`${w}:author`) === o.author) {
      // 删除自己尚未被接受的插入：直接移除即可
      this.removeRunAndCleanup(run, p.el);
      return null;
    }
    let inner = "";
    for (const c of run.children) {
      if (c.type === "el" && c.prefix === w && c.local === "t") {
        let s = "";
        for (const t of c.children) if (t.type === "text") s += t.text;
        inner += this.textElementXml(`${w}:delText`, s);
      } else if (c.type === "el" && c.prefix === w && c.local === "instrText") {
        let s = "";
        for (const t of c.children) if (t.type === "text") s += t.text;
        inner += this.textElementXml(`${w}:delInstrText`, s);
      } else inner += serializeNode(c, src);
    }
    const xml = `<${w}:del ${this.revAttrs(w, o)}><${run.name}${this.attrsXml(run)}>${inner}</${run.name}></${w}:del>`;
    const [wrapper] = parseFragment(xml);
    replaceNode(run, [wrapper]);
    return wrapper;
  }

  /** 把一个修改块（hunk）应用到段落上 */
  private applyHunk(p: ParaEntry, h: Hunk, o: EditOptions): void {
    const w = this.w(p.part);
    let { segs } = this.segments(p);
    for (const s of segs) {
      if (s.kind === "atom" && s.start < h.end && s.end > h.start) {
        throw new DocError(
          `修改范围包含不可编辑对象 ${s.text}。请只修改它前面或后面的文字，并在 old_text/new_text 中原样保留 ${s.text}。`,
          "protected"
        );
      }
      if (s.kind === "zw" && s.hard && s.start > h.start && s.start < h.end) {
        throw new DocError("修改范围跨越了域（自动编号/交叉引用/目录等）或批注的边界，请缩小修改范围，分别修改边界两侧的文字。", "protected");
      }
    }

    if (h.start < h.end) {
      this.splitAt(p, h.start);
      this.splitAt(p, h.end);
      segs = this.segments(p).segs;
    }

    // 汇总每个 run 覆盖的区间
    const info = new Map<XElement, { min: number; max: number; content: boolean; hard: boolean; order: number }>();
    segs.forEach((s, idx) => {
      if (!s.run) return;
      const cur = info.get(s.run) ?? { min: Infinity, max: -Infinity, content: false, hard: false, order: idx };
      cur.min = Math.min(cur.min, s.start);
      cur.max = Math.max(cur.max, s.end);
      if (s.end > s.start) cur.content = true;
      if (s.hard) cur.hard = true;
      info.set(s.run, cur);
    });
    const runs = [...info.entries()].sort((a, b) => a[1].order - b[1].order);

    const covered = h.start < h.end
      ? runs.filter(([, i]) =>
          (i.content && i.min >= h.start && i.max <= h.end && !i.hard) ||
          (!i.content && !i.hard && i.min > h.start && i.max < h.end)
        ).map(([r]) => r)
      : [];

    const contentRuns = runs.filter(([, i]) => i.content);
    const preceding = [...contentRuns].reverse().find(([, i]) => i.max <= h.start)?.[0] ?? null;
    const following = contentRuns.find(([, i]) => i.min >= h.end)?.[0] ?? null;
    const firstCoveredContent = covered.find((r) => info.get(r)!.content) ?? null;

    const formatSource = firstCoveredContent ?? preceding ?? following;
    const rPrXml = formatSource ? this.rPrXmlOf(formatSource, p.part) : this.paraMarkRPrXml(p);
    const newRun = h.text ? this.buildRunXml(w, rPrXml, h.text) : "";
    const wrapIns = (xml: string) => (o.track ? `<${w}:ins ${this.revAttrs(w, o)}>${xml}</${w}:ins>` : xml);

    if (covered.length) {
      const first = covered[0];
      const firstParent = first.parent;
      if (!o.track) {
        if (newRun) insertBefore(first, parseFragment(newRun));
        for (const r of covered) this.removeRunAndCleanup(r, p.el);
      } else {
        // 修订模式：先删除（<w:del>），再在其后插入（<w:ins>），与 Word 自身的修订顺序一致。
        // 新的 <w:ins> 不能嵌套在别人的 <w:ins> 里，所以插入点总是提升到修订容器之外。
        void firstParent;
        let placeholder: XNode | null = null;
        if (newRun) {
          [placeholder] = parseFragment(`<${w}:proofErr ${w}:type="__anchor"/>`);
          insertBefore(this.anchorOf(first, p.el), [placeholder]);
        }
        let lastWrapper: XNode | null = null;
        for (const r of covered) {
          const wrapper = this.trackDeleteRun(r, p, o);
          if (wrapper) lastWrapper = wrapper;
        }
        if (newRun && placeholder) {
          const ins = parseFragment(wrapIns(newRun));
          if (lastWrapper && lastWrapper.parent) {
            let target: XNode = lastWrapper;
            while (target.parent && target.parent !== p.el && target.parent.type === "el" && ["ins", "moveTo"].includes(target.parent.local)) {
              target = target.parent;
            }
            insertAfter(target, ins);
          } else insertBefore(placeholder, ins);
          removeNode(placeholder);
        }
      }
      return;
    }

    if (!newRun) return;
    const frag = parseFragment(wrapIns(newRun));
    const inLinkLike = (r: XElement | null) => !!r && r.parent !== p.el && r.parent?.type === "el" && ["hyperlink", "fldSimple"].includes(r.parent.local);
    if (preceding && !(inLinkLike(preceding) && following && following.parent !== preceding.parent)) {
      insertAfter(o.track ? this.anchorOf(preceding, p.el) : preceding, frag);
    } else if (following) {
      insertBefore(o.track ? this.anchorOf(following, p.el) : following, frag);
    } else {
      appendChild(p.el, frag);
    }
  }

  private ensureChild(parent: XElement, local: string, order: string[], w: string, first = false): XElement {
    const existing = parent.child(`${w}:${local}`);
    if (existing) return existing;
    const [el] = parseFragment(`<${w}:${local}/>`) as XElement[];
    const idx = first ? 0 : orderedInsertIndex(parent, local, order, w);
    parent.children.splice(idx, 0, el);
    el.parent = parent;
    markDirty(parent);
    if (parent.selfClosing) parent.attrsDirty = true;
    return el;
  }

  private setProp(container: XElement, local: string, attrs: Record<string, string> | null, order: string[], w: string): void {
    const old = container.child(`${w}:${local}`);
    if (old) removeNode(old);
    if (attrs === null) return;
    const attrXml = Object.entries(attrs).map(([k, v]) => ` ${w}:${k}="${encodeAttr(v)}"`).join("");
    const [el] = parseFragment(`<${w}:${local}${attrXml}/>`);
    const idx = orderedInsertIndex(container, local, order, w);
    container.children.splice(idx, 0, el);
    el.parent = container;
    if (container.selfClosing) container.attrsDirty = true;
    markDirty(container);
  }

  private ensurePPr(pEl: XElement, w: string): XElement {
    return this.ensureChild(pEl, "pPr", [], w, true);
  }

  private findRange(p: ParaEntry, text?: string): { start: number; end: number } {
    const full = this.textOf(p);
    if (!text) return { start: 0, end: full.length };
    const ms = findAll(full, text);
    if (ms.length === 0) throw new DocError(`在段落 ${p.ref} 中找不到"${clip(text, 60)}"。段落原文："${clip(full, 160)}"`, "not_found");
    if (ms.length > 1) throw new DocError(`"${clip(text, 60)}"在段落 ${p.ref} 中出现了 ${ms.length} 次，请提供更长、唯一的片段。`, "ambiguous");
    return ms[0];
  }

  private coveredRuns(p: ParaEntry, start: number, end: number): XElement[] {
    const { segs } = this.segments(p);
    const out: XElement[] = [];
    for (const s of segs) {
      if (!s.run || s.end <= s.start) continue;
      if (s.start >= start && s.end <= end && !out.includes(s.run)) out.push(s.run);
    }
    return out;
  }

  private isEditable(p: ParaEntry): void {
    const w = this.w(p.part);
    if (p.el.closest(`${w}:del`) || p.el.closest(`${w}:moveFrom`)) {
      throw new DocError(`段落 ${p.ref} 位于已删除的修订内容中，不能编辑`, "protected");
    }
  }

  private newParaId(): string {
    const used = new Set(this.paras.map((p) => p.paraId).filter(Boolean));
    for (;;) {
      const id = Math.floor(Math.random() * 0x7fffffff).toString(16).toUpperCase().padStart(8, "0");
      if (!used.has(id)) return id;
    }
  }

  private supportsW14(part: string): boolean {
    return this.part(part).root.attrs.some(([k, v]) => k === "xmlns:w14" && v === NS.w14);
  }

  // -------------------------------------------------------------------------
  // 写接口
  // -------------------------------------------------------------------------

  replaceText(p: { ref?: string; oldText: string; newText: string; replaceAll?: boolean }, o: EditOptions): EditResult {
    if (!p.oldText) throw new DocError("old_text 不能为空。若要在段落中插入文字，请把插入点前后的一小段文字作为 old_text。", "invalid");
    if (p.oldText === p.newText) throw new DocError("old_text 与 new_text 完全相同，无需修改。", "invalid");
    const candidates = p.ref ? [this.resolve(p.ref)] : this.paras;
    const matches: Array<{ pe: ParaEntry; start: number; end: number }> = [];
    for (const pe of candidates) {
      for (const m of findAll(this.textOf(pe), p.oldText)) matches.push({ pe, start: m.start, end: m.end });
    }
    if (matches.length === 0) {
      const hint = p.oldText.includes("\n")
        ? " 注意：段落之间的换行不能出现在 old_text 中，跨段落的修改请逐段进行。"
        : "";
      const scope = p.ref ? `段落 ${p.ref}（原文："${clip(this.textOf(candidates[0]), 200)}"）` : "整篇文档";
      throw new DocError(`在${scope}中找不到 old_text"${clip(p.oldText, 80)}"。请先用 doc_search 或 doc_read 确认原文。${hint}`, "not_found");
    }
    if (matches.length > 1 && !p.replaceAll) {
      const list = matches.slice(0, 6).map((m) => `  ${m.pe.ref}（${m.pe.location}）："${clip(this.textOf(m.pe), 70)}"`).join("\n");
      throw new DocError(
        `old_text 出现了 ${matches.length} 次，无法确定修改哪一处。请传入 ref 指定段落，或提供更长的唯一片段；若确实要全部替换，设置 replace_all=true。\n${list}`,
        "ambiguous"
      );
    }

    const byPara = new Map<ParaEntry, Array<{ start: number; end: number }>>();
    for (const m of matches) {
      this.isEditable(m.pe);
      const arr = byPara.get(m.pe) ?? [];
      arr.push({ start: m.start, end: m.end });
      byPara.set(m.pe, arr);
    }
    const preview: EditResult["preview"] = [];
    for (const [pe, ms] of byPara) {
      const before = this.textOf(pe);
      const hunks: Hunk[] = [];
      for (const m of ms) {
        const actual = before.slice(m.start, m.end);
        for (const h of computeHunks(actual, p.newText)) hunks.push({ start: h.start + m.start, end: h.end + m.start, text: h.text });
      }
      hunks.sort((a, b) => b.start - a.start);
      for (const h of hunks) this.applyHunk(pe, h, o);
      preview.push({ ref: pe.ref, before, after: this.textOf(pe) });
    }
    return {
      changedRefs: [...byPara.keys()].map((x) => x.ref),
      summary: `已修改 ${matches.length} 处（${byPara.size} 个段落）${o.track ? "，以修订形式记录" : ""}。`,
      preview,
      structural: false,
    };
  }

  formatText(p: { ref: string; text?: string; format: RunFormat }, o: EditOptions): EditResult {
    const pe = this.resolve(p.ref);
    this.isEditable(pe);
    const w = this.w(pe.part);
    const { start, end } = this.findRange(pe, p.text);
    if (start === end) throw new DocError("段落为空，没有可设置格式的文字", "invalid");
    this.splitAt(pe, start);
    this.splitAt(pe, end);
    const runs = this.coveredRuns(pe, start, end);
    const f = p.format;
    let charStyleId: string | undefined;
    if (f.char_style) {
      const s = this.styles.resolve(f.char_style, "character");
      if (!s) throw new DocError(`找不到字符样式"${f.char_style}"。可用 doc_styles 查看可用样式。`, "not_found");
      charStyleId = s.id;
    }
    for (const run of runs) {
      const rPr = this.ensureChild(run, "rPr", [], w, true);
      if (o.track && !rPr.child(`${w}:rPrChange`)) {
        const old = rPr.children.map((c) => serializeNode(c, this.src(pe.part))).join("");
        const [chg] = parseFragment(`<${w}:rPrChange ${this.revAttrs(w, o)}><${w}:rPr>${old}</${w}:rPr></${w}:rPrChange>`);
        rPr.children.push(chg);
        chg.parent = rPr;
        markDirty(rPr);
      }
      const bool = (local: string, v: boolean | undefined, cs?: string) => {
        if (v === undefined) return;
        this.setProp(rPr, local, v ? {} : { val: "0" }, RPR_ORDER, w);
        if (cs) this.setProp(rPr, cs, v ? {} : { val: "0" }, RPR_ORDER, w);
      };
      if (charStyleId) this.setProp(rPr, "rStyle", { val: charStyleId }, RPR_ORDER, w);
      bool("b", f.bold, "bCs");
      bool("i", f.italic, "iCs");
      bool("strike", f.strike);
      bool("caps", f.all_caps);
      bool("smallCaps", f.small_caps);
      if (f.underline) this.setProp(rPr, "u", { val: f.underline }, RPR_ORDER, w);
      if (f.color) this.setProp(rPr, "color", { val: f.color.replace(/^#/, "").toUpperCase() }, RPR_ORDER, w);
      if (f.highlight) this.setProp(rPr, "highlight", { val: f.highlight }, RPR_ORDER, w);
      if (f.size_pt) {
        const hp = String(Math.round(f.size_pt * 2));
        this.setProp(rPr, "sz", { val: hp }, RPR_ORDER, w);
        this.setProp(rPr, "szCs", { val: hp }, RPR_ORDER, w);
      }
      if (f.font) this.setProp(rPr, "rFonts", { ascii: f.font, hAnsi: f.font, eastAsia: f.font, cs: f.font }, RPR_ORDER, w);
      if (f.vertical) this.setProp(rPr, "vertAlign", { val: f.vertical }, RPR_ORDER, w);
    }
    const txt = this.textOf(pe).slice(start, end);
    return {
      changedRefs: [pe.ref],
      summary: `已为段落 ${pe.ref} 中"${clip(txt, 40)}"设置格式${o.track ? "（记录为格式修订）" : ""}。`,
      preview: [{ ref: pe.ref, before: `（格式）${clip(txt, 80)}`, after: `（新格式）${JSON.stringify(f)}` }],
      structural: false,
    };
  }

  setParagraph(p: { refs: string[]; format: ParagraphFormat }, o: EditOptions): EditResult {
    const f = p.format;
    let styleId: string | undefined;
    if (f.style) {
      const s = this.styles.resolve(f.style, "paragraph");
      if (!s) {
        const near = this.listStyles().filter((x) => x.type === "paragraph").slice(0, 15).map((x) => x.name).join("、");
        throw new DocError(`找不到段落样式"${f.style}"。常用样式：${near}`, "not_found");
      }
      styleId = s.id;
    }
    const preview: EditResult["preview"] = [];
    for (const ref of p.refs) {
      const pe = this.resolve(ref);
      this.isEditable(pe);
      const w = this.w(pe.part);
      const before = this.blockInfo(pe).style ?? "";
      const pPr = this.ensurePPr(pe.el, w);
      if (o.track && !pPr.child(`${w}:pPrChange`)) {
        const skip = new Set(["rPr", "sectPr", "pPrChange"]);
        const old = pPr.children
          .filter((c) => !(c.type === "el" && skip.has(c.local)))
          .map((c) => serializeNode(c, this.src(pe.part))).join("");
        const [chg] = parseFragment(`<${w}:pPrChange ${this.revAttrs(w, o)}><${w}:pPr>${old}</${w}:pPr></${w}:pPrChange>`);
        pPr.children.push(chg);
        chg.parent = pPr;
        markDirty(pPr);
      }
      if (styleId) this.setProp(pPr, "pStyle", { val: styleId }, PPR_ORDER, w);
      if (f.alignment) {
        const map = { left: "left", center: "center", right: "right", justify: "both" } as const;
        this.setProp(pPr, "jc", { val: map[f.alignment] }, PPR_ORDER, w);
      }
      const mergeAttrs = (local: string, attrs: Record<string, string | undefined>) => {
        const defined = Object.entries(attrs).filter((e): e is [string, string] => e[1] !== undefined);
        if (!defined.length) return;
        const el = this.ensureChild(pPr, local, PPR_ORDER, w);
        for (const [k, v] of defined) el.setAttr(`${w}:${k}`, v);
      };
      mergeAttrs("spacing", {
        before: f.space_before_pt !== undefined ? String(ptToTwips(f.space_before_pt)) : undefined,
        after: f.space_after_pt !== undefined ? String(ptToTwips(f.space_after_pt)) : undefined,
        line: f.line_spacing !== undefined ? String(Math.round(f.line_spacing * 240)) : undefined,
        lineRule: f.line_spacing !== undefined ? "auto" : undefined,
      });
      mergeAttrs("ind", {
        left: f.indent_left_pt !== undefined ? String(ptToTwips(f.indent_left_pt)) : undefined,
        right: f.indent_right_pt !== undefined ? String(ptToTwips(f.indent_right_pt)) : undefined,
        firstLine: f.first_line_pt !== undefined ? String(ptToTwips(f.first_line_pt)) : undefined,
        hanging: f.hanging_pt !== undefined ? String(ptToTwips(f.hanging_pt)) : undefined,
      });
      if (f.first_line_pt !== undefined) pPr.child(`${w}:ind`)?.removeAttr(`${w}:hanging`);
      if (f.hanging_pt !== undefined) pPr.child(`${w}:ind`)?.removeAttr(`${w}:firstLine`);
      if (f.keep_with_next !== undefined) this.setProp(pPr, "keepNext", f.keep_with_next ? {} : { val: "0" }, PPR_ORDER, w);
      if (f.page_break_before !== undefined) this.setProp(pPr, "pageBreakBefore", f.page_break_before ? {} : { val: "0" }, PPR_ORDER, w);
      preview.push({ ref: pe.ref, before: `样式 ${before}`, after: `样式 ${this.blockInfo(pe).style}；${JSON.stringify(f)}` });
    }
    return { changedRefs: p.refs, summary: `已调整 ${p.refs.length} 个段落的段落格式${o.track ? "（记录为格式修订）" : ""}。`, preview, structural: false };
  }

  insertBlocks(p: { anchor: string; position: "before" | "after"; blocks: NewBlock[] }, o: EditOptions): EditResult {
    if (!p.blocks.length) throw new DocError("blocks 不能为空", "invalid");
    const anchor = this.resolve(p.anchor);
    const w = this.w(anchor.part);
    const src = this.src(anchor.part);
    const w14 = this.supportsW14(anchor.part);
    const styleIds: Array<string | undefined> = [];
    let xml = "";
    for (const b of p.blocks) {
      let styleId: string | undefined;
      if (b.style) {
        const s = this.styles.resolve(b.style, "paragraph");
        if (!s) throw new DocError(`找不到段落样式"${b.style}"。可用 doc_styles 查看。`, "not_found");
        styleId = s.id;
      }
      // 格式模板的选择顺序：显式 like → 文档中离锚点最近的同样式段落 → 锚点本身。
      // 用同样式的现有段落做模板，新段落的段落格式和文字格式就与文档中已有的同类段落完全一致。
      let like: ParaEntry = anchor;
      let minimal = false;
      if (b.like) like = this.resolve(b.like);
      else if (styleId && styleId !== this.styleIdOf(anchor)) {
        const tpl = this.findStyleTemplate(styleId, anchor);
        if (tpl) like = tpl;
        else minimal = true; // 文档中没有该样式的段落：只设置样式，其余交给样式定义
      }
      if (like.part !== anchor.part) throw new DocError("格式模板段落必须和锚点段落位于同一部件（正文/页眉/脚注）", "invalid");
      let pPrXml = "";
      const likePPr = minimal ? undefined : like.el.child(`${w}:pPr`);
      const insMark = o.track ? `<${w}:ins ${this.revAttrs(w, o)}/>` : "";
      if (likePPr) {
        const kids = likePPr.children.filter((c) => !(c.type === "el" && ["sectPr", "pPrChange", "rPr"].includes(c.local)));
        let inner = kids.map((k) => serializeNode(k, src)).join("");
        const markRPr = this.paraMarkRPrXml(like).replace(new RegExp(`^<${w}:rPr>|</${w}:rPr>$`, "g"), "");
        if (markRPr || insMark) inner += `<${w}:rPr>${insMark}${markRPr}</${w}:rPr>`;
        pPrXml = `<${w}:pPr>${inner}</${w}:pPr>`;
      } else if (insMark) {
        pPrXml = `<${w}:pPr><${w}:rPr>${insMark}</${w}:rPr></${w}:pPr>`;
      }
      const rPrXml = minimal ? "" : this.dominantRPrXml(like);
      let runXml = b.text ? this.buildRunXml(w, rPrXml, b.text) : "";
      if (runXml && o.track) runXml = `<${w}:ins ${this.revAttrs(w, o)}>${runXml}</${w}:ins>`;
      const id = w14 ? this.newParaId() : "";
      const pAttrs = id ? ` w14:paraId="${id}" w14:textId="77777777"` : "";
      xml += `<${w}:p${pAttrs}>${pPrXml}${runXml}</${w}:p>`;
      styleIds.push(styleId);
    }
    const nodes = parseFragment(xml) as XElement[];
    if (p.position === "before") insertBefore(anchor.el, nodes);
    else insertAfter(anchor.el, nodes);
    nodes.forEach((n, i) => {
      const sid = styleIds[i];
      if (sid) {
        const pPr = this.ensurePPr(n, w);
        this.setProp(pPr, "pStyle", { val: sid }, PPR_ORDER, w);
      }
    });
    this._structureVersion++;
    this.reindex();
    const refs = nodes.map((n) => this.paras.find((x) => x.el === n)!.ref);
    return {
      changedRefs: refs,
      summary: `已在 ${anchor.ref} ${p.position === "before" ? "之前" : "之后"}插入 ${nodes.length} 个段落，新段落引用：${refs.join("、")}${o.track ? "（记录为插入修订）" : ""}。文档结构已变化，旧的 P 序号引用失效（#开头的引用仍有效）。`,
      preview: refs.map((r, i) => ({ ref: r, before: "", after: p.blocks[i].text })),
      structural: true,
    };
  }

  deleteBlocks(p: { refs: string[] }, o: EditOptions): EditResult {
    const entries = p.refs.map((r) => this.resolve(r));
    for (const pe of entries) {
      this.isEditable(pe);
      const w = this.w(pe.part);
      if (pe.el.child(`${w}:pPr`)?.child(`${w}:sectPr`)) {
        throw new DocError(`段落 ${pe.ref} 携带分节符（页面设置/分栏/页眉页脚绑定），删除会破坏版式。可以用 doc_replace_text 清空其文字。`, "protected");
      }
      const parent = pe.el.parent as XElement;
      if (parent.local === "tc" && parent.childrenNamed(`${w}:p`).length <= 1) {
        throw new DocError(`段落 ${pe.ref} 是表格单元格中唯一的段落，Word 要求单元格至少有一个段落。可以用 doc_replace_text 清空其文字。`, "protected");
      }
    }
    const preview = entries.map((pe) => ({ ref: pe.ref, before: this.textOf(pe), after: "" }));
    for (const pe of entries) {
      const w = this.w(pe.part);
      if (!o.track) {
        removeNode(pe.el);
        continue;
      }
      const runs: XElement[] = [];
      const collect = (n: XElement) => {
        for (const c of n.children) {
          if (c.type !== "el" || c.prefix !== w) continue;
          if (c.local === "r") runs.push(c);
          else if (TRANSPARENT_CONTAINERS.has(c.local)) collect(c);
        }
      };
      collect(pe.el);
      for (const r of runs) this.trackDeleteRun(r, pe, o);
      const pPr = this.ensurePPr(pe.el, w);
      const rPr = this.ensureChild(pPr, "rPr", PPR_ORDER, w);
      const [del] = parseFragment(`<${w}:del ${this.revAttrs(w, o)}/>`);
      rPr.children.unshift(del);
      del.parent = rPr;
      markDirty(rPr);
    }
    this._structureVersion++;
    this.reindex();
    return {
      changedRefs: entries.map((e) => e.ref),
      summary: `已删除 ${entries.length} 个段落${o.track ? "（记录为删除修订，可在 Word 中拒绝以恢复）" : ""}。文档结构已变化，请重新读取以获得最新引用。`,
      preview,
      structural: true,
    };
  }

  insertTableRow(p: { ref: string; position: "before" | "after"; cells: string[] }, o: EditOptions): EditResult {
    const pe = this.resolve(p.ref);
    const w = this.w(pe.part);
    const src = this.src(pe.part);
    const tr = pe.el.closest(`${w}:tr`);
    if (!tr) throw new DocError(`段落 ${pe.ref} 不在表格中。请传入目标表格某一行中任意单元格段落的引用。`, "invalid");
    const tcs = tr.childrenNamed(`${w}:tc`);
    if (p.cells.length !== tcs.length) {
      throw new DocError(`该行有 ${tcs.length} 个单元格，但 cells 提供了 ${p.cells.length} 个值，请一一对应（可用空字符串）。`, "invalid");
    }
    const [row] = parseFragment(serializeNode(tr, src)) as XElement[];
    // 去掉纵向合并标记，避免新行意外并入上方合并单元格
    for (const d of [...row.descendants()]) if (d.local === "vMerge") removeNode(d);
    const w14 = this.supportsW14(pe.part);
    row.childrenNamed(`${w}:tc`).forEach((tc, i) => {
      const paras = tc.childrenNamed(`${w}:p`);
      const firstP = paras[0];
      paras.slice(1).forEach((x) => removeNode(x));
      const templateRun = [...firstP.descendants()].find((d) => d.local === "r" && d.prefix === w && d.child(`${w}:t`));
      const rPr = templateRun?.child(`${w}:rPr`);
      const rPrXml = rPr ? serializeNode(rPr, "") : "";
      for (const c of [...firstP.children]) if (!(c.type === "el" && c.local === "pPr")) removeNode(c);
      if (w14) {
        firstP.setAttr("w14:paraId", this.newParaId());
        firstP.setAttr("w14:textId", "77777777");
      }
      if (p.cells[i]) {
        let run = this.buildRunXml(w, rPrXml, p.cells[i]);
        if (o.track) run = `<${w}:ins ${this.revAttrs(w, o)}>${run}</${w}:ins>`;
        appendChild(firstP, parseFragment(run));
      }
    });
    if (o.track) {
      let trPr = row.child(`${w}:trPr`);
      if (!trPr) {
        [trPr] = parseFragment(`<${w}:trPr/>`) as XElement[];
        const ex = row.child(`${w}:tblPrEx`);
        if (ex) insertAfter(ex, [trPr]);
        else {
          row.children.unshift(trPr);
          trPr.parent = row;
        }
      }
      const old = trPr.child(`${w}:ins`);
      if (old) removeNode(old);
      const [insMark] = parseFragment(`<${w}:ins ${this.revAttrs(w, o)}/>`);
      trPr.children.splice(orderedInsertIndex(trPr, "ins", TRPR_ORDER, w), 0, insMark);
      insMark.parent = trPr;
      trPr.attrsDirty = true;
      markDirty(trPr);
    }
    if (p.position === "before") insertBefore(tr, [row]);
    else insertAfter(tr, [row]);
    this._structureVersion++;
    this.reindex();
    const refs = row.childrenNamed(`${w}:tc`).map((tc) => this.paras.find((x) => x.el === tc.childrenNamed(`${w}:p`)[0])!.ref);
    return {
      changedRefs: refs,
      summary: `已插入表格行（${tcs.length} 列），单元格引用：${refs.join("、")}${o.track ? "（记录为插入修订）" : ""}。`,
      preview: [{ ref: refs[0], before: "", after: p.cells.join(" | ") }],
      structural: true,
    };
  }

  private commentsPartName(): string | null {
    const rels = parseRels(this.optText(relsPathFor(this.mainPart)));
    const r = rels.find((x) => x.type === REL_TYPES.comments);
    return r ? resolveTarget(this.mainPart, r.target) : null;
  }

  private ensureCommentsPart(): string {
    const existing = this.commentsPartName();
    if (existing && this.pkg.has(existing)) return existing;
    const name = this.mainPart.replace(/[^/]+$/, "comments.xml");
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:comments xmlns:w="${NS.w}"></w:comments>`;
    this.pkg.write(name, Buffer.from(xml, "utf8"));

    const relsName = relsPathFor(this.mainPart);
    const relsDoc = this.part(relsName);
    const ids = new Set(relsDoc.root.elements().map((e) => e.getAttr("Id")));
    let n = ids.size + 1;
    while (ids.has(`rId${n}`)) n++;
    appendChild(relsDoc.root, parseFragment(`<Relationship Id="rId${n}" Type="${REL_TYPES.comments}" Target="comments.xml"/>`));

    const ct = this.part("[Content_Types].xml");
    appendChild(ct.root, parseFragment(`<Override PartName="/${name}" ContentType="${CT_COMMENTS}"/>`));
    this.prefixes.delete(name);
    this.parts.delete(name);
    return name;
  }

  addComment(p: { ref: string; text?: string; comment: string }, o: EditOptions): EditResult {
    const pe = this.resolve(p.ref);
    if (pe.part !== this.mainPart) throw new DocError("目前只支持在正文段落上添加批注", "unsupported");
    const w = this.w(pe.part);
    const { start, end } = this.findRange(pe, p.text);
    if (end > start) {
      this.splitAt(pe, start);
      this.splitAt(pe, end);
    }
    const runs = this.coveredRuns(pe, start, end);
    const commentsPart = this.ensureCommentsPart();
    const cw = this.w(commentsPart);
    const id = this.revId();
    const initials = o.author.replace(/[^A-Za-z一-鿿]/g, "").slice(0, 2) || "AI";
    const paras = p.comment.split(/\n+/).map((line) =>
      `<${cw}:p><${cw}:r><${cw}:annotationRef/></${cw}:r><${cw}:r>${this.textElementXml(`${cw}:t`, line)}</${cw}:r></${cw}:p>`
    );
    appendChild(
      this.part(commentsPart).root,
      parseFragment(`<${cw}:comment ${cw}:id="${id}" ${cw}:author="${encodeAttr(o.author)}" ${cw}:date="${o.date}" ${cw}:initials="${encodeAttr(initials)}">${paras.join("")}</${cw}:comment>`)
    );
    const startMark = parseFragment(`<${w}:commentRangeStart ${w}:id="${id}"/>`);
    const endMarks = parseFragment(`<${w}:commentRangeEnd ${w}:id="${id}"/><${w}:r><${w}:commentReference ${w}:id="${id}"/></${w}:r>`);
    const top = (n: XElement): XNode => {
      let x: XElement = n;
      while (x.parent && x.parent !== pe.el && x.parent.type === "el") x = x.parent;
      return x;
    };
    if (runs.length) {
      insertBefore(top(runs[0]), startMark);
      insertAfter(top(runs[runs.length - 1]), endMarks);
    } else {
      const pPr = pe.el.child(`${w}:pPr`);
      if (pPr) insertAfter(pPr, startMark);
      else {
        pe.el.children.unshift(...startMark);
        startMark.forEach((m) => (m.parent = pe.el));
      }
      appendChild(pe.el, endMarks);
    }
    const target = this.textOf(pe).slice(start, end);
    return {
      changedRefs: [pe.ref],
      summary: `已在段落 ${pe.ref}${p.text ? `的"${clip(target, 40)}"` : ""}上添加批注。`,
      preview: [{ ref: pe.ref, before: clip(target, 80), after: `💬 ${p.comment}` }],
      structural: false,
    };
  }

  // =========================================================================
  // 扩展编辑能力：表格 / 图片 / 图形 / 图表 / 公式 / 分隔符 / 页面设置 / 页眉页脚 /
  //               列表 / 脚注 / 超链接 / 目录 / 样式 / 修订审阅 / 批注管理 / 移动段落
  // =========================================================================

  private docPrCounter = -1;

  private nextDocPrId(): number {
    if (this.docPrCounter < 0) {
      let max = 0;
      for (const part of this.contentParts) {
        for (const d of this.part(part).root.descendants()) {
          if (d.local === "docPr" || d.local === "cNvPr") {
            const id = Number(d.getAttr("id"));
            if (Number.isFinite(id) && id > max) max = id;
          }
        }
      }
      this.docPrCounter = max + 1;
    }
    return this.docPrCounter++;
  }

  private relsDoc(owner = this.mainPart): XDocument {
    const name = relsPathFor(owner);
    if (!this.pkg.has(name)) {
      this.pkg.write(name, Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${NS.rels}"></Relationships>`, "utf8"));
      this.parts.delete(name);
    }
    return this.part(name);
  }

  private addRel(type: string, target: string, opts: { owner?: string; external?: boolean } = {}): string {
    const doc = this.relsDoc(opts.owner);
    const ids = new Set(doc.root.elements().map((e) => e.getAttr("Id")));
    let n = ids.size + 1;
    while (ids.has(`rId${n}`)) n++;
    appendChild(doc.root, parseFragment(`<Relationship Id="rId${n}" Type="${type}" Target="${encodeAttr(target)}"${opts.external ? ' TargetMode="External"' : ""}/>`));
    return `rId${n}`;
  }

  private ensureDefaultCT(ext: string, ct: string): void {
    const doc = this.part("[Content_Types].xml");
    if (doc.root.elements().some((e) => e.local === "Default" && e.getAttr("Extension")?.toLowerCase() === ext.toLowerCase())) return;
    const first = doc.root.elements()[0];
    const node = parseFragment(`<Default Extension="${ext}" ContentType="${ct}"/>`);
    if (first) insertBefore(first, node);
    else appendChild(doc.root, node);
  }

  private ensureOverrideCT(partName: string, ct: string): void {
    const doc = this.part("[Content_Types].xml");
    if (doc.root.elements().some((e) => e.local === "Override" && e.getAttr("PartName") === `/${partName}`)) return;
    appendChild(doc.root, parseFragment(`<Override PartName="/${partName}" ContentType="${ct}"/>`));
  }

  /** 新部件的唯一名称，如 word/media/image3.png */
  private uniquePartName(dir: string, base: string, ext: string): string {
    for (let n = 1; ; n++) {
      const name = `${dir}/${base}${n}.${ext}`;
      if (!this.pkg.has(name)) return name;
    }
  }

  /** 部件相对主文档的路径（关系 Target 用） */
  private relTarget(partName: string, owner = this.mainPart): string {
    const dir = owner.includes("/") ? owner.slice(0, owner.lastIndexOf("/") + 1) : "";
    return partName.startsWith(dir) ? partName.slice(dir.length) : `/${partName}`;
  }

  private addPart(name: string, data: Buffer | string): void {
    this.pkg.write(name, typeof data === "string" ? Buffer.from(data, "utf8") : data);
    this.parts.delete(name);
    this.prefixes.delete(name);
  }

  private body(): XElement {
    const w = this.w(this.mainPart);
    const b = this.part(this.mainPart).root.child(`${w}:body`);
    if (!b) throw new DocError("文档缺少正文（w:body）", "invalid");
    return b;
  }

  /** 段落所属的分节属性（之后第一个带 sectPr 的正文段落，或文末的 sectPr） */
  private sectPrFor(pe?: ParaEntry): XElement | undefined {
    const w = this.w(this.mainPart);
    const body = this.body();
    if (pe && pe.part === this.mainPart) {
      let top: XElement = pe.el;
      while (top.parent && top.parent !== body && top.parent.type === "el") top = top.parent as XElement;
      const kids = body.elements();
      for (let k = Math.max(0, kids.indexOf(top)); k < kids.length; k++) {
        const sp = kids[k].local === "p" ? kids[k].child(`${w}:pPr`)?.child(`${w}:sectPr`) : undefined;
        if (sp) return sp;
      }
    }
    return body.child(`${w}:sectPr`);
  }

  private allSectPrs(): XElement[] {
    const w = this.w(this.mainPart);
    const out: XElement[] = [];
    for (const d of this.body().descendants()) if (d.prefix === w && d.local === "sectPr" && !d.closest(`${w}:sectPrChange`)) out.push(d);
    return out;
  }

  /** 可用文字宽度（twips）：页宽 − 左右边距；多栏时为单栏宽度；表格内为单元格宽度 */
  private textWidthTwips(pe?: ParaEntry): number {
    const w = this.w(this.mainPart);
    const sp = this.sectPrFor(pe);
    const pgW = Number(sp?.child(`${w}:pgSz`)?.getAttr(`${w}:w`) ?? 11906);
    const mar = sp?.child(`${w}:pgMar`);
    const num = (k: string, d: number) => { const v = Number(mar?.getAttr(`${w}:${k}`)); return Number.isFinite(v) ? v : d; };
    let width = pgW - num("left", 1800) - num("right", 1800) - num("gutter", 0);
    const cols = sp?.child(`${w}:cols`);
    const n = Number(cols?.getAttr(`${w}:num`) ?? 1);
    if (n > 1) width = (width - Number(cols?.getAttr(`${w}:space`) ?? 425) * (n - 1)) / n;
    if (pe?.container === "table") {
      const tcW = pe.el.closest(`${w}:tc`)?.child(`${w}:tcPr`)?.child(`${w}:tcW`);
      const v = Number(tcW?.getAttr(`${w}:w`));
      if (tcW?.getAttr(`${w}:type`) === "dxa" && v > 0) width = v - 216;
    }
    return Math.max(1440, Math.round(width));
  }

  private insMark(w: string, o: EditOptions): string {
    return o.track ? `<${w}:ins ${this.revAttrs(w, o)}/>` : "";
  }

  private trackRun(w: string, runXml: string, o: EditOptions): string {
    return o.track && runXml ? `<${w}:ins ${this.revAttrs(w, o)}>${runXml}</${w}:ins>` : runXml;
  }

  /** 生成一个新段落：pPr 内容（按 schema 顺序给出）+ 段内内容；修订模式下标记为插入 */
  private paraXml(part: string, pPrInner: string, content: string, o: EditOptions): string {
    const w = this.w(part);
    const id = this.supportsW14(part) ? ` w14:paraId="${this.newParaId()}" w14:textId="77777777"` : "";
    const mark = this.insMark(w, o);
    const pPr = pPrInner || mark ? `<${w}:pPr>${pPrInner}${mark ? `<${w}:rPr>${mark}</${w}:rPr>` : ""}</${w}:pPr>` : "";
    return `<${w}:p${id}>${pPr}${content}</${w}:p>`;
  }

  /** 锚点段落的字体与字号（新表格、题注沿用正文字体） */
  private baseFontRPr(pe: ParaEntry, sizePt?: number, extra = ""): string {
    const w = this.w(pe.part);
    const dom = this.dominantRPrXml(pe);
    const fonts = new RegExp(`<${w}:rFonts[^>]*/>`).exec(dom)?.[0] ?? "";
    const sz = sizePt ? `<${w}:sz ${w}:val="${Math.round(sizePt * 2)}"/><${w}:szCs ${w}:val="${Math.round(sizePt * 2)}"/>` : (new RegExp(`<${w}:sz [^>]*/>`).exec(dom)?.[0] ?? "") + (new RegExp(`<${w}:szCs [^>]*/>`).exec(dom)?.[0] ?? "");
    const inner = fonts + extra + sz;
    return inner ? `<${w}:rPr>${inner}</${w}:rPr>` : "";
  }

  /** 题注段落：优先使用文档的"题注 / Caption"样式 */
  private captionXml(anchor: ParaEntry, text: string, o: EditOptions): string {
    const w = this.w(anchor.part);
    const st = this.styles.resolve("caption", "paragraph") ?? this.styles.resolve("题注", "paragraph");
    const pPr = st ? `<${w}:pStyle ${w}:val="${st.id}"/><${w}:jc ${w}:val="center"/>` : `<${w}:spacing ${w}:before="60" ${w}:after="120"/><${w}:jc ${w}:val="center"/>`;
    const rPr = st ? "" : this.baseFontRPr(anchor, 9);
    return this.paraXml(anchor.part, pPr, this.trackRun(w, `<${w}:r>${rPr}${runContent(w, text)}</${w}:r>`, o), o);
  }

  /** 把块级 XML 插到锚点段落前/后，重建索引，返回新插入的顶层元素 */
  private insertBlockXml(anchor: ParaEntry, position: "before" | "after", xml: string): XElement[] {
    this.isEditable(anchor);
    const w = this.w(anchor.part);
    const nodes = parseFragment(xml).filter((n): n is XElement => n.type === "el");
    if (position === "before") insertBefore(anchor.el, nodes);
    else insertAfter(anchor.el, nodes);
    const parent = anchor.el.parent as XElement;
    if (parent?.local === "tc") {
      const last = parent.elements().filter((e) => e.local !== "tcPr").pop();
      if (last?.local === "tbl") appendChild(parent, parseFragment(this.paraXml(anchor.part, "", "", { track: false, author: "", date: "" })));
    }
    void w;
    this._structureVersion++;
    this.reindex();
    return nodes;
  }

  private refsOf(nodes: XElement[]): string[] {
    const set = new Set<XElement>();
    for (const n of nodes) {
      if (n.local === "p") set.add(n);
      for (const d of n.descendants()) if (d.local === "p") set.add(d);
    }
    return this.paras.filter((p) => set.has(p.el)).map((p) => p.ref);
  }

  private objectParaXml(anchor: ParaEntry, runXml: string, align: "left" | "center" | "right", o: EditOptions, keepNext: boolean): string {
    const w = this.w(anchor.part);
    const pPr = `${keepNext ? `<${w}:keepNext/>` : ""}<${w}:spacing ${w}:before="120" ${w}:after="${keepNext ? 60 : 120}" ${w}:line="240" ${w}:lineRule="auto"/><${w}:ind ${w}:firstLine="0" ${w}:left="0" ${w}:right="0"/><${w}:jc ${w}:val="${align}"/>`;
    return this.paraXml(anchor.part, pPr, this.trackRun(w, runXml, o), o);
  }

  // ------------------------------- 表格 -------------------------------

  insertTable(p: {
    anchor: string; position: "before" | "after"; rows: string[][]; header?: boolean;
    style?: "grid" | "three_line" | "plain" | "banded"; align?: "left" | "center" | "right"; widths?: number[];
    width_pct?: number; font_size_pt?: number; caption?: string; caption_position?: "above" | "below";
  }, o: EditOptions): EditResult {
    const anchor = this.resolve(p.anchor);
    const w = this.w(anchor.part);
    const rows = p.rows.filter((r) => Array.isArray(r));
    if (!rows.length) throw new DocError("rows 不能为空", "invalid");
    const nCols = Math.max(...rows.map((r) => r.length));
    if (nCols < 1 || nCols > 63) throw new DocError("表格列数需在 1–63 之间", "invalid");
    const header = p.header ?? true;
    const style = p.style ?? "grid";
    const total = Math.round(this.textWidthTwips(anchor) * Math.min(1, Math.max(0.1, (p.width_pct ?? 100) / 100)));
    const rel = p.widths?.length === nCols && p.widths.every((x) => x > 0) ? p.widths : Array.from({ length: nCols }, () => 1);
    const sum = rel.reduce((a, b) => a + b, 0);
    const widths = rel.map((x) => Math.max(200, Math.round((total * x) / sum)));
    const B = (sz: number) => `${w}:val="single" ${w}:sz="${sz}" ${w}:space="0" ${w}:color="000000"`;
    const tblBorders = style === "grid" || style === "banded"
      ? `<${w}:tblBorders><${w}:top ${B(4)}/><${w}:left ${B(4)}/><${w}:bottom ${B(4)}/><${w}:right ${B(4)}/><${w}:insideH ${B(4)}/><${w}:insideV ${B(4)}/></${w}:tblBorders>`
      : style === "three_line"
        ? `<${w}:tblBorders><${w}:top ${B(12)}/><${w}:bottom ${B(12)}/></${w}:tblBorders>`
        : `<${w}:tblBorders><${w}:top ${w}:val="nil"/><${w}:left ${w}:val="nil"/><${w}:bottom ${w}:val="nil"/><${w}:right ${w}:val="nil"/><${w}:insideH ${w}:val="nil"/><${w}:insideV ${w}:val="nil"/></${w}:tblBorders>`;
    const jc = p.align ?? "center";
    const tblPr = `<${w}:tblPr><${w}:tblW ${w}:w="${total}" ${w}:type="dxa"/><${w}:jc ${w}:val="${jc}"/>${tblBorders}<${w}:tblLayout ${w}:type="fixed"/><${w}:tblCellMar><${w}:left ${w}:w="108" ${w}:type="dxa"/><${w}:right ${w}:w="108" ${w}:type="dxa"/></${w}:tblCellMar><${w}:tblLook ${w}:val="04A0" ${w}:firstRow="1" ${w}:lastRow="0" ${w}:firstColumn="1" ${w}:lastColumn="0" ${w}:noHBand="0" ${w}:noVBand="1"/></${w}:tblPr>`;
    const grid = `<${w}:tblGrid>${widths.map((x) => `<${w}:gridCol ${w}:w="${x}"/>`).join("")}</${w}:tblGrid>`;
    const rowXml = rows.map((r, ri) => {
      const isHead = header && ri === 0;
      const trPr = `<${w}:trPr><${w}:cantSplit/>${isHead ? `<${w}:tblHeader/>` : ""}${o.track ? `<${w}:ins ${this.revAttrs(w, o)}/>` : ""}</${w}:trPr>`;
      const cells = Array.from({ length: nCols }, (_, ci) => {
        const text = String(r[ci] ?? "");
        const shade = (style === "banded" && isHead) ? "D9E2F3" : (style === "grid" && isHead) ? "F2F2F2" : (style === "banded" && ri % 2 === 0 && !isHead) ? "F2F2F2" : "";
        const tcBorders = style === "three_line" && isHead ? `<${w}:tcBorders><${w}:bottom ${B(6)}/></${w}:tcBorders>` : "";
        const tcPr = `<${w}:tcPr><${w}:tcW ${w}:w="${widths[ci]}" ${w}:type="dxa"/>${tcBorders}${shade ? `<${w}:shd ${w}:val="clear" ${w}:color="auto" ${w}:fill="${shade}"/>` : ""}<${w}:vAlign ${w}:val="center"/></${w}:tcPr>`;
        const num = /^[\s\-+−]?[\d.,%]+\s*$/.test(text);
        const cellJc = isHead || num ? "center" : "left";
        const pPr = `<${w}:spacing ${w}:before="40" ${w}:after="40" ${w}:line="240" ${w}:lineRule="auto"/><${w}:ind ${w}:left="0" ${w}:right="0" ${w}:firstLine="0"/><${w}:jc ${w}:val="${cellJc}"/>`;
        const rPr = this.baseFontRPr(anchor, p.font_size_pt, isHead ? `<${w}:b/><${w}:bCs/>` : "");
        const content = text ? this.trackRun(w, `<${w}:r>${rPr}${runContent(w, text)}</${w}:r>`, o) : "";
        return `<${w}:tc>${tcPr}${this.paraXml(anchor.part, pPr, content, o)}</${w}:tc>`;
      }).join("");
      return `<${w}:tr>${trPr}${cells}</${w}:tr>`;
    }).join("");
    let xml = `<${w}:tbl>${tblPr}${grid}${rowXml}</${w}:tbl>`;
    if (p.caption) {
      const cap = this.captionXml(anchor, p.caption, o);
      xml = (p.caption_position ?? "above") === "above" ? cap + xml : xml + cap;
    }
    // 表格后面紧跟一个表格时 Word 会把两者合并：插在锚点之前时补一个空段落隔开
    const nodes = this.insertBlockXml(anchor, p.position, xml);
    const refs = this.refsOf(nodes);
    return {
      changedRefs: refs,
      summary: `已插入 ${rows.length} 行 × ${nCols} 列的表格（${{ grid: "网格线", three_line: "三线表", plain: "无边框", banded: "表头底色+隔行底纹" }[style]}）${p.caption ? "，含题注" : ""}${o.track ? "（记录为插入修订）" : ""}。单元格引用：${refs.slice(0, 12).join("、")}${refs.length > 12 ? "…" : ""}`,
      preview: [{ ref: refs[0] ?? "", before: "", after: rows.map((r) => r.join(" | ")).join("\n") }],
      structural: true,
    };
  }

  /** 表格结构编辑：删除行/列、插入列、合并单元格、底纹、对齐、边框、删除表格 */
  editTable(p: {
    ref: string; action: "delete_row" | "delete_column" | "insert_column" | "merge" | "shade" | "align" | "borders" | "delete_table" | "set_width" | "repeat_header";
    to_ref?: string; position?: "before" | "after"; cells?: string[]; color?: string;
    horizontal?: "left" | "center" | "right"; vertical?: "top" | "center" | "bottom"; style?: "grid" | "three_line" | "plain" | "outer";
    width_pct?: number; scope?: "cell" | "row" | "column" | "table";
  }, o: EditOptions): EditResult {
    const pe = this.resolve(p.ref);
    const w = this.w(pe.part);
    const tc = pe.el.closest(`${w}:tc`);
    const tr = pe.el.closest(`${w}:tr`);
    const tbl = pe.el.closest(`${w}:tbl`);
    if (!tc || !tr || !tbl) throw new DocError(`段落 ${pe.ref} 不在表格中。请传入表格某个单元格里的段落引用（doc_read part="表格1"）。`, "invalid");
    if (o.track && ["delete_column", "insert_column", "merge", "delete_table"].includes(p.action)) {
      throw new DocError("修订模式下暂不支持这类表格结构修改（Word 的列/合并修订记录不完整）。请先关闭修订模式，或改用逐行操作。", "unsupported");
    }
    const rows = tbl.childrenNamed(`${w}:tr`);
    const cellsOf = (row: XElement) => row.childrenNamed(`${w}:tc`);
    const span = (c: XElement) => Number(c.child(`${w}:tcPr`)?.child(`${w}:gridSpan`)?.getAttr(`${w}:val`) ?? 1);
    /** 单元格在网格中的起始列 */
    const gridCol = (row: XElement, c: XElement) => {
      const gb = Number(row.child(`${w}:trPr`)?.child(`${w}:gridBefore`)?.getAttr(`${w}:val`) ?? 0);
      let col = gb;
      for (const x of cellsOf(row)) { if (x === c) return col; col += span(x); }
      return col;
    };
    const cellAt = (row: XElement, col: number) => {
      let c0 = Number(row.child(`${w}:trPr`)?.child(`${w}:gridBefore`)?.getAttr(`${w}:val`) ?? 0);
      for (const x of cellsOf(row)) { const s = span(x); if (col >= c0 && col < c0 + s) return x; c0 += s; }
      return undefined;
    };
    const tcPrOf = (c: XElement) => this.ensureChild(c, "tcPr", [], w, true);
    const col = gridCol(tr, tc);
    let summary = "";
    const before = this.textOf(pe);
    const scopeCells = (): XElement[] => {
      switch (p.scope ?? "cell") {
        case "row": return cellsOf(tr);
        case "column": return rows.map((r) => cellAt(r, col)).filter((x): x is XElement => !!x);
        case "table": return rows.flatMap(cellsOf);
        default: return [tc];
      }
    };
    switch (p.action) {
      case "delete_row": {
        if (rows.length <= 1) throw new DocError("这是表格唯一的一行；要删除整个表格请用 action=delete_table。", "invalid");
        if (o.track) {
          const trPr = tr.child(`${w}:trPr`) ?? (() => { const [x] = parseFragment(`<${w}:trPr/>`) as XElement[]; tr.children.unshift(x); x.parent = tr; markDirty(tr); return x; })();
          const [del] = parseFragment(`<${w}:del ${this.revAttrs(w, o)}/>`);
          trPr.children.splice(orderedInsertIndex(trPr, "del", TRPR_ORDER, w), 0, del);
          del.parent = trPr; trPr.attrsDirty = true; markDirty(trPr);
          for (const r of [...tr.descendants()].filter((d) => d.local === "r" && d.prefix === w && !d.closest(`${w}:del`))) this.trackDeleteRun(r, pe, o);
          summary = "已删除该行（记录为删除修订）";
        } else {
          removeNode(tr);
          summary = "已删除该行";
        }
        break;
      }
      case "delete_column": {
        const grid = tbl.child(`${w}:tblGrid`);
        const gcols = grid?.childrenNamed(`${w}:gridCol`) ?? [];
        if (gcols.length <= 1) throw new DocError("表格只有一列；要删除整个表格请用 action=delete_table。", "invalid");
        for (const r of rows) {
          const c = cellAt(r, col);
          if (!c) continue;
          if (span(c) > 1) {
            const gs = c.child(`${w}:tcPr`)!.child(`${w}:gridSpan`)!;
            gs.setAttr(`${w}:val`, String(span(c) - 1));
          } else removeNode(c);
          if (!cellsOf(r).length) removeNode(r);
        }
        if (gcols[col]) removeNode(gcols[col]);
        summary = `已删除第 ${col + 1} 列`;
        break;
      }
      case "insert_column": {
        const grid = tbl.child(`${w}:tblGrid`);
        const gcols = grid?.childrenNamed(`${w}:gridCol`) ?? [];
        const at = (p.position ?? "after") === "after" ? col + span(tc) - 1 : col;
        const width = Number(gcols[at]?.getAttr(`${w}:w`) ?? 1200);
        rows.forEach((r, ri) => {
          const ref = cellAt(r, at);
          if (!ref) return;
          const [clone] = parseFragment(serializeNode(ref, this.src(pe.part))) as XElement[];
          for (const d of [...clone.descendants()]) if (d.local === "vMerge" || d.local === "gridSpan") removeNode(d);
          const paras = clone.childrenNamed(`${w}:p`);
          paras.slice(1).forEach((x) => removeNode(x));
          const firstP = paras[0];
          const tplRun = [...firstP.descendants()].find((d) => d.local === "r" && d.child(`${w}:t`));
          const rPr = tplRun?.child(`${w}:rPr`);
          for (const c of [...firstP.children]) if (!(c.type === "el" && c.local === "pPr")) removeNode(c);
          if (this.supportsW14(pe.part)) { firstP.setAttr("w14:paraId", this.newParaId()); firstP.setAttr("w14:textId", "77777777"); }
          const text = p.cells?.[ri] ?? "";
          if (text) appendChild(firstP, parseFragment(`<${w}:r>${rPr ? serializeNode(rPr, "") : ""}${runContent(w, text)}</${w}:r>`));
          const tcW = clone.child(`${w}:tcPr`)?.child(`${w}:tcW`);
          if (tcW) { tcW.setAttr(`${w}:w`, String(width)); tcW.setAttr(`${w}:type`, "dxa"); }
          if ((p.position ?? "after") === "after" && span(ref) === 1) insertAfter(ref, [clone]);
          else insertBefore(ref, [clone]);
        });
        if (grid && gcols[at]) {
          const [g] = parseFragment(`<${w}:gridCol ${w}:w="${width}"/>`);
          if ((p.position ?? "after") === "after") insertAfter(gcols[at], [g]); else insertBefore(gcols[at], [g]);
        }
        const tblW = tbl.child(`${w}:tblPr`)?.child(`${w}:tblW`);
        if (tblW?.getAttr(`${w}:type`) === "dxa") tblW.setAttr(`${w}:w`, String(Number(tblW.getAttr(`${w}:w`)) + width));
        summary = `已在第 ${col + 1} 列${(p.position ?? "after") === "after" ? "之后" : "之前"}插入一列`;
        break;
      }
      case "merge": {
        if (!p.to_ref) throw new DocError("合并单元格需要 to_ref：矩形区域另一个角的单元格段落引用。", "invalid");
        const other = this.resolve(p.to_ref);
        const tc2 = other.el.closest(`${w}:tc`), tr2 = other.el.closest(`${w}:tr`);
        if (!tc2 || !tr2 || other.el.closest(`${w}:tbl`) !== tbl) throw new DocError("to_ref 必须是同一个表格中的单元格。", "invalid");
        const r1 = Math.min(rows.indexOf(tr), rows.indexOf(tr2)), r2 = Math.max(rows.indexOf(tr), rows.indexOf(tr2));
        const c1 = Math.min(col, gridCol(tr2, tc2)), c2 = Math.max(col + span(tc) - 1, gridCol(tr2, tc2) + span(tc2) - 1);
        for (let ri = r1; ri <= r2; ri++) {
          const row = rows[ri];
          const first = cellAt(row, c1);
          if (!first || gridCol(row, first) !== c1) throw new DocError("合并区域的边界穿过了已合并的单元格，请调整范围。", "invalid");
          // 横向：把 c1..c2 合成一个 gridSpan 单元格，文字并入第一个
          let k = gridCol(row, first) + span(first);
          while (k <= c2) {
            const next = cellAt(row, k);
            if (!next) break;
            const ns = span(next);
            for (const para of next.childrenNamed(`${w}:p`)) if (this.textOfEl(para, pe.part).trim()) { removeNode(para); appendChild(first, [para]); }
            removeNode(next);
            k += ns;
          }
          const pr = tcPrOf(first);
          this.setProp(pr, "gridSpan", c2 - c1 + 1 > 1 ? { val: String(c2 - c1 + 1) } : null, TCPR_ORDER, w);
          const tcW = pr.child(`${w}:tcW`);
          const gcols = tbl.child(`${w}:tblGrid`)?.childrenNamed(`${w}:gridCol`) ?? [];
          const sumW = gcols.slice(c1, c2 + 1).reduce((a, g) => a + Number(g.getAttr(`${w}:w`) ?? 0), 0);
          if (tcW && sumW) { tcW.setAttr(`${w}:w`, String(sumW)); tcW.setAttr(`${w}:type`, "dxa"); }
          // 纵向
          if (r2 > r1) this.setProp(pr, "vMerge", ri === r1 ? { val: "restart" } : {}, TCPR_ORDER, w);
          if (ri > r1) {
            const top = cellAt(rows[r1], c1)!;
            for (const para of first.childrenNamed(`${w}:p`)) if (this.textOfEl(para, pe.part).trim()) { removeNode(para); appendChild(top, [para]); }
            if (!first.childrenNamed(`${w}:p`).length) appendChild(first, parseFragment(this.paraXml(pe.part, "", "", o)));
          }
          // 合并后第一个单元格里去掉多余的空段落
          const ps = first.childrenNamed(`${w}:p`);
          if (ps.length > 1) for (const x of ps) if (!this.textOfEl(x, pe.part).trim() && first.childrenNamed(`${w}:p`).length > 1) removeNode(x);
        }
        summary = `已合并第 ${r1 + 1}–${r2 + 1} 行、第 ${c1 + 1}–${c2 + 1} 列的单元格`;
        break;
      }
      case "shade": {
        const color = (p.color ?? "").replace(/^#/, "").toUpperCase();
        if (color && !/^[0-9A-F]{6}$/.test(color) && color !== "NONE") throw new DocError("color 需为 6 位十六进制颜色（如 D9E2F3）或 none", "invalid");
        for (const c of scopeCells()) this.setProp(tcPrOf(c), "shd", color && color !== "NONE" ? { val: "clear", color: "auto", fill: color } : null, TCPR_ORDER, w);
        summary = `已设置${{ cell: "单元格", row: "整行", column: "整列", table: "整个表格" }[p.scope ?? "cell"]}底纹 ${color || "（清除）"}`;
        break;
      }
      case "align": {
        for (const c of scopeCells()) {
          if (p.vertical) this.setProp(tcPrOf(c), "vAlign", { val: p.vertical }, TCPR_ORDER, w);
          if (p.horizontal) for (const para of c.childrenNamed(`${w}:p`)) this.setProp(this.ensurePPr(para, w), "jc", { val: p.horizontal === "left" ? "left" : p.horizontal }, PPR_ORDER, w);
        }
        summary = `已设置对齐（${[p.horizontal && `水平 ${p.horizontal}`, p.vertical && `垂直 ${p.vertical}`].filter(Boolean).join("，")}）`;
        break;
      }
      case "borders": {
        const tblPr = this.ensureChild(tbl, "tblPr", [], w, true);
        const B = (sz: number) => ({ val: "single", sz: String(sz), space: "0", color: "000000" });
        const old = tblPr.child(`${w}:tblBorders`);
        if (old) removeNode(old);
        const st = p.style ?? "grid";
        const spec: Record<string, Record<string, string> | null> = st === "grid"
          ? { top: B(4), left: B(4), bottom: B(4), right: B(4), insideH: B(4), insideV: B(4) }
          : st === "three_line" ? { top: B(12), left: { val: "nil" }, bottom: B(12), right: { val: "nil" }, insideH: { val: "nil" }, insideV: { val: "nil" } }
          : st === "outer" ? { top: B(8), left: B(8), bottom: B(8), right: B(8), insideH: { val: "nil" }, insideV: { val: "nil" } }
          : { top: { val: "nil" }, left: { val: "nil" }, bottom: { val: "nil" }, right: { val: "nil" }, insideH: { val: "nil" }, insideV: { val: "nil" } };
        const inner = Object.entries(spec).map(([k, a]) => `<${w}:${k}${Object.entries(a!).map(([x, v]) => ` ${w}:${x}="${v}"`).join("")}/>`).join("");
        const [bEl] = parseFragment(`<${w}:tblBorders>${inner}</${w}:tblBorders>`);
        tblPr.children.splice(orderedInsertIndex(tblPr, "tblBorders", TBLPR_ORDER, w), 0, bEl);
        bEl.parent = tblPr; markDirty(tblPr);
        // 单元格自带的边框会覆盖表格边框：清除
        for (const c of rows.flatMap(cellsOf)) { const b = c.child(`${w}:tcPr`)?.child(`${w}:tcBorders`); if (b) removeNode(b); }
        if (st === "three_line") {
          for (const c of cellsOf(rows[0])) this.setProp(tcPrOf(c), "tcBorders", null, TCPR_ORDER, w);
          for (const c of cellsOf(rows[0])) {
            const [bb] = parseFragment(`<${w}:tcBorders><${w}:bottom ${w}:val="single" ${w}:sz="6" ${w}:space="0" ${w}:color="000000"/></${w}:tcBorders>`);
            const pr = tcPrOf(c);
            pr.children.splice(orderedInsertIndex(pr, "tcBorders", TCPR_ORDER, w), 0, bb); bb.parent = pr; markDirty(pr);
          }
        }
        summary = `已把表格边框设为${{ grid: "网格线", three_line: "三线表", plain: "无边框", outer: "仅外框" }[st]}`;
        break;
      }
      case "set_width": {
        const pct = Math.min(100, Math.max(10, p.width_pct ?? 100));
        const total = Math.round(this.textWidthTwips(pe) * pct / 100);
        const tblPr = this.ensureChild(tbl, "tblPr", [], w, true);
        this.setProp(tblPr, "tblW", { w: String(total), type: "dxa" }, TBLPR_ORDER, w);
        const gcols = tbl.child(`${w}:tblGrid`)?.childrenNamed(`${w}:gridCol`) ?? [];
        const old = gcols.reduce((a, g) => a + Number(g.getAttr(`${w}:w`) ?? 0), 0) || 1;
        const ratio = total / old;
        gcols.forEach((g) => g.setAttr(`${w}:w`, String(Math.round(Number(g.getAttr(`${w}:w`) ?? 0) * ratio))));
        for (const c of rows.flatMap(cellsOf)) {
          const tcW = c.child(`${w}:tcPr`)?.child(`${w}:tcW`);
          if (tcW?.getAttr(`${w}:type`) === "dxa") tcW.setAttr(`${w}:w`, String(Math.round(Number(tcW.getAttr(`${w}:w`)) * ratio)));
        }
        summary = `已把表格宽度设为版心的 ${pct}%`;
        break;
      }
      case "repeat_header": {
        const first = rows[0];
        const trPr = first.child(`${w}:trPr`) ?? (() => { const [x] = parseFragment(`<${w}:trPr/>`) as XElement[]; const ex = first.child(`${w}:tblPrEx`); if (ex) insertAfter(ex, [x]); else { first.children.unshift(x); x.parent = first; markDirty(first); } return x; })();
        this.setProp(trPr, "tblHeader", {}, TRPR_ORDER, w);
        summary = "已设置表头行在每页重复";
        break;
      }
      case "delete_table": {
        if (o.track) throw new DocError("修订模式下请逐行删除。", "unsupported");
        const parent = tbl.parent as XElement;
        removeNode(tbl);
        if (parent.local === "tc" && !parent.childrenNamed(`${w}:p`).length) appendChild(parent, parseFragment(this.paraXml(pe.part, "", "", o)));
        summary = "已删除整个表格";
        break;
      }
    }
    const structural = ["delete_row", "delete_column", "insert_column", "merge", "delete_table"].includes(p.action);
    if (structural) this._structureVersion++;
    this.reindex();
    return { changedRefs: structural ? [] : [pe.ref], summary: `${summary}${structural ? "。表格结构已变化，请重新读取以获得最新引用。" : "。"}`, preview: [{ ref: pe.ref, before: `（表格）${clip(before, 60)}`, after: summary }], structural };
  }

  private textOfEl(pEl: XElement, part: string): string {
    const pe = this.paras.find((x) => x.el === pEl);
    if (pe) return this.textOf(pe);
    void part;
    return [...pEl.descendants()].filter((d) => d.local === "t").map((d) => innerText(d)).join("");
  }

  // ------------------------------- 图片 / 图形 / 图表 -------------------------------

  /** 写入图片部件并登记关系，返回关系 ID */
  private addImagePart(data: Buffer, owner: string): { rId: string; info: NonNullable<ReturnType<typeof imageInfo>> } {
    const info = imageInfo(data);
    if (!info) throw new DocError("无法识别的图片格式（支持 PNG / JPEG / GIF / BMP）。", "invalid");
    const ext = info.ext === "jpeg" ? "jpeg" : info.ext;
    const name = this.uniquePartName(this.mainPart.replace(/[^/]+$/, "media").replace(/^\/?/, ""), "image", ext);
    this.addPart(name, data);
    this.ensureDefaultCT(ext, info.mime);
    return { rId: this.addRel(REL_EXTRA.image, this.relTarget(name), { owner }), info };
  }

  /** 新式对象外面包一层图片后备（浏览器预览等不认识形状/图表的阅读器显示图片） */
  private withFallback(w: string, owner: string, drawing: string, fallbackPng: Buffer | undefined, cx: number, cy: number, descr?: string): string {
    if (!fallbackPng) return drawing;
    try {
      const { rId } = this.addImagePart(fallbackPng, owner);
      const id = this.nextDocPrId();
      return alternateContentXml(drawing, pictureXml({ w, rId, cx, cy, id, name: `后备图片 ${id}`, descr }));
    } catch {
      return drawing;
    }
  }

  insertImage(p: { anchor: string; position: "before" | "after"; data: Buffer; width_pt?: number; height_pt?: number; align?: "left" | "center" | "right"; caption?: string; alt?: string; name?: string }, o: EditOptions): EditResult {
    const anchor = this.resolve(p.anchor);
    const w = this.w(anchor.part);
    const { rId, info } = this.addImagePart(p.data, anchor.part);
    const maxW = this.textWidthTwips(anchor) / 20;
    let wpt = p.width_pt ?? info.width * 0.75;
    let hpt = p.height_pt ?? (wpt * info.height) / Math.max(1, info.width);
    if (p.width_pt && !p.height_pt) hpt = (p.width_pt * info.height) / Math.max(1, info.width);
    if (!p.width_pt && p.height_pt) wpt = (p.height_pt * info.width) / Math.max(1, info.height);
    if (wpt > maxW) { hpt = (hpt * maxW) / wpt; wpt = maxW; }
    const id = this.nextDocPrId();
    const drawing = pictureXml({ w, rId, cx: emu(wpt), cy: emu(hpt), id, name: p.name ?? `图片 ${id}`, descr: p.alt ?? p.caption });
    let xml = this.objectParaXml(anchor, `<${w}:r>${drawing}</${w}:r>`, p.align ?? "center", o, !!p.caption);
    if (p.caption) xml += this.captionXml(anchor, p.caption, o);
    const nodes = this.insertBlockXml(anchor, p.position, xml);
    const refs = this.refsOf(nodes);
    return {
      changedRefs: refs,
      summary: `已插入图片（${Math.round(wpt)}×${Math.round(hpt)} pt，原图 ${info.width}×${info.height} 像素）${p.caption ? "及题注" : ""}，段落引用：${refs.join("、")}。`,
      preview: [{ ref: refs[0], before: "", after: `⟨图片⟩${p.caption ? `\n${p.caption}` : ""}` }],
      structural: true,
    };
  }

  insertShapes(p: { anchor: string; position: "before" | "after"; width_pt: number; height_pt: number; shapes: ShapeSpec[]; caption?: string; alt?: string; align?: "left" | "center" | "right"; fallbackPng?: Buffer }, o: EditOptions): EditResult {
    const anchor = this.resolve(p.anchor);
    const w = this.w(anchor.part);
    if (!p.shapes.length) throw new DocError("shapes 不能为空", "invalid");
    const maxW = this.textWidthTwips(anchor) / 20;
    let shapes = p.shapes;
    let W = p.width_pt, H = p.height_pt;
    if (W > maxW) {
      const k = maxW / W;
      shapes = shapes.map((s) => ({ ...s, x: s.x * k, y: s.y * k, w: s.w !== undefined ? s.w * k : undefined, h: s.h !== undefined ? s.h * k : undefined, x2: s.x2 !== undefined ? s.x2 * k : undefined, y2: s.y2 !== undefined ? s.y2 * k : undefined, font_size: s.font_size ? Math.max(6, s.font_size * Math.max(k, 0.8)) : s.font_size }));
      W = maxW; H = H * k;
    }
    const font = /<[^>]*rFonts[^>]*eastAsia="([^"]+)"/.exec(this.dominantRPrXml(anchor))?.[1];
    const id = this.nextDocPrId();
    const alt = p.alt ?? shapes.map((s) => s.text).filter(Boolean).join(" / ");
    const drawing = this.withFallback(w, anchor.part, shapesGroupXml({ w, id, name: `图形 ${id}`, descr: alt, width: W, height: H, shapes, font }), p.fallbackPng, emu(W), emu(H), alt);
    let xml = this.objectParaXml(anchor, `<${w}:r>${drawing}</${w}:r>`, p.align ?? "center", o, !!p.caption);
    if (p.caption) xml += this.captionXml(anchor, p.caption, o);
    const nodes = this.insertBlockXml(anchor, p.position, xml);
    const refs = this.refsOf(nodes);
    return {
      changedRefs: refs,
      summary: `已插入由 ${shapes.length} 个形状组成的矢量图（${Math.round(W)}×${Math.round(H)} pt，可在 Word 中逐个选中修改）${p.caption ? "及题注" : ""}。`,
      preview: [{ ref: refs[0], before: "", after: `⟨图形⟩ ${clip(alt, 120)}${p.caption ? `\n${p.caption}` : ""}` }],
      structural: true,
    };
  }

  insertChart(p: { anchor: string; position: "before" | "after"; chart: ChartSpec; width_pt?: number; height_pt?: number; caption?: string; fallbackPng?: Buffer }, o: EditOptions): EditResult {
    const anchor = this.resolve(p.anchor);
    const w = this.w(anchor.part);
    const c = p.chart;
    if (!c.categories.length || !c.series.length) throw new DocError("图表需要至少一个分类和一个数据系列", "invalid");
    for (const s of c.series) if (s.values.length !== c.categories.length) throw new DocError(`系列"${s.name}"有 ${s.values.length} 个值，但分类有 ${c.categories.length} 个，需一一对应。`, "invalid");
    if (c.series.length > 20 || c.categories.length > 200) throw new DocError("系列最多 20 个、分类最多 200 个", "invalid");
    const dir = this.mainPart.replace(/[^/]+$/, "");
    const chartName = this.uniquePartName(`${dir}charts`, "chart", "xml");
    const n = /chart(\d+)\.xml$/.exec(chartName)![1];
    const wbName = this.uniquePartName(`${dir}embeddings`, `Microsoft_Excel_Worksheet`, "xlsx");
    this.addPart(chartName, chartXml(c));
    this.addPart(wbName, chartWorkbook(c));
    this.ensureOverrideCT(chartName, CT.chart);
    this.ensureDefaultCT("xlsx", CT.xlsx);
    this.addPart(`${dir}charts/_rels/chart${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${REL_EXTRA.package}" Target="../embeddings/${wbName.split("/").pop()}"/></Relationships>`);
    const rId = this.addRel(REL_EXTRA.chart, this.relTarget(chartName), { owner: anchor.part });
    const maxW = this.textWidthTwips(anchor) / 20;
    const wpt = Math.min(maxW, p.width_pt ?? Math.min(maxW, 430));
    const hpt = p.height_pt ?? wpt * 0.6;
    const id = this.nextDocPrId();
    const alt = `${c.title ?? "图表"}：${c.series.map((s) => `${s.name}（${s.values.join("、")}）`).join("；")}`;
    const chartDrawing = this.withFallback(w, anchor.part, chartInlineXml({ w, rId, cx: emu(wpt), cy: emu(hpt), id, name: `图表 ${id}`, descr: alt }), p.fallbackPng, emu(wpt), emu(hpt), alt);
    let xml = this.objectParaXml(anchor, `<${w}:r>${chartDrawing}</${w}:r>`, "center", o, !!p.caption);
    if (p.caption) xml += this.captionXml(anchor, p.caption, o);
    const nodes = this.insertBlockXml(anchor, p.position, xml);
    const refs = this.refsOf(nodes);
    return {
      changedRefs: refs,
      summary: `已插入原生 Word 图表（${c.type}，${c.series.length} 个系列 × ${c.categories.length} 个分类），数据内嵌为 Excel 工作簿，可在 Word 中右键"编辑数据"。`,
      preview: [{ ref: refs[0], before: "", after: `⟨图表⟩ ${clip(alt, 160)}${p.caption ? `\n${p.caption}` : ""}` }],
      structural: true,
    };
  }

  // ------------------------------- 公式 -------------------------------

  insertEquation(p: { anchor?: string; position?: "before" | "after"; ref?: string; after_text?: string; latex: string; display?: boolean; number?: string }, o: EditOptions): EditResult {
    const display = p.display ?? !p.ref;
    if (display) {
      if (!p.anchor) throw new DocError("行间公式需要 anchor（插在哪个段落前后）", "invalid");
      const anchor = this.resolve(p.anchor);
      const w = this.w(anchor.part);
      let xml = equationParagraphXml({ w, latex: p.latex, display: true, number: p.number, textWidthTwips: this.textWidthTwips(anchor), rPr: this.baseFontRPr(anchor) });
      // 补上段落 ID 与修订标记
      xml = xml.replace(`<${w}:p>`, `<${w}:p${this.supportsW14(anchor.part) ? ` w14:paraId="${this.newParaId()}" w14:textId="77777777"` : ""}>`);
      if (o.track) {
        xml = xml.replace(`</${w}:pPr>`, `<${w}:rPr>${this.insMark(w, o)}</${w}:rPr></${w}:pPr>`);
        xml = xml.replace(new RegExp(`<${w}:r>[\\s\\S]*?</${w}:r>`, "g"), (r) => this.trackRun(w, r, o));
      }
      const nodes = this.insertBlockXml(anchor, p.position ?? "after", xml);
      const refs = this.refsOf(nodes);
      return { changedRefs: refs, summary: `已插入行间公式${p.number ? `（编号 ${p.number}，与公式在同一行右端）` : ""}，可在 Word 中直接编辑。`, preview: [{ ref: refs[0], before: "", after: `⟨公式⟩ ${p.latex}${p.number ? `   ${p.number}` : ""}` }], structural: true };
    }
    // 行内公式：插在段落中某段文字之后（或段尾）
    const pe = this.resolve(p.ref!);
    this.isEditable(pe);
    const w = this.w(pe.part);
    const full = this.textOf(pe);
    const pos = p.after_text ? this.findRange(pe, p.after_text).end : full.length;
    if (pos > 0 && pos < full.length) this.splitAt(pe, pos);
    const omml = parseFragment(equationParagraphXml({ w, latex: p.latex, display: false, textWidthTwips: 0 }));
    const { segs } = this.segments(pe);
    const before = [...segs].reverse().find((s) => s.run && s.end <= pos && s.end > s.start);
    if (before?.run && pos > 0) insertAfter(this.anchorOf(before.run, pe.el), omml);
    else {
      const pPr = pe.el.child(`${w}:pPr`);
      if (pPr) insertAfter(pPr, omml);
      else { pe.el.children.unshift(...omml); omml.forEach((x) => (x.parent = pe.el)); markDirty(pe.el); }
    }
    return { changedRefs: [pe.ref], summary: `已在段落 ${pe.ref} 中插入行内公式。`, preview: [{ ref: pe.ref, before: full, after: this.textOf(pe) }], structural: false };
  }

  // ------------------------------- 分隔符 / 页面设置 -------------------------------

  insertBreak(p: { anchor: string; position: "before" | "after"; kind: "page" | "column" | "section_next_page" | "section_continuous" | "section_odd_page" | "section_even_page" }, o: EditOptions): EditResult {
    const anchor = this.resolve(p.anchor);
    if (anchor.part !== this.mainPart || anchor.container !== "body") throw new DocError("分隔符只能插在正文段落（不在表格、页眉页脚中）前后。", "invalid");
    const w = this.w(anchor.part);
    if (p.kind === "page" || p.kind === "column") {
      const xml = this.paraXml(anchor.part, "", this.trackRun(w, `<${w}:r><${w}:br ${w}:type="${p.kind}"/></${w}:r>`, o), o);
      const nodes = this.insertBlockXml(anchor, p.position, xml);
      return { changedRefs: this.refsOf(nodes), summary: `已插入${p.kind === "page" ? "分页符" : "分栏符"}。`, preview: [{ ref: "", before: "", after: p.kind === "page" ? "⟨分页符⟩" : "⟨分栏符⟩" }], structural: true };
    }
    // 分节符：新段落带上当前节的属性（结束前半部分），原来的节属性改为新的起始方式
    const cur = this.sectPrFor(anchor);
    if (!cur) throw new DocError("文档缺少分节属性（sectPr），无法插入分节符。", "invalid");
    const clone = serializeNode(cur, this.src(anchor.part)).replace(new RegExp(`<${w}:sectPrChange[\\s\\S]*?</${w}:sectPrChange>`), "");
    const xml = this.paraXml(anchor.part, clone, "", o);
    const nodes = this.insertBlockXml(anchor, p.position, xml);
    const type = { section_next_page: "nextPage", section_continuous: "continuous", section_odd_page: "oddPage", section_even_page: "evenPage" }[p.kind];
    this.setProp(cur, "type", { val: type }, SECTPR_ORDER, w);
    return { changedRefs: this.refsOf(nodes), summary: `已插入分节符（${{ nextPage: "下一页", continuous: "连续", oddPage: "奇数页", evenPage: "偶数页" }[type]}）。之后可以用 doc_page_setup 单独设置新节的纸张方向、边距、分栏。`, preview: [{ ref: "", before: "", after: "⟨分节符⟩" }], structural: true };
  }

  pageSetup(p: { ref?: string; orientation?: "portrait" | "landscape"; paper?: "A4" | "A3" | "A5" | "B5" | "Letter" | "Legal"; margins_pt?: { top?: number; bottom?: number; left?: number; right?: number; header?: number; footer?: number }; columns?: number; column_gap_pt?: number; line_numbers?: boolean; v_align?: "top" | "center" | "bottom" }, _o: EditOptions): EditResult {
    const w = this.w(this.mainPart);
    const targets = p.ref ? [this.sectPrFor(this.resolve(p.ref))].filter((x): x is XElement => !!x) : this.allSectPrs();
    if (!targets.length) throw new DocError("文档缺少分节属性（sectPr）。", "invalid");
    const PAPER: Record<string, [number, number]> = { A4: [11906, 16838], A3: [16838, 23811], A5: [8391, 11906], B5: [10319, 14571], Letter: [12240, 15840], Legal: [12240, 20160] };
    const changes: string[] = [];
    for (const sp of targets) {
      const pgSz = this.ensureChild(sp, "pgSz", SECTPR_ORDER, w);
      let W = Number(pgSz.getAttr(`${w}:w`) ?? 11906), H = Number(pgSz.getAttr(`${w}:h`) ?? 16838);
      if (p.paper) { [W, H] = PAPER[p.paper]; if ((pgSz.getAttr(`${w}:orient`) === "landscape" && p.orientation !== "portrait") || p.orientation === "landscape") [W, H] = [H, W]; }
      if (p.orientation) {
        const land = p.orientation === "landscape";
        if (land !== W > H) [W, H] = [H, W];
        if (land) pgSz.setAttr(`${w}:orient`, "landscape"); else pgSz.removeAttr(`${w}:orient`);
        // 横向时左右与上下边距互换，保持版心比例
        const mar = sp.child(`${w}:pgMar`);
        if (mar && !p.margins_pt) {
          const [t, b, l, r] = ["top", "bottom", "left", "right"].map((k) => mar.getAttr(`${w}:${k}`));
          if (t && b && l && r && land !== (Number(l) > Number(t))) { mar.setAttr(`${w}:top`, l); mar.setAttr(`${w}:bottom`, r); mar.setAttr(`${w}:left`, t); mar.setAttr(`${w}:right`, b); }
        }
      }
      pgSz.setAttr(`${w}:w`, String(W));
      pgSz.setAttr(`${w}:h`, String(H));
      if (p.margins_pt) {
        const mar = this.ensureChild(sp, "pgMar", SECTPR_ORDER, w);
        for (const [k, v] of Object.entries(p.margins_pt)) if (typeof v === "number") mar.setAttr(`${w}:${k}`, String(twip(v)));
        for (const k of ["top", "bottom", "left", "right", "header", "footer", "gutter"]) if (!mar.getAttr(`${w}:${k}`)) mar.setAttr(`${w}:${k}`, k === "gutter" ? "0" : k === "header" || k === "footer" ? "851" : "1440");
      }
      if (p.columns) {
        const cols = this.ensureChild(sp, "cols", SECTPR_ORDER, w);
        for (const c of [...cols.children]) removeNode(c);
        cols.setAttr(`${w}:num`, String(p.columns));
        cols.setAttr(`${w}:space`, String(twip(p.column_gap_pt ?? 21)));
        cols.removeAttr(`${w}:equalWidth`);
      }
      if (p.line_numbers !== undefined) this.setProp(sp, "lnNumType", p.line_numbers ? { countBy: "1", restart: "newPage" } : null, SECTPR_ORDER, w);
      if (p.v_align) this.setProp(sp, "vAlign", { val: p.v_align === "center" ? "center" : p.v_align }, SECTPR_ORDER, w);
      markDirty(sp);
    }
    if (p.paper) changes.push(`纸张 ${p.paper}`);
    if (p.orientation) changes.push(p.orientation === "landscape" ? "横向" : "纵向");
    if (p.margins_pt) changes.push(`边距 ${Object.entries(p.margins_pt).map(([k, v]) => `${k}=${v}pt`).join(" ")}`);
    if (p.columns) changes.push(`${p.columns} 栏`);
    if (p.line_numbers !== undefined) changes.push(p.line_numbers ? "显示行号" : "取消行号");
    if (p.v_align) changes.push(`页面垂直对齐 ${p.v_align}`);
    return { changedRefs: [], summary: `已修改${p.ref ? "所在节" : `全部 ${targets.length} 节`}的页面设置：${changes.join("，") || "（无变化）"}。`, preview: [{ ref: "", before: "页面设置", after: changes.join("，") }], structural: false };
  }

  // ------------------------------- 页眉页脚 -------------------------------

  setHeaderFooter(p: { kind: "header" | "footer"; text: string; align?: "left" | "center" | "right"; ref?: string; first_page?: boolean }, o: EditOptions): EditResult {
    const w = this.w(this.mainPart);
    const targets = p.ref ? [this.sectPrFor(this.resolve(p.ref))].filter((x): x is XElement => !!x) : this.allSectPrs();
    if (!targets.length) throw new DocError("文档缺少分节属性（sectPr）。", "invalid");
    const dir = this.mainPart.replace(/[^/]+$/, "");
    const name = this.uniquePartName(dir.replace(/\/$/, ""), p.kind, "xml");
    const styleName = p.kind === "header" ? "header" : "footer";
    const st = this.styles.resolve(styleName, "paragraph") ?? this.styles.resolve(p.kind === "header" ? "页眉" : "页脚", "paragraph");
    // {PAGE} {NUMPAGES} {SECTIONPAGES} → 域
    const fld = (instr: string) => `<w:fldSimple w:instr=" ${instr} "><w:r><w:t>1</w:t></w:r></w:fldSimple>`;
    const paras = p.text.split("\n").map((line) => {
      const pieces = line.split(/(\{PAGE\}|\{NUMPAGES\}|\{SECTIONPAGES\}|\{DATE\})/);
      const content = pieces.filter(Boolean).map((x) => x === "{PAGE}" ? fld("PAGE") : x === "{NUMPAGES}" ? fld("NUMPAGES") : x === "{SECTIONPAGES}" ? fld("SECTIONPAGES") : x === "{DATE}" ? fld('DATE \\@ "yyyy-MM-dd"') : `<w:r>${runContent("w", x)}</w:r>`).join("");
      return `<w:p><w:pPr>${st ? `<w:pStyle w:val="${st.id}"/>` : ""}<w:jc w:val="${p.align ?? "center"}"/></w:pPr>${content}</w:p>`;
    }).join("");
    const root = p.kind === "header" ? "hdr" : "ftr";
    this.addPart(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:${root} xmlns:w="${NS.w}" xmlns:r="${NS.r}">${paras || "<w:p/>"}</w:${root}>`);
    this.ensureOverrideCT(name, p.kind === "header" ? CT.header : CT.footer);
    const rId = this.addRel(p.kind === "header" ? REL_TYPES.header : REL_TYPES.footer, this.relTarget(name));
    const refName = p.kind === "header" ? "headerReference" : "footerReference";
    const type = p.first_page ? "first" : "default";
    for (const sp of targets) {
      for (const old of sp.childrenNamed(`${w}:${refName}`)) if ((old.getAttr(`${w}:type`) ?? "default") === type) removeNode(old);
      const [node] = parseFragment(`<${w}:${refName} ${w}:type="${type}" r:id="${rId}" xmlns:r="${NS.r}"/>`);
      // 页眉引用在前、页脚引用在后
      const kids = sp.children;
      let idx = 0;
      for (let k = 0; k < kids.length; k++) {
        const c = kids[k];
        if (c.type !== "el") continue;
        if (c.local === "headerReference" || (p.kind === "footer" && c.local === "footerReference")) idx = k + 1;
      }
      kids.splice(idx, 0, node);
      node.parent = sp;
      markDirty(sp);
      if (p.first_page) this.setProp(sp, "titlePg", {}, SECTPR_ORDER, w);
    }
    // 新部件加入可编辑范围（排在最后，已有段落的序号引用不受影响）
    this.contentParts.push(name);
    this.reindex();
    void o;
    return { changedRefs: [], summary: `已为${p.ref ? "所在节" : `全部 ${targets.length} 节`}设置${p.first_page ? "首页" : ""}${p.kind === "header" ? "页眉" : "页脚"}：${p.text.replace(/\n/g, " ⏎ ")}（{PAGE} 等已转换为自动更新的页码域）。`, preview: [{ ref: "", before: "", after: `${p.kind === "header" ? "页眉" : "页脚"}：${p.text}` }], structural: false };
  }

  // ------------------------------- 列表 -------------------------------

  private numberingPartName(create: boolean): string | null {
    const rels = parseRels(this.optText(relsPathFor(this.mainPart)) ?? null);
    const r = rels.find((x) => x.type === REL_TYPES.numbering);
    if (r) {
      const name = resolveTarget(this.mainPart, r.target);
      if (this.pkg.has(name)) return name;
    }
    if (!create) return null;
    const name = this.mainPart.replace(/[^/]+$/, "numbering.xml");
    this.addPart(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:numbering xmlns:w="${NS.w}"></w:numbering>`);
    this.ensureOverrideCT(name, CT.numbering);
    this.addRel(REL_TYPES.numbering, this.relTarget(name));
    return name;
  }

  setList(p: { refs: string[]; kind: "bullet" | "number" | "chinese" | "outline" | "none"; level?: number; restart?: boolean }, o: EditOptions): EditResult {
    const entries = p.refs.map((r) => this.resolve(r));
    const level = Math.max(0, Math.min(8, (p.level ?? 1) - 1));
    let numId: string | null = null;
    if (p.kind !== "none") {
      const part = this.numberingPartName(true)!;
      const doc = this.part(part);
      const nw = this.w(part);
      const abs = doc.root.elements().filter((e) => e.local === "abstractNum");
      const nums = doc.root.elements().filter((e) => e.local === "num");
      // 复用本工具之前创建的同类定义（连续编号）
      const tag = `docagent-${p.kind}`;
      let absId: string | undefined = abs.find((a) => a.child(`${nw}:name`)?.getAttr(`${nw}:val`) === tag)?.getAttr(`${nw}:abstractNumId`);
      if (!absId) {
        const id = Math.max(-1, ...abs.map((a) => Number(a.getAttr(`${nw}:abstractNumId`)))) + 1;
        absId = String(id);
        const xml = abstractNumXml(nw, id, p.kind).replace(`<${nw}:multiLevelType`, `<${nw}:name ${nw}:val="${tag}"/><${nw}:multiLevelType`)
          // name 必须在 multiLevelType 之后：按 schema 调整顺序
          .replace(new RegExp(`(<${nw}:name [^>]*/>)(<${nw}:multiLevelType [^>]*/>)`), "$2$1");
        const node = parseFragment(xml);
        const lastAbs = abs[abs.length - 1];
        if (lastAbs) insertAfter(lastAbs, node);
        else if (nums[0]) insertBefore(nums[0], node);
        else appendChild(doc.root, node);
      }
      const existingNum = nums.find((n) => n.child(`${nw}:abstractNumId`)?.getAttr(`${nw}:val`) === absId && !n.child(`${nw}:lvlOverride`));
      if (existingNum && !p.restart) numId = existingNum.getAttr(`${nw}:numId`)!;
      else {
        const id = Math.max(0, ...nums.map((n) => Number(n.getAttr(`${nw}:numId`)))) + 1;
        numId = String(id);
        const override = p.restart ? `<${nw}:lvlOverride ${nw}:ilvl="${level}"><${nw}:startOverride ${nw}:val="1"/></${nw}:lvlOverride>` : "";
        appendChild(doc.root, parseFragment(`<${nw}:num ${nw}:numId="${id}"><${nw}:abstractNumId ${nw}:val="${absId}"/>${override}</${nw}:num>`));
      }
    }
    const preview: EditResult["preview"] = [];
    for (const pe of entries) {
      this.isEditable(pe);
      const w = this.w(pe.part);
      const pPr = this.ensurePPr(pe.el, w);
      if (o.track && !pPr.child(`${w}:pPrChange`)) {
        const old = pPr.children.filter((c) => !(c.type === "el" && ["rPr", "sectPr", "pPrChange"].includes(c.local))).map((c) => serializeNode(c, this.src(pe.part))).join("");
        const [chg] = parseFragment(`<${w}:pPrChange ${this.revAttrs(w, o)}><${w}:pPr>${old}</${w}:pPr></${w}:pPrChange>`);
        pPr.children.push(chg); chg.parent = pPr; markDirty(pPr);
      }
      const old = pPr.child(`${w}:numPr`);
      if (old) removeNode(old);
      const styleHasNum = this.styles.get(this.styleIdOf(pe))?.hasNumbering;
      if (p.kind === "none") {
        if (styleHasNum) {
          this.setProp(pPr, "numPr", {}, PPR_ORDER, w);
          this.fillNumPr(pPr, w, "0", "0");
        }
      } else {
        const [np] = parseFragment(`<${w}:numPr><${w}:ilvl ${w}:val="${level}"/><${w}:numId ${w}:val="${numId}"/></${w}:numPr>`);
        pPr.children.splice(orderedInsertIndex(pPr, "numPr", PPR_ORDER, w), 0, np);
        np.parent = pPr; markDirty(pPr);
        // 列表段落去掉首行缩进，由编号定义的悬挂缩进决定
        const ind = pPr.child(`${w}:ind`);
        if (ind) removeNode(ind);
      }
      preview.push({ ref: pe.ref, before: this.textOf(pe), after: `${p.kind === "none" ? "（取消编号）" : p.kind === "bullet" ? "• " : "1. "}${this.textOf(pe)}` });
    }
    const label = { bullet: "项目符号", number: "数字编号", chinese: "中文编号（一、（一）1.）", outline: "多级编号（1. 1.1 1.1.1）", none: "取消编号" }[p.kind];
    return { changedRefs: entries.map((e) => e.ref), summary: `已为 ${entries.length} 个段落设置${label}${p.kind !== "none" ? `（第 ${level + 1} 级${p.restart ? "，从 1 重新开始" : ""}）` : ""}。`, preview, structural: false };
  }

  private fillNumPr(pPr: XElement, w: string, ilvl: string, numId: string): void {
    const np = pPr.child(`${w}:numPr`);
    if (!np) return;
    appendChild(np, parseFragment(`<${w}:ilvl ${w}:val="${ilvl}"/><${w}:numId ${w}:val="${numId}"/>`));
  }

  // ------------------------------- 脚注 / 尾注 -------------------------------

  insertNote(p: { ref: string; after_text?: string; text: string; kind?: "footnote" | "endnote" }, o: EditOptions): EditResult {
    const pe = this.resolve(p.ref);
    if (pe.part !== this.mainPart) throw new DocError("脚注只能加在正文段落上", "unsupported");
    this.isEditable(pe);
    const kind = p.kind ?? "footnote";
    const w = this.w(pe.part);
    // 1) 找到 / 创建脚注部件
    const relType = kind === "footnote" ? REL_TYPES.footnotes : REL_TYPES.endnotes;
    const rels = parseRels(this.optText(relsPathFor(this.mainPart)));
    let part = rels.filter((r) => r.type === relType).map((r) => resolveTarget(this.mainPart, r.target)).find((x) => this.pkg.has(x));
    const tag = kind;
    if (!part) {
      part = this.mainPart.replace(/[^/]+$/, `${kind}s.xml`);
      const sep = `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}><w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`;
      this.addPart(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:${tag}s xmlns:w="${NS.w}" xmlns:r="${NS.r}">${sep}</w:${tag}s>`);
      this.ensureOverrideCT(part, kind === "footnote" ? CT.footnotes : CT.endnotes);
      this.addRel(relType, this.relTarget(part));
      this.contentParts.push(part);
    }
    const nd = this.part(part);
    const nw = this.w(part);
    const id = Math.max(0, ...nd.root.elements().map((e) => Number(e.getAttr(`${nw}:id`)) || 0)) + 1;
    const textSt = this.styles.resolve(kind === "footnote" ? "footnote text" : "endnote text", "paragraph");
    const refSt = this.styles.resolve(kind === "footnote" ? "footnote reference" : "endnote reference", "character");
    const refRPr = refSt ? `<${nw}:rPr><${nw}:rStyle ${nw}:val="${refSt.id}"/></${nw}:rPr>` : `<${nw}:rPr><${nw}:vertAlign ${nw}:val="superscript"/></${nw}:rPr>`;
    const noteXml = `<${nw}:${tag} ${nw}:id="${id}"><${nw}:p><${nw}:pPr>${textSt ? `<${nw}:pStyle ${nw}:val="${textSt.id}"/>` : `<${nw}:spacing ${nw}:after="0"/>`}</${nw}:pPr><${nw}:r>${refRPr}<${nw}:${tag}Ref/></${nw}:r><${nw}:r>${textSt ? "" : `<${nw}:rPr><${nw}:sz ${nw}:val="18"/></${nw}:rPr>`}${runContent(nw, " " + p.text)}</${nw}:r></${nw}:p></${nw}:${tag}>`;
    appendChild(nd.root, parseFragment(noteXml));
    // 2) 在正文中插入引用标记
    const full = this.textOf(pe);
    const pos = p.after_text ? this.findRange(pe, p.after_text).end : full.replace(/[。．.!！?？;；,，]+$/, "").length;
    if (pos > 0 && pos < full.length) this.splitAt(pe, pos);
    const wRefRPr = refSt ? `<${w}:rPr><${w}:rStyle ${w}:val="${refSt.id}"/></${w}:rPr>` : `<${w}:rPr><${w}:vertAlign ${w}:val="superscript"/></${w}:rPr>`;
    const refRun = parseFragment(this.trackRun(w, `<${w}:r>${wRefRPr}<${w}:${tag}Reference ${w}:id="${id}"/></${w}:r>`, o));
    const { segs } = this.segments(pe);
    const prev = [...segs].reverse().find((s) => s.run && s.end <= pos && s.end > s.start);
    if (prev?.run) insertAfter(this.anchorOf(prev.run, pe.el), refRun);
    else appendChild(pe.el, refRun);
    // 新的注释段落追加在脚注/尾注部件末尾：只有排在它后面的部件（尾注）的序号引用会变化
    if (kind === "footnote" && this.contentParts.some((x) => /endnotes\.xml$/.test(x))) this._structureVersion++;
    this.reindex();
    return { changedRefs: [pe.ref], summary: `已在段落 ${pe.ref}${p.after_text ? `"${clip(p.after_text, 30)}"之后` : "句末"}插入${kind === "footnote" ? "脚注" : "尾注"} ${id}：${clip(p.text, 80)}`, preview: [{ ref: pe.ref, before: full, after: this.textOf(pe) + `\n［${kind === "footnote" ? "脚注" : "尾注"}${id}］${p.text}` }], structural: false };
  }

  // ------------------------------- 超链接 -------------------------------

  insertLink(p: { ref: string; text: string; url: string }, o: EditOptions): EditResult {
    const pe = this.resolve(p.ref);
    this.isEditable(pe);
    if (!/^(https?:\/\/|mailto:)/i.test(p.url)) throw new DocError("链接地址需以 http://、https:// 或 mailto: 开头", "invalid");
    const w = this.w(pe.part);
    const { start, end } = this.findRange(pe, p.text);
    this.splitAt(pe, start);
    this.splitAt(pe, end);
    const runs = this.coveredRuns(pe, start, end);
    if (!runs.length) throw new DocError("找不到可以加链接的文字", "not_found");
    if (runs.some((r) => r.closest(`${w}:hyperlink`))) throw new DocError("这段文字已经是超链接", "invalid");
    const top = runs.map((r) => this.anchorOf(r, pe.el) as XElement);
    if (new Set(top.map((t) => t.parent)).size > 1) throw new DocError("文字跨越了不同的结构（修订/域），请缩小范围", "protected");
    const rId = this.addRel(REL_EXTRA.hyperlink, p.url, { owner: pe.part, external: true });
    const linkSt = this.styles.resolve("hyperlink", "character");
    for (const r of runs) {
      const rPr = this.ensureChild(r, "rPr", [], w, true);
      if (linkSt) this.setProp(rPr, "rStyle", { val: linkSt.id }, RPR_ORDER, w);
      else {
        this.setProp(rPr, "color", { val: "0563C1" }, RPR_ORDER, w);
        this.setProp(rPr, "u", { val: "single" }, RPR_ORDER, w);
      }
    }
    const xml = `<${w}:hyperlink r:id="${rId}" xmlns:r="${NS.r}" ${w}:history="1">${top.map((t) => serializeNode(t, this.src(pe.part))).join("")}</${w}:hyperlink>`;
    const [link] = parseFragment(xml);
    insertBefore(top[0], [link]);
    for (const t of top) removeNode(t);
    void o;
    this.reindex();
    return { changedRefs: [pe.ref], summary: `已为"${clip(p.text, 40)}"添加超链接 ${p.url}。`, preview: [{ ref: pe.ref, before: p.text, after: `${p.text} → ${p.url}` }], structural: false };
  }

  // ------------------------------- 目录 -------------------------------

  insertToc(p: { anchor: string; position: "before" | "after"; levels?: number; title?: string }, o: EditOptions): EditResult {
    const anchor = this.resolve(p.anchor);
    if (anchor.part !== this.mainPart) throw new DocError("目录只能插在正文中", "invalid");
    const w = this.w(anchor.part);
    const levels = Math.max(1, Math.min(9, p.levels ?? 3));
    const heads = this.summary().headings.filter((h) => h.level >= 1 && h.level <= levels && h.location === "正文");
    const tocSt = (lv: number) => this.styles.resolve(`toc ${lv}`, "paragraph") ?? this.styles.resolve(`目录 ${lv}`, "paragraph");
    const width = this.textWidthTwips(anchor);
    const rPr = this.baseFontRPr(anchor);
    let xml = "";
    if (p.title !== "") {
      const st = this.styles.resolve("TOC Heading", "paragraph") ?? this.styles.resolve("目录标题", "paragraph");
      xml += this.paraXml(anchor.part, st ? `<${w}:pStyle ${w}:val="${st.id}"/>` : `<${w}:spacing ${w}:before="240" ${w}:after="120"/><${w}:jc ${w}:val="center"/>`, this.trackRun(w, `<${w}:r>${st ? "" : this.baseFontRPr(anchor, 16, `<${w}:b/><${w}:bCs/>`)}${runContent(w, p.title ?? "目  录")}</${w}:r>`, o), o);
    }
    const begin = `<${w}:r><${w}:fldChar ${w}:fldCharType="begin" ${w}:dirty="true"/></${w}:r><${w}:r><${w}:instrText xml:space="preserve"> TOC \\o "1-${levels}" \\h \\z \\u </${w}:instrText></${w}:r><${w}:r><${w}:fldChar ${w}:fldCharType="separate"/></${w}:r>`;
    const end = `<${w}:r><${w}:fldChar ${w}:fldCharType="end"/></${w}:r>`;
    const entries = heads.length ? heads : [{ level: 1, text: "（打开文档后右键目录 → 更新域，生成目录）" }];
    entries.forEach((h, k) => {
      const st = tocSt(h.level);
      const pPr = `${st ? `<${w}:pStyle ${w}:val="${st.id}"/>` : ""}<${w}:tabs><${w}:tab ${w}:val="right" ${w}:leader="dot" ${w}:pos="${width}"/></${w}:tabs>${st ? "" : `<${w}:spacing ${w}:after="60"/><${w}:ind ${w}:left="${(h.level - 1) * 420}"/>`}`;
      const text = h.text.replace(/⟨[^⟩]*⟩/g, "").trim();
      const body = `<${w}:r>${rPr}${runContent(w, text)}</${w}:r><${w}:r>${rPr}<${w}:tab/></${w}:r>`;
      xml += this.paraXml(anchor.part, pPr, this.trackRun(w, (k === 0 ? begin : "") + body + (k === entries.length - 1 ? end : ""), o), o);
    });
    const nodes = this.insertBlockXml(anchor, p.position, xml);
    this.setUpdateFieldsOnOpen();
    return { changedRefs: this.refsOf(nodes), summary: `已插入目录（${levels} 级标题，${heads.length} 个条目）。目录是 Word 的 TOC 域：打开文件时 Word 会提示更新域，更新后页码自动填好。`, preview: [{ ref: "", before: "", after: entries.map((h) => `${"  ".repeat(h.level - 1)}${h.text}`).join("\n") }], structural: true };
  }

  /** settings.xml 里设置 updateFields，Word 打开时自动刷新目录 / 页码等域 */
  private setUpdateFieldsOnOpen(): void {
    const rels = parseRels(this.optText(relsPathFor(this.mainPart)));
    const r = rels.find((x) => x.type === REL_EXTRA.settings);
    if (!r) return;
    const name = resolveTarget(this.mainPart, r.target);
    if (!this.pkg.has(name)) return;
    const doc = this.part(name);
    const w = this.w(name);
    if (doc.root.child(`${w}:updateFields`)) return;
    // CT_Settings 中 updateFields 位于 hdrShapeDefaults / footnotePr 等之前、trackRevisions 等之后：
    // 放在第一个"必须在它之后"的元素前面
    const after = new Set(["hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"]);
    const [node] = parseFragment(`<${w}:updateFields ${w}:val="true"/>`);
    const kids = doc.root.children;
    const idx = kids.findIndex((c) => c.type === "el" && after.has(c.local));
    if (idx === -1) appendChild(doc.root, [node]);
    else { kids.splice(idx, 0, node); node.parent = doc.root; markDirty(doc.root); }
  }

  // ------------------------------- 样式 -------------------------------

  modifyStyle(p: { style: string; type?: "paragraph" | "character"; create?: boolean; based_on?: string; font?: string; size_pt?: number; bold?: boolean; italic?: boolean; color?: string; alignment?: "left" | "center" | "right" | "justify"; space_before_pt?: number; space_after_pt?: number; line_spacing?: number; first_line_pt?: number; indent_left_pt?: number; keep_with_next?: boolean; outline_level?: number }, _o: EditOptions): EditResult {
    if (!this.stylesPart) throw new DocError("文档没有样式表（styles.xml）", "unsupported");
    const part = this.stylesPart;
    const doc = this.part(part);
    const w = this.w(part);
    const type = p.type ?? "paragraph";
    let def = this.styles.resolve(p.style, type);
    let el: XElement | undefined;
    let created = false;
    if (def) el = doc.root.elements().find((e) => e.local === "style" && e.getAttr(`${w}:styleId`) === def!.id);
    if (!el) {
      if (!p.create) throw new DocError(`找不到样式"${p.style}"。设置 create=true 可以新建；用 doc_styles 查看已有样式。`, "not_found");
      const base = p.based_on ? this.styles.resolve(p.based_on, type) : this.styles.defaultParagraphStyle();
      const id = p.style.replace(/[^A-Za-z0-9]/g, "") || `AgentStyle${doc.root.elements().length}`;
      const xml = `<${w}:style ${w}:type="${type}" ${w}:customStyle="1" ${w}:styleId="${encodeAttr(id)}"><${w}:name ${w}:val="${encodeAttr(p.style)}"/>${base && type === "paragraph" ? `<${w}:basedOn ${w}:val="${base.id}"/>` : ""}<${w}:qFormat/></${w}:style>`;
      [el] = parseFragment(xml) as XElement[];
      appendChild(doc.root, [el]);
      created = true;
    }
    const changes: string[] = [];
    if (type === "paragraph") {
      const pPr = this.ensureChild(el, "pPr", STYLE_ORDER, w);
      if (p.alignment) { this.setProp(pPr, "jc", { val: p.alignment === "justify" ? "both" : p.alignment }, PPR_ORDER, w); changes.push(`对齐 ${p.alignment}`); }
      const sp: Record<string, string> = {};
      if (p.space_before_pt !== undefined) { sp.before = String(twip(p.space_before_pt)); changes.push(`段前 ${p.space_before_pt}pt`); }
      if (p.space_after_pt !== undefined) { sp.after = String(twip(p.space_after_pt)); changes.push(`段后 ${p.space_after_pt}pt`); }
      if (p.line_spacing !== undefined) { sp.line = String(Math.round(p.line_spacing * 240)); sp.lineRule = "auto"; changes.push(`行距 ${p.line_spacing} 倍`); }
      if (Object.keys(sp).length) { const e = this.ensureChild(pPr, "spacing", PPR_ORDER, w); for (const [k, v] of Object.entries(sp)) e.setAttr(`${w}:${k}`, v); }
      const ind: Record<string, string> = {};
      if (p.first_line_pt !== undefined) { ind.firstLine = String(twip(p.first_line_pt)); changes.push(`首行缩进 ${p.first_line_pt}pt`); }
      if (p.indent_left_pt !== undefined) { ind.left = String(twip(p.indent_left_pt)); changes.push(`左缩进 ${p.indent_left_pt}pt`); }
      if (Object.keys(ind).length) { const e = this.ensureChild(pPr, "ind", PPR_ORDER, w); for (const [k, v] of Object.entries(ind)) e.setAttr(`${w}:${k}`, v); if (ind.firstLine) e.removeAttr(`${w}:hanging`); }
      if (p.keep_with_next !== undefined) { this.setProp(pPr, "keepNext", p.keep_with_next ? {} : { val: "0" }, PPR_ORDER, w); changes.push(p.keep_with_next ? "与下段同页" : "取消与下段同页"); }
      if (p.outline_level !== undefined) { this.setProp(pPr, "outlineLvl", p.outline_level > 0 ? { val: String(p.outline_level - 1) } : null, PPR_ORDER, w); changes.push(`大纲级别 ${p.outline_level}`); }
      if (!pPr.elements().length) removeNode(pPr);
    }
    const rPr = this.ensureChild(el, "rPr", STYLE_ORDER, w);
    if (p.font) { this.setProp(rPr, "rFonts", { ascii: p.font, hAnsi: p.font, eastAsia: p.font, cs: p.font }, RPR_ORDER, w); changes.push(`字体 ${p.font}`); }
    if (p.bold !== undefined) { this.setProp(rPr, "b", p.bold ? {} : { val: "0" }, RPR_ORDER, w); this.setProp(rPr, "bCs", p.bold ? {} : { val: "0" }, RPR_ORDER, w); changes.push(p.bold ? "加粗" : "不加粗"); }
    if (p.italic !== undefined) { this.setProp(rPr, "i", p.italic ? {} : { val: "0" }, RPR_ORDER, w); changes.push(p.italic ? "斜体" : "非斜体"); }
    if (p.color) { this.setProp(rPr, "color", { val: p.color.replace(/^#/, "").toUpperCase() }, RPR_ORDER, w); changes.push(`颜色 ${p.color}`); }
    if (p.size_pt) { const hp = String(Math.round(p.size_pt * 2)); this.setProp(rPr, "sz", { val: hp }, RPR_ORDER, w); this.setProp(rPr, "szCs", { val: hp }, RPR_ORDER, w); changes.push(`字号 ${p.size_pt}pt`); }
    if (!rPr.elements().length) removeNode(rPr);
    markDirty(el);
    // 刷新样式表
    this.styles = new StyleSheet(serializeDocument(doc));
    const uses = this.paras.filter((x) => this.styleIdOf(x) === (def?.id ?? el!.getAttr(`${w}:styleId`))).length;
    def = this.styles.resolve(p.style, type);
    return { changedRefs: [], summary: `${created ? "已新建" : "已修改"}样式"${p.style}"：${changes.join("，") || "（无属性变化）"}。文档中使用该样式的 ${uses} 个段落会一起变化（直接格式优先于样式）。`, preview: [{ ref: "", before: `样式 ${p.style}`, after: changes.join("，") }], structural: false };
  }

  // ------------------------------- 修订审阅 -------------------------------

  reviewChanges(p: { action: "accept" | "reject"; refs?: string[]; author?: string }, _o: EditOptions): EditResult {
    const scopeParas = p.refs?.length ? p.refs.map((r) => this.resolve(r)) : null;
    let count = 0;
    const parts = scopeParas ? [...new Set(scopeParas.map((x) => x.part))] : this.contentParts;
    const accept = p.action === "accept";
    for (const part of parts) {
      const w = this.w(part);
      const byAuthor = (el: XElement) => !p.author || el.getAttr(`${w}:author`) === p.author;
      const roots: XElement[] = scopeParas ? scopeParas.filter((x) => x.part === part).map((x) => x.el) : [this.part(part).root];
      for (const root of roots) {
        const isMark = (el: XElement) => ["rPr", "trPr"].includes((el.parent as XElement)?.local ?? "");
        const revs = () => [root, ...root.descendants()].filter((d) => d.prefix === w && ["ins", "del", "moveFrom", "moveTo"].includes(d.local) && byAuthor(d));
        // 第一遍：段落内的插入/删除内容
        for (const el of revs().filter((e) => !isMark(e))) {
          if (!el.parent) continue;
          const isAdd = el.local === "ins" || el.local === "moveTo";
          if (accept === isAdd) {
            insertBefore(el, [...el.children]);
            removeNode(el);
          } else removeNode(el);
          count++;
        }
        // 第二遍：段落标记、表格行标记
        for (const el of revs().filter(isMark)) {
          if (!el.parent) continue;
          const isAdd = el.local === "ins" || el.local === "moveTo";
          const owner = el.parent as XElement;
          removeNode(el);
          count++;
          if (accept === isAdd) continue;
          if (owner.local === "trPr") {
            const tr = owner.parent as XElement;
            if (tr?.parent) removeNode(tr);
          } else {
            const pEl = (owner.parent as XElement | null)?.parent as XElement | undefined;
            if (pEl?.local === "p" && pEl.parent && ![...pEl.descendants()].some((d) => (d.local === "t" && d.prefix === w) || d.local === "drawing")) {
              const parent = pEl.parent as XElement;
              if (!(parent.local === "tc" && parent.childrenNamed(`${w}:p`).length <= 1)) removeNode(pEl);
            }
          }
        }
        // 格式修订：接受 = 去掉记录；拒绝 = 恢复记录里的旧属性
        for (const el of [root, ...root.descendants()].filter((d) => d.prefix === w && /PrChange$/.test(d.local) && byAuthor(d))) {
          if (!el.parent) continue;
          const owner = el.parent as XElement;
          if (!accept) {
            const oldPr = el.elements()[0];
            const keep = owner.local === "pPr" ? owner.children.filter((c) => c.type === "el" && ["rPr", "sectPr"].includes(c.local)) : [];
            for (const c of [...owner.children]) removeNode(c);
            if (oldPr) appendChild(owner, [...oldPr.children]);
            if (keep.length) appendChild(owner, keep);
          } else removeNode(el);
          count++;
        }
      }
      // 拒绝删除后，留下的 delText / delInstrText 改回普通文字
      if (!accept) {
        for (const d of [...this.part(part).root.descendants()]) {
          if (d.prefix !== w || d.closest(`${w}:del`) || d.closest(`${w}:moveFrom`)) continue;
          if (d.local === "delText") { d.name = `${w}:t`; d.attrsDirty = true; markDirty(d); }
          else if (d.local === "delInstrText") { d.name = `${w}:instrText`; d.attrsDirty = true; markDirty(d); }
        }
      }
      markDirty(this.part(part).root);
    }
    this._structureVersion++;
    this.reindex();
    return { changedRefs: [], summary: count ? `已${accept ? "接受" : "拒绝"} ${count} 处修订${p.author ? `（作者 ${p.author}）` : ""}${scopeParas ? `（范围：${scopeParas.length} 个段落）` : "（全文）"}。` : "范围内没有修订标记。", preview: [{ ref: "", before: "修订", after: `${accept ? "接受" : "拒绝"} ${count} 处` }], structural: true };
  }

  // ------------------------------- 批注管理 -------------------------------

  listComments(): Array<{ id: string; author: string; date: string; text: string; ref?: string; anchor: string }> {
    const cp = this.commentsPartName();
    if (!cp || !this.pkg.has(cp)) return [];
    const cw = this.w(cp);
    const w = this.w(this.mainPart);
    const out: Array<{ id: string; author: string; date: string; text: string; ref?: string; anchor: string }> = [];
    for (const c of this.part(cp).root.elements()) {
      if (c.local !== "comment") continue;
      const id = c.getAttr(`${cw}:id`) ?? "";
      const text = c.elements().map((p) => innerText(p)).join("\n").trim();
      const pe = this.paras.find((x) => x.part === this.mainPart && [...x.el.descendants()].some((d) => d.prefix === w && (d.local === "commentRangeStart" || d.local === "commentReference") && d.getAttr(`${w}:id`) === id));
      out.push({ id, author: c.getAttr(`${cw}:author`) ?? "", date: c.getAttr(`${cw}:date`) ?? "", text, ref: pe?.ref, anchor: pe ? clip(this.textOf(pe), 60) : "" });
    }
    return out;
  }

  deleteComments(p: { ids?: string[]; all?: boolean }): EditResult {
    const cp = this.commentsPartName();
    if (!cp || !this.pkg.has(cp)) throw new DocError("文档中没有批注", "not_found");
    const cw = this.w(cp);
    const w = this.w(this.mainPart);
    const want = new Set(p.all ? this.listComments().map((c) => c.id) : (p.ids ?? []).map(String));
    if (!want.size) throw new DocError("请提供要删除的批注 id（doc_comments action=list 查看），或 all=true", "invalid");
    let n = 0;
    for (const c of [...this.part(cp).root.elements()]) if (c.local === "comment" && want.has(c.getAttr(`${cw}:id`) ?? "")) { removeNode(c); n++; }
    for (const d of [...this.part(this.mainPart).root.descendants()]) {
      if (d.prefix !== w || !d.parent) continue;
      if ((d.local === "commentRangeStart" || d.local === "commentRangeEnd") && want.has(d.getAttr(`${w}:id`) ?? "")) removeNode(d);
      else if (d.local === "commentReference" && want.has(d.getAttr(`${w}:id`) ?? "")) {
        const run = d.parent as XElement;
        if (run.local === "r" && run.elements().every((e) => e.local === "rPr" || e === d)) removeNode(run);
        else removeNode(d);
      }
    }
    this.reindex();
    return { changedRefs: [], summary: `已删除 ${n} 条批注。`, preview: [{ ref: "", before: `${n} 条批注`, after: "" }], structural: false };
  }

  // ------------------------------- 移动段落 -------------------------------

  moveBlocks(p: { refs: string[]; anchor: string; position: "before" | "after" }, o: EditOptions): EditResult {
    const anchor = this.resolve(p.anchor);
    const items = p.refs.map((r) => this.resolve(r));
    for (const pe of items) {
      this.isEditable(pe);
      if (pe.part !== anchor.part || pe.container !== anchor.container) throw new DocError("只能在同一部件、同类容器内移动段落（正文到正文、同一表格单元格内）。", "invalid");
      if (pe.el.child(`${this.w(pe.part)}:pPr`)?.child(`${this.w(pe.part)}:sectPr`)) throw new DocError(`段落 ${pe.ref} 携带分节符，不能移动`, "protected");
      if (pe === anchor) throw new DocError("锚点不能是被移动的段落之一", "invalid");
    }
    if (o.track) throw new DocError("修订模式下暂不支持移动段落（Word 的移动修订需要成对标记）。请先关闭修订模式。", "unsupported");
    const ordered = [...items].sort((a, b) => a.ordinal - b.ordinal);
    for (const pe of ordered) removeNode(pe.el);
    if (p.position === "before") insertBefore(anchor.el, ordered.map((x) => x.el));
    else insertAfter(anchor.el, ordered.map((x) => x.el));
    this._structureVersion++;
    this.reindex();
    return { changedRefs: ordered.map((x) => this.paras.find((y) => y.el === x.el)!.ref), summary: `已把 ${ordered.length} 个段落移动到 ${anchor.ref} ${p.position === "before" ? "之前" : "之后"}。`, preview: ordered.map((x) => ({ ref: x.ref, before: "（原位置）", after: clip(this.textOf(x), 80) })), structural: true };
  }

  // -------------------------------------------------------------------------
  // 序列化与保真校验
  // -------------------------------------------------------------------------

  serialize(): Buffer {
    for (const [name, doc] of this.parts) {
      if (doc.dirty) this.pkg.write(name, Buffer.from(serializeDocument(doc), "utf8"));
    }
    return this.pkg.toBuffer();
  }

  convertedLayout(): "exact" | "flow" | undefined {
    const txt = (n: string) => (this.pkg.has(n) ? this.pkg.read(n).toString("utf8") : "");
    const m = /name="DocAgentLayout"[^>]*>\s*<vt:lpwstr>(exact|flow)</.exec(txt("docProps/custom.xml"));
    if (m) return m[1] as "exact" | "flow";
    // 早期版本转换的文档没有自定义属性：按"PDF 转换"标记与逐页分页样式判断
    if (txt("docProps/app.xml").includes("文案 Agent PDF 转换")) return txt(this.mainPart).includes('w:val="PageStart"') ? "exact" : "flow";
    return undefined;
  }

  fidelity(original: Buffer): FidelityReport {
    const orig = new ZipPackage(original);
    const cur = new ZipPackage(this.serialize());
    const problems: string[] = [];
    const modified: FidelityReport["modifiedParts"] = [];
    const added: string[] = [];
    let identical = 0;
    for (const name of cur.names()) {
      let data: Buffer;
      try {
        data = cur.read(name);
      } catch (e: any) {
        problems.push(`部件 ${name} 无法解压：${e.message}`);
        continue;
      }
      if (!orig.has(name)) {
        added.push(name);
        if (name.endsWith(".xml") || name.endsWith(".rels")) this.checkWellFormed(name, data, problems);
        continue;
      }
      const before = orig.read(name);
      if (before.equals(data)) {
        identical++;
        continue;
      }
      if (name.endsWith(".xml") || name.endsWith(".rels")) this.checkWellFormed(name, data, problems);
      modified.push({ name, changedBlocks: this.changedBlocks(before.toString("utf8"), data.toString("utf8")) });
    }
    for (const name of orig.names()) if (!cur.has(name)) problems.push(`部件 ${name} 在编辑后丢失`);
    // 修订 ID 唯一性
    const seen = new Map<string, number>();
    const main = parseXml(cur.readText(this.mainPart));
    const w = this.w(this.mainPart);
    for (const d of main.root.descendants()) {
      if (d.prefix === w && (d.local === "ins" || d.local === "del")) {
        const id = d.getAttr(`${w}:id`) ?? "";
        seen.set(id, (seen.get(id) ?? 0) + 1);
      }
    }
    for (const [id, n] of seen) if (n > 1) problems.push(`修订 ID ${id} 重复 ${n} 次`);
    return {
      ok: problems.length === 0,
      totalParts: cur.names().length,
      identicalParts: identical,
      modifiedParts: modified,
      addedParts: added,
      problems,
    };
  }

  private checkWellFormed(name: string, data: Buffer, problems: string[]): void {
    try {
      parseXml(data.toString("utf8"));
    } catch (e: any) {
      problems.push(`${name} 不是合法 XML：${e.message}`);
    }
  }

  /** 对比两个版本的段落级差异，返回变化段落的描述 */
  /** 段落级对比：按顺序做序列差分，报告"第 N 段 修改 / 新增 / 删除"及文字片段 */
  private changedBlocks(before: string, after: string): string[] {
    const collect = (xml: string) => {
      const doc = parseXml(xml);
      const out: Array<{ key: string; xml: string; text: string }> = [];
      const walk = (n: XElement) => {
        for (const c of n.children) {
          if (c.type !== "el") continue;
          if (c.local === "p") {
            const text = [...c.descendants()].filter((d) => d.local === "t").map((d) => innerText(d)).join("");
            out.push({ key: c.getAttr("w14:paraId") ?? "", xml: serializeNode(c, doc.source), text });
          } else walk(c);
        }
      };
      walk(doc.root);
      return out;
    };
    let a: Array<{ key: string; xml: string; text: string }>, b: Array<{ key: string; xml: string; text: string }>;
    try {
      a = collect(before);
      b = collect(after);
    } catch {
      return ["（无法解析，无法给出段落级对比）"];
    }
    const snip = (t: string) => `「${t.length > 18 ? t.slice(0, 18) + "…" : t || "空段落"}」`;
    const label = (x: { key: string }, n: number) => (x.key ? `#${x.key}` : `第${n}段`);
    const parts = diffArrays(a.map((x) => x.xml), b.map((x) => x.xml));
    const out: string[] = [];
    let ai = 0, bi = 0;
    for (let i = 0; i < parts.length; i++) {
      const pt = parts[i];
      const n = pt.count ?? pt.value.length;
      if (!pt.added && !pt.removed) { ai += n; bi += n; continue; }
      if (pt.removed && parts[i + 1]?.added) {
        const m = parts[i + 1].count ?? parts[i + 1].value.length;
        const pairs = Math.min(n, m);
        for (let k = 0; k < pairs; k++) out.push(`${label(b[bi + k], bi + k + 1)}${snip(b[bi + k].text)} 已修改`);
        for (let k = pairs; k < m; k++) out.push(`${label(b[bi + k], bi + k + 1)}${snip(b[bi + k].text)} 新增`);
        for (let k = pairs; k < n; k++) out.push(`原第${ai + k + 1}段${snip(a[ai + k].text)} 已删除`);
        ai += n; bi += m; i++;
      } else if (pt.added) {
        for (let k = 0; k < n; k++) out.push(`${label(b[bi + k], bi + k + 1)}${snip(b[bi + k].text)} 新增`);
        bi += n;
      } else {
        for (let k = 0; k < n; k++) out.push(`原第${ai + k + 1}段${snip(a[ai + k].text)} 已删除`);
        ai += n;
      }
    }
    return out;
  }

}
