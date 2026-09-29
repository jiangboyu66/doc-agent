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
  resolveTarget, orderedInsertIndex, ptToTwips,
} from "./ooxml.js";
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
  };
  private _structureVersion = 0;
  /** 结构版本号；外部（会话重载 / 干跑副本）设置时需要重建引用表 */
  get structureVersion(): number { return this._structureVersion; }
  set structureVersion(v: number) {
    if (v === this._structureVersion) return;
    this._structureVersion = v;
    if (this.paras) this.reindex();
  }

  private pkg: ZipPackage;
  private parts = new Map<string, XDocument>();
  private prefixes = new Map<string, string>();
  readonly mainPart: string;
  private contentParts: string[] = [];
  private styles: StyleSheet;
  private paras: ParaEntry[] = [];
  private byRef = new Map<string, ParaEntry>();
  private nextRevisionId = -1;

  constructor(buf: Buffer) {
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

  private reindex(): void {
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

  private resolve(ref: string): ParaEntry {
    const r = ref.trim();
    const direct = this.byRef.get(r) ?? this.byRef.get(`#${r}`);
    if (direct) return direct;
    const m = /^P(\d+)(?:@v(\d+))?$/.exec(r);
    if (m) {
      if (m[2] !== undefined && Number(m[2]) !== this.structureVersion) {
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

  // -------------------------------------------------------------------------
  // 序列化与保真校验
  // -------------------------------------------------------------------------

  serialize(): Buffer {
    for (const [name, doc] of this.parts) {
      if (doc.dirty) this.pkg.write(name, Buffer.from(serializeDocument(doc), "utf8"));
    }
    return this.pkg.toBuffer();
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
