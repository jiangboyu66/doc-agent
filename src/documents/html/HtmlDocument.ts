/**
 * HTML 原生适配器
 *
 * 用 parse5（浏览器同款 HTML 解析算法）解析并拿到每个节点在源文件中的字符位置，
 * 修改时只对文本节点所在的源码区间做拼接——标签、属性、CSS、脚本、缩进、注释全部原样保留。
 */

import { parse } from "parse5";
import type { DefaultTreeAdapterMap } from "parse5";
import { findAll, computeHunks, clip } from "../textMatch.js";
import {
  DocError, type DocumentAdapter, type DocumentCapabilities, type BlockInfo, type SearchHit, type DocumentSummary,
  type EditOptions, type EditResult, type RunFormat, type ParagraphFormat, type NewBlock, type FidelityReport,
} from "../types.js";

type Node = DefaultTreeAdapterMap["node"];
type Element = DefaultTreeAdapterMap["element"];
type TextNode = DefaultTreeAdapterMap["textNode"];

const BLOCK_TAGS = new Set(["p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "td", "th", "dt", "dd", "pre", "figcaption", "caption", "title", "blockquote", "summary"]);
const SKIP_TAGS = new Set(["script", "style", "template", "noscript", "svg", "math"]);
const CONTAINER_BLOCKS = new Set(["div", "section", "article", "header", "footer", "main", "aside", "nav", "body", "blockquote"]);

interface TextSeg {
  node: TextNode;
  start: number; // 在块渲染文本中的位置
  end: number;
  /** 解码后文本的每个字符在源码中的偏移，长度 = 文本长度 + 1 */
  map: number[] | null;
}

interface HBlock {
  el: Element;
  segs: TextSeg[];
  text: string;
}

export class HtmlDocument implements DocumentAdapter {
  readonly format = "html" as const;
  readonly capabilities: DocumentCapabilities = {
    replaceText: true, formatText: true, paragraphProps: true, insertBlocks: true, deleteBlocks: true,
    tables: false, comments: true, trackChanges: false, styles: false, rawInsert: true,
  };
  structureVersion = 0;
  private src: string;
  private blocks: HBlock[] = [];

  constructor(buf: Buffer) {
    this.src = buf.toString("utf8");
    this.parse();
  }

  private parse(): void {
    const doc = parse(this.src, { sourceCodeLocationInfo: true });
    const blocks: HBlock[] = [];
    const collect = (el: Element, block: HBlock) => {
      for (const c of el.childNodes) {
        if (c.nodeName === "#text") {
          const t = c as TextNode;
          const loc = t.sourceCodeLocation;
          const start = block.text.length;
          block.text += t.value;
          block.segs.push({ node: t, start, end: block.text.length, map: loc ? this.alignMap(t.value, loc.startOffset, loc.endOffset) : null });
        } else if ("tagName" in c) {
          const e = c as Element;
          if (SKIP_TAGS.has(e.tagName) || BLOCK_TAGS.has(e.tagName) || (CONTAINER_BLOCKS.has(e.tagName) && this.hasDirectText(e))) continue;
          collect(e, block);
        }
      }
    };
    const walk = (n: Node) => {
      if (!("childNodes" in n)) return;
      for (const c of n.childNodes) {
        if (!("tagName" in c)) continue;
        const e = c as Element;
        if (SKIP_TAGS.has(e.tagName)) continue;
        if (BLOCK_TAGS.has(e.tagName) || (CONTAINER_BLOCKS.has(e.tagName) && this.hasDirectText(e))) {
          const b: HBlock = { el: e, segs: [], text: "" };
          collect(e, b);
          if (b.text.trim() && e.sourceCodeLocation) blocks.push(b);
        }
        walk(e);
      }
    };
    walk(doc);
    this.blocks = blocks;
  }

  private hasDirectText(e: Element): boolean {
    return e.childNodes.some((c) => c.nodeName === "#text" && (c as TextNode).value.trim());
  }

  /** 把解码文本逐字符对齐到源码偏移（处理 &amp; &nbsp; &#123; 等实体） */
  private alignMap(decoded: string, s: number, e: number): number[] | null {
    const raw = this.src.slice(s, e);
    const map: number[] = [];
    let i = 0, j = 0;
    while (j < decoded.length) {
      if (i >= raw.length) return null;
      if (raw[i] === "&") {
        const semi = raw.indexOf(";", i);
        if (semi !== -1 && semi - i <= 32) {
          const cp = decoded.codePointAt(j)!;
          const width = cp > 0xffff ? 2 : 1;
          for (let k = 0; k < width; k++) map.push(s + i);
          j += width;
          i = semi + 1;
          continue;
        }
      }
      if (raw[i] !== decoded[j]) {
        if (raw[i] === "\r" && decoded[j] === "\n") { i++; continue; }
        return null;
      }
      map.push(s + i);
      i++;
      j++;
    }
    map.push(s + i);
    return map;
  }

  private ref(i: number) {
    return `H${i + 1}@v${this.structureVersion}`;
  }

  private resolve(ref: string): number {
    const m = /^H(\d+)(?:@v(\d+))?$/.exec(ref.trim());
    if (!m) throw new DocError(`无效的块引用 ${ref}`, "not_found");
    if (m[2] !== undefined && Number(m[2]) !== this.structureVersion) throw new DocError(`块引用 ${ref} 已过期，请重新读取。`, "stale_ref");
    const i = Number(m[1]) - 1;
    if (!this.blocks[i]) throw new DocError(`找不到块 ${ref}`, "not_found");
    return i;
  }

  private display(t: string) {
    return t.replace(/\s+/g, " ").trim();
  }

  private info(i: number): BlockInfo {
    const b = this.blocks[i];
    const tag = b.el.tagName;
    const h = /^h(\d)$/.exec(tag);
    return {
      ref: this.ref(i), location: "正文",
      kind: h ? "heading" : tag === "li" ? "list" : tag === "td" || tag === "th" ? "table-cell" : tag === "pre" ? "code" : tag === "blockquote" ? "quote" : tag === "title" ? "title" : "paragraph",
      level: h ? Number(h[1]) : undefined,
      style: `<${tag}>`,
      text: this.display(b.text),
    };
  }

  summary(): DocumentSummary {
    return {
      format: "html", blockCount: this.blocks.length,
      charCount: this.blocks.reduce((s, b) => s + b.text.replace(/\s/g, "").length, 0),
      headings: this.blocks.map((_, i) => this.info(i)).filter((b) => b.kind === "heading" || b.kind === "title").map((b) => ({ ref: b.ref, level: b.level ?? 0, text: b.text, location: b.location })),
      parts: ["源文件"], tables: (this.src.match(/<table\b/gi) ?? []).length, images: (this.src.match(/<img\b/gi) ?? []).length,
      sections: 1, trackedChanges: 0, comments: 0, notes: ["HTML 的外观由标签与 CSS 决定；修改只触及文本节点源码。"],
    };
  }

  listBlocks(): BlockInfo[] { return this.blocks.map((_, i) => this.info(i)); }
  getBlock(ref: string): BlockInfo { return this.info(this.resolve(ref)); }
  listStyles() { return []; }
  describeFormat(ref: string): string {
    const b = this.blocks[this.resolve(ref)];
    const loc = b.el.sourceCodeLocation!;
    return `块 ${ref}：<${b.el.tagName}>，源码：\n${clip(this.src.slice(loc.startOffset, loc.endOffset), 600)}`;
  }

  search(query: string, opts: { regex?: boolean; caseSensitive?: boolean }): SearchHit[] {
    const hits: SearchHit[] = [];
    this.blocks.forEach((b, i) => {
      const t = this.display(b.text);
      const ms = opts.regex
        ? [...t.matchAll(new RegExp(query, opts.caseSensitive ? "g" : "gi"))].map((m) => ({ start: m.index ?? 0, end: (m.index ?? 0) + m[0].length }))
        : findAll(t, query, opts.caseSensitive ?? false);
      for (const m of ms) hits.push({ ref: this.ref(i), location: "正文", text: t, index: m.start, match: t.slice(m.start, m.end) });
    });
    return hits;
  }

  private esc(s: string) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  /** 把块内 [start,end) 的文字替换为 text，返回源码拼接操作 */
  private planHunk(b: HBlock, start: number, end: number, text: string): Array<{ s: number; e: number; t: string }> {
    const ops: Array<{ s: number; e: number; t: string }> = [];
    const touched = b.segs.filter((sg) => sg.end > start && sg.start < end);
    if (!touched.length) {
      const host = [...b.segs].reverse().find((sg) => sg.end <= start) ?? b.segs.find((sg) => sg.start >= end);
      if (!host?.map) throw new DocError("无法定位插入位置", "unsupported");
      const off = Math.max(0, Math.min(host.map.length - 1, start - host.start));
      ops.push({ s: host.map[off], e: host.map[off], t: this.esc(text) });
      return ops;
    }
    touched.forEach((sg, k) => {
      if (!sg.map) throw new DocError("该文本包含无法精确定位的字符实体，暂不支持修改", "unsupported");
      const a = Math.max(start, sg.start) - sg.start;
      const z = Math.min(end, sg.end) - sg.start;
      ops.push({ s: sg.map[a], e: sg.map[z], t: k === 0 ? this.esc(text) : "" });
    });
    return ops;
  }

  private apply(ops: Array<{ s: number; e: number; t: string }>) {
    for (const op of [...ops].sort((x, y) => y.s - x.s)) this.src = this.src.slice(0, op.s) + op.t + this.src.slice(op.e);
  }

  replaceText(p: { ref?: string; oldText: string; newText: string; replaceAll?: boolean }, _o: EditOptions): EditResult {
    const scope = p.ref ? [this.resolve(p.ref)] : this.blocks.map((_, i) => i);
    const matches: Array<{ i: number; start: number; end: number }> = [];
    for (const i of scope) for (const m of findAll(this.blocks[i].text, p.oldText)) matches.push({ i, ...m });
    if (!matches.length) throw new DocError(`找不到 old_text"${clip(p.oldText, 80)}"`, "not_found");
    if (matches.length > 1 && !p.replaceAll) throw new DocError(`old_text 出现了 ${matches.length} 次，请指定 ref、提供更长片段或设置 replace_all=true。`, "ambiguous");
    const before = new Map(matches.map((m) => [m.i, this.display(this.blocks[m.i].text)]));
    const ops: Array<{ s: number; e: number; t: string }> = [];
    for (const m of matches) {
      const b = this.blocks[m.i];
      for (const h of computeHunks(b.text.slice(m.start, m.end), p.newText)) ops.push(...this.planHunk(b, m.start + h.start, m.start + h.end, h.text));
    }
    this.apply(ops);
    this.parse();
    return {
      changedRefs: [...before.keys()].map((i) => this.ref(i)), summary: `已修改 ${matches.length} 处。`,
      preview: [...before.entries()].map(([i, t]) => ({ ref: this.ref(i), before: t, after: this.blocks[i] ? this.display(this.blocks[i].text) : "" })),
      structural: false,
    };
  }

  formatText(p: { ref: string; text?: string; format: RunFormat }, _o: EditOptions): EditResult {
    const i = this.resolve(p.ref);
    const b = this.blocks[i];
    const ms = p.text ? findAll(b.text, p.text) : [{ start: 0, end: b.text.length, fuzzy: false }];
    if (ms.length !== 1) throw new DocError("目标文字不存在或不唯一", ms.length ? "ambiguous" : "not_found");
    const sg = b.segs.find((s) => s.start <= ms[0].start && s.end >= ms[0].end);
    if (!sg?.map) throw new DocError("设置格式的文字必须位于同一个文本节点内（不能跨越已有标签）", "unsupported");
    const f = p.format;
    const styles: string[] = [];
    if (f.color) styles.push(`color:#${f.color.replace(/^#/, "")}`);
    if (f.size_pt) styles.push(`font-size:${f.size_pt}pt`);
    if (f.font) styles.push(`font-family:${f.font}`);
    let open = "", close = "";
    const wrap = (tag: string) => { open += `<${tag}>`; close = `</${tag}>` + close; };
    if (f.bold) wrap("strong");
    if (f.italic) wrap("em");
    if (f.underline && f.underline !== "none") wrap("u");
    if (f.strike) wrap("s");
    if (f.highlight) wrap("mark");
    if (f.vertical === "superscript") wrap("sup");
    if (f.vertical === "subscript") wrap("sub");
    if (styles.length) { open += `<span style="${styles.join(";")}">`; close = "</span>" + close; }
    if (!open) throw new DocError("没有可应用的格式", "invalid");
    const s = sg.map[ms[0].start - sg.start], e = sg.map[ms[0].end - sg.start];
    const before = this.display(b.text);
    this.apply([{ s: e, e, t: close }, { s, e: s, t: open }]);
    this.parse();
    return { changedRefs: [p.ref], summary: "已添加格式标签。", preview: [{ ref: p.ref, before, after: `${open}${clip(p.text ?? before, 60)}${close}` }], structural: false };
  }

  setParagraph(p: { refs: string[]; format: ParagraphFormat }, _o: EditOptions): EditResult {
    const tag = p.format.style ? (/^h[1-6]$|^p$|^li$|^blockquote$|^pre$/i.exec(p.format.style)?.[0] ?? (/(\d)$/.exec(p.format.style) ? `h${/(\d)$/.exec(p.format.style)![1]}` : undefined)) : undefined;
    if (p.format.style && !tag) throw new DocError("HTML 的 style 参数请使用标签名：p / h1–h6 / li / blockquote / pre", "invalid");
    const ops: Array<{ s: number; e: number; t: string }> = [];
    for (const ref of p.refs) {
      const b = this.blocks[this.resolve(ref)];
      const loc = b.el.sourceCodeLocation!;
      if (tag && loc.startTag && loc.endTag) {
        const st = this.src.slice(loc.startTag.startOffset, loc.startTag.endOffset);
        ops.push({ s: loc.startTag.startOffset, e: loc.startTag.endOffset, t: st.replace(/^<\w+/, `<${tag.toLowerCase()}`) });
        ops.push({ s: loc.endTag.startOffset, e: loc.endTag.endOffset, t: `</${tag.toLowerCase()}>` });
      }
      if (p.format.alignment && loc.startTag) {
        const align = p.format.alignment === "justify" ? "justify" : p.format.alignment;
        const st = tag ? null : this.src.slice(loc.startTag.startOffset, loc.startTag.endOffset);
        if (!st) throw new DocError("修改标签与对齐请分两次调用", "invalid");
        const next = /\sstyle="/i.test(st) ? st.replace(/\sstyle="([^"]*)"/i, (_m, v) => ` style="${v.replace(/text-align:[^;]*;?/i, "")}${v && !v.trim().endsWith(";") ? ";" : ""}text-align:${align}"`) : st.replace(/\/?>$/, (m) => ` style="text-align:${align}"${m}`);
        ops.push({ s: loc.startTag.startOffset, e: loc.startTag.endOffset, t: next });
      }
    }
    this.apply(ops);
    this.parse();
    return { changedRefs: p.refs, summary: "已调整段落标签/对齐。", preview: [], structural: false };
  }

  private indentBefore(offset: number): string {
    const lineStart = this.src.lastIndexOf("\n", offset - 1) + 1;
    const ws = /^[ \t]*/.exec(this.src.slice(lineStart, offset))?.[0] ?? "";
    return ws;
  }

  insertBlocks(p: { anchor: string; position: "before" | "after"; blocks: NewBlock[] }, _o: EditOptions): EditResult {
    const b = this.blocks[this.resolve(p.anchor)];
    const loc = b.el.sourceCodeLocation!;
    const indent = this.indentBefore(loc.startOffset);
    const eol = this.src.includes("\r\n") ? "\r\n" : "\n";
    const html = p.blocks.map((nb) => {
      const tpl = nb.like ? this.blocks[this.resolve(nb.like)].el : b.el;
      const tag = nb.style ? (/^h[1-6]$|^p$|^li$/i.exec(nb.style)?.[0] ?? tpl.tagName) : tpl.tagName;
      const tl = tpl.sourceCodeLocation!;
      const openSrc = tl.startTag ? this.src.slice(tl.startTag.startOffset, tl.startTag.endOffset) : `<${tag}>`;
      const open = tag.toLowerCase() === tpl.tagName ? openSrc.replace(/\sid="[^"]*"/i, "") : `<${tag.toLowerCase()}>`;
      return `${open}${this.esc(nb.text)}</${tag.toLowerCase()}>`;
    });
    const joined = html.join(eol + indent);
    if (p.position === "after") this.apply([{ s: loc.endOffset, e: loc.endOffset, t: eol + indent + joined }]);
    else this.apply([{ s: loc.startOffset, e: loc.startOffset, t: joined + eol + indent }]);
    this.structureVersion++;
    this.parse();
    return { changedRefs: [], summary: `已插入 ${p.blocks.length} 个块，请重新读取以获得新引用。`, preview: [{ ref: "", before: "", after: joined }], structural: true };
  }

  deleteBlocks(p: { refs: string[] }, _o: EditOptions): EditResult {
    const ops: Array<{ s: number; e: number; t: string }> = [];
    const preview: EditResult["preview"] = [];
    for (const ref of p.refs) {
      const b = this.blocks[this.resolve(ref)];
      const loc = b.el.sourceCodeLocation!;
      let s = loc.startOffset;
      const lineStart = this.src.lastIndexOf("\n", s - 1) + 1;
      if (/^[ \t]*$/.test(this.src.slice(lineStart, s))) s = lineStart;
      let e = loc.endOffset;
      const m = /^[ \t]*\r?\n/.exec(this.src.slice(e));
      if (m && s === lineStart) e += m[0].length;
      ops.push({ s, e, t: "" });
      preview.push({ ref, before: this.display(b.text), after: "" });
    }
    this.apply(ops);
    this.structureVersion++;
    this.parse();
    return { changedRefs: p.refs, summary: `已删除 ${p.refs.length} 个块。`, preview, structural: true };
  }

  /** 插入一段原始 HTML（表格、图片、SVG、公式等由工具层生成），沿用锚点所在行的缩进 */
  insertRaw(p: { anchor: string; position: "before" | "after"; markup: string; label?: string }, _o: EditOptions): EditResult {
    const b = this.blocks[this.resolve(p.anchor)];
    for (let e: any = b.el; e; e = e.parentNode) if (e.tagName === "head") throw new DocError(`块 ${p.anchor} 位于 <head> 中，内容只能插在正文（<body>）里的块前后。`, "invalid");
    const loc = b.el.sourceCodeLocation!;
    const indent = this.indentBefore(loc.startOffset);
    const eol = this.src.includes("\r\n") ? "\r\n" : "\n";
    const markup = p.markup.split(/\r?\n/).join(eol + indent);
    if (p.position === "after") this.apply([{ s: loc.endOffset, e: loc.endOffset, t: eol + indent + markup }]);
    else this.apply([{ s: loc.startOffset, e: loc.startOffset, t: markup + eol + indent }]);
    this.structureVersion++;
    this.parse();
    return { changedRefs: [], summary: `已插入${p.label ?? "内容"}，请重新读取以获得新引用。`, preview: [{ ref: "", before: "", after: clip(markup, 400) }], structural: true };
  }

  insertTableRow(): EditResult {
    throw new DocError("HTML 表格行请用 doc_insert_blocks 在某个单元格所在行附近插入，或直接编辑源码。", "unsupported");
  }

  addComment(p: { ref: string; text?: string; comment: string }, o: EditOptions): EditResult {
    const b = this.blocks[this.resolve(p.ref)];
    const loc = b.el.sourceCodeLocation!;
    const note = `<!-- 批注（${o.author}）：${p.comment.replace(/--/g, "—")} -->`;
    this.apply([{ s: loc.startOffset, e: loc.startOffset, t: note }]);
    this.parse();
    return { changedRefs: [p.ref], summary: "已以 HTML 注释形式添加批注。", preview: [{ ref: p.ref, before: "", after: note }], structural: false };
  }

  serialize(): Buffer {
    return Buffer.from(this.src, "utf8");
  }

  fidelity(original: Buffer): FidelityReport {
    const a = original.toString("utf8");
    let pre = 0;
    while (pre < a.length && pre < this.src.length && a[pre] === this.src[pre]) pre++;
    let suf = 0;
    while (suf < a.length - pre && suf < this.src.length - pre && a[a.length - 1 - suf] === this.src[this.src.length - 1 - suf]) suf++;
    const same = a === this.src;
    let problems: string[] = [];
    try { parse(this.src); } catch (e: any) { problems = [`HTML 解析失败：${e.message}`]; }
    return {
      ok: !problems.length, totalParts: 1, identicalParts: same ? 1 : 0,
      modifiedParts: same ? [] : [{ name: "源文件", changedBlocks: [`源码字符 ${pre}–${this.src.length - suf} 之间有改动，其余 ${pre + suf} 个字符逐字不变`] }],
      addedParts: [], problems,
    };
  }
}
