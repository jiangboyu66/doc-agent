/**
 * Markdown 原生适配器
 *
 * Markdown 文件本身就是"源格式"，这里所有修改都是对源文本的精确拼接（splice）：
 * 没被修改的字节——包括换行符风格（CRLF/LF）、缩进、行尾空格、front matter——全部原样保留。
 */

import { findAll, computeHunks, clip } from "../textMatch.js";
import {
  DocError, type DocumentAdapter, type DocumentCapabilities, type BlockInfo, type SearchHit, type DocumentSummary,
  type EditOptions, type EditResult, type RunFormat, type ParagraphFormat, type NewBlock, type FidelityReport,
} from "../types.js";

interface MdBlock {
  start: number; // 源文本字符偏移
  end: number;
  kind: BlockInfo["kind"];
  level?: number;
}

export class MarkdownDocument implements DocumentAdapter {
  readonly format = "markdown" as const;
  readonly capabilities: DocumentCapabilities = {
    replaceText: true, formatText: true, paragraphProps: true, insertBlocks: true, deleteBlocks: true,
    tables: false, comments: true, trackChanges: false, styles: false, rawInsert: true, links: true, move: true,
  };
  structureVersion = 0;
  private src: string;
  private eol: string;
  private blocks: MdBlock[] = [];

  constructor(buf: Buffer) {
    this.src = buf.toString("utf8");
    this.eol = this.src.includes("\r\n") ? "\r\n" : "\n";
    this.parse();
  }

  private parse(): void {
    const lines: Array<{ start: number; end: number; text: string }> = [];
    let pos = 0;
    const re = /\r?\n/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(this.src))) {
      lines.push({ start: pos, end: m.index, text: this.src.slice(pos, m.index) });
      pos = m.index + m[0].length;
    }
    if (pos <= this.src.length) lines.push({ start: pos, end: this.src.length, text: this.src.slice(pos) });

    const blocks: MdBlock[] = [];
    let i = 0;
    if (lines[0]?.text === "---") {
      let j = 1;
      while (j < lines.length && lines[j].text !== "---") j++;
      if (j < lines.length) {
        blocks.push({ start: lines[0].start, end: lines[j].end, kind: "other" });
        i = j + 1;
      }
    }
    while (i < lines.length) {
      const t = lines[i].text;
      if (!t.trim()) { i++; continue; }
      const fence = /^\s*(```|~~~)/.exec(t);
      if (fence) {
        let j = i + 1;
        while (j < lines.length && !lines[j].text.trim().startsWith(fence[1])) j++;
        const endLine = Math.min(j, lines.length - 1);
        blocks.push({ start: lines[i].start, end: lines[endLine].end, kind: "code" });
        i = endLine + 1;
        continue;
      }
      const h = /^(#{1,6})\s/.exec(t);
      if (h) { blocks.push({ start: lines[i].start, end: lines[i].end, kind: "heading", level: h[1].length }); i++; continue; }
      if (/^\s*([-*+]|\d+[.)])\s/.test(t)) { blocks.push({ start: lines[i].start, end: lines[i].end, kind: "list" }); i++; continue; }
      if (/^\s*\|/.test(t)) { blocks.push({ start: lines[i].start, end: lines[i].end, kind: "table-cell" }); i++; continue; }
      // 普通段落 / 引用：连续非空、非特殊行
      let j = i;
      while (
        j + 1 < lines.length && lines[j + 1].text.trim() &&
        !/^(#{1,6}\s|\s*([-*+]|\d+[.)])\s|\s*\||\s*(```|~~~))/.test(lines[j + 1].text)
      ) j++;
      blocks.push({ start: lines[i].start, end: lines[j].end, kind: /^\s*>/.test(t) ? "quote" : "paragraph" });
      i = j + 1;
    }
    this.blocks = blocks;
  }

  private ref(i: number): string {
    return `B${i + 1}@v${this.structureVersion}`;
  }

  private resolve(ref: string): number {
    const m = /^B(\d+)(?:@v(\d+))?$/.exec(ref.trim());
    if (!m) throw new DocError(`无效的块引用 ${ref}`, "not_found");
    if (m[2] !== undefined && Number(m[2]) !== this.structureVersion) {
      throw new DocError(`块引用 ${ref} 已过期（文档结构已变化，当前 v${this.structureVersion}），请重新读取。`, "stale_ref");
    }
    const i = Number(m[1]) - 1;
    if (!this.blocks[i]) throw new DocError(`找不到块 ${ref}`, "not_found");
    return i;
  }

  private text(b: MdBlock): string {
    return this.src.slice(b.start, b.end);
  }

  private info(i: number): BlockInfo {
    const b = this.blocks[i];
    return { ref: this.ref(i), location: "正文", kind: b.kind, level: b.level, text: this.text(b) };
  }

  summary(): DocumentSummary {
    const headings = this.blocks
      .map((b, i) => ({ b, i }))
      .filter(({ b }) => b.kind === "heading")
      .map(({ b, i }) => ({ ref: this.ref(i), level: b.level ?? 1, text: this.text(b).replace(/^#+\s*/, ""), location: "正文" }));
    return {
      format: "markdown", blockCount: this.blocks.length, charCount: this.src.replace(/\s/g, "").length, headings,
      parts: ["源文件"], tables: this.blocks.filter((b) => b.kind === "table-cell").length ? 1 : 0,
      images: (this.src.match(/!\[[^\]]*\]\(/g) ?? []).length, sections: 1, trackedChanges: 0, comments: (this.src.match(/<!--\s*批注/g) ?? []).length,
      notes: ["Markdown 的格式由标记语法表达（**粗体**、# 标题），修改格式即修改标记。"],
    };
  }

  listBlocks(): BlockInfo[] {
    return this.blocks.map((_, i) => this.info(i));
  }
  getBlock(ref: string): BlockInfo {
    return this.info(this.resolve(ref));
  }
  describeFormat(ref: string): string {
    const b = this.info(this.resolve(ref));
    return `块 ${b.ref}：类型 ${b.kind}${b.level ? `（${b.level} 级标题）` : ""}。Markdown 的格式直接体现在源文本标记中：\n${b.text}`;
  }
  listStyles() {
    return [];
  }

  search(query: string, opts: { regex?: boolean; caseSensitive?: boolean }): SearchHit[] {
    const hits: SearchHit[] = [];
    this.blocks.forEach((b, i) => {
      const t = this.text(b);
      if (opts.regex) {
        for (const m of t.matchAll(new RegExp(query, opts.caseSensitive ? "g" : "gi"))) {
          hits.push({ ref: this.ref(i), location: "正文", text: t, index: m.index ?? 0, match: m[0] });
        }
      } else {
        for (const m of findAll(t, query, opts.caseSensitive ?? false)) {
          hits.push({ ref: this.ref(i), location: "正文", text: t, index: m.start, match: t.slice(m.start, m.end) });
        }
      }
    });
    return hits;
  }

  private splice(start: number, end: number, text: string): void {
    this.src = this.src.slice(0, start) + text + this.src.slice(end);
  }

  private reparse(structural: boolean): void {
    if (structural) this.structureVersion++;
    this.parse();
  }

  replaceText(p: { ref?: string; oldText: string; newText: string; replaceAll?: boolean }, _o: EditOptions): EditResult {
    const scopeIdx = p.ref ? [this.resolve(p.ref)] : this.blocks.map((_, i) => i);
    const matches: Array<{ i: number; start: number; end: number }> = [];
    for (const i of scopeIdx) {
      const b = this.blocks[i];
      for (const m of findAll(this.text(b), p.oldText)) matches.push({ i, start: b.start + m.start, end: b.start + m.end });
    }
    if (!matches.length) throw new DocError(`找不到 old_text"${clip(p.oldText, 80)}"，请先用 doc_search 确认原文。`, "not_found");
    if (matches.length > 1 && !p.replaceAll) {
      throw new DocError(`old_text 出现了 ${matches.length} 次（${matches.slice(0, 6).map((m) => this.ref(m.i)).join("、")}），请指定 ref 或提供更长片段，或设置 replace_all=true。`, "ambiguous");
    }
    const before = new Map(matches.map((m) => [m.i, this.text(this.blocks[m.i])]));
    const lineCount = this.src.split(/\r?\n/).length;
    for (const m of [...matches].sort((a, b) => b.start - a.start)) {
      const actual = this.src.slice(m.start, m.end);
      const hunks = computeHunks(actual, p.newText.replace(/\r?\n/g, this.eol)).sort((a, b) => b.start - a.start);
      for (const h of hunks) this.splice(m.start + h.start, m.start + h.end, h.text);
    }
    const structural = this.src.split(/\r?\n/).length !== lineCount;
    const refsBefore = [...before.keys()].map((i) => this.ref(i));
    this.reparse(structural);
    return {
      changedRefs: refsBefore,
      summary: `已修改 ${matches.length} 处。${structural ? "行数发生变化，块引用已更新，请重新读取。" : ""}`,
      preview: [...before.entries()].map(([i, t]) => ({ ref: this.ref(i), before: t, after: structural ? "（结构已变化，请重新读取）" : this.text(this.blocks[i]) })),
      structural,
    };
  }

  formatText(p: { ref: string; text?: string; format: RunFormat }, _o: EditOptions): EditResult {
    const i = this.resolve(p.ref);
    const b = this.blocks[i];
    const t = this.text(b);
    const ms = p.text ? findAll(t, p.text) : [{ start: 0, end: t.length, fuzzy: false }];
    if (ms.length !== 1) throw new DocError(ms.length ? "目标文字不唯一，请提供更长片段" : "找不到目标文字", ms.length ? "ambiguous" : "not_found");
    let open = "", close = "";
    if (p.format.bold) { open += "**"; close = "**" + close; }
    if (p.format.italic) { open += "*"; close = "*" + close; }
    if (p.format.strike) { open += "~~"; close = "~~" + close; }
    if (!open) throw new DocError("Markdown 只支持加粗、斜体、删除线；颜色/字号等需要 HTML 标记，可直接用 doc_replace_text 写入。", "unsupported");
    const s = b.start + ms[0].start, e = b.start + ms[0].end;
    const beforeText = t;
    this.splice(e, e, close);
    this.splice(s, s, open);
    this.reparse(false);
    return { changedRefs: [this.ref(i)], summary: "已添加 Markdown 格式标记。", preview: [{ ref: this.ref(i), before: beforeText, after: this.text(this.blocks[i]) }], structural: false };
  }

  setParagraph(p: { refs: string[]; format: ParagraphFormat }, _o: EditOptions): EditResult {
    const lvl = /^(heading|标题)\s*(\d)$/i.exec(p.format.style ?? "")?.[2] ?? (/^h(\d)$/i.exec(p.format.style ?? "")?.[1]);
    const toPara = /^(normal|正文|paragraph|p)$/i.test(p.format.style ?? "");
    if (!lvl && !toPara) throw new DocError("Markdown 只支持把块改为标题（style=\"Heading 2\" 或 \"h2\"）或正文（style=\"正文\"）。", "unsupported");
    const idxs = p.refs.map((r) => this.resolve(r)).sort((a, b) => b - a);
    const preview: EditResult["preview"] = [];
    for (const i of idxs) {
      const b = this.blocks[i];
      const t = this.text(b);
      const body = t.replace(/^#{1,6}\s+/, "");
      const next = lvl ? `${"#".repeat(Number(lvl))} ${body}` : body;
      this.splice(b.start, b.end, next);
      preview.push({ ref: this.ref(i), before: t, after: next });
    }
    this.reparse(false);
    return { changedRefs: p.refs, summary: `已调整 ${p.refs.length} 个块的层级。`, preview, structural: false };
  }

  insertBlocks(p: { anchor: string; position: "before" | "after"; blocks: NewBlock[] }, _o: EditOptions): EditResult {
    const i = this.resolve(p.anchor);
    const b = this.blocks[i];
    const render = (nb: NewBlock) => {
      const lvl = /(\d)$/.exec(nb.style ?? "")?.[1];
      if (nb.style && /heading|标题|^h\d/i.test(nb.style) && lvl) return `${"#".repeat(Number(lvl))} ${nb.text}`;
      if (nb.style && /list|列表/i.test(nb.style)) return `- ${nb.text}`;
      return nb.text;
    };
    const joiner = b.kind === "list" && p.blocks.every((x) => /list|列表/i.test(x.style ?? "")) ? this.eol : this.eol + this.eol;
    const content = p.blocks.map(render).join(joiner);
    if (p.position === "after") this.splice(b.end, b.end, joiner + content);
    else this.splice(b.start, b.start, content + joiner);
    this.reparse(true);
    return { changedRefs: [], summary: `已插入 ${p.blocks.length} 个块。结构已变化，请重新读取以获得新引用。`, preview: [{ ref: "", before: "", after: content }], structural: true };
  }

  deleteBlocks(p: { refs: string[] }, _o: EditOptions): EditResult {
    const idxs = p.refs.map((r) => this.resolve(r)).sort((a, b) => b - a);
    const preview: EditResult["preview"] = [];
    for (const i of idxs) {
      const b = this.blocks[i];
      preview.push({ ref: this.ref(i), before: this.text(b), after: "" });
      let end = b.end;
      const rest = this.src.slice(end);
      const m = /^(\r?\n)+/.exec(rest);
      if (m) end += m[0].length;
      this.splice(b.start, end, "");
    }
    this.reparse(true);
    return { changedRefs: p.refs, summary: `已删除 ${idxs.length} 个块。`, preview, structural: true };
  }

  /** 插入一段原始 Markdown（表格、图片、公式、SVG 等由工具层生成） */
  insertRaw(p: { anchor: string; position: "before" | "after"; markup: string; label?: string }, _o: EditOptions): EditResult {
    const b = this.blocks[this.resolve(p.anchor)];
    const content = p.markup.replace(/\r?\n/g, this.eol);
    if (p.position === "after") this.splice(b.end, b.end, this.eol + this.eol + content);
    else this.splice(b.start, b.start, content + this.eol + this.eol);
    this.reparse(true);
    return { changedRefs: [], summary: `已插入${p.label ?? "内容"}。结构已变化，请重新读取以获得新引用。`, preview: [{ ref: "", before: "", after: clip(content, 400) }], structural: true };
  }

  /** Markdown 超链接：把文字改为 [文字](地址) */
  insertLink(p: { ref: string; text: string; url: string }, o: EditOptions): EditResult {
    return this.replaceText({ ref: p.ref, oldText: p.text, newText: `[${p.text}](${p.url})` }, o);
  }

  moveBlocks(p: { refs: string[]; anchor: string; position: "before" | "after" }, _o: EditOptions): EditResult {
    const idx = [...new Set(p.refs.map((r) => this.resolve(r)))].sort((a, b) => a - b);
    const target = this.resolve(p.anchor);
    if (idx.includes(target)) throw new DocError("锚点不能是被移动的块之一", "invalid");
    const texts = idx.map((i) => this.text(this.blocks[i]));
    const t = this.blocks[target];
    const insertAt = p.position === "after" ? t.end : t.start;
    // 先在目标位置插入，再从后往前删除原块（偏移量从后往前处理不会互相影响）
    const ops: Array<{ s: number; e: number; t: string }> = [];
    const joined = texts.join(this.eol + this.eol);
    ops.push({ s: insertAt, e: insertAt, t: p.position === "after" ? this.eol + this.eol + joined : joined + this.eol + this.eol });
    for (const i of idx) {
      const b = this.blocks[i];
      let e = b.end;
      const m = /^(\r?\n)+/.exec(this.src.slice(e));
      if (m) e += m[0].length;
      ops.push({ s: b.start, e, t: "" });
    }
    ops.sort((a, b) => b.s - a.s || b.e - a.e);
    for (const op of ops) this.splice(op.s, op.e, op.t);
    this.reparse(true);
    return { changedRefs: [], summary: `已移动 ${idx.length} 个块。结构已变化，请重新读取。`, preview: texts.map((x) => ({ ref: "", before: "（原位置）", after: clip(x, 80) })), structural: true };
  }

  insertTableRow(): EditResult {
    throw new DocError("Markdown 表格行请用 doc_insert_blocks 插入一行 | a | b | 形式的文本。", "unsupported");
  }

  addComment(p: { ref: string; text?: string; comment: string }, o: EditOptions): EditResult {
    const i = this.resolve(p.ref);
    const b = this.blocks[i];
    const note = `<!-- 批注（${o.author}）：${p.comment.replace(/--/g, "—")} -->`;
    this.splice(b.start, b.start, note + this.eol);
    this.reparse(true);
    return { changedRefs: [p.ref], summary: "已以 HTML 注释形式添加批注（渲染时不可见）。", preview: [{ ref: p.ref, before: "", after: note }], structural: true };
  }

  serialize(): Buffer {
    return Buffer.from(this.src, "utf8");
  }

  fidelity(original: Buffer): FidelityReport {
    const a = original.toString("utf8").split(/\r?\n/);
    const b = this.src.split(/\r?\n/);
    let prefix = 0;
    while (prefix < a.length && prefix < b.length && a[prefix] === b[prefix]) prefix++;
    let suffix = 0;
    while (suffix < a.length - prefix && suffix < b.length - prefix && a[a.length - 1 - suffix] === b[b.length - 1 - suffix]) suffix++;
    const same = original.equals(this.serialize());
    return {
      ok: true, totalParts: 1, identicalParts: same ? 1 : 0,
      modifiedParts: same ? [] : [{ name: "源文件", changedBlocks: [`第 ${prefix + 1}–${b.length - suffix} 行（其余 ${prefix + suffix} 行逐字节未变）`] }],
      addedParts: [], problems: [],
    };
  }
}
