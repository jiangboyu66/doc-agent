/**
 * PDF 只读适配器
 *
 * PDF 是"最终版式"格式：文字以绝对坐标的字形绘制，没有段落与样式结构。任何工具都无法在保持
 * 版式 100% 不变的前提下改写 PDF 中的文字（换一个字就可能需要重排整行甚至整页）。
 * 因此这里如实地只提供读取、搜索与审阅能力；写类工具在 PDF 上不启用（isEnabled=false），
 * 模型看不到这些工具，也就不会尝试做做不到的事。需要编辑时，应让用户提供原始 Word 文件。
 */

import {
  DocError, type DocumentAdapter, type DocumentCapabilities, type BlockInfo, type SearchHit, type DocumentSummary,
  type EditResult, type FidelityReport,
} from "../types.js";
import { findAll } from "../textMatch.js";

interface PBlock { page: number; text: string }

export class PdfDocument implements DocumentAdapter {
  readonly format = "pdf" as const;
  readonly capabilities: DocumentCapabilities = {
    replaceText: false, formatText: false, paragraphProps: false, insertBlocks: false, deleteBlocks: false,
    tables: false, comments: false, trackChanges: false, styles: false,
  };
  structureVersion = 0;
  private constructor(private readonly buf: Buffer, private readonly blocks: PBlock[], private readonly pages: number) {}

  static async load(buf: Buffer): Promise<PdfDocument> {
    const pdfjs: any = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const doc = await pdfjs.getDocument({ data: new Uint8Array(buf), isEvalSupported: false, useSystemFonts: false }).promise;
    const blocks: PBlock[] = [];
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const tc = await page.getTextContent();
      const lines: Array<{ y: number; h: number; text: string }> = [];
      let cur: { y: number; h: number; text: string } | null = null;
      for (const it of tc.items as Array<{ str: string; transform: number[]; height: number; hasEOL: boolean }>) {
        const y = it.transform[5];
        if (!cur || Math.abs(cur.y - y) > Math.max(2, (it.height || cur.h) * 0.5)) {
          if (cur && cur.text.trim()) lines.push(cur);
          cur = { y, h: it.height || 10, text: "" };
        }
        cur.text += it.str;
        if (it.hasEOL) {
          if (cur.text.trim()) lines.push(cur);
          cur = null;
        }
      }
      if (cur && cur.text.trim()) lines.push(cur);
      let para = "";
      let lastY: number | null = null;
      let lastH = 10;
      for (const l of lines) {
        const gap = lastY === null ? 0 : lastY - l.y;
        if (para && (gap > lastH * 1.8 || gap < 0)) {
          blocks.push({ page: n, text: para.trim() });
          para = "";
        }
        para += (para && !para.endsWith("-") ? " " : "") + l.text.trim();
        lastY = l.y;
        lastH = l.h;
      }
      if (para.trim()) blocks.push({ page: n, text: para.trim() });
    }
    return new PdfDocument(buf, blocks, doc.numPages);
  }

  private ref(i: number) { return `D${i + 1}`; }
  private resolve(ref: string): number {
    const m = /^D(\d+)$/.exec(ref.trim());
    const i = m ? Number(m[1]) - 1 : -1;
    if (!this.blocks[i]) throw new DocError(`找不到块 ${ref}`, "not_found");
    return i;
  }
  private info(i: number): BlockInfo {
    const b = this.blocks[i];
    return { ref: this.ref(i), location: `第 ${b.page} 页`, kind: "paragraph", text: b.text };
  }

  summary(): DocumentSummary {
    return {
      format: "pdf", blockCount: this.blocks.length, charCount: this.blocks.reduce((s, b) => s + b.text.replace(/\s/g, "").length, 0),
      headings: [], parts: [`${this.pages} 页`], tables: 0, images: 0, sections: this.pages, trackedChanges: 0, comments: 0,
      notes: [
        "PDF 为最终版式格式，本系统对 PDF 只提供读取、搜索与审阅，不提供改写（任何改写都无法保证版式不变）。",
        this.blocks.length === 0 ? "未提取到文字：可能是扫描件（图片型 PDF），需要 OCR。" : "段落由文字坐标推断，可能与视觉分段略有差异。",
      ],
    };
  }
  listBlocks(): BlockInfo[] { return this.blocks.map((_, i) => this.info(i)); }
  getBlock(ref: string): BlockInfo { return this.info(this.resolve(ref)); }
  describeFormat(ref: string): string { return `${ref}：PDF 不含可编辑的样式信息。`; }
  listStyles() { return []; }
  search(query: string, opts: { regex?: boolean; caseSensitive?: boolean }): SearchHit[] {
    const hits: SearchHit[] = [];
    this.blocks.forEach((b, i) => {
      const ms = opts.regex
        ? [...b.text.matchAll(new RegExp(query, opts.caseSensitive ? "g" : "gi"))].map((m) => ({ start: m.index ?? 0, end: (m.index ?? 0) + m[0].length }))
        : findAll(b.text, query, opts.caseSensitive ?? false);
      for (const m of ms) hits.push({ ref: this.ref(i), location: `第 ${b.page} 页`, text: b.text, index: m.start, match: b.text.slice(m.start, m.end) });
    });
    return hits;
  }
  private ro(): never {
    throw new DocError("PDF 是只读格式，无法在保持版式不变的前提下修改文字。请提供原始 Word/Markdown/HTML 文件进行编辑。", "unsupported");
  }
  replaceText(): EditResult { return this.ro(); }
  formatText(): EditResult { return this.ro(); }
  setParagraph(): EditResult { return this.ro(); }
  insertBlocks(): EditResult { return this.ro(); }
  deleteBlocks(): EditResult { return this.ro(); }
  insertTableRow(): EditResult { return this.ro(); }
  addComment(): EditResult { return this.ro(); }
  serialize(): Buffer { return this.buf; }
  fidelity(): FidelityReport {
    return { ok: true, totalParts: 1, identicalParts: 1, modifiedParts: [], addedParts: [], problems: [] };
  }
}
