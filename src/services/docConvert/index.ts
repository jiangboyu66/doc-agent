/**
 * 旧版 Word 文档（.doc，Word 97-2003 二进制格式）→ .docx。
 *
 * 两个引擎：
 *   builtin      内置的纯 Node 解析器（parse.ts + writer.ts）：不依赖任何外部程序，覆盖正文、表格、样式、编号、页眉页脚、
 *                脚注尾注、批注、修订、域、书签、图文框、兼容性选项、行内与浮动图片、文本框、自选图形与形状组合；
 *                与 Word 自己"另存为 .docx"的结果逐页核对，分页与换行位置一致。
 *   libreoffice  LibreOffice 的 Word 二进制格式过滤器：能读 Word 6.0 / 95 和艺术字，但图文框、兼容性选项等会丢，版面走样。
 * 默认 auto：内置引擎优先，读不了的（Word 6.0 / 95、损坏）交给 LibreOffice。也可用环境变量 DOC_ENGINE=builtin|libreoffice 固定。
 */

import { DocFile, DocFileError } from "./parse.js";
import { writeDocx, visibleChars, type DocConvertReport } from "./writer.js";
import { ZipPackage } from "../../documents/docx/zip.js";

export { DocFileError, type DocConvertReport };

export type DocEngine = "auto" | "builtin" | "libreoffice";

let officeAvailable: Promise<boolean> | undefined;
async function hasOffice(): Promise<boolean> {
  officeAvailable ??= import("../convert.js").then((m) => m.checkExternalTools()).then((t) => !!t.soffice).catch(() => false);
  return officeAvailable;
}

/** 内置引擎 */
export function docToDocxBuiltin(buf: Buffer): { data: Buffer; report: DocConvertReport } {
  try {
    return writeDocx(new DocFile(buf));
  } catch (e) {
    if (e instanceof DocFileError) throw e;
    // 结构异常的文件（越界、链表损坏）：统一报"已损坏"，不把内部错误抛给用户
    throw new DocFileError(`Word 文档已损坏或结构异常，内置引擎无法转换（${(e as Error).message.slice(0, 80)}）。可以在 Word 中打开后另存为 .docx 再上传。`, "corrupt");
  }
}

/** 文件头是否是 OLE2 复合文档（.doc 的容器） */
export function isOle2(buf: Buffer): boolean {
  return buf.length >= 8 && buf.readUInt32LE(0) === 0xe011cfd0 && buf.readUInt32LE(4) === 0xe11ab1a1;
}

export async function docToDocx(buf: Buffer, opts: { engine?: DocEngine } = {}): Promise<{ data: Buffer; report: DocConvertReport }> {
  const engine = opts.engine ?? ((process.env.DOC_ENGINE as DocEngine | undefined) || "auto");
  // 有些 .doc 其实是改了扩展名的 RTF、HTML 或 .docx：交给能读它的引擎
  if (!isOle2(buf)) {
    if (buf.subarray(0, 2).toString("latin1") === "PK") throw new DocFileError("这个文件其实是 .docx（扩展名写成了 .doc），请把扩展名改为 .docx 后再上传。", "unsupported");
    if (engine !== "builtin" && (await hasOffice())) return viaOffice(buf, "这个 .doc 文件其实是 RTF / 网页等其他格式，已由 LibreOffice 转换。");
    throw new DocFileError("不是 Word 97-2003 文档（文件可能是改了扩展名的 RTF、网页或已损坏）。", "not-ole");
  }
  if (engine === "libreoffice") return viaOffice(buf);
  // 内置引擎优先：与 Word 自己"另存为 .docx"的结果逐页核对过（分页、换行位置一致）；
  // LibreOffice 的转换会丢掉图文框、兼容性选项等，版面明显走样。内置引擎读不了的（Word 6.0 / 95、损坏）再交给 LibreOffice。
  try {
    return docToDocxBuiltin(buf);
  } catch (e) {
    if (engine === "builtin") throw e;
    if (e instanceof DocFileError && (e.code === "old-version" || e.code === "corrupt") && (await hasOffice())) {
      try {
        const r = await viaOffice(buf);
        r.report.notes.unshift(e.code === "old-version" ? "这是 Word 6.0 / 95 格式，已由 LibreOffice 转换。" : `内置引擎读取失败（${e.message.replace(/[。.]$/, "")}），已由 LibreOffice 转换。`);
        return r;
      } catch { /* LibreOffice 也不行：给出内置引擎的说明 */ }
    }
    if (e instanceof DocFileError && e.code === "old-version") throw new DocFileError("这是 Word 6.0 / 95 或更早版本的文档：内置引擎只支持 Word 97-2003 格式。请安装 LibreOffice（服务器上）后重新上传，或在 Word 中另存为 .docx。", "old-version");
    if (e instanceof DocFileError && e.code === "corrupt" && !e.message.includes("另存")) throw new DocFileError(`${e.message.replace(/[。.]$/, "")}。可以在 Word 中打开后另存为 .docx 再上传。`, "corrupt");
    throw e;
  }
}

async function viaOffice(buf: Buffer, note?: string): Promise<{ data: Buffer; report: DocConvertReport }> {
  const { convertViaOffice } = await import("../convert.js");
  const data = await convertViaOffice(buf, "doc", "docx", "MS Word 2007 XML");
  let chars = 0;
  try { const d = new DocFile(buf); chars = visibleChars(d.text.slice(0, d.ccp.text)); } catch { /* 旧版本或其他格式：不核对 */ }
  return {
    data,
    report: { engine: "libreoffice", paragraphs: 0, tables: 0, images: 0, floating: 0, textBoxes: 0, shapes: 0, footnotes: 0, endnotes: 0, comments: 0, revisions: 0, sections: 0, chars, written: 0, notes: note ? [note] : [] },
  };
}

/** 转换报告的文字说明（显示在会话开头） */
export function describeDocReport(r: DocConvertReport, coverage?: number): string {
  const parts: string[] = [];
  if (r.engine === "libreoffice") parts.push("原文件是 Word 97-2003 格式（.doc），已由 LibreOffice 转换为 .docx");
  else {
    const what = [
      `${r.paragraphs} 个段落`, r.tables ? `${r.tables} 个表格` : "", r.images ? `${r.images} 张图片${r.floating ? `（其中浮动 ${r.floating} 张）` : ""}` : "",
      r.textBoxes ? `${r.textBoxes} 个文本框` : "", r.shapes ? `${r.shapes} 个图形` : "", r.footnotes ? `${r.footnotes} 个脚注` : "", r.endnotes ? `${r.endnotes} 个尾注` : "",
      r.comments ? `${r.comments} 条批注` : "", r.revisions ? `${r.revisions} 处修订` : "", r.sections > 1 ? `${r.sections} 节` : "",
    ].filter(Boolean).join("、");
    parts.push(`原文件是 Word 97-2003 格式（.doc），已由内置引擎转换为 .docx（${what}）`);
  }
  parts[0] += "，之后按 Word（.docx）编辑与下载；上传的 .doc 原样保留，可在「下载 → 原始文件」取回，也可导出为 .doc。";
  if (coverage !== undefined) parts.push(`文字完整性 ${(100 * coverage).toFixed(1)}%。`);
  if (r.notes.length) parts.push(`注意：${r.notes.map((n) => n.replace(/[。.]$/, "")).join("；")}。`);
  return parts.join("");
}

/** .docx 正文里看得见的字符数（不含域代码、文本框、空白），用于核对 LibreOffice 转换的文字完整性 */
export function docxVisibleChars(docx: Buffer): number {
  const xml = new ZipPackage(docx).readText("word/document.xml")
    .replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, "")
    .replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, "");
  // 符号字体的字符（Symbol / Wingdings）写成 <w:sym>：各算一个字
  let n = (xml.match(/<w:sym /g) ?? []).length;
  for (const m of xml.matchAll(/<w:(?:t|delText)(?: [^>]*)?>([^<]*)<\/w:(?:t|delText)>/g)) {
    const t = m[1].replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-f]+);/gi, "x");
    n += visibleChars(t);
  }
  return n;
}

/**
 * 上传的 .doc：转换为 .docx，返回新文件名、引擎、给用户看的报告与文字完整性（0–1）。
 * 文字完整性 = 写进 .docx 正文的可见字符 / .doc 正文的可见字符（不含域代码、空白）。
 */
export async function convertUploadedDoc(buf: Buffer, filename: string, opts: { engine?: DocEngine } = {}): Promise<{ data: Buffer; filename: string; engine: string; report: string; coverage: number; detail?: DocConvertReport }> {
  const docxName = filename.replace(/\.(doc|dot)$/i, "") + ".docx";
  // 其实是 .docx（只是扩展名写成了 .doc）：直接按 Word 文档打开
  if (buf.subarray(0, 4).toString("latin1") === "PK\x03\x04") {
    return { data: buf, filename: docxName, engine: "none", report: "这个文件其实是 .docx（扩展名写成了 .doc），已直接按 Word（.docx）打开，内容没有经过转换。", coverage: 1 };
  }
  const { data, report } = await docToDocx(buf, opts);
  let coverage = 1;
  if (report.chars) {
    const written = report.engine === "libreoffice" ? docxVisibleChars(data) : report.written;
    coverage = Math.min(1, written / report.chars);
  }
  return {
    data,
    filename: docxName,
    engine: report.engine,
    report: describeDocReport(report, report.chars ? coverage : undefined),
    coverage,
    detail: report,
  };
}
