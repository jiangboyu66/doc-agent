import path from "node:path";
import { DocxDocument } from "./docx/DocxDocument.js";
import { MarkdownDocument } from "./markdown/MarkdownDocument.js";
import { HtmlDocument } from "./html/HtmlDocument.js";
import { PdfDocument } from "./pdf/PdfDocument.js";
import { DocError, type DocFormat, type DocumentAdapter } from "./types.js";

/** 旧版 Word（Word 97-2003 二进制 .doc / .dot）：上传时先转换为 .docx（见 services/docConvert），之后按 Word 文档编辑 */
export function isLegacyWord(filename: string): boolean {
  return /\.(doc|dot)$/i.test(filename);
}

export function detectFormat(filename: string): DocFormat {
  const ext = path.extname(filename).toLowerCase();
  if (ext === ".docx" || ext === ".dotx" || ext === ".docm") return "docx";
  if (ext === ".pdf") return "pdf";
  if (ext === ".html" || ext === ".htm" || ext === ".xhtml") return "html";
  if (ext === ".md" || ext === ".markdown" || ext === ".txt") return "markdown";
  if (ext === ".doc" || ext === ".dot") throw new DocError("旧版 .doc 需要先转换为 .docx 才能编辑（上传时会自动转换）。", "unsupported");
  throw new DocError(`不支持的文件类型：${ext || "(无扩展名)"}。支持 .docx / .doc / .md / .html / .pdf`, "unsupported");
}

export async function loadDocument(buf: Buffer, format: DocFormat): Promise<DocumentAdapter> {
  switch (format) {
    case "docx": return new DocxDocument(buf);
    case "markdown": return new MarkdownDocument(buf);
    case "html": return new HtmlDocument(buf);
    case "pdf": return PdfDocument.load(buf);
  }
}

export * from "./types.js";
