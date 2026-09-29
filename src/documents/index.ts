import path from "node:path";
import { DocxDocument } from "./docx/DocxDocument.js";
import { MarkdownDocument } from "./markdown/MarkdownDocument.js";
import { HtmlDocument } from "./html/HtmlDocument.js";
import { PdfDocument } from "./pdf/PdfDocument.js";
import { DocError, type DocFormat, type DocumentAdapter } from "./types.js";

export function detectFormat(filename: string): DocFormat {
  const ext = path.extname(filename).toLowerCase();
  if (ext === ".docx" || ext === ".dotx" || ext === ".docm") return "docx";
  if (ext === ".pdf") return "pdf";
  if (ext === ".html" || ext === ".htm" || ext === ".xhtml") return "html";
  if (ext === ".md" || ext === ".markdown" || ext === ".txt") return "markdown";
  if (ext === ".doc") throw new DocError("旧版 .doc（Word 97-2003）是二进制格式，请先在 Word 中另存为 .docx 再上传，以保证编辑后格式完全不变。", "unsupported");
  throw new DocError(`不支持的文件类型：${ext || "(无扩展名)"}。支持 .docx / .md / .html / .pdf`, "unsupported");
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
