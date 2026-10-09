/**
 * PDF → Word 转换服务
 *
 * 三个引擎：
 *   - builtin（默认）：内置引擎，纯 Node 实现，不依赖外部程序；按绝对位置还原段落、表格、分栏、图片；
 *   - pdf2docx：Python 库（需 pip install pdf2docx），适合结构复杂但无网格线的表格；
 *   - libreoffice：LibreOffice 的 PDF 导入，每行文字为独立文本框，版面接近但不便编辑。
 *
 * 每次转换都会做"文字完整性校验"：逐字符比对 PDF 中的文字与生成的 Word 中的文字，报告缺失的字符。
 */

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { extractPdf, type ImageBox, type PageModel } from "./extract.js";
import { buildDocModel, type Block, type DocModel, type LayoutMode, type Para } from "./layout.js";
import { writeDocx } from "./docxWriter.js";
import { rasterAvailable, renderRegions } from "./render.js";
import { DocxDocument } from "../../documents/docx/DocxDocument.js";

const execFileAsync = promisify(execFile);

export type PdfEngine = "builtin" | "pdf2docx" | "libreoffice";

export interface ConversionReport {
  engine: PdfEngine;
  pages: number;
  paragraphs: number;
  tables: number;
  images: number;
  columnPages: number;
  /** 按原样渲染为图片的矢量图 / 行间公式数量 */
  figures: number;
  equations: number;
  /** PDF 中非空白字符总数 */
  textTotal: number;
  /** 生成的 Word 中找回的比例（0–1） */
  textCoverage: number;
  /** 缺失字符示例 */
  missing: string;
  notes: string[];
  ms: number;
}

const ENGINE_LABEL: Record<PdfEngine, string> = { builtin: "内置引擎", pdf2docx: "pdf2docx", libreoffice: "LibreOffice" };

function charCounts(s: string): Map<string, number> {
  const m = new Map<string, number>();
  // 控制字符不是文字（Type 3 字体的乱码编码）：不计
  for (const ch of s) if (ch.trim() && !/[\u0000-\u001f\u007f-\u009f]/.test(ch)) m.set(ch, (m.get(ch) ?? 0) + 1);
  return m;
}

/** 生成的 docx 中的全部文字（正文、表格、文本框、页眉页脚、脚注） */
function docxText(buf: Buffer): string {
  const d = new DocxDocument(buf);
  return d.listBlocks().map((b) => b.text.replace(/⟨[^⟩]*⟩/g, "")).join("\n");
}

function verify(pages: PageModel[], docx: Buffer, rasterText = "") {
  const src = charCounts(pages.flatMap((p) => p.spans.filter((s) => !s.img).map((s) => s.text)).join(""));
  // 渲染为图片的区域：文字保存在替代文字中，计为已保留
  const out = charCounts(docxText(docx) + rasterText);
  // 并入正文字母的重音符号变成了组合字符（F + ˆ → F̂）：按原来的重音符号计数
  const ACC: Array<[string, string[]]> = [["\u0302", ["ˆ", "^"]], ["\u0303", ["˜", "~"]], ["\u0304", ["¯"]], ["\u0307", ["˙"]], ["\u0308", ["¨"]], ["\u0301", ["´"]], ["\u0300", ["`"]], ["\u030C", ["ˇ"]], ["\u20D7", ["→", "⃗"]]];
  for (const [comb, chars] of ACC) {
    let n = out.get(comb) ?? 0;
    for (const ch of chars) { if (!n) break; const need = Math.max(0, (src.get(ch) ?? 0) - (out.get(ch) ?? 0)); const k = Math.min(n, need); out.set(ch, (out.get(ch) ?? 0) + k); n -= k; }
  }
  let total = 0, lost = 0;
  const missing: string[] = [];
  for (const [ch, n] of src) {
    total += n;
    const got = out.get(ch) ?? 0;
    if (got < n) {
      lost += n - got;
      if (missing.length < 30) missing.push(ch);
    }
  }
  return { total, coverage: total ? 1 - lost / total : 1, missing: missing.join("") };
}

// ---------------------------------------------------------------------------

/** 文档模型中所有待渲染的图片区域（含表格单元格、文本框、页眉页脚中的段落，以及段内的行内公式图片） */
function renderBoxes(model: DocModel): ImageBox[] {
  const out: ImageBox[] = [];
  const seen = new Set<ImageBox>();
  const add = (im?: ImageBox) => { if (im?.render && !seen.has(im)) { seen.add(im); out.push(im); } };
  const para = (b: Para) => {
    for (const im of [...(b.images ?? []), ...(b.inlineImages ?? []), ...(b.hfImages ?? []).map((f) => f.img), ...(b.floats ?? []).map((f) => f.img)]) add(im);
    for (const r of b.runs) add(r.img);
    for (const t of b.textBoxes ?? []) para(t.para);
  };
  const block = (b: Block) => {
    if (b.kind === "p") para(b);
    else for (const row of b.rows) for (const c of row.cells) for (const x of c.blocks) block(x as Block);
  };
  model.blocks.forEach(block);
  for (const set of [model.headers, model.footers]) for (const k of ["default", "first", "even"] as const) for (const p of (set as any)?.[k] ?? []) para(p);
  return out;
}

async function runBuiltin(buf: Buffer, title: string, pages: PageModel[], mode: LayoutMode) {
  const notes: string[] = [];
  let model = buildDocModel(pages, { mode, raster: await rasterAvailable() });
  const regions = renderBoxes(model);
  if (regions.length) {
    const ok = await renderRegions(buf, regions).catch(() => 0);
    if (ok < regions.length) {
      // 渲染失败：退回纯文字重建，保证内容不丢
      model = buildDocModel(pages, { mode, raster: false });
      notes.push("矢量图 / 公式渲染失败，已按文字方式重建（公式与流程图的版式可能有偏差）。");
    } else {
      const parts = [model.stats.figures ? `${model.stats.figures} 幅矢量图（流程图、曲线图等）` : "", model.stats.equations ? `${model.stats.equations} 个行间公式` : ""].filter(Boolean);
      if (!parts.length && !regions.some((r) => r.render!.kind === "dropcap")) parts.push(`${regions.length} 处图形区域`);
      const dc = regions.filter((r) => r.render!.kind === "dropcap").length;
      if (dc) parts.push(`${dc} 个首字下沉`);
      notes.push(`${parts.join("、")}按原样渲染为 288 dpi 高清图片（透明背景）贴回原位；其中的文字保存在图片的替代文字里。`);
    }
  }
  if (model.stats.inlineMath) notes.push(`${model.stats.inlineMath} 处没有文字编码的行内公式 / 符号（Type 3 位图字体）以图片形式插回原文行内。`);
  const data = writeDocx(model, title);
  if (model.stats.vectorShapes) notes.push(`${model.stats.vectorShapes} 处零散的矢量图形（不构成完整的图）未能转换。`);
  if (model.stats.rotatedText) notes.push(`${model.stats.rotatedText} 段旋转文字未转换。`);
  if (model.stats.columnPages) notes.push(`${model.stats.columnPages} 页检测到双栏排版，已用 Word 分栏还原。`);
  return { data, stats: model.stats, notes, rasterText: model.rasterText };
}

async function findPython(): Promise<string[] | null> {
  const candidates: string[][] = [];
  if (process.env.PYTHON_PATH) candidates.push([process.env.PYTHON_PATH]);
  candidates.push(["python"], ["python3"]);
  if (process.platform === "win32") candidates.push(["py", "-3"]);
  for (const c of candidates) {
    try {
      await execFileAsync(c[0], [...c.slice(1), "-c", "import pdf2docx"], { timeout: 20_000, windowsHide: true });
      return c;
    } catch { /* 下一个 */ }
  }
  return null;
}

async function runPdf2docx(buf: Buffer): Promise<Buffer> {
  const py = await findPython();
  if (!py) throw new Error("未找到可用的 pdf2docx。请先安装 Python，再运行 pip install pdf2docx；或在 .env 中用 PYTHON_PATH 指定 Python 路径。");
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "doc-agent-pdf-"));
  try {
    const src = path.join(dir, "in.pdf"), out = path.join(dir, "out.docx");
    await fs.writeFile(src, buf);
    const script = "import sys,logging\nlogging.disable(logging.CRITICAL)\nfrom pdf2docx import Converter\ncv=Converter(sys.argv[1])\ncv.convert(sys.argv[2])\ncv.close()";
    await execFileAsync(py[0], [...py.slice(1), "-c", script, src, out], { timeout: 10 * 60_000, windowsHide: true, maxBuffer: 64 * 1024 * 1024 });
    return await fs.readFile(out);
  } finally {
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

async function runLibreOffice(buf: Buffer): Promise<Buffer> {
  const soffice = process.env.SOFFICE_PATH || "soffice";
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "doc-agent-pdf-"));
  try {
    const src = path.join(dir, "in.pdf");
    await fs.writeFile(src, buf);
    await execFileAsync(soffice, [
      "--headless", "--norestore", "--nolockcheck", `-env:UserInstallation=${pathToFileURL(path.join(dir, "profile")).href}`,
      "--infilter=writer_pdf_import", "--convert-to", "docx:MS Word 2007 XML", "--outdir", dir, src,
    ], { timeout: Number(process.env.SOFFICE_TIMEOUT_MS || 120_000), windowsHide: true, killSignal: "SIGKILL" });
    return await fs.readFile(path.join(dir, "in.docx"));
  } catch (e: any) {
    if (e?.code === "ENOENT") throw new Error("未找到 LibreOffice（soffice）。请安装后重试，或在 .env 中设置 SOFFICE_PATH。");
    throw new Error(`LibreOffice 转换失败：${(e?.stderr || e?.message || "").toString().slice(0, 300)}`);
  } finally {
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

export async function pdfToDocx(buf: Buffer, opts: { engine?: PdfEngine; title?: string; mode?: LayoutMode } = {}): Promise<{ data: Buffer; report: ConversionReport }> {
  const engine = opts.engine ?? "builtin";
  const t0 = Date.now();
  const pages = await extractPdf(buf);
  const textChars = pages.reduce((a, p) => a + p.spans.reduce((b, s) => b + s.text.replace(/\s/g, "").length, 0), 0);
  const notes: string[] = [];
  if (pages.length && textChars < 20 * pages.length && pages.some((p) => p.images.length)) {
    notes.push("这份 PDF 几乎没有可提取的文字，可能是扫描件（整页图片）。转换结果只包含图片，文字需要先做 OCR 识别。");
  }

  let data: Buffer;
  let rasterText = "";
  let stats = { pages: pages.length, paragraphs: 0, tables: 0, images: 0, columnPages: 0, figures: 0, equations: 0 };
  if (engine === "builtin") {
    const r = await runBuiltin(buf, opts.title ?? "文档", pages, opts.mode ?? "exact");
    data = r.data;
    rasterText = r.rasterText;
    stats = r.stats;
    notes.push(...r.notes);
  } else if (engine === "pdf2docx") {
    data = await runPdf2docx(buf);
  } else {
    data = await runLibreOffice(buf);
    notes.push("LibreOffice 引擎把每行文字放在独立的文本框里：版面接近，但修改时文字不会自动换行，建议优先使用内置引擎。");
  }

  const v = verify(pages, data, rasterText);
  if (v.coverage < 0.995 && v.missing) notes.unshift(`有 ${Math.round((1 - v.coverage) * v.total)} 个字符未能转换（例如：${v.missing}），请对照原 PDF 检查。`);
  if (engine !== "builtin") {
    const d = new DocxDocument(data);
    const s = d.summary();
    stats = { ...stats, paragraphs: s.blockCount };
  }
  return {
    data,
    report: {
      engine, ...stats,
      textTotal: v.total,
      textCoverage: v.coverage,
      missing: v.missing,
      notes,
      ms: Date.now() - t0,
    },
  };
}

export function describeReport(r: ConversionReport): string {
  return [
    `已用${ENGINE_LABEL[r.engine]}把 PDF 转换为 Word（${(r.ms / 1000).toFixed(1)} 秒）：${r.pages} 页，${r.paragraphs} 个段落` +
      (r.tables ? `，${r.tables} 个表格` : "") + (r.images ? `，${r.images} 张图片` : "") + "。",
    `文字完整性：${(r.textCoverage * 100).toFixed(r.textCoverage === 1 ? 0 : 2)}%（共 ${r.textTotal} 个字符）。`,
    ...r.notes,
  ].join("\n");
}

export async function hasPdf2docx(): Promise<boolean> {
  return !!(await findPython());
}
