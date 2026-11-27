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
import { looksLikeText } from "./scanLayer.js";
import { extractPdf, formatPageRange, MAX_PDF_PAGES, type ImageBox, type PageModel } from "./extract.js";

export { parsePageRange, formatPageRange, MAX_PDF_PAGES } from "./extract.js";
import { buildDocModel, type Block, type DocModel, type LayoutMode, type Para } from "./layout.js";
import { writeDocx } from "./docxWriter.js";
import { rasterAvailable, renderRegions } from "./render.js";
import { DocxDocument } from "../../documents/docx/DocxDocument.js";

const execFileAsync = promisify(execFile);

export type PdfEngine = "builtin" | "pdf2docx" | "libreoffice";

export interface ConversionReport {
  engine: PdfEngine;
  /** 本次转换的页数 */
  pages: number;
  /** PDF 的总页数 */
  totalPages?: number;
  /** 只转换了一部分页时：转换的页码（"1–20"、"1–5、8"） */
  pageRange?: string;
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
  /**
   * 扫描页（整页是一张图）怎么处理的，页码从 1 起；没有这样的页时省略。
   *   text：文字画在扫描图上面的页，正文取自这些文字，图里剩下的东西裁成图片；
   *   picture：整页保留为一张图片，且页面上没有成段的文字要读（附图），或文字已随图片保存（表单页）；
   *   unread：整页保留为一张图片，但页面看起来是文字（或无法判断）——这些文字没有被读出来，需要先做文字识别。
   *   unknown（1.2.0 起，没有时省略）：unread 里无法判断的页（彩色、灰阶照片式、几条拼成的扫描图拆不开）：可能是附图，
   *     调用方替换前要自己确认是文字页；
   *   blocks（1.2.0 起，转换时要求了 scanBlocks、没有时省略）：unread 里按原位置分块的页——整页图换成抹掉了字的底图
   *     （pdf2word-scan N），每块文字一个原位文本框（pdf2word-block N.K），框里是这块的原图（pdf2word-text N.K）。
   */
  scanPages?: { text: number[]; picture: number[]; unread: number[]; unknown?: number[]; blocks?: number[] };
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

function verify(pages: PageModel[], docx: Buffer, rasterText = "", lostText = "") {
  // lostText：没有转成正文、也不在渲染出的图片里的旋转文字。它也是原文，计入分母、记为缺失（之前不计，
  // 丢了文字仍报 100%）
  const src = charCounts(pages.flatMap((p) => p.spans.filter((s) => !s.img).map((s) => s.text)).join("") + lostText);
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

async function runBuiltin(buf: Buffer, title: string, pages: PageModel[], mode: LayoutMode, unread: ReadonlySet<number> = new Set(), scanBlocks = false) {
  const notes: string[] = [];
  const raster = await rasterAvailable();
  let model = buildDocModel(pages, { mode, raster, unread, scanBlocks });
  const regions = renderBoxes(model);
  if (regions.length) {
    const rotation = new Map(pages.filter((p) => p.rotation !== undefined).map((p) => [p.index, p.rotation!] as [number, number]));
    const ok = await renderRegions(buf, regions, undefined, rotation).catch(() => 0);
    if (ok < regions.length) {
      // 渲染失败：退回纯文字重建，保证内容不丢
      model = buildDocModel(pages, { mode, raster: false, unread, scanBlocks });
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
  if (model.stats.vectorShapes) notes.push(raster
    ? `${model.stats.vectorShapes} 处零散的矢量图形（不构成完整的图）未能转换。`
    // 不能渲染时，完整的矢量图、文字被转成轮廓的表格也都算在这里：明确说出来，而不是混在"零散图形"里
    : `${model.stats.vectorShapes} 处矢量图形未能转换：本环境不能渲染，矢量图、以轮廓绘制文字的表格和公式没有保留，请对照原 PDF。`);
  // 旋转文字（竖排、斜放）不转成正文：图里的（坐标轴标题等）随渲染出的图保留，其余的在 Word 里没有了，逐段说出来
  const figures = renderBoxes(model).filter((r) => r.render?.kind === "figure" && r.png.length);
  const rotated = pages.flatMap((pg) => (pg.rotated ?? []).map((t) => {
    t.inFigure = figures.some((f) => f.render!.page === pg.index && t.x0 >= f.x0 - 1 && t.x1 <= f.x1 + 1 && t.y0 >= f.y0 - 1 && t.y1 <= f.y1 + 1);
    return { ...t, page: pg.index + 1 };
  }));
  const lost = rotated.filter((t) => !t.inFigure && t.text.trim());
  if (lost.length) {
    const lostPages = [...new Set(lost.map((t) => t.page))];
    const sample = lost.slice(0, 3).map((t) => `“${t.text.trim().slice(0, 20)}”`).join("、");
    notes.push(`第 ${formatPageRange(lostPages)} 页有 ${lost.length} 段旋转文字（竖排或斜放）没有转成正文，例如 ${sample}，请对照原 PDF。`);
  }
  if (rotated.length > lost.length) notes.push(`${rotated.length - lost.length} 段旋转文字是图中的标注，随图保留。`);
  if (model.stats.columnPages) notes.push(`${model.stats.columnPages} 页检测到双栏排版，已用 Word 分栏还原。`);
  if (model.notes) notes.push(...model.notes);
  return { data, stats: model.stats, notes, rasterText: model.rasterText, lostText: lost.map((t) => t.text).join("") };
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

async function runPdf2docx(buf: Buffer, pages?: number[]): Promise<Buffer> {
  const py = await findPython();
  if (!py) throw new Error("未找到可用的 pdf2docx。请先安装 Python，再运行 pip install pdf2docx；或在 .env 中用 PYTHON_PATH 指定 Python 路径。");
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "doc-agent-pdf-"));
  try {
    const src = path.join(dir, "in.pdf"), out = path.join(dir, "out.docx");
    await fs.writeFile(src, buf);
    // 只转换部分页时传 0 起的页码列表
    const script = "import sys,json,logging\nlogging.disable(logging.CRITICAL)\nfrom pdf2docx import Converter\ncv=Converter(sys.argv[1])\np=json.loads(sys.argv[3])\ncv.convert(sys.argv[2],pages=p) if p else cv.convert(sys.argv[2])\ncv.close()";
    await execFileAsync(py[0], [...py.slice(1), "-c", script, src, out, JSON.stringify(pages?.map((p) => p - 1) ?? null)], { timeout: 10 * 60_000, windowsHide: true, maxBuffer: 64 * 1024 * 1024 });
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

/**
 * pages：只转换这些页（"1-20"、"1-5,8,10-12" 或从 1 开始的页码列表）；省略时转换全部（单次最多 MAX_PDF_PAGES 页）
 */
/**
 * scanBlocks（可编辑排版）：没有文字层、看得出是文字的扫描页按原位置分块——抹掉了字的底图垫底，每块文字一个原位文本框，
 * 框里是这块的原图，留给调用方识别成文字（AltaLex 翻译用：图留在原处，译文落回原来的位置）。不要求时这类页整页保留成图
 */
export async function pdfToDocx(buf: Buffer, opts: { engine?: PdfEngine; title?: string; mode?: LayoutMode; pages?: string | number[]; maxPages?: number; scanBlocks?: boolean } = {}): Promise<{ data: Buffer; report: ConversionReport }> {
  const cap = opts.maxPages ?? MAX_PDF_PAGES;
  const engine = opts.engine ?? "builtin";
  const t0 = Date.now();
  const info = { total: 0, pages: [] as number[], truncated: false };
  const partial = opts.pages !== undefined && (Array.isArray(opts.pages) ? opts.pages.length > 0 : !!opts.pages.trim());
  if (partial && engine === "libreoffice") throw new Error("LibreOffice 引擎不支持只转换部分页，请改用内置引擎。");
  const pages = await extractPdf(buf, { pages: partial ? opts.pages : undefined, info, maxPages: cap, scanBlocks: opts.scanBlocks && (opts.mode ?? "exact") === "flow" });
  // LibreOffice 引擎总是转换整份 PDF（只是完整性校验只核对前面这些页）
  if (engine === "libreoffice") info.truncated = false;
  const textChars = pages.reduce((a, p) => a + p.spans.reduce((b, s) => b + s.text.replace(/\s/g, "").length, 0), 0);
  const notes: string[] = [];
  const ocrPages = pages.filter((p) => (p.ocrChars ?? 0) > 0);
  if (ocrPages.length) {
    const n = ocrPages.reduce((a, p) => a + (p.ocrChars ?? 0), 0);
    notes.push(`${ocrPages.length} 页是带 OCR 识别层的扫描件：每页按原样保留为整页图片；压在图上的不可见识别文字（${n} 个字符）没有转成正文，否则会与图里的字重叠成两层。`);
  }
  // 几乎没有文字层的扫描件：说明放在这里，内容在转换后写（内置引擎要看哪些页按原位置分了块）
  const allScan = !ocrPages.length && pages.length > 0 && textChars < 20 * pages.length && pages.some((p) => p.images.length);
  const allScanAt = notes.length;
  let allScanNote = "这份 PDF 几乎没有可提取的文字，可能是扫描件（整页图片）。转换结果每页保留为一张整页图片，文字需要先做 OCR 识别。";

  // 只转换了一部分页（指定了页码范围，或超过单次上限被截断）：如实写进报告，文字完整性只针对这些页
  const subset = partial || info.truncated;
  const range = subset ? formatPageRange(info.pages) : undefined;
  if (info.truncated) {
    const rest = info.pages.length ? info.pages[info.pages.length - 1] + 1 : 1;
    const left = rest === info.total ? `${rest}` : `${rest}–${info.total}`, hint = rest === info.total ? `${rest}` : `${rest}-${info.total}`;
    notes.unshift(partial
      ? `指定的范围超过单次转换上限 ${cap} 页，只转换了第 ${range} 页；其余页请再指定页码范围转换。`
      : `这份 PDF 共 ${info.total} 页，超过单次转换上限 ${cap} 页：只转换了第 ${range} 页，第 ${left} 页没有转换，需要时请指定页码范围（如 ${hint}）另行转换。`);
  }

  let data: Buffer;
  let rasterText = "", lostText = "";
  let stats = { pages: pages.length, paragraphs: 0, tables: 0, images: 0, columnPages: 0, figures: 0, equations: 0 };
  let scanPages: ConversionReport["scanPages"];
  if (engine === "builtin") {
    // 没有文字层的整页图片：附图（整页保留成图），还是看得出是成行文字、没读出来的页。后者的整页图在 Word 里
    // 标上 pdf2word-page N，调用方识别出文字后可以替换
    const picture: number[] = [], unread: number[] = [], unknown: number[] = [];
    for (const pg of pages) {
      if (!pg.images.length || pg.spans.reduce((n, sp) => n + sp.text.replace(/\s/g, "").length, 0) >= 10) continue;
      if (!pg.images.some((im) => im.x1 - im.x0 >= 0.5 * pg.width && im.y1 - im.y0 >= 0.3 * pg.height)) continue;
      if (pg.scan && !looksLikeText(pg.scan)) picture.push(pg.index + 1);
      else {
        unread.push(pg.index + 1);
        // 拆不开的（彩色、灰度照片式、上下拼条的扫描图）无法判断：按"文字没被读出来"算，另外列出，调用方替换前自己确认
        if (!pg.scan) unknown.push(pg.index + 1);
      }
    }
    const r = await runBuiltin(buf, opts.title ?? "文档", pages, opts.mode ?? "exact", new Set(unread.map((n) => n - 1)), !!opts.scanBlocks);
    data = r.data;
    rasterText = r.rasterText;
    lostText = r.lostText;
    const { scan, ...rest } = r.stats;
    stats = rest;
    // 扫描页的说明排在前面：它说的是正文从哪来、哪些页没有转成文字，比版式上的说明要紧
    // 文字页：拆开扫描图的页和表单页（文字都是 PDF 里已有的识别文字）
    const text = [...(scan?.textPages ?? []), ...(scan?.formPages ?? [])].sort((a, b) => a - b);
    const blockPages = scan?.blockPages ?? [];
    if (text.length || picture.length || unread.length) scanPages = { text, picture, unread, ...(unknown.length ? { unknown } : {}), ...(blockPages.length ? { blocks: blockPages } : {}) };
    const reoriented = pages.filter((pg) => pg.reoriented).map((pg) => pg.index + 1);
    if (reoriented.length) notes.push(`第 ${formatPageRange(reoriented)} 页的扫描图与页面横竖不一致、比页面大（PDF 阅读器里这几页显示不全）：按扫描图的完整大小和方向转换。`);
    if (allScan && blockPages.length) {
      const rest = unread.filter((n) => !blockPages.includes(n));
      allScanNote = "这份 PDF 几乎没有可提取的文字，是扫描件。" + [
        picture.length ? `第 ${formatPageRange(picture)} 页是附图，整页保留成图片` : "",
        blockPages.length ? `第 ${formatPageRange(blockPages)} 页的文字按原位置分成 ${scan?.textBlocks ?? 0} 块放进文本框（框里先是这块的原图，需要识别成文字），附图、条码、表格线留在垫底的图里` : "",
        rest.length ? `第 ${formatPageRange(rest)} 页每页保留为一张整页图片，文字需要先做 OCR 识别` : "",
      ].filter(Boolean).join("；") + "。";
    }
    if (scan?.textPages.length) {
      notes.push(`第 ${formatPageRange(scan.textPages)} 页的文字画在整页扫描图上面（识别出的字已从图里抹掉）：整页扫描图不再垫在文字下面，` +
        [scan.blocks ? `图里剩下的附图、条码等 ${scan.blocks} 处裁成图片放回原位` : "", scan.inline ? `${scan.inline} 处没被识别成文字的字以小图片插回原来的行里（图片名 pdf2word-glyph）` : ""].filter(Boolean).join("，") +
        (scan.header ? `；每页相同的页眉页脚块（公开号等 ${scan.header} 处）写进了 Word 页眉页脚` : "") +
        (scan.pageNumbers ? `；印成图的页码换成了页码域` : "") +
        (scan.furniture ? `；另有 ${scan.furniture} 处页眉页脚块没有保留` : "") +
        (scan.stray ? `；${scan.stray} 处零散墨迹没有保留` : "") +
        "。正文来自 PDF 里已有的识别文字，可能有识别错误，重要的数字、日期请对照原 PDF。");
    }
    if (scan?.formPages.length) notes.push(`第 ${formatPageRange(scan.formPages)} 页是带表格线的表单：格线保留为垫底的图片，文字按原位置放在文本框里` +
      (scan.formWords ? `；格里 ${scan.formWords} 处没被识别成文字的词从格线图里取出，以小图片放回原位的文本框（图片名 pdf2word-glyph）` : "") + "。");
    // 整份都是整页图片时前面已有说明；这里只说混在文字页中间的
    if (picture.length && textChars >= 20 * pages.length) notes.push(`第 ${formatPageRange(picture)} 页没有文字层（附图等），按整页图片保留。`);
    const whole = unread.filter((n) => !blockPages.includes(n));
    if (whole.length && textChars >= 20 * pages.length) notes.push(`第 ${formatPageRange(whole)} 页是没有做过文字识别的扫描页，按整页图片保留（图片名 pdf2word-page 加页码），其中的文字没有转成正文。`);
    if (blockPages.length && textChars >= 20 * pages.length) notes.push(`第 ${formatPageRange(blockPages)} 页是没有做过文字识别的扫描页：文字按原位置分块放进文本框（框里先是这块的原图，需要识别成文字），附图、条码、表格线留在垫底的图里。`);
    notes.push(...r.notes);
  } else if (engine === "pdf2docx") {
    data = await runPdf2docx(buf, subset ? info.pages : undefined);
  } else {
    data = await runLibreOffice(buf);
    notes.push("LibreOffice 引擎把每行文字放在独立的文本框里：版面接近，但修改时文字不会自动换行，建议优先使用内置引擎。");
  }

  // 原来的位置：截断说明（unshift）在它前面
  if (allScan) notes.splice(allScanAt + (info.truncated ? 1 : 0), 0, allScanNote);
  const v = verify(pages, data, rasterText, lostText);
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
      totalPages: info.total,
      ...(range ? { pageRange: range } : {}),
      textTotal: v.total,
      textCoverage: v.coverage,
      missing: v.missing,
      ...(scanPages ? { scanPages } : {}),
      notes,
      ms: Date.now() - t0,
    },
  };
}

export function describeReport(r: ConversionReport): string {
  return [
    `已用${ENGINE_LABEL[r.engine]}把 PDF 转换为 Word（${(r.ms / 1000).toFixed(1)} 秒）：` +
      (r.pageRange ? `第 ${r.pageRange} 页（共 ${r.pages} 页，原 PDF 共 ${r.totalPages} 页）` : `${r.pages} 页`) + `，${r.paragraphs} 个段落` +
      (r.tables ? `，${r.tables} 个表格` : "") + (r.images ? `，${r.images} 张图片` : "") + "。",
    `文字完整性：${(r.textCoverage * 100).toFixed(r.textCoverage === 1 ? 0 : 2)}%（${r.pageRange ? "转换的这些页" : "共"} ${r.textTotal} 个字符）。`,
    ...r.notes,
  ].join("\n");
}

export async function hasPdf2docx(): Promise<boolean> {
  return !!(await findPython());
}
