/**
 * PDF → 版面元素提取（pdf.js）
 *
 * 输出每页的：文字片段（位置/字体/字号/粗斜体/颜色）、水平与垂直线段（表格边框、下划线、分隔线）、
 * 填充矩形（底纹）、位图（编码为 PNG）。坐标统一为"页面左上角为原点、单位 pt、y 向下"。
 */

import { pdfjsDataOptions } from "./render.js";
import zlib from "node:zlib";
import { crc32 } from "../../documents/docx/zip.js";
import { symbolChar } from "../../documents/docx/symbolFont.js";
import { isSafeFamily, isSmallCapsFont, parseFontName, measureWord, type FontInfo } from "./fonts.js";
import { splitScan, type ScanSplit } from "./scanLayer.js";

export { parseFontName };

export interface Span {
  text: string;
  x0: number; x1: number;
  top: number; bottom: number; baseline: number;
  size: number;
  font: string;
  bold: boolean;
  italic: boolean;
  color: string; // RRGGBB
  /** 按字体字宽计算的自然宽度（pt）；实际宽度明显更大 → 两端对齐拉开的行；稳定偏差 → 字符间距 */
  natural?: number;
  /** 该字体一个空格的宽度（pt） */
  spaceW?: number;
  /** PDF 中的原字体名（去掉子集前缀） */
  rawFont?: string;
  /** 数学字体（Computer Modern 数学、Cambria Math 等） */
  math?: boolean;
  /** 在 Word 目标字体中的宽度（pt）：用于逐行不折行校验；无内置字宽时为空 */
  wordW?: number;
  /** Word 目标字体中空格的宽度（pt） */
  wordSpaceW?: number;
  underline?: boolean;
  strike?: boolean;
  /** 没有文字编码的墨迹（Type 3 位图字形组成的行内公式 / 符号）：作为行内图片随文字排版，text 为占位符 U+FFFC */
  img?: ImageBox;
  /** Type 3 字体（字形由 PDF 自己描述，字符编码常不可靠；Word 中也没有这个字体） */
  type3?: boolean;
  /** 小型大写字母（小写字母用的是比普通小写字形更宽的小型大写字形，如 Elsevier 的 "ARTICLE INFO"） */
  smallCaps?: boolean;
  /** 字距加宽的文字（"A R T I C L E"）每个字符之后的额外间距（pt） */
  tracking?: number;
  /** 文字底色（Word 的突出显示 / 字符底纹，PDF 里是紧贴文字的填充矩形），RRGGBB */
  bg?: string;
  /** 填写线（表格里"人民币（大写）______"文字旁边、与文字同一基线的横线）：没有文字，写成带下划线的制表符 */
  blank?: boolean;
}

/**
 * 图像蒙版（stencil mask）：Type 3 位图字体的字形（老式 TeX/DVI 生成的 PDF 中公式、符号常用这种字体）、模板图像。
 * 它们没有文字编码，也不在路径里——不单独收集的话，公式和行内符号会整片消失。
 */
export interface MaskBox {
  x0: number; y0: number; x1: number; y1: number;
  color: string;
  /** 1 位位图，按行打包（每行 ceil(w/8) 字节），0 = 墨迹；纯色蒙版（填满整个单位方块）为 null */
  bits: Uint8Array | null;
  w: number; h: number;
  /** 页面坐标 → 位图像素坐标：u = a(X−ox) + b(Y−oy)，v = c(X−ox) + d(Y−oy)；px = u·w，py = (1−v)·h */
  inv: { ox: number; oy: number; a: number; b: number; c: number; d: number };
  /** 没有位图的字形（Type 3 字体的矢量字形）：只能由画布渲染 */
  vector?: boolean;
}

export interface Seg { x0: number; y0: number; x1: number; y1: number; width: number; color: string; horizontal: boolean }
export interface FillRect { x0: number; y0: number; x1: number; y1: number; color: string }
export interface ImageBox {
  x0: number; y0: number; x1: number; y1: number;
  png: Buffer; pxW: number; pxH: number;
  behind?: boolean;
  /** 需要从 PDF 页面渲染的区域（矢量图、公式、首字下沉）：渲染后填入 png */
  render?: { page: number; kind: "figure" | "equation" | "dropcap" };
  /** 图片中包含的文字（替代文字，供屏幕阅读器与 Agent 理解） */
  alt?: string;
  /** 文字环绕时与旁边文字的间距（pt），按原 PDF 中图片与文字的实测距离；Word 默认只留 2pt，文字几乎贴着图片 */
  wrap?: { l: number; r: number; b: number };
  /** 从扫描页的位图层上裁下来的一块（见 scanLayer.ts）：graphic 是图、条码、线框，glyphs 是没被识别成文字的字 */
  scan?: "graphic" | "glyphs";
  /** 落在一行文字里的漏认字：那一行的基线与字号（作为行内图片插回这一行） */
  line?: { baseline: number; size: number };
  /** 固定在页面坐标上、不进入文字流（版心之外的页边装饰） */
  pinned?: boolean;
  /**
   * 写进 Word 时给调用方的标记（图片名 pdf2word-glyph / pdf2word-page N / pdf2word-form N）：glyph 是句子里没被识别成文字的字，
   * page 是第 N 页没有文字层、看得出是成行文字的整页图（调用方可以把它们识别成文字后替换掉）；
   * form 是第 N 页表单垫在文本框下面的格线图（格里的字已是文字，图不必再翻译）
   */
  role?: "glyph" | "page" | "form";
  /** role 为 page、form 时的页码（从 1 起） */
  pageNo?: number;
}

/** 一条已绘制的路径（用于识别矢量图） */
export interface PathBox {
  x0: number; y0: number; x1: number; y1: number;
  curved: boolean;
  /** 含斜线（非水平/垂直） */
  diagonal: boolean;
  filled: boolean;
  stroked: boolean;
  fillColor: string;
  strokeColor: string;
}

export interface PageModel {
  index: number;
  width: number;
  height: number;
  /** 按与页面 /Rotate 不同的方向读取时的视口旋转角（渲染图片区域时用同一方向） */
  rotation?: number;
  spans: Span[];
  segs: Seg[];
  fills: FillRect[];
  images: ImageBox[];
  /** 无法转换的矢量图形（曲线等）数量 */
  vectorShapes: number;
  /** 全部已绘制的路径（含曲线、斜线） */
  paths: PathBox[];
  /** 图像蒙版（Type 3 位图字形等） */
  masks: MaskBox[];
  rotatedText: number;
  /** 压在扫描图上的不可见文字（OCR 识别层，文字渲染模式 3 / 7）的字符数：不转成正文 */
  ocrChars?: number;
  /**
   * 旋转文字（坐标轴的竖排标题、表头竖排、页边竖排说明、斜放的水印 / 印章）：外接框之外记下文字、角度和几何。
   * angle：基线方向，页面坐标（y 向下）里顺时针的度数，−90 是从下往上读；ox / oy：基线起点；len：沿基线的长度。
   * 落在渲染成图片的区域里的标 raster（随图片保存，文字进替代文字），其余的转成按原角度摆放的文本框或竖排单元格
   */
  rotated?: RotatedText[];
  /** 铺满页面的黑白扫描图拆开后的结果（见 scanLayer.ts）；image 是 images 里的那张整页图 */
  scan?: ScanSplit & { image: ImageBox };
}

type M = [number, number, number, number, number, number];
const mul = (m: M, n: M): M => [
  m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5],
];
const apply = (m: M, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const hex = (c: any): string => {
  if (typeof c === "string") return c.replace("#", "").toUpperCase().padStart(6, "0");
  const r = Number(c?.[0] ?? 0), g = Number(c?.[1] ?? 0), b = Number(c?.[2] ?? 0);
  return [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();
};
/** 半透明的颜色（ExtGState 的 ca / CA）：按压在白纸上的效果折算成不透明色（Word 的文字、底纹没有透明度） */
const overWhite = (c: string, a: number): string => {
  if (!(a < 0.999)) return c;
  const n = parseInt(c, 16), k = Math.max(0, a);
  return hex([(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => v * k + 255 * (1 - k)));
};
const isWhite = (c: string) => {
  const n = parseInt(c, 16);
  return ((n >> 16) & 255) > 245 && ((n >> 8) & 255) > 245 && (n & 255) > 245;
};

/** Symbol 类字体用私有码位（U+F0xx）编码符号：映射回真实的 Unicode（Φ、→、×…），并改用常规字体显示 */
function mapSymbolFont(text: string, font: string): { text: string; font: string } {
  if (!/symbol/i.test(font) || !/[\uf020-\uf0ff]/.test(text)) return { text, font };
  const out = text.replace(/[\uf020-\uf0ff]/g, (c) => symbolChar("Symbol", c.charCodeAt(0).toString(16)) ?? c);
  return /[\uf020-\uf0ff]/.test(out) ? { text: out, font } : { text: out, font: "Cambria Math" };
}

const LIGATURES: Record<string, string> = { "\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl", "\ufb03": "ffi", "\ufb04": "ffl", "\ufb05": "st", "\ufb06": "st" };
const expandLigatures = (s: string) => s.replace(/[\ufb00-\ufb06]/g, (c) => LIGATURES[c] ?? c);

// ---------------------------------------------------------------------------
// PNG 编码（不依赖原生库）
// ---------------------------------------------------------------------------

function pngChunk(type: string, data: Buffer): Buffer {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const td = Buffer.concat([Buffer.from(type, "latin1"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td) >>> 0, 0);
  return Buffer.concat([len, td, crc]);
}

export function encodePng(width: number, height: number, rgba: Uint8Array | Uint8ClampedArray, channels: 1 | 3 | 4): Buffer {
  const colorType = channels === 4 ? 6 : channels === 3 ? 2 : 0;
  const stride = width * channels;
  const raw = Buffer.alloc((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    raw[y * (stride + 1)] = 0;
    Buffer.from(rgba.buffer, rgba.byteOffset + y * stride, stride).copy(raw, y * (stride + 1) + 1);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = colorType;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", zlib.deflateSync(raw, { level: 6 })),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

/** 位图解码为逐像素数据（1 / 3 / 4 通道）；1 位灰度展开成每像素 1 字节 */
function decodeImage(img: any): { data: Uint8Array | Uint8ClampedArray; w: number; h: number; ch: 1 | 3 | 4 } | null {
  if (!img || !img.width || !img.height || !img.data) return null;
  const { width: w, height: h, data, kind } = img;
  if (kind === 3 || data.length === w * h * 4) return { data, w, h, ch: 4 };
  if (kind === 2 || data.length === w * h * 3) return { data, w, h, ch: 3 };
  if (kind === 1) {
    // 1 位灰度，按行打包
    const rowBytes = Math.ceil(w / 8);
    const out = new Uint8Array(w * h);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) out[y * w + x] = data[y * rowBytes + (x >> 3)] & (0x80 >> (x & 7)) ? 255 : 0;
    return { data: out, w, h, ch: 1 };
  }
  return null;
}

/** 取像素区域 [c0,c1) × [r0,r1)（图像行从上到下） */
function cropPixels(d: { data: Uint8Array | Uint8ClampedArray; w: number; h: number; ch: 1 | 3 | 4 }, c0: number, r0: number, c1: number, r1: number) {
  const w = c1 - c0, h = r1 - r0, out = new Uint8Array(w * h * d.ch);
  for (let y = 0; y < h; y++) out.set(d.data.subarray(((r0 + y) * d.w + c0) * d.ch, ((r0 + y) * d.w + c1) * d.ch), y * w * d.ch);
  return { data: out, w, h, ch: d.ch };
}

// ---------------------------------------------------------------------------

let pdfjsPromise: Promise<any> | null = null;
const loadPdfjs = () => (pdfjsPromise ??= import("pdfjs-dist/legacy/build/pdf.mjs"));

const getObj = (store: any, id: string): Promise<any> =>
  new Promise((resolve) => {
    try {
      if (store.has(id)) return resolve(store.get(id));
      store.get(id, resolve);
      setTimeout(() => resolve(null), 5000);
    } catch {
      resolve(null);
    }
  });

/** 单次转换的页数上限：再多的话 Word 文件与预览都过于庞大，按页码范围分批转换 */
export const MAX_PDF_PAGES = 500;

/**
 * 页码范围："1-20"、"1-5, 8, 10-12"、"5-"（到最后一页）、"-20"（到第 20 页），也接受 "1–20"、"1~20"、"第1至20页"、"1，3"。
 * 返回从 1 开始、升序去重、落在 [1, total] 内的页码；没有一页有效时报错
 */
export function parsePageRange(spec: string | number[], total: number): number[] {
  const out = new Set<number>();
  if (Array.isArray(spec)) {
    for (const p of spec) if (Number.isInteger(p) && p >= 1 && p <= total) out.add(p);
  } else {
    const norm = spec.replace(/[０-９]/g, (c) => String(c.charCodeAt(0) - 0xff10)).replace(/[第页頁\s]/g, "")
      .replace(/[–—－~～至到]/g, "-").replace(/[，、;；]/g, ",").replace(/前(\d+)/g, "1-$1");
    for (const part of norm.split(",").filter(Boolean)) {
      const m = /^(\d+)?(-)?(\d+)?$/.exec(part);
      if (!m || (!m[1] && !m[3])) throw new Error(`页码范围 "${spec}" 无法识别：请写成 "1-20" 或 "1-5,8,10-12"。`);
      const a = m[1] ? Number(m[1]) : 1, b = m[2] ? (m[3] ? Number(m[3]) : total) : a;
      for (let p = Math.max(1, Math.min(a, b)); p <= Math.min(total, Math.max(a, b)); p++) out.add(p);
    }
  }
  if (!out.size) throw new Error(`页码范围 "${Array.isArray(spec) ? spec.join(",") : spec}" 不在这份 PDF 的页码内（共 ${total} 页）。`);
  return [...out].sort((a, b) => a - b);
}

/** 页码列表 → 简短写法："1–20"、"1–5、8、10–12" */
export function formatPageRange(pages: number[]): string {
  const parts: string[] = [];
  for (let i = 0; i < pages.length; ) {
    let j = i;
    while (j + 1 < pages.length && pages[j + 1] === pages[j] + 1) j++;
    parts.push(j > i ? `${pages[i]}–${pages[j]}` : `${pages[i]}`);
    i = j + 1;
  }
  return parts.join("、");
}

/**
 * 提取 PDF 各页的版面元素。
 * pages：只提取这些页（页码范围或从 1 开始的页码列表）；省略时从第 1 页起。单次最多 maxPages 页（默认 MAX_PDF_PAGES），
 * 超出的不提取——info 里记下 PDF 总页数、实际提取的页码、是否因上限被截断，由调用方如实告诉用户
 */
export async function extractPdf(buf: Buffer, opts: { maxPages?: number; pages?: string | number[]; info?: { total: number; pages: number[]; truncated: boolean } } = {}): Promise<PageModel[]> {
  const pdfjs = await loadPdfjs();
  const { OPS, Util } = pdfjs;
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf), isOffscreenCanvasSupported: false, disableFontFace: true, useSystemFonts: false, verbosity: 0, ...pdfjsDataOptions() }).promise;
  const pages: PageModel[] = [];
  try {
    const all = opts.pages !== undefined && (Array.isArray(opts.pages) ? opts.pages.length : opts.pages.trim())
      ? parsePageRange(opts.pages, doc.numPages)
      : Array.from({ length: doc.numPages }, (_, i) => i + 1);
    const list = all.slice(0, opts.maxPages ?? MAX_PDF_PAGES);
    if (opts.info) Object.assign(opts.info, { total: doc.numPages, pages: list, truncated: list.length < all.length });
    for (const p of list) pages.push(await extractPage(await doc.getPage(p), p - 1, OPS, Util));
  } finally {
    await doc.destroy().catch(() => {});
  }
  return pages;
}

export interface RotatedText {
  x0: number; y0: number; x1: number; y1: number; size: number;
  text: string; angle: number; ox: number; oy: number; len: number;
  color: string; family: string; bold: boolean; italic: boolean;
  /** 字距加宽 / 紧缩（pt，每个字之后） */
  tracking?: number;
  raster?: boolean;
}

type SubPath = { kind: "rect"; pts: Array<[number, number]> } | { kind: "poly"; pts: Array<[number, number]>; ctrl: Array<[number, number]>; curved: boolean; closed: boolean };

/**
 * 一条填充路径里互相重叠的水平矩形按填充规则算出真正涂色的区域。
 * 网页导出的 PDF（WeasyPrint、浏览器打印）把边框画成"外框 + 内框"的一条路径：奇偶规则（或非零规则下
 * 内框方向相反）只涂两框之间的一圈。逐个子路径当实心块，外框就成了盖住文字的整块黑底、整行红条。
 * 不重叠的矩形原样返回（表格各格的底色、柱状图的柱子）；返回 null 表示没有需要处理的重叠
 */
export function resolveRectFill(subs: SubPath[], evenOdd: boolean): SubPath[] | null {
  type R = { x0: number; y0: number; x1: number; y1: number; sign: number; sp: SubPath };
  const rects: R[] = [];
  const others: SubPath[] = [];
  for (const sp of subs) {
    const p = sp.pts;
    const pts = sp.kind === "poly" && p.length === 5 && Math.abs(p[0][0] - p[4][0]) < 0.05 && Math.abs(p[0][1] - p[4][1]) < 0.05 ? p.slice(0, 4) : p;
    const rectLike = pts.length === 4 && !(sp.kind === "poly" && sp.curved) && pts.every((pt, i) => {
      const q = pts[(i + 1) % 4];
      return Math.abs(pt[0] - q[0]) < 0.05 || Math.abs(pt[1] - q[1]) < 0.05;
    });
    if (!rectLike) { others.push(sp); continue; }
    const xs = pts.map((q) => q[0]), ys = pts.map((q) => q[1]);
    const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
    if (x1 - x0 < 0.01 || y1 - y0 < 0.01) { others.push(sp); continue; }
    let a = 0;
    for (let i = 0; i < 4; i++) a += pts[i][0] * pts[(i + 1) % 4][1] - pts[(i + 1) % 4][0] * pts[i][1];
    rects.push({ x0, y0, x1, y1, sign: a >= 0 ? 1 : -1, sp });
  }
  if (rects.length < 2 || rects.length > 400) return null;
  // 互相重叠（面积 > 0）的矩形归成一组；只有一个矩形的组原样保留
  const parent = rects.map((_, i) => i);
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  let any = false;
  for (let i = 0; i < rects.length; i++) for (let j = i + 1; j < rects.length; j++) {
    const a = rects[i], b = rects[j];
    if (Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0) > 0.01 && Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0) > 0.01) { parent[find(i)] = find(j); any = true; }
  }
  if (!any) return null;
  const groups = new Map<number, R[]>();
  rects.forEach((r, i) => { const g = find(i); if (!groups.has(g)) groups.set(g, []); groups.get(g)!.push(r); });
  const out: SubPath[] = [...others];
  for (const g of groups.values()) {
    if (g.length === 1 || g.length > 200) { out.push(...g.map((r) => r.sp)); continue; }
    // 按所有边的坐标切成小格，逐格按填充规则判断涂不涂，再把涂色的格合并成矩形
    const uniq = (v: number[]) => v.sort((a, b) => a - b).filter((x, i, arr) => i === 0 || x - arr[i - 1] > 0.01);
    const xs = uniq(g.flatMap((r) => [r.x0, r.x1])), ys = uniq(g.flatMap((r) => [r.y0, r.y1]));
    const on: boolean[][] = [];
    for (let j = 0; j + 1 < ys.length; j++) {
      const row: boolean[] = [];
      const cy = (ys[j] + ys[j + 1]) / 2;
      for (let i = 0; i + 1 < xs.length; i++) {
        const cx = (xs[i] + xs[i + 1]) / 2;
        let n = 0, w = 0;
        for (const r of g) if (cx > r.x0 && cx < r.x1 && cy > r.y0 && cy < r.y1) { n++; w += r.sign; }
        row.push(evenOdd ? n % 2 === 1 : w !== 0);
      }
      on.push(row);
    }
    // 每一行连续涂色的格合成横条，上下几行完全相同的横条再合成一个矩形
    let open = new Map<string, { x0: number; x1: number; y0: number; y1: number }>();
    const flush = (keep: Set<string>) => {
      for (const [k, r] of open) if (!keep.has(k)) {
        out.push({ kind: "rect", pts: [[r.x0, r.y0], [r.x1, r.y0], [r.x1, r.y1], [r.x0, r.y1]] });
        open.delete(k);
      }
    };
    for (let j = 0; j < on.length; j++) {
      const runs: Array<[number, number]> = [];
      for (let i = 0; i < on[j].length; i++) {
        if (!on[j][i]) continue;
        let e = i;
        while (e + 1 < on[j].length && on[j][e + 1]) e++;
        runs.push([i, e]);
        i = e;
      }
      const keys = new Set(runs.map(([a, b]) => `${a}:${b}`));
      flush(keys);
      for (const [a, b] of runs) {
        const k = `${a}:${b}`;
        if (open.has(k)) open.get(k)!.y1 = ys[j + 1];
        else open.set(k, { x0: xs[a], x1: xs[b + 1], y0: ys[j], y1: ys[j + 1] });
      }
    }
    flush(new Set());
    open = new Map();
  }
  return out;
}

async function extractPage(page: any, index: number, OPS: any, Util: any): Promise<PageModel> {
  let vp = page.getViewport({ scale: 1 });
  let rotation: number | undefined;
  // 页面带 /Rotate 90 / 180 / 270、内容却按未旋转的版面排（合并、扫描时转了页面方向）：显示出来整页文字都横躺着
  // 或倒着，这样的文字转换不了，整页会丢。这样的页面按文字自己的方向读（90 / 270 时页面宽高随之互换）
  if ((((page.rotate ?? 0) % 360) + 360) % 360 !== 0) {
    const items = ((await page.getTextContent({ includeMarkedContent: false, disableNormalization: true })).items as any[]).filter((it) => it.str?.trim());
    // 正着的文字：基线水平、从左往右（倒着的文字基线也水平，但方向相反、字头朝下）
    const upright = (T: M) => {
      let up = 0, all = 0;
      for (const it of items) {
        const t = Util.transform(T, it.transform) as M, s = Math.hypot(t[2], t[3]), n = it.str.trim().length;
        all += n;
        if (Math.abs(t[1]) <= 0.01 * s && Math.abs(t[2]) <= 0.01 * s && t[0] > 0 && t[3] < 0) up += n;
      }
      return all ? up / all : 1;
    };
    const flat = page.getViewport({ scale: 1, rotation: 0 });
    if (upright(vp.transform) < 0.2 && upright(flat.transform) >= 0.8) { vp = flat; rotation = 0; }
  }
  const VT = vp.transform as M;
  const ops = await page.getOperatorList();
  const model: PageModel = { index, width: vp.width, height: vp.height, ...(rotation !== undefined ? { rotation } : {}), spans: [], segs: [], fills: [], images: [], vectorShapes: 0, rotatedText: 0, paths: [], masks: [] };
  // 铺满页面的黑白扫描图（只有一张时）：像素留到文字读完
  let scanned: { gray: { data: Uint8Array | Uint8ClampedArray; w: number; h: number }; box: { x0: number; y0: number; x1: number; y1: number }; image: ImageBox } | null = null;
  let scanPages = 0;

  // ---------- 操作符遍历：图形、图片、文字颜色 ----------
  // fill / stroke：折算了透明度的颜色（各处直接使用）；fill0 / stroke0、fa / sa：原色与透明度
  interface GState { ctm: M; fill: string; stroke: string; fill0: string; stroke0: string; fa: number; sa: number; lw: number; trm: number; clip?: { x0: number; y0: number; x1: number; y1: number } }
  let gs: GState = { ctm: [1, 0, 0, 1, 0, 0], fill: "000000", stroke: "000000", fill0: "000000", stroke0: "000000", fa: 1, sa: 1, lw: 1, trm: 0 };
  const setFill = (c: string) => { gs.fill0 = c; gs.fill = overWhite(c, gs.fa); };
  const setStroke = (c: string) => { gs.stroke0 = c; gs.stroke = overWhite(c, gs.sa); };
  const stack: GState[] = [];
  // rot：字形的基线不是水平的（旋转文字）。旋转文字单独按顺序对齐颜色，横排文字的对齐不受影响
  const textColors: Array<{ ch: string; color: string; fakeBold: boolean; adv: number; spaceBefore: number; font: string; hidden: boolean; clip?: { x0: number; y0: number; x1: number; y1: number }; rot?: boolean }> = [];
  let tm: M = [1, 0, 0, 1, 0, 0]; // 文字矩阵（只用来判断字形的方向）
  const spaceEm = new Map<string, number>(); // 各字体空格的字宽（em）
  let curFont = "";
  // 字形宽度的单位由字体矩阵决定：多数字体是 1/1000 em，但也有 1/2048（Elsevier 排版的 Charis SIL 等）、
  // Type 3 字体任意。按 1/1000 计会把字宽算成 2 倍，进而把字符间距压成每字 −3 磅，文字挤成一团
  const scaleOf = new Map<string, number>();
  const advScale = (name: string) => {
    if (!scaleOf.has(name)) {
      let k = 0.001;
      try {
        const fm = page.commonObjs.has(name) ? page.commonObjs.get(name)?.fontMatrix : undefined;
        if (Array.isArray(fm) && typeof fm[0] === "number" && fm[0] > 0 && Math.abs(fm[1] ?? 0) < 1e-9) k = fm[0];
      } catch { /* 取不到字体矩阵时按 1/1000 */ }
      scaleOf.set(name, k);
    }
    return scaleOf.get(name)!;
  };
  let curScale = 0.001;
  let pendingPath: SubPath[] = [];

  const toPage = (x: number, y: number) => apply(VT, ...apply(gs.ctm, x, y));
  const scale = () => Math.sqrt(Math.abs(gs.ctm[0] * gs.ctm[3] - gs.ctm[1] * gs.ctm[2])) || 1;

  const paint = (stroke: boolean, fill: boolean, evenOdd = false) => {
    // 填充路径里互相重叠的矩形先按填充规则算出涂色区域（边框的"外框 + 内框"只涂一圈）；
    // 描边仍沿原来的子路径
    const resolved = fill ? resolveRectFill(pendingPath, evenOdd) : null;
    if (resolved) {
      const all = pendingPath;
      if (stroke) { pendingPath = all; paintSubs(pendingPath, true, false); }
      paintSubs(resolved, false, true);
      pendingPath = [];
      return;
    }
    paintSubs(pendingPath, stroke, fill);
    pendingPath = [];
  };
  const paintSubs = (subs: SubPath[], stroke: boolean, fill: boolean) => {
    for (const sp of subs) {
      const xs = sp.pts.map((p) => p[0]), ys = sp.pts.map((p) => p[1]);
      const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
      const w = x1 - x0, h = y1 - y0;
      // 每一小段都是水平或竖直的（按段长的比例判断：曲线图的折线由大量很短的斜段组成，不能用绝对阈值）
      const axisAligned = sp.kind === "rect" || (!sp.curved && sp.pts.every((pt, i) => {
        if (i === 0) return true;
        const dx = Math.abs(pt[0] - sp.pts[i - 1][0]), dy = Math.abs(pt[1] - sp.pts[i - 1][1]);
        return Math.min(dx, dy) <= 0.1 * Math.max(dx, dy) + 0.05;
      }));
      if ((stroke || fill) && (w > 0.3 || h > 0.3) && !(fill && !stroke && isWhite(gs.fill))) {
        const bb = sp.kind === "poly" && sp.ctrl.length ? [...sp.pts, ...sp.ctrl] : sp.pts;
        const bx = bb.map((p) => p[0]), by = bb.map((p) => p[1]);
        model.paths.push({
          x0: Math.min(...bx), y0: Math.min(...by), x1: Math.max(...bx), y1: Math.max(...by),
          curved: sp.kind === "poly" && sp.curved, diagonal: !axisAligned && !(sp.kind === "poly" && sp.curved),
          filled: fill, stroked: stroke, fillColor: gs.fill, strokeColor: gs.stroke,
        });
      }
      if (!axisAligned) {
        if ((stroke || fill) && w * h > 4) model.vectorShapes++;
        continue;
      }
      const isRectLike = sp.kind === "rect" || sp.pts.length >= 4;
      if (fill) {
        if (Math.min(w, h) <= 2.5 && Math.max(w, h) > 2) {
          model.segs.push({ x0, y0, x1, y1, width: Math.min(w, h) || 0.5, color: gs.fill, horizontal: w >= h });
        } else if (isRectLike && w > 2 && h > 2 && !isWhite(gs.fill)) {
          model.fills.push({ x0, y0, x1, y1, color: gs.fill });
        } else if (isRectLike && w * h >= 0.85 * model.width * model.height && gs.fill !== "FFFFFF" && !/^F[EF]F[EF]F[EF]$/.test(gs.fill)) {
          // 铺满页面的浅色底（米色、浅灰的页面颜色）：接近白色也是有意设的页面颜色，记下来（作为页面背景色）；纯白的不记
          model.fills.push({ x0, y0, x1, y1, color: gs.fill });
        }
      }
      if (stroke) {
        const lw = Math.max(0.25, gs.lw * scale());
        const pts = sp.pts.slice();
        if (sp.kind === "rect" || (sp.kind === "poly" && sp.closed)) pts.push(pts[0]);
        for (let i = 1; i < pts.length; i++) {
          const [ax, ay] = pts[i - 1], [bx, by] = pts[i];
          const horiz = Math.abs(ay - by) < 0.6, vert = Math.abs(ax - bx) < 0.6;
          if (!horiz && !vert) continue;
          const len = horiz ? Math.abs(bx - ax) : Math.abs(by - ay);
          if (len < 1) continue;
          // 很粗的"线"其实是色块（出版社用一根 60 多磅宽的线画刊头的灰底）：按色块处理。
          // 当作线会被识别成穿过刊名的删除线，也不会画出底色
          if (lw > 4 && len > 4 * lw / 4 && !isWhite(gs.stroke)) {
            const mx = (ax + bx) / 2, my = (ay + by) / 2;
            model.fills.push(horiz
              ? { x0: Math.min(ax, bx), y0: my - lw / 2, x1: Math.max(ax, bx), y1: my + lw / 2, color: gs.stroke }
              : { x0: mx - lw / 2, y0: Math.min(ay, by), x1: mx + lw / 2, y1: Math.max(ay, by), color: gs.stroke });
            continue;
          }
          model.segs.push({ x0: Math.min(ax, bx), y0: Math.min(ay, by), x1: Math.max(ax, bx), y1: Math.max(ay, by), width: lw, color: gs.stroke, horizontal: horiz });
        }
      }
    }
  };

  // 图像蒙版：单位方块经 ctm、视口变换映射到页面；记下外接框与反向映射，供合成图片
  const addMask = (ctm: M, img: any) => {
    const F = (u: number, v: number) => apply(VT, ...apply(ctm, u, v));
    const [ox, oy] = F(0, 0), [ux, uy] = F(1, 0), [vx, vy] = F(0, 1), [qx, qy] = F(1, 1);
    const det = (ux - ox) * (vy - oy) - (vx - ox) * (uy - oy);
    if (Math.abs(det) < 1e-9) return;
    const x0 = Math.min(ox, ux, vx, qx), x1 = Math.max(ox, ux, vx, qx), y0 = Math.min(oy, uy, vy, qy), y1 = Math.max(oy, uy, vy, qy);
    if (x1 - x0 < 0.05 && y1 - y0 < 0.05) return;
    const w = img?.width | 0, h = img?.height | 0;
    const bits = img?.data instanceof Uint8Array && w > 0 && h > 0 && img.data.length >= Math.ceil(w / 8) * h ? img.data : null;
    model.masks.push({
      x0, y0, x1, y1, color: gs.fill, bits, w: bits ? w : 1, h: bits ? h : 1,
      inv: { ox, oy, a: (vy - oy) / det, b: -(vx - ox) / det, c: -(uy - oy) / det, d: (ux - ox) / det },
    });
  };
  const maskData = async (a: any) => (a && typeof a.data === "string" ? await getObj(String(a.data).startsWith("g_") ? page.commonObjs : page.objs, a.data) : a);

  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i];
    const args = ops.argsArray[i];
    switch (fn) {
      case OPS.save: stack.push({ ...gs, ctm: [...gs.ctm] as M }); break;
      case OPS.restore: gs = stack.pop() ?? gs; break;
      case OPS.transform: gs.ctm = mul(gs.ctm, args as M); break;
      case OPS.setFillRGBColor: setFill(hex(args)); break;
      case OPS.setStrokeRGBColor: setStroke(hex(args)); break;
      case OPS.setFillGray: setFill(hex([args[0] * 255, args[0] * 255, args[0] * 255])); break;
      case OPS.setStrokeGray: setStroke(hex([args[0] * 255, args[0] * 255, args[0] * 255])); break;
      case OPS.setGState: {
        // 透明度：半透明的水印、印章、底色块按压在白纸上的颜色输出
        for (const [k, v] of Array.isArray(args?.[0]) ? args[0] : []) {
          if (k === "ca" && typeof v === "number") { gs.fa = v; setFill(gs.fill0); }
          else if (k === "CA" && typeof v === "number") { gs.sa = v; setStroke(gs.stroke0); }
        }
        break;
      }
      case OPS.setLineWidth: gs.lw = args[0]; break;
      case OPS.setTextRenderingMode: gs.trm = args[0]; break;
      case OPS.setFont: curFont = String(args[0]); curScale = advScale(curFont); break;
      case OPS.beginText: tm = [1, 0, 0, 1, 0, 0]; break;
      case OPS.setTextMatrix: if (Array.isArray(args) && args.length >= 6) tm = args.slice(0, 6) as M; break;
      case OPS.showText:
      case OPS.showSpacedText: {
        const glyphs = args[0] ?? [];
        let pendingSpace = 0;
        const [bx, by] = (() => { const a = [tm[0], tm[1]], b = [gs.ctm[0] * a[0] + gs.ctm[2] * a[1], gs.ctm[1] * a[0] + gs.ctm[3] * a[1]]; return [VT[0] * b[0] + VT[2] * b[1], VT[1] * b[0] + VT[3] * b[1]]; })();
        const rot = Math.abs(by) > 0.01 * Math.hypot(bx, by) || bx < 0;
        for (const g of glyphs) {
          if (!g || typeof g !== "object") continue;
          const w = typeof g.width === "number" ? g.width * curScale : 0;
          // 注意：pdf.js 把字符码 32 一律标为 isSpace，但子集字体里码 32 可能是任意字母（如 m、A），只能以 Unicode 判断
          const uni = String(g.unicode ?? "");
          if (!uni.trim()) {
            if (uni === " " || (g.isSpace && !uni)) { pendingSpace += w; if (w > 0) spaceEm.set(curFont, w); }
            continue;
          }
          // 按 UTF-16 码元逐个记录，与字符串 indexOf 的下标一一对应
          const u = uni;
          for (let k = 0; k < u.length; k++) {
            if (!u[k].trim()) continue;
            textColors.push({ ch: u[k], color: gs.trm === 1 ? gs.stroke : gs.fill, fakeBold: gs.trm === 2, adv: k === 0 ? w : 0, spaceBefore: k === 0 ? pendingSpace : 0, font: curFont, hidden: gs.trm === 3 || gs.trm === 7, clip: gs.clip, ...(rot ? { rot } : {}) });
          }
          pendingSpace = 0;
        }
        break;
      }
      case OPS.constructPath: {
        const [subOps, coords] = args as [number[], number[]];
        let k = 0;
        let cur: { kind: "poly"; pts: Array<[number, number]>; ctrl: Array<[number, number]>; curved: boolean; closed: boolean } | null = null;
        for (const op of subOps) {
          if (op === OPS.rectangle) {
            const [x, y, w, h] = coords.slice(k, k + 4);
            k += 4;
            pendingPath.push({ kind: "rect", pts: [toPage(x, y), toPage(x + w, y), toPage(x + w, y + h), toPage(x, y + h)] });
          } else if (op === OPS.moveTo) {
            cur = { kind: "poly", pts: [toPage(coords[k], coords[k + 1])], ctrl: [], curved: false, closed: false };
            pendingPath.push(cur);
            k += 2;
          } else if (op === OPS.lineTo) {
            if (!cur) { cur = { kind: "poly", pts: [], ctrl: [], curved: false, closed: false }; pendingPath.push(cur); }
            cur.pts.push(toPage(coords[k], coords[k + 1]));
            k += 2;
          } else if (op === OPS.curveTo) {
            if (cur) { cur.curved = true; cur.ctrl.push(toPage(coords[k], coords[k + 1]), toPage(coords[k + 2], coords[k + 3])); cur.pts.push(toPage(coords[k + 4], coords[k + 5])); }
            k += 6;
          } else if (op === OPS.curveTo2 || op === OPS.curveTo3) {
            if (cur) { cur.curved = true; cur.ctrl.push(toPage(coords[k], coords[k + 1])); cur.pts.push(toPage(coords[k + 2], coords[k + 3])); }
            k += 4;
          } else if (op === OPS.closePath) {
            if (cur) cur.closed = true;
          }
        }
        break;
      }
      case OPS.stroke: case OPS.closeStroke: paint(true, false); break;
      case OPS.fill: paint(false, true); break;
      case OPS.eoFill: paint(false, true, true); break;
      case OPS.fillStroke: case OPS.closeFillStroke: paint(true, true); break;
      case OPS.eoFillStroke: case OPS.closeEOFillStroke: paint(true, true, true); break;
      case OPS.endPath: pendingPath = []; break;
      case OPS.clip: case OPS.eoClip: {
        // 裁剪区域只取水平 / 竖直的矩形（照片的取景框）；其他形状的裁剪不处理
        // 只有一个点的子路径（pdfTeX 在 closepath 之后又写了一个 moveto）不构成形状，忽略
        const subs = pendingPath.filter((sp) => sp.kind === "rect" || sp.pts.length >= 2);
        const pts = subs.flatMap((sp) => sp.pts);
        const rectLike = subs.length > 0 && subs.every((sp) => sp.kind === "rect" || (!sp.curved && sp.pts.length >= 4 && sp.pts.every((pt, i) => {
          const q = sp.pts[(i + 1) % sp.pts.length];
          return Math.abs(pt[0] - q[0]) < 0.05 || Math.abs(pt[1] - q[1]) < 0.05;
        })));
        if (rectLike && pts.length) {
          const c = { x0: Math.min(...pts.map((p) => p[0])), y0: Math.min(...pts.map((p) => p[1])), x1: Math.max(...pts.map((p) => p[0])), y1: Math.max(...pts.map((p) => p[1])) };
          const o = gs.clip;
          gs = { ...gs, clip: o ? { x0: Math.max(o.x0, c.x0), y0: Math.max(o.y0, c.y0), x1: Math.min(o.x1, c.x1), y1: Math.min(o.y1, c.y1) } : c };
        }
        break;
      }
      // 表单 XObject（LaTeX \includegraphics[trim, clip] 插入的 PDF / 图片都包在其中）：自带变换矩阵，
      // 并按 BBox 裁剪。不处理时，裁剪掉的部分也会被当成图片的一部分，图片比原文大一圈、盖住旁边的图和文字
      case OPS.paintFormXObjectBegin: {
        stack.push({ ...gs, ctm: [...gs.ctm] as M });
        const [matrix, bbox] = (args ?? []) as [ArrayLike<number> | null, ArrayLike<number> | null];
        if (matrix && matrix.length === 6) gs.ctm = mul(gs.ctm, Array.from(matrix) as M);
        if (bbox && bbox.length === 4) {
          const [bx0, by0, bx1, by1] = Array.from(bbox);
          const cs = [toPage(bx0, by0), toPage(bx1, by0), toPage(bx1, by1), toPage(bx0, by1)];
          // 只取映射后仍是水平 / 竖直矩形的 BBox（旋转的表单不裁剪）
          const axis = Math.abs(cs[0][1] - cs[1][1]) < 0.05 || Math.abs(cs[0][0] - cs[1][0]) < 0.05;
          if (axis) {
            const c = { x0: Math.min(...cs.map((p) => p[0])), y0: Math.min(...cs.map((p) => p[1])), x1: Math.max(...cs.map((p) => p[0])), y1: Math.max(...cs.map((p) => p[1])) };
            const o = gs.clip;
            gs = { ...gs, clip: o ? { x0: Math.max(o.x0, c.x0), y0: Math.max(o.y0, c.y0), x1: Math.min(o.x1, c.x1), y1: Math.min(o.y1, c.y1) } : c };
          }
        }
        break;
      }
      case OPS.paintFormXObjectEnd: gs = stack.pop() ?? gs; break;
      case OPS.paintImageMaskXObject: {
        try { addMask(gs.ctm, await maskData(args[0])); } catch { /* 忽略损坏的蒙版 */ }
        break;
      }
      case OPS.paintImageMaskXObjectGroup: {
        for (const im of (args[0] ?? []) as any[]) {
          try { addMask(im?.transform ? mul(gs.ctm, im.transform as M) : gs.ctm, await maskData(im)); } catch { /* 忽略 */ }
        }
        break;
      }
      case OPS.paintImageMaskXObjectRepeat: {
        const [im, sx, kx, ky, sy, pos] = args as [any, number, number, number, number, number[]];
        try {
          const data = await maskData(im);
          for (let k = 0; k + 1 < (pos?.length ?? 0); k += 2) addMask(mul(gs.ctm, [sx, kx, ky, sy, pos[k], pos[k + 1]]), data);
        } catch { /* 忽略 */ }
        break;
      }
      case OPS.paintSolidColorImageMask: addMask(gs.ctm, null); break;
      case OPS.paintImageXObject:
      case OPS.paintInlineImageXObject: {
        const img = fn === OPS.paintInlineImageXObject ? args[0] : await getObj(String(args[0]).startsWith("g_") ? page.commonObjs : page.objs, args[0]);
        let dec = decodeImage(img);
        const corners = [toPage(0, 0), toPage(1, 0), toPage(0, 1), toPage(1, 1)];
        const xs = corners.map((c) => c[0]), ys = corners.map((c) => c[1]);
        let box = { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
        // 照片常以更大的尺寸放置、再用取景框裁剪（IEEE 作者照片：方形原图裁成 1 × 1.25 英寸的竖幅）。
        // 只取取景框内可见的部分，否则图片比实际显示的大，会盖住旁边的文字、与相邻照片被误并
        const cl = gs.clip;
        if (dec && cl && box.x1 - box.x0 > 2 && box.y1 - box.y0 > 2) {
          const vis = { x0: Math.max(box.x0, cl.x0), y0: Math.max(box.y0, cl.y0), x1: Math.min(box.x1, cl.x1), y1: Math.min(box.y1, cl.y1) };
          if (vis.x1 - vis.x0 <= 2 || vis.y1 - vis.y0 <= 2) { dec = null; break; } // 完全被裁掉
          const axisAligned = Math.abs(corners[1][1] - corners[0][1]) < 0.05 && Math.abs(corners[2][0] - corners[0][0]) < 0.05;
          const cut = vis.x0 - box.x0 > 0.5 || box.x1 - vis.x1 > 0.5 || vis.y0 - box.y0 > 0.5 || box.y1 - vis.y1 > 0.5;
          if (axisAligned && cut) {
            // 页面坐标 → 图像单位方块：u 向右、v 向上；像素行从上到下
            const [ox, oy] = corners[0], du = corners[1][0] - ox, dv = corners[2][1] - oy;
            const us = [(vis.x0 - ox) / du, (vis.x1 - ox) / du].sort((a, b) => a - b);
            const vs = [(vis.y0 - oy) / dv, (vis.y1 - oy) / dv].sort((a, b) => a - b);
            const c0 = Math.max(0, Math.floor(us[0] * dec.w + 1e-6)), c1 = Math.min(dec.w, Math.ceil(us[1] * dec.w - 1e-6));
            const r0 = Math.max(0, Math.floor((1 - vs[1]) * dec.h + 1e-6)), r1 = Math.min(dec.h, Math.ceil((1 - vs[0]) * dec.h - 1e-6));
            if (c1 - c0 >= 1 && r1 - r0 >= 1) {
              dec = cropPixels(dec, c0, r0, c1, r1);
              box = vis;
            }
          }
        }
        const png = dec && { png: encodePng(dec.w, dec.h, dec.data, dec.ch), w: dec.w, h: dec.h };
        if (png && box.x1 - box.x0 > 2 && box.y1 - box.y0 > 2) {
          const im: ImageBox = { ...box, png: png.png, pxW: png.w, pxH: png.h };
          model.images.push(im);
          // 铺满页面的单通道位图（黑白扫描）：留着像素，等文字读完后把它拆开；一页有两张以上（上下拼条）的不拆
          const straight = Math.abs(corners[1][1] - corners[0][1]) < 0.05 && Math.abs(corners[2][0] - corners[0][0]) < 0.05 && corners[1][0] > corners[0][0] && corners[2][1] < corners[0][1];
          if (dec!.ch === 1 && straight && box.x1 - box.x0 >= 0.85 * model.width && box.y1 - box.y0 >= 0.85 * model.height) {
            scanPages++;
            scanned = scanPages === 1 ? { gray: { data: dec!.data, w: dec!.w, h: dec!.h }, box, image: im } : null;
          }
        } else if (!png) model.vectorShapes++;
        break;
      }
    }
  }

  // ---------- 文字 ----------
  // 不做 Unicode 规范化，保证与操作符流中的字形一一对应；连字（ﬁ ﬂ）单独展开
  const tc = await page.getTextContent({ includeMarkedContent: false, disableNormalization: true });
  const fontInfo = new Map<string, FontInfo>();
  const rawNames = new Map<string, string>();
  const type3 = new Map<string, boolean>();
  const isType3 = (fontName: string) => {
    if (!type3.has(fontName)) {
      let t = false;
      try { t = page.commonObjs.has(fontName) && !!page.commonObjs.get(fontName)?.isType3Font; } catch { /* 忽略 */ }
      type3.set(fontName, t);
    }
    return type3.get(fontName)!;
  };
  const infoOf = (fontName: string): FontInfo => {
    if (!fontInfo.has(fontName)) {
      let raw: string | undefined;
      try { raw = page.commonObjs.has(fontName) ? page.commonObjs.get(fontName)?.name : undefined; } catch { /* 忽略 */ }
      rawNames.set(fontName, (raw ?? "").replace(/^[A-Z]{6}\+/, ""));
      const info = parseFontName(raw ?? tc.styles[fontName]?.fontFamily);
      // 出版社的定制字体（Gulliver、Charter 变体、……）：族名不在 Word 常见字体里时，按 PDF 里字体的类别（衬线 / 无衬线 / 等宽）
      // 回退到同类的 Word 字体，否则读者的电脑上会被随意替换成宽度完全不同的字体，整页折行错乱
      if (!info.math && !isSafeFamily(info.family) && !isType3(fontName)) {
        const cls = tc.styles[fontName]?.fontFamily;
        info.family = cls === "monospace" ? "Courier New" : cls === "sans-serif" ? "Arial" : "Times New Roman";
      }
      fontInfo.set(fontName, info);
    }
    return fontInfo.get(fontName)!;
  };


  // 每个字体（按原字体名，跨子集）里每个小写字母最常见的字宽：小型大写字形的字宽明显不同
  const advMode = new Map<string, number>();
  {
    const counts = new Map<string, Map<number, number>>();
    for (const t of textColors) {
      if (!/^[a-z]$/.test(t.ch) || !t.adv) continue;
      infoOf(t.font);
      const k = `${rawNames.get(t.font) ?? t.font}|${t.ch}`;
      const a = Math.round(t.adv * 1000);
      if (!counts.has(k)) counts.set(k, new Map());
      counts.get(k)!.set(a, (counts.get(k)!.get(a) ?? 0) + 1);
    }
    for (const [k, m] of counts) advMode.set(k, [...m.entries()].sort((a, b) => b[1] - a[1])[0][0] / 1000);
  }

  let ci = 0; // textColors 游标
  const opText = textColors.map((t) => t.ch).join("");
  // 旋转的字形单独成一条字符流，旋转文字按顺序在里面找自己的颜色（单个字母在全文里随便一找会找到别处的字）
  const rotIdx = textColors.flatMap((t, i) => (t.rot ? [i] : []));
  const rotText = rotIdx.map((i) => textColors[i].ch).join("");
  let rci = 0;
  for (const it of tc.items as any[]) {
    if (!it.str || !it.str.trim()) continue;
    const t = Util.transform(VT, it.transform) as M;
    const size = Math.hypot(t[2], t[3]);
    if (size < 0.5) continue;
    // 基线不水平的，或水平但倒着（从右往左、字头朝下）的：旋转文字
    if (Math.abs(t[1]) > 0.01 * size || Math.abs(t[2]) > 0.01 * size || t[0] < 0) {
      model.rotatedText++;
      const em = Math.hypot(it.transform[0], it.transform[1]) || size;
      const wEm = (it.width || em * it.str.length * 0.5) / em;
      const cs = [apply(t, 0, -0.22), apply(t, wEm, -0.22), apply(t, 0, 0.85), apply(t, wEm, 0.85)];
      // 颜色：在旋转字形的字符流里按顺序找到这段文字（横排文字的游标不动）
      const joined = [...it.str].filter((c) => c.trim()).join("");
      let k = joined ? rotText.indexOf(joined, rci) : -1;
      // 字符流里是连字等合写的字形时整段对不上：用开头几个字定位
      if (k === -1 && joined.length > 3) k = rotText.indexOf(joined.slice(0, 3), rci);
      const at = k !== -1 ? rotIdx[k] : -1;
      if (k !== -1) rci = k + Math.max(1, rotText.startsWith(joined, k) ? joined.length : joined.length - 2);
      // 不可见的文字（文字渲染模式 3 / 7，扫描件的识别层）：不转换
      if (at !== -1 && rotIdx.slice(k, k + joined.length).every((i) => textColors[i].hidden)) continue;
      const fr = infoOf(it.fontName);
      const [ox, oy] = apply(t, 0, 0);
      const emX = Math.hypot(t[0], t[1]), len = wEm * emX;
      let text = expandLigatures(it.str.replace(/\u00ad/g, ""));
      let tracking: number | undefined;
      // 字距加宽的旋转文字（印章 "PAID"、水印 "CONFIDENTIAL"）：pdf.js 按间距在字母之间插了空格（"PA I D"）。
      // 按字形流还原：只在真正有空格字形的地方留空格，加宽量（实际长度减去字形自然宽度）折成字符间距
      if (k !== -1 && rotText.startsWith(joined, k)) {
        const gl = rotIdx.slice(k, k + joined.length).map((i) => textColors[i]);
        const real = gl.map((g, j) => (j > 0 && g.spaceBefore > 0 ? " " : "") + g.ch).join("");
        const nat = gl.reduce((a, g, j) => a + g.adv + (j > 0 && g.spaceBefore > 0 ? g.spaceBefore : 0), 0) * emX;
        const inserted = (it.str.trim().match(/\s+/g) ?? []).length > (real.match(/ /g) ?? []).length;
        if (nat > 0 && (inserted || (!/\s/.test(it.str.trim()) && gl.length >= 3))) {
          const tr = (len - nat) / gl.length;
          if (Math.abs(tr) > 0.03 * size) tracking = tr;
          if (inserted) text = expandLigatures(real);
        }
      }
      (model.rotated ??= []).push({
        x0: Math.min(...cs.map((c) => c[0])), y0: Math.min(...cs.map((c) => c[1])), x1: Math.max(...cs.map((c) => c[0])), y1: Math.max(...cs.map((c) => c[1])), size,
        text, angle: (Math.atan2(t[1], t[0]) * 180) / Math.PI, ox, oy, len,
        color: at !== -1 ? textColors[at].color : "000000", family: fr.family, bold: fr.bold || (at !== -1 && textColors[at].fakeBold), italic: fr.italic,
        ...(tracking !== undefined ? { tracking } : {}),
      });
      continue;
    }
    const style = tc.styles[it.fontName] ?? {};
    const ascent = typeof style.ascent === "number" && style.ascent > 0.5 ? Math.min(style.ascent, 1.05) : 0.85;
    const descent = typeof style.descent === "number" && style.descent < 0 ? Math.max(style.descent, -0.35) : -0.22;
    const baseline = t[5];
    const x0 = t[4];
    const width = it.width || size * it.str.length * 0.5;
    const fi = infoOf(it.fontName);

    // 与操作符流中的字符对齐，取得颜色与"伪粗体"（描边加粗）：用前几个字符在操作符字符流里定位
    let color = "000000", fakeBold = false, smallCaps = false, hidden = false;
    let natural: number | undefined;
    // 一个文字项里颜色 / 伪粗体有变化（行首蓝色的引用编号 "[23]" 接着黑色正文）：pdf.js 按字体合并文字项、不看颜色，
    // 整项取第一个字的颜色会把整行正文染成引用的蓝色。按变化处切成几段，各自的颜色、位置、自然字宽
    let pieces: Array<{ str: string; x0: number; x1: number; color: string; fakeBold: boolean; natural?: number; clipped?: boolean }> = [];
    let clippedOut = false;
    // 字距加宽的文字（"a r t i c l e"）：pdf.js 按间距在每个字母之间插入了空格，还原成一个词，加宽由字符间距体现
    const spaced = /^(?:\S ){3,}\S$/u.test(it.str);
    const chars = [...it.str].filter((c) => c.trim());
    if (chars.length) {
      const joined = chars.join("");
      // 文字项与操作符流的顺序一致：优先在游标处原位匹配；对不上时只在附近小范围内重新同步，
      // 避免"•"这类短文字误配到后面很远的同一字符上
      const probe = joined.slice(0, Math.min(6, joined.length));
      const fits = (pos: number) => opText.startsWith(probe, pos);
      let at = -1;
      if (fits(ci)) at = ci;
      else {
        for (let d = 1; d <= 80 && at === -1; d++) {
          if (fits(ci + d)) at = ci + d;
          else if (probe.length >= 3 && ci - d >= 0 && fits(ci - d)) at = ci - d;
        }
        if (at === -1 && probe.length >= 6) {
          const far = opText.indexOf(joined.slice(0, Math.min(16, joined.length)), ci);
          if (far !== -1 && far - ci < 4000) at = far;
        }
      }
      if (at !== -1) {
        color = textColors[at].color;
        fakeBold = textColors[at].fakeBold;
        // 逐字核对并累计自然字宽；记下每个字形的起点（em）与它之前已计入的空格字形数
        let pos = at, em = 0, spaces = 0, matched = 0, lower = 0, wide = 0;
        const starts: number[] = [], spacesAt: number[] = [];
        for (const ch of joined) {
          if (opText[pos] !== ch) break;
          if (/^[a-z]$/.test(ch) && textColors[pos].adv) {
            const mode = advMode.get(`${rawNames.get(textColors[pos].font) ?? textColors[pos].font}|${ch}`);
            lower++;
            if (mode && textColors[pos].adv > 1.06 * mode) wide++;
          }
          const lead = pos > at ? textColors[pos].spaceBefore : 0;
          if (lead > 0) spaces++;
          starts.push(em + lead);
          spacesAt.push(spaces);
          em += textColors[pos].adv + lead;
          pos++;
          matched++;
        }
        ci = pos;
        // 不可见的文字（渲染模式 3 / 7：扫描件的 OCR 识别层）
        if (matched > 0) hidden = textColors.slice(at, at + matched).filter((t) => t.hidden).length > 0.5 * matched;
        if (matched === joined.length && em > 0) {
          // 用 TJ 位移代替空格字形的 PDF：按字体空格宽度补上词间空格
          const inner = spaced ? 0 : (it.str.trim().match(/\s+/g) ?? []).length;
          const spEm = spaceEm.get(it.fontName) ?? 0.25;
          const missing = spaces < inner;
          if (missing) em += (inner - spaces) * spEm;
          natural = em * size;
          // 同一字形的字宽总是相同，普通文字不会出现偏差；有一半以上的小写字母变宽即为小型大写
          smallCaps = lower >= 3 && wide >= Math.max(2, 0.6 * lower);
          // 落在裁剪框外的字看不见（网页导出的"…………"引导点被 overflow: hidden 截在金额前面）：不转换。
          // 整段转出来会穿过金额、把一行撑长，逐页保留时挤出页面
          const clipped = (j: number) => {
            const c = textColors[at + j].clip;
            if (!c) return false;
            const cx = x0 + (width * (starts[j] + textColors[at + j].adv / 2)) / em, cy = baseline - 0.3 * size;
            return cx < c.x0 - 0.5 || cx > c.x1 + 0.5 || cy < c.y0 - 0.5 * size || cy > c.y1 + 0.5 * size;
          };
          const hiddenAt = [...joined].map((_, j) => clipped(j));
          if (hiddenAt.every(Boolean)) { clippedOut = true; }
          // 按颜色 / 伪粗体 / 是否在裁剪框外切段（只处理 BMP 字符：操作符流按 UTF-16 码元记录）
          const keyOf = (j: number) => `${textColors[at + j].color}|${textColors[at + j].fakeBold ? 1 : 0}|${hiddenAt[j] ? 1 : 0}`;
          // 跨过表格竖线的文字项：相邻两格的文字间距不大时，pdf.js 当作词间空格把它们合成一项（"…testing 2027-02-14"），
          // 整项按中心落进左边一格、右边一格变空。在竖线所在的词间空白处切开
          const mid = baseline - 0.3 * size;
          const vlines = model.segs.filter((g) => !g.horizontal && g.x0 > x0 && g.x1 < x0 + width && g.y1 - g.y0 >= 0.8 * size && g.y0 <= mid && g.y1 >= mid);
          if (!spaced && !/[\ud800-\udfff]/.test(it.str) && joined.length > 1 && !clippedOut && (vlines.length || [...joined].some((_, j) => j > 0 && keyOf(j) !== keyOf(j - 1)))) {
            // 每个非空白字符在 it.str 中的下标，以及它前面 it.str 里的空白段数（TJ 位移代替空格时据此补上空格宽度）
            const idx: number[] = [], gapsBefore: number[] = [];
            let gaps = 0, prevSpace = false;
            for (let k = 0; k < it.str.length; k++) {
              if (!it.str[k].trim()) { if (!prevSpace && idx.length) gaps++; prevSpace = true; continue; }
              prevSpace = false;
              idx.push(k);
              gapsBefore.push(gaps);
            }
            const startEm = (j: number) => starts[j] + (missing ? Math.max(0, gapsBefore[j] - spacesAt[j]) * spEm : 0);
            const atRule = (j: number) => {
              if (!vlines.length || gapsBefore[j] === gapsBefore[j - 1]) return false;
              const a = x0 + (width * (startEm(j - 1) + textColors[at + j - 1].adv)) / em, b = x0 + (width * startEm(j)) / em;
              // 两个操作合成的一项里，两段之间的真实间距按一个空格估算，位置会差几磅：两边各放宽半个字号
              return vlines.some((g) => (g.x0 + g.x1) / 2 > a - 0.5 * size && (g.x0 + g.x1) / 2 < b + 0.5 * size);
            };
            const cuts = [0, ...[...joined].map((_, j) => j).filter((j) => j > 0 && (keyOf(j) !== keyOf(j - 1) || atRule(j)))];
            if (cuts.length > 1) pieces = cuts.map((j, k) => {
              const next = cuts[k + 1];
              const f0 = k === 0 ? 0 : startEm(j) / em, f1 = next === undefined ? 1 : startEm(next) / em;
              return {
                str: it.str.slice(k === 0 ? 0 : idx[j], next === undefined ? it.str.length : idx[next]),
                x0: x0 + width * f0, x1: x0 + width * f1,
                color: textColors[at + j].color, fakeBold: textColors[at + j].fakeBold, natural: natural! * (f1 - f0),
                clipped: hiddenAt[j],
              };
            }).filter((pc) => !pc.clipped);
          }
        }
      } else if (textColors[ci]) {
        color = textColors[ci].color;
      }
    }

    if (clippedOut) continue;
    if (!pieces.length) pieces = [{ str: it.str, x0, x1: x0 + width, color, fakeBold, natural }];
    // 扫描图上的 OCR 识别层：看不见的文字压在整页扫描图上，转成正文会与图里的字重叠成两层；不转换，只计数
    if (hidden) {
      const cx = x0 + width / 2, cy = baseline - 0.3 * size;
      if (model.images.some((im) => cx >= im.x0 && cx <= im.x1 && cy >= im.y0 && cy <= im.y1)) {
        model.ocrChars = (model.ocrChars ?? 0) + it.str.replace(/\s/g, "").length;
        continue;
      }
    }
    for (const pc of pieces) {
      const sym = mapSymbolFont(expandLigatures((spaced ? pc.str.replace(/ /g, "") : pc.str).replace(/\u00ad/g, "")), fi.family);
      if (!sym.text.trim()) continue;
      const bold = fi.bold || pc.fakeBold, w = pc.x1 - pc.x0, nat = pc.natural;
      const wordW = measureWord(sym.text, sym.font, bold, fi.italic, Math.round(size * 2) / 2) ?? undefined;
      const chars = Math.max(1, sym.text.length);
      // 等宽字体被替换成更宽的 Courier New（TeX 的打字机字体 CMTT / SFTT 每字 0.525 em，Courier New 是 0.6 em）：
      // 按差值收紧字符间距，使文字仍占原来的宽度，否则并排的作者邮箱、居中的网址会互相挤压、折到下一行。
      // 等宽字体每个字的差值相同，同字号的片段收紧量一致，不会把文字切成碎片；比例字体由段落的不折行校验处理
      const fit = sym.font === "Courier New" && nat && wordW && !spaced && wordW > 1.04 * Math.min(w, nat)
        ? Math.max(-0.12 * size, (Math.min(w, nat) - wordW) / chars) : undefined;
      model.spans.push({
        text: sym.text,
        x0: pc.x0, x1: pc.x1,
        top: baseline - ascent * size,
        bottom: baseline - descent * size,
        baseline,
        size: Math.round(size * 2) / 2,
        font: sym.font,
        bold,
        italic: fi.italic,
        color: pc.color,
        natural: nat,
        spaceW: (spaceEm.get(it.fontName) ?? 0.25) * size,
        rawFont: rawNames.get(it.fontName),
        math: fi.math,
        type3: isType3(it.fontName) || undefined,
        wordW,
        wordSpaceW: measureWord(" ", sym.font, bold, fi.italic, Math.round(size * 2) / 2) ?? undefined,
        smallCaps: smallCaps || (isSmallCapsFont(rawNames.get(it.fontName)) && /[a-z]/.test(sym.text) && !/[A-Z]/.test(sym.text)) || undefined,
        // 字距加宽 / 紧缩：pdf.js 插了空格的逐字母排列，或单个词（不受两端对齐的词间距影响）的实际宽度明显偏离自然宽度
        tracking: nat && (spaced || (!/\s/.test(sym.text) && sym.text.length >= 3 && Math.abs(w / nat - 1) > 0.12))
          ? (w - (wordW && wordW > 1.04 * nat ? wordW : nat)) / chars : fit,
      });
    }
  }
  model.spans = mergeLetterSpaced(model.spans);
  // 裁剪框（CropBox）之外的东西看不见：整个落在页面外的文字、图片、线条、色块不转换。
  // 视口已按裁剪框换算，页面就是 [0, 宽] × [0, 高]；以前它们按负坐标排进版面，把页面上的内容整块往下挤，一页变三页
  {
    const W = model.width, H = model.height;
    const on = (x0: number, y0: number, x1: number, y1: number, pad = 0) => x1 + pad > 0.5 && x0 - pad < W - 0.5 && y1 + pad > 0.5 && y0 - pad < H - 0.5;
    model.spans = model.spans.filter((s) => on(s.x0, s.top, s.x1, s.bottom));
    model.images = model.images.filter((b) => on(b.x0, b.y0, b.x1, b.y1));
    model.fills = model.fills.filter((b) => on(b.x0, b.y0, b.x1, b.y1));
    model.segs = model.segs.filter((g) => on(Math.min(g.x0, g.x1), Math.min(g.y0, g.y1), Math.max(g.x0, g.x1), Math.max(g.y0, g.y1), g.width / 2));
    model.masks = model.masks.filter((b) => on(b.x0, b.y0, b.x1, b.y1));
    const gone = model.paths.filter((b) => !on(b.x0, b.y0, b.x1, b.y1));
    if (gone.length) {
      model.paths = model.paths.filter((b) => on(b.x0, b.y0, b.x1, b.y1));
      model.vectorShapes = Math.max(0, model.vectorShapes - gone.filter((p) => (p.curved || p.diagonal) && (p.x1 - p.x0) * (p.y1 - p.y0) > 4).length);
    }
    if (model.rotated) model.rotated = model.rotated.filter((b) => on(b.x0, b.y0, b.x1, b.y1));
  }
  if (scanned) {
    const split = splitScan(scanned.gray, scanned.box, model.spans.map((sp) => ({ x0: sp.x0, y0: sp.top, x1: sp.x1, y1: sp.bottom, text: sp.text, baseline: sp.baseline, size: sp.size })));
    if (split) model.scan = { ...split, image: scanned.image };
  }
  return model;
}

/**
 * 逐字母定位的加宽字距标题（"A R T I C L E  I N F O"）：PDF 把每个字母画成独立的文字项，
 * 字母之间的空隙就是字距。≥ 4 个同字体、同字号、同基线的单字符文字项、间隙规则 → 合成词，加宽由字符间距体现；
 * 间隙明显更大（> 1.7 倍字母间距）处是词间空格
 */
function mergeLetterSpaced(spans: Span[]): Span[] {
  const idx = spans.map((_, i) => i).filter((i) => spans[i].text.trim().length === 1 && !spans[i].img && !spans[i].math);
  idx.sort((a, b) => spans[a].baseline - spans[b].baseline || spans[a].x0 - spans[b].x0);
  const drop = new Set<number>(), add: Span[] = [];
  for (let a = 0; a < idx.length; ) {
    let b = a + 1;
    const f = spans[idx[a]];
    while (b < idx.length) {
      const p = spans[idx[b - 1]], q = spans[idx[b]];
      const gap = q.x0 - p.x1;
      if (Math.abs(q.baseline - f.baseline) > 0.3 || q.font !== f.font || q.size !== f.size || q.bold !== f.bold || q.italic !== f.italic || q.color !== f.color || gap < 0.08 * f.size || gap > 1.2 * f.size) break;
      // 两个字母之间夹着别的文字项（R₁¹, R₂¹ 的上下标、逗号）：是公式里各自独立的字母，不是加宽字距
      if (spans.some((o) => o !== p && o !== q && o.text.trim() && o.x0 >= p.x1 - 0.3 && o.x1 <= q.x0 + 0.3 && Math.abs(o.baseline - f.baseline) <= 0.8 * f.size)) break;
      b++;
    }
    const run = idx.slice(a, b);
    a = b > a + 1 ? b : a + 1;
    if (run.length < 4) continue;
    const gaps = run.slice(1).map((k, i) => spans[k].x0 - spans[run[i]].x1);
    const med = [...gaps].sort((x, y) => x - y)[Math.floor(gaps.length / 2)];
    // 字母间距必须规则：绝大多数间隙与中位数相差不大（排除"每个字母恰好分开"的偶然情形）
    if (gaps.filter((g) => g <= 1.7 * med + 0.3 && Math.abs(g - med) <= 0.35 * Math.max(med, 0.5)).length < 0.6 * gaps.length) continue;
    const words: number[][] = [[run[0]]];
    run.slice(1).forEach((k, i) => (gaps[i] > 1.7 * med + 0.3 ? words.push([k]) : words[words.length - 1].push(k)));
    for (const w of words) {
      const ss = w.map((k) => spans[k]);
      const first = ss[0], last = ss[ss.length - 1];
      const natural = ss.every((x) => x.natural) ? ss.reduce((n, x) => n + x.natural!, 0) : undefined;
      add.push({ ...first, text: ss.map((x) => x.text).join(""), x1: last.x1, natural, wordW: ss.every((x) => x.wordW) ? ss.reduce((n, x) => n + x.wordW!, 0) : undefined,
        tracking: ss.length > 1 && natural ? (last.x1 - first.x0 - natural) / ss.length : undefined, top: Math.min(...ss.map((x) => x.top)), bottom: Math.max(...ss.map((x) => x.bottom)),
        smallCaps: ss.every((x) => x.smallCaps) || undefined });
    }
    for (const k of run) drop.add(k);
  }
  return drop.size ? [...spans.filter((_, i) => !drop.has(i)), ...add] : spans;
}
