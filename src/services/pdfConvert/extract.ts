/**
 * PDF → 版面元素提取（pdf.js）
 *
 * 输出每页的：文字片段（位置/字体/字号/粗斜体/颜色）、水平与垂直线段（表格边框、下划线、分隔线）、
 * 填充矩形（底纹）、位图（编码为 PNG）。坐标统一为"页面左上角为原点、单位 pt、y 向下"。
 */

import zlib from "node:zlib";
import { crc32 } from "../../documents/docx/zip.js";
import { symbolChar } from "../../documents/docx/symbolFont.js";
import { parseFontName, measureWord, type FontInfo } from "./fonts.js";

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
  spans: Span[];
  segs: Seg[];
  fills: FillRect[];
  images: ImageBox[];
  /** 无法转换的矢量图形（曲线等）数量 */
  vectorShapes: number;
  /** 全部已绘制的路径（含曲线、斜线） */
  paths: PathBox[];
  rotatedText: number;
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

function imageToPng(img: any): { png: Buffer; w: number; h: number } | null {
  if (!img || !img.width || !img.height || !img.data) return null;
  const { width: w, height: h, data, kind } = img;
  if (kind === 3 || data.length === w * h * 4) return { png: encodePng(w, h, data, 4), w, h };
  if (kind === 2 || data.length === w * h * 3) return { png: encodePng(w, h, data, 3), w, h };
  if (kind === 1) {
    // 1 位灰度，按行打包
    const rowBytes = Math.ceil(w / 8);
    const out = new Uint8Array(w * h);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) out[y * w + x] = data[y * rowBytes + (x >> 3)] & (0x80 >> (x & 7)) ? 255 : 0;
    return { png: encodePng(w, h, out, 1), w, h };
  }
  return null;
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

export async function extractPdf(buf: Buffer, opts: { maxPages?: number } = {}): Promise<PageModel[]> {
  const pdfjs = await loadPdfjs();
  const { OPS, Util } = pdfjs;
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf), isOffscreenCanvasSupported: false, disableFontFace: true, useSystemFonts: false, verbosity: 0 }).promise;
  const pages: PageModel[] = [];
  const n = Math.min(doc.numPages, opts.maxPages ?? 500);
  try {
    for (let p = 1; p <= n; p++) pages.push(await extractPage(await doc.getPage(p), p - 1, OPS, Util));
  } finally {
    await doc.destroy().catch(() => {});
  }
  return pages;
}

async function extractPage(page: any, index: number, OPS: any, Util: any): Promise<PageModel> {
  const vp = page.getViewport({ scale: 1 });
  const VT = vp.transform as M;
  const ops = await page.getOperatorList();
  const model: PageModel = { index, width: vp.width, height: vp.height, spans: [], segs: [], fills: [], images: [], vectorShapes: 0, rotatedText: 0, paths: [] };

  // ---------- 操作符遍历：图形、图片、文字颜色 ----------
  interface GState { ctm: M; fill: string; stroke: string; lw: number; trm: number }
  let gs: GState = { ctm: [1, 0, 0, 1, 0, 0], fill: "000000", stroke: "000000", lw: 1, trm: 0 };
  const stack: GState[] = [];
  const textColors: Array<{ ch: string; color: string; fakeBold: boolean; adv: number; spaceBefore: number }> = [];
  const spaceEm = new Map<string, number>(); // 各字体空格的字宽（em）
  let curFont = "";
  let pendingPath: Array<{ kind: "rect"; pts: Array<[number, number]> } | { kind: "poly"; pts: Array<[number, number]>; ctrl: Array<[number, number]>; curved: boolean; closed: boolean }> = [];

  const toPage = (x: number, y: number) => apply(VT, ...apply(gs.ctm, x, y));
  const scale = () => Math.sqrt(Math.abs(gs.ctm[0] * gs.ctm[3] - gs.ctm[1] * gs.ctm[2])) || 1;

  const paint = (stroke: boolean, fill: boolean) => {
    for (const sp of pendingPath) {
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
          model.segs.push({ x0: Math.min(ax, bx), y0: Math.min(ay, by), x1: Math.max(ax, bx), y1: Math.max(ay, by), width: lw, color: gs.stroke, horizontal: horiz });
        }
      }
    }
    pendingPath = [];
  };

  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i];
    const args = ops.argsArray[i];
    switch (fn) {
      case OPS.save: stack.push({ ...gs, ctm: [...gs.ctm] as M }); break;
      case OPS.restore: gs = stack.pop() ?? gs; break;
      case OPS.transform: gs.ctm = mul(gs.ctm, args as M); break;
      case OPS.setFillRGBColor: gs.fill = hex(args); break;
      case OPS.setStrokeRGBColor: gs.stroke = hex(args); break;
      case OPS.setFillGray: gs.fill = hex([args[0] * 255, args[0] * 255, args[0] * 255]); break;
      case OPS.setStrokeGray: gs.stroke = hex([args[0] * 255, args[0] * 255, args[0] * 255]); break;
      case OPS.setLineWidth: gs.lw = args[0]; break;
      case OPS.setTextRenderingMode: gs.trm = args[0]; break;
      case OPS.setFont: curFont = String(args[0]); break;
      case OPS.showText:
      case OPS.showSpacedText: {
        const glyphs = args[0] ?? [];
        let pendingSpace = 0;
        for (const g of glyphs) {
          if (!g || typeof g !== "object") continue;
          const w = typeof g.width === "number" ? g.width / 1000 : 0;
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
            textColors.push({ ch: u[k], color: gs.trm === 1 ? gs.stroke : gs.fill, fakeBold: gs.trm === 2, adv: k === 0 ? w : 0, spaceBefore: k === 0 ? pendingSpace : 0 });
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
      case OPS.fill: case OPS.eoFill: paint(false, true); break;
      case OPS.fillStroke: case OPS.eoFillStroke: case OPS.closeFillStroke: case OPS.closeEOFillStroke: paint(true, true); break;
      case OPS.endPath: pendingPath = []; break;
      case OPS.paintImageXObject:
      case OPS.paintInlineImageXObject: {
        const img = fn === OPS.paintInlineImageXObject ? args[0] : await getObj(String(args[0]).startsWith("g_") ? page.commonObjs : page.objs, args[0]);
        const png = imageToPng(img);
        const corners = [toPage(0, 0), toPage(1, 0), toPage(0, 1), toPage(1, 1)];
        const xs = corners.map((c) => c[0]), ys = corners.map((c) => c[1]);
        const box = { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
        if (png && box.x1 - box.x0 > 2 && box.y1 - box.y0 > 2) model.images.push({ ...box, png: png.png, pxW: png.w, pxH: png.h });
        else if (!png) model.vectorShapes++;
        break;
      }
    }
  }

  // ---------- 文字 ----------
  // 不做 Unicode 规范化，保证与操作符流中的字形一一对应；连字（ﬁ ﬂ）单独展开
  const tc = await page.getTextContent({ includeMarkedContent: false, disableNormalization: true });
  const fontInfo = new Map<string, FontInfo>();
  const rawNames = new Map<string, string>();
  const infoOf = (fontName: string): FontInfo => {
    if (!fontInfo.has(fontName)) {
      let raw: string | undefined;
      try { raw = page.commonObjs.has(fontName) ? page.commonObjs.get(fontName)?.name : undefined; } catch { /* 忽略 */ }
      rawNames.set(fontName, (raw ?? "").replace(/^[A-Z]{6}\+/, ""));
      fontInfo.set(fontName, parseFontName(raw ?? tc.styles[fontName]?.fontFamily));
    }
    return fontInfo.get(fontName)!;
  };


  let ci = 0; // textColors 游标
  const opText = textColors.map((t) => t.ch).join("");
  for (const it of tc.items as any[]) {
    if (!it.str || !it.str.trim()) continue;
    const t = Util.transform(VT, it.transform) as M;
    const size = Math.hypot(t[2], t[3]);
    if (size < 0.5) continue;
    if (Math.abs(t[1]) > 0.01 * size || Math.abs(t[2]) > 0.01 * size) {
      model.rotatedText++;
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
    let color = "000000", fakeBold = false;
    let natural: number | undefined;
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
        // 逐字核对并累计自然字宽
        let pos = at, em = 0, spaces = 0, matched = 0;
        for (const ch of joined) {
          if (opText[pos] !== ch) break;
          em += textColors[pos].adv + (pos > at ? textColors[pos].spaceBefore : 0);
          if (pos > at && textColors[pos].spaceBefore > 0) spaces++;
          pos++;
          matched++;
        }
        ci = pos;
        if (matched === joined.length && em > 0) {
          // 用 TJ 位移代替空格字形的 PDF：按字体空格宽度补上词间空格
          const inner = (it.str.trim().match(/\s+/g) ?? []).length;
          if (spaces < inner) em += (inner - spaces) * (spaceEm.get(it.fontName) ?? 0.25);
          natural = em * size;
        }
      } else if (textColors[ci]) {
        color = textColors[ci].color;
      }
    }

    const sym = mapSymbolFont(expandLigatures(it.str.replace(/\u00ad/g, "")), fi.family);
    model.spans.push({
      text: sym.text,
      x0, x1: x0 + width,
      top: baseline - ascent * size,
      bottom: baseline - descent * size,
      baseline,
      size: Math.round(size * 2) / 2,
      font: sym.font,
      bold: fi.bold || fakeBold,
      italic: fi.italic,
      color,
      natural,
      spaceW: (spaceEm.get(it.fontName) ?? 0.25) * size,
      rawFont: rawNames.get(it.fontName),
      math: fi.math,
      wordW: measureWord(sym.text, sym.font, fi.bold || fakeBold, fi.italic, Math.round(size * 2) / 2) ?? undefined,
      wordSpaceW: measureWord(" ", sym.font, fi.bold || fakeBold, fi.italic, Math.round(size * 2) / 2) ?? undefined,
    });
  }
  return model;
}
