/**
 * 按区域渲染 PDF 页面（pdf.js + @napi-rs/canvas）
 *
 * 用于矢量图、行间公式、首字下沉：以 288 dpi 渲染整页（透明背景），再裁出各区域编码为 PNG。
 * @napi-rs/canvas 是预编译的原生模块（Windows / macOS / Linux 均有），加载失败时 rasterAvailable() 返回 false，
 * 转换自动退回"不渲染"的模式（这些区域按文字处理）。
 */

import type { ImageBox } from "./extract.js";

let napi: any | null | undefined;

async function loadCanvas(): Promise<any | null> {
  if (napi !== undefined) return napi;
  try {
    const mod: any = await import("@napi-rs/canvas");
    const m = mod.createCanvas ? mod : mod.default;
    m.createCanvas(2, 2).getContext("2d");
    // 与画布同一实例的 Path2D / DOMMatrix（pdf.js 自带的兜底可能加载了另一份原生模块，画布不认）
    const g = globalThis as any;
    g.DOMMatrix = m.DOMMatrix;
    g.Path2D = m.Path2D;
    g.ImageData = m.ImageData;
    napi = m;
  } catch {
    napi = null;
  }
  return napi;
}

/** 供其他模块使用的画布模块（不可用时为 null） */
export async function canvasModule(): Promise<any | null> {
  return loadCanvas();
}

export async function rasterAvailable(): Promise<boolean> {
  if (process.env.PDF_RASTER === "0") return false;
  return !!(await loadCanvas());
}

/** 渲染所有待渲染的区域，填入 png / pxW / pxH；返回成功渲染的数量 */
export async function renderRegions(buf: Buffer, boxes: ImageBox[], dpi = 288): Promise<number> {
  const todo = boxes.filter((b) => b.render && !b.png.length);
  if (!todo.length) return 0;
  const m = await loadCanvas();
  if (!m) return 0;
  class CanvasFactory {
    create(w: number, h: number) {
      const canvas = m.createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h)));
      return { canvas, context: canvas.getContext("2d") };
    }
    reset(cc: any, w: number, h: number) { cc.canvas.width = Math.max(1, Math.ceil(w)); cc.canvas.height = Math.max(1, Math.ceil(h)); }
    destroy(cc: any) { cc.canvas.width = 0; cc.canvas.height = 0; cc.canvas = null; cc.context = null; }
  }
  const pdfjs: any = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(buf), CanvasFactory, disableFontFace: true, useSystemFonts: false,
    isOffscreenCanvasSupported: false, verbosity: 0,
  }).promise;
  const scale = dpi / 72;
  let done = 0;
  try {
    const byPage = new Map<number, ImageBox[]>();
    for (const b of todo) {
      const k = b.render!.page;
      if (!byPage.has(k)) byPage.set(k, []);
      byPage.get(k)!.push(b);
    }
    for (const [pi, list] of byPage) {
      try {
        const page = await doc.getPage(pi + 1);
        const vp = page.getViewport({ scale });
        const f = new CanvasFactory();
        const cc = f.create(vp.width, vp.height);
        // 透明背景：贴回 Word 后不会遮住相邻的文字
        await page.render({ canvasContext: cc.context, viewport: vp, background: "rgba(0,0,0,0)" }).promise;
        for (const b of list) {
          const x = Math.max(0, Math.floor(b.x0 * scale)), y = Math.max(0, Math.floor(b.y0 * scale));
          const w = Math.min(cc.canvas.width - x, Math.ceil((b.x1 - b.x0) * scale)), h = Math.min(cc.canvas.height - y, Math.ceil((b.y1 - b.y0) * scale));
          if (w < 2 || h < 2) continue;
          const out = m.createCanvas(w, h);
          out.getContext("2d").drawImage(cc.canvas, x, y, w, h, 0, 0, w, h);
          b.png = out.toBuffer("image/png");
          b.pxW = w; b.pxH = h;
          done++;
        }
        f.destroy(cc);
        page.cleanup();
      } catch (e) { if (process.env.DEBUG_RENDER) console.error(e); /* 这一页渲染失败：其区域保持空白（写出时跳过） */ }
    }
  } finally {
    await doc.destroy().catch(() => {});
  }
  return done;
}

/** 把 SVG / WebP 等 Word 不能直接嵌入的图片转成 PNG（需要 @napi-rs/canvas；不可用时返回 null） */
export async function convertToPng(data: Buffer, maxPx = 2400): Promise<{ png: Buffer; width: number; height: number } | null> {
  const m = await loadCanvas();
  if (!m) return null;
  try {
    const img = await m.loadImage(data);
    let w = img.width || 800, h = img.height || 600;
    // 矢量图（SVG）按 2 倍分辨率栅格化，保证打印清晰
    const isSvg = /^\s*(<\?xml|<svg)/i.test(data.subarray(0, 200).toString("utf8"));
    const k = Math.min(isSvg ? 2 : 1, maxPx / Math.max(w, h));
    const cw = Math.max(1, Math.round(w * k)), ch = Math.max(1, Math.round(h * k));
    const c = m.createCanvas(cw, ch);
    c.getContext("2d").drawImage(img, 0, 0, cw, ch);
    return { png: c.toBuffer("image/png"), width: w, height: h };
  } catch {
    return null;
  }
}
