/**
 * 形状组合 / 原生图表的 PNG 后备图
 *
 * Word、WPS、LibreOffice 直接显示可编辑的形状和图表；浏览器内的快速预览（docx-preview）不认识它们，
 * 会显示这里画出的后备图片。用 Canvas 2D 直接绘制（字体按系统可用的中文字体回退），
 * @napi-rs/canvas 不可用时返回 undefined，文档里只保留可编辑对象。
 */

import type { ShapeSpec, ChartSpec } from "../documents/docx/build.js";
import { canvasModule } from "./pdfConvert/render.js";

const FONT = `"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", "Noto Sans CJK JP", "Source Han Sans SC", "WenQuanYi Zen Hei", SimHei, sans-serif`;
const PALETTE = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E"];
const K = 2; // 每磅 2 像素（约 144 dpi）
const col = (c: string | undefined, d: string) => (c && c !== "none" ? `#${c.replace(/^#/, "")}` : d);

export async function shapesPng(shapes: ShapeSpec[], width: number, height: number): Promise<Buffer | undefined> {
  const m = await canvasModule();
  if (!m) return undefined;
  try {
    const c = m.createCanvas(Math.max(1, Math.ceil(width * K)), Math.max(1, Math.ceil(height * K)));
    const g = c.getContext("2d");
    g.scale(K, K);
    for (const s of shapes) {
      const isLine = s.type === "line" || s.type === "arrow" || s.type === "doubleArrow";
      g.save();
      g.lineWidth = s.stroke_width ?? (isLine ? 1.25 : 1);
      g.strokeStyle = col(s.stroke, isLine ? "#404040" : "#2F5597");
      g.setLineDash(s.dash === "dash" ? [5, 3] : s.dash === "dot" ? [1.5, 2] : []);
      if (isLine) {
        const x1 = s.x, y1 = s.y, x2 = s.x2 ?? s.x + (s.w ?? 60), y2 = s.y2 ?? s.y + (s.h ?? 0);
        g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke();
        g.setLineDash([]);
        const head = (fx: number, fy: number, tx: number, ty: number) => {
          const a = Math.atan2(ty - fy, tx - fx), L = 6;
          g.fillStyle = g.strokeStyle;
          g.beginPath(); g.moveTo(tx, ty); g.lineTo(tx - L * Math.cos(a - 0.4), ty - L * Math.sin(a - 0.4)); g.lineTo(tx - L * Math.cos(a + 0.4), ty - L * Math.sin(a + 0.4)); g.closePath(); g.fill();
        };
        if (s.type !== "line") head(x1, y1, x2, y2);
        if (s.type === "doubleArrow") head(x2, y2, x1, y1);
        g.restore();
        continue;
      }
      const x = s.x, y = s.y, w = s.w ?? 80, h = s.h ?? 40;
      g.beginPath();
      switch (s.type) {
        case "ellipse": case "cloud": g.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2); break;
        case "diamond": g.moveTo(x + w / 2, y); g.lineTo(x + w, y + h / 2); g.lineTo(x + w / 2, y + h); g.lineTo(x, y + h / 2); g.closePath(); break;
        case "triangle": g.moveTo(x + w / 2, y); g.lineTo(x + w, y + h); g.lineTo(x, y + h); g.closePath(); break;
        case "parallelogram": g.moveTo(x + w * 0.2, y); g.lineTo(x + w, y); g.lineTo(x + w * 0.8, y + h); g.lineTo(x, y + h); g.closePath(); break;
        case "hexagon": g.moveTo(x + w * 0.25, y); g.lineTo(x + w * 0.75, y); g.lineTo(x + w, y + h / 2); g.lineTo(x + w * 0.75, y + h); g.lineTo(x + w * 0.25, y + h); g.lineTo(x, y + h / 2); g.closePath(); break;
        case "roundRect": case "cylinder": g.roundRect(x, y, w, h, Math.min(w, h) * (s.type === "cylinder" ? 0.3 : 0.18)); break;
        default: g.rect(x, y, w, h);
      }
      if (s.type !== "text" && s.fill !== "none") { g.fillStyle = col(s.fill, "#DEEBF7"); g.fill(); }
      if (s.type !== "text" && s.stroke !== "none") g.stroke();
      if (s.text) {
        const fs = s.font_size ?? 10.5;
        g.font = `${s.bold ? "bold " : ""}${fs}px ${FONT}`;
        g.fillStyle = col(s.font_color, "#1F1F1F");
        g.textBaseline = "middle";
        g.textAlign = s.align === "left" ? "left" : s.align === "right" ? "right" : "center";
        const tx = s.align === "left" ? x + 4 : s.align === "right" ? x + w - 4 : x + w / 2;
        const lines = s.text.split("\n");
        lines.forEach((l, i) => g.fillText(l, tx, y + h / 2 + (i - (lines.length - 1) / 2) * fs * 1.3));
      }
      g.restore();
    }
    return c.toBuffer("image/png");
  } catch {
    return undefined;
  }
}

export async function chartPng(spec: ChartSpec, widthPt: number, heightPt: number): Promise<Buffer | undefined> {
  const m = await canvasModule();
  if (!m) return undefined;
  try {
    const W = widthPt, H = heightPt;
    const c = m.createCanvas(Math.ceil(W * K), Math.ceil(H * K));
    const g = c.getContext("2d");
    g.scale(K, K);
    g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
    g.strokeStyle = "#D9D9D9"; g.lineWidth = 0.75; g.strokeRect(0.5, 0.5, W - 1, H - 1);
    const text = (t: string, x: number, y: number, size: number, align: CanvasTextAlign = "center", bold = false, color = "#404040") => {
      g.font = `${bold ? "bold " : ""}${size}px ${FONT}`; g.fillStyle = color; g.textAlign = align; g.textBaseline = "middle"; g.fillText(t, x, y);
    };
    const top = spec.title ? 26 : 10;
    if (spec.title) text(spec.title, W / 2, 14, 11, "center", true, "#1F1F1F");
    const pie = spec.type === "pie" || spec.type === "doughnut";
    const legendNames = pie ? spec.categories : spec.series.length > 1 ? spec.series.map((s) => s.name) : [];
    const legendH = legendNames.length && spec.legend !== "none" ? 16 : 0;
    if (legendH) {
      const item = Math.min(90, (W - 20) / legendNames.length);
      const start = W / 2 - (legendNames.length * item) / 2;
      legendNames.forEach((n, i) => { g.fillStyle = PALETTE[i % PALETTE.length]; g.fillRect(start + i * item, H - 13, 7, 7); text(n, start + i * item + 10, H - 9.5, 7.5, "left"); });
    }
    if (pie) {
      const vals = spec.series[0].values.map((v) => Math.max(0, v));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = W / 2, cy = top + (H - top - legendH) / 2, r = Math.max(5, Math.min(W, H - top - legendH) / 2 - 8);
      let a0 = -Math.PI / 2;
      vals.forEach((v, i) => {
        const a1 = a0 + (v / total) * Math.PI * 2;
        g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, r, a0, a1); g.closePath();
        g.fillStyle = PALETTE[i % PALETTE.length]; g.fill(); g.strokeStyle = "#fff"; g.lineWidth = 1; g.stroke();
        if (spec.show_values) { const am = (a0 + a1) / 2; text(`${Math.round((v / total) * 1000) / 10}%`, cx + r * 0.65 * Math.cos(am), cy + r * 0.65 * Math.sin(am), 8, "center", false, "#fff"); }
        a0 = a1;
      });
      if (spec.type === "doughnut") { g.beginPath(); g.arc(cx, cy, r * 0.55, 0, Math.PI * 2); g.fillStyle = "#fff"; g.fill(); }
      return c.toBuffer("image/png");
    }
    const left = 34, right = 10, bottom = 20 + legendH;
    const pw = W - left - right, ph = H - top - bottom;
    const all = spec.series.flatMap((s) => s.values);
    const max = Math.max(0, ...all), min = Math.min(0, ...all);
    const span = max - min || 1;
    const horiz = spec.type === "bar";
    const n = spec.categories.length;
    const band = (horiz ? ph : pw) / Math.max(1, n);
    const vpos = (v: number) => (horiz ? left + ((v - min) / span) * pw : top + ph - ((v - min) / span) * ph);
    g.lineWidth = 0.5;
    for (let k = 0; k <= 4; k++) {
      const v = min + (span * k) / 4, p = vpos(v);
      g.strokeStyle = "#E5E5E5"; g.beginPath();
      if (horiz) { g.moveTo(p, top); g.lineTo(p, top + ph); } else { g.moveTo(left, p); g.lineTo(left + pw, p); }
      g.stroke();
      const label = String(Math.round(v * 100) / 100);
      if (horiz) text(label, p, top + ph + 8, 7); else text(label, left - 4, p, 7, "right");
    }
    spec.categories.forEach((cat, i) => (horiz ? text(cat, left - 4, top + band * (i + 0.5), 7, "right") : text(cat, left + band * (i + 0.5), top + ph + 9, 7)));
    if (spec.type === "bar" || spec.type === "column") {
      const ms = spec.series.length;
      const bw = (band * 0.7) / (spec.stacked ? 1 : ms);
      const acc = new Array(n).fill(0);
      spec.series.forEach((s, k) => s.values.forEach((v, i) => {
        g.fillStyle = PALETTE[k % PALETTE.length];
        const base = spec.stacked ? acc[i] : 0;
        const a = vpos(base), b = vpos(base + v);
        const off = band * i + band * 0.15 + (spec.stacked ? 0 : bw * k);
        if (horiz) g.fillRect(Math.min(a, b), top + off, Math.abs(b - a), bw);
        else g.fillRect(left + off, Math.min(a, b), bw, Math.abs(b - a));
        if (spec.show_values) horiz ? text(String(v), b + 8, top + off + bw / 2, 7) : text(String(v), left + off + bw / 2, b - 6, 7);
        if (spec.stacked) acc[i] += v;
      }));
    } else {
      spec.series.forEach((s, k) => {
        const pts = s.values.map((v, i) => [left + band * (i + 0.5), vpos(v)] as const);
        g.strokeStyle = g.fillStyle = PALETTE[k % PALETTE.length];
        if (spec.type === "area") {
          g.globalAlpha = 0.45; g.beginPath(); g.moveTo(pts[0][0], vpos(0)); pts.forEach(([x, y]) => g.lineTo(x, y)); g.lineTo(pts[pts.length - 1][0], vpos(0)); g.closePath(); g.fill(); g.globalAlpha = 1;
        }
        if (spec.type !== "scatter") { g.lineWidth = 1.8; g.beginPath(); pts.forEach(([x, y], i) => (i ? g.lineTo(x, y) : g.moveTo(x, y))); g.stroke(); }
        pts.forEach(([x, y]) => { g.beginPath(); g.arc(x, y, 2.4, 0, Math.PI * 2); g.fill(); });
      });
    }
    g.strokeStyle = "#7F7F7F"; g.lineWidth = 0.75; g.beginPath();
    if (horiz) { g.moveTo(left, top); g.lineTo(left, top + ph); } else { g.moveTo(left, top + ph); g.lineTo(left + pw, top + ph); }
    g.stroke();
    return c.toBuffer("image/png");
  } catch {
    return undefined;
  }
}
