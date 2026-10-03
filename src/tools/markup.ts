/**
 * Markdown / HTML 的插入标记生成：表格、图片（data URI）、SVG 图形与图表、公式、目录
 *
 * Word 文档用原生对象（DocxDocument 里实现）；Markdown 与 HTML 没有"对象"，这里生成等价的源码片段，
 * 由适配器的 insertRaw 原样拼接进源文件。
 */

import type { ShapeSpec, ChartSpec } from "../documents/docx/build.js";

const escHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const escCell = (s: string) => s.replace(/\|/g, "\\|").replace(/\r?\n/g, "<br>");

export function tableMarkdown(rows: string[][], header = true, caption?: string): string {
  const n = Math.max(...rows.map((r) => r.length));
  const fill = (r: string[]) => Array.from({ length: n }, (_, i) => escCell(String(r[i] ?? "")));
  const head = header ? fill(rows[0]) : Array.from({ length: n }, () => " ");
  const body = header ? rows.slice(1) : rows;
  const lines = [`| ${head.join(" | ")} |`, `| ${head.map(() => "---").join(" | ")} |`, ...body.map((r) => `| ${fill(r).join(" | ")} |`)];
  return (caption ? `**${caption}**\n\n` : "") + lines.join("\n");
}

export function tableHtml(rows: string[][], header = true, caption?: string, style = "grid"): string {
  const n = Math.max(...rows.map((r) => r.length));
  const border = style === "plain" ? "" : style === "three_line" ? ' style="border-collapse:collapse;border-top:2px solid #000;border-bottom:2px solid #000"' : ' border="1" style="border-collapse:collapse"';
  const cell = (tag: string, v: string) => `<${tag} style="padding:4px 8px${style === "three_line" && tag === "th" ? ";border-bottom:1px solid #000" : ""}">${escHtml(v).replace(/\r?\n/g, "<br>")}</${tag}>`;
  const row = (r: string[], tag: string) => `<tr>${Array.from({ length: n }, (_, i) => cell(tag, String(r[i] ?? ""))).join("")}</tr>`;
  const head = header ? `<thead>${row(rows[0], "th")}</thead>` : "";
  const body = `<tbody>${(header ? rows.slice(1) : rows).map((r) => row(r, "td")).join("")}</tbody>`;
  return `<table${border}>${caption ? `<caption>${escHtml(caption)}</caption>` : ""}${head}${body}</table>`;
}

export function imageMarkup(format: "markdown" | "html", dataUri: string, alt: string, widthPt?: number, caption?: string): string {
  const wpx = widthPt ? Math.round(widthPt / 0.75) : undefined;
  if (format === "markdown") {
    const img = wpx ? `<img src="${dataUri}" alt="${escHtml(alt)}" width="${wpx}">` : `![${alt.replace(/[[\]]/g, "")}](${dataUri})`;
    return caption ? `${img}\n\n*${caption}*` : img;
  }
  const img = `<img src="${dataUri}" alt="${escHtml(alt)}"${wpx ? ` width="${wpx}"` : ""}>`;
  return caption ? `<figure>${img}<figcaption>${escHtml(caption)}</figcaption></figure>` : `<p>${img}</p>`;
}

const PALETTE = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E"];
const FONT = `font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, sans-serif"`;

export function shapesSvg(shapes: ShapeSpec[], width: number, height: number): string {
  const px = (v: number) => Math.round((v / 0.75) * 10) / 10;
  const W = px(width), H = px(height);
  const out: string[] = [`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">`, `<defs><marker id="ah" markerWidth="10" markerHeight="8" refX="9" refY="4" orient="auto"><path d="M0,0 L10,4 L0,8 z" fill="#404040"/></marker></defs>`];
  for (const s of shapes) {
    const stroke = s.stroke === "none" ? "none" : `#${(s.stroke ?? (s.type.includes("rrow") || s.type === "line" ? "404040" : "2F5597")).replace(/^#/, "")}`;
    const fill = s.type === "text" || s.fill === "none" ? "none" : `#${(s.fill ?? "DEEBF7").replace(/^#/, "")}`;
    const sw = (s.stroke_width ?? 1) / 0.75;
    const dash = s.dash === "dash" ? ' stroke-dasharray="6 4"' : s.dash === "dot" ? ' stroke-dasharray="2 3"' : "";
    const x = px(s.x), y = px(s.y), w = px(s.w ?? 80), h = px(s.h ?? 40);
    if (s.type === "line" || s.type === "arrow" || s.type === "doubleArrow") {
      const x2 = px(s.x2 ?? s.x + (s.w ?? 60)), y2 = px(s.y2 ?? s.y + (s.h ?? 0));
      out.push(`<line x1="${x}" y1="${y}" x2="${x2}" y2="${y2}" stroke="${stroke}" stroke-width="${sw}"${dash}${s.type !== "line" ? ' marker-end="url(#ah)"' : ""}${s.type === "doubleArrow" ? ' marker-start="url(#ah)"' : ""}/>`);
      continue;
    }
    const common = `fill="${fill}" stroke="${s.type === "text" ? "none" : stroke}" stroke-width="${sw}"${dash}`;
    if (s.type === "ellipse") out.push(`<ellipse cx="${x + w / 2}" cy="${y + h / 2}" rx="${w / 2}" ry="${h / 2}" ${common}/>`);
    else if (s.type === "diamond") out.push(`<polygon points="${x + w / 2},${y} ${x + w},${y + h / 2} ${x + w / 2},${y + h} ${x},${y + h / 2}" ${common}/>`);
    else if (s.type === "triangle") out.push(`<polygon points="${x + w / 2},${y} ${x + w},${y + h} ${x},${y + h}" ${common}/>`);
    else if (s.type === "parallelogram") out.push(`<polygon points="${x + w * 0.2},${y} ${x + w},${y} ${x + w * 0.8},${y + h} ${x},${y + h}" ${common}/>`);
    else out.push(`<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${s.type === "roundRect" ? Math.min(w, h) * 0.18 : s.type === "cylinder" ? 10 : 0}" ${common}/>`);
    if (s.text) {
      const fs = (s.font_size ?? 10.5) / 0.75;
      const lines = s.text.split("\n");
      const anchor = s.align === "left" ? "start" : s.align === "right" ? "end" : "middle";
      const tx = s.align === "left" ? x + 6 : s.align === "right" ? x + w - 6 : x + w / 2;
      lines.forEach((l, i) => out.push(`<text x="${tx}" y="${y + h / 2 + (i - (lines.length - 1) / 2) * fs * 1.25 + fs * 0.35}" font-size="${fs}" text-anchor="${anchor}" fill="#${(s.font_color ?? "1F1F1F").replace(/^#/, "")}"${s.bold ? ' font-weight="bold"' : ""} ${FONT}>${escHtml(l)}</text>`));
    }
  }
  out.push("</svg>");
  return out.join("");
}

export function chartSvg(spec: ChartSpec, widthPt = 430, heightPt = 260): string {
  const W = Math.round(widthPt / 0.75), H = Math.round(heightPt / 0.75);
  const out: string[] = [`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" ${FONT}>`, `<rect width="${W}" height="${H}" fill="#fff"/>`];
  const top = spec.title ? 36 : 14;
  if (spec.title) out.push(`<text x="${W / 2}" y="24" font-size="16" font-weight="bold" text-anchor="middle">${escHtml(spec.title)}</text>`);
  const legendH = spec.series.length > 1 || spec.type === "pie" || spec.type === "doughnut" ? 24 : 0;
  if (spec.type === "pie" || spec.type === "doughnut") {
    const vals = spec.series[0].values.map((v) => Math.max(0, v));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = W / 2, cy = top + (H - top - legendH) / 2, r = Math.min(W, H - top - legendH) / 2 - 10;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (v / total) * Math.PI * 2;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      out.push(`<path d="M${cx},${cy} L${cx + r * Math.cos(a0)},${cy + r * Math.sin(a0)} A${r},${r} 0 ${large} 1 ${cx + r * Math.cos(a1)},${cy + r * Math.sin(a1)} z" fill="${PALETTE[i % PALETTE.length]}" stroke="#fff"/>`);
      if (spec.show_values) { const am = (a0 + a1) / 2; out.push(`<text x="${cx + r * 0.65 * Math.cos(am)}" y="${cy + r * 0.65 * Math.sin(am)}" font-size="11" text-anchor="middle" fill="#fff">${Math.round((v / total) * 1000) / 10}%</text>`); }
      a0 = a1;
    });
    if (spec.type === "doughnut") out.push(`<circle cx="${cx}" cy="${cy}" r="${r * 0.55}" fill="#fff"/>`);
    out.push(legend(spec.categories, W, H - 10));
    out.push("</svg>");
    return out.join("");
  }
  const left = 48, right = 14, bottom = 30 + legendH;
  const pw = W - left - right, ph = H - top - bottom;
  const all = spec.series.flatMap((s) => s.values);
  const max = Math.max(0, ...all), min = Math.min(0, ...all);
  const span = max - min || 1;
  const y = (v: number) => top + ph - ((v - min) / span) * ph;
  for (let k = 0; k <= 4; k++) {
    const v = min + (span * k) / 4;
    out.push(`<line x1="${left}" y1="${y(v)}" x2="${left + pw}" y2="${y(v)}" stroke="#ddd"/><text x="${left - 6}" y="${y(v) + 4}" font-size="11" text-anchor="end">${Math.round(v * 100) / 100}</text>`);
  }
  const n = spec.categories.length;
  const band = pw / Math.max(1, n);
  spec.categories.forEach((c, i) => out.push(`<text x="${left + band * (i + 0.5)}" y="${top + ph + 16}" font-size="11" text-anchor="middle">${escHtml(c)}</text>`));
  if (spec.type === "bar" || spec.type === "column") {
    const m = spec.series.length;
    const bw = (band * 0.7) / m;
    spec.series.forEach((s, k) => s.values.forEach((v, i) => {
      const x = left + band * i + band * 0.15 + bw * k;
      out.push(`<rect x="${x}" y="${Math.min(y(v), y(0))}" width="${bw}" height="${Math.abs(y(v) - y(0))}" fill="${PALETTE[k % PALETTE.length]}"/>`);
      if (spec.show_values) out.push(`<text x="${x + bw / 2}" y="${y(v) - 4}" font-size="10" text-anchor="middle">${v}</text>`);
    }));
  } else {
    spec.series.forEach((s, k) => {
      const pts = s.values.map((v, i) => `${left + band * (i + 0.5)},${y(v)}`).join(" ");
      if (spec.type === "area") out.push(`<polygon points="${left + band * 0.5},${y(0)} ${pts} ${left + band * (n - 0.5)},${y(0)}" fill="${PALETTE[k % PALETTE.length]}" fill-opacity="0.45"/>`);
      if (spec.type !== "scatter") out.push(`<polyline points="${pts}" fill="none" stroke="${PALETTE[k % PALETTE.length]}" stroke-width="2.5"/>`);
      s.values.forEach((v, i) => out.push(`<circle cx="${left + band * (i + 0.5)}" cy="${y(v)}" r="3.5" fill="${PALETTE[k % PALETTE.length]}"/>`));
    });
  }
  out.push(`<line x1="${left}" y1="${top + ph}" x2="${left + pw}" y2="${top + ph}" stroke="#666"/>`);
  if (legendH) out.push(legend(spec.series.map((s) => s.name), W, H - 8));
  out.push("</svg>");
  return out.join("");
}

function legend(names: string[], W: number, y: number): string {
  const item = 110;
  const start = W / 2 - (names.length * item) / 2;
  return names.map((n, i) => `<rect x="${start + i * item}" y="${y - 10}" width="12" height="12" fill="${PALETTE[i % PALETTE.length]}"/><text x="${start + i * item + 16}" y="${y}" font-size="11">${escHtml(n)}</text>`).join("");
}

export const svgDataUri = (svg: string) => `data:image/svg+xml;base64,${Buffer.from(svg, "utf8").toString("base64")}`;
