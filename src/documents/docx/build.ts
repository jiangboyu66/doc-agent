/**
 * OOXML 构件生成器（纯函数，只产出 XML 字符串）
 *
 * 表格、图片、矢量图形（形状组合）、原生图表（含内嵌 Excel 数据，可在 Word 中"编辑数据"）、
 * 公式（LaTeX → OMML）、编号定义。由 DocxDocument 负责把它们放进文档包、登记关系与内容类型。
 *
 * 所有命名空间都在元素上就地声明：插入到任何来源的 docx（Word、WPS、LibreOffice、python-docx 生成）
 * 都不依赖根元素是否预先声明了 wp / a / pic / wps / wpg / c / m。
 */

import { encodeAttr, encodeText } from "./xml.js";
import { createZip } from "./zip.js";

export const DNS = {
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
};

export const EMU_PER_PT = 12700;
export const emu = (pt: number) => Math.max(0, Math.round(pt * EMU_PER_PT));
export const twip = (pt: number) => Math.round(pt * 20);

const t = (w: string, text: string) => `<${w}:t${/^\s|\s$/.test(text) ? ' xml:space="preserve"' : ""}>${encodeText(text)}</${w}:t>`;
/** 文字 → run 内容：\n 换行、\t 制表符 */
export function runContent(w: string, text: string): string {
  return text.split(/(\n|\t)/).filter((x) => x !== "").map((x) => (x === "\n" ? `<${w}:br/>` : x === "\t" ? `<${w}:tab/>` : t(w, x))).join("");
}
export const hex6 = (c: string | undefined, fallback: string) => {
  const v = (c ?? "").replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(v) ? v : fallback;
};

// ---------------------------------------------------------------------------
// 图片尺寸
// ---------------------------------------------------------------------------

export interface ImageInfo { width: number; height: number; ext: "png" | "jpeg" | "gif" | "bmp"; mime: string }

export function imageInfo(buf: Buffer): ImageInfo | null {
  if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) {
    return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20), ext: "png", mime: "image/png" };
  }
  if (buf.length > 10 && buf.toString("ascii", 0, 3) === "GIF") {
    return { width: buf.readUInt16LE(6), height: buf.readUInt16LE(8), ext: "gif", mime: "image/gif" };
  }
  if (buf.length > 26 && buf.toString("ascii", 0, 2) === "BM") {
    return { width: buf.readInt32LE(18), height: Math.abs(buf.readInt32LE(22)), ext: "bmp", mime: "image/bmp" };
  }
  if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
    let i = 2;
    while (i + 9 < buf.length) {
      if (buf[i] !== 0xff) { i++; continue; }
      const marker = buf[i + 1];
      const len = buf.readUInt16BE(i + 2);
      if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)) {
        return { height: buf.readUInt16BE(i + 5), width: buf.readUInt16BE(i + 7), ext: "jpeg", mime: "image/jpeg" };
      }
      i += 2 + len;
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// 内嵌对象（wp:inline）外壳
// ---------------------------------------------------------------------------

export function inlineDrawing(p: { w: string; cx: number; cy: number; id: number; name: string; descr?: string; graphicData: string; uri: string; locks?: string }): string {
  const descr = p.descr ? ` descr="${encodeAttr(p.descr.slice(0, 2000))}"` : "";
  return (
    `<${p.w}:drawing><wp:inline xmlns:wp="${DNS.wp}" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${p.cx}" cy="${p.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${p.id}" name="${encodeAttr(p.name)}"${descr}/>` +
    `<wp:cNvGraphicFramePr>${p.locks ?? ""}</wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="${DNS.a}"><a:graphicData uri="${p.uri}">${p.graphicData}</a:graphicData></a:graphic>` +
    `</wp:inline></${p.w}:drawing>`
  );
}

export function pictureXml(p: { w: string; rId: string; cx: number; cy: number; id: number; name: string; descr?: string }): string {
  const pic =
    `<pic:pic xmlns:pic="${DNS.pic}"><pic:nvPicPr><pic:cNvPr id="0" name="${encodeAttr(p.name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${DNS.r}" r:embed="${p.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${p.cx}" cy="${p.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`;
  return inlineDrawing({
    w: p.w, cx: p.cx, cy: p.cy, id: p.id, name: p.name, descr: p.descr, graphicData: pic,
    uri: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    locks: `<a:graphicFrameLocks xmlns:a="${DNS.a}" noChangeAspect="1"/>`,
  });
}

// ---------------------------------------------------------------------------
// 矢量图形：形状组合（Word 中可逐个选中、移动、改色、改文字）
// ---------------------------------------------------------------------------

export type ShapeType = "rect" | "roundRect" | "ellipse" | "diamond" | "triangle" | "parallelogram" | "hexagon" | "cylinder" | "cloud" | "line" | "arrow" | "doubleArrow" | "text";

export interface ShapeSpec {
  type: ShapeType;
  /** 位置与大小（pt，相对画布左上角）；线条用 x,y → x2,y2 */
  x: number; y: number; w?: number; h?: number; x2?: number; y2?: number;
  fill?: string; stroke?: string; stroke_width?: number; dash?: "solid" | "dash" | "dot";
  text?: string; font_size?: number; font_color?: string; bold?: boolean; align?: "left" | "center" | "right";
}

const PRST: Record<string, string> = {
  rect: "rect", roundRect: "roundRect", ellipse: "ellipse", diamond: "diamond", triangle: "triangle",
  parallelogram: "parallelogram", hexagon: "hexagon", cylinder: "can", cloud: "cloud", text: "rect",
};

/** 一组形状 → 内嵌的组合图形 */
export function shapesGroupXml(p: { w: string; id: number; name: string; descr?: string; width: number; height: number; shapes: ShapeSpec[]; font?: string }): string {
  const W = emu(p.width), H = emu(p.height);
  let sid = 1;
  const kids = p.shapes.map((s) => shapeXml(p.w, s, ++sid, p.font)).join("");
  const group =
    `<wpg:wgp xmlns:wpg="${DNS.wpg}" xmlns:wps="${DNS.wps}"><wpg:cNvGrpSpPr/>` +
    `<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/><a:chOff x="0" y="0"/><a:chExt cx="${W}" cy="${H}"/></a:xfrm></wpg:grpSpPr>` +
    kids + `</wpg:wgp>`;
  return inlineDrawing({ w: p.w, cx: W, cy: H, id: p.id, name: p.name, descr: p.descr, graphicData: group, uri: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" });
}

function lnXml(s: ShapeSpec, isLine: boolean): string {
  if (!isLine && s.stroke === "none") return `<a:ln><a:noFill/></a:ln>`;
  const color = hex6(s.stroke, isLine ? "404040" : "2F5597");
  const wd = emu(s.stroke_width ?? (isLine ? 1.25 : 1));
  const dash = s.dash === "dash" ? `<a:prstDash val="dash"/>` : s.dash === "dot" ? `<a:prstDash val="sysDot"/>` : "";
  const head = s.type === "doubleArrow" ? `<a:headEnd type="triangle" w="med" len="med"/>` : "";
  const tail = s.type === "arrow" || s.type === "doubleArrow" ? `<a:tailEnd type="triangle" w="med" len="med"/>` : "";
  return `<a:ln w="${wd}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${dash}${head}${tail}</a:ln>`;
}

function shapeXml(w: string, s: ShapeSpec, id: number, font?: string): string {
  const isLine = s.type === "line" || s.type === "arrow" || s.type === "doubleArrow";
  let xfrm: string;
  let geom: string;
  if (isLine) {
    const x1 = s.x, y1 = s.y, x2 = s.x2 ?? s.x + (s.w ?? 60), y2 = s.y2 ?? s.y + (s.h ?? 0);
    const flipH = x2 < x1 ? ' flipH="1"' : "", flipV = y2 < y1 ? ' flipV="1"' : "";
    xfrm = `<a:xfrm${flipH}${flipV}><a:off x="${emu(Math.min(x1, x2))}" y="${emu(Math.min(y1, y2))}"/><a:ext cx="${emu(Math.abs(x2 - x1))}" cy="${emu(Math.abs(y2 - y1))}"/></a:xfrm>`;
    geom = `<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>`;
  } else {
    xfrm = `<a:xfrm><a:off x="${emu(s.x)}" y="${emu(s.y)}"/><a:ext cx="${emu(s.w ?? 80)}" cy="${emu(s.h ?? 40)}"/></a:xfrm>`;
    const fill = s.type === "text" || s.fill === "none" ? "<a:noFill/>" : `<a:solidFill><a:srgbClr val="${hex6(s.fill, "DEEBF7")}"/></a:solidFill>`;
    geom = `<a:prstGeom prst="${PRST[s.type] ?? "rect"}"><a:avLst/></a:prstGeom>${fill}`;
  }
  const ln = s.type === "text" ? `<a:ln><a:noFill/></a:ln>` : lnXml(s, isLine);
  let txbx = "";
  if (s.text && !isLine) {
    const size = Math.round((s.font_size ?? 10.5) * 2);
    const color = hex6(s.font_color, "1F1F1F");
    const jc = s.align ?? "center";
    const fonts = font ? `<${w}:rFonts ${w}:ascii="${encodeAttr(font)}" ${w}:hAnsi="${encodeAttr(font)}" ${w}:eastAsia="${encodeAttr(font)}"/>` : "";
    const rPr = `<${w}:rPr>${fonts}${s.bold ? `<${w}:b/>` : ""}<${w}:color ${w}:val="${color}"/><${w}:sz ${w}:val="${size}"/><${w}:szCs ${w}:val="${size}"/></${w}:rPr>`;
    const paras = s.text.split("\n").map((line) =>
      `<${w}:p><${w}:pPr><${w}:spacing ${w}:before="0" ${w}:after="0" ${w}:line="240" ${w}:lineRule="auto"/><${w}:jc ${w}:val="${jc}"/></${w}:pPr>${line ? `<${w}:r>${rPr}${t(w, line)}</${w}:r>` : ""}</${w}:p>`
    ).join("");
    txbx = `<wps:txbx><${w}:txbxContent>${paras}</${w}:txbxContent></wps:txbx>`;
  }
  const body = `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="45720" tIns="22860" rIns="45720" bIns="22860" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`;
  const cNv = isLine ? `<wps:cNvCnPr/>` : `<wps:cNvSpPr${s.type === "text" ? ' txBox="1"' : ""}/>`;
  return `<wps:wsp><wps:cNvPr id="${id}" name="${isLine ? "连接线" : "形状"} ${id}"/>${cNv}<wps:spPr>${xfrm}${geom}${ln}</wps:spPr>${txbx}${body}</wps:wsp>`;
}

// ---------------------------------------------------------------------------
// 流程图 / 结构图：节点 + 连线 → 自动分层布局 → 形状
// ---------------------------------------------------------------------------

export interface DiagramNode { id: string; text: string; shape?: "rect" | "roundRect" | "ellipse" | "diamond" | "parallelogram" | "cylinder"; fill?: string }
export interface DiagramEdge { from: string; to: string; label?: string; dashed?: boolean }

export function layoutDiagram(p: { nodes: DiagramNode[]; edges: DiagramEdge[]; direction?: "TB" | "LR"; maxWidth: number; fontSize?: number }): { shapes: ShapeSpec[]; width: number; height: number } {
  const ids = new Set(p.nodes.map((n) => n.id));
  for (const e of p.edges) if (!ids.has(e.from) || !ids.has(e.to)) throw new Error(`连线 ${e.from} → ${e.to} 引用了不存在的节点`);
  // 分层：最长路径（忽略回边）
  const level = new Map<string, number>(p.nodes.map((n) => [n.id, 0]));
  const order = p.nodes.map((n) => n.id);
  for (let iter = 0; iter < p.nodes.length; iter++) {
    let changed = false;
    for (const e of p.edges) {
      if (order.indexOf(e.to) < order.indexOf(e.from) && level.get(e.to)! <= level.get(e.from)!) continue; // 回边
      const nl = level.get(e.from)! + 1;
      if (nl > level.get(e.to)!) { level.set(e.to, nl); changed = true; }
    }
    if (!changed) break;
  }
  const layers: string[][] = [];
  for (const n of p.nodes) (layers[level.get(n.id)!] ??= []).push(n.id);
  const fs = p.fontSize ?? 10.5;
  const lr = p.direction === "LR";
  const nodeW = Math.min(150, Math.max(70, ...p.nodes.map((n) => Math.max(...n.text.split("\n").map((l) => estimateWidth(l, fs))) + 20)));
  const nodeH = Math.max(32, ...p.nodes.map((n) => n.text.split("\n").length * fs * 1.35 + 14));
  const gapMain = lr ? 46 : 34, gapCross = 22;
  const maxPerLayer = Math.max(...layers.map((l) => l.length));
  let w = lr ? layers.length * nodeW + (layers.length - 1) * gapMain : maxPerLayer * nodeW + (maxPerLayer - 1) * gapCross;
  let h = lr ? maxPerLayer * nodeH + (maxPerLayer - 1) * gapCross : layers.length * nodeH + (layers.length - 1) * gapMain;
  const pos = new Map<string, { x: number; y: number }>();
  layers.forEach((layer, li) => {
    const span = layer.length * (lr ? nodeH : nodeW) + (layer.length - 1) * gapCross;
    const crossTotal = lr ? h : w;
    layer.forEach((id, k) => {
      const c = (crossTotal - span) / 2 + k * ((lr ? nodeH : nodeW) + gapCross);
      const m = li * ((lr ? nodeW : nodeH) + gapMain);
      pos.set(id, lr ? { x: m, y: c } : { x: c, y: m });
    });
  });
  // 超出可用宽度时等比缩小
  const scale = Math.min(1, p.maxWidth / w);
  const S = (v: number) => v * scale;
  const shapes: ShapeSpec[] = [];
  const palette = ["DEEBF7", "E2F0D9", "FFF2CC", "FCE4D6", "EDE3F7", "DDEBF7"];
  for (const n of p.nodes) {
    const q = pos.get(n.id)!;
    shapes.push({
      type: n.shape ?? "roundRect", x: S(q.x), y: S(q.y), w: S(nodeW), h: S(nodeH),
      fill: n.fill ?? palette[level.get(n.id)! % palette.length], stroke: "5B7BA6", text: n.text, font_size: Math.max(7, fs * Math.max(scale, 0.75)),
    });
  }
  for (const e of p.edges) {
    const a = pos.get(e.from)!, b = pos.get(e.to)!;
    let x1: number, y1: number, x2: number, y2: number;
    if (lr) {
      const fwd = b.x >= a.x;
      x1 = fwd ? a.x + nodeW : a.x; y1 = a.y + nodeH / 2; x2 = fwd ? b.x : b.x + nodeW; y2 = b.y + nodeH / 2;
      if (b.x === a.x) { x1 = a.x + nodeW / 2; x2 = b.x + nodeW / 2; y1 = b.y > a.y ? a.y + nodeH : a.y; y2 = b.y > a.y ? b.y : b.y + nodeH; }
    } else {
      const fwd = b.y >= a.y;
      x1 = a.x + nodeW / 2; y1 = fwd ? a.y + nodeH : a.y; x2 = b.x + nodeW / 2; y2 = fwd ? b.y : b.y + nodeH;
      if (b.y === a.y) { y1 = a.y + nodeH / 2; y2 = b.y + nodeH / 2; x1 = b.x > a.x ? a.x + nodeW : a.x; x2 = b.x > a.x ? b.x : b.x + nodeW; }
    }
    shapes.push({ type: "arrow", x: S(x1), y: S(y1), x2: S(x2), y2: S(y2), stroke: "404040", dash: e.dashed ? "dash" : "solid" });
    if (e.label) {
      const lw = estimateWidth(e.label, fs * 0.85) + 8;
      shapes.push({ type: "text", x: S((x1 + x2) / 2 - lw / 2 + (lr ? 0 : 6 + lw / 2)), y: S((y1 + y2) / 2 - fs), w: S(lw), h: S(fs * 1.6), text: e.label, font_size: Math.max(7, fs * 0.85 * Math.max(scale, 0.75)), font_color: "404040" });
    }
  }
  w = S(w); h = S(h);
  return { shapes, width: w + 2, height: h + 2 };
}

/** 粗略估计文字宽度（pt）：中日韩全角 1em，其余约 0.55em */
export function estimateWidth(text: string, size: number): number {
  let em = 0;
  for (const ch of text) em += /[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]/.test(ch) ? 1 : 0.55;
  return em * size;
}

// ---------------------------------------------------------------------------
// 原生图表（DrawingML Chart + 内嵌 Excel 工作簿）
// ---------------------------------------------------------------------------

export interface ChartSpec {
  type: "bar" | "column" | "line" | "pie" | "area" | "scatter" | "doughnut";
  title?: string;
  categories: string[];
  series: Array<{ name: string; values: number[] }>;
  x_title?: string;
  y_title?: string;
  stacked?: boolean;
  show_values?: boolean;
  legend?: "bottom" | "right" | "top" | "none";
}

const PALETTE = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "264478", "9E480E", "636363", "997300"];
const colName = (i: number) => String.fromCharCode(65 + i);
const xmlNum = (v: number) => (Number.isFinite(v) ? String(v) : "");

function richText(text: string, size: number, bold: boolean): string {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="${bold ? 1 : 0}"/></a:pPr><a:r><a:rPr lang="zh-CN" sz="${size}" b="${bold ? 1 : 0}"/><a:t>${encodeText(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
}

export function chartXml(spec: ChartSpec): string {
  const n = spec.categories.length;
  const catRef = () =>
    `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f><c:strCache><c:ptCount val="${n}"/>${spec.categories.map((c, i) => `<c:pt idx="${i}"><c:v>${encodeText(c)}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>`;
  const numRef = (tag: string, col: string, values: number[]) =>
    `<c:${tag}><c:numRef><c:f>Sheet1!$${col}$2:$${col}$${n + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${xmlNum(v)}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:${tag}>`;
  const xNumRef = () =>
    `<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${spec.categories.map((c, i) => `<c:pt idx="${i}"><c:v>${xmlNum(Number(c))}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:xVal>`;
  const txRef = (k: number, name: string) =>
    `<c:tx><c:strRef><c:f>Sheet1!$${colName(k + 1)}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${encodeText(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
  const fill = (c: string) => `<c:spPr><a:solidFill><a:srgbClr val="${c}"/></a:solidFill></c:spPr>`;
  const lineSp = (c: string) => `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:round/></a:ln></c:spPr>`;
  const dLbls = spec.show_values
    ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${spec.type === "pie" || spec.type === "doughnut" ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`
    : "";
  const ax1 = 50010001, ax2 = 50010002;
  let plot = "";
  const type = spec.type;
  if (type === "bar" || type === "column") {
    const grouping = spec.stacked ? "stacked" : "clustered";
    plot = `<c:barChart><c:barDir val="${type === "bar" ? "bar" : "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>` +
      spec.series.map((s, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${txRef(k, s.name)}${fill(PALETTE[k % PALETTE.length])}<c:invertIfNegative val="0"/>${dLbls}${catRef()}${numRef("val", colName(k + 1), s.values)}</c:ser>`).join("") +
      `<c:gapWidth val="${spec.stacked ? 80 : 150}"/>${spec.stacked ? '<c:overlap val="100"/>' : ""}<c:axId val="${ax1}"/><c:axId val="${ax2}"/></c:barChart>`;
  } else if (type === "line") {
    plot = `<c:lineChart><c:grouping val="${spec.stacked ? "stacked" : "standard"}"/><c:varyColors val="0"/>` +
      spec.series.map((s, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${txRef(k, s.name)}${lineSp(PALETTE[k % PALETTE.length])}<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>${dLbls}${catRef()}${numRef("val", colName(k + 1), s.values)}<c:smooth val="0"/></c:ser>`).join("") +
      `<c:marker val="1"/><c:axId val="${ax1}"/><c:axId val="${ax2}"/></c:lineChart>`;
  } else if (type === "area") {
    plot = `<c:areaChart><c:grouping val="${spec.stacked ? "stacked" : "standard"}"/><c:varyColors val="0"/>` +
      spec.series.map((s, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${txRef(k, s.name)}${fill(PALETTE[k % PALETTE.length])}${dLbls}${catRef()}${numRef("val", colName(k + 1), s.values)}</c:ser>`).join("") +
      `<c:axId val="${ax1}"/><c:axId val="${ax2}"/></c:areaChart>`;
  } else if (type === "pie" || type === "doughnut") {
    const s = spec.series[0];
    const pts = spec.categories.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>${fill(PALETTE[i % PALETTE.length])}</c:dPt>`).join("");
    const ser = `<c:ser><c:idx val="0"/><c:order val="0"/>${txRef(0, s.name)}${pts}${dLbls}${catRef()}${numRef("val", "B", s.values)}</c:ser>`;
    plot = type === "pie"
      ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
  } else if (type === "scatter") {
    plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>` +
      spec.series.map((s, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${txRef(k, s.name)}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="${PALETTE[k % PALETTE.length]}"/></a:solidFill></c:spPr></c:marker>${dLbls}${xNumRef()}${numRef("yVal", colName(k + 1), s.values)}<c:smooth val="0"/></c:ser>`).join("") +
      `<c:axId val="${ax1}"/><c:axId val="${ax2}"/></c:scatterChart>`;
  }
  const axTitle = (s?: string) => (s ? `<c:title>${richText(s, 1000, false)}<c:overlay val="0"/></c:title>` : "");
  const grid = `<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>`;
  let axes = "";
  if (type === "scatter") {
    axes = `<c:valAx><c:axId val="${ax1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${axTitle(spec.x_title)}<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${ax2}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>` +
      `<c:valAx><c:axId val="${ax2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>${grid}${axTitle(spec.y_title)}<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${ax1}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
  } else if (type !== "pie" && type !== "doughnut") {
    const horiz = type === "bar";
    axes = `<c:catAx><c:axId val="${ax1}"/><c:scaling><c:orientation val="${horiz ? "maxMin" : "minMax"}"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? "l" : "b"}"/>${axTitle(spec.x_title)}<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${ax2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
      `<c:valAx><c:axId val="${ax2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? "t" : "l"}"/>${grid}${axTitle(spec.y_title)}<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${ax1}"/><c:crosses val="${horiz ? "max" : "autoZero"}"/><c:crossBetween val="between"/></c:valAx>`;
  }
  const legendPos = { bottom: "b", right: "r", top: "t" } as const;
  const showLegend = spec.legend !== "none" && (spec.series.length > 1 || type === "pie" || type === "doughnut");
  const legend = showLegend ? `<c:legend><c:legendPos val="${legendPos[(spec.legend as "bottom") ?? "bottom"] ?? "b"}"/><c:overlay val="0"/></c:legend>` : "";
  const title = spec.title ? `<c:title>${richText(spec.title, 1400, true)}<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : `<c:autoTitleDeleted val="1"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<c:chartSpace xmlns:c="${DNS.c}" xmlns:a="${DNS.a}" xmlns:r="${DNS.r}"><c:date1904 val="0"/><c:lang val="zh-CN"/><c:roundedCorners val="0"/>` +
    `<c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:pPr><a:endParaRPr lang="zh-CN"/></a:p></c:txPr>` +
    `<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
}

export function chartWorkbook(spec: ChartSpec): Buffer {
  const esc = (s: string) => encodeText(s);
  const cell = (ref: string, v: string | number) =>
    typeof v === "number" ? `<c r="${ref}"><v>${xmlNum(v)}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
  const rows: string[] = [];
  rows.push(`<row r="1">${cell("A1", " ")}${spec.series.map((s, k) => cell(`${colName(k + 1)}1`, s.name)).join("")}</row>`);
  spec.categories.forEach((c, i) => {
    const r = i + 2;
    const first = spec.type === "scatter" && Number.isFinite(Number(c)) ? Number(c) : c;
    rows.push(`<row r="${r}">${cell(`A${r}`, first)}${spec.series.map((s, k) => cell(`${colName(k + 1)}${r}`, s.values[i] ?? NaN)).join("")}</row>`);
  });
  const files = [
    { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>` },
    { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: "xl/workbook.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${DNS.r}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>` },
    { name: "xl/worksheets/sheet1.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join("")}</sheetData></worksheet>` },
  ];
  return createZip(files.map((f) => ({ name: f.name, data: Buffer.from(f.data, "utf8") })));
}

export function chartInlineXml(p: { w: string; rId: string; cx: number; cy: number; id: number; name: string; descr?: string }): string {
  return inlineDrawing({
    w: p.w, cx: p.cx, cy: p.cy, id: p.id, name: p.name, descr: p.descr,
    uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    graphicData: `<c:chart xmlns:c="${DNS.c}" xmlns:r="${DNS.r}" r:id="${p.rId}"/>`,
  });
}

// ---------------------------------------------------------------------------
// 编号（项目符号 / 多级编号）
// ---------------------------------------------------------------------------

export function abstractNumXml(w: string, id: number, kind: "bullet" | "number" | "chinese" | "outline"): string {
  const bullets = ["●", "○", "■", "●", "○", "■", "●", "○", "■"];
  const lvls = Array.from({ length: 9 }, (_, i) => {
    const ind = 420 * (i + 1);
    let fmt = "decimal", text = `%${i + 1}.`, font = "";
    if (kind === "bullet") {
      fmt = "bullet"; text = bullets[i];
      font = `<${w}:rPr><${w}:rFonts ${w}:ascii="Arial" ${w}:hAnsi="Arial" ${w}:hint="default"/><${w}:sz ${w}:val="16"/></${w}:rPr>`;
    } else if (kind === "chinese") {
      fmt = i === 0 ? "chineseCounting" : i === 1 ? "chineseCounting" : "decimal";
      text = i === 0 ? "%1、" : i === 1 ? "（%2）" : `%${i + 1}.`;
    } else if (kind === "outline") {
      fmt = "decimal"; text = Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join(".") + (i === 0 ? "." : "");
    } else {
      fmt = i % 3 === 0 ? "decimal" : i % 3 === 1 ? "lowerLetter" : "lowerRoman";
      text = i % 3 === 0 ? `%${i + 1}.` : `%${i + 1})`;
    }
    return `<${w}:lvl ${w}:ilvl="${i}"><${w}:start ${w}:val="1"/><${w}:numFmt ${w}:val="${fmt}"/><${w}:lvlText ${w}:val="${encodeAttr(text)}"/><${w}:lvlJc ${w}:val="left"/><${w}:pPr><${w}:ind ${w}:left="${ind}" ${w}:hanging="${kind === "chinese" && i < 2 ? 420 : 420}"/></${w}:pPr>${font}</${w}:lvl>`;
  }).join("");
  return `<${w}:abstractNum ${w}:abstractNumId="${id}"><${w}:multiLevelType ${w}:val="${kind === "bullet" ? "hybridMultilevel" : "multilevel"}"/>${lvls}</${w}:abstractNum>`;
}

// ---------------------------------------------------------------------------
// 公式：LaTeX（常用子集）→ OMML
// ---------------------------------------------------------------------------

const GREEK: Record<string, string> = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ϵ", varepsilon: "ε", zeta: "ζ", eta: "η", theta: "θ", vartheta: "ϑ",
  iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π", varpi: "ϖ", rho: "ρ", varrho: "ϱ", sigma: "σ",
  varsigma: "ς", tau: "τ", upsilon: "υ", phi: "ϕ", varphi: "φ", chi: "χ", psi: "ψ", omega: "ω",
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ", Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
};
const SYMBOLS: Record<string, string> = {
  cdot: "⋅", times: "×", div: "÷", pm: "±", mp: "∓", leq: "≤", le: "≤", geq: "≥", ge: "≥", neq: "≠", ne: "≠", approx: "≈",
  equiv: "≡", sim: "∼", simeq: "≃", propto: "∝", infty: "∞", partial: "∂", nabla: "∇", to: "→", rightarrow: "→", leftarrow: "←",
  Rightarrow: "⇒", Leftarrow: "⇐", leftrightarrow: "↔", Leftrightarrow: "⇔", mapsto: "↦", in: "∈", notin: "∉", subset: "⊂",
  subseteq: "⊆", supset: "⊃", supseteq: "⊇", cup: "∪", cap: "∩", emptyset: "∅", forall: "∀", exists: "∃", neg: "¬",
  land: "∧", lor: "∨", wedge: "∧", vee: "∨", oplus: "⊕", otimes: "⊗", circ: "∘", bullet: "∙", star: "⋆", ast: "∗",
  ldots: "…", cdots: "⋯", vdots: "⋮", ddots: "⋱", dots: "…", prime: "′", angle: "∠", perp: "⊥", parallel: "∥", mid: "∣",
  langle: "⟨", rangle: "⟩", lceil: "⌈", rceil: "⌉", lfloor: "⌊", rfloor: "⌋", hbar: "ℏ", ell: "ℓ", Re: "ℜ", Im: "ℑ",
  aleph: "ℵ", degree: "°", quad: "\u2003", qquad: "\u2003\u2003", ",": "\u2009", ";": "\u2005", ":": "\u2004", "!": "", " ": " ",
  "{": "{", "}": "}", "%": "%", "&": "&", "#": "#", "_": "_", "|": "‖", lbrace: "{", rbrace: "}", vert: "|", Vert: "‖",
  ll: "≪", gg: "≫", therefore: "∴", because: "∵", triangle: "△", square: "□",
};
const NARY: Record<string, string> = { sum: "∑", prod: "∏", coprod: "∐", int: "∫", iint: "∬", iiint: "∭", oint: "∮", bigcup: "⋃", bigcap: "⋂" };
const FUNCS = new Set(["sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "log", "ln", "lg", "exp", "max", "min", "sup", "inf", "lim", "det", "dim", "ker", "deg", "gcd", "arg", "Pr", "mod"]);
const ACCENTS: Record<string, string> = { hat: "̂", widehat: "̂", bar: "̅", overline: "̅", vec: "⃗", dot: "̇", ddot: "̈", tilde: "̃", widetilde: "̃" };

export function latexToOmml(latex: string): string {
  let i = 0;
  const s = latex.trim();
  const run = (text: string, plain = false) =>
    text ? `<m:r>${plain ? "<m:rPr><m:sty m:val=\"p\"/></m:rPr>" : ""}<m:t xml:space="preserve">${encodeText(text)}</m:t></m:r>` : "";

  const skipWs = () => { while (i < s.length && /\s/.test(s[i])) i++; };

  const readGroup = (): string => {
    skipWs();
    if (s[i] === "{") {
      i++;
      const out = parseSeq("}");
      i++; // }
      return out;
    }
    return parseAtom();
  };
  const readRawGroup = (): string => {
    skipWs();
    if (s[i] !== "{") return s[i++] ?? "";
    let depth = 0, start = ++i;
    for (; i < s.length; i++) {
      if (s[i] === "{") depth++;
      else if (s[i] === "}") { if (depth === 0) break; depth--; }
    }
    const raw = s.slice(start, i);
    i++;
    return raw;
  };
  const readCommand = (): string => {
    i++; // backslash
    if (i < s.length && !/[A-Za-z]/.test(s[i])) return s[i++];
    const st = i;
    while (i < s.length && /[A-Za-z]/.test(s[i])) i++;
    return s.slice(st, i);
  };

  const scripts = (base: string): string => {
    let sub: string | null = null, sup: string | null = null;
    for (;;) {
      skipWs();
      if (s[i] === "_" && sub === null) { i++; sub = readGroup(); continue; }
      if (s[i] === "^" && sup === null) { i++; sup = readGroup(); continue; }
      if (s[i] === "'" && sup === null) { i++; sup = run("′"); continue; }
      break;
    }
    if (sub !== null && sup !== null) return `<m:sSubSup><m:e>${base}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`;
    if (sub !== null) return `<m:sSub><m:e>${base}</m:e><m:sub>${sub}</m:sub></m:sSub>`;
    if (sup !== null) return `<m:sSup><m:e>${base}</m:e><m:sup>${sup}</m:sup></m:sSup>`;
    return base;
  };

  const DELIMS: Record<string, string> = { "(": "(", ")": ")", "[": "[", "]": "]", "\\{": "{", "\\}": "}", "|": "|", "\\|": "‖", ".": "", "\\langle": "⟨", "\\rangle": "⟩", "\\lfloor": "⌊", "\\rfloor": "⌋", "\\lceil": "⌈", "\\rceil": "⌉" };
  const readDelim = (): string => {
    skipWs();
    if (s[i] === "\\") {
      const save = i;
      const cmd = "\\" + readCommand();
      if (cmd in DELIMS) return DELIMS[cmd];
      i = save;
      return "";
    }
    const ch = s[i++];
    return DELIMS[ch] ?? ch;
  };

  const parseAtom = (): string => {
    skipWs();
    const ch = s[i];
    if (ch === undefined) return "";
    if (ch === "{") return readGroup();
    if (ch === "\\") {
      const cmd = readCommand();
      if (cmd === "frac" || cmd === "dfrac" || cmd === "tfrac") {
        const num = readGroup(), den = readGroup();
        return `<m:f><m:num>${num}</m:num><m:den>${den}</m:den></m:f>`;
      }
      if (cmd === "binom") {
        const a = readGroup(), b = readGroup();
        return `<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${a}</m:num><m:den>${b}</m:den></m:f></m:e></m:d>`;
      }
      if (cmd === "sqrt") {
        skipWs();
        let deg = "";
        if (s[i] === "[") {
          const end = s.indexOf("]", i);
          const inner = s.slice(i + 1, end);
          i = end + 1;
          deg = latexToOmmlInner(inner);
        }
        const body = readGroup();
        return deg
          ? `<m:rad><m:deg>${deg}</m:deg><m:e>${body}</m:e></m:rad>`
          : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${body}</m:e></m:rad>`;
      }
      if (cmd in NARY) {
        let sub = "", sup = "";
        for (;;) {
          skipWs();
          if (s[i] === "_") { i++; sub = readGroup(); continue; }
          if (s[i] === "^") { i++; sup = readGroup(); continue; }
          if (s.startsWith("\\limits", i)) { i += 7; continue; }
          break;
        }
        const body = parseOperand();
        const hide = `${sub ? "" : '<m:subHide m:val="1"/>'}${sup ? "" : '<m:supHide m:val="1"/>'}`;
        return `<m:nary><m:naryPr><m:chr m:val="${NARY[cmd]}"/>${cmd.includes("int") ? "" : '<m:limLoc m:val="undOvr"/>'}${hide}</m:naryPr><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup><m:e>${body}</m:e></m:nary>`;
      }
      if (cmd === "left") {
        const open = readDelim();
        const inner = parseSeq("\\right");
        i += 6; // \right
        const close = readDelim();
        return `<m:d><m:dPr><m:begChr m:val="${encodeAttr(open)}"/><m:endChr m:val="${encodeAttr(close)}"/></m:dPr><m:e>${inner}</m:e></m:d>`;
      }
      if (cmd === "text" || cmd === "mathrm" || cmd === "textrm" || cmd === "operatorname" || cmd === "mbox") return run(readRawGroup(), true);
      if (cmd === "mathbf" || cmd === "boldsymbol" || cmd === "bm") return `<m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>${encodeText(readRawGroup())}</m:t></m:r>`;
      if (cmd === "mathit") return run(readRawGroup());
      if (cmd === "mathbb") {
        const map: Record<string, string> = { R: "ℝ", N: "ℕ", Z: "ℤ", Q: "ℚ", C: "ℂ", P: "ℙ", E: "𝔼" };
        return run([...readRawGroup()].map((c) => map[c] ?? c).join(""));
      }
      if (cmd === "mathcal") return `<m:r><m:rPr><m:scr m:val="script"/></m:rPr><m:t>${encodeText(readRawGroup())}</m:t></m:r>`;
      if (cmd in ACCENTS) {
        const body = readGroup();
        if (cmd === "overline" || cmd === "bar") return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${body}</m:e></m:bar>`;
        return `<m:acc><m:accPr><m:chr m:val="${ACCENTS[cmd]}"/></m:accPr><m:e>${body}</m:e></m:acc>`;
      }
      if (cmd === "underline") return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${readGroup()}</m:e></m:bar>`;
      if (FUNCS.has(cmd)) {
        let name = run(cmd, true);
        name = scripts(name);
        return `<m:func><m:fName>${name}</m:fName><m:e>${parseOperand()}</m:e></m:func>`;
      }
      if (cmd in GREEK) return run(GREEK[cmd]);
      if (cmd in SYMBOLS) return run(SYMBOLS[cmd], true);
      if (cmd === "displaystyle" || cmd === "limits" || cmd === "nolimits" || cmd === "big" || cmd === "Big" || cmd === "bigg" || cmd === "Bigg") return "";
      return run(cmd, true);
    }
    if (ch === "}") return "";
    i++;
    if (/[0-9.]/.test(ch)) {
      let num = ch;
      while (i < s.length && /[0-9.]/.test(s[i])) num += s[i++];
      return run(num);
    }
    if (/[A-Za-z]/.test(ch)) return run(ch);
    if ("+-=<>()[]|,;:!/*".includes(ch)) return run(ch === "-" ? "−" : ch === "*" ? "∗" : ch, true);
    return run(ch);
  };

  /** 求和、函数等后面跟的"主体"：到下一个同级的 + - = 或分组结束为止 */
  const parseOperand = (): string => {
    let out = "";
    for (;;) {
      skipWs();
      if (i >= s.length || s[i] === "}" || s.startsWith("\\right", i) || "+-=<>,".includes(s[i]) || s.startsWith("\\\\", i)) break;
      if (s[i] === "\\") {
        const save = i;
        const cmd = readCommand();
        i = save;
        if (["leq", "geq", "neq", "approx", "le", "ge", "ne", "to", "pm", "cdot", "times", "quad", "qquad", "in", "equiv", "rightarrow"].includes(cmd)) break;
      }
      out += scripts(parseAtom());
    }
    return out;
  };

  const parseSeq = (until?: string): string => {
    let out = "";
    while (i < s.length) {
      skipWs();
      if (until && s.startsWith(until, i)) break;
      if (s[i] === "}" && until === "}") break;
      if (s[i] === "&") { i++; continue; }
      if (s.startsWith("\\\\", i)) { i += 2; out += run(" "); continue; }
      out += scripts(parseAtom());
    }
    return out;
  };

  function latexToOmmlInner(src: string): string {
    return latexToOmml(src);
  }

  return parseSeq();
}

/** 公式段落内容：display 时居中；带编号时用"居中制表位 + 右对齐制表位"把编号放在同一行末尾 */
export function equationParagraphXml(p: { w: string; latex: string; display: boolean; number?: string; textWidthTwips: number; pPrExtra?: string; rPr?: string }): string {
  const w = p.w;
  const omml = `<m:oMath xmlns:m="${DNS.m}">${latexToOmml(p.latex)}</m:oMath>`;
  if (!p.display) return omml;
  if (p.number) {
    const tabs = `<${w}:tabs><${w}:tab ${w}:val="center" ${w}:pos="${Math.round(p.textWidthTwips / 2)}"/><${w}:tab ${w}:val="right" ${w}:pos="${p.textWidthTwips}"/></${w}:tabs>`;
    return `<${w}:p><${w}:pPr>${tabs}${p.pPrExtra ?? ""}</${w}:pPr><${w}:r>${p.rPr ?? ""}<${w}:tab/></${w}:r>${omml}<${w}:r>${p.rPr ?? ""}<${w}:tab/>${t(w, p.number)}</${w}:r></${w}:p>`;
  }
  return `<${w}:p><${w}:pPr>${p.pPrExtra ?? ""}<${w}:jc ${w}:val="center"/></${w}:pPr><m:oMathPara xmlns:m="${DNS.m}">${omml.replace(` xmlns:m="${DNS.m}"`, "")}</m:oMathPara></${w}:p>`;
}

/**
 * 兼容外壳：新式对象（形状组合、图表）+ 图片后备。
 * Word / WPS / LibreOffice 认识 wps 命名空间，使用 Choice 中的可编辑对象；
 * 不认识的阅读器（浏览器内的快速预览等）显示 Fallback 中的 PNG 图片。
 */
export function alternateContentXml(choice: string, fallback: string): string {
  return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="${DNS.wps}"><mc:Choice Requires="wps">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
}
