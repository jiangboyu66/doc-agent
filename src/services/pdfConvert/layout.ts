/**
 * 版面重建：把 PDF 的"字在哪里、线在哪里"还原成 Word 的"段落、表格、分栏"
 *
 * 核心思路：Word 是流式排版，PDF 是绝对定位。为了在 Word 里得到同样的版面：
 *   - 表格：由边框线段构成的网格还原（含合并单元格、每条边框的有无/粗细/颜色、底纹、垂直对齐）；
 *   - 段落：按基线聚成行、按行距/缩进/是否折行聚成段，推断对齐方式、首行/悬挂缩进、制表位；
 *   - 位置：每段采用"固定行距 = PDF 实测行距"，段前距 = 与上一块的实测间距，使每一行的基线都落在原位置；
 *   - 分栏：检测贯穿版心的空白栏间距，用连续分节 + 分栏还原双栏排版；
 *   - 图片：以页面绝对坐标浮动定位（不参与文字流），位置与原 PDF 一致；
 *   - 每一页以分页开始，页与页之间不会互相推挤。
 */

import type { FillRect, ImageBox, PageModel, Seg, Span } from "./extract.js";
import { collectHyphenated, collectWords, isSoftHyphen, wordCounts } from "../../documents/hyphen.js";
import { detectDropCaps, detectEquations, detectFigures, detectImageGrids, detectMaskInk, rawBase, rasterBox, removeInside, spansText, type Box, type MathCtx } from "./raster.js";

export interface Run {
  text: string;
  tab?: boolean;
  /** 行尾换行（逐行一致模式） */
  br?: boolean;
  /** 字符间距（pt，每个字符之后），还原 PDF 中的紧缩/加宽 */
  spacing?: number;
  /** 本行在原 PDF 中被压紧（两端对齐时压缩了词间距）：逐行一致模式下按实测值紧缩，保证不会折行 */
  lineSpacing?: number;
  /** 本段实测的字符间距（样式中的紧缩/加宽） */
  paraSpacing?: number;
  font: string;
  size: number;
  bold: boolean;
  italic: boolean;
  color: string;
  underline?: boolean;
  strike?: boolean;
  vertAlign?: "superscript" | "subscript";
  /** 域（页眉页脚中的页码） */
  field?: "PAGE";
  /** 原 PDF 的分页位置（写成 lastRenderedPageBreak，浏览器预览据此分页） */
  pageMark?: boolean;
  /** 行内图片（没有文字编码的公式 / 符号）：text 为占位符 U+FFFC */
  img?: ImageBox;
  /** 行内图片底边在基线以下的距离（pt），写成字符位置（降低） */
  imgLower?: number;
  /** 行内图片的缩放比例（比行距略高的图缩小到正好放进一行，不撑高行距、不被固定行距裁掉） */
  imgScale?: number;
}

export interface TabStop { pos: number; align: "left" | "right" }
export interface Border { width: number; color: string }

export interface Para {
  kind: "p";
  runs: Run[];
  align: "left" | "center" | "right" | "both";
  indLeft: number;
  indRight: number;
  firstLine: number;
  spaceBefore: number;
  lineHeight: number;
  /** 制表位（相对版心/栏的左边界，pt）；行末顶到右边界的编号等用右对齐制表位 */
  tabs: TabStop[];
  heading?: number;
  pageBreakBefore?: boolean;
  columnBreakBefore?: boolean;
  borderBottom?: Border;
  images?: ImageBox[];
  tiny?: boolean;
  sectionEnd?: SectionProps;
  size: number;
  /** 分隔线、竖线、色块等图形：页面绝对坐标，不参与文字流 */
  shapes?: Array<{ kind: "line" | "rect"; x0: number; y0: number; x1: number; y1: number; width: number; color: string; behind: boolean }>;
  /** 页眉页脚等：以页面绝对坐标定位的文本框（挂在页首锚点段落上，不参与文字流，不会被挤动） */
  textBoxes?: Array<{ x: number; y: number; w: number; h: number; para: Para }>;
  /** 逐行一致模式：每一行的 run 范围、自然宽度与可用宽度，用于最终的"不折行"校验 */
  lineFits?: Array<{ start: number; end: number; natural: number; chars: number; avail: number }>;
  /** 独占一段的行内图片（行间公式、图）：随文字流移动，前后文字增减时不会与文字重叠 */
  inlineImages?: ImageBox[];
  /** 随段落浮动、文字环绕的图片（首字下沉、作者照片等旁边有文字的图）：相对栏左边界 / 段落顶部定位 */
  floats?: Array<{ img: ImageBox; dx: number; dy: number }>;
  /** 区域（栏 / 页）的最后一段、第一段：流式模式下用于合并被分栏、分页切断的段落 */
  regionEnd?: boolean;
  regionStart?: boolean;
  /** 最后一行排满到右边界（段落可能在此被切断） */
  lastFull?: boolean;
  /** 原 PDF 中新的一页从这一段开始（流式模式）：写成 lastRenderedPageBreak，浏览器预览据此分页 */
  pageMark?: boolean;
  /** 与下段同页（表题与表格、图与图题不分开） */
  keepNext?: boolean;
  /** 页眉中的图片（徽标等）：相对页边距左边界 / 段落顶部定位，不环绕文字 */
  hfImages?: Array<{ img: ImageBox; dx: number; dy: number }>;
}

export interface Cell {
  width: number;
  gridSpan: number;
  vMerge?: "restart" | "continue";
  borders: { top: Border | null; left: Border | null; bottom: Border | null; right: Border | null };
  shade?: string;
  vAlign: "top" | "center" | "bottom";
  blocks: Para[];
}

export interface Table {
  kind: "table";
  /** 行高规则：逐行一致模式下固定行高（与原文完全一致，不会被撑高） */
  exactRows?: boolean;
  indent: number;
  colWidths: number[];
  rows: Array<{ height: number; cells: Cell[] }>;
  /** 浮动表格（页面绝对坐标，pt）：旁边并排有文字时使用，不占文字流 */
  float?: { x: number; y: number };
}

export type Block = Para | Table;

export interface SectionProps {
  cols: number; colWidths: number[]; colGap: number;
  /** 本节从第一页开始（流式模式）：首页使用"首页"页眉页脚 */
  titlePg?: boolean;
}

/** 页眉 / 页脚（流式模式）：首页、奇数页（默认）、偶数页 */
export interface HeaderSet { default?: Para[]; first?: Para[]; even?: Para[] }

export interface DocModel {
  pageW: number;
  pageH: number;
  margins: { top: number; bottom: number; left: number; right: number };
  blocks: Block[];
  finalSection: SectionProps;
  bodyFont: string;
  bodySize: number;
  stats: { pages: number; paragraphs: number; tables: number; images: number; vectorShapes: number; rotatedText: number; columnPages: number; figures: number; equations: number; inlineMath?: number };
  /** 渲染为图片的区域中的文字（计入完整性校验） */
  rasterText: string;
  mode: LayoutMode;
  /** 流式模式：真正的页眉页脚（含页码域） */
  headers?: HeaderSet;
  footers?: HeaderSet;
  headerDist?: number;
  footerDist?: number;
  /** 第一页的页码（PDF 中印刷的页码不从 1 开始时） */
  pageNumStart?: number;
}

// ---------------------------------------------------------------------------
// 行
// ---------------------------------------------------------------------------

export interface Line {
  spans: Span[];
  x0: number; x1: number;
  top: number; bottom: number;
  baseline: number;
  size: number;
}

const CELL_MARGIN = 2.85; // pt，单元格左右边距（57 twips）
const BULLET = /^(\[\d{1,4}\]|[•·▪■□◆◇○●➢➤►▸\-–—*]|\(?\d{1,3}[.)、]|\(?[a-zA-Z][.)]|\([ivxIVX]+\)|[ivxIVX]+[.)]|[一二三四五六七八九十]+[、.．])$/;
const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]/;

/**
 * 两端对齐的文字，大部分行的边界落在同一个坐标上（栏/版心的边）；个别行（长网址、悬挂标点）会越界。
 * 取出现最多的坐标作为边界；找不到明显的众数时退回最大/最小值。
 */
function edgeOf(values: number[], side: "max" | "min"): number {
  if (!values.length) return 0;
  const counts = new Map<number, number>();
  for (const v of values) { const k = Math.round(v * 2) / 2; counts.set(k, (counts.get(k) ?? 0) + 1); }
  // 允许 ±0.5pt 的抖动
  const score = (k: number) => (counts.get(k) ?? 0) + (counts.get(k - 0.5) ?? 0) + (counts.get(k + 0.5) ?? 0);
  const need = Math.max(3, Math.ceil(values.length * 0.15));
  const cands = [...counts.keys()].filter((k) => score(k) >= need);
  if (!cands.length) return side === "max" ? Math.max(...values) : Math.min(...values);
  return side === "max" ? Math.max(...cands) : Math.min(...cands);
}

function median(xs: number[]): number {
  if (!xs.length) return 0;
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

function finishLine(spans: Span[]): Line {
  spans.sort((a, b) => a.x0 - b.x0);
  const bySize = new Map<number, number>();
  for (const s of spans) bySize.set(s.size, (bySize.get(s.size) ?? 0) + s.text.trim().length);
  const size = [...bySize.entries()].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0][0];
  const main = spans.filter((s) => Math.abs(s.size - size) < 0.6);
  const baseline = median(main.map((s) => s.baseline));
  return {
    spans,
    x0: Math.min(...spans.map((s) => s.x0)),
    x1: Math.max(...spans.map((s) => s.x1)),
    top: Math.min(...main.map((s) => s.top)),
    bottom: Math.max(...main.map((s) => s.bottom)),
    baseline,
    size,
  };
}

export function buildLines(spans: Span[]): Line[] {
  const sorted = [...spans].sort((a, b) => a.baseline - b.baseline || a.x0 - b.x0);
  const groups: Array<{ spans: Span[]; top: number; bottom: number; baseline: number }> = [];
  for (const s of sorted) {
    let best: (typeof groups)[number] | null = null;
    let bestD = Infinity;
    for (let i = groups.length - 1; i >= 0 && i >= groups.length - 12; i--) {
      const g = groups[i];
      const overlap = Math.min(g.bottom, s.bottom) - Math.max(g.top, s.top);
      const minH = Math.min(g.bottom - g.top, s.bottom - s.top);
      if (overlap < 0.45 * minH) continue;
      // 同字号的文字必须共享基线（行距小于字框高度时，仅看重叠会误并到上一行）；小字号的上下标允许偏移
      const gSize = Math.max(...g.spans.map((x) => x.size));
      const d0 = Math.abs(g.baseline - s.baseline);
      if (Math.abs(gSize - s.size) < 0.6 ? d0 > 0.2 * s.size : d0 > 0.5 * Math.max(gSize, s.size)) continue;
      const clash = g.spans.some((o) => Math.min(o.x1, s.x1) - Math.max(o.x0, s.x0) > 0.3 * Math.min(o.x1 - o.x0, s.x1 - s.x0));
      if (clash) continue;
      const d = Math.abs(g.baseline - s.baseline);
      if (d < bestD) { bestD = d; best = g; }
    }
    if (best) {
      const prevMax = Math.max(...best.spans.map((x) => x.size));
      best.spans.push(s);
      // 以较大字号的片段为准更新行的垂直范围与基线（上下标不扩大行，也不决定基线）
      if (s.size > prevMax + 0.5) {
        best.top = s.top;
        best.bottom = s.bottom;
        best.baseline = s.baseline;
      } else if (s.size >= prevMax - 0.5) {
        best.top = Math.min(best.top, s.top);
        best.bottom = Math.max(best.bottom, s.bottom);
      }
    } else groups.push({ spans: [s], top: s.top, bottom: s.bottom, baseline: s.baseline });
  }
  attachMarks(groups);
  return groups.map((g) => finishLine(g.spans)).sort((a, b) => a.baseline - b.baseline || a.x0 - b.x0);
}

/** 重音符号 → 组合字符（F + ˆ → F̂） */
const COMBINING: Record<string, string> = { "ˆ": "\u0302", "^": "\u0302", "˜": "\u0303", "~": "\u0303", "¯": "\u0304", "˙": "\u0307", "¨": "\u0308", "´": "\u0301", "`": "\u0300", "ˇ": "\u030C", "→": "\u20D7", "⃗": "\u20D7" };

/**
 * 行内公式的上下标、重音符号与正文字母上下叠放（W 的上标 Q 与下标 e 在同一横坐标），
 * 成行时会因"横向重叠"被分到单独的一行，转换后变成正文上方多出的一行碎片、行距被撑大。
 * 这里把这类"碎片行"并回它所属的正文行：重音变成组合字符，上下标作为上标 / 下标片段。
 */
function attachMarks(groups: Array<{ spans: Span[]; top: number; bottom: number; baseline: number }>) {
  const sizeOf = (g: { spans: Span[] }) => Math.max(...g.spans.map((x) => x.size));
  const isAccent = (t: string) => /^[ˆ^˜~¯˙¨´`ˇ→⃗\s]+$/.test(t);
  for (let i = groups.length - 1; i >= 0; i--) {
    const g = groups[i];
    const text = g.spans.map((x) => x.text).join("").trim();
    if (!text || text.replace(/\s/g, "").length > 30) continue;
    // 重音符号的字号与正文相同，不参与判断
    const plain = g.spans.filter((x) => !isAccent(x.text));
    const gs = plain.length ? Math.max(...plain.map((x) => x.size)) : sizeOf(g);
    // 找竖直方向紧挨着、字号更大、横向覆盖它的正文行
    let host: (typeof groups)[number] | null = null, best = Infinity;
    for (const h of groups) {
      if (h === g) continue;
      const hs = sizeOf(h);
      const allAccent = g.spans.every((x) => isAccent(x.text));
      if (!allAccent && gs > 0.85 * hs) continue;
      const hx0 = Math.min(...h.spans.map((x) => x.x0)), hx1 = Math.max(...h.spans.map((x) => x.x1));
      const gx0 = Math.min(...g.spans.map((x) => x.x0)), gx1 = Math.max(...g.spans.map((x) => x.x1));
      // 行末的上标（i 的 th）可能略超出正文行
      if (gx0 < hx0 - 2 || gx1 > hx1 + 2.5 * hs) continue;
      const d = Math.abs(h.baseline - g.baseline);
      if (d > 0.75 * hs) continue;
      if (d < best) { best = d; host = h; }
    }
    if (!host) continue;
    for (const sp of g.spans) {
      if (isAccent(sp.text)) {
        const ch = COMBINING[sp.text.trim()[0]];
        const cx = (sp.x0 + sp.x1) / 2;
        const k = host.spans.findIndex((o) => o.x0 - 0.5 <= cx && o.x1 + 0.5 >= cx && o.text.trim());
        if (ch && k >= 0) {
          const o = host.spans[k];
          const n = o.text.length;
          const idx = Math.min(n - 1, Math.max(0, Math.floor(((cx - o.x0) / Math.max(0.1, o.x1 - o.x0)) * n)));
          host.spans[k] = { ...o, text: o.text.slice(0, idx + 1) + ch + o.text.slice(idx + 1) };
          continue;
        }
      }
      host.spans.push(sp);
    }
    groups.splice(i, 1);
  }
}

// ---------------------------------------------------------------------------
// 表格检测
// ---------------------------------------------------------------------------

interface HSeg { y: number; x0: number; x1: number; w: number; color: string; src: Seg[] }
interface VSeg { x: number; y0: number; y1: number; w: number; color: string; src: Seg[] }

interface TableDraft {
  x0: number; x1: number; y0: number; y1: number;
  xs: number[]; ys: number[];
  hs: HSeg[]; vs: VSeg[];
  spans: Span[];
}

function cluster(values: number[], tol: number): number[] {
  const s = [...values].sort((a, b) => a - b);
  const out: number[][] = [];
  for (const v of s) {
    const last = out[out.length - 1];
    if (last && v - last[last.length - 1] <= tol) last.push(v);
    else out.push([v]);
  }
  return out.map((g) => g.reduce((a, b) => a + b, 0) / g.length);
}

function mergeH(segs: Seg[]): HSeg[] {
  const hs = segs.filter((s) => s.horizontal && s.x1 - s.x0 >= 2).map((s) => ({ y: (s.y0 + s.y1) / 2, x0: s.x0, x1: s.x1, w: s.width, color: s.color, src: [s] }));
  hs.sort((a, b) => a.y - b.y || a.x0 - b.x0);
  const out: HSeg[] = [];
  for (const h of hs) {
    const m = out.find((o) => Math.abs(o.y - h.y) <= 1 && h.x0 <= o.x1 + 1.5 && h.x1 >= o.x0 - 1.5);
    if (m) { m.x0 = Math.min(m.x0, h.x0); m.x1 = Math.max(m.x1, h.x1); m.src.push(...h.src); m.w = Math.max(m.w, h.w); }
    else out.push(h);
  }
  return out;
}

function mergeV(segs: Seg[]): VSeg[] {
  const vs = segs.filter((s) => !s.horizontal && s.y1 - s.y0 >= 2).map((s) => ({ x: (s.x0 + s.x1) / 2, y0: s.y0, y1: s.y1, w: s.width, color: s.color, src: [s] }));
  vs.sort((a, b) => a.x - b.x || a.y0 - b.y0);
  const out: VSeg[] = [];
  for (const v of vs) {
    const m = out.find((o) => Math.abs(o.x - v.x) <= 1 && v.y0 <= o.y1 + 1.5 && v.y1 >= o.y0 - 1.5);
    if (m) { m.y0 = Math.min(m.y0, v.y0); m.y1 = Math.max(m.y1, v.y1); m.src.push(...v.src); m.w = Math.max(m.w, v.w); }
    else out.push(v);
  }
  return out;
}

/**
 * 曲线图 / 柱状图的网格线也是横竖线构成的网格，但不是表格：网格内有曲线、斜线（折线）或多块彩色填充，
 * 而几乎没有文字（刻度标签在网格外）。这类区域留给矢量图识别，整体渲染成图片。
 */
function isChartGrid(page: PageModel, x0: number, y0: number, x1: number, y1: number, cells: number): boolean {
  const inside = (b: { x0: number; y0: number; x1: number; y1: number }) => b.x0 >= x0 - 4 && b.x1 <= x1 + 4 && b.y0 >= y0 - 4 && b.y1 <= y1 + 4;
  // 折线 / 曲线：大部分落在网格内即可（线宽、超出坐标轴的少许部分不影响）
  const mostlyIn = (p: { x0: number; y0: number; x1: number; y1: number }) => {
    const w = Math.max(0, Math.min(p.x1, x1) - Math.max(p.x0, x0)), h = Math.max(0, Math.min(p.y1, y1) - Math.max(p.y0, y0));
    return w >= 0.8 * (p.x1 - p.x0) && h >= 0.8 * (p.y1 - p.y0);
  };
  const curves = page.paths.filter((p) => (p.curved || p.diagonal) && (inside(p) || mostlyIn(p)) && Math.max(p.x1 - p.x0, p.y1 - p.y0) > 0.15 * Math.min(x1 - x0, y1 - y0)).length;
  const gray = (c: string) => { const n = parseInt(c, 16); const r = n >> 16, g = (n >> 8) & 255, b = n & 255; return Math.max(r, g, b) - Math.min(r, g, b) < 24; };
  const colored = page.fills.filter((f) => inside(f) && /^[0-9a-f]{6}$/i.test(f.color) && !gray(f.color)).length;
  if (!curves && colored < 3) return false;
  const textChars = page.spans.filter((sp) => { const cx = (sp.x0 + sp.x1) / 2, cy = (sp.top + sp.bottom) / 2; return cx > x0 && cx < x1 && cy > y0 && cy < y1; })
    .reduce((n, sp) => n + sp.text.replace(/\s/g, "").length, 0);
  return textChars <= Math.max(6, cells * 2);
}

function detectTables(page: PageModel): { tables: TableDraft[]; used: Set<Seg> } {
  const hs = mergeH(page.segs);
  const vs = mergeV(page.segs);
  const n = hs.length + vs.length;
  const parent = Array.from({ length: n }, (_, i) => i);
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  for (let i = 0; i < hs.length; i++) {
    for (let j = 0; j < vs.length; j++) {
      const h = hs[i], v = vs[j];
      if (v.x >= h.x0 - 2.5 && v.x <= h.x1 + 2.5 && h.y >= v.y0 - 2.5 && h.y <= v.y1 + 2.5) parent[find(i)] = find(hs.length + j);
    }
  }
  const comps = new Map<number, { hs: HSeg[]; vs: VSeg[] }>();
  hs.forEach((h, i) => { const r = find(i); if (!comps.has(r)) comps.set(r, { hs: [], vs: [] }); comps.get(r)!.hs.push(h); });
  vs.forEach((v, j) => { const r = find(hs.length + j); if (!comps.has(r)) comps.set(r, { hs: [], vs: [] }); comps.get(r)!.vs.push(v); });

  const tables: TableDraft[] = [];
  const used = new Set<Seg>();
  for (const c of comps.values()) {
    if (c.hs.length < 2 || c.vs.length < 2) continue;
    const xs = cluster(c.vs.map((v) => v.x), 2);
    const ys = cluster(c.hs.map((h) => h.y), 2);
    if (xs.length < 2 || ys.length < 2) continue;
    const x0 = xs[0], x1 = xs[xs.length - 1], y0 = ys[0], y1 = ys[ys.length - 1];
    if (x1 - x0 < 20 || y1 - y0 < 8) continue;
    if (isChartGrid(page, x0, y0, x1, y1, (xs.length - 1) * (ys.length - 1))) continue;
    const t: TableDraft = { x0, x1, y0, y1, xs, ys, hs: c.hs, vs: c.vs, spans: [] };
    tables.push(t);
    for (const h of c.hs) h.src.forEach((s) => used.add(s));
    for (const v of c.vs) v.src.forEach((s) => used.add(s));
  }
  // 文字归属：中心点落在表格内
  for (const s of page.spans) {
    const cx = (s.x0 + s.x1) / 2, cy = (s.top + s.bottom) / 2;
    const t = tables.find((t) => cx > t.x0 && cx < t.x1 && cy > t.y0 && cy < t.y1);
    if (t) t.spans.push(s);
  }
  return { tables, used };
}

/**
 * 三线表（只有横线、没有竖线，学术论文常见）：同宽的横线 ≥2 条夹住的区域，
 * 列由纵向空白分隔，行以"第一列有文字"开始（没有第一列文字的行是上一行单元格的续行）。
 * 只有真实画出的横线成为边框，其余边框为空（用 color = "none" 的虚拟线段表示）。
 */
function detectRuledTables(page: PageModel, used: Set<Seg>, taken: Set<Span>): TableDraft[] {
  const hs = mergeH(page.segs.filter((s) => !used.has(s))).filter((h) => h.x1 - h.x0 >= 60).sort((a, b) => a.y - b.y);
  const groups: HSeg[][] = [];
  for (const h of hs) {
    const g = groups.find((g) => Math.abs(g[0].x0 - h.x0) <= 3 && Math.abs(g[0].x1 - h.x1) <= 3);
    if (g) g.push(h);
    else groups.push([h]);
  }
  const out: TableDraft[] = [];
  for (const g of groups) {
    if (g.length < 2) continue;
    const x0 = Math.min(...g.map((h) => h.x0)), x1 = Math.max(...g.map((h) => h.x1));
    const y0 = g[0].y, y1 = g[g.length - 1].y;
    if (y1 - y0 > 0.5 * page.height || y1 - y0 < 10) continue;
    const spans = page.spans.filter((s) => {
      const cx = (s.x0 + s.x1) / 2, cy = (s.top + s.bottom) / 2;
      return !taken.has(s) && cx > x0 - 2 && cx < x1 + 2 && cy > y0 && cy < y1;
    });
    if (spans.length < 4) continue;
    // 列：文字在横向上的投影之间的空白
    const iv = spans.map((s) => [s.x0, s.x1]).sort((a, b) => a[0] - b[0]);
    const merged: number[][] = [];
    for (const [a, b] of iv) {
      const last = merged[merged.length - 1];
      if (last && a <= last[1] + 4) last[1] = Math.max(last[1], b);
      else merged.push([a, b]);
    }
    if (merged.length < 2 || merged.length > 12) continue;
    const xs = [x0, ...merged.slice(1).map((m, i) => (merged[i][1] + m[0]) / 2), x1];
    const colOf = (sp: Span) => { const c = (sp.x0 + sp.x1) / 2; let k = 0; while (k < xs.length - 2 && c >= xs[k + 1]) k++; return k; };
    // 行
    const lines = buildLines(spans);
    const rows: Line[][] = [];
    for (const l of lines) {
      if (!rows.length || l.spans.some((sp) => colOf(sp) === 0)) rows.push([l]);
      else rows[rows.length - 1].push(l);
    }
    if (rows.length < 2) continue;
    const filled = rows.filter((r) => new Set(r.flatMap((l) => l.spans.map(colOf))).size >= 2).length;
    if (filled < 2) continue;
    const ys = [y0];
    const hsDraft: HSeg[] = [...g];
    for (let i = 1; i < rows.length; i++) {
      const aBottom = Math.max(...rows[i - 1].map((l) => l.bottom)), bTop = Math.min(...rows[i].map((l) => l.top));
      const rule = g.find((h) => h.y >= aBottom - 2 && h.y <= bTop + 2);
      const y = rule ? rule.y : (aBottom + bTop) / 2;
      ys.push(y);
      if (!rule) hsDraft.push({ y, x0, x1, w: 0, color: "none", src: [] });
    }
    ys.push(y1);
    // 中间若还有未落在行边界上的横线（例如表头下的横线恰在行内），忽略即可
    const vs: VSeg[] = xs.map((x) => ({ x, y0, y1, w: 0, color: "none", src: [] }));
    out.push({ x0, x1, y0, y1, xs, ys, hs: hsDraft, vs, spans });
    for (const h of g) h.src.forEach((sg) => used.add(sg));
    for (const sp of spans) taken.add(sp);
  }
  return out;
}

function coverageH(hs: HSeg[], y: number, xa: number, xb: number): HSeg | null {
  const cand = hs.filter((h) => Math.abs(h.y - y) <= 2 && h.x1 > xa && h.x0 < xb);
  const cov = cand.reduce((a, h) => a + Math.min(h.x1, xb) - Math.max(h.x0, xa), 0);
  return cov >= 0.6 * (xb - xa) ? cand.sort((a, b) => b.w - a.w)[0] : null;
}
function coverageV(vs: VSeg[], x: number, ya: number, yb: number): VSeg | null {
  const cand = vs.filter((v) => Math.abs(v.x - x) <= 2 && v.y1 > ya && v.y0 < yb);
  const cov = cand.reduce((a, v) => a + Math.min(v.y1, yb) - Math.max(v.y0, ya), 0);
  return cov >= 0.6 * (yb - ya) ? cand.sort((a, b) => b.w - a.w)[0] : null;
}

// ---------------------------------------------------------------------------
// 段落
// ---------------------------------------------------------------------------

interface Region {
  x0: number;
  x1: number;
  /** 逐行一致模式下，左对齐段落允许向右越界的余量（页边距 / 栏间距），用于防止意外折行 */
  overhang?: number;
  /** 表格单元格：文字可以伸入单元格左右边距（窄列中的"Mean"等不致折行） */
  cell?: boolean;
}

const isCentered = (l: Line, r: Region) => {
  const lg = l.x0 - r.x0, rg = r.x1 - l.x1, w = r.x1 - r.x0;
  return lg > Math.max(4, 0.04 * w) && Math.abs(lg - rg) <= Math.max(3, 0.025 * w);
};
const isRight = (l: Line, r: Region) => r.x1 - l.x1 <= 3 && l.x0 - r.x0 > 0.2 * (r.x1 - r.x0);

function hasBigGap(l: Line): boolean {
  for (let i = 1; i < l.spans.length; i++) if (l.spans[i].x0 - l.spans[i - 1].x1 > Math.max(1.2 * l.size, 12)) return true;
  return false;
}

/** 行首第一个"词"的宽度（中文按单字） */
function firstWordWidth(l: Line): number {
  const s = l.spans[0];
  const t = s.text.trimStart();
  const m = CJK.test(t[0] ?? "") ? t[0] : (/^\S+/.exec(t)?.[0] ?? t);
  return ((s.x1 - s.x0) * m.length) / Math.max(1, s.text.length);
}

/**
 * 每一行所在文字块的右边界：附近若有 ≥3 行落在同一右边界（两端对齐的正文、缩进的摘要），用这条边界；
 * 否则用区域右边界。用于判断"上一行是自动折行还是手动换行/段落结束"。
 */
function localEdges(lines: Line[], r: Region): number[] {
  return lines.map((l, i) => {
    const near = lines.slice(Math.max(0, i - 4), i + 5).filter((o) => Math.abs(o.x0 - l.x0) <= 20 && Math.abs(o.size - l.size) <= 0.15 * l.size);
    const edge = Math.max(...near.map((o) => o.x1));
    return near.filter((o) => edge - o.x1 <= 1.5).length >= 3 ? edge : r.x1;
  });
}

function groupParagraphs(lines: Line[], r: Region): Array<{ lines: Line[]; edge: number }> {
  const paras: Line[][] = [];
  const paraEdge: number[] = [];
  const edges = localEdges(lines, r);
  lines.forEach((l, idx) => {
    const P = paras[paras.length - 1];
    if (P) {
      const a = P[P.length - 1];
      const gap = l.baseline - a.baseline;
      const pitch = P.length >= 2 ? P[P.length - 1].baseline - P[P.length - 2].baseline : 0;
      const sizeOk = Math.abs(l.size - a.size) <= 0.15 * a.size;
      const gapOk = gap > 0.5 * a.size && gap <= 1.75 * Math.max(a.size, l.size) && (!pitch || Math.abs(gap - pitch) <= 0.2 * pitch + 0.6);
      const firstText = l.spans[0].text.trim().split(/\s+/)[0] ?? "";
      const startsBullet = BULLET.test(firstText) && l.spans.length > 1;
      // 自动折行的判据：下一行的第一个词放不进上一行剩余的空间
      const edge = Math.max(edges[idx - 1] ?? r.x1, a.x1);
      const wrapped = a.x1 + firstWordWidth(l) + 0.25 * a.size > edge - 0.5;
      const bodyLeft = P.length >= 2 ? P[1].x0 : null;
      const aIsBullet = BULLET.test(P[0].spans[0].text.trim()) && P[0].spans.length > 1;
      const leftOk = bodyLeft !== null
        ? Math.abs(l.x0 - bodyLeft) <= 3
        : l.x0 <= P[0].x0 + 3 || (aIsBullet && Math.abs(l.x0 - P[0].spans[1].x0) <= 3);
      const centered = isCentered(a, r) && isCentered(l, r) && !hasBigGap(a) && !hasBigGap(l) && a.x1 + firstWordWidth(l) + 0.25 * a.size > r.x1 - (a.x0 - r.x0);
      if (sizeOk && gapOk && !startsBullet && !hasBigGap(l) && ((wrapped && leftOk && !hasBigGap(a)) || centered)) {
        P.push(l);
        paraEdge[paraEdge.length - 1] = Math.max(paraEdge[paraEdge.length - 1], edges[idx]);
        return;
      }
    }
    paras.push([l]);
    paraEdge.push(edges[idx]);
  });
  return paras.map((lines, i) => ({ lines, edge: paraEdge[i] }));
}

/**
 * 行内公式图片放进固定行距的一行：Word 的固定行距中基线以上约"行距 − 下沉"、以下约下沉（≈0.21 字号），
 * 超出的部分会被裁掉，改"最小值"又会撑高行、把栏挤出页面。
 *   - 总高度不超过行距：只调整升降（偏移通常不到 1pt），整图落进行框；
 *   - 略高于行距（≤1.6 倍，如带上下限的求和号）：等比缩小到正好放下；
 *   - 更高的：可编辑排版保持原尺寸（段落改用"最小值"行距）；逐行一致模式也缩放（不能撑高行）。
 */
function fitInline(img: ImageBox, l: Line, lineHeight: number, mode: LayoutMode): { imgLower: number; imgScale?: number } {
  const h = img.y1 - img.y0, lower = img.y1 - l.baseline;
  const d = 0.21 * l.size, L = lineHeight;
  const place = (hh: number, low: number) => Math.min(Math.max(low, hh - (L - d) + 0.15), d - 0.15);
  if (h <= L - 0.3) return { imgLower: place(h, lower) };
  // 逐行一致模式：行、栏、页与原文一一对应，任何一行被撑高都会把栏底的行挤到下一栏 / 下一页，一律缩放
  if (h <= 1.6 * L || mode === "exact") {
    const k = (L - 0.3) / h;
    return { imgLower: place(h * k, lower * k), imgScale: k };
  }
  return { imgLower: lower };
}

/** 换行、制表符等控制 run 的格式：不带行内图片 */
const plainFmt = (f: Omit<Run, "text" | "tab">): Omit<Run, "text" | "tab"> => { const { img: _i, imgLower: _l, imgScale: _s, ...rest } = f; return rest; };

function spanFormat(s: Span, lineSize: number, lineBaseline: number): Omit<Run, "text" | "tab"> {
  let vertAlign: Run["vertAlign"];
  let size = s.size;
  if (s.size < 0.85 * lineSize) {
    if (s.baseline < lineBaseline - 0.12 * lineSize) vertAlign = "superscript";
    else if (s.baseline > lineBaseline + 0.08 * lineSize) vertAlign = "subscript";
    if (vertAlign) size = Math.max(s.size / 0.66, s.size);
  }
  if (s.img) return { font: s.font, size: lineSize, bold: false, italic: false, color: s.color, img: s.img, imgLower: s.img.y1 - lineBaseline };
  return { font: s.font, size, bold: s.bold, italic: s.italic, color: s.color, underline: s.underline, strike: s.strike, vertAlign };
}

const sameFmt = (a: Run, b: Omit<Run, "text" | "tab">) =>
  !a.tab && !a.br && !a.img && !b.img && a.font === b.font && a.size === b.size && a.bold === b.bold && a.italic === b.italic && a.color === b.color && !!a.underline === !!b.underline && !!a.strike === !!b.strike && a.vertAlign === b.vertAlign;

function pushText(runs: Run[], text: string, fmt: Omit<Run, "text" | "tab">) {
  if (fmt.img) {
    // 行内图片：前面的空格单独成一段文字，图片自成一个 run
    const lead = /^\s+/.exec(text)?.[0];
    if (lead) pushText(runs, lead, plainFmt(fmt));
    runs.push({ text: "\uFFFC", ...fmt });
    return;
  }
  const last = runs[runs.length - 1];
  if (last && sameFmt(last, fmt)) last.text += text;
  else runs.push({ text, ...fmt });
}

/**
 * 行的实测：实际宽度、自然宽度（PDF 原字体的字宽 + 词间空格），以及在 Word 目标字体中的宽度
 * （字体被替换时两者不同，例如 Computer Modern → Times New Roman）。跳过行首带制表间距的列表符号。
 */
function lineMeasure(l: Line): { actual: number; natural: number; word: number; chars: number } | null {
  let spans = l.spans;
  if (spans.length > 1 && BULLET.test(spans[0].text.trim()) && spans[1].x0 - spans[0].x1 > 0.45 * l.size) spans = spans.slice(1);
  if (spans.some((s) => !s.natural)) return null;
  let natural = 0, word = 0, chars = 0;
  spans.forEach((s, i) => {
    natural += s.natural!;
    word += s.wordW ?? s.natural!;
    chars += s.text.length;
    if (i > 0) {
      const gap = s.x0 - spans[i - 1].x1;
      if (gap > 0.1 * s.size && !/\s$/.test(spans[i - 1].text) && !/^\s/.test(s.text)) {
        natural += Math.min(s.spaceW ?? 0.25 * s.size, spans[i - 1].spaceW ?? 0.25 * s.size);
        word += Math.min(s.wordSpaceW ?? s.spaceW ?? 0.25 * s.size, spans[i - 1].wordSpaceW ?? spans[i - 1].spaceW ?? 0.25 * s.size);
        chars++;
      }
    }
  });
  return { actual: spans[spans.length - 1].x1 - spans[0].x0, natural, word, chars };
}

/** 行被拉伸的程度：实际宽度 / 自然宽度（无法测量时为 null） */
function stretch(l: Line): number | null {
  const m = lineMeasure(l);
  return m && m.natural > 0 ? m.actual / m.natural : null;
}

/**
 * 两端对齐判断：只在有正面证据时判定——非末行被明显拉宽（实测字宽），
 * 或者（无法测量时）至少 3 行非末行落在同一右边界上。避免把两行标题误判成两端对齐而被拉散。
 */
function isJustified(lines: Line[], paraRight: number): boolean {
  const body = lines.slice(0, -1);
  const atEdge = body.filter((l) => Math.abs(paraRight - l.x1) <= 1.5);
  if (atEdge.length < Math.max(1, body.length * 0.8)) return false;
  // 非两端对齐的行，实测宽度与自然宽度一致（±0.1%）；两端对齐的行会被拉宽，或被压缩词间距（Word 2013+ / LibreOffice 的智能两端对齐）
  const ratios = atEdge.map(stretch).filter((r): r is number => r !== null);
  if (ratios.length) return median(ratios.map((r) => Math.abs(r - 1))) >= 0.006;
  return atEdge.length >= 3;
}

export type LayoutMode = "exact" | "flow";

interface Ctx {
  mode: LayoutMode;
  /** 字符间距样本：字体键 → 每字符的宽度差（pt） */
  spacing: Map<string, number[]>;
  /** 全文的单词表 / 带连字符的写法：流式模式下判断行尾连字符是否为排版断词 */
  words: Set<string>;
  hyphenated: Set<string>;
  counts: Map<string, number>;
  /** 去掉的排版断词连字符（计入完整性校验） */
  softHyphens: number;
}

const fontKey = (f: { font: string; size: number; bold: boolean; italic: boolean }) => `${f.font}|${f.size}|${f.bold ? 1 : 0}|${f.italic ? 1 : 0}`;

/** 行 → 段落属性（对齐、缩进、行距、制表位、文字 run） */
function buildPara(lines: Line[], r: Region, blockEdge: number, ctx: Ctx): Para {
  const W = r.x1 - r.x0;
  const first = lines[0];
  const size = median(lines.map((l) => l.size));
  const pitch = lines.length >= 2 ? median(lines.slice(1).map((l, i) => l.baseline - lines[i].baseline)) : 0;
  // 多行段落：行距就是 PDF 实测的基线间距；单行：按字号的常规单倍行距
  const lineHeight = pitch || size * 1.17;

  let align: Para["align"] = "left";
  // 段落自身的右边界（缩进的段落如摘要，右边界在版心之内）
  const paraRight = lines.length >= 2 ? edgeOf(lines.slice(0, -1).map((l) => l.x1), "max") : first.x1;
  // 顺序：居中 → 两端对齐 → 右对齐。图片旁边的两端对齐文字左边齐、右边也齐，不能误判成右对齐；
  // 真正的右对齐多行文字左边参差不齐
  const leftSpread = Math.max(...lines.map((l) => l.x0)) - Math.min(...lines.map((l) => l.x0));
  // 多行居中（论文标题等）：各行中心重合、起点各不相同。不依赖版心——版心可能被页码等撑宽，标题是相对正文居中的；
  // 两端对齐的段落各行起点相同（首行缩进的那一行中心偏开），左对齐的参差段落起点也相同
  const centers = lines.map((l) => (l.x0 + l.x1) / 2);
  const cMean = centers.reduce((a, b) => a + b, 0) / Math.max(1, centers.length);
  const centeredBlock = lines.length >= 2 && centers.every((c) => Math.abs(c - cMean) <= 2) &&
    Math.max(...lines.map((l) => l.x0)) - Math.min(...lines.map((l) => l.x0)) > 3;
  if ((lines.every((l) => isCentered(l, r)) || centeredBlock) && !lines.some(hasBigGap)) align = "center";
  else if (lines.length >= 2 && isJustified(lines, paraRight)) align = "both";
  else if (lines.every((l) => isRight(l, r)) && !lines.some(hasBigGap) && (lines.length === 1 || leftSpread > 3)) align = "right";

  let indLeft = 0, firstLine = 0, indRight = 0;
  const tabs: TabStop[] = [];
  const runs: Run[] = [];

  if (align === "left" || align === "both") {
    const bulletHang = BULLET.test(first.spans[0].text.trim()) && first.spans.length > 1 && first.spans[1].x0 - first.spans[0].x1 > Math.max(0.45 * size, 2 * (first.spans[0].spaceW ?? 0.25 * size));
    const bodyX = lines.length >= 2 ? lines[1].x0 : bulletHang ? first.spans[1].x0 : first.x0;
    indLeft = Math.max(0, bodyX - r.x0);
    firstLine = first.x0 - r.x0 - indLeft;
    // 右缩进：段落所在文字块有自己的右边界（缩进的摘要等）时，用这条边界；普通段落以区域右边界为准
    // 逐行一致模式下，非两端对齐的段落行尾已固定，不需要右缩进（留出余量防止意外折行）
    const edge = align === "both" ? Math.max(paraRight, blockEdge) : blockEdge;
    if (lines.length >= 2 && edge < r.x1 - Math.max(3 * size, 0.08 * W) && (ctx.mode === "flow" || align === "both")) indRight = Math.max(0, r.x1 - edge - 0.3);
  } else if (align === "right") {
    indRight = Math.max(r.cell ? -(r.overhang ?? 0) : 0, r.x1 - Math.max(...lines.map((l) => l.x1)));
  } else if (r.cell && r.overhang && ctx.mode === "exact") {
    // 居中：两侧对称地伸入边距，居中位置不变
    indLeft = indRight = -r.overhang;
  }
  // 逐行一致模式：行尾已由换行符固定，左对齐段落向右留出越界余量不会改变排版，只会防止因字宽差异导致的意外折行
  if (ctx.mode === "exact" && align === "left" && r.overhang && indRight === 0) indRight = -r.overhang;

  const lineRanges: Array<{ line: Line; start: number; end: number }> = [];
  lines.forEach((l, li) => {
    if (li > 0) {
      const prev = runs[runs.length - 1];
      if (ctx.mode === "exact") {
        if (prev && !prev.tab) prev.text = prev.text.replace(/\s+$/, "");
        runs.push({ text: "", br: true, ...plainFmt(spanFormat(lines[li - 1].spans[lines[li - 1].spans.length - 1], lines[li - 1].size, lines[li - 1].baseline)) });
      } else {
        if (prev && !prev.tab) prev.text = prev.text.replace(/\s+$/, "");
        const prevCh = prev?.text.slice(-1) ?? "";
        const nextText = l.spans[0].text.trimStart();
        const nextCh = nextText[0] ?? "";
        if (prev && !prev.tab && /[A-Za-z]-$/.test(prev.text) && /^[A-Za-z]/.test(nextText) && isSoftHyphen(prev.text.slice(0, -1), nextText, ctx.words, ctx.hyphenated, ctx.counts)) {
          // 排版断词（"sharp-" + "ness"）：去掉连字符直接相连，重新排版后不会在行中出现多余的连字符
          prev.text = prev.text.slice(0, -1);
          ctx.softHyphens++;
        } else if (prev && !prev.tab && prevCh !== "-" && prevCh !== "‐" && !CJK.test(prevCh) && !CJK.test(nextCh)) prev.text += " ";
      }
    }
    // 字符间距采样：未被两端对齐拉伸的行（左对齐段落的每行、两端对齐段落的末行）
    if (align !== "both" || li === lines.length - 1) {
      for (const sp of l.spans) {
        if (!sp.natural || sp.text.length < 6) continue;
        const ratio = (sp.x1 - sp.x0) / sp.natural;
        if (ratio < 0.85 || ratio > 1.15) continue;
        const k = fontKey(sp);
        if (!ctx.spacing.has(k)) ctx.spacing.set(k, []);
        ctx.spacing.get(k)!.push(((sp.x1 - sp.x0) - sp.natural) / sp.text.length);
      }
    }
    const lineStart = runs.length;
    l.spans.forEach((s, si) => {
      const fmt = spanFormat(s, l.size, l.baseline);
      if (s.img) Object.assign(fmt, fitInline(s.img, l, lineHeight, ctx.mode));
      let text = li === 0 && si === 0 ? s.text.replace(/^\s+/, "") : s.text;
      if (si > 0) {
        const p = l.spans[si - 1];
        const gap = s.x0 - p.x1;
        const leftAligned = align === "left" || align === "both";
        if (leftAligned && gap > Math.max(1.2 * l.size, 12)) {
          // 大间距 → 制表符，制表位就是原文字的横坐标；
          // 最后一段顶到右边界（公式编号"(1)"、页码等）→ 右对齐制表位，放在段尾的位置，
          // 不受字宽差异影响，永远不会被挤到下一行
          const isLastSeg = !l.spans.slice(si + 1).some((o, k) => o.x0 - l.spans[si + k].x1 > Math.max(1.2 * l.size, 12));
          const tab: TabStop = isLastSeg && r.x1 - l.x1 <= 3 && !r.cell
            ? { pos: Math.round((Math.min(l.x1, r.x1) - r.x0) * 20) / 20, align: "right" }
            : { pos: Math.round((s.x0 - r.x0) * 20) / 20, align: "left" };
          if (!tabs.some((t) => Math.abs(t.pos - tab.pos) < 1)) tabs.push(tab);
          const last = runs[runs.length - 1];
          if (last) last.text = last.text.replace(/\s+$/, "");
          runs.push({ text: "", tab: true, ...plainFmt(fmt) });
          text = text.replace(/^\s+/, "");
        } else if (gap > 0.18 * l.size && !/\s$/.test(runs[runs.length - 1]?.text ?? "") && !/^\s/.test(text)) {
          const pc = p.text.slice(-1), nc = text[0] ?? "";
          if (!(CJK.test(pc) && CJK.test(nc)) || gap > 0.6 * l.size) text = " " + text;
        } else if (gap < 0.05 * l.size && /\s$/.test(runs[runs.length - 1]?.text ?? "") && /^\s/.test(text)) {
          text = text.replace(/^\s+/, "");
        }
      }
      if (text) pushText(runs, text, fmt);
    });
    if (ctx.mode === "exact") lineRanges.push({ line: l, start: lineStart, end: runs.length });
    if (ctx.mode === "exact") {
      const m = lineMeasure(l);
      if (m && m.natural > 0 && m.actual / m.natural < 0.998) {
        const delta = (m.actual - m.natural) / Math.max(1, m.chars);
        const d = Math.floor(delta * 20) / 20;
        for (const r of runs.slice(lineStart)) if (!r.tab && !r.br) r.lineSpacing = d;
      }
    }
  });
  if (runs.length) {
    const last = runs[runs.length - 1];
    if (!last.tab) last.text = last.text.replace(/\s+$/, "");
  }
  // 本段的字符间距：取未被两端对齐拉伸的行（左对齐段落的全部行、两端对齐段落的末行）实测
  const plain = align === "both" ? lines.slice(-1) : lines;
  let dw = 0, dc = 0;
  for (const l of plain) {
    const m = lineMeasure(l);
    if (!m || m.natural <= 0) continue;
    const ratio = m.actual / m.natural;
    if (ratio < 0.85 || ratio > 1.15) continue;
    dw += m.actual - m.natural;
    dc += m.chars;
  }
  if (dc >= 12) {
    const d = Math.round((dw / dc) * 20) / 20;
    for (const r of runs) if (!r.tab && !r.br) r.paraSpacing = d;
  }

  // 每行的可用宽度（区域宽度减去缩进；首行另计首行缩进/悬挂）
  const lineFits: NonNullable<Para["lineFits"]> = [];
  if (ctx.mode === "exact" && !tabs.length && align !== "center") {
    lineRanges.forEach(({ line, start, end }, i) => {
      const m = lineMeasure(line);
      if (!m) return;
      const bulletSkipped = line.spans.length > 1 && BULLET.test(line.spans[0].text.trim()) && line.spans[1].x0 - line.spans[0].x1 > 0.45 * line.size;
      const startX = r.x0 + indLeft + (i === 0 && !bulletSkipped ? firstLine : 0);
      // 不折行校验按 Word 目标字体的字宽计算
      lineFits.push({ start, end, natural: m.word, chars: m.chars, avail: r.x1 - indRight - startX });
    });
  }
  // 单行且带制表位的段落（如"居中的公式 + 右端编号"）：起始缩进也改成制表位、缩进归零。
  // Word 的制表位从版心左边算起，而部分预览器从缩进处算起，统一成"无缩进 + 制表位"后各处一致
  if (lines.length === 1 && tabs.length && align === "left" && indLeft + firstLine > 0.5) {
    const lead = Math.round((indLeft + firstLine) * 20) / 20;
    if (!tabs.some((t) => Math.abs(t.pos - lead) < 1) && tabs.every((t) => t.pos > lead)) {
      tabs.push({ pos: lead, align: "left" });
      runs.unshift({ text: "", tab: true, ...plainFmt(spanFormat(first.spans[0], first.size, first.baseline)) });
      lineFits.forEach((f) => { f.start++; f.end++; });
      indLeft = 0; firstLine = 0;
    }
  }
  const kept = runs.map((x) => x.tab || x.br || !!x.text);
  // run 过滤后下标会变化：换算成过滤后的下标
  const remap = (k: number) => kept.slice(0, k).filter(Boolean).length;
  const lastLine = lines[lines.length - 1];
  return {
    // 末行排满到右边界：以文字块自己的右边缘为准（版心可能被页码等撑宽，比正文的右边缘更靠右）
    lastFull: Math.min(r.x1, blockEdge > r.x0 ? blockEdge : r.x1) - lastLine.x1 <= Math.max(3, 0.04 * W),
    kind: "p", runs: runs.filter((_, k) => kept[k]), align, indLeft, indRight, firstLine,
    spaceBefore: 0, lineHeight, tabs: tabs.sort((a, b) => a.pos - b.pos), size,
    lineFits: lineFits.map((f) => ({ ...f, start: remap(f.start), end: remap(f.end) })),
  };
}

/** Word 固定行距下，基线距行框顶部的距离（额外空间加在文字上方） */
const baselineOffset = (lineHeight: number, size: number) => lineHeight - 0.22 * size;

/** 把一组行在给定区域里排成段落，并按实测位置计算段前距。返回段落与结束位置 */
function layoutLines(lines: Line[], r: Region, cursor: number, ctx: Ctx, floats?: Map<Line, ImageBox[]>): { paras: Para[]; cursor: number } {
  const paras: Para[] = [];
  for (const { lines: group, edge } of groupParagraphs(lines, r)) {
    const p = buildPara(group, r, edge, ctx);
    const top = group[0].baseline - baselineOffset(p.lineHeight, p.size);
    p.spaceBefore = Math.max(0, top - cursor);
    // 旁边的浮动图片挂在这一段上：相对栏左边界、段落顶部定位
    for (const l of group) for (const img of floats?.get(l) ?? []) (p.floats ??= []).push({ img, dx: img.x0 - r.x0, dy: img.y0 - top });
    paras.push(p);
    cursor = Math.max(cursor, top) + p.lineHeight * group.length;
  }
  return { paras, cursor };
}

/** 独占一段的行内图片：段前距 = 与上一块的实测间距，左缩进 = 图片的横坐标（并排的多张图用制表位定位） */
function imagePara(imgs: ImageBox[], r: Region, cursor: number): Para {
  const sorted = [...imgs].sort((a, b) => a.x0 - b.x0);
  const y0 = Math.min(...sorted.map((i) => i.y0)), y1 = Math.max(...sorted.map((i) => i.y1));
  const p = tinyPara({ inlineImages: sorted, lineHeight: y1 - y0 });
  p.spaceBefore = Math.max(0, y0 - cursor);
  if (sorted.length === 1) p.indLeft = Math.max(-(r.overhang ?? 0), sorted[0].x0 - r.x0);
  else p.tabs = sorted.map((i) => ({ pos: Math.max(0, Math.round((i.x0 - r.x0) * 20) / 20), align: "left" as const }));
  return p;
}

/** 一个区域的第一段 / 最后一段文字（跳过占位空段）：流式模式下合并被切断的段落 */
function markRegion(bs: Block[]) {
  const real = bs.filter((b) => !(b.kind === "p" && b.tiny && !b.inlineImages && !b.borderBottom));
  const f = real[0], l = real[real.length - 1];
  if (f?.kind === "p" && !f.tiny) f.regionStart = true;
  if (l?.kind === "p" && !l.tiny) l.regionEnd = true;
}

// ---------------------------------------------------------------------------
// 表格 → Word 表格
// ---------------------------------------------------------------------------

function buildTable(t: TableDraft, fills: FillRect[], originX: number, ctx: Ctx): Table {
  const { xs, ys, hs, vs } = t;
  const nR = ys.length - 1, nC = xs.length - 1;
  const border = (s: HSeg | VSeg | null): Border | null => (s && s.color !== "none" ? { width: s.w, color: s.color } : null);

  // 每行的横向合并：相邻两格之间没有竖线 → 合并
  const rowGroups: Array<Array<{ c0: number; c1: number }>> = [];
  for (let r = 0; r < nR; r++) {
    const groups: Array<{ c0: number; c1: number }> = [];
    let c0 = 0;
    for (let c = 0; c < nC; c++) {
      const last = c === nC - 1;
      if (last || coverageV(vs, xs[c + 1], ys[r], ys[r + 1])) {
        groups.push({ c0, c1: c });
        c0 = c + 1;
      }
    }
    rowGroups.push(groups);
  }

  // 纵向合并：上下两格之间没有横线且横向跨度一致
  const vmerge: Array<Map<number, "restart" | "continue">> = rowGroups.map(() => new Map());
  for (let r = 1; r < nR; r++) {
    for (const g of rowGroups[r]) {
      const above = rowGroups[r - 1].find((a) => a.c0 === g.c0 && a.c1 === g.c1);
      if (above && !coverageH(hs, ys[r], xs[g.c0], xs[g.c1 + 1])) {
        if (!vmerge[r - 1].has(g.c0)) vmerge[r - 1].set(g.c0, "restart");
        vmerge[r].set(g.c0, "continue");
      }
    }
  }

  // 每个（合并后的）单元格区域收集文字
  const cellBox = (r: number, g: { c0: number; c1: number }) => {
    let r1 = r;
    while (r1 + 1 < nR && vmerge[r1 + 1].get(g.c0) === "continue") r1++;
    return { x0: xs[g.c0], x1: xs[g.c1 + 1], y0: ys[r], y1: ys[r1 + 1] };
  };

  const rows: Table["rows"] = [];
  for (let r = 0; r < nR; r++) {
    const cells: Cell[] = [];
    for (const g of rowGroups[r]) {
      const vm = vmerge[r].get(g.c0);
      const box = cellBox(r, g);
      const cell: Cell = {
        width: xs[g.c1 + 1] - xs[g.c0],
        gridSpan: g.c1 - g.c0 + 1,
        vMerge: vm,
        borders: {
          top: border(coverageH(hs, ys[r], xs[g.c0], xs[g.c1 + 1])),
          bottom: border(coverageH(hs, ys[r + 1], xs[g.c0], xs[g.c1 + 1])),
          left: border(coverageV(vs, xs[g.c0], ys[r], ys[r + 1])),
          right: border(coverageV(vs, xs[g.c1 + 1], ys[r], ys[r + 1])),
        },
        vAlign: "top",
        blocks: [],
      };
      const fill = fills.find((f) => f.x0 <= box.x0 + 2 && f.x1 >= box.x1 - 2 && f.y0 <= box.y0 + 2 && f.y1 >= box.y1 - 2)
        ?? fills.find((f) => (Math.min(f.x1, box.x1) - Math.max(f.x0, box.x0)) * (Math.min(f.y1, box.y1) - Math.max(f.y0, box.y0)) > 0.8 * (box.x1 - box.x0) * (box.y1 - box.y0));
      if (fill) cell.shade = fill.color;

      if (vm !== "continue") {
        const spans = t.spans.filter((s) => {
          const cx = (s.x0 + s.x1) / 2, cy = (s.top + s.bottom) / 2;
          return cx >= box.x0 && cx < box.x1 && cy >= box.y0 && cy < box.y1;
        });
        if (spans.length) {
          const lines = buildLines(spans);
          const region: Region = { x0: box.x0 + CELL_MARGIN, x1: box.x1 - CELL_MARGIN, overhang: CELL_MARGIN - 0.3, cell: true };
          const contentTop = lines[0].top, contentBottom = lines[lines.length - 1].bottom;
          const topGap = contentTop - box.y0, bottomGap = box.y1 - contentBottom;
          const h = box.y1 - box.y0;
          if (topGap > 0.15 * h && Math.abs(topGap - bottomGap) <= Math.max(3, 0.12 * h)) cell.vAlign = "center";
          else if (bottomGap < 0.5 * lines[0].size && topGap > 0.3 * h) cell.vAlign = "bottom";
          // 单元格内容从上边框下方开始（边框线宽占据单元格空间）
          const { paras } = layoutLines(lines, region, box.y0 + (cell.borders.top?.width ?? 0), ctx);
          if (cell.vAlign !== "top" && paras.length) paras[0].spaceBefore = 0;
          cell.blocks = paras;
        }
      }
      cells.push(cell);
    }
    rows.push({ height: ys[r + 1] - ys[r], cells });
  }
  return { kind: "table", indent: t.x0 - originX, colWidths: xs.slice(1).map((x, i) => x - xs[i]), rows, exactRows: ctx.mode === "exact" };
}

// ---------------------------------------------------------------------------
// 下划线 / 删除线 / 分隔线
// ---------------------------------------------------------------------------

function applyTextDecorations(spans: Span[], segs: Seg[], used: Set<Seg>): HSeg[] {
  const free = mergeH(segs.filter((s) => !used.has(s)));
  const rules: HSeg[] = [];
  for (const h of free) {
    let consumed = false;
    for (const s of spans) {
      const overlap = Math.min(h.x1, s.x1) - Math.max(h.x0, s.x0);
      if (overlap < 0.5 * (s.x1 - s.x0)) continue;
      if (h.y >= s.baseline - 0.5 && h.y <= s.baseline + 0.3 * s.size) { s.underline = true; consumed = true; }
      else if (h.y < s.baseline - 0.15 * s.size && h.y > s.baseline - 0.45 * s.size) { s.strike = true; consumed = true; }
    }
    if (!consumed && h.x1 - h.x0 >= 20) rules.push(h);
  }
  return rules;
}

// ---------------------------------------------------------------------------
// 分栏检测
// ---------------------------------------------------------------------------

type LineSide = "cross" | "two" | "L" | "R";

function classify(l: Line, x: number): LineSide {
  if (l.spans.some((s) => s.x0 < x - 1 && s.x1 > x + 1)) return "cross";
  const left = l.spans.filter((s) => (s.x0 + s.x1) / 2 < x), right = l.spans.filter((s) => (s.x0 + s.x1) / 2 >= x);
  if (left.length && right.length) {
    const gap = Math.min(...right.map((s) => s.x0)) - Math.max(...left.map((s) => s.x1));
    // 栏间距可能只有一个字宽（LaTeX 默认 10pt）：两端对齐的词间距通常小于 0.8 个字宽
    return gap < 0.8 * l.size ? "cross" : "two";
  }
  return left.length ? "L" : "R";
}

/**
 * 分栏检测：寻找一条纵向空白（栏间距），使尽可能多的行在其两侧都有文字而不跨越它。
 * 之后按行的纵向顺序切带：不跨越栏间距、且含至少 2 行"两侧都有文字"的连续行才算分栏带，
 * 这样页首的通栏标题/摘要与下面的双栏正文可以共存于同一页。
 */
export function detectGutter(lines: Line[], r: Region): { g0: number; g1: number; x: number } | null {
  if (lines.length < 6) return null;
  const W = r.x1 - r.x0;
  // 逐点打分：左侧整行都在 x 左边的"长行"数、右侧整行都在 x 右边的"长行"数取较小者，减去跨越 x 的行数。
  // 不要求左右两栏的基线对齐（很多双栏文档两栏基线并不对齐），取最高分区间的中点作为分栏线。
  const scores: Array<{ x: number; two: number }> = [];
  for (let x = r.x0 + 0.3 * W; x <= r.x0 + 0.7 * W; x += 1) {
    let L = 0, R = 0, cross = 0, two = 0;
    for (const l of lines) {
      const k = classify(l, x);
      if (k === "cross") { cross++; continue; }
      if (k === "two") two++;
      const left = l.spans.filter((s) => (s.x0 + s.x1) / 2 < x), right = l.spans.filter((s) => (s.x0 + s.x1) / 2 >= x);
      if (left.length && Math.max(...left.map((s) => s.x1)) - r.x0 > 0.2 * W) L++;
      if (right.length && r.x1 - Math.min(...right.map((s) => s.x0)) > 0.2 * W) R++;
    }
    // 两种证据取其大：两侧同行都有文字的行数（页首有通栏内容时可靠），或左右长行数减去跨栏行数（两栏基线不对齐时可靠）
    scores.push({ x, two: Math.max(two, Math.min(L, R) - cross) });
  }
  const top = Math.max(0, ...scores.map((s) => s.two));
  if (top < 4 || top < 0.15 * lines.length) return null;
  let bestRun: { a: number; b: number } | null = null;
  for (let i = 0; i < scores.length; ) {
    if (scores[i].two !== top) { i++; continue; }
    let j = i;
    while (j + 1 < scores.length && scores[j + 1].two === top) j++;
    if (!bestRun || j - i > bestRun.b - bestRun.a) bestRun = { a: i, b: j };
    i = j + 1;
  }
  const best = { x: (scores[bestRun!.a].x + scores[bestRun!.b].x) / 2, two: top };  const lefts: number[] = [], rights: number[] = [];
  for (const l of lines) {
    const k = classify(l, best.x);
    if (k === "cross") continue;
    const left = l.spans.filter((s) => (s.x0 + s.x1) / 2 < best.x), right = l.spans.filter((s) => (s.x0 + s.x1) / 2 >= best.x);
    if (left.length) lefts.push(Math.max(...left.map((s) => s.x1)));
    if (right.length) rights.push(Math.min(...right.map((s) => s.x0)));
  }
  const g0 = edgeOf(lefts, "max"), g1 = edgeOf(rights, "min");
  if (g1 - g0 < 6) return null;
  return { g0, g1, x: best.x };
}

// ---------------------------------------------------------------------------
// 整体
// ---------------------------------------------------------------------------

/**
 * 画成短横线的破折号（参考文献里表示"同上作者"的 "——"、排版软件画的长破折号）：
 * 位于某行文字的字身中部（基线上 0.15–0.65 字号）、处在两段文字之间的空隙里、长度 1–6 个字号 → 还原成 "—" 字符。
 * 否则它既不是文字的下划线 / 删除线，也不构成表格，会被丢掉
 */
function dashRules(pg: PageModel): void {
  const keep: typeof pg.segs = [];
  const added: Span[] = [];
  for (const g of pg.segs) {
    const y = (g.y0 + g.y1) / 2, len = g.x1 - g.x0;
    const row = g.horizontal && g.width <= 1.5 ? pg.spans.filter((s) => s.text.trim() && !s.img && y > s.baseline - 0.65 * s.size && y < s.baseline - 0.15 * s.size) : [];
    const near = row.filter((s) => s.x1 > g.x0 - 2.5 * s.size && s.x0 < g.x1 + 2.5 * s.size);
    const ref = near[0];
    const overlapped = row.some((s) => Math.min(s.x1, g.x1) - Math.max(s.x0, g.x0) > 0.5);
    if (!ref || overlapped || len < 0.8 * ref.size || len > 6 * ref.size) { keep.push(g); continue; }
    const n = Math.max(1, Math.round(len / ref.size)); // 长破折号宽 1 个字号
    added.push({ ...ref, text: "\u2014".repeat(n), x0: g.x0, x1: g.x1, bold: false, italic: false, natural: n * ref.size, wordW: n * ref.size, img: undefined, type3: undefined });
  }
  if (added.length) { pg.segs = keep; pg.spans = [...pg.spans, ...added]; }
}

export function buildDocModel(pages0: PageModel[], opts: { mode?: LayoutMode; raster?: boolean } = {}): DocModel {
  // 识别渲染区域时会从页面中移除元素：在副本上进行，原始提取结果保持不变（用于完整性校验、失败时回退）
  // 识别时会从页面中移除元素、插入行内公式：总在副本上进行，失败回退时原始结果不受影响
  const pages = pages0.map((p) => ({ ...p }));
  const allText = pages0.map((p) => p.spans.map((s) => s.text).join(" ")).join("\n");
  const ctx: Ctx = { mode: opts.mode ?? "exact", spacing: new Map(), words: collectWords(allText), hyphenated: new Set(), counts: wordCounts(allText), softHyphens: 0 };
  // 行尾断开的写法不算"本身带连字符"：只收集行中间出现的
  ctx.hyphenated = collectHyphenated(pages0.map((p) => p.spans.map((s) => s.text.replace(/[A-Za-z]+-\s*$/, "")).join(" ")).join("\n"));
  const flow = ctx.mode === "flow";
  const pageW = pages[0]?.width ?? 595, pageH = pages[0]?.height ?? 842;

  // 正文字号 / 字体（按字符数的众数）：用于识别图注、公式
  const sizeN = new Map<number, number>(), baseN = new Map<string, number>();
  for (const pg of pages) for (const s of pg.spans) {
    const n = s.text.replace(/\s/g, "").length;
    sizeN.set(s.size, (sizeN.get(s.size) ?? 0) + n);
    if (!s.math) baseN.set(rawBase(s), (baseN.get(rawBase(s)) ?? 0) + n);
  }
  const top1 = <K,>(m: Map<K, number>) => [...m.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
  const bodyBase = top1(baseN) ?? "";
  const mathCtx: MathCtx = { bodySize: top1(sizeN) ?? 10, bodyBase, bodyIsTeX: /^(CM|SF|LM|EC)/.test(bodyBase) };
  const rasterText: string[] = [];
  let figureCount = 0, equationCount = 0, inlineMathCount = 0;

  // 预处理：矢量图（渲染为图片）、表格、装饰线
  const pre = pages.map((pg) => {
    let figures: ImageBox[] = [];
    if (opts.raster) {
      const t0 = detectTables(pg);
      const tb: Box[] = t0.tables.map((t) => ({ x0: t.x0, y0: t.y0, x1: t.x1, y1: t.y1 }));
      tb.push(...detectRuledTables(pg, t0.used, new Set(t0.tables.flatMap((t) => t.spans))).map((t) => ({ x0: t.x0, y0: t.y0, x1: t.x1, y1: t.y1 })));
      figures = detectFigures(pg, mathCtx.bodySize, tb);
      figures.push(...detectImageGrids(pg, mathCtx.bodySize));
      // 带文字标注的位图（子图编号 (a)(b)…、坐标轴文字压在图片上）：连同标注整体渲染成一张图，
      // 否则标注会被当成正文、图片被当成"文字下方的底图"
      for (const im of [...pg.images]) {
        if (im.render) continue;
        const over = pg.spans.filter((s) => s.x1 > im.x0 && s.x0 < im.x1 && s.bottom > im.y0 && s.top < im.y1);
        if (!over.length) continue;
        const inside = over.every((s) => s.x0 >= im.x0 - 2 && s.x1 <= im.x1 + 2 && s.top >= im.y0 - 2 && s.bottom <= im.y1 + 2);
        const chars = over.reduce((n, s) => n + s.text.replace(/\s/g, "").length, 0);
        if (!inside || chars > 300) continue;
        const gone = removeInside(pg, im);
        figures.push(rasterBox({ x0: im.x0, y0: im.y0, x1: im.x1, y1: im.y1 }, pg.index, "figure", spansText(gone)));
      }
      for (const f of figures) if (f.alt) rasterText.push(f.alt);
      figureCount += figures.length;
    }
    // 没有可靠文字编码的字形（Type 3 字体画的公式、符号）：独立的公式成图，行内的作为行内图片插回正文行
    if (pg.masks?.length || pg.spans.some((sp) => sp.type3)) {
      const ink = detectMaskInk(pg, mathCtx.bodySize, !!opts.raster);
      figures.push(...ink.display);
      if (ink.inline.length) pg.spans = [...pg.spans, ...ink.inline];
      rasterText.push(...ink.alt);
      equationCount += ink.display.length;
      inlineMathCount += ink.inline.length;
    }
    dashRules(pg);
    const { tables, used } = detectTables(pg);
    const inTable = new Set(tables.flatMap((t) => t.spans));
    tables.push(...detectRuledTables(pg, used, inTable));
    const free = pg.spans.filter((s) => !inTable.has(s));
    const rules = applyTextDecorations(free, pg.segs, used);
    return { pg, tables, free, rules, figures };
  });

  // 版心：全部页面的内容边界
  const xs0: number[] = [], xs1: number[] = [], tops: number[] = [], bottoms: number[] = [];
  const tableX0: number[] = [], tableX1: number[] = [];
  for (const { pg, tables, free } of pre) {
    // 以"行"为单位统计左右边界（片段级统计会被行内的词间断开干扰）
    for (const l of buildLines(free)) { xs0.push(l.x0); xs1.push(l.x1); }
    for (const t of tables) { tableX0.push(t.x0); tableX1.push(t.x1); }
    const ys = [...free.map((s) => s.top), ...tables.map((t) => t.y0), ...pg.images.map((i) => i.y0)];
    const ye = [...free.map((s) => s.bottom), ...tables.map((t) => t.y1)];
    if (ys.length) tops.push(Math.min(...ys));
    if (ye.length) bottoms.push(Math.max(...ye));
  }
  const left = Math.max(9, Math.min(pageW * 0.4, xs0.length ? Math.min(edgeOf(xs0, "min"), ...tableX0) : 72));
  const maxX1 = xs1.length ? Math.max(edgeOf(xs1, "max"), ...tableX1) : pageW - 72;
  const right = Math.min(pageW - 9, Math.max(maxX1, pageW - left));
  const marginTop = Math.max(9, Math.min(...(tops.length ? tops : [72])) - 4);
  // 底边距留得尽量小：每页都以分页开始，多出的空间不会被使用，但能吸收排版误差，避免页脚被挤到下一页
  const marginBottom = Math.max(2, Math.min(14, pageH - Math.max(...(bottoms.length ? bottoms : [pageH - 72])) - 4));
  const body: Region = { x0: left, x1: right, overhang: Math.max(0, Math.min(36, pageW - right - 4)) };

  const blocks: Block[] = [];
  const stats = { pages: pages.length, paragraphs: 0, tables: 0, images: 0, vectorShapes: 0, rotatedText: 0, columnPages: 0, figures: 0, equations: 0 };
  let curSection: SectionProps = { cols: 1, colWidths: [right - left], colGap: 0, titlePg: flow };
  let curPage = 0;
  const sectKey = (x: SectionProps) => JSON.stringify([x.cols, x.colWidths, x.colGap]);

  // 分节：以一个极小段落承载上一节的属性；返回占用的高度，计入纵向位置
  let freshPage = false;
  const endSection = (next: SectionProps): number => {
    // 分栏节不跨页延续：每页的第一个分栏带都另起一节（跨页延续时，右栏顶部的段前距在部分排版引擎中会丢失）
    const force = !flow && freshPage && next.cols > 1;
    freshPage = false;
    if (!force && sectKey(next) === sectKey(curSection)) return 0;
    blocks.push(tinyPara({ sectionEnd: curSection }));
    curSection = { ...next, titlePg: flow && curPage === 0 };
    return TINY;
  };

  // 页眉 / 页脚：页面最上 / 最下 10% 内、与正文之间有明显空白的行，改为绝对定位，不参与文字流
  const pageLines = pre.map((x) => buildLines(x.free));
  const floating = pageLines.map((ls, i) => {
    const H = pre[i].pg.height;
    const out: Line[] = [];
    // 页脚：从下往上
    for (let k = ls.length - 1; k >= 0; k--) {
      const l = ls[k], prev = ls[k - 1];
      if (l.top < H * 0.9) break;
      if (!prev || l.top - prev.bottom > 1.2 * l.size || out.length) out.push(l);
      else break;
    }
    // 页眉：从上往下
    for (let k = 0; k < ls.length; k++) {
      const l = ls[k], next = ls[k + 1];
      if (l.bottom > H * 0.1) break;
      if (!next || next.top - l.bottom > 1.2 * l.size) out.push(l);
      else break;
    }
    return out;
  });
  // 流式模式：只有在多页重复出现的（页眉、页码、刊名卷期）才写进页眉页脚；只在一页出现的栏底标题、正文末行等放回正文
  if (flow && pages.length > 1) {
    const norm = (l: Line) => l.spans.map((sp) => sp.text).join(" ").replace(/\d+/g, "#").replace(/\s+/g, " ").trim();
    const seen = new Map<string, Set<number>>();
    floating.forEach((ls, i) => { for (const l of ls) { const k = norm(l); if (!seen.has(k)) seen.set(k, new Set()); seen.get(k)!.add(i); } });
    floating.forEach((ls, i) => { floating[i] = ls.filter((l) => (seen.get(norm(l))?.size ?? 0) >= 2); });
  }
  pageLines.forEach((ls, i) => {
    const f = new Set(floating[i]);
    pageLines[i] = ls.filter((l) => !f.has(l));
  });

  // 分栏在整篇文档中通常一致：先逐页检测，再用最常见的栏间距补全"只有一栏有文字"的页面
  const gutters = pageLines.map((ls) => detectGutter(ls, body));
  const known = gutters.filter((g): g is NonNullable<typeof g> => !!g);
  if (known.length) {
    // 全文共识：栏间距位置取中位数；栏边界取各页中最宽的左栏右边界 / 最靠左的右栏左边界的中位数
    const med = (xs: number[]) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
    const docG = { x: med(known.map((g) => g.x)), g0: med(known.map((g) => g.g0)), g1: med(known.map((g) => g.g1)) };
    pageLines.forEach((ls, i) => {
      const g = gutters[i];
      if (g) {
        // 本页的栏间距与全文一致时，用全文的栏边界（列表、参考文献等不齐的页面自己估计的边界偏窄）
        if (known.length >= 2 && Math.abs(g.x - docG.x) <= 12) gutters[i] = { ...docG };
        return;
      }
      if (ls.length < 4) return;
      const cross = ls.filter((l) => classify(l, docG.x) === "cross").length;
      if (cross <= 0.3 * ls.length) gutters[i] = { ...docG };
    });
  }

  // 流式模式：上下边距按正文（不含页眉页脚）计算，页眉页脚写进真正的页眉页脚部件
  let mTop = marginTop, mBottom = marginBottom;
  let headerDist: number | undefined, footerDist: number | undefined;
  if (flow) {
    const bt: number[] = [], bb: number[] = [], ht: number[] = [], fb: number[] = [];
    pre.forEach(({ pg, tables, figures }, i) => {
      const ls = pageLines[i];
      const pics = [...pg.images.filter((im) => !im.behind && im.y0 > pg.height * 0.1 && im.y1 < pg.height * 0.9), ...figures];
      const ys = [...ls.map((l) => l.top), ...tables.map((t) => t.y0), ...pics.map((f) => f.y0)];
      const ye = [...ls.map((l) => l.bottom), ...tables.map((t) => t.y1), ...pics.map((f) => f.y1)];
      if (ys.length) bt.push(Math.min(...ys));
      if (ye.length) bb.push(Math.max(...ye));
      for (const l of floating[i]) if (l.top < pg.height / 2) ht.push(l.top); else fb.push(l.bottom);
    });
    mTop = Math.max(9, (bt.length ? Math.min(...bt) : 72) - 2);
    mBottom = Math.max(9, pageH - (bb.length ? Math.max(...bb) : pageH - 72) - 2);
    if (ht.length) headerDist = Math.max(4, Math.min(...ht) - 1);
    if (fb.length) footerDist = Math.max(4, pageH - Math.max(...fb) - 1);
  }
  // 流式模式：每页的页眉页脚文字由页眉页脚部件（含页码域）统一表示，计入完整性校验
  if (flow) for (const fl of floating) for (const l of fl) rasterText.push(l.spans.map((sp) => sp.text).join(""));
  // 全文正文的上下界（按页统计会被"整页大图 + 一行图注"的页面误导）
  const allBody = pageLines.flat();
  const bodyTop = allBody.length ? Math.min(...allBody.map((l) => l.top)) : 0;
  const bodyBottom = allBody.length ? Math.max(...allBody.map((l) => l.bottom)) : pageH;
  /** 每页的页眉 / 页脚行（流式模式） */
  const hf: Array<{ head: Line[]; foot: Line[]; imgs: ImageBox[]; mid: number }> = [];
  /** 首页栏底的注释（作者单位、基金等）：流式模式下放进首页页脚，正文增减时位置不变 */
  let notes: Para[] = [];
  let notesBottom = 0;

  for (const [pi, { pg, tables, free, rules, figures }] of pre.entries()) {
    curPage = pi;
    stats.images += pg.images.length;
    // 与文字重叠的图片（水印、底图）放在文字下方
    // 文字真正压在图片上（文字中心落在图片内部）才算底图；只是边缘相接的（作者照片旁边的简介）不算
    for (const img of pg.images) img.behind = pg.spans.some((s) => {
      const cx = (s.x0 + s.x1) / 2, cy = (s.top + s.bottom) / 2;
      return s.text.trim() && cx > img.x0 + 2 && cx < img.x1 - 2 && cy > img.y0 + 2 && cy < img.y1 - 2;
    });
    stats.vectorShapes += pg.vectorShapes;
    stats.rotatedText += pg.rotatedText;
    stats.tables += tables.length;

    const lines = pageLines[pi];
    const gutter = gutters[pi];
    // 正文的上下界：之外（页眉页脚区域）的图形、图片属于页面装饰，以页面绝对坐标定位
    const inMarginZone = (y0: number, y1: number) => y1 <= bodyTop + 1 || y0 >= bodyBottom - 1;

    // 页首锚点段落：承载本页的浮动图片与页眉页脚文本框，并负责分页（逐页一致模式）
    // 表格之外的线条与色块：作为绝对定位的图形放在页首锚点段落上（不占文字流，位置与原文完全一致）
    // 流式模式：正文区域的横线改为段落边框（随文字移动），竖线、色块等装饰只保留首页页眉页脚区域的
    const inTableArea = (x0: number, y0: number, x1: number, y1: number) =>
      tables.some((t) => x0 >= t.x0 - 2 && x1 <= t.x1 + 2 && y0 >= t.y0 - 2 && y1 <= t.y1 + 2);
    const keepShape = (y0: number, y1: number) => !flow || (pi === 0 && inMarginZone(y0, y1));
    const shapes: NonNullable<Para["shapes"]> = [];
    const ruleSegs: HSeg[] = [];
    for (const h of rules) {
      if (flow && !inMarginZone(h.y, h.y)) ruleSegs.push(h);
      else if (keepShape(h.y, h.y)) shapes.push({ kind: "line", x0: h.x0, y0: h.y, x1: h.x1, y1: h.y, width: h.w, color: h.color, behind: false });
    }
    const usedV = new Set(tables.flatMap((t) => t.vs.flatMap((v) => v.src)));
    for (const v of mergeV(pg.segs.filter((sg) => !sg.horizontal && !usedV.has(sg)))) {
      if (v.y1 - v.y0 >= 8 && !inTableArea(v.x - 1, v.y0, v.x + 1, v.y1) && keepShape(v.y0, v.y1)) shapes.push({ kind: "line", x0: v.x, y0: v.y0, x1: v.x, y1: v.y1, width: v.w, color: v.color, behind: false });
    }
    for (const f of pg.fills) {
      if (inTableArea(f.x0, f.y0, f.x1, f.y1) || !keepShape(f.y0, f.y1)) continue;
      shapes.push({ kind: "rect", x0: f.x0, y0: f.y0, x1: f.x1, y1: f.y1, width: 0, color: f.color, behind: true });
    }
    const anchor = tinyPara({ pageBreakBefore: !flow && pg.index > 0, images: pg.images.filter((im) => im.behind), textBoxes: [], shapes });
    freshPage = true;
    const anchorAt = blocks.length;
    blocks.push(anchor);
    let cursor = (flow ? mTop : marginTop) + TINY;
    if (flow) {
      hf.push({ head: floating[pi].filter((l) => l.top < pg.height / 2), foot: floating[pi].filter((l) => l.top >= pg.height / 2), imgs: [], mid: pg.height / 2 });
    } else {
      // 页眉页脚：每一段（按大间距切开）一个绝对定位的文字框
      for (const l of floating[pi]) {
        const groups: Span[][] = [];
        for (const sp of l.spans) {
          const g = groups[groups.length - 1];
          if (g && sp.x0 - g[g.length - 1].x1 < 3 * l.size) g.push(sp);
          else groups.push([sp]);
        }
        for (const g of groups) {
          const gl = finishLine(g);
          const w = gl.x1 - gl.x0 + 6;
          const p = buildPara([gl], { x0: gl.x0, x1: gl.x0 + w }, gl.x0 + w, ctx);
          p.align = "left"; p.indLeft = 0; p.indRight = 0; p.firstLine = 0; p.tabs = [];
          const lh = gl.size * 1.17;
          p.lineHeight = lh;
          anchor.textBoxes!.push({ x: gl.x0, y: gl.baseline - baselineOffset(lh, gl.size), w, h: lh, para: p });
        }
      }
    }

    // 条目：行 / 表格 / 分隔线 / 图片，按纵向位置排序后切成"通栏带"与"分栏带"
    type Item = PageItem;
    const items: Array<Item & { top: number }> = [];
    // 行的类型：跨栏 / 两侧 / 仅左 / 仅右；连续的非跨栏行里至少有 2 行"两侧"才算分栏带
    const kinds = lines.map((l) => (gutter ? classify(l, gutter.x) : "cross"));
    // 通栏段落的末行（图题、摘要的最后一行很短，只落在左半边）：与上一行行距相同、左端对齐或同样居中 → 仍属通栏
    for (let i = 1; i < lines.length; i++) {
      if (kinds[i] === "cross" || kinds[i] === "two" || kinds[i - 1] !== "cross") continue;
      const l = lines[i], p = lines[i - 1];
      const pitch = l.baseline - p.baseline;
      const sameStyle = Math.abs(l.size - p.size) < 0.3;
      const aligned = Math.abs(l.x0 - p.x0) < 2 || Math.abs((l.x0 + l.x1) / 2 - (p.x0 + p.x1) / 2) < 3;
      const nextGap = lines[i + 1] ? lines[i + 1].baseline - l.baseline : Infinity;
      if (sameStyle && aligned && pitch > 0.9 * l.size && pitch < 1.4 * l.size && nextGap > pitch + 0.5) kinds[i] = "cross";
    }
    // 分栏行的判定：按"跨栏行"与大的纵向空白把行切成若干段；一段属于分栏区域，当且仅当
    //   段内至少 2 行两侧都有文字，或者至少 2 行恰好止于栏的右边界（两端对齐的栏内正文，如图下方只有一栏有字的区域）。
    // 通栏的标题、摘要末行等短行不会被误切进分栏节里，段落也不会被拆散。
    const colLine = new Array(lines.length).fill(false);
    if (gutter) {
      const atColEdge = (l: Line, k: LineSide) => (k === "L" ? Math.abs(l.x1 - gutter.g0) <= 2 : k === "R" ? Math.abs(l.x1 - right) <= 2 : false);
      for (let i = 0; i < lines.length; ) {
        if (kinds[i] === "cross") { i++; continue; }
        let j = i + 1;
        while (j < lines.length && kinds[j] !== "cross" && lines[j].top - lines[j - 1].bottom < 2.5 * Math.max(lines[j].size, lines[j - 1].size)) j++;
        const seg = lines.slice(i, j), ks = kinds.slice(i, j);
        const two = ks.filter((k) => k === "two").length;
        const edge = seg.filter((l, k) => atColEdge(l, ks[k])).length;
        const both = ks.filter((k) => k === "L").length >= 2 && ks.filter((k) => k === "R").length >= 2;
        // 只有一栏有文字、另一栏并排的是表格或图：同样属于分栏区域（例如左栏的表题紧挨右栏的表格）
        const segTop = Math.min(...seg.map((l) => l.top)), segBottom = Math.max(...seg.map((l) => l.bottom));
        // 或者紧挨着栏内的表格/图（两栏表题并排、表题与表格之间没有跨栏内容）
        const onlySide = ks.every((k) => k === "L") ? "L" : ks.every((k) => k === "R") ? "R" : null;
        const facing = [...tables, ...figures, ...pg.images.filter((im) => !im.render)].some((t) => {
          const side = t.x1 <= gutter.x + 2 ? "L" : t.x0 >= gutter.x - 2 ? "R" : "full";
          if (side === "full") return false;
          if (onlySide && side !== onlySide) return t.y0 < segBottom + 6 && t.y1 > segTop - 6;
          // 栏内图表紧挨着的题注（在图下方或表上方）：属于这一栏
          if (onlySide && side === onlySide) return (segTop - t.y1 > -2 && segTop - t.y1 < 18) || (t.y0 - segBottom > -2 && t.y0 - segBottom < 18);
          return !onlySide && t.y0 < segBottom + 12 && t.y1 > segTop - 12;
        });
        // 只有一栏有文字、且像栏内正文（有接近整栏宽的行）：例如最后一页左栏剩下的几行
        const colW = (gutter.g0 - left);
        const sideBody = !!onlySide && seg.length >= 3 && seg.every((l) => l.size <= mathCtx.bodySize * 1.15) &&
          seg.filter((l) => l.x1 - l.x0 > 0.85 * colW).length >= 2 && seg.every((l) => l.x1 - l.x0 < 1.05 * colW);
        if (two >= 2 || edge >= 2 || both || facing || sideBody) for (let k = i; k < j; k++) colLine[k] = true;
        i = j;
      }
    }
    // 分栏行：把每一栏的文字片段收集起来，在栏内重新成行（整页成行时，左右栏基线相近的片段可能被错误地配对）
    const leftSpans: Span[] = [], rightSpans: Span[] = [];
    lines.forEach((l, i) => {
      if (!gutter || !colLine[i]) { items.push({ kind: "line", line: l, side: "full", top: l.top }); return; }
      for (const sp of l.spans) ((sp.x0 + sp.x1) / 2 < gutter.x ? leftSpans : rightSpans).push(sp);
    });
    for (const l of buildLines(leftSpans)) items.push({ kind: "line", line: l, side: "L", top: l.top });
    for (const l of buildLines(rightSpans)) items.push({ kind: "line", line: l, side: "R", top: l.top });
    // 表格、分隔线、图片：完全落在某一栏内的归入该栏，否则为通栏
    const sideOf = (x0: number, x1: number): "full" | "L" | "R" =>
      !gutter ? "full" : x1 <= gutter.x + 2 ? "L" : x0 >= gutter.x - 2 ? "R" : "full";
    const regionOf = (side: "full" | "L" | "R"): Region =>
      side === "full" || !gutter ? { x0: left, x1: right } : side === "L" ? { x0: left, x1: gutter.g0 } : { x0: gutter.g1, x1: right };
    for (const t of tables) items.push({ kind: "table", t, side: sideOf(t.x0, t.x1), top: t.y0 });

    // 旁边有文字的图片（流式模式）：挂在段落上、文字环绕；旁边那几行的缩进由环绕产生，行本身改为从栏边界起排
    const floatBeside = (img: ImageBox, side: "full" | "L" | "R", besideLines: Array<Extract<Item, { kind: "line" }> & { top: number }>) => {
      const reg = regionOf(side);
      // 环绕间距：改写行边界之前，量图片与旁边文字行、下方文字行的实测距离
      const clamp = (v: number, hi: number) => Math.min(hi, Math.max(2, v));
      let gapR = Infinity, gapL = Infinity, gapB = Infinity;
      for (const it of besideLines) {
        if (it.line.x0 >= img.x1 - 6) gapR = Math.min(gapR, it.line.x0 - img.x1);
        else gapL = Math.min(gapL, img.x0 - it.line.x1);
      }
      for (const it of items) {
        if (it.kind !== "line" || besideLines.includes(it as any) || !(side === "full" || it.side === side || it.side === "full")) continue;
        if (it.line.top >= img.y1 - 1 && it.line.x0 < img.x1 + 1) gapB = Math.min(gapB, it.line.top - img.y1);
      }
      img.wrap = {
        r: Number.isFinite(gapR) ? clamp(gapR, 30) : 2,
        l: Number.isFinite(gapL) ? clamp(gapL, 30) : 2,
        // 下方文字行的 top 含字体上方留白，只取一半，避免整行文字被多推下去
        b: gapB < 20 ? clamp(gapB * 0.5, 8) : 2,
      };
      for (const it of besideLines) it.line = it.line.x0 >= img.x1 - 6 ? { ...it.line, x0: reg.x0 } : { ...it.line, x1: reg.x1 };
      items.push({ kind: "float", img, side, top: img.y0 - 0.01 });
    };
    const besideOf = (img: { x0: number; y0: number; x1: number; y1: number }, side: "full" | "L" | "R") =>
      items.filter((it): it is Extract<Item, { kind: "line" }> & { top: number } =>
        it.kind === "line" && (side === "full" || it.side === side || it.side === "full") &&
        it.line.bottom > img.y0 + 2 && it.line.top < img.y1 - 2 && (it.line.x0 >= img.x1 - 6 || it.line.x1 <= img.x0 + 6));

    // 行间公式、首字下沉：按区域识别，渲染为图片，从文字流中移除
    //   公式：独占一段的行内图片（随文字流移动，正文增减时不会与文字重叠）；
    //   首字下沉：逐行一致模式按页面坐标定位，流式模式挂在段落上、文字环绕
    const rasters: Box[] = [];
    if (opts.raster) {
      const extra: Box[] = [...pg.segs, ...pg.fills, ...pg.paths];
      for (const side of ["full", "L", "R"] as const) {
        const reg = regionOf(side);
        const sideLines = items.filter((it): it is Extract<typeof it, { kind: "line" }> & { top: number } => it.kind === "line" && it.side === side).map((it) => it.line);
        if (!sideLines.length) continue;
        const found = [
          ...detectEquations(sideLines, reg, mathCtx, extra).map((e) => ({ ...e, kind: "equation" as const })),
          ...detectDropCaps(sideLines, mathCtx).map((d) => ({ lines: [d.line], box: d.box, kind: "dropcap" as const, rest: d.rest })),
        ];
        for (const f of found) {
          const gone = new Set(f.lines);
          for (let k = items.length - 1; k >= 0; k--) {
            const it = items[k];
            if (it.kind === "line" && gone.has(it.line)) items.splice(k, 1);
          }
          // 首字母与正文并在同一行时，只取首字母成图，其余正文放回
          const rest = "rest" in f ? f.rest : undefined;
          if (rest) items.push({ kind: "line", line: rest, side, top: rest.top });
          const capSpans = f.lines.flatMap((l) => l.spans).filter((sp) => !rest || !rest.spans.includes(sp));
          const alt = spansText(capSpans);
          if (alt) rasterText.push(alt);
          const box = rasterBox(f.box, pg.index, f.kind, alt);
          if (f.kind === "equation") {
            items.push({ kind: "image", imgs: [box], side, top: f.box.y0 });
            equationCount++;
          } else if (flow) {
            floatBeside(box, side, besideOf(f.box, side));
          } else {
            // 首字下沉前后的文字不能被并成同一段
            items.push({ kind: "break", side, top: f.box.y0 + 0.01 });
            anchor.images!.push(box);
          }
          rasters.push(f.box);
        }
      }
      if (rasters.length) {
        const within = (x0: number, y0: number, x1: number, y1: number) => rasters.some((r) => x0 >= r.x0 - 1 && x1 <= r.x1 + 1 && y0 >= r.y0 - 1 && y1 <= r.y1 + 1);
        anchor.shapes = anchor.shapes!.filter((sh) => !within(Math.min(sh.x0, sh.x1), Math.min(sh.y0, sh.y1), Math.max(sh.x0, sh.x1), Math.max(sh.y0, sh.y1)));
        for (let k = ruleSegs.length - 1; k >= 0; k--) if (within(ruleSegs[k].x0, ruleSegs[k].y, ruleSegs[k].x1, ruleSegs[k].y)) ruleSegs.splice(k, 1);
      }
    }
    // 流式模式：正文区域的横线 → 段落边框
    for (const h of ruleSegs) items.push({ kind: "rule", h, side: sideOf(h.x0, h.x1), top: h.y });

    // 图片（照片、矢量图）：
    //   - 页眉页脚区域的（徽标等）：页面绝对坐标；
    //   - 旁边有文字的：逐行一致模式用页面绝对坐标，流式模式挂在段落上、文字环绕；
    //   - 其余：独占一段的行内图片，并排的几张放在同一段里
    const pics = [...pg.images.filter((im) => !im.behind), ...figures].sort((a, b) => a.y0 - b.y0);
    const rest: ImageBox[] = [];
    for (const img of pics) {
      const side = sideOf(img.x0, img.x1);
      // 流式模式：页眉页脚区域的图片（徽标）写进页眉页脚，每页都在原位置，不随正文移动
      if (inMarginZone(img.y0, img.y1)) { (flow ? hf[pi].imgs : anchor.images!).push(img); continue; }
      const beside = besideOf(img, side);
      if (beside.length) {
        if (flow) floatBeside(img, side, beside);
        else anchor.images!.push(img);
        continue;
      }
      rest.push(img);
    }
    const groups: ImageBox[][] = [];
    for (const img of rest) {
      const g = groups.find((gr) => gr.some((o) => sideOf(o.x0, o.x1) === sideOf(img.x0, img.x1) &&
        Math.min(o.y1, img.y1) - Math.max(o.y0, img.y0) > 0.5 * Math.min(o.y1 - o.y0, img.y1 - img.y0)));
      if (g) g.push(img); else groups.push([img]);
    }
    for (const g of groups) {
      const overlap = g.some((a, i) => g.some((b, j) => j > i && Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0) > 2));
      if (overlap) { anchor.images!.push(...g); continue; } // 互相叠放的图片：保持页面绝对坐标
      const x0 = Math.min(...g.map((i) => i.x0)), x1 = Math.max(...g.map((i) => i.x1));
      items.push({ kind: "image", imgs: g, side: sideOf(x0, x1), top: Math.min(...g.map((i) => i.y0)) });
    }
    items.sort((a, b) => a.top - b.top);

    // 首页栏底的注释（流式模式）：栏内最后几行字号小于正文、上方有空白或短横线隔开 → 移到首页页脚
    if (flow && pi === 0) {
      for (const side of ["L", "R", "full"] as const) {
        const own = items.filter((it) => it.side === side);
        let k = own.length;
        const small = (it: Item) => it.kind === "line" && it.line.size <= mathCtx.bodySize * 0.92;
        while (k > 0 && small(own[k - 1])) k--;
        if (k === own.length || k === 0) continue;
        const first = own[k] as Extract<Item, { kind: "line" }> & { top: number };
        const prev = own[k - 1];
        if (first.line.top < pg.height * 0.65 || /^(fig(ure)?|table|图|表)\b/i.test(spansText(first.line.spans))) continue;
        let start = k;
        if (prev.kind === "rule") start = k - 1;
        else if (!(prev.kind === "line" && first.line.top - prev.line.bottom > 0.8 * first.line.size)) continue;
        const moved = own.slice(start);
        for (const it of moved) items.splice(items.indexOf(it), 1);
        const reg = regionOf(side);
        const out = layoutItems(moved, reg, moved[0].top - 0.01, [], ctx);
        for (const b of out.blocks) {
          if (b.kind !== "p") continue;
          b.indLeft += reg.x0 - left;
          b.indRight += right - reg.x1;
          b.tabs = b.tabs.map((t) => ({ ...t, pos: t.pos + reg.x0 - left }));
          if (!notes.length && b === out.blocks[0]) b.spaceBefore = 0;
          notes.push(b);
        }
        notesBottom = Math.max(notesBottom, ...moved.map((it) => (it.kind === "line" ? it.line.bottom : it.top)));
      }
    }

    // 通栏条目把页面切成若干带；分栏条目只有在两侧都有内容的区域才构成分栏带
    const bands: Array<{ cols: boolean; items: typeof items }> = [];
    for (const it of items) {
      const cols = it.side !== "full";
      const last = bands[bands.length - 1];
      if (last && last.cols === cols) last.items.push(it);
      else bands.push({ cols, items: [it] });
    }
    if (bands.some((b) => b.cols)) stats.columnPages++;
    // 流式模式：每页的第一块紧接上一页的内容，不保留页面顶部的空白
    if (flow && pi > 0 && items.length) cursor = items[0].top;
    const pageFirst = blocks.length;

    for (const band of bands) {
      if (!band.cols) {
        cursor += endSection({ cols: 1, colWidths: [right - left], colGap: 0 });
        const out = layoutItems(band.items, body, cursor, pg.fills, ctx);
        if (flow) markRegion(out.blocks);
        blocks.push(...out.blocks);
        cursor = out.cursor;
      } else {
        const g = gutter!;
        const lw = g.g0 - left + 0.3, rw = right - g.g1 + 0.3;
        if (flow) {
          // 两栏共同的上方空白放在分节之前（单栏节里的空段），左右两栏的顶部对齐
          const gap = Math.min(...band.items.map((i) => i.top)) - cursor - 3;
          if (gap > 2) { blocks.push(spacer(gap)); cursor += gap; }
        }
        cursor += endSection({ cols: 2, colWidths: [Math.round(lw * 2) / 2, Math.round(rw * 2) / 2], colGap: Math.round((g.g1 - g.g0 - 0.6) * 2) / 2 });
        const start = cursor;
        const L = layoutItems(band.items.filter((i) => i.side === "L"), { x0: left, x1: left + lw, overhang: Math.max(0, g.g1 - g.g0 - 4) }, start, pg.fills, ctx);
        // 流式模式：右栏紧接左栏排（同一个分栏节内连续流动），段前距从本带顶部算起
        const rStart = flow ? Math.min(...band.items.map((i) => i.top)) : start;
        const R = layoutItems(band.items.filter((i) => i.side === "R"), { x0: g.g1 - 0.3, x1: right, overhang: body.overhang }, rStart, pg.fills, ctx);
        if (flow) {
          markRegion(L.blocks);
          markRegion(R.blocks);
          blocks.push(...L.blocks, ...R.blocks);
          cursor = Math.max(L.cursor, R.cursor);
          continue;
        }
        // 段前距统一改为"固定高度的空段"：
        //   - LibreOffice 把分栏节第一段的段前距放在栏区域之上（右栏顶部随之下移），分栏符后的段前距则被吞掉；
        //   - Word 则保留两者。空段的行高在两种排版引擎中都不会被压缩，位置一致。
        const hoistable = (f: Block | undefined): f is Para => f?.kind === "p" && (!f.tiny || !!f.inlineImages) && f.spaceBefore > 0.5;
        const hoist = (bs: Block[], lead: Para[]) => {
          const f = bs[0];
          if (hoistable(f)) {
            lead.push(spacer(f.spaceBefore - (lead.length ? TINY : 0)));
            f.spaceBefore = 0;
          }
          return [...lead, ...bs];
        };
        blocks.push(...hoist(L.blocks, []));
        if (R.blocks.length) {
          // 右栏：先放一个承载分栏符的极小段落，再放占位空段与右栏内容
          const lead = [tinyPara({ columnBreakBefore: true })];
          blocks.push(...(hoistable(R.blocks[0]) ? hoist(R.blocks, lead) : [...lead, ...R.blocks]));
          R.cursor += TINY;
        }
        cursor = Math.max(L.cursor, R.cursor);
      }
    }
    // 流式模式：记下原 PDF 的分页位置
    if (flow && pi > 0) {
      const f = blocks.slice(pageFirst).find((b): b is Para => b.kind === "p" && (!b.tiny || !!b.inlineImages));
      if (f) f.pageMark = true;
    }
    // 流式模式：没有挂任何东西的页首锚点段落不需要
    if (flow && !anchor.images!.length && !anchor.shapes!.length && !anchor.textBoxes!.length) blocks.splice(anchorAt, 1);
  }

  // 流式模式：合并被分栏、分页切断的段落（左栏末段 + 右栏首段、上一页末段 + 下一页首段）
  if (flow) mergeSplitParagraphs(blocks, ctx);
  // 流式模式：表题与下面的表格 / 表格图片、图片与下面的图题保持在同一页
  if (flow) {
    const capRe = /^\s*((table|tab\.)\s*[\dIVXLC]+|TABLE\s+[IVXLC\d]+|表\s*\d+)/i, figRe = /^\s*((fig\.?|figure)\s*\d+|图\s*\d+)/i;
    const real = blocks.filter((b) => !(b.kind === "p" && b.tiny && !b.inlineImages));
    real.forEach((b, i) => {
      const n = real[i + 1];
      if (!n || b.kind !== "p") return;
      const t = textOfPara(b);
      const nIsObj = n.kind === "table" || (n.kind === "p" && !!n.inlineImages);
      if (capRe.test(t) && nIsObj) b.keepNext = true;
      if (b.inlineImages && n.kind === "p" && figRe.test(textOfPara(n))) b.keepNext = true;
      // 题注分成了两段（第二段是题注的续行）
      if ((capRe.test(t) || figRe.test(t)) && n.kind === "p" && !n.inlineImages && n.size === b.size && textOfPara(n).length < 200 && real[i + 2] && (real[i + 2].kind === "table" || (real[i + 2] as Para).inlineImages)) b.keepNext = (n as Para).keepNext = true;
    });
  }

  // 字符间距：同一字体/字号/字形的样本中位数明显偏离 0 时应用（例如样式里设置了紧缩 0.1 磅）
  const paras = collectParas(blocks);
  const spacingOf = new Map<string, number>();
  for (const [k, v] of ctx.spacing) {
    const m = median(v);
    if (v.length >= 2 && Math.abs(m) >= 0.03) spacingOf.set(k, Math.round(m * 20) / 20);
  }
  for (const p of paras) for (const r of p.runs) {
    if (r.tab || r.br) continue;
    const sp = spacingOf.get(fontKey(r));
    const base = r.paraSpacing ?? sp ?? 0;
    const v = r.lineSpacing !== undefined ? Math.min(r.lineSpacing, base) : base;
    r.spacing = Math.abs(v) >= 0.03 ? v : undefined;
  }
  // 逐行一致模式的安全余量：两端对齐段落的文字略微收紧 0.04 pt/字，Word 会把行重新拉满到栏宽，
  // 视觉上不变，但可避免因度量误差导致某一行比原来长出一点而被折成两行
  if (ctx.mode === "exact") {
    for (const p of paras) {
      if (p.align !== "both") continue;
      for (const r of p.runs) if (!r.tab && !r.br) r.spacing = Math.round(((r.spacing ?? 0) - 0.04) * 100) / 100;
    }
    // 不折行校验：自然宽度 + 字符间距 ≤ 可用宽度；超出的行按差值再收紧（字宽数据来自 PDF 自身的字体度量）
    for (const p of paras) {
      for (const f of p.lineFits ?? []) {
        const rs = p.runs.slice(f.start, f.end).filter((r) => !r.tab && !r.br);
        if (!rs.length) continue;
        const sp = Math.min(...rs.map((r) => r.spacing ?? 0));
        const need = f.natural + sp * f.chars;
        if (need > f.avail - 0.3) {
          const v = Math.floor(((f.avail - 0.3 - f.natural) / Math.max(1, f.chars)) * 20) / 20;
          if (v > -1) for (const r of rs) r.spacing = Math.min(r.spacing ?? 0, v);
        }
      }
      delete p.lineFits;
    }
  }

  // 标题识别：字号明显大于正文的短段落
  stats.paragraphs = paras.filter((p) => !p.tiny).length;
  const sizeChars = new Map<number, number>();
  const fontChars = new Map<string, number>();
  for (const p of paras) for (const r of p.runs) {
    sizeChars.set(r.size, (sizeChars.get(r.size) ?? 0) + r.text.length);
    fontChars.set(r.font, (fontChars.get(r.font) ?? 0) + r.text.length);
  }
  const bodySize = [...sizeChars.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 11;
  const bodyFont = [...fontChars.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "Times New Roman";
  const topLevel = blocks.filter((b): b is Para => b.kind === "p" && !b.tiny);
  // 标题：按"视觉样式"（字号、粗细、字体、颜色）分组；字号明显大于正文，或带编号且样式与正文不同的短段落
  const sig = (p: Para) => {
    const r = p.runs.find((x) => x.text.trim()) ?? p.runs[0];
    return r ? `${p.size}|${r.bold ? 1 : 0}|${r.font}|${r.color}|${r.italic ? 1 : 0}` : "";
  };
  const cands = topLevel.filter((p) => isHeadingCandidate(p, bodySize, bodyFont));
  const sigs = [...new Set(cands.map(sig))].sort((a, b) => Number(b.split("|")[0]) - Number(a.split("|")[0]) || Number(b.split("|")[1]) - Number(a.split("|")[1]));
  // 同一样式出现过多（像正文一样频繁）的不算标题
  const counts = new Map<string, number>();
  for (const p of topLevel) counts.set(sig(p), (counts.get(sig(p)) ?? 0) + 1);
  const valid = sigs.filter((g) => (counts.get(g) ?? 0) <= Math.max(3, topLevel.length * 0.15));
  for (const p of cands) {
    const k = valid.indexOf(sig(p));
    if (k >= 0) p.heading = Math.min(3, k + 1);
  }

  // 流式模式：页眉页脚
  const hfOut = flow ? buildHeaders(hf, notes, notesBottom, { x0: left, x1: right }, right, ctx, headerDist ?? mTop / 2) : undefined;
  // 去掉的排版断词连字符：计入完整性校验
  if (ctx.softHyphens) rasterText.push("-".repeat(ctx.softHyphens));

  return {
    pageW, pageH,
    margins: flow ? { top: mTop, bottom: mBottom, left, right: pageW - right } : { top: marginTop, bottom: marginBottom, left, right: pageW - right },
    blocks,
    finalSection: curSection,
    bodyFont, bodySize,
    stats: { ...stats, figures: figureCount, equations: equationCount, inlineMath: inlineMathCount },
    rasterText: rasterText.join("\n"),
    mode: ctx.mode,
    ...(hfOut ? { headers: hfOut.headers, footers: hfOut.footers, pageNumStart: hfOut.pageNumStart, headerDist, footerDist } : {}),
  };
}

const textOfPara = (p: Para) => p.runs.map((r) => (r.tab ? "\t" : r.text)).join("");

/** 流式模式：合并被分栏、分页切断的段落 */
function mergeSplitParagraphs(blocks: Block[], ctx: Ctx) {
  const body = blocks.filter((b): b is Para => b.kind === "p" && !b.tiny && textOfPara(b).length > 40);
  if (!body.length) return;
  const sizeN = new Map<number, number>(), fontN = new Map<string, number>();
  for (const p of body) {
    sizeN.set(p.size, (sizeN.get(p.size) ?? 0) + 1);
    const f = p.runs.find((r) => r.text.trim())?.font ?? "";
    fontN.set(f, (fontN.get(f) ?? 0) + 1);
  }
  const bodySize = [...sizeN.entries()].sort((a, b) => b[1] - a[1])[0][0];
  const bodyFont = [...fontN.entries()].sort((a, b) => b[1] - a[1])[0][0];
  // 正文段落普遍首行缩进（如 IEEE）：栏首没有缩进的段落就是上一栏的延续
  const indented = body.filter((p) => Math.abs(p.size - bodySize) < 0.3 && p.firstLine > 0.5 * p.size).length;
  const indentStyle = indented >= 3 && indented >= 0.25 * body.length;

  const canMerge = (a: Para, b: Para) => {
    if (!a.lastFull || a.tabs.length || b.tabs.length || b.floats?.length) return false;
    if (!(a.align === "both" || a.align === "left") || !(b.align === "both" || b.align === "left")) return false;
    if (Math.abs(a.size - b.size) > 0.3 || Math.abs(b.firstLine) > 0.3 * b.size) return false;
    if (isHeadingCandidate(a, bodySize, bodyFont) || isHeadingCandidate(b, bodySize, bodyFont)) return false;
    const at = textOfPara(a).trimEnd(), bt = textOfPara(b).trimStart();
    if (!at || !bt || BULLET.test(bt.split(/\s/)[0])) return false;
    // 图题、表题自成一段（"TABLE V: …""Fig. 2: …"），不能并进上一段
    if (/^((fig\.?|figure|table)\s*[\dIVXLC]+[a-z]?\s*[:.．：]|TABLE\s+[IVXLC\d]+\b|[图表]\s*\d+)/i.test(bt)) return false;
    return indentStyle || !/[.!?。！？:：]["'”’)\]]?$/.test(at) || /^[a-z]/.test(bt);
  };
  const out: Block[] = [];
  let last: Para | null = null;
  for (const b of blocks) {
    if (b.kind === "p" && b.tiny) {
      if (b.sectionEnd || b.inlineImages || b.borderBottom) last = null;
      out.push(b);
      continue;
    }
    if (b.kind === "p" && last && b.regionStart && last.regionEnd && canMerge(last, b)) {
      const tail = last.runs[last.runs.length - 1], head = b.runs[0];
      if (tail && head && !tail.tab && !head.tab) {
        tail.text = tail.text.replace(/\s+$/, "");
        const ht = head.text.replace(/^\s+/, "");
        if (/[A-Za-z]-$/.test(tail.text) && /^[A-Za-z]/.test(ht) && isSoftHyphen(tail.text.slice(0, -1), ht, ctx.words, ctx.hyphenated, ctx.counts)) {
          tail.text = tail.text.slice(0, -1);
          ctx.softHyphens++;
        } else if (!/[-‐]$/.test(tail.text) && !(CJK.test(tail.text.slice(-1)) && CJK.test(ht[0] ?? ""))) tail.text += " ";
        head.text = ht;
      }
      if (b.pageMark && head) last.runs.push({ ...head, text: "", pageMark: true });
      last.runs.push(...b.runs.filter((r) => r.tab || r.text));
      last.regionEnd = b.regionEnd;
      last.lastFull = b.lastFull;
      continue;
    }
    out.push(b);
    last = b.kind === "p" ? b : null;
  }
  blocks.splice(0, blocks.length, ...out);
}

/** 流式模式：由每页的页眉页脚行生成首页 / 奇数页 / 偶数页的页眉页脚，页码改为 PAGE 域 */
function buildHeaders(hf: Array<{ head: Line[]; foot: Line[]; imgs: ImageBox[]; mid: number }>, notes: Para[], notesBottom: number, body: Region, right: number, ctx: Ctx, headerTop: number):
  { headers: HeaderSet; footers: HeaderSet; pageNumStart?: number } | undefined {
  const n = hf.length;
  if (!hf.some((x) => x.head.length || x.foot.length || x.imgs.length) && !notes.length) return undefined;
  const lineText = (l: Line) => l.spans.map((s) => s.text).join(" ");
  // 页码：印刷的数字 - 页序号 在多数页上一致
  let pageNumStart: number | undefined;
  if (n >= 2) {
    const off = new Map<number, Set<number>>();
    hf.forEach((x, i) => {
      for (const l of [...x.head, ...x.foot]) for (const m of lineText(l).matchAll(/(?<![\d.,])\d{1,4}(?![\d.,]?\d)/g)) {
        const k = Number(m[0]) - i;
        if (!off.has(k)) off.set(k, new Set());
        off.get(k)!.add(i);
      }
    });
    const best = [...off.entries()].sort((a, b) => b[1].size - a[1].size)[0];
    if (best && best[1].size >= Math.max(2, Math.ceil(n * 0.5))) pageNumStart = best[0];
  }
  const printed = (i: number) => i + (pageNumStart ?? 1);

  const markPage = (p: Para, num: number) => {
    const s = String(num);
    const re = new RegExp(`(?<!\\d)${s}(?!\\d)`, "g");
    let idx = p.runs.findIndex((r) => !r.tab && r.text.trim() === s);
    let at = -1;
    if (idx < 0) {
      for (let k = p.runs.length - 1; k >= 0 && idx < 0; k--) {
        const r = p.runs[k];
        if (r.tab) continue;
        const ms = [...r.text.matchAll(re)];
        if (ms.length) { idx = k; at = ms[ms.length - 1].index!; }
      }
    } else at = p.runs[idx].text.indexOf(s);
    if (idx < 0) return false;
    const r = p.runs[idx];
    const parts: Run[] = [];
    if (at > 0) parts.push({ ...r, text: r.text.slice(0, at) });
    parts.push({ ...r, text: s, field: "PAGE" });
    if (at + s.length < r.text.length) parts.push({ ...r, text: r.text.slice(at + s.length) });
    p.runs.splice(idx, 1, ...parts);
    return true;
  };
  const mk = (ls: Line[], i: number, kind: "head" | "foot" = "head"): Para[] => {
    const out = mkText(ls, i);
    const imgs = (hf[i]?.imgs ?? []).filter((im) => (kind === "head") === ((im.y0 + im.y1) / 2 < hf[i].mid));
    if (imgs.length) {
      if (!out.length) out.push(tinyPara({}));
      const host = out[0];
      if (kind === "head") {
        // 相对页边距 / 第一段顶部：浏览器预览按偏移量直接定位，与 Word 一致
        const top = ls.length ? Math.min(...ls.map((l) => l.baseline - baselineOffset(l.size * 1.17, l.size))) : headerTop;
        host.hfImages = imgs.map((img) => ({ img, dx: img.x0 - body.x0, dy: img.y0 - top }));
      } else host.images = imgs;
    }
    return out;
  };
  const mkText = (ls: Line[], i: number): Para[] => {
    const out: Para[] = [];
    let cur: number | null = null;
    for (const l of [...ls].sort((a, b) => a.top - b.top)) {
      const p = buildPara([l], body, right, ctx);
      p.lineHeight = Math.round(l.size * 1.17 * 20) / 20;
      const top = l.baseline - baselineOffset(p.lineHeight, l.size);
      p.spaceBefore = cur === null ? 0 : Math.max(0, top - cur);
      cur = top + p.lineHeight;
      if (pageNumStart !== undefined) markPage(p, printed(i));
      delete p.lineFits;
      out.push(p);
    }
    return out;
  };
  const imgSig = (p: Para) => [...(p.hfImages ?? []).map((h) => h.img), ...(p.images ?? [])].map((im) => `${Math.round(im.x0)},${Math.round(im.y0)},${Math.round(im.x1)},${Math.round(im.y1)}`).join(";");
  const sigOf = (ps: Para[]) => ps.map((p) => `${p.align}|${p.tabs.map((t) => t.align + t.pos).join(",")}|${p.runs.map((r) => (r.field ? "#" : r.tab ? "\t" : r.text)).join("")}|${imgSig(p)}`).join("\n");
  const sig = (i: number) => sigOf(mk(hf[i].head, i)) + "||" + sigOf(mk(hf[i].foot, i, "foot"));
  const rest = [...Array(n).keys()].slice(1);
  const odd = rest.filter((i) => printed(i) % 2 === 1), even = rest.filter((i) => printed(i) % 2 === 0);
  // 每组取出现最多的页眉页脚（个别页面不同时以多数为准）
  const sigs = new Map<number, string>();
  const sigAt = (i: number) => { if (!sigs.has(i)) sigs.set(i, sig(i)); return sigs.get(i)!; };
  const major = (ids: number[]) => {
    const n = new Map<string, number[]>();
    for (const i of ids) { const k = sigAt(i); if (!n.has(k)) n.set(k, []); n.get(k)!.push(i); }
    return [...n.values()].sort((a, b) => b.length - a.length)[0]?.[0];
  };
  const oddRep = major(odd), evenRep = major(even);
  const evenOdd = oddRep !== undefined && evenRep !== undefined && sigAt(oddRep) !== sigAt(evenRep);
  const defPage = evenOdd ? oddRep! : major(rest) ?? 0;
  const rep0 = evenOdd && printed(0) % 2 === 0 ? evenRep! : defPage;
  const titlePg = notes.length > 0 || (n > 1 && sigAt(0) !== sigAt(rep0));
  const headers: HeaderSet = { default: mk(hf[defPage].head, defPage) };
  const footers: HeaderSet = { default: mk(hf[defPage].foot, defPage, "foot") };
  if (evenOdd) { headers.even = mk(hf[evenRep!].head, evenRep!); footers.even = mk(hf[evenRep!].foot, evenRep!, "foot"); }
  if (titlePg) {
    headers.first = mk(hf[0].head, 0);
    const f0 = mk(hf[0].foot, 0, "foot");
    if (notes.length && f0.length && hf[0].foot.length) {
      const top = Math.min(...hf[0].foot.map((l) => l.top));
      f0[0].spaceBefore = Math.max(0, top - notesBottom - 2);
    }
    footers.first = [...notes, ...f0];
  }
  return { headers, footers, pageNumStart: pageNumStart !== undefined && pageNumStart !== 1 ? pageNumStart : undefined };
}

type PageItem = { side: "full" | "L" | "R" } & (
  | { kind: "line"; line: Line }
  | { kind: "table"; t: TableDraft }
  | { kind: "rule"; h: HSeg }
  | { kind: "break" }
  /** 独占一段的行内图片（行间公式、图） */
  | { kind: "image"; imgs: ImageBox[] }
  /** 旁边有文字的图片：挂在紧随其后的那一行所在的段落上，文字环绕 */
  | { kind: "float"; img: ImageBox }
);

/** 在一个区域（整页版心 / 某一栏）内按纵向顺序排布行、表格、分隔线 */
function layoutItems(items: Array<PageItem & { top: number }>, r: Region, cursor: number, fills: FillRect[], ctx: Ctx): { blocks: Block[]; cursor: number } {
  const blocks: Block[] = [];
  let pending: Line[] = [];
  const floats = new Map<Line, ImageBox[]>();
  let waiting: ImageBox[] = [];
  const flush = () => {
    if (!pending.length) return;
    const out = layoutLines(pending, r, cursor, ctx, floats);
    blocks.push(...out.paras);
    cursor = out.cursor;
    pending = [];
  };
  for (const it of items) {
    if (it.kind === "line") {
      if (waiting.length) { floats.set(it.line, waiting); waiting = []; }
      pending.push(it.line);
    } else if (it.kind === "float") waiting.push(it.img);
    else if (it.kind === "image") {
      flush();
      const p = imagePara(it.imgs, r, cursor);
      blocks.push(p);
      cursor = Math.max(cursor, Math.min(...it.imgs.map((i) => i.y0))) + p.lineHeight;
    } else if (it.kind === "break") flush();
    else if (it.kind === "table") {
      flush();
      const t = it.t;
      // 旁边并排有文字（如"照片框 + 作者简介"）：表格改为浮动定位，不占文字流，旁边的文字保持原位
      const beside = items.some((o) => o.kind === "line" && o.line.bottom > t.y0 + 2 && o.line.top < t.y1 - 2 && (o.line.x0 >= t.x1 - 1 || o.line.x1 <= t.x0 + 1));
      if (beside) {
        const tb = buildTable(t, fills, r.x0, ctx);
        tb.float = { x: t.x0, y: t.y0 };
        blocks.push(tb);
        continue;
      }
      const gap = it.t.y0 - cursor;
      if (gap > 1) { blocks.push(spacer(gap)); cursor += gap; }
      blocks.push(buildTable(it.t, fills, r.x0, ctx));
      cursor = it.t.y1;
    } else {
      flush();
      const p = tinyPara({ borderBottom: { width: it.h.w, color: it.h.color } });
      p.spaceBefore = Math.max(0, it.h.y - cursor - 1);
      p.indLeft = Math.max(0, it.h.x0 - r.x0);
      p.indRight = Math.max(0, r.x1 - it.h.x1);
      blocks.push(p);
      cursor = Math.max(cursor, it.h.y);
    }
  }
  flush();
  // 后面没有文字可挂的浮动图片：改为独占一段
  for (const img of waiting) {
    const p = imagePara([img], r, cursor);
    blocks.push(p);
    cursor = Math.max(cursor, img.y0) + p.lineHeight;
  }
  return { blocks, cursor };
}

const HEADING_NUM = /^(第[一二三四五六七八九十百\d]+[章节部分篇]|[一二三四五六七八九十]+[、.．]|（[一二三四五六七八九十]+）|\d+(\.\d+){0,3}[.)]?\s|[IVX]+\.\s|[A-Z]\.\s|APPENDIX|ACKNOWLEDG|REFERENCES)/;

function isHeadingCandidate(p: Para, bodySize: number, bodyFont: string): boolean {
  const text = p.runs.map((r) => r.text).join("").trim();
  if (!text || text.length > 100 || p.runs.some((r) => r.br)) return false;
  if (/[。；;，,]$/.test(text)) return false;
  const runs = p.runs.filter((r) => r.text.trim());
  const bold = runs.every((r) => r.bold);
  const distinct = bold || runs.some((r) => r.font !== bodyFont || r.color !== "000000");
  if (p.size >= bodySize * 1.2 && (bold || p.size >= bodySize * 1.45)) return true;
  if (p.size >= bodySize * 1.1 && distinct) return true;
  return HEADING_NUM.test(text) && distinct && text.length <= 60 && p.size >= bodySize * 0.8;
}

function collectParas(blocks: Block[]): Para[] {
  const out: Para[] = [];
  for (const b of blocks) {
    if (b.kind === "p") { out.push(b); for (const t of b.textBoxes ?? []) out.push(t.para); }
    else for (const r of b.rows) for (const c of r.cells) out.push(...c.blocks);
  }
  return out;
}

const TINY = 1;
function tinyPara(extra: Partial<Para>): Para {
  return { kind: "p", runs: [], align: "left", indLeft: 0, indRight: 0, firstLine: 0, spaceBefore: 0, lineHeight: TINY, tabs: [], tiny: true, size: 1, ...extra };
}
function spacer(height: number): Para {
  return tinyPara({ lineHeight: Math.max(1, height) });
}
