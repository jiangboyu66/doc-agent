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

import type { FillRect, ImageBox, PageModel, PathBox, Seg, Span } from "./extract.js";
import { collectHyphenated, collectWords, isSoftHyphen, wordCounts } from "../../documents/hyphen.js";
import { romanPage, settleFurniture, toRoman, type FrontMatter, type Furniture, type FurniturePick } from "./furniture.js";
import { measureWord } from "./fonts.js";
import { detectDropCaps, detectEquations, detectIconPaths, detectFigures, detectImageGrids, detectMaskInk, rawBase, rasterBox, removeInside, spansText, textHighlights, type Box, type MathCtx } from "./raster.js";
import type { ScanPiece, ScanTextBlock } from "./scanLayer.js";

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
  smallCaps?: boolean;
  /** 字距加宽的文字实测的字符间距（pt），优先于按样式统计的间距 */
  tracking?: number;
  /** 域（页眉页脚中的页码、总页数） */
  field?: "PAGE" | "NUMPAGES";
  /** 域的数字格式：页码补零到的位数（PAGE \# "000000"） */
  fieldDigits?: number;
  /** 页码写成罗马数字（PAGE \* roman / ROMAN：前置部分的页脚"~i~"） */
  fieldFormat?: "roman" | "ROMAN";
  /** 原 PDF 的分页位置（写成 lastRenderedPageBreak，浏览器预览据此分页） */
  pageMark?: boolean;
  /** 行内图片（没有文字编码的公式 / 符号）：text 为占位符 U+FFFC */
  img?: ImageBox;
  /** 行内图片底边在基线以下的距离（pt），写成字符位置（降低） */
  imgLower?: number;
  /** 行内图片的缩放比例（比行距略高的图缩小到正好放进一行，不撑高行距、不被固定行距裁掉） */
  imgScale?: number;
  /** 文字底色（突出显示 / 字符底纹），RRGGBB */
  shade?: string;
}

export interface TabStop { pos: number; align: "left" | "right"; /** 点状前导符（目录的引导点） */ leader?: "dot" }
export interface Border { width: number; color: string }

export interface Para {
  kind: "p";
  runs: Run[];
  /** 制表位只用于首行编号后的正文（编号比悬挂缩进宽）：跨页合并时不算"带制表位的段落" */
  labelTab?: boolean;
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
  textBoxes?: Array<{ x: number; y: number; w: number; h: number; para: Para; /** 框里的文字按框宽折行、框随文字长高（没有文字层的页上的文字块） */ wrap?: boolean; /** 给调用方的框名（pdf2word-block 页.块） */ name?: string; /** 给调用方的说明（wp:docPr 的 descr：文字块的版面位置、行数、字号、行距的 JSON） */ descr?: string }>;
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
  /** 前置部分（罗马数字页码）的节：用前置部分的页眉页脚 */
  front?: boolean;
  /** 正文的第一节（前置部分之后）：从下一页开始，页码重新编号 */
  restart?: boolean;
  /** 前置部分的第一节（前面有封面）：从下一页开始，按罗马数字从前置部分第一页印的数编起 */
  frontStart?: boolean;
  /** 本节的纸张大小（与全文不同时：横向的附图页夹在纵向的页中间）；省略时用全文的 */
  pageW?: number;
  pageH?: number;
  /** 本节从新的一页开始（纸张大小变了的节：分节符本身换页） */
  newPage?: boolean;
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
  stats: {
    pages: number; paragraphs: number; tables: number; images: number; vectorShapes: number; rotatedText: number; columnPages: number; figures: number; equations: number; inlineMath?: number;
    /**
     * 扫描页的位图层（可编辑排版）：拆开的文字页、格线垫底 + 文字放进文本框的表单页（页码从 1 起）；作为行内图片插回的漏认字、
     * 没保留的页眉页脚块与零散墨迹的个数，换成页码文字的页数
     */
    scan?: { textPages: number[]; formPages: number[]; blocks: number; inline: number; formWords: number; furniture: number; stray: number; pageNumbers: number; header: number; blockPages?: number[]; textBlocks?: number };
  };
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
  /**
   * 前置部分（流式模式）：这几页的页眉页脚、罗马数字页码的格式与起始值（前置部分第一页印的数，目录第一页印 i 时是 1）；
   * bodyStart：正文第一页的页码
   */
  front?: { headers: HeaderSet; footers: HeaderSet; fmt: "lowerRoman" | "upperRoman"; start: number; bodyStart: number };
  /** 需要告诉用户的版面取舍（如写不进 Word 页眉页脚的内容） */
  notes?: string[];
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
/**
 * 合同、法规里的条款编号：1.1 / 1.1.1 / 第一条 / （一） / (12) / (aa) / Article 5 / Section 2.1。
 * 律所文件的多级编号几乎都是"编号 + 制表符 + 悬挂缩进"，续行从编号后的正文位置起排
 */
const CLAUSE_LABEL = /^(\d{1,3}(\.\d{1,3}){1,4}\.?|第[一二三四五六七八九十百零〇\d]+[条章节款项]|[（(]\s*(\d{1,3}|[a-zA-Z]{1,2}|[ivxlcIVXLC]{1,6}|[一二三四五六七八九十]{1,3})\s*[)）]|[一二三四五六七八九十]{1,3}[、.．]|(Article|ARTICLE|Section|SECTION|Clause|CLAUSE)\s+\d{1,3}(\.\d{1,3})*\.?)$/;
/**
 * 行首是独立成段的编号 / 列表符号（单独一个文字片段）。minGap：编号与正文之间至少隔多宽（pt）。
 * 编号比悬挂缩进略窄时（"10.1"、"(iii)"），制表符只留下两三磅的间距，与普通空格差不多宽
 */
function labelLead(l: Line, minGap = 0): boolean {
  if (l.spans.length < 2) return false;
  const t = l.spans[0].text.trim(), gap = l.spans[1].x0 - l.spans[0].x1;
  return (BULLET.test(t) || CLAUSE_LABEL.test(t)) && gap > minGap;
}
/** 行首是编号、后面隔着制表间距（不是普通的词间空格）：新的条款 / 列表项 */
const labelHead = (l: Line): boolean => labelLead(l, 0.3 * l.size);
const BULLET = /^(\[ ?\d(?: ?\d){0,3} ?\]|[•·▪■□◆◇○●◦‣⁃∙➢➤►▸☐☑☒\-–—*]|\(?\d{1,3}[.)、]|\(?[a-zA-Z][.)]|\([ivxIVX]+\)|[ivxIVX]+[.)]|[一二三四五六七八九十]+[、.．])$/;
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
      // 重音符号（ȳ 的横线）与它下面的字母横向重叠，但属于同一行：不算冲突，成行后并成组合字符
      const accentLike = (x: Span) => ACCENT_ONLY.test(x.text) && !x.img;
      // 单独一个标点的片段与别的字横向重叠不算冲突：扫描件 OCR 的标点字框偏大，"（" 压在后面的 "S" 上、"：" 压在
      // 下一个词 "frame" 的开头上。算作冲突的话同一行被拆成两行，在 Word 里只能一上一下，后面的内容整块下移。
      // 同一个标点叠在自己的副本上（重复绘制的文字层）仍算冲突
      const lonePunct = (x: Span) => !x.img && LONE_PUNCT.test(x.text.trim());
      const punctOverlap = (o: Span) => (lonePunct(o) || lonePunct(s)) && o.text.trim() !== s.text.trim();
      // 识别文字层的字框偏大：后一个片段的开头压在前一个片段的末尾（"護が可肯" 的"肯"与后面的"巨"重叠 3pt），后一个片段
      // 伸出前一个的右端、重叠不到它自身宽度的一半、字号相同：是同一行接着写，不是叠印。算作冲突的话这一行从"巨"起另成一行，
      // 在流式排版里另起一段（2026-11-27 WO2017038541A1 首页 (84)）
      const tailOverlap = (o: Span) => {
        const ov = Math.min(o.x1, s.x1) - Math.max(o.x0, s.x0);
        return Math.abs(o.size - s.size) < 0.6 && s.x0 >= o.x0 + 0.6 * (o.x1 - o.x0) && s.x1 > o.x1 + 0.3 * (s.x1 - s.x0) && ov <= 0.45 * (s.x1 - s.x0);
      };
      const clash = g.spans.some((o) => Math.min(o.x1, s.x1) - Math.max(o.x0, s.x0) > 0.3 * Math.min(o.x1 - o.x0, s.x1 - s.x0) && !punctOverlap(o) && !tailOverlap(o) && !(accentLike(o) !== accentLike(s) && Math.abs(o.baseline - s.baseline) <= 0.45 * Math.max(o.size, s.size)));
      if (clash) continue;
      // 上下标优先并入左右紧挨着它的那一行：另一栏恰好同基线的行不能把它抢走（左栏的 M_θ* 并进右栏那一行）
      const small = s.size < 0.85 * gSize;
      const adjacent = !small || g.spans.some((o) => Math.max(o.x0 - s.x1, s.x0 - o.x1) < 1.5 * gSize);
      // 不挨着、基线也不同的小字不是这一行的上下标：扫描件首页的大号公开号（14 pt）纵跨左边上下两行小字
      // （"(43) 国際公開日" / "2017 年 3 月 9 日"），按上下标并进来会把这两行搅成一行
      if (small && !adjacent && d0 > 0.2 * s.size) continue;
      const d = Math.abs(g.baseline - s.baseline) + (adjacent ? 0 : 0.6 * gSize);
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
  for (const g of groups) g.spans = foldAccents(g.spans);
  attachMarks(groups);
  return groups.map((g) => finishLine(g.spans)).sort((a, b) => a.baseline - b.baseline || a.x0 - b.x0);
}

// 单独一个标点符号（Unicode 标点类，不含重音、数学符号）
const LONE_PUNCT = /^\p{P}$/u;
// 重音符号：独立的修饰字符，或组合用重音（BLEX 等大型定界符字体里的 ̂ ̃，U+0300–U+036F）
const ACCENT_ONLY = /^[ˆ^˜~¯˙¨´`ˇ→⃗\u0300-\u036f\s]+$/;
/** 根号（BLEX 的 √，字号与正文相同、垂直位置略高）：与重音一样要并回它所在的正文行 */
const RADICAL = /^[√∛∜]$/;

/** 同一行里叠在字母上方的重音符号（单独的文字项）并入它下面的字母：ȳ、F̂ */
function foldAccents(spans: Span[]): Span[] {
  if (spans.length < 2) return spans;
  const accents = spans.filter((x) => ACCENT_ONLY.test(x.text) && !x.img && x.text.trim());
  if (!accents.length || accents.length === spans.length) return spans;
  const out = [...spans];
  for (const ac of accents) {
    const cx = (ac.x0 + ac.x1) / 2;
    const k = out.findIndex((o) => o !== ac && !ACCENT_ONLY.test(o.text) && !o.img && o.text.trim() && o.x0 - 0.5 <= cx && o.x1 + 0.5 >= cx);
    const first = ac.text.trim()[0];
    const ch = /[\u0300-\u036f]/.test(first) ? first : COMBINING[first];
    if (k < 0 || !ch) continue;
    const o = out[k], n = o.text.length;
    const idx = Math.min(n - 1, Math.max(0, Math.floor(((cx - o.x0) / Math.max(0.1, o.x1 - o.x0)) * n)));
    out[k] = { ...o, text: o.text.slice(0, idx + 1) + ch + o.text.slice(idx + 1) };
    out.splice(out.indexOf(ac), 1);
  }
  return out;
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
  const isAccent = (t: string) => ACCENT_ONLY.test(t);
  const isRadical = (t: string) => RADICAL.test(t.trim());
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
      const allAccent = g.spans.every((x) => isAccent(x.text) || isRadical(x.text));
      if (!allAccent && gs > 0.85 * hs) continue;
      const hx0 = Math.min(...h.spans.map((x) => x.x0)), hx1 = Math.max(...h.spans.map((x) => x.x1));
      const gx0 = Math.min(...g.spans.map((x) => x.x0)), gx1 = Math.max(...g.spans.map((x) => x.x1));
      // 行末的上标（i 的 th）可能略超出正文行
      if (gx0 < hx0 - 2 || gx1 > hx1 + 2.5 * hs) continue;
      let d = Math.abs(h.baseline - g.baseline);
      // 根号的基线比所在行高，可能离上一行更近：与该行的文字横向重叠（压在字上）、而另一行紧接着就是被开方的式子时，
      // 属于后者（"squared error √‖x − x̂‖²" 的根号被并进上一行，上一行中间出现大段空白、根号与式子分离）
      let radicand = false;
      if (allAccent && g.spans.some((x) => isRadical(x.text))) {
        const clash = g.spans.some((x) => h.spans.some((o) => Math.min(o.x1, x.x1) - Math.max(o.x0, x.x0) > 0.3 * Math.min(o.x1 - o.x0, x.x1 - x.x0)));
        radicand = g.spans.some((x) => isRadical(x.text) && h.spans.some((o) => o.x0 >= x.x1 - 1.5 && o.x0 <= x.x1 + 2 && o.bottom > x.bottom));
        if (clash && !radicand) d += hs;
      }
      if (d > (radicand ? 1.0 : 0.75) * hs) continue;
      if (d < best) { best = d; host = h; }
    }
    if (!host) continue;
    for (const sp of g.spans) {
      if (isAccent(sp.text)) {
        const first0 = sp.text.trim()[0];
        const ch = /[\u0300-\u036f]/.test(first0) ? first0 : COMBINING[first0];
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
  // 同一栏里上下相邻的两张三线表，横线的左右端点几乎相同，会被并成一组；
  // 两条横线之间若夹着成行的正文或表题（"Table 2 …"），说明这是两张表的分界，从这里切开
  const split: HSeg[][] = [];
  for (const g of groups) {
    let cur: HSeg[] = [g[0]];
    for (let k = 1; k < g.length; k++) {
      const a = g[k - 1], b = g[k];
      const gx0 = Math.min(a.x0, b.x0), gx1 = Math.max(a.x1, b.x1);
      const between = page.spans.filter((sp) => !taken.has(sp) && sp.text.trim() && (sp.top + sp.bottom) / 2 > a.y && (sp.top + sp.bottom) / 2 < b.y && sp.x1 > gx0 - 2 && sp.x0 < gx1 + 2);
      const boundary = buildLines(between).some((l) => {
        const t = l.spans.map((sp) => sp.text).join(" ").trim();
        if (/^(table|tab\.)\s*[\dIVXLC]+\b|^表\s*\d/i.test(t)) return true;
        let maxGap = 0;
        for (let i = 1; i < l.spans.length; i++) maxGap = Math.max(maxGap, l.spans[i].x0 - l.spans[i - 1].x1);
        return l.x1 - l.x0 > 0.6 * (gx1 - gx0) && maxGap < Math.max(4, 0.6 * l.size) && l.spans.length >= 6;
      });
      if (boundary) { split.push(cur); cur = [b]; } else cur.push(b);
    }
    split.push(cur);
  }
  const out: TableDraft[] = [];
  for (const g of split) {
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

interface ScanPlan {
  /** 线框铺满页面、文字大半在框里的表单页：格线留作垫底图片，文字按原位置放进文本框 */
  form: boolean;
  /** 从位图层裁下来、要保留的块（含每页都一样、写进页眉页脚的块） */
  images: ImageBox[];
  /** 没保留的页眉页脚块 */
  furniture: number;
  /** 每页都不同的窄条：印成图的页码，换成这一页的页码文字（由页眉页脚的处理写成页码域） */
  pageNumber?: { box: Box; text: string };
  /** images 里每页都一样、留给页眉页脚的块 */
  header: Set<ImageBox>;
  /**
   * images 里页顶窄条（公开号、条码）的块：只有一两页有、不是每页重复的页眉，留在这一页。落在正文区域里时（其他页的正文排得
   * 比它高），这一页从新的一页开始：可编辑排版不保留原文的分页，上一页排满（没有提前换页）时这一页的内容接着排，页顶的窄条
   * 就跑到了上一页正文的末尾（WO2017038541A1 节选第 2 页的"WO 2017/038541 A1"和条码排在了第 1 页的底部，2026-11-27）
   */
  topStrip?: Set<ImageBox>;
}

/**
 * 可编辑排版下，哪些扫描页把位图层拆开处理（见 scanLayer.ts）：整页的黑白位图上面画着文字的页。
 * 没有文字的页（附图）不在其中，仍按整页图片保留。
 *   - 表单页（一个线框占了大半页、文字大半在框里：国际检索报告）：框线和字合起来才看得懂，重排文字会把字从格子里挪走。
 *     能渲染时整页成图；不能渲染时不动（整页位图垫在文字下面）。
 *   - 其余的页：整页位图不要，裁下来的块当图片。每页顶上、底下重复出现的块（印成图的公开号、页码）是页眉页脚，不进正文。
 */
function planScanPages(pages: PageModel[]): Map<PageModel, ScanPlan> {
  const plans = new Map<PageModel, ScanPlan>();
  const cand = pages.filter((pg) => pg.scan && pg.images.includes(pg.scan.image) && pg.spans.some((sp) => sp.text.trim() && !sp.img));
  // 页眉页脚位置上的块：整块在第一行文字之上（且在页面顶部 10% 内），或在最后一行文字之下（且在底部 8% 内）
  const strips = new Map<PageModel, Set<ScanPiece>>();
  for (const pg of cand) {
    const real = pg.spans.filter((sp) => sp.text.trim());
    const top = Math.min(...real.map((sp) => sp.top)), bottom = Math.max(...real.map((sp) => sp.bottom));
    const set = new Set(pg.scan!.pieces.filter((p) => p.y1 <= Math.min(0.1 * pg.height, top - 1) || p.y0 >= Math.max(0.92 * pg.height, bottom + 1)));
    if (set.size) strips.set(pg, set);
  }
  // 三页以上都有：是每页重复的页眉页脚。只有一两页有的（首页顶上的徽标）留在正文里。
  // 按位置分组：每页一模一样的（公开号、申请号）保留，写进 Word 页眉页脚；每页都不一样的窄条是页码
  const keep = new Set<ScanPiece>();
  const numbers = new Map<ScanPiece, string>();
  if (strips.size >= 3) {
    const groups = new Map<string, Array<{ pg: PageModel; p: ScanPiece }>>();
    for (const [pg, set] of strips) for (const p of set) {
      const key = `${(p.y0 + p.y1) / 2 < pg.height / 2 ? "t" : "b"}${Math.round((p.x0 + p.x1) / 2 / 20)}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key)!.push({ pg, p });
    }
    for (const list of groups.values()) {
      if (list.length < 3) continue;
      const counts = new Map<string, number>();
      const sig = (p: ScanPiece) => p.png.toString("base64");
      for (const { p } of list) counts.set(sig(p), (counts.get(sig(p)) ?? 0) + 1);
      const [topSig, topCount] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
      if (topCount >= Math.max(3, 0.6 * list.length)) {
        for (const { p } of list) if (sig(p) === topSig) keep.add(p);
        continue;
      }
      if (counts.size === list.length) for (const [p, text] of pageNumberTexts(list)) numbers.set(p, text);
    }
  }
  for (const pg of cand) {
    const sc = pg.scan!;
    const real = pg.spans.filter((sp) => sp.text.trim());
    const chars = (list: Span[]) => list.reduce((n, sp) => n + sp.text.replace(/\s/g, "").length, 0);
    const within = (p: Box) => real.filter((sp) => { const cx = (sp.x0 + sp.x1) / 2, cy = (sp.top + sp.bottom) / 2; return cx >= p.x0 && cx <= p.x1 && cy >= p.y0 && cy <= p.y1; });
    const form = sc.pieces.some((p) => p.kind === "graphic" && p.x1 - p.x0 >= 0.6 * pg.width && p.y1 - p.y0 >= 0.45 * pg.height && chars(within(p)) >= 0.5 * chars(real));
    if (form) {
      plans.set(pg, { form: true, images: [], furniture: 0, header: new Set() });
      continue;
    }
    const drop = strips.size >= 3 ? strips.get(pg) : undefined;
    const header = new Set<ImageBox>(), top = new Set<ImageBox>();
    const images: ImageBox[] = sc.pieces.filter((p) => !drop?.has(p) || keep.has(p)).map((p) => {
      const im: ImageBox = { x0: p.x0, y0: p.y0, x1: p.x1, y1: p.y1, png: p.png, pxW: p.pxW, pxH: p.pxH, scan: p.kind, ...(p.line ? { line: p.line } : {}) };
      if (keep.has(p)) header.add(im);
      // 页顶的窄条留在了这一页（不足三页有，不进页眉）
      else if (!drop && strips.get(pg)?.has(p) && (p.y0 + p.y1) / 2 < pg.height / 2) top.add(im);
      return im;
    });
    const num = [...(drop ?? [])].find((p) => numbers.has(p));
    plans.set(pg, {
      form: false, images, header, furniture: [...(drop ?? [])].filter((p) => !keep.has(p) && !numbers.has(p)).length,
      ...(num ? { pageNumber: { box: { x0: num.x0, y0: num.y0, x1: num.x1, y1: num.y1 }, text: numbers.get(num)! } } : {}),
      ...(top.size ? { topStrip: top } : {}),
    });
  }
  return plans;
}

/**
 * 每页都不同的一组窄条是不是页码、各是几：页码逐页加一，图里的数字个数（连通块数）只在 9→10、99→100 处变。
 * 由位数变化的位置推出起始页码，所有页的位数都对得上才算；推不出来（不是页码、或页码不连续）返回空
 */
function pageNumberTexts(list: Array<{ pg: PageModel; p: ScanPiece }>): Map<ScanPiece, string> {
  const out = new Map<ScanPiece, string>();
  const sorted = [...list].sort((a, b) => a.pg.index - b.pg.index);
  if (sorted.some(({ p, pg }) => p.x1 - p.x0 > 0.12 * pg.width || p.glyphs < 1 || p.glyphs > 4)) return out;
  // 位数：连通块数；两个数字粘在一起时只算一块，再按宽度核对（一位数的宽度取单块中较窄的一半的中位数）
  const widths = sorted.filter(({ p }) => p.glyphs === 1).map(({ p }) => p.x1 - p.x0).sort((a, b) => a - b);
  const one = widths.length ? widths[Math.floor((widths.length - 1) / 4)] : 0;
  const digits = sorted.map(({ p }) => Math.max(p.glyphs, one ? (p.x1 - p.x0 > 2.4 * one ? 3 : p.x1 - p.x0 > 1.45 * one ? 2 : 1) : 1));
  // 候选起点：第一页的页码 n0，使得每一页 n0 + (页序差) 的位数与图里的数字个数一致
  for (let n0 = 0; n0 <= 9999; n0++) {
    const ok = sorted.every(({ pg }, k) => String(n0 + pg.index - sorted[0].pg.index).length === digits[k]);
    if (!ok) continue;
    // 起点不唯一（全是一位数、没有位数变化）时推不出来
    const next = sorted.every(({ pg }, k) => String(n0 + 1 + pg.index - sorted[0].pg.index).length === digits[k]);
    if (next) return out;
    sorted.forEach(({ pg, p }) => out.set(p, String(n0 + pg.index - sorted[0].pg.index)));
    return out;
  }
  return out;
}

/**
 * 整页大小的图片（扫描件的每一页、整页的底图）：一张或几张拼在一起（扫描仪常把一页切成几条）覆盖页面宽、高的 85% 以上。
 * 它们是页面背景，按页面坐标绝对定位、不进入文字流。当成独占一段的行内图片时，图比版心高：
 * 逐页保留模式下页首的分页锚点段落之后放不下，整图被挤到下一页（4 页的扫描信变成 8 页、空白页与图片页交替）；
 * 可编辑排版下图片从上边距处开始排，下半截被页面裁掉
 */
function pageBackgrounds(pg: PageModel): ImageBox[] {
  const W = pg.width, H = pg.height;
  const clip = (im: ImageBox) => ({ x0: Math.max(0, im.x0), y0: Math.max(0, im.y0), x1: Math.min(W, im.x1), y1: Math.min(H, im.y1) });
  const ims = pg.images.filter((im) => { const c = clip(im); return !im.render && (c.x1 - c.x0) * (c.y1 - c.y0) > 0.02 * W * H; });
  if (!ims.length) return [];
  const cs = ims.map(clip);
  const u = { x0: Math.min(...cs.map((c) => c.x0)), y0: Math.min(...cs.map((c) => c.y0)), x1: Math.max(...cs.map((c) => c.x1)), y1: Math.max(...cs.map((c) => c.y1)) };
  const covered = cs.reduce((a, c) => a + (c.x1 - c.x0) * (c.y1 - c.y0), 0);
  return u.x1 - u.x0 >= 0.85 * W && u.y1 - u.y0 >= 0.85 * H && covered >= 0.7 * W * H ? ims : [];
}

/**
 * 夹在文字行中的小图标（作者名后的 ORCID 图标、行内的小符号图片）：改成该行里的行内图片。
 * 否则它们被当成独立的图片，各占一段，把下面的行整体推下去，图标本身也和名字脱开
 */
function inlineIcons(pg: PageModel): number {
  const keep: ImageBox[] = [];
  for (const im of pg.images) {
    const h = im.y1 - im.y0, w = im.x1 - im.x0;
    const row = pg.spans.filter((sp) => sp.text.trim() && !sp.img && Math.abs(sp.baseline - im.y1) <= 0.35 * sp.size + 1 && h <= 1.6 * sp.size && w <= 3 * sp.size);
    const near = row.filter((sp) => Math.min(Math.abs(sp.x1 - im.x0), Math.abs(sp.x0 - im.x1)) < 1.5 * sp.size);
    const over = row.some((sp) => Math.min(sp.x1, im.x1) - Math.max(sp.x0, im.x0) > 0.5 * Math.min(w, sp.x1 - sp.x0));
    if (im.behind || !near.length || over) { keep.push(im); continue; }
    const ref = near.reduce((a, b) => (b.size > a.size ? b : a));
    const size = Math.max(...row.map((sp) => sp.size));
    pg.spans.push({
      text: "\uFFFC", x0: im.x0, x1: im.x1, top: ref.baseline - 0.75 * size, bottom: ref.baseline + 0.22 * size, baseline: ref.baseline, size,
      font: ref.font, bold: false, italic: false, color: "000000", natural: w, wordW: w, spaceW: ref.spaceW, wordSpaceW: ref.wordSpaceW, img: im,
    });
  }
  const n = pg.images.length - keep.length;
  pg.images = keep;
  return n;
}

/**
 * 左右并排的两个文字块（Elsevier 等期刊首页的"文章信息 / 关键词"与"摘要"）：
 * 一段连续的行里有一条所有行都不跨越的竖向空白，且这条空白不在页面中部（不是正文的双栏），
 * 两侧各有至少 3 行实在的文字 → 写成一张无边框的两格表格，两侧各自成段。
 * 否则两侧的行会按纵坐标交错排在一起：左侧的短行把右侧段落切碎，行距也被打乱。
 * 两侧标题下各自的横线（"ARTICLE INFO" / "ABSTRACT" 下的线）作为表头行的下边框。
 */
function detectSideBySide(page: PageModel, used: Set<Seg>, taken: Set<Span>): TableDraft[] {
  const free = page.spans.filter((sp) => !taken.has(sp) && sp.text.trim() && !sp.img);
  if (free.length < 12) return [];
  const lines = buildLines(free).sort((a, b) => a.top - b.top);
  const X0 = Math.min(...lines.map((l) => l.x0)), X1 = Math.max(...lines.map((l) => l.x1)), W = X1 - X0;
  if (W < 200) return [];
  const lo = X0 + 0.12 * W, hi = X0 + 0.88 * W;
  type Iv = [number, number];
  const cut = (ivs: Iv[], a: number, b: number): Iv[] => ivs.flatMap(([p, q]): Iv[] => (b <= p || a >= q ? [[p, q]] : [...(a > p ? [[p, a] as Iv] : []), ...(b < q ? [[b, q] as Iv] : [])]));
  const offCenter = (iv: Iv) => { const c = (iv[0] + iv[1]) / 2; return c < X0 + 0.4 * W || c > X0 + 0.6 * W; };
  // 本页正文的双栏：分界落在栏边界上（左栏几行短句 + 右栏正文）只是普通的分栏，不是并排的文字块
  const pageGutter = detectGutter(lines, { x0: X0, x1: X1 });
  const out: TableDraft[] = [];
  let i = 0;
  while (i < lines.length) {
    let corridor: Iv[] = [[lo, hi]];
    for (const sp of lines[i].spans) corridor = cut(corridor, sp.x0 - 3, sp.x1 + 3);
    corridor = corridor.filter((iv) => iv[1] - iv[0] >= 8 && offCenter(iv));
    let j = i + 1;
    while (corridor.length && j < lines.length && lines[j].top - lines[j - 1].bottom < 2.5 * Math.max(lines[j].size, lines[j - 1].size)) {
      let next = corridor;
      for (const sp of lines[j].spans) next = cut(next, sp.x0 - 3, sp.x1 + 3);
      next = next.filter((iv) => iv[1] - iv[0] >= 8 && offCenter(iv));
      if (!next.length) break;
      corridor = next;
      j++;
    }
    const run = lines.slice(i, j);
    // 取最宽的空白作为分界
    const gap = corridor.sort((a, b) => (b[1] - b[0]) - (a[1] - a[0]))[0];
    if (gap && run.length >= 6) {
      const mid = (gap[0] + gap[1]) / 2;
      // 分界取两侧文字之间空白的正中（空白带受扫描范围限制，不一定居中）
      const lEdge = Math.max(...run.flatMap((l) => l.spans.filter((sp) => sp.x1 <= mid).map((sp) => sp.x1)), -Infinity);
      const rEdge = Math.min(...run.flatMap((l) => l.spans.filter((sp) => sp.x0 >= mid).map((sp) => sp.x0)), Infinity);
      const gx = Number.isFinite(lEdge) && Number.isFinite(rEdge) ? (lEdge + rEdge) / 2 : mid;
      const side = (l: Line, left: boolean) => l.spans.filter((sp) => (left ? sp.x1 <= gx : sp.x0 >= gx));
      const chars = (left: boolean) => run.reduce((n, l) => n + side(l, left).reduce((m, sp) => m + sp.text.trim().length, 0), 0);
      const nL = run.filter((l) => side(l, true).length).length, nR = run.filter((l) => side(l, false).length).length;
      const onColumns = !!pageGutter && (Math.abs(rEdge - pageGutter.g1) <= 4 || Math.abs(lEdge - pageGutter.g0) <= 4);
      // 整块都落在本页双栏的某一栏里（没有一行跨过栏间距）：这条空白是栏内的（著录项目"标签 · 值"之间的空隙），
      // 不是两个并排的文字块。拆成两格表格会把左右两栏的文字搅在一起
      const withinColumns = !!pageGutter && run.every((l) => classify(l, pageGutter.x) !== "cross");
      if (nL >= 3 && nR >= 3 && chars(true) >= 40 && chars(false) >= 40 && !onColumns && !withinColumns) {
        const spans = run.flatMap((l) => l.spans);
        const x0 = Math.min(...spans.map((sp) => sp.x0)) - CELL_MARGIN, x1 = Math.max(...spans.map((sp) => sp.x1)) + CELL_MARGIN;
        const y0 = Math.min(...spans.map((sp) => sp.top)) - 1, y1 = Math.max(...spans.map((sp) => sp.bottom)) + 1;
        // 两侧各自的标题横线：落在块内、只在一侧、两侧同一高度 → 表头行的下边框
        const isL = (h: HSeg) => (h.x0 + h.x1) / 2 < gx && h.x1 < rEdge - 1, isR = (h: HSeg) => (h.x0 + h.x1) / 2 > gx && h.x0 > lEdge + 1;
        const rules = mergeH(page.segs.filter((sg) => !used.has(sg))).filter((h) => h.y > y0 && h.y < y1 && h.x1 - h.x0 > 30 && (isL(h) || isR(h)));
        const ys = [y0];
        const hs: HSeg[] = [];
        // 必要条件：两侧标题下各有一条同高的横线（"ARTICLE INFO / ABSTRACT"）。没有它，文字间的竖向空白多半只是
        // 公式内部的空隙、居中的图题、两栏正文本身——把这些并排的行拆成两格表格会打乱整页版式
        const ry = rules.length >= 1 ? rules[0].y : null;
        const bothRules = rules.length >= 2 && rules.some(isL) && rules.some(isR) && rules.every((h) => Math.abs(h.y - rules[0].y) <= 1.5);
        // 首页上半部的前置信息块（≥ 8 行，各自成段）：标题下的横线可只有一侧有（左侧的 ARTICLE INFO 有、右侧的 ABSTRACT 没有）
        const frontMatter = page.index === 0 && y0 < 0.5 * page.height && run.length >= 8 && nL >= 4 && nR >= 4;
        if (!(bothRules || frontMatter) || (ry !== null && spans.some((sp) => sp.top < ry && sp.bottom > ry))) { i = j; continue; }
        if (ry !== null && (bothRules || rules.every((h) => Math.abs(h.y - ry) <= 1.5))) {
          ys.push(ry);
          for (const h of rules) {
            hs.push({ ...h, y: ry, x0: isL(h) ? x0 : gx, x1: isL(h) ? gx : x1 });
            h.src.forEach((sg) => used.add(sg));
          }
        }
        ys.push(y1);
        const xs = [x0, gx, x1];
        out.push({ x0, x1, y0, y1, xs, ys, hs, vs: xs.map((x) => ({ x, y0, y1, w: 0, color: "none", src: [] })), spans });
        for (const sp of spans) taken.add(sp);
      }
    }
    i = j;
  }
  return out;
}

/**
 * 两栏释义表（招股书、合同的"释义 / DEFINITIONS"：左栏词条，右栏释义，每条词条与它的释义第一行同高）：
 * 写成一张无边框的两列表格，一条一行。不能当成正文的双栏：分栏后词条连续排在左栏、释义连续排在右栏，
 * 读者对不上哪条释义属于哪个词条，相邻的释义还会并成一段（2026-11-27 LTB-07：48 条释义与词条错位）。
 * 条件：一段连续的行都只从两个固定位置起排（左栏起点 / 右栏起点），没有行跨过两栏之间的空白；左栏比右栏窄；
 * 左栏每段（词条）的第一行，都有右栏一段（释义）的第一行与它同高，这样的条目至少 4 条。
 * 正文的双栏不满足"段首同高"：两栏的段落各排各的。
 * 写成表格的只有两种清单，其余照旧按制表位写成段落：
 *   - 会被当成双栏切开的：本页测出的栏间距，或全文共识的栏间距（docGutterX，本页自己测不出时分栏按它切，
 *     LTB-07 第 7 页），正好落在词条与释义之间；
 *   - 词条栏宽（释义从版心宽度 25% 以外起排）、有条目折行、又不是条款编号的：不分栏时每个折行各成一段，
 *     两行的词条与释义的第二行拼进同一段，一句话被切成几段分别翻译。
 *   释义起点靠左的邮件抬头（From / To / Subject，LTB-14）、条款编号（悬挂缩进另有处理）不改成表格；
 *   文字压在整页底图上的页（扫描的表单，scan-text.pdf 第 6 页）也不改：精确排版按原位置把每行放在底图的格子上，
 *   换成表格会和底图错开。
 * 同一份文件里词条、释义起点相同的清单（listCols，跨页的释义）一并写成表格：前一页的释义改成表格后，
 * 全文共识的栏间距随之消失，后一页就不再被切开，同一份释义一页是表格、一页是段落。
 */
type DefinitionCols = { A: number; B: number };
function detectDefinitionList(page: PageModel, taken: Set<Span>, docGutterX?: number, listCols: DefinitionCols[] = [], found?: DefinitionCols[]): TableDraft[] {
  const free = page.spans.filter((sp) => !taken.has(sp) && sp.text.trim() && !sp.img);
  if (free.length < 12) return [];
  const lines = buildLines(free).sort((a, b) => a.top - b.top);
  const X0 = Math.min(...lines.map((l) => l.x0)), X1 = Math.max(...lines.map((l) => l.x1)), W = X1 - X0;
  if (W < 200) return [];
  const pageGutter = detectGutter(lines, { x0: X0, x1: X1 });
  const cuts = [pageGutter?.x, docGutterX].filter((x): x is number => x !== undefined);
  // 一行里按明显的空白切成几块（左栏词条、右栏释义各一块）。
  // 招股书封面的"项目 : 内容"列表里，单独的冒号紧贴右栏：块的起点按冒号之后的文字算，冒号另记
  type Piece = { spans: Span[]; x0: number; x1: number; top: number; bottom: number; colon?: Span };
  const isColon = (sp: Span) => /^[:：]$/.test(sp.text.trim());
  const pieces = (l: Line): Piece[] => {
    const out: Piece[] = [];
    for (const sp of [...l.spans].sort((a, b) => a.x0 - b.x0)) {
      const last = out[out.length - 1];
      if (last && sp.x0 - last.x1 < Math.max(8, 1.5 * sp.size)) {
        last.spans.push(sp); last.x1 = Math.max(last.x1, sp.x1); last.top = Math.min(last.top, sp.top); last.bottom = Math.max(last.bottom, sp.bottom);
        if (last.colon && last.spans.length === 2) last.x0 = sp.x0;
      } else out.push({ spans: [sp], x0: sp.x0, x1: sp.x1, top: sp.top, bottom: sp.bottom, ...(isColon(sp) ? { colon: sp } : {}) });
    }
    // 只有一个冒号、后面没有文字的块不算冒号块
    for (const p of out) if (p.colon && p.spans.length === 1) delete p.colon;
    return out;
  };
  const split = lines.map((l) => ({ l, ps: pieces(l) }));
  // 右栏起点：版心 15%–60% 处最常见的块起点。没有的话再看 60%–92%：定义术语索引（"Term … Location"）的右栏是一列
  // 短的出处（"Section 7.1(d)"），从版心 80% 处起排（2026-11-27 APA：索引排满三页，被当成正文双栏，术语与条款号对不上）
  const rankStarts = (lo: number, hi: number) => {
    const starts = new Map<number, number>();
    for (const { ps } of split) for (const p of ps) if (p.x0 > X0 + lo * W && p.x0 < X0 + hi * W) starts.set(Math.round(p.x0), (starts.get(Math.round(p.x0)) ?? 0) + 1);
    return [...starts.entries()].sort((a, b) => b[1] - a[1]);
  };
  let ranked = rankStarts(0.15, 0.6);
  const index = !ranked.length || ranked[0][1] < 6;
  if (index) ranked = rankStarts(0.6, 0.92);
  if (!ranked.length || ranked[0][1] < 6) return [];
  const B = ranked[0][0];
  const atB = (x: number) => Math.abs(x - B) <= 2;
  // 左栏起点：右栏起点左侧最常见的块起点
  const lstarts = new Map<number, number>();
  for (const { ps } of split) for (const p of ps) if (p.x1 < B - 6) lstarts.set(Math.round(p.x0), (lstarts.get(Math.round(p.x0)) ?? 0) + 1);
  const lranked = [...lstarts.entries()].sort((a, b) => b[1] - a[1]);
  if (!lranked.length) return [];
  const A = lranked[0][0];
  const atA = (x: number) => Math.abs(x - A) <= 3;
  // 每行：左块（起于 A、止于右栏之前）、右块（起于 B），二者至少有一个，别无他块
  type Row = { l: Line; L?: Piece; R?: Piece };
  const rowOf = ({ l, ps }: { l: Line; ps: Piece[] }): Row | null => {
    if (!ps.length || ps.length > 2) return null;
    const L = ps.find((p) => atA(p.x0) && p.x1 < B - 6), R = ps.find((p) => atB(p.x0));
    if ((L ? 1 : 0) + (R ? 1 : 0) !== ps.length) return null;
    return { l, L, R };
  };
  const out: TableDraft[] = [];
  let i = 0;
  while (i < split.length) {
    const run: Row[] = [];
    let j = i;
    for (; j < split.length; j++) {
      const r = rowOf(split[j]);
      if (!r) break;
      // 行距过大（另起一块内容）也结束
      if (run.length && r.l.top - run[run.length - 1].l.bottom > 3 * r.l.size) break;
      run.push(r);
    }
    i = Math.max(j, i + 1);
    if (run.length < 8) continue;
    const left = run.filter((r) => r.L).map((r) => r.L!), right = run.filter((r) => r.R).map((r) => r.R!);
    if (left.length < 4 || right.length < 4) continue;
    // 冒号列：有词条的行，右块都以同一位置的冒号开头
    const withL = run.filter((r) => r.L && r.R);
    const colons = withL.map((r) => r.R!.colon).filter((c): c is Span => !!c);
    const colonCol = colons.length >= 4 && colons.length >= 0.8 * withL.length && Math.max(...colons.map((c) => c.x0)) - Math.min(...colons.map((c) => c.x0)) <= 2;
    const leftW = Math.max(...left.map((p) => p.x1)) - A, rightW = Math.max(...right.map((p) => p.x1)) - B;
    const text = (p: Piece) => p.spans.map((sp) => sp.text).join("").trim();
    // 索引：右栏每一块都是与左栏某一行同高的短出处（不折行），这时右栏比左栏窄
    const isIndex = index && right.every((q) => text(q).length <= 40 && left.some((p) => Math.abs(p.top - q.top) <= 1.5));
    if (index && !isIndex) continue;
    if (!isIndex && leftW >= (colonCol ? 1 : 0.75) * rightW) continue;
    // 词条：左栏的段首。词条的续行正好在上一行下方一个行距处；新的词条更低（释义多行，或条目之间有段距），
    // 或者是紧接上一个词条的右引号另起的引号（条目之间没有段距时）。行距取这一段里相邻两行的最小间距。
    // 右栏不按行距分段：多行释义之后往往没有额外的段距（LTB-07）
    const steps = run.slice(1).map((r, k) => r.l.top - run[k].l.top).filter((d, k) => d > 0.5 * run[k + 1].l.size);
    const lineH = steps.length ? Math.min(...steps) : 0;
    const quoted = (p: Piece, prev: Piece) => /^[“"「『]/.test(text(p)) && /[”"」』]$/.test(text(prev));
    // 有冒号列时，带冒号的行就是一个条目的第一行（项目名称折行时续行没有冒号，行距也一样）
    // 索引：右栏有出处的每一行就是一条（一行一条，行距都相同）
    const ls = isIndex
      ? left.filter((p) => right.some((q) => Math.abs(q.top - p.top) <= 1.5))
      : colonCol
      ? withL.filter((r) => r.R!.colon).map((r) => r.L!)
      : left.filter((p, k) => k === 0 || p.top - left[k - 1].top > 1.12 * lineH + 0.5 || quoted(p, left[k - 1]));
    // 左栏大多是一行的词条（正文双栏的左栏大多是续行）；每个词条都有右栏的一行与它同高
    if (ls.length < 0.6 * left.length) continue;
    const aligned = ls.filter((p) => right.some((q) => Math.abs(q.top - p.top) <= 1.5));
    if (aligned.length < 4 || aligned.length < 0.8 * ls.length) continue;
    // 表格从第一条开始：之前只在一栏的行留在正文
    const first = aligned[0].top;
    const rows = run.filter((r) => r.l.top >= first - 1.5);
    const spans = rows.flatMap((r) => r.l.spans);
    const x0 = Math.min(...spans.map((sp) => sp.x0)) - CELL_MARGIN, x1 = Math.max(...spans.map((sp) => sp.x1)) + CELL_MARGIN;
    const cx0 = colonCol ? Math.min(...colons.map((c) => c.x0)) : B, cx1 = colonCol ? Math.max(...colons.map((c) => c.x1)) : B;
    const leftX1 = Math.max(...left.map((p) => p.x1));
    const sameCols = listCols.some((c) => Math.abs(c.A - A) <= 3 && Math.abs(c.B - B) <= 2);
    const cut = cuts.some((x) => x >= leftX1 - 1 && x <= cx0 + 1);
    const wraps = rows.length > aligned.length;
    const numbered = aligned.filter((p) => CLAUSE_LABEL.test(text(p))).length >= 0.5 * aligned.length;
    if (!sameCols && !cut && !(wraps && !numbered && cx0 > X0 + 0.25 * W)) continue;
    found?.push({ A, B });
    const gx = (leftX1 + cx0) / 2;
    const y0 = Math.min(...spans.map((sp) => sp.top)) - 1, y1 = Math.max(...spans.map((sp) => sp.bottom)) + 1;
    // 行界：下一条的第一行之上、上一条最后一行之下的正中
    const ys = [y0];
    for (const p of aligned.slice(1)) {
      const above = rows.filter((r) => r.l.top < p.top - 1.5);
      const lastBottom = Math.max(...above.map((r) => r.l.bottom));
      ys.push((lastBottom + Math.min(p.top, ...rows.filter((r) => Math.abs(r.l.top - p.top) <= 1.5).map((r) => r.l.top))) / 2);
    }
    ys.push(y1);
    const xs = colonCol ? [x0, gx, (cx1 + B) / 2, x1] : [x0, gx, x1];
    // 行界写成不可见的横线（无横线时上下格会被纵向合并成一格）
    const hs: HSeg[] = ys.map((y) => ({ y, x0, x1, w: 0, color: "none", src: [] }));
    out.push({ x0, x1, y0, y1, xs, ys, hs, vs: xs.map((x) => ({ x, y0, y1, w: 0, color: "none", src: [] })), spans });
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

/** 行尾是引导点和页码的行（目录的一条）："Section 1.1. Definitions……1""第一章 总则 · · · · · ii" */
const CONTENTS_END = /\s*(?:[.·…]\s?){4,}\s*(\d{1,4}|[ivxlcdm]{1,7})\s*$/i;
const contentsLine = (l: Line) => CONTENTS_END.test(l.spans.map((sp) => sp.text).join(""));

/**
 * 目录的一条（流式模式）：引导点是文字，译文长短一变，页码就不在行尾了。末尾的引导点和页码前的空白改成右对齐、
 * 带点状前导符的制表位（位置是原来页码的右端），页码跟在后面。返回去掉的引导点（计入完整性校验）；没有改时为 null
 */
function contentsTab(runs: Run[], tabs: TabStop[], pos: number): string | null {
  // 制表符算空白（页码前已有一个大间距写成的制表符时，它连同引导点一起换掉）
  const all = runs.map((x) => (x.tab ? "\t" : x.br || x.img ? "\u0001" : x.text)).join("");
  const m = all.match(CONTENTS_END);
  if (!m || m.index === undefined || m.index === 0) return null;
  let drop = all.length - m.index;
  let numFmt: Run | undefined;
  while (drop > 0 && runs.length) {
    const last = runs[runs.length - 1];
    const len = last.tab || last.br || last.img ? 1 : last.text.length;
    if (!numFmt && !last.tab && !last.br && !last.img && last.text.trim()) numFmt = last;
    if (len <= drop) { runs.pop(); drop -= len; }
    else { last.text = last.text.slice(0, last.text.length - drop); drop = 0; }
  }
  const tail = runs[runs.length - 1];
  if (!tail || !numFmt) return null;
  if (!tail.tab && !tail.br) tail.text = tail.text.replace(/\s+$/, "");
  runs.push({ ...plainFmt(numFmt), text: "", tab: true, underline: undefined, strike: undefined }, { ...plainFmt(numFmt), text: m[1], underline: undefined });
  const keep = tabs.filter((t) => t.pos < pos - 2);
  tabs.splice(0, tabs.length, ...keep, { pos, align: "right", leader: "dot" });
  return m[0].replace(/[^.·…]/g, "");
}

/** 行内有制表间距（并排的标签、签字栏）；skipLabel：行首编号后的制表间距不算（悬挂缩进的条款） */
function hasBigGap(l: Line, skipLabel = false): boolean {
  if (proseLines.has(l)) return false;
  const from = skipLabel && labelLead(l) ? 2 : 1;
  for (let i = from; i < l.spans.length; i++) if (l.spans[i].x0 - l.spans[i - 1].x1 > Math.max(1.2 * l.size, 12)) return true;
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
/**
 * 行的正文从哪里起：编号伸在左边距里的行（专利的"[0006]"、条款号），按编号后正文的起点算，
 * 才能和它下面从正文左边界起排的续行比较
 */
const textStart = (l: Line) => (labelLead(l) ? Math.min(l.spans[1].x0, l.x0 + 0.5 * (l.x1 - l.x0)) : l.x0);

function localEdges(lines: Line[], r: Region): number[] {
  return lines.map((l, i) => {
    const near = lines.slice(Math.max(0, i - 4), i + 5).filter((o) => Math.abs(o.x0 - l.x0) <= 20 && Math.abs(o.size - l.size) <= 0.15 * l.size);
    const edge = Math.max(...near.map((o) => o.x1));
    return near.filter((o) => edge - o.x1 <= 1.5).length >= 3 ? edge : r.x1;
  });
}

/** 行内的上标片段（比本行字号小、基线抬高）：脚注引用号 */
function superscripts(l: Line): string[] {
  return l.spans.filter((s) => s.size <= 0.85 * l.size && s.baseline < l.baseline - 0.12 * l.size).map((s) => s.text.trim()).filter((t) => /^(\d{1,3}|[*†‡§])$/.test(t));
}

/** 注释行开头的编号（"12 See clause…"、"¹²…"、"* …"）；没有则为 "" */
function noteLabel(l: Line): string {
  const m = /^\s*(\d{1,3}|[*†‡§])(?=[\s.)]|[A-Za-z\u4e00-\u9fff])/.exec(l.spans[0]?.text ?? "");
  return m ? m[1] : "";
}

/** 行的文字（片段之间按间距补空格） */
const lineString = (l: Line) => l.spans.map((s) => s.text).join(" ").replace(/\s+/g, " ").trim();

/**
 * 西文行停在句子中间：至少占到行宽（到 edge）的 60%，以字母、数字或逗号结尾，含小写字母（不是全大写的标题）。
 * 中文按字折行，行尾留白就是段落结束或手动换行，不在此列
 */
function openEnded(a: Line, edge: number): boolean {
  const ta = lineString(a);
  if (!ta || CJK.test(ta.slice(-1)) || a.x1 - a.x0 < 0.6 * (edge - a.x0)) return false;
  return /[A-Za-z0-9,]$/.test(ta) && /[a-z]/.test(ta);
}

/** 上一行 a 没写完一句话、下一行 b 接着写：b 以字母、数字、左括号或引号开头，a 的末尾与 b 的开头粗体一致 */
function continuesSentence(a: Line, b: Line, edge: number): boolean {
  const tb = lineString(b);
  if (!tb || CJK.test(tb[0]) || !/^[A-Za-z0-9("'\u2018\u201c]/.test(tb) || !openEnded(a, edge)) return false;
  return a.spans[a.spans.length - 1].bold === b.spans[0].bold;
}

/**
 * 判断自动折行用的右边界：左对齐（右边参差）的正文没有 3 行以上落在同一右边界，localEdges 退回区域右边界；
 * 区域常被页眉页脚、页码或更宽的表格撑得比实际折行处更靠右，句末恰好在行尾的那一行会被当成段落结束。
 * 附近有 ≥3 行同起点、同字号，最长的一行离区域右边界不远（≤3 个字号或版心宽的 8%）时，以最长那一行为准
 */
function wrapEdges(lines: Line[], r: Region, edges: number[]): number[] {
  return lines.map((l, i) => {
    if (edges[i] < r.x1) return edges[i];
    const near = lines.slice(Math.max(0, i - 4), i + 5).filter((o) => Math.abs(textStart(o) - textStart(l)) <= 20 && Math.abs(o.size - l.size) <= 0.15 * l.size);
    if (near.length < 3) return edges[i];
    const longest = Math.max(...near.map((o) => o.x1));
    const slack = r.x1 - longest;
    return slack > 1.5 && slack <= Math.max(3 * l.size, 0.08 * (r.x1 - r.x0)) ? longest : edges[i];
  });
}

/** 行里最大的字号（不算行内图片）：识别文字层里中日文字和西文混排时两种字号，行的字号按字数取，同一段的各行忽大忽小 */
const lineEm = (l: Line) => Math.max(l.size, ...l.spans.filter((s) => !s.img && s.text.trim()).map((s) => s.size));

/**
 * 被拉开的正文行（流式模式）。两端对齐的窄栏、扫描件的识别文字层（词宽按替代字体算，比扫描图上的字窄）里，
 * 行内常有一两处词距超过制表间距的门槛 max(1.2×字号, 12pt)：按制表位处理，这一行单独成段、在原位置设制表位，
 * 上下行也不并进来，译文变短后行中、行尾留下大片空白，一句话被拆到几段里分别翻
 * （2026-11-27 WO2017038541A1 首页：英文摘要 19 行里 7 行各有一处 12.1–13.6pt 的词距，只是同一行普通词距的
 * 1.1–1.9 倍，19 行成了 12 段；英文发明名称一行、(72)(74)(84) 著录项目同样）。下面这些行里不超过 3.5 个字宽的空隙
 * 按词距算，不设制表位、不断段：
 *   - 两端对齐的非末行：行尾到本栏右边界（容差 2 个字宽），下一行从本行起点或正文起点接着排（字号相近、
 *     行距不超过 1.75 个字宽、不以编号开头）；至少 3 处空隙；
 *   - 紧跟这样一行的下一行（段落末行等）：从上一行起点或正文起点起排，至少 2 处空隙，最宽不超过 2.5 个字宽；
 *   - 词距均匀的一行（发明名称）：至少 4 处空隙，最宽的不超过中位数的 1.6 倍、不超过 3 个字宽，多数片段是词。
 * 宽空隙在前后 3 行内又有两行与它对齐（空隙后的字从同一处起排，或空隙前的字在同一处结束：没认出的分栏、
 * 表单里对齐的填写栏），仍是制表间距
 */
const proseLines = new WeakSet<Line>();

function markProseLines(lines: Line[], r: Region, edges: number[], ctx?: Ctx): void {
  if (ctx?.mode !== "flow" || !lines.length) return;
  const limit = (l: Line) => Math.max(1.2 * l.size, 12);
  const gapsOf = (l: Line) => {
    const out: Array<{ x0: number; x1: number }> = [];
    for (let k = labelLead(l) ? 2 : 1; k < l.spans.length; k++) out.push({ x0: l.spans[k - 1].x1, x1: l.spans[k].x0 });
    return out;
  };
  const wide = lines.map((l) => gapsOf(l).filter((g) => g.x1 - g.x0 > limit(l)));
  const aligned = (i: number) => wide[i].some((g) => {
    let n = 0;
    for (let j = Math.max(0, i - 3); j <= Math.min(lines.length - 1, i + 3); j++) {
      // 对齐：空隙后的字从同一处起排（制表位、右栏的左边界），或空隙前的字在同一处结束（左栏的右边界）
      if (j !== i && wide[j].some((h) => Math.min(g.x1, h.x1) > Math.max(g.x0, h.x0) && (Math.abs(g.x1 - h.x1) <= 2 || Math.abs(g.x0 - h.x0) <= 2))) n++;
    }
    return n >= 2;
  });
  const continues = (l: Line, n: Line | undefined): boolean => !!n && !labelHead(n) &&
    Math.abs(lineEm(n) - lineEm(l)) <= 0.15 * lineEm(l) && n.baseline > l.baseline && n.baseline - l.baseline <= 1.75 * lineEm(l) &&
    (Math.abs(n.x0 - l.x0) <= 3 || Math.abs(n.x0 - textStart(l)) <= 3);
  const body = lines.map((l, i) => l.x1 >= Math.min(edges[i], r.x1) - Math.max(2 * lineEm(l), 0.04 * (edges[i] - textStart(l))) && continues(l, lines[i + 1]));
  lines.forEach((l, i) => {
    if (!wide[i].length || aligned(i)) return;
    const gaps = gapsOf(l).map((g) => g.x1 - g.x0);
    const max = Math.max(...gaps), em = lineEm(l);
    if (body[i] && gaps.length >= 3 && max <= 3.5 * em) { proseLines.add(l); return; }
    if (i > 0 && body[i - 1] && continues(lines[i - 1], l) && gaps.length >= 2 && max <= 2.5 * em) { proseLines.add(l); return; }
    const pos = gaps.filter((g) => g > 0).sort((a, b) => a - b);
    const med = pos.length ? pos[Math.floor(pos.length / 2)] : 0;
    const words = l.spans.slice(labelLead(l) ? 1 : 0).filter((s) => !s.img && s.text.trim());
    if (pos.length >= 4 && max <= 1.6 * med && max <= 3 * em && words.filter((s) => /\p{L}.*\p{L}/u.test(s.text)).length >= 0.6 * words.length) proseLines.add(l);
  });
}

function groupParagraphs(lines: Line[], r: Region, ctx?: Ctx): Array<{ lines: Line[]; edge: number }> {
  const paras: Line[][] = [];
  const paraEdge: number[] = [];
  const edges = localEdges(lines, r);
  markProseLines(lines, r, edges, ctx);
  const wraps = wrapEdges(lines, r, edges);
  // 这组行里占多数的字号（按字数）：只有正文字号的行才按"句子没写完"接到上一行（标题、题注另行判断）
  const bySize = new Map<number, number>();
  for (const l of lines) bySize.set(pitchKey(l.size), (bySize.get(pitchKey(l.size)) ?? 0) + lineChars(l));
  const bodySize = [...bySize.entries()].sort((x, y) => y[1] - x[1])[0]?.[0] ?? 0;
  lines.forEach((l, idx) => {
    const P = paras[paras.length - 1];
    if (P) {
      const a = P[P.length - 1];
      const gap = l.baseline - a.baseline;
      const pitch = P.length >= 2 ? P[P.length - 1].baseline - P[P.length - 2].baseline : 0;
      // 识别文字层里中日文字和西文混排的行（著录项目），行的字号按字数取、同一段各行忽大忽小：按行里最大的字号也算同一字号
      const sizeOk = Math.abs(l.size - a.size) <= 0.15 * a.size ||
        (ctx?.mode === "flow" && Math.abs(lineEm(l) - lineEm(a)) <= 0.15 * lineEm(a) && Math.abs(l.size - a.size) <= 0.35 * lineEm(a));
      // 行距上限：西文正文的行距约为字号的 1.2 倍，1.75 倍足以区分"段内换行"与"段间空行"；
      // 中文合同常用 1.5 倍 / 2 倍行距（基线间距是字号的 2.1–2.6 倍），按 1.75 倍判断会把每一行都切成一段。
      // 同字号的正文在全文中有一个占绝对多数的行距：不超过它（容差 6%）的都是段内换行
      const docPitch = ctx?.pitch?.get(pitchKey(a.size));
      const maxGap = Math.max(1.75 * Math.max(a.size, l.size), docPitch ? 1.06 * docPitch + 0.6 : 0);
      const gapOk = gap > 0.5 * a.size && gap <= maxGap && (!pitch || Math.abs(gap - pitch) <= 0.2 * pitch + 0.6);
      const firstText = l.spans[0].text.trim().split(/\s+/)[0] ?? "";
      // 新的条款 / 列表项（"1.2⇥…""(b)⇥…"）另起一段
      const startsBullet = (BULLET.test(firstText) && l.spans.length > 1) || labelHead(l);
      // 自动折行的判据：下一行的第一个词放不进上一行剩余的空间
      const edge = Math.max(wraps[idx - 1] ?? r.x1, a.x1);
      const wrapped = a.x1 + firstWordWidth(l) + 0.25 * a.size > edge - 0.5;
      const bodyLeft = P.length >= 2 ? P[1].x0 : null;
      // 首行以编号开头（悬挂缩进）：续行与编号后的正文对齐
      const aIsBullet = (BULLET.test(P[0].spans[0].text.trim()) && P[0].spans.length > 1) || labelLead(P[0], 0.1 * P[0].size);
      const leftOk = bodyLeft !== null
        ? Math.abs(l.x0 - bodyLeft) <= 3
        : l.x0 <= P[0].x0 + 3 || (aIsBullet && Math.abs(l.x0 - P[0].spans[1].x0) <= 3) ||
          // 编号比悬挂缩进还宽（"10.12""72.152"）：首行正文被推到下一个制表位，续行仍从悬挂缩进处起排、落在编号下方
          (aIsBullet && l.x0 > P[0].x0 + 3 && l.x0 < P[0].spans[1].x0 - 3 && P[0].spans[0].x1 > l.x0 - 1) ||
          // 编号伸在左边距里、首行正文空一两个字起排（日文专利"[0006]本技術…"，2026-11-26 WO2017038541A1）：续行从正文
          // 左边界起排，落在编号和首行正文之间。原来每段的首行单独成段，余下的行另成一段
          (aIsBullet && l.x0 > P[0].spans[0].x1 + 3 && l.x0 < P[0].spans[1].x0 - 3 && P[0].spans[1].x0 - l.x0 <= 2.2 * P[0].size);
      const centered = isCentered(a, r) && isCentered(l, r) && !hasBigGap(a) && !hasBigGap(l) && a.x1 + firstWordWidth(l) + 0.25 * a.size > r.x1 - (a.x0 - r.x0);
      // 折行后恰好以"[30]""(b)"开头的续行（高亮的占位符、行文中的列举自成一个文字片段）：后面只隔一个普通空格、
      // 行首与段落正文对齐，是上一行的延续，不是新的列表项（列表项的编号后是制表间距，或者编号伸出正文左边界）
      const inlineToken = startsBullet && l.spans.length > 1 && l.spans[1].x0 - l.spans[0].x1 <= 0.3 * l.size && (P.length >= 2 || aIsBullet) && wrapped && leftOk;
      // 左对齐（右边参差）的西文正文：区域右边界常比实际折行处更靠右（页眉页脚、页码伸到页边，文本框留有内边距），
      // 下一行的短词（and、of、the）看起来放得进上一行，按上面的判据会在句子中间断段。上一行占了大半行宽、
      // 停在词或逗号上（没有句末标点）、下一行接着是字母或数字、字体粗细不变，就是同一句话的折行
      // （标题等非正文字号的行、两行都居中的行另按原来的判据，不在此列）
      const runOn = !wrapped && leftOk && !hasBigGap(a, true) && pitchKey(a.size) === bodySize && !(isCentered(a, r) && isCentered(l, r)) &&
        continuesSentence(a, l, edge);
      // 目录（流式模式）：每条自成一段——上一行以引导点和页码结束就另起一段；条目的标题折成两行时（只有最后一行带引导点），
      // 前面的行接到这一条里
      const contents = ctx?.mode === "flow";
      if (contents && contentsLine(a)) { paras.push([l]); paraEdge.push(edges[idx]); return; }
      if (contents && contentsLine(l) && !P.some(contentsLine) && sizeOk && gapOk && leftOk && !hasBigGap(a, true) && !isCentered(a, r)) {
        P.push(l);
        paraEdge[paraEdge.length - 1] = Math.max(paraEdge[paraEdge.length - 1], edges[idx]);
        return;
      }
      if (sizeOk && gapOk && (!startsBullet || inlineToken) && !hasBigGap(l) && (((wrapped || runOn) && leftOk && !hasBigGap(a, true)) || centered)) {
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
  return { font: s.font, size, bold: s.bold, italic: s.italic, color: s.color, underline: s.underline, strike: s.strike, vertAlign, smallCaps: s.smallCaps, tracking: s.tracking, shade: s.bg };
}

const sameFmt = (a: Run, b: Omit<Run, "text" | "tab">) =>
  !a.tab && !a.br && !a.img && !b.img && a.font === b.font && a.size === b.size && a.bold === b.bold && a.italic === b.italic && a.color === b.color && !!a.underline === !!b.underline && !!a.strike === !!b.strike && a.vertAlign === b.vertAlign && !!a.smallCaps === !!b.smallCaps && a.tracking === b.tracking && a.shade === b.shade;

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
  if (spans.length > 1 && (BULLET.test(spans[0].text.trim()) || labelHead(l)) && spans[1].x0 - spans[0].x1 > 0.45 * l.size) spans = spans.slice(1);
  if (spans.some((s) => !s.natural)) return null;
  let natural = 0, word = 0, chars = 0, tabGaps = 0, reach = -Infinity;
  spans.forEach((s, i) => {
    natural += s.natural!;
    word += s.wordW ?? s.natural!;
    chars += s.text.length;
    const before = reach;
    reach = Math.max(reach, s.x1);
    if (i > 0) {
      const gap = s.x0 - spans[i - 1].x1;
      // 制表间距（并排的几位作者的邮箱、"公式 … 编号"）不是字距：不计入实际宽度，否则整行被算成"加宽"，
      // 每个字都被加上字符间距、在 Word 里变宽后互相挤到下一行。
      // 以此前所有文字的最右端为准：与前文重叠的符号（伸进上一行的根号）后面不算空隙
      const open = s.x0 - before;
      if (open > Math.max(1.2 * l.size, 12) && !proseLines.has(l)) { tabGaps += open; return; }
      if (gap > 0.1 * s.size && !/\s$/.test(spans[i - 1].text) && !/^\s/.test(s.text)) {
        natural += Math.min(s.spaceW ?? 0.25 * s.size, spans[i - 1].spaceW ?? 0.25 * s.size);
        word += Math.min(s.wordSpaceW ?? s.spaceW ?? 0.25 * s.size, spans[i - 1].wordSpaceW ?? spans[i - 1].spaceW ?? 0.25 * s.size);
        chars++;
      }
    }
  });
  return { actual: spans[spans.length - 1].x1 - spans[0].x0 - tabGaps, natural, word, chars };
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
  // 被拉开的识别文字行：行尾按替代字体的字宽算，参差一两个字宽
  const tol = lines.some((l) => proseLines.has(l)) ? Math.max(1.5, 2 * lineEm(lines[0])) : 1.5;
  const atEdge = body.filter((l) => Math.abs(paraRight - l.x1) <= tol);
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
  /** 改成制表位前导符的目录引导点（计入完整性校验） */
  leaders?: string;
  /** 全文各字号正文的主流行距（基线间距，pt）：判断段内换行的上限 */
  pitch?: Map<number, number>;
  /** 逐行一致模式：本页页首锚点段落上的文本框（排不进文字流原位置的行放在这里） */
  boxes?: NonNullable<Para["textBoxes"]>;
}

const pitchKey = (size: number) => Math.round(size * 2) / 2;

/**
 * 全文各字号正文的主流行距：每一行与它正下方（横向重叠、同字号）最近一行的基线间距，取占多数的值。
 * 多栏文档左右栏的行横向不重叠，不会互相配对；段间距、标题前后的空白是少数，不影响众数
 */
function dominantPitch(pages: Line[][]): Map<number, number> {
  const samples = new Map<number, number[]>();
  for (const ls of pages) {
    const sorted = [...ls].sort((a, b) => a.baseline - b.baseline);
    sorted.forEach((a, i) => {
      if (a.x1 - a.x0 < 6 * a.size) return;
      for (let j = i + 1; j < sorted.length; j++) {
        const b = sorted[j];
        const d = b.baseline - a.baseline;
        if (d > 3.5 * a.size) break;
        if (Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0) < 0.5 * Math.min(a.x1 - a.x0, b.x1 - b.x0)) continue;
        if (Math.abs(a.size - b.size) <= 0.15 * a.size && d >= 0.9 * a.size) {
          const k = pitchKey(a.size);
          if (!samples.has(k)) samples.set(k, []);
          samples.get(k)!.push(d);
        }
        break;
      }
    });
  }
  const out = new Map<number, number>();
  for (const [k, ds] of samples) {
    if (ds.length < 8) continue;
    const bins = new Map<number, number>();
    for (const d of ds) { const b = Math.round(d * 2) / 2; bins.set(b, (bins.get(b) ?? 0) + 1); }
    const score = (b: number) => (bins.get(b) ?? 0) + (bins.get(b - 0.5) ?? 0) + (bins.get(b + 0.5) ?? 0);
    const best = [...bins.keys()].sort((x, y) => score(y) - score(x) || x - y)[0];
    if (score(best) < 0.4 * ds.length) continue;
    out.set(k, median(ds.filter((d) => Math.abs(d - best) <= 0.75)));
  }
  return out;
}

const fontKey = (f: { font: string; size: number; bold: boolean; italic: boolean }) => `${f.font}|${f.size}|${f.bold ? 1 : 0}|${f.italic ? 1 : 0}`;

/** 行 → 段落属性（对齐、缩进、行距、制表位、文字 run） */
function buildPara(lines: Line[], r: Region, blockEdge: number, ctx: Ctx): Para {
  const W = r.x1 - r.x0;
  const first = lines[0];
  const size = median(lines.map((l) => l.size));
  const pitch = lines.length >= 2 ? median(lines.slice(1).map((l, i) => l.baseline - lines[i].baseline)) : 0;
  // 多行段落：行距就是 PDF 实测的基线间距；单行：按字号的常规单倍行距。
  // 单行按本行最大的字号算（著录项目里编号 8.5 pt、内容 6.5 pt 混排）：按中位字号算的行高放不下大字，
  // 只能写成"最小值"行距，排版时行高按字体实际高度撑开（CJK 回退字体更高），一行行累积下移，压到底图里的横线
  const big = lines.length === 1 ? Math.max(size, ...first.spans.filter((s) => s.text.trim() && !s.img).map((s) => s.size)) : size;
  const lineHeight = pitch || big * 1.17;

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
  // 各行起点相同的多行文字是一栏正文，不是居中：版心常被首页著录项目、表格撑得比正文宽，正文满行恰好左右留白相近时
  // 会被当成居中（2026-11-26 WO2017038541A1：全书 57 段正文居中，译文换行后整段居中排）。单独一行满到所在文字块
  // 右边界（附近 3 行以上同起点、同右界）的，也是正文
  const columnText = lines.length >= 2 ? leftSpread <= 3 : blockEdge < r.x1 - 1 && Math.abs(first.x1 - blockEdge) <= 1.5 && first.x1 - first.x0 > 0.6 * W;
  if (((lines.every((l) => isCentered(l, r)) && !columnText) || centeredBlock) && !lines.some((l) => hasBigGap(l))) align = "center";
  else if (lines.length >= 2 && isJustified(lines, paraRight)) align = "both";
  else if (lines.every((l) => isRight(l, r)) && !lines.some((l) => hasBigGap(l)) && (lines.length === 1 || leftSpread > 3)) align = "right";

  let indLeft = 0, firstLine = 0, indRight = 0;
  const tabs: TabStop[] = [];
  const runs: Run[] = [];

  if (align === "left" || align === "both") {
    const bulletHang = (BULLET.test(first.spans[0].text.trim()) || labelHead(first)) && first.spans.length > 1 && first.spans[1].x0 - first.spans[0].x1 > Math.max(0.45 * size, 2 * (first.spans[0].spaceW ?? 0.25 * size));
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
  // 悬挂缩进的条款（"1.2⇥正文…"，续行与正文对齐）：编号后写制表符，由悬挂缩进的位置对齐，
  // 不另设制表位（另设制表位的段落跨页时不能与下一页的续段合并）；改写编号、正文后对齐也不会乱
  const hangTab = (align === "left" || align === "both") && labelLead(first, 0.1 * size) && firstLine < -0.5 &&
    (Math.abs(first.spans[1].x0 - (r.x0 + indLeft)) <= 2 || (lines.length >= 2 && first.spans[0].x1 > r.x0 + indLeft - 1 && first.spans[1].x0 > r.x0 + indLeft));

  let labelTab = false;
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
      if (li === 0 && si === 1 && hangTab) {
        const last = runs[runs.length - 1];
        if (last && !last.tab) last.text = last.text.replace(/\s+$/, "");
        // 编号比悬挂缩进宽：原文正文被推到了后面的制表位，按原位置设一个制表位（Word 默认制表位的间隔与原文不一定相同）
        if (Math.abs(s.x0 - (r.x0 + indLeft)) > 2) {
          const pos = Math.round((s.x0 - r.x0) * 20) / 20;
          if (!tabs.some((t) => Math.abs(t.pos - pos) < 1)) tabs.push({ pos, align: "left" });
          labelTab = true;
        }
        runs.push({ text: "", tab: true, ...plainFmt(fmt), underline: undefined, strike: undefined, shade: undefined });
        text = text.replace(/^\s+/, "");
      } else if (si > 0) {
        const p = l.spans[si - 1];
        const gap = s.x0 - p.x1;
        const leftAligned = align === "left" || align === "both";
        if (leftAligned && gap > Math.max(1.2 * l.size, 12) && !proseLines.has(l)) {
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
          // 制表符只在前后文字都带下划线 / 删除线时才带（同一条线从前一段文字延续到后一段），否则空白处会多出一条线
          runs.push({ text: "", tab: true, ...plainFmt(fmt), underline: fmt.underline && p.underline, strike: fmt.strike && p.strike, shade: fmt.shade && fmt.shade === p.bg ? fmt.shade : undefined });
          text = text.replace(/^\s+/, "");
        } else if (gap > 0.18 * l.size && !/\s$/.test(runs[runs.length - 1]?.text ?? "") && !/^\s/.test(text)) {
          const pc = p.text.slice(-1), nc = text[0] ?? "";
          if (!(CJK.test(pc) && CJK.test(nc)) || gap > 0.6 * l.size) {
            const prevRun = runs[runs.length - 1];
            if (prevRun && !prevRun.tab && !prevRun.br && !prevRun.img && (prevRun.shade ?? "") !== (fmt.shade ?? "") && fmt.shade) prevRun.text += " ";
            else text = " " + text;
          }
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
  if (ctx.mode === "exact" && (!tabs.length || labelTab) && align !== "center") {
    lineRanges.forEach(({ line, start, end }, i) => {
      const m = lineMeasure(line);
      if (!m) return;
      const bulletSkipped = line.spans.length > 1 && (BULLET.test(line.spans[0].text.trim()) || labelHead(line)) && (line.spans[1].x0 - line.spans[0].x1 > 0.45 * line.size || (i === 0 && hangTab));
      const startX = i === 0 && labelTab ? r.x0 + tabs[0].pos : r.x0 + indLeft + (i === 0 && !bulletSkipped ? firstLine : 0);
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
      // 编号后的制表符原本停在悬挂缩进处（没有另设制表位）：缩进归零后这个位置要补成制表位，
      // 否则编号后的文字跳到行尾的右对齐制表位上，最后一段被挤到下一行
      const hang = Math.round(indLeft * 20) / 20;
      if (hangTab && !labelTab && !tabs.some((t) => Math.abs(t.pos - hang) < 1)) tabs.push({ pos: hang, align: "left" });
      runs.unshift({ text: "", tab: true, ...plainFmt(spanFormat(first.spans[0], first.size, first.baseline)), underline: undefined, strike: undefined });
      lineFits.forEach((f) => { f.start++; f.end++; });
      indLeft = 0; firstLine = 0;
    }
  }
  const lastLine = lines[lines.length - 1];
  // 目录的一条（2026-11-27 APA：每条都排满到右边界，整章的条目被当成一段正文的折行并成一段，译文里引导点丢掉、
  // 页码贴在标题后面）：groupParagraphs 已让每条自成一段，这里把引导点改成带前导符的右对齐制表位
  if (ctx.mode === "flow" && (align === "left" || align === "both") && contentsLine(lastLine)) {
    const dots = contentsTab(runs, tabs, Math.round((Math.min(lastLine.x1, r.x1) - r.x0) * 20) / 20);
    if (dots !== null) { align = "left"; ctx.leaders = (ctx.leaders ?? "") + dots; }
  }
  const kept = runs.map((x) => x.tab || x.br || !!x.text);
  // run 过滤后下标会变化：换算成过滤后的下标
  const remap = (k: number) => kept.slice(0, k).filter(Boolean).length;
  return {
    // 末行排满到右边界：以文字块自己的右边缘为准（版心可能被页码等撑宽，比正文的右边缘更靠右）
    // 左对齐的西文段落末行停在句子中间（没有句末标点、占了大半行宽）：下一页 / 下一栏接着写，也算排满
    lastFull: Math.min(r.x1, blockEdge > r.x0 ? blockEdge : r.x1) - lastLine.x1 <= Math.max(3, 0.04 * W) ||
      (align === "left" && openEnded(lastLine, Math.min(r.x1, blockEdge > r.x0 ? blockEdge : r.x1))),
    kind: "p", runs: runs.filter((_, k) => kept[k]), align, indLeft, indRight, firstLine, ...(labelTab && tabs.length === 1 ? { labelTab: true } : {}),
    spaceBefore: 0, lineHeight, tabs: tabs.sort((a, b) => a.pos - b.pos), size,
    lineFits: lineFits.map((f) => ({ ...f, start: remap(f.start), end: remap(f.end) })),
  };
}

/** Word 固定行距下，基线距行框顶部的距离（额外空间加在文字上方） */
const baselineOffset = (lineHeight: number, size: number) => lineHeight - 0.22 * size;

/** 把一组行在给定区域里排成段落，并按实测位置计算段前距。返回段落与结束位置 */
function layoutLines(lines: Line[], r: Region, cursor: number, ctx: Ctx, floats?: Map<Line, ImageBox[]>, boxes?: NonNullable<Para["textBoxes"]>): { paras: Para[]; cursor: number } {
  const paras: Para[] = [];
  // 逐行一致模式：原文里处于同一高度、却没能成为同一行的两行（扫描件 OCR 识别出的徽标碎字、叠在正文行首的杂点
  // "："、并排的两个小块）。排进文字流只能一前一后，后一行连同下面的内容整块往下挤，压到底图里的横线、图上。
  // 字少的那一行改为按原位置的文本框（与页眉页脚一样不参与文字流），字多的留在文字流里
  if (boxes) {
    const byTop = [...lines].sort((a, b) => a.top - b.top);
    const minor = new Set<Line>();
    for (let i = 0; i < byTop.length; i++) {
      for (let j = i + 1; j < byTop.length && byTop[j].top < byTop[i].bottom; j++) {
        const a = byTop[i], b = byTop[j];
        if (minor.has(a) || minor.has(b) || !sameRow(a, b)) continue;
        const m = lineChars(a) <= lineChars(b) ? a : b;
        if (lineChars(m) <= 12 && !floats?.get(m)?.length) minor.add(m);
      }
    }
    if (minor.size) {
      for (const l of minor) boxes.push(...lineBoxes(l, ctx));
      lines = lines.filter((l) => !minor.has(l));
    }
  }
  for (const { lines: group, edge } of groupParagraphs(lines, r, ctx)) {
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

/** 一行的可见字符数 */
const lineChars = (l: Line) => l.spans.reduce((n, sp) => n + (sp.img ? 1 : sp.text.replace(/\s/g, "").length), 0);
/** 两行在原文里处于同一高度（纵向重叠超过较矮一行的一半） */
const sameRow = (a: Line, b: Line) => Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top) >= 0.5 * Math.min(a.bottom - a.top, b.bottom - b.top);

/** 一行文字 → 按页面坐标定位的文本框（按大间距切开，每一段一个） */
function lineBoxes(l: Line, ctx: Ctx): NonNullable<Para["textBoxes"]> {
  const groups: Span[][] = [];
  for (const sp of l.spans) {
    const g = groups[groups.length - 1];
    if (g && sp.x0 - g[g.length - 1].x1 < 3 * l.size) g.push(sp);
    else groups.push([sp]);
  }
  return groups.map((g) => {
    const gl = finishLine(g);
    const w = gl.x1 - gl.x0 + 6;
    const p = buildPara([gl], { x0: gl.x0, x1: gl.x0 + w }, gl.x0 + w, ctx);
    p.align = "left"; p.indLeft = 0; p.indRight = 0; p.firstLine = 0; p.tabs = [];
    // 行高按最大的字号：混排的大字（叠在 "input image" 后面的 9 pt "："）超过固定行高时 Word 改用"最小值"行距，文字下沉
    const big = Math.max(gl.size, ...g.filter((sp) => sp.text.trim() && !sp.img).map((sp) => sp.size));
    const lh = big * 1.17;
    p.lineHeight = lh;
    return { x: gl.x0, y: gl.baseline - baselineOffset(lh, big), w, h: lh, para: p };
  });
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

/**
 * 线只在文字项的一部分下面：pdf.js 把同一行里同一字体的文字合成一个文字项，定义条款的第一行
 * "“Accounts Payable” shall mean (i) all trade accounts payable…"是一项，下划线（一个很扁的填充矩形）只在被定义的词下面。
 * applyTextDecorations 要求线至少盖住文字项的一半，盖不到的被当成分隔线，写成两行之间一个带下边框的空段落：定义条款
 * 从第一行之后断成两段（第一行单独成段、常被判成居中），译文换了字宽以后线与文字错开（2026-11-27 APA：357 条这样的线，
 * 26 条定义的译文被切成两半、中间补了"……"）。
 * 在线的起止处把文字项切开（按字宽找最近的字符边界，不切在词中间的空格上），线下那一段带下划线，这条线不再是分隔线。
 * 逐字母加宽排的词（"W I T N E S S E T H"）每个字母下面一小段线：相距不到一个字宽的几段合起来算一条。
 * 表格里的文字（keep）不动：表格已按原来的文字项分好格。
 */
function splitUnderlinedParts(spans: Span[], segs: Seg[], used: Set<Seg>, keep: Set<Span>): Span[] {
  const hs = mergeH(segs.filter((s) => !used.has(s)));
  if (!hs.length) return spans;
  const out: Span[] = [];
  for (const s of spans) {
    const width = s.x1 - s.x0;
    // 线在下划线的位置（与 applyTextDecorations 相同），整条落在这个文字项的左右范围内
    const band = keep.has(s) || s.img || !s.text.trim() || width <= 0 ? [] : hs
      .filter((h) => h.y >= s.baseline - 0.5 && h.y <= s.baseline + 0.3 * s.size && h.x0 >= s.x0 - 1 && h.x1 <= s.x1 + 1 && h.x1 - h.x0 >= 0.25 * s.size)
      .sort((a, b) => a.x0 - b.x0);
    if (!band.length) { out.push(s); continue; }
    const groups: HSeg[][] = [];
    for (const h of band) {
      const g = groups[groups.length - 1];
      if (g && h.x0 - Math.max(...g.map((x) => x.x1)) <= s.size) g.push(h);
      else groups.push([h]);
    }
    // 字符边界的横坐标：按目标字体的字宽分配文字项的实际宽度（两端对齐拉开的词间距按比例摊开）；没有字宽数据时按字数平分
    const chars = [...s.text];
    const ws = chars.map((ch) => measureWord(ch, s.font, s.bold, s.italic, s.size));
    const known = ws.every((x) => x !== null && x > 0);
    const unit = known ? (ws as number[]) : chars.map(() => 1);
    const total = unit.reduce((a, b) => a + b, 0);
    const xs = [s.x0];
    for (const u of unit) xs.push(xs[xs.length - 1] + (u / total) * width);
    const nearest = (x: number) => xs.reduce((best, v, k) => (Math.abs(v - x) < Math.abs(xs[best] - x) ? k : best), 0);
    const cuts: Array<[number, number]> = [];
    for (const g of groups) {
      let a = nearest(Math.min(...g.map((h) => h.x0))), b = nearest(Math.max(...g.map((h) => h.x1)));
      while (a < b && !chars[a].trim()) a++;
      while (b > a && !chars[b - 1].trim()) b--;
      if (b <= a || (cuts.length && a < cuts[cuts.length - 1][1])) continue;
      cuts.push([a, b]);
      for (const h of g) for (const src of h.src) used.add(src);
    }
    if (!cuts.length) { out.push(s); continue; }
    const piece = (a: number, b: number, underline: boolean): Span | null => {
      if (b <= a) return null;
      const text = chars.slice(a, b).join("");
      const f = (xs[b] - xs[a]) / width;
      return {
        ...s, text, x0: xs[a], x1: xs[b],
        natural: s.natural !== undefined ? s.natural * f : undefined,
        wordW: measureWord(text, s.font, s.bold, s.italic, s.size) ?? (s.wordW !== undefined ? s.wordW * f : undefined),
        underline: underline || s.underline || undefined,
      };
    };
    let at = 0;
    for (const [a, b] of cuts) {
      const before = piece(at, a, false), mid = piece(a, b, true);
      if (before) out.push(before);
      if (mid) out.push(mid);
      at = b;
    }
    const rest = piece(at, chars.length, false);
    if (rest) out.push(rest);
  }
  return out;
}

function applyTextDecorations(spans: Span[], segs: Seg[], used: Set<Seg>): HSeg[] {
  const free = mergeH(segs.filter((s) => !used.has(s)));
  const rules: HSeg[] = [];
  for (const h of free) {
    const under: Span[] = [], strike: Span[] = [];
    for (const s of spans) {
      const overlap = Math.min(h.x1, s.x1) - Math.max(h.x0, s.x0);
      if (overlap < 0.5 * (s.x1 - s.x0)) continue;
      if (h.y >= s.baseline - 0.5 && h.y <= s.baseline + 0.3 * s.size) under.push(s);
      else if (h.y < s.baseline - 0.15 * s.size && h.y > s.baseline - 0.45 * s.size) strike.push(s);
    }
    const marked = [...under, ...strike];
    // 下划线 / 删除线与文字等长；明显长出文字一截的是线（表单的填写线、无框表格的列线恰好从表头文字下方经过）：
    // 当成下划线会只剩文字下面那一段，伸出的部分丢失，前面的制表符也跟着带上下划线
    const over = marked.length ? Math.max(Math.min(...marked.map((s) => s.x0)) - h.x0, h.x1 - Math.max(...marked.map((s) => s.x1))) : 0;
    if (marked.length && over <= Math.max(2, 0.6 * Math.max(...marked.map((s) => s.size)))) {
      for (const s of under) s.underline = true;
      for (const s of strike) s.strike = true;
    } else if (h.x1 - h.x0 >= 20) rules.push(h);
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
  const scores: Array<{ x: number; two: number; eff: number }> = [];
  for (let x = r.x0 + 0.3 * W; x <= r.x0 + 0.7 * W; x += 1) {
    let L = 0, R = 0, cross = 0, two = 0;
    for (const l of lines) {
      const k = classify(l, x);
      if (k === "cross") { cross++; continue; }
      const left = l.spans.filter((s) => (s.x0 + s.x1) / 2 < x), right = l.spans.filter((s) => (s.x0 + s.x1) / 2 >= x);
      // 左边只有紧挨着 x 的一小段（右栏里按网格对齐的国家代码清单 "AE, AG, AL…"，第一个代码落在 x 左边）：
      // 这是从 x 左边不远处起头、跨过 x 的一行，不是左右两栏各有文字
      if (k === "two" && Math.min(...left.map((s) => s.x0)) > r.x0 + 0.5 * (x - r.x0)) { cross++; continue; }
      if (k === "two") two++;
      if (left.length && Math.max(...left.map((s) => s.x1)) - r.x0 > 0.2 * W) L++;
      if (right.length && r.x1 - Math.min(...right.map((s) => s.x0)) > 0.2 * W) R++;
    }
    // 两种证据取其大：两侧同行都有文字的行数（页首有通栏内容时可靠），或左右长行数减去跨栏行数（两栏基线不对齐时可靠）。
    // eff：再减去横跨这里的行数。扫描件的 OCR 文字字距很宽、日文按字格排版、国家代码这类清单按网格对齐，
    // 栏内的字距空隙也会连成一条"竖向空白"，两侧都有文字的行数甚至比真正的栏间距还多——但总有大量的行横跨它；
    // 真正的栏间距只被通栏的标题等少数行跨越
    scores.push({ x, two: Math.max(two, Math.min(L, R) - cross), eff: Math.max(two, Math.min(L, R)) - cross });
  }
  const top = Math.max(0, ...scores.map((s) => s.two));
  if (top < 4 || top < 0.15 * lines.length) return null;
  // 候选：eff 接近最高的位置；其中取离版心中线最近的（两栏版式大多对称；右栏的条目编号与正文之间的空隙也几乎不被跨越，
  // 但它偏在一侧），再取 eff 高的、连续区间长的
  const maxEff = Math.max(...scores.map((sc) => sc.eff));
  const nearEff = (sc: { eff: number }) => sc.eff >= maxEff - Math.max(2, 0.2 * Math.abs(maxEff));
  const mid = (r.x0 + r.x1) / 2;
  type Run = { a: number; b: number; two: number; eff: number };
  const center = (c: Run) => Math.abs((scores[c.a].x + scores[c.b].x) / 2 - mid);
  const better = (c: Run, o: Run) => {
    if (Math.abs(center(c) - center(o)) > 4) return center(c) < center(o);
    return c.eff !== o.eff ? c.eff > o.eff : c.b - c.a > o.b - o.a;
  };
  let bestRun: Run | null = null;
  for (let i = 0; i < scores.length; ) {
    const sc = scores[i];
    if (!nearEff(sc)) { i++; continue; }
    let j = i;
    while (j + 1 < scores.length && scores[j + 1].two === sc.two && scores[j + 1].eff === sc.eff) j++;
    const run = { a: i, b: j, two: sc.two, eff: sc.eff };
    if (!bestRun || better(run, bestRun)) bestRun = run;
    i = j + 1;
  }
  if (!bestRun || bestRun.two < 4 || bestRun.two < 0.15 * lines.length) return null;
  const best = { x: (scores[bestRun.a].x + scores[bestRun.b].x) / 2, two: bestRun.two };
  if (bestRun.eff <= 0) {
    // 横跨这条线的行不比被它分开的行少：要有成片的双栏（上下相连的 4 行以上两侧都有文字，中间没有横跨的行）才算分栏。
    // 按制表位对齐的清单（附图标记说明：每行几组"部件名，编号"）里，隔几行就有两三行的空隙上下对齐，
    // 当成分栏后这几行被拆成左右两栏，读序错乱
    let run = 0, longest = 0;
    for (const l of [...lines].sort((a, b) => a.top - b.top)) {
      const k = classify(l, best.x);
      if (k === "cross") run = 0;
      else if (k === "two") longest = Math.max(longest, ++run);
    }
    if (longest < 4) return null;
  }
  // 不跨栏的行：分栏线左侧 / 右侧的文字片段
  const parts: Array<{ l: Line; L: Span[]; R: Span[] }> = [];
  for (const l of lines) {
    if (classify(l, best.x) === "cross") continue;
    parts.push({ l, L: l.spans.filter((s) => (s.x0 + s.x1) / 2 < best.x), R: l.spans.filter((s) => (s.x0 + s.x1) / 2 >= best.x) });
  }
  const Ls = parts.filter((p) => p.L.length), Rs = parts.filter((p) => p.R.length);
  const endL = (p: { L: Span[] }) => Math.max(...p.L.map((s) => s.x1)), startR = (p: { R: Span[] }) => Math.min(...p.R.map((s) => s.x0));
  const g0 = edgeOf(Ls.map(endL), "max"), g1 = edgeOf(Rs.map(startR), "min");
  if (g1 - g0 < 6) return null;
  // 栏间距必须真是空白，两栏里必须是成段的正文。签字区、表单里左右并排的短标签（"(Signature)      (Signature)"）、
  // 无框表格的表头也是"两侧都有文字"的行，几行就够上面的打分门槛；误判成分栏后，栏边界与这些标签对不上
  // （左栏只到编号"1."的行尾），标签被挤出栏外、整片错位。这类行按单栏 + 制表位还原才对：
  //   - 文字伸进栏间距（越过左栏右边界 / 右栏左边界）的行不能多：真正的分栏只有个别长网址、悬挂标点越界；
  //   - 两栏都要有排满整栏的行（至少一半，最多要求 3 行）：并排的短标签、编号列排不满栏
  const tol = (l: Line) => Math.max(3, 0.5 * l.size);
  const intrude = parts.filter(({ l }) => l.spans.some((s) => s.x1 > g0 + tol(l) && s.x0 < g1 - tol(l))).length;
  if (intrude > Math.max(2, 0.2 * parts.length)) return null;
  const colL = g0 - Math.min(...Ls.flatMap((p) => p.L.map((s) => s.x0))), colR = Math.max(...Rs.flatMap((p) => p.R.map((s) => s.x1))) - g1;
  if (colL < 0.15 * W || colR < 0.15 * W) return null;
  const fullL = Ls.filter((p) => endL(p) <= g0 + tol(p.l) && endL(p) - Math.min(...p.L.map((s) => s.x0)) >= 0.6 * colL).length;
  const fullR = Rs.filter((p) => startR(p) >= g1 - tol(p.l) && Math.max(...p.R.map((s) => s.x1)) - startR(p) >= 0.6 * colR).length;
  const need = (n: number) => Math.min(3, Math.ceil(0.5 * n));
  if (fullL < need(Ls.length) || fullR < need(Rs.length)) return null;
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
    // 分数线：正上方和正下方都紧挨着文字（分子、分母），不是破折号
    const stacked = (above: boolean) => pg.spans.some((s) => s.text.trim() && s.x1 > g.x0 && s.x0 < g.x1 &&
      (above ? s.baseline <= y + 0.5 && s.baseline > y - 1.2 * s.size : s.top >= y - 1 && s.top < y + 1.2 * s.size));
    if (!ref || overlapped || len < 0.8 * ref.size || len > 6 * ref.size || (stacked(true) && stacked(false))) { keep.push(g); continue; }
    const n = Math.max(1, Math.round(len / ref.size)); // 长破折号宽 1 个字号
    added.push({ ...ref, text: "\u2014".repeat(n), x0: g.x0, x1: g.x1, bold: false, italic: false, natural: n * ref.size, wordW: n * ref.size, img: undefined, type3: undefined });
  }
  if (added.length) { pg.segs = keep; pg.spans = [...pg.spans, ...added]; }
}

/** 复选框字符的字体：与 Word 自带的复选框内容控件相同（☐ / ☒ 用 MS Gothic） */
const CHECKBOX_FONT = "MS Gothic";

/**
 * 表单里画出来的复选框（合同、申请表的勾选项）：四条边围成的小方框 → "☐"，框里画了勾 / 叉的 → "☒"。
 * 方框的上下边太短，既不是下划线也构不成表格，原先被丢掉、只剩左右两条竖线（"| |"）；流式模式下整个方框都不见了。
 * 还原成字符后随文字排版，在 Word 里也能直接改成 ☒ 打勾。
 * 条件：边长 4–20pt、近似正方形、四边齐全；不与其他线相连（表格的格子、连成一排的输入方格留给表格识别）；
 * 框里没有文字、色块、图片；同一行有文字（取它的字号与基线）。
 * 同样大小、孤立的方框但不是复选框（框住的字母 / 编号、同一行没有文字）：返回它的四条边，由调用方画成完整的方框
 */
function checkBoxes(pg: PageModel): Seg[] {
  type B = { x0: number; y0: number; x1: number; y1: number };
  const mid = (a: number, b: number) => (a + b) / 2;
  const near = (a: number, b: number) => Math.abs(a - b) <= 1.5;
  const same = (a: Seg, b: Seg) => a.horizontal === b.horizontal && near(a.x0, b.x0) && near(a.x1, b.x1) && near(a.y0, b.y0) && near(a.y1, b.y1);
  const H = pg.segs.filter((s) => s.horizontal), V = pg.segs.filter((s) => !s.horizontal);
  const goneSegs = new Set<Seg>(), gonePaths = new Set<PathBox>();
  const added: Span[] = [], frames: Seg[] = [];
  let marked = 0;
  for (const top of H) {
    const w = top.x1 - top.x0, yT = mid(top.y0, top.y1);
    if (goneSegs.has(top) || w < 4 || w > 20) continue;
    // 底边与两条竖边：竖边要恰好连起上下两条边（一列上下紧挨的方框，下一个框的顶边不能当成这个框的底边）
    let found: { box: B; sides: Seg[] } | null = null;
    for (const bottom of H) {
      const yB = mid(bottom.y0, bottom.y1);
      if (bottom === top || goneSegs.has(bottom) || !near(bottom.x0, top.x0) || !near(bottom.x1, top.x1) || yB - yT < 0.7 * w || yB - yT > 1.4 * w) continue;
      const side = (x: number) => V.find((v) => !goneSegs.has(v) && near(mid(v.x0, v.x1), x) && near(v.y0, yT) && near(v.y1, yB));
      const l = side(top.x0), r = side(top.x1);
      if (l && r) { found = { box: { x0: top.x0, y0: yT, x1: top.x1, y1: yB }, sides: [top, bottom, l, r] }; break; }
    }
    if (!found) continue;
    const { box, sides } = found;
    const pad = 1.5;
    const touches = (b: B) => b.x1 > box.x0 - pad && b.x0 < box.x1 + pad && b.y1 > box.y0 - pad && b.y0 < box.y1 + pad;
    const within = (b: B) => b.x0 >= box.x0 - pad && b.x1 <= box.x1 + pad && b.y0 >= box.y0 - pad && b.y1 <= box.y1 + pad;
    // 四条边（同一条边描了两遍的也算）之外，不能再有线段碰到方框
    const edges = pg.segs.filter((s) => sides.some((e) => same(s, e)));
    if (pg.segs.some((s) => !edges.includes(s) && touches(s))) continue;
    // 框上的勾 / 叉（斜线、曲线，勾常常伸出框外一点）：打了勾的复选框。与更大的图形相连的不是复选框；
    // 圆角方框角上的小圆弧不是勾
    const h = box.y1 - box.y0, cy = mid(box.y0, box.y1);
    const marks = pg.paths.filter((p) => (p.diagonal || p.curved) && touches(p) && Math.max(p.x1 - p.x0, p.y1 - p.y0) >= 0.3 * Math.min(w, h));
    const reach = { x0: box.x0 - 0.5 * w, y0: box.y0 - 0.5 * h, x1: box.x1 + 0.5 * w, y1: box.y1 + 0.5 * h };
    if (marks.some((p) => p.x0 < reach.x0 || p.x1 > reach.x1 || p.y0 < reach.y0 || p.y1 > reach.y1)) continue;
    const inBox = (s: Span) => mid(s.x0, s.x1) > box.x0 && mid(s.x0, s.x1) < box.x1 && mid(s.top, s.bottom) > box.y0 && mid(s.top, s.bottom) < box.y1;
    const occupied = pg.spans.some((s) => s.text.trim() && inBox(s)) || pg.fills.some(within) || pg.images.some(within);
    // 同一行的文字：竖直方向覆盖方框中心，字号与方框相称
    const row = occupied ? [] : pg.spans.filter((s) => s.text.trim() && !s.img && s.top < cy && s.bottom > cy && h >= 0.4 * s.size && h <= 1.6 * s.size);
    edges.forEach((s) => goneSegs.add(s));
    if (!row.length) {
      // 不是复选框：上下边留在线段里会被丢掉（或被当成框内文字的下划线），只剩左右两条竖线——改由调用方画出四条边
      frames.push(...sides);
      continue;
    }
    const dist = (s: Span) => Math.max(s.x0 - box.x1, box.x0 - s.x1);
    const ref = row.reduce((a, b) => (dist(b) < dist(a) ? b : a));
    for (const p of pg.paths) if (within(p) || marks.includes(p)) gonePaths.add(p);
    // 勾 / 叉在提取时记作"未能转换的矢量图形"，现在已还原
    marked += marks.filter((p) => (p.x1 - p.x0) * (p.y1 - p.y0) > 4).length;
    added.push({
      text: marks.length ? "☒" : "☐", x0: box.x0, x1: box.x1, top: ref.top, bottom: ref.bottom, baseline: ref.baseline, size: ref.size,
      font: CHECKBOX_FONT, bold: false, italic: false, color: top.color,
      natural: box.x1 - box.x0, wordW: ref.size, spaceW: ref.spaceW, wordSpaceW: ref.wordSpaceW,
    });
  }
  if (!goneSegs.size) return frames;
  pg.segs = pg.segs.filter((s) => !goneSegs.has(s));
  pg.paths = pg.paths.filter((p) => !gonePaths.has(p));
  pg.vectorShapes = Math.max(0, pg.vectorShapes - marked);
  if (added.length) pg.spans = [...pg.spans, ...added];
  return frames;
}

/** unread：没有文字层、看得出是成行文字的页（页序号，从 0 起）；它们的整页图在 Word 里标上 pdf2word-page N */
export function buildDocModel(pages0: PageModel[], opts: { mode?: LayoutMode; raster?: boolean; unread?: ReadonlySet<number>; scanBlocks?: boolean } = {}): DocModel {
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
  let figureCount = 0, equationCount = 0, inlineMathCount = 0, iconCount = 0;

  // 扫描页的位图层（见 scanLayer.ts）：可编辑排版下，文字画在位图上面的页不把整页位图垫在文字下面
  const scanPlan = flow ? planScanPages(pages) : new Map<PageModel, ScanPlan>();
  const scanStats = { textPages: [] as number[], formPages: [] as number[], blocks: 0, inline: 0, formWords: 0, furniture: 0, stray: 0, pageNumbers: 0, header: 0, blockPages: [] as number[], textBlocks: 0 };

  // 全文共识的栏间距（与下面分栏时的共识同一做法：≥ 20% 的文字页测得出栏间距时取其中位数），释义清单检测用：
  // 本页自己测不出栏间距、却会按共识被切成两栏的释义页（LTB-07 第 7 页）
  const docGutterX = (() => {
    const ls = pages.map((pg) => buildLines(pg.spans.filter((sp) => sp.text.trim() && !sp.img)));
    const found = ls.filter((l) => l.length >= 6).map((l) => detectGutter(l, { x0: Math.min(...l.map((x) => x.x0)), x1: Math.max(...l.map((x) => x.x1)) }));
    const xs = found.filter((g): g is NonNullable<typeof g> => !!g).map((g) => g.x).sort((a, b) => a - b);
    return xs.length && xs.length >= 0.2 * found.length ? xs[Math.floor(xs.length / 2)] : undefined;
  })();
  // 会被切成两栏的释义清单的词条、释义起点（表格里的文字不算）
  const listCols: DefinitionCols[] = [];
  for (const pg of pages) if (!pageBackgrounds(pg).length) detectDefinitionList(pg, new Set(detectTables(pg).tables.flatMap((t) => t.spans)), docGutterX, [], listCols);

  // 预处理：矢量图（渲染为图片）、表格、装饰线
  const pre = pages.map((pg) => {
    let figures: ImageBox[] = [];
    // 整页的扫描图 / 底图：作为页面背景单独处理，不参与图片、组图、版心的识别
    let bgs = pageBackgrounds(pg);
    if (bgs.length) { const b = new Set(bgs); pg.images = pg.images.filter((im) => !b.has(im)); }
    // 没有文字层、看得出是成行文字的扫描页，可编辑排版下要求分块时（opts.scanBlocks）：整页图换成抹掉了字的底图
    // （pdf2word-scan N，附图、条码、表格线留着），每块文字一个原位文本框，框里先放这块的原图（pdf2word-text N.K），
    // 调用方识别后换成文字，翻译后文字仍在原来的位置
    const blocks = flow && opts.scanBlocks && opts.unread?.has(pg.index) && pg.scan?.textBlocks?.blocks.length ? pg.scan.textBlocks : undefined;
    if (blocks && bgs.includes(pg.scan!.image)) {
      const whole = pg.scan!.image;
      const bg: ImageBox = { ...whole, png: blocks.background, role: "scan", pageNo: pg.index + 1 };
      scanStats.blockPages.push(pg.index + 1);
      scanStats.textBlocks += blocks.blocks.length;
      return { pg, tables: [], free: [], rules: [], figures: [], frames: [], bgs: bgs.map((im) => (im === whole ? bg : im)), formLines: undefined as Line[] | undefined, scanBlocks: blocks.blocks.map((b, k) => scanBlockBox(b, pg.index + 1, k + 1)) };
    }
    // 没有文字层、看得出是成行文字的扫描页：整页图标上页码，调用方识别出文字后可以替换
    if (opts.unread?.has(pg.index)) for (const im of bgs) { im.role = "page"; im.pageNo = pg.index + 1; }
    const plan = scanPlan.get(pg);
    if (plan) {
      const whole = pg.scan!.image;
      bgs = bgs.filter((im) => im !== whole);
      pg.images = pg.images.filter((im) => im !== whole);
      scanStats.furniture += plan.furniture;
      scanStats.header += plan.header.size;
      if (plan.form) {
        // 表单页：线框和框里的字合起来才看得懂。扫描图（识别出的字已从图里抹掉，剩下格线）整页垫在下面，
        // 每行文字按原位置放进文本框（独占一页）：翻译只换文字，格线和位置不动
        scanStats.formPages.push(pg.index + 1);
        // 格里没被识别成文字的词（只在扫描图里）：从格线图上取出来（垫底的换成抹掉它们的图），作为行内小图片放回原位、
        // 进那一行的文本框，标成 pdf2word-glyph，调用方识别成文字后和格里的字一样翻译
        const words = pg.scan!.formWords;
        const bg: ImageBox = words ? { ...whole, png: words.background } : whole;
        bg.role = "form";
        bg.pageNo = pg.index + 1;
        if (words) {
          const real = pg.spans.filter((sp) => sp.text.trim());
          const sizes = real.map((sp) => sp.size).sort((a, b) => a - b);
          const usual = sizes.length ? sizes[Math.floor(sizes.length / 2)] : 10;
          const added: Span[] = [];
          for (const p of words.pieces) {
            const im: ImageBox = { x0: p.x0, y0: p.y0, x1: p.x1, y1: p.y1, png: p.png, pxW: p.pxW, pxH: p.pxH, scan: p.kind, role: "glyph", ...(p.line ? { line: p.line } : {}) };
            // 独占一处的词：字号取本页文字的众数，基线在词的下沿
            const size = p.line?.size ?? usual, baseline = p.line?.baseline ?? p.y1;
            const near = (sp: Span) => Math.abs(sp.baseline - baseline) + Math.max(0, sp.x0 - p.x1, p.x0 - sp.x1) / 100;
            const ref = real.length ? real.reduce((a, b) => (near(b) < near(a) ? b : a)) : undefined;
            added.push({
              text: "\uFFFC", x0: p.x0, x1: p.x1, top: baseline - 0.75 * size, bottom: baseline + 0.22 * size, baseline, size,
              font: ref?.font ?? "SimSun", bold: false, italic: false, color: "000000", natural: p.x1 - p.x0, wordW: p.x1 - p.x0,
              spaceW: ref?.spaceW ?? 0.25 * size, wordSpaceW: ref?.wordSpaceW, img: im,
            });
          }
          pg.spans = [...pg.spans, ...added];
          scanStats.formWords += added.length;
        }
        return { pg, tables: [], free: [], rules: [], figures: [], frames: [], bgs: [bg], formLines: buildLines(pg.spans.filter((sp) => sp.text.trim())), scanBlocks: undefined as NonNullable<Para["textBoxes"]> | undefined };
      } else {
        scanStats.textPages.push(pg.index + 1);
        const real = pg.spans.filter((sp) => sp.text.trim());
        const tx0 = Math.min(...real.map((sp) => sp.x0)), tx1 = Math.max(...real.map((sp) => sp.x1));
        if (plan.pageNumber) {
          // 印成图的页码：换成这一页的页码文字，由页眉页脚的处理认成页码、写成页码域
          const { box, text } = plan.pageNumber;
          const size = Math.round((box.y1 - box.y0) / 0.72 * 2) / 2;
          const ref = real[0];
          pg.spans.push({
            text, x0: box.x0, x1: box.x1, top: box.y1 - 0.75 * size, bottom: box.y1 + 0.22 * size, baseline: box.y1, size,
            font: ref.font, bold: false, italic: false, color: "000000", natural: box.x1 - box.x0, wordW: box.x1 - box.x0, spaceW: ref.spaceW, wordSpaceW: ref.wordSpaceW,
          });
          scanStats.pageNumbers++;
        }
        for (const im of plan.images) {
          if (plan.header.has(im)) {
            // 每页都一样的页眉页脚块：在页眉页脚区域，由页眉页脚的处理写进 Word 页眉页脚
            pg.images.push(im);
          } else if (im.line) {
            im.role = "glyph";
            // 落在文字行里的漏认字：作为行内图片插回那一行
            const ref = real.reduce((a, b) => (Math.abs(b.baseline - im.line!.baseline) + Math.max(0, b.x0 - im.x1, im.x0 - b.x1) / 100 < Math.abs(a.baseline - im.line!.baseline) + Math.max(0, a.x0 - im.x1, im.x0 - a.x1) / 100 ? b : a));
            const size = im.line.size;
            pg.spans.push({
              text: "\uFFFC", x0: im.x0, x1: im.x1, top: im.line.baseline - 0.75 * size, bottom: im.line.baseline + 0.22 * size, baseline: im.line.baseline, size,
              font: ref.font, bold: false, italic: false, color: "000000", natural: im.x1 - im.x0, wordW: im.x1 - im.x0, spaceW: ref.spaceW, wordSpaceW: ref.wordSpaceW, img: im,
            });
            scanStats.inline++;
          } else if (im.scan === "glyphs" && im.x1 - im.x0 <= 16 && im.y1 - im.y0 <= 16) {
            // 换到下一行行首、独占一行的标点（句末的"。"）：接在上一行末尾，作为行内图片。
            // 其余不挨着任何一行文字的零散墨迹（污点、重影）不保留
            const own = real.some((sp) => Math.min(sp.bottom, im.y1) - Math.max(sp.top, im.y0) > 0);
            const above = own ? [] : real.filter((sp) => sp.baseline <= im.y0 + 1 && im.y0 - sp.baseline <= 2.2 * sp.size);
            const last = above.length ? Math.max(...above.map((sp) => sp.baseline)) : NaN;
            const row = above.filter((sp) => Math.abs(sp.baseline - last) < 1);
            const end = row.reduce<Span | undefined>((a, b) => (!a || b.x1 > a.x1 ? b : a), undefined);
            // 在行首：不比上一行的起点靠右两个字以上（段落编号栏右边、悬在正文左边的句号也算）
            if (end && im.x0 <= Math.min(...row.map((sp) => sp.x0)) + 2 * end.size) {
              im.role = "glyph";
              const w = im.x1 - im.x0, x0 = end.x1 + 0.3;
              pg.spans.push({
                text: "\uFFFC", x0, x1: x0 + w, top: end.baseline - 0.75 * end.size, bottom: end.baseline + 0.22 * end.size, baseline: end.baseline, size: end.size,
                font: end.font, bold: false, italic: false, color: "000000", natural: w, wordW: w, spaceW: end.spaceW, wordSpaceW: end.wordSpaceW, img: im,
              });
              scanStats.inline++;
            } else scanStats.stray++;
          } else if (im.x1 <= tx0 - 2 || im.x0 >= tx1 + 2) {
            // 版心左右之外的（页边的竖排条码）：首页固定在原位置，其余页不要（与页边的装饰线条同样处理）
            if (pg.index === pages[0].index) { im.pinned = true; pg.images.push(im); scanStats.blocks++; } else scanStats.furniture++;
          } else { pg.images.push(im); scanStats.blocks++; }
        }
      }
    }
    // 文字底色（高亮的占位符等）先认出来：它们是文字的属性，不能被当成图形截成图片
    textHighlights(pg);
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
        figures.push({ ...rasterBox({ x0: im.x0, y0: im.y0, x1: im.x1, y1: im.y1 }, pg.index, "figure", spansText(gone)), ...(im.scan ? { scan: im.scan } : {}) });
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
    const frames = checkBoxes(pg);
    // 作者名后的矢量小图标（ORCID）：渲染成图后作为行内图片
    if (opts.raster) {
      const t1 = detectTables(pg).tables.map((t) => ({ x0: t.x0, y0: t.y0, x1: t.x1, y1: t.y1 }));
      for (const box of detectIconPaths(pg, t1)) { removeInside(pg, box); pg.images.push(box); }
    }
    iconCount += inlineIcons(pg);
    const { tables, used } = detectTables(pg);
    const inTable = new Set(tables.flatMap((t) => t.spans));
    tables.push(...detectRuledTables(pg, used, inTable));
    tables.push(...detectSideBySide(pg, used, inTable));
    if (!bgs.length) tables.push(...detectDefinitionList(pg, inTable, docGutterX, listCols));
    pg.spans = splitUnderlinedParts(pg.spans, pg.segs, used, inTable);
    const free = pg.spans.filter((s) => !inTable.has(s));
    // 删除线 / 下划线也标在表格单元格里的文字上（2026-11-26 LTB-03 附件一：表内整行删除线丢失，原来只看表格外的文字）；
    // 没对上任何文字的线照旧
    const rules = applyTextDecorations(pg.spans, pg.segs, used);
    return { pg, tables, free, rules, figures, frames, bgs, formLines: undefined as Line[] | undefined, scanBlocks: undefined as NonNullable<Para["textBoxes"]> | undefined };
  });

  // 版心：全部页面的内容边界
  const xs0: number[] = [], xs1: number[] = [], tops: number[] = [], bottoms: number[] = [];
  const tableX0: number[] = [], tableX1: number[] = [];
  for (const { pg, tables, free } of pre) {
    // 以"行"为单位统计左右边界（片段级统计会被行内的词间断开干扰）
    const pls = buildLines(free);
    for (const l of pls) { xs0.push(l.x0); xs1.push(l.x1); }
    for (const t of tables) { tableX0.push(t.x0); tableX1.push(t.x1); }
    // 双栏的页面：两栏都必须落在版心里（栏宽 = 栏间距到版心边界）。只有一页是双栏时（扫描件首页的著录项目），
    // 全文统计会把它比正文宽的左右边界当成离群值滤掉，左栏从版心外起排、整页错乱。和表格一样把两栏的外边界算进版心
    if (pls.length >= 6) {
      const g = detectGutter(pls, { x0: Math.min(...pls.map((l) => l.x0)), x1: Math.max(...pls.map((l) => l.x1)) });
      if (g) {
        const side = (l: Line) => classify(l, g.x);
        const lx = pls.filter((l) => side(l) === "L" || side(l) === "two").map((l) => l.x0);
        const rx = pls.filter((l) => side(l) === "R" || side(l) === "two").map((l) => l.x1);
        if (lx.length >= 4 && rx.length >= 4) { tableX0.push(edgeOf(lx, "min")); tableX1.push(edgeOf(rx, "max")); }
      }
    }
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
    // 分栏节沿用本页的纸张大小（横向页面上的分栏仍是横向）；它接着上一节排，不另起一页
    curSection = { ...next, pageW: curSection.pageW, pageH: curSection.pageH, titlePg: flow && curPage === 0, ...(front && curPage >= front.first && curPage <= front.last ? { front: true } : {}) };
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
      // 页脚可以有好几行（期刊首页的 doi、收稿日期、版权行），但都与第一行页脚同样字号、且比正文小；
      // 遇到正文字号的行就停：页码上方紧挨着的正文末行、栏底的公式不能被卷进页脚
      const footLike = out.length > 0 && Math.abs(l.size - out[0].size) < 0.6 && l.size < 0.95 * mathCtx.bodySize;
      if (!prev || l.top - prev.bottom > 1.2 * l.size || footLike) out.push(l);
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
  // 行的身份：数字归一；只印罗马数字页码的行（"~i~""~ii~"）也归一，几页的页码行才认得出是同一行
  const norm = (l: Line) => {
    const t = l.spans.map((sp) => sp.text).join(" ").replace(/\s+/g, " ").trim();
    const rp = romanPage(t);
    return (rp ? t.slice(0, rp.at) + "#R" + t.slice(rp.at + rp.token.length) : t).replace(/\d+/g, "#");
  };
  // 两行紧挨着的页脚（"文档名" + "Page N"）：最下面一行与上一行几乎没有间距，上面逐行往上判断时第一行就停了，
  // 整组页脚留在正文里——流式模式下它连同上方的大段空白排在正文中间，下一页的内容接在它后面。
  // 最下面一组紧挨着、同样小字号、整组与上方正文之间有明显空白的行，在多页重复出现（数字归一后文字相同）就是页脚；
  // 只在一页出现的（栏底的脚注）不动
  const footBlocks = pageLines.map((ls, i) => {
    const H = pre[i].pg.height;
    if (floating[i].some((l) => l.top >= H / 2)) return [];
    const out: Line[] = [];
    for (let k = ls.length - 1; k >= 0; k--) {
      const l = ls[k], below = out[out.length - 1];
      if (l.top < H * 0.9 || l.size >= 0.95 * mathCtx.bodySize) break;
      if (below && (Math.abs(l.size - below.size) >= 0.6 || below.top - l.bottom > 1.2 * l.size)) break;
      out.push(l);
    }
    const above = ls[ls.length - 1 - out.length], first = out[out.length - 1];
    return out.length >= 2 && (!above || first.top - above.bottom > 1.2 * first.size) ? out : [];
  });
  const footPages = new Map<string, Set<number>>();
  const addFoot = (l: Line, i: number) => { const k = norm(l); if (!footPages.has(k)) footPages.set(k, new Set()); footPages.get(k)!.add(i); };
  floating.forEach((ls, i) => { for (const l of ls) if (l.top >= pre[i].pg.height / 2) addFoot(l, i); });
  footBlocks.forEach((b, i) => b.forEach((l) => addFoot(l, i)));
  footBlocks.forEach((b, i) => { if (b.length && b.every((l) => (footPages.get(norm(l))?.size ?? 0) >= 2)) floating[i].push(...b); });
  // 第三条判据：在多页的同一高度重复出现。
  // 中文合同的正文常常排满到版心最底部，页脚紧挨在下面（与正文之间只有几磅空白）；页脚两行左右各一行、
  // 两行字号不同、或与正文同字号时，上面两条判据都认不出来——页脚留在正文里，跨页的条款被它截断。
  // 页眉页脚每页都在同一高度（±3pt）、文字相同（只有页码等数字不同）：页面最下方 / 最上方连续的这种行就是页脚 / 页眉
  if (pages.length >= 2) {
    const nP = pages.length;
    const numsOf = (l: Line) => {
      const t = l.spans.map((sp) => sp.text).join(" ");
      const rp = romanPage(t.replace(/\s+/g, " ").trim());
      return rp ? [rp.value] : (t.match(/\d+/g) ?? []).map(Number);
    };
    // 页码：某一个位置上的数字随页序递增（印刷的数字 − 页序号在这一组里都相同）。
    // 只是数字不同的正文行（"34.4 对于前款所述事项…"与"47.6 对于前款所述事项…"恰好排在页首同一高度）不算
    const pageCounter = (grp: Array<{ i: number; l: Line }>) => {
      const ns = grp.map((c) => numsOf(c.l));
      const len = Math.min(...ns.map((x) => x.length));
      for (let j = 0; j < len; j++) {
        const off = new Map<number, number>();
        grp.forEach((c, k) => { const d = ns[k][j] - pre[c.i].pg.index; off.set(d, (off.get(d) ?? 0) + 1); });
        if (Math.max(...off.values()) >= Math.max(2, 0.8 * grp.length) && new Set(ns.map((x) => x[j])).size >= 2) return true;
      }
      return false;
    };
    const groups = new Map<string, Array<{ i: number; l: Line }>>();
    pageLines.forEach((ls, i) => {
      const H = pre[i].pg.height;
      for (const l of ls) {
        if (l.top < 0.8 * H && l.bottom > 0.2 * H) continue;
        const k = norm(l);
        if (!k) continue;
        if (!groups.has(k)) groups.set(k, []);
        groups.get(k)!.push({ i, l });
      }
    });
    const repeated = new Set<Line>();
    for (const [k, cs] of groups) {
      cs.sort((a, b) => a.l.baseline - b.l.baseline);
      for (let s = 0; s < cs.length; ) {
        let e = s + 1;
        while (e < cs.length && cs[e].l.baseline - cs[s].l.baseline <= 3) e++;
        const grp = cs.slice(s, e);
        s = e;
        const n = new Set(grp.map((c) => c.i)).size;
        if (n < 2) continue;
        // 带页码的行：数字随页序递增；其余的行：至少四成的页面都有（奇偶页不同的页脚各占一半）
        const counter = k.includes("#") && pageCounter(grp);
        if (!counter && n < Math.max(2, 0.4 * nP)) continue;
        for (const c of grp) repeated.add(c.l);
      }
    }
    if (repeated.size) pageLines.forEach((ls, i) => {
      const H = pre[i].pg.height, fl = new Set(floating[i]);
      const byY = [...ls].sort((a, b) => a.baseline - b.baseline);
      const take = (l: Line) => { if (!fl.has(l)) { floating[i].push(l); fl.add(l); } };
      for (let k = byY.length - 1; k >= 0 && byY[k].top >= 0.8 * H && (repeated.has(byY[k]) || fl.has(byY[k])); k--) take(byY[k]);
      for (let k = 0; k < byY.length && byY[k].bottom <= 0.2 * H && (repeated.has(byY[k]) || fl.has(byY[k])); k++) take(byY[k]);
    });
  }
  // 流式模式：只有 Word 页眉页脚能在该页原样再现的行才写进页眉页脚（数字逐位比较，只有页码位可以逐页不同）；
  // 只在一页出现的栏底标题、页首的条款标题、附件标题放回正文。见 furniture.ts
  let furniture: Furniture | undefined;
  /** 前置部分（罗马数字页码的几页）：单独一节，正文另起一节从 1 编页 */
  let front: FrontMatter | undefined;
  if (flow && pages.length > 1) {
    furniture = settleFurniture(floating, pre.map((x) => x.pg.height / 2), pre.map((x) => x.pg.index + 1));
    // 前置部分之后还有正文页才分节（整份文件都是罗马数字页码时就是一节）
    if (furniture.front && furniture.front.last < pages.length - 1) {
      front = furniture.front;
      // 前置部分从第一页开始时，第一节就是它；前面有封面时封面另成一节
      if (front.first === 0) curSection.front = true;
    }
    floating.forEach((_, i) => { floating[i] = furniture!.keep[i]; });
    // 写不进 Word 页眉页脚的页眉页脚（另一节的页眉、Bates 号）：不回正文，文字计入完整性校验，notes 里说明
    for (const ls of furniture.hidden) for (const l of ls) rasterText.push(l.spans.map((sp) => sp.text).join(""));
  }
  pageLines.forEach((ls, i) => {
    const f = new Set([...floating[i], ...(furniture?.hidden[i] ?? [])]);
    pageLines[i] = ls.filter((l) => !f.has(l));
  });
  ctx.pitch = dominantPitch(pageLines);

  // 分栏在整篇文档中通常一致：先逐页检测，再用最常见的栏间距补全"只有一栏有文字"的页面
  const gutters = pageLines.map((ls) => detectGutter(ls, body));
  const known = gutters.filter((g): g is NonNullable<typeof g> => !!g);
  // 全文共识只在分栏确实是本文的版式时才用：100 页里只有 1 页测出"栏间距"（字距很宽的单栏正文的误检），
  // 不能拿它去切其他页（扫描件首页的著录项目被这样从左栏中间切开，两栏文字交错、撑高版面，压到底图里的横线）
  const textPages = pageLines.filter((ls) => ls.length >= 6).length;
  if (known.length && known.length >= 0.2 * textPages) {
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
      const cross = ls.filter((l) => classify(l, docG.x) === "cross").length;
      if (ls.length >= 4 && cross <= 0.3 * ls.length) { gutters[i] = { ...docG }; return; }
      // 图页：通栏大图、通栏图题之后页面下方只剩几行双栏正文，行太少、又被通栏的图题拉高了"跨栏"比例。
      // 只要有 ≥ 2 行两侧都有文字、且左半止于全文的栏边界、右半起于全文的栏边界（两端对齐的双栏正文），就沿用全文的分栏
      const aligned = ls.filter((l) => {
        if (classify(l, docG.x) !== "two") return false;
        const L = l.spans.filter((sp) => (sp.x0 + sp.x1) / 2 < docG.x), R = l.spans.filter((sp) => (sp.x0 + sp.x1) / 2 >= docG.x);
        return Math.abs(Math.max(...L.map((sp) => sp.x1)) - docG.g0) <= 4 && Math.abs(Math.min(...R.map((sp) => sp.x0)) - docG.g1) <= 4;
      }).length;
      if (aligned >= 2) gutters[i] = { ...docG };
    });
    // 全文共识的栏边界要贴合本页的实际文字：期刊奇偶页的版心左右互换（偶页右栏从 302 起、奇页从 312 起），
    // 沿用共识值会让右栏最左的行落到"栏外"，被当成跨栏行：栏内的行间公式（整栏宽、右端带编号）被拆成零散的碎片
    pageLines.forEach((ls, i) => {
      const g = gutters[i];
      if (!g) return;
      // 每行取最靠近分栏线的那一侧的边：右栏的行首 x0 / 左栏的行尾 x1
      const ownR = ls.filter((l) => classify(l, g.x) !== "cross" && l.spans.some((sp) => (sp.x0 + sp.x1) / 2 >= g.x)).map((l) => Math.min(...l.spans.filter((sp) => (sp.x0 + sp.x1) / 2 >= g.x).map((sp) => sp.x0)));
      const ownL = ls.filter((l) => classify(l, g.x) !== "cross" && l.spans.some((sp) => (sp.x0 + sp.x1) / 2 < g.x)).map((l) => Math.max(...l.spans.filter((sp) => (sp.x0 + sp.x1) / 2 < g.x).map((sp) => sp.x1)));
      const g1 = ownR.length >= 4 ? Math.min(g.g1, edgeOf(ownR, "min")) : g.g1;
      const g0 = ownL.length >= 4 ? Math.max(g.g0, edgeOf(ownL, "max")) : g.g0;
      if (g1 - g0 >= 6 && (g1 !== g.g1 || g0 !== g.g0)) gutters[i] = { g0, g1, x: g.x > g0 && g.x < g1 ? g.x : (g0 + g1) / 2 };
    });
  }

  // 流式模式：上下边距按正文（不含页眉页脚）计算，页眉页脚写进真正的页眉页脚部件
  let mTop = marginTop, mBottom = marginBottom;
  let headerDist: number | undefined, footerDist: number | undefined;
  if (flow) {
    const bt: number[] = [], bb: number[] = [], ht: number[] = [], fb: number[] = [], ft: number[] = [];
    pre.forEach(({ pg, tables, figures }, i) => {
      const ls = pageLines[i];
      const pics = [...pg.images.filter((im) => !im.behind && im.y0 > pg.height * 0.1 && im.y1 < pg.height * 0.9), ...figures];
      const ys = [...ls.map((l) => l.top), ...tables.map((t) => t.y0), ...pics.map((f) => f.y0)];
      const ye = [...ls.map((l) => l.bottom), ...tables.map((t) => t.y1), ...pics.map((f) => f.y1)];
      if (ys.length) bt.push(Math.min(...ys));
      if (ye.length) bb.push(Math.max(...ye));
      for (const l of floating[i]) if (l.top < pg.height / 2) ht.push(l.top); else { fb.push(l.bottom); ft.push(l.top); }
    });
    mTop = Math.max(9, (bt.length ? Math.min(...bt) : 72) - 2);
    mBottom = Math.max(9, pageH - (bb.length ? Math.max(...bb) : pageH - 72) - 2);
    // 可编辑排版的版面元素比原文略高（同一高度并排的几条线各占一个极小段落），版心按原文正文的最低处算会正好排满：
    // 原文排到最低处的那一页（主动换页后单独成页的签字页）会被挤出一行到下一页。
    // 正文与页脚之间空着的地方留出至多两行余量（不会碰到页脚）
    if (ft.length && bb.length) {
      const room = Math.min(...ft) - Math.max(...bb) - 0.5 * mathCtx.bodySize;
      if (room > 0) mBottom = Math.max(9, mBottom - Math.min(room, 2.4 * mathCtx.bodySize));
    }
    if (ht.length) headerDist = Math.max(4, Math.min(...ht) - 1);
    if (fb.length) footerDist = Math.max(4, pageH - Math.max(...fb) - 1);
  }
  // 全文正文的上下界（按页统计会被"整页大图 + 一行图注"的页面误导）
  const allBody = pageLines.flat();
  const bodyTop = allBody.length ? Math.min(...allBody.map((l) => l.top)) : 0;
  const bodyBottom = allBody.length ? Math.max(...allBody.map((l) => l.bottom)) : pageH;
  // 流式模式：原文提前换页的地方（这一页下方大片空着、最后一段已经结束就换页，如"签字见下页"）保留分页符，
  // 否则下一页的内容会接着排上来。自然换页最多空出几行（孤行控制、放不下的图表），远不到版心的两成；
  // 最后一行排满到栏的右边界的，段落多半接到下一页，也不是主动换页
  const pageEnd = pre.map(({ pg, tables, figures }, i) => Math.max(0, ...pageLines[i].map((l) => l.bottom), ...tables.map((t) => t.y1),
    ...[...pg.images, ...figures].filter((im) => im.y0 < bodyBottom).map((im) => im.y1)));
  const endsMidParagraph = (i: number) => {
    const ls = pageLines[i];
    if (!ls.length) return false;
    const low = Math.max(...ls.map((l) => l.bottom));
    if (low < pageEnd[i] - 2) return false; // 页面最后是表格 / 图
    const last = ls.filter((l) => l.bottom >= low - 2).sort((a, b) => b.x1 - a.x1)[0];
    // 目录的最后一条排满到右边界（引导点），但它不接到下一页（2026-11-27 APA：目录最后一页下方空着，附表清单没有另起一页）
    if (contentsLine(last)) return false;
    const g = gutters[i], edge = g && last.x1 <= g.x ? g.g0 : right;
    return edge - last.x1 <= Math.max(3, 0.03 * (right - left));
  };
  // 目录在上一页结束、这一页不再是目录（附表清单、正文）：另起一页，不看上一页空出多少
  const lowest = (ls: Line[]) => ls.length ? ls.reduce((a, b) => (b.bottom > a.bottom ? b : a)) : undefined;
  const highest = (ls: Line[]) => ls.length ? ls.reduce((a, b) => (b.top < a.top ? b : a)) : undefined;
  const contentsEnd = (pi: number) => {
    const last = lowest(pageLines[pi - 1]), next = highest(pageLines[pi]);
    return !!last && contentsLine(last) && pageEnd[pi - 1] - last.bottom <= 2 && !(next && contentsLine(next)) &&
      !(next && pageLines[pi].some((l) => contentsLine(l) && l.top - next.bottom < 2 * l.size));
  };
  const brokeEarly = (pi: number) => flow && pi > 0 && ((!endsMidParagraph(pi - 1) &&
    bodyBottom - pageEnd[pi - 1] > Math.max(0.2 * (bodyBottom - bodyTop), 8 * mathCtx.bodySize)) || contentsEnd(pi));
  /** 每页的页眉 / 页脚行（流式模式） */
  const hf: Array<{ head: Line[]; foot: Line[]; imgs: ImageBox[]; mid: number }> = [];
  /** 首页栏底的注释（作者单位、基金等）：流式模式下放进首页页脚，正文增减时位置不变 */
  let notes: Para[] = [];
  let notesBottom = 0;

  for (const [pi, { pg, tables, free, rules, figures, frames, bgs }] of pre.entries()) {
    curPage = pi;
    stats.images += pg.images.length + bgs.length;
    // 整页背景：有文字压在上面时放在文字下方（表单底图），没有文字的扫描页浮于文字上方（在 Word 里点一下就能选中、裁剪）
    for (const im of bgs) im.behind = pg.spans.some((s) => s.text.trim() && s.x1 > im.x0 && s.x0 < im.x1 && s.bottom > im.y0 && s.top < im.y1);
    // 与文字重叠的图片（水印、底图）放在文字下方
    // 文字真正压在图片上（文字中心落在图片内部）才算底图；只是边缘相接的（作者照片旁边的简介）不算
    for (const img of pg.images) img.behind = !!img.pinned || pg.spans.some((s) => {
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
    // 不是复选框的小方框（框住的字母、编号）：四条边都画出来
    for (const s of frames) {
      if (inTableArea(s.x0, s.y0, s.x1, s.y1) || !keepShape(s.y0, s.y1)) continue;
      const x = (s.x0 + s.x1) / 2, y = (s.y0 + s.y1) / 2;
      shapes.push({ kind: "line", x0: s.horizontal ? s.x0 : x, y0: s.horizontal ? y : s.y0, x1: s.horizontal ? s.x1 : x, y1: s.horizontal ? y : s.y1, width: s.width, color: s.color, behind: false });
    }
    for (const f of pg.fills) {
      if (inTableArea(f.x0, f.y0, f.x1, f.y1) || !keepShape(f.y0, f.y1)) continue;
      shapes.push({ kind: "rect", x0: f.x0, y0: f.y0, x1: f.x1, y1: f.y1, width: 0, color: f.color, behind: true });
    }
    // 逐页保留：除了第一页，每页都从新的一页开始（按本次转换的顺序；只转换部分页时第一页不一定是原 PDF 的第 1 页）
    // 可编辑排版：整页背景（扫描页）所在的页、以及它后面的一页都从新的一页开始，否则下一页的内容会排到扫描图上
    const bgBreak = flow && pi > 0 && (bgs.length > 0 || pre[pi - 1].bgs.length > 0);
    // 前置部分的第一页（前面有封面）：封面一节到此结束，前置部分一节从下一页开始、按罗马数字编页；
    // 正文的第一页：前置部分一节到此结束，正文一节从下一页开始、从 1 编页（分节符本身换页，不再另加分页）
    const frontStart = !!front && pi > 0 && pi === front.first;
    const bodyStart = !!front && pi === front.last + 1;
    if (frontStart) {
      blocks.push(tinyPara({ sectionEnd: { ...curSection } }));
      curSection = { ...curSection, front: true, frontStart: true, titlePg: false };
    } else if (bodyStart) {
      blocks.push(tinyPara({ sectionEnd: { ...curSection, front: true } }));
      curSection = { ...curSection, front: undefined, frontStart: undefined, titlePg: false, restart: true };
    }
    // 页顶的窄条（公开号、条码）落在正文区域里的扫描页：同样从新的一页开始，窄条才在页顶（见 ScanPlan.topStrip）。
    // 落在页眉区域的（其他页的正文都排在它下面）写进这一页的页眉，不占正文，不必换页
    const stripBreak = flow && pi > 0 && [...(scanPlan.get(pg)?.topStrip ?? [])].some((im) => !inMarginZone(im.y0, im.y1));
    // 可编辑排版：纸张大小与全文不同的整页图页（横向的附图页夹在纵向的页中间）单独成节，用自己的纸张；
    // 回到全文纸张的下一页同样另起一节。分节符本身换页，不再另加分页
    let sized = false;
    if (flow && pi > 0 && !frontStart && !bodyStart) {
      const own = bgs.length > 0 && Math.abs(pg.width - pageW) + Math.abs(pg.height - pageH) > 0.05 * (pageW + pageH);
      const want = own ? `${pg.width}x${pg.height}` : "";
      const have = curSection.pageW !== undefined ? `${curSection.pageW}x${curSection.pageH}` : "";
      if (want !== have) {
        blocks.push(tinyPara({ sectionEnd: { ...curSection } }));
        curSection = { ...curSection, pageW: own ? pg.width : undefined, pageH: own ? pg.height : undefined, newPage: true, titlePg: false, restart: undefined, frontStart: undefined };
        sized = true;
      }
    }
    const anchor = tinyPara({ pageBreakBefore: !frontStart && !bodyStart && !sized && ((!flow && pi > 0) || brokeEarly(pi) || bgBreak || stripBreak), images: [...bgs, ...pg.images.filter((im) => im.behind)], textBoxes: [], shapes });
    freshPage = true;
    const anchorAt = blocks.length;
    blocks.push(anchor);
    let cursor = (flow ? mTop : marginTop) + TINY;
    if (flow) {
      hf.push({ head: floating[pi].filter((l) => l.top < pg.height / 2), foot: floating[pi].filter((l) => l.top >= pg.height / 2), imgs: [], mid: pg.height / 2 });
    } else {
      // 页眉页脚：每一段（按大间距切开）一个绝对定位的文字框
      for (const l of floating[pi]) anchor.textBoxes!.push(...lineBoxes(l, ctx));
    }
    ctx.boxes = flow ? undefined : anchor.textBoxes;
    // 表单页：每行文字一个按原位置定位的文本框，压在垫底的格线图上
    for (const l of pre[pi].formLines ?? []) anchor.textBoxes!.push(...lineBoxes(l, ctx));
    // 没有文字层的文字页：每块文字一个按原位置定位的文本框，压在抹掉了字的底图上
    for (const tb of pre[pi].scanBlocks ?? []) anchor.textBoxes!.push(tb);

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
      const parts = (l: Line) => ({ L: l.spans.filter((sp) => (sp.x0 + sp.x1) / 2 < gutter.x), R: l.spans.filter((sp) => (sp.x0 + sp.x1) / 2 >= gutter.x) });
      for (let i = 0; i < lines.length; ) {
        if (kinds[i] === "cross") { i++; continue; }
        let j = i + 1;
        let lastL = kinds[i] === "R" ? -Infinity : lines[i].bottom;
        while (j < lines.length && kinds[j] !== "cross" && lines[j].top - lines[j - 1].bottom < 2.5 * Math.max(lines[j].size, lines[j - 1].size)) {
          // 两侧都有字、却离左栏的上一行很远、左右两半又都伸进了栏间距的一行：不是两栏里的一行，是分栏区域下面的一行通栏文字，
          // 字距恰好落在栏间距上（WO2017038541A1 首页著录项目下面的 "(54) Title: VIDEO PROCESSING … METHOD, AND PROGRAM"，
          // 被切成左右两栏的两行，2026-11-27）。从这里另起一段再判断
          if (kinds[j] === "two" && lines[j].top - lastL >= 2.5 * lines[j].size) {
            const { L, R } = parts(lines[j]);
            if (Math.max(...L.map((sp) => sp.x1)) > gutter.g0 + 1 && Math.min(...R.map((sp) => sp.x0)) < gutter.g1 - 1) break;
          }
          if (kinds[j] !== "R") lastL = lines[j].bottom;
          j++;
        }
        const seg = lines.slice(i, j), ks = kinds.slice(i, j);
        const two = ks.filter((k) => k === "two").length;
        const edge = seg.filter((l, k) => atColEdge(l, ks[k])).length;
        const both = ks.filter((k) => k === "L").length >= 2 && ks.filter((k) => k === "R").length >= 2;
        // 只有一栏有文字、另一栏并排的是表格或图：同样属于分栏区域（例如左栏的表题紧挨右栏的表格）
        const segTop = Math.min(...seg.map((l) => l.top)), segBottom = Math.max(...seg.map((l) => l.bottom));
        // 或者紧挨着栏内的表格/图（两栏表题并排、表题与表格之间没有跨栏内容）
        const onlySide = ks.every((k) => k === "L") ? "L" : ks.every((k) => k === "R") ? "R" : null;
        // 版心外固定在页边的图（首页左边距里竖排的公开号条码）不在哪一栏旁边
        const facing = [...tables, ...figures, ...pg.images.filter((im) => !im.render && !im.pinned)].some((t) => {
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
    // 本页分栏正文实际的左 / 右边界（奇偶页版心左右互换，全文最左 / 最右的值对本页偏出几个点）：
    // 判断"右端带编号的行间公式"、居中、缩进都以栏的实际边界为准
    const colEdge = (() => {
      if (!gutter) return { l: left, r: right };
      const rs = lines.filter((l) => l.x0 >= gutter.g1 - 1 && l.x1 - l.x0 > 0.5 * (right - gutter.g1)).map((l) => l.x1);
      const ls = lines.filter((l) => l.x1 <= gutter.g0 + 1 && l.x1 - l.x0 > 0.5 * (gutter.g0 - left)).map((l) => l.x0);
      return { l: ls.length >= 4 ? Math.max(left, edgeOf(ls, "min")) : left, r: rs.length >= 4 ? Math.min(right, edgeOf(rs, "max")) : right };
    })();
    // 比第一页窄的页（横向文件里夹着的纵向页、按文字方向读的 /Rotate 页）：右边界按同样的右边距从本页右边缘量起，
    // 否则整页的行都比"版心"短得多，每行都被当成段落结束
    const pageRight = pg.width < pageW - 1 ? Math.max(left + 0.3 * (pg.width - left), Math.min(right, pg.width - (pageW - right))) : right;
    const regionOf = (side: "full" | "L" | "R"): Region =>
      side === "full" || !gutter ? { x0: left, x1: pageRight } : side === "L" ? { x0: colEdge.l, x1: gutter.g0 } : { x0: gutter.g1, x1: colEdge.r };
    for (const t of tables) items.push({ kind: "table", t, side: sideOf(t.x0, t.x1), top: t.y0 });

    // 旁边有文字的图片（流式模式）：挂在段落上、文字环绕；旁边那几行的缩进由环绕产生，行本身改为从栏边界起排
    const floatedLines = new Map<Extract<Item, { kind: "line" }> & { top: number }, { line: Line; reg: Region }>();
    const floatedImgs: ImageBox[] = [];
    const floatBeside = (img: ImageBox, side0: "full" | "L" | "R", besideLines: Array<Extract<Item, { kind: "line" }> & { top: number }>) => {
      // 旁边是通栏文字（期刊页眉：左徽标、右封面、中间刊名）：图片随通栏文字排，否则它单独落进分栏带，找不到可挂的段落
      const side = besideLines.length && besideLines.every((it) => it.side === "full") ? "full" : side0;
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
      for (const it of besideLines) {
        if (!floatedLines.has(it)) floatedLines.set(it, { line: it.line, reg });
        it.line = it.line.x0 >= img.x1 - 6 ? { ...it.line, x0: reg.x0 } : { ...it.line, x1: reg.x1 };
      }
      floatedImgs.push(img);
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
    // 从扫描图上裁下来的块：外接框含留白和线框，略微越过栏间距的中线不算通栏；没碰到另一栏的文字就属于这一栏
    const sideOfImage = (img: ImageBox): "full" | "L" | "R" =>
      img.scan && gutter ? (img.x1 <= gutter.g1 - 1 ? "L" : img.x0 >= gutter.g0 + 1 ? "R" : "full") : sideOf(img.x0, img.x1);
    for (const img of pics) {
      const side = sideOfImage(img);
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
    // 夹在两侧浮动图片之间、原本居中的行（期刊页眉的刊名）：在 Word 里居中于两图之间的空白，
    // 行的范围按版心对称放置，段落识别为居中
    for (const [it, { line: o, reg }] of floatedLines) {
      const overlaps = floatedImgs.filter((im) => im.y1 > o.top && im.y0 < o.bottom);
      const a = Math.max(reg.x0, ...overlaps.filter((im) => im.x1 <= o.x0 + 6).map((im) => im.x1));
      const b = Math.min(reg.x1, ...overlaps.filter((im) => im.x0 >= o.x1 - 6).map((im) => im.x0));
      const c = (o.x0 + o.x1) / 2, w = o.x1 - o.x0;
      if (Math.abs(c - (a + b) / 2) <= 3 && o.x0 - a > 4 && b - o.x1 > 4) {
        const rc = (reg.x0 + reg.x1) / 2;
        it.line = { ...it.line, x0: rc - w / 2, x1: rc + w / 2 };
      }
    }
    const groups: ImageBox[][] = [];
    for (const img of rest) {
      const g = groups.find((gr) => gr.some((o) => sideOfImage(o) === sideOfImage(img) &&
        Math.min(o.y1, img.y1) - Math.max(o.y0, img.y0) > 0.5 * Math.min(o.y1 - o.y0, img.y1 - img.y0)));
      if (g) g.push(img); else groups.push([img]);
    }
    for (const g of groups) {
      const overlap = g.some((a, i) => g.some((b, j) => j > i && Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0) > 2));
      if (overlap) { anchor.images!.push(...g); continue; } // 互相叠放的图片：保持页面绝对坐标
      const x0 = Math.min(...g.map((i) => i.x0)), x1 = Math.max(...g.map((i) => i.x1));
      // 与上面判断旁边有没有文字时用同一个归属：扫描图上裁下来的附图略微越过栏间距的中线（WO2017038541A1 首页的附图右缘
      // 在 325，栏间距中线 316），按通栏排会把右栏的摘要从附图处切成两段、各成一节（2026-11-27）
      const sides = new Set(g.map(sideOfImage));
      items.push({ kind: "image", imgs: g, side: sides.size === 1 ? [...sides][0] : sideOf(x0, x1), top: Math.min(...g.map((i) => i.y0)) });
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
        // 正文里有上标引用号（"…仲裁。¹"）对应的脚注：与后面各页的脚注一样留在本页正文末尾（小字号段落），
        // 不移进首页页脚——否则只有第一页的脚注进页脚、其余页的在正文里，同一份文件两种放法。
        // 首页页脚只收论文首页那种没有引用号的注释（作者单位、收稿日期）
        const refs = new Set(own.slice(0, k).flatMap((it) => (it.kind === "line" ? superscripts(it.line) : [])));
        const labelled = moved.filter((it) => it.kind === "line" && refs.has(noteLabel(it.line)));
        if (refs.size && labelled.length && labelled.length >= 0.5 * moved.filter((it) => it.kind === "line" && noteLabel(it.line)).length) continue;
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

    // 逐行一致模式：与分栏行处于同一高度的通栏行（扫描件首页：徽标下跨过栏间的 "P O | P C T"，与左栏的
    // "2017 年 3 月 9 日"、"W 画" 同高）。按带排版只能一带接一带，同一高度的几行被排成上下几行，
    // 后面的带整块下移，压到底图里的横线上。两行中字少的改为按原位置的文本框，不参与文字流
    if (ctx.boxes) {
      type LineItem = Extract<Item, { kind: "line" }> & { top: number };
      const lineItems = items.filter((it): it is LineItem => it.kind === "line");
      const gone = new Set<LineItem>();
      for (const a of lineItems) {
        if (a.side !== "full") continue;
        for (const b of lineItems) {
          if (b.side === "full" || !sameRow(a.line, b.line)) continue;
          const m = lineChars(a.line) <= lineChars(b.line) ? a : b;
          if (lineChars(m.line) <= 12) gone.add(m);
        }
      }
      for (const m of gone) {
        ctx.boxes.push(...lineBoxes(m.line, ctx));
        items.splice(items.indexOf(m), 1);
      }
    } else if (flow) {
      // 流式模式：与通栏行同一高度、只有几个字的分栏行并进那一行（徽标下的 "WIPO | PCT" 被识别成 "W 画" 和
      // "P O … P C T"，前者落在左栏范围、后者跨过栏间距）。否则它被挤到下面的分栏带里、占掉左栏的第一行，
      // 左栏的 (51) 比右栏的 (74) 低一行（2026-11-27 WO2017038541A1 首页）
      type LineItem = Extract<Item, { kind: "line" }> & { top: number };
      const lineItems = items.filter((it): it is LineItem => it.kind === "line");
      for (const a of lineItems) {
        if (a.side !== "full") continue;
        for (const b of lineItems) {
          if (b.side === "full" || !items.includes(b) || lineChars(b.line) > 12 || !sameRow(a.line, b.line)) continue;
          if (Math.abs(b.line.baseline - a.line.baseline) > 0.3 * Math.max(a.line.size, b.line.size)) continue;
          a.line = {
            ...a.line, spans: [...a.line.spans, ...b.line.spans].sort((p, q) => p.x0 - q.x0),
            x0: Math.min(a.line.x0, b.line.x0), x1: Math.max(a.line.x1, b.line.x1),
            top: Math.min(a.line.top, b.line.top), bottom: Math.max(a.line.bottom, b.line.bottom),
          };
          a.top = Math.min(a.top, b.top);
          items.splice(items.indexOf(b), 1);
        }
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
    // （原文在这里主动换页的，新的一页从页首排起，保留原来的上边距）
    if (flow && pi > 0 && items.length && !anchor.pageBreakBefore) cursor = items[0].top;
    const pageFirst = blocks.length;

    for (const band of bands) {
      if (!band.cols) {
        cursor += endSection({ cols: 1, colWidths: [right - left], colGap: 0 });
        const out = layoutItems(band.items, pageRight === right ? body : { ...body, x1: pageRight }, cursor, pg.fills, ctx);
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
          // 左栏以图片、表格结束（左栏是附图，右栏是摘要）：右栏的文字在它旁边，不是接着它往下排的续文。分栏符把右栏留在右栏；
          // 不加时 Word 平衡两栏，右栏开头的几行排到了左栏的附图下面（WO2017038541A1 首页，2026-11-27）
          const lastL = [...L.blocks].reverse().find((b) => !(b.kind === "p" && b.tiny && !b.inlineImages));
          const beside = R.blocks.length > 0 && !!lastL && (lastL.kind === "table" || (lastL.kind === "p" && !!lastL.inlineImages));
          // 右栏另起一项（著录项目 "(74) 代理人…"、以编号开头的条目），带又不高（不是满页的两栏正文）：右栏不是左栏的续文，
          // 也加分栏符。不加时由 Word / LibreOffice 平衡两栏决定右栏从哪一行起，译文左右两栏长短变化不一样，分栏点会移动；
          // 被平衡排到栏顶的那一段，段前距在栏顶被去掉（2026-11-27 WO2017038541A1 首页：(74) 比左栏的 (51) 高出一行）
          const firstR = band.items.find((i) => i.side === "R");
          const bandH = Math.max(...band.items.map((i) => (i.kind === "line" ? i.line.bottom : i.top))) - Math.min(...band.items.map((i) => i.top));
          const fresh = !beside && R.blocks.length > 0 && firstR?.kind === "line" && labelHead(firstR.line) && bandH <= 0.5 * pg.height;
          const lead: Para[] = beside || fresh ? [tinyPara({ columnBreakBefore: true })] : [];
          const f = R.blocks[0];
          // 右栏顶部离带顶的距离：分栏符后的段前距 LibreOffice 会去掉，改成固定高度的空段（两种排版引擎都保留）
          if (fresh && f?.kind === "p" && f.spaceBefore > 2 * TINY) { lead.push(spacer(f.spaceBefore - TINY)); f.spaceBefore = 0; }
          blocks.push(...L.blocks, ...lead, ...R.blocks);
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
    if (flow && !anchor.pageBreakBefore && !anchor.images!.length && !anchor.shapes!.length && !anchor.textBoxes!.length) blocks.splice(anchorAt, 1);
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
    const base = r.tracking !== undefined ? Math.round(r.tracking * 20) / 20 : r.paraSpacing ?? sp ?? 0;
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
  const hfOut = flow ? buildHeaders(hf, notes, notesBottom, { x0: left, x1: right }, right, ctx, headerDist ?? mTop / 2, furniture, front) : undefined;
  // 完整性校验：每页计入 Word 在这一页实际显示的页眉页脚（页码域按该页的页码展开），而不是判成页眉页脚的原文——
  // 判错的行若没有写进去，缺的字会如实计为未转换
  if (hfOut) {
    const shown = (ps: Para[] | undefined, i: number) => (ps ?? []).map((p) => p.runs.map((r) =>
      r.field === "PAGE" ? String(i + hfOut.start).padStart(r.fieldDigits ?? 0, "0") : r.field === "NUMPAGES" ? String(pre.length) : r.tab ? "" : r.text).join("")).join("\n");
    const fo = hfOut.front;
    // 前置部分第 i 页印的罗马数字：它第一页印 fo.start
    const romanAt = (i: number, upper: boolean) => { const t = toRoman(fo!.start + i - front!.first); return upper ? t.toUpperCase() : t; };
    const shownFront = (ps: Para[] | undefined, i: number) => (ps ?? []).map((p) => p.runs.map((r) =>
      r.field === "PAGE" && r.fieldFormat ? romanAt(i, r.fieldFormat === "ROMAN")
        : r.field === "PAGE" ? String(i + hfOut.start) : r.tab ? "" : r.text).join("")).join("\n");
    pre.forEach((_, i) => {
      const part = i === 0 && hfOut.headers.first ? "first" : hfOut.headers.even && (i + hfOut.start) % 2 === 0 ? "even" : "default";
      // 前置部分的页（封面用首页的那一套，与正文相同）
      if (fo && front && i >= front.first && i <= front.last && part !== "first") {
        rasterText.push(shownFront(fo.headers.default, i), shownFront(fo.footers.default, i));
        return;
      }
      rasterText.push(shown(hfOut.headers[part], i), shown(hfOut.footers[part], i));
    });
  }
  // 去掉的排版断词连字符：计入完整性校验
  if (ctx.softHyphens) rasterText.push("-".repeat(ctx.softHyphens));
  if (ctx.leaders) rasterText.push(ctx.leaders);

  return {
    pageW, pageH,
    margins: flow ? { top: mTop, bottom: mBottom, left, right: pageW - right } : { top: marginTop, bottom: marginBottom, left, right: pageW - right },
    blocks,
    finalSection: curSection,
    bodyFont, bodySize,
    stats: {
      ...stats, images: stats.images + iconCount, figures: figureCount, equations: equationCount, inlineMath: inlineMathCount,
      ...(scanStats.textPages.length || scanStats.formPages.length || scanStats.blockPages.length ? { scan: scanStats } : {}),
    },
    rasterText: rasterText.join("\n"),
    mode: ctx.mode,
    ...(hfOut ? { headers: hfOut.headers, footers: hfOut.footers, pageNumStart: hfOut.front ? undefined : hfOut.pageNumStart, headerDist, footerDist } : {}),
    ...(hfOut?.front ? { front: hfOut.front } : {}),
    ...(furniture?.notes.length ? { notes: furniture.notes } : {}),
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
    if (!a.lastFull || (a.tabs.length && !a.labelTab) || b.tabs.length || b.floats?.length) return false;
    const at = textOfPara(a).trimEnd(), bt = textOfPara(b).trimStart();
    // 下一页页首只剩一行续行时，左右的空白可能恰好接近对称，被误判成居中：
    // 上一段在句子中间断开、这一行以句末标点收尾，仍是续行（居中的标题不以句号结尾）
    const strayCenter = b.align === "center" && !b.runs.some((r) => r.br) && /[。.!?！？；;]["'”’)\]]?$/.test(bt) && !/[.!?。！？:：；;]["'”’)\]]?$/.test(at);
    if (!(a.align === "both" || a.align === "left") || !(b.align === "both" || b.align === "left" || strayCenter)) return false;
    if (Math.abs(a.size - b.size) > 0.3 || Math.abs(b.firstLine) > 0.3 * b.size) return false;
    if (isHeadingCandidate(a, bodySize, bodyFont) || isHeadingCandidate(b, bodySize, bodyFont)) return false;
    if (!at || !bt || BULLET.test(bt.split(/\s/)[0]) || CLAUSE_LABEL.test(bt.split(/\s/)[0])) return false;
    // 图题、表题自成一段（"TABLE V: …""Fig. 2: …"），不能并进上一段
    if (/^((fig\.?|figure|table)\s*[\dIVXLC]+[a-z]?\s*[:.．：]|TABLE\s+[IVXLC\d]+\b|[图表]\s*\d+)/i.test(bt)) return false;
    // 悬挂缩进的条款在页首接着排：下一页第一段没有编号、左缩进与上一页末段的正文对齐，就是同一条款的续行
    // （新的条款都以编号开头，已在上面排除）
    const hangCont = a.firstLine < -0.5 && Math.abs(b.indLeft - a.indLeft) <= 1.5 && Math.abs(b.firstLine) <= 0.5;
    return indentStyle || hangCont || !/[.!?。！？:：]["'”’)\]]?$/.test(at) || /^[a-z]/.test(bt);
  };
  const out: Block[] = [];
  let last: Para | null = null;
  for (const b of blocks) {
    if (b.kind === "p" && b.tiny) {
      if (b.sectionEnd || b.inlineImages || b.borderBottom || b.pageBreakBefore) last = null;
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
function buildHeaders(hf: Array<{ head: Line[]; foot: Line[]; imgs: ImageBox[]; mid: number }>, notes: Para[], notesBottom: number, body: Region, right: number, ctx: Ctx, headerTop: number,
  furniture?: Furniture, front?: FrontMatter):
  { headers: HeaderSet; footers: HeaderSet; pageNumStart?: number; start: number; front?: NonNullable<DocModel["front"]> } | undefined {
  const n = hf.length;
  if (!hf.some((x) => x.head.length || x.foot.length || x.imgs.length) && !notes.length) return undefined;
  const lineText = (l: Line) => l.spans.map((s) => s.text).join(" ");
  // 页码：印刷的数字 - 页序号 在多数页上一致
  let pageNumStart: number | undefined;
  const pick: FurniturePick | undefined = furniture?.pick;
  if (pick) pageNumStart = pick.pageNumStart;
  else if (n >= 2) {
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
  // 总页数（"第 3 页 共 12 页""Page 3 of 12""3 / 12"）：页码从 1 起、每页都印着同一个数且等于页数时写成 NUMPAGES 域，
  // 可编辑排版重新分页、增删内容后仍然正确。只转换了部分页时印的总数大于页数，保持原文
  const TOTAL = new RegExp(`(?:共\\s*|of\\s+|/\\s*|／\\s*)${n}(?!\\d)`);
  const totalPages = pageNumStart === 1 && n >= 2 &&
    hf.filter((x) => [...x.head, ...x.foot].some((l) => TOTAL.test(lineText(l).replace(/\s+/g, " ")))).length >= Math.max(2, Math.ceil(n * 0.5));

  const markPage = (p: Para, num: number, field: "PAGE" | "NUMPAGES" = "PAGE", digits = 0) => {
    const s = String(num).padStart(digits, "0");
    const re = new RegExp(`(?<!\\d)${s}(?!\\d)`, "g");
    let idx = field === "PAGE" ? p.runs.findIndex((r) => !r.tab && !r.field && r.text.trim() === s) : -1;
    let at = -1;
    if (idx < 0) {
      for (let k = p.runs.length - 1; k >= 0 && idx < 0; k--) {
        const r = p.runs[k];
        if (r.tab || r.field) continue;
        const ms = [...r.text.matchAll(re)];
        if (ms.length) { idx = k; at = ms[ms.length - 1].index!; }
      }
    } else at = p.runs[idx].text.indexOf(s);
    if (idx < 0) return false;
    const r = p.runs[idx];
    const parts: Run[] = [];
    if (at > 0) parts.push({ ...r, text: r.text.slice(0, at) });
    // 补零的页码：各页页码位数相同时，前导零写成文字、后面接普通 PAGE 域（LibreOffice 等不支持数字格式开关的
    // 程序也显示正确）；位数会变（第 9 页到第 10 页）时才用 PAGE \# "000000"
    const lead = digits > String(num).length ? s.slice(0, digits - String(num).length) : "";
    const sameWidth = String(printed(0)).length === String(printed(n - 1)).length;
    if (lead && sameWidth) parts.push({ ...r, text: lead }, { ...r, text: String(num), field });
    else parts.push({ ...r, text: s, field, ...(lead ? { fieldDigits: digits } : {}) });
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
      // 先标总页数（最后一页的页码与总页数相同，总数在后）
      if (totalPages && TOTAL.test(p.runs.map((r) => r.text).join("").replace(/\s+/g, " "))) markPage(p, n, "NUMPAGES");
      // 页码域只放在真正的页码位上（furniture.ts 逐位确认过的行）；页眉里恰好等于该页页码的版本号、条款号保持原文
      if (pageNumStart !== undefined && (!furniture || furniture.pageNumberLines.has(l))) markPage(p, printed(i));
      const digits = furniture?.paddedPageLines.get(l);
      if (pageNumStart !== undefined && digits) markPage(p, printed(i), "PAGE", digits);
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
  // 代表页：furniture.ts 已按文字身份选定（其余页只留下了与代表页相同的行，每一行都能原样再现）。
  // 页眉页脚里只有图片（扫描页上每页相同的公开号条）时它没有依据，按多数页选
  const byText = pick && hf.some((x) => x.head.length || x.foot.length) ? pick : undefined;
  const oddRep = byText ? byText.defPage : major(odd), evenRep = byText ? byText.evenRep : major(even);
  const evenOdd = byText ? byText.evenOdd : oddRep !== undefined && evenRep !== undefined && sigAt(oddRep) !== sigAt(evenRep);
  const defPage = byText ? byText.defPage : evenOdd ? oddRep! : major(rest) ?? 0;
  const rep0 = evenOdd && printed(0) % 2 === 0 ? evenRep! : defPage;
  const titlePg = notes.length > 0 || (byText ? byText.titlePg : n > 1 && sigAt(0) !== sigAt(rep0));
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
  // 前置部分：取它第一页的页眉页脚，罗马数字页码换成罗马数字格式的 PAGE 域；封面仍用首页那一套
  let frontOut: NonNullable<DocModel["front"]> | undefined;
  if (front) {
    const i = front.first;
    const fHeaders: HeaderSet = { default: mk(hf[i].head, i) }, fFooters: HeaderSet = { default: mk(hf[i].foot, i, "foot") };
    const token = front.upper ? toRoman(front.value).toUpperCase() : toRoman(front.value);
    const re = new RegExp(`(?<![A-Za-z])${token}(?![A-Za-z])`);
    let marked = false;
    for (const p of [...fFooters.default!, ...fHeaders.default!]) {
      const k = p.runs.findIndex((r) => !r.tab && !r.field && re.test(r.text));
      if (k < 0 || marked) continue;
      const r = p.runs[k], at = r.text.search(re);
      const parts: Run[] = [];
      if (at > 0) parts.push({ ...r, text: r.text.slice(0, at) });
      parts.push({ ...r, text: token, field: "PAGE", fieldFormat: front.upper ? "ROMAN" : "roman" });
      if (at + token.length < r.text.length) parts.push({ ...r, text: r.text.slice(at + token.length) });
      p.runs.splice(k, 1, ...parts);
      marked = true;
    }
    if (marked) {
      if (titlePg) { fHeaders.first = headers.first; fFooters.first = footers.first; }
      frontOut = { headers: fHeaders, footers: fFooters, fmt: front.upper ? "upperRoman" : "lowerRoman", start: front.value, bodyStart: printed(front.last + 1) };
    }
  }
  return { headers, footers, pageNumStart: pageNumStart !== undefined && pageNumStart !== 1 ? pageNumStart : undefined, start: pageNumStart ?? 1, ...(frontOut ? { front: frontOut } : {}) };
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
    const out = layoutLines(pending, r, cursor, ctx, floats, ctx.boxes);
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
      const gap = Math.max(0, it.h.y - cursor - 1);
      p.indLeft = Math.max(0, it.h.x0 - r.x0);
      p.indRight = Math.max(0, r.x1 - it.h.x1);
      // 相邻两段的下边框、缩进都相同时，Word / LibreOffice 把它们并成一组、只在组末画一条线（表单里连着的三条填写线只剩一条）：
      // 中间隔一个没有边框的空段，空段的高度就是原来的段前距
      const prev = blocks[blocks.length - 1];
      const sep = prev?.kind === "p" && !!prev.borderBottom && Math.abs(prev.indLeft - p.indLeft) < 0.05 && Math.abs(prev.indRight - p.indRight) < 0.05;
      if (sep) blocks.push(spacer(gap));
      else p.spaceBefore = gap;
      blocks.push(p);
      // 同一高度并排的几条线（左右两条签字线、无框表格每列一条）各占一个极小段落：位置按它们实际占的高度累计，
      // 否则每多一条线下面的内容就往下错 1pt，一页下来累积十几磅，单独成页的签字页被挤出一行到下一页
      cursor += (sep ? Math.max(1, gap) : gap) + TINY;
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

/**
 * 没有文字层的文字页上的一块文字（ScanTextBlock）：按原位置定位的文本框，框里一段，段里先放这块的原图（pdf2word-text 页.块，
 * 带上估计的字体、字号）。调用方把它识别成文字、换成文字 run 时沿用这些格式；段落的行距是原文的行距（图在时按"最小值"，
 * 不裁图），两端对齐与否按原文。框按框宽折行、随文字长高，译文比原文长时不会被裁掉
 */
function scanBlockBox(b: ScanTextBlock, page: number, no: number): NonNullable<Para["textBoxes"]>[number] {
  const img: ImageBox = { x0: b.x0, y0: b.y0, x1: b.x1, y1: b.y1, png: b.png, pxW: b.pxW, pxH: b.pxH, role: "text", pageNo: page, blockNo: no };
  const run: Run = { text: "\uFFFC", font: "Times New Roman", size: b.size, bold: false, italic: false, color: "000000", img };
  // 单行块：行距不超过块本身的高度。行比图高时图落在基线上、整块往下沉（页眉、表格里的数值下沉 2.5–6 pt）；
  // 多行块的图比行距高得多，行高就是图高，不受影响
  const lineHeight = b.lines > 1 ? b.pitch : Math.min(b.pitch, b.y1 - b.y0);
  const para: Para = { kind: "p", runs: [run], align: b.align, indLeft: 0, indRight: 0, firstLine: 0, spaceBefore: 0, lineHeight, tabs: [], size: b.size };
  // 调用方回填译文、接续跨栏的段落时要用的版面信息（见 cli/pdf2word.ts 的契约）
  const descr = JSON.stringify({ col: b.col, band: b.band, ...(b.furniture ? { furniture: true } : {}), lines: b.lines, size: b.size, pitch: Math.round(b.pitch * 100) / 100, align: b.align });
  return { x: b.x0, y: b.y0, w: b.x1 - b.x0, h: b.y1 - b.y0, para, wrap: true, name: `pdf2word-block ${page}.${no}`, descr };
}

const TINY = 1;
function tinyPara(extra: Partial<Para>): Para {
  return { kind: "p", runs: [], align: "left", indLeft: 0, indRight: 0, firstLine: 0, spaceBefore: 0, lineHeight: TINY, tabs: [], tiny: true, size: 1, ...extra };
}
function spacer(height: number): Para {
  return tinyPara({ lineHeight: Math.max(1, height) });
}
