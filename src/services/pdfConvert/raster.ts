/**
 * 需要"按原样渲染"的区域：矢量图、行间公式、首字下沉
 *
 * Word 无法用段落/表格还原这些内容（流程图的圆角框与箭头、TeX 公式的分式/根号/大括号……），
 * 逐字搬运只会得到错位的碎片。做法是：识别出区域 → 从文字流中移除 → 用 PDF 渲染出该区域的高清图片（透明背景），
 * 以页面绝对坐标贴回原位；区域内的文字保存在图片的替代文字里，并计入完整性报告。
 */

import { encodePng, type ImageBox, type MaskBox, type PageModel, type PathBox, type Span } from "./extract.js";
import { measureWord } from "./fonts.js";
import type { Line } from "./layout.js";

export interface Box { x0: number; y0: number; x1: number; y1: number }

const PAD = 1;
const CAPTION = /^\s*(Fig\.?|Figure|FIG\.?|TABLE|Table|Algorithm|图|表)\s*[\dIVX]/;

const centerIn = (s: Span, r: Box, tol = 0.5) => {
  const cx = (s.x0 + s.x1) / 2, cy = (s.top + s.bottom) / 2;
  return cx >= r.x0 - tol && cx <= r.x1 + tol && cy >= r.y0 - tol && cy <= r.y1 + tol;
};
const inside = (b: Box, r: Box, tol = 1) => b.x0 >= r.x0 - tol && b.x1 <= r.x1 + tol && b.y0 >= r.y0 - tol && b.y1 <= r.y1 + tol;
const near = (a: Box, b: Box, gap: number) => a.x0 <= b.x1 + gap && b.x0 <= a.x1 + gap && a.y0 <= b.y1 + gap && b.y0 <= a.y1 + gap;
const union = (a: Box, b: Box): Box => ({ x0: Math.min(a.x0, b.x0), y0: Math.min(a.y0, b.y0), x1: Math.max(a.x1, b.x1), y1: Math.max(a.y1, b.y1) });
const area = (b: Box) => Math.max(0, b.x1 - b.x0) * Math.max(0, b.y1 - b.y0);
const overlapArea = (a: Box, b: Box) => Math.max(0, Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0)) * Math.max(0, Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0));
const spanBox = (s: Span): Box => ({ x0: s.x0, y0: s.top, x1: s.x1, y1: s.bottom });
const isGray = (c: string) => {
  const n = parseInt(c, 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  return Math.max(r, g, b) - Math.min(r, g, b) < 12;
};

/** 按阅读顺序拼出一组文字片段（替代文字） */
export function spansText(spans: Span[]): string {
  const sorted = [...spans].sort((a, b) => a.baseline - b.baseline || a.x0 - b.x0);
  let out = "", lastBase = -Infinity, lastX = 0;
  for (const s of sorted) {
    if (out && Math.abs(s.baseline - lastBase) > 0.5 * s.size) out += "\n";
    else if (out && s.x0 - lastX > 0.15 * s.size) out += " ";
    out += s.text;
    lastBase = s.baseline; lastX = s.x1;
  }
  return out.replace(/[ \t]+/g, " ").trim();
}

function rasterBox(b: Box, page: number, kind: NonNullable<ImageBox["render"]>["kind"], alt: string): ImageBox {
  return { x0: b.x0, y0: b.y0, x1: b.x1, y1: b.y1, png: Buffer.alloc(0), pxW: 0, pxH: 0, render: { page, kind }, alt };
}

/** 从页面中移除落在区域内的全部元素（文字、线段、色块、位图、路径） */
export function removeInside(pg: PageModel, r: Box): Span[] {
  const gone = pg.spans.filter((s) => centerIn(s, r));
  const g = new Set(gone);
  pg.spans = pg.spans.filter((s) => !g.has(s));
  pg.segs = pg.segs.filter((s) => !inside(s, r));
  pg.fills = pg.fills.filter((f) => !inside(f, r));
  pg.images = pg.images.filter((i) => i.render || !inside(i, r, 2));
  pg.masks = (pg.masks ?? []).filter((m) => !inside(m, r, 1));
  const keep = pg.paths.filter((p) => !inside(p, r));
  const lostCurvy = pg.paths.filter((p) => inside(p, r) && (p.curved || p.diagonal) && (p.x1 - p.x0) * (p.y1 - p.y0) > 4).length;
  pg.vectorShapes = Math.max(0, pg.vectorShapes - lostCurvy);
  pg.paths = keep;
  return gone;
}

// ---------------------------------------------------------------------------
// 矢量图
// ---------------------------------------------------------------------------

/**
 * 矢量图识别：把已绘制的路径与位图按邻近关系聚类；含曲线/斜线/彩色填充的簇（流程图、曲线图、示意图）即为图。
 * 表格（只有横平竖直的线、落在已识别的表格范围内）不算；包含大量正文字号文字的簇（彩色文字框）不算。
 * 图内的文字、以及紧贴图边的小字号标注一并归入图。
 */
export function detectFigures(pg: PageModel, bodySize: number, tableBoxes: Box[]): ImageBox[] {
  type El = { box: Box; curvy: number; colored: number; paths: number; image: boolean };
  const els: El[] = [];
  const pageArea = pg.width * pg.height;
  for (const p of pg.paths) {
    // 贯穿整页的细线（页眉线、栏线）不参与聚类
    // （横跨正文的表格横线要保留：它把表格的各列连成一簇）
    const thinH = p.y1 - p.y0 < 3, thinV = p.x1 - p.x0 < 3;
    const inMargin = p.y1 < 0.08 * pg.height || p.y0 > 0.92 * pg.height;
    if ((thinH && (p.x1 - p.x0 > 0.8 * pg.width || (inMargin && p.x1 - p.x0 > 0.4 * pg.width))) || (thinV && p.y1 - p.y0 > 0.6 * pg.height)) continue;
    const curvy = p.curved || p.diagonal ? 1 : 0;
    const colored = p.filled && !isGray(p.fillColor) ? 1 : (p.stroked && !isGray(p.strokeColor) ? 1 : 0);
    els.push({ box: p, curvy, colored, paths: 1, image: false });
  }
  for (const im of pg.images) if (!im.render) els.push({ box: im, curvy: 0, colored: 0, paths: 0, image: true });
  if (!els.length) return [];

  // 聚类：包围盒相距 ≤ 6pt 的元素归为一簇（反复合并直至稳定）
  let clusters: Array<{ box: Box; els: El[] }> = els.map((e) => ({ box: { x0: e.box.x0, y0: e.box.y0, x1: e.box.x1, y1: e.box.y1 }, els: [e] }));
  const GAP = 6;
  for (let changed = true; changed; ) {
    changed = false;
    clusters.sort((a, b) => a.box.y0 - b.box.y0);
    const out: typeof clusters = [];
    for (const c of clusters) {
      const hit = out.find((o) => near(o.box, c.box, GAP));
      if (hit) { hit.box = union(hit.box, c.box); hit.els.push(...c.els); changed = true; }
      else out.push(c);
    }
    clusters = out;
  }

  const figs: Box[] = [];
  // 文字被转成矢量轮廓的表格（出版社处理过的 PDF 常见）：几条等宽的横线之间密布着小的填充曲线（字形），
  // 却几乎没有真正的文字 → 整张表作为一幅图渲染，否则表格内容会整个丢失
  {
    const rules = pg.paths.filter((p) => p.y1 - p.y0 < 2.5 && p.x1 - p.x0 > 40).sort((a, b) => a.y0 - b.y0);
    const used = new Set<number>();
    for (let i = 0; i < rules.length; i++) {
      if (used.has(i)) continue;
      const grp = [i];
      for (let j = i + 1; j < rules.length; j++) {
        const a = rules[grp[grp.length - 1]], b = rules[j];
        if (!(Math.abs(b.x0 - rules[i].x0) < 4 && Math.abs(b.x1 - rules[i].x1) < 4)) continue;
        // 两条横线之间有真正的文字（下一张表的表题等）：不是同一张表
        const between = pg.spans.filter((sp) => { const cy = (sp.top + sp.bottom) / 2, cx = (sp.x0 + sp.x1) / 2; return cy > a.y1 && cy < b.y0 && cx > a.x0 && cx < a.x1; })
          .reduce((n, sp) => n + sp.text.replace(/\s/g, "").length, 0);
        if (between > 3 || b.y0 - a.y0 > 0.4 * pg.height) break;
        grp.push(j);
      }
      if (grp.length < 2) continue;
      const top = rules[grp[0]], bot = rules[grp[grp.length - 1]];
      const box: Box = { x0: Math.min(top.x0, bot.x0), y0: top.y0, x1: Math.max(top.x1, bot.x1), y1: bot.y1 };
      if (box.y1 - box.y0 < 12) continue;
      const glyphs = pg.paths.filter((p) => p.filled && (p.curved || p.diagonal) && p.x1 - p.x0 < 20 && p.y1 - p.y0 < 20 && inside(p, box, 2)).length;
      const realText = pg.spans.filter((sp) => centerIn(sp, box)).reduce((n, sp) => n + sp.text.replace(/\s/g, "").length, 0);
      if (glyphs >= 20 && realText < 0.1 * glyphs) {
        grp.forEach((k) => used.add(k));
        figs.push({ x0: box.x0 - 1, y0: box.y0 - 1, x1: box.x1 + 1, y1: box.y1 + 1 });
      }
    }
  }
  for (const c of clusters) {
    if (figs.some((f) => inside(c.box, f, 2))) continue;
    const curvy = c.els.reduce((a, e) => a + e.curvy, 0);
    const colored = c.els.reduce((a, e) => a + e.colored, 0);
    const images = c.els.filter((e) => e.image);
    const paths = c.els.reduce((a, e) => a + e.paths, 0);
    const w = c.box.x1 - c.box.x0, h = c.box.y1 - c.box.y0;
    if (w < 24 || h < 14 || area(c.box) > 0.85 * pageArea) continue;
    // 位图只有叠加了矢量标注（箭头、框线）时才整体渲染；单独的位图按原图嵌入即可
    const annotatedImage = images.some((im) => c.els.filter((e) => !e.image && overlapArea(e.box, im.box) > 0).length >= 3);
    const seed = curvy >= 2 || (curvy >= 1 && paths >= 3) || colored >= 2 || annotatedImage;
    if (!seed) continue;
    // 与已识别的表格大面积重合：只有曲线/斜线足够多（流程图被误识别为表格）时才算图
    const tbl = tableBoxes.reduce((a, t) => a + overlapArea(t, c.box), 0);
    if (tbl > 0.6 * area(c.box) && curvy < 4) continue;
    figs.push({ ...c.box });
  }
  if (!figs.length) return [];

  // 合并相互接近的图区域，吸收图内文字与紧贴图边的小字号标注
  const out: ImageBox[] = [];
  const merged: Box[] = [];
  // 同一行并排的子图（曲线图的多个面板）：纵向范围大体重合、横向间隔不大 → 合成一幅图
  const sameRow = (a: Box, b: Box) => {
    const ov = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
    const gap = Math.max(a.x0, b.x0) - Math.min(a.x1, b.x1);
    const minH = Math.min(a.y1 - a.y0, b.y1 - b.y0);
    // 子图面板顶端、底端基本齐平；只是部分重叠的（左右两栏各一张表）不算
    const aligned = Math.abs(a.y0 - b.y0) < 0.2 * minH + 4 && Math.abs(a.y1 - b.y1) < 0.2 * minH + 4;
    return ov > 0.5 * minH && gap < 0.1 * pg.width && aligned;
  };
  // 合并后的范围里出现了两者之外的正文（例如左右两栏各有一张表，中间是正文）：不是同一幅图
  const textBetween = (a: Box, b: Box) => {
    const u = union(a, b);
    return pg.spans.some((s) => s.size >= 0.9 * bodySize && s.text.trim().length > 1 && centerIn(s, u) && !centerIn(s, a, 2) && !centerIn(s, b, 2));
  };
  // 各自带有题注（TABLE 5 / TABLE 6 分别在左右两栏）：是两个独立的图表
  const ownCaption = (b: Box) => pg.spans.some((s) => CAPTION.test(s.text) && s.x0 < b.x1 && s.x1 > b.x0 && s.x0 >= b.x0 - 20 && (Math.abs(s.bottom - b.y0) < 40 || Math.abs(s.top - b.y1) < 40));
  const mergeAll = (boxes: Box[]) => {
    for (let changed = true; changed; ) {
      changed = false;
      for (let i = 0; i < boxes.length && !changed; i++) for (let j = i + 1; j < boxes.length; j++) {
        if (near(boxes[i], boxes[j], 4) || (sameRow(boxes[i], boxes[j]) && !textBetween(boxes[i], boxes[j]) && !(ownCaption(boxes[i]) && ownCaption(boxes[j])))) { boxes[i] = union(boxes[i], boxes[j]); boxes.splice(j, 1); changed = true; break; }
      }
    }
    return boxes;
  };
  merged.push(...mergeAll(figs.sort((a, b) => a.y0 - b.y0).map((f) => ({ ...f }))));
  for (const r0 of merged) {
    let r = { ...r0 };
    for (let iter = 0; iter < 4; iter++) {
      let grown = false;
      for (const s of pg.spans) {
        if (centerIn(s, r, 0.5)) { const u = union(r, spanBox(s)); if (area(u) > area(r) + 0.01) { r = u; grown = true; } continue; }
        // 图边的小字标注：字号明显小于正文、紧贴图框、不是题注
        if (s.size < 0.9 * bodySize && near(spanBox(s), r, 6) && !CAPTION.test(s.text)) {
          r = union(r, spanBox(s)); grown = true;
        }
      }
      if (!grown) break;
    }
    const inner = pg.spans.filter((s) => centerIn(s, r));
    // 正文字号的文字过多：这是带底色/边框的文字块，而不是图
    const bodyChars = inner.filter((s) => s.size >= 0.95 * bodySize).reduce((a, s) => a + s.text.trim().length, 0);
    if (bodyChars > 160) continue;
    const box = { x0: r.x0 - PAD, y0: r.y0 - PAD, x1: r.x1 + PAD, y1: r.y1 + PAD };
    const alt = spansText(inner);
    removeInside(pg, box);
    out.push(rasterBox(box, pg.index, "figure", alt));
  }
  return out;
}

// ---------------------------------------------------------------------------
// 行间公式、首字下沉
// ---------------------------------------------------------------------------

export interface MathCtx {
  /** 正文字体的原始族名（如 NimbusRomNo9L、CMR）：与之不同的字体视为公式字体 */
  bodyBase: string;
  bodyIsTeX: boolean;
  bodySize: number;
}

export const rawBase = (s: Span) => (s.rawFont ?? s.font).split(/[-,+]/)[0].replace(/\d+$/, "").toUpperCase();
const CJK_RE = /[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef\u3000-\u303f]/;
const TEX_BASE = /^(CM|LM|MSAM|MSBM|EU[FRSX]|RSFS|STMARY|WASY|SF[A-Z]{2}\d|TC[A-Z]{2}\d|TXMI|TXSY|TXEX|PXMI|PXSY|PXEX|STIX|XITSMATH|CAMBRIAMATH|LATINMODERNMATH|ASANAMATH|R?MT(MI|SY|EX|MS|MB|EXT|SYN|SYB|MIB)|BLEX|BMTEX|EUCLID|MTPRO|MATHTIME)/;
/** 关系 / 运算符号：一行公式的标志 */
const REL = /[=≈≃≅≡≠≤≥≪≫∈∉⊂⊆←→↦∝]|:=/;
/** 公式中常见的运算符名称（不是英文正文） */
const OPNAME = new Set(["softmax", "concat", "argmax", "argmin", "sigmoid", "tanh", "relu", "mean", "clamp", "clip", "exp", "log", "max", "min", "sum", "where", "and", "for", "with", "head"]);
const EQ_NUM = /\(\s*\d{1,3}[a-z]?\s*\)\s*$/;
const isBigOp = (s: Span) => /^(CMEX|LMMATHEXTENSION|EUEX)/.test(rawBase(s)) || (/[∑∏∫∮⋃⋂]/.test(s.text) && !!s.math);

function lineInfo(l: Line, r: { x0: number; x1: number }, ctx: MathCtx) {
  let prose = "", math = 0;
  for (const s of l.spans) {
    // 公式字体：数学字体，或正文不是 TeX 字体时出现的 TeX 字体（CMR 的数字、运算符名等）；含中日韩文字的片段一律视为正文
    const nonProse = !CJK_RE.test(s.text) && (s.math || isBigOp(s) || (!ctx.bodyIsTeX && TEX_BASE.test(rawBase(s))));
    if (nonProse) math += s.text.replace(/\s/g, "").length;
    else prose += (prose && !/\s$/.test(prose) ? " " : "") + s.text; // 片段之间加空格：下标 i、变量 F 不被拼成"单词"
  }
  const text = l.spans.map((s) => s.text).join("").trim();
  const numbered = EQ_NUM.test(text) && r.x1 - l.x1 < 4;
  const proseCore = prose.replace(EQ_NUM, "").replace(/[\s.,;:()[\]{}=+\-−–·×/|]/g, "");
  const words = (prose.match(/[A-Za-z]{4,}/g) ?? []).length;
  // 英文正文单词：小写开头、4 个字母以上、不是运算符名称（ReLU、Conv、Upscale 等大写开头的函数名不算）
  const engWords = (prose.match(/(?<![A-Za-z])[a-z]{4,}(?![A-Za-z])/g) ?? []).filter((w) => !OPNAME.has(w)).length;
  // 公式里的变量：斜体的单个字母 / 短下标（正文斜体字体排的变量，如 Times Italic 的 F、rs）
  const vars = l.spans.filter((s) => s.italic && /^[A-Za-z]{1,3}$/.test(s.text.trim())).length;
  const rel = REL.test(text);
  const total = math + proseCore.length;
  const onlyNumber = /^\(\s*\d{1,3}[a-z]?\s*\)$/.test(text);
  const w = r.x1 - r.x0;
  const indented = l.x0 - r.x0 > 0.06 * w;
  const strong = !onlyNumber && total > 0 && (
    (math >= 0.6 * total && (words === 0 || (words <= 1 && indented)))
    // 关系符号 + 缩进 + 没有英文正文单词 + 有变量或编号：MathTime / 正文字体排的公式（变量用正文斜体）
    || (rel && (indented || numbered || l.x0 - r.x0 > 0.04 * w) && engWords === 0 && (vars >= 1 || numbered || math >= 2) && !/[.;,]\s+[A-Z][a-z]{3,}/.test(text))
  );
  // 分式的分子/分母、上下标等碎片：很短、没有单词、缩进
  const fragment = !strong && !onlyNumber && total <= 12 && words === 0 && indented;
  const big = l.spans.some(isBigOp);
  return { strong, fragment, onlyNumber, numbered, indented, big, text };
}

/**
 * 在一个区域（某一栏/整页版心）内识别行间公式：连续的"公式行"（公式字体占多数、几乎没有单词），
 * 且整组缩进居中、或右端带 (n) 编号、或含大型运算符/定界符。
 * 返回每个公式块的行与渲染区域（纵向不越过相邻的正文行）。
 */
export function detectEquations(lines: Line[], r: { x0: number; x1: number }, ctx: MathCtx, extra: Box[]): Array<{ lines: Line[]; box: Box }> {
  const ls = [...lines].sort((a, b) => a.top - b.top || a.x0 - b.x0);
  const info = ls.map((l) => lineInfo(l, r, ctx));
  const blocks: Array<{ lines: Line[]; box: Box }> = [];
  for (let i = 0; i < ls.length; ) {
    if (!(info[i].strong || info[i].fragment || info[i].onlyNumber)) { i++; continue; }
    let j = i + 1;
    while (j < ls.length && (info[j].strong || info[j].fragment || info[j].onlyNumber) && ls[j].top - Math.max(...ls.slice(i, j).map((l) => l.bottom)) < 1.2 * ls[j].size) j++;
    const bl = ls.slice(i, j), bi = info.slice(i, j);
    const ok = bi.some((x) => x.strong) && bi.some((x) => x.numbered || x.onlyNumber || x.big || (x.strong && x.indented));
    if (ok) {
      let box: Box = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
      for (const l of bl) for (const s of l.spans) {
        box = union(box, spanBox(s));
        // 大型定界符/运算符的字形远超字号给出的上下界
        if (isBigOp(s)) box = union(box, { x0: s.x0, y0: s.top, x1: s.x1, y1: s.baseline + 1.4 * s.size });
      }
      // 分数线、根号横线等：落在公式范围内的线段/色块/路径
      for (const e of extra) {
        if (e.x0 >= box.x0 - 3 && e.x1 <= box.x1 + 3 && e.y0 >= box.y0 - 4 && e.y1 <= box.y1 + 4) box = union(box, e);
      }
      // 不越过相邻的正文行
      const prev = ls.slice(0, i).reverse().find((l) => l.x1 > box.x0 && l.x0 < box.x1);
      const next = ls.slice(j).find((l) => l.x1 > box.x0 && l.x0 < box.x1);
      box.y0 = Math.max(box.y0 - 1.5, prev ? prev.baseline + 0.28 * prev.size : -Infinity);
      box.y1 = Math.min(box.y1 + 1.5, next ? next.baseline - 0.8 * next.size : Infinity);
      box.x0 = Math.max(box.x0 - 1.5, r.x0 - 3);
      box.x1 = Math.min(box.x1 + 1.5, r.x1 + 3);
      if (box.y1 - box.y0 > 4) blocks.push({ lines: bl, box });
    }
    i = j;
  }
  return blocks;
}

/**
 * 首字下沉：只含 1–2 个字符、字号 ≥ 正文 1.8 倍、右侧并排有 ≥2 行正文的行。
 * 首字母的基线常与第二行正文对齐，会被并进同一行（"T rapid disintegration…"）：这种行拆成首字母 + 其余正文（rest）。
 */
export function detectDropCaps(lines: Line[], ctx: MathCtx): Array<{ line: Line; box: Box; rest?: Line }> {
  const out: Array<{ line: Line; box: Box; rest?: Line }> = [];
  for (const l of lines) {
    let cap = l, rest: Line | undefined;
    const t = l.spans.map((s) => s.text).join("").trim();
    if (!t || t.length > 2 || l.size < 1.8 * ctx.bodySize) {
      const s0 = [...l.spans].sort((a, b) => a.x0 - b.x0)[0];
      const t0 = s0?.text.trim() ?? "";
      if (!s0 || s0.img || !t0 || t0.length > 2 || s0.size < 1.8 * ctx.bodySize || l.spans.length < 2) continue;
      const others = l.spans.filter((s) => s !== s0);
      if (!others.every((s) => s.size < 0.6 * s0.size && s.x0 >= s0.x1 - 1)) continue;
      cap = lineOf([s0]);
      rest = lineOf(others);
    }
    const h = cap.bottom - cap.top;
    const beside = lines.filter((o) => o !== l && o.x0 >= cap.x1 - 1 && o.x0 - cap.x1 < 2 * cap.size && o.top < cap.bottom && o.bottom > cap.top + 0.15 * h);
    if (beside.length + (rest ? 1 : 0) < 2) continue;
    out.push({ line: l, box: { x0: cap.x0 - 1, y0: cap.top - 0.5, x1: cap.x1 + 1, y1: cap.baseline + 0.1 * cap.size }, rest });
  }
  return out;
}

/** 由几段文字组成一行（与版面模块的行结构一致；字号取字数最多的） */
function lineOf(spans: Span[]): Line {
  const ss = [...spans].sort((a, b) => a.x0 - b.x0);
  const bySize = new Map<number, number>();
  for (const s of ss) bySize.set(s.size, (bySize.get(s.size) ?? 0) + Math.max(1, s.text.trim().length));
  const size = [...bySize.entries()].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0][0];
  const main = ss.filter((s) => Math.abs(s.size - size) < 0.6);
  const bl = main.map((s) => s.baseline).sort((a, b) => a - b);
  return {
    spans: ss, x0: Math.min(...ss.map((s) => s.x0)), x1: Math.max(...ss.map((s) => s.x1)),
    top: Math.min(...main.map((s) => s.top)), bottom: Math.max(...main.map((s) => s.bottom)),
    baseline: bl[Math.floor(bl.length / 2)], size,
  };
}

/**
 * 多张照片排成的组图（子图 (a)(b)… 标注在图下）：整体渲染成一张图，保持原来的行列排布；
 * 否则在分栏页面里各张照片会被拆到左右两栏、各占一段。
 */
export function detectImageGrids(pg: PageModel, bodySize: number): ImageBox[] {
  const ims = pg.images.filter((i) => !i.render && (i.x1 - i.x0) * (i.y1 - i.y0) > 400);
  if (ims.length < 2) return [];
  let groups: Box[][] = ims.map((i) => [{ x0: i.x0, y0: i.y0, x1: i.x1, y1: i.y1 }]);
  const bb = (g: Box[]) => g.reduce((a, b) => union(a, b));
  for (let changed = true; changed; ) {
    changed = false;
    outer: for (let i = 0; i < groups.length; i++) for (let j = i + 1; j < groups.length; j++) {
      const A = bb(groups[i]), B = bb(groups[j]);
      const row = Math.min(A.y1, B.y1) - Math.max(A.y0, B.y0) > 0.5 * Math.min(A.y1 - A.y0, B.y1 - B.y0) && Math.max(A.x0, B.x0) - Math.min(A.x1, B.x1) < 0.15 * pg.width;
      if (near(A, B, 3 * bodySize) || row) { groups[i].push(...groups[j]); groups.splice(j, 1); changed = true; break outer; }
    }
  }
  const LABEL = /^\(?[a-zA-Z0-9]{1,3}[).]?$/;
  const out: ImageBox[] = [];
  for (const g of groups) {
    if (g.length < 2) continue;
    const r0 = bb(g);
    let r = r0;
    // 子图编号：图内或紧贴图下方的短标注。
    // 紧贴图下方的标注以原图边界为准（不随扩展逐行级联），并且那一行里只能有编号：
    // 图题第一行常以 "(a) …" 开头，若把它当成子图编号，范围会一路扩进图题，整组图被"夹着正文"否决
    const labelLine = (s: Span) => pg.spans.every((o) => !o.text.trim() || LABEL.test(o.text.trim()) || Math.abs((o.top + o.bottom) / 2 - (s.top + s.bottom) / 2) > 0.4 * (s.bottom - s.top + 1));
    for (const s of pg.spans) {
      const t = s.text.trim();
      if (!t) continue;
      if (centerIn(s, r, 0.5) || (LABEL.test(t) && s.x0 >= r0.x0 - 2 && s.x1 <= r0.x1 + 2 && s.top >= r0.y1 - 2 && s.top - r0.y1 < 1.6 * bodySize && labelLine(s))) r = union(r, spanBox(s));
    }
    const inner = pg.spans.filter((s) => centerIn(s, r));
    if (inner.filter((s) => !LABEL.test(s.text.trim())).reduce((a, s) => a + s.text.trim().length, 0) > 80) continue; // 照片之间夹着正文：不是组图
    const box = { x0: r.x0 - PAD, y0: r.y0 - PAD, x1: r.x1 + PAD, y1: r.y1 + PAD };
    const alt = spansText(inner);
    removeInside(pg, box);
    out.push(rasterBox(box, pg.index, "figure", alt));
  }
  return out;
}

export function pathBoxes(pg: PageModel): Box[] {
  return [...pg.segs, ...pg.fills, ...pg.paths.filter((p: PathBox) => !p.curved)];
}

export { rasterBox };

// ---------------------------------------------------------------------------
// 没有文字编码的字形墨迹（Type 3 位图字体）
// ---------------------------------------------------------------------------

/**
 * 把落在 box 内的图像蒙版合成为透明背景的 PNG（纯 JS，不依赖画布；2×2 超采样抗锯齿）。
 * 画布不可用或渲染失败时，公式与符号仍能以图片保留下来。
 */
export function compositeMasks(masks: MaskBox[], box: Box, scale = 4): { png: Buffer; w: number; h: number } | null {
  const W = Math.max(1, Math.ceil((box.x1 - box.x0) * scale)), H = Math.max(1, Math.ceil((box.y1 - box.y0) * scale));
  if (W * H > 25e6) return null;
  const rgba = new Uint8Array(W * H * 4);
  const SS = [0.25, 0.75];
  let any = false;
  for (const m of masks) {
    if (m.vector || m.x1 <= box.x0 || m.x0 >= box.x1 || m.y1 <= box.y0 || m.y0 >= box.y1) continue;
    const col = /^[0-9a-f]{6}$/i.test(m.color) ? m.color : "000000";
    const r = parseInt(col.slice(0, 2), 16), g = parseInt(col.slice(2, 4), 16), b = parseInt(col.slice(4, 6), 16);
    const i0 = Math.max(0, Math.floor((m.x0 - box.x0) * scale)), i1 = Math.min(W, Math.ceil((m.x1 - box.x0) * scale));
    const j0 = Math.max(0, Math.floor((m.y0 - box.y0) * scale)), j1 = Math.min(H, Math.ceil((m.y1 - box.y0) * scale));
    const rb = Math.ceil(m.w / 8);
    const { ox, oy, a, b: bb, c, d } = m.inv;
    for (let j = j0; j < j1; j++) for (let i = i0; i < i1; i++) {
      let hit = 0;
      for (const sy of SS) for (const sx of SS) {
        const dx = box.x0 + (i + sx) / scale - ox, dy = box.y0 + (j + sy) / scale - oy;
        const u = a * dx + bb * dy, v = c * dx + d * dy;
        if (u < 0 || u >= 1 || v < 0 || v >= 1) continue;
        if (!m.bits) { hit++; continue; }
        const px = Math.min(m.w - 1, Math.floor(u * m.w)), py = Math.min(m.h - 1, Math.floor((1 - v) * m.h));
        if (!(m.bits[py * rb + (px >> 3)] & (0x80 >> (px & 7)))) hit++;
      }
      if (!hit) continue;
      const k = (j * W + i) * 4, alpha = Math.round((255 * hit) / 4);
      if (alpha > rgba[k + 3]) { rgba[k] = r; rgba[k + 1] = g; rgba[k + 2] = b; rgba[k + 3] = alpha; any = true; }
    }
  }
  if (!any) return null;
  return { png: encodePng(W, H, rgba, 4), w: W, h: H };
}

export interface MaskInk {
  /** 嵌在正文行里的公式 / 符号：带行内图片的占位文字段（text 为 U+FFFC），随文字排版 */
  inline: Span[];
  /** 独占一块的公式（连同其中的 "if"、编号等短文字）：独立成段的图片 */
  display: ImageBox[];
  /** 并进公式图片的文字（计入完整性校验） */
  alt: string[];
}

/**
 * 没有文字编码的字形（Type 3 位图字体画的公式、符号）→ 图片。
 *   1. 字形按位置聚成墨迹块（同一行内相邻、或上下紧贴——分式、上下标、重音）；
 *   2. 墨迹块所在的那一行（同一栏、连续的一段）里有足够多的正文文字 → 行内符号，作为行内图片插回正文行；
 *   3. 其余是独立成块的公式：上下相邻的合成一块，连同其中的短文字（"if"、"otherwise"、编号）整体成图。
 * raster 为真时独立公式用画布按原样渲染（含其中的文字）；否则由蒙版直接合成（其中的文字留作正文）。
 */
export function detectMaskInk(pg: PageModel, bodySize: number, raster: boolean): MaskInk {
  const out: MaskInk = { inline: [], display: [], alt: [] };
  const masks = (pg.masks ?? []).filter((m) => m.x1 > 0 && m.x0 < pg.width && m.y1 > 0 && m.y0 < pg.height);
  const B = bodySize || 10;
  // Type 3 字体的文字：字形由 PDF 自己描述（多为公式符号），字符编码常是乱码（\u0000、\u0012），Word 里也没有这个字体。
  // 这页只有少量 Type 3 文字（公式）时，全部当作墨迹按原样渲染；正文本身也用 Type 3 时，编码正常的保留为文字，只把乱码的当作墨迹。
  // 没有画布时无法渲染：保留为文字，换成正文字体并去掉乱码
  const t3 = pg.spans.filter((sp) => sp.type3 && sp.text.trim());
  if (t3.length) {
    const chars = (l: Span[]) => l.reduce((n, sp) => n + sp.text.replace(/\s/g, "").length, 0);
    const share = chars(t3) / Math.max(1, chars(pg.spans.filter((sp) => sp.text.trim())));
    const junk = (sp: Span) => /[\u0000-\u001f\u007f-\u009f\ue000-\uf8ff\ufffd]/.test(sp.text);
    const body = pg.spans.filter((sp) => !sp.type3 && sp.text.trim());
    const bodyFont = body.length ? body.reduce((a, b) => (b.text.length > a.text.length ? b : a)).font : "Times New Roman";
    const asInk = raster ? t3.filter((sp) => share < 0.3 || junk(sp)) : [];
    const gone = new Set(asInk);
    for (const sp of asInk) {
      masks.push({ x0: sp.x0, y0: sp.top, x1: sp.x1, y1: sp.bottom, color: sp.color, bits: null, w: 1, h: 1, vector: true, inv: { ox: 0, oy: 0, a: 0, b: 0, c: 0, d: 0 } });
      out.alt.push(sp.text.replace(/[\u0000-\u001f]/g, ""));
    }
    pg.spans = pg.spans.filter((sp) => !gone.has(sp)).map((sp) => {
      if (!sp.type3) return sp;
      const t = sp.text.replace(/[\u0000-\u001f\u007f-\u009f\ue000-\uf8ff\ufffd]/g, "");
      return { ...sp, text: t, font: bodyFont, type3: undefined };
    }).filter((sp) => sp.text.length > 0);
  }
  if (!masks.length) return out;
  // 紧贴字形的短横线（分数线、向量下划线、上横线）是公式的一部分：并入墨迹，不再当作正文的下划线 / 删除线
  const solid = (x0: number, y0: number, x1: number, y1: number, color: string): MaskBox => ({
    x0, y0, x1, y1, color, bits: null, w: 1, h: 1,
    inv: { ox: x0, oy: y1, a: 1 / Math.max(1e-6, x1 - x0), b: 0, c: 0, d: -1 / Math.max(1e-6, y1 - y0) },
  });
  const nearInk = (x0: number, y0: number, x1: number, y1: number) =>
    masks.some((m) => Math.min(m.x1, x1) - Math.max(m.x0, x0) > 0.3 * Math.min(m.x1 - m.x0, x1 - x0) && (Math.abs(m.y1 - y0) <= 3 || Math.abs(m.y0 - y1) <= 3 || (m.y0 < y1 && m.y1 > y0)));
  const isRule = (w: number, h: number) => h <= 1.5 && w >= 1.5 && w <= 8 * B;
  const ruleSegs = pg.segs.filter((g) => g.horizontal && isRule(g.x1 - g.x0, Math.max(g.y1 - g.y0, g.width)) && nearInk(g.x0, g.y0 - g.width / 2, g.x1, g.y1 + g.width / 2));
  const ruleFills = pg.fills.filter((f) => isRule(f.x1 - f.x0, f.y1 - f.y0) && nearInk(f.x0, f.y0, f.x1, f.y1));
  for (const g of ruleSegs) { const hw = Math.max(0.2, g.width / 2); masks.push(solid(g.x0, (g.y0 + g.y1) / 2 - hw, g.x1, (g.y0 + g.y1) / 2 + hw, g.color)); }
  for (const f of ruleFills) masks.push(solid(f.x0, f.y0, f.x1, f.y1, f.color));
  if (ruleSegs.length) { const gone = new Set(ruleSegs); pg.segs = pg.segs.filter((g) => !gone.has(g)); }
  if (ruleFills.length) { const gone = new Set(ruleFills); pg.fills = pg.fills.filter((f) => !gone.has(f)); }
  const composite = (box: Box, kind: "figure" | "equation") => {
    // 含 Type 3 矢量字形的区域只能由画布渲染（转换时统一渲染）；其余由蒙版直接合成
    if (raster && masks.some((m) => m.vector && m.x1 > box.x0 && m.x0 < box.x1 && m.y1 > box.y0 && m.y0 < box.y1)) return rasterBox(box, pg.index, kind, "");
    const c = compositeMasks(masks, box);
    return c ? { ...box, png: c.png, pxW: c.w, pxH: c.h, alt: "" } as ImageBox : null;
  };

  // 大块的模板图像（徽标等）：直接成图
  const big = masks.filter((m) => m.x1 - m.x0 > 6 * B && m.y1 - m.y0 > 6 * B);
  for (const m of big) { const im = composite(m, "figure"); if (im) out.display.push(im); }
  const small = masks.filter((m) => !big.includes(m));
  if (!small.length) return out;

  const text = pg.spans.filter((s) => !s.img && s.text.trim());
  // 分栏：中线两侧都有大量文字、几乎没有文字跨过中线时，按两栏之间的空白分开
  const xs0 = text.map((s) => s.x0), xs1 = text.map((s) => s.x1);
  const tx0 = Math.min(...xs0, pg.width), tx1 = Math.max(...xs1, 0);
  const midX = (tx0 + tx1) / 2;
  const leftS = text.filter((s) => s.x1 <= midX), rightS = text.filter((s) => s.x0 >= midX);
  const crossS = text.filter((s) => s.x0 < midX - 2 && s.x1 > midX + 2);
  const split = leftS.length >= 8 && rightS.length >= 8 && crossS.length < 0.1 * text.length
    ? (Math.max(...leftS.map((s) => s.x1)) + Math.min(...rightS.map((s) => s.x0))) / 2 : null;
  const colOf = (x0: number, x1: number, bl: number, size: number) => {
    if (split === null) return { x0: tx0, x1: tx1 };
    // 这一行有文字跨过分栏线（通栏的标题、摘要）：整行算一栏
    if (text.some((s) => Math.abs(s.baseline - bl) < 0.35 * size && s.x0 < split - 2 && s.x1 > split + 2)) return { x0: tx0, x1: tx1 };
    return (x0 + x1) / 2 < split ? { x0: tx0, x1: split } : { x0: split, x1: tx1 };
  };
  const sameCol = (a: { x0: number; x1: number }, b: { x0: number; x1: number }) =>
    split === null || ((a.x0 + a.x1) / 2 < split) === ((b.x0 + b.x1) / 2 < split) || a.x0 < split - 2 && a.x1 > split + 2;

  // 1. 聚类
  // 字形明确所在的正文行：字身带（基线上 0.7 字号到基线下 0.15 字号）覆盖了它 60% 以上的高度。
  // 上下相邻的两个字形各自明确属于不同的行时不能拼合（上一行的下标、下一行的上标）；
  // 求和号的上下限不在任何一行的字身带里，可以与求和号拼合
  // 同一栏里的正文行（不限横向距离：行内公式可能离同一行的文字很远）
  const textRows = text.map((s) => ({ bl: s.baseline, size: s.size, x0: s.x0, x1: s.x1 }));
  const firmRow = (m: MaskBox): number | null => {
    const h = m.y1 - m.y0;
    for (const r of textRows) {
      if (!sameCol(r, m)) continue;
      const ov = Math.min(r.bl + 0.15 * r.size, m.y1) - Math.max(r.bl - 0.7 * r.size, m.y0);
      if (ov >= 0.6 * h) return Math.round(r.bl * 2) / 2;
    }
    return null;
  };
  const firm = small.map(firmRow);
  // 离得最近的正文行（按字身带中心）：不在任何一行字身带里的字形（上下限、分式的分子分母）归属它
  const nearRow = (m: MaskBox): number | null => {
    const cy = (m.y0 + m.y1) / 2;
    let best: number | null = null, bd = Infinity;
    for (const r of textRows) {
      if (!sameCol(r, m)) continue;
      const d = Math.abs(r.bl - 0.275 * r.size - cy);
      if (d < bd && d < 1.2 * r.size) { bd = d; best = Math.round(r.bl * 2) / 2; }
    }
    return best;
  };
  // 不在任何字身带里的字形（上下限、重音、分子分母、分数线）跟随上下紧贴着它的字形；没有紧贴的才按最近的行
  const touching = (a: MaskBox, b: MaskBox) => {
    const hOv = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
    const vGap = Math.max(a.y0, b.y0) - Math.min(a.y1, b.y1);
    return hOv >= 0.3 * Math.min(a.x1 - a.x0, b.x1 - b.x0) && vGap <= 0.3 * B ? vGap : Infinity;
  };
  const home: Array<number | null> = small.map((_, i) => firm[i]);
  for (let pass = 0; pass < 4; pass++) {
    let changed = false;
    small.forEach((m, i) => {
      if (firm[i] !== null) return;
      let bestGap = Infinity, bestHome: number | null = null;
      small.forEach((o, j) => {
        if (j === i || home[j] === null) return;
        const g = touching(m, o);
        if (g < bestGap) { bestGap = g; bestHome = home[j]; }
      });
      if (bestHome !== null && bestHome !== home[i]) { home[i] = bestHome; changed = true; }
    });
    if (!changed) break;
  }
  small.forEach((m, i) => { if (home[i] === null) home[i] = nearRow(m); });
  const bigGlyph = (m: MaskBox) => m.y1 - m.y0 >= 0.95 * B; // 求和号、积分号、大括号：上下限总是并入
  const parent = small.map((_, i) => i);
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  const order = small.map((_, i) => i).sort((i, j) => small[i].x0 - small[j].x0);
  for (let oi = 0; oi < order.length; oi++) {
    const a = small[order[oi]];
    for (let oj = oi + 1; oj < order.length; oj++) {
      const b = small[order[oj]];
      if (b.x0 - a.x1 > 0.35 * B) break;
      const hGap = Math.max(a.x0, b.x0) - Math.min(a.x1, b.x1);
      const vOv = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
      const minH = Math.min(a.y1 - a.y0, b.y1 - b.y0);
      const sameRow = vOv > 0.25 * minH && hGap <= 0.35 * B;
      const minW = Math.min(a.x1 - a.x0, b.x1 - b.x0);
      const ha = home[order[oi]], hb = home[order[oj]];
      const sameHome = ha === null || hb === null || Math.abs(ha - hb) <= 1;
      const stacked = -hGap >= 0.3 * minW && -vOv <= 0.3 * B && (sameHome || ((bigGlyph(a) || bigGlyph(b)) && (firm[order[oi]] === null || firm[order[oj]] === null)));
      if (sameRow || stacked) parent[find(order[oi])] = find(order[oj]);
    }
  }
  const groups = new Map<number, Box>();
  small.forEach((m, i) => { const k = find(i); const g = groups.get(k); groups.set(k, g ? union(g, m) : { x0: m.x0, y0: m.y0, x1: m.x1, y1: m.y1 }); });
  const clusters = [...groups.values()];

  // 2. 行内 / 独立
  // 每个墨迹块所属的正文行：同一栏内，字身带（基线上 0.7 字号到基线下 0.15 字号）与它重叠最多、中心最近的那一行。
  // 带上下限的求和号会同时碰到上下两行，按中心距离取本行
  const rowKey = (bl: number) => Math.round(bl * 2) / 2;
  const assign = new Map<Box, { baseline: number; size: number; col: { x0: number; x1: number } }>();
  for (const c of clusters) {
    const cMid = (c.y0 + c.y1) / 2;
    let best: { baseline: number; size: number; score: number } | null = null;
    for (const s of text) {
      if (Math.abs(s.baseline - cMid) > 1.5 * Math.max(s.size, c.y1 - c.y0)) continue;
      const col = colOf(s.x0, s.x1, s.baseline, s.size);
      const cc = (c.x0 + c.x1) / 2;
      if (cc < col.x0 - 2 || cc > col.x1 + 2) continue;
      const band0 = s.baseline - 0.7 * s.size, band1 = s.baseline + 0.15 * s.size;
      const ov = Math.min(band1, c.y1) - Math.max(band0, c.y0);
      if (ov <= 0) continue;
      const score = ov - 0.5 * Math.abs((band0 + band1) / 2 - cMid);
      if (!best || score > best.score + 1e-6) best = { baseline: s.baseline, size: s.size, score };
    }
    if (best) assign.set(c, { baseline: best.baseline, size: best.size, col: colOf(c.x0, c.x1, best.baseline, best.size) });
  }
  // 每一行（同一栏）的文字与墨迹：字数、覆盖的宽度
  const rowStat = new Map<string, { chars: number; covered: number; spans: Span[]; colW: number; left: boolean; right: boolean; cont: boolean }>();
  const keyOf = (bl: number, col: { x0: number }) => `${rowKey(bl)}@${Math.round(col.x0)}`;
  for (const r of assign.values()) {
    const k = keyOf(r.baseline, r.col);
    if (rowStat.has(k)) continue;
    // 单独的公式编号 "(17)" 不算正文（它让公式行"排到了右边缘"）
    const spans = text.filter((s) => Math.abs(s.baseline - r.baseline) < 0.35 * r.size && (s.x0 + s.x1) / 2 >= r.col.x0 - 2 && (s.x0 + s.x1) / 2 <= r.col.x1 + 2 && !/^\(\s*\d{1,3}[a-z]?\s*\)$/.test(s.text.trim()));
    const inks = [...assign.entries()].filter(([, o]) => keyOf(o.baseline, o.col) === k).map(([b]) => b);
    const covered = spans.reduce((w, s) => w + (s.x1 - s.x0), 0) + inks.reduce((w, b) => w + (b.x1 - b.x0), 0);
    const all = [...spans, ...inks];
    const rx0 = Math.min(...all.map((b) => b.x0)), rx1 = Math.max(...all.map((b) => b.x1));
    rowStat.set(k, {
      chars: spans.reduce((n, s) => n + s.text.replace(/\s/g, "").length, 0), covered, spans, colW: r.col.x1 - r.col.x0,
      // 从栏的左边缘（容许段落首行缩进）排起 / 排到栏的右边缘
      left: rx0 <= r.col.x0 + 1.6 * B, right: rx1 >= r.col.x1 - 1.6 * B,
      // 段落的续行：上一行（同一栏、一个行距之内）是排满整栏的正文
      cont: (() => {
        const above = (bl: number) => r.baseline - bl > 0.5 * r.size && r.baseline - bl < 1.6 * r.size;
        const inCol = (b: { x0: number; x1: number }) => (b.x0 + b.x1) / 2 >= r.col.x0 - 2 && (b.x0 + b.x1) / 2 <= r.col.x1 + 2;
        // 上一行的文字与墨迹（行尾常常就是公式）
        const up: Array<{ x0: number; x1: number }> = [
          ...text.filter((t) => above(t.baseline) && inCol(t)),
          ...[...assign.entries()].filter(([b, o]) => above(o.baseline) && inCol(b)).map(([b]) => b),
        ];
        return up.length > 0 && Math.min(...up.map((t) => t.x0)) <= r.col.x0 + 1.6 * B && Math.max(...up.map((t) => t.x1)) >= r.col.x1 - 1.6 * B;
      })(),
    });
  }
  const inlineC: Array<{ c: Box; spans: Span[]; baseline: number; size: number }> = [];
  // 成段的正文：3 个词以上（公式里的文字只有 "if"、"otherwise"、"for all" 这类短语）
  const prose = (sp: Span) => !/^\(\s*\d{1,3}[a-z]?\s*\)$/.test(sp.text.trim()) && sp.text.trim().split(/\s+/).filter((w) => /[A-Za-z\u4e00-\u9fff]{2,}/.test(w)).length >= 3;
  const displayC: Box[] = [];
  for (const c of clusters) {
    const r = assign.get(c);
    const st = r ? rowStat.get(keyOf(r.baseline, r.col)) : undefined;
    // 正文行：有文字，并且从栏的左边缘排到右边缘（两端对齐的整行），或从左边缘起有一定字数（段落末行），或文字很多。
    // 独立公式的行缩进 / 居中、文字很少（"if"、"otherwise"、编号），不满足这些条件
    const body = !!st && st.chars >= 1 && ((st.left && st.right) || ((st.left || st.right) && st.chars >= 6) || (st.left && st.cont) || st.chars >= 20 || st.spans.some(prose));
    if (r && st && body && c.y1 - c.y0 <= 2.4 * B) inlineC.push({ c, spans: st.spans, baseline: r.baseline, size: r.size });
    else displayC.push(c);
  }

  // 行内：占位文字段 + 由蒙版合成的图片
  const makeInline = (c: Box, spans: Span[], baseline: number, size: number) => {
    const ref = spans.length ? spans.reduce((a, b) => (Math.min(Math.abs(a.x1 - c.x0), Math.abs(a.x0 - c.x1)) <= Math.min(Math.abs(b.x1 - c.x0), Math.abs(b.x0 - c.x1)) ? a : b)) : null;
    const box = { x0: c.x0 - 0.2, y0: c.y0 - 0.2, x1: c.x1 + 0.2, y1: c.y1 + 0.2 };
    const img = composite(box, "equation");
    if (!img) return;
    const w = box.x1 - box.x0;
    out.inline.push({
      text: "\uFFFC", x0: box.x0, x1: box.x1, top: baseline - 0.75 * size, bottom: baseline + 0.22 * size, baseline, size,
      font: ref?.font ?? "Times New Roman", bold: false, italic: false, color: "000000",
      natural: w, wordW: w, spaceW: ref?.spaceW ?? 0.25 * size, wordSpaceW: ref?.wordSpaceW, img,
    });
  };
  for (const { c, spans, baseline, size } of inlineC) makeInline(c, spans, baseline, size);

  // 一段文字把行内公式的位置也包在里面（PDF 用一个大的字间距跳过公式，文字提取时只留下一个空格）：
  // 在最接近公式位置的空格处把文字拆成两段，公式图片插在中间
  const splitAround = (sp: Span, c: Box): Span[] | null => {
    const t = sp.text;
    const width = (str: string) => measureWord(str, sp.font, sp.bold, sp.italic, sp.size) ?? str.length * 0.45 * sp.size;
    const total = width(t), actual = sp.x1 - sp.x0;
    if (total <= 0) return null;
    const k = Math.min(1.15, Math.max(0.85, (actual - (c.x1 - c.x0)) / total));
    // 候选：左半段排在公式之前、右半段排在公式之后（容许 1.5pt 误差）；
    // 其中优先右半段以标点开头的（"l̂." "t⁽ᵐ⁾," 这类行内公式后紧跟标点），其次左半段末尾最接近公式
    let best = -1, bd = Infinity;
    for (let i = 1; i < t.length - 1; i++) {
      if (t[i] !== " ") continue;
      const E = sp.x0 + width(t.slice(0, i)) * k, S = sp.x1 - width(t.slice(i + 1)) * k;
      const v = Math.max(0, E - c.x0) + Math.max(0, c.x1 - S);
      if (v > 1.5) continue;
      const d = Math.abs(E + 0.25 * sp.size - c.x0) - (/^[.,;:!?)\]]/.test(t.slice(i + 1)) ? 2 * B : 0);
      if (d < bd) { bd = d; best = i; }
    }
    if (best < 0) return null;
    const left = t.slice(0, best), right = t.slice(best + 1);
    if (!left.trim() || !right.trim()) return null;
    const wl = width(left), wr = width(right);
    const part = (str: string, x0: number, x1: number, w: number): Span => ({
      ...sp, text: str, x0, x1,
      natural: sp.natural !== undefined ? sp.natural * (w / total) : undefined,
      wordW: sp.wordW !== undefined ? sp.wordW * (w / total) : undefined,
    });
    return [part(left, sp.x0, Math.min(c.x0 - 0.1, sp.x0 + wl * k), wl), part(right, Math.max(c.x1 + 0.1, sp.x1 - wr * k), sp.x1, wr)];
  };
  for (const p of out.inline) {
    const host = pg.spans.find((t) => !t.img && Math.abs(t.baseline - p.baseline) <= 0.3 * p.size && t.x0 < p.x0 - 0.5 && t.x1 > p.x1 + 0.5);
    if (!host) continue;
    const parts = splitAround(host, p.img!);
    if (!parts) continue;
    const i = pg.spans.indexOf(host);
    pg.spans = [...pg.spans.slice(0, i), ...parts, ...pg.spans.slice(i + 1)];
    const j = text.indexOf(host);
    if (j >= 0) text.splice(j, 1, ...parts);
  }

  // 行内图片与同一行紧挨着的文字可能有横向重叠（波浪号伸到后面的逗号上方）：占位的范围收窄到不重叠，
  // 否则分行时两者会被当成不同的行。图片本身仍按原宽度排版（natural 不变）
  for (const p of out.inline) {
    const w = p.x1 - p.x0;
    for (const t of text) {
      if (Math.abs(t.baseline - p.baseline) > 0.3 * p.size) continue;
      const ov = Math.min(p.x1, t.x1) - Math.max(p.x0, t.x0);
      if (ov <= 0) continue;
      if ((t.x0 + t.x1) / 2 >= (p.x0 + p.x1) / 2) p.x1 = Math.max(p.x0 + 0.3 * w, t.x0 - 0.01);
      else p.x0 = Math.min(p.x1 - 0.3 * w, t.x1 + 0.01);
    }
  }

  // 3. 独立公式：上下相邻的墨迹块合成一块
  let blocks = displayC.map((c) => ({ box: { ...c }, members: [c] }));
  for (let changed = true; changed; ) {
    changed = false;
    outer: for (let i = 0; i < blocks.length; i++) for (let j = i + 1; j < blocks.length; j++) {
      const a = blocks[i].box, b = blocks[j].box;
      const hGap = Math.max(a.x0, b.x0) - Math.min(a.x1, b.x1);
      const vGap = Math.max(a.y0, b.y0) - Math.min(a.y1, b.y1);
      if (hGap <= 3 * B && vGap <= 0.9 * B && sameCol(a, b)) {
        blocks[i] = { box: union(a, b), members: [...blocks[i].members, ...blocks[j].members] };
        blocks.splice(j, 1); changed = true; break outer;
      }
    }
  }
  // 兜底：成图的范围会盖住成段的正文 → 判断有误（行内公式所在的行没认出来），退回行内图片，正文绝不并进图里
  const fallbackInline = (members: Box[], covers: Span[]) => {
    for (const c of members) {
      const cm = (c.y0 + c.y1) / 2;
      const ref = covers.reduce((a, b) => (Math.abs(a.baseline - 0.3 * a.size - cm) <= Math.abs(b.baseline - 0.3 * b.size - cm) ? a : b));
      makeInline(c, covers.filter((sp) => Math.abs(sp.baseline - ref.baseline) < 0.35 * ref.size), ref.baseline, ref.size);
    }
  };
  const proseIn = (r: Box) => text.filter((sp) => prose(sp) && (sp.x0 + sp.x1) / 2 > r.x0 && (sp.x0 + sp.x1) / 2 < r.x1 && (sp.top + sp.bottom) / 2 > r.y0 && (sp.top + sp.bottom) / 2 < r.y1);
  for (const { box: b0, members } of blocks) {
    const covers0 = proseIn(b0);
    if (covers0.length) { fallbackInline(members, covers0); continue; }
    let b = { ...b0 };
    if (raster) {
      // 所在栏（按分栏线；公式本身跨过分栏线时为整个版心）
      const crosses = split !== null && b.x0 < split - 2 && b.x1 > split + 2;
      const colX0 = split === null || crosses ? tx0 : (b.x0 + b.x1) / 2 < split ? tx0 : split;
      const colX1 = split === null || crosses ? tx1 : (b.x0 + b.x1) / 2 < split ? split : tx1;
      // 公式中的短文字（"if"、"otherwise"、"for all"、编号）一并成图；成段的文字（3 个词以上）不并入
      const absorbed = text.filter((s) => {
        const cy = (s.top + s.bottom) / 2;
        return cy >= b.y0 - 0.2 * B && cy <= b.y1 + 0.2 * B && s.x0 >= colX0 - 2 && s.x1 <= colX1 + 2 && !prose(s);
      });
      for (const s of absorbed) b = union(b, spanBox(s));
      // 上下留 1.5pt 边，但不越过相邻的正文行（否则会把上一行字母的下伸部分也截进图里）
      const others = text.filter((s) => !absorbed.includes(s) && s.x1 > b.x0 && s.x0 < b.x1);
      const above = Math.max(-Infinity, ...others.filter((s) => s.baseline <= b.y0 + 0.2 * s.size).map((s) => s.baseline + 0.28 * s.size));
      const below = Math.min(Infinity, ...others.filter((s) => s.baseline - 0.8 * s.size >= b.y1 - 0.2 * s.size).map((s) => s.baseline - 0.8 * s.size));
      const box = { x0: b.x0 - 1.5, y0: Math.max(b.y0 - 1.5, Math.min(b.y0, above)), x1: b.x1 + 1.5, y1: Math.min(b.y1 + 1.5, Math.max(b.y1, below)) };
      const covers = proseIn(box);
      if (covers.length) { fallbackInline(members, covers); continue; }
      const gone = removeInside(pg, box);
      const alt = spansText(gone);
      if (alt) out.alt.push(alt);
      out.display.push(rasterBox(box, pg.index, "equation", alt));
    } else {
      const im = composite({ x0: b.x0 - 1, y0: b.y0 - 1, x1: b.x1 + 1, y1: b.y1 + 1 }, "equation");
      if (im) out.display.push(im);
    }
  }
  return out;
}
