/**
 * 需要"按原样渲染"的区域：矢量图、行间公式、首字下沉
 *
 * Word 无法用段落/表格还原这些内容（流程图的圆角框与箭头、TeX 公式的分式/根号/大括号……），
 * 逐字搬运只会得到错位的碎片。做法是：识别出区域 → 从文字流中移除 → 用 PDF 渲染出该区域的高清图片（透明背景），
 * 以页面绝对坐标贴回原位；区域内的文字保存在图片的替代文字里，并计入完整性报告。
 */

import type { ImageBox, PageModel, PathBox, Span } from "./extract.js";
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

/** 首字下沉：只含 1–2 个字符、字号 ≥ 正文 1.8 倍、右侧并排有 ≥2 行正文的行 */
export function detectDropCaps(lines: Line[], ctx: MathCtx): Array<{ line: Line; box: Box }> {
  const out: Array<{ line: Line; box: Box }> = [];
  for (const l of lines) {
    const t = l.spans.map((s) => s.text).join("").trim();
    if (!t || t.length > 2 || l.size < 1.8 * ctx.bodySize) continue;
    const h = l.bottom - l.top;
    const beside = lines.filter((o) => o !== l && o.x0 >= l.x1 - 1 && o.x0 - l.x1 < 2 * l.size && o.top < l.bottom && o.bottom > l.top + 0.15 * h);
    if (beside.length < 2) continue;
    out.push({ line: l, box: { x0: l.x0 - 1, y0: l.top - 0.5, x1: l.x1 + 1, y1: l.baseline + 0.1 * l.size } });
  }
  return out;
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
    let r = bb(g);
    // 子图编号：图内或紧贴图下方的短标注
    for (const s of pg.spans) {
      const t = s.text.trim();
      if (!t) continue;
      if (centerIn(s, r, 0.5) || (LABEL.test(t) && s.x0 >= r.x0 - 2 && s.x1 <= r.x1 + 2 && s.top >= r.y1 - 2 && s.top - r.y1 < 1.6 * bodySize)) r = union(r, spanBox(s));
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
