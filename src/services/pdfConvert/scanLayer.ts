/**
 * 扫描页的位图层。
 *
 * 一类扫描件（WIPO 的 PCT 公布文本是典型）不是"扫描图 + 压在上面的识别文字"，而是把识别出来的字从图里抹掉、改用字体文字
 * 画在原处：整页位图里只剩没被识别成文字的东西——附图、条码、徽标、表格线、个别漏认的字——文字画在上面。
 * 两层互补，合起来才是完整的一页。把整页位图垫在文字下面只在逐页保留的版面里成立；可编辑排版下文字重排、
 * 翻译后长度变化，图里的东西就不在原来的字旁边了。
 *
 * 这里把位图层拆开：按连通块找出墨迹，细长的线和噪点不要，其余的按相邻关系聚成"块"，每块从位图上裁成一张图。
 * 排版时大的块（附图、条码）当作普通图片，落在文字行里的小块（漏认的字）作为行内图片插回那一行。
 *
 * 没有文字层的整页扫描图也在这里看一眼：墨迹成行（很多行、每行很多字）的是没做过文字识别的文字页，
 * 其余的是图。调用方据此区分"附图页可以按图片保留"和"这页的文字还没被读出来"。
 */
import { encodePng } from "./extract.js";

export interface ScanPiece {
  x0: number; y0: number; x1: number; y1: number;
  /** graphic：含比一个字大的连通块（附图、条码、表格线框）；glyphs：只有字形大小的墨迹（没被识别成文字的字） */
  kind: "graphic" | "glyphs";
  /** 字形大小的连通块个数 */
  glyphs: number;
  /**
   * 落在一行文字里的漏认字（左右紧挨着同一行的文字）：那一行的基线与字号。这样的块按整行的高度裁（上到字顶、下到基线），
   * 作为行内图片插回去时字的高低位置不变（连字符仍在半高处，而不是落到基线上）
   */
  line?: { baseline: number; size: number };
  png: Buffer; pxW: number; pxH: number;
}

export interface ScanSplit {
  pieces: ScanPiece[];
  /** 细长的线（表格线、分隔线）：不保留，只记位置 */
  rules: Array<{ x0: number; y0: number; x1: number; y1: number }>;
  /** 字框下有墨迹的文字片段所占比例：接近 0 是"字已从图里抹掉"，接近 1 是"图里还有字、文字是叠上去的一层" */
  inkUnderText: number;
  /** 成行的字形：长度 ≥ 150 pt、至少 12 个字形连通块的行数（横、竖都算）。文字页有十几行以上，图上的标注没有这么长的行 */
  textLines: number;
  /** 墨迹占页面的比例 */
  ink: number;
  /** 字形大小的连通块数、比字大的连通块（线框、箭头、曲线）数，以及后者的外接框合计占位图面积的比例 */
  glyphCount: number;
  graphicCount: number;
  graphicArea: number;
}

interface Gray { data: Uint8Array | Uint8ClampedArray; w: number; h: number }
interface Box { x0: number; y0: number; x1: number; y1: number }
interface TextBox extends Box { text: string; baseline: number; size: number }

/** 连通块上限：超过的是照片、灰度底纹这类不是黑白线条的图，不拆 */
const MAX_COMPONENTS = 6000;
/** 块的上限：再多就不是"图里剩下的几样东西"了 */
const MAX_PIECES = 400;

/**
 * 拆开一张铺满页面的灰度位图。gray：每像素 1 字节（< 128 为墨迹），行从上到下；box：位图在页面上的位置（pt，y 向下）。
 * texts：页面上的文字片段。返回 null：不是黑白线条图，或文字与墨迹的关系说不清（一部分字下有墨迹、一部分没有）。
 */
export function splitScan(gray: Gray, box: Box, texts: TextBox[]): ScanSplit | null {
  const { w, h, data } = gray;
  const bw = box.x1 - box.x0, bh = box.y1 - box.y0;
  if (w < 200 || h < 200 || bw <= 0 || bh <= 0) return null;
  // 网格：约每 pt 2 格（300 dpi 的图每格 2 × 2 像素）。格内有一个墨迹像素就算有墨迹
  const f = Math.max(1, Math.round(w / bw / 2));
  const gw = Math.ceil(w / f), gh = Math.ceil(h / f);
  const ux = bw / gw, uy = bh / gh;
  const grid = new Uint8Array(gw * gh);
  let dark = 0;
  for (let y = 0; y < h; y++) {
    const row = y * w, gr = Math.floor(y / f) * gw;
    for (let x = 0; x < w; x++) if (data[row + x] < 128 && !grid[gr + ((x / f) | 0)]) { grid[gr + ((x / f) | 0)] = 1; dark++; }
  }
  const ink = dark / (gw * gh);
  // 墨迹超过三成：反相的图、照片、底色，不是白底上的线条
  if (ink > 0.3) return null;

  // 连通块（8 邻接）
  interface Comp { c0: number; r0: number; c1: number; r1: number; n: number }
  const comps: Comp[] = [];
  const stack = new Int32Array(dark + 1);
  for (let start = 0; start < grid.length && comps.length <= MAX_COMPONENTS; start++) {
    if (grid[start] !== 1) continue;
    let top = 0;
    stack[top++] = start;
    grid[start] = 2;
    const cp: Comp = { c0: gw, r0: gh, c1: -1, r1: -1, n: 0 };
    while (top) {
      const at = stack[--top], r = (at / gw) | 0, c = at - r * gw;
      cp.n++;
      if (c < cp.c0) cp.c0 = c;
      if (c > cp.c1) cp.c1 = c;
      if (r < cp.r0) cp.r0 = r;
      if (r > cp.r1) cp.r1 = r;
      for (let dr = -1; dr <= 1; dr++) {
        const rr = r + dr;
        if (rr < 0 || rr >= gh) continue;
        for (let dc = -1; dc <= 1; dc++) {
          const cc = c + dc;
          if (cc < 0 || cc >= gw) continue;
          const k = rr * gw + cc;
          if (grid[k] === 1) { grid[k] = 2; stack[top++] = k; }
        }
      }
    }
    comps.push(cp);
  }
  if (comps.length > MAX_COMPONENTS) return null;

  // 连通块 → 页面坐标，分类
  type Kind = "glyph" | "graphic";
  const items: Array<Box & { kind: Kind; n: number }> = [];
  const rules: Box[] = [];
  for (const cp of comps) {
    const b = { x0: box.x0 + cp.c0 * ux, y0: box.y0 + cp.r0 * uy, x1: box.x0 + (cp.c1 + 1) * ux, y1: box.y0 + (cp.r1 + 1) * uy };
    const cw = b.x1 - b.x0, ch = b.y1 - b.y0;
    if (cw < 1.3 && ch < 1.3) continue; // 噪点
    // 细长的线：粗细不到 2.5 pt，或不到长度的 1.2%（扫描略歪的长线会跨几行像素）
    const thin = Math.max(2.5, 0.012 * Math.max(cw, ch));
    if ((ch <= thin && cw >= 30) || (cw <= thin && ch >= 30)) { rules.push(b); continue; }
    items.push({ ...b, kind: cw > 30 || ch > 30 ? "graphic" : "glyph", n: cp.n });
  }

  // 文字与墨迹的关系：字框下有墨迹的片段占多少
  const glyphs = items.filter((it) => it.kind === "glyph");
  const inText = (it: Box) => {
    const cx = (it.x0 + it.x1) / 2, cy = (it.y0 + it.y1) / 2;
    return texts.some((t) => cx >= t.x0 && cx <= t.x1 && cy >= t.y0 && cy <= t.y1);
  };
  const real = texts.filter((t) => t.text.trim());
  let withInk = 0;
  for (const t of real) if (glyphs.some((g) => { const cx = (g.x0 + g.x1) / 2, cy = (g.y0 + g.y1) / 2; return cx >= t.x0 && cx <= t.x1 && cy >= t.y0 && cy <= t.y1; })) withInk++;
  const inkUnderText = real.length ? withInk / real.length : 0;
  if (real.length >= 10 && inkUnderText > 0.2 && inkUnderText < 0.6) return null;
  // 图里还有字、文字是叠上去的一层：字框下的墨迹就是那些字，不重复保留
  const duplicated = inkUnderText >= 0.6;
  const kept = duplicated ? items.filter((it) => !(it.kind === "glyph" && inText(it))) : items;

  // 成行的字形：把横向（竖向）相邻的字形连成行，数足够长的行
  const textLines = countLines(kept.filter((it) => it.kind === "glyph"));

  // 聚成块：外接框相距不远的并在一起（一行字、相邻的几行、图和它旁边的标注）
  const parent = kept.map((_, i) => i);
  const find = (i: number): number => { while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; } return i; };
  const GX = 4.5, GY = 2.6;
  const order = kept.map((_, i) => i).sort((a, b) => kept[a].x0 - kept[b].x0);
  for (let a = 0; a < order.length; a++) {
    const p = kept[order[a]];
    // 图形的外接框里可能包着别的块（线框里的字），不能按 x 提前截断；字形只和右边不远的比
    for (let b = a + 1; b < order.length; b++) {
      const q = kept[order[b]];
      if (p.kind === "glyph" && q.x0 - p.x1 > GX) break;
      const gx = Math.max(0, Math.max(p.x0, q.x0) - Math.min(p.x1, q.x1)), gy = Math.max(0, Math.max(p.y0, q.y0) - Math.min(p.y1, q.y1));
      if (gx <= GX && gy <= GY) { const ra = find(order[a]), rb = find(order[b]); if (ra !== rb) parent[ra] = rb; }
    }
  }
  const groups = new Map<number, number[]>();
  for (let i = 0; i < kept.length; i++) { const r = find(i); (groups.get(r) ?? groups.set(r, []).get(r)!).push(i); }
  // 外接框互相包含、交叠的块再并一次（图的外框把里面的标注块整个包住）
  let boxes = [...groups.values()].map((ids) => {
    const m = ids.map((i) => kept[i]);
    return {
      x0: Math.min(...m.map((o) => o.x0)), y0: Math.min(...m.map((o) => o.y0)), x1: Math.max(...m.map((o) => o.x1)), y1: Math.max(...m.map((o) => o.y1)),
      graphic: m.some((o) => o.kind === "graphic"), glyphs: m.filter((o) => o.kind === "glyph").length, n: m.reduce((s, o) => s + o.n, 0),
    };
  });
  for (let merged = true; merged;) {
    merged = false;
    outer: for (let i = 0; i < boxes.length; i++) {
      for (let j = 0; j < boxes.length; j++) {
        if (i === j) continue;
        const a = boxes[i], b = boxes[j];
        if (!a.graphic) continue;
        const ix = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0), iy = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
        if (ix <= 0 || iy <= 0) continue;
        // 只并"大半落在图的外接框里"的块：擦过一个角的不算
        if (ix * iy < 0.5 * (b.x1 - b.x0) * (b.y1 - b.y0)) continue;
        boxes[i] = { x0: Math.min(a.x0, b.x0), y0: Math.min(a.y0, b.y0), x1: Math.max(a.x1, b.x1), y1: Math.max(a.y1, b.y1), graphic: true, glyphs: a.glyphs + b.glyphs, n: a.n + b.n };
        boxes.splice(j, 1);
        merged = true;
        break outer;
      }
    }
  }
  // 只有两三格墨迹的碎屑不要
  boxes = boxes.filter((b) => b.n >= 6 || b.graphic);
  if (boxes.length > MAX_PIECES) return null;

  const pieces: ScanPiece[] = boxes.map((b) => {
    // 落在文字行里的漏认字：与它同一行、左右紧挨着的文字片段
    let line: ScanPiece["line"];
    if (!b.graphic && !duplicated) {
      const cy = (b.y0 + b.y1) / 2, bw2 = b.x1 - b.x0, bh2 = b.y1 - b.y0;
      let best = Infinity;
      for (const t of real) {
        const th = t.y1 - t.y0;
        if (cy < t.y0 - 0.15 * th || cy > t.y1 + 0.15 * th || bh2 > 1.8 * th || bw2 > 14 * t.size) continue;
        const gap = Math.max(t.x0 - b.x1, b.x0 - t.x1);
        // 压在文字上的（重影、同一个字画了两遍）不算漏认
        if (gap < 0 && Math.min(t.x1, b.x1) - Math.max(t.x0, b.x0) > 0.5 * Math.min(bw2, t.x1 - t.x0)) { line = undefined; best = -1; break; }
        if (gap <= 2.5 * t.size && gap < best) { best = gap; line = { baseline: t.baseline, size: t.size }; }
      }
    }
    // 四周留 1 pt，不贴着笔画裁；行内的块按行高裁
    const x0 = Math.max(box.x0, b.x0 - (line ? 0.5 : 1)), x1 = Math.min(box.x1, b.x1 + (line ? 0.5 : 1));
    const y0 = Math.max(box.y0, line ? Math.min(b.y0, line.baseline - 0.8 * line.size) : b.y0 - 1), y1 = Math.min(box.y1, line ? Math.max(b.y1, line.baseline) : b.y1 + 1);
    const c0 = Math.max(0, Math.floor(((x0 - box.x0) / bw) * w)), c1 = Math.min(w, Math.ceil(((x1 - box.x0) / bw) * w));
    const r0 = Math.max(0, Math.floor(((y0 - box.y0) / bh) * h)), r1 = Math.min(h, Math.ceil(((y1 - box.y0) / bh) * h));
    const pw = Math.max(1, c1 - c0), ph = Math.max(1, r1 - r0), out = new Uint8Array(pw * ph);
    for (let y = 0; y < ph; y++) out.set(data.subarray((r0 + y) * w + c0, (r0 + y) * w + c0 + pw), y * pw);
    return { x0, y0, x1, y1, kind: b.graphic ? "graphic" : "glyphs", glyphs: b.glyphs, ...(line ? { line } : {}), png: encodePng(pw, ph, out, 1), pxW: pw, pxH: ph };
  });
  pieces.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  const graphics = kept.filter((it) => it.kind === "graphic");
  const graphicArea = Math.min(1, graphics.reduce((a, g) => a + (g.x1 - g.x0) * (g.y1 - g.y0), 0) / (bw * bh));
  return { pieces, rules, inkUnderText, textLines, ink, glyphCount: kept.length - graphics.length, graphicCount: graphics.length, graphicArea };
}

/** 成行的字形：相邻（间距不超过约一个字）且对齐的字形连成一行；数长度 ≥ 150 pt、至少 12 个字形的行，横竖都算 */
function countLines(glyphs: Box[]): number {
  let n = 0;
  for (const vertical of [false, true]) {
    // 竖排时把坐标转置，按同一套逻辑数
    const g = glyphs.map((b) => (vertical ? { x0: b.y0, x1: b.y1, y0: b.x0, y1: b.x1 } : b)).sort((a, b) => a.x0 - b.x0);
    const used = new Uint8Array(g.length);
    for (let i = 0; i < g.length; i++) {
      if (used[i]) continue;
      used[i] = 1;
      let count = 1, x0 = g[i].x0, x1 = g[i].x1, y0 = g[i].y0, y1 = g[i].y1;
      for (let j = i + 1; j < g.length; j++) {
        if (used[j]) continue;
        const q = g[j];
        if (q.x0 - x1 > 14) break;
        // 同一行：纵向范围与行大半重合
        const ov = Math.min(y1, q.y1) - Math.max(y0, q.y0);
        if (ov < 0.4 * Math.min(y1 - y0, q.y1 - q.y0)) continue;
        used[j] = 1;
        count++;
        x1 = Math.max(x1, q.x1);
        // 行高不随着并进来的字无限长高（否则一行会吞掉上下行）
        if (y1 - y0 < 14) { y0 = Math.min(y0, q.y0); y1 = Math.max(y1, q.y1); }
      }
      if (count >= 12 && x1 - x0 >= 150 && y1 - y0 <= 24) n++;
    }
  }
  return n;
}

/**
 * 没有文字层的整页扫描图看起来是不是文字页：成行的字很多；或者满页都是字形、没有比字大的线条。
 * 附图页有线框、箭头、曲线（外接框占了页面的相当一部分），标注都是短的。
 */
export function looksLikeText(sc: Pick<ScanSplit, "textLines" | "glyphCount" | "graphicArea">): boolean {
  return sc.textLines >= 12 || (sc.graphicArea < 0.05 && sc.glyphCount >= 60);
}
