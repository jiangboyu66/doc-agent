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
import { encodePng, encodePngBits } from "./extract.js";

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
  /**
   * 表单页（一个线框占了大半页、文字大半在框里：国际检索报告）才有。表单页的位图整页垫在文本框下面，格里没被识别成文字的词
   * （"カテゴリー*""パテントファミリー""データベー"）留在图里，既不是文字、也没人读它，翻译后仍是原文（WO2017038541A1 的
   * 检索报告，2026-11-27）。这里把它们从位图上取出来：pieces 是每个词裁成的小图（落在一行文字里的带上那一行的基线与字号），
   * background 是抹掉这些词之后的整页位图（PNG，与原图同样大小）。调用方把词作为行内小图片放回原位，由它识别成文字。
   */
  formWords?: { pieces: ScanPiece[]; background: Buffer };
  /**
   * 没有文字层、看得出是文字的页（splitScan 的 opts.blocks）：文字按栏、按段分成的块，每块从位图上裁成一张图，留给调用方识别；
   * background 是抹掉这些字之后的整页位图（PNG，与原图同样大小）：附图、条码、表格线和附图框里的标注留着。
   * 块放进原位置的文本框，识别、翻译后文字落回原来的地方，图留在底图上（US20040163766A1，2026-11-27）
   */
  textBlocks?: { blocks: ScanTextBlock[]; background: Buffer };
}

/** 没有文字层的文字页上的一块文字（见 ScanSplit.textBlocks），页面坐标（pt，y 向下） */
export interface ScanTextBlock {
  x0: number; y0: number; x1: number; y1: number;
  /** 行数、行距（相邻两行中线的距离）与估计的字号（pt） */
  lines: number; pitch: number; size: number;
  /** 三行以上、除末行外各行右端都对齐：两端对齐 */
  align: "left" | "both";
  /**
   * 版面位置，给调用方接续跨栏、跨页的段落用：col 是左栏、右栏或通栏（没有分栏的页、跨栏的块都是 X），band 是第几段版面
   * （横贯全页的大段空白、通栏的块把页面分成上下几段，从 0 起），furniture 是页眉页脚（页顶、页底只有单行块的那一段：
   * 公开号、日期、页码），不属于正文的段落流
   */
  col: "L" | "R" | "X";
  band: number;
  furniture?: boolean;
  png: Buffer; pxW: number; pxH: number;
}

interface Gray { data: Uint8Array | Uint8ClampedArray; w: number; h: number }
interface Box { x0: number; y0: number; x1: number; y1: number }
interface TextBox extends Box { text: string; baseline: number; size: number }

/** 连通块上限：超过的是照片、灰度底纹这类不是黑白线条的图，不拆 */
const MAX_COMPONENTS = 6000;
/**
 * 没有文字层的页：字都在图里，一页双栏正文就有七八千个连通块（每个字一块）。墨迹不到三成、又有这么多块的仍是黑白文字，
 * 照片和底纹在墨迹那一关已经出局
 */
const MAX_BARE_COMPONENTS = 60000;
/** 块的上限：再多就不是"图里剩下的几样东西"了 */
const MAX_PIECES = 400;

/**
 * 拆开一张铺满页面的灰度位图。gray：每像素 1 字节（< 128 为墨迹），行从上到下；box：位图在页面上的位置（pt，y 向下）。
 * texts：页面上的文字片段。返回 null：不是黑白线条图，或文字与墨迹的关系说不清（一部分字下有墨迹、一部分没有）。
 */
export function splitScan(gray: Gray, box: Box, texts: TextBox[], opts: { blocks?: boolean } = {}): ScanSplit | null {
  const { w, h, data } = gray;
  const bw = box.x1 - box.x0, bh = box.y1 - box.y0;
  if (w < 200 || h < 200 || bw <= 0 || bh <= 0) return null;
  // 没有文字层：图里的字就是全部的字
  const bare = !texts.some((t) => t.text.trim());
  const maxComponents = bare ? MAX_BARE_COMPONENTS : MAX_COMPONENTS;
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
  // 每格属于哪个连通块（编号从 1 起）：表单页要把格里的词从位图上抹掉时，按块抹，不碰挨着的格线
  const label = new Uint16Array(gw * gh);
  for (let start = 0; start < grid.length && comps.length <= maxComponents; start++) {
    if (grid[start] !== 1) continue;
    let top = 0;
    const id = comps.length + 1;
    stack[top++] = start;
    grid[start] = 2;
    label[start] = id;
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
          if (grid[k] === 1) { grid[k] = 2; label[k] = id; stack[top++] = k; }
        }
      }
    }
    comps.push(cp);
  }
  if (comps.length > maxComponents) return null;

  // 连通块 → 页面坐标，分类
  type Kind = "glyph" | "graphic";
  const items: Array<Box & { kind: Kind; n: number; id: number }> = [];
  const rules: Box[] = [];
  for (let ci = 0; ci < comps.length; ci++) {
    const cp = comps[ci];
    const b = { x0: box.x0 + cp.c0 * ux, y0: box.y0 + cp.r0 * uy, x1: box.x0 + (cp.c1 + 1) * ux, y1: box.y0 + (cp.r1 + 1) * uy };
    const cw = b.x1 - b.x0, ch = b.y1 - b.y0;
    if (cw < 1.3 && ch < 1.3) continue; // 噪点
    // 细长的线：粗细不到 2.5 pt，或不到长度的 1.2%（扫描略歪的长线会跨几行像素）
    const thin = Math.max(2.5, 0.012 * Math.max(cw, ch));
    if ((ch <= thin && cw >= 30) || (cw <= thin && ch >= 30)) { rules.push(b); continue; }
    items.push({ ...b, kind: cw > 30 || ch > 30 ? "graphic" : "glyph", n: cp.n, id: ci + 1 });
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
  // 附图、条码、线框所在的区域（框里的标注随图留在底图上）
  const graphicBoxes: Box[] = boxes.filter((b) => b.graphic);
  // 没有文字层的页不拿块拼版面（整页图保留，或按 textBlocks 分块）：只裁图形块，块再多也照样分得出附图页和文字页
  if (bare) boxes = graphicBoxes.length > MAX_PIECES ? [] : boxes.filter((b) => b.graphic);
  else if (boxes.length > MAX_PIECES) return null;

  /** 落在文字行里的漏认字：与它同一行、左右紧挨着的文字片段的基线与字号；压在文字上的（重影）返回 null */
  const lineOf = (b: Box): ScanPiece["line"] | null => {
    const cy = (b.y0 + b.y1) / 2, bw2 = b.x1 - b.x0, bh2 = b.y1 - b.y0;
    let line: ScanPiece["line"];
    let best = Infinity;
    for (const t of real) {
      const th = t.y1 - t.y0;
      if (cy < t.y0 - 0.15 * th || cy > t.y1 + 0.15 * th || bh2 > 1.8 * th || bw2 > 14 * t.size) continue;
      const gap = Math.max(t.x0 - b.x1, b.x0 - t.x1);
      // 压在文字上的（重影、同一个字画了两遍）不算漏认
      if (gap < 0 && Math.min(t.x1, b.x1) - Math.max(t.x0, b.x0) > 0.5 * Math.min(bw2, t.x1 - t.x0)) return null;
      if (gap <= 2.5 * t.size && gap < best) { best = gap; line = { baseline: t.baseline, size: t.size }; }
    }
    return line;
  };
  /** 从位图上裁下一块：四周留 1 pt，不贴着笔画裁；行内的块按行高裁 */
  const cut = (b: Box, kind: ScanPiece["kind"], glyphs: number, line: ScanPiece["line"]): ScanPiece => {
    const x0 = Math.max(box.x0, b.x0 - (line ? 0.5 : 1)), x1 = Math.min(box.x1, b.x1 + (line ? 0.5 : 1));
    const y0 = Math.max(box.y0, line ? Math.min(b.y0, line.baseline - 0.8 * line.size) : b.y0 - 1), y1 = Math.min(box.y1, line ? Math.max(b.y1, line.baseline) : b.y1 + 1);
    const c0 = Math.max(0, Math.floor(((x0 - box.x0) / bw) * w)), c1 = Math.min(w, Math.ceil(((x1 - box.x0) / bw) * w));
    const r0 = Math.max(0, Math.floor(((y0 - box.y0) / bh) * h)), r1 = Math.min(h, Math.ceil(((y1 - box.y0) / bh) * h));
    const pw = Math.max(1, c1 - c0), ph = Math.max(1, r1 - r0), out = new Uint8Array(pw * ph);
    for (let y = 0; y < ph; y++) out.set(data.subarray((r0 + y) * w + c0, (r0 + y) * w + c0 + pw), y * pw);
    return { x0, y0, x1, y1, kind, glyphs, ...(line ? { line } : {}), png: encodePng(pw, ph, out, 1), pxW: pw, pxH: ph };
  };
  const pieces: ScanPiece[] = boxes.map((b) => cut(b, b.graphic ? "graphic" : "glyphs", b.glyphs, !b.graphic && !duplicated ? lineOf(b) ?? undefined : undefined));
  pieces.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  const graphics = kept.filter((it) => it.kind === "graphic");
  const graphicArea = Math.min(1, graphics.reduce((a, g) => a + (g.x1 - g.x0) * (g.y1 - g.y0), 0) / (bw * bh));
  const split: ScanSplit = { pieces, rules, inkUnderText, textLines, ink, glyphCount: kept.length - graphics.length, graphicCount: graphics.length, graphicArea };

  // 表单页（与 layout.ts 的判断相同：一个线框占了页面宽的六成、高的四成半以上，页上一半以上的字在框里）：
  // 格线连成一个大块，格里的漏认字在上面"并块"时都被并进了格线，只能另外按字聚成词
  const chars = (list: TextBox[]) => list.reduce((n, t) => n + t.text.replace(/\s/g, "").length, 0);
  const inside = (p: Box) => real.filter((t) => { const cx = (t.x0 + t.x1) / 2, cy = (t.y0 + t.y1) / 2; return cx >= p.x0 && cx <= p.x1 && cy >= p.y0 && cy <= p.y1; });
  const frames = duplicated || !real.length ? [] : pieces.filter((p) => p.kind === "graphic" && p.x1 - p.x0 >= 0.6 * bw && p.y1 - p.y0 >= 0.45 * bh && chars(inside(p)) >= 0.5 * chars(real));
  if (frames.length) {
    const gl = kept.filter((it) => it.kind === "glyph");
    const up = gl.map((_, i) => i);
    const root = (i: number): number => { while (up[i] !== i) { up[i] = up[up[i]]; i = up[i]; } return i; };
    const byX = gl.map((_, i) => i).sort((a, b) => gl[a].x0 - gl[b].x0);
    for (let a = 0; a < byX.length; a++) {
      const p = gl[byX[a]];
      for (let b = a + 1; b < byX.length; b++) {
        const q = gl[byX[b]];
        if (q.x0 - p.x1 > GX) break;
        const gx = Math.max(0, Math.max(p.x0, q.x0) - Math.min(p.x1, q.x1)), gy = Math.max(0, Math.max(p.y0, q.y0) - Math.min(p.y1, q.y1));
        if (gx <= GX && gy <= GY) { const ra = root(byX[a]), rb = root(byX[b]); if (ra !== rb) up[ra] = rb; }
      }
    }
    const words = new Map<number, number[]>();
    for (let i = 0; i < gl.length; i++) { const r = root(i); (words.get(r) ?? words.set(r, []).get(r)!).push(i); }
    const lifted = new Uint8Array(comps.length + 1);
    const out: ScanPiece[] = [];
    for (const ids of words.values()) {
      const m = ids.map((i) => gl[i]);
      const b = { x0: Math.min(...m.map((o) => o.x0)), y0: Math.min(...m.map((o) => o.y0)), x1: Math.max(...m.map((o) => o.x1)), y1: Math.max(...m.map((o) => o.y1)) };
      const hh = b.y1 - b.y0, cx = (b.x0 + b.x1) / 2, cy = (b.y0 + b.y1) / 2;
      // 格里的词：落在线框里（框外页顶、页底的公开号、页码是页面装饰，照旧留在图里）；至少两个字形、其中至少两个有词高的
      // 一半上下（不是一串点、一条虚线），词高像一行字（4–24 pt）。单个的字形（类别栏的 "A"、复选框）留在格线图里：
      // 它们多半不是要翻译的字，读错了反而会丢
      if (!frames.some((p) => cx >= p.x0 && cx <= p.x1 && cy >= p.y0 && cy <= p.y1)) continue;
      if (m.length < 2 || m.filter((o) => o.y1 - o.y0 >= 0.45 * hh).length < 2 || hh < 4 || hh > 24 || m.reduce((n, o) => n + o.n, 0) < 6) continue;
      const line = lineOf(b);
      if (line === null) continue;
      out.push(cut(b, "glyphs", m.length, line));
      for (const o of m) lifted[o.id] = 1;
    }
    if (out.length) {
      // 抹掉取出来的词：只抹这些连通块自己的像素
      const cleared = Uint8Array.from(data);
      for (let y = 0; y < h; y++) {
        const row = y * w, gr = Math.floor(y / f) * gw;
        for (let x = 0; x < w; x++) if (data[row + x] < 128 && lifted[label[gr + ((x / f) | 0)]]) cleared[row + x] = 255;
      }
      out.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
      split.formWords = { pieces: out, background: encodePng(w, h, cleared, 1) };
    }
  }
  if (opts.blocks && bare && looksLikeText(split)) {
    const blocks = scanTextBlocks(gray, box, f, gw, label, comps, kept, graphicBoxes, rules);
    if (blocks) split.textBlocks = blocks;
  }
  return split;
}

/**
 * 没有文字层的文字页分块（ScanSplit.textBlocks）。附图、条码、线框区域里的字形（附图标注）留在图上；其余字形：
 *   1. 成行：横向相邻（间隔不超过约一个半字高）、纵向大半重合的连成行段；
 *   2. 分栏：本页中间一段没有行段横穿的空白是栏间距（页码、个别通栏的字可以压在上面）；同一行、同一栏里的行段并成一行
 *      （段落编号 [0011] 和后面的正文、两端对齐拉开的词距），中间隔着竖的表格线的不并（表格各格各是各的）；
 *   3. 成块：同一栏里上下相邻的行（行距不超过本页行距的 1.3 倍：段间距更大的另起一块）、左右大半重合、中间没有横的表格线的
 *      并成一块；通栏的行只和通栏的行并；
 *   4. 竖排的字（横向页面页边竖着的页眉：一行一个字形、块又细又高）不分块，留在图上，由图片翻译处理。
 * 每块的字号按行距推算（多行块：行距 ÷ 1.2；单行块：按字高与本页多行块的"字号 / 字高"之比）。
 * 底图只抹掉分了块的字形自己的像素（按连通块，挨着的线条、框线不动）。
 */
function scanTextBlocks(gray: Gray, box: Box, f: number, gw: number, label: Uint16Array,
  comps: Array<{ c0: number; r0: number; c1: number; r1: number; n: number }>,
  kept: Array<Box & { kind: "glyph" | "graphic"; n: number; id: number }>, graphicBoxes: Box[], rules: Box[]): ScanSplit["textBlocks"] {
  const { w, h, data } = gray;
  const bw = box.x1 - box.x0, bh = box.y1 - box.y0;
  const components = comps.length;
  const gh = Math.ceil(h / f), ux = bw / gw, uy = bh / gh;
  const within = (it: Box, g: Box) => { const cx = (it.x0 + it.x1) / 2, cy = (it.y0 + it.y1) / 2; return cx >= g.x0 - 1 && cx <= g.x1 + 1 && cy >= g.y0 - 1 && cy <= g.y1 + 1; };
  const median = (v: number[]) => { const s = [...v].sort((a, b) => a - b); return s.length ? s[Math.floor(s.length / 2)] : 0; };
  // 字高：去掉点、逗号这类小墨迹后的中位数
  // 中位数只算一次：放在 filter 里每个元素都重排一遍，一页四五千个字形要三四秒（2026-11-27）
  const all = kept.filter((it) => it.kind === "glyph").map((g) => g.y1 - g.y0), mAll = median(all);
  const hMed = median(all.filter((v) => v >= 0.5 * mAll)) || mAll;
  if (!hMed) return undefined;
  // 粘连成一片的一行字（全大写的小标题 SUMMARY OF THE INVENTION 字母挨着，连成一个 135 pt 宽的连通块）也是字：
  // 矮（不到 1.8 个字高）、扁（宽是高的两倍以上）、墨迹稀（实心的横条不算）。附图里的扁零件（极板、栅网）在附图的范围里，不受影响
  const cell = (it: Box) => Math.max(1, ((it.x1 - it.x0) * (it.y1 - it.y0)) / (bw / (w / f)) / (bh / (h / f)));
  const textLike = (it: Box & { kind: string; n: number }) => it.kind === "graphic" && it.y1 - it.y0 <= 1.8 * hMed && it.x1 - it.x0 >= 2 * (it.y1 - it.y0) && it.n / cell(it) <= 0.6;
  // 条码的竖条、竖线的碎段：又高又细，不是字（留在底图上）
  const bar = (it: Box) => it.y1 - it.y0 >= 2.5 * hMed && it.y1 - it.y0 >= 4 * (it.x1 - it.x0);
  const px = (x: number) => Math.round(((x - box.x0) / bw) * w), py = (y: number) => Math.round(((y - box.y0) / bh) * h);
  const dark = (c: number, r: number) => c >= 0 && c < w && r >= 0 && r < h && data[r * w + c] < 128;
  // 带下划线的字：下划线把碰到它的下伸字母、括号连成一个扁长的连通块（图 8 表里的"13.8" plasma … (Prior Art)"，括号从升部
  // 伸到下划线，块高是字高的两倍多）。按 textLike 的 1.8 倍字高算不是字，留在底图上：块图里少了 p 和括号，块图白底又把底图上的
  // 这几个字盖住，换成译文后它们和下划线一起露出来（2026-11-27）。认法：不到 3 个字高、宽是高的 4 倍以上、墨迹稀，
  // 底部三分之一里几乎每一列都有墨迹（下划线；扫描略歪，一条 330 pt 长的线跨好几行像素，不按某一行横贯来认）。
  // 表格线带着碰到它的字母也是这个样子，挂到行上时再看宽度合不合（见下）
  const underlined = (it: Box & { kind: string; n: number }) => {
    const ht = it.y1 - it.y0;
    if (it.kind !== "graphic" || ht > 3 * hMed || it.x1 - it.x0 < 4 * ht || it.n / cell(it) > 0.35) return false;
    const c0 = px(it.x0), c1 = px(it.x1) - 1, r0 = py(it.y1 - ht / 3), r1 = py(it.y1) - 1;
    let n = 0;
    for (let c = c0; c <= c1; c++) for (let r = r0; r <= r1; r++) if (dark(c, r)) { n++; break; }
    return n >= 0.85 * (c1 - c0 + 1);
  };
  const lined = new Set(kept.filter((it) => underlined(it)));
  const candidates = kept.filter((it) => (it.kind === "glyph" || textLike(it) || lined.has(it as any)) && !bar(it));
  // 附图的范围：含有不像字的图形连通块（高的、实心的）的图形块；只由粘连的字组成的"图形块"不算。
  // 框里有成行文字的（表格、文本框：两行以上、每行 12 个字形以上、宽 80 pt 以上）也不算附图，字照样分块，框线留在底图上
  // （US20040163766A1 的图 8、图 9 两张表，2026-11-27）；附图里散落的标注连不成这样的行
  const rowsIn = (g: Box) => {
    const inside = candidates.filter((it) => within(it, g)).sort((a, b) => a.x0 - b.x0);
    const rows: Array<Box & { n: number }> = [];
    for (const it of inside) {
      const r = rows.find((q) => it.x0 - q.x1 <= Math.max(4, 1.6 * hMed) && Math.min(q.y1, it.y1) - Math.max(q.y0, it.y0) >= 0.4 * Math.min(q.y1 - q.y0, it.y1 - it.y0));
      if (r) { r.x1 = Math.max(r.x1, it.x1); r.y0 = Math.min(r.y0, it.y0); r.y1 = Math.max(r.y1, it.y1); r.n++; } else rows.push({ x0: it.x0, y0: it.y0, x1: it.x1, y1: it.y1, n: 1 });
    }
    return rows.filter((r) => r.n >= 12 && r.x1 - r.x0 >= 80).length;
  };
  const figures = graphicBoxes.filter((g) => kept.some((it) => it.kind === "graphic" && !textLike(it) && !lined.has(it as any) && within(it, g)) && rowsIn(g) < 2);
  const glyphs = candidates.filter((it) => !figures.some((g) => within(it, g)));
  if (glyphs.length < 2) return undefined;
  const gapX = Math.max(4, Math.min(14, 1.6 * hMed));
  // 表格线、框线：在位图上看两段文字之间有没有一道贯穿的墨迹（竖线贯穿行高的八成、横线贯穿重叠宽度的八成）。
  // 粗的表格线是图形连通块，细的是 rules，短的竖线碎段被当成字形——一律按墨迹看（US20040163766A1 图 9 的粗表格线，2026-11-27）
  // 一道"墙"、一道"地板"至少有一个字高（宽）：两段都只是引导点（Int. Cl.⁷ …… C23F 1/00）时，重叠的那几行像素只有点那么高，
  // 夹在中间的一个点（比 1.3 pt 小，不算字形）就占满了，引导点被当成隔着表格线、拆成九块（2026-11-27）
  const hPx = Math.max(3, py(box.y0 + hMed) - py(box.y0)), wPx = Math.max(3, px(box.x0 + hMed) - px(box.x0));
  const wallBetween = (a: Box, b: Box) => {
    const [l, r] = a.x1 <= b.x0 ? [a, b] : [b, a];
    const c0 = px(l.x1) + 1, c1 = px(r.x0) - 1;
    if (c1 < c0) return false;
    const r0 = py(Math.max(a.y0, b.y0)), r1 = py(Math.min(a.y1, b.y1));
    let [top, bottom] = r1 - r0 >= 3 ? [r0, r1] : [py(Math.min(a.y0, b.y0)), py(Math.max(a.y1, b.y1))];
    if (bottom - top + 1 < hPx) { top = Math.round((top + bottom - hPx) / 2); bottom = top + hPx - 1; }
    for (let c = c0; c <= c1; c++) {
      let n = 0;
      for (let rr = top; rr <= bottom; rr++) if (dark(c, rr)) n++;
      if (n >= 0.8 * (bottom - top + 1)) return true;
    }
    return false;
  };
  const floorBetween = (a: Box, b: Box) => {
    const [u, d] = a.y1 <= b.y0 ? [a, b] : [b, a];
    const r0 = py(u.y1) + 1, r1 = py(d.y0) - 1;
    let c0 = px(Math.max(a.x0, b.x0)), c1 = px(Math.min(a.x1, b.x1));
    if (r1 < r0 || c1 - c0 < 3) return false;
    if (c1 - c0 + 1 < wPx) { c0 = Math.round((c0 + c1 - wPx) / 2); c1 = c0 + wPx - 1; }
    for (let rr = r0; rr <= r1; rr++) {
      let n = 0;
      for (let c = c0; c <= c1; c++) if (dark(c, rr)) n++;
      if (n >= 0.8 * (c1 - c0 + 1)) return true;
    }
    return false;
  };

  interface Ln { x0: number; y0: number; x1: number; y1: number; items: typeof glyphs; side?: "L" | "R" | "X" }
  const vOverlap = (a: Box, b: Box) => Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
  let lines: Ln[] = [];
  for (const g of [...glyphs].sort((a, b) => a.x0 - b.x0)) {
    let best: Ln | undefined, bestOv = 0;
    for (const l of lines) {
      if (g.x0 - l.x1 > gapX || g.x1 < l.x0 - gapX) continue;
      const ov = vOverlap(l, g);
      if (ov >= 0.4 * Math.min(l.y1 - l.y0, g.y1 - g.y0) && ov > bestOv && !(g.x0 - l.x1 > 1.5 && wallBetween(l, g))) { best = l; bestOv = ov; }
    }
    if (!best) { lines.push({ x0: g.x0, y0: g.y0, x1: g.x1, y1: g.y1, items: [g] }); continue; }
    best.items.push(g);
    best.x0 = Math.min(best.x0, g.x0);
    best.x1 = Math.max(best.x1, g.x1);
    // 行高不随标点、括号、上下标无限长高（否则一行会吞掉上下两行）
    if (best.y1 - best.y0 < 1.6 * hMed) { best.y0 = Math.min(best.y0, g.y0); best.y1 = Math.max(best.y1, g.y1); }
  }

  // 竖排的字（横向页面页边转了 90° 的页眉 "Patent Application Publication … Sheet 9 of 14"）：一行只有一两个字形的行
  // 上下排成一长条。整条不分块，留在图上交给图片翻译；拆成零碎的小块，识别出来也是横着的碎字
  // 竖排的字一个挨一个（字距、词距都不到 1.5 个字高）；一列项目符号、表格里一列一位数的行，上下隔着一两行，不是竖排的字
  // （图 8 表里左边一列 8 个"•"被当成竖排的字整列去掉，项目符号留在底图上，被下一块的块图盖住，2026-11-27）
  {
    const singles = lines.filter((l) => l.items.length <= 2).sort((a, b) => a.x0 + a.x1 - b.x0 - b.x1);
    const bands: Box[] = [];
    for (let i = 0; i < singles.length;) {
      const cx = (singles[i].x0 + singles[i].x1) / 2;
      let j = i;
      while (j < singles.length && (singles[j].x0 + singles[j].x1) / 2 - cx <= 1.5 * hMed) j++;
      const group = singles.slice(i, j).sort((a, b) => a.y0 - b.y0);
      const y0 = Math.min(...group.map((l) => l.y0)), y1 = Math.max(...group.map((l) => l.y1));
      const gaps = group.slice(1).map((l, k) => l.y0 - group[k].y1);
      if (group.length >= 6 && y1 - y0 >= 60 && median(gaps) <= 1.5 * hMed) bands.push({ x0: Math.min(...group.map((l) => l.x0)) - 1, y0: y0 - 1, x1: Math.max(...group.map((l) => l.x1)) + 1, y1: y1 + 1 });
      i = j;
    }
    if (bands.length) lines = lines.filter((l) => !bands.some((b) => l.x0 >= b.x0 && l.x1 <= b.x1 && l.y0 >= b.y0 && l.y1 <= b.y1));
  }
  if (lines.length < 1) return undefined;

  // 栏间距：页面中间（左右各 30% 以内）最宽的一段没有行段横穿的空白，宽 5 pt 以上、两边都有不少行段。
  // 允许少数行段压在上面（页码、条码下的号码、个别通栏的字）；大号的字（首页的 Patent Application Publication）不算
  const gutter = (() => {
    const x0 = box.x0 + 0.3 * bw, x1 = box.x0 + 0.7 * bw, n = Math.max(0, Math.ceil(x1 - x0));
    if (n <= 0 || lines.length < 8) return undefined;
    const cover = new Uint16Array(n);
    // 大号字：行高超过一般行（5 个字形以上的行）行高的 1.6 倍
    const usual = median(lines.filter((l) => l.items.length >= 5).map((l) => l.y1 - l.y0)) || 2 * hMed;
    for (const l of lines) {
      if (l.y1 - l.y0 > 1.6 * usual) continue;
      for (let x = Math.max(0, Math.floor(l.x0 - x0)); x < Math.min(n, Math.ceil(l.x1 - x0)); x++) cover[x]++;
    }
    const tol = Math.max(2, Math.floor(0.04 * lines.length));
    let best: { a: number; b: number } | undefined;
    for (let a = 0; a < n; ) {
      if (cover[a] > tol) { a++; continue; }
      let b = a;
      while (b < n && cover[b] <= tol) b++;
      if (!best || b - a > best.b - best.a) best = { a, b };
      a = b;
    }
    if (!best || best.b - best.a < 5) return undefined;
    const g = { x0: x0 + best.a, x1: x0 + best.b };
    const left = lines.filter((l) => l.x1 <= g.x0 + 1).length, right = lines.filter((l) => l.x0 >= g.x1 - 1).length;
    return left >= 6 && right >= 6 ? g : undefined;
  })();
  // 哪一栏：按栏间距的中线分；只有左右都越过中线 2 pt 以上的才算通栏。行首挂在栏外的引号、扫描略歪时，栏边的行会伸进
  // 栏间距一两点，按栏间距的边界算会被当成通栏，与本栏的其他行段分开（2026-11-27：段中间冒出只框着"“steering"的小块）
  const sideOf = (l: Box): "L" | "R" | "X" => {
    if (!gutter) return "X";
    const mid = (gutter.x0 + gutter.x1) / 2;
    return l.x1 <= mid + 2 ? "L" : l.x0 >= mid - 2 ? "R" : "X";
  };
  // 同一行、同一栏的行段并成一行；隔着竖的表格线、框线的不并
  const ruleBetween = (a: Box, b: Box) => wallBetween(a, b);
  for (let merged = true; merged;) {
    merged = false;
    lines.sort((a, b) => a.x0 - b.x0);
    outer: for (let i = 0; i < lines.length; i++) {
      for (let j = i + 1; j < lines.length; j++) {
        const a = lines[i], b = lines[j];
        if (b.x0 - a.x1 > 6 * hMed) continue;
        if (vOverlap(a, b) < 0.5 * Math.min(a.y1 - a.y0, b.y1 - b.y0)) continue;
        const sa = sideOf(a), sb = sideOf(b);
        if (sa !== sb && !(gutter && sa === "X" && sb === "X")) continue;
        if (ruleBetween(a, b)) continue;
        a.items.push(...b.items);
        a.x0 = Math.min(a.x0, b.x0); a.x1 = Math.max(a.x1, b.x1);
        a.y0 = Math.min(a.y0, b.y0); a.y1 = Math.max(a.y1, b.y1);
        lines.splice(j, 1);
        merged = true;
        break outer;
      }
    }
  }
  // 下划线式的连通块只留在宽度相合的行上：比这一行别的字左右多出两个字高以上的（表格线带着碰到它的字母）拿掉，留在底图上
  if (lined.size) {
    for (const l of lines) {
      if (!l.items.some((it) => lined.has(it as any))) continue;
      const own = l.items.filter((it) => !lined.has(it as any));
      const ox0 = Math.min(...own.map((it) => it.x0)), ox1 = Math.max(...own.map((it) => it.x1));
      l.items = l.items.filter((it) => !lined.has(it as any) || (own.length > 0 && it.x0 >= ox0 - 2 * hMed && it.x1 <= ox1 + 2 * hMed));
      if (l.items.length) { l.x0 = Math.min(...l.items.map((it) => it.x0)); l.x1 = Math.max(...l.items.map((it) => it.x1)); }
    }
    lines = lines.filter((l) => l.items.length);
  }
  for (const l of lines) l.side = sideOf(l);

  // 本页行距：同一栏里上下相邻两行基线的距离的中位数。基线取行里各字形下沿的中位数：下伸的字母、下标、括号是少数，
  // 不会像行框的中线那样被引号、上下标带偏（2026-11-27：带 V_p、“pancake” 的行被当成另一段）
  const baseOf = new Map<Ln, number>();
  for (const l of lines) baseOf.set(l, median(l.items.map((it) => it.y1)));
  const centre = (l: Ln) => baseOf.get(l) ?? l.y1;
  const steps: number[] = [];
  for (const side of ["L", "R", "X"] as const) {
    const col = lines.filter((l) => l.side === side).sort((a, b) => centre(a) - centre(b));
    for (let k = 1; k < col.length; k++) {
      const d = centre(col[k]) - centre(col[k - 1]);
      if (d > 0.8 * hMed && d < 4 * hMed && Math.min(col[k].x1, col[k - 1].x1) - Math.max(col[k].x0, col[k - 1].x0) > 0) steps.push(d);
    }
  }
  const pagePitch = median(steps) || 2 * hMed;

  interface Bk { lines: Ln[]; x0: number; y0: number; x1: number; y1: number }
  const blocks: Bk[] = [];
  const ruleUnder = (a: Box, b: Box) => floorBetween(a, b);
  for (const l of [...lines].sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0)) {
    let best: Bk | undefined, bestOv = 0;
    for (const b of blocks) {
      const p = b.lines[b.lines.length - 1];
      if (p.side !== l.side) continue;
      const d = centre(l) - centre(p);
      if (d < 0.5 * pagePitch || d > 1.3 * pagePitch) continue;
      const ov = Math.min(l.x1, p.x1) - Math.max(l.x0, p.x0);
      if (ov < 0.6 * Math.min(l.x1 - l.x0, p.x1 - p.x0)) continue;
      // 比块宽出去的行不并（另一栏、旁边的格），居中排的除外：表头格里"Center / Divergence / Angle"一行比一行宽，
      // 中线对齐（图 9，2026-11-27）
      const centred = Math.abs(l.x0 + l.x1 - b.x0 - b.x1) <= 2 * hMed;
      if ((!centred && (l.x0 < b.x0 - 2 * hMed || l.x1 > b.x1 + 2 * hMed)) || ruleUnder(p, l)) continue;
      if (ov > bestOv) { best = b; bestOv = ov; }
    }
    if (!best) { blocks.push({ lines: [l], x0: l.x0, y0: l.y0, x1: l.x1, y1: l.y1 }); continue; }
    best.lines.push(l);
    best.x0 = Math.min(best.x0, l.x0); best.x1 = Math.max(best.x1, l.x1);
    best.y0 = Math.min(best.y0, l.y0); best.y1 = Math.max(best.y1, l.y1);
  }
  // 与更大的块大半重叠的块并进去（没并进行里的零星行段）：否则识别时大块的框里也看得到这几个字，换成文字后会重复
  for (let merged = true; merged;) {
    merged = false;
    blocks.sort((a, b) => (b.x1 - b.x0) * (b.y1 - b.y0) - (a.x1 - a.x0) * (a.y1 - a.y0));
    outer: for (let i = 0; i < blocks.length; i++) {
      for (let j = i + 1; j < blocks.length; j++) {
        const a = blocks[i], b = blocks[j];
        const ix = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0), iy = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
        if (ix <= 0 || iy <= 0 || ix * iy < 0.3 * (b.x1 - b.x0) * (b.y1 - b.y0)) continue;
        a.lines.push(...b.lines);
        a.lines.sort((x, y) => centre(x) - centre(y));
        a.x0 = Math.min(a.x0, b.x0); a.x1 = Math.max(a.x1, b.x1);
        a.y0 = Math.min(a.y0, b.y0); a.y1 = Math.max(a.y1, b.y1);
        blocks.splice(j, 1);
        merged = true;
        break outer;
      }
    }
  }
  // 零星的小块（i 的点、重音、上标挤出来的一两个字形）：并进挨着的块（四周半个字高以内）
  const usualLine = median(blocks.flatMap((b) => b.lines.filter((l) => l.items.length >= 5).map((l) => l.y1 - l.y0))) || 2 * hMed;
  for (let i = blocks.length - 1; i >= 0; i--) {
    const b = blocks[i];
    if (b.lines.reduce((n, l) => n + l.items.length, 0) > 3 || b.y1 - b.y0 >= 0.8 * usualLine) continue;
    const cx = (b.x0 + b.x1) / 2, cy = (b.y0 + b.y1) / 2, m = 0.6 * hMed;
    const host = blocks.find((o) => o !== b && cx >= o.x0 - m && cx <= o.x1 + m && cy >= o.y0 - m && cy <= o.y1 + m && o.lines.reduce((n, l) => n + l.items.length, 0) > 3);
    if (!host) continue;
    host.lines.push(...b.lines);
    host.lines.sort((x, y) => centre(x) - centre(y));
    host.x0 = Math.min(host.x0, b.x0); host.x1 = Math.max(host.x1, b.x1);
    host.y0 = Math.min(host.y0, b.y0); host.y1 = Math.max(host.y1, b.y1);
    blocks.splice(i, 1);
  }
  const keep = blocks.filter((b) => {
    const items = b.lines.reduce((n, l) => n + l.items.length, 0);
    // 竖排的字：一行一个字形、块又细又高；竖排的一个词（转了 90° 的页眉里的 "2004"）：高是宽的两倍以上、每行不过两个字形
    if (b.lines.length >= 3 && median(b.lines.map((l) => l.items.length)) <= 1.5 && b.y1 - b.y0 > 3 * (b.x1 - b.x0)) return false;
    if (b.y1 - b.y0 > 2 * (b.x1 - b.x0) && b.y1 - b.y0 > 1.5 * hMed && median(b.lines.map((l) => l.items.length)) <= 2) return false;
    // 孤立的小墨迹（污点、半个标点）
    return items >= 2 || b.y1 - b.y0 >= 0.6 * hMed;
  });
  if (!keep.length) return undefined;

  // 行距只看像样的行（5 个字形以上）：并进块里的上下标、零星行段夹在中间会把行距算小
  const pitchOf = (b: Bk) => {
    const full = b.lines.filter((l) => l.items.length >= 5 || b.lines.length <= 3).sort((x, y) => centre(x) - centre(y));
    // 同一行里分开的两段（表里一行的左右两格、隔得很开的数值）基线差不多，不算行距
    const steps = full.slice(1).map((l, k) => centre(l) - centre(full[k])).filter((d) => d > 0.5 * hMed);
    if (!steps.length) return 0;
    // 基线落在网格上（300 dpi 的图每格约 0.5 pt）：9.9 pt 的行距量出来是 9.6、10.1 交替，取中位数总偏向一边。
    // 取中位数上下一成半以内的平均值（段与段之间的空行不算）
    const m = median(steps), near = steps.filter((d) => Math.abs(d - m) <= 0.15 * m);
    return near.reduce((a, d) => a + d, 0) / near.length;
  };
  // 字号按字高推：字高取"基线以上的高度"靠上的一段，即大写字母、升部、括号的高度，常见的正文字体是字号的 0.66–0.76，按 0.7 算。
  // 小写为主的正文行与全大写的页眉、附图号量的是同一个高度，下伸的字母只算基线以上的部分。
  // 不按行距推：专利公报正文的行距只有字号的 1.03 倍（9.6 pt 的字、9.9 pt 的行距），按"行距 ≈ 1.2 倍字号"推小了一成半；
  // 也不取字形高度的中位数：正文行量到的是 x 高度、全大写的行是大写高度，单行的页眉按正文的比例换算大了四成（2026-11-27）。
  // 多行块的字号不超过行距（字比行距大，行就叠在一起了）
  // 取上 30% 到上 5% 之间的平均值：字框落在约 0.5 pt 的网格上，单取一个分位数，字高就差出一成（字号在 8.5 与 9.5 之间跳）
  const capOf = (b: Bk) => {
    const v = b.lines.flatMap((l) => l.items.map((it) => centre(l) - it.y0)).filter((x) => x > 0.3 * hMed).sort((x, y) => x - y);
    if (!v.length) return hMed;
    const top = v.slice(Math.floor(0.7 * v.length), Math.max(Math.floor(0.7 * v.length) + 1, Math.ceil(0.95 * v.length)));
    return top.reduce((a, x) => a + x, 0) / top.length;
  };
  const half = (v: number) => Math.round(v * 2) / 2;
  // 正文的字高按全页算：同一种正文，各段单独量出来的字高差一两个网格（0.5 pt），字号就在 9 与 9.5 之间跳。三行以上的块
  // 取全页的中位数，字高与它相差一成以内的多行块都按它算；标题、页眉、表格里的数值仍按各自的字高
  const bodyCap = median(keep.filter((b) => b.lines.filter((l) => l.items.length >= 5).length >= 3).map(capOf));
  const heightOf = (b: Bk) => {
    const own = capOf(b);
    return bodyCap && b.lines.filter((l) => l.items.length >= 5).length >= 2 && Math.abs(own - bodyCap) <= 0.1 * bodyCap ? bodyCap : own;
  };
  // 阅读顺序：通栏的块（跨栏的标题、栏间距里的页码）和横贯全页的大段空白（页眉行与正文之间，1.5 倍行距以上）把页面分成
  // 上下几段，每段里先左栏从上到下、再右栏从上到下。按纵坐标排时左右栏交错（左栏 [0001] 后面是右栏的 [0007]），
  // 按块的先后翻译时上下文跳来跳去，栏尾接到下一栏的句子也断开；页眉行右边的日期也不该排到左栏正文的后面
  const cuts: number[] = [];
  {
    let reach = -Infinity;
    for (const b of [...keep].sort((a, c) => a.y0 - c.y0)) {
      if (reach > -Infinity && b.y0 - reach >= 1.5 * pagePitch) cuts.push((reach + b.y0) / 2);
      reach = Math.max(reach, b.y1);
    }
  }
  const bandOf = (b: Bk) => cuts.filter((c) => c < b.y0).length;
  const sideB = (b: Bk) => (gutter && b.x0 >= gutter.x0 - 1 && b.x1 <= gutter.x1 + 1 ? "X" : sideOf(b));
  // 栏里按行排：上下大半重叠的块（表里同一行的格、一行里分开的两段）从左到右，不按上沿差的一两个像素排
  // （图 9 表头的三格按上沿排成了从右到左，2026-11-27）
  const rows = (bs: Bk[]) => {
    const sorted = [...bs].sort((a, c) => a.y0 - c.y0 || a.x0 - c.x0), res: Bk[] = [];
    for (let i = 0; i < sorted.length;) {
      const head = sorted[i];
      let j = i + 1;
      while (j < sorted.length && Math.min(head.y1, sorted[j].y1) - Math.max(head.y0, sorted[j].y0) >= 0.5 * Math.min(head.y1 - head.y0, sorted[j].y1 - sorted[j].y0)) j++;
      res.push(...sorted.slice(i, j).sort((a, c) => a.x0 - c.x0));
      i = j;
    }
    return res;
  };
  const ordered: Bk[] = [];
  let band: Bk[] = [], at = 0;
  const flush = () => { ordered.push(...rows(band.filter((b) => sideB(b) !== "R")), ...rows(band.filter((b) => sideB(b) === "R"))); band = []; };
  for (const b of [...keep].sort((a, c) => bandOf(a) - bandOf(c) || a.y0 - c.y0 || a.x0 - c.x0)) {
    if (bandOf(b) !== at) { flush(); at = bandOf(b); }
    if (gutter && sideB(b) === "X") { flush(); ordered.push(b); } else band.push(b);
  }
  flush();
  // 版面位置（块的范围扩到完整字形之前定下来）：栏、第几段、页眉页脚
  const realLines = (b: Bk) => b.lines.filter((l) => l.y1 - l.y0 >= 0.5 * hMed).length || 1;
  const place = new Map<Bk, { col: "L" | "R" | "X"; band: number }>(ordered.map((b) => [b, { col: sideB(b), band: bandOf(b) }]));
  const furnitureBands = new Set<number>();
  {
    const used = [...new Set(ordered.map(bandOf))].sort((a, c) => a - c);
    if (used.length > 1) {
      const top = ordered.filter((b) => bandOf(b) === used[0]), bottom = ordered.filter((b) => bandOf(b) === used[used.length - 1]);
      if (top.every((b) => realLines(b) === 1) && Math.max(...top.map((b) => b.y1)) <= box.y0 + 0.12 * bh) furnitureBands.add(used[0]);
      if (bottom.every((b) => realLines(b) === 1) && Math.min(...bottom.map((b) => b.y0)) >= box.y0 + 0.9 * bh) furnitureBands.add(used[used.length - 1]);
    }
  }
  // 每个连通块属于哪一块（从 1 起）：块图只画这一块自己的字，框与框有重叠时，换成文字后不会在别的框里留下这几个字的影子
  const owner = new Uint16Array(components + 1);
  const out: ScanTextBlock[] = [];
  // 块的范围按其中字形的完整外接框算：行框的高度封了顶（不让一行吞掉上下行），连着下划线的字母、高括号会伸出行框；
  // 按行框裁图会把伸出去的部分裁掉，底图上又已抹掉——那几个字就没了（图 8 标题里的 p 和括号，2026-11-27）
  for (const b of ordered) for (const l of b.lines) for (const it of l.items) {
    b.x0 = Math.min(b.x0, it.x0); b.x1 = Math.max(b.x1, it.x1);
    b.y0 = Math.min(b.y0, it.y0); b.y1 = Math.max(b.y1, it.y1);
  }
  ordered.forEach((b, k) => { for (const l of b.lines) for (const it of l.items) owner[it.id] = k + 1; });
  // 整个落在块里、没分进行的小墨迹归这一块：比 1.3 pt 小、拆连通块时当成噪点丢掉的（i 上的点、小号字的句点、扫描的污点），
  // 没分进行的零星字形，以及块里的细线（与字分开的下划线）。块图是白底的，不归进来时块图里的 i 没有点（识别成 ı），
  // 底图上的点又被块图盖住，换成译文后散在译文上（US20040163766A1 的正文，每页上千个像素，2026-11-27）。
  // 图形连通块（附图、表格线）不动
  // 首行 i 上的点、行尾的句点会探出字形的外接框一点：小墨迹放宽到块外 0.4 个字高，归进来后块的范围跟着扩大
  {
    const graphic = new Uint8Array(components + 1), glyph = new Uint8Array(components + 1);
    for (const it of kept) (it.kind === "graphic" ? graphic : glyph)[it.id] = 1;
    const grow: Array<[number, number, number, number] | undefined> = ordered.map(() => undefined);
    for (let id = 1; id <= components; id++) {
      if (owner[id] || graphic[id]) continue;
      const cp = comps[id - 1];
      const x0 = box.x0 + cp.c0 * ux, y0 = box.y0 + cp.r0 * uy, x1 = box.x0 + (cp.c1 + 1) * ux, y1 = box.y0 + (cp.r1 + 1) * uy;
      const m = !glyph[id] && x1 - x0 < 1.3 && y1 - y0 < 1.3 ? Math.max(0.5, 0.4 * hMed) : 0.5;
      const k = ordered.findIndex((b) => x0 >= b.x0 - m && x1 <= b.x1 + m && y0 >= b.y0 - m && y1 <= b.y1 + m);
      if (k < 0) continue;
      owner[id] = k + 1;
      const g = grow[k] ?? [x0, y0, x1, y1];
      grow[k] = [Math.min(g[0], x0), Math.min(g[1], y0), Math.max(g[2], x1), Math.max(g[3], y1)];
    }
    for (const [k, g] of grow.entries()) {
      if (!g) continue;
      const b = ordered[k];
      b.x0 = Math.min(b.x0, g[0]); b.y0 = Math.min(b.y0, g[1]); b.x1 = Math.max(b.x1, g[2]); b.y1 = Math.max(b.y1, g[3]);
    }
  }
  for (const [k, b] of ordered.entries()) {
    const pitch = pitchOf(b);
    const size = Math.max(5, Math.min(40, half(Math.min(heightOf(b) / 0.7, pitch > 0 ? pitch : Infinity))));
    // 两端对齐：除末行外，大多数行的右端齐在栏的右边上（取行右端的上四分位数，行尾挂出去的连字符、引号不算）
    const body = b.lines.filter((l) => l.items.length >= 5).sort((x, y) => centre(x) - centre(y));
    const edge = body.length ? [...body.map((l) => l.x1)].sort((x, y) => x - y)[Math.floor(0.75 * (body.length - 1))] : b.x1;
    const ends = body.slice(0, -1).filter((l) => l.x1 >= edge - Math.max(1.5, 0.3 * hMed)).length;
    const align: ScanTextBlock["align"] = body.length >= 3 && ends >= 0.7 * (body.length - 1) ? "both" : "left";
    // 四周留一点空：框正好贴着字时，换成文字后最后一个字母可能折到下一行
    const x0 = Math.max(box.x0, b.x0 - 1), y0 = Math.max(box.y0, b.y0 - 1), x1 = Math.min(box.x1, b.x1 + 1.5), y1 = Math.min(box.y1, b.y1 + 1);
    const c0 = Math.max(0, Math.floor(((x0 - box.x0) / bw) * w)), c1 = Math.min(w, Math.ceil(((x1 - box.x0) / bw) * w));
    const r0 = Math.max(0, Math.floor(((y0 - box.y0) / bh) * h)), r1 = Math.min(h, Math.ceil(((y1 - box.y0) / bh) * h));
    const pw = Math.max(1, c1 - c0), ph = Math.max(1, r1 - r0), crop = new Uint8Array(pw * ph).fill(255);
    for (let y = 0; y < ph; y++) {
      const row = (r0 + y) * w, gr = Math.floor((r0 + y) / f) * gw;
      for (let x = 0; x < pw; x++) if (data[row + c0 + x] < 128 && owner[label[gr + (((c0 + x) / f) | 0)]] === k + 1) crop[y * pw + x] = data[row + c0 + x];
    }
    // 行数不算单独成行的 i 上的点、上标（比半个字高还矮的行段）
    const { col, band } = place.get(b)!;
    out.push({ x0, y0, x1, y1, lines: realLines(b), pitch: pitch > 0 ? pitch : 1.2 * size, size, align, col, band,
      ...(furnitureBands.has(band) ? { furniture: true } : {}), png: encodePngBits(pw, ph, crop), pxW: pw, pxH: ph });
  }
  // 底图：只抹掉分了块的字形自己的像素
  const cleared = Uint8Array.from(data);
  for (let y = 0; y < h; y++) {
    const row = y * w, gr = Math.floor(y / f) * gw;
    for (let x = 0; x < w; x++) if (data[row + x] < 128 && owner[label[gr + ((x / f) | 0)]]) cleared[row + x] = 255;
  }
  return { blocks: out, background: encodePngBits(w, h, cleared) };
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
