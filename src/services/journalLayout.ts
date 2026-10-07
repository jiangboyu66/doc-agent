/**
 * 学术期刊排版（Word）
 *
 * 目标（与 LaTeX 期刊模板的版面规则一致）：
 *   1. 浮动体：宽于一栏的图、表占满整行（单栏分节、左右居中），放在页面顶部；不宽的图表放在栏内，
 *      当前栏放得下就留在原处，放不下就移到下一栏顶部——文字继续排满当前栏，不留大块空白；
 *   2. 不被打断：表格每行不跨页、整表与表题保持在同一页；图与图题不分离；
 *   3. 栏底平齐：所有段落的"段前距 + 行高 × 行数 + 段后距"凑成正文行距的整数倍（基线网格），
 *      公式、栏内图表同样凑整，左右两栏的每一行都在同一高度，排满的栏底部对齐；关闭孤行控制。
 *
 * 浮动体的位置需要知道分页结果：用 LibreOffice 把当前文档渲染成 PDF，按文字顺序把 PDF 中的每个词
 * 对应回文档中的段落与字符位置（"版面映射"），据此找到"下一页顶部""下一栏顶部"在文档中的位置；
 * 那里若正好在段落中间，就把段落拆成两段（前半段末行两端对齐，后半段不缩进），浮动体插在中间。
 * 每放一个浮动体重新渲染一次。没有 LibreOffice 时只做不依赖分页的部分。
 *
 * 再次运行时先还原上一次的排版（拆开的段落合并回去、插入的分节去掉、浮动体回到引用它的位置之后），
 * 因此修改内容后可以反复运行。
 */

import { ZipPackage } from "../documents/docx/zip.js";
import {
  parseXml, parseFragment, serializeDocument, serializeNode, removeNode, insertBefore, insertAfter, appendChild,
  innerText, encodeText, markDirty, type XDocument, type XElement, type XNode,
} from "../documents/docx/xml.js";
import { REL_TYPES, parseRels, resolveTarget, relsPathFor, orderedInsertIndex, PPR_ORDER, TRPR_ORDER, SECTPR_ORDER, TBLPR_ORDER } from "../documents/docx/ooxml.js";
import { extractPdf, type PageModel } from "./pdfConvert/extract.js";
import { buildLines, type Line } from "./pdfConvert/layout.js";

// ---------------------------------------------------------------------------
// 选项与报告
// ---------------------------------------------------------------------------

export interface JournalOptions {
  /** auto：宽图表占满整行放页顶、栏内图表放得下留原处否则移到下一栏顶；wide：只处理宽图表；none：不移动图表 */
  floats?: "auto" | "wide" | "none";
  /** 基线网格：段距、公式、栏内图表凑成行距整数倍，两栏逐行对齐（默认开） */
  grid?: boolean;
  /** 正文统一为固定行距（单倍行距的正文改为固定值，网格才能精确；默认开） */
  exactLines?: boolean;
  /** 关闭孤行控制（默认开：Word 的孤行控制会让两栏底部差一行） */
  widowOff?: boolean;
  /** 渲染器：docx → PDF（LibreOffice）。不提供时不做依赖分页的浮动体定位 */
  render?: (docx: Buffer) => Promise<Buffer>;
}

export interface FloatReport { kind: "figure" | "table"; caption: string; wide: boolean; placement: string; page?: number }

export interface JournalReport {
  floats: FloatReport[];
  pitch: number | null;
  gridAdjusted: number;
  keepTogether: number;
  pages: number | null;
  /** 中间页（不含最后一页）中留白超过正文区 15% 的页码 */
  whitespacePages: number[];
  /** 两栏底部没有对齐的页码 */
  unevenPages: number[];
  renders: number;
  notes: string[];
}

// ---------------------------------------------------------------------------
// 题注识别
// ---------------------------------------------------------------------------

const FIG_CAP = /^\s*((fig\.?|figure)\s*\d+[a-z]?\s*[:.．。：]|图\s*\d+)/i;
const TAB_CAP_I = /^\s*((table|tab\.)\s*[\dIVXLC]+\s*[:.．。：]|表\s*\d+)/i;
/** 全大写的 TABLE I（IEEE 表题常换行，不带冒号）；区分大小写，避免把 "Table V shows…" 当成表题 */
const TAB_CAP_U = /^\s*TABLE\s+[IVXLC\d]+\b/;
const TAB_CAP = { test: (s: string) => TAB_CAP_I.test(s) || TAB_CAP_U.test(s) } as RegExp;

const TWIP_PER_PT = 20;
const EMU_PER_PT = 12700;

// ---------------------------------------------------------------------------
// 样式（只解析排版需要的：段距、行距、对齐、字号）
// ---------------------------------------------------------------------------

interface Sty { basedOn?: string; before?: number; after?: number; line?: number; lineRule?: string; jc?: string; sz?: number }

class Styles {
  private map = new Map<string, Sty>();
  private def: Sty = {};
  private defaultPara?: string;
  constructor(xml: string | null) {
    if (!xml) return;
    const d = parseXml(xml);
    const root = d.root;
    const p = root.prefix;
    const q = (l: string) => `${p}:${l}`;
    const read = (pPr?: XElement, rPr?: XElement): Sty => {
      const s: Sty = {};
      const sp = pPr?.child(q("spacing"));
      const num = (v?: string) => (v !== undefined && v !== "" && Number.isFinite(Number(v)) ? Number(v) : undefined);
      if (sp) {
        s.before = num(sp.getAttr(q("before")));
        s.after = num(sp.getAttr(q("after")));
        s.line = num(sp.getAttr(q("line")));
        s.lineRule = sp.getAttr(q("lineRule"));
      }
      const jc = pPr?.child(q("jc"))?.getAttr(q("val"));
      if (jc) s.jc = jc;
      const sz = num(rPr?.child(q("sz"))?.getAttr(q("val")));
      if (sz !== undefined) s.sz = sz;
      return s;
    };
    const dd = root.child(q("docDefaults"));
    this.def = {
      ...read(dd?.child(q("pPrDefault"))?.child(q("pPr")), undefined),
      ...strip(read(undefined, dd?.child(q("rPrDefault"))?.child(q("rPr")))),
    };
    for (const st of root.childrenNamed(q("style"))) {
      if (st.getAttr(q("type")) !== "paragraph") continue;
      const id = st.getAttr(q("styleId")) ?? "";
      const s = read(st.child(q("pPr")), st.child(q("rPr")));
      s.basedOn = st.child(q("basedOn"))?.getAttr(q("val"));
      this.map.set(id, strip(s));
      if (st.getAttr(q("default")) === "1") this.defaultPara = id;
    }
  }
  resolve(id?: string): Sty {
    const chain: Sty[] = [];
    const seen = new Set<string>();
    let cur = id ?? this.defaultPara;
    while (cur && !seen.has(cur)) {
      seen.add(cur);
      const s = this.map.get(cur);
      if (!s) break;
      chain.unshift(s);
      cur = s.basedOn;
    }
    return Object.assign({}, this.def, ...chain.map(strip));
  }
}

function strip<T extends object>(o: T): T {
  const out: any = {};
  for (const [k, v] of Object.entries(o)) if (v !== undefined && k !== "basedOn") out[k] = v;
  return out;
}

// ---------------------------------------------------------------------------
// 文档模型
// ---------------------------------------------------------------------------

interface Float {
  kind: "figure" | "table";
  nodes: XElement[];
  caption: string;
  wide: boolean;
  /** 栏宽（pt），用于估算题注行数 */
  colW: number;
  /** 引用它的位置：浮动体原来所在位置之前的那个正文元素 */
  anchor: XElement | null;
  id: number;
}

interface Geo { pageW: number; pageH: number; top: number; bottom: number; left: number; right: number; cols: number; gap: number }

class JDoc {
  private pkg: ZipPackage;
  private main: string;
  xml: XDocument;
  w: string;
  body: XElement;
  styles: Styles;
  private bm = 0;

  constructor(buf: Buffer) {
    this.pkg = new ZipPackage(buf);
    const rootRels = parseRels(this.pkg.has("_rels/.rels") ? this.pkg.readText("_rels/.rels") : null);
    const m = rootRels.find((r) => r.type === REL_TYPES.officeDocument);
    this.main = m ? resolveTarget("", m.target) : "word/document.xml";
    this.xml = parseXml(this.pkg.readText(this.main));
    const root = this.xml.root;
    this.w = root.prefix || "w";
    const b = root.child(this.q("body"));
    if (!b) throw new Error("文档缺少正文");
    this.body = b;
    const rels = parseRels(this.pkg.has(relsPathFor(this.main)) ? this.pkg.readText(relsPathFor(this.main)) : null);
    const st = rels.find((r) => r.type === REL_TYPES.styles);
    const stPart = st ? resolveTarget(this.main, st.target) : null;
    this.styles = new Styles(stPart && this.pkg.has(stPart) ? this.pkg.readText(stPart) : null);
    for (const d of this.body.descendants()) {
      if (d.local === "bookmarkStart" || d.local === "bookmarkEnd") {
        const id = Number(d.getAttr(this.q("id")));
        if (Number.isFinite(id)) this.bm = Math.max(this.bm, id);
      }
    }
  }

  q = (l: string) => `${this.w}:${l}`;

  serialize(): Buffer {
    this.pkg.write(this.main, Buffer.from(serializeDocument(this.xml), "utf8"));
    return this.pkg.toBuffer();
  }

  // ---------------- 基础判断 ----------------

  top(): XElement[] { return this.body.elements(); }
  isP = (e: XElement) => e.prefix === this.w && e.local === "p";
  isTbl = (e: XElement) => e.prefix === this.w && e.local === "tbl";
  sect = (e: XElement) => (this.isP(e) ? e.child(this.q("pPr"))?.child(this.q("sectPr")) : undefined);
  hasObj(e: XElement): boolean {
    for (const d of e.descendants()) if (d.local === "drawing" || d.local === "pict" || d.local === "object" || d.local === "oMath") return true;
    return false;
  }
  inlineDrawings(e: XElement): XElement[] {
    const out: XElement[] = [];
    for (const d of e.descendants()) if (d.local === "inline" && (d.parent as XElement)?.local === "drawing") out.push(d);
    return out;
  }
  extent(inline: XElement): { w: number; h: number } {
    const ex = inline.elements().find((x) => x.local === "extent");
    return { w: Number(ex?.getAttr("cx") ?? 0) / EMU_PER_PT, h: Number(ex?.getAttr("cy") ?? 0) / EMU_PER_PT };
  }

  /** 段落 / 元素的渲染文字（与版面映射、拆段使用同一套规则） */
  text(n: XElement): string {
    let s = "";
    const walk = (e: XElement) => {
      for (const c of e.children) {
        if (c.type !== "el") continue;
        const l = c.local;
        if (c.prefix === this.w) {
          if (l === "t") { s += innerText(c); continue; }
          if (l === "tab" || l === "ptab") { s += "\t"; continue; }
          if (l === "br" || l === "cr") { s += "\n"; continue; }
          if (l === "noBreakHyphen") { s += "-"; continue; }
          if (l === "delText" || l === "instrText" || l === "del" || l === "moveFrom" || l === "pPr" || l === "rPr" || l === "txbxContent" || l === "sectPr") continue;
        }
        if (l === "AlternateContent") { const ch = c.elements()[0]; if (ch) walk(ch); continue; }
        walk(c);
      }
    };
    walk(n);
    return s;
  }
  tiny = (e: XElement) => this.isP(e) && !this.sect(e) && !this.hasObj(e) && !this.text(e).trim();
  /** 只承载分节符的空段落 */
  sectOnly = (e: XElement) => this.isP(e) && !!this.sect(e) && !this.hasObj(e) && !this.text(e).trim();

  topOf(n: XElement): XElement {
    let cur: XElement = n;
    while (cur.parent && cur.parent !== this.body && cur.parent.type === "el") cur = cur.parent as XElement;
    return cur;
  }
  next(e: XElement): XElement | undefined { const k = this.top(); return k[k.indexOf(e) + 1]; }
  prev(e: XElement): XElement | undefined { const k = this.top(); const i = k.indexOf(e); return i > 0 ? k[i - 1] : undefined; }

  /** 元素所在节的 sectPr（之后第一个带分节符的段落，或正文末尾的 sectPr） */
  sectFor(e: XElement): XElement {
    const k = this.top();
    for (let i = Math.max(0, k.indexOf(e)); i < k.length; i++) {
      const s = this.sect(k[i]);
      if (s) return s;
      if (k[i].local === "sectPr") return k[i];
    }
    return this.body.child(this.q("sectPr"))!;
  }
  colsOf(s?: XElement): number { return Number(s?.child(this.q("cols"))?.getAttr(this.q("num")) ?? 1) || 1; }
  geo(s?: XElement): Geo {
    const sp = s ?? this.body.child(this.q("sectPr"));
    const n = (el: XElement | undefined, a: string, d: number) => { const v = Number(el?.getAttr(this.q(a))); return Number.isFinite(v) ? v / TWIP_PER_PT : d; };
    const sz = sp?.child(this.q("pgSz")), mar = sp?.child(this.q("pgMar")), cols = sp?.child(this.q("cols"));
    return {
      pageW: n(sz, "w", 595), pageH: n(sz, "h", 842),
      top: n(mar, "top", 72), bottom: n(mar, "bottom", 72), left: n(mar, "left", 90), right: n(mar, "right", 90),
      cols: this.colsOf(sp), gap: n(cols, "space", 21),
    };
  }
  colWidth(s?: XElement): number {
    const g = this.geo(s);
    const body = g.pageW - g.left - g.right;
    return g.cols > 1 ? (body - g.gap * (g.cols - 1)) / g.cols : body;
  }

  // ---------------- 段落属性 ----------------

  pPr(p: XElement): XElement {
    let pPr = p.child(this.q("pPr"));
    if (!pPr) {
      [pPr] = parseFragment(`<${this.q("pPr")}/>`) as XElement[];
      p.children.unshift(pPr);
      pPr.parent = p;
      if (p.selfClosing) p.attrsDirty = true;
      markDirty(p);
    }
    return pPr;
  }
  prop(container: XElement, local: string, order: string[]): XElement {
    let el = container.child(this.q(local));
    if (!el) {
      [el] = parseFragment(`<${this.q(local)}/>`) as XElement[];
      const idx = orderedInsertIndex(container, local, order, this.w);
      container.children.splice(idx, 0, el);
      el.parent = container;
      if (container.selfClosing) container.attrsDirty = true;
      markDirty(el);
    }
    return el;
  }
  setPProp(p: XElement, local: string, attrs: Record<string, string | null>) {
    const el = this.prop(this.pPr(p), local, PPR_ORDER);
    for (const [k, v] of Object.entries(attrs)) {
      if (v === null) el.removeAttr(this.q(k));
      else el.setAttr(this.q(k), v);
    }
  }
  removePProp(p: XElement, local: string) {
    const el = p.child(this.q("pPr"))?.child(this.q(local));
    if (el) removeNode(el);
  }

  /** 段落的实际排版属性（直接格式 > 段落样式 > 文档默认） */
  eff(p: XElement): { before: number; after: number; lh: number; exact: boolean; sz: number; jc: string; lineTw: number; rule: string } {
    const pPr = p.child(this.q("pPr"));
    const st = this.styles.resolve(pPr?.child(this.q("pStyle"))?.getAttr(this.q("val")));
    const sp = pPr?.child(this.q("spacing"));
    const num = (v?: string) => (v !== undefined && v !== "" && Number.isFinite(Number(v)) ? Number(v) : undefined);
    const before = num(sp?.getAttr(this.q("before"))) ?? st.before ?? 0;
    const after = num(sp?.getAttr(this.q("after"))) ?? st.after ?? 0;
    const line = num(sp?.getAttr(this.q("line"))) ?? st.line ?? 240;
    const rule = sp?.getAttr(this.q("lineRule")) ?? (sp?.getAttr(this.q("line")) ? "auto" : st.lineRule ?? "auto");
    // 字号：段内文字最多的 run
    const count = new Map<number, number>();
    for (const r of p.descendants()) {
      if (r.local !== "r" || r.prefix !== this.w) continue;
      const sz = num(r.child(this.q("rPr"))?.child(this.q("sz"))?.getAttr(this.q("val")));
      const len = this.text(r).length;
      if (!len) continue;
      const k = sz ?? num(pPr?.child(this.q("rPr"))?.child(this.q("sz"))?.getAttr(this.q("val"))) ?? st.sz ?? 20;
      count.set(k, (count.get(k) ?? 0) + len);
    }
    const sz = ([...count.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? st.sz ?? 20) / 2;
    const exact = rule === "exact" || rule === "atLeast";
    const lh = exact ? line / TWIP_PER_PT : (line / 240) * sz * 1.15;
    const jc = pPr?.child(this.q("jc"))?.getAttr(this.q("val")) ?? st.jc ?? "left";
    return { before: before / TWIP_PER_PT, after: after / TWIP_PER_PT, lh, exact, sz, jc, lineTw: line, rule };
  }

  newBookmark(name: string): string {
    const id = ++this.bm;
    return `<${this.q("bookmarkStart")} ${this.q("id")}="${id}" ${this.q("name")}="${name}${id}"/><${this.q("bookmarkEnd")} ${this.q("id")}="${id}"/>`;
  }
  /** 段落末尾加一个隐藏书签（标记本工具做过的改动，再次运行时据此还原） */
  mark(p: XElement, name: string) {
    appendChild(p, parseFragment(this.newBookmark(name)));
  }
  findMark(e: XElement, prefix: string): XElement | undefined {
    for (const d of e.descendants()) if (d.local === "bookmarkStart" && (d.getAttr(this.q("name")) ?? "").startsWith(prefix)) return d;
    return undefined;
  }
  dropMark(start: XElement) {
    const id = start.getAttr(this.q("id"));
    for (const d of [...this.body.descendants()]) if (d.local === "bookmarkEnd" && d.getAttr(this.q("id")) === id) { removeNode(d); break; }
    removeNode(start);
  }

  clone(e: XElement): XElement { return parseFragment(serializeNode(e, this.xml.source))[0] as XElement; }

  // ---------------- 分节 ----------------

  /** 固定高度的空段（图表之后的留白） */
  spacerPara(heightPt: number): XElement {
    const q = this.q;
    return parseFragment(`<${q("p")}><${q("pPr")}><${q("spacing")} ${q("before")}="0" ${q("after")}="0" ${q("line")}="${Math.max(20, Math.round(heightPt * TWIP_PER_PT))}" ${q("lineRule")}="exact"/><${q("rPr")}><${q("sz")} ${q("val")}="2"/><${q("szCs")} ${q("val")}="2"/></${q("rPr")}></${q("pPr")}></${q("p")}>`)[0] as XElement;
  }

  /** 一个只承载分节符的空段落；cols=1 时为单栏节 */
  sectPara(base: XElement, opt: { single: boolean; keepFirst: boolean }): XElement {
    const s = this.clone(base);
    for (const k of [...s.elements()]) {
      if (k.local === "sectPrChange") removeNode(k);
      if (!opt.keepFirst && (k.local === "titlePg" || k.local === "pgNumType")) removeNode(k);
    }
    const type = this.prop(s, "type", SECTPR_ORDER);
    type.setAttr(this.q("val"), "continuous");
    if (opt.single) {
      const old = s.child(this.q("cols"));
      const [cols] = parseFragment(`<${this.q("cols")} ${this.q("space")}="425"/>`) as XElement[];
      if (old) { insertBefore(old, [cols]); removeNode(old); }
      else { const idx = orderedInsertIndex(s, "cols", SECTPR_ORDER, this.w); s.children.splice(idx, 0, cols); cols.parent = s; }
    }
    const q = this.q;
    const xml =
      `<${q("p")}><${q("pPr")}><${q("spacing")} ${q("before")}="0" ${q("after")}="0" ${q("line")}="20" ${q("lineRule")}="exact"/>` +
      `<${q("rPr")}><${q("sz")} ${q("val")}="2"/><${q("szCs")} ${q("val")}="2"/></${q("rPr")}>${serializeNode(s, this.xml.source)}</${q("pPr")}>` +
      `${this.newBookmark("_JLsect")}</${q("p")}>`;
    return parseFragment(xml)[0] as XElement;
  }

  /** "首页不同"与起始页码属于从第一页开始的那一节：删除 / 合并节时转给下一节 */
  transferFirst(from: XElement, to: XElement | undefined): void {
    if (!to || to === from) return;
    for (const l of ["titlePg", "pgNumType"]) {
      const el = from.child(this.q(l));
      if (!el || to.child(this.q(l))) continue;
      const c = this.clone(el);
      const idx = orderedInsertIndex(to, l, SECTPR_ORDER, this.w);
      to.children.splice(idx, 0, c);
      c.parent = to;
      markDirty(c);
    }
  }

  sameLayout(a: XElement, b: XElement): boolean {
    const key = (s: XElement) => ["cols", "pgSz", "pgMar"].map((l) => { const e = s.child(this.q(l)); return e ? serializeNode(e, this.xml.source).replace(/\s+/g, " ") : ""; }).join("|");
    return key(a) === key(b);
  }

  /** 去掉空节、合并相邻的同版式节 */
  cleanSections(): void {
    for (let changed = true; changed; ) {
      changed = false;
      const kids = this.top();
      // 节：以带分节符的段落结尾
      let start = 0;
      for (let i = 0; i < kids.length; i++) {
        const s = this.sect(kids[i]);
        if (!s) continue;
        const nodes = kids.slice(start, i);
        const nextS = this.sectFor(kids[i + 1] ?? kids[i]);
        // 空节（只有空段落）：删掉，首页标记转给下一节
        if (this.sectOnly(kids[i]) && nodes.every((n) => this.tiny(n))) {
          this.transferFirst(s, nextS);
          for (const n of nodes) removeNode(n);
          removeNode(kids[i]);
          changed = true;
          break;
        }
        // 与下一节版式相同：合并
        if (nextS && nextS !== s && this.sameLayout(s, nextS) && this.sectOnly(kids[i])) {
          this.transferFirst(s, nextS);
          removeNode(kids[i]);
          changed = true;
          break;
        }
        start = i + 1;
      }
    }
  }

  // ---------------- 拆段 / 合段 ----------------

  /** 一个子元素贡献的文字长度（与 text() 的规则一致；text() 只看后代，所以 w:t 等自身要单独算） */
  private childLen(c: XNode): number {
    if (c.type !== "el") return 0;
    if (c.prefix === this.w) {
      if (c.local === "t") return innerText(c).length;
      if (["tab", "ptab", "br", "cr", "noBreakHyphen"].includes(c.local)) return 1;
      if (["delText", "instrText", "del", "moveFrom", "pPr", "rPr", "txbxContent", "sectPr"].includes(c.local)) return 0;
    }
    return this.text(c).length;
  }

  /** 在字符偏移处把段落拆成两段，返回后半段；偏移在段首/段尾时不拆，返回 null */
  splitPara(p: XElement, off: number, plain = false): XElement | null {
    if (off <= 0) return null;
    const total = this.text(p).length;
    if (off >= total) return null;
    let pos = 0, at = -1;
    for (let i = 0; i < p.children.length; i++) {
      const c = p.children[i];
      if (c.type !== "el" || c.local === "pPr") continue;
      const len = this.childLen(c);
      if (pos + len > off) {
        if (c.local === "r" && c.prefix === this.w && off > pos) {
          const parts = this.splitRun(c, off - pos);
          if (parts) { insertBefore(c, parts); removeNode(c); at = p.children.indexOf(parts[1]); }
          else at = i;
        } else at = i;
        break;
      }
      pos += len;
      if (pos === off) { at = i + 1; break; }
    }
    if (at < 0 || at >= p.children.length) return null;
    const moved = p.children.slice(at);
    // 后半段：段落属性去掉分节、编号、修订记录；不缩进、无段前距
    const pPr = p.child(this.q("pPr"));
    const tail = parseFragment(`<${this.q("p")}>${pPr ? serializeNode(pPr, this.xml.source) : ""}</${this.q("p")}>`)[0] as XElement;
    for (const m of moved) removeNode(m);
    appendChild(tail, moved);
    const tp = tail.child(this.q("pPr"));
    if (tp) for (const k of [...tp.elements()]) if (k.local === "numPr" || k.local === "pPrChange") removeNode(k);
    this.setPProp(tail, "spacing", { before: "0", beforeLines: null, beforeAutospacing: null });
    const ind = tail.child(this.q("pPr"))?.child(this.q("ind"));
    if (ind) { ind.removeAttr(this.q("firstLine")); ind.removeAttr(this.q("firstLineChars")); ind.removeAttr(this.q("hanging")); ind.removeAttr(this.q("hangingChars")); }
    if (plain) { this.removePProp(p, "sectPr"); insertAfter(p, [tail]); return tail; }
    // 前半段：两端对齐的段落末行也拉满（分散对齐），看起来与跨栏连续的段落一样
    const e = this.eff(p);
    const orig = p.child(this.q("pPr"))?.child(this.q("jc"))?.getAttr(this.q("val")) ?? "none";
    if (e.jc === "both" || e.jc === "justify") this.setPProp(p, "jc", { val: "distribute" });
    this.removePProp(p, "sectPr");
    this.setPProp(p, "spacing", { after: "0", afterLines: null, afterAutospacing: null });
    this.mark(p, `_JLsplit_${orig}_`);
    insertAfter(p, [tail]);
    return tail;
  }

  private splitRun(r: XElement, off: number): [XElement, XElement] | null {
    const src = this.xml.source;
    const rPr = r.child(this.q("rPr"));
    const rPrXml = rPr ? serializeNode(rPr, src) : "";
    let a = "", b = "", pos = 0, done = false;
    for (const c of r.children) {
      if (c.type !== "el") continue;
      if (c.local === "rPr") continue;
      const len = this.childLen(c);
      if (done || pos >= off) { b += serializeNode(c, src); continue; }
      if (c.local === "t" && pos + len > off) {
        const t = innerText(c);
        const k = off - pos;
        a += `<${this.q("t")} xml:space="preserve">${encodeText(t.slice(0, k))}</${this.q("t")}>`;
        b += `<${this.q("t")} xml:space="preserve">${encodeText(t.slice(k))}</${this.q("t")}>`;
        done = true;
        pos += len;
        continue;
      }
      a += serializeNode(c, src);
      pos += len;
      if (pos >= off) done = true;
    }
    if (!a || !b) return null;
    const mk = (inner: string) => parseFragment(`<${this.q("r")}>${rPrXml}${inner}</${this.q("r")}>`)[0] as XElement;
    return [mk(a), mk(b)];
  }

  /** 还原上一次拆开的段落；返回 后半段 → 前半段 的对应（引用位置据此换算） */
  rejoinSplits(): Map<XElement, XElement> {
    const map = new Map<XElement, XElement>();
    for (const head of this.top()) {
      if (!this.isP(head)) continue;
      const m = this.findMark(head, "_JLsplit_");
      if (!m) continue;
      const orig = /^_JLsplit_(.*?)_\d+$/.exec(m.getAttr(this.q("name")) ?? "")?.[1] ?? "none";
      this.dropMark(m);
      const tail = this.next(head);
      if (orig === "none") this.removePProp(head, "jc"); else this.setPProp(head, "jc", { val: orig });
      if (!tail || !this.isP(tail) || this.sect(tail)) continue;
      // 恢复段后距：取后半段的（后半段保留了原段落的段后距）
      const tsp = tail.child(this.q("pPr"))?.child(this.q("spacing"));
      const after = tsp?.getAttr(this.q("after"));
      if (after !== undefined) this.setPProp(head, "spacing", { after });
      const moved = tail.children.filter((c) => !(c.type === "el" && c.local === "pPr"));
      for (const c of moved) removeNode(c);
      appendChild(head, moved);
      removeNode(tail);
      map.set(tail, head);
    }
    return map;
  }

  // ---------------- 修复：并进正文段落末尾的表题 ----------------

  /** 早期转换结果中，表题（TABLE V: …）可能被并进了表格上方那段正文的末尾：拆出来自成一段 */
  repairCaptions(): number {
    let n = 0;
    for (const t of this.top()) {
      if (!this.isTbl(t)) continue;
      let prev = this.prev(t);
      while (prev && this.tiny(prev)) prev = this.prev(prev);
      if (!prev || !this.isP(prev) || this.sect(prev)) continue;
      const text = this.text(prev);
      if (TAB_CAP.test(text)) continue;
      const re = /(TABLE\s+[IVXLC\d]+\s*[:.：]|Table\s+\d+\s*[:.：]|表\s*\d+[\s:：])/g;
      let at = -1;
      for (const m of text.matchAll(re)) if (m.index! >= 20 && /[\s.;:，。；]$/.test(text.slice(0, m.index!))) at = m.index!;
      if (at < 0) continue;
      const cap = this.splitPara(prev, at, true);
      if (cap) { this.setPProp(cap, "jc", { val: "center" }); n++; }
    }
    return n;
  }

  /** Word / LibreOffice 兼容选项：连续分节前不平衡分栏（先排满左栏再排右栏） */
  noColumnBalance(): void {
    const rels = parseRels(this.pkg.has(relsPathFor(this.main)) ? this.pkg.readText(relsPathFor(this.main)) : null);
    const st = rels.find((r) => r.type.endsWith("/settings"));
    if (!st) return;
    const part = resolveTarget(this.main, st.target);
    if (!this.pkg.has(part)) return;
    let xml = this.pkg.readText(part);
    if (/<\w+:noColumnBalance\b/.test(xml)) return;
    const pre = /<(\w+):settings\b/.exec(xml)?.[1] ?? "w";
    const el = `<${pre}:noColumnBalance/>`;
    if (new RegExp(`<${pre}:compat\\s*/>`).test(xml)) xml = xml.replace(new RegExp(`<${pre}:compat\\s*/>`), `<${pre}:compat>${el}</${pre}:compat>`);
    else if (xml.includes(`<${pre}:compat>`)) {
      // 放在 compat 中 noColumnBalance 之前的那几个元素之后
      const m = new RegExp(`<${pre}:compat>((?:<${pre}:(?:useSingleBorderforContiguousCells|wpJustification|noTabHangInd|noLeading|spaceForUL)\\b[^>]*/>)*)`).exec(xml)!;
      xml = xml.replace(m[0], m[0] + el);
    } else {
      const after = ["docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"];
      const re = new RegExp(`<${pre}:(${after.join("|")})\\b`);
      const m = re.exec(xml);
      const ins = `<${pre}:compat>${el}</${pre}:compat>`;
      xml = m ? xml.slice(0, m.index) + ins + xml.slice(m.index) : xml.replace(`</${pre}:settings>`, `${ins}</${pre}:settings>`);
    }
    this.pkg.write(part, Buffer.from(xml, "utf8"));
  }

  // ---------------- 浮动体 ----------------

  detectFloats(mode: "auto" | "wide"): Float[] {
    const kids = this.top();
    const out: Float[] = [];
    const used = new Set<XElement>();
    const isFigPara = (e: XElement) => this.isP(e) && !this.sect(e) && this.inlineDrawings(e).length > 0 && this.text(e).replace(/\s/g, "").length <= 3;
    const capOf = (e: XElement | undefined, re: RegExp) => (e && this.isP(e) && !this.sect(e) && !used.has(e) && re.test(this.text(e)) && this.text(e).length < 700 ? e : undefined);
    const step = (i: number, d: number) => { let j = i + d; while (kids[j] && (this.tiny(kids[j]) || this.sectOnly(kids[j]))) j += d; return j; };
    let id = 0;
    for (let i = 0; i < kids.length; i++) {
      const e = kids[i];
      if (used.has(e)) continue;
      let kind: Float["kind"] | null = null;
      if (this.isTbl(e)) kind = "table";
      else if (isFigPara(e)) kind = "figure";
      if (!kind) continue;
      const re = kind === "figure" ? FIG_CAP : TAB_CAP;
      const pi = step(i, -1), ni = step(i, 1);
      let lo = i, hi = i;
      // 并排 / 上下排的多张子图：连续的图段落
      if (kind === "figure") while (kids[hi + 1] && (isFigPara(kids[hi + 1]) || (this.tiny(kids[hi + 1]) && kids[hi + 2] && isFigPara(kids[hi + 2])))) hi++;
      const after = kind === "figure" ? step(hi, 1) : ni;
      let cap = kind === "figure" ? capOf(kids[after], re) ?? capOf(kids[pi], re) : capOf(kids[pi], re) ?? capOf(kids[ni], re);
      // 以图片形式存在的表格（PDF 中表格文字是矢量轮廓，转换时整体渲染成图）：表题在上
      if (!cap && kind === "figure") { cap = capOf(kids[pi], TAB_CAP) ?? capOf(kids[after], TAB_CAP); if (cap) kind = "table"; }
      if (!cap) continue;
      const ci = kids.indexOf(cap);
      lo = Math.min(lo, ci); hi = Math.max(hi, ci);
      const nodes = kids.slice(lo, hi + 1).filter((n) => !this.sect(n));
      if (nodes.some((n) => used.has(n) || this.sect(n))) continue;
      nodes.forEach((n) => used.add(n));
      // 宽度：图取最宽的图片，表取表宽
      let width = 0;
      if (!nodes.some((n) => this.isTbl(n))) {
        for (const n of nodes) {
          const ds = this.inlineDrawings(n);
          const rowW = ds.reduce((a, d) => a + this.extent(d).w, 0);
          width = Math.max(width, rowW);
        }
      } else {
        const tbl = nodes.find((n) => this.isTbl(n))!;
        const tw = tbl.child(this.q("tblPr"))?.child(this.q("tblW"));
        const grid = tbl.child(this.q("tblGrid"))?.childrenNamed(this.q("gridCol")).reduce((a, g) => a + Number(g.getAttr(this.q("w")) ?? 0), 0) ?? 0;
        width = (tw?.getAttr(this.q("type")) === "dxa" ? Number(tw.getAttr(this.q("w"))) : grid) / TWIP_PER_PT;
      }
      // 所在的正文版式：浮动体单独占一个单栏节时（转换结果常见），看相邻的正文节
      // 浮动体会被放回多栏正文里：只要文档有多栏节、它比一栏宽，就是通栏浮动体
      let ctx = this.contextSect(nodes[0], nodes);
      if (this.colsOf(ctx) < 2) {
        const multi = [...this.body.descendants()].find((d) => d.local === "sectPr" && d.prefix === this.w && this.colsOf(d) > 1);
        if (multi) ctx = multi;
      }
      const colW = this.colWidth(ctx);
      const wide = this.colsOf(ctx) > 1 && width > colW * 1.02;
      if (mode === "wide" && !wide) continue;
      // 上一次运行留下的引用位置
      const fm = nodes.map((n) => this.findMark(n, "_JLfloat")).find(Boolean);
      let anchor: XElement | null = null;
      if (fm) {
        const n = /_JLfloat(\d+)_/.exec(fm.getAttr(this.q("name")) ?? "")?.[1];
        for (const d of this.body.descendants()) if (d.local === "bookmarkStart" && (d.getAttr(this.q("name")) ?? "").startsWith(`_JLanchor${n}_`)) { anchor = this.topOf(d); this.dropMark(d); break; }
        this.dropMark(fm);
      }
      if (!anchor) {
        for (let j = lo - 1; j >= 0; j--) { const k = kids[j]; if (!this.tiny(k) && !this.sectOnly(k) && !used.has(k)) { anchor = k; break; } }
      }
      out.push({ kind, nodes, caption: this.text(cap).trim().slice(0, 80), wide, colW, anchor, id: ++id });
    }
    return out;
  }

  /** 浮动体所处的正文节：若它所在的节里只有它自己（单独成节），取后面（或前面）第一个有正文的节 */
  contextSect(first: XElement, own: XElement[]): XElement {
    const kids = this.top();
    const own2 = new Set(own);
    const sectRange = (s: XElement) => {
      const end = kids.findIndex((k) => this.sect(k) === s || k === s);
      let start = end - 1;
      while (start >= 0 && !this.sect(kids[start])) start--;
      return kids.slice(start + 1, end < 0 ? kids.length : end + 1);
    };
    const s = this.sectFor(first);
    const content = (ss: XElement) => sectRange(ss).some((n) => !own2.has(n) && !this.tiny(n) && !this.sectOnly(n) && !this.isFloatish(n));
    if (this.colsOf(s) > 1 || content(s)) return s;
    // 往后找
    const end = kids.findIndex((k) => this.sect(k) === s);
    for (let i = end + 1; i < kids.length; i++) {
      const ss = this.sect(kids[i]) ?? (kids[i].local === "sectPr" ? kids[i] : undefined);
      if (ss && content(ss)) return ss;
    }
    const fin = this.body.child(this.q("sectPr"));
    return fin && content(fin) ? fin : s;
  }
  private isFloatish(n: XElement): boolean {
    if (this.isTbl(n)) return true;
    if (this.isP(n) && this.inlineDrawings(n).length && this.text(n).replace(/\s/g, "").length <= 3) return true;
    const t = this.isP(n) ? this.text(n) : "";
    return FIG_CAP.test(t) || TAB_CAP.test(t);
  }

  /** 浮动体从正文中取出（连同紧挨着的空段落），返回其引用位置 */
  detach(f: Float): void {
    const first = f.nodes[0], last = f.nodes[f.nodes.length - 1];
    const around: XElement[] = [];
    for (let p = this.prev(first); p && this.tiny(p); p = this.prev(p)) around.push(p);
    for (let n = this.next(last); n && this.tiny(n); n = this.next(n)) around.push(n);
    let follow = this.next(around.length ? around[around.length - 1] : last);
    if (follow && around.includes(follow)) follow = undefined;
    for (const n of f.nodes) removeNode(n);
    for (const n of around) removeNode(n);
    // 紧跟在浮动体后面的正文段落：原来的段前距是与图表之间的空白，图表移走后不再需要
    if (follow && this.isP(follow) && !this.sect(follow) && this.text(follow).trim()) this.setPProp(follow, "spacing", { before: "0", beforeLines: null, beforeAutospacing: null });
  }

  /** 浮动体内部：保持在一起、不跨页；宽浮动体居中、去掉左缩进 */
  keepTogether(f: Float): number {
    let n = 0;
    const capFirst = f.nodes[0] && !this.isTbl(f.nodes[0]) && !this.inlineDrawings(f.nodes[0]).length;
    f.nodes.forEach((node, i) => {
      const lastNode = i === f.nodes.length - 1;
      if (this.isTbl(node)) {
        const rows = node.childrenNamed(this.q("tr"));
        rows.forEach((tr, ri) => {
          let trPr = tr.child(this.q("trPr"));
          if (!trPr) { [trPr] = parseFragment(`<${this.q("trPr")}/>`) as XElement[]; const tcIdx = tr.children.findIndex((c) => c.type === "el" && c.local !== "tblPrEx"); tr.children.splice(Math.max(0, tcIdx), 0, trPr); trPr.parent = tr; markDirty(trPr); }
          this.prop(trPr, "cantSplit", TRPR_ORDER);
          if (ri < rows.length - 1 || !lastNode) for (const p of tr.descendants()) if (this.isP(p)) this.setPProp(p, "keepNext", {});
          n++;
        });
        if (f.wide) {
          const tblPr = node.child(this.q("tblPr"));
          const jc = tblPr?.child(this.q("jc"));
          if (tblPr) {
            if (jc) jc.setAttr(this.q("val"), "center");
            else { const [j] = parseFragment(`<${this.q("jc")} ${this.q("val")}="center"/>`) as XElement[]; const idx = orderedInsertIndex(tblPr, "jc", TBLPR_ORDER, this.w); tblPr.children.splice(idx, 0, j); j.parent = tblPr; markDirty(j); }
            const ind = tblPr.child(this.q("tblInd"));
            if (ind) removeNode(ind);
          }
        }
      } else {
        if (!lastNode) this.setPProp(node, "keepNext", {});
        this.setPProp(node, "keepLines", {});
        if (f.wide && this.inlineDrawings(node).length) {
          this.setPProp(node, "jc", { val: "center" });
          this.setPProp(node, "ind", { left: "0", right: "0", firstLine: "0", hanging: null });
        }
        n++;
      }
    });
    // 浮动体的上下间距：顶部不留空，底部空一行（由网格对齐时设定），题注在上时题注与表之间不变
    if (capFirst) { /* 表题在上：保持 */ }
    return n;
  }

  /** 把浮动体插到 before 之前；宽浮动体单独成一个单栏节 */
  insertFloat(f: Float, before: XElement): void {
    if (!f.wide) { insertBefore(before, f.nodes); return; }
    const S = this.sectFor(before);
    const prev = this.prev(before);
    const parts: XElement[] = [];
    if (prev && !this.sect(prev)) {
      const p1 = this.sectPara(S, { single: false, keepFirst: true });
      parts.push(p1);
      // 首页标记属于从第一页开始的那一节（前半节），后半节不再是"首页"
      for (const l of ["titlePg", "pgNumType"]) { const el = S.child(this.q(l)); if (el) removeNode(el); }
    }
    parts.push(...f.nodes);
    parts.push(this.sectPara(S, { single: true, keepFirst: false }));
    insertBefore(before, parts);
  }

  /** 记下引用位置（再次运行时浮动体回到这里之后重新定位） */
  markAnchor(f: Float): void {
    if (!f.anchor || !f.anchor.parent) return;
    const n = ++this.bm;
    const target = this.isP(f.anchor) ? f.anchor : [...f.anchor.descendants()].filter((d) => this.isP(d)).pop();
    const host = f.nodes.find((x) => this.isP(x));
    if (!target || !host) return;
    appendChild(target, parseFragment(`<${this.q("bookmarkStart")} ${this.q("id")}="${n}" ${this.q("name")}="_JLanchor${n}_"/><${this.q("bookmarkEnd")} ${this.q("id")}="${n}"/>`));
    const m = ++this.bm;
    appendChild(host, parseFragment(`<${this.q("bookmarkStart")} ${this.q("id")}="${m}" ${this.q("name")}="_JLfloat${n}_"/><${this.q("bookmarkEnd")} ${this.q("id")}="${m}"/>`));
  }

  /** 正文中位于多栏节里的顶层段落 */
  columnParas(): XElement[] {
    const out: XElement[] = [];
    let buf: XElement[] = [];
    for (const k of this.top()) {
      if (k.local === "sectPr") break;
      buf.push(k);
      const s = this.sect(k);
      if (s) { if (this.colsOf(s) > 1) out.push(...buf.filter((x) => this.isP(x))); buf = []; }
    }
    const fin = this.body.child(this.q("sectPr"));
    if (buf.length && this.colsOf(fin) > 1) out.push(...buf.filter((x) => this.isP(x)));
    return out;
  }
}

// ---------------------------------------------------------------------------
// 版面映射：PDF 中的词 ↔ 文档中的段落与字符位置
// ---------------------------------------------------------------------------

interface PTok { t: string; page: number; col: "L" | "R" | "F"; line: number; top: number; bottom: number; zone: number }
interface DTok { t: string; p: XElement; off: number }

interface LayoutMap {
  pages: PageModel[];
  ptoks: PTok[];
  dtoks: DTok[];
  /** PDF 词 → 文档词 */
  p2d: Int32Array;
  /** 文档词 → PDF 词 */
  d2p: Int32Array;
  geo: Geo;
  /** 每页每栏的内容底边（pt） */
  colBottoms: Array<{ L: number; R: number; F: number; any: number; hasCols: boolean }>;
}

const TOKEN = /[\p{L}\p{N}]+/gu;

function tokenize(s: string): Array<{ t: string; off: number }> {
  const out: Array<{ t: string; off: number }> = [];
  for (const m of s.matchAll(TOKEN)) out.push({ t: m[0].toLowerCase(), off: m.index! });
  return out;
}

async function buildMap(jd: JDoc, render: (b: Buffer) => Promise<Buffer>): Promise<LayoutMap> {
  const pdf = await render(jd.serialize());
  const pages = await extractPdf(pdf);
  // 分栏几何取文档中的多栏节（最后一节可能是单栏的）
  const multi = [...jd.body.descendants()].find((d) => d.local === "sectPr" && d.prefix === jd.w && jd.colsOf(d) > 1);
  const geo = jd.geo(multi);
  const mid = (geo.left + (geo.pageW - geo.right)) / 2;
  const bodyTop = geo.top - 3, bodyBot = geo.pageH - geo.bottom + 3;
  const ptoks: PTok[] = [];
  const colBottoms: LayoutMap["colBottoms"] = [];
  let lineNo = 0;
  pages.forEach((pg, pi) => {
    const raw = buildLines(pg.spans.filter((s) => (s.top + s.bottom) / 2 > bodyTop && (s.top + s.bottom) / 2 < bodyBot));
    // 左右栏同一基线的文字可能被并成一行：在栏间距处拆开
    const lines: Array<Line & { col: "L" | "R" | "F" }> = [];
    for (const l of raw) {
      const L = l.spans.filter((s) => (s.x0 + s.x1) / 2 < mid), R = l.spans.filter((s) => (s.x0 + s.x1) / 2 >= mid);
      const crossing = l.spans.some((s) => s.x0 < mid - 2 && s.x1 > mid + 2);
      if (!crossing && L.length && R.length) {
        lines.push({ ...l, spans: L, x0: Math.min(...L.map((s) => s.x0)), x1: Math.max(...L.map((s) => s.x1)), col: "L" });
        lines.push({ ...l, spans: R, x0: Math.min(...R.map((s) => s.x0)), x1: Math.max(...R.map((s) => s.x1)), col: "R" });
      } else lines.push({ ...l, col: crossing || geo.cols < 2 ? (geo.cols < 2 ? "L" : "F") : L.length ? "L" : "R" });
    }
    lines.sort((a, b) => a.top - b.top);
    // 通栏行把页面切成若干区；区内先左栏后右栏
    const zones: Array<{ cols: boolean; lines: typeof lines }> = [];
    for (const l of lines) {
      const cols = l.col !== "F";
      const z = zones[zones.length - 1];
      if (z && z.cols === cols) z.lines.push(l);
      else zones.push({ cols, lines: [l] });
    }
    const cb = { L: 0, R: 0, F: 0, any: 0, hasCols: false };
    zones.forEach((z, zi) => {
      const ordered = z.cols ? [...z.lines.filter((l) => l.col === "L"), ...z.lines.filter((l) => l.col === "R")] : z.lines;
      for (const l of ordered) {
        const text = l.spans.map((s) => s.text).join(" ");
        lineNo++;
        for (const t of tokenize(text)) ptoks.push({ t: t.t, page: pi + 1, col: l.col, line: lineNo, top: l.top, bottom: l.bottom, zone: zi });
        // 行尾连字符断词：与下一行开头合并为一个词
        if (/[A-Za-z]-$/.test(text.trim())) (ptoks[ptoks.length - 1] as any).hy = true;
        cb[l.col] = Math.max(cb[l.col], l.bottom);
        cb.any = Math.max(cb.any, l.bottom);
        if (z.cols) cb.hasCols = true;
      }
    });
    for (const im of pg.images) if (im.y1 < bodyBot && im.y0 > bodyTop) cb.any = Math.max(cb.any, im.y1);
    colBottoms.push(cb);
  });
  // 合并断词
  const merged: PTok[] = [];
  for (let i = 0; i < ptoks.length; i++) {
    const t = ptoks[i];
    if ((t as any).hy && ptoks[i + 1] && /^[a-z]/.test(ptoks[i + 1].t)) { merged.push({ ...t, t: t.t + ptoks[i + 1].t }); i++; }
    else merged.push(t);
  }
  // 文档词
  const dtoks: DTok[] = [];
  for (const d of jd.body.descendants()) {
    if (!jd.isP(d) || d.closest(jd.q("txbxContent"))) continue;
    for (const t of tokenize(jd.text(d))) dtoks.push({ t: t.t, p: d, off: t.off });
  }
  const p2d = align(dtoks.map((x) => x.t), merged.map((x) => x.t));
  const d2p = new Int32Array(dtoks.length).fill(-1);
  p2d.forEach((d, i) => { if (d >= 0 && d2p[d] < 0) d2p[d] = i; });
  // 栏底：只算正文里的文字（页脚里的注释不算）与图片（公式、图）
  const bottoms = pages.map(() => ({ L: 0, R: 0, F: 0, any: 0, hasCols: false }));
  merged.forEach((t, i) => {
    if (p2d[i] < 0) return;
    const b = bottoms[t.page - 1];
    b[t.col] = Math.max(b[t.col], t.bottom);
    b.any = Math.max(b.any, t.bottom);
    if (t.col !== "F") b.hasCols = true;
  });
  pages.forEach((pg, pi) => {
    const b = bottoms[pi];
    for (const im of pg.images) {
      if (im.y1 > bodyBot || im.y0 < bodyTop) continue;
      const c = geo.cols < 2 ? "L" : im.x0 < mid - 2 && im.x1 > mid + 2 ? "F" : (im.x0 + im.x1) / 2 < mid ? "L" : "R";
      b[c] = Math.max(b[c], im.y1);
      b.any = Math.max(b.any, im.y1);
    }
  });
  void colBottoms;
  return { pages, ptoks: merged, dtoks, p2d, d2p, geo, colBottoms: bottoms };
}

/** 词序列对齐：相同就前进，不同时用后面的三连词重新同步 */
function align(doc: string[], pdf: string[]): Int32Array {
  const res = new Int32Array(pdf.length).fill(-1);
  const key = (a: string[], i: number) => `${a[i]} ${a[i + 1]} ${a[i + 2]}`;
  const idx = new Map<string, number[]>();
  for (let i = 0; i + 2 < doc.length; i++) {
    const k = key(doc, i);
    const arr = idx.get(k);
    if (arr) arr.push(i); else idx.set(k, [i]);
  }
  let i = 0, j = 0;
  while (i < pdf.length && j < doc.length) {
    if (pdf[i] === doc[j]) { res[i] = j; i++; j++; continue; }
    let found = false;
    for (let di = 0; di < 80 && i + di + 2 < pdf.length; di++) {
      const c = idx.get(key(pdf, i + di));
      if (!c) continue;
      const hit = c.find((x) => x >= j && x - j < 1500);
      if (hit !== undefined) { i += di; j = hit; found = true; break; }
    }
    if (!found) i++;
  }
  return res;
}

/** 某元素（段落或表格）最后一个词在 PDF 中的位置 */
function posAfter(map: LayoutMap, jd: JDoc, el: XElement): PTok | null {
  const paras = jd.isP(el) ? [el] : [...el.descendants()].filter((d) => jd.isP(d));
  const set = new Set(paras);
  let last = -1;
  map.dtoks.forEach((t, i) => { if (set.has(t.p)) last = i; });
  if (last < 0) {
    // 没有文字（图片段落）：取它之前最近的词
    const order = [...jd.body.descendants()].filter((d) => jd.isP(d));
    const pos = order.indexOf(paras[paras.length - 1] ?? el);
    for (let k = map.dtoks.length - 1; k >= 0; k--) if (order.indexOf(map.dtoks[k].p) <= pos) { last = k; break; }
  }
  for (let k = last; k >= 0; k--) if (map.d2p[k] >= 0) return map.ptoks[map.d2p[k]];
  return null;
}

/** 某页某栏（多栏区）第一个词对应的文档位置 */
function firstOf(map: LayoutMap, page: number, col: "L" | "R" | "any", skip?: (p: XElement) => boolean): { p: XElement; off: number; tok: PTok } | null {
  for (let i = 0; i < map.ptoks.length; i++) {
    const t = map.ptoks[i];
    if (t.page < page) continue;
    if (t.page > page) break;
    if (col !== "any" && t.col !== col) continue;
    if (col === "any" && t.col === "F") continue;
    const d = map.p2d[i];
    if (d >= 0 && !skip?.(map.dtoks[d].p)) return { p: map.dtoks[d].p, off: map.dtoks[d].off, tok: t };
  }
  return null;
}

// ---------------------------------------------------------------------------
// 主流程
// ---------------------------------------------------------------------------

export async function applyJournalLayout(buf: Buffer, opts: JournalOptions = {}): Promise<{ data: Buffer; report: JournalReport }> {
  const jd = new JDoc(buf);
  const mode = opts.floats ?? "auto";
  const report: JournalReport = { floats: [], pitch: null, gridAdjusted: 0, keepTogether: 0, pages: null, whitespacePages: [], unevenPages: [], renders: 0, notes: [] };
  const render = opts.render;
  const doRender = async () => { report.renders++; return buildMap(jd, render!); };

  // 0) 修复早期转换中并进正文的表题；分栏不再在分节前平衡（页顶通栏图表之前的那一页排满）
  const repaired = jd.repairCaptions();
  if (repaired) report.notes.push(`把 ${repaired} 个并进正文段落末尾的表题拆成了独立段落。`);
  jd.noColumnBalance();
  // 1) 还原上一次的排版：取出浮动体 → 去掉插入的分节 → 合并拆开的段落
  const floats = mode === "none" ? [] : jd.detectFloats(mode);
  // 上一次留下、这次没有用到的引用标记
  for (const d of [...jd.body.descendants()]) {
    if (d.local === "bookmarkStart" && /^_JL(float|anchor)\d+_/.test(d.getAttr(jd.q("name")) ?? "") && d.parent) jd.dropMark(d);
  }
  for (const f of floats) jd.detach(f);
  for (const k of [...jd.top()]) {
    if (jd.sectOnly(k) && jd.findMark(k, "_JLsect")) {
      const nx = jd.next(k);
      jd.transferFirst(jd.sect(k)!, nx ? jd.sectFor(nx) : undefined);
      removeNode(k);
    }
  }
  jd.cleanSections();
  const rejoined = jd.rejoinSplits();
  for (const f of floats) if (f.anchor && rejoined.has(f.anchor)) f.anchor = rejoined.get(f.anchor)!;
  jd.cleanSections();

  // 2) 浮动体内部不拆开
  for (const f of floats) report.keepTogether += jd.keepTogether(f);
  // 其余表格（没有题注的）也不跨页断行
  for (const t of jd.top()) if (jd.isTbl(t)) for (const tr of t.childrenNamed(jd.q("tr"))) {
    const trPr = tr.child(jd.q("trPr"));
    if (trPr) jd.prop(trPr, "cantSplit", TRPR_ORDER);
  }

  // 3) 关闭孤行控制、正文固定行距、基线网格
  const colParas = jd.columnParas();
  const bodyParas = colParas.filter((p) => jd.text(p).trim().length > 60 && !jd.hasObj(p));
  const lhCount = new Map<number, number>();
  for (const p of bodyParas) {
    const e = jd.eff(p);
    const k = Math.round(e.lh * 20) / 20;
    lhCount.set(k, (lhCount.get(k) ?? 0) + jd.text(p).length);
  }
  const pitch = [...lhCount.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
  report.pitch = pitch;
  if (opts.widowOff !== false) for (const p of colParas) jd.setPProp(p, "widowControl", { val: "0" });
  if (pitch && opts.exactLines !== false) {
    // 正文字号、单倍行距的段落改为固定行距 = 网格行距
    const bodySz = (() => { const c = new Map<number, number>(); for (const p of bodyParas) { const e = jd.eff(p); c.set(e.sz, (c.get(e.sz) ?? 0) + 1); } return [...c.entries()].sort((a, b) => b[1] - a[1])[0]?.[0]; })();
    for (const p of colParas) {
      const e = jd.eff(p);
      if (!e.exact && Math.abs(e.sz - (bodySz ?? e.sz)) < 0.3 && jd.text(p).trim()) jd.setPProp(p, "spacing", { line: String(Math.round(pitch * TWIP_PER_PT)), lineRule: "exact" });
    }
  }

  // 测量：每段实际占几行（需要渲染）
  let lineCount: Map<XElement, number> | null = null;
  if (render && pitch && opts.grid !== false) {
    try {
      const m = await doRender();
      lineCount = new Map();
      const seen = new Map<XElement, Set<number>>();
      m.p2d.forEach((d, i) => { if (d < 0) return; const p = m.dtoks[d].p; if (!seen.has(p)) seen.set(p, new Set()); seen.get(p)!.add(m.ptoks[i].line); });
      for (const [p, s] of seen) lineCount.set(p, s.size);
    } catch (e: any) {
      report.notes.push(`渲染失败（${String(e?.message ?? e).slice(0, 120)}），行数按估算。`);
    }
  }

  if (pitch && opts.grid !== false) {
    const P = pitch;
    const geo = jd.geo();
    const colW = jd.colWidth();
    const estLines = (p: XElement, sz: number) => {
      const t = jd.text(p);
      const cjk = (t.match(/[⺀-鿿＀-￯]/g) ?? []).length;
      const w = (t.length - cjk) * 0.48 * sz + cjk * sz;
      return Math.max(1, Math.ceil(w / (colW - 2)));
    };
    const setSpacing = (p: XElement, before: number, after: number) => {
      jd.setPProp(p, "spacing", {
        before: String(Math.max(0, Math.round(before * TWIP_PER_PT))), after: String(Math.max(0, Math.round(after * TWIP_PER_PT))),
        beforeLines: null, afterLines: null, beforeAutospacing: null, afterAutospacing: null,
      });
      report.gridAdjusted++;
    };
    for (const p of colParas) {
      if (jd.sect(p)) continue;
      const e = jd.eff(p);
      const text = jd.text(p).trim();
      const imgs = jd.inlineDrawings(p);
      if (!text && !imgs.length) {
        if (jd.hasObj(p)) continue;
        // 空段（占位空白）：高度凑整；凑整后为 0 的删掉
        const h = e.before + e.lh + e.after;
        const k = Math.round(h / P);
        if (k <= 0) { removeNode(p); continue; }
        jd.setPProp(p, "spacing", { before: "0", after: "0", line: String(Math.round(k * P * TWIP_PER_PT)), lineRule: "exact" });
        report.gridAdjusted++;
        continue;
      }
      if (imgs.length && text.replace(/\s/g, "").length <= 3) {
        // 公式 / 图片段落：图片高度 + 段距凑成整数行
        const ih = Math.max(...imgs.map((d) => jd.extent(d).h));
        const h = e.before + ih + e.after;
        const H = Math.ceil((h - 0.5) / P) * P;
        const extra = H - h;
        setSpacing(p, e.before + extra / 2, e.after + extra / 2);
        continue;
      }
      const n = lineCount?.get(p) ?? estLines(p, e.sz);
      const content = n * e.lh;
      if (Math.abs(e.lh - P) < 0.3) {
        const k = Math.round((e.before + e.after) / P);
        if (Math.abs(e.before + e.after - k * P) < 0.05) continue;
        // 段距凑成整数行：优先保留段后距
        const total = k * P;
        const after = Math.min(e.after, total);
        setSpacing(p, total - after, after);
      } else {
        const h = e.before + content + e.after;
        let H = Math.round(h / P) * P;
        if (H < content) H = Math.ceil(content / P) * P;
        let before = H - content - e.after, after = e.after;
        if (before < 0) { after += before; before = 0; }
        if (after < 0) { after = 0; }
        setSpacing(p, before, after);
      }
    }
    // 栏内浮动体：整体高度凑整（最后一个元素的段后距补足）
    for (const f of floats) {
      if (f.wide) continue;
      let h = 0;
      for (const n of f.nodes) {
        if (jd.isTbl(n)) { h += n.childrenNamed(jd.q("tr")).length * P; continue; }
        const e = jd.eff(n);
        const imgs = jd.inlineDrawings(n);
        h += e.before + e.after + (imgs.length ? Math.max(...imgs.map((d) => jd.extent(d).h)) : estLines(n, e.sz) * e.lh);
      }
      // 与正文之间至少空半行；整体凑成整数行
      const H = Math.ceil((h + 0.5 * P - 0.5) / P) * P;
      const lastNode = f.nodes[f.nodes.length - 1];
      if (jd.isP(lastNode)) {
        const e = jd.eff(lastNode);
        setSpacing(lastNode, e.before, e.after + (H - h));
      } else {
        // 最后是表格（表题在上）：空白一半加在表题段前，一半用表格后的空段
        const capP = f.nodes.find((x) => jd.isP(x));
        const extra = H - h, top = Math.floor(extra / 2);
        if (capP) { const e = jd.eff(capP); setSpacing(capP, e.before + top, e.after); }
        // 浮动体此时已从正文中取出，空段随浮动体一起插回
        f.nodes.push(jd.spacerPara(Math.max(1, extra - top)));
      }
    }
    // 宽浮动体：上方不留空、下方空一行
    for (const f of floats) {
      if (!f.wide) continue;
      const firstP = f.nodes.find((x) => jd.isP(x));
      const lastP = [...f.nodes].reverse().find((x) => jd.isP(x));
      if (firstP && firstP === f.nodes[0]) jd.setPProp(firstP, "spacing", { before: "0", beforeLines: null });
      if (lastP && lastP === f.nodes[f.nodes.length - 1]) jd.setPProp(lastP, "spacing", { after: String(Math.round(P * TWIP_PER_PT)), afterLines: null });
      else f.nodes.push(jd.spacerPara(P));
    }
    // 正文区高度除以行距的余数至少 3pt，容纳分节用的 1pt 空段而不挤出一行
    const bodyH = geo.pageH - geo.top - geo.bottom;
    const r = bodyH - Math.floor(bodyH / P) * P;
    if (r < 3) {
      const nb = Math.max(0, geo.bottom - (3 - r));
      for (const s of jd.body.descendants()) if (s.local === "sectPr" && s.prefix === jd.w) s.child(jd.q("pgMar"))?.setAttr(jd.q("bottom"), String(Math.round(nb * TWIP_PER_PT)));
    }
  }

  // 4) 浮动体定位
  let lastPlaced: XElement | null = null;
  const insertAfterAnchor = (f: Float) => {
    let before: XElement | undefined;
    if (f.anchor && f.anchor.parent) before = jd.next(f.anchor);
    else before = jd.top()[0];
    if (lastPlaced && before && jd.top().indexOf(before) <= jd.top().indexOf(lastPlaced)) before = jd.next(lastPlaced);
    jd.insertFloat(f, before ?? jd.body.child(jd.q("sectPr"))!);
  };
  const floatNodes = new Set(floats.flatMap((f) => f.nodes));
  const inFloat = (p: XElement) => floatNodes.has(jd.topOf(p));
  for (const f of floats) {
    let placement = "原位置";
    let page: number | undefined;
    if (!render) {
      insertAfterAnchor(f);
      placement = f.wide ? "通栏（原位置）" : "原位置";
    } else {
      try {
        const m = await doRender();
        const pos = f.anchor ? posAfter(m, jd, f.anchor) : null;
        let target: { p: XElement; off: number; tok: PTok } | null = null;
        if (!pos) {
          insertAfterAnchor(f);
        } else if (f.wide && jd.colsOf(jd.sectFor(f.anchor!)) > 1) {
          // 通栏图表放页顶：引用在本页前段（阅读顺序前 35%）→ 本页顶部，否则下一页顶部；第一页不放（标题在上）
          const onPage = m.ptoks.filter((t) => t.page === pos.page && t.col !== "F");
          const before = onPage.filter((t) => t.line < pos.line || (t.line === pos.line && t.top <= pos.top)).length;
          let pg = pos.page > 1 && onPage.length && before / onPage.length < 0.35 ? pos.page : pos.page + 1;
          if (pg <= 1) pg = 2;
          // 引用在最后一页：没有"下一页"，放在本页顶部（叠在已有的通栏图表之下）
          if (pg > m.pages.length) pg = Math.max(2, pos.page);
          target = firstOf(m, pg, "L", inFloat) ?? firstOf(m, pg, "any", inFloat);
          placement = `第 ${pg} 页顶部（通栏）`;
          page = pg;
          if (!target) { insertAfterAnchor(f); placement = "文末（通栏）"; page = undefined; }
        } else {
          // 栏内图表：当前栏剩余空间放得下 → 原位置；否则下一栏顶部
          const P = pitch ?? 12;
          const bodyBot = m.geo.pageH - m.geo.bottom;
          const remain = bodyBot - pos.bottom;
          let need = 0;
          for (const n of f.nodes) {
            if (jd.isTbl(n)) { need += n.childrenNamed(jd.q("tr")).length * P * 1.1; continue; }
            const e = jd.eff(n);
            const imgs = jd.inlineDrawings(n);
            need += e.before + e.after + (imgs.length ? Math.max(...imgs.map((d) => jd.extent(d).h)) : Math.ceil(jd.text(n).length * 0.48 * e.sz / f.colW) * e.lh);
          }
          if (remain >= need + 2) {
            insertAfterAnchor(f);
            placement = `第 ${pos.page} 页原位置`;
            page = pos.page;
          } else if (m.geo.cols > 1 && pos.col === "L") {
            target = firstOf(m, pos.page, "R", inFloat);
            placement = `第 ${pos.page} 页右栏顶部`;
            page = pos.page;
          } else {
            target = firstOf(m, pos.page + 1, "L", inFloat) ?? firstOf(m, pos.page + 1, "any", inFloat);
            placement = `第 ${pos.page + 1} 页${m.geo.cols > 1 ? "左栏" : ""}顶部`;
            page = pos.page + 1;
          }
          if (!target && placement.includes("顶部")) { insertAfterAnchor(f); placement = "原位置（文末）"; }
        }
        if (target) {
          let host = jd.topOf(target.p);
          if (host === target.p && target.off > 0) host = jd.splitPara(target.p, target.off) ?? jd.next(target.p) ?? host;
          if (lastPlaced && jd.top().indexOf(host) <= jd.top().indexOf(lastPlaced)) host = jd.next(lastPlaced) ?? host;
          // 不能放在引用之前太远：目标在引用之前（同页顶部）是允许的；但不能早于上一个浮动体
          jd.insertFloat(f, host);
        }
      } catch (e: any) {
        report.notes.push(`定位"${f.caption}"时渲染失败（${String(e?.message ?? e).slice(0, 100)}），放在原位置。`);
        insertAfterAnchor(f);
      }
    }
    // 保险：无论哪条路径，浮动体都必须回到文档里
    if (!f.nodes[0].parent) { insertAfterAnchor(f); placement += "（回退到原位置）"; }
    lastPlaced = f.nodes[f.nodes.length - 1];
    if (f.wide) { const n = jd.next(lastPlaced); if (n && jd.sectOnly(n)) lastPlaced = n; }
    jd.markAnchor(f);
    report.floats.push({ kind: f.kind, caption: f.caption, wide: f.wide, placement, page });
  }
  jd.cleanSections();

  // 5) 最终检查
  if (render) {
    try {
      const m = await doRender();
      report.pages = m.pages.length;
      const g = m.geo;
      const bodyH = g.pageH - g.top - g.bottom, bodyBot = g.pageH - g.bottom;
      m.colBottoms.forEach((cb, i) => {
        if (i === m.colBottoms.length - 1) return;
        if (bodyBot - cb.any > 0.15 * bodyH) report.whitespacePages.push(i + 1);
        if (cb.hasCols && cb.L && cb.R && Math.abs(cb.L - cb.R) > (pitch ?? 12) * 0.6 && Math.max(cb.L, cb.R) > bodyBot - 2 * (pitch ?? 12)) report.unevenPages.push(i + 1);
      });
      // 每个浮动体最终所在的页
      for (const [k, f] of floats.entries()) {
        const cap = f.nodes.find((n) => jd.isP(n) && (FIG_CAP.test(jd.text(n)) || TAB_CAP.test(jd.text(n))));
        if (!cap) continue;
        const idx = m.dtoks.findIndex((t) => t.p === cap);
        if (idx >= 0 && m.d2p[idx] >= 0) report.floats[k].page = m.ptoks[m.d2p[idx]].page;
      }
    } catch (e: any) {
      report.notes.push(`最终检查渲染失败：${String(e?.message ?? e).slice(0, 120)}`);
    }
  } else {
    report.notes.push("未检测到 LibreOffice：浮动体按原位置排版（已通栏 / 不跨页 / 网格对齐），没有按分页结果移动到页顶或栏顶。安装 LibreOffice 后再运行一次即可。");
  }
  return { data: jd.serialize(), report };
}

export function describeJournalReport(r: JournalReport): string {
  const lines: string[] = [];
  const wide = r.floats.filter((f) => f.wide).length;
  lines.push(`学术期刊排版完成${r.pages ? `：共 ${r.pages} 页` : ""}。`);
  if (r.floats.length) {
    lines.push(`图表 ${r.floats.length} 个（通栏 ${wide} 个、栏内 ${r.floats.length - wide} 个），均设置为不跨页、与题注不分离：`);
    for (const f of r.floats) lines.push(`  - ${f.kind === "figure" ? "图" : "表"}「${f.caption}」→ ${f.placement}${f.page && !f.placement.includes(`第 ${f.page} 页`) ? `（最终在第 ${f.page} 页）` : ""}`);
  } else lines.push("没有找到带题注的图表（题注需以 Fig. 1: / Figure 1. / TABLE I / 图 1 / 表 1 开头）。");
  if (r.pitch) lines.push(`基线网格：行距 ${r.pitch} pt，调整了 ${r.gridAdjusted} 处段距 / 公式 / 图表间距，使左右栏逐行对齐。`);
  if (r.keepTogether) lines.push(`不断开设置：${r.keepTogether} 处（表格行不跨页、图表与题注保持同页）。`);
  if (r.pages !== null) {
    lines.push(r.whitespacePages.length ? `仍有较大留白的页：第 ${r.whitespacePages.join("、")} 页（通常是放不下的大图表导致，可考虑缩小该图表）。` : "除最后一页外，没有大面积留白的页面。");
    lines.push(r.unevenPages.length ? `两栏底部仍未对齐的页：第 ${r.unevenPages.join("、")} 页（多为放不下的公式或标题挪到了下一栏，栏底空出一两行）。` : "各页两栏底部对齐。");
  }
  lines.push(...r.notes);
  return lines.join("\n");
}
/** 供测试脚本检查内部结果 */
export { JDoc as _JDoc, buildMap as _buildMap };
