/**
 * 内置 .doc → .docx 转换：把 DocFile 解析出的内容写成 Word 2007+ 文档。
 *
 * 结构对应：
 *   段落 / 表格（含嵌套表格、合并单元格）、节（页面大小、页边距、分栏、页码、文档网格）、
 *   样式（含基于关系、标题大纲级别）、列表编号（numbering.xml）、页眉页脚（首页 / 奇偶页）、
 *   脚注、尾注、批注（含批注范围）、书签、修订（插入 / 删除）、域（页码、目录、超链接等保持为域）、
 *   行内图片（PNG / JPEG / EMF / WMF / BMP / TIFF）、浮动图片与文本框。
 * 没有转换的内容（形状组合、艺术字、OLE 对象的可编辑数据……）记入报告。
 */

import { createZip } from "../../documents/docx/zip.js";
import { DocFile, type Anchor, type Blip, type ListLevel, type ShapeInfo, type StyleDef, type Story } from "./parse.js";
import { charXml, esc, NUMFMT, paraXml, sectXml, border, shading, cv, type CharCtx } from "./props.js";
import { S, last, type Sprm } from "./sprm.js";

const NS =
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14"';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const R_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const EMU = 635; // 每 twip

export interface DocConvertReport {
  engine: "builtin" | "libreoffice";
  paragraphs: number;
  tables: number;
  images: number;
  floating: number;
  textBoxes: number;
  footnotes: number;
  endnotes: number;
  comments: number;
  revisions: number;
  sections: number;
  /** 原文字符数（正文 + 其他故事，不含控制字符）与写进 Word 的字符数 */
  chars: number;
  written: number;
  notes: string[];
}

interface Para { start: number; end: number; mark: string; istd: number; sprms: Sprm[]; itap: number; ttp: boolean; innerTtp: boolean; innerCell: boolean }
type StoryKind = "main" | "hdr" | "note" | "comment" | "txbx";

export function writeDocx(doc: DocFile): { data: Buffer; report: DocConvertReport } {
  return new Builder(doc).build();
}

class Builder {
  private readonly styleIds = new Map<number, string>();
  private readonly fonts: string[];
  private readonly files: Array<{ name: string; data: Buffer }> = [];
  private readonly rels = new Map<string, Array<{ id: string; type: string; target: string; mode?: string }>>();
  private readonly overrides: Array<{ part: string; type: string }> = [];
  private part = "word/document.xml";
  private drawingId = 1;
  private revId = 1;
  private media = 0;
  private readonly mediaCache = new Map<Blip, string>();
  private readonly listIndex = new Map<number, number>();
  private readonly sectEnds: Set<number>;
  private readonly footnoteAt = new Map<number, number>();
  private readonly endnoteAt = new Map<number, number>();
  private readonly commentAt = new Map<number, number>();
  /** 在某个 CP 之前插入的标记（书签、批注范围） */
  private readonly marks = new Map<number, string[]>();
  private readonly report: DocConvertReport;
  private inTextBox = false;
  private readonly charCtx: CharCtx;
  private headerCount = 0;
  private footerCount = 0;
  private readonly hddStart: number;

  constructor(private readonly d: DocFile) {
    this.fonts = d.fonts.map((f) => f.name);
    this.sectEnds = new Set(d.sections.map((s) => s.cpEnd));
    this.hddStart = d.storyStarts().hdd;
    this.report = { engine: "builtin", paragraphs: 0, tables: 0, images: 0, floating: 0, textBoxes: 0, footnotes: d.footnotes.length, endnotes: d.endnotes.length,
      comments: d.comments.length, revisions: 0, sections: d.sections.length, chars: 0, written: 0, notes: [...d.notes] };
    this.charCtx = { fonts: this.fonts, styleIds: this.styleIds, styleToggle: () => false };
    d.footnotes.forEach((n, i) => this.footnoteAt.set(n.ref, i + 1));
    d.endnotes.forEach((n, i) => this.endnoteAt.set(n.ref, i + 1));
    d.comments.forEach((c, i) => {
      this.commentAt.set(c.ref, i);
      if (c.range && c.range.start <= c.ref && c.range.end <= c.ref + 1) {
        this.mark(c.range.start, `<w:commentRangeStart w:id="${i}"/>`);
        this.mark(Math.min(c.range.end, c.ref), `<w:commentRangeEnd w:id="${i}"/>`);
      }
    });
    d.lists.forEach((l, i) => this.listIndex.set(l.lsid, i));
  }

  private mark(cp: number, xml: string) {
    if (!this.marks.has(cp)) this.marks.set(cp, []);
    this.marks.get(cp)!.push(xml);
  }

  // ---------------------------------------------------------------------------

  build(): { data: Buffer; report: DocConvertReport } {
    const d = this.d;
    this.readBookmarks();
    const styles = this.stylesXml();
    const numbering = this.numberingXml();

    // 正文：按节分段
    const paras = this.paragraphs(0, d.ccp.text);
    const body: string[] = [];
    let si = 0, cur: Para[] = [];
    const flush = (sect: string | undefined) => {
      let xml = this.blocks(cur, 0, "main");
      if (sect) {
        // 节属性放在该节最后一个段落里；最后是表格时补一个空段落
        const at = xml.lastIndexOf("\u0000SECT\u0000");
        if (at >= 0 && xml.endsWith("</w:p>") && !xml.slice(at).includes("<w:p ") && !xml.slice(at).includes("<w:p>")) xml = xml.slice(0, at) + sect + xml.slice(at + 6);
        else xml += `<w:p><w:pPr>${sect}</w:pPr></w:p>`;
      }
      body.push(xml.replace(/\u0000SECT\u0000/g, ""));
      cur = [];
    };
    for (const p of paras) {
      cur.push(p);
      while (si < d.sections.length - 1 && p.end >= d.sections[si].cpEnd) {
        flush(this.sectPr(si));
        si++;
      }
    }
    flush(undefined);
    const finalSect = this.sectPr(Math.min(si, d.sections.length - 1));
    const document = `${XML}<w:document ${NS}><w:body>${body.join("").replace(/<w:pPr><\/w:pPr>/g, "")}${finalSect}</w:body></w:document>`;
    this.files.push({ name: "word/document.xml", data: Buffer.from(document, "utf8") });
    this.files.push({ name: "word/styles.xml", data: Buffer.from(styles, "utf8") });
    if (numbering) {
      this.files.push({ name: "word/numbering.xml", data: Buffer.from(numbering, "utf8") });
      this.rel("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml");
      this.overrides.push({ part: "/word/numbering.xml", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" });
    }
    this.notesParts();
    this.commentsPart();
    this.settingsPart();
    this.rel("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");

    // 字数核对：正文里看得见的字符（不含域代码、控制字符、空白）vs 写进 Word 正文的字符
    this.report.chars = visibleChars(this.storyText(0, d.ccp.text));

    const types =
      `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="emf" ContentType="image/x-emf"/>` +
      `<Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="bmp" ContentType="image/bmp"/><Default Extension="tiff" ContentType="image/tiff"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
      this.overrides.map((o) => `<Override PartName="${o.part}" ContentType="${o.type}"/>`).join("") + `</Types>`;
    const out: Array<{ name: string; data: Buffer }> = [
      { name: "[Content_Types].xml", data: Buffer.from(types, "utf8") },
      { name: "_rels/.rels", data: Buffer.from(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`, "utf8") },
      ...this.files,
    ];
    for (const [part, list] of this.rels) {
      const dir = part.slice(0, part.lastIndexOf("/") + 1), base = part.slice(part.lastIndexOf("/") + 1);
      out.push({ name: `${dir}_rels/${base}.rels`, data: Buffer.from(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.mode ? ` TargetMode="${r.mode}"` : ""}/>`).join("")}</Relationships>`, "utf8") });
    }
    return { data: createZip(out), report: this.report };
  }

  private rel(part: string, type: string, target: string, mode?: string): string {
    if (!this.rels.has(part)) this.rels.set(part, []);
    const list = this.rels.get(part)!;
    const found = list.find((r) => r.type === type && r.target === target);
    if (found) return found.id;
    const id = `rId${list.length + 1}`;
    list.push({ id, type, target, mode });
    return id;
  }

  private storyText(a: number, b: number): string {
    return this.d.text.slice(a, b);
  }

  // ---------------------------------------------------------------------------
  // 段落切分与表格
  // ---------------------------------------------------------------------------

  private paragraphs(start: number, end: number): Para[] {
    const t = this.d.text, out: Para[] = [];
    end = Math.min(end, t.length);
    let s = start;
    const mk = (cpMark: number, endCp: number, mark: string) => {
      const pp = this.d.paraProps(cpMark);
      const itapS = last(pp.sprms, S.PItap);
      const inTable = last(pp.sprms, S.PFInTable);
      const itap = itapS ? itapS.arg.readInt32LE(0) : inTable && inTable.arg[0] ? 1 : 0;
      const flag = (op: number) => { const x = last(pp.sprms, op); return !!x && x.arg[0] !== 0; };
      out.push({ start: s, end: endCp, mark, istd: pp.istd, sprms: pp.sprms, itap: Math.max(0, Math.min(itap, 64)), ttp: flag(S.PFTtp), innerTtp: flag(S.PFInnerTtp), innerCell: flag(S.PFInnerTableCell) });
    };
    for (let cp = start; cp < end; cp++) {
      const ch = t[cp];
      if (ch === "\r" || ch === "\x07" || (ch === "\x0c" && this.sectEnds.has(cp + 1))) { mk(cp, cp + 1, ch); s = cp + 1; }
    }
    if (s < end) mk(end - 1, end, "");
    return out;
  }

  private blocks(paras: Para[], depth: number, kind: StoryKind): string {
    const out: string[] = [];
    for (let i = 0; i < paras.length; ) {
      if (paras[i].itap > depth) {
        let j = i;
        while (j < paras.length && paras[j].itap > depth) j++;
        out.push(this.table(paras.slice(i, j), depth + 1, kind));
        i = j;
      } else out.push(this.para(paras[i++], kind));
    }
    return out.join("");
  }

  private table(paras: Para[], depth: number, kind: StoryKind): string {
    type Row = { cells: Para[][]; ttp?: Para };
    const rows: Row[] = [];
    let cells: Para[][] = [], cell: Para[] = [];
    for (const p of paras) {
      const rowEnd = p.itap === depth && (depth === 1 ? p.ttp : p.innerTtp);
      const cellEnd = p.itap === depth && (depth === 1 ? p.mark === "\x07" : p.innerCell);
      if (rowEnd) { if (cell.length) cells.push(cell); rows.push({ cells, ttp: p }); cells = []; cell = []; continue; }
      cell.push(p);
      if (cellEnd) { cells.push(cell); cell = []; }
    }
    if (cell.length) cells.push(cell);
    if (cells.length) rows.push({ cells });
    if (!rows.length) return "";
    this.report.tables++;

    // 每行的单元格边界（twips）与单元格属性
    interface Cell { x0: number; x1: number; tcgrf: number; brc: Array<string | undefined>; shd?: string; paras: Para[]; vAlign?: number }
    const rowInfo = rows.map((r) => {
      const sp = r.ttp?.sprms ?? [];
      const def = last(sp, S.TDefTable);
      let centers: number[] = [];
      const tcs: Array<{ tcgrf: number; brc: Array<string | undefined> }> = [];
      if (def && def.arg.length >= 1) {
        const n = def.arg[0];
        for (let k = 0; k <= n && 1 + k * 2 + 2 <= def.arg.length; k++) centers.push(def.arg.readInt16LE(1 + k * 2));
        const tcAt = 1 + (n + 1) * 2;
        for (let k = 0; k < n && tcAt + k * 20 + 20 <= def.arg.length; k++) {
          const o = tcAt + k * 20;
          tcs.push({ tcgrf: def.arg.readUInt16LE(o), brc: ["top", "left", "bottom", "right"].map((tag, q) => border(tag, def.arg, o + 4 + q * 4, false)) });
        }
      }
      if (centers.length < 2) {
        // 没有单元格定义：平均分配版心宽度
        const n = Math.max(1, r.cells.length);
        centers = Array.from({ length: n + 1 }, (_, k) => Math.round((k * 9000) / n));
      }
      const gap = last(sp, S.TDxaGapHalf)?.arg.readInt16LE(0) ?? 108;
      // 单元格底纹
      const shds: Array<string | undefined> = [];
      for (const s of sp) {
        if (s.op === S.TDefTableShd80) for (let k = 0; k + 2 <= s.arg.length; k += 2) shds[k / 2] = shading(s.arg, k, false);
        else if (s.op === S.TDefTableShd || s.op === S.TDefTableShd2nd || s.op === S.TDefTableShd3rd) {
          const base = s.op === S.TDefTableShd ? 0 : s.op === S.TDefTableShd2nd ? 22 : 44;
          for (let k = 0; k + 10 <= s.arg.length; k += 10) shds[base + k / 10] = shading(s.arg, k, true);
        }
      }
      // 单独设置的单元格边框、竖直对齐、竖直合并
      const vAligns: Array<number | undefined> = [];
      for (const s of sp) {
        if ((s.op === S.TSetBrc80 || s.op === S.TSetBrc) && s.arg.length >= 3) {
          const wide = s.op === S.TSetBrc;
          const [f, l, which] = [s.arg[0], s.arg[1], s.arg[2]];
          for (let k = f; k < l && k < tcs.length; k++) ["top", "left", "bottom", "right"].forEach((tag, q) => { if (which & (1 << q)) tcs[k].brc[q] = border(tag, s.arg, 3, wide); });
        } else if (s.op === S.TVertAlign && s.arg.length >= 3) {
          for (let k = s.arg[0]; k < s.arg[1]; k++) vAligns[k] = s.arg[2];
        } else if (s.op === S.TVertMerge && s.arg.length >= 2 && tcs[s.arg[0]]) {
          const t = tcs[s.arg[0]];
          t.tcgrf = (t.tcgrf & ~(3 << 5)) | ((s.arg[1] === 3 ? 3 : s.arg[1] ? 2 : 0) << 5);
        }
      }
      const out: Cell[] = [];
      for (let k = 0; k < r.cells.length; k++) {
        const x0 = centers[Math.min(k, centers.length - 2)], x1 = centers[Math.min(k + 1, centers.length - 1)];
        const tc = tcs[k] ?? { tcgrf: 0, brc: [] };
        out.push({ x0, x1: Math.max(x1, x0 + 1), tcgrf: tc.tcgrf, brc: tc.brc, shd: shds[k], paras: r.cells[k], vAlign: vAligns[k] });
      }
      // 旧式横向合并：首格 + 若干被合并的格 → 一个跨列的格
      const merged: Cell[] = [];
      for (const c of out) {
        const h = c.tcgrf & 3;
        if (h >= 2 && merged.length) { merged[merged.length - 1].x1 = c.x1; continue; }
        merged.push(c);
      }
      const height = last(sp, S.TDyaRowHeight)?.arg.readInt16LE(0) ?? 0;
      const header = !!last(sp, S.TTableHeader)?.arg[0];
      const cantSplit = !!last(sp, S.TFCantSplit, S.TFCantSplit90)?.arg[0];
      const jc = last(sp, S.TJc, S.TJc90)?.arg.readUInt16LE(0) ?? 0;
      const tb = last(sp, S.TTableBorders, S.TTableBorders80);
      return { cells: merged, gap, height, header, cantSplit, jc, tb, left: centers[0] };
    });

    // 表格网格：所有行的单元格边界合并去重
    const xs = [...new Set(rowInfo.flatMap((r) => r.cells.flatMap((c) => [c.x0, c.x1])))].sort((a, b) => a - b);
    const grid: number[] = [];
    for (const x of xs) if (!grid.length || x - grid[grid.length - 1] > 4) grid.push(x);
    const col = (x: number) => { let best = 0; for (let k = 0; k < grid.length; k++) if (Math.abs(grid[k] - x) < Math.abs(grid[best] - x)) best = k; return best; };

    const first = rowInfo[0];
    let tblBorders = "";
    if (first.tb) {
      const wide = first.tb.op === S.TTableBorders, step = wide ? 8 : 4;
      tblBorders = `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map((tag, q) => border(tag, first.tb!.arg, q * step, wide) ?? "").join("")}</w:tblBorders>`;
    }
    const jc = ["left", "center", "right"][first.jc] ?? "left";
    const tblPr =
      `<w:tblPr><w:tblW w:w="0" w:type="auto"/>${first.jc ? `<w:jc w:val="${jc}"/>` : ""}` +
      `<w:tblInd w:w="${first.left + first.gap}" w:type="dxa"/>${tblBorders}<w:tblLayout w:type="fixed"/>` +
      `<w:tblCellMar><w:left w:w="${first.gap}" w:type="dxa"/><w:right w:w="${first.gap}" w:type="dxa"/></w:tblCellMar>` +
      `<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>`;
    const tblGrid = `<w:tblGrid>${grid.slice(1).map((x, k) => `<w:gridCol w:w="${Math.max(1, x - grid[k])}"/>`).join("")}</w:tblGrid>`;
    const rowsXml = rowInfo.map((r) => {
      const trPr = (r.cantSplit ? "<w:cantSplit/>" : "") +
        (r.height ? `<w:trHeight w:val="${Math.abs(r.height)}" w:hRule="${r.height < 0 ? "exact" : "atLeast"}"/>` : "") + (r.header ? "<w:tblHeader/>" : "");
      const c0 = col(r.cells[0]?.x0 ?? grid[0]);
      const before = c0 > 0 ? `<w:gridBefore w:val="${c0}"/><w:wBefore w:w="${grid[c0] - grid[0]}" w:type="dxa"/>` : "";
      const cellsXml = r.cells.map((c) => {
        const a = col(c.x0), b = Math.max(a + 1, col(c.x1));
        const vm = (c.tcgrf >> 5) & 3;
        const brc = ["top", "left", "bottom", "right"].map((_, q) => c.brc[q] ?? "").join("");
        const va = c.vAlign ?? (c.tcgrf >> 7) & 3;
        const tcPr = `<w:tcPr><w:tcW w:w="${c.x1 - c.x0}" w:type="dxa"/>${b - a > 1 ? `<w:gridSpan w:val="${b - a}"/>` : ""}` +
          (vm === 3 ? '<w:vMerge w:val="restart"/>' : vm ? "<w:vMerge/>" : "") +
          (brc ? `<w:tcBorders>${brc}</w:tcBorders>` : "") + (c.shd ?? "") + (va ? `<w:vAlign w:val="${["top", "center", "bottom"][va] ?? "top"}"/>` : "") + `</w:tcPr>`;
        let inner = this.blocks(c.paras, depth, kind);
        if (!inner.includes("<w:p>") && !inner.includes("<w:p ") || inner.endsWith("</w:tbl>")) inner += "<w:p/>";
        return `<w:tc>${tcPr}${inner}</w:tc>`;
      }).join("");
      return `<w:tr>${trPr || before ? `<w:trPr>${before}${trPr}</w:trPr>` : ""}${cellsXml || "<w:tc><w:p/></w:tc>"}</w:tr>`;
    }).join("");
    return `<w:tbl>${tblPr}${tblGrid}${rowsXml}</w:tbl>`;
  }

  // ---------------------------------------------------------------------------
  // 段落与文字
  // ---------------------------------------------------------------------------

  private fieldDepth: Array<{ instr: boolean }> = [];

  private para(p: Para, kind: StoryKind): string {
    this.report.paragraphs++;
    const ppr = paraXml(p.sprms, { styleIds: this.styleIds, numId: (ilfo) => this.numId(ilfo) }, p.istd);
    this.charCtx.styleToggle = (op) => this.styleToggle(p.istd, op);
    const runs = this.runs(p.start, p.mark ? p.end - 1 : p.end, kind);
    // 段落标记的字符格式（空段落的行高由它决定）
    const markRuns = p.mark ? this.d.charRuns(p.end - 1, p.end) : [];
    const mrpr = markRuns.length ? charXml(markRuns[0].sprms, this.charCtx) : "";
    const tail = p.mark ? this.marksAt(p.end - 1) : "";
    return `<w:p><w:pPr>${ppr}${mrpr ? `<w:rPr>${mrpr}</w:rPr>` : ""}\u0000SECT\u0000</w:pPr>${runs}${tail}</w:p>`;
  }

  private marksAt(cp: number): string {
    const m = this.marks.get(cp);
    if (!m) return "";
    this.marks.delete(cp);
    return m.join("");
  }

  /** 段落样式 / 字符样式里开关属性（粗体、斜体……）的值：0x81 表示"与样式相反" */
  private styleToggle(istd: number, op: number): boolean {
    const seen = new Set<number>();
    for (let i: number | undefined = istd; i !== undefined && i < 0x0fff && !seen.has(i); ) {
      seen.add(i);
      const st: StyleDef | undefined = this.d.styles[i];
      if (!st) break;
      const s = last(st.chpx, op);
      if (s) return s.arg[0] === 1 || s.arg[0] === 0x81;
      i = st.base;
    }
    return false;
  }

  private runs(a: number, b: number, kind: StoryKind): string {
    const out: string[] = [];
    const t = this.d.text;
    for (const r of this.d.charRuns(a, b)) {
      const sp = r.sprms;
      const flag = (op: number) => { const x = last(sp, op); return !!x && x.arg[0] !== 0 && x.arg[0] !== 0x80; };
      const spec = flag(S.CFSpec);
      const del = flag(S.CFRMarkDel), ins = flag(S.CFRMarkIns);
      const rpr = charXml(sp, this.charCtx);
      const rp = rpr ? `<w:rPr>${rpr}</w:rPr>` : "";
      let buf = "";
      const pieces: string[] = [];
      const flush = () => {
        if (!buf) return;
        const tag = this.fieldDepth.length && this.fieldDepth[this.fieldDepth.length - 1].instr ? "instrText" : del ? "delText" : "t";
        pieces.push(`<w:r>${rp}<w:${tag} xml:space="preserve">${esc(buf)}</w:${tag}></w:r>`);
        if (kind === "main" && tag !== "instrText") this.report.written += buf.replace(/[\s　]/g, "").length;
        buf = "";
      };
      const el = (x: string) => { flush(); pieces.push(`<w:r>${rp}${x}</w:r>`); };
      for (let cp = r.cp; cp < r.end; cp++) {
        const m = this.marks.get(cp);
        if (m) { flush(); pieces.push(this.marksAt(cp)); }
        const ch = t[cp];
        if (ch === undefined) break;
        const code = ch.charCodeAt(0);
        const inInstr = this.fieldDepth.length > 0 && this.fieldDepth[this.fieldDepth.length - 1].instr;
        if (code >= 0x20 && !(code === 0x28 && last(sp, S.CSymbol))) {
          if (code >= 0xd800 && code <= 0xdbff) { buf += ch + (t[cp + 1] ?? ""); cp++; continue; }
          if (code === 0xfffe || code === 0xffff) continue;
          buf += ch;
          continue;
        }
        switch (code) {
          case 0x09: el("<w:tab/>"); break;
          case 0x0b: el("<w:br/>"); break;
          case 0x0c: el('<w:br w:type="page"/>'); break;
          case 0x0e: el('<w:br w:type="column"/>'); break;
          case 0x1e: el("<w:noBreakHyphen/>"); break;
          case 0x1f: el("<w:softHyphen/>"); break;
          case 0x13: el('<w:fldChar w:fldCharType="begin"/>'); this.fieldDepth.push({ instr: true }); break;
          case 0x14: if (this.fieldDepth.length) { el('<w:fldChar w:fldCharType="separate"/>'); this.fieldDepth[this.fieldDepth.length - 1].instr = false; } break;
          case 0x15: if (this.fieldDepth.length) { el('<w:fldChar w:fldCharType="end"/>'); this.fieldDepth.pop(); } break;
          case 0x28: {
            const sym = last(sp, S.CSymbol)!;
            const font = this.fonts[sym.arg.readUInt16LE(0)] ?? "Symbol";
            el(`<w:sym w:font="${esc(font)}" w:char="${sym.arg.readUInt16LE(2).toString(16).toUpperCase().padStart(4, "0")}"/>`);
            break;
          }
          case 0x01: {
            if (!spec || inInstr || flag(S.CFData)) break;
            const loc = last(sp, S.CPicLocation);
            const pic = loc ? this.d.picture(loc.arg.readUInt32LE(0)) : undefined;
            if (pic?.blip) {
              const c = pic.crop;
              el(this.drawing(pic.blip, pic.w, pic.h, undefined, c.some(Boolean) ? [c[0] / (pic.w + c[0] + c[2]), c[1] / (pic.h + c[1] + c[3]), c[2] / (pic.w + c[0] + c[2]), c[3] / (pic.h + c[1] + c[3])] : undefined));
            } else if (pic?.empty) this.note("有图片框在原文件中就是空的（没有图片数据），已略过");
            else if (loc) this.note("有图片的格式无法识别，未转换");
            break;
          }
          case 0x08: if (spec && !inInstr) { const x = this.floating(cp, kind); if (x) el(x); } break;
          case 0x02: {
            if (!spec) break;
            if (kind === "main") {
              const fn = this.footnoteAt.get(cp), en = this.endnoteAt.get(cp);
              if (fn) el(`<w:footnoteReference w:id="${fn}"/>`);
              else if (en) el(`<w:endnoteReference w:id="${en}"/>`);
            } else if (kind === "note") el(this.noteRef);
            break;
          }
          case 0x05: {
            if (!spec) break;
            if (kind === "main") { const c = this.commentAt.get(cp); if (c !== undefined) el(`<w:commentReference w:id="${c}"/>`); }
            else if (kind === "comment") el("<w:annotationRef/>");
            break;
          }
          default: break; // 其他控制字符（脚注分隔符等）
        }
      }
      flush();
      if (!pieces.length) continue;
      if (del || ins) {
        this.report.revisions++;
        const who = last(sp, del ? S.CIbstRMarkDel : S.CIbstRMark) ?? last(sp, S.CIbstRMark);
        const when = last(sp, del ? S.CDttmRMarkDel : S.CDttmRMark) ?? last(sp, S.CDttmRMark);
        const author = esc(this.d.rmarkAuthors[who ? who.arg.readUInt16LE(0) : 0] ?? "作者");
        const date = when ? dttm(when.arg.readUInt32LE(0)) : undefined;
        out.push(`<w:${del ? "del" : "ins"} w:id="${this.revId++}" w:author="${author}"${date ? ` w:date="${date}"` : ""}>${pieces.join("")}</w:${del ? "del" : "ins"}>`);
      } else out.push(pieces.join(""));
    }
    return out.join("");
  }

  private noteRef = "<w:footnoteRef/>";

  private note(s: string) {
    if (!this.report.notes.includes(s)) this.report.notes.push(s);
  }

  // ---------------------------------------------------------------------------
  // 图片
  // ---------------------------------------------------------------------------

  private mediaTarget(blip: Blip): string {
    let name = this.mediaCache.get(blip);
    if (!name) {
      name = `media/image${++this.media}.${blip.ext}`;
      this.mediaCache.set(blip, name);
      this.files.push({ name: `word/${name}`, data: blip.data });
    }
    return name;
  }

  /** 图片：行内（anchor 省略）或浮动；crop 为左上右下的裁剪比例 */
  private drawing(blip: Blip, wTw: number, hTw: number, anchor?: { a: Anchor; shape?: ShapeInfo }, crop?: number[]): string {
    const rid = this.rel(this.part, R_IMAGE, this.mediaTarget(blip));
    const id = this.drawingId++;
    const cx = Math.max(1, Math.round(wTw * EMU)), cy = Math.max(1, Math.round(hTw * EMU));
    const src = crop && crop.some((x) => x > 0) ? `<a:srcRect l="${Math.round(crop[0] * 100000)}" t="${Math.round(crop[1] * 100000)}" r="${Math.round(crop[2] * 100000)}" b="${Math.round(crop[3] * 100000)}"/>` : "";
    const graphic =
      `<wp:docPr id="${id}" name="图片 ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}.${blip.ext}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/>${src}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
    if (!anchor) {
      this.report.images++;
      return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${graphic}</wp:inline></w:drawing>`;
    }
    this.report.images++;
    this.report.floating++;
    return `<w:drawing>${this.anchorOpen(anchor.a, id, cx, cy)}${graphic}</wp:anchor></w:drawing>`;
  }

  private anchorOpen(a: Anchor, id: number, cx: number, cy: number): string {
    const rh = ["margin", "page", "column"][a.bx] ?? "column";
    const rv = ["margin", "page", "paragraph"][a.by] ?? "paragraph";
    const behind = a.below && a.wr === 3;
    const wrap = a.wr === 3 ? "<wp:wrapNone/>" : a.wr === 1 ? "<wp:wrapTopAndBottom/>" : '<wp:wrapSquare wrapText="bothSides"/>';
    return `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251658240 + id}" behindDoc="${behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${rh}"><wp:posOffset>${Math.round(a.x0 * EMU)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="${rv}"><wp:posOffset>${Math.round(a.y0 * EMU)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}`;
  }

  /** 0x08：浮动图形（图片、文本框） */
  private floating(cp: number, kind: StoryKind): string | undefined {
    if (this.inTextBox || kind === "txbx") return undefined;
    const hdr = kind === "hdr";
    const list = hdr ? this.d.anchorsHdr : this.d.anchorsMain;
    const rel = hdr ? cp - this.hddStart : cp;
    const a = list.find((x) => x.cp === rel);
    if (!a) return undefined;
    const shape = this.d.shapes.get(a.spid);
    if (!shape) { this.note("部分浮动图形缺少绘图数据，未转换"); return undefined; }
    const pib = shape.props.get(0x104);
    const txid = shape.props.get(0x80);
    const w = Math.max(1, a.x1 - a.x0), h = Math.max(1, a.y1 - a.y0);
    if (pib && !shape.group) {
      const blip = this.d.blipAt(pib);
      if (!blip) { this.note("部分浮动图片的数据缺失，未转换"); return undefined; }
      const f = (k: number) => { const v = shape.props.get(k); return v ? (v | 0) / 65536 : 0; };
      return this.drawing(blip, w, h, { a, shape }, [f(0x102), f(0x100), f(0x103), f(0x101)]);
    }
    if (txid !== undefined && !shape.group) {
      const map = hdr ? this.d.txbx.hdr : this.d.txbx.main;
      const story = map.get(a.spid) ?? (hdr ? this.d.txbx.hdrList : this.d.txbx.mainList)[(txid >>> 16) - 1];
      if (story) return this.textBox(a, shape, story);
    }
    this.note(shape.group ? "形状组合（多个图形组合在一起）未转换" : "自选图形（线条、箭头、色块等）未转换");
    return undefined;
  }

  private textBox(a: Anchor, shape: ShapeInfo, story: Story): string {
    this.report.textBoxes++;
    const id = this.drawingId++;
    const cx = Math.max(1, Math.round((a.x1 - a.x0) * EMU)), cy = Math.max(1, Math.round((a.y1 - a.y0) * EMU));
    const prev = this.inTextBox;
    this.inTextBox = true;
    const saved = this.fieldDepth;
    this.fieldDepth = [];
    // 故事末尾多出的一个段落标记不属于文本框
    let end = story.end;
    if (end > story.start && this.d.text[end - 1] === "\r" && end - 1 > story.start && this.d.text[end - 2] === "\r") end--;
    let inner = this.blocks(this.paragraphs(story.start, end), 0, "txbx").replace(/\u0000SECT\u0000/g, "");
    if (!inner) inner = "<w:p/>";
    this.fieldDepth = saved;
    this.inTextBox = prev;
    const p = shape.props;
    const bools = p.get(0x1bf), lineBools = p.get(0x1ff);
    const filled = bools === undefined || !(bools & 0x100000) || (bools & 0x10) !== 0;
    const lined = lineBools === undefined || !(lineBools & 0x80000) || (lineBools & 0x8) !== 0;
    const col = (v: number | undefined, def: string) => {
      if (v === undefined || v >>> 24 > 0x10) return def;
      const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0); b[3] = 0;
      return cv(b);
    };
    const fill = filled ? `<a:solidFill><a:srgbClr val="${col(p.get(0x181), "FFFFFF")}"/></a:solidFill>` : "<a:noFill/>";
    const line = lined ? `<a:ln w="${p.get(0x1cb) ?? 9525}"><a:solidFill><a:srgbClr val="${col(p.get(0x1c0), "000000")}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
    const ins = (k: number, d: number) => p.get(k) ?? d;
    return `<w:drawing>${this.anchorOpen(a, id, cx, cy)}<wp:docPr id="${id}" name="文本框 ${id}"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
      `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}${line}</wps:spPr>` +
      `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx>` +
      `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${ins(0x81, 91440)}" tIns="${ins(0x82, 45720)}" rIns="${ins(0x83, 91440)}" bIns="${ins(0x84, 45720)}" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>` +
      `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  }

  // ---------------------------------------------------------------------------
  // 节、页眉页脚
  // ---------------------------------------------------------------------------

  private sectPr(k: number): string {
    const sec = this.d.sections[k];
    const refs: string[] = [];
    const base = 6 + k * 6;
    const kinds: Array<[number, "header" | "footer", "even" | "default" | "first"]> = [
      [0, "header", "even"], [1, "header", "default"], [2, "footer", "even"], [3, "footer", "default"], [4, "header", "first"], [5, "footer", "first"],
    ];
    for (const [off, kind, type] of kinds) {
      const i = base + off;
      if (i + 1 >= this.d.hdd.length) continue;
      const s = this.hddStart + this.d.hdd[i], e = this.hddStart + this.d.hdd[i + 1];
      if (e <= s) continue;
      const id = this.hdrFtrPart(kind, s, e);
      refs.push(`<w:${kind}Reference w:type="${type}" r:id="${id}"/>`);
    }
    return sectXml(sec.sprms, refs.join(""));
  }

  private hdrFtrPart(kind: "header" | "footer", s: number, e: number): string {
    // 每个页眉页脚故事末尾多一个段落标记（故事结束符），不是内容
    if (e - s >= 2 && this.d.text[e - 1] === "\r") e--;
    const n = kind === "header" ? ++this.headerCount : ++this.footerCount;
    const name = `${kind}${n}.xml`;
    const prev = this.part;
    this.part = `word/${name}`;
    const saved = this.fieldDepth;
    this.fieldDepth = [];
    let inner = this.blocks(this.paragraphs(s, e), 0, "hdr").replace(/\u0000SECT\u0000/g, "").replace(/<w:pPr><\/w:pPr>/g, "");
    if (!inner) inner = "<w:p/>";
    this.fieldDepth = saved;
    this.part = prev;
    const tag = kind === "header" ? "hdr" : "ftr";
    this.files.push({ name: `word/${name}`, data: Buffer.from(`${XML}<w:${tag} ${NS}>${inner}</w:${tag}>`, "utf8") });
    this.overrides.push({ part: `/word/${name}`, type: `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml` });
    return this.rel("word/document.xml", `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`, name);
  }

  // ---------------------------------------------------------------------------
  // 脚注、尾注、批注、设置
  // ---------------------------------------------------------------------------

  private notesParts(): void {
    for (const [list, tag, kind] of [[this.d.footnotes, "footnote", "footnotes"], [this.d.endnotes, "endnote", "endnotes"]] as const) {
      if (!list.length) continue;
      const prev = this.part;
      this.part = `word/${kind}.xml`;
      this.noteRef = `<w:${tag}Ref/>`;
      const items = list.map((n, i) => {
        let end = n.story.end;
        const saved = this.fieldDepth;
        this.fieldDepth = [];
        let inner = this.blocks(this.paragraphs(n.story.start, end), 0, "note").replace(/\u0000SECT\u0000/g, "").replace(/<w:pPr><\/w:pPr>/g, "");
        this.fieldDepth = saved;
        if (!inner) inner = "<w:p/>";
        void end;
        return `<w:${tag} w:id="${i + 1}">${inner}</w:${tag}>`;
      }).join("");
      this.part = prev;
      const seps = `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}>` +
        `<w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`;
      this.files.push({ name: `word/${kind}.xml`, data: Buffer.from(`${XML}<w:${kind} ${NS}>${seps}${items}</w:${kind}>`, "utf8") });
      this.overrides.push({ part: `/word/${kind}.xml`, type: `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml` });
      this.rel("word/document.xml", `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`, `${kind}.xml`);
    }
    this.noteRef = "<w:footnoteRef/>";
  }

  private commentsPart(): void {
    if (!this.d.comments.length) return;
    const prev = this.part;
    this.part = "word/comments.xml";
    const items = this.d.comments.map((c, i) => {
      const saved = this.fieldDepth;
      this.fieldDepth = [];
      let inner = this.blocks(this.paragraphs(c.story.start, c.story.end), 0, "comment").replace(/\u0000SECT\u0000/g, "").replace(/<w:pPr><\/w:pPr>/g, "");
      this.fieldDepth = saved;
      if (!inner) inner = "<w:p/>";
      return `<w:comment w:id="${i}" w:author="${esc(c.author || c.initials || "作者")}" w:initials="${esc(c.initials)}">${inner}</w:comment>`;
    }).join("");
    this.part = prev;
    this.files.push({ name: "word/comments.xml", data: Buffer.from(`${XML}<w:comments ${NS}>${items}</w:comments>`, "utf8") });
    this.overrides.push({ part: "/word/comments.xml", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml" });
    this.rel("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml");
  }

  private settingsPart(): void {
    const d = this.d;
    const notes = (d.footnotes.length ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : "") +
      (d.endnotes.length ? '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>' : "");
    const xml = `${XML}<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:defaultTabStop w:val="${d.dop.dxaTab}"/>${d.dop.facingPages ? "<w:evenAndOddHeaders/>" : ""}` +
      `<w:characterSpacingControl w:val="doNotCompress"/>${notes}` +
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat></w:settings>`;
    this.files.push({ name: "word/settings.xml", data: Buffer.from(xml, "utf8") });
    this.rel("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml");
  }

  /** 书签（目录、交叉引用的跳转目标） */
  private readBookmarks(): void {
    try {
      const d = this.d as unknown as { table: Buffer; word: Buffer };
      const w = d.word;
      // 重新读取 FIB 里书签相关的三项（SttbfBkmk 21、PlcfBkf 22、PlcfBkl 23）
      let p = 32;
      p += 2 + w.readUInt16LE(p) * 2;
      const cslw = w.readUInt16LE(p); p += 2 + cslw * 4;
      const cb = w.readUInt16LE(p); p += 2;
      const fl = (i: number) => (i < cb ? [w.readUInt32LE(p + i * 8), w.readUInt32LE(p + i * 8 + 4)] : [0, 0]);
      const [fcN, lcbN] = fl(21), [fcF, lcbF] = fl(22), [fcL, lcbL] = fl(23);
      if (!lcbN || !lcbF || !lcbL) return;
      const t = d.table;
      const names: string[] = [];
      const ext = t.readUInt16LE(fcN) === 0xffff;
      const cData = t.readUInt16LE(fcN + (ext ? 2 : 0));
      const cbExtra = t.readUInt16LE(fcN + (ext ? 4 : 2));
      let o = fcN + (ext ? 6 : 4);
      for (let i = 0; i < cData && o < fcN + lcbN; i++) {
        const cch = t.readUInt16LE(o);
        names.push(t.toString("utf16le", o + 2, o + 2 + cch * 2));
        o += 2 + cch * 2 + cbExtra;
      }
      const nF = Math.floor((lcbF - 4) / 8), nL = Math.floor(lcbL / 4) - 1;
      for (let i = 0; i < nF && i < names.length; i++) {
        const s = t.readUInt32LE(fcF + i * 4), ibkl = t.readUInt16LE(fcF + 4 * (nF + 1) + i * 4);
        if (ibkl >= nL) continue;
        const e = t.readUInt32LE(fcL + ibkl * 4);
        if (s > this.d.ccp.text || e > this.d.ccp.text) continue;
        const name = names[i].replace(/[^\p{L}\p{N}_]/gu, "_").slice(0, 40) || `bm${i}`;
        this.mark(s, `<w:bookmarkStart w:id="${10000 + i}" w:name="${esc(name)}"/>`);
        this.mark(e, `<w:bookmarkEnd w:id="${10000 + i}"/>`);
      }
    } catch { /* 书签是可选的 */ }
  }

  // ---------------------------------------------------------------------------
  // 样式与编号
  // ---------------------------------------------------------------------------

  private stylesXml(): string {
    const d = this.d;
    const used = new Set<string>();
    // 样式名必须唯一（"Normal,Normalc" 去掉别名后会与真正的 Normal 重名，后者被覆盖）
    const names = new Map<number, string>();
    const usedNames = new Set<string>();
    d.styles.forEach((st, i) => {
      if (!st) return;
      let name = st.name.split(",")[0].trim() || `Style${i}`;
      if (usedNames.has(name.toLowerCase())) name = st.name.trim();
      for (let k = 2; usedNames.has(name.toLowerCase()); k++) name = `${st.name.split(",")[0].trim()} (${k})`;
      usedNames.add(name.toLowerCase());
      names.set(i, name);
    });
    d.styles.forEach((st, i) => {
      if (!st) return;
      const name = names.get(i)!;
      let id = /^[A-Za-z0-9 _\-.]+$/.test(name) ? name.replace(/(^|\s+)(\w)/g, (_, _s, c: string) => c.toUpperCase()).replace(/[^A-Za-z0-9]/g, "") : `s${i}`;
      if (!id) id = `s${i}`;
      while (used.has(id.toLowerCase())) id += "x";
      used.add(id.toLowerCase());
      this.styleIds.set(i, id);
    });
    const df = d.defaultFonts;
    const font = (i: number) => esc(this.fonts[i] ?? "");
    const pctx = { styleIds: this.styleIds, numId: (ilfo: number) => this.numId(ilfo) };
    const cctx: CharCtx = { fonts: this.fonts, styleIds: this.styleIds, styleToggle: () => false };
    const items = d.styles.map((st, i) => {
      if (!st) return "";
      const type = ["", "paragraph", "character", "table", "numbering"][st.stk];
      if (!type) return "";
      const name = names.get(i)!;
      const isDefault = st.sti === 0 || st.sti === 65 || st.sti === 105 || st.sti === 107;
      const base = this.validBase(i) ? `<w:basedOn w:val="${this.styleIds.get(st.base)}"/>` : "";
      const nx = st.next < 0x0fff ? d.styles[st.next] : undefined;
      const next = type === "paragraph" && nx && nx.stk === 1 && st.next !== i && this.styleIds.get(st.next) ? `<w:next w:val="${this.styleIds.get(st.next)}"/>` : "";
      let ppr = st.papx.length ? paraXml(st.papx, pctx) : "";
      if (type === "paragraph" && st.sti >= 1 && st.sti <= 9 && !ppr.includes("outlineLvl")) ppr = ppr + `<w:outlineLvl w:val="${st.sti - 1}"/>`;
      const rpr = st.chpx.length ? charXml(st.chpx, cctx, { skipStyle: true }) : "";
      return `<w:style w:type="${type}"${isDefault ? ' w:default="1"' : ""} w:styleId="${this.styleIds.get(i)}"><w:name w:val="${esc(name)}"/>${base}${next}` +
        (st.sti >= 1 && st.sti <= 9 || st.sti === 0 ? "<w:qFormat/>" : "") +
        (ppr && type !== "character" ? `<w:pPr>${ppr}</w:pPr>` : "") + (rpr && type !== "numbering" ? `<w:rPr>${rpr}</w:rPr>` : "") +
        (type === "table" ? '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>' : "") +
        `</w:style>`;
    }).join("");
    return `${XML}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font(df.ascii)}" w:eastAsia="${font(df.fe)}" w:hAnsi="${font(df.other)}" w:cs="${font(df.other)}"/>` +
      `<w:kern w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>${items}</w:styles>`;
  }

  /** "基于"关系是否可用：基础样式存在、类型相同、沿链不会绕回自己（损坏的文件里出现过 A 基于 B、B 基于 A，LibreOffice 读到会死循环） */
  private validBase(i: number): boolean {
    const st = this.d.styles[i];
    if (!st || st.sti === 0 || st.base >= 0x0fff || st.base === i) return false;
    const b = this.d.styles[st.base];
    if (!b || b.stk !== st.stk || !this.styleIds.has(st.base)) return false;
    const seen = new Set([i]);
    for (let k: number | undefined = st.base, n = 0; k !== undefined && k < 0x0fff && n < 64; n++) {
      if (seen.has(k)) return false;
      seen.add(k);
      const x: StyleDef | undefined = this.d.styles[k];
      if (!x || x.base >= 0x0fff) return true;
      k = x.base;
    }
    return true;
  }

  private numId(ilfo: number): number | undefined {
    if (ilfo < 1 || ilfo > this.d.overrides.length) return undefined;
    return this.listIndex.has(this.d.overrides[ilfo - 1].lsid) ? ilfo : undefined;
  }

  private lvlXml(k: number, l: ListLevel, styleOf?: number): string {
    const pctx = { styleIds: this.styleIds, numId: () => undefined };
    const cctx: CharCtx = { fonts: this.fonts, styleIds: this.styleIds, styleToggle: () => false };
    let text = "";
    for (const ch of l.text) text += ch.charCodeAt(0) < 9 ? `%${ch.charCodeAt(0) + 1}` : ch;
    const fmt = NUMFMT[l.nfc] ?? "decimal";
    const ppr = paraXml(l.papx, pctx).replace(/<w:numPr>.*?<\/w:numPr>/, "");
    const rpr = charXml(l.chpx, cctx, { skipStyle: true });
    return `<w:lvl w:ilvl="${k}"><w:start w:val="${l.start}"/><w:numFmt w:val="${fmt}"/>` +
      (l.noRestart ? '<w:lvlRestart w:val="0"/>' : "") +
      (styleOf !== undefined && this.styleIds.get(styleOf) ? `<w:pStyle w:val="${this.styleIds.get(styleOf)}"/>` : "") +
      (l.legal ? "<w:isLgl/>" : "") + `<w:suff w:val="${["tab", "space", "nothing"][l.follow] ?? "tab"}"/>` +
      `<w:lvlText w:val="${esc(text)}"/><w:lvlJc w:val="${["left", "center", "right"][l.jc] ?? "left"}"/>` +
      (ppr ? `<w:pPr>${ppr}</w:pPr>` : "") + (rpr ? `<w:rPr>${rpr}</w:rPr>` : "") + `</w:lvl>`;
  }

  private numberingXml(): string | undefined {
    const d = this.d;
    if (!d.lists.length || !d.overrides.length) return undefined;
    const abs = d.lists.map((l, i) =>
      `<w:abstractNum w:abstractNumId="${i}"><w:multiLevelType w:val="${l.simple ? "singleLevel" : "multilevel"}"/>` +
      l.levels.map((lv, k) => this.lvlXml(k, lv)).join("") + `</w:abstractNum>`).join("");
    const nums = d.overrides.map((o, i) => {
      const a = this.listIndex.get(o.lsid);
      if (a === undefined) return "";
      const ov = [...o.levels.entries()].map(([k, x]) =>
        `<w:lvlOverride w:ilvl="${k}">${x.startAt !== undefined ? `<w:startOverride w:val="${x.startAt}"/>` : ""}${x.lvl ? this.lvlXml(k, x.lvl) : ""}</w:lvlOverride>`).join("");
      return `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${a}"/>${ov}</w:num>`;
    }).join("");
    return `${XML}<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abs}${nums}</w:numbering>`;
  }
}

/** DTTM → ISO 时间 */
function dttm(v: number): string | undefined {
  if (!v) return undefined;
  const min = v & 0x3f, hr = (v >> 6) & 0x1f, dom = (v >> 11) & 0x1f, mon = (v >> 16) & 0xf, yr = ((v >> 20) & 0x1ff) + 1900;
  if (!mon || !dom) return undefined;
  const p = (n: number) => String(n).padStart(2, "0");
  return `${yr}-${p(mon)}-${p(dom)}T${p(hr)}:${p(min)}:00Z`;
}

/** 可见字符数：跳过域代码（0x13 与 0x14 / 0x15 之间）、控制字符与空白 */
export function visibleChars(t: string): number {
  let n = 0;
  const stack: boolean[] = [];
  for (const ch of t) {
    const c = ch.charCodeAt(0);
    if (c === 0x13) { stack.push(true); continue; }
    if (c === 0x14) { if (stack.length) stack[stack.length - 1] = false; continue; }
    if (c === 0x15) { stack.pop(); continue; }
    if (stack.length && stack[stack.length - 1]) continue;
    if (c < 0x20 || c === 0x3000 || /\s/.test(ch)) continue;
    n++;
  }
  return n;
}
