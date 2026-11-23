/**
 * Word 97-2003 二进制文档（.doc，[MS-DOC]）解析：把文字、格式、样式、列表、节、各个文字"故事"与图片读出来。
 *
 * .doc 的结构：
 *   WordDocument 流  FIB（文件信息块，所有其他结构的索引）+ 文字 + 字符 / 段落格式页（FKP）
 *   0Table / 1Table  片段表（CP → 文字在流中的位置）、样式表、字体表、列表、节、页眉页脚 / 脚注 / 批注的位置表……
 *   Data 流          图片（PICF + OfficeArt）、超大的段落格式
 * 文字按 CP（字符位置）编号：正文、脚注、页眉页脚、批注、尾注、文本框依次排在同一条 CP 轴上，称为"故事"。
 */

import zlib from "node:zlib";
import { CompoundFile, DocFileError } from "./cfb.js";
import { parseGrpprl, S, type Sprm } from "./sprm.js";

export { DocFileError };

export interface Piece { cpStart: number; cpEnd: number; fc: number; compressed: boolean; prm: Sprm[] }
interface FkpRun { fcStart: number; fcEnd: number; sprms: Sprm[]; istd: number }
export interface StyleDef { istd: number; sti: number; stk: number; name: string; base: number; next: number; papx: Sprm[]; chpx: Sprm[]; tapx: Sprm[] }
export interface ListLevel { start: number; nfc: number; jc: number; legal: boolean; noRestart: boolean; follow: number; text: string; nums: number[]; papx: Sprm[]; chpx: Sprm[]; restartLim: number }
export interface ListDef { lsid: number; simple: boolean; levels: ListLevel[] }
export interface ListOverride { lsid: number; levels: Map<number, { startAt?: number; lvl?: ListLevel }> }
export interface Section { cpEnd: number; sprms: Sprm[] }
export interface Blip { ext: "png" | "jpeg" | "emf" | "wmf" | "bmp" | "tiff"; data: Buffer }
/** 浮动图形的锚点（SPA） */
export interface Anchor { cp: number; spid: number; x0: number; y0: number; x1: number; y1: number; bx: number; by: number; wr: number; wrk: number; below: boolean }
export interface ShapeInfo {
  spid: number; type: number; props: Map<number, number>;
  /** 是组合本身，或在组合里 */
  group: boolean;
  /** FSP 标志：0x1 组合、0x2 组合内的形状、0x40 水平翻转、0x80 垂直翻转 */
  flags: number;
  /** 组合内形状在组合坐标系里的位置（左、上、右、下） */
  child?: [number, number, number, number];
  /** 组合的坐标系（左、上、右、下） */
  coords?: [number, number, number, number];
  /** 组合内的形状（按绘制顺序） */
  children?: number[];
  /** 复杂属性（如自由曲线的顶点、路径） */
  complex: Map<number, Buffer>;
}
export interface Story { start: number; end: number }
export interface Note { ref: number; story: Story; auto: boolean }
export interface Comment { ref: number; story: Story; author: string; initials: string; range?: Story }
export interface PicInfo { blip?: Blip; w: number; h: number; crop: [number, number, number, number]; /** 空的图片框：形状里没有引用任何图片（原文件中就是空的） */ empty?: boolean }

const td1252 = new TextDecoder("windows-1252");

export class DocFile {
  readonly word: Buffer;
  readonly table: Buffer;
  readonly data: Buffer;
  readonly nFib: number;
  readonly ccp: { text: number; ftn: number; hdd: number; mcr: number; atn: number; edn: number; txbx: number; hdrTxbx: number };
  readonly pieces: Piece[] = [];
  readonly text: string;
  private readonly chpx: FkpRun[];
  private readonly papx: FkpRun[];
  readonly styles: Array<StyleDef | undefined> = [];
  readonly defaultFonts: { ascii: number; fe: number; other: number };
  readonly fonts: Array<{ name: string; charset: number; family: number; pitch: number }> = [];
  readonly lists: ListDef[] = [];
  readonly overrides: ListOverride[] = [];
  readonly sections: Section[] = [];
  /** 页眉页脚故事的 CP 边界（相对页眉故事起点） */
  readonly hdd: number[] = [];
  readonly footnotes: Note[];
  readonly endnotes: Note[];
  readonly comments: Comment[];
  readonly rmarkAuthors: string[];
  readonly dop: { dxaTab: number; facingPages: boolean; compat: string[] };
  /** 文本框故事（正文 / 页眉）：FTXBXS 里的形状 id → 故事 */
  readonly txbx: { main: Map<number, Story>; hdr: Map<number, Story>; mainList: Story[]; hdrList: Story[] };
  readonly anchorsMain: Anchor[];
  readonly anchorsHdr: Anchor[];
  readonly shapes = new Map<number, ShapeInfo>();
  private readonly bstore: Array<Blip | undefined> = [];
  /** 解析中遇到但没有转换的内容（写进转换报告） */
  readonly notes: string[] = [];

  constructor(buf: Buffer) {
    const cf = new CompoundFile(buf);
    if (!cf.has("WordDocument")) throw new DocFileError("不是 Word 文档（复合文档里没有 WordDocument 流）。", "not-ole");
    const w = this.word = cf.stream("WordDocument");
    if (w.length < 0x200) throw new DocFileError("Word 文档已损坏（文件信息块不完整）。");
    const ident = w.readUInt16LE(0);
    const flags = w.readUInt16LE(0x0a);
    let nFib = w.readUInt16LE(2);
    if (flags & 0x0100) throw new DocFileError("文档设置了打开密码（已加密），无法读取。请在 Word 中取消密码后再上传。", "encrypted");
    if (ident !== 0xa5ec || nFib < 0xc1) throw new DocFileError("这是 Word 6.0 / 95 或更早版本的文档，内置引擎不支持（需要安装 LibreOffice）。", "old-version");
    const tableName = flags & 0x0200 ? "1Table" : "0Table";
    if (!cf.has(tableName)) throw new DocFileError(`Word 文档已损坏（缺少 ${tableName} 流）。`);
    this.table = cf.stream(tableName);
    this.data = cf.has("Data") ? cf.stream("Data") : Buffer.alloc(0);

    // FIB：FibBase(32) → fibRgW → fibRgLw → fibRgFcLcb → fibRgCswNew
    let p = 32;
    const csw = w.readUInt16LE(p); p += 2 + csw * 2;
    const cslw = w.readUInt16LE(p); p += 2;
    const rgLw = p; p += cslw * 4;
    const cbRgFcLcb = w.readUInt16LE(p); p += 2;
    const rgFcLcb = p; p += cbRgFcLcb * 8;
    if (p + 2 <= w.length) {
      const cswNew = w.readUInt16LE(p);
      if (cswNew > 0 && p + 4 <= w.length) nFib = Math.max(nFib, w.readUInt16LE(p + 2));
    }
    this.nFib = nFib;
    const lw = (i: number) => (i < cslw ? w.readInt32LE(rgLw + i * 4) : 0);
    this.ccp = { text: lw(3), ftn: lw(4), hdd: lw(5), mcr: lw(6), atn: lw(7), edn: lw(8), txbx: lw(9), hdrTxbx: lw(10) };
    const fcLcb = (i: number): [number, number] => (i < cbRgFcLcb ? [w.readUInt32LE(rgFcLcb + i * 8), w.readUInt32LE(rgFcLcb + i * 8 + 4)] : [0, 0]);
    const tbl = (i: number): Buffer => {
      const [fc, lcb] = fcLcb(i);
      if (!lcb || fc + lcb > this.table.length) return Buffer.alloc(0);
      return this.table.subarray(fc, fc + lcb);
    };

    // 片段表
    const clx = tbl(33);
    const prcs: Buffer[] = [];
    let plcPcd: Buffer | undefined;
    for (let i = 0; i < clx.length; ) {
      if (clx[i] === 1) { const cb = clx.readInt16LE(i + 1); prcs.push(clx.subarray(i + 3, i + 3 + cb)); i += 3 + cb; }
      else if (clx[i] === 2) { const lcb = clx.readUInt32LE(i + 1); plcPcd = clx.subarray(i + 5, i + 5 + lcb); break; }
      else break;
    }
    if (!plcPcd || plcPcd.length < 16) throw new DocFileError("Word 文档已损坏（缺少片段表）。");
    const nPcd = Math.floor((plcPcd.length - 4) / 12);
    const parts: string[] = [];
    for (let k = 0; k < nPcd; k++) {
      const cpStart = plcPcd.readUInt32LE(k * 4), cpEnd = plcPcd.readUInt32LE(k * 4 + 4);
      const d = 4 * (nPcd + 1) + k * 8;
      const fcRaw = plcPcd.readUInt32LE(d + 2);
      const compressed = (fcRaw & 0x40000000) !== 0;
      const fc = compressed ? (fcRaw & 0x3fffffff) / 2 : fcRaw & 0x3fffffff;
      const prmRaw = plcPcd.readUInt16LE(d + 6);
      // Prm：复杂形式（最低位 1）指向 Clx 里的第 igrpprl 组 sprm；简单形式（单条 sprm）极少出现，忽略
      const prm = prmRaw & 1 && prcs[prmRaw >> 1] ? parseGrpprl(prcs[prmRaw >> 1]) : [];
      if (cpEnd < cpStart) continue;
      const n = cpEnd - cpStart;
      const bytes = w.subarray(fc, fc + n * (compressed ? 1 : 2));
      parts.push(compressed ? td1252.decode(bytes) : bytes.toString("utf16le"));
      const got = parts[parts.length - 1].length;
      if (got < n) parts[parts.length - 1] += "\r".repeat(n - got); // 越界的片段补段落标记，CP 不错位
      this.pieces.push({ cpStart, cpEnd, fc, compressed, prm });
    }
    this.text = parts.join("");

    // 字符 / 段落格式页
    this.chpx = this.readFkps(tbl(12), "chp");
    this.papx = this.readFkps(tbl(13), "pap");

    // 样式表
    const stsh = tbl(1);
    let df = { ascii: 0, fe: 0, other: 0 };
    if (stsh.length > 4) {
      const cbStshi = stsh.readUInt16LE(0);
      const cstd = stsh.readUInt16LE(2), cbBase = stsh.readUInt16LE(4);
      if (cbStshi >= 18) df = { ascii: stsh.readUInt16LE(2 + 12), fe: stsh.readUInt16LE(2 + 14), other: stsh.readUInt16LE(2 + 16) };
      let o = 2 + cbStshi;
      for (let i = 0; i < cstd && o + 2 <= stsh.length; i++) {
        const cb = stsh.readUInt16LE(o); o += 2;
        if (cb && o + cb <= stsh.length) {
          try { this.styles[i] = parseStd(stsh.subarray(o, o + cb), cbBase, i); } catch { /* 损坏的样式跳过 */ }
        }
        o += cb;
      }
    }
    this.defaultFonts = df;

    // 字体表
    const ffn = tbl(15);
    if (ffn.length >= 4) {
      let cData = ffn.readUInt16LE(0), o = 4;
      if (cData === 0xffff) { cData = ffn.readUInt16LE(2); o = 6; }
      for (let i = 0; i < cData && o < ffn.length; i++) {
        const cb = ffn[o];
        const f = ffn.subarray(o + 1, o + 1 + cb);
        o += 1 + cb;
        if (f.length < 40) { this.fonts.push({ name: "", charset: 0, family: 0, pitch: 0 }); continue; }
        let end = 39;
        while (end + 1 < f.length && f.readUInt16LE(end) !== 0) end += 2;
        this.fonts.push({ name: f.toString("utf16le", 39, end).replace(/[\u0000-\u001f]/g, "").trim(), charset: f[3], family: (f[0] >> 4) & 7, pitch: f[0] & 3 });
      }
    }

    // 列表
    this.readLists(fcLcb(73), fcLcb(74));

    // 节
    const sed = tbl(6);
    if (sed.length >= 16) {
      const n = Math.floor((sed.length - 4) / 16);
      for (let k = 0; k < n; k++) {
        const cpEnd = sed.readUInt32LE((k + 1) * 4);
        const fcSepx = sed.readUInt32LE(4 * (n + 1) + k * 12 + 2);
        let sprms: Sprm[] = [];
        if (fcSepx !== 0xffffffff && fcSepx + 2 <= w.length) {
          const cb = w.readInt16LE(fcSepx);
          if (cb > 0) sprms = parseGrpprl(w, fcSepx + 2, fcSepx + 2 + cb);
        }
        this.sections.push({ cpEnd, sprms });
      }
    }
    if (!this.sections.length) this.sections.push({ cpEnd: this.ccp.text, sprms: [] });

    // 页眉页脚
    const hdd = tbl(11);
    for (let o = 0; o + 4 <= hdd.length; o += 4) this.hdd.push(hdd.readUInt32LE(o));

    // 脚注、尾注、批注
    const starts = this.storyStarts();
    this.footnotes = this.readNotes(tbl(2), tbl(3), starts.ftn);
    this.endnotes = this.readNotes(tbl(46), tbl(47), starts.edn);
    this.comments = this.readComments(tbl(4), tbl(5), tbl(36), starts.atn, tbl(37), tbl(42), tbl(43));
    this.rmarkAuthors = readSttb(tbl(51));

    // 文档属性
    const dop = tbl(31);
    this.dop = { dxaTab: dop.length >= 12 ? dop.readUInt16LE(10) || 720 : 720, facingPages: dop.length > 0 && (dop[0] & 1) === 1, compat: readCompat(dop) };

    // 文本框故事、浮动图形
    this.txbx = { main: new Map(), hdr: new Map(), mainList: [], hdrList: [] };
    this.readTxbx(tbl(56), starts.txbx, this.txbx.main, this.txbx.mainList);
    this.readTxbx(tbl(58), starts.hdrTxbx, this.txbx.hdr, this.txbx.hdrList);
    this.anchorsMain = readSpa(tbl(40));
    this.anchorsHdr = readSpa(tbl(41));
    try { this.readOfficeArt(tbl(50)); } catch { this.notes.push("浮动图形的绘图数据已损坏，浮动图片未转换"); }
  }

  /** 各故事在 CP 轴上的起点 */
  storyStarts() {
    const c = this.ccp;
    const ftn = c.text, hdd = ftn + c.ftn, atn = hdd + c.hdd + c.mcr, edn = atn + c.atn, txbx = edn + c.edn, hdrTxbx = txbx + c.txbx;
    return { main: 0, ftn, hdd, atn, edn, txbx, hdrTxbx, end: hdrTxbx + c.hdrTxbx };
  }

  // ---------------------------------------------------------------------------
  // 格式页（FKP）
  // ---------------------------------------------------------------------------

  private readFkps(plc: Buffer, kind: "chp" | "pap"): FkpRun[] {
    const out: FkpRun[] = [];
    if (plc.length < 12) return out;
    const n = Math.floor((plc.length - 4) / 8);
    const seen = new Set<number>();
    for (let k = 0; k < n; k++) {
      const pn = plc.readUInt32LE(4 * (n + 1) + k * 4) & 0x3fffff;
      if (seen.has(pn)) continue;
      seen.add(pn);
      const off = pn * 512;
      if (off + 512 > this.word.length) continue;
      const page = this.word.subarray(off, off + 512);
      const cnt = page[511];
      // 条目数与页内结构不符（损坏的文件）：跳过这一页
      if ((cnt + 1) * 4 + cnt * (kind === "chp" ? 1 : 13) > 511) continue;
      for (let i = 0; i < cnt; i++) {
        const fcStart = page.readUInt32LE(i * 4), fcEnd = page.readUInt32LE((i + 1) * 4);
        if (kind === "chp") {
          const b = page[(cnt + 1) * 4 + i] * 2;
          const cb = b ? page[b] : 0;
          out.push({ fcStart, fcEnd, istd: 0, sprms: b ? parseGrpprl(page, b + 1, b + 1 + cb) : [] });
        } else {
          const b = page[(cnt + 1) * 4 + i * 13] * 2;
          if (!b) { out.push({ fcStart, fcEnd, istd: 0, sprms: [] }); continue; }
          let cb = page[b], st = b + 1, len: number;
          if (cb) len = 2 * cb - 1;
          else { cb = page[b + 1]; st = b + 2; len = 2 * cb; }
          if (len < 2 || st + len > 512) { out.push({ fcStart, fcEnd, istd: 0, sprms: [] }); continue; }
          const istd = page.readUInt16LE(st);
          let sprms = parseGrpprl(page, st + 2, st + len);
          const huge = sprms.find((s) => s.op === S.PHugePapx);
          if (huge && this.data.length) {
            const fc = huge.arg.readUInt32LE(0);
            if (fc + 2 <= this.data.length) {
              const hcb = this.data.readUInt16LE(fc);
              sprms = [...sprms.filter((s) => s.op !== S.PHugePapx), ...parseGrpprl(this.data, fc + 2, fc + 2 + hcb)];
            }
          }
          out.push({ fcStart, fcEnd, istd, sprms });
        }
      }
    }
    out.sort((a, b) => a.fcStart - b.fcStart);
    return out;
  }

  private findRun(runs: FkpRun[], fc: number): FkpRun | undefined {
    let lo = 0, hi = runs.length - 1;
    while (lo <= hi) {
      const m = (lo + hi) >> 1;
      if (runs[m].fcEnd <= fc) lo = m + 1;
      else if (runs[m].fcStart > fc) hi = m - 1;
      else return runs[m];
    }
    return undefined;
  }

  private nextRunStart(runs: FkpRun[], fc: number): number | undefined {
    let lo = 0, hi = runs.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (runs[m].fcStart <= fc) lo = m + 1; else hi = m; }
    return runs[lo]?.fcStart;
  }

  pieceAt(cp: number): Piece | undefined {
    let lo = 0, hi = this.pieces.length - 1;
    while (lo <= hi) {
      const m = (lo + hi) >> 1, p = this.pieces[m];
      if (p.cpEnd <= cp) lo = m + 1;
      else if (p.cpStart > cp) hi = m - 1;
      else return p;
    }
    return undefined;
  }

  /** [a, b) 内的字符格式片段 */
  charRuns(a: number, b: number): Array<{ cp: number; end: number; sprms: Sprm[] }> {
    const out: Array<{ cp: number; end: number; sprms: Sprm[] }> = [];
    let cp = a;
    while (cp < b) {
      const p = this.pieceAt(cp);
      if (!p) { out.push({ cp, end: b, sprms: [] }); break; }
      const bpc = p.compressed ? 1 : 2;
      const fc = p.fc + (cp - p.cpStart) * bpc;
      const r = this.findRun(this.chpx, fc);
      let end = Math.min(b, p.cpEnd);
      const lim = r ? r.fcEnd : this.nextRunStart(this.chpx, fc);
      if (lim !== undefined) end = Math.min(end, cp + Math.max(1, Math.ceil((lim - fc) / bpc)));
      const prm = p.prm.filter((s) => (s.op >> 10 & 7) === 2);
      out.push({ cp, end, sprms: prm.length ? [...(r?.sprms ?? []), ...prm] : r?.sprms ?? [] });
      cp = end;
    }
    return out;
  }

  /** 以 cp 处的段落标记结尾的段落的格式：样式 + sprm */
  paraProps(cp: number): { istd: number; sprms: Sprm[] } {
    const p = this.pieceAt(cp);
    if (!p) return { istd: 0, sprms: [] };
    const fc = p.fc + (cp - p.cpStart) * (p.compressed ? 1 : 2);
    const r = this.findRun(this.papx, fc);
    const prm = p.prm.filter((s) => (s.op >> 10 & 7) === 1 || (s.op >> 10 & 7) === 5);
    return { istd: r?.istd ?? 0, sprms: prm.length ? [...(r?.sprms ?? []), ...prm] : r?.sprms ?? [] };
  }

  // ---------------------------------------------------------------------------
  // 列表
  // ---------------------------------------------------------------------------

  private readLists([fcLst, lcbLst]: [number, number], [fcLfo, lcbLfo]: [number, number]): void {
    const t = this.table;
    if (lcbLst >= 2 && fcLst + lcbLst <= t.length) {
      const n = t.readUInt16LE(fcLst);
      let o = fcLst + 2;
      const lstfs: Array<{ lsid: number; simple: boolean }> = [];
      for (let i = 0; i < n && o + 28 <= t.length; i++, o += 28) lstfs.push({ lsid: t.readInt32LE(o), simple: (t[o + 26] & 1) === 1 });
      // LVL 紧跟在 PlfLst 之后
      for (const l of lstfs) {
        const levels: ListLevel[] = [];
        for (let k = 0; k < (l.simple ? 1 : 9) && o + 28 <= t.length; k++) {
          const [lvl, next] = readLvl(t, o);
          levels.push(lvl);
          o = next;
        }
        this.lists.push({ lsid: l.lsid, simple: l.simple, levels });
      }
    }
    if (lcbLfo >= 4 && fcLfo + lcbLfo <= t.length) {
      const n = t.readUInt32LE(fcLfo);
      const lfos: Array<{ lsid: number; clfolvl: number }> = [];
      let o = fcLfo + 4;
      for (let i = 0; i < n && o + 16 <= t.length; i++, o += 16) lfos.push({ lsid: t.readInt32LE(o), clfolvl: t[o + 12] });
      for (const lf of lfos) {
        const levels = new Map<number, { startAt?: number; lvl?: ListLevel }>();
        if (o + 4 <= t.length) o += 4; // LFOData.cp
        for (let k = 0; k < lf.clfolvl && o + 8 <= t.length; k++) {
          const startAt = t.readInt32LE(o);
          const fl = t.readUInt32LE(o + 4);
          o += 8;
          const ilvl = fl & 0xf, fStartAt = (fl >> 4) & 1, fFormatting = (fl >> 5) & 1;
          const ov: { startAt?: number; lvl?: ListLevel } = {};
          if (fStartAt) ov.startAt = startAt;
          if (fFormatting && o + 28 <= t.length) { const [lvl, next] = readLvl(t, o); ov.lvl = lvl; o = next; }
          levels.set(ilvl, ov);
        }
        this.overrides.push({ lsid: lf.lsid, levels });
      }
    }
  }

  // ---------------------------------------------------------------------------
  // 脚注、尾注、批注、文本框
  // ---------------------------------------------------------------------------

  private readNotes(ref: Buffer, txt: Buffer, start: number): Note[] {
    const out: Note[] = [];
    if (ref.length < 6 || txt.length < 8) return out;
    const n = Math.floor((ref.length - 4) / 6);
    for (let k = 0; k < n; k++) {
      const cp = ref.readUInt32LE(k * 4);
      const auto = ref.readInt16LE(4 * (n + 1) + k * 2) !== 0;
      if ((k + 1) * 4 + 4 > txt.length) break;
      out.push({ ref: cp, auto, story: { start: start + txt.readUInt32LE(k * 4), end: start + txt.readUInt32LE((k + 1) * 4) } });
    }
    return out;
  }

  private readComments(ref: Buffer, txt: Buffer, owners: Buffer, start: number, atnBkmk: Buffer, bkf: Buffer, bkl: Buffer): Comment[] {
    const out: Comment[] = [];
    if (ref.length < 34 || txt.length < 8) return out;
    const names: string[] = [];
    for (let o = 0; o + 2 <= owners.length; ) {
      const cch = owners.readUInt16LE(o);
      names.push(owners.toString("utf16le", o + 2, o + 2 + cch * 2));
      o += 2 + cch * 2;
    }
    // 批注范围：SttbfAtnBkmk 的附加数据里是 lTag；PlcfAtnBkf / Bkl 给出起止 CP
    const tags = new Map<number, Story>();
    try {
      if (atnBkmk.length >= 6 && bkf.length >= 8) {
        const cData = atnBkmk.readUInt16LE(2), cbExtra = atnBkmk.readUInt16LE(4);
        const nb = Math.floor((bkf.length - 4) / 8);
        const nl = Math.floor(bkl.length / 4) - 1;
        let o = 6;
        for (let i = 0; i < cData && o + 2 <= atnBkmk.length; i++) {
          const cch = atnBkmk.readUInt16LE(o); o += 2 + cch * 2;
          const lTag = cbExtra >= 6 ? atnBkmk.readInt32LE(o + 2) : -1;
          o += cbExtra;
          if (i < nb) {
            const s = bkf.readUInt32LE(i * 4), ibkl = bkf.readUInt16LE(4 * (nb + 1) + i * 4);
            if (ibkl < nl) tags.set(lTag, { start: s, end: bkl.readUInt32LE(ibkl * 4) });
          }
        }
      }
    } catch { /* 范围是可选的 */ }
    const n = Math.floor((ref.length - 4) / 34);
    for (let k = 0; k < n; k++) {
      const cp = ref.readUInt32LE(k * 4);
      const d = 4 * (n + 1) + k * 30;
      const cchI = Math.min(9, ref.readUInt16LE(d));
      const initials = ref.toString("utf16le", d + 2, d + 2 + cchI * 2);
      const ibst = ref.readUInt16LE(d + 20);
      const tag = ref.readInt32LE(d + 26);
      if ((k + 1) * 4 + 4 > txt.length) break;
      out.push({ ref: cp, author: names[ibst] ?? initials ?? "", initials, range: tag >= 0 ? tags.get(tag) : undefined,
        story: { start: start + txt.readUInt32LE(k * 4), end: start + txt.readUInt32LE((k + 1) * 4) } });
    }
    return out;
  }

  private readTxbx(plc: Buffer, start: number, byId: Map<number, Story>, list: Story[]): void {
    if (plc.length < 30) return;
    const n = Math.floor((plc.length - 4) / 26);
    for (let k = 0; k < n; k++) {
      const s = { start: start + plc.readUInt32LE(k * 4), end: start + plc.readUInt32LE((k + 1) * 4) };
      const lid = plc.readInt32LE(4 * (n + 1) + k * 22 + 14);
      list.push(s);
      byId.set(lid, s);
    }
  }

  // ---------------------------------------------------------------------------
  // 图片
  // ---------------------------------------------------------------------------

  private readOfficeArt(buf: Buffer): void {
    if (buf.length < 8) return;
    const walk = (s: number, e: number, fn: (type: number, inst: number, o: number, len: number) => void) => {
      for (let o = s; o + 8 <= e; ) {
        const vi = buf.readUInt16LE(o), type = buf.readUInt16LE(o + 2), len = buf.readUInt32LE(o + 4);
        if (o + 8 + len > e) break;
        fn(type, vi >> 4, o + 8, len);
        o += 8 + len;
      }
    };
    let o = 0;
    // DggContainer：BStore
    if (buf.readUInt16LE(2) === 0xf000) {
      const len = buf.readUInt32LE(4);
      walk(8, 8 + len, (type, _i, s, l) => {
        if (type !== 0xf001) return;
        walk(s, s + l, (t2, _i2, s2, l2) => {
          if (t2 === 0xf007) {
            const cbName = buf[s2 + 33];
            const size = buf.readUInt32LE(s2 + 20), foDelay = buf.readUInt32LE(s2 + 28);
            let blip: Blip | undefined;
            if (l2 > 36 + cbName) blip = readBlip(buf, s2 + 36 + cbName, s2 + l2);
            else if (foDelay !== 0xffffffff && foDelay + 8 <= this.word.length) blip = readBlip(this.word, foDelay, Math.min(this.word.length, foDelay + size + 8));
            this.bstore.push(blip);
          } else this.bstore.push(readBlip(buf, s2 - 8, s2 + l2));
        });
      });
      o = 8 + len;
    }
    // 绘图：dgglbl(1) + DgContainer
    while (o + 9 <= buf.length) {
      o += 1;
      if (buf.readUInt16LE(o + 2) !== 0xf002) break;
      const len = buf.readUInt32LE(o + 4);
      // 形状组：最外层（depth 0）是整页的"总组"，其中的形状都是顶层形状；
      // 嵌套的组（depth ≥ 1）第一个形状是组本身（锚点指向它），其余是组内形状（含更深的组）
      const spgr = (s: number, e: number, depth: number): number | undefined => {
        let self: number | undefined;
        walk(s, e, (t3, _i3, s3, l3) => {
          let id: number | undefined;
          if (t3 === 0xf004) id = this.readShape(buf, s3, l3, walk, depth > 1 || (depth === 1 && self !== undefined));
          else if (t3 === 0xf003) id = spgr(s3, s3 + l3, depth + 1);
          if (id === undefined || depth === 0) return;
          if (self === undefined) self = id;
          else this.shapes.get(self)?.children?.push(id);
        });
        return self;
      };
      walk(o + 8, o + 8 + len, (type, _inst, s2, l2) => {
        if (type === 0xf003) spgr(s2, s2 + l2, 0);
        else if (type === 0xf004) this.readShape(buf, s2, l2, walk, false);
      });
      o += 8 + len;
    }
  }

  private readShape(buf: Buffer, s: number, l: number, walk: (s: number, e: number, fn: (type: number, inst: number, o: number, len: number) => void) => void, inGroup: boolean): number | undefined {
    let spid = -1, type = 0, flags = 0;
    let child: ShapeInfo["child"], coords: ShapeInfo["coords"];
    const props = new Map<number, number>();
    const complex = new Map<number, Buffer>();
    const rect = (o: number): [number, number, number, number] => [buf.readInt32LE(o), buf.readInt32LE(o + 4), buf.readInt32LE(o + 8), buf.readInt32LE(o + 12)];
    walk(s, s + l, (t, inst, o, len) => {
      if (t === 0xf00a && len >= 8) { spid = buf.readUInt32LE(o); type = inst; flags = buf.readUInt32LE(o + 4); }
      else if (t === 0xf009 && len >= 16) coords = rect(o);
      else if (t === 0xf00f && len >= 16) child = rect(o);
      else if (t === 0xf00b || t === 0xf121 || t === 0xf122) {
        // 属性表：每项 6 字节；带 fComplex 标志的属性，数据按顺序接在表后面
        let data = o + inst * 6;
        for (let k = 0; k < inst && k * 6 + 6 <= len; k++) {
          const pid = buf.readUInt16LE(o + k * 6), val = buf.readUInt32LE(o + k * 6 + 2);
          props.set(pid & 0x3fff, val);
          if (pid & 0x8000) {
            // 几何数组（IMsoArray）：有的程序写的长度不含 6 字节的数组头
            let size = val;
            const id = pid & 0x3fff;
            if (id >= 0x145 && id <= 0x158 && data + 6 <= o + len) {
              const n = buf.readUInt16LE(data), cb = buf.readUInt16LE(data + 4);
              if (n * (cb === 0xfff0 ? 4 : cb) === val && data + val + 6 <= o + len) size = val + 6;
            }
            if (data + size <= o + len) complex.set(id, buf.subarray(data, data + size));
            data += size;
          }
        }
      }
    });
    if (spid < 0) return undefined;
    if (!this.shapes.has(spid)) {
      this.shapes.set(spid, { spid, type, props, group: (flags & 1) === 1 || inGroup, flags, child, coords, children: flags & 1 ? [] : undefined, complex });
    }
    return spid;
  }

  blipAt(index1: number): Blip | undefined {
    return this.bstore[index1 - 1];
  }

  /** 行内图片：0x01 字符的 sprmCPicLocation 指向 Data 流里的 PICF */
  picture(fc: number): PicInfo | undefined {
    const d = this.data;
    if (fc + 68 > d.length) return undefined;
    const lcb = d.readUInt32LE(fc), cbHeader = d.readUInt16LE(fc + 4);
    if (lcb < 68 || cbHeader < 68 || fc + lcb > d.length) return undefined;
    const mm = d.readInt16LE(fc + 6);
    const dxaGoal = d.readInt16LE(fc + 28), dyaGoal = d.readInt16LE(fc + 30);
    const mx = d.readUInt16LE(fc + 32), my = d.readUInt16LE(fc + 34);
    const crop: [number, number, number, number] = [d.readInt16LE(fc + 36), d.readInt16LE(fc + 38), d.readInt16LE(fc + 40), d.readInt16LE(fc + 42)];
    const w = Math.max(1, ((dxaGoal - crop[0] - crop[2]) * (mx || 1000)) / 1000);
    const h = Math.max(1, ((dyaGoal - crop[1] - crop[3]) * (my || 1000)) / 1000);
    let o = fc + cbHeader;
    if (mm === 0x66) o += 1 + d[o];
    const end = fc + lcb;
    let blip: Blip | undefined;
    let empty = false;
    // OfficeArtInlineSpContainer：形状 + 一个或几个 FBSE（内嵌图片）
    if (o + 8 <= end && d.readUInt16LE(o + 2) === 0xf004) {
      let pib = 0;
      const spLen = d.readUInt32LE(o + 4);
      for (let q = o + 8; q + 8 <= o + 8 + spLen; ) {
        const vi = d.readUInt16LE(q), t = d.readUInt16LE(q + 2), len = d.readUInt32LE(q + 4);
        if (t === 0xf00b) for (let k = 0; k < vi >> 4; k++) { if ((d.readUInt16LE(q + 8 + k * 6) & 0x3fff) === 0x104) pib = d.readUInt32LE(q + 8 + k * 6 + 2); }
        q += 8 + len;
      }
      const fbses: Array<Blip | undefined> = [];
      for (let q = o + 8 + spLen; q + 8 <= end; ) {
        const t = d.readUInt16LE(q + 2), len = d.readUInt32LE(q + 4);
        if (t === 0xf007) { const cbName = d[q + 8 + 33]; fbses.push(readBlip(d, q + 8 + 36 + cbName, Math.min(end, q + 8 + len))); }
        else if (t >= 0xf018 && t <= 0xf117) fbses.push(readBlip(d, q, Math.min(end, q + 8 + len)));
        q += 8 + len;
      }
      blip = fbses[pib ? pib - 1 : 0] ?? fbses.find(Boolean) ?? (pib ? this.blipAt(pib) : undefined);
      empty = !pib && !fbses.length;
    } else if (o < end) {
      // 早期的图片：PICF 后面直接是 Windows 图元文件
      const raw = d.subarray(o, end);
      if (raw.length > 18 && (raw.readUInt16LE(0) === 1 || raw.readUInt16LE(0) === 2) && raw.readUInt16LE(2) === 9) blip = { ext: "wmf", data: raw };
    }
    return { blip, w, h, crop, empty };
  }
}

// ---------------------------------------------------------------------------

function parseStd(b: Buffer, cbBase: number, istd: number): StyleDef {
  const sti = b.readUInt16LE(0) & 0x0fff;
  const w1 = b.readUInt16LE(2), w2 = b.readUInt16LE(4);
  const stk = w1 & 0xf, base = w1 >> 4, cupx = w2 & 0xf, next = w2 >> 4;
  let o = cbBase;
  const cch = b.readUInt16LE(o);
  const name = b.toString("utf16le", o + 2, o + 2 + cch * 2);
  o += 2 + cch * 2 + 2;
  const upx: Buffer[] = [];
  for (let k = 0; k < cupx && o + 2 <= b.length; k++) {
    const cb = b.readUInt16LE(o);
    upx.push(b.subarray(o + 2, o + 2 + cb));
    o += 2 + cb;
    if (cb & 1) o++;
  }
  const g = (x?: Buffer, skip = 0) => (x && x.length > skip ? parseGrpprl(x, skip) : []);
  let papx: Sprm[] = [], chpx: Sprm[] = [], tapx: Sprm[] = [];
  if (stk === 1) { papx = g(upx[0], 2); chpx = g(upx[1]); }
  else if (stk === 2) chpx = g(upx[0]);
  else if (stk === 3) { tapx = g(upx[0]); papx = g(upx[1], 2); chpx = g(upx[2]); }
  else if (stk === 4) papx = g(upx[0], 2);
  return { istd, sti, stk, name, base, next, papx, chpx, tapx };
}

function readLvl(t: Buffer, o: number): [ListLevel, number] {
  const start = t.readInt32LE(o), nfc = t[o + 4], fl = t[o + 5];
  const nums = [...t.subarray(o + 6, o + 15)].filter((x) => x > 0);
  const follow = t[o + 15];
  const cbChpx = t[o + 24], cbPapx = t[o + 25], restartLim = t[o + 26];
  let p = o + 28;
  const papx = parseGrpprl(t, p, p + cbPapx); p += cbPapx;
  const chpx = parseGrpprl(t, p, p + cbChpx); p += cbChpx;
  const cch = p + 2 <= t.length ? t.readUInt16LE(p) : 0;
  const text = t.toString("utf16le", p + 2, Math.min(t.length, p + 2 + cch * 2));
  p += 2 + cch * 2;
  return [{ start, nfc, jc: fl & 3, legal: (fl & 4) !== 0, noRestart: (fl & 8) !== 0, follow, text, nums, papx, chpx, restartLim }, p];
}

/** STTB（字符串表）：fExtend 0xFFFF 表示 UTF-16 */
function readSttb(b: Buffer): string[] {
  const out: string[] = [];
  if (b.length < 6) return out;
  const ext = b.readUInt16LE(0) === 0xffff;
  const cData = ext ? b.readUInt16LE(2) : b.readUInt16LE(0);
  const cbExtra = ext ? b.readUInt16LE(4) : b.readUInt16LE(2);
  let o = ext ? 6 : 4;
  for (let i = 0; i < cData && o < b.length; i++) {
    if (ext) { const cch = b.readUInt16LE(o); out.push(b.toString("utf16le", o + 2, o + 2 + cch * 2)); o += 2 + cch * 2; }
    else { const cch = b[o]; out.push(b.toString("latin1", o + 1, o + 1 + cch)); o += 1 + cch; }
    o += cbExtra;
  }
  return out;
}

function readSpa(plc: Buffer): Anchor[] {
  const out: Anchor[] = [];
  if (plc.length < 34) return out;
  const n = Math.floor((plc.length - 4) / 30);
  for (let k = 0; k < n; k++) {
    const d = 4 * (n + 1) + k * 26;
    const fl = plc.readUInt16LE(d + 20);
    out.push({
      cp: plc.readUInt32LE(k * 4), spid: plc.readInt32LE(d),
      x0: plc.readInt32LE(d + 4), y0: plc.readInt32LE(d + 8), x1: plc.readInt32LE(d + 12), y1: plc.readInt32LE(d + 16),
      bx: (fl >> 1) & 3, by: (fl >> 3) & 3, wr: (fl >> 5) & 0xf, wrk: (fl >> 9) & 0xf, below: (fl & 0x4000) !== 0,
    });
  }
  return out;
}

/** 读取一条 OfficeArt 图片记录（从记录头开始） */
function readBlip(b: Buffer, o: number, end: number): Blip | undefined {
  if (o + 8 > end || o + 8 > b.length) return undefined;
  const inst = b.readUInt16LE(o) >> 4, type = b.readUInt16LE(o + 2), len = b.readUInt32LE(o + 4);
  const s = o + 8, e = Math.min(end, s + len, b.length);
  const uids = inst & 1 ? 32 : 16;
  switch (type) {
    case 0xf01a: case 0xf01b: {
      const h = s + uids;
      if (h + 34 > e) return undefined;
      const cbSize = b.readUInt32LE(h), cbSave = b.readUInt32LE(h + 28), compression = b[h + 32];
      let data = b.subarray(h + 34, Math.min(e, h + 34 + cbSave));
      if (compression === 0) {
        try { data = zlib.inflateSync(data); } catch { return undefined; }
      }
      if (type === 0xf01b) {
        // 图元文件缺少"可放置"文件头时补上（Word / LibreOffice 都能识别）
        const l = b.readInt32LE(h + 4), t = b.readInt32LE(h + 8), r = b.readInt32LE(h + 12), bt = b.readInt32LE(h + 16);
        return { ext: "wmf", data: placeableWmf(data.subarray(0, cbSize || data.length), l, t, r, bt, b.readInt32LE(h + 20)) };
      }
      return { ext: "emf", data: data.subarray(0, cbSize || data.length) };
    }
    case 0xf01d: case 0xf02a: return { ext: "jpeg", data: b.subarray(s + uids + 1, e) };
    case 0xf01e: return { ext: "png", data: b.subarray(s + uids + 1, e) };
    case 0xf029: return { ext: "tiff", data: b.subarray(s + uids + 1, e) };
    case 0xf01f: {
      // DIB → BMP：补 14 字节文件头
      const dib = b.subarray(s + uids + 1, e);
      if (dib.length < 40) return undefined;
      const hdr = Buffer.alloc(14);
      const bits = dib.readUInt16LE(14), clr = dib.readUInt32LE(32), hs = dib.readUInt32LE(0);
      const pal = (clr || (bits <= 8 ? 1 << bits : 0)) * 4;
      hdr.write("BM", 0, "latin1"); hdr.writeUInt32LE(14 + dib.length, 2); hdr.writeUInt32LE(14 + hs + pal, 10);
      return { ext: "bmp", data: Buffer.concat([hdr, dib]) };
    }
  }
  return undefined;
}

/** 补"可放置"图元文件头：外框取图元文件的逻辑坐标，每英寸的单位数由显示宽度（EMU）换算 */
function placeableWmf(data: Buffer, l: number, t: number, r: number, b: number, emuW: number): Buffer {
  if (data.length >= 4 && data.readUInt32LE(0) === 0x9ac6cdd7) return data;
  const clamp = (v: number) => Math.max(-32768, Math.min(32767, v));
  const h = Buffer.alloc(22);
  h.writeUInt32LE(0x9ac6cdd7, 0);
  h.writeInt16LE(clamp(l), 6); h.writeInt16LE(clamp(t), 8); h.writeInt16LE(clamp(r), 10); h.writeInt16LE(clamp(b), 12);
  const inch = emuW > 0 && r > l ? Math.round(((r - l) * 914400) / emuW) : 1440;
  h.writeUInt16LE(Math.max(1, Math.min(65535, inch)), 14);
  let x = 0;
  for (let i = 0; i < 20; i += 2) x ^= h.readUInt16LE(i);
  h.writeUInt16LE(x, 20);
  return Buffer.concat([h, data]);
}

/**
 * 兼容性选项（Dop2000 起的 Copts，位于 Dop 第 508 字节；其中的 copts80 必须与 Dop97 第 84 字节处相同，用来校验）。
 * 只取 Copts 中 copts80 之后的 33 个开关，按位序对应 OOXML 的 <w:compat> 子元素；
 * fDontAdjustLineHeightInTable 的含义与 adjustLineHeightInTable 相反。这些开关决定 Word 的排版细节
 * （编号后的制表位、表格自动调整、东亚断行等），丢掉会让同一份文档在 Word 里重新排版。
 */
const COPTS = [
  "shapeLayoutLikeWW8", "footnoteLayoutLikeWW8", "doNotUseHTMLParagraphAutoSpacing", "!adjustLineHeightInTable", "forgetLastTabAlignment",
  "autoSpaceLikeWord95", "alignTablesRowByRow", "layoutRawTableWidth", "layoutTableRowsApart", "useWord97LineBreakRules",
  "doNotBreakWrappedTables", "doNotSnapToGridInCell", "selectFldWithFirstOrLastChar", "applyBreakingRules", "doNotWrapTextWithPunct",
  "doNotUseEastAsianBreakRules", "useWord2002TableStyleRules", "growAutofit", "useNormalStyleForList", "doNotUseIndentAsNumberingTabStop",
  "useAltKinsokuLineBreakRules", "allowSpaceOfSameStyleInTable", "doNotSuppressIndentation", "doNotAutofitConstrainedTables",
  "autofitToFirstFixedWidthCell", "underlineTabInNumList", "displayHangulFixedWidth", "splitPgBreakAndParaMark", "doNotVertAlignCellWithSp",
  "doNotBreakConstrainedForcedTable", "doNotVertAlignInTxbx", "useAnsiKerningPairs", "cachedColBalance",
];
export function readCompat(dop: Buffer): string[] {
  if (dop.length < 544 || dop.readUInt32LE(84) !== dop.readUInt32LE(508)) return [];
  const lo = dop.readUInt32LE(512), hi = dop.readUInt32LE(516);
  const out: string[] = [];
  COPTS.forEach((name, i) => {
    const on = i < 32 ? (lo >>> i) & 1 : (hi >>> (i - 32)) & 1;
    if (name.startsWith("!")) { if (!on) out.push(name.slice(1)); }
    else if (on) out.push(name);
  });
  return out;
}
