/**
 * .doc 的格式指令（sprm）→ Word 2007+ 的 XML 属性（rPr / pPr / tblPr / trPr / tcPr / sectPr）。
 * 属性按 schema 规定的顺序输出（Word 遇到顺序不对的元素会报"文件已损坏"）。
 */

import { PPR_ORDER, RPR_ORDER } from "../../documents/docx/ooxml.js";
import { S, type Sprm } from "./sprm.js";

/** XML 转义；去掉 XML 不允许的字符（控制字符、孤立的代理项）——损坏的文件里样式名、编号文字中会出现 */
export const esc = (s: string) =>
  s.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** 旧版 16 色索引 */
const ICO = ["auto", "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF", "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0"];
const HIGHLIGHT = ["", "black", "blue", "cyan", "green", "magenta", "red", "yellow", "white", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray"];
const UNDERLINE: Record<number, string> = {
  1: "single", 2: "words", 3: "double", 4: "dotted", 6: "thick", 7: "dash", 9: "dotDash", 10: "dotDotDash", 11: "wave",
  20: "dottedHeavy", 23: "dashedHeavy", 25: "dashDotHeavy", 26: "dashDotDotHeavy", 27: "wavyHeavy", 39: "dashLong", 43: "wavyDouble", 55: "dashLongHeavy",
};
const BORDER: Record<number, string> = {
  1: "single", 2: "thick", 3: "double", 5: "single", 6: "dotted", 7: "dashed", 8: "dotDash", 9: "dotDotDash", 10: "triple",
  11: "thinThickSmallGap", 12: "thickThinSmallGap", 13: "thinThickThinSmallGap", 14: "thinThickMediumGap", 15: "thickThinMediumGap",
  16: "thinThickThinMediumGap", 17: "thinThickLargeGap", 18: "thickThinLargeGap", 19: "thinThickThinLargeGap", 20: "wave", 21: "doubleWave",
  22: "dashSmallGap", 23: "dashDotStroked", 24: "threeDEmboss", 25: "threeDEngrave", 26: "outset", 27: "inset",
};
const PCT = ["clear", "solid", "pct5", "pct10", "pct20", "pct25", "pct30", "pct40", "pct50", "pct60", "pct70", "pct75", "pct80", "pct90",
  "horzStripe", "vertStripe", "reverseDiagStripe", "diagStripe", "horzCross", "diagCross", "thinHorzStripe", "thinVertStripe",
  "thinReverseDiagStripe", "thinDiagStripe", "thinHorzCross", "thinDiagCross"];

/** Windows 语言代码（LCID）→ BCP 47：决定断行规则、拼写检查语言（中文按中文规则处理行首行尾标点） */
const LCID: Record<number, string> = {
  0x0409: "en-US", 0x0809: "en-GB", 0x0c09: "en-AU", 0x1009: "en-CA", 0x0804: "zh-CN", 0x0404: "zh-TW", 0x0c04: "zh-HK", 0x1004: "zh-SG", 0x1404: "zh-MO",
  0x0411: "ja-JP", 0x0412: "ko-KR", 0x0419: "ru-RU", 0x0407: "de-DE", 0x040c: "fr-FR", 0x0c0a: "es-ES", 0x040a: "es-ES", 0x0410: "it-IT", 0x0416: "pt-BR",
  0x0816: "pt-PT", 0x0413: "nl-NL", 0x0415: "pl-PL", 0x0405: "cs-CZ", 0x041b: "sk-SK", 0x040e: "hu-HU", 0x0418: "ro-RO", 0x0402: "bg-BG", 0x0408: "el-GR",
  0x041f: "tr-TR", 0x041d: "sv-SE", 0x0406: "da-DK", 0x0414: "nb-NO", 0x040b: "fi-FI", 0x0422: "uk-UA", 0x040d: "he-IL", 0x0401: "ar-SA", 0x041e: "th-TH",
  0x042a: "vi-VN", 0x0421: "id-ID", 0x0439: "hi-IN", 0x0424: "sl-SI", 0x041a: "hr-HR", 0x081a: "sr-Latn-CS", 0x0c1a: "sr-Cyrl-CS", 0x0425: "et-EE", 0x0426: "lv-LV", 0x0427: "lt-LT",
};

/** COLORREF：R G B + 标志字节（0xFF = 自动） */
export function cv(b: Buffer, o = 0): string {
  if (b.length < o + 4) return "auto";
  if (b[o + 3] === 0xff) return "auto";
  return [b[o], b[o + 1], b[o + 2]].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
}
const ico = (i: number) => ICO[i] ?? "auto";

/** 边框：旧版 Brc80（4 字节）或新版 Brc（8 字节）。返回 undefined 表示"未指定" */
export function border(tag: string, b: Buffer, o = 0, wide = b.length - o >= 8): string | undefined {
  if (wide) {
    if (b.length < o + 8) return undefined;
    const color = cv(b, o), width = b[o + 4], type = b[o + 5], space = b[o + 6] & 0x1f, shadow = (b[o + 6] & 0x20) !== 0;
    if (type === 0xff || type === 0) return `<w:${tag} w:val="${type === 0xff ? "nil" : "none"}"/>`;
    return `<w:${tag} w:val="${BORDER[type] ?? "single"}" w:sz="${width || 4}" w:space="${space}" w:color="${color}"${shadow ? ' w:shadow="1"' : ""}/>`;
  }
  if (b.length < o + 4) return undefined;
  const v = b.readUInt32LE(o);
  if (v === 0) return undefined;
  if (v === 0xffffffff) return `<w:${tag} w:val="nil"/>`;
  const width = b[o], type = b[o + 1], color = ico(b[o + 2]), space = b[o + 3] & 0x1f;
  if (!type) return `<w:${tag} w:val="none"/>`;
  return `<w:${tag} w:val="${BORDER[type] ?? "single"}" w:sz="${width || 4}" w:space="${space}" w:color="${color}"/>`;
}

/** 底纹：SHD80（2 字节，颜色索引）或 SHD（10 字节，RGB） */
export function shading(b: Buffer, o = 0, wide = b.length - o >= 10): string | undefined {
  let fore: string, back: string, pat: number;
  if (wide) {
    if (b.length < o + 10) return undefined;
    fore = cv(b, o); back = cv(b, o + 4); pat = b.readUInt16LE(o + 8);
  } else {
    if (b.length < o + 2) return undefined;
    const v = b.readUInt16LE(o);
    fore = ico(v & 0x1f); back = ico((v >> 5) & 0x1f); pat = v >> 10;
  }
  if (pat === 0xffff) return undefined;
  if (pat === 0 && back === "auto") return undefined;
  // 实心底纹：用前景色铺满（直接写成填充色，快速预览也能显示）
  if (pat === 1) return `<w:shd w:val="clear" w:color="auto" w:fill="${fore === "auto" ? "000000" : fore}"/>`;
  return `<w:shd w:val="${PCT[pat] ?? "clear"}" w:color="${fore}" w:fill="${back}"/>`;
}

/** 按 schema 顺序拼接属性 */
function ordered(map: Map<string, string>, order: string[]): string {
  return [...map.entries()].sort((a, b) => order.indexOf(a[0]) - order.indexOf(b[0])).map(([, v]) => v).join("");
}

export interface CharCtx {
  fonts: string[];
  styleIds: Map<number, string>;
  /** 开关属性取"与样式相反"（0x81）时需要样式里的值 */
  styleToggle: (op: number) => boolean;
}

const TOGGLES: Array<[number, string]> = [
  [S.CFBold, "b"], [S.CFBoldBi, "bCs"], [S.CFItalic, "i"], [S.CFItalicBi, "iCs"], [S.CFCaps, "caps"], [S.CFSmallCaps, "smallCaps"],
  [S.CFStrike, "strike"], [S.CFOutline, "outline"], [S.CFShadow, "shadow"], [S.CFEmboss, "emboss"], [S.CFImprint, "imprint"], [S.CFVanish, "vanish"],
];

/** 字符格式 → rPr 子元素（不含外层 <w:rPr>） */
export function charXml(sprms: Sprm[], ctx: CharCtx, opts: { skipStyle?: boolean } = {}): string {
  const m = new Map<string, string>();
  const fonts: Record<string, string> = {};
  const lang: Record<string, string> = {};
  let color: string | undefined;
  for (const s of sprms) {
    const a = s.arg;
    const tog = TOGGLES.find(([op]) => op === s.op);
    if (tog) {
      const v = a[0];
      if (v === 0x80) { m.delete(tog[1]); continue; }
      const on = v === 0x81 ? !ctx.styleToggle(s.op) : v === 1;
      m.set(tog[1], on ? `<w:${tog[1]}/>` : `<w:${tog[1]} w:val="0"/>`);
      continue;
    }
    switch (s.op) {
      case S.CIstd: if (!opts.skipStyle) { const id = ctx.styleIds.get(a.readUInt16LE(0)); if (id) m.set("rStyle", `<w:rStyle w:val="${id}"/>`); } break;
      case S.CFDStrike: m.set("dstrike", a[0] ? "<w:dstrike/>" : '<w:dstrike w:val="0"/>'); break;
      case S.CHps: m.set("sz", `<w:sz w:val="${a.readUInt16LE(0)}"/>`); break;
      case S.CHpsBi: m.set("szCs", `<w:szCs w:val="${a.readUInt16LE(0)}"/>`); break;
      case S.CKul: { const u = a[0]; m.set("u", `<w:u w:val="${u ? UNDERLINE[u] ?? "single" : "none"}"/>`); break; }
      case S.CIco: color = ico(a[0]); break;
      case S.CCv: color = cv(a); break;
      case S.CIss: m.set("vertAlign", `<w:vertAlign w:val="${["baseline", "superscript", "subscript"][a[0]] ?? "baseline"}"/>`); break;
      case S.CHighlight: m.set("highlight", a[0] ? `<w:highlight w:val="${HIGHLIGHT[a[0]] ?? "yellow"}"/>` : '<w:highlight w:val="none"/>'); break;
      case S.CRgFtc0: fonts.ascii = ctx.fonts[a.readUInt16LE(0)] ?? ""; break;
      case S.CRgFtc1: fonts.eastAsia = ctx.fonts[a.readUInt16LE(0)] ?? ""; break;
      case S.CRgFtc2: fonts.hAnsi = ctx.fonts[a.readUInt16LE(0)] ?? ""; break;
      case S.CFtcBi: fonts.cs = ctx.fonts[a.readUInt16LE(0)] ?? ""; break;
      case S.CDxaSpace: m.set("spacing", `<w:spacing w:val="${a.readInt16LE(0)}"/>`); break;
      case S.CHpsPos: m.set("position", `<w:position w:val="${a.readInt16LE(0)}"/>`); break;
      case S.CCharScale: m.set("w", `<w:w w:val="${a.readUInt16LE(0)}"/>`); break;
      case S.CShd80: { const x = shading(a, 0, false); if (x) m.set("shd", x); break; }
      case S.CShd: { const x = shading(a, 0, true); if (x) m.set("shd", x); break; }
      case S.CBrc80: case S.CBrc: { const x = border("bdr", a, 0, s.op === S.CBrc); if (x) m.set("bdr", x); break; }
      case S.CRgLid0_80: case S.CRgLid0: { const t = LCID[a.readUInt16LE(0)]; if (t) lang.val = t; break; }
      case S.CRgLid1_80: case S.CRgLid1: { const t = LCID[a.readUInt16LE(0)]; if (t) lang.eastAsia = t; break; }
      case S.CLidBi: { const t = LCID[a.readUInt16LE(0)]; if (t) lang.bidi = t; break; }
    }
  }
  const fa = Object.entries(fonts).filter(([, v]) => v);
  if (fa.length) m.set("rFonts", `<w:rFonts ${fa.map(([k, v]) => `w:${k}="${esc(v)}"`).join(" ")}/>`);
  if (color) m.set("color", `<w:color w:val="${color}"/>`);
  const la = Object.entries(lang);
  if (la.length) m.set("lang", `<w:lang ${la.map(([k, v]) => `w:${k}="${v}"`).join(" ")}/>`);
  return ordered(m, RPR_ORDER);
}

export interface ParaCtx {
  styleIds: Map<number, string>;
  /** 段落的列表编号：返回 numId（已在 numbering.xml 中定义）或 undefined */
  numId: (ilfo: number) => number | undefined;
}

/** 段落格式 → pPr 子元素（不含外层 <w:pPr>，不含 rPr / sectPr） */
export function paraXml(sprms: Sprm[], ctx: ParaCtx, istd?: number): string {
  const m = new Map<string, string>();
  const ind: Record<string, number> = {};
  const sp: Record<string, string | number> = {};
  const bdr: Record<string, string> = {};
  let ilvl: number | undefined, ilfo: number | undefined;
  let tabs: Map<number, string> | undefined;
  if (istd !== undefined) { const id = ctx.styleIds.get(istd); if (id) m.set("pStyle", `<w:pStyle w:val="${id}"/>`); }
  const flag = (k: string, v: number) => m.set(k, v ? `<w:${k}/>` : `<w:${k} w:val="0"/>`);
  for (const s of sprms) {
    const a = s.arg;
    switch (s.op) {
      case S.PIstd: { const id = ctx.styleIds.get(a.readUInt16LE(0)); if (id && istd === undefined) m.set("pStyle", `<w:pStyle w:val="${id}"/>`); break; }
      case S.PJc80: case S.PJc: m.set("jc", `<w:jc w:val="${["left", "center", "right", "both", "distribute", "both", "distribute", "lowKashida", "thaiDistribute", "highKashida"][a[0]] ?? "left"}"/>`); break;
      case S.PFKeep: flag("keepLines", a[0]); break;
      case S.PFKeepFollow: flag("keepNext", a[0]); break;
      case S.PFPageBreakBefore: flag("pageBreakBefore", a[0]); break;
      case S.PFWidowControl: flag("widowControl", a[0]); break;
      case S.PFKinsoku: flag("kinsoku", a[0]); break;
      case S.PFAutoSpaceDE: flag("autoSpaceDE", a[0]); break;
      case S.PFAutoSpaceDN: flag("autoSpaceDN", a[0]); break;
      case S.PFBiDi: flag("bidi", a[0]); break;
      case S.PFContextualSpacing: flag("contextualSpacing", a[0]); break;
      case S.PFNoAutoHyph: flag("suppressAutoHyphens", a[0]); break;
      case S.POutLvl: if (a[0] < 9) m.set("outlineLvl", `<w:outlineLvl w:val="${a[0]}"/>`); break;
      case S.PIlvl: ilvl = a[0]; break;
      case S.PIlfo: ilfo = a.readInt16LE(0); break;
      case S.PDxaLeft80: case S.PDxaLeft: ind.left = a.readInt16LE(0); break;
      case S.PDxaRight80: case S.PDxaRight: ind.right = a.readInt16LE(0); break;
      case S.PDxaLeft180: case S.PDxaLeft1: ind.first = a.readInt16LE(0); break;
      case S.PDxcLeft: ind.leftChars = a.readInt16LE(0); break;
      case S.PDxcRight: ind.rightChars = a.readInt16LE(0); break;
      case S.PDxcLeft1: ind.firstChars = a.readInt16LE(0); break;
      case S.PDyaBefore: sp.before = a.readUInt16LE(0); break;
      case S.PDyaAfter: sp.after = a.readUInt16LE(0); break;
      case S.PDylBefore: sp.beforeLines = a.readInt16LE(0); break;
      case S.PDylAfter: sp.afterLines = a.readInt16LE(0); break;
      case S.PFDyaBeforeAuto: sp.beforeAutospacing = a[0] ? 1 : 0; break;
      case S.PFDyaAfterAuto: sp.afterAutospacing = a[0] ? 1 : 0; break;
      case S.PDyaLine: {
        const line = a.readInt16LE(0), mult = a.readInt16LE(2);
        if (mult) { sp.line = line; sp.lineRule = "auto"; }
        else { sp.line = Math.abs(line); sp.lineRule = line < 0 ? "exact" : "atLeast"; }
        break;
      }
      case S.PBrcTop80: case S.PBrcTop: { const x = border("top", a, 0, s.op === S.PBrcTop); if (x) bdr.top = x; break; }
      case S.PBrcLeft80: case S.PBrcLeft: { const x = border("left", a, 0, s.op === S.PBrcLeft); if (x) bdr.left = x; break; }
      case S.PBrcBottom80: case S.PBrcBottom: { const x = border("bottom", a, 0, s.op === S.PBrcBottom); if (x) bdr.bottom = x; break; }
      case S.PBrcRight80: case S.PBrcRight: { const x = border("right", a, 0, s.op === S.PBrcRight); if (x) bdr.right = x; break; }
      case S.PBrcBetween80: case S.PBrcBetween: { const x = border("between", a, 0, s.op === S.PBrcBetween); if (x) bdr.between = x; break; }
      case S.PShd80: { const x = shading(a, 0, false); if (x) m.set("shd", x); break; }
      case S.PShd: { const x = shading(a, 0, true); if (x) m.set("shd", x); break; }
      case S.PChgTabsPapx: case S.PChgTabs: tabs = applyTabs(tabs ?? new Map(), a, s.op === S.PChgTabs); break;
    }
  }
  if (ilfo !== undefined || ilvl !== undefined) {
    if (ilfo === 0 || ilfo === 2047 || (ilfo !== undefined && ilfo < 0)) m.set("numPr", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>');
    else if (ilfo !== undefined) {
      const id = ctx.numId(ilfo);
      if (id !== undefined) m.set("numPr", `<w:numPr><w:ilvl w:val="${ilvl ?? 0}"/><w:numId w:val="${id}"/></w:numPr>`);
    } else m.set("numPr", `<w:numPr><w:ilvl w:val="${ilvl}"/></w:numPr>`);
  }
  if (Object.keys(ind).length) {
    const at: string[] = [];
    if (ind.leftChars !== undefined) at.push(`w:leftChars="${ind.leftChars}"`);
    if (ind.left !== undefined) at.push(`w:left="${ind.left}"`);
    if (ind.rightChars !== undefined) at.push(`w:rightChars="${ind.rightChars}"`);
    if (ind.right !== undefined) at.push(`w:right="${ind.right}"`);
    if (ind.firstChars !== undefined) at.push(ind.firstChars < 0 ? `w:hangingChars="${-ind.firstChars}"` : `w:firstLineChars="${ind.firstChars}"`);
    if (ind.first !== undefined) at.push(ind.first < 0 ? `w:hanging="${-ind.first}"` : `w:firstLine="${ind.first}"`);
    m.set("ind", `<w:ind ${at.join(" ")}/>`);
  }
  if (Object.keys(sp).length) {
    const at = ["before", "beforeLines", "beforeAutospacing", "after", "afterLines", "afterAutospacing", "line", "lineRule"]
      .filter((k) => sp[k] !== undefined).map((k) => `w:${k}="${sp[k]}"`);
    m.set("spacing", `<w:spacing ${at.join(" ")}/>`);
  }
  const bk = ["top", "left", "bottom", "right", "between"].filter((k) => bdr[k]);
  if (bk.length) m.set("pBdr", `<w:pBdr>${bk.map((k) => bdr[k]).join("")}</w:pBdr>`);
  if (tabs?.size) m.set("tabs", `<w:tabs>${[...tabs.entries()].sort((x, y) => x[0] - y[0]).map(([, v]) => v).join("")}</w:tabs>`);
  return ordered(m, PPR_ORDER);
}

const TAB_JC = ["left", "center", "right", "decimal", "bar", "left", "left", "left"];
const TAB_LEADER = ["none", "dot", "hyphen", "underscore", "heavy", "middleDot", "none", "none"];

function applyTabs(tabs: Map<number, string>, a: Buffer, withClose: boolean): Map<number, string> {
  let o = 0;
  if (a.length < 2) return tabs;
  const nDel = a[o++];
  for (let i = 0; i < nDel && o + 2 <= a.length; i++, o += 2) {
    const pos = a.readInt16LE(o);
    tabs.set(pos, `<w:tab w:val="clear" w:pos="${pos}"/>`);
  }
  if (withClose) o += nDel * 2; // rgdxaClose
  if (o >= a.length) return tabs;
  const nAdd = a[o++];
  const posAt = o, tbdAt = o + nAdd * 2;
  for (let i = 0; i < nAdd && tbdAt + i < a.length; i++) {
    const pos = a.readInt16LE(posAt + i * 2), tbd = a[tbdAt + i];
    tabs.set(pos, `<w:tab w:val="${TAB_JC[tbd & 7]}" w:leader="${TAB_LEADER[(tbd >> 3) & 7]}" w:pos="${pos}"/>`);
  }
  return tabs;
}

// ---------------------------------------------------------------------------
// 节
// ---------------------------------------------------------------------------

export function sectXml(sprms: Sprm[], refs: string): string {
  const v: Record<string, number> = { bkc: 2, xaPage: 12240, yaPage: 15840, left: 1800, right: 1800, top: 1440, bottom: 1440, hdr: 720, ftr: 720, gutter: 0, cols: 0, colSpace: 720 };
  let orient = 0, titlePg = false, nfc: number | undefined, pgStart: number | undefined, restart = false, grid: number | undefined, pitch: number | undefined, charSpace: number | undefined, vjc: number | undefined, sep = false;
  const colW: Array<{ w?: number; sp?: number }> = [];
  let evenly = true;
  for (const s of sprms) {
    const a = s.arg;
    switch (s.op) {
      case S.SBkc: v.bkc = a[0]; break;
      case S.SFTitlePage: titlePg = a[0] !== 0; break;
      case S.SCcolumns: v.cols = a.readUInt16LE(0); break;
      case S.SDxaColumns: v.colSpace = a.readUInt16LE(0); break;
      case S.SNfcPgn: nfc = a[0]; break;
      case S.SFPgnRestart: restart = a[0] !== 0; break;
      case S.SPgnStart97: pgStart = a.readUInt16LE(0); break;
      case S.SPgnStart: pgStart = a.readUInt32LE(0); break;
      case S.SBOrientation: orient = a[0]; break;
      case S.SXaPage: v.xaPage = a.readUInt16LE(0); break;
      case S.SYaPage: v.yaPage = a.readUInt16LE(0); break;
      case S.SDxaLeft: v.left = a.readUInt16LE(0); break;
      case S.SDxaRight: v.right = a.readUInt16LE(0); break;
      case S.SDyaTop: v.top = a.readInt16LE(0); break;
      case S.SDyaBottom: v.bottom = a.readInt16LE(0); break;
      case S.SDzaGutter: v.gutter = a.readUInt16LE(0); break;
      case S.SDyaHdrTop: v.hdr = a.readUInt16LE(0); break;
      case S.SDyaHdrBottom: v.ftr = a.readUInt16LE(0); break;
      case S.SLBetween: sep = a[0] !== 0; break;
      case S.SVjc: vjc = a[0]; break;
      case S.SClm: grid = a.readUInt16LE(0); break;
      case S.SDyaLinePitch: pitch = a.readUInt16LE(0); break;
      case S.SDxtCharSpace: charSpace = a.readInt32LE(0); break;
      case S.SFEvenlySpaced: evenly = a[0] !== 0; break;
      case S.SDxaColWidth: if (a.length >= 3) (colW[a[0]] ??= {}).w = a.readUInt16LE(1); break;
      case S.SDxaColSpacing: if (a.length >= 3) (colW[a[0]] ??= {}).sp = a.readUInt16LE(1); break;
    }
  }
  const type = ["continuous", "nextColumn", "nextPage", "evenPage", "oddPage"][v.bkc] ?? "nextPage";
  const fmt = nfc !== undefined ? NUMFMT[nfc] : undefined;
  const pg = `<w:pgNumType${fmt && fmt !== "decimal" ? ` w:fmt="${fmt}"` : ""}${restart && pgStart !== undefined ? ` w:start="${pgStart}"` : ""}/>`;
  const n = v.cols + 1;
  const cols = n > 1 && !evenly && colW.length
    ? `<w:cols w:num="${n}" w:space="${v.colSpace}" w:equalWidth="0"${sep ? ' w:sep="1"' : ""}>${Array.from({ length: n }, (_, i) => `<w:col w:w="${colW[i]?.w ?? 0}"${i < n - 1 ? ` w:space="${colW[i]?.sp ?? v.colSpace}"` : ""}/>`).join("")}</w:cols>`
    : `<w:cols w:space="${v.colSpace}"${n > 1 ? ` w:num="${n}"` : ""}${sep ? ' w:sep="1"' : ""}/>`;
  const gridType = grid === 1 ? "lines" : grid === 2 ? "linesAndChars" : grid === 3 ? "snapToChars" : undefined;
  return (
    `<w:sectPr>${refs}<w:type w:val="${type}"/>` +
    `<w:pgSz w:w="${v.xaPage}" w:h="${v.yaPage}"${orient === 2 ? ' w:orient="landscape"' : ""}/>` +
    `<w:pgMar w:top="${v.top}" w:right="${v.right}" w:bottom="${v.bottom}" w:left="${v.left}" w:header="${v.hdr}" w:footer="${v.ftr}" w:gutter="${v.gutter}"/>` +
    pg + cols + (vjc ? `<w:vAlign w:val="${["top", "center", "both", "bottom"][vjc] ?? "top"}"/>` : "") + (titlePg ? "<w:titlePg/>" : "") +
    (gridType ? `<w:docGrid w:type="${gridType}"${pitch ? ` w:linePitch="${pitch}"` : ""}${charSpace !== undefined && gridType !== "lines" ? ` w:charSpace="${charSpace}"` : ""}/>` : pitch ? `<w:docGrid w:linePitch="${pitch}"/>` : "") +
    `</w:sectPr>`
  );
}

/** 编号格式（MSONFC → ST_NumberFormat，两者的顺序相同） */
export const NUMFMT: Record<number, string> = Object.fromEntries([
  "decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter", "ordinal", "cardinalText", "ordinalText", "hex", "chicago",
  "ideographDigital", "japaneseCounting", "aiueo", "iroha", "decimalFullWidth", "decimalHalfWidth", "japaneseLegal", "japaneseDigitalTenThousand",
  "decimalEnclosedCircle", "decimalFullWidth2", "aiueoFullWidth", "irohaFullWidth", "decimalZero", "bullet", "ganada", "chosung",
  "decimalEnclosedFullstop", "decimalEnclosedParen", "decimalEnclosedCircleChinese", "ideographEnclosedCircle", "ideographTraditional",
  "ideographZodiac", "ideographZodiacTraditional", "taiwaneseCounting", "ideographLegalTraditional", "taiwaneseCountingThousand",
  "taiwaneseDigital", "chineseCounting", "chineseLegalSimplified", "chineseCountingThousand", "koreanDigital", "koreanCounting",
  "koreanLegal", "koreanDigital2", "vietnameseCounting", "russianLower", "russianUpper", "none", "numberInDash", "hebrew1", "hebrew2",
  "arabicAlpha", "arabicAbjad", "hindiVowels", "hindiConsonants", "hindiNumbers", "hindiCounting", "thaiLetters", "thaiNumbers", "thaiCounting",
].map((f, i) => [i, f]).concat([[255, "none"]]));
