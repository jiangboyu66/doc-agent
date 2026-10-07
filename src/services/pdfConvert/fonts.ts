/**
 * 字体：PDF 字体名 → Word 字体，以及按 Word 字体计算文字宽度
 *
 * PDF 里的字体名五花八门：子集前缀（ABCDEF+）、PostScript 标准 14 字体（Times-Roman）、
 * LaTeX 的 Computer Modern（CMR10 / CMMI10 / CMSY10 …）、URW/Nimbus 替代字体（NimbusRomNo9L-Medi）、
 * TeX Gyre、Liberation、Windows 字体（TimesNewRomanPS-BoldItalicMT）、GBK 编码的中文字体名……
 * 这里统一解析成 Word 中可用的字体 + 粗体/斜体，并标记数学字体。
 *
 * 另外内置 Times New Roman / Arial / Courier New 的逐字宽度（取自度量完全兼容的 Liberation 字体），
 * 用于"逐行一致"模式下的不折行校验：字体被替换时（例如 Computer Modern → Times New Roman），
 * 按 Word 里真正使用的字体计算宽度，而不是 PDF 原字体的宽度。
 */

import fs from "node:fs";
import { fileURLToPath } from "node:url";

export interface FontInfo {
  family: string;
  bold: boolean;
  italic: boolean;
  /** 数学字体（公式） */
  math: boolean;
}

const MAP: Record<string, string> = {
  // Windows / 通用
  timesnewroman: "Times New Roman", timesnewromanps: "Times New Roman", times: "Times New Roman", timesroman: "Times New Roman", timesten: "Times New Roman",
  arial: "Arial", helvetica: "Arial", arialnarrow: "Arial Narrow", arialunicodems: "Arial Unicode MS", helveticaneue: "Arial",
  couriernew: "Courier New", courier: "Courier New", calibri: "Calibri", calibrilight: "Calibri Light", cambria: "Cambria",
  cambriamath: "Cambria Math", georgia: "Georgia", verdana: "Verdana", tahoma: "Tahoma", garamond: "Garamond",
  bookantiqua: "Book Antiqua", palatino: "Palatino Linotype", palatinolinotype: "Palatino Linotype", centurygothic: "Century Gothic",
  centuryschoolbook: "Century Schoolbook", bookmanoldstyle: "Bookman Old Style", segoeui: "Segoe UI", segoeuisymbol: "Segoe UI Symbol",
  symbol: "Symbol", opensymbol: "Symbol", starsymbol: "Symbol", wingdings: "Wingdings", zapfdingbats: "Wingdings", dingbats: "Wingdings",
  aptos: "Aptos", consolas: "Consolas", trebuchetms: "Trebuchet MS", lucidaconsole: "Lucida Console",
  // 中日韩
  simsun: "SimSun", nsimsun: "NSimSun", simhei: "SimHei", kaiti: "KaiTi", kaitigb2312: "KaiTi_GB2312", fangsong: "FangSong",
  fangsonggb2312: "FangSong_GB2312", microsoftyahei: "Microsoft YaHei", microsoftjhenghei: "Microsoft JhengHei",
  dengxian: "DengXian", stsong: "STSong", stheiti: "STHeiti", stkaiti: "STKaiti", stfangsong: "STFangsong", stzhongsong: "STZhongsong",
  pmingliu: "PMingLiU", mingliu: "MingLiU", msmincho: "MS Mincho", msgothic: "MS Gothic", malgungothic: "Malgun Gothic",
  notoserifcjksc: "SimSun", notosanscjksc: "Microsoft YaHei", sourcehanserifsc: "SimSun", sourcehansanssc: "Microsoft YaHei",
  // 度量兼容的开源替代字体 → Word 原字体（字宽完全一致）
  liberationserif: "Times New Roman", liberationsans: "Arial", liberationsansnarrow: "Arial Narrow", liberationmono: "Courier New",
  tinos: "Times New Roman", arimo: "Arial", cousine: "Courier New", carlito: "Calibri", caladea: "Cambria",
  texgyretermes: "Times New Roman", texgyreheros: "Arial", texgyrecursor: "Courier New", texgyrepagella: "Palatino Linotype",
  texgyrebonum: "Bookman Old Style", texgyreschola: "Century Schoolbook", texgyreadventor: "Century Gothic",
  nimbusroman: "Times New Roman", nimbusromno9l: "Times New Roman", nimbusromanno9l: "Times New Roman",
  nimbussans: "Arial", nimbussanl: "Arial", nimbussansnarrow: "Arial Narrow", nimbusmono: "Courier New", nimbusmonl: "Courier New", nimbusmonops: "Courier New",
  urwpalladiol: "Palatino Linotype", p052: "Palatino Linotype", urwbookmanl: "Bookman Old Style", urwbookman: "Bookman Old Style",
  urwgothicl: "Century Gothic", urwgothic: "Century Gothic", centuryschl: "Century Schoolbook", c059: "Century Schoolbook",
  urwchanceryl: "Monotype Corsiva", z003: "Monotype Corsiva",
  freeserif: "Times New Roman", freesans: "Arial", freemono: "Courier New",
  dejavuserif: "Cambria", dejavusans: "Verdana", dejavusansmono: "Consolas",
  stixgeneral: "Times New Roman", stixtwotext: "Times New Roman", stixmath: "Cambria Math", stixtwomath: "Cambria Math", xitsmath: "Cambria Math",
  latinmodernmath: "Cambria Math", lmmath: "Cambria Math", asanamath: "Cambria Math",
};

/** LaTeX Computer Modern / Latin Modern / EC 字体：按字体代号识别族、粗细、斜体与是否数学字体 */
function texFont(name: string): FontInfo | null {
  const n = name.toUpperCase();
  // Latin Modern：LMRoman10-Bold、LMSans10-Regular、LMMono10-Italic、LMMathItalic10-Regular
  let m = /^LM(ROMAN|SANS|MONO|MATHITALIC|MATHSYMBOLS|MATHEXTENSION|ROMANSLANT|ROMANCAPS|ROMANDEMI)\d*/.exec(n);
  if (m) {
    const k = m[1];
    const style = n.slice(m[0].length);
    const bold = /BOLD|DEMI/.test(style + k), italic = /ITALIC|OBLIQUE|SLANT/.test(style + k) || k === "MATHITALIC";
    if (k.startsWith("MATH")) return { family: "Cambria Math", bold, italic: k === "MATHITALIC", math: true };
    return { family: k === "SANS" ? "Arial" : k === "MONO" ? "Courier New" : "Times New Roman", bold, italic, math: false };
  }
  // Computer Modern 与 AMS 字体：CMR10 CMBX12 CMTI10 CMSL10 CMSS10 CMSSBX10 CMTT10 CMMI10 CMMIB10 CMSY10 CMBSY10 CMEX10 MSAM10 MSBM10 EUFM10
  m = /^(CMMIB|CMBSY|CMSSBX|CMSSI|CMSSDC|CMSS|CMBXTI|CMBXSL|CMBX|CMTI|CMSL|CMTT|CMITT|CMSLTT|CMCSC|CMMI|CMSY|CMEX|CMR|CMU|MSAM|MSBM|EUFM|EUFB|EUSM|EURM|RSFS|STMARY|WASY)(\d+)?$/.exec(n);
  if (m) {
    const k = m[1];
    if (/^(CMMI|CMMIB|CMSY|CMBSY|CMEX|MSAM|MSBM|EUFM|EUFB|EUSM|EURM|RSFS|STMARY|WASY)$/.test(k)) {
      return { family: "Cambria Math", bold: k === "CMMIB" || k === "CMBSY" || k === "EUFB", italic: k === "CMMI" || k === "CMMIB", math: true };
    }
    const bold = /BX|SSBX/.test(k), italic = /TI|SL|SSI|ITT/.test(k);
    const family = /^CMSS/.test(k) ? "Arial" : /TT$/.test(k) ? "Courier New" : "Times New Roman";
    return { family, bold, italic, math: false };
  }
  // EC / TC（T1 编码的 CM）：SFRM1000 SFBX1200 SFTI1000 SFTT1000 SFSS1000
  m = /^(SF|TC)(RM|BX|TI|SL|TT|SS|SX|BI|CC)(\d+)$/.exec(n);
  if (m) {
    const k = m[2];
    return { family: k === "SS" || k === "SX" ? "Arial" : k === "TT" ? "Courier New" : "Times New Roman", bold: /BX|SX|BI/.test(k), italic: /TI|SL|BI/.test(k), math: false };
  }
  return null;
}

export function parseFontName(raw: string | undefined): FontInfo {
  let name = (raw ?? "").replace(/^[A-Z]{6}\+/, "").trim();
  // GBK 十六进制编码的中文字体名（如 #CB#CE#CC#E5 = 宋体）
  if (/#[0-9A-F]{2}/i.test(name)) {
    try {
      const bytes = Buffer.from(name.replace(/#([0-9A-F]{2})/gi, (_m, h) => String.fromCharCode(parseInt(h, 16))), "latin1");
      name = new TextDecoder("gbk").decode(bytes);
    } catch { /* 保持原样 */ }
  }
  const tex = texFont(name.split(/[-,+]/)[0]) ?? texFont(name.replace(/[-_,].*$/, ""));
  if (tex) {
    const style = name.includes("-") ? name.slice(name.indexOf("-") + 1) : "";
    if (/bold/i.test(style)) tex.bold = true;
    if (/ital|oblique/i.test(style)) tex.italic = true;
    return tex;
  }

  // 族名与样式：Family-Style / Family,Style / FamilyStyle（如 ArialMT、Arial-BoldItalicMT、TimesNewRomanPS-BoldMT）
  const sep = name.search(/[-,]/);
  let family = sep > 0 ? name.slice(0, sep) : name;
  let style = sep > 0 ? name.slice(sep + 1) : "";
  const fused = /^(.*?)(Bold|Italic|Oblique|BoldItalic|BoldOblique|Black|Medium|Semibold|Light)(MT|PS)?$/.exec(family);
  if (!style && fused && fused[1].length > 2) { family = fused[1]; style = fused[2]; }
  family = family.replace(/(PSMT|PS|MT|Std|Pro|LT|OT)$/g, "").replace(/(PSMT|PS|MT)$/g, "");
  style = style.replace(/(PSMT|PS|MT)$/g, "");

  // 样式缩写：Bold / Medi(Nimbus) / Demi / Semibold / Black / Heavy / Bd；Ital / Italic / Oblique / Obli / It / Slanted
  const bold = /bold|black|heavy|semibold|demi|^medi(ital)?$|^bd|extrabold|ultrabold|^b$|^bi$/i.test(style) || /Bold$/.test(family);
  const italic = /ital|oblique|obli|^it$|slant|^i$|^bi$|kursiv/i.test(style) || /Italic$/.test(family);
  const narrow = /cond|narrow/i.test(style + family);

  const key = family.toLowerCase().replace(/[\s_]/g, "");
  // MathTime（IEEE / Elsevier 等出版社的数学字体）：MTSYN 符号、RMTMI / MTMI 数学斜体、MTEX / BLEX 大型定界符
  if (/^(r?mt(mi|mib|sy|syn|syb|ex|ext|ms|mb)|blex|bmtex|mtpro\w*|mathtime\w*)$/.test(key.replace(/\d+$/, ""))) {
    return { family: "Cambria Math", bold: /b$|bold/i.test(key + style), italic: /mi/.test(key), math: true };
  }
  let out = MAP[key] ?? MAP[key.replace(/\d+$/, "")];
  // 出版社定制字体（Times LT Std、Formata、Giovanni……）：按族名归到 Word 中一定有的同类字体，否则在没有该字体的电脑上会被随意替换
  if (!out) {
    if (/^(times|tmsrm|nimbusrom|ptmr|tirom)/.test(key)) out = "Times New Roman";
    else if (/^(helvetica|helv|arialmt|phvr|formata|frutiger|univers|myriad|futura|gillsans|optima|syntax|segoe|opensans|roboto|lato|sourcesans|sans)/.test(key) || /sans|grotesk/.test(key)) out = narrow ? "Arial Narrow" : "Arial";
    else if (/^(minion|giovanni|utopia|charter|baskerville|caslon|sabon|stempel|bembo|janson|plantin|lucidabright|stix|mtimes|adobetimes|serif)/.test(key)) out = "Times New Roman";
    else if (/^(courier|mono|lucidatypewriter|inconsolata|menlo)/.test(key) || /mono$/.test(key)) out = "Courier New";
  }
  if (out === "Arial" && narrow) out = "Arial Narrow";
  if (!out) out = /^[\x20-\x7e]+$/.test(family) && !/\s/.test(family) ? family.replace(/([a-z])([A-Z])/g, "$1 $2") : family;
  const math = /math/i.test(out) || /^(Symbol)$/.test(out);
  // 无衬线显示字体的 Medium（Formata Medium、Helvetica Neue Medium 等）笔画已接近粗体，替换成 Arial 时用粗体才接近原样
  const medium = /(^|[^a-z])(md|med|medium)([^a-z]|$)|Md$|Medium$|Med$/.test(style) || /(Md|Medium|Med)$/.test(raw?.split(/[-,]/)[0] ?? "") || /OTFMd|Md(It)?$/.test(name.split(/[-,]/)[0]);
  const sansOut = out === "Arial" || out === "Arial Narrow";
  return { family: out.trim() || "Times New Roman", bold: bold || (sansOut && medium), italic: italic || /(It|Italic)$/.test(name.split(/[-,]/)[0]), math };
}

// ---------------------------------------------------------------------------
// Word 字体的字宽
// ---------------------------------------------------------------------------

type Seg = [number, number[]];
let metrics: Record<string, Seg[]> | null = null;
function loadMetrics() {
  if (!metrics) {
    try {
      metrics = JSON.parse(fs.readFileSync(fileURLToPath(new URL("./fontMetrics.json", import.meta.url)), "utf8"));
    } catch {
      metrics = {};
    }
  }
  return metrics!;
}

const CJK = /[⺀-鿿豈-﫿＀-￯　-〿가-힯]/;

/** 有内置字宽表的 Word 字体 */
export function hasMetrics(family: string): boolean {
  return family === "Times New Roman" || family === "Arial" || family === "Courier New";
}

/**
 * 文字在 Word 字体中的宽度（pt）。字宽表里没有的字符返回 null（由调用方回退到 PDF 字宽）；
 * 中日韩全角字符按 1 em 计。
 */
export function measureWord(text: string, family: string, bold: boolean, italic: boolean, size: number): number | null {
  if (!hasMetrics(family)) return null;
  const segs = loadMetrics()[`${family}|${bold ? "b" : ""}${italic ? "i" : ""}`];
  if (!segs) return null;
  let em = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (CJK.test(ch)) { em += 1000; continue; }
    let w = -1;
    for (const [start, vals] of segs) {
      if (cp >= start && cp < start + vals.length) { w = vals[cp - start]; break; }
    }
    if (w < 0) return null;
    em += w;
  }
  return (em / 1000) * size;
}
