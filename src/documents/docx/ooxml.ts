/**
 * OOXML 常量、元素顺序表与样式解析
 *
 * WordprocessingML 的 schema 对很多子元素有严格顺序要求（例如 w:rPr 里 w:b 必须在 w:sz 之前）。
 * 插入属性时如果顺序错了，Word 会报"文件已损坏"。这里把 schema 规定的顺序写成表，所有属性
 * 修改都按表插入到正确位置。
 */

import { XElement, XDocument, parseXml } from "./xml.js";

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
};

export const REL_TYPES = {
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
  comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
};

export const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

/** CT_RPr 子元素顺序 */
export const RPR_ORDER = [
  "ins", "del", "moveFrom", "moveTo",
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
  "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
  "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
  "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
];

/** CT_PPr 子元素顺序 */
export const PPR_ORDER = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
  "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
  "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
  "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
  "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
];

/** CT_TrPr 子元素顺序（修订标记 ins/del 在最后） */
export const TRPR_ORDER = [
  "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight",
  "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange",
];

/** 在父元素中按 schema 顺序找到新子元素应插入的位置下标 */
export function orderedInsertIndex(parent: XElement, localName: string, order: string[], prefix: string): number {
  const rank = order.indexOf(localName);
  const kids = parent.children;
  for (let i = 0; i < kids.length; i++) {
    const k = kids[i];
    if (k.type !== "el" || k.prefix !== prefix) continue;
    const r = order.indexOf(k.local);
    if (r > rank) return i;
  }
  return kids.length;
}

// ---------------- 样式 ----------------

export interface StyleDef {
  id: string;
  name: string;
  type: string;
  basedOn?: string;
  outlineLvl?: number;
  hasNumbering: boolean;
  isDefault: boolean;
}

export class StyleSheet {
  readonly byId = new Map<string, StyleDef>();
  constructor(xml: string | null) {
    if (!xml) return;
    const doc: XDocument = parseXml(xml);
    for (const s of doc.root.elements()) {
      if (s.local !== "style") continue;
      const p = s.prefix;
      const id = s.getAttr(`${p}:styleId`) ?? "";
      const name = s.child(`${p}:name`)?.getAttr(`${p}:val`) ?? id;
      const pPr = s.child(`${p}:pPr`);
      const ol = pPr?.child(`${p}:outlineLvl`)?.getAttr(`${p}:val`);
      this.byId.set(id, {
        id,
        name,
        type: s.getAttr(`${p}:type`) ?? "paragraph",
        basedOn: s.child(`${p}:basedOn`)?.getAttr(`${p}:val`),
        outlineLvl: ol !== undefined ? Number(ol) : undefined,
        hasNumbering: !!pPr?.child(`${p}:numPr`),
        isDefault: s.getAttr(`${p}:default`) === "1",
      });
    }
  }

  get(id: string | undefined): StyleDef | undefined {
    return id ? this.byId.get(id) : undefined;
  }

  defaultParagraphStyle(): StyleDef | undefined {
    for (const s of this.byId.values()) if (s.type === "paragraph" && s.isDefault) return s;
    return undefined;
  }

  /** 支持按样式 ID 或显示名称查找（不区分大小写） */
  resolve(idOrName: string, type?: string): StyleDef | undefined {
    const direct = this.byId.get(idOrName);
    if (direct && (!type || direct.type === type)) return direct;
    const lower = idOrName.toLowerCase();
    for (const s of this.byId.values()) {
      if ((!type || s.type === type) && (s.name.toLowerCase() === lower || s.id.toLowerCase() === lower)) return s;
    }
    return undefined;
  }

  /** 沿 basedOn 链查找大纲级别 */
  outlineLevel(id: string | undefined): number | undefined {
    const seen = new Set<string>();
    let cur = this.get(id);
    while (cur && !seen.has(cur.id)) {
      if (cur.outlineLvl !== undefined) return cur.outlineLvl;
      seen.add(cur.id);
      cur = this.get(cur.basedOn);
    }
    return undefined;
  }

  /**
   * 标题级别推断：先看 schema 意义上的大纲级别，再按样式名称启发式识别。
   * 很多期刊模板（如 IEEE Access 的 H1_List / H2_First / H3）不设置 outlineLvl，只能靠名称。
   */
  headingLevel(id: string | undefined): number | undefined {
    const ol = this.outlineLevel(id);
    if (ol !== undefined && ol < 9) return ol + 1;
    const seen = new Set<string>();
    let cur = this.get(id);
    while (cur && !seen.has(cur.id)) {
      const n = cur.name;
      let m = /^heading\s*(\d)$/i.exec(n) || /^标题\s*(\d)$/.exec(n);
      if (m) return Number(m[1]);
      m = /^H(\d)(?:[_\s(]|$)/.exec(n);
      if (m) return Number(m[1]);
      seen.add(cur.id);
      cur = this.get(cur.basedOn);
    }
    return undefined;
  }

  isTitle(id: string | undefined): boolean {
    const s = this.get(id);
    return !!s && /^(title|paper title|标题)$/i.test(s.name);
  }
}

// ---------------- 关系文件 ----------------

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export function parseRels(xml: string | null): Relationship[] {
  if (!xml) return [];
  const doc = parseXml(xml);
  return doc.root.elements().map((r) => ({
    id: r.getAttr("Id") ?? "",
    type: r.getAttr("Type") ?? "",
    target: r.getAttr("Target") ?? "",
    external: r.getAttr("TargetMode") === "External",
  }));
}

export function relsPathFor(part: string): string {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}

export function resolveTarget(fromPart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const base = fromPart.slice(0, fromPart.lastIndexOf("/") + 1);
  const parts = (base + target).split("/");
  const out: string[] = [];
  for (const p of parts) {
    if (p === "..") out.pop();
    else if (p !== "." && p !== "") out.push(p);
  }
  return out.join("/");
}

/** 修订日期格式：Word 使用不带毫秒的 ISO 8601 */
export function revisionDate(d = new Date()): string {
  return d.toISOString().replace(/\.\d{3}Z$/, "Z");
}

export function ptToTwips(pt: number): number {
  return Math.round(pt * 20);
}
