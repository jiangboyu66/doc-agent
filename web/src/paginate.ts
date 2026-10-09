/**
 * 快速预览的分页
 *
 * docx-preview 不会自己分页：只在分页符、分节符，以及文档里记录的"上次分页位置"处分页。
 * 文档内容增减后（例如扩写摘要），那一页会被撑长，两栏的底部也对不齐。这里按 Word 的规则重新分页：
 *   - 每页正文区高度固定；多栏节先填满左栏再填右栏（column-fill: auto），节的最后一页两栏平衡；
 *   - 放不下的段落按行拆到下一页：前半段末行两端对齐，续段不再首行缩进、没有段前距；
 *   - 表格、图片等整体移到下一页；比一整页还高的内容单独占一页（页面随之变长）；
 *   - 页眉页脚取原渲染中对应页（首页 / 偶数页 / 奇数页）的那一份。
 * 内容本来就放得下的页面（例如逐页保留原排版的转换结果）保持不变。
 */

import JSZip from "jszip";

const MAX_PAGES = 3000;

/**
 * 去掉文档里记录的"上次分页位置"（lastRenderedPageBreak）后再渲染正文：由 paginate() 重新分页。
 * （docx-preview 的 ignoreLastRenderedPageBreak 选项会把每个分节都拆成单独一页，不能用。）
 */
export async function stripRenderedBreaks(blob: Blob): Promise<Blob> {
  return (await prepareDocx(blob)).blob;
}

/** 页眉页脚中的页码域占位符：docx-preview 不渲染域，渲染后由 numberPages() 换成实际页码 */
const PH_PAGE = "\uE000PAGE\uE000";
const PH_NUMPAGES = "\uE000NUMPAGES\uE000";

function markFields(xml: string): string {
  // 简单域 <w:fldSimple w:instr=" PAGE ">…</w:fldSimple>
  xml = xml.replace(/<(\w+):fldSimple\b([^>]*?)\binstr="([^"]*)"([^>]*?)(?:\/>|>([\s\S]*?)<\/\1:fldSimple>)/g, (m, p: string, _a: string, instr: string, _b: string, inner?: string) => {
    const ph = /\bNUMPAGES\b/i.test(instr) ? PH_NUMPAGES : /\bPAGE\b/i.test(instr) ? PH_PAGE : null;
    if (!ph) return m;
    const rPr = inner?.match(new RegExp(`<${p}:rPr>[\\s\\S]*?</${p}:rPr>`))?.[0] ?? "";
    return `<${p}:r>${rPr}<${p}:t>${ph}</${p}:t></${p}:r>`;
  });
  // 复杂域：PAGE / NUMPAGES 指令之后、separate 之后的第一段结果文字
  xml = xml.replace(/(<(\w+):instrText[^>]*>\s*(PAGE|NUMPAGES)\b[^<]*<\/\2:instrText>(?:(?!fldCharType="end")[\s\S])*?fldCharType="separate"(?:(?!fldCharType="end")[\s\S])*?<\2:t(?:\s[^>]*)?>)[^<]*(<\/\2:t>)/g,
    (_m, pre: string, _p: string, kind: string, post: string) => pre + (kind.toUpperCase() === "NUMPAGES" ? PH_NUMPAGES : PH_PAGE) + post);
  return xml;
}

/**
 * 渲染前的准备：
 *   - 正文渲染用的副本去掉旧分页位置（由 paginate 重新分页）；
 *   - 页眉页脚中的页码域换成占位符（两份渲染都需要），渲染后填入实际页码；
 *   - 读取"连续分节前不平衡分栏"兼容选项（期刊排版会打开）与起始页码。
 */
export async function prepareDocx(blob: Blob): Promise<{ blob: Blob; plain: Blob; balance: boolean; pageStart: number; wraps: WrapDist[]; anchors: PageAnchor[]; expandBreaks: boolean }> {
  const zip = await JSZip.loadAsync(blob);
  const settings = await zip.file("word/settings.xml")?.async("string");
  const balance = !(settings && /<\w+:noColumnBalance(\s|\/|>)/.test(settings) && !/<\w+:noColumnBalance\s+\w+:val="(0|false)"/.test(settings));
  const mime = blob.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  let changed = false;
  for (const name of Object.keys(zip.files)) {
    if (!/^word\/(header|footer)\d*\.xml$/.test(name)) continue;
    const x = await zip.file(name)!.async("string");
    const y = markFields(x);
    if (y !== x) { zip.file(name, y); changed = true; }
  }
  const f = zip.file("word/document.xml");
  let xml = f ? await f.async("string") : "";
  const pageStart = Number(/<\w+:pgNumType\b[^>]*\w+:start="(\d+)"/.exec(xml)?.[1] ?? 1) || 1;
  const wraps = wrapDistances(xml);
  // 手动换行结束的两端对齐行：Word 默认也会两端对齐（除非打开了 doNotExpandShiftReturn 兼容选项）
  const expandBreaks = !(settings && /<\w+:doNotExpandShiftReturn(\s|\/|>)/.test(settings) && !/<\w+:doNotExpandShiftReturn\s+\w+:val="(0|false)"/.test(settings));
  try {
    const marked = markKeepFlags(xml, await zip.file("word/styles.xml")?.async("string") ?? "");
    if (marked !== xml) { xml = marked; zip.file("word/document.xml", xml); changed = true; }
  } catch { /* 解析失败：保持原样 */ }
  let anchors: PageAnchor[] = [];
  try {
    const r = liftPageAnchors(xml);
    if (r.anchors.length) { xml = r.xml; anchors = r.anchors; zip.file("word/document.xml", xml); changed = true; }
  } catch { /* 解析失败：保持原样 */ }
  const plain = changed ? await zip.generateAsync({ type: "blob", mimeType: mime }) : blob;
  if (!xml.includes("lastRenderedPageBreak")) return { blob: plain, plain, balance, pageStart, wraps, anchors, expandBreaks };
  zip.file("word/document.xml", xml.replace(/<w:lastRenderedPageBreak\s*\/>/g, ""));
  return { blob: await zip.generateAsync({ type: "blob", mimeType: mime }), plain, balance, pageStart, wraps, anchors, expandBreaks };
}

// ---------------------------------------------------------------------------
// 段前分页 / 与下段同页 / 段中不分页
// ---------------------------------------------------------------------------

/**
 * docx-preview 只认段落样式里的"段前分页"，段落上直接设置的 w:pageBreakBefore 被忽略；
 * "与下段同页"（keepNext）、"段中不分页"（keepLines）完全不支持。Agent 调整分页时改的正是这些属性，
 * 不处理的话快速预览与 Word / 精确版式对不上。
 * 渲染前在这类段落开头放一个书签（名字带上标志 B / N / L），paginate() 据此按 Word 的规则分页。
 * 段落属性优先于样式（含 basedOn 继承链），w:val="0" / "false" 表示关闭。
 */
function markKeepFlags(xml: string, stylesXml: string): string {
  if (!/(pageBreakBefore|keepNext|keepLines)/.test(xml + stylesXml)) return xml;
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const parse = (x: string) => { const d = new DOMParser().parseFromString(x, "application/xml"); return d.getElementsByTagName("parsererror").length ? null : d; };
  const doc = parse(xml);
  if (!doc) return xml;
  const kid = (e: Element | null | undefined, local: string) => (e ? Array.from(e.children).find((c) => c.localName === local) : undefined);
  const on = (e: Element | undefined): boolean | undefined => {
    if (!e) return undefined;
    const v = e.getAttributeNS(W, "val") ?? e.getAttribute("w:val");
    return !(v === "0" || v === "false" || v === "off");
  };
  type Flags = { B?: boolean; N?: boolean; L?: boolean };
  const read = (pPr: Element | undefined): Flags => ({ B: on(kid(pPr, "pageBreakBefore")), N: on(kid(pPr, "keepNext")), L: on(kid(pPr, "keepLines")) });
  // 样式（含继承）
  const styles = new Map<string, { flags: Flags; basedOn?: string }>();
  let defaultPara: string | undefined;
  const sdoc = stylesXml ? parse(stylesXml) : null;
  if (sdoc) {
    for (const st of Array.from(sdoc.getElementsByTagNameNS(W, "style"))) {
      if (st.getAttributeNS(W, "type") !== "paragraph") continue;
      const id = st.getAttributeNS(W, "styleId") ?? "";
      styles.set(id, { flags: read(kid(st, "pPr")), basedOn: kid(st, "basedOn")?.getAttributeNS(W, "val") ?? undefined });
      if (st.getAttributeNS(W, "default") === "1") defaultPara = id;
    }
  }
  const styleFlags = (id: string | undefined): Flags => {
    const out: Flags = {};
    const seen = new Set<string>();
    for (let cur = id ?? defaultPara; cur && !seen.has(cur); cur = styles.get(cur)?.basedOn) {
      seen.add(cur);
      const f = styles.get(cur)?.flags;
      if (!f) break;
      for (const k of ["B", "N", "L"] as const) if (out[k] === undefined && f[k] !== undefined) out[k] = f[k];
    }
    return out;
  };
  let n = 0;
  for (const p of Array.from(doc.getElementsByTagNameNS(W, "p"))) {
    // 表格单元格、文本框里的段落不参与分页
    let anc = p.parentElement, nested = false;
    while (anc) { if (anc.localName === "tc" || anc.localName === "txbxContent") { nested = true; break; } anc = anc.parentElement; }
    if (nested) continue;
    const pPr = kid(p, "pPr");
    const own = read(pPr), sty = styleFlags(kid(pPr, "pStyle")?.getAttributeNS(W, "val") ?? undefined);
    const f = (k: keyof Flags) => (own[k] ?? sty[k]) === true;
    const tag = (f("B") ? "B" : "") + (f("N") ? "N" : "") + (f("L") ? "L" : "");
    if (!tag) continue;
    const bs = doc.createElementNS(W, "w:bookmarkStart");
    bs.setAttributeNS(W, "w:id", String(800000 + n)); bs.setAttributeNS(W, "w:name", `_pvk_${tag}_${n}`);
    const be = doc.createElementNS(W, "w:bookmarkEnd");
    be.setAttributeNS(W, "w:id", String(800000 + n));
    p.insertBefore(be, pPr ? pPr.nextSibling : p.firstChild);
    p.insertBefore(bs, be);
    n++;
  }
  return n ? new XMLSerializer().serializeToString(doc) : xml;
}

/** 渲染后：把书签标志写到所在段落的 data-pvk 上（拆段时随段落一起复制） */
function tagKeepFlags(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>('[id^="_pvk_"]').forEach((m) => {
    const p = m.closest<HTMLElement>("p");
    const tag = /^_pvk_([BNL]+)_/.exec(m.id)?.[1];
    if (p && tag) p.dataset.pvk = tag;
  });
}
const hasKeep = (el: Element | null | undefined, flag: "B" | "N" | "L") => !!el && el instanceof HTMLElement && (el.dataset.pvk ?? "").includes(flag);

// ---------------------------------------------------------------------------
// 按页面坐标定位的图片 / 文本框 / 图形
// ---------------------------------------------------------------------------

/**
 * 不环绕文字、相对页面（或页边距）定位的对象。docx-preview 不支持这种定位（图片画在所在行内、文本框和图形根本不画），
 * 逐页保留原排版的文档里首字下沉、页眉页脚文字、插图、分隔线都是这样放的。
 * 渲染前把它们换成 docx-preview 能渲染的形式（图片 → 行内图片，文本框 → 普通段落），前面放书签作标记；
 * 渲染后由 placePageAnchors() 按原坐标绝对定位到所在页面上。
 */
export interface PageAnchor {
  k: number;
  kind: "pic" | "text" | "shape";
  /** 相对页面 / 页边距的位置与尺寸（pt） */
  x: number; y: number; w: number; h: number;
  relH: "page" | "margin"; relV: "page" | "margin";
  behind: boolean;
  /** 文本框的段落数 */
  paras?: number;
  /** 图形：直线 / 矩形的线色、填充色、线宽 */
  shape?: { prst: string; fill?: string; line?: string; lw: number; flipV: boolean };
}

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

function liftPageAnchors(xml: string): { xml: string; anchors: PageAnchor[] } {
  if (!xml.includes("wp:anchor") && !xml.includes(":anchor")) return { xml, anchors: [] };
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) return { xml, anchors: [] };
  const kids = (e: Element, local: string) => Array.from(e.children).filter((c) => c.localName === local);
  const kid = (e: Element | undefined, local: string) => (e ? kids(e, local)[0] : undefined);
  const desc = (e: Element, local: string) => Array.from(e.getElementsByTagName("*")).filter((c) => c.localName === local);
  const up = (e: Element, local: string) => { let c: Element | null = e.parentElement; while (c && c.localName !== local) c = c.parentElement; return c; };
  const emu = (v: string | null | undefined) => (Number(v) || 0) / 12700;
  const anchors: PageAnchor[] = [];
  let bmId = 900000;
  const bookmark = (name: string) => {
    const s = doc.createElementNS(NS_W, "w:bookmarkStart");
    s.setAttributeNS(NS_W, "w:id", String(++bmId)); s.setAttributeNS(NS_W, "w:name", name);
    const e = doc.createElementNS(NS_W, "w:bookmarkEnd");
    e.setAttributeNS(NS_W, "w:id", String(bmId));
    return [s, e];
  };
  for (const a of Array.from(doc.getElementsByTagNameNS(NS_WP, "anchor"))) {
    if (!kid(a, "wrapNone")) continue; // 环绕文字的浮动对象交给 docx-preview 的浮动排版
    if (up(a, "txbxContent") || up(a, "AlternateContent")) continue;
    const ph = kid(a, "positionH"), pv = kid(a, "positionV");
    const relH = ph?.getAttribute("relativeFrom"), relV = pv?.getAttribute("relativeFrom");
    const offH = kid(ph, "posOffset"), offV = kid(pv, "posOffset");
    if (!offH || !offV || !(relH === "page" || relH === "margin") || !(relV === "page" || relV === "margin")) continue;
    const run = up(a, "r"), para = up(a, "p");
    if (!run || !para || !run.parentNode) continue;
    const ext = kid(a, "extent");
    const item: PageAnchor = {
      k: anchors.length, kind: "pic", x: emu(offH.textContent), y: emu(offV.textContent), w: emu(ext?.getAttribute("cx")), h: emu(ext?.getAttribute("cy")),
      relH, relV, behind: a.getAttribute("behindDoc") === "1",
    };
    const graphic = desc(a, "graphic")[0];
    const gd = graphic ? kid(graphic, "graphicData") : undefined;
    const txbx = desc(a, "txbxContent")[0];
    const drawing = up(a, "drawing");
    if (gd && kid(gd, "pic")) {
      // 图片 → 行内图片（同样的尺寸与图片数据）
      const inline = doc.createElementNS(NS_WP, "wp:inline");
      for (const at of ["distT", "distB", "distL", "distR"]) inline.setAttribute(at, "0");
      for (const local of ["extent", "effectExtent", "docPr", "cNvGraphicFramePr"]) { const c = kid(a, local); if (c) inline.appendChild(c); }
      inline.appendChild(graphic!);
      a.replaceWith(inline);
      const [s, e] = bookmark(`_pva_${item.k}`);
      run.parentNode.insertBefore(s, run); run.parentNode.insertBefore(e, run);
    } else if (txbx) {
      // 文本框 → 紧跟在所在段落之后的普通段落（各带标记），渲染后再移到原位置
      const ps = kids(txbx, "p");
      let at: Node = para;
      ps.forEach((p, i) => {
        const [s, e] = bookmark(`_pvt_${item.k}_${i}`);
        const ppr = kid(p, "pPr");
        p.insertBefore(e, ppr ? ppr.nextSibling : p.firstChild); p.insertBefore(s, e);
        para.parentNode!.insertBefore(p, at.nextSibling); at = p;
      });
      item.kind = "text"; item.paras = ps.length;
      drawing?.remove();
    } else if (gd && desc(gd, "wsp").length) {
      // 直线、矩形等图形
      const geom = desc(gd, "prstGeom")[0]?.getAttribute("prst") ?? "rect";
      const spPr = desc(gd, "spPr")[0];
      const fillClr = spPr && kid(spPr, "solidFill") ? desc(kid(spPr, "solidFill")!, "srgbClr")[0]?.getAttribute("val") ?? undefined : undefined;
      const ln = spPr ? kid(spPr, "ln") : undefined;
      const lineClr = ln && !kid(ln, "noFill") ? desc(ln, "srgbClr")[0]?.getAttribute("val") ?? "000000" : undefined;
      const xfrm = spPr ? kid(spPr, "xfrm") : undefined;
      item.kind = "shape";
      item.shape = { prst: geom, fill: fillClr, line: lineClr, lw: emu(ln?.getAttribute("w")) || 0.75, flipV: xfrm?.getAttribute("flipV") === "1" };
      const [s, e] = bookmark(`_pvs_${item.k}`);
      run.parentNode.insertBefore(s, run); run.parentNode.insertBefore(e, run);
      drawing?.remove();
    } else continue;
    anchors.push(item);
  }
  return { xml: new XMLSerializer().serializeToString(doc), anchors };
}

/** 把 liftPageAnchors() 标记过的对象按原坐标绝对定位到所在页面上（放在页面元素下，不参与正文排版和分页） */
export function placePageAnchors(root: HTMLElement, anchors: PageAnchor[]): void {
  for (const it of anchors) {
    const marker = root.querySelector<HTMLElement>(it.kind === "text" ? `[id="_pvt_${it.k}_0"]` : it.kind === "pic" ? `[id="_pva_${it.k}"]` : `[id="_pvs_${it.k}"]`);
    const sec = marker?.closest<HTMLElement>("section.docx");
    if (!marker || !sec) continue;
    const cs = getComputedStyle(sec);
    const px2pt = 0.75;
    const dx = it.relH === "margin" ? parseFloat(cs.paddingLeft) * px2pt : 0;
    const dy = it.relV === "margin" ? parseFloat(cs.paddingTop) * px2pt : 0;
    const box = document.createElement("div");
    box.className = "pv-abs";
    box.style.cssText = `position:absolute;left:${it.x + dx}pt;top:${it.y + dy}pt;width:${it.w}pt;min-height:${it.h}pt;z-index:${it.behind ? 0 : 2};pointer-events:none;`;
    if (it.kind === "pic") {
      const run = marker.nextElementSibling;
      if (!run) continue;
      box.appendChild(run);
    } else if (it.kind === "text") {
      for (let i = 0; i < (it.paras ?? 1); i++) {
        const p = root.querySelector<HTMLElement>(`[id="_pvt_${it.k}_${i}"]`)?.closest<HTMLElement>("p");
        if (!p) continue;
        p.style.margin = "0";
        box.appendChild(p);
      }
    } else if (it.shape) {
      const sh = it.shape;
      if (sh.prst === "line") {
        // 直线：外框的对角线；水平 / 竖直线最常见
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 ${Math.max(it.w, 0.01)} ${Math.max(it.h, 0.01)}" preserveAspectRatio="none" style="overflow:visible;position:absolute;left:0;top:0">` +
          `<line x1="0" y1="${sh.flipV ? it.h : 0}" x2="${it.w}" y2="${sh.flipV ? 0 : it.h}" stroke="#${sh.line ?? "000000"}" stroke-width="${sh.lw}" vector-effect="non-scaling-stroke"/></svg>`;
        box.innerHTML = svg;
        box.style.height = `${it.h}pt`;
      } else {
        box.style.height = `${it.h}pt`;
        if (sh.fill) box.style.background = `#${sh.fill}`;
        if (sh.line) box.style.border = `${sh.lw}pt solid #${sh.line}`;
        box.style.boxSizing = "border-box";
      }
    }
    sec.appendChild(box);
  }
}

/**
 * 固定行距的段落：Word 中每行高度固定，CSS 却会被撑高——
 *   - 段落自身字号形成的"支柱"与大字号文字按基线对齐时，行框高于行距（标题、大字号的行）；
 *   - 上标 / 下标的升降也会撑高行框。
 * 逐行累积下来，正文越往下越比 PDF 低，按页面坐标放置的首字下沉、插图就和文字对不上。
 * 做法：段落字号设为段内最大的字号（各段文字先固定自己的字号），上下标的行高设为 0。
 */
export function fixExactLineBoxes(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>("p").forEach((p) => {
    const lh = p.style.lineHeight;
    if (!lh || !/(pt|px)$/.test(lh) || p.style.minHeight !== lh) return; // docx-preview 对固定行距同时设置 line-height 与 min-height
    let max = 0;
    const texts = Array.from(p.querySelectorAll<HTMLElement>("span")).filter((sp) => Array.from(sp.childNodes).some((n) => n.nodeType === 3 && n.textContent!.trim()));
    for (const sp of texts) {
      const cs = getComputedStyle(sp);
      const fs = parseFloat(cs.fontSize);
      sp.style.fontSize = cs.fontSize;
      const va = cs.verticalAlign;
      if (va !== "baseline" && !/^-?0(px)?$/.test(va)) sp.style.lineHeight = "0";
      else if (fs > max) max = fs;
    }
    p.querySelectorAll<HTMLElement>("sup, sub").forEach((x) => { x.style.lineHeight = "0"; });
    if (max > 0) p.style.fontSize = `${max}px`;
  });
}

/**
 * 以手动换行结束的两端对齐行：浏览器不对齐（text-align: justify 不作用于强制换行之前的行），Word 默认会对齐。
 * 逐页保留原排版的文档每一行都以换行结束，不处理的话整页行尾参差不齐。
 * 做法：text-align-last: justify 让这些行对齐；段落真正的最后一行包进一个行内块，保持左对齐。
 */
export function justifyForcedBreaks(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>("p").forEach((p) => {
    const brs = p.querySelectorAll("br");
    if (!brs.length || getComputedStyle(p).textAlign !== "justify") return;
    p.style.textAlignLast = "justify";
    let top: Node = brs[brs.length - 1];
    while (top.parentNode && top.parentNode !== p) top = top.parentNode;
    const tail: Node[] = [];
    for (let n = top.nextSibling; n; n = n.nextSibling) tail.push(n);
    if (!tail.some((n) => (n.textContent ?? "").trim() || (n instanceof Element && n.querySelector("img")))) return;
    const box = document.createElement("span");
    box.style.cssText = "display:inline-block;text-align:left;text-align-last:auto;text-indent:0;";
    p.insertBefore(box, tail[0]);
    for (const n of tail) box.appendChild(n);
  });
}

/**
 * "最小值"行距（w:lineRule="atLeast"）：docx-preview 把它渲染成 line-height: calc(100% + N pt)，
 * 即字号再加 N pt——约为 Word 里实际行距的 1.9 倍，和固定行距 / 单倍行距的段落排在一起时行距忽松忽紧。
 * Word 的语义是"至少 N pt，字体自然行高更大时取自然行高"，这里改写成 max(N pt, 1.15em)。
 * 同时处理样式表（<style>）和段落上的内联样式。
 */
const AT_LEAST = /calc\(\s*100%\s*\+\s*(-?[\d.]+)(pt|px)\s*\)/g; // 内联样式经浏览器规范化后单位是 px
const atLeastValue = (_m: string, n: string, unit: string) => `max(${n}${unit}, 1.15em)`;
export function fixAtLeastSpacing(root: HTMLElement): void {
  root.querySelectorAll("style").forEach((st) => {
    const t = st.textContent ?? "";
    if (t.includes("calc(")) st.textContent = t.replace(/line-height\s*:\s*calc\([^;}]*\)/g, (m) => m.replace(AT_LEAST, atLeastValue));
  });
  root.querySelectorAll<HTMLElement>('[style*="calc("]').forEach((el) => {
    const v = el.style.lineHeight;
    if (v && v.startsWith("calc(")) el.style.lineHeight = v.replace(AT_LEAST, atLeastValue);
  });
}

/** 文字环绕图片与文字的间距（pt） */
export interface WrapDist { t: number; b: number; l: number; r: number }

/**
 * 文字环绕（四周型 / 紧密型 / 穿越型）图片的环绕间距，按在正文中出现的顺序。
 * docx-preview 把这类图片渲染成 float: left，却忽略 distT/distB/distL/distR，文字会紧贴图片边缘。
 */
function wrapDistances(xml: string): WrapDist[] {
  const out: WrapDist[] = [];
  const emuToPt = (v: string | undefined) => (v ? Number(v) / 12700 : 0) || 0;
  for (const m of xml.matchAll(/<wp:anchor\b([^>]*)>([\s\S]*?)<\/wp:anchor>/g)) {
    if (!/<wp:wrap(Square|Tight|Through)\b/.test(m[2])) continue; // 与 docx-preview 浮动的条件一致：不是 wrapNone / wrapTopAndBottom
    const a = (n: string) => new RegExp(`\\b${n}="(\\d+)"`).exec(m[1])?.[1];
    out.push({ t: emuToPt(a("distT")), b: emuToPt(a("distB")), l: emuToPt(a("distL")), r: emuToPt(a("distR")) });
  }
  return out;
}

/**
 * 把环绕间距加到 docx-preview 渲染出的浮动图片上（按出现顺序一一对应；数量对不上就不处理，避免错位）。
 * 图片靠左浮动：文字在它右边和下边，留右边距、下边距；左边距不加，以免图片偏离栏边界。
 */
export function applyWrapDistances(root: HTMLElement, wraps: WrapDist[]): void {
  if (!wraps.length) return;
  const floats = Array.from(root.querySelectorAll<HTMLElement>("div")).filter((d) => {
    const f = d.style.float;
    return (f === "left" || f === "right") && d.style.display === "inline-block" && !d.closest("header, footer") && !!d.querySelector("img, svg");
  });
  if (floats.length !== wraps.length) return;
  // 以前转换出的文档只写了 Word 默认的 2pt，文字贴着图片；这种情况下预览至少留 6pt
  const side = (v: number) => (v <= 2.5 ? 6 : v);
  floats.forEach((d, i) => {
    const w = wraps[i];
    if (d.style.float === "right") d.style.margin = `${w.t}pt 0 ${w.b}pt ${side(w.l)}pt`;
    else d.style.margin = `${w.t}pt ${side(w.r)}pt ${w.b}pt 0`;
  });
}

/** 填入页眉页脚中的页码，并在每页下方加"第 n 页 / 共 N 页"标注；返回总页数 */
export function numberPages(root: HTMLElement, start = 1): number {
  const wrapper = root.querySelector<HTMLElement>(".docx-wrapper") ?? root;
  wrapper.querySelectorAll(".page-label").forEach((x) => x.remove());
  const pages = Array.from(wrapper.children).filter((x): x is HTMLElement => x instanceof HTMLElement && x.tagName === "SECTION");
  pages.forEach((pg, i) => {
    const walker = document.createTreeWalker(pg, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const t = n.textContent ?? "";
      if (t.includes("\uE000")) n.textContent = t.split(PH_PAGE).join(String(start + i)).split(PH_NUMPAGES).join(String(pages.length));
    }
    pg.dataset.page = String(i + 1);
    const label = document.createElement("div");
    label.className = "page-label";
    label.textContent = `第 ${i + 1} 页 / 共 ${pages.length} 页`;
    pg.after(label);
  });
  return pages.length;
}

function pxHeight(el: HTMLElement): number {
  const v = parseFloat(getComputedStyle(el).minHeight);
  return Number.isFinite(v) && v > 0 ? v : el.offsetHeight;
}

/** 文本偏移 → (文本节点, 节点内偏移) */
function locate(root: Node, k: number): { node: Node; offset: number } | null {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let pos = 0;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const len = n.textContent?.length ?? 0;
    if (pos + len >= k) return { node: n, offset: k - pos };
    pos += len;
  }
  return null;
}

/** 复制段落，只保留前 k 个字符（head）或第 k 个字符之后的部分（tail） */
function cut(el: HTMLElement, k: number, part: "head" | "tail"): HTMLElement {
  const cl = el.cloneNode(true) as HTMLElement;
  const at = locate(cl, k);
  if (!at) return cl;
  const r = document.createRange();
  if (part === "head") { r.setStart(at.node, at.offset); r.setEnd(cl, cl.childNodes.length); }
  else { r.setStart(cl, 0); r.setEnd(at.node, at.offset); }
  r.deleteContents();
  return cl;
}

const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]/;

export function paginate(content: HTMLElement, hf: HTMLElement | null, opts: { balance?: boolean } = {}): number {
  const balance = opts.balance !== false;
  const wrapper = content.querySelector<HTMLElement>(".docx-wrapper") ?? content;
  const groups = Array.from(wrapper.children).filter((x): x is HTMLElement => x instanceof HTMLElement && x.tagName === "SECTION");
  const hfPages = hf ? Array.from(hf.querySelectorAll<HTMLElement>("section")) : [];
  const pickHF = (i: number): HTMLElement | null => {
    if (!hfPages.length) return null;
    if (i < hfPages.length) return hfPages[i];
    const same = hfPages.filter((_, j) => j >= 1 && j % 2 === i % 2);
    return same[same.length - 1] ?? hfPages[hfPages.length - 1];
  };
  let pageNo = 0, made = 0;
  tagKeepFlags(wrapper);

  for (const src of groups) {
    const pageH = pxHeight(src);
    // 本来就是一页的内容不动。文档自己逐页分页时（逐页保留原排版的转换结果、手动分页的文档），
    // 每一段在浏览器里可能因字宽差异略长于一页，只有明显超过两页的才重新分页
    const limit = groups.length > 1 ? 2 : 1.02;
    // 节中间有段前分页的段落（docx-preview 不认段落上直接设置的段前分页）时也要重新分页
    const forced = Array.from(src.querySelectorAll<HTMLElement>("p[data-pvk*='B']")).some((p) => p !== src.querySelector("p"));
    if (src.scrollHeight <= pageH * limit + 2 && !forced) { pageNo++; continue; }

    let page!: HTMLElement, body!: HTMLElement;
    const remaining = (after: HTMLElement | null) =>
      body.getBoundingClientRect().bottom - (after ? after.getBoundingClientRect().bottom : body.getBoundingClientRect().top);
    const newPage = () => {
      if (made++ > MAX_PAGES) throw new Error("分页过多");
      page = src.cloneNode(false) as HTMLElement;
      page.style.height = `${pageH}px`;
      const ref = pickHF(pageNo) ?? src;
      const header = ref.querySelector<HTMLElement>(":scope > header") ?? src.querySelector<HTMLElement>(":scope > header");
      const footer = ref.querySelector<HTMLElement>(":scope > footer") ?? src.querySelector<HTMLElement>(":scope > footer");
      body = document.createElement("div");
      body.className = "pg-body";
      body.style.cssText = "flex:1 1 auto;min-height:0;overflow:hidden;position:relative;z-index:1;";
      if (header) page.appendChild(header.cloneNode(true));
      page.appendChild(body);
      if (footer) page.appendChild(footer.cloneNode(true));
      wrapper.insertBefore(page, src);
      pageNo++;
    };
    newPage();

    // 放不放得下：看最后放入的元素——多栏时它的任何一段落进了第 cols+1 栏（容器右侧之外）即溢出；
    // 单栏时看它的底边。不用 scrollWidth：略宽于栏的公式图片会让 scrollWidth 变大，却并没有溢出到下一栏
    const fits = (c: HTMLElement, cols: number) => {
      const el = c.lastElementChild as HTMLElement | null;
      if (!el) return true;
      const box = c.getBoundingClientRect();
      const rects = Array.from(el.getClientRects());
      if (!rects.length) return true;
      if (cols >= 2) {
        const colW = box.width / cols;
        return rects.every((r) => r.left < box.right - 0.5 * colW && r.bottom <= box.bottom + 1);
      }
      return rects.every((r) => r.bottom <= box.bottom + 1);
    };

    // 按行拆段落：二分查找放得下的最长前缀（只在词边界处断开）
    const split = (el: HTMLElement, c: HTMLElement, cols: number): [HTMLElement, HTMLElement] | null => {
      if (el.tagName !== "P" || hasKeep(el, "L")) return null; // 段中不分页
      const text = el.textContent ?? "";
      if (text.trim().length < 2) return null;
      const cuts: number[] = [];
      for (let i = 1; i < text.length; i++) {
        const a = text[i - 1];
        if (/\s/.test(a) || a === "-" || a === "‐" || CJK.test(a) || CJK.test(text[i])) cuts.push(i);
      }
      let lo = 0, hi = cuts.length - 1, best = -1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const head = cut(el, cuts[mid], "head");
        c.appendChild(head);
        const ok = fits(c, cols);
        c.removeChild(head);
        if (ok) { best = mid; lo = mid + 1; } else hi = mid - 1;
      }
      if (best < 0) return null;
      const k = cuts[best];
      const head = cut(el, k, "head"), tail = cut(el, k, "tail");
      if (getComputedStyle(el).textAlign === "justify") head.style.textAlignLast = "justify";
      head.style.marginBottom = "0";
      tail.style.textIndent = "0";
      tail.style.marginTop = "0";
      // 段前分页只作用于段落的第一部分；前半段的"与下段同页"由后半段自然满足
      if (tail.dataset.pvk) tail.dataset.pvk = tail.dataset.pvk.replace("B", "");
      head.dataset.pvHead = "1";
      return [head, tail];
    };

    // 按页面坐标定位的对象（placePageAnchors）不参与排版，放到这一节分出的第一页上
    const absolutes = Array.from(src.children).filter((x): x is HTMLElement => x instanceof HTMLElement && x.classList.contains("pv-abs"));
    const firstPage = page;
    const children = Array.from(src.children).filter((x): x is HTMLElement => x instanceof HTMLElement && x.tagName !== "HEADER" && x.tagName !== "FOOTER" && !x.classList.contains("pv-abs"));
    let last: HTMLElement | null = null; // 本页最后一个容器
    let lastCols = 0;
    /** 推迟到下一页顶部的通栏图表（与 LaTeX 的页顶浮动体一致：文字继续排满本页，不留空白） */
    const deferred: HTMLElement[] = [];
    const isFloatBand = (a: HTMLElement) => {
      if (a.tagName !== "ARTICLE" || (parseInt(a.style.columnCount || "1", 10) || 1) > 1) return false;
      if (!a.querySelector("img,svg,table")) return false;
      // 题注不长、没有成段的正文（表格内的文字不算）
      return Array.from(a.children).every((k) => k.tagName === "TABLE" || (k.textContent ?? "").length < 700);
    };
    const flushDeferred = () => {
      while (deferred.length) placeArticle(deferred.shift()!, null);
    };

    const placeArticle = (child: HTMLElement, reuse: HTMLElement | null) => {
      const isArticle = child.tagName === "ARTICLE";
      const cols = isArticle ? parseInt(child.style.columnCount || "1", 10) || 1 : 1;
      const open = () => {
        const c = (isArticle ? child.cloneNode(false) : document.createElement("div")) as HTMLElement;
        c.style.columnFill = "auto";
        c.style.height = `${Math.max(1, remaining(last))}px`;
        c.style.overflow = "hidden";
        c.style.marginBottom = "0";
        body.appendChild(c);
        return c;
      };
      // 节结束（或换页前的最后一段）：多栏平衡、高度收紧到内容
      const close = (c: HTMLElement, full: boolean) => {
        if (!full) {
          if (balance || cols < 2) {
            c.style.height = "auto";
            c.style.columnFill = "balance";
          } else {
            // 不平衡：保持整栏高度、先排满左栏再排右栏；内容全在左栏时，高度收紧到左栏内容（后面的节紧接其下）
            const box = c.getBoundingClientRect();
            const colW = box.width / cols;
            let maxBottom = box.top, inLeftOnly = true;
            for (const el of Array.from(c.children)) {
              for (const r of Array.from(el.getClientRects())) {
                if (r.left >= box.left + colW * 0.75) inLeftOnly = false;
                maxBottom = Math.max(maxBottom, r.bottom);
              }
            }
            if (inLeftOnly) c.style.height = `${Math.max(1, maxBottom - box.top)}px`;
          }
        }
        last = c;
        lastCols = cols;
      };
      let c: HTMLElement;
      if (reuse) {
        // 接着上一段同栏数的正文排（中间的通栏图表推迟到了下一页）
        reuse.style.columnFill = "auto";
        reuse.style.height = `${Math.max(1, body.getBoundingClientRect().bottom - reuse.getBoundingClientRect().top)}px`;
        c = reuse;
      } else c = open();
      const queue: HTMLElement[] = isArticle ? Array.from(child.children) as HTMLElement[] : [child];
      while (queue.length) {
        const item = queue.shift()!;
        // 段前分页：本页已有内容时从新页开始（只处理一次）
        const pageHasContent = !!c.childElementCount || !!last || !!body.querySelector("p,table,img,svg");
        if (hasKeep(item, "B") && !item.dataset.pvBroke && pageHasContent) {
          item.dataset.pvBroke = "1";
          queue.unshift(item);
          close(c, true);
          newPage();
          last = null;
          lastCols = 0;
          flushDeferred();
          c = open();
          continue;
        }
        c.appendChild(item);
        if (fits(c, cols)) continue;
        c.removeChild(item);
        const parts = split(item, c, cols);
        if (parts) {
          c.appendChild(parts[0]);
          queue.unshift(parts[1]);
        } else if (c.childElementCount || last) {
          // 整段移到下一页：前面紧挨着的"与下段同页"段落随它一起走（整页都是这样的段落时不移，与 Word 一致）
          const moved: HTMLElement[] = [];
          while (c.lastElementChild instanceof HTMLElement && hasKeep(c.lastElementChild, "N") && !c.lastElementChild.dataset.pvHead) {
            const el = c.lastElementChild;
            if (c.childElementCount === 1 && !last) break;
            moved.unshift(el);
            c.removeChild(el);
          }
          queue.unshift(...moved, item);
        } else if (!c.childElementCount && !last) {
          // 一整页都放不下（大表格、大图）：单独占一页，页面随之变长
          c.appendChild(item);
          c.style.height = "auto";
          c.style.overflow = "visible";
          page.style.height = "";
          body.style.overflow = "visible";
        }
        close(c, true);
        newPage();
        last = null;
        lastCols = 0;
        flushDeferred();
        c = open();
      }
      close(c, false);
    };

    for (const child of children) {
      const cols = child.tagName === "ARTICLE" ? parseInt(child.style.columnCount || "1", 10) || 1 : 1;
      // 通栏图表：页面上已经有内容时推迟到下一页顶部
      if (isFloatBand(child) && last) { deferred.push(child); continue; }
      const reuse = deferred.length && cols > 1 && last && lastCols === cols && body.contains(last) ? last : null;
      placeArticle(child, reuse);
    }
    // 文末还有推迟的图表：放得下就放在本页，否则下一页
    flushDeferred();
    for (const a of absolutes) firstPage.appendChild(a);
    // 空的末页（恰好排满时）去掉
    if (body && !body.textContent?.trim() && !body.querySelector("img,svg,table") && page !== firstPage) page.remove();
    src.remove();
  }
  return pageNo;
}
