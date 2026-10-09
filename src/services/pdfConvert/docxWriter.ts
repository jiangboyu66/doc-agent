/**
 * 文档模型 → .docx（从零生成 OOXML）
 *
 * 所有属性按 OOXML Schema 顺序输出；关闭孤行控制、中西文自动间距、网格对齐等会改变分页和字距的选项，
 * 使 Word 的排版尽量逐行对应原 PDF。
 */

import { createZip } from "../../documents/docx/zip.js";
import type { Block, Border, Cell, DocModel, HeaderSet, Para, Run, SectionProps, Table } from "./layout.js";
import type { ImageBox } from "./extract.js";

const tw = (pt: number) => Math.round(pt * 20); // twips
const emu = (pt: number) => Math.round(pt * 12700);
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
/** 属性值：去掉 XML 不允许的控制字符与私有区字形码，换行保留为 &#10; */
const escAttr = (s: string) => esc(s.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff\ue000-\uf8ff]/g, "").slice(0, 4000)).replace(/\r?\n/g, "&#10;");
// XML 1.0 不允许的控制字符
const clean = (s: string) => s.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, "");

const NS =
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14"';

class Writer {
  private media: Array<{ name: string; data: Buffer; rid: string; part: string }> = [];
  /** 当前正在生成的部件（图片关系写进该部件的 .rels） */
  private relPart = "document.xml";
  private drawingId = 1;
  private paraSeq = 0x10000000;
  private sectCount = 0;
  /** 页眉页脚部件：在第一节的 sectPr 中引用 */
  private hfParts: Array<{ kind: "header" | "footer"; type: "default" | "first" | "even"; name: string; rid: string; paras: Para[] }> = [];

  constructor(private m: DocModel) {
    let k = 0;
    const add = (kind: "header" | "footer", set: HeaderSet | undefined) => {
      if (!set) return;
      for (const type of ["default", "first", "even"] as const) {
        const paras = set[type];
        if (!paras) continue;
        k++;
        this.hfParts.push({ kind, type, name: `${kind}${k}.xml`, rid: `rIdHf${k}`, paras });
      }
    };
    add("header", m.headers);
    add("footer", m.footers);
  }

  private rPr(r: Run): string {
    const parts: string[] = [];
    const f = esc(r.font);
    parts.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`);
    if (r.bold) parts.push("<w:b/><w:bCs/>");
    if (r.italic) parts.push("<w:i/><w:iCs/>");
    if (r.strike) parts.push("<w:strike/>");
    if (r.color && r.color !== "000000") parts.push(`<w:color w:val="${r.color}"/>`);
    if (r.spacing) parts.push(`<w:spacing w:val="${Math.round(r.spacing * 20)}"/>`);
    const hp = Math.max(2, Math.round(r.size * 2));
    parts.push(`<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`);
    if (r.underline) parts.push('<w:u w:val="single"/>');
    if (r.vertAlign) parts.push(`<w:vertAlign w:val="${r.vertAlign}"/>`);
    return `<w:rPr>${parts.join("")}</w:rPr>`;
  }

  private run(r: Run): string {
    if (r.tab) return `<w:r>${this.rPr(r)}<w:tab/></w:r>`;
    // 逐行一致模式的排版换行：clear="none"（与普通换行等效）作为标记，编辑时可识别并去掉
    if (r.br) return `<w:r>${this.rPr(r)}<w:br w:clear="none"/></w:r>`;
    if (r.pageMark) return "<w:r><w:lastRenderedPageBreak/></w:r>";
    if (r.field === "PAGE") return `<w:fldSimple w:instr=" PAGE "><w:r>${this.rPr(r)}<w:t>${esc(r.text)}</w:t></w:r></w:fldSimple>`;
    return `<w:r>${this.rPr(r)}<w:t xml:space="preserve">${esc(clean(r.text))}</w:t></w:r>`;
  }

  private border(tag: string, b: Border | null | undefined, space = 0): string {
    if (!b) return `<w:${tag} w:val="nil"/>`;
    const sz = Math.max(2, Math.min(96, Math.round(b.width * 8)));
    return `<w:${tag} w:val="single" w:sz="${sz}" w:space="${space}" w:color="${b.color}"/>`;
  }

  /** 图片的公共部分：docPr + graphic；返回 [id, xml] */
  private picture(img: ImageBox): { id: number; cx: number; cy: number; xml: string } {
    const rid = `rIdImg${this.media.length + 1}`;
    const name = `image${this.media.length + 1}.png`;
    this.media.push({ name, data: img.png, rid, part: this.relPart });
    const id = this.drawingId++;
    const cx = emu(img.x1 - img.x0), cy = emu(img.y1 - img.y0);
    const xml =
      `<wp:docPr id="${id}" name="图片 ${id}"${img.alt ? ` descr="${escAttr(img.alt)}"` : ""}/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic>`;
    return { id, cx, cy, xml };
  }

  /** 以页面绝对坐标浮动定位的图片 */
  private image(img: ImageBox, textOverlaps: boolean): string {
    const { id, cx, cy, xml } = this.picture(img);
    return (
      `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${251658240 + id}" behindDoc="${textOverlaps ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(img.x0)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(img.y0)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
      `${xml}</wp:anchor></w:drawing></w:r>`
    );
  }

  /** 行内图片（随文字流移动） */
  private inlineImage(img: ImageBox): string {
    const { cx, cy, xml } = this.picture(img);
    return (
      `<w:r><w:rPr><w:noProof/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${xml}</wp:inline></w:drawing></w:r>`
    );
  }

  /** 页眉中的图片：相对页边距 / 段落顶部定位，浮于文字上方 */
  private hfImage(f: NonNullable<Para["hfImages"]>[number]): string {
    const { id, cx, cy, xml } = this.picture(f.img);
    return (
      `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${251658240 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>${emu(f.dx)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(f.dy)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
      `${xml}</wp:anchor></w:drawing></w:r>`
    );
  }

  /** 随段落浮动、四周型文字环绕的图片：相对栏左边界 / 段落顶部定位 */
  private floatImage(f: NonNullable<Para["floats"]>[number]): string {
    const { id, cx, cy, xml } = this.picture(f.img);
    return (
      `<w:r><w:drawing><wp:anchor distT="0" distB="${emu(f.img.wrap?.b ?? 2)}" distL="${emu(f.img.wrap?.l ?? 2)}" distR="${emu(f.img.wrap?.r ?? 2)}" simplePos="0" relativeHeight="${251658240 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${emu(f.dx)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(f.dy)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>` +
      `${xml}</wp:anchor></w:drawing></w:r>`
    );
  }

  /** 绝对定位的线条 / 色块 */
  private shape(sh: NonNullable<Para["shapes"]>[number]): string {
    const id = this.drawingId++;
    const w = Math.max(0, sh.x1 - sh.x0), h = Math.max(0, sh.y1 - sh.y0);
    // 线条的坐标是中心线；色块按外框
    const x = sh.kind === "line" ? sh.x0 : sh.x0, y = sh.kind === "line" ? sh.y0 : sh.y0;
    const cx = emu(w), cy = emu(h);
    const geom = sh.kind === "line"
      ? `<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="${emu(Math.max(0.25, sh.width))}"><a:solidFill><a:srgbClr val="${sh.color}"/></a:solidFill></a:ln>`
      : `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${sh.color}"/></a:solidFill><a:ln><a:noFill/></a:ln>`;
    return (
      `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${(sh.behind ? 251650000 : 251658240) + id}" behindDoc="${sh.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
      `<wp:docPr id="${id}" name="${sh.kind === "line" ? "直线" : "矩形"} ${id}"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>` +
      `<wps:spPr><a:xfrm${sh.kind === "line" && h > w ? "" : ""}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>${geom}</wps:spPr>` +
      `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
    );
  }

  /** 绝对定位的文本框（页眉页脚等）：无边框、无填充、零内边距、不参与文字环绕 */
  private textBox(tb: { x: number; y: number; w: number; h: number; para: Para }): string {
    const id = this.drawingId++;
    const cx = emu(tb.w), cy = emu(tb.h);
    return (
      `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${251658240 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(tb.x)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(tb.y)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
      `<wp:docPr id="${id}" name="文本框 ${id}"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
      `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
      `<wps:txbx><w:txbxContent>${this.para(tb.para)}</w:txbxContent></wps:txbx>` +
      `<wps:bodyPr rot="0" vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>` +
      `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
    );
  }

  private sectPr(s: SectionProps, continuous: boolean): string {
    const m = this.m;
    // 每一节都显式引用页眉页脚（与继承等效；浏览器预览只认本节的引用）；起始页码只在第一节；
    // 从第一页开始的节标记"首页不同"
    const first = this.sectCount++ === 0;
    const refs = this.hfParts.length ? this.hfParts.map((h) => `<w:${h.kind}Reference w:type="${h.type}" r:id="${h.rid}"/>`).sort((a, b) => (a.startsWith("<w:header") ? 0 : 1) - (b.startsWith("<w:header") ? 0 : 1)).join("") : "";
    const pgNum = first && m.pageNumStart !== undefined ? `<w:pgNumType w:start="${m.pageNumStart}"/>` : "";
    const titlePg = s.titlePg && this.hfParts.some((h) => h.type === "first") ? "<w:titlePg/>" : "";
    const cols = s.cols > 1
      ? `<w:cols w:num="${s.cols}" w:space="${tw(s.colGap)}" w:equalWidth="0">${s.colWidths.map((w, i) => `<w:col w:w="${tw(w)}"${i < s.colWidths.length - 1 ? ` w:space="${tw(s.colGap)}"` : ""}/>`).join("")}</w:cols>`
      : '<w:cols w:space="425"/>';
    return (
      `<w:sectPr>${refs}${continuous ? '<w:type w:val="continuous"/>' : ""}` +
      `<w:pgSz w:w="${tw(m.pageW)}" w:h="${tw(m.pageH)}"${m.pageW > m.pageH ? ' w:orient="landscape"' : ""}/>` +
      `<w:pgMar w:top="${tw(m.margins.top)}" w:right="${tw(m.margins.right)}" w:bottom="${tw(m.margins.bottom)}" w:left="${tw(m.margins.left)}" w:header="${tw(m.headerDist ?? 0)}" w:footer="${tw(m.footerDist ?? 0)}" w:gutter="0"/>` +
      `${pgNum}${cols}${titlePg}</w:sectPr>`
    );
  }

  para(p: Para): string {
    const ppr: string[] = [];
    // 分页用样式表达（"页首"样式自带段前分页）：Word、LibreOffice 与浏览器预览都能正确识别
    if (p.heading) ppr.push(`<w:pStyle w:val="Heading${p.heading}"/>`);
    else if (p.pageBreakBefore) ppr.push(`<w:pStyle w:val="PageStart"/>`);
    if (p.keepNext) ppr.push("<w:keepNext/>");
    if (p.pageBreakBefore && p.heading) ppr.push("<w:pageBreakBefore/>");

    ppr.push('<w:widowControl w:val="0"/>');
    if (p.borderBottom) ppr.push(`<w:pBdr>${this.border("bottom", p.borderBottom, 0)}</w:pBdr>`);
    if (p.tabs.length) ppr.push(`<w:tabs>${p.tabs.map((t) => `<w:tab w:val="${t.align}" w:pos="${tw(t.pos)}"/>`).join("")}</w:tabs>`);
    ppr.push('<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/>');
    // 图片段落：单倍行距（行高自动等于图片高度；不用"最小值"——浏览器预览会把它当成行高，图片上方多出一大块空白），其余为固定行距
    ppr.push(p.inlineImages?.length
      ? `<w:spacing w:before="${tw(p.spaceBefore)}" w:after="0" w:line="240" w:lineRule="auto"/>`
      : `<w:spacing w:before="${tw(p.spaceBefore)}" w:after="0" w:line="${Math.max(20, tw(p.lineHeight))}" w:lineRule="exact"/>`);
    const ind: string[] = [];
    if (p.indLeft) ind.push(`w:left="${tw(p.indLeft)}"`);
    if (p.indRight) ind.push(`w:right="${tw(p.indRight)}"`);
    if (p.firstLine > 0.5) ind.push(`w:firstLine="${tw(p.firstLine)}"`);
    else if (p.firstLine < -0.5) ind.push(`w:hanging="${tw(-p.firstLine)}"`);
    if (ind.length) ppr.push(`<w:ind ${ind.join(" ")}/>`);
    if (p.align !== "left") ppr.push(`<w:jc w:val="${p.align}"/>`);
    if (p.tiny) ppr.push('<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>');
    if (p.sectionEnd) ppr.push(this.sectPr(p.sectionEnd, true));

    const body: string[] = [];
    if (p.columnBreakBefore) body.push('<w:r><w:br w:type="column"/></w:r>');
    if (p.pageMark) body.push("<w:r><w:lastRenderedPageBreak/></w:r>");
    for (const img of p.images ?? []) if (img.png.length) body.push(this.image(img, !!img.behind));
    for (const f of p.floats ?? []) if (f.img.png.length) body.push(this.floatImage(f));
    for (const f of p.hfImages ?? []) if (f.img.png.length) body.push(this.hfImage(f));
    for (const img of p.inlineImages ?? []) {
      if (!img.png.length) continue;
      // 并排的多张图：每张前面一个制表符，定位到原横坐标
      if ((p.inlineImages?.length ?? 0) > 1) body.push('<w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:tab/></w:r>');
      body.push(this.inlineImage(img));
    }
    for (const sh of p.shapes ?? []) body.push(this.shape(sh));
    for (const tb of p.textBoxes ?? []) body.push(this.textBox(tb));
    for (const r of p.runs) body.push(this.run(r));
    const pid = (this.paraSeq++).toString(16).toUpperCase().padStart(8, "0");
    return `<w:p w14:paraId="${pid}" w14:textId="77777777"><w:pPr>${ppr.join("")}</w:pPr>${body.join("")}</w:p>`;
  }

  private cell(c: Cell): string {
    const pr: string[] = [`<w:tcW w:w="${tw(c.width)}" w:type="dxa"/>`];
    if (c.gridSpan > 1) pr.push(`<w:gridSpan w:val="${c.gridSpan}"/>`);
    if (c.vMerge === "restart") pr.push('<w:vMerge w:val="restart"/>');
    else if (c.vMerge === "continue") pr.push("<w:vMerge/>");
    pr.push(`<w:tcBorders>${this.border("top", c.borders.top)}${this.border("left", c.borders.left)}${this.border("bottom", c.borders.bottom)}${this.border("right", c.borders.right)}</w:tcBorders>`);
    if (c.shade) pr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${c.shade}"/>`);
    if (c.vAlign !== "top") pr.push(`<w:vAlign w:val="${c.vAlign}"/>`);
    const content = c.blocks.length ? c.blocks.map((p) => this.para(p)).join("") : this.para({ kind: "p", runs: [], align: "left", indLeft: 0, indRight: 0, firstLine: 0, spaceBefore: 0, lineHeight: 12, tabs: [], size: 10 });
    return `<w:tc><w:tcPr>${pr.join("")}</w:tcPr>${content}</w:tc>`;
  }

  table(t: Table): string {
    const total = t.colWidths.reduce((a, b) => a + b, 0);
    const tblPr =
      `<w:tblPr>` +
      (t.float
        ? `<w:tblpPr w:leftFromText="0" w:rightFromText="0" w:topFromText="0" w:bottomFromText="0" w:vertAnchor="page" w:horzAnchor="page" w:tblpX="${tw(t.float.x)}" w:tblpY="${tw(t.float.y)}"/><w:tblOverlap w:val="overlap"/><w:tblW w:w="${tw(total)}" w:type="dxa"/>`
        : `<w:tblW w:w="${tw(total)}" w:type="dxa"/><w:tblInd w:w="${tw(t.indent)}" w:type="dxa"/>`) +
      `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="57" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="57" w:type="dxa"/></w:tblCellMar>` +
      `<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>`;
    const grid = `<w:tblGrid>${t.colWidths.map((w) => `<w:gridCol w:w="${tw(w)}"/>`).join("")}</w:tblGrid>`;
    const rows = t.rows.map((r) => `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${tw(r.height)}" w:hRule="${t.exactRows ? "exact" : "atLeast"}"/></w:trPr>${r.cells.map((c) => this.cell(c)).join("")}</w:tr>`).join("");
    return `<w:tbl>${tblPr}${grid}${rows}</w:tbl>`;
  }

  document(): string {
    const out: string[] = [];
    const tiny: Para = { kind: "p", runs: [], align: "left", indLeft: 0, indRight: 0, firstLine: 0, spaceBefore: 0, lineHeight: 1, tabs: [], tiny: true, size: 1 };
    this.m.blocks.forEach((b, i) => {
      // 相邻的两个表格之间必须有段落，否则 Word 会把它们合并成一个表格
      if (b.kind === "table" && this.m.blocks[i - 1]?.kind === "table") out.push(this.para(tiny));
      out.push(b.kind === "p" ? this.para(b) : this.table(b));
    });
    // 表格不能作为 body 的最后一个元素之前的节尾；Word 要求 body 以段落或 sectPr 结束
    if (this.m.blocks[this.m.blocks.length - 1]?.kind === "table") out.push(this.para({ kind: "p", runs: [], align: "left", indLeft: 0, indRight: 0, firstLine: 0, spaceBefore: 0, lineHeight: 1, tabs: [], tiny: true, size: 1 }));
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NS}><w:body>${out.join("")}${this.sectPr(this.m.finalSection, this.m.blocks.some((b) => b.kind === "p" && !!b.sectionEnd))}</w:body></w:document>`;
  }

  styles(): string {
    const f = esc(this.m.bodyFont);
    const hp = Math.round(this.m.bodySize * 2);
    const heading = (n: number) =>
      `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
      `<w:pPr><w:keepNext/><w:outlineLvl w:val="${n - 1}"/></w:pPr></w:style>`;
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/><w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
      heading(1) + heading(2) + heading(3) +
      `<w:style w:type="paragraph" w:customStyle="1" w:styleId="PageStart"><w:name w:val="Page Start"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:semiHidden/>` +
      `<w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:style>` +
      `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>` +
      `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>` +
      `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
      `</w:styles>`
    );
  }

  /** 页眉 / 页脚部件 */
  private hfXml(h: { kind: "header" | "footer"; paras: Para[]; name: string }): string {
    const tag = h.kind === "header" ? "hdr" : "ftr";
    const empty: Para = { kind: "p", runs: [], align: "left", indLeft: 0, indRight: 0, firstLine: 0, spaceBefore: 0, lineHeight: 1, tabs: [], tiny: true, size: 1 };
    this.relPart = (h as any).name ?? "document.xml";
    const body = (h.paras.length ? h.paras : [empty]).map((p) => this.para({ ...p, inlineImages: undefined, floats: undefined, textBoxes: undefined, shapes: undefined })).join("");
    this.relPart = "document.xml";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${tag} ${NS}>${body}</w:${tag}>`;
  }

  package(title: string): Buffer {
    const doc = this.document();
    const hf = this.hfParts.map((h) => ({ ...h, xml: this.hfXml(h) }));
    const media = this.media;
    const flow = this.m.mode === "flow";
    const files: Array<{ name: string; data: Buffer }> = [];
    const add = (name: string, s: string) => files.push({ name, data: Buffer.from(s, "utf8") });
    add("[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
      `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>` +
      hf.map((h) => `<Override PartName="/word/${h.name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${h.kind}+xml"/>`).join("") +
      `</Types>`);
    add("_rels/.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
      `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>`);
    add("word/document.xml", doc);
    add("word/styles.xml", this.styles());
    add("word/settings.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      // 流式模式：自动断词（行尾排版断词已去掉，重新排版时由 Word 断词，行宽与原文接近）
      `<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/>${flow ? '<w:autoHyphenation/>' : ""}` +
      `${hf.some((h) => h.type === "even") ? "<w:evenAndOddHeaders/>" : ""}<w:characterSpacingControl w:val="doNotCompress"/>` +
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
    add("word/_rels/document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` +
      media.filter((m) => m.part === "document.xml").map((m) => `<Relationship Id="${m.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.name}"/>`).join("") +
      hf.map((h) => `<Relationship Id="${h.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${h.kind}" Target="${h.name}"/>`).join("") +
      `</Relationships>`);
    const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    add("docProps/core.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>${esc(title)}</dc:title><dc:creator>文案 Agent</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    add("docProps/app.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>文案 Agent PDF 转换</Application><Pages>${this.m.stats.pages}</Pages></Properties>`);
    // 自定义属性：记录版式模式，编辑时据此提示（逐行一致模式不适合大幅增删文字）
    add("docProps/custom.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="DocAgentLayout"><vt:lpwstr>${this.m.mode}</vt:lpwstr></property></Properties>`);
    for (const h of hf) {
      add(`word/${h.name}`, h.xml);
      const own = media.filter((m) => m.part === h.name);
      if (own.length) add(`word/_rels/${h.name}.rels`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        own.map((m) => `<Relationship Id="${m.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.name}"/>`).join("") + `</Relationships>`);
    }
    for (const m of media) files.push({ name: `word/media/${m.name}`, data: m.data });
    return createZip(files);
  }
}

export function writeDocx(model: DocModel, title: string): Buffer {
  return new Writer(model).package(title);
}

export type { Block };
