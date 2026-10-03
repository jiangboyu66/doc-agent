/**
 * 扩展编辑工具：表格、图片、矢量图形、流程图、图表、公式、分隔符、页面设置、页眉页脚、
 * 项目符号与编号、脚注、超链接、目录、样式、修订审阅、批注管理、移动段落
 *
 * Word 文档使用原生对象（表格、DrawingML 形状组合、原生图表 + 内嵌 Excel、OMML 公式、编号定义、
 * 页码域、TOC 域），在 Word 里都可以继续编辑；Markdown / HTML 生成等价的源码片段插入。
 * 每个工具都只在当前文档支持时出现在模型的工具列表里（isEnabled）。
 */

import { z } from "zod";
import { buildTool, type Tool, type ToolUseContext } from "../Tool.js";
import { DocError, type DocumentAdapter } from "../documents/types.js";
import { editTool, reason } from "./docTools.js";
import { layoutDiagram, imageInfo, type ShapeSpec, type ChartSpec } from "../documents/docx/build.js";
import { tableMarkdown, tableHtml, imageMarkup, shapesSvg, chartSvg, svgDataUri } from "./markup.js";
import { shapesPng, chartPng } from "../services/fallbackImages.js";

const caps = (ctx: ToolUseContext) => ctx.session.doc.capabilities;
const fmt = (ctx: ToolUseContext) => ctx.session.meta.format;
const isRaw = (ctx: ToolUseContext) => !!caps(ctx).rawInsert;
const position = z.enum(["before", "after"]).describe("插在锚点段落之前还是之后");
const anchor = z.string().describe("锚点段落引用（来自 doc_read / doc_search / doc_outline）");
const color = z.string().regex(/^#?([0-9A-Fa-f]{6}|none)$/).describe("6 位十六进制颜色，如 4472C4；none 表示无");


function need<K extends keyof DocumentAdapter>(doc: DocumentAdapter, k: K): NonNullable<DocumentAdapter[K]> {
  const f = doc[k];
  if (typeof f !== "function") throw new DocError("当前文档格式不支持这个操作", "unsupported");
  return (f as Function).bind(doc) as NonNullable<DocumentAdapter[K]>;
}

/** Markdown / HTML：生成标记后插入 */
function rawInsert(doc: DocumentAdapter, ctx: ToolUseContext, p: { anchor: string; position: "before" | "after" }, markup: { markdown: string; html: string }, label: string) {
  const f = fmt(ctx);
  return need(doc, "insertRaw")({ ...p, markup: f === "markdown" ? markup.markdown : markup.html, label }, ctx.editOptions());
}

// ---------------------------------------------------------------------------
// 表格
// ---------------------------------------------------------------------------

export const DocInsertTableTool = editTool({
  name: "doc_insert_table",
  capability: "tableEdit",
  isEnabled: (ctx) => !!caps(ctx).tableEdit || isRaw(ctx),
  title: (i) => `插入表格（${i.rows?.length ?? 0} 行）`,
  inputSchema: z.object({
    anchor, position,
    rows: z.array(z.array(z.string())).min(1).max(500).describe("表格内容，按行给出；第一行默认为表头。单元格内换行用 \\n"),
    header: z.boolean().optional().describe("第一行是否为表头（加粗、每页重复），默认 true"),
    style: z.enum(["grid", "three_line", "plain", "banded"]).optional().describe("grid 网格线（默认）；three_line 学术三线表；plain 无边框；banded 表头底色 + 隔行底纹"),
    align: z.enum(["left", "center", "right"]).optional().describe("表格在页面中的位置，默认居中"),
    widths: z.array(z.number().positive()).optional().describe("各列相对宽度，如 [1,2,1]；默认等宽"),
    width_pct: z.number().min(10).max(100).optional().describe("表格宽度占版心的百分比，默认 100"),
    font_size_pt: z.number().min(5).max(30).optional().describe("表格文字字号，默认沿用正文"),
    caption: z.string().optional().describe("题注，如「表 1 实验结果」"),
    caption_position: z.enum(["above", "below"]).optional().describe("题注位置，表格惯例在上方（默认）"),
    reason,
  }),
  description: () => `新建一个表格，插在锚点段落之前或之后。字体沿用正文；可选网格线 / 三线表 / 无边框 / 隔行底纹样式与题注。
- 先用 doc_read 找到锚点（通常是表格前一段或题注所在段落）；
- 数字列自动居中，表头加粗并在跨页时重复；
- 已有表格加一行用 doc_insert_table_row，改表格结构（删行列、合并、底纹、边框）用 doc_edit_table，改单元格文字用 doc_replace_text。`,
  apply: (doc, i, ctx) => {
    if (caps(ctx).tableEdit) return need(doc, "insertTable")(i, ctx.editOptions());
    return rawInsert(doc, ctx, i, { markdown: tableMarkdown(i.rows, i.header ?? true, i.caption), html: tableHtml(i.rows, i.header ?? true, i.caption, i.style) }, "表格");
  },
});

export const DocEditTableTool = editTool({
  name: "doc_edit_table",
  capability: "tableEdit",
  title: (i) => `修改表格（${i.action ?? ""}）`,
  inputSchema: z.object({
    ref: z.string().describe("表格中某个单元格里的段落引用（doc_read part=\"表格1\" 获取）"),
    action: z.enum(["delete_row", "delete_column", "insert_column", "merge", "shade", "align", "borders", "set_width", "repeat_header", "delete_table"]),
    to_ref: z.string().optional().describe("merge：合并区域另一个角的单元格段落引用"),
    position: position.optional().describe("insert_column：插在该列之前或之后"),
    cells: z.array(z.string()).optional().describe("insert_column：新列每一行的文字（从第一行起）"),
    color: color.optional().describe("shade：底纹颜色"),
    horizontal: z.enum(["left", "center", "right"]).optional().describe("align：水平对齐"),
    vertical: z.enum(["top", "center", "bottom"]).optional().describe("align：垂直对齐"),
    scope: z.enum(["cell", "row", "column", "table"]).optional().describe("shade / align 的作用范围，默认 cell"),
    style: z.enum(["grid", "three_line", "plain", "outer"]).optional().describe("borders：边框样式"),
    width_pct: z.number().min(10).max(100).optional().describe("set_width：表格宽度占版心百分比"),
    reason,
  }),
  destructive: (i) => i.action === "delete_table" || i.action === "delete_column" || i.action === "delete_row",
  permissionContent: (i) => i.action,
  description: () => `修改已有表格的结构与外观：
- delete_row / delete_column：删除 ref 所在的行 / 列；
- insert_column：在 ref 所在列之前/之后插入一列（cells 给出各行文字）；
- merge：合并 ref 与 to_ref 之间的矩形区域（文字并入左上角单元格）；
- shade：设置底纹（scope 可选单元格/整行/整列/整表）；align：设置水平/垂直对齐；
- borders：整表边框改为网格线 / 三线表 / 无边框 / 仅外框；set_width：调整表格宽度；
- repeat_header：首行在每页重复；delete_table：删除整个表格。
结构变化后旧的 P 序号引用失效，请重新读取。修改单元格文字用 doc_replace_text。`,
  apply: (doc, i, ctx) => need(doc, "editTable")(i, ctx.editOptions()),
});

// ---------------------------------------------------------------------------
// 图片 / 图形 / 流程图 / 图表 / 公式
// ---------------------------------------------------------------------------

async function loadImageAsset(ctx: ToolUseContext, name: string): Promise<{ data: Buffer; mime: string; w: number; h: number }> {
  const raw = await ctx.session.readAsset(name);
  const info = imageInfo(raw);
  if (info) return { data: raw, mime: info.mime, w: info.width, h: info.height };
  const { convertToPng } = await import("../services/pdfConvert/render.js");
  const png = await convertToPng(raw);
  if (!png) throw new DocError(`素材"${name}"的格式无法直接插入（Word 支持 PNG / JPEG / GIF / BMP；SVG、WebP 需要图像组件 @napi-rs/canvas 转换）。`, "unsupported");
  return { data: png.png, mime: "image/png", w: png.width, h: png.height };
}

export const DocInsertImageTool = editTool({
  name: "doc_insert_image",
  capability: "media",
  isEnabled: (ctx) => !!caps(ctx).media || isRaw(ctx),
  title: (i) => `插入图片 ${i.asset ?? ""}`,
  inputSchema: z.object({
    anchor, position,
    asset: z.string().describe("用户上传的图片素材文件名（见系统提示中的「已上传素材」）"),
    width_pt: z.number().min(10).max(2000).optional().describe("显示宽度（磅，1 cm ≈ 28.35 pt）；省略则按原尺寸、超出版心时缩小到版心宽度"),
    height_pt: z.number().min(10).max(2000).optional().describe("显示高度；只给宽或高时按比例缩放"),
    align: z.enum(["left", "center", "right"]).optional().describe("默认居中"),
    caption: z.string().optional().describe("图题，如「图 1 系统架构」，放在图片下方"),
    alt: z.string().optional().describe("替代文字（无障碍阅读）"),
    reason,
  }),
  description: () => `把用户上传的图片插入文档（单独成段，默认居中，可加图题）。
- 用户通过输入框上方的「图片」按钮上传，文件名会出现在系统提示的「已上传素材」里；没有素材时请让用户先上传；
- Word 中为嵌入图片，可拖动调整大小；Markdown / HTML 中以内嵌 data URI 保存，文件自包含。`,
  apply: async (doc, i, ctx) => {
    const img = await loadImageAsset(ctx, i.asset);
    if (caps(ctx).media) return need(doc, "insertImage")({ anchor: i.anchor, position: i.position, data: img.data, width_pt: i.width_pt, height_pt: i.height_pt, align: i.align, caption: i.caption, alt: i.alt, name: i.asset }, ctx.editOptions());
    const uri = `data:${img.mime};base64,${img.data.toString("base64")}`;
    const markup = imageMarkup(fmt(ctx) as "markdown" | "html", uri, i.alt ?? i.caption ?? i.asset, i.width_pt, i.caption);
    return rawInsert(doc, ctx, i, { markdown: markup, html: markup }, "图片");
  },
});

const shapeSchema = z.object({
  type: z.enum(["rect", "roundRect", "ellipse", "diamond", "triangle", "parallelogram", "hexagon", "cylinder", "cloud", "line", "arrow", "doubleArrow", "text"]),
  x: z.number().describe("左上角横坐标（磅，相对画布）；线条为起点"),
  y: z.number(),
  w: z.number().positive().optional().describe("宽度（线条可用 x2,y2 代替）"),
  h: z.number().positive().optional(),
  x2: z.number().optional().describe("线条 / 箭头终点"),
  y2: z.number().optional(),
  fill: color.optional(), stroke: color.optional(),
  stroke_width: z.number().min(0.25).max(12).optional(),
  dash: z.enum(["solid", "dash", "dot"]).optional(),
  text: z.string().optional().describe("形状里的文字，\\n 换行"),
  font_size: z.number().min(5).max(72).optional(),
  font_color: color.optional(),
  bold: z.boolean().optional(),
  align: z.enum(["left", "center", "right"]).optional(),
});

export const DocDrawTool = editTool({
  name: "doc_draw",
  capability: "media",
  isEnabled: (ctx) => !!caps(ctx).media || isRaw(ctx),
  title: (i) => `绘制图形（${i.shapes?.length ?? 0} 个形状）`,
  inputSchema: z.object({
    anchor, position,
    width_pt: z.number().min(20).max(2000).describe("画布宽度（磅），超过版心会自动等比缩小"),
    height_pt: z.number().min(10).max(2000).describe("画布高度（磅）"),
    shapes: z.array(shapeSchema).min(1).max(300),
    caption: z.string().optional(),
    alt: z.string().optional(),
    reason,
  }),
  description: () => `在文档中绘制矢量图形：矩形、圆角矩形、椭圆、菱形、三角形、平行四边形、六边形、圆柱、云形、直线、箭头、双向箭头、文字框，任意组合（示意图、标注图、简单插画、版式装饰）。
- 坐标单位是磅（pt），原点在画布左上角；A4 正文宽约 450 pt；
- Word 中是可编辑的形状组合（可逐个选中、改色、改文字）；Markdown / HTML 中为内嵌 SVG；
- 节点 + 连线的流程图 / 结构图请用 doc_insert_diagram（自动排版），数据图表请用 doc_insert_chart。`,
  apply: async (doc, i, ctx) => {
    if (caps(ctx).media) return need(doc, "insertShapes")({ anchor: i.anchor, position: i.position, width_pt: i.width_pt, height_pt: i.height_pt, shapes: i.shapes as ShapeSpec[], caption: i.caption, alt: i.alt, fallbackPng: await shapesPng(i.shapes as ShapeSpec[], i.width_pt, i.height_pt) }, ctx.editOptions());
    const svg = shapesSvg(i.shapes as ShapeSpec[], i.width_pt, i.height_pt);
    return rawInsert(doc, ctx, i, { markdown: imageMarkup("markdown", svgDataUri(svg), i.alt ?? "图形", i.width_pt, i.caption), html: i.caption ? `<figure>${svg}<figcaption>${i.caption}</figcaption></figure>` : svg }, "图形");
  },
});

export const DocInsertDiagramTool = editTool({
  name: "doc_insert_diagram",
  capability: "media",
  isEnabled: (ctx) => !!caps(ctx).media || isRaw(ctx),
  title: (i) => `插入流程图（${i.nodes?.length ?? 0} 个节点）`,
  inputSchema: z.object({
    anchor, position,
    nodes: z.array(z.object({
      id: z.string(), text: z.string(),
      shape: z.enum(["rect", "roundRect", "ellipse", "diamond", "parallelogram", "cylinder"]).optional().describe("开始/结束用 ellipse，判断用 diamond，输入输出用 parallelogram，数据库用 cylinder，其余 roundRect（默认）"),
      fill: color.optional(),
    })).min(1).max(60),
    edges: z.array(z.object({ from: z.string(), to: z.string(), label: z.string().optional(), dashed: z.boolean().optional() })).max(150),
    direction: z.enum(["TB", "LR"]).optional().describe("TB 自上而下（默认），LR 自左向右"),
    font_size: z.number().min(6).max(20).optional(),
    caption: z.string().optional(),
    reason,
  }),
  description: () => `根据"节点 + 连线"自动排版并插入流程图 / 结构图 / 架构图 / 组织结构图，无需计算坐标。
- 节点按连线自动分层排列，超宽时等比缩小到版心宽度；连线可加文字（如"是/否"）和虚线；
- Word 中是可编辑的形状组合；Markdown / HTML 中为内嵌 SVG。`,
  apply: async (doc, i, ctx) => {
    const maxW = fmt(ctx) === "docx" ? 450 : 600;
    let laid;
    try {
      laid = layoutDiagram({ nodes: i.nodes, edges: i.edges, direction: i.direction, maxWidth: maxW, fontSize: i.font_size });
    } catch (e: any) {
      throw new DocError(e.message, "invalid");
    }
    const alt = `流程图：${i.edges.map((e) => `${i.nodes.find((n) => n.id === e.from)?.text} → ${i.nodes.find((n) => n.id === e.to)?.text}`).join("；")}`;
    if (caps(ctx).media) return need(doc, "insertShapes")({ anchor: i.anchor, position: i.position, width_pt: laid.width, height_pt: laid.height, shapes: laid.shapes, caption: i.caption, alt, fallbackPng: await shapesPng(laid.shapes, laid.width, laid.height) }, ctx.editOptions());
    const svg = shapesSvg(laid.shapes, laid.width, laid.height);
    return rawInsert(doc, ctx, i, { markdown: imageMarkup("markdown", svgDataUri(svg), alt, undefined, i.caption), html: i.caption ? `<figure>${svg}<figcaption>${i.caption}</figcaption></figure>` : svg }, "流程图");
  },
});

export const DocInsertChartTool = editTool({
  name: "doc_insert_chart",
  capability: "media",
  isEnabled: (ctx) => !!caps(ctx).media || isRaw(ctx),
  title: (i) => `插入图表（${i.type ?? ""}）`,
  inputSchema: z.object({
    anchor, position,
    type: z.enum(["column", "bar", "line", "pie", "doughnut", "area", "scatter"]).describe("column 柱状图、bar 条形图、line 折线图、pie 饼图、doughnut 环形图、area 面积图、scatter 散点图"),
    title: z.string().optional(),
    categories: z.array(z.string()).min(1).max(200).describe("分类轴标签（散点图为 x 数值）"),
    series: z.array(z.object({ name: z.string(), values: z.array(z.number()) })).min(1).max(20).describe("数据系列，每个系列的值与分类一一对应；饼图只用第一个系列"),
    x_title: z.string().optional(), y_title: z.string().optional(),
    stacked: z.boolean().optional().describe("柱状/条形/折线/面积图是否堆叠"),
    show_values: z.boolean().optional().describe("显示数据标签"),
    legend: z.enum(["bottom", "right", "top", "none"]).optional(),
    width_pt: z.number().min(100).max(1000).optional(), height_pt: z.number().min(60).max(1000).optional(),
    caption: z.string().optional().describe("图题，如「图 2 各季度销售额」"),
    reason,
  }),
  description: () => `根据数据插入图表。Word 中是原生图表（数据内嵌为 Excel 工作簿，用户可在 Word 里右键"编辑数据"、改样式），Markdown / HTML 中为内嵌 SVG。
数据必须来自文档或用户提供的数字，不要编造数据；数据来自文档中的表格时，先 doc_read 读取再填入。`,
  apply: async (doc, i, ctx) => {
    const spec: ChartSpec = { type: i.type, title: i.title, categories: i.categories, series: i.series, x_title: i.x_title, y_title: i.y_title, stacked: i.stacked, show_values: i.show_values, legend: i.legend };
    if (caps(ctx).media) {
      const wpt = i.width_pt ?? 430, hpt = i.height_pt ?? wpt * 0.6;
      const ok = spec.series.every((s) => s.values.length === spec.categories.length);
      return need(doc, "insertChart")({ anchor: i.anchor, position: i.position, chart: spec, width_pt: i.width_pt, height_pt: i.height_pt, caption: i.caption, fallbackPng: ok ? await chartPng(spec, wpt, hpt) : undefined }, ctx.editOptions());
    }
    for (const s of spec.series) if (s.values.length !== spec.categories.length) throw new DocError(`系列"${s.name}"的值个数与分类个数不一致`, "invalid");
    const svg = chartSvg(spec, i.width_pt, i.height_pt);
    return rawInsert(doc, ctx, i, { markdown: imageMarkup("markdown", svgDataUri(svg), i.title ?? "图表", undefined, i.caption), html: i.caption ? `<figure>${svg}<figcaption>${i.caption}</figcaption></figure>` : svg }, "图表");
  },
});

export const DocInsertEquationTool = editTool({
  name: "doc_insert_equation",
  capability: "equations",
  isEnabled: (ctx) => !!caps(ctx).equations || isRaw(ctx),
  title: (i) => `插入公式${i.number ? ` ${i.number}` : ""}`,
  inputSchema: z.object({
    latex: z.string().min(1).describe("LaTeX 公式（不含 $ 符号），如 E = mc^2、\\frac{a}{b}、\\sum_{i=1}^{n} x_i"),
    display: z.boolean().optional().describe("true 行间公式（单独成段，默认）；false 行内公式（插在段落文字中）"),
    anchor: z.string().optional().describe("行间公式：插在哪个段落前后"),
    position: position.optional(),
    number: z.string().optional().describe("行间公式编号，如 (1)，与公式在同一行右端"),
    ref: z.string().optional().describe("行内公式：所在段落引用"),
    after_text: z.string().optional().describe("行内公式：插在段落中这段文字之后（省略则段尾）"),
    reason,
  }),
  description: () => `插入数学公式。Word 中转换为原生公式（OMML），可在 Word 公式编辑器里继续修改；Markdown 中写成 $$…$$ / $…$，HTML 中写成 \\[…\\]（MathJax 可渲染）。
支持常用 LaTeX：分式 \\frac、根式 \\sqrt[n]{}、上下标 ^ _、求和积分 \\sum \\prod \\int \\oint（含上下限）、\\left( \\right)、希腊字母、运算符与关系符、\\text{}、\\mathbf、\\mathbb、\\hat \\bar \\vec 等、常用函数 \\sin \\log \\lim \\max。
带编号的行间公式会把编号放在同一行右端（居中制表位 + 右对齐制表位），不会折行。`,
  apply: (doc, i, ctx) => {
    if (caps(ctx).equations) return need(doc, "insertEquation")(i, ctx.editOptions());
    const display = i.display ?? !i.ref;
    if (!display) {
      if (!i.ref) throw new DocError("行内公式需要 ref", "invalid");
      const tex = fmt(ctx) === "markdown" ? `$${i.latex}$` : `\\(${i.latex}\\)`;
      if (i.after_text) return doc.replaceText({ ref: i.ref, oldText: i.after_text, newText: `${i.after_text}${tex}` }, ctx.editOptions());
      throw new DocError("Markdown / HTML 行内公式请给出 after_text（插在哪段文字之后）", "invalid");
    }
    if (!i.anchor) throw new DocError("行间公式需要 anchor", "invalid");
    const num = i.number ? ` \\tag{${i.number.replace(/[()]/g, "")}}` : "";
    return rawInsert(doc, ctx, { anchor: i.anchor, position: i.position ?? "after" }, { markdown: `$$\n${i.latex}${num}\n$$`, html: `<p class="math">\\[${i.latex}${num}\\]</p>` }, "公式");
  },
});

// ---------------------------------------------------------------------------
// 分隔符 / 页面设置 / 页眉页脚 / 目录
// ---------------------------------------------------------------------------

export const DocInsertBreakTool = editTool({
  name: "doc_insert_break",
  capability: "layout",
  isEnabled: (ctx) => !!caps(ctx).layout || isRaw(ctx),
  title: (i) => `插入${{ page: "分页符", column: "分栏符" }[i.kind as "page"] ?? "分节符"}`,
  inputSchema: z.object({
    anchor, position,
    kind: z.enum(["page", "column", "section_next_page", "section_continuous", "section_odd_page", "section_even_page"]).describe("page 分页符；column 分栏符；section_* 分节符（下一页 / 连续 / 奇数页 / 偶数页）"),
    reason,
  }),
  description: () => `插入分页符、分栏符或分节符。
- 让某部分另起一页：分页符；
- 需要某几页横向、不同页边距、不同分栏或不同页眉页脚：在前后各插一个分节符，再对该节调用 doc_page_setup / doc_header_footer（传入该节中某段的 ref）；
- 也可以用 doc_set_paragraph 的 page_break_before 让标题段落总在新页开始。
Markdown / HTML 只支持分页符（打印时生效）。`,
  apply: (doc, i, ctx) => {
    if (caps(ctx).layout) return need(doc, "insertBreak")(i, ctx.editOptions());
    if (i.kind !== "page") throw new DocError("Markdown / HTML 只支持分页符", "unsupported");
    const div = `<div style="page-break-after: always; break-after: page;"></div>`;
    return rawInsert(doc, ctx, i, { markdown: div, html: div }, "分页符");
  },
});

export const DocPageSetupTool = editTool({
  name: "doc_page_setup",
  capability: "layout",
  title: () => "页面设置",
  inputSchema: z.object({
    ref: z.string().optional().describe("只修改这个段落所在的节；省略则修改全部节"),
    orientation: z.enum(["portrait", "landscape"]).optional(),
    paper: z.enum(["A4", "A3", "A5", "B5", "Letter", "Legal"]).optional(),
    margins_pt: z.object({ top: z.number().min(0).max(400).optional(), bottom: z.number().min(0).max(400).optional(), left: z.number().min(0).max(400).optional(), right: z.number().min(0).max(400).optional(), header: z.number().min(0).max(300).optional(), footer: z.number().min(0).max(300).optional() }).optional().describe("页边距（磅，1 cm ≈ 28.35 pt；Word 默认上下 72、左右 90）"),
    columns: z.number().int().min(1).max(6).optional().describe("分栏数"),
    column_gap_pt: z.number().min(0).max(144).optional(),
    line_numbers: z.boolean().optional().describe("显示行号（投稿审阅常用）"),
    v_align: z.enum(["top", "center", "bottom"]).optional().describe("页面垂直对齐（封面常用 center）"),
    reason,
  }),
  description: () => `修改页面设置：纸张大小、纸张方向（横向/纵向）、页边距、分栏、行号、页面垂直对齐。
默认作用于全部节；只想改某几页时，先用 doc_insert_break 插入分节符，再传入该节中某段的 ref。`,
  apply: (doc, i, ctx) => need(doc, "pageSetup")(i, ctx.editOptions()),
});

export const DocHeaderFooterTool = editTool({
  name: "doc_header_footer",
  capability: "layout",
  title: (i) => `设置${i.kind === "footer" ? "页脚" : "页眉"}`,
  inputSchema: z.object({
    kind: z.enum(["header", "footer"]),
    text: z.string().describe("内容；可用占位符 {PAGE} 当前页码、{NUMPAGES} 总页数、{SECTIONPAGES} 本节页数、{DATE} 日期，\\n 换行。例：第 {PAGE} 页 共 {NUMPAGES} 页"),
    align: z.enum(["left", "center", "right"]).optional().describe("默认居中"),
    ref: z.string().optional().describe("只设置这个段落所在的节；省略则全部节"),
    first_page: z.boolean().optional().describe("设置首页单独的页眉/页脚（封面常用：首页留空）"),
    reason,
  }),
  destructive: () => true,
  description: () => `设置页眉或页脚（会替换原有的默认页眉/页脚），支持自动页码域。
- 只改页眉页脚里的个别文字时，用 doc_read part="页眉"/"页脚" 读取后 doc_replace_text 修改，保留原有格式；
- 需要完全重设（加页码、换内容）时使用本工具。`,
  apply: (doc, i, ctx) => need(doc, "setHeaderFooter")(i, ctx.editOptions()),
});

export const DocInsertTocTool = editTool({
  name: "doc_insert_toc",
  capability: "layout",
  isEnabled: (ctx) => !!caps(ctx).layout || fmt(ctx) === "markdown",
  title: () => "插入目录",
  inputSchema: z.object({
    anchor, position,
    levels: z.number().int().min(1).max(9).optional().describe("包含几级标题，默认 3"),
    title: z.string().optional().describe("目录标题，默认「目  录」；空字符串表示不加标题"),
    reason,
  }),
  description: () => `插入目录。Word 中是自动目录（TOC 域）：按标题样式生成，打开文件时提示更新以填入页码，之后标题变化也可一键更新；Markdown 中生成带锚点链接的标题列表。
标题必须使用标题样式（大纲中能看到的段落）才会进入目录。`,
  apply: (doc, i, ctx) => {
    if (caps(ctx).layout) return need(doc, "insertToc")(i, ctx.editOptions());
    const heads = doc.summary().headings.filter((h) => h.level <= (i.levels ?? 3));
    const slug = (t: string) => t.trim().toLowerCase().replace(/[^\p{L}\p{N}\s-]/gu, "").replace(/\s+/g, "-");
    const md = `${i.title === "" ? "" : `**${i.title ?? "目录"}**\n\n`}${heads.map((h) => `${"  ".repeat(Math.max(0, h.level - 1))}- [${h.text}](#${slug(h.text)})`).join("\n")}`;
    return rawInsert(doc, ctx, i, { markdown: md, html: md }, "目录");
  },
});

// ---------------------------------------------------------------------------
// 列表 / 脚注 / 超链接 / 样式
// ---------------------------------------------------------------------------

export const DocSetListTool = editTool({
  name: "doc_set_list",
  capability: "lists",
  title: (i) => `设置编号（${(i.refs ?? []).length} 段）`,
  inputSchema: z.object({
    refs: z.array(z.string()).min(1).max(300),
    kind: z.enum(["bullet", "number", "chinese", "outline", "none"]).describe("bullet 项目符号；number 1. a) i)；chinese 一、（一）1.；outline 1. 1.1 1.1.1；none 取消编号"),
    level: z.number().int().min(1).max(9).optional().describe("列表级别，1 为最外层"),
    restart: z.boolean().optional().describe("从 1 重新开始编号（默认接续本文档之前同类编号）"),
    reason,
  }),
  description: () => `为段落设置项目符号或自动编号（Word 原生编号，增删条目后序号自动更新），或取消编号。
多个段落一次设置可以保证序号连续；不同级别分开调用并指定 level。`,
  apply: (doc, i, ctx) => need(doc, "setList")(i, ctx.editOptions()),
});

export const DocInsertFootnoteTool = editTool({
  name: "doc_insert_footnote",
  capability: "notes",
  title: (i) => `插入${i.kind === "endnote" ? "尾注" : "脚注"}`,
  inputSchema: z.object({
    ref: z.string().describe("正文段落引用"),
    after_text: z.string().optional().describe("注号放在这段文字之后（须在段内唯一）；省略则放在句末标点之前"),
    text: z.string().min(1).describe("注释内容"),
    kind: z.enum(["footnote", "endnote"]).optional().describe("footnote 脚注（页面底部，默认）；endnote 尾注（文末）"),
    reason,
  }),
  description: () => `插入脚注或尾注：在正文中加上标注号，注释内容放在页面底部 / 文末，编号由 Word 自动维护。`,
  apply: (doc, i, ctx) => need(doc, "insertNote")(i, ctx.editOptions()),
});

export const DocInsertLinkTool = editTool({
  name: "doc_insert_link",
  capability: "links",
  title: (i) => `添加链接 ${i.url ?? ""}`,
  inputSchema: z.object({
    ref: z.string(),
    text: z.string().describe("段落中要变成链接的文字（须唯一）"),
    url: z.string().describe("链接地址：http(s):// 或 mailto:"),
    reason,
  }),
  description: () => `把段落中的一段文字设为超链接（Word 使用文档的"超链接"字符样式；Markdown 改写为 [文字](地址)）。`,
  apply: (doc, i, ctx) => need(doc, "insertLink")(i, ctx.editOptions()),
});

export const DocModifyStyleTool = editTool({
  name: "doc_modify_style",
  capability: "styleEdit",
  title: (i) => `修改样式「${i.style ?? ""}」`,
  inputSchema: z.object({
    style: z.string().describe("样式名称或 ID（doc_styles 查看），如 正文 / Normal、标题 1 / Heading 1"),
    type: z.enum(["paragraph", "character"]).optional(),
    create: z.boolean().optional().describe("样式不存在时新建"),
    based_on: z.string().optional().describe("新建时基于哪个样式"),
    font: z.string().optional(), size_pt: z.number().min(1).max(200).optional(),
    bold: z.boolean().optional(), italic: z.boolean().optional(), color: z.string().regex(/^#?[0-9A-Fa-f]{6}$/).optional(),
    alignment: z.enum(["left", "center", "right", "justify"]).optional(),
    space_before_pt: z.number().min(0).max(1000).optional(), space_after_pt: z.number().min(0).max(1000).optional(),
    line_spacing: z.number().min(0.5).max(5).optional(),
    first_line_pt: z.number().min(0).max(200).optional(), indent_left_pt: z.number().min(0).max(500).optional(),
    keep_with_next: z.boolean().optional(),
    outline_level: z.number().int().min(0).max(9).optional().describe("大纲级别（1–9，0 表示正文级别）：决定是否进入目录"),
    reason,
  }),
  destructive: () => true,
  description: () => `修改（或新建）样式定义。修改后，文档中所有使用该样式的段落一起变化——这是统一全文排版（如"正文改为宋体小四、1.5 倍行距、首行缩进 2 字"）最省事、最规范的方式。
- 先用 doc_styles 查看样式名称与使用次数；
- 段落上的直接格式优先于样式，个别段落不变时用 doc_inspect 检查；
- 中文排版常用：小四 = 12pt，五号 = 10.5pt，四号 = 14pt，三号 = 16pt；首行缩进 2 字 ≈ 字号 × 2。`,
  apply: (doc, i, ctx) => need(doc, "modifyStyle")(i, ctx.editOptions()),
});

// ---------------------------------------------------------------------------
// 审阅：修订、批注；移动段落
// ---------------------------------------------------------------------------

export const DocReviewChangesTool = editTool({
  name: "doc_review_changes",
  capability: "revisions",
  title: (i) => `${i.action === "reject" ? "拒绝" : "接受"}修订`,
  inputSchema: z.object({
    action: z.enum(["accept", "reject"]),
    refs: z.array(z.string()).optional().describe("只处理这些段落中的修订；省略则全文"),
    author: z.string().optional().describe("只处理某位作者的修订"),
    reason,
  }),
  destructive: () => true,
  description: () => `接受或拒绝文档中的修订（插入、删除、格式修改、表格行修订）。省略 refs 表示全文。
这是用户在 Word"审阅"里才会做的决定：只有用户明确要求时才调用。`,
  apply: (doc, i, ctx) => need(doc, "reviewChanges")(i, ctx.editOptions()),
});

export const DocCommentsTool = buildTool({
  name: "doc_comments",
  category: "edit",
  description: () => `查看或删除批注。action=list 列出全部批注（id、作者、内容、所在段落）；action=delete 删除指定 id（或 all=true 全部删除）。
添加批注用 doc_add_comment。删除批注需要用户明确要求。`,
  inputSchema: z.object({
    action: z.enum(["list", "delete"]),
    ids: z.array(z.string()).optional(),
    all: z.boolean().optional(),
  }),
  userFacingName: (i) => (i.action === "delete" ? `删除批注${i.all ? "（全部）" : ""}` : "查看批注"),
  isEnabled: (ctx) => ctx.session.meta.format === "docx",
  isReadOnly: (i) => i.action === "list",
  isConcurrencySafe: (i) => i.action === "list",
  isDestructive: (i) => i.action === "delete",
  checkPermissions: async (i) => (i.action === "list" ? { behavior: "allow" } : { behavior: "passthrough" }),
  async call(i, ctx) {
    const doc = ctx.session.doc;
    if (i.action === "list") {
      const list = need(doc, "listComments")();
      if (!list.length) return { content: "文档中没有批注。" };
      return { content: list.map((c) => `[id=${c.id}] ${c.author}${c.date ? `（${c.date.slice(0, 10)}）` : ""}${c.ref ? ` @${c.ref}「${c.anchor}」` : ""}：${c.text}`).join("\n") };
    }
    const r = need(doc, "deleteComments")({ ids: i.ids, all: i.all });
    return { content: r.summary, docChange: { label: r.summary, changedRefs: [], structural: false } };
  },
});

export const DocMoveBlocksTool = editTool({
  name: "doc_move_blocks",
  capability: "move",
  title: (i) => `移动 ${(i.refs ?? []).length} 个段落`,
  inputSchema: z.object({
    refs: z.array(z.string()).min(1).max(200).describe("要移动的段落（保持原有先后顺序）"),
    anchor, position,
    reason,
  }),
  description: () => `把段落整体移动到另一个位置（调整章节顺序、把结论提前等），段落内容与格式完全不变。比"删除再插入"更安全：不会丢失段内格式、编号、批注。`,
  apply: (doc, i, ctx) => need(doc, "moveBlocks")(i, ctx.editOptions()),
});

export const EDITING_TOOLS: Tool[] = [
  DocInsertTableTool, DocEditTableTool, DocInsertImageTool, DocDrawTool, DocInsertDiagramTool, DocInsertChartTool,
  DocInsertEquationTool, DocInsertBreakTool, DocPageSetupTool, DocHeaderFooterTool, DocInsertTocTool, DocSetListTool,
  DocInsertFootnoteTool, DocInsertLinkTool, DocModifyStyleTool, DocReviewChangesTool, DocCommentsTool, DocMoveBlocksTool,
] as Tool[];
