/**
 * 文档工具集
 *
 * 每个工具的 description 都按"工具说明书六要素"撰写：功能、前置条件、格式要求、边界情况、
 * 危险提示、推荐与不推荐的用法（参考 cc-src-learning 第四章对 FileEditTool/BashTool 提示词的分析）。
 */

import { z } from "zod";
import { buildTool, type Tool, type ToolUseContext, type ValidationResult } from "../Tool.js";
import { DocError, type BlockInfo, type DocumentAdapter, type EditResult } from "../documents/types.js";
import { clip } from "../documents/textMatch.js";
import { exportDocument } from "../services/convert.js";

// ---------------------------------------------------------------------------
// 公共辅助
// ---------------------------------------------------------------------------

const MAX_PARA_CHARS = 1500;

function markRead(ctx: ToolUseContext) {
  if (!ctx.session.meta.hasRead) ctx.session.meta.hasRead = true;
}

function requireRead(ctx: ToolUseContext): ValidationResult {
  return ctx.session.meta.hasRead
    ? { ok: true }
    : { ok: false, message: "修改前必须先读取文档：请先调用 doc_outline、doc_read 或 doc_search 了解原文，再进行修改。" };
}

function formatBlock(b: BlockInfo, full = false): string {
  const meta = [b.location, b.style, b.kind === "heading" ? `标题${b.level ?? ""}` : b.kind === "title" ? "文档标题" : null, ...(b.flags ?? [])].filter(Boolean).join(" · ");
  let text = b.text.replace(/\n/g, "⏎");
  if (!full && text.length > MAX_PARA_CHARS) text = text.slice(0, MAX_PARA_CHARS) + `…（已截断，共 ${b.text.length} 字，可用 doc_read 的 ref + full=true 查看全文）`;
  return `[${b.ref}] (${meta}) ${text || "（空段落）"}`;
}

function editResultText(r: EditResult): string {
  const lines = [r.summary];
  for (const p of r.preview.slice(0, 8)) {
    if (p.before && p.after) lines.push(`- ${p.ref}\n  改前：${clip(p.before, 300)}\n  改后：${clip(p.after, 300)}`);
    else if (p.after) lines.push(`- ${p.ref || "新内容"}：${clip(p.after, 300)}`);
    else lines.push(`- ${p.ref} 已删除：${clip(p.before, 200)}`);
  }
  if (r.preview.length > 8) lines.push(`…另有 ${r.preview.length - 8} 处`);
  return lines.join("\n");
}

function cap(ctx: ToolUseContext, k: keyof DocumentAdapter["capabilities"]) {
  return ctx.session.doc.capabilities[k];
}

/** 写工具的通用骨架：先读后改校验 + 干跑预览 + 执行 + 生成版本 */
function editTool<I extends { reason?: string }>(def: {
  name: string;
  capability: keyof DocumentAdapter["capabilities"];
  description: (ctx: ToolUseContext) => string;
  inputSchema: z.ZodType<I>;
  title: (i: Partial<I>) => string;
  apply: (doc: DocumentAdapter, input: I, ctx: ToolUseContext) => EditResult;
  destructive?: (i: I) => boolean;
  permissionContent?: (i: I) => string;
  checkPermissions?: Tool<I>["checkPermissions"];
}): Tool<I> {
  return buildTool<I>({
    name: def.name,
    category: "edit",
    description: def.description,
    inputSchema: def.inputSchema,
    userFacingName: def.title,
    isEnabled: (ctx) => cap(ctx, def.capability),
    isDestructive: (i) => def.destructive?.(i) ?? false,
    permissionContent: def.permissionContent,
    checkPermissions: def.checkPermissions,
    validateInput: async (_i, ctx) => requireRead(ctx),
    async preview(input, ctx) {
      return (await ctx.session.dryRun((doc) => def.apply(doc, input, ctx))).preview;
    },
    async call(input, ctx) {
      const r = def.apply(ctx.session.doc, input, ctx);
      return {
        content: editResultText(r),
        preview: r.preview,
        docChange: { label: `${def.title(input)}${input.reason ? "：" + input.reason : ""}`, changedRefs: r.changedRefs, structural: r.structural },
      };
    },
  });
}

// ---------------------------------------------------------------------------
// 只读工具
// ---------------------------------------------------------------------------

export const DocOutlineTool = buildTool({
  name: "doc_outline",
  category: "read",
  description: () => `获取文档概览：格式、段落数、字数、标题大纲（含段落引用）、表格/图片/分节数量、页眉页脚等部件、检测到的模板、已有修订与批注数量。

使用时机：
- 任何任务开始时先调用它，了解文档结构，再决定读取哪些部分。
- 处理长文档时，用大纲里的引用配合 doc_read 按章节读取，不要一次性读取全文。
这是只读操作，可以和其它只读工具并行调用。`,
  inputSchema: z.object({}),
  userFacingName: () => "查看文档大纲",
  isReadOnly: () => true,
  isConcurrencySafe: () => true,
  async call(_i, ctx) {
    markRead(ctx);
    const s = ctx.session.doc.summary();
    const lines = [
      `文件：${ctx.session.meta.filename}（${s.format}）｜段落/块 ${s.blockCount} 个｜正文约 ${s.charCount} 字`,
      `表格 ${s.tables}｜图片 ${s.images}｜分节 ${s.sections}｜已有修订标记 ${s.trackedChanges}｜批注 ${s.comments}`,
    ];
    if (s.template) lines.push(`检测到模板：${s.template.name}（依据：${s.template.evidence.join("；")}）`);
    if (s.parts.length > 1) lines.push(`文档部件：${s.parts.join("、")}`);
    lines.push("", "标题大纲：");
    if (!s.headings.length) lines.push("（未识别到标题结构，可用 doc_read 顺序浏览）");
    for (const h of s.headings) lines.push(`${"  ".repeat(Math.max(0, h.level - 1))}- [${h.ref}] ${clip(h.text.replace(/⟨分栏符⟩|⟨分页符⟩/g, "").trim(), 100)}`);
    if (s.notes.length) lines.push("", ...s.notes.map((n) => `注意：${n}`));
    return { content: lines.join("\n"), data: s };
  },
});

export const DocReadTool = buildTool({
  name: "doc_read",
  category: "read",
  description: () => `按顺序读取文档的段落/块，每行格式：[引用] (位置 · 样式 · 标记) 文本。

参数：
- ref：从这个段落开始读（引用来自 doc_outline / doc_search / 之前的 doc_read）；
- offset：或者从第几个段落开始（1 起）；二者都不给则从头开始；
- limit：读取多少段（默认 40，最大 200）；
- part：只读某一部分，按位置前缀过滤，例如 "页眉"、"页脚"、"脚注"、"表格1"；
- full：为 true 时不截断超长段落。

文本中的特殊记号：⟨图片⟩⟨公式⟩⟨文本框⟩⟨脚注N⟩⟨符号X⟩⟨分页符⟩⟨分栏符⟩ 代表不可编辑的对象，修改文字时必须原样保留它们；⏎ 表示段内换行，\\t 为制表符。
段落引用以 # 开头的是稳定引用；形如 P12@v3 的是临时引用，插入/删除段落后会失效，需重新读取。
这是只读操作，可以并行调用。`,
  inputSchema: z.object({
    ref: z.string().optional(),
    offset: z.number().int().min(0).optional().describe("从第几段开始（1 起算）"),
    limit: z.number().int().min(1).max(200).optional(),
    part: z.string().optional(),
    full: z.boolean().optional(),
  }),
  userFacingName: (i) => (i.ref ? `读取 ${i.ref} 起的段落` : i.part ? `读取${i.part}` : "读取文档"),
  isReadOnly: () => true,
  isConcurrencySafe: () => true,
  async call(i, ctx) {
    markRead(ctx);
    const all = ctx.session.doc.listBlocks();
    const pool = i.part ? all.filter((b) => b.location.startsWith(i.part!) || b.location.includes(i.part!)) : all;
    let start = 0;
    if (i.ref) {
      const target = ctx.session.doc.getBlock(i.ref).ref;
      start = pool.findIndex((b) => b.ref === target);
      if (start === -1) throw new DocError(`段落 ${i.ref} 不在所选范围内`, "not_found");
    } else if (i.offset) start = Math.max(0, i.offset - 1);
    const limit = i.limit ?? 40;
    const slice = pool.slice(start, start + limit);
    const lines = slice.map((b) => formatBlock(b, i.full));
    const end = start + slice.length;
    lines.push(`\n（第 ${start + 1}–${end} 段，共 ${pool.length} 段${end < pool.length ? `；继续读取请用 offset=${end + 1}` : "；已到末尾"}）`);
    return { content: lines.join("\n") };
  },
});

export const DocSearchTool = buildTool({
  name: "doc_search",
  category: "read",
  description: () => `在整篇文档（含表格、页眉页脚、脚注）中搜索文字，返回命中段落的引用与上下文，命中部分用【】标出。

- 默认忽略大小写，并容忍弯/直引号、全/半角破折号、空白差异；
- regex=true 时按 JavaScript 正则表达式搜索；
- 在修改前用它确认原文的准确写法与出现次数：出现多次时，修改需要指定 ref 或提供更长的唯一片段。
这是只读操作，可以并行调用。`,
  inputSchema: z.object({
    query: z.string().min(1),
    regex: z.boolean().optional(),
    case_sensitive: z.boolean().optional(),
    max_results: z.number().int().min(1).max(200).optional(),
  }),
  userFacingName: (i) => `搜索"${clip(i.query ?? "", 30)}"`,
  isReadOnly: () => true,
  isConcurrencySafe: () => true,
  async call(i, ctx) {
    markRead(ctx);
    const hits = ctx.session.doc.search(i.query, { regex: i.regex, caseSensitive: i.case_sensitive });
    if (!hits.length) return { content: `未找到"${i.query}"。可以尝试更短的关键词或 regex=true。` };
    const max = i.max_results ?? 30;
    const lines = hits.slice(0, max).map((h) => {
      const a = Math.max(0, h.index - 60);
      const b = Math.min(h.text.length, h.index + h.match.length + 60);
      const ctxText = `${a > 0 ? "…" : ""}${h.text.slice(a, h.index)}【${h.match}】${h.text.slice(h.index + h.match.length, b)}${b < h.text.length ? "…" : ""}`;
      return `[${h.ref}] (${h.location}) ${ctxText.replace(/\n/g, "⏎")}`;
    });
    const paras = new Set(hits.map((h) => h.ref)).size;
    lines.push(`\n共 ${hits.length} 处命中，分布在 ${paras} 个段落${hits.length > max ? `（仅显示前 ${max} 处）` : ""}。`);
    return { content: lines.join("\n") };
  },
});

export const DocInspectTool = buildTool({
  name: "doc_inspect",
  category: "read",
  description: () => `查看某个段落的详细格式：段落样式、对齐、间距、缩进，以及段内每个 run（格式一致的一段文字）的加粗/斜体/上下标/字号/字体/颜色。
在需要保证修改后的格式与原文一致、或排查"为什么这段看起来不一样"时使用。只读，可并行。`,
  inputSchema: z.object({ ref: z.string() }),
  userFacingName: (i) => `查看 ${i.ref} 的格式`,
  isReadOnly: () => true,
  isConcurrencySafe: () => true,
  async call(i, ctx) {
    markRead(ctx);
    return { content: ctx.session.doc.describeFormat(i.ref) };
  },
});

export const DocStylesTool = buildTool({
  name: "doc_styles",
  category: "read",
  description: () => `列出文档中定义的段落样式与字符样式，以及每个样式在文档中被使用的次数。
插入新段落或调整段落样式之前调用，选择文档中已在使用的样式（例如论文模板的专用标题样式），保证与全文一致。只读，可并行。`,
  inputSchema: z.object({ type: z.enum(["paragraph", "character", "all"]).optional() }),
  userFacingName: () => "查看样式列表",
  isEnabled: (ctx) => cap(ctx, "styles"),
  isReadOnly: () => true,
  isConcurrencySafe: () => true,
  async call(i, ctx) {
    markRead(ctx);
    const list = ctx.session.doc.listStyles().filter((s) => !i.type || i.type === "all" || s.type === i.type);
    const used = list.filter((s) => s.inUse > 0);
    const unused = list.filter((s) => s.inUse === 0);
    return {
      content: [
        "文档中正在使用的样式（名称 ｜ ID ｜ 类型 ｜ 使用次数）：",
        ...used.map((s) => `- ${s.name} ｜ ${s.id} ｜ ${s.type === "paragraph" ? "段落" : "字符"} ｜ ${s.inUse}`),
        "",
        `另有 ${unused.length} 个已定义但未使用的样式：${unused.slice(0, 40).map((s) => s.name).join("、")}${unused.length > 40 ? "…" : ""}`,
      ].join("\n"),
    };
  },
});

export const DocVerifyTool = buildTool({
  name: "doc_verify",
  category: "read",
  description: () => `保真校验：对比当前文档与用户上传的原始文件，报告哪些部件/段落发生了变化、其余部分是否逐字节保持原样，并检查 XML 结构、修订 ID 等是否有效。
render=true 时额外用 LibreOffice 渲染一次，确认文件可以被正常打开（较慢）。
在完成一批修改后、交付前调用。只读。`,
  inputSchema: z.object({ render: z.boolean().optional() }),
  userFacingName: () => "保真校验",
  isReadOnly: () => true,
  isConcurrencySafe: (i) => !i.render,
  async call(i, ctx) {
    const rep = ctx.session.doc.fidelity(await ctx.session.original());
    const lines = [
      rep.ok ? "✓ 结构校验通过" : "✗ 发现问题",
      `部件：共 ${rep.totalParts} 个，其中 ${rep.identicalParts} 个与原文件逐字节一致`,
    ];
    for (const m of rep.modifiedParts) lines.push(`- 已修改 ${m.name}${m.changedBlocks.length ? `：变化段落 ${m.changedBlocks.slice(0, 30).join("、")}${m.changedBlocks.length > 30 ? "…" : ""}` : "（清单/关系文件的登记变化）"}`);
    if (rep.addedParts.length) lines.push(`- 新增部件：${rep.addedParts.join("、")}`);
    for (const p of rep.problems) lines.push(`! ${p}`);
    if (i.render && ctx.session.meta.format !== "pdf") {
      try {
        const out = await exportDocument(ctx.session.current(), ctx.session.meta.format, "pdf");
        lines.push(`✓ LibreOffice 渲染成功（PDF ${Math.round(out.data.length / 1024)} KB），文件可正常打开`);
      } catch (e: any) {
        lines.push(`! 渲染检查失败：${e.message}`);
      }
    }
    return { content: lines.join("\n"), data: rep, isError: !rep.ok };
  },
});

// ---------------------------------------------------------------------------
// 写工具
// ---------------------------------------------------------------------------

const reason = z.string().describe("一句话说明这次修改的目的，会展示给用户确认").optional();

export const DocReplaceTextTool = editTool({
  name: "doc_replace_text",
  capability: "replaceText",
  title: (i) => `修改文字${i.ref ? `（${i.ref}）` : ""}`,
  inputSchema: z.object({
    old_text: z.string().min(1).describe("要修改的原文片段，必须与文档原文一致（弯直引号、空白差异可容忍）"),
    new_text: z.string().describe("修改后的文字；为空字符串表示删除这段文字"),
    ref: z.string().optional().describe("限定在某个段落内查找（强烈建议提供）"),
    replace_all: z.boolean().optional().describe("old_text 出现多次时全部替换"),
    reason,
  }),
  destructive: (i) => !!i.replace_all,
  permissionContent: (i) => (i.replace_all ? "replace_all" : ""),
  description: () => `在段落内做精确的文字修改，完整保留原有格式。

工作原理：你给出原文片段 old_text 和修改后的 new_text，系统计算两者的最小差异，只改动真正变化的词；未变化的字符原样保留在原来的格式里（加粗、斜体、上标、字体、颜色、超链接、域都不受影响）。新写入的文字继承被替换位置的格式。

使用要求：
- 修改前必须已经用 doc_read / doc_search 读过原文。old_text 要从读取结果中复制，不要凭记忆编写。
- 强烈建议同时给出 ref，把查找限定在一个段落里。
- old_text 必须在范围内唯一；出现多次会失败并列出所有位置，这时请给 ref 或更长的片段，确实要全部替换才用 replace_all=true。
- 只包含需要修改的那一句或一小段即可，不要把整段长文都放进 old_text/new_text。
- 不能跨段落修改：段落之间的分隔不是换行符，跨段修改请逐段进行。
- ⟨图片⟩⟨公式⟩⟨脚注N⟩⟨符号X⟩ 等对象不可编辑：修改它附近的文字时，在 old_text 和 new_text 中都原样保留它。
- 不要用它修改格式（加粗、字号、样式）——格式请用 doc_format_text / doc_set_paragraph。
- 修改失败时根据错误信息调整参数重试，不要反复用相同参数调用。`,
  apply: (doc, i, ctx) => doc.replaceText({ ref: i.ref, oldText: i.old_text, newText: i.new_text, replaceAll: i.replace_all }, ctx.editOptions()),
});

export const DocFormatTextTool = editTool({
  name: "doc_format_text",
  capability: "formatText",
  title: (i) => `设置文字格式（${i.ref ?? ""}）`,
  inputSchema: z.object({
    ref: z.string(),
    text: z.string().optional().describe("段落中要设置格式的文字（须唯一）；省略则作用于整段"),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.enum(["single", "double", "none"]).optional(),
    strike: z.boolean().optional(),
    color: z.string().regex(/^#?[0-9A-Fa-f]{6}$/).optional().describe("十六进制颜色，如 C00000"),
    highlight: z.enum(["yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "lightGray", "darkGray", "black"]).optional(),
    size_pt: z.number().min(1).max(400).optional(),
    font: z.string().optional(),
    vertical: z.enum(["superscript", "subscript", "baseline"]).optional(),
    small_caps: z.boolean().optional(),
    all_caps: z.boolean().optional(),
    char_style: z.string().optional().describe("应用文档中已有的字符样式（名称或 ID）"),
    reason,
  }),
  description: () => `为段落中的一段文字设置字符格式（加粗、斜体、下划线、删除线、颜色、高亮、字号、字体、上下标、大小写、字符样式），不改变文字内容。
只修改指定的属性，其它格式保持不变。优先使用文档已有的字符样式（char_style），与全文保持一致。
开启修订模式时，格式修改会记录为可接受/拒绝的格式修订。`,
  apply: (doc, i, ctx) => {
    const { ref, text, reason: _r, ...format } = i;
    return doc.formatText({ ref, text, format }, ctx.editOptions());
  },
});

export const DocSetParagraphTool = editTool({
  name: "doc_set_paragraph",
  capability: "paragraphProps",
  title: (i) => `调整段落格式（${(i.refs ?? []).length} 段）`,
  inputSchema: z.object({
    refs: z.array(z.string()).min(1),
    style: z.string().optional().describe("段落样式名称或 ID，必须是文档中已定义的样式（先用 doc_styles 查看）"),
    alignment: z.enum(["left", "center", "right", "justify"]).optional(),
    space_before_pt: z.number().min(0).max(1000).optional(),
    space_after_pt: z.number().min(0).max(1000).optional(),
    line_spacing: z.number().min(0.5).max(5).optional().describe("行距倍数，如 1.5"),
    indent_left_pt: z.number().optional(),
    indent_right_pt: z.number().optional(),
    first_line_pt: z.number().optional().describe("首行缩进（磅），中文正文首行缩进两字约 21–24 磅"),
    hanging_pt: z.number().optional(),
    keep_with_next: z.boolean().optional(),
    page_break_before: z.boolean().optional(),
    reason,
  }),
  description: () => `调整一个或多个段落的段落级格式：段落样式、对齐、段前段后间距、行距、缩进、与下段同页、段前分页。
- 统一格式时优先"套用文档已有的样式"（style），而不是逐项设置直接格式；这样与模板定义保持一致。
- 只修改指定的属性，未指定的保持原样。
- 批量调整前先用 doc_inspect 看一个正确段落的格式作为参照。`,
  apply: (doc, i, ctx) => {
    const { refs, reason: _r, ...format } = i;
    return doc.setParagraph({ refs, format }, ctx.editOptions());
  },
});

export const DocInsertBlocksTool = editTool({
  name: "doc_insert_blocks",
  capability: "insertBlocks",
  title: (i) => `插入 ${(i.blocks ?? []).length} 个段落`,
  inputSchema: z.object({
    anchor: z.string().describe("锚点段落引用"),
    position: z.enum(["before", "after"]),
    blocks: z.array(z.object({
      text: z.string().describe("段落文字；\\n 为段内换行"),
      style: z.string().optional().describe("段落样式（名称或 ID）；省略则与锚点段落相同"),
      like: z.string().optional().describe("以哪个已有段落为格式模板"),
    })).min(1).max(50),
    reason,
  }),
  description: () => `在锚点段落之前或之后插入新段落（新章节、补充说明、列表项、标题等）。

格式来源（保证新内容与文档浑然一体）：
- 默认完全复制锚点段落的段落格式和主体文字格式；
- 指定 style 时，自动寻找文档中离锚点最近的同样式段落作为模板，新段落与已有同类段落外观完全一致（例如新标题与已有标题一致、自动编号延续）；
- 也可以用 like 显式指定模板段落。
插入后返回新段落的引用。文档结构发生变化：P12@v3 这类临时引用会失效，#开头的引用仍有效。
不要用它来"替换"已有段落——修改已有文字请用 doc_replace_text。`,
  apply: (doc, i, ctx) => doc.insertBlocks({ anchor: i.anchor, position: i.position, blocks: i.blocks }, ctx.editOptions()),
});

export const DocDeleteBlocksTool = editTool({
  name: "doc_delete_blocks",
  capability: "deleteBlocks",
  title: (i) => `删除 ${(i.refs ?? []).length} 个段落`,
  inputSchema: z.object({ refs: z.array(z.string()).min(1).max(200), reason }),
  destructive: () => true,
  checkPermissions: async (i) =>
    i.refs.length > 10 ? { behavior: "ask", reason: `将一次删除 ${i.refs.length} 个段落，请确认。` } : { behavior: "passthrough" },
  description: () => `删除整个段落。这是破坏性操作，总是需要用户确认（除非用户明确授权）。
- 只删除用户明确要求删除的内容；如果只是要删掉段落中的几个词，用 doc_replace_text。
- 携带分节符的段落、表格单元格中唯一的段落受保护，不能删除（可以清空其文字）。
- 开启修订模式时，删除记录为修订，用户可以在 Word 中拒绝以恢复。`,
  apply: (doc, i, ctx) => doc.deleteBlocks({ refs: i.refs }, ctx.editOptions()),
});

export const DocInsertTableRowTool = editTool({
  name: "doc_insert_table_row",
  capability: "tables",
  title: () => "插入表格行",
  inputSchema: z.object({
    ref: z.string().describe("表格中某一行里任意单元格段落的引用，新行将复制这一行的结构与格式"),
    position: z.enum(["before", "after"]),
    cells: z.array(z.string()).min(1).describe("新行每个单元格的文字，数量必须等于该行单元格数"),
    reason,
  }),
  description: () => `在表格中插入一行：复制参照行的单元格宽度、边框、底纹、字体格式，填入新的文字。
先用 doc_read（part="表格1"）查看表格结构，确认列数和参照行。`,
  apply: (doc, i, ctx) => doc.insertTableRow({ ref: i.ref, position: i.position, cells: i.cells }, ctx.editOptions()),
});

export const DocAddCommentTool = editTool({
  name: "doc_add_comment",
  capability: "comments",
  title: (i) => `添加批注（${i.ref ?? ""}）`,
  inputSchema: z.object({
    ref: z.string(),
    text: z.string().optional().describe("批注锚定的文字（须在段落内唯一）；省略则批注整段"),
    comment: z.string().min(1),
    reason,
  }),
  description: () => `在段落（或段落中的一段文字）上添加批注，不修改正文。
适用于：审阅意见、需要作者确认的问题、不确定是否应该修改的地方。用户明确只要"提意见"而不要直接改时，用批注代替修改。`,
  apply: (doc, i, ctx) => doc.addComment({ ref: i.ref, text: i.text, comment: i.comment }, ctx.editOptions()),
});

export const DocExportTool = buildTool({
  name: "doc_export",
  category: "export",
  description: () => `把当前文档导出为文件供用户下载：docx / pdf / html / markdown。
- 导出与原文件相同的格式 = 编辑后的原文件本身，未修改部分逐字节不变（推荐）；
- Word 导出 PDF 由 LibreOffice 按 Word 版式渲染；
- 跨格式导出（如 Word → Markdown）只能近似保留结构，会在结果中注明。
导出不修改文档本身。`,
  inputSchema: z.object({ format: z.enum(["docx", "pdf", "html", "markdown"]), filename: z.string().optional() }),
  userFacingName: (i) => `导出 ${String(i.format ?? "").toUpperCase()}`,
  isEnabled: (ctx) => ctx.session.meta.format !== "pdf",
  permissionContent: (i) => i.format,
  checkPermissions: async () => ({ behavior: "allow" }),
  async call(i, ctx) {
    const out = await exportDocument(ctx.session.current(), ctx.session.meta.format, i.format);
    const base = (i.filename || ctx.session.meta.filename).replace(/\.[^.]+$/, "").replace(/[\\/:*?"<>|]/g, "_");
    const name = `${base}-v${ctx.session.meta.currentVersion}.${out.ext}`;
    const file = await ctx.session.writeExport(name, out.data);
    return {
      content: `已导出 ${name}（${Math.round(out.data.length / 1024)} KB）。${out.note}`,
      data: { name, file, url: `/api/sessions/${ctx.session.id}/exports/${encodeURIComponent(name)}`, lossy: out.lossy },
    };
  },
});

export const DOC_TOOLS: Tool[] = [
  DocOutlineTool, DocReadTool, DocSearchTool, DocInspectTool, DocStylesTool, DocVerifyTool,
  DocReplaceTextTool, DocFormatTextTool, DocSetParagraphTool, DocInsertBlocksTool, DocDeleteBlocksTool,
  DocInsertTableRowTool, DocAddCommentTool, DocExportTool,
] as Tool[];
