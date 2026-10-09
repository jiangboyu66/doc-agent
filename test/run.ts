/**
 * 自动化测试（npm test）
 *
 *  1. 引擎保真：字节级往返、最小修改、格式继承、修订/批注、引用过期检测（docx / markdown / html）
 *  2. 代理主循环：用"脚本化的模拟模型"跑完整的 QueryEngine——并行只读工具、权限确认（允许 / 记住 / 拒绝）、
 *     plan 模式拦截、先读后改、并行子代理、保真守卫、消息配对修复、压缩。
 *
 * 不调用真实的 DeepSeek API，不需要 API Key。设置 DOC_AGENT_TEST_DOCX=某个.docx 可额外对真实文档做往返与编辑测试。
 */

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

// 数据目录、用户目录必须在导入任何模块之前指向临时目录（PATHS 在模块加载时确定）
const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "doc-agent-test-"));
process.env.DOC_AGENT_DATA_DIR = path.join(tmp, "data");
process.env.DOC_AGENT_USER_DIR = path.join(tmp, "user");

const here = path.dirname(fileURLToPath(import.meta.url));
const fixture = (n: string) => fs.readFile(path.join(here, "fixtures", n));

const { DocxDocument } = await import("../src/documents/docx/DocxDocument.js");
const { MarkdownDocument } = await import("../src/documents/markdown/MarkdownDocument.js");
const { HtmlDocument } = await import("../src/documents/html/HtmlDocument.js");
const { ZipPackage } = await import("../src/documents/docx/zip.js");
const { computeHunks } = await import("../src/documents/textMatch.js");
const { Session } = await import("../src/session/Session.js");
const { QueryEngine } = await import("../src/QueryEngine.js");
const { bootstrap } = await import("../src/bootstrap.js");
const { normalizeMessagesForAPI } = await import("../src/services/api/normalize.js");
const { microCompact } = await import("../src/services/compact.js");
const { processUserInput } = await import("../src/commands.js");
type ApiMessage = import("../src/services/api/deepseek.js").ApiMessage;
type ModelClient = import("../src/services/api/deepseek.js").ModelClient;
type CompletionResult = import("../src/services/api/deepseek.js").CompletionResult;
type EngineEvent = import("../src/bridge/protocol.js").EngineEvent;
type PermissionRequest = import("../src/bridge/protocol.js").PermissionRequest;
type PermissionResponse = import("../src/bridge/protocol.js").PermissionResponse;

// ---------------------------------------------------------------------------
// 迷你测试框架
// ---------------------------------------------------------------------------

const results: Array<{ name: string; ok: boolean; err?: string }> = [];
async function test(name: string, fn: () => unknown | Promise<unknown>) {
  try {
    await fn();
    results.push({ name, ok: true });
    console.log(`  \x1b[32m✓\x1b[0m ${name}`);
  } catch (e: any) {
    results.push({ name, ok: false, err: e?.stack ?? String(e) });
    console.log(`  \x1b[31m✗ ${name}\x1b[0m\n    ${String(e?.message ?? e).split("\n").join("\n    ")}`);
  }
}
function assert(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}
function eq<T>(a: T, b: T, msg: string) {
  if (a !== b) throw new Error(`${msg}\n      期望：${JSON.stringify(b)}\n      实际：${JSON.stringify(a)}`);
}
const count = (s: string, sub: string | RegExp) => (typeof sub === "string" ? s.split(sub).length - 1 : (s.match(new RegExp(sub, "g")) ?? []).length);
const docText = (buf: Buffer) => new DocxDocument(buf).listBlocks().map((b) => b.text).join("\n");
const docXml = (buf: Buffer) => new ZipPackage(buf).readText("word/document.xml");
const O = (track = false) => ({ track, author: "测试 作者", date: "2026-09-29T00:00:00Z" });

// ---------------------------------------------------------------------------
console.log("\n文本差分");
// ---------------------------------------------------------------------------

await test("最小差分只覆盖真正变化的字符", () => {
  const hunks = computeHunks("First A. Author1, Fellow", "Zhang San1, Fellow");
  assert(hunks.length >= 1, "应产生差分");
  const old = "First A. Author1, Fellow";
  for (const h of hunks) assert(!old.slice(h.start, h.end).includes("1,"), `上标数字不应落入修改区：${JSON.stringify(h)}`);
});

await test("差分位置按旧文本计算（中英混排）", () => {
  const a = "Another paragraph with a link and much more.";
  const b = "Another paragraph with a link and so much more.";
  const hunks = computeHunks(a, b);
  let out = a;
  for (const h of [...hunks].sort((x, y) => y.start - x.start)) out = out.slice(0, h.start) + h.text + out.slice(h.end);
  eq(out, b, "应用差分后应得到新文本");
});

// ---------------------------------------------------------------------------
console.log("\nWord（.docx）");
// ---------------------------------------------------------------------------

const sampleDocx = await fixture("sample.docx");

await test("未修改时字节级完全一致", () => {
  const d = new DocxDocument(sampleDocx);
  assert(Buffer.compare(d.serialize(), sampleDocx) === 0, "序列化结果与原文件不是字节一致");
});

await test("替换文字：只改动一个段落，其余部件字节一致，加粗 run 保留", () => {
  const d = new DocxDocument(sampleDocx);
  const r = d.replaceText({ oldText: "智能代理", newText: "文档代理" }, O());
  eq(r.changedRefs.length, 1, "应只修改一个段落");
  const rep = d.fidelity(sampleDocx);
  assert(rep.ok, `保真检查失败：${rep.problems.join("；")}`);
  eq(rep.modifiedParts.length, 1, "应只有 document.xml 被修改");
  eq(rep.modifiedParts[0].changedBlocks.length, 1, "document.xml 中应只有一个段落变化");
  const xml = docXml(d.serialize());
  // 最小差分只改"智能"→"文档"，"代理"所在的 run 原样保留
  assert(xml.includes("<w:t>文档</w:t>") && !xml.includes("智能"), "文字未替换");
  const block = d.listBlocks().find((b) => b.text.includes("文档代理"))!;
  assert(d.describeFormat(block.ref).includes("加粗"), "同段落中的加粗 run 应保持加粗");
});

await test("直引号可以匹配文档中的弯引号", () => {
  const d = new DocxDocument(sampleDocx);
  d.replaceText({ oldText: "It's a well-known pangram.", newText: "It's a famous pangram." }, O());
  assert(d.listBlocks().some((b) => b.text.includes("famous pangram")), "替换失败");
});

await test("修改相邻文字时上标 run 原样保留", () => {
  const d = new DocxDocument(sampleDocx);
  const before = count(docXml(sampleDocx), 'w:val="superscript"');
  d.replaceText({ oldText: "后面是普通文字", newText: "后面是正文文字" }, O());
  const xml = docXml(d.serialize());
  eq(count(xml, 'w:val="superscript"'), before, "上标数量变化");
  assert(/<w:vertAlign w:val="superscript"\/><\/w:rPr><w:t>2<\/w:t>/.test(xml), "上标 2 丢失");
  assert(/<w:color w:val="C00000"\/>[\s\S]*?<w:t>红色强调文字<\/w:t>/.test(xml), "红色格式丢失");
});

await test("修订模式生成 w:ins / w:del，作者含空格也正确", () => {
  const d = new DocxDocument(sampleDocx);
  d.replaceText({ oldText: "最小化修改", newText: "最小化、可追溯的修改" }, O(true));
  const xml = docXml(d.serialize());
  assert(count(xml, "<w:ins ") >= 1, "缺少 w:ins");
  assert(/w:author="测试 作者"/.test(xml), "作者名错误");
  assert(d.fidelity(sampleDocx).ok, "修订模式下保真检查失败");
});

await test("修订模式下删除文字用 w:delText", () => {
  const d = new DocxDocument(sampleDocx);
  d.replaceText({ oldText: "第一步：读取文档。", newText: "第一步：读取。" }, O(true));
  const xml = docXml(d.serialize());
  assert(/<w:del [^>]*>[\s\S]*?<w:delText[^>]*>文档<\/w:delText>/.test(xml), "删除的文字应在 w:delText 中");
});

await test("文字格式 / 段落格式（修订模式带 rPrChange / pPrChange）", () => {
  const d = new DocxDocument(sampleDocx);
  const ref = d.search("结论段落", {})[0].ref;
  d.formatText({ ref, text: "结论", format: { bold: true, color: "1F4E79" } }, O(true));
  d.setParagraph({ refs: [ref], format: { alignment: "justify", space_after_pt: 6 } }, O(true));
  const xml = docXml(d.serialize());
  assert(xml.includes("w:rPrChange") && xml.includes("w:pPrChange"), "缺少格式修订记录");
  assert(xml.includes('<w:jc w:val="both"/>'), "对齐未生效");
});

await test("插入段落沿用同样式段落格式；插入后旧的序号引用仍指向原段落，段落被删后报过期", () => {
  const d = new DocxDocument(sampleDocx);
  const anchor = d.search("第一步", {})[0].ref;
  const oldRef = d.search("结论段落", {})[0].ref;
  const r = d.insertBlocks({ anchor, position: "after", blocks: [{ text: "第一步半：规划修改。" }, { text: "3. 结果", style: "Heading 1" }] }, O());
  eq(r.changedRefs.length, 2, "应插入两个段落");
  const blocks = d.listBlocks();
  const inserted = blocks.find((b) => b.text === "第一步半：规划修改。")!;
  eq(inserted.style, blocks.find((b) => b.text.startsWith("第一步："))!.style, "新段落应沿用列表样式");
  eq(blocks.find((b) => b.text === "3. 结果")!.kind, "heading", "新标题应识别为标题");
  if (/^P\d+@v\d+$/.test(oldRef)) {
    assert(d.getBlock(oldRef).text.startsWith("结论段落"), "插入后旧引用应仍指向原来的段落");
    d.deleteBlocks({ refs: [d.search("结论段落", {})[0].ref] }, O());
    let code = "";
    try { d.getBlock(oldRef); } catch (e: any) { code = e.code; }
    eq(code, "stale_ref", "段落删除后旧引用应报过期");
  }
});

await test("删除段落 / 表格插行 / 批注", () => {
  const d = new DocxDocument(sampleDocx);
  d.deleteBlocks({ refs: [d.search("The quick brown fox", {})[0].ref] }, O());
  const cell = d.search("速度", {})[0];
  d.insertTableRow({ ref: cell.ref, position: "after", cells: ["a", "加速度", "m/s²"] }, O(true));
  d.addComment({ ref: d.search("深度学习", {})[0].ref, text: "深度学习", comment: "请补充参考文献" }, O());
  const out = d.serialize();
  const pkg = new ZipPackage(out);
  assert(pkg.readText("word/comments.xml").includes("请补充参考文献"), "批注内容缺失");
  assert(pkg.readText("[Content_Types].xml").includes("comments+xml"), "缺少批注的内容类型");
  assert(pkg.readText("word/_rels/document.xml.rels").includes("comments.xml"), "缺少批注关系");
  const xml = docXml(out);
  assert(!xml.includes("The quick brown fox"), "段落未删除");
  assert(xml.includes("加速度"), "表格行未插入");
  const rep = d.fidelity(sampleDocx);
  assert(rep.ok, rep.problems.join("；"));
  // 重新加载能正常解析
  const again = new DocxDocument(out);
  assert(again.listBlocks().some((b) => b.text === "加速度"), "重新加载后找不到新增的行");
});

await test("找不到 / 多处匹配时给出可操作的错误", () => {
  const d = new DocxDocument(sampleDocx);
  let e1: any, e2: any;
  try { d.replaceText({ oldText: "不存在的句子", newText: "x" }, O()); } catch (e) { e1 = e; }
  try { d.replaceText({ oldText: "第", newText: "x" }, O()); } catch (e) { e2 = e; }
  eq(e1?.code, "not_found", "应报 not_found");
  eq(e2?.code, "ambiguous", "应报 ambiguous");
});

if (process.env.DOC_AGENT_TEST_DOCX) {
  await test(`真实文档往返：${path.basename(process.env.DOC_AGENT_TEST_DOCX)}`, async () => {
    const buf = await fs.readFile(process.env.DOC_AGENT_TEST_DOCX!);
    const d = new DocxDocument(buf);
    assert(Buffer.compare(d.serialize(), buf) === 0, "往返不一致");
    const b = d.listBlocks().find((x) => x.text.length > 30 && !x.text.includes("⟨"))!;
    const word = b.text.slice(5, 15);
    d.replaceText({ ref: b.ref, oldText: word, newText: word + "（改）" }, O(true));
    const rep = d.fidelity(buf);
    assert(rep.ok, rep.problems.join("；"));
    console.log(`      ${rep.identicalParts}/${rep.totalParts} 个部件字节一致`);
  });
}

// ---------------------------------------------------------------------------
console.log("\nMarkdown / HTML");
// ---------------------------------------------------------------------------

await test("Markdown：往返一致，修改保留行内标记与 CRLF 换行", async () => {
  const src = (await fixture("sample.md")).toString("utf8").replace(/\r?\n/g, "\r\n");
  const d = new MarkdownDocument(Buffer.from(src));
  eq(d.serialize().toString("utf8"), src, "往返不一致");
  d.replaceText({ oldText: "第一段", newText: "首段" }, O());
  const out = d.serialize().toString("utf8");
  assert(out.includes("**首段**"), "加粗标记丢失");
  assert(out.includes("[链接](https://example.com)"), "链接被破坏");
  eq(count(out, "\r\n"), count(src, "\r\n"), "换行符数量变化");
  assert(!/[^\r]\n/.test(out), "混入了 LF 换行");
});

await test("HTML：实体、属性与未修改部分原样保留", async () => {
  const src = (await fixture("sample.html")).toString("utf8");
  const d = new HtmlDocument(Buffer.from(src));
  eq(d.serialize().toString("utf8"), src, "往返不一致");
  d.replaceText({ oldText: "and much more", newText: "and so much more" }, O());
  d.replaceText({ oldText: "副标题", newText: "小标题" }, O());
  const out = d.serialize().toString("utf8");
  assert(out.includes("and so much more"), `替换结果错误：${out.match(/Another.*<\/p>/)?.[0]}`);
  assert(out.includes("标题 &amp; 小标题"), "实体被破坏");
  assert(out.includes('<p class="note">这是 <b>加粗</b> 的段落，包含&nbsp;实体。</p>'), "未修改段落发生变化");
  assert(out.includes("<style>p.note { color: #555; }</style>"), "样式表发生变化");
});

// ---------------------------------------------------------------------------
console.log("\nPDF → Word");
// ---------------------------------------------------------------------------

const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
const { parseFontName } = await import("../src/services/pdfConvert/extract.js");

await test("字体名解析：子集前缀、粗斜体、替代字体、GBK 编码中文名", () => {
  const a = parseFontName("ABCDEF+TimesNewRomanPS-BoldItalicMT");
  eq(a.family, "Times New Roman", "字体族");
  assert(a.bold && a.italic, "应识别粗斜体");
  eq(parseFontName("BAAAAA+LiberationSerif").family, "Times New Roman", "度量兼容字体应映射回原字体");
  eq(parseFontName("CAAAAA+Carlito-Bold").family, "Calibri", "Carlito → Calibri");
  eq(parseFontName("#CB#CE#CC#E5").family, "宋体", "GBK 编码字体名应解码");
});

for (const [name, expect] of [["form.pdf", { tables: 1, columns: false }], ["columns.pdf", { tables: 0, columns: true }], ["sample.pdf", { tables: 1, columns: false }]] as const) {
  for (const mode of ["exact", "flow"] as const) {
    await test(`${name}（${mode === "exact" ? "保留原排版" : "便于改写"}）：文字 100% 保留，结构还原，可被编辑引擎打开`, async () => {
      const { data, report } = await pdfToDocx(await fixture(name), { mode });
      eq(report.textCoverage, 1, `文字完整性，缺失：${report.missing}`);
      eq(report.tables, expect.tables, "表格数");
      eq(report.columnPages > 0, expect.columns, "分栏检测");
      const d = new DocxDocument(data);
      assert(Buffer.compare(d.serialize(), data) === 0, "生成的 docx 往返不一致");
      const rep = d.fidelity(data);
      assert(rep.ok, rep.problems.join("；"));
      // 能在转换结果上做修订模式的编辑
      const b = d.listBlocks().find((x) => x.text.replace(/\s/g, "").length > 12 && !x.text.includes("⟨"))!;
      const word = b.text.split(/\s+/).find((w) => w.length >= 4)!;
      d.replaceText({ ref: b.ref, oldText: word, newText: word + "X" }, O(true));
      assert(d.fidelity(data).ok, "编辑后结构校验失败");
    });
  }
}

await test("LaTeX 论文：矢量图与行间公式按原样渲染为图片贴回原位，文字计入替代文字；渲染不可用时退回文字重建", async () => {
  const pdf = await fixture("math.pdf");
  const { data, report } = await pdfToDocx(pdf);
  eq(report.figures, 1, "矢量图数量");
  eq(report.equations, 1, "行间公式数量");
  eq(report.textCoverage, 1, `文字完整性，缺失：${report.missing}`);
  const { ZipPackage } = await import("../src/documents/docx/zip.js");
  const zp = new ZipPackage(data);
  const xml = zp.readText("word/document.xml");
  assert(/descr="[^"]*Encoder[^"]*"/.test(xml), "图片替代文字应包含图中文字");
  assert(!/>Encoder</.test(xml), "图中文字不应再以正文形式出现");
  const pngs = zp.names().filter((k) => k.startsWith("word/media/"));
  assert(pngs.length >= 2, "应生成渲染图片");
  assert(new DocxDocument(data).fidelity(data).ok, "结构校验失败");
  // 关闭渲染：按文字重建，文字仍然完整
  process.env.PDF_RASTER = "0";
  try {
    const r2 = await pdfToDocx(pdf);
    eq(r2.report.figures + r2.report.equations, 0, "关闭渲染后不应有渲染区域");
    assert(r2.report.textCoverage > 0.99, `回退后文字完整性 ${r2.report.textCoverage}`);
  } finally {
    delete process.env.PDF_RASTER;
  }
});

await test("公式编号与公式在同一行：渲染为一张图（编号在图内）；不渲染时右端编号用右对齐制表位", async () => {
  {
    const { data } = await pdfToDocx(await fixture("equation-number.pdf"));
    const { ZipPackage: ZP } = await import("../src/documents/docx/zip.js");
    const xml = new ZP(data).readText("word/document.xml");
    assert(/<wp:inline\b(?:(?!<\/wp:inline>).)*descr="[^"]*\(1\)/.test(xml), "公式与编号应渲染为同一张行内图片");
  }
  process.env.PDF_RASTER = "0";
  const { data, report } = await pdfToDocx(await fixture("equation-number.pdf"));
  delete process.env.PDF_RASTER;
  eq(report.textCoverage, 1, `文字完整性，缺失：${report.missing}`);
  const { ZipPackage } = await import("../src/documents/docx/zip.js");
  const xml = new ZipPackage(data).readText("word/document.xml");
  const para = xml.match(/<w:p [^>]*>(?:(?!<\/w:p>).)*mc(?:(?!<\/w:p>).)*<\/w:p>/)?.[0] ?? "";
  assert(para.includes(">(1)<"), "编号 (1) 应与公式在同一段落");
  assert(/<w:tab w:val="right" w:pos="\d+"\/>/.test(para), "编号应使用右对齐制表位");
  assert(!/<w:ind [^>]*w:left="[1-9]/.test(para), "单行公式段落不应使用左缩进（预览器会把制表位算错）");
});

await test("可编辑（流式）转换：不逐页分页、不插分栏符，公式为行内图片，扩写后文档仍合法；版式模式可识别", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const cols = await pdfToDocx(await fixture("columns.pdf"), { mode: "flow" });
  eq(cols.report.textCoverage, 1, `文字完整性，缺失：${cols.report.missing}`);
  const pkg = new ZipPackage(cols.data);
  const xml = pkg.readText("word/document.xml");
  assert(!xml.includes('w:val="PageStart"') && !xml.includes("<w:pageBreakBefore/>"), "流式模式不应逐页强制分页");
  assert(!xml.includes('w:type="column"'), "流式模式不应插入分栏符（文字跨栏连续流动）");
  assert(!/<w:br w:clear="none"\/>/.test(xml), "流式模式段落内不应有排版换行");
  assert(xml.includes("<w:lastRenderedPageBreak/>"), "应记录原 PDF 的分页位置（浏览器预览据此分页）");
  assert(pkg.readText("word/settings.xml").includes("<w:autoHyphenation/>"), "流式模式开启自动断词");
  const d = new DocxDocument(cols.data);
  eq(d.convertedLayout(), "flow", "版式模式");
  // 扩写一段：文档仍可解析，段落数不变
  const p = d.listBlocks().filter((b) => b.text.length > 200)[0];
  d.replaceText({ ref: p.ref, oldText: p.text, newText: p.text + " 扩写的内容。".repeat(30) }, { track: false, author: "t", date: new Date().toISOString() } as any);
  const d2 = new DocxDocument(d.serialize());
  eq(d2.listBlocks().length, d.listBlocks().length, "扩写后段落数");

  // 行间公式：两种模式都是行内图片（随文字流移动，不会与文字重叠）
  for (const mode of ["flow", "exact"] as const) {
    const m = await pdfToDocx(await fixture("math.pdf"), { mode });
    const mx = new ZipPackage(m.data).readText("word/document.xml");
    assert(/<wp:inline\b/.test(mx), `${mode}：公式应为行内图片`);
    assert(!/<wp:anchor\b(?:(?!<\/wp:anchor>).)*descr=/.test(mx), `${mode}：公式/图不应按页面坐标浮动`);
  }
  const ex = await pdfToDocx(await fixture("columns.pdf"), { mode: "exact" });
  eq(new DocxDocument(ex.data).convertedLayout(), "exact", "逐页一致模式可识别");
  assert(new ZipPackage(ex.data).readText("word/document.xml").includes('<w:br w:clear="none"/>'), "逐页一致模式的排版换行带标记");
});

await test("曲线图的网格线不当作表格、组图整体成图、图题末行不被切进分栏（两种模式）", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  for (const mode of ["flow", "exact"] as const) {
    const { data, report } = await pdfToDocx(await fixture("figures.pdf"), { mode });
    eq(report.textCoverage, 1, `${mode} 文字完整性，缺失：${report.missing}`);
    eq(report.tables, 1, `${mode}：只有一个真正的表格（曲线图网格不算表格）`);
    eq(report.figures, 2, `${mode}：七联曲线图合成一幅、2×2 组图合成一幅`);
    const d = new DocxDocument(data);
    const caps = d.listBlocks().filter((b) => /^Figure \d/.test(b.text.trim()));
    eq(caps.length, 2, `${mode}：两个图题`);
    assert(caps.some((b) => b.text.includes("decreases smoothly")), `${mode}：图题末行应与图题在同一段`);
    assert(!d.listBlocks().some((b) => /^(Total loss|Charbonnier|\(a\))$/.test(b.text.trim())), `${mode}：图中文字不应散落为正文`);
  }
});

await test("学术期刊排版：通栏图表单独成节、不跨页、网格对齐、不平衡分栏；可反复运行", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const { applyJournalLayout } = await import("../src/services/journalLayout.js");
  const { data } = await pdfToDocx(await fixture("journal.pdf"), { mode: "flow" });
  const r1 = await applyJournalLayout(data, {});
  eq(r1.report.floats.length, 3, "两幅图 + 一张表");
  assert(r1.report.floats.some((f) => f.wide && f.kind === "figure"), "七联曲线图应为通栏");
  assert(r1.report.floats.some((f) => !f.wide && f.kind === "table"), "表格为栏内浮动体");
  assert(r1.report.pitch !== null && r1.report.gridAdjusted > 0, "基线网格");
  const pkg = new ZipPackage(r1.data);
  const xml = pkg.readText("word/document.xml");
  assert(pkg.readText("word/settings.xml").includes("<w:noColumnBalance/>"), "不平衡分栏");
  assert(xml.includes("<w:cantSplit/>") && xml.includes("<w:keepNext/>"), "表格行不跨页、与题注同页");
  // 通栏图在单栏节里：图片段落之后的第一个分节符是单栏
  const figAt = xml.indexOf("<wp:inline");
  const nextSect = xml.slice(figAt).match(/<w:sectPr>.*?<\/w:sectPr>/)?.[0] ?? "";
  assert(!/w:num="2"/.test(nextSect), "通栏图所在的节应为单栏");
  const d1 = new DocxDocument(r1.data);
  const text1 = d1.listBlocks().map((b) => b.text).join("\n").replace(/\s+/g, " ");
  // 再运行一次：还原后重新排，文字不变、书签不重复累积
  const r2 = await applyJournalLayout(r1.data, {});
  eq(r2.report.floats.length, 3, "再次运行仍识别全部图表");
  const d2 = new DocxDocument(r2.data);
  eq(d2.listBlocks().map((b) => b.text).join("\n").replace(/\s+/g, " "), text1, "再次运行文字不变");
  const marks = (new ZipPackage(r2.data).readText("word/document.xml").match(/_JLanchor/g) ?? []).length;
  eq(marks, 3, "引用标记不累积");
  // 工具：会话内运行后文档可继续读取
  process.env.DOC_AGENT_NO_OFFICE = "1";
  const s = await Session.create(data, "论文.docx", { mode: "bypassPermissions", trackChanges: false, author: "测试" });
  const { findTool } = await import("../src/tools.js");
  const tool = findTool("doc_journal_layout")!;
  const res = await tool.call({} as any, { session: s } as any);
  assert(String(res.content).includes("学术期刊排版完成") && (res as any).docChange?.structural, `工具输出异常：${res.content}`);
  assert(s.doc.listBlocks().length > 10, "排版后文档可继续读取");
  delete process.env.DOC_AGENT_NO_OFFICE;
  // 有 LibreOffice 时：按分页结果定位，通栏图在页顶，中间页没有大面积留白
  const { checkExternalTools, toPdfViaOffice } = await import("../src/services/convert.js");
  if ((await checkExternalTools()).soffice) {
    const r3 = await applyJournalLayout(data, { render: (b) => toPdfViaOffice(b, "docx") });
    assert(r3.report.floats.some((f) => f.wide && /页顶部/.test(f.placement)), `通栏图应放到页顶：${JSON.stringify(r3.report.floats)}`);
    eq(r3.report.whitespacePages.join(","), "", "中间页不应有大面积留白");
  }
});

await test("文字被转成矢量轮廓的表格整体渲染为图片（内容不丢失）；重新转换时保留之前的润色修改", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const { data, report } = await pdfToDocx(await fixture("outlined-table.pdf"), { mode: "flow" });
  assert(report.figures >= 1, `轮廓文字表格应渲染为图片：${JSON.stringify(report)}`);
  const d = new DocxDocument(data);
  const blocks = d.listBlocks();
  const capIdx = blocks.findIndex((b) => /Table 1/.test(b.text));
  assert(capIdx >= 0 && blocks.slice(capIdx + 1, capIdx + 3).some((b) => b.text.includes("⟨图片⟩")), "表题之后应是表格图片");

  // 重新转换保留修改：会话内转换 → 润色一段 → 从原 PDF 重新转换
  const s = await Session.create(await fixture("columns.pdf"), "论文.pdf", { mode: "bypassPermissions", trackChanges: false, author: "测试" });
  const { convertPdfInPlace } = await import("../src/session/convertPdf.js");
  const r1 = await convertPdfInPlace(s, { mode: "flow" });
  await s.commit(r1.label, []);
  const target = s.doc.listBlocks().find((b) => b.text.length > 120)!;
  const words = target.text.split(" ");
  const oldFrag = words.slice(2, 6).join(" ");
  s.doc.replaceText({ ref: target.ref, oldText: oldFrag, newText: oldFrag + " (polished wording)" }, { track: false, author: "t", date: new Date().toISOString() } as any);
  await s.commit("润色", [target.ref]);
  const r2 = await convertPdfInPlace(s, { mode: "flow" });
  await s.commit(r2.label, []);
  assert(r2.text.includes("段落修改搬到新文档"), `报告应说明搬运了修改：${r2.text}`);
  assert(s.doc.listBlocks().some((b) => b.text.includes("(polished wording)")), "重新转换后应保留润色修改");
});

await test("行尾连字符：排版断词去掉，复合词 / 缩写 / 多段复合词保留", async () => {
  const { collectWords, collectHyphenated, isSoftHyphen, wordCounts } = await import("../src/documents/hyphen.js");
  const text = "The sharpness of edges. A noise-dominated region. attention is gated by a gated unit; attention again. image-to-image";
  const words = collectWords(text), hy = collectHyphenated(text), counts = wordCounts(text);
  const soft = (a: string, b: string) => isSoftHyphen(a, b, words, hy, counts);
  assert(soft("sharp", "ness"), "sharp-ness 是断词");
  assert(soft("convolu", "tional"), "音节断词");
  assert(!soft("noise", "dominated"), "文中出现过带连字符的写法");
  assert(!soft("NDCT", "consistent"), "缩写 + 连字符");
  assert(!soft("attention", "gated"), "两半都是文中的独立单词");
  assert(!soft("image-to", "image"), "多段复合词的中间");
  assert(!soft("CNN", "Transformer"), "大写开头的专有名词复合");
});

await test("逐页一致方式转换的 Word：提示模型先换成可编辑排版；doc_convert_to_word 可从原 PDF 重新转换", async () => {
  const s = await Session.create(await fixture("columns.pdf"), "论文.pdf", { mode: "bypassPermissions", trackChanges: false, author: "测试" });
  const { convertPdfInPlace } = await import("../src/session/convertPdf.js");
  const r = await convertPdfInPlace(s, { mode: "exact" });
  await s.commit(r.label, []);
  const { buildTurnReminder } = await import("../src/context.js");
  assert(buildTurnReminder(s).includes("版式提示") && buildTurnReminder(s).includes("doc_convert_to_word(layout="), "应提示先转换为可编辑排版");
  const { findTool } = await import("../src/tools.js");
  const tool = findTool("doc_convert_to_word")!;
  assert(tool.isEnabled({ session: s } as any), "已转换的 Word 会话中应可重新转换");
  const res = await tool.call({} as any, { session: s } as any);
  assert(String(res.content).includes("100%"), `应报告文字完整性：${res.content}`);
  await s.commit((res as any).docChange.label, []);
  eq(s.doc.convertedLayout?.(), "flow", "默认重新转换为流式排版");
  assert(!buildTurnReminder(s).includes("版式提示"), "流式排版不再提示");
  eq(s.meta.filename, "论文.docx", "文件名");
  await s.rollback(0);
  eq(s.meta.format, "pdf", "仍可回滚到原 PDF");
});

await test("PDF 直接导出为 Word 文件（两种排版方式），导出命令与 Agent 导出工具在 PDF 会话中可用", async () => {
  const { exportDocument } = await import("../src/services/convert.js");
  const pdf = await fixture("form.pdf");
  for (const pdfMode of ["exact", "flow"] as const) {
    const out = await exportDocument(pdf, "pdf", "docx", { pdfMode });
    eq(out.ext, "docx", "扩展名");
    assert(out.note.includes("100%"), `应报告文字完整性：${out.note}`);
    assert(new DocxDocument(out.data).listBlocks().some((b) => b.text.includes("PENGGUNAAN")), "内容缺失");
  }
  const s = await Session.create(pdf, "声明表.pdf", { mode: "default", trackChanges: false, author: "测试" });
  const rt = await bootstrap({ permissionMode: "default" });
  const r = await processUserInput("/export docx", { session: s, runtime: rt, compact: async () => ({ before: 0, after: 0 }) });
  assert(r.kind === "local" && r.output.includes("声明表") && r.output.includes("100%"), `/export 输出异常：${(r as any).output}`);
  const { findTool } = await import("../src/tools.js");
  const tool = findTool("doc_export")!;
  assert(tool.isEnabled({ session: s } as any), "PDF 会话中应可使用导出工具");
  const res = await tool.call({ format: "docx", pdf_layout: "exact" } as any, { session: s } as any);
  const dname = (res.data as any)?.name;
  assert(dname === "声明表.docx", `导出文件名应为 声明表.docx，实际 ${dname}`);
});

await test("PDF 会话转换为新的 Word 会话，原 PDF 会话不变", async () => {
  const { convertPdfSession } = await import("../src/session/convertPdf.js");
  const src = await Session.create(await fixture("form.pdf"), "表单.pdf", { mode: "default", trackChanges: false, author: "测试" });
  const { session, report } = await convertPdfSession(src, { mode: "exact", defaults: { mode: "default", trackChanges: true, author: "测试" } });
  eq(session.meta.format, "docx", "新会话应为 Word");
  eq(session.meta.filename, "表单.docx", "文件名");
  eq(session.meta.origin?.fromSession, src.id, "记录来源会话");
  assert(report.textCoverage === 1, "文字完整");
  eq(src.meta.format, "pdf", "原会话不变");
  const again = await Session.load(session.id);
  assert(again.doc.listBlocks().some((b) => b.text.includes("PENGGUNAAN")), "重新加载后内容完整");
});

// ---------------------------------------------------------------------------
console.log("\n代理主循环（模拟模型）");
// ---------------------------------------------------------------------------

type Step = (req: { system: string; messages: ApiMessage[]; tools: Array<{ function: { name: string } }> }) => Partial<CompletionResult>;
let callSeq = 0;
const tc = (name: string, args: unknown) => ({ id: `call_${++callSeq}`, type: "function" as const, function: { name, arguments: JSON.stringify(args) } });

/** 脚本化模型：主代理按脚本逐步回答；子代理（工具列表里没有 agent）统一返回一段审校意见 */
class MockClient implements ModelClient {
  model = "deepseek-flash";
  thinking = true;
  requests: Array<{ messages: ApiMessage[]; tools: string[] }> = [];
  subActive = 0;
  subMaxConcurrent = 0;
  constructor(private steps: Step[]) {}
  async complete(req: any, opts: any): Promise<CompletionResult> {
    const tools = req.tools.map((t: any) => t.function.name);
    this.requests.push({ messages: structuredClone(req.messages), tools });
    const base = { content: "", reasoning: "（思考）", toolCalls: [], finishReason: "stop", usage: { promptTokens: 1000, cachedTokens: 800, completionTokens: 50 } };
    if (!tools.includes("agent") && !tools.includes("todo_write")) {
      // 子代理
      this.subActive++;
      this.subMaxConcurrent = Math.max(this.subMaxConcurrent, this.subActive);
      await new Promise((r) => setTimeout(r, 30));
      this.subActive--;
      return { ...base, content: "审校完成：未发现问题。" };
    }
    const step = this.steps.shift();
    if (!step) return { ...base, content: "完成。" };
    const r = step(req);
    if (r.content) opts.onChunk?.({ kind: "text", text: r.content });
    return { ...base, ...r } as CompletionResult;
  }
}

const runtime = await bootstrap({ permissionMode: "default" });

async function newSession(name = "sample.docx") {
  return Session.create(await fixture(name), name, { mode: "default", trackChanges: false, author: "测试" });
}

async function drive(session: Awaited<ReturnType<typeof Session.create>>, client: ModelClient, prompt: string, decide: (r: PermissionRequest) => PermissionResponse) {
  const events: EngineEvent[] = [];
  const asked: PermissionRequest[] = [];
  const engine = new QueryEngine(session, {
    client, settings: runtime.settings, skills: runtime.skills, agents: runtime.agents, memory: "",
    canUseTool: async (req) => { asked.push(req); return decide(req); },
  });
  for await (const e of engine.submitMessage(prompt, new AbortController().signal)) events.push(e);
  return { events, asked, engine };
}

/** 历史中每个 tool_call 都有且只有一个对应结果，且紧随其后 */
function assertPaired(history: ApiMessage[]) {
  for (let i = 0; i < history.length; i++) {
    const m = history[i];
    if (m.role !== "assistant" || !m.tool_calls?.length) continue;
    const ids = m.tool_calls.map((t) => t.id);
    const following = history.slice(i + 1, i + 1 + ids.length);
    for (const id of ids) assert(following.some((f) => f.role === "tool" && f.tool_call_id === id), `tool_call ${id} 缺少紧随的结果`);
  }
}

await test("完整流程：并行读取 → 预览确认 → 修改 → 并行子代理审校", async () => {
  const s = await newSession();
  let findRef = "";
  const client = new MockClient([
    () => ({ toolCalls: [tc("doc_outline", {}), tc("doc_search", { query: "智能代理" })] }),
    (req) => {
      const last = req.messages[req.messages.length - 1] as any;
      findRef = /(#[0-9A-F]{8}|P\d+@v\d+)/.exec(last.content)?.[1] ?? "";
      return { toolCalls: [tc("doc_replace_text", { ref: findRef, old_text: "智能代理", new_text: "文档代理", reason: "术语统一" })] };
    },
    () => ({ toolCalls: [tc("agent", { subagent_type: "reviewer", description: "审前半", prompt: "审校前半部分" }), tc("agent", { subagent_type: "reviewer", description: "审后半", prompt: "审校后半部分" })] }),
    () => ({ content: "已把“智能代理”改为“文档代理”，审校未发现问题。" }),
  ]);
  const { events, asked } = await drive(s, client, "把智能代理改成文档代理，然后审校", () => ({ decision: "allow", remember: "session" }));
  assert(findRef, "模型应从搜索结果拿到段落引用");
  eq(asked.length, 1, "应只请求一次确认");
  assert(asked[0].preview?.[0]?.after.includes("文档代理"), "确认请求应带有干跑预览");
  eq(s.meta.currentVersion, 1, "应提交一个新版本");
  assert(s.meta.sessionRules.some((r) => r.startsWith("doc_replace_text")), "“本次会话都允许”应记录会话规则");
  assert(docText(s.current()).includes("文档代理"), "文档未修改");
  assert(events.some((e) => e.type === "doc_changed"), "缺少 doc_changed 事件");
  eq(events.filter((e) => e.type === "subagent" && e.status === "start").length, 2, "应启动两个子代理");
  eq(client.subMaxConcurrent, 2, "两个只读子代理应并行执行");
  eq(events[events.length - 1].type, "done", "最后一个事件应是 done");
  assertPaired(s.history);
  // 带思考模式时每个 assistant 消息都回传了 reasoning_content
  for (const r of client.requests) for (const m of r.messages) if (m.role === "assistant") assert(typeof (m as any).reasoning_content === "string", "assistant 消息缺少 reasoning_content");
  // 系统提醒只出现在用户消息里，系统提示词保持稳定（利于前缀缓存）
  assert((s.history[0] as any).content.includes("<system-reminder>"), "用户消息应附带 system-reminder");
});

await test("会话规则生效：同类修改不再询问；/undo 回到上一版本", async () => {
  const s = await newSession();
  s.meta.sessionRules.push("doc_replace_text");
  const client = new MockClient([
    () => ({ toolCalls: [tc("doc_read", { offset: 0, limit: 50 })] }),
    () => ({ toolCalls: [tc("doc_replace_text", { old_text: "最小化修改", new_text: "精确修改" })] }),
    () => ({ content: "好了" }),
  ]);
  const { asked } = await drive(s, client, "改一下", () => ({ decision: "deny" }));
  eq(asked.length, 0, "已有会话规则时不应再询问");
  eq(s.meta.currentVersion, 1, "应生成版本 1");
  const out = await processUserInput("/undo", { session: s, runtime, compact: async () => ({ before: 0, after: 0 }) });
  eq(out.kind, "local", "/undo 应是本地命令");
  assert(!docText(s.current()).includes("精确修改") && docText(s.current()).includes("最小化修改"), "/undo 后文档应回到修改前");
});

await test("拒绝并反馈：文档不变，反馈回传给模型", async () => {
  const s = await newSession();
  const client = new MockClient([
    () => ({ toolCalls: [tc("doc_read", {})] }),
    () => ({ toolCalls: [tc("doc_replace_text", { old_text: "深度学习", new_text: "机器学习" })] }),
    () => ({ content: "明白，不改了。" }),
  ]);
  await drive(s, client, "改术语", () => ({ decision: "deny", feedback: "保留原术语" }));
  eq(s.meta.currentVersion, 0, "被拒绝时不应产生新版本");
  const toolMsgs = s.history.filter((m) => m.role === "tool").map((m) => (m as any).content as string);
  assert(toolMsgs.some((c) => c.includes("保留原术语")), "用户反馈应回传给模型");
  assertPaired(s.history);
});

await test("先读后改：未读取文档时写工具被拒绝", async () => {
  const s = await newSession();
  const client = new MockClient([
    () => ({ toolCalls: [tc("doc_replace_text", { old_text: "深度学习", new_text: "机器学习" })] }),
    () => ({ content: "好" }),
  ]);
  const { asked } = await drive(s, client, "直接改", () => ({ decision: "allow" }));
  eq(asked.length, 0, "未读取时不应走到确认环节");
  eq(s.meta.currentVersion, 0, "不应修改");
});

await test("plan 模式：写工具被拦截，exit_plan_mode 经确认后切换模式", async () => {
  const s = await newSession();
  s.meta.mode = "plan";
  const client = new MockClient([
    (req) => {
      assert(req.tools.some((t) => t.function.name === "exit_plan_mode"), "plan 模式下应提供 exit_plan_mode");
      return { toolCalls: [tc("doc_read", {}), tc("doc_replace_text", { old_text: "深度学习", new_text: "机器学习" })] };
    },
    () => ({ toolCalls: [tc("exit_plan_mode", { plan: "1. 把深度学习改为机器学习" })] }),
    () => ({ content: "计划已确认" }),
  ]);
  const { asked, events } = await drive(s, client, "先出计划", () => ({ decision: "allow" }));
  eq(s.meta.currentVersion, 0, "plan 模式下不应修改文档");
  eq(asked.length, 1, "只有 exit_plan_mode 需要确认");
  eq(asked[0].tool, "exit_plan_mode", "确认的应是计划");
  eq(s.meta.mode, "default", "确认计划后应退出 plan 模式");
  assert(events.some((e) => e.type === "mode_changed"), "缺少 mode_changed 事件");
});

await test("acceptEdits 模式：非破坏性修改自动放行，删除仍需确认", async () => {
  const s = await newSession();
  s.meta.mode = "acceptEdits";
  let ref = "";
  const client = new MockClient([
    () => ({ toolCalls: [tc("doc_search", { query: "pangram" })] }),
    (req) => {
      ref = /(#[0-9A-F]{8}|P\d+@v\d+)/.exec((req.messages[req.messages.length - 1] as any).content)![1];
      return { toolCalls: [tc("doc_replace_text", { old_text: "lazy dog", new_text: "sleepy dog" }), tc("doc_delete_blocks", { refs: [ref] })] };
    },
    () => ({ content: "ok" }),
  ]);
  const { asked } = await drive(s, client, "改", () => ({ decision: "deny" }));
  eq(asked.length, 1, "只有删除需要确认");
  eq(asked[0].tool, "doc_delete_blocks", "应确认删除操作");
  eq(s.meta.currentVersion, 1, "替换应自动生效");
});

await test("工具报错时文档回滚、错误信息回传模型继续", async () => {
  const s = await newSession();
  s.meta.mode = "bypassPermissions";
  const client = new MockClient([
    () => ({ toolCalls: [tc("doc_read", {})] }),
    () => ({ toolCalls: [tc("doc_replace_text", { old_text: "第", new_text: "X" }), tc("doc_replace_text", { old_text: "结论段落", new_text: "总结段落" })] }),
    () => ({ content: "ok" }),
  ]);
  await drive(s, client, "改", () => ({ decision: "allow" }));
  const toolMsgs = s.history.filter((m) => m.role === "tool").map((m) => (m as any).content as string);
  assert(toolMsgs.some((c) => c.startsWith("错误：") && c.includes("处")), "多处匹配的错误应回传");
  eq(s.meta.currentVersion, 1, "第二个修改应成功");
  const text = docText(s.current());
  assert(!text.includes("X") && text.includes("总结段落"), "失败的修改不应残留");
});

await test("干跑预览不会因为结构版本被误判为过期引用", async () => {
  const s = await newSession();
  s.meta.mode = "bypassPermissions";
  // 先插入段落让结构版本递增，再在 default 模式下用新引用修改（需要干跑预览）
  const client = new MockClient([
    () => ({ toolCalls: [tc("doc_read", {})] }),
    (req) => {
      const anchor = /(#[0-9A-F]{8}|P\d+@v\d+)\]?[^\n]*第一步/.exec((req.messages[req.messages.length - 1] as any).content)?.[1];
      return { toolCalls: [tc("doc_insert_blocks", { anchor, position: "after", blocks: [{ text: "新增一步。" }] })] };
    },
    () => { s.meta.mode = "default"; return { toolCalls: [tc("doc_search", { query: "结论段落" })] }; },
    (req) => {
      const ref = /(#[0-9A-F]{8}|P\d+@v\d+)/.exec((req.messages[req.messages.length - 1] as any).content)![1];
      return { toolCalls: [tc("doc_replace_text", { ref, old_text: "结论", new_text: "总结" })] };
    },
    () => ({ content: "ok" }),
  ]);
  const { asked } = await drive(s, client, "改", () => ({ decision: "allow" }));
  eq(s.doc.structureVersion >= 1, true, "插入后结构版本应递增");
  eq(asked.length, 1, "修改应请求确认（而不是在干跑时报错）");
  assert(asked[0].preview?.length, "应有预览");
  assert(docText(s.current()).includes("总结段落"), "修改未生效");
  // 重新加载会话，结构版本保持
  const again = await Session.load(s.id);
  eq(again.doc.structureVersion, s.doc.structureVersion, "重新加载后结构版本应保持");
});

await test("消息规范化：补齐缺失的工具结果、丢弃孤立结果", () => {
  const broken: ApiMessage[] = [
    { role: "user", content: "hi" },
    { role: "assistant", content: null, tool_calls: [{ id: "a", type: "function", function: { name: "doc_read", arguments: "{}" } }] },
    { role: "user", content: "again" },
    { role: "tool", tool_call_id: "zzz", content: "orphan" },
  ];
  const n = normalizeMessagesForAPI(broken, true);
  assertPaired(n);
  assert(!n.some((m) => m.role === "tool" && m.tool_call_id === "zzz"), "孤立的工具结果应被丢弃");
  assert(n.filter((m) => m.role === "assistant").every((m) => (m as any).reasoning_content !== undefined), "思考模式下应补 reasoning_content");
});

await test("微压缩：只折叠较早的大段工具输出", () => {
  const h: ApiMessage[] = [];
  for (let i = 0; i < 10; i++) {
    h.push({ role: "assistant", content: null, tool_calls: [{ id: `t${i}`, type: "function", function: { name: "doc_read", arguments: "{}" } }] });
    h.push({ role: "tool", tool_call_id: `t${i}`, content: "x".repeat(3000) });
  }
  const saved = microCompact(h, 6);
  assert(saved > 0, "应节省上下文");
  eq(h.filter((m) => m.role === "tool" && m.content.length > 2000).length, 6, "最近 6 个结果应保留");
});

await test("本地斜杠命令不调用模型", async () => {
  const s = await newSession();
  for (const cmd of ["/help", "/cost", "/mode acceptEdits", "/history", "/verify", "/skills", "/agents"]) {
    const r = await processUserInput(cmd, { session: s, runtime, compact: async () => ({ before: 0, after: 0 }) });
    eq(r.kind, "local", `${cmd} 应为本地命令`);
  }
  eq(s.meta.mode, "acceptEdits", "/mode 应切换模式");
  const r = await processUserInput("/proofread 第二节", { session: s, runtime, compact: async () => ({ before: 0, after: 0 }) });
  eq(r.kind, "prompt", "技能命令应转成提示词");
});

// ---------------------------------------------------------------------------
console.log("\n编辑 / 排版 / 绘图工具");
// ---------------------------------------------------------------------------

const allow = () => ({ decision: "allow" as const });
const lastRefOf = (r: { messages: ApiMessage[] }) => {
  const tool = [...r.messages].reverse().find((m) => m.role === "tool") as { content: string } | undefined;
  return [...(tool?.content ?? "").matchAll(/\[((?:#[0-9A-F]{8})|(?:P\d+@v\d+))\]/g)].map((m) => m[1]);
};

await test("PDF 会话内转换：Agent 调用 doc_convert_to_word 后写工具立即可用，并能接着修改；可回滚到 PDF", async () => {
  const s = await Session.create(await fixture("form.pdf"), "表单.pdf", { mode: "bypassPermissions", trackChanges: false, author: "测试" });
  let toolsAfter: string[] = [];
  const client = new MockClient([
    (req) => { assert(!req.tools.some((t) => t.function.name === "doc_replace_text"), "PDF 中不应有写工具"); return { toolCalls: [tc("doc_convert_to_word", { layout: "exact" })] }; },
    (req) => { toolsAfter = req.tools.map((t) => t.function.name); return { toolCalls: [tc("doc_search", { query: "PENGGUNAAN", case_sensitive: true })] }; },
    (req) => ({ toolCalls: [tc("doc_replace_text", { ref: lastRefOf(req)[0], old_text: "PENGGUNAAN", new_text: "PENGGUNAAN（已修改）" })] }),
    () => ({ content: "已完成。" }),
  ]);
  const { events } = await drive(s, client, "把 PENGGUNAAN 改一下", allow);
  const errs = events.filter((e) => e.type === "tool_result" && !e.ok).map((e: any) => `${e.name}: ${e.content}`);
  assert(!errs.length, errs.join("\n"));
  assert(toolsAfter.includes("doc_replace_text") && toolsAfter.includes("doc_insert_table"), "转换后应出现写工具与编辑工具");
  eq(s.meta.format, "docx", "会话格式应变为 Word");
  eq(s.meta.sourceFormat, "pdf", "记录原始格式");
  assert(s.doc.listBlocks().some((b) => b.text.includes("PENGGUNAAN（已修改）")), "转换后的修改应生效");
  assert(s.doc.fidelity(await s.original()).ok, "保真校验（以转换结果为基准）");
  eq((await s.uploaded()).ext, "pdf", "原始文件仍是上传的 PDF");
  await s.rollback(0);
  eq(s.meta.format, "pdf", "回滚到 v0 后恢复为 PDF");
  eq(s.meta.filename, "表单.pdf", "文件名随之恢复");
});

await test("Word 编辑工具：表格 / 图表 / 流程图 / 绘图 / 公式 / 编号 / 脚注 / 链接 / 目录 / 分节 / 页面设置 / 页眉页脚 / 样式，结构合法", async () => {
  const s = await newSession();
  s.meta.mode = "bypassPermissions";
  const name = await s.saveAsset("测试 图片.png", Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP4z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg==", "base64"));
  let refs: string[] = [];
  const client = new MockClient([
    () => ({ toolCalls: [tc("doc_read", { limit: 60 })] }),
    (req) => {
      refs = lastRefOf(req);
      const a = refs[3], b = refs[refs.length - 1];
      return { toolCalls: [
        tc("doc_insert_table", { anchor: a, position: "after", rows: [["指标", "数值"], ["A", "1"], ["B", "2"]], style: "three_line", caption: "表 1 测试" }),
        tc("doc_insert_chart", { anchor: b, position: "after", type: "line", title: "趋势", categories: ["1月", "2月", "3月"], series: [{ name: "销量", values: [3, 5, 4] }] }),
        tc("doc_insert_diagram", { anchor: b, position: "after", nodes: [{ id: "a", text: "开始", shape: "ellipse" }, { id: "b", text: "处理" }, { id: "c", text: "结束", shape: "ellipse" }], edges: [{ from: "a", to: "b" }, { from: "b", to: "c" }], caption: "图 1 流程" }),
        tc("doc_draw", { anchor: b, position: "after", width_pt: 200, height_pt: 60, shapes: [{ type: "rect", x: 0, y: 0, w: 80, h: 40, text: "A" }, { type: "arrow", x: 80, y: 20, x2: 140, y2: 20 }] }),
        tc("doc_insert_image", { anchor: b, position: "after", asset: name, width_pt: 60, caption: "图 2 素材" }),
        tc("doc_insert_equation", { anchor: b, position: "after", latex: "\\frac{a}{b} = \\sqrt{c}", number: "(1)" }),
        tc("doc_header_footer", { kind: "footer", text: "第 {PAGE} 页" }),
        tc("doc_page_setup", { orientation: "landscape", paper: "A4" }),
        tc("doc_modify_style", { style: "Normal", font: "宋体", size_pt: 12, line_spacing: 1.5 }),
      ] };
    },
    () => ({ toolCalls: [tc("doc_read", { limit: 80 })] }),
    (req) => {
      const r = lastRefOf(req);
      return { toolCalls: [
        tc("doc_set_list", { refs: [r[r.length - 1]], kind: "number" }),
        tc("doc_insert_toc", { anchor: r[0], position: "after", levels: 2 }),
        tc("doc_insert_break", { anchor: r[0], position: "after", kind: "section_next_page" }),
      ] };
    },
    () => ({ toolCalls: [tc("doc_search", { query: "quick brown" })] }),
    (req) => ({ toolCalls: [
      tc("doc_insert_footnote", { ref: lastRefOf(req)[0], after_text: "pangram", text: "全字母句" }),
      tc("doc_insert_link", { ref: lastRefOf(req)[0], text: "lazy dog", url: "https://example.com" }),
    ] }),
    () => ({ content: "完成。" }),
  ]);
  const { events } = await drive(s, client, "排版", allow);
  const errors = events.filter((e) => e.type === "tool_result" && !e.ok).map((e: any) => `${e.name}: ${e.content}`);
  assert(!errors.length, errors.join("\n"));
  const data = s.current();
  const d = new DocxDocument(data);
  const rep = d.fidelity(await s.original());
  assert(rep.ok, rep.problems.join("；"));
  const { ZipPackage } = await import("../src/documents/docx/zip.js");
  const z = new ZipPackage(data);
  const xml = z.readText("word/document.xml");
  for (const [what, ok] of [
    ["表格", (xml.match(/<w:tbl>/g) ?? []).length >= 2], ["图表", z.names().some((n) => /charts\/chart\d+\.xml$/.test(n)) && z.names().some((n) => n.endsWith(".xlsx"))],
    ["形状组合", xml.includes("wordprocessingGroup")], ["图片", z.names().some((n) => n.startsWith("word/media/"))], ["公式", xml.includes("<m:f>") && xml.includes("<m:rad>")],
    ["页脚页码", z.names().some((n) => /footer\d+\.xml/.test(n))], ["横向", xml.includes('w:orient="landscape"')], ["编号", z.has("word/numbering.xml") || xml.includes("<w:numPr>")],
    ["目录", xml.includes("TOC \\o")], ["分节", (xml.match(/<w:sectPr/g) ?? []).length >= 2], ["脚注", xml.includes("footnoteReference")], ["链接", xml.includes("<w:hyperlink")],
  ] as const) assert(ok, `缺少${what}`);
  // 每次修改都生成了版本
  assert(s.meta.versions.length >= 14, `版本数 ${s.meta.versions.length}`);
});

await test("Markdown / HTML：表格、图表（SVG）、流程图、公式、分页符以源码形式插入，其余内容不变", async () => {
  for (const name of ["sample.md", "sample.html"]) {
    const s = await newSession(name);
    s.meta.mode = "bypassPermissions";
    const before = (await s.original()).toString("utf8");
    const client = new MockClient([
      () => ({ toolCalls: [tc("doc_read", {})] }),
      (req) => {
        const r = lastRefOf(req).length ? lastRefOf(req) : [...((req.messages.at(-1) as any).content as string).matchAll(/\[(B\d+@v\d+|H\d+@v\d+|[A-Z]\d+@v\d+)\]/g)].map((m) => m[1]);
        const a = r[r.length - 1];
        return { toolCalls: [
          tc("doc_insert_table", { anchor: a, position: "after", rows: [["a", "b"], ["1", "2"]] }),
        ] };
      },
      () => ({ toolCalls: [tc("doc_read", {})] }),
      (req) => {
        const r = [...((req.messages.at(-1) as any).content as string).matchAll(/\[([A-Z]\d+@v\d+)\]/g)].map((m) => m[1]);
        return { toolCalls: [tc("doc_insert_chart", { anchor: r[r.length - 1], position: "after", type: "column", categories: ["x", "y"], series: [{ name: "s", values: [1, 2] }] })] };
      },
      () => ({ toolCalls: [tc("doc_read", {})] }),
      (req) => {
        const r = [...((req.messages.at(-1) as any).content as string).matchAll(/\[([A-Z]\d+@v\d+)\]/g)].map((m) => m[1]);
        return { toolCalls: [tc("doc_insert_equation", { anchor: r[r.length - 1], position: "before", latex: "E = mc^2" })] };
      },
      () => ({ content: "完成。" }),
    ]);
    const { events } = await drive(s, client, "插入", allow);
    const errors = events.filter((e) => e.type === "tool_result" && !e.ok).map((e: any) => `${e.name}: ${e.content}`);
    assert(!errors.length, `${name}：${errors.join("\n")}`);
    const after = s.current().toString("utf8");
    assert(name.endsWith(".md") ? after.includes("| a | b |") && after.includes("$$") : after.includes("<table") && after.includes("\\["), `${name} 未插入预期标记`);
    assert(after.includes("<svg") || after.includes("data:image/svg+xml"), `${name} 缺少 SVG 图表`);
    // 原文的每一行都还在
    for (const line of before.split(/\r?\n/).filter((l) => l.trim())) assert(after.includes(line), `${name} 原文被改动：${line}`);
  }
});

await test("修订模式下插入的表格 / 图片 / 公式可整体拒绝，恢复原文", async () => {
  const s = await newSession();
  s.meta.trackChanges = true;
  const o = { track: true, author: "测试", date: "2026-01-01T00:00:00Z" };
  const doc = s.doc as import("../src/documents/types.js").DocumentAdapter;
  const orig = doc.listBlocks().map((b) => b.text).join("|");
  const a = doc.listBlocks().find((b) => b.text.includes("quick brown"))!.ref;
  doc.insertTable!({ anchor: a, position: "after", rows: [["x", "y"], ["1", "2"]], caption: "表 X" }, o);
  doc.insertEquation!({ anchor: doc.listBlocks().find((b) => b.text.includes("quick brown"))!.ref, position: "after", latex: "a^2", number: "(9)" }, o);
  doc.reviewChanges!({ action: "reject" }, o);
  eq(doc.listBlocks().map((b) => b.text).join("|"), orig, "拒绝全部修订后应与原文一致");
});

await test("思考模式开关：会话级设置覆盖全局，关闭后请求不带思考", async () => {
  const s = await newSession();
  const seen: Array<boolean | undefined> = [];
  const client = new MockClient([() => ({ content: "好的。" })]);
  const orig = client.complete.bind(client);
  client.complete = async (req: any, opts: any) => { seen.push(req.thinking); return orig(req, opts); };
  s.meta.thinking = false;
  await drive(s, client, "你好", allow);
  eq(seen[0], false, "关闭思考模式后请求应带 thinking=false");
  s.meta.thinking = true;
  await drive(s, new (class extends MockClient {})([() => ({ content: "好" })]), "再来", allow);
  const { buildTranscript } = await import("../src/session/transcript.js");
  const t = buildTranscript(s.history);
  assert(t.some((x) => x.kind === "assistant" && (x as any).reasoning), "对话记录应保留思考过程（刷新后可展开）");
});

await test("位图切片拼成的组图（图题以 (a) 开头）合并为一张图，图题仍是正文", async () => {
  const { detectImageGrids } = await import("../src/services/pdfConvert/raster.js");
  const png = Buffer.alloc(0);
  const tile = (x0: number, y0: number, x1: number, y1: number) => ({ x0, y0, x1, y1, png, pxW: 10, pxH: 10 });
  const span = (text: string, x0: number, x1: number, top: number) => ({ text, x0, x1, top, bottom: top + 8, baseline: top + 6.5, size: 8, font: "Arial", bold: false, italic: false, color: "000000" });
  const pg: any = {
    index: 0, width: 576, height: 783, segs: [], fills: [], paths: [], vectorShapes: 0, rotatedText: 0,
    // Figure 1：7 张相邻的位图切片，另有页眉里的 logo
    images: [tile(461, 18, 538, 32), tile(36, 66, 144, 135), tile(144, 66, 251, 135), tile(251, 66, 359, 135), tile(359, 66, 467, 135), tile(467, 66, 538, 170), tile(36, 135, 287, 170), tile(287, 135, 467, 170)],
    spans: [
      // 图题第一行里就有 "(a)"：不能把它当成子图编号，否则范围会扩进图题
      span("FIGURE 1.", 36, 80, 176), span("The proposed convolution-free network for 3D medical image segmentation. Left: An overall schematic of the method:", 82, 480, 176),
      span("(a)", 482, 494, 176), span("an", 496, 536, 176),
      span("image block is divided into patches, (b) each patch is reshaped into a vector and embedded into a lower dimension, (c) positional", 36, 536, 187),
    ],
  };
  const grids = detectImageGrids(pg, 10);
  eq(grids.length, 1, "7 张切片应合并成一张图");
  assert(grids[0].y0 > 60 && grids[0].y1 < 175 && grids[0].x0 < 40 && grids[0].x1 > 534, `组图范围应是切片的外接矩形，不含图题：${JSON.stringify([grids[0].x0, grids[0].y0, grids[0].x1, grids[0].y1])}`);
  assert(pg.spans.some((x: any) => x.text === "FIGURE 1."), "图题仍保留为正文");
});

// ---------------------------------------------------------------------------

await fs.rm(tmp, { recursive: true, force: true });
const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} 通过`);
if (failed.length) {
  for (const f of failed) console.log(`\n\x1b[31m${f.name}\x1b[0m\n${f.err}`);
  process.exit(1);
}
