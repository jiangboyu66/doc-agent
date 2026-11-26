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

await test("只转换部分页：页码范围写法；导出 / 会话内转换只处理这些页，报告与文件名注明页码，重新转换沿用范围；超过单次上限时报告没转换的页", async () => {
  const { parsePageRange, pdfToDocx, describeReport } = await import("../src/services/pdfConvert/index.js");
  eq(parsePageRange("1-20", 611).length, 20, "1-20");
  eq(parsePageRange("第3至5页", 10).join(","), "3,4,5", "中文写法");
  eq(parsePageRange("前2页", 10).join(","), "1,2", "前 N 页");
  eq(parsePageRange("8-, 1，3", 10).join(","), "1,3,8,9,10", "到最后一页、多段、全角逗号");
  eq(parsePageRange("5-30", 10).join(","), "5,6,7,8,9,10", "超出的部分截到最后一页");
  let threw = "";
  try { parsePageRange("700-720", 611); } catch (e: any) { threw = e.message; }
  assert(threw.includes("611"), `超出页码范围应报错并给出总页数：${threw}`);

  const pdf = await fixture("colstart.pdf"); // 3 页
  const whole = await pdfToDocx(pdf, { mode: "exact" });
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode, pages: "2-3" });
    eq(report.pages, 2, `${mode}：转换页数`);
    eq(report.pageRange, "2–3", `${mode}：页码`);
    eq(report.totalPages, 3, `${mode}：原 PDF 总页数`);
    eq(report.textCoverage, 1, `${mode}：这些页文字完整`);
    assert(report.textTotal < whole.report.textTotal, `${mode}：只含第 2–3 页的文字`);
    assert(describeReport(report).includes("第 2–3 页（共 2 页，原 PDF 共 3 页）"), describeReport(report));
    // 逐页保留：第一页（原 PDF 第 2 页）前面不分页，只在原第 3 页前分页
    if (mode === "exact") eq((new ZipPackage(data).readText("word/document.xml").match(/w:val="PageStart"/g) ?? []).length, 1, "逐页保留：只在第二页前分页");
  }
  // 整份超过单次上限：只转换前面的页，报告写明哪些页没有转换
  const cut = await pdfToDocx(pdf, { mode: "exact", maxPages: 2 });
  eq(cut.report.pages, 2, "上限内的页数");
  assert(describeReport(cut.report).includes("共 3 页，超过单次转换上限 2 页：只转换了第 1–2 页，第 3 页没有转换"), describeReport(cut.report));

  // Agent 工具：导出指定页、会话内转换指定页，重新转换沿用范围
  const { findTool } = await import("../src/tools.js");
  const s = await Session.create(pdf, "论文.pdf", { mode: "bypassPermissions", trackChanges: false, author: "测试" });
  const exp = await findTool("doc_export")!.call({ format: "docx", pages: "2-3" } as any, { session: s } as any);
  eq((exp.data as any)?.name, "论文（第2–3页）.docx", "导出文件名注明页码");
  const conv = findTool("doc_convert_to_word")!;
  const r1 = await conv.call({ layout: "flow", pages: "2-3" } as any, { session: s } as any);
  assert(String(r1.content).includes("第 2–3 页"), `报告应注明页码：${r1.content}`);
  await s.commit((r1 as any).docChange.label, []);
  eq(s.meta.filename, "论文（第2–3页）.docx", "会话文件名注明页码");
  eq(s.meta.origin?.pages, "2-3", "记下页码范围");
  const r2 = await conv.call({ layout: "exact" } as any, { session: s } as any);
  assert(String(r2.content).includes("第 2–3 页"), `重新转换应沿用页码范围：${r2.content}`);
});

await test("LaTeX 论文：裁剪图片按 BBox / 裁剪框取可见部分，示意图整体成图，纵轴竖排标题与刻度包进图里，打字机字体的邮箱收紧到原宽度", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const { extractPdf } = await import("../src/services/pdfConvert/extract.js");
  const pdf = await fixture("texpaper.pdf");
  // 提取：表单 XObject 的 BBox、带多余 moveto 的裁剪路径都要生效
  const [pg] = await extractPdf(pdf);
  const boxes = pg.images.map((im) => [im.x0, im.y0, im.x1, im.y1].map((v) => Math.round(v)));
  assert(boxes.some((b) => b.join() === "312,320,462,520"), `\\includegraphics[clip] 的图只取 BBox 内的 150 × 200：${JSON.stringify(boxes)}`);
  for (const x of [182, 343, 432]) assert(boxes.some((b) => b.join() === `${x},610,${x + 20},630`), `节点旁的小图只取裁剪框内的 20 × 20：${JSON.stringify(boxes)}`);
  assert(boxes.every((b) => b[2] - b[0] <= 201), `没有未裁剪的大图：${JSON.stringify(boxes)}`);
  const EMU = 12700;
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    const xml = new ZipPackage(data).readText("word/document.xml");
    const texts = (xml.match(/<w:t[^>]*>[^<]*<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, "")).join("|");
    // 图片尺寸：最宽的是左边 200 pt 的图，裁剪后的图不应比原文大
    const widths = [...xml.matchAll(/<wp:extent cx="(\d+)"/g)].map((m) => Number(m[1]) / EMU);
    assert(widths.every((w) => w <= 400), `${mode}：图片宽度 ${widths.map((w) => w.toFixed(0)).join(",")}`);
    // 示意图整体渲染成一幅图：节点之间的省略号、节点里的字母不再散落成文字
    assert(!texts.includes("· · ·"), `${mode}：示意图的省略号不应成为正文文字`);
    // 曲线图：刻度数字不散落到图外
    assert(!/\|(?:2|4|6|8)\|/.test(texts) && !texts.includes("|200|"), `${mode}：刻度数字应在图里：${texts.slice(-200)}`);
    // 打字机字体：Courier New 收紧到原宽度（-0.75 pt/字），不能被加宽
    const mono = (xml.match(/<w:r>(?:(?!<\/w:r>)[\s\S])*Courier New(?:(?!<\/w:r>)[\s\S])*<\/w:r>/g) ?? []).filter((r) => /<w:t[^>]*>[^<]*[@/][^<]*<\/w:t>/.test(r));
    assert(mono.length >= 4, `${mode}：邮箱与网址应是 Courier New：${mono.length}`);
    for (const r of mono) {
      const sp = Number(r.match(/<w:spacing w:val="(-?\d+)"/)?.[1] ?? 0);
      assert(sp <= -12 && sp >= -18, `${mode}：等宽文字的字符间距应约为 -15（-0.75 pt）：${sp} ${r.replace(/<[^>]+>/g, "")}`);
    }
    if (mode === "exact") {
      // 三个邮箱在同一段、以制表位分隔（不折行）
      const para = (xml.match(/<w:p [^>]*>(?:(?!<\/w:p>)[\s\S])*alicezhang@(?:(?!<\/w:p>)[\s\S])*<\/w:p>/) ?? [""])[0];
      assert(para.includes("boli@") && para.includes("carols@") && !para.includes("<w:br"), "exact：三个邮箱在同一行");
      // 曲线图的图片包含左侧竖排的纵轴标题（x ≈ 148 pt）
      const anchors = [...xml.matchAll(/<wp:anchor[\s\S]*?<\/wp:anchor>/g)].map((m) => m[0]).filter((a) => /descr="[^"]*200[^"]*400/.test(a));
      assert(anchors.length === 1, `exact：曲线图应是一幅图：${anchors.length}`);
      const h = Number(anchors[0].match(/<wp:positionH relativeFrom="page"><wp:posOffset>(-?\d+)</)?.[1] ?? NaN) / EMU;
      assert(h <= 149, `exact：曲线图左边缘应包含纵轴标题：${h.toFixed(1)}`);
    }
  }
});

await test("期刊里蓝色的引用编号：只有编号是蓝色，同一文字项里后面的黑色正文不被染蓝（两种模式）", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("citecolor.pdf");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    const xml = new ZipPackage(data).readText("word/document.xml");
    const runs = (xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? []).map((r) => ({ text: (r.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, "")).join(""), blue: /<w:color w:val="0000FF"\/>/.test(r) }));
    const blue = runs.filter((r) => r.blue && r.text.trim()).map((r) => r.text.trim());
    assert(blue.length >= 6 && blue.every((t) => /^[[\]\d,\s]+$/.test(t)), `${mode}：蓝色的只应是引用编号：${JSON.stringify(blue)}`);
    for (const t of ["[18]", "[23]", "[29]", "[27], [28]"]) assert(blue.includes(t), `${mode}：${t} 应是蓝色：${JSON.stringify(blue)}`);
    assert(runs.filter((r) => /While the transformer method|In particular, TransUNet and TransFuse/.test(r.text)).every((r) => !r.blue), `${mode}：引用后面的正文不应是蓝色`);
    assert(/<w:t[^>]*>\[23\]<\/w:t><\/w:r><w:r>(?:(?!<\/w:r>)[\s\S])*<w:t[^>]*>\. While/.test(xml), `${mode}：[23] 与后面的正文紧挨着，中间不多出空格`);
  }
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

await test("作者简介页：方形照片被取景框裁成竖幅，只取可见部分；相邻两位作者的照片不合并，正文不被盖住", async () => {
  const { extractPdf } = await import("../src/services/pdfConvert/extract.js");
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("bios.pdf");
  const [pg] = await extractPdf(pdf);
  const photos = pg.images.filter((im) => im.y0 > 50);
  eq(photos.length, 5, "5 位作者的照片");
  for (const im of photos) {
    assert(Math.abs(im.x1 - im.x0 - 72) < 1.5 && Math.abs(im.y1 - im.y0 - 90) < 1.5, `照片应是取景框内的 72 × 90：${JSON.stringify([im.x0, im.y0, im.x1, im.y1])}`);
    assert(Math.abs(im.pxW / im.pxH - 72 / 90) < 0.05, `像素也应按取景框裁剪：${im.pxW}×${im.pxH}`);
  }
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.figures, 0, `${mode}：两张相邻的照片不应被合并成一张图`);
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    const xml = new ZipPackage(data).readText("word/document.xml");
    const sizes = [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g)].map((m) => [+m[1] / 12700, +m[2] / 12700]).filter(([w, h]) => w > 50 && h > 50 && h < 200);
    eq(sizes.length, 5, `${mode}：5 张照片`);
    for (const [w, h] of sizes) assert(Math.abs(w - 72) < 1.5 && Math.abs(h - 90) < 1.5, `${mode}：照片应是 72 × 90 pt：${w} × ${h}`);
  }
});

await test("图页：通栏大图与图题之后页面下方只剩几行双栏正文，沿用全文的分栏，不把左右两栏拼成一行", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("colstart.pdf");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    const xml = new ZipPackage(data).readText("word/document.xml");
    const paras = (xml.match(/<w:p [^>]*>[\s\S]*?<\/w:p>/g) ?? []).map((p) => p.replace(/<w:tab\/>/g, "\t").replace(/<[^>]+>/g, ""));
    const joined = paras.filter((t) => t.includes("medical images with transformer") && t.includes("\t"));
    eq(joined.length, 0, `${mode}：左右两栏不应被制表符拼成一行：${JSON.stringify(joined)}`);
    assert(!paras.some((t) => /size of\s+transformer encoders reduces/.test(t)), `${mode}：左栏末与右栏首不应在同一段`);
  }
});

await test("字形单位为 1/2048 em 的字体（Elsevier 的 Charis SIL）：按字体矩阵计算字宽，不把字符间距压成负值、文字不挤在一起", async () => {
  const { extractPdf } = await import("../src/services/pdfConvert/extract.js");
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("fontmatrix.pdf");
  const [pg] = await extractPdf(pdf);
  eq(pg.spans.length, 3, "3 行文字");
  for (const sp of pg.spans) assert(sp.natural && Math.abs(sp.natural / (sp.x1 - sp.x0) - 1) < 0.03, `自然字宽应与实际宽度一致：${sp.text} ${sp.natural} / ${sp.x1 - sp.x0}`);
  for (const mode of ["exact", "flow"] as const) {
    const { data } = await pdfToDocx(pdf, { mode });
    const xml = new ZipPackage(data).readText("word/document.xml");
    const tight = [...xml.matchAll(/<w:spacing w:val="(-?\d+)"\/>/g)].map((m) => +m[1]).filter((v) => v < -10);
    eq(tight.length, 0, `${mode}：不应出现每字紧缩超过 0.5 磅的字符间距：${tight.join(",")}`);
  }
});

await test("期刊首页：文章信息 / 摘要左右并排、字距加宽的标题、作者名后的小图标、上下相邻的两张三线表、页码上方带分式的公式", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("journal-front.pdf");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    const xml = new ZipPackage(data).readText("word/document.xml");
    const tables = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? [];
    const textOf = (x: string) => x.replace(/<[^>]+>/g, "");
    // 并排的两块各成一格，摘要不被关键词的行切碎
    const side = tables.find((t) => textOf(t).includes("Keywords:"));
    assert(side && /tomography reduces/.test(textOf(side)), `${mode}：关键词与摘要应在同一张无边框表格的两格里`);
    const cells = side!.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? [];
    assert(!cells.some((c) => textOf(c).includes("Keywords:") && textOf(c).includes("tomography")), `${mode}：关键词与摘要不应在同一格`);
    // 两张三线表分开识别，不退回成散落的文字
    eq(tables.filter((t) => textOf(t).includes("SSIM")).length, 2, `${mode}：两张表都应识别成表格`);
    // 字距加宽：字符间距写进 Word
    assert(/<w:spacing w:val="([4-9]\d)"\/>(?:(?!<\/w:rPr>).)*<\/w:rPr><w:t[^>]*>ARTICLE/.test(xml), `${mode}：加宽的字距应写成字符间距`);
    assert(/<w:spacing w:val="([4-9]\d)"\/>(?:(?!<\/w:rPr>).)*<\/w:rPr><w:t[^>]*>ABSTRACT/.test(xml), `${mode}：逐字母定位的标题应合成一个词并带字符间距`);
    // 作者名后的小图标是该行的行内图片，不单独成段
    const authors = (xml.match(/<w:p [^>]*>(?:(?!<\/w:p>)[\s\S])*Zihui Zhao[\s\S]*?<\/w:p>/) ?? [""])[0];
    eq((authors.match(/<wp:inline/g) ?? []).length, 4, `${mode}：4 个图标都应在作者行里`);
    eq(report.images, 4, `${mode}：图标计入图片数`);
    // 分数线不被当成破折号；带分式的公式整体成图；页码上方的正文与公式不被卷进页脚
    assert(!textOf(xml).includes("\u2014"), `${mode}：分数线不应变成破折号`);
    assert(report.equations >= 1, `${mode}：带分式的公式应整体成图`);
    const boxes = (xml.match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g) ?? []).map(textOf);
    assert(!boxes.some((t) => /radiation|frequency/.test(t)), `${mode}：正文不应被当成页脚放进文本框：${boxes.join(" | ")}`);
  }
});

await test("期刊刊头灰底（粗描边线）、自定义字体回退、逐字母字距标题、奇偶页版心偏移、分数线公式、字母上方的重音", async () => {
  const { extractPdf } = await import("../src/services/pdfConvert/extract.js");
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const { isSafeFamily } = await import("../src/services/pdfConvert/fonts.js");
  const pdf = await fixture("journal-eq.pdf");
  const pages = await extractPdf(pdf);
  // 60 多磅宽的描边线是色块，不是线：不进线段（否则被识别成穿过刊名的删除线），而是底色
  assert(!pages[0].segs.some((g) => g.horizontal && g.width > 20), "粗描边线不应作为线段");
  assert(pages[0].fills.some((f) => f.y1 - f.y0 > 50 && f.x1 - f.x0 > 300), "粗描边线应成为底色块");
  // 自定义字体（族名不在 Word 常见字体里）回退到同类的常见字体
  for (const p of pages) for (const sp of p.spans) assert(sp.math || isSafeFamily(sp.font), `字体应是 Word 里常见的：${sp.font}`);
  // 逐字母定位的标题合成一个词、带字符间距
  assert(pages[0].spans.some((sp) => sp.text === "ARTICLE" && (sp.tracking ?? 0) > 2), `逐字母标题应合成 ARTICLE：${pages[0].spans.filter((sp) => sp.baseline > 150 && sp.baseline < 170).map((sp) => sp.text).join("|")}`);
  // 公式里隔着上下标的同字母（R₁¹, R₂¹, …）各自独立，不能被并成加宽字距的词 "RRRR"
  assert(!pages[2].spans.some((sp) => sp.text === "RRRR"), "隔着上下标的 R 不应被并成一个词");
  eq(pages[2].spans.filter((sp) => /(^|\s)R$/.test(sp.text)).length, 4, "四个 R 各自独立");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    const xml = new ZipPackage(data).readText("word/document.xml");
    // 刊头灰底：写成底纹 / 色块，而不是画在刊名上的线
    assert(!/<a:ln w="\d{6,}"/.test(xml), `${mode}：不应有 60 磅宽的线`);
    // 分数线公式整体成图（page 2 两条：分式、重音在内；page 3 一条）
    assert(report.equations >= 2, `${mode}：栏内的分式公式应整体成图：${report.equations}`);
    const text = xml.replace(/<wp:inline[\s\S]*?<\/wp:inline>/g, "[IMG]").replace(/<[^>]+>/g, "");
    assert(!/\u2014/.test(text), `${mode}：分数线不应变成破折号`);
    // 公式编号并进公式图片，不单独留下一行 "(10)"、"(8)"
    assert(!/(^|\s)\(10\)(\s|$)/.test(text) || report.equations >= 3, `${mode}：公式编号应并进公式图片`);
  }
});

await test("上下叠放的两幅图（各由左右两张位图拼成）：图题夹在中间时不并成一组；矢量小图标嵌回作者行；多行公式续行带编号", async () => {
  const { detectImageGrids, detectIconPaths } = await import("../src/services/pdfConvert/raster.js");
  const { extractPdf } = await import("../src/services/pdfConvert/extract.js");
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("journal-fig.pdf");
  const pages = await extractPdf(pdf);
  // 第 2 页：4 张位图一行两张，上下两幅；图题夹在两幅之间 → 两个组，且组的范围不含任何一幅图的图题
  const grids = detectImageGrids({ ...pages[1], images: [...pages[1].images], spans: [...pages[1].spans] } as any, 10);
  eq(grids.length, 2, `两幅图各自成组：${JSON.stringify(grids.map((g) => [g.y0, g.y1]))}`);
  assert(grids[0].y1 < grids[1].y0, "两组上下分开，不重叠");
  // 第 1 页：矢量小图标（4 个）
  const icons = detectIconPaths({ ...pages[0] } as any, []);
  eq(icons.length, 4, `作者名后的矢量小图标：${icons.length}`);
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    const xml = new ZipPackage(data).readText("word/document.xml");
    const author = (xml.match(/<w:p [^>]*>(?:(?!<\/w:p>)[\s\S])*Bo Su[\s\S]*?<\/w:p>/) ?? [""])[0];
    eq((author.match(/<wp:inline/g) ?? []).length, 4, `${mode}：4 个小图标都应在作者行里`);
    eq(report.figures >= 2, true, `${mode}：两幅图整体成图：${report.figures}`);
    // 第 3 页：带编号的续行并进公式图片，不单独留下一行 "(10)"
    const text = xml.replace(/<wp:inline[\s\S]*?<\/wp:inline>/g, "[IMG]").replace(/<[^>]+>/g, "");
    assert(!/(^|\s)\(10\)(\s|$)/.test(text), `${mode}：公式编号应并进公式图片：${text.match(/.{20}\(10\).{20}/)?.[0]}`);
  }
});

await test("所有夹具 PDF：两种排版方式都不报错、文字完整；LibreOffice 渲染的页数与原 PDF 一致（不出现留白 / 溢页）", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const { extractPdf } = await import("../src/services/pdfConvert/extract.js");
  const { execFile } = await import("node:child_process");
  const { promisify } = await import("node:util");
  const run = promisify(execFile);
  const soffice = process.env.SOFFICE_PATH || "soffice";
  const hasSoffice = await run(soffice, ["--version"]).then(() => true, () => false);
  const names = ["columns", "journal", "figures", "math", "journal-front", "journal-eq", "journal-fig", "bios", "colstart", "signature", "citecolor", "texpaper", "contract-cn", "scan-letter", "furniture-cn", "bates"];
  const dir = await fs.mkdtemp(path.join(tmp, "pages-"));
  const bad: string[] = [];
  for (const name of names) {
    const pdf = await fixture(`${name}.pdf`);
    const pdfPages = (await extractPdf(pdf)).length;
    for (const mode of ["exact", "flow"] as const) {
      const { data, report } = await pdfToDocx(pdf, { mode });
      if (report.textCoverage < 1) bad.push(`${name}/${mode}：文字缺失 ${report.missing}`);
      if (!hasSoffice) continue;
      const f = path.join(dir, `${name}-${mode}.docx`);
      await fs.writeFile(f, data);
      await run(soffice, ["--headless", "--convert-to", "pdf", "--outdir", dir, f], { timeout: 120000 }).catch(() => {});
      const out = await fs.readFile(path.join(dir, `${name}-${mode}.pdf`)).catch(() => null);
      if (!out) continue;
      const n = (await extractPdf(out)).length;
      // 逐页保留方式页数必须与原件一致；可编辑排版允许多 / 少一页（重新换行）
      if (mode === "exact" ? n !== pdfPages : Math.abs(n - pdfPages) > 1) bad.push(`${name}/${mode}：原 ${pdfPages} 页，转换后 ${n} 页`);
    }
  }
  eq(bad.length, 0, bad.join("；"));
});

await test("Type 3 字形（位图蒙版 / 乱码编码）的公式与符号：行内的作为行内图片插回原行，独立的成图（含编号），不出现 Type3 字体、文字不丢", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const { rasterAvailable } = await import("../src/services/pdfConvert/render.js");
  const pdf = await fixture("type3-math.pdf");
  const paraText = (p: string) => p.replace(/<wp:inline[\s\S]*?<\/wp:inline>/g, "[IMG]").replace(/<[^>]+>/g, "");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    const xml = new ZipPackage(data).readText("word/document.xml");
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    assert(!/w:ascii="Type3"/.test(xml), `${mode}：Word 中没有 "Type3" 这个字体，不应写进文档`);
    assert(report.equations >= 1, `${mode}：独立成行的位图公式应成图：${JSON.stringify(report)}`);
    const paras = (xml.match(/<w:p [^>]*>[\s\S]*?<\/w:p>/g) ?? []).map(paraText);
    // 行内的位图字形：图片在原来那一行的文字之间
    assert(paras.some((t) => /loss of transaction\s*\[IMG\]\s*, which is computed below/.test(t)), `${mode}：行内位图字形应插回原行：${paras.join(" | ")}`);
    // 公式编号并进公式图片，不单独成段
    assert(!paras.some((t) => t.trim() === "(1)"), `${mode}：公式编号应并进公式图片`);
    // 固定行距的行不能被行内图片撑高（逐行一致模式按页面排满，一行变高就会把栏底的行挤到下一页）
    if (mode === "exact") assert(!/<w:p [^>]*>(?:(?!<\/w:p>)[\s\S])*w:lineRule="atLeast"(?:(?!<\/w:p>)[\s\S])*loss of transaction/.test(xml), "行内图片应放进固定行距");
    if (await rasterAvailable()) {
      // Type 3 矢量字形（编码为控制字符）按原样渲染成行内图片
      assert(paras.some((t) => /equals the total loss\s*\[IMG\]\s*of the system/.test(t)), `${mode}：Type 3 矢量字形应渲染成行内图片：${paras.join(" | ")}`);
    }
  }
  // 没有画布时：位图字形由蒙版直接合成，公式与符号仍以图片保留
  const prev = process.env.PDF_RASTER;
  process.env.PDF_RASTER = "0";
  try {
    const { data, report } = await pdfToDocx(pdf, { mode: "flow" });
    const pkg = new ZipPackage(data);
    const xml = pkg.readText("word/document.xml");
    eq(report.textCoverage, 1, "无画布：文字完整性");
    assert(!/w:ascii="Type3"/.test(xml), "无画布：不出现 Type3 字体");
    const paras = (xml.match(/<w:p [^>]*>[\s\S]*?<\/w:p>/g) ?? []).map(paraText);
    assert(paras.some((t) => /loss of transaction\s*\[IMG\]/.test(t)), "无画布：行内位图字形仍以图片插回原行");
    assert(report.equations >= 1, "无画布：独立公式仍由蒙版合成成图");
  } finally {
    if (prev === undefined) delete process.env.PDF_RASTER; else process.env.PDF_RASTER = prev;
  }
});

await test("中英左右对照 PDF（中文用浏览器嵌入的 Type 3 字体）：中文栏保留为文字，不当公式成图；表格单元格里的整行删除线保留", async () => {
  // 2026-11-26 LTB-v1 LTB-03（合成测试文件的第 1、4 页）：一页上中文的字数远少于英文的字母数，Type 3 文字占比低于 0.3，
  // 整栏中文被当成公式成图；附件一表内带删除线的一行丢了删除线（删除线检测原来只看表格以外的文字）
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("bilingual-type3.pdf");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    const xml = new ZipPackage(data).readText("word/document.xml");
    eq(report.textCoverage, 1, `${mode}：文字完整性`);
    eq(report.equations, 0, `${mode}：中文正文不应成公式图：${JSON.stringify(report)}`);
    eq(report.inlineMath ?? 0, 0, `${mode}：中文正文不应成行内图片：${JSON.stringify(report)}`);
    const text = xml.replace(/<[^>]+>/g, "");
    assert(text.includes("供应商委任经销商为产品在区域内的独家经销商"), `${mode}：中文栏应是文字`);
    assert(text.includes("The Supplier appoints the Distributor"), `${mode}：英文栏应是文字`);
    const row = (xml.match(/<w:tr[ >](?:(?!<\/w:tr>)[\s\S])*LY-SP60A[\s\S]*?<\/w:tr>/) ?? [""])[0];
    const runs = row.match(/<w:r[ >][\s\S]*?<\/w:r>/g) ?? [];
    const struck = runs.filter((r) => /<w:strike\/>/.test(r)).map((r) => r.replace(/<[^>]+>/g, "")).join("");
    for (const cell of ["LY-SP60A", "7.35", "6.90"]) assert(struck.includes(cell), `${mode}：删除线那一行的 ${cell} 应带删除线：${struck}`);
    assert(!runs.some((r) => /同意删除/.test(r.replace(/<[^>]+>/g, "")) && /<w:strike\/>/.test(r)), `${mode}：行末的手写备注没有删除线`);
  }
});

await test("首字下沉与正文同基线时仍识别为首字下沉；多行居中的标题不被误判为两端对齐；画成短横线的破折号还原为文字", async () => {
  const { detectDropCaps } = await import("../src/services/pdfConvert/raster.js");
  const sp = (text: string, x0: number, x1: number, baseline: number, size: number) => ({ text, x0, x1, top: baseline - 0.75 * size, bottom: baseline + 0.22 * size, baseline, size, font: "Times New Roman", bold: false, italic: false, color: "000000" });
  const line = (spans: any[]) => ({ spans, x0: Math.min(...spans.map((s) => s.x0)), x1: Math.max(...spans.map((s) => s.x1)), top: Math.min(...spans.map((s) => s.top)), bottom: Math.max(...spans.map((s) => s.bottom)), baseline: spans[spans.length - 1].baseline, size: spans[spans.length - 1].size });
  // "T" 的基线与第二行正文对齐，被并进第二行："T rapid disintegration…"
  const l1 = line([sp("HE open access transmission regime is spearheading the", 60, 290, 410, 10)]);
  const l2 = line([sp("T", 40, 57, 422, 30), sp("rapid disintegration of the well-entrenched vertically inte-", 60, 290, 422, 10)]);
  const l3 = line([sp("grated structure of the electric power industry. The entry of a", 40, 290, 434, 10)]);
  const caps = detectDropCaps([l1, l2, l3] as any, { bodySize: 10, bodyBase: "TIMES", bodyIsTeX: false });
  eq(caps.length, 1, "应识别出首字下沉");
  assert(caps[0].rest && caps[0].rest.spans.length === 1 && caps[0].rest.spans[0].text.startsWith("rapid"), "其余正文应放回原行");

  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  for (const mode of ["exact", "flow"] as const) {
    const { data } = await pdfToDocx(await fixture("type3-math.pdf"), { mode });
    const xml = new ZipPackage(data).readText("word/document.xml");
    const title = (xml.match(/<w:p [^>]*>(?:(?!<\/w:p>)[\s\S])*?A Physical-Flow-Based Approach(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/) ?? [""])[0];
    assert(/<w:jc w:val="center"\/>/.test(title), `${mode}：两行居中的标题应居中（版心被页码撑宽也不影响）`);
    // 参考文献编号 "[2]" 与正文之间是制表符（悬挂缩进的列表）：编号与破折号之间可以隔着制表符所在的 run
    assert(/\[2\](\s|<[^>]+>)*——/.test(xml.replace(/<w:tab\/>/g, " ")), `${mode}：画成短横线的破折号应还原为 "——"`);
  }
});

await test("合同签字页：左右并排的签字标签、无框表头不当成双栏；描边的小方框还原为 ☐ / ☒；比表头文字长的填写线不当成下划线；两行页脚、主动换页、连着的填写线（可编辑排版）", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("signature.pdf");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.textCoverage, 1, `${mode}：文字完整性，缺失：${report.missing}`);
    eq(report.columnPages, 0, `${mode}：签字区、表头不应被当成双栏`);
    const xml = new ZipPackage(data).readText("word/document.xml");
    const paras = (xml.match(/<w:p [^>]*>[\s\S]*?<\/w:p>/g) ?? []).map((p) => p.replace(/<w:tab\/>/g, "\t").replace(/<[^>]+>/g, ""));
    // 左右两个签字标签在同一段里，用制表位定位（误判成分栏时左侧标签被挤进只有编号宽的左栏，折成好几行）
    for (const pair of ["(Signature)\t(Signature)", "(Printed Name)\t(Printed Name)", "(Title)\t(Date Signed)"]) {
      assert(paras.some((t) => t.replace(/^\t/, "") === pair), `${mode}：签字标签 ${pair} 应在同一行：${paras.join(" | ")}`);
    }
    // 复选框：3 个空框、1 个打了叉的框，与选项文字同段；字体与 Word 的复选框控件相同；方框的竖边不再单独画出（"| |"）
    eq((xml.match(/☐/g) ?? []).length, 3, `${mode}：空复选框个数`);
    eq((xml.match(/☒/g) ?? []).length, 1, `${mode}：已勾选的复选框个数`);
    assert(paras.some((t) => /^\t?☒\tNo prior works\.$/.test(t)) && paras.filter((t) => /^\t?☐\tAdditional sheets attached\.$/.test(t)).length === 2, `${mode}：复选框应与选项文字在同一段：${paras.join(" | ")}`);
    assert(/<w:rFonts w:ascii="MS Gothic"[^>]*\/>(?:(?!<\/w:r>)[\s\S])*☐/.test(xml), `${mode}：复选框应使用 MS Gothic`);
    // 表头下的填写线比表头文字长：是线，不是下划线（否则伸出文字的部分丢失，表头前的制表符也会带上下划线）；
    // 框住字母的方框，底边也不是字母的下划线
    assert(!xml.includes("<w:u "), `${mode}：表头、框内字母不应带下划线`);
    assert(paras.some((t) => t === "See Exhibit A"), `${mode}：框里的字母是正文`);
    // 逐页保留：线条 = 签字线 6 + 分隔线 1 + 填写行 3 + 表格填写线 12 + 框住字母的方框 4 条边；
    // 竖线只有那个方框的两条边（复选框的竖边不再单独留下，"| |"）
    if (mode === "exact") {
      eq((xml.match(/name="直线 /g) ?? []).length, 26, "逐页保留：线条数");
      eq((xml.match(/<wp:extent cx="0" /g) ?? []).length, 2, "逐页保留：竖线数");
    }
    // 紧挨着的两行页脚（"Consulting Agreement" / "Page N"）：是页脚，不留在正文里
    // （逐页保留：页面绝对坐标的文本框；可编辑排版：真正的页脚，正文里不再夹着页脚和它上方的大段空白）
    const pkg = new ZipPackage(data);
    assert(!xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, "").includes("Consulting Agreement"), `${mode}：页脚不应留在正文里`);
    if (mode === "flow") {
      assert(pkg.names().some((n) => /^word\/footer\d*\.xml$/.test(n) && pkg.readText(n).includes("Consulting Agreement")), "可编辑排版：页脚应写进页脚部件");
      // 原文在 "[Signatures to follow on next page]" 之后主动换页：保留分页符，签字页不被提到上一页
      const brk = xml.indexOf('w:val="PageStart"');
      assert(brk > xml.indexOf("[Signatures to follow on next page]") && brk < xml.indexOf("This Agreement takes effect"), "可编辑排版：签字页之前应保留分页符");
      eq((xml.match(/w:val="PageStart"/g) ?? []).length, 1, "可编辑排版：只在主动换页处分页");
      // 连着的三条相同的填写线（段落下边框）：相邻两段边框、缩进都相同会被 Word / LibreOffice 并成一条，中间要隔开
      const blocks = (xml.match(/<w:p [^>]*>[\s\S]*?<\/w:p>/g) ?? []).map((p) => ({ bdr: p.includes("<w:pBdr>"), ind: p.match(/<w:ind [^>]*\/>/)?.[0] ?? "" }));
      assert(!blocks.some((b, i) => i > 0 && b.bdr && blocks[i - 1].bdr && b.ind === blocks[i - 1].ind), "可编辑排版：相同的填写线不应相邻（会被合并成一条）");
      const sameInd = new Map<string, number>();
      for (const b of blocks) if (b.bdr) sameInd.set(b.ind, (sameInd.get(b.ind) ?? 0) + 1);
      assert([...sameInd.values()].some((n) => n >= 3), "可编辑排版：三条相同的填写线都应保留");
    }
  }
});

await test("中文合同（律所草稿）：1.5 倍行距的条款合成一段、多级编号悬挂缩进；高亮的方括号占位符是可编辑文字；紧贴正文的左右两行页脚；跨页条款", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("contract-cn.pdf");
  const truth: string[] = JSON.parse(await fs.readFile(path.join(here, "fixtures", "contract-cn.truth.json"), "utf8"));
  const norm = (s: string) => s.replace(/[\s\u3000\u00a0]/g, "");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    eq(report.textCoverage, 1, `${mode}：文字完整性，缺失：${report.missing}`);
    // 占位符背后的高亮矩形不是图形：不截成图片
    eq(report.images + report.figures, 0, `${mode}：占位符不应变成图片`);
    const pkg = new ZipPackage(data);
    const xml = pkg.readText("word/document.xml");
    assert(!xml.includes("<pic:pic>"), `${mode}：文档里不应有图片`);
    const body = xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, "");
    const paras = (body.match(/<w:p [^>]*>[\s\S]*?<\/w:p>/g) ?? []).map((p) => norm(p.replace(/<[^>]+>/g, ""))).filter(Boolean);
    // 条款完整：每个条款恰好是一段。逐页保留模式每页独立，跨页的条款（1.5 跨第 1、2 页，第 2、3 页之间还有一条）按页分成两段
    const left = new Map<string, number>();
    for (const p of paras) left.set(p, (left.get(p) ?? 0) + 1);
    const split = truth.filter((t) => { const k = norm(t), n = left.get(k) ?? 0; if (n) left.set(k, n - 1); return !n; });
    assert(split.length <= (mode === "flow" ? 0 : 2), `${mode}：${split.length} 个条款被切开：${split.map((t) => t.slice(0, 12)).join(" | ")}`);
    // 占位符是文字（个数与原文一致），并保留黄色突出显示
    const all = paras.join("");
    for (const ph of ["[30日]", "[]", "[●]"]) eq(all.split(ph).length - 1, truth.join("").replace(/\s/g, "").split(ph).length - 1, `${mode}：占位符 ${ph} 的个数`);
    assert((xml.match(/<w:highlight w:val="yellow"\/>/g) ?? []).length >= 8, `${mode}：占位符应保留黄色突出显示`);
    // 两行页脚（左 "保密文件 · 律师工作草稿"、右 "第 N 页 / 共 3 页"）紧挨着正文末行：是页脚，不留在正文里、不截断跨页的条款
    assert(!paras.some((p) => p.includes("保密文件") || /第\d页\/共3页/.test(p)), `${mode}：页脚不应留在正文里`);
    if (mode === "flow") {
      const foot = pkg.names().filter((n) => /^word\/footer\d*\.xml$/.test(n)).map((n) => pkg.readText(n)).find((t) => t.includes("保密文件")) ?? "";
      assert(foot, "可编辑排版：页脚应写进页脚部件");
      // "第 N 页 / 共 3 页"：页码与总页数都是域，重新分页、增删内容后仍然正确
      assert(/instr=" PAGE "/.test(foot) && /instr=" NUMPAGES "/.test(foot), "可编辑排版：页脚的页码、总页数应为 PAGE / NUMPAGES 域");
      // 悬挂缩进：编号后是制表符，续行与正文对齐
      const p11 = (xml.match(/<w:p [^>]*>[\s\S]*?<\/w:p>/g) ?? []).find((p) => />1\.1<\/w:t>/.test(p)) ?? "";
      assert(/w:hanging="\d+"/.test(p11) && p11.includes("<w:tab/>"), "可编辑排版：条款 1.1 应为悬挂缩进、编号后接制表符");
    }
  }
});

await test("页眉页脚只收 Word 能逐页再现的行：页首的条款标题、附件标题留在正文；页眉里的版本号不写成页码域；Bates 号不写成固定号码", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const norm = (s: string) => s.replace(/[\s\u3000\u00a0]/g, "");
  const parts = (pkg: InstanceType<typeof ZipPackage>, re: RegExp) => pkg.names().filter((n) => re.test(n)).map((n) => pkg.readText(n));
  const textOf = (xml: string) => norm(xml.replace(/<w:instrText[^>]*>[^<]*<\/w:instrText>/g, "").replace(/<[^>]+>/g, ""));
  // 第 2、3 页页首 "第4条"、"第5条"，第 4–6 页页首 "附件1"–"附件3"：数字逐页加一，但只在一页出现，是正文
  {
    const { data, report } = await pdfToDocx(await fixture("furniture-cn.pdf"), { mode: "flow" });
    eq(report.textCoverage, 1, `文字完整性，缺失：${report.missing}`);
    eq(report.notes.length, 0, `不应有页眉页脚说明：${report.notes.join(" | ")}`);
    const pkg = new ZipPackage(data);
    const paras = (pkg.readText("word/document.xml").match(/<w:p [^>]*>[\s\S]*?<\/w:p>/g) ?? []).map(textOf);
    for (const h of ["第4条", "第5条", "附件1", "附件2", "附件3"]) eq(paras.filter((p) => p === h).length, 1, `正文里应有单独成段的 "${h}"`);
    const heads = parts(pkg, /^word\/header\d*\.xml$/), feet = parts(pkg, /^word\/footer\d*\.xml$/);
    assert(heads.length && heads.every((h) => !/第\d条|附件/.test(textOf(h))), `页眉里不应有条款或附件标题：${heads.map(textOf).join(" | ")}`);
    // "第3版" 是版本号：恰好等于第 3 页的页码，也不写成 PAGE 域
    assert(heads.some((h) => textOf(h).includes("设备采购合同（第3版）")) && heads.every((h) => !/instr=" ?PAGE/.test(h)), "页眉应原样是 设备采购合同（第3版），不含页码域");
    assert(feet.some((f) => /instr=" PAGE "/.test(f) && /instr=" NUMPAGES "/.test(f)), "页脚的页码、总页数应为 PAGE / NUMPAGES 域");
  }
  // Bates 号逐页加一、却不是页码：Word 页脚只能写一个固定号码——不写进页脚、不回正文，notes 里说明
  {
    const { data, report } = await pdfToDocx(await fixture("bates.pdf"), { mode: "flow" });
    eq(report.textCoverage, 1, `文字完整性，缺失：${report.missing}`);
    const pkg = new ZipPackage(data);
    assert(!textOf(pkg.readText("word/document.xml")).includes("ABC000"), "Bates 号不应回到正文");
    assert(parts(pkg, /^word\/(header|footer)\d*\.xml$/).every((x) => !textOf(x).includes("ABC000")), "Bates 号不应写成固定号码的页脚");
    assert(parts(pkg, /^word\/header\d*\.xml$/).some((h) => textOf(h).includes("ProductionVolume1")), "页眉应保留");
    assert(report.notes.some((n) => n.includes("ABC0001234") && n.includes("没有写进页脚")), `notes 应说明 Bates 号没有写进页脚：${report.notes.join(" | ")}`);
  }
});

await test("页眉页脚取舍（settleFurniture）：另一节的页眉三页以上不回正文并说明、两页时留在正文；完整性校验只认 Word 实际显示的页眉页脚", async () => {
  const { settleFurniture } = await import("../src/services/pdfConvert/furniture.js");
  const line = (text: string, top: number) => ({ spans: [{ text, x0: 72, x1: 72 + 5 * text.length, size: 9 }], x0: 72, x1: 72 + 5 * text.length, top, bottom: top + 9, baseline: top + 7, size: 9 }) as any;
  const doc = (n: number, sched: number[]) => Array.from({ length: n }, (_, i) => [
    line(sched.includes(i) ? "Schedule 1 — Pricing" : "Master Services Agreement", 30),
    line(`Page ${i + 1}`, 780),
  ]);
  const three = settleFurniture(doc(8, [5, 6, 7]), Array(8).fill(396), [1, 2, 3, 4, 5, 6, 7, 8]);
  eq(three.hidden.flat().map((l: any) => l.spans[0].text).join("|"), "Schedule 1 — Pricing|Schedule 1 — Pricing|Schedule 1 — Pricing", "另一节的页眉（三页）：写不进去，不回正文");
  assert(three.notes.length === 1 && three.notes[0].includes("第 6–8 页的页眉"), `应说明第 6–8 页的页眉没有写进去：${three.notes.join(" | ")}`);
  eq(three.keep.flat().length, 8 + 5, "其余页的页眉与全部页码写进页眉页脚");
  eq([...three.pageNumberLines].length, 8, "只有页码行带页码位");
  const two = settleFurniture(doc(8, [6, 7]), Array(8).fill(396), [1, 2, 3, 4, 5, 6, 7, 8]);
  eq(two.hidden.flat().length, 0, "只有两页的不同页眉：证据不足，留在正文");
  eq(two.keep.flat().length, 6 + 8, "两页的不同页眉既不写进页眉，也不隐藏");
  eq(two.notes.length, 0, "留在正文的不需要说明");
});

await test("扫描件：整页扫描图按页面坐标铺满原页（不再 3 页变 6 页）；OCR 隐藏文字层不转成正文；上下两条拼成的整页", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("scan-letter.pdf");
  for (const mode of ["exact", "flow"] as const) {
    const { data, report } = await pdfToDocx(pdf, { mode });
    const xml = new ZipPackage(data).readText("word/document.xml");
    eq((xml.match(/<wp:inline/g) ?? []).length, 0, `${mode}：整页扫描图不应作为行内图片（比版心高，会被挤到下一页）`);
    eq((xml.match(/<pic:pic>/g) ?? []).length, 4, `${mode}：4 张扫描图（第 3 页上下两条）都保留`);
    eq((xml.match(/<wp:positionV relativeFrom="page">/g) ?? []).length, 4, `${mode}：扫描图按页面坐标定位`);
    // 第 2、3 页各从新的一页开始（可编辑排版下扫描页没有文字流，不强制分页会叠在同一页上）
    eq((xml.match(/w:val="PageStart"/g) ?? []).length, 2, `${mode}：每页扫描图各占一页`);
    // OCR 识别层看不见：不转成正文（否则与图里的字重叠成两层）
    eq((xml.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, "")).join("").trim(), "", `${mode}：OCR 识别层不应转成正文`);
    assert(report.notes.some((n) => n.includes("OCR")), `${mode}：报告应说明 OCR 识别层未转换：${report.notes.join(" | ")}`);
  }
});

await test("扫描件首页（看得见的 OCR 文字层）：著录项目的两栏不被栏内空隙、网格对齐的代码清单搅乱；大号公开号不吞并两侧小字；与左栏同高、跨过栏间的徽标碎字按原位置成文本框；OCR 标点字框压住下一个字不拆行", async () => {
  const { buildLines, detectGutter, buildDocModel } = await import("../src/services/pdfConvert/layout.js");
  // OCR 文字的字框偏高（上沿约一个字号）
  const sp = (text: string, x0: number, x1: number, bl: number, size: number) => ({ text, x0, x1, top: bl - size, bottom: bl + 0.27 * size, baseline: bl, size, font: "MS-Mincho", bold: false, italic: false, color: "000000" });
  const text = (p: any) => p.runs.map((r: any) => (r.tab ? "\t" : r.br ? "\n" : r.text)).join("");

  // 标点的字框压住下一个字（"（" 压在 "S" 上、"：" 压在 "frame" 上）仍是同一行：拆成两行时 Word 里只能一上一下，
  // 后面的内容整块下移；同一个标点叠在自己的副本上（重复绘制的文字层）仍要拆开
  eq(buildLines([sp("ステップ", 115, 160, 605, 10), sp("（", 165.2, 175.4, 605, 10), sp("S", 171.7, 178.5, 605, 10), sp("1 0 8", 185.4, 213.6, 605, 10), sp("）", 215, 225, 605, 10)] as any).length, 1, "（S108）应在同一行");
  eq(buildLines([sp("input image", 335.5, 375, 616, 6.5), sp("：", 381, 389.3, 616, 6.5), sp("frame so that the subject", 381.6, 470, 616, 6.5)] as any).length, 1, "input image ：frame … 应在同一行");
  const dup = buildLines([sp("abc", 50, 70, 100, 10), sp("abc", 50.2, 70.2, 100, 10), sp("，", 70, 76, 100, 10), sp("，", 70.2, 76.2, 100, 10)] as any);
  assert(dup.length === 2 && dup.every((l) => l.spans.filter((s) => s.text === "，").length === 1), `重复绘制的两层文字各成一行、各带一个逗号：${dup.map((l) => l.spans.map((s) => s.text).join("")).join(" | ")}`);

  // 栏间距：右栏是按网格对齐的国家代码清单（第一个代码与后面的代码之间连成一条竖向空白，两侧都有文字的行比真正的
  // 栏间距还多），左右两栏基线不对齐。真正的栏间距只被个别行跨过；网格空白则被每一个标题行跨过、且左侧只有一个代码
  const codes = "AE AG AL AM AO AT AU AZ BA BB BG BH BN BR BW BY BZ CA CH CL CN CO CR CU CZ DE DK DM DO DZ EC EE EG ES FI GB GD GE".split(" ");
  const gs: any[] = [];
  for (let i = 0; i < 10; i++) gs.push(sp(`(${21 + i})`, 74.2, 90, 160 + 16.3 * i, 6.5), sp("国際出願番号：", 96, 160, 160 + 16.3 * i, 8.5), sp("PCT/JP2099/012345", 240, 309.6, 160 + 16.3 * i, 6.5));
  for (let i = 0, c = 0; i < 24; i++) {
    const bl = 164.5 + 10.2 * i;
    if (i % 6 === 0) gs.push(sp(`(8${i % 10})`, 318.5, 336, bl, 6.5), sp("指定国（表示のない限り、全ての種類の国内保", 340, 552, bl, 8.5));
    else for (let k = 0; k < 12; k++) gs.push(sp(codes[c++ % codes.length] + ",", 339 + 18 * k, 350 + 18 * k, bl, 6.5));
  }
  const g = detectGutter(buildLines(gs), { x0: 74.2, x1: 552 });
  assert(g && Math.abs(g.x - 314) < 3, `栏间距应在两栏之间（x≈314），实际：${JSON.stringify(g)}`);

  // 整页：页首左栏两行小字、右栏大号公开号（纵跨左边两行）；徽标下的 "W 画"（栏间左侧）与 "P O | P C T"（跨过栏间距）
  // 与左栏的日期同高；下面是左右两栏的著录项目。正文行从 x = 72 起，标签行从 74.2 起（编号前有 2pt 的缩进）
  const spans: any[] = [
    sp("(43)", 128.2, 143, 116.6, 8.5), sp("国際公開日", 147, 197.3, 116.6, 8.5),
    sp("2099", 95.5, 120, 128.9, 9.5), sp("年", 124, 133.5, 128.9, 9.5), sp("3", 137, 142, 128.9, 9.5), sp("月", 146, 155.5, 128.9, 9.5), sp("日", 170, 179.5, 128.9, 9.5), sp("（09.03.2099)", 183, 226.4, 128.9, 9.5),
    sp("(10)", 406.8, 422, 109, 8.5), sp("国際公開番号", 426, 484.6, 109, 8.5), sp("W O 2099/012345 A 1", 376.1, 517.1, 125.5, 14),
    sp("W", 251.8, 258.4, 134.6, 8.5), sp("画", 262.6, 275.8, 134.6, 8.5), sp("P", 272.9, 279.6, 134.6, 9.5), sp("O", 282.5, 289.5, 134.6, 9.5), sp("|", 296, 298, 134.6, 9.5), sp("P C T", 305.3, 331.9, 134.6, 9.5),
  ];
  for (let i = 0; i < 12; i++) spans.push(sp(`(${21 + i})`, 74.2, 89, 161.1 + 16.4 * i, 6.5), sp("国際出願番号：", 96, 160, 161.1 + 16.4 * i, 8.5), sp("PCT/JP2099/012345", 240, 309.6, 161.1 + 16.4 * i, 6.5));
  for (let i = 0; i < 18; i++) spans.push(...(i % 6 === 0 ? [sp(`(7${i % 10})`, 318.5, 334, 160.6 + 10.3 * i, 6.5)] : []), sp("東京都新宿区西新宿 7 丁目 5 番 25 号 西新宿木村屋ビル 9 階", 339.4, 552.7, 160.6 + 10.3 * i, 8.5));
  // 右栏一行里 OCR 多认出的一个字，与 "区" 叠在一起（同一高度、不同的行）
  spans.push(sp("巨", 405, 413, 160.6 + 10.3 * 3 + 0.4, 8.5));
  for (let i = 0; i < 7; i++) spans.push(sp("東京都港区港南 1 丁目 7 番 1 号 ソニー株式会社内 Tokyo (JP).", 72, 300, 365 + 10.2 * i, 8.5));
  const page = { index: 0, width: 595, height: 842, spans, segs: [], fills: [], images: [], vectorShapes: 0, paths: [], masks: [], rotatedText: 0 };
  const model = buildDocModel([page as any], { mode: "exact" });
  eq(model.stats.columnPages, 1, "著录项目应识别为两栏");
  const paras = model.blocks.filter((b: any) => b.kind === "p") as any[];
  const flow = paras.map(text);
  const boxes = paras.flatMap((p) => (p.textBoxes ?? []).map((t: any) => text(t.para)));
  assert(flow.includes("(43) 国際公開日") && flow.includes("W O 2099/012345 A 1"), `大号公开号与左边的小字各自成行：${flow.slice(0, 8).join(" | ")}`);
  for (let i = 21; i <= 32; i++) assert(flow.some((t) => t.includes(`(${i})`) && t.includes("国際出願番号：") && t.includes("PCT/JP2099/012345")), `(${i}) 的标签与值在左栏同一段：${flow.filter((t) => t.includes(`(${i})`)).join(" | ")}`);
  assert(!flow.some((t) => t.includes("PCT/JP") && t.includes("東京都新宿区")), "左栏的值不应混进右栏");
  // 与左栏同高的碎字按原位置成文本框；排进文字流只能一前一后，后面的两栏整块下移，压到底图里的横线上
  assert(boxes.includes("W 画") && boxes.includes("P O | P C T") && !flow.some((t) => /画|P C T/.test(t)), `徽标碎字应是文本框：框 ${boxes.join(" | ")}；正文 ${flow.filter((t) => /画|P C T/.test(t)).join(" | ")}`);
  assert(boxes.includes("巨") && !flow.some((t) => t.includes("巨")), `栏内与正文行叠在一起的碎字应是文本框：框 ${boxes.join(" | ")}`);
  // 编号 6.5 pt、标签 8.5 pt 混排的单行：行高按最大的字号，不能写成"最小值"行距（排版时按字体实际高度撑开、逐行下移）
  for (const p of paras.filter((p) => !p.tiny && p.runs.length && !p.runs.some((r: any) => r.br))) {
    const big = Math.max(...p.runs.filter((r: any) => !r.tab && r.text?.trim()).map((r: any) => r.size ?? 0));
    assert(big <= 1.08 * p.lineHeight, `单行段落的行高放得下最大的字：${JSON.stringify(text(p))} 行高 ${p.lineHeight} 字号 ${big}`);
  }
  // 编号行（单行、编号后接制表符）：缩进改写成制表位后，编号后的制表位也要保留，否则 "国際出願番号：" 跳到行尾的右对齐制表位上
  for (const p of paras.filter((p) => text(p).includes("国際出願番号：") && !text(p).includes("\n") && p.indLeft === 0 && p.firstLine === 0)) {
    const tabRuns = p.runs.filter((r: any) => r.tab).length;
    assert(tabRuns <= p.tabs.length, `每个制表符都有制表位：${JSON.stringify(text(p))} 制表位 ${JSON.stringify(p.tabs)}`);
  }
});

// ---------------------------------------------------------------------------
console.log("\n旧版 Word（.doc）");
// ---------------------------------------------------------------------------

const legacyDoc = await fixture("contract-features-cn.doc");
/** 粗略的良构检查：标签成对、嵌套正确 */
function wellFormed(xml: string): string | null {
  const stack: string[] = [];
  for (const m of xml.replace(/<\?[\s\S]*?\?>|<!--[\s\S]*?-->/g, "").matchAll(/<(\/?)([\w:.-]+)(?:\s[^>]*?)?(\/?)>/g)) {
    if (m[3]) continue;
    if (!m[1]) stack.push(m[2]);
    else if (stack.pop() !== m[2]) return `</${m[2]}> 不匹配`;
  }
  return stack.length ? `<${stack[stack.length - 1]}> 未闭合` : null;
}

await test(".doc 内置引擎：正文文字一字不差；自动编号、表格合并与底纹、脚注、批注、修订、行内 / 浮动图片、页眉页脚域、横向第二节都转成 Word 原生结构", async () => {
  const { docToDocx } = await import("../src/services/docConvert/index.js");
  const { data, report } = await docToDocx(legacyDoc, { engine: "builtin" });
  eq(report.engine, "builtin", "引擎");
  assert(report.chars > 1000 && report.written === report.chars, `文字完整性：${report.written}/${report.chars}`);
  eq(report.notes.join("；"), "", "不应有未转换的内容");
  const pkg = new ZipPackage(data);
  for (const n of pkg.names().filter((x) => /\.(xml|rels)$/.test(x))) {
    const bad = wellFormed(pkg.readText(n));
    assert(!bad, `${n} 不是良构的 XML：${bad}`);
  }
  const xml = docXml(data);
  const text = docText(data);
  for (const t of ["设备采购合同", "合同总价款为人民币[ ]元（大写：[ ]）", "本合同一式肆份，甲乙双方各执贰份，具有同等法律效力。", "附件二：交货进度表（横向页面）", "三坐标测量机"]) {
    assert(text.includes(t), `正文应包含「${t}」`);
  }
  // 自动编号：第一条 / 1.1 / (a)，仍是 Word 编号（不是写死的文字）
  const numbering = pkg.readText("word/numbering.xml");
  for (const lvl of ["第%1条", "%1.%2", "(%3)"]) assert(numbering.includes(`<w:lvlText w:val="${lvl}"/>`), `编号格式 ${lvl}`);
  assert(count(xml, "<w:numPr>") >= 30, `编号段落：${count(xml, "<w:numPr>")}`);
  assert(!/>第[一二三四五六]条</.test(xml), "条款编号不应写成文字");
  // 黄色高亮的方括号占位符
  eq(count(xml, '<w:highlight w:val="yellow"/>'), 2, "占位符高亮");
  // 表格：表头底纹、合计行横向合并、纵向合并
  eq(count(xml, "<w:tbl>"), 2, "表格个数");
  eq(count(xml, 'w:fill="D9D9D9"'), 4, "表头底纹");
  assert(xml.includes('<w:gridSpan w:val="3"/>'), "合计行横向合并三列");
  assert(xml.includes('<w:vMerge w:val="restart"/>') && /<w:vMerge\/>/.test(xml), "纵向合并");
  // 脚注、批注、修订
  assert(pkg.readText("word/footnotes.xml").includes("含增值税，税率按开票时适用的法定税率计算。"), "脚注内容");
  eq(count(xml, "<w:footnoteReference "), 1, "脚注引用");
  const comments = pkg.readText("word/comments.xml");
  assert(comments.includes("请客户确认份数。") && comments.includes('w:author="李律师"'), "批注内容与作者");
  assert(xml.includes("<w:commentRangeStart ") && xml.includes("<w:commentReference "), "批注范围");
  assert(/<w:del [^>]*w:author="张律师"[^>]*>[\s\S]*?<w:delText[^>]*>有效期一年。<\/w:delText>/.test(xml), "删除修订");
  assert(/<w:ins [^>]*w:author="张律师"[^>]*>[\s\S]*?有效期至双方义务履行完毕之日止。/.test(xml), "插入修订");
  // 图片：行内一张，浮动（盖章，压在文字上）一张
  eq(count(xml, "<pic:pic"), 2, "图片个数");
  eq(count(xml, "<wp:inline"), 1, "行内图片");
  eq(count(xml, "<wp:anchor "), 1, "浮动图片");
  assert(/<wp:anchor [^>]*behindDoc="0"[\s\S]*?<wp:wrapNone\/>/.test(xml), "盖章浮于文字上方、不绕排");
  // 页眉（下边框）与页脚（第 PAGE 页 共 NUMPAGES 页）
  const parts = pkg.names();
  const header = parts.filter((n) => /^word\/header\d+\.xml$/.test(n)).map((n) => pkg.readText(n)).find((t) => t.includes("北京某某律师事务所")) ?? "";
  assert(header && /<w:pBdr>.*<w:bottom w:val="single"/.test(header), "页眉文字与下边框");
  const footer = parts.filter((n) => /^word\/footer\d+\.xml$/.test(n)).map((n) => pkg.readText(n)).find((t) => t.includes(" PAGE ")) ?? "";
  assert(footer.includes(" NUMPAGES "), "页脚的页码与总页数是域");
  // 两节：第二节横向
  eq(count(xml, "<w:sectPr"), 2, "节数");
  assert(/<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/.test(xml), `第二节横向 A4：${xml.match(/<w:pgSz [^>]*>/g)}`);
  // 中文字体；语言标记与原文一致（影响中文断行，与原文不同会导致换行位置、页数变化）
  assert(xml.includes('w:eastAsia="宋体"') && pkg.readText("word/styles.xml").includes("宋体"), "中文字体");
  assert(/<w:lang w:val="en-US" w:eastAsia="en-US"/.test(xml + pkg.readText("word/styles.xml")), "语言标记");
});

await test("上传 .doc：自动转换为 Word 会话，上传的原件逐字节保留；编辑、保真校验、回滚、重新加载与普通 Word 一致", async () => {
  const prev = process.env.DOC_ENGINE;
  process.env.DOC_ENGINE = "builtin";
  try {
    const s = await Session.create(legacyDoc, "设备采购合同.doc", { mode: "bypassPermissions", trackChanges: false, author: "测试" });
    eq(s.meta.format, "docx", "会话格式");
    eq(s.meta.filename, "设备采购合同.docx", "文件名");
    eq(s.meta.sourceFile, "设备采购合同.doc", "原文件名");
    eq(s.meta.sourceExt, "doc", "原件格式");
    eq(s.meta.origin?.engine, "builtin", "转换引擎");
    eq(s.meta.origin?.coverage, 1, "文字完整性");
    assert(/Word 97-2003/.test(s.meta.origin!.report) && s.meta.origin!.report.includes("文字完整性 100.0%"), `转换报告：${s.meta.origin!.report}`);
    assert(s.meta.versions[0].label.includes("设备采购合同.doc"), "v0 标注来源");
    const up = await s.uploaded();
    eq(up.ext, "doc", "原始文件扩展名");
    eq(up.filename, "设备采购合同.doc", "原始文件名");
    assert(up.data.equals(legacyDoc), "上传的 .doc 应逐字节保留");
    // 模型知道文件来源
    const { buildTurnReminder } = await import("../src/context.js");
    assert(buildTurnReminder(s).includes("旧版 Word 97-2003"), "当轮提醒应说明文件由 .doc 转换而来");
    // 编辑：与普通 .docx 一样，未修改的部件保持不变
    s.doc.replaceText({ oldText: "本合同一式肆份", newText: "本合同一式陆份" }, O(true));
    const fid = s.doc.fidelity(await s.original());
    assert(fid.ok, `保真校验：${fid.problems.join("；")}`);
    assert(fid.modifiedParts.every((p) => p.name === "word/document.xml"), `只应改动正文：${fid.modifiedParts.map((p) => p.name)}`);
    await s.commit("改份数", []);
    const re = await Session.load(s.id);
    assert(docText(re.current()).includes("本合同一式陆份"), "重新加载后保留修改");
    assert((await re.uploaded()).data.equals(legacyDoc), "重新加载后原件仍可下载");
    await re.rollback(0);
    assert(docText(re.current()).includes("本合同一式肆份"), "回滚到转换结果");
    eq(re.meta.format, "docx", "回滚后仍是 Word");
  } finally {
    if (prev === undefined) delete process.env.DOC_ENGINE; else process.env.DOC_ENGINE = prev;
  }
});

await test(".doc 异常情况：加密、Word 6 / 95、不是 Word 的文件给出明确提示；扩展名写错的 .docx 直接打开；有 LibreOffice 时可导出 .doc", async () => {
  const { convertUploadedDoc, DocFileError } = await import("../src/services/docConvert/index.js");
  const rejects = async (fn: () => Promise<unknown>, re: RegExp, what: string) => {
    let err: any;
    try { await fn(); } catch (e) { err = e; }
    assert(err instanceof DocFileError && re.test(err.message), `${what}：${err?.message ?? "没有报错"}`);
  };
  // FIB 在 WordDocument 流开头（扇区边界上）
  let fib = -1;
  for (let o = 512; o + 12 < legacyDoc.length; o += 512) if (legacyDoc.readUInt16LE(o) === 0xa5ec && legacyDoc.readUInt16LE(o + 2) >= 0xc1) { fib = o; break; }
  assert(fib > 0, "应找到 FIB");
  const encrypted = Buffer.from(legacyDoc);
  encrypted.writeUInt16LE(encrypted.readUInt16LE(fib + 0x0a) | 0x0100, fib + 0x0a);
  await rejects(() => convertUploadedDoc(encrypted, "a.doc", { engine: "builtin" }), /密码/, "加密文档");
  const word95 = Buffer.from(legacyDoc);
  word95.writeUInt16LE(0x65, fib + 2);
  await rejects(() => convertUploadedDoc(word95, "a.doc", { engine: "builtin" }), /Word 6\.0 \/ 95/, "Word 95 文档");
  await rejects(() => convertUploadedDoc(Buffer.from("{\\rtf1\\ansi hello}"), "a.doc", { engine: "builtin" }), /不是 Word 97-2003 文档/, "不是 Word 文件");
  // 其实是 .docx：不转换，原样打开
  const s = await Session.create(sampleDocx, "论文.doc", { mode: "default", trackChanges: false, author: "测试" });
  eq(s.meta.filename, "论文.docx", "按 .docx 打开");
  assert(s.current().equals(sampleDocx), "内容不经转换");
  assert(s.meta.origin!.report.includes("扩展名写成了 .doc"), `报告：${s.meta.origin!.report}`);
  // 导出为 .doc（LibreOffice 另存为 Word 97-2003），导出的文件内置引擎能读回、文字不丢
  const { checkExternalTools, exportDocument } = await import("../src/services/convert.js");
  if ((await checkExternalTools()).soffice) {
    const { docToDocx } = await import("../src/services/docConvert/index.js");
    const src = (await docToDocx(legacyDoc, { engine: "builtin" })).data;
    const out = await exportDocument(src, "docx", "doc");
    eq(out.ext, "doc", "扩展名");
    eq(out.data.readUInt32LE(0), 0xe011cfd0, "导出的是 OLE2 复合文档");
    const back = await docToDocx(out.data, { engine: "builtin" });
    eq(back.report.written, back.report.chars, "导出的 .doc 文字完整");
    assert(docText(back.data).includes("本合同一式肆份"), "导出的 .doc 能读回");
  }
});

await test(".doc 版面细节：图文框（页脚里定位的页码）、东亚排版开关、字体提示、字距调整、兼容性选项与 Word 自己的转换一致", async () => {
  const { paraXml, charXml } = await import("../src/services/docConvert/props.js");
  const { S } = await import("../src/services/docConvert/sprm.js");
  const { readCompat } = await import("../src/services/docConvert/parse.js");
  const sp = (op: number, hex: string) => ({ op, arg: Buffer.from(hex, "hex") });
  // 一份真实合同页脚里的页码段落：右侧定位的图文框，文字环绕（Word 另存为 .docx 的结果是同样的 framePr）
  const ppr = paraXml([sp(S.PDxaAbs, "3e23"), sp(S.PDyaAbs, "6100"), sp(S.PDxaWidth, "b005"), sp(S.PPc, "a0"), sp(S.PWr, "02"),
    sp(S.PFOverflowPunct, "00"), sp(S.PFAdjustRight, "00"), sp(S.PWAlignFont, "0200")], { styleIds: new Map(), numId: () => undefined });
  assert(ppr.includes('<w:framePr w:w="1456" w:wrap="around" w:hAnchor="page" w:vAnchor="text" w:x="9022" w:y="97"/>'), `图文框：${ppr}`);
  assert(/<w:framePr [^>]*\/><w:overflowPunct w:val="0"\/><w:adjustRightInd w:val="0"\/><w:textAlignment w:val="baseline"\/>/.test(ppr), `东亚排版开关与元素顺序：${ppr}`);
  // 对齐方式：XAS / YAS 的负数
  const centered = paraXml([sp(S.PDxaAbs, "fcff"), sp(S.PDyaAbs, "f4ff"), sp(S.PPc, "50")], { styleIds: new Map(), numId: () => undefined });
  assert(centered.includes('w:hAnchor="margin" w:vAnchor="page" w:xAlign="center" w:yAlign="bottom"'), `对齐方式：${centered}`);
  const rpr = charXml([sp(S.CRgFtc0, "0000"), sp(S.CIdctHint, "01"), sp(S.CHpsKern, "0200")], { fonts: ["Times New Roman"], styleIds: new Map(), styleToggle: () => false });
  assert(rpr.includes('<w:rFonts w:ascii="Times New Roman" w:hint="eastAsia"/>') && rpr.includes('<w:kern w:val="2"/>'), `字体提示与字距：${rpr}`);
  // 兼容性选项：同一份合同的 Dop（Word 2007 写的），Word 另存为 .docx 时输出的就是这些开关
  const dop = Buffer.alloc(674);
  dop.writeUInt32LE(0x0010f000, 84); dop.writeUInt32LE(0x0010f000, 508); dop.writeUInt32LE(0xfdfc0008, 512); dop.writeUInt32LE(1, 516);
  eq(readCompat(dop).join(","), "useNormalStyleForList,doNotUseIndentAsNumberingTabStop,useAltKinsokuLineBreakRules,allowSpaceOfSameStyleInTable," +
    "doNotSuppressIndentation,doNotAutofitConstrainedTables,autofitToFirstFixedWidthCell,displayHangulFixedWidth,splitPgBreakAndParaMark," +
    "doNotVertAlignCellWithSp,doNotBreakConstrainedForcedTable,doNotVertAlignInTxbx,useAnsiKerningPairs,cachedColBalance", "兼容性选项");
  dop.writeUInt32LE(0, 508);
  eq(readCompat(dop).length, 0, "copts80 对不上时不读（不是 Word 2000 以后的 Dop）");
});

await test(".doc 形状组合：矩形、带箭头的连接线、自由曲线都转成 Word 的组合图形，组内坐标换算成 EMU", async () => {
  const { docToDocx } = await import("../src/services/docConvert/index.js");
  const { data, report } = await docToDocx(await fixture("shapes-group.doc"), { engine: "builtin" });
  const xml = docXml(data);
  eq(count(xml, "<wpg:wgp>"), 1, "形状组合");
  const group = xml.slice(xml.indexOf("<wpg:wgp>"), xml.indexOf("</wpg:wgp>"));
  eq(count(group, "<wps:wsp>"), 3, "组内三个形状");
  assert(/<a:chOff x="0" y="0"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/.test(group), "组的坐标系与外框一致（EMU）");
  assert(/<a:custGeom>.*<a:moveTo><a:pt x="0" y="100"\/><\/a:moveTo><a:lnTo><a:pt x="50" y="0"\/><\/a:lnTo><a:lnTo><a:pt x="100" y="100"\/><\/a:lnTo>/.test(group), "自由曲线三角形的路径");
  assert(group.includes('<a:prstGeom prst="line">') && group.includes('<a:tailEnd type="triangle"'), "带箭头的连接线");
  assert(group.includes('<a:srgbClr val="00FF00">'), "填充色");
  assert(report.shapes >= 4 && !report.notes.some((n) => n.includes("未转换")), `报告：${report.shapes} 个图形；${report.notes.join("；")}`);
  for (const t of ["示意图如下。", "图后的正文。"]) assert(docText(data).includes(t), `正文：${t}`);
});

await test("PWB 回归（pwb-cases.py）：左对齐正文不在短行处断段、跨页段落合并；补零的 Bates 号写成页码域；表格竖线两侧的文字不并格；/Rotate 页按文字方向转换；首页带引用号的脚注留在正文", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const plain = (xml: string) => (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [])
    .map((p) => p.replace(/<w:instrText[^>]*>[^<]*<\/w:instrText>/g, "").replace(/<w:tab\/>/g, "\t").replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim()).filter(Boolean);
  const footers = (pkg: InstanceType<typeof ZipPackage>) => pkg.names().filter((n) => /^word\/footer\d*\.xml$/.test(n)).map((n) => pkg.readText(n));
  // 左对齐正文排在留有内边距的文本框里，页眉横线、页脚伸到页边：区域右边界比折行处靠右；第 7 段跨页
  {
    const { data, report } = await pdfToDocx(await fixture("ragged.pdf"), { mode: "flow" });
    eq(report.textCoverage, 1, `文字完整性，缺失：${report.missing}`);
    const pkg = new ZipPackage(data);
    const paras = plain(pkg.readText("word/document.xml"));
    const body = paras.filter((p) => p.length > 200);
    eq(body.length, 12, `12 段正文各成一段：${paras.map((p) => p.slice(-24)).join(" | ")}`);
    assert(body.every((p) => /\.$/.test(p)), "每段都以句号结束（没有在句子中间断开）");
    assert(footers(pkg).some((f) => /ALX<\/w:t>[\s\S]*?00000<\/w:t>[\s\S]*?instr=" PAGE "/.test(f)), "Bates 号写成 ALX00000 + PAGE 域");
    assert(!paras.some((p) => p.includes("ALX00000")), "Bates 号不回正文");
  }
  // 说明列的文字离右边框线很近，pdf.js 把它和日期当作词间空格合成一项
  {
    const { data } = await pdfToDocx(await fixture("cells.pdf"), { mode: "flow" });
    const rows = (docXml(data).match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) ?? []).map((r) => (r.match(/<w:tc[ >][\s\S]*?<\/w:tc>/g) ?? []).map((c) => c.replace(/<[^>]+>/g, "").trim()));
    for (let i = 1; i <= 5; i++) {
      assert(rows.some((r) => r[1] === `Delivery of work package ${i} and acceptance testing` && r[2] === `2027-0${i}-14`), `第 ${i} 行：说明与日期应各在一格：${JSON.stringify(rows[i])}`);
    }
  }
  // 第 2 页带 /Rotate 90：按文字自己的方向读，正文不当成旋转文字丢掉，也不按第一页（横向）的版心判断折行
  {
    const { data, report } = await pdfToDocx(await fixture("rotated.pdf"), { mode: "flow" });
    eq(report.textCoverage, 1, `文字完整性，缺失：${report.missing}`);
    assert(!report.notes.some((n) => n.includes("旋转文字")), `不应有旋转文字未转换：${report.notes.join(" | ")}`);
    eq(plain(docXml(data)).filter((p) => p.length > 200).length, 4, "第 2 页 4 段正文各成一段");
  }
  // 各页页底的脚注都留在本页正文末尾，不只把第一页的移进首页页脚
  {
    const { data } = await pdfToDocx(await fixture("footnotes.pdf"), { mode: "flow" });
    const pkg = new ZipPackage(data);
    assert(footers(pkg).every((f) => !f.includes("See clause")), "脚注不进页脚");
    const paras = plain(pkg.readText("word/document.xml"));
    for (let k = 1; k <= 9; k++) eq(paras.filter((p) => p.startsWith(`${k} See clause`)).length, 1, `脚注 ${k} 在正文里单独成段`);
  }
});

await test("按制表位对齐的清单（附图标记说明：每行几组“部件名，编号”）不是分栏：隔几行才有两行的空隙上下对齐，不切成左右两栏；通栏内容多于双栏内容、但双栏成片时仍是分栏", async () => {
  const { buildLines, detectGutter } = await import("../src/services/pdfConvert/layout.js");
  const sp = (text: string, x0: number, x1: number, bl: number, size = 10) => ({ text, x0, x1, top: bl - size, bottom: bl + 0.27 * size, baseline: bl, size, font: "MS-Mincho", bold: false, italic: false, color: "000000" });
  // 12 行清单：每 4 行里有 2 行在 x≈324 处留空（两侧都有文字），另外 2 行有一组“部件名”横跨这里。
  // 当成分栏后这两行被拆成左右两栏，读成“理装置，13 统合…｜ラ装置，30…｜20，20_1…カメ｜雲台，40…”，行序错乱
  const list: any[] = [];
  for (let r = 0; r < 12; r++) {
    const bl = 101.3 + 24 * r;
    if (r % 4 < 2) list.push(sp("理装置，", 113.8, 160.4, bl), sp("1 3", 177.4, 193.4, bl), sp("統合映像処理装置，", 210, 316.4, bl), sp("2 0 ， 2 0 _ 1乃至2 0 — N", 332.2, 483.1, bl), sp("カ メ", 498.7, 520, bl));
    else list.push(sp("映像処理サーバ，", 126.2, 219.5, bl), sp("6 0", 236.2, 253.4, bl), sp("情報端末装置，", 269.8, 352.4, bl), sp("7 0", 368.4, 385.4, bl), sp("ネ ッ トワ ー ク ，", 402.7, 483.5, bl), sp("8 0", 499.9, 517.4, bl));
  }
  eq(detectGutter(buildLines(list as any), { x0: 113.8, x1: 520 }), null, "清单不是分栏");
  // 通栏的摘要 14 行 + 下面双栏正文 8 行：横跨栏间的行比双栏的行多，但双栏的 8 行上下相连
  const mixed: any[] = [];
  for (let r = 0; r < 14; r++) mixed.push(sp("この技術は映像処理装置に関するものであり入力画像フレームを処理する方法を提供する", 72, 523, 100 + 14 * r));
  for (let r = 0; r < 8; r++) mixed.push(sp("左の段の本文がここに続いて書かれている行", 72, 288, 320 + 14 * r), sp("右の段の本文がここに続いて書かれている行", 307, 523, 320 + 14 * r));
  const g = detectGutter(buildLines(mixed as any), { x0: 72, x1: 523 });
  assert(g && g.g0 <= 289 && g.g1 >= 306, `成片的双栏仍应识别：${JSON.stringify(g)}`);
});

await test("文字替换式扫描件（scan-text.pdf，WIPO 专利小册子的做法）：整页扫描图拆开，附图成图、只在图里的字留在原来的句子里、每页相同的页眉写进 Word 页眉；表单页格线垫底、文字放进原位文本框；没有文字层的页分清是附图还是没做识别的文字，并给调用方标出来", async () => {
  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  const pdf = await fixture("scan-text.pdf");
  const paraText = (p: string) => p.replace(/<wp:inline[\s\S]*?<\/wp:inline>/g, "[IMG]").replace(/<[^>]+>/g, "");
  const read = async (mode: "exact" | "flow") => {
    const { data, report } = await pdfToDocx(pdf, { mode });
    const pkg = new ZipPackage(data);
    const xml = pkg.readText("word/document.xml");
    const headers = pkg.names().filter((n) => /^word\/header\d*\.xml$/.test(n)).map((n) => pkg.readText(n)).join("");
    return { report, xml, headers, paras: (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).map(paraText), media: pkg.names().filter((n) => n.startsWith("word/media/")).length };
  };
  const flow = await read("flow");
  eq(flow.report.textCoverage, 1, "文字完整性");
  // 表单页（第 6 页）的文字仍是文字：算文字页，不算整页成图
  eq(JSON.stringify(flow.report.scanPages), JSON.stringify({ text: [1, 2, 3, 6], picture: [4], unread: [5] }), "各页的归类");
  // 文字页不再把整页扫描图垫在文字下面；只有表单页的格线图垫在文字下面
  eq((flow.xml.match(/behindDoc="1"/g) ?? []).length, 1, "只有表单页垫底");
  // 只在图里的字：小图片插在原来那句话的中间，前后文字不断开，并标上 pdf2word-glyph
  assert(flow.paras.some((t) => /document number\s*\[IMG\]\s*of the series/.test(t)), `第 1 页句中的字应留在原行：${flow.paras.join(" | ").slice(0, 600)}`);
  assert(flow.paras.some((t) => /according to claim\s*\[IMG\]\s*, wherein the frames are stored/.test(t)), "第 2 页第一处权利要求编号应留在原行");
  assert(flow.paras.some((t) => /according to claim\s*\[IMG\]\s*, wherein the angle is corrected/.test(t)), "第 2 页第二处权利要求编号应留在原行");
  eq((flow.xml.match(/name="pdf2word-glyph \d+"/g) ?? []).length, 3, "句中的字标成 pdf2word-glyph");
  // 没有文字层、看得出是文字的第 5 页：整页图标成 pdf2word-page 5，调用方识别后可以替换
  eq((flow.xml.match(/name="pdf2word-page 5"/g) ?? []).length, 1, "没读出来的页标成 pdf2word-page 5");
  assert(!/pdf2word-page 4/.test(flow.xml), "附图页不标");
  // 表单页：每行文字一个原位文本框；垫底的格线图标成 pdf2word-form 6（格里的字已是文字，图不必再翻译）
  eq((flow.xml.match(/name="pdf2word-form 6"/g) ?? []).length, 1, "表单页的格线图标成 pdf2word-form 6");
  for (const line of ["Box No. 1", "Box No. 12", "The applicant confirms the statement made in this box of the form"]) {
    assert(new RegExp(`<w:txbxContent>(?:(?!</w:txbxContent>)[\\s\\S])*${line.replace(/\./g, "\\.")}`).test(flow.xml), `表单文字应在文本框里：${line}`);
  }
  // 每页相同的页眉块（公开号条）写进 Word 页眉，不在正文里重复
  assert(/<w:drawing>/.test(flow.headers), "页眉里应有公开号条的图片");
  eq(flow.media, 10, "图片数：1 幅附图 + 3 个字 + 页眉 3 块（WO、2099/000001、PCT 号）+ 第 4、5、6 页各一张整页图");
  const notes = flow.report.notes.join("\n");
  assert(/第 1–3 页的文字画在整页扫描图上面/.test(notes) && /可能有识别错误/.test(notes), `应说明正文来自 PDF 已有的识别文字：${notes}`);
  assert(/写进了 Word 页眉页脚/.test(notes), "应说明页眉块写进了页眉");
  assert(/第 6 页是带表格线的表单：格线保留为垫底的图片，文字按原位置放在文本框里/.test(notes), "应说明表单页的做法");
  assert(/第 4 页没有文字层/.test(notes), "应说明哪些页是整页附图");
  assert(/第 5 页是没有做过文字识别的扫描页/.test(notes), "应说明哪些页的文字没有转成正文");
  // 逐行一致模式不重排文字，整页图照旧垫在文字下面；没有文字层的页同样要归类、标出
  const exact = await read("exact");
  eq(exact.report.textCoverage, 1, "exact：文字完整性");
  eq(JSON.stringify(exact.report.scanPages), JSON.stringify({ text: [], picture: [4], unread: [5] }), "exact：各页的归类");
  eq((exact.xml.match(/behindDoc="1"/g) ?? []).length, 4, "exact：有文字的 4 页仍以整页图衬底");
  eq((exact.xml.match(/name="pdf2word-page 5"/g) ?? []).length, 1, "exact：没读出来的页也标出");
  // 没有画布时结果相同（表单页不再靠渲染）
  const prev = process.env.PDF_RASTER;
  process.env.PDF_RASTER = "0";
  try {
    const plain = await read("flow");
    eq(plain.report.textCoverage, 1, "无画布：文字完整性");
    eq(JSON.stringify(plain.report.scanPages), JSON.stringify({ text: [1, 2, 3, 6], picture: [4], unread: [5] }), "无画布：各页的归类");
    assert(plain.paras.some((t) => /according to claim\s*\[IMG\]\s*, wherein the frames/.test(t)), "无画布：句中的字仍留在原行");
  } finally {
    if (prev === undefined) delete process.env.PDF_RASTER; else process.env.PDF_RASTER = prev;
  }
});

await test("扫描页印成图的页码：按位数变化推出编号，写成页码域；两位数的数字粘成一块时按宽度算位数；推不出来时不写", async () => {
  const { buildDocModel } = await import("../src/services/pdfConvert/layout.js");
  const { encodePng } = await import("../src/services/pdfConvert/extract.js");
  // 只构造版面需要的字段：每页一行正文，页顶中间一块"页码"（宽度按位数，第 13 页起两位数粘成一块）
  const mk = (index: number, printed: number | null) => {
    const png = encodePng(4, 4, new Uint8Array(16), 1);
    const digits = printed === null ? 0 : String(printed).length;
    const w = 6 * digits, merged = printed !== null && printed >= 13 && printed <= 15;
    const pieces = printed === null ? [] : [{ x0: 300, y0: 40, x1: 300 + w, y1: 49, kind: "glyphs", glyphs: merged ? 1 : digits, png: Buffer.from(png.toString("base64") + index, "utf8"), pxW: 4, pxH: 4 }];
    const scanImage = { x0: 0, y0: 0, x1: 595, y1: 842, png, pxW: 4, pxH: 4 };
    const span = (text: string, y: number) => ({ text, x0: 72, x1: 500, top: y - 9, bottom: y + 2.5, baseline: y, size: 11, font: "Times New Roman", bold: false, italic: false, color: "000000", natural: 428, wordW: 428, spaceW: 3, wordSpaceW: 3 });
    return {
      index, width: 595, height: 842, spans: [span(`Paragraph text on page ${index + 1} that runs across the line.`, 120), span("Another line of the same paragraph continues here.", 136)],
      images: [scanImage], segs: [], fills: [], paths: [], vectorShapes: 0, rotatedText: 0,
      scan: { pieces, rules: [], inkUnderText: 0, textLines: 0, ink: 0.01, glyphCount: 1, graphicCount: 0, graphicArea: 0, image: scanImage },
    };
  };
  // 原文第 3 页起印 1（前两页不印）：第 3–20 页印 1–18
  const pages = [mk(0, null), mk(1, null), ...Array.from({ length: 18 }, (_, k) => mk(k + 2, k + 1))];
  const model = buildDocModel(pages as any, { mode: "flow" });
  eq(model.stats.scan?.pageNumbers, 18, "18 页的页码换成文字");
  const head = (model.headers?.default ?? []).flatMap((p: any) => p.runs).map((r: any) => (r.field ? `{${r.field}}` : r.text)).join("");
  assert(head.includes("{PAGE}"), `页眉应有页码域：${head}`);
  // 全是一位数（看不出起点）：不写页码
  const short = buildDocModel([mk(0, 4), mk(1, 5), mk(2, 6)] as any, { mode: "flow" });
  eq(short.stats.scan?.pageNumbers ?? 0, 0, "推不出起点时不写页码");
});

await test("扫描图拆分（scanLayer）：墨迹太多、碎片太多或和文字大面积重叠一半的图不拆；整页都是成行的字时看得出是文字", async () => {
  const { splitScan, looksLikeText } = await import("../src/services/pdfConvert/scanLayer.js");
  const W = 600, H = 800, box = { x0: 0, y0: 0, x1: 300, y1: 400 };
  const blank = () => new Uint8Array(W * H).fill(255);
  const rect = (d: Uint8Array, x0: number, y0: number, x1: number, y1: number) => { for (let y = y0; y < y1; y++) d.fill(0, y * W + x0, y * W + x1); };
  // 一张照片（大片深色）不是文字替换式的扫描页
  const dark = blank();
  rect(dark, 0, 0, W, H / 2);
  eq(splitScan({ data: dark, w: W, h: H }, box, []), null, "墨迹超过三成的图不拆");
  // 空白页：没有东西可拆，也不像文字
  const empty = splitScan({ data: blank(), w: W, h: H }, box, []);
  assert(empty && empty.pieces.length === 0 && !looksLikeText(empty), "空白页没有碎片");
  // 30 行、每行 24 个小方块（字）
  const page = blank();
  for (let r = 0; r < 30; r++) for (let k = 0; k < 24; k++) rect(page, 40 + k * 18, 60 + r * 22, 40 + k * 18 + 14, 60 + r * 22 + 12);
  const text = splitScan({ data: page, w: W, h: H }, box, []);
  assert(text && looksLikeText(text), `成行的字应看得出是文字：${text && JSON.stringify({ lines: text.textLines, glyphs: text.glyphCount })}`);
  const drawing = blank();
  rect(drawing, 100, 100, 500, 104); rect(drawing, 100, 500, 500, 504); rect(drawing, 100, 100, 104, 504); rect(drawing, 496, 100, 500, 504);
  for (let k = 0; k < 6; k++) rect(drawing, 150 + k * 20, 300, 162 + k * 20, 314);
  const fig = splitScan({ data: drawing, w: W, h: H }, box, []);
  assert(fig && !looksLikeText(fig) && fig.graphicCount >= 1, `一幅带几个标注的附图不是文字页：${fig && JSON.stringify({ lines: fig.textLines, glyphs: fig.glyphCount, graphics: fig.graphicCount })}`);
});

// ---------------------------------------------------------------------------

await fs.rm(tmp, { recursive: true, force: true });
const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} 通过`);
if (failed.length) {
  for (const f of failed) console.log(`\n\x1b[31m${f.name}\x1b[0m\n${f.err}`);
  process.exit(1);
}
