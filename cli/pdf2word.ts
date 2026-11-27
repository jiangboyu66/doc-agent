/**
 * pdf2word：把 doc-agent 的内置 PDF → Word 引擎包装成独立的命令行工具，供外部系统（AltaLex）在沙箱里调用。
 * 打包：npm run build:pdf2word（scripts/build-pdf2word.mjs），产物是单文件加上它旁边的数据：pdf.js 的 CMap / 标准字体，
 * 以及把矢量图、轮廓文字绘制的表格、行间公式渲染成图片的 @napi-rs/canvas（只带 linux-x64 glibc 的原生模块）。
 *
 *   node pdf2word.mjs --version
 *   node pdf2word.mjs convert <in.pdf> <out.docx> [--mode flow|exact] [--pages 1-20,25] [--max-pages N] [--scan-blocks]
 *
 * 契约 v1（有任何变化，次版本号加一，调用方的适配器测试同步修改）：
 *   - stdout 恰好一行 JSON（UTF-8，以 \n 结尾），除此之外没有任何输出；告警、诊断一律写 stderr。
 *       成功：{"ok":true,"tool":"pdf2word","version","commit","mode","bytes","report":ConversionReport}
 *       失败：{"ok":false,"tool":"pdf2word","version","commit","code","message"}
 *     code：usage | input_unreadable | not_pdf | encrypted | conversion_failed | output_unwritable | engine_incomplete
 *   - --version：stdout 一行 "pdf2word <version> (doc-agent <commit>)"，退出码 0。
 *   - 退出码：0 成功；2 失败；64 参数错误。
 *   - 只写 <out.docx> 这一个文件；只用内置引擎；不访问网络，不需要环境变量，不调用 Python / LibreOffice。
 *   - 发布包带了本平台的渲染组件时它必须能加载，否则 engine_incomplete（不静默丢掉矢量图和表格）；
 *     没带的平台不渲染，矢量内容按文字重建，未能保留的在 report.notes 里写明。
 *   - report.scanPages（1.1.0 起；没有这类页时不出现）：{"text":[…],"picture":[…],"unread":[…]}，页码从 1 起、升序。
 *       text     文字画在整页黑白扫描图上面的页（含表单页）：扫描图已拆开，正文是 PDF 里已有的识别文字
 *       picture  没有文字层的附图页：整页保留成一张图片，内容没有丢
 *       unread   没有文字层、图里看得出是成行文字（或无法判断）的页：按整页图片保留，文字没有转成正文
 *       unknown  （1.2.0 起，没有时省略）unread 里无法判断的页（拆不开的彩色、灰阶照片式、拼条扫描图）：也可能是附图
 *       blocks   （1.2.0 起，--scan-blocks 且有这类页时）unread 里按原位置分了块的页（见下面的图片名）
 *     调用方据此判断没有文字层的页是否都已如实成图（unread 非空说明还有文字没读出来）。
 *   - --scan-blocks（1.2.0 起，只用于 flow）：unread 里能拆开的页不再整页保留成图——整页图换成抹掉了字的底图，
 *     每块文字（按栏、按段）一个按原位置定位的文本框，框里放这块的原图。调用方识别后换成文字，图留在原处、文字留在原位。
 *   - 转了 90° 的倍数放置的图片（1.2.0 起）按页面方向转正写进 Word；页上只有一张扫描图、图框与页面框横竖互换时，
 *     这一页按扫描图的大小和方向转换（PDF 阅读器里这样的页显示不全），report.notes 里写明。
 *   - Word 里给调用方的图片名（1.1.0 起，wp:docPr 的 name）：
 *       pdf2word-glyph <id>  句子里只以图片存在的字（行内图片，在原来的位置）
 *       pdf2word-page <n>    第 n 页（unread）的整页图
 *       pdf2word-form <n>    第 n 页表单垫在原位文本框下面的格线图
 *       pdf2word-scan <n>    （1.2.0）第 n 页（blocks）抹掉了字的底图：附图、条码、表格线、附图框里的标注留着
 *       pdf2word-text <n>.<k>（1.2.0）第 n 页第 k 块文字的原图，在名为 pdf2word-block <n>.<k> 的原位文本框里
 *     调用方可以把 glyph、page、text 识别成文字后替换掉；其余图片照旧叫"图片 <id>"。
 *   - pdf2word-block <n>.<k> 的 wp:docPr descr（1.2.0）是一行 JSON：{"col","band","furniture"?,"lines","size","pitch","align"}。
 *       k 是阅读顺序（页眉一行，然后每段版面里先左栏从上到下、再右栏；同一行的格从左到右）。
 *       col      "L" 左栏 / "R" 右栏 / "X" 通栏（不分栏的页、跨栏的块、栏间距里的页码）
 *       band     第几段版面（横贯全页的大段空白、通栏的块把页面分成上下几段，从 0 起）
 *       furniture 页眉页脚（页顶、页底只有单行块的那一段：公开号、日期、页码），不属于正文的段落流；不是时省略
 *       lines、size、pitch、align  行数、估计的字号与行距（pt）、"both" 两端对齐 / "left"
 *     调用方据此把一栏末尾没说完的段落接到下一栏、下一页的开头，并按原来的字号、行距回填译文。
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
// pdf.js 的 worker 必须打进同一个文件：Node 下 pdf.js 用"假 worker"，先找 globalThis.pdfjsWorker，
// 找不到才按 workerSrc 动态加载 pdf.worker.mjs——单文件发布包里没有这个文件
// @ts-ignore pdf.js 没有为 worker 提供类型声明
import * as pdfjsWorker from "pdfjs-dist/legacy/build/pdf.worker.mjs";

/** 打包时由 esbuild 的 define 写入；直接用 tsx 运行源码时没有 */
declare const __PDF2WORD_VERSION__: string | undefined;
declare const __PDF2WORD_COMMIT__: string | undefined;
const BUNDLED = typeof __PDF2WORD_VERSION__ === "string";
const VERSION = typeof __PDF2WORD_VERSION__ === "string" ? __PDF2WORD_VERSION__ : "0.0.0-dev";
const COMMIT = typeof __PDF2WORD_COMMIT__ === "string" ? __PDF2WORD_COMMIT__ : "unknown";

// ---------------------------------------------------------------------------
// 以下三件事必须在引擎模块加载之前完成（引擎在 convert 里动态导入）
// ---------------------------------------------------------------------------

// 1. stdout 只留给最后那一行 JSON。缺少 @napi-rs/canvas 时 pdf.js 用 console.log 打印告警：
//    console.log 改写到 stderr；事先绑定了 console.log 的代码最终也经过 stdout.write，一并改写
const stdoutFd = 1;
console.log = console.error;
console.info = console.error;
console.debug = console.error;
process.stdout.write = process.stderr.write.bind(process.stderr) as typeof process.stdout.write;
// 2. worker 与 CLI 在同一个文件里
(globalThis as any).pdfjsWorker = pdfjsWorker;
// 3. 矢量图 / 公式区域：能加载 @napi-rs/canvas 就渲染成图片（发布包里它与单文件并列）。PDF_RASTER=0 只供调试时关闭

type Code = "usage" | "input_unreadable" | "not_pdf" | "encrypted" | "conversion_failed" | "output_unwritable" | "engine_incomplete";

class CliError extends Error {
  constructor(readonly code: Code, message: string) {
    super(message);
  }
}

const USAGE = [
  "用法：",
  "  pdf2word --version",
  "  pdf2word convert <in.pdf> <out.docx> [--mode flow|exact] [--pages 1-20,25] [--max-pages N] [--scan-blocks]",
].join("\n");

function writeStdout(text: string) {
  const b = Buffer.from(text, "utf8");
  for (let off = 0; off < b.length; ) {
    try {
      off += fs.writeSync(stdoutFd, b, off);
    } catch (e: any) {
      if (e?.code !== "EAGAIN") throw e;
    }
  }
}

let finished = false;
function finish({ ok, ...rest }: { ok: boolean } & Record<string, unknown>, exitCode: number) {
  if (finished) return;
  finished = true;
  writeStdout(JSON.stringify({ ok, tool: "pdf2word", version: VERSION, commit: COMMIT, ...rest }) + "\n");
  process.exit(exitCode);
}

function fail(code: Code, message: string) {
  finish({ ok: false, code, message }, code === "usage" ? 64 : 2);
}

// ---------------------------------------------------------------------------

interface Args {
  input: string;
  output: string;
  mode: "flow" | "exact";
  pages?: string;
  maxPages?: number;
  scanBlocks?: boolean;
}

function parseArgs(argv: string[]): Args {
  const [cmd, ...rest] = argv;
  if (cmd !== "convert") throw new CliError("usage", cmd ? `未知命令：${cmd}` : "缺少命令");
  const positional: string[] = [];
  const args: Partial<Args> = { mode: "flow" };
  for (let i = 0; i < rest.length; i++) {
    const a = rest[i];
    if (a === "--mode" || a === "--pages" || a === "--max-pages") {
      const v = rest[++i];
      if (v === undefined) throw new CliError("usage", `${a} 缺少取值`);
      if (a === "--mode") {
        if (v !== "flow" && v !== "exact") throw new CliError("usage", `--mode 只能是 flow 或 exact：${v}`);
        args.mode = v;
      } else if (a === "--pages") {
        if (!v.trim()) throw new CliError("usage", "--pages 不能为空");
        args.pages = v;
      } else {
        if (!/^[1-9]\d{0,5}$/.test(v)) throw new CliError("usage", `--max-pages 必须是正整数：${v}`);
        args.maxPages = Number(v);
      }
    } else if (a === "--scan-blocks") {
      args.scanBlocks = true;
    } else if (a.startsWith("-")) {
      throw new CliError("usage", `未知选项：${a}`);
    } else {
      positional.push(a);
    }
  }
  if (positional.length !== 2) throw new CliError("usage", "需要两个参数：<in.pdf> <out.docx>");
  [args.input, args.output] = positional;
  return args as Args;
}

const hasFile = (dir: string, ext: RegExp) => {
  try {
    return fs.readdirSync(dir).some((n) => ext.test(n));
  } catch {
    return false;
  }
};

/**
 * 读入 PDF 之前确认引擎完整。pdf.js 找不到 CMap / 标准字体时不报错：未嵌入字体的中文 PDF 转出来是 0 个字，
 * 完整性校验却仍是 100%——这里改为明确失败，不静默降级
 */
async function checkEngine() {
  if (typeof (process as any).getBuiltinModule !== "function") {
    // pdf.js 用 process.getBuiltinModule("fs") 读取 CMap 与标准字体（Node 20.16 起才有）
    throw new CliError("engine_incomplete", `需要 Node 20.16 或更高版本（当前 ${process.version}）：pdf.js 无法读取 CMap 与标准字体`);
  }
  const { pdfjsDataOptions } = await import("../src/services/pdfConvert/render.js");
  const opts = pdfjsDataOptions();
  const need: Array<[string, unknown, RegExp]> = [["cmaps", opts.cMapUrl, /\.bcmap$/], ["standard_fonts", opts.standardFontDataUrl, /\.(pfb|ttf|otf)$/i]];
  for (const [name, dir, ext] of need) {
    if (typeof dir !== "string") throw new CliError("engine_incomplete", `找不到 pdfjs-dist（${name}/ 无从定位）`);
    if (!hasFile(dir, ext)) throw new CliError("engine_incomplete", `pdf.js 数据目录 ${name}/ 缺失或为空：${dir}`);
  }
  if (BUNDLED) {
    // 发布包里 fontMetrics.json 与 pdf2word.mjs 并列（引擎按 import.meta.url 找它；找不到时静默改用估算的字宽）
    const metrics = fileURLToPath(new URL("./fontMetrics.json", import.meta.url));
    if (!fs.existsSync(metrics)) throw new CliError("engine_incomplete", `缺少 ${path.basename(metrics)}`);
    // 发布包带了本平台的渲染组件，它就必须能加载：加载不了时矢量图、轮廓文字的表格会整块丢失，而文字完整性仍是 100%
    const platform = canvasPlatform();
    const shipped = fileURLToPath(new URL(`./node_modules/@napi-rs/canvas-${platform}/package.json`, import.meta.url));
    if (process.env.PDF_RASTER !== "0" && fs.existsSync(shipped)) {
      const { rasterAvailable } = await import("../src/services/pdfConvert/render.js");
      if (!(await rasterAvailable())) throw new CliError("engine_incomplete", `渲染组件 @napi-rs/canvas-${platform} 无法加载`);
    }
  }
}

/** @napi-rs/canvas 原生模块的平台名（linux 区分 glibc / musl，判据与 @napi-rs/canvas 自己的一致） */
function canvasPlatform(): string {
  if (process.platform !== "linux") return `${process.platform}-${process.arch}`;
  let musl = false;
  try { musl = fs.readFileSync("/usr/bin/ldd", "utf8").includes("musl"); } catch { /* 没有 ldd：按 glibc */ }
  return `linux-${process.arch}-${musl ? "musl" : "gnu"}`;
}

function isPasswordError(e: any): boolean {
  return e?.name === "PasswordException" || /password/i.test(String(e?.message ?? ""));
}

/** 只写 out 这一个文件；打不开时不碰已有文件，写到一半失败时删掉写了一半的文件 */
function writeOutput(out: string, data: Buffer) {
  let fd: number;
  try {
    fd = fs.openSync(out, "w");
  } catch (e: any) {
    throw new CliError("output_unwritable", `无法写入输出文件（${e?.code ?? e?.message}）`);
  }
  try {
    for (let off = 0; off < data.length; ) off += fs.writeSync(fd, data, off);
    fs.closeSync(fd);
  } catch (e: any) {
    try { fs.closeSync(fd); } catch { /* 已关闭 */ }
    try { fs.unlinkSync(out); } catch { /* 删不掉也只能如此 */ }
    throw new CliError("output_unwritable", `写入输出文件失败（${e?.code ?? e?.message}）`);
  }
}

async function convert(args: Args) {
  await checkEngine();

  let buf: Buffer;
  try {
    buf = fs.readFileSync(args.input);
  } catch (e: any) {
    throw new CliError("input_unreadable", `无法读取输入文件（${e?.code ?? e?.message}）`);
  }
  if (!buf.subarray(0, 1024).includes("%PDF-")) throw new CliError("not_pdf", "输入文件不是 PDF（前 1024 字节里没有 %PDF- 文件头）");

  const { pdfToDocx } = await import("../src/services/pdfConvert/index.js");
  let result: Awaited<ReturnType<typeof pdfToDocx>>;
  try {
    result = await pdfToDocx(buf, { engine: "builtin", mode: args.mode, pages: args.pages, maxPages: args.maxPages, scanBlocks: args.scanBlocks });
  } catch (e: any) {
    if (isPasswordError(e)) throw new CliError("encrypted", "PDF 有密码保护");
    // 页码范围写法无法识别，或不在这份 PDF 的页码内：调用方的参数问题，引擎的原话说得清楚
    if (args.pages && /^页码范围/.test(String(e?.message ?? ""))) throw new CliError("usage", String(e.message));
    console.error(e?.stack ?? e);
    throw new CliError("conversion_failed", String(e?.message ?? e).slice(0, 500));
  }
  writeOutput(args.output, result.data);
  finish({ ok: true, mode: args.mode, bytes: result.data.length, report: result.report }, 0);
}

async function main(argv: string[]) {
  if (argv[0] === "--version") {
    if (argv.length > 1) throw new CliError("usage", "--version 不接受其他参数");
    writeStdout(`pdf2word ${VERSION} (doc-agent ${COMMIT})\n`);
    process.exit(0);
  }
  if (argv[0] === "--help" || argv[0] === "-h") throw new CliError("usage", "显示用法");
  await convert(parseArgs(argv));
}

process.on("unhandledRejection", (e: any) => console.error("pdf2word: unhandled rejection:", e?.stack ?? e));
process.on("uncaughtException", (e: any) => {
  console.error(e?.stack ?? e);
  fail("conversion_failed", String(e?.message ?? e).slice(0, 500));
});

main(process.argv.slice(2)).catch((e: any) => {
  if (e instanceof CliError) {
    if (e.code === "usage") console.error(`${e.message}\n${USAGE}`);
    fail(e.code, e.message);
  } else {
    console.error(e?.stack ?? e);
    fail("conversion_failed", String(e?.message ?? e).slice(0, 500));
  }
});
