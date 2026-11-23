/**
 * pdf2word：把 doc-agent 的内置 PDF → Word 引擎包装成独立的命令行工具，供外部系统（AltaLex）在沙箱里调用。
 * 打包：npm run build:pdf2word（scripts/build-pdf2word.mjs），产物是不依赖 node_modules 的单文件。
 *
 *   node pdf2word.mjs --version
 *   node pdf2word.mjs convert <in.pdf> <out.docx> [--mode flow|exact] [--pages 1-20,25] [--max-pages N]
 *
 * 契约 v1（有任何变化，次版本号加一，调用方的适配器测试同步修改）：
 *   - stdout 恰好一行 JSON（UTF-8，以 \n 结尾），除此之外没有任何输出；告警、诊断一律写 stderr。
 *       成功：{"ok":true,"tool":"pdf2word","version","commit","mode","bytes","report":ConversionReport}
 *       失败：{"ok":false,"tool":"pdf2word","version","commit","code","message"}
 *     code：usage | input_unreadable | not_pdf | encrypted | conversion_failed | output_unwritable | engine_incomplete
 *   - --version：stdout 一行 "pdf2word <version> (doc-agent <commit>)"，退出码 0。
 *   - 退出码：0 成功；2 失败；64 参数错误。
 *   - 只写 <out.docx> 这一个文件；只用内置引擎；不访问网络，不需要环境变量，不调用 Python / LibreOffice。
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
// 3. 不渲染矢量图 / 公式区域（发布包不带 @napi-rs/canvas）；这些区域按文字处理或不输出
process.env.PDF_RASTER ??= "0";

type Code = "usage" | "input_unreadable" | "not_pdf" | "encrypted" | "conversion_failed" | "output_unwritable" | "engine_incomplete";

class CliError extends Error {
  constructor(readonly code: Code, message: string) {
    super(message);
  }
}

const USAGE = [
  "用法：",
  "  pdf2word --version",
  "  pdf2word convert <in.pdf> <out.docx> [--mode flow|exact] [--pages 1-20,25] [--max-pages N]",
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
  }
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
    result = await pdfToDocx(buf, { engine: "builtin", mode: args.mode, pages: args.pages, maxPages: args.maxPages });
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
