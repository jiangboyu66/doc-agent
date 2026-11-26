/**
 * 打包 pdf2word 发布包（命令行契约见 cli/pdf2word.ts）
 *
 *   npm run build:pdf2word                 构建 dist/pdf2word-<版本>.tgz 与 .sha256，再在只含解压结果的空目录里跑冒烟测试
 *   npm run build:pdf2word -- --no-smoke   只构建
 *
 * 发布包（运行只需要 Node ≥ 20.16；渲染组件只带 linux-x64 glibc 的原生模块，其他平台照常转换，只是不渲染）：
 *   pdf2word/
 *     pdf2word.mjs               esbuild 单文件：CLI + 引擎 + pdf.js + worker（@napi-rs/canvas 不打进单文件，在下面的 node_modules 里）
 *     pdf2word.mjs.LEGAL.txt     打进单文件的第三方代码的许可证注释
 *     fontMetrics.json           Word 字体的字宽表（引擎按 import.meta.url 在单文件旁边找它）
 *     node_modules/pdfjs-dist/{package.json, LICENSE, cmaps/, standard_fonts/}
 *                                pdf.js 的 CMap 与标准字体（引擎用 createRequire 从单文件所在目录定位）
 *     node_modules/@napi-rs/canvas/{package.json, LICENSE, *.js}
 *     node_modules/@napi-rs/canvas-linux-x64-gnu/{package.json, skia.linux-x64-gnu.node}
 *                                把矢量图、文字被转成轮廓的表格、行间公式渲染成图片（没有它时这些内容整块丢失，
 *                                文字完整性却仍是 100%）。CLI 发现带了本平台的原生模块却加载不了时报 engine_incomplete
 * 可复现：文件 mtime 固定为 HEAD 的提交时间，属主 0:0，按名字排序，gzip -n；同一提交连续构建两次，摘要相同。
 * 打包需要 GNU tar 与 gzip（Linux、WSL 或 CI）。
 */

import { build } from "esbuild";
import { execFileSync, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { createRequire } from "node:module";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";

/**
 * pdf2word 自己的版本号（与 doc-agent 的版本号无关）：主版本对应契约 v1；契约有任何变化，次版本号加一；
 * 只有引擎更新时修订号加一。发版标签为 pdf2word-v<版本>，CI 会核对标签与这里一致。
 */
const VERSION = "1.1.1";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(root, "dist");
const stage = path.join(dist, "pdf2word-stage");
const pkg = path.join(stage, "pdf2word");
const name = `pdf2word-${VERSION}.tgz`;
const tgz = path.join(dist, name);
const smoke = !process.argv.includes("--no-smoke");

const git = (...args) => {
  try {
    return execFileSync("git", args, { cwd: root, encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim();
  } catch {
    return "";
  }
};
const head = git("rev-parse", "--short=7", "HEAD");
const COMMIT = head ? head + (git("status", "--porcelain", "--untracked-files=no") ? "-dirty" : "") : "unknown";
const MTIME = Number(git("log", "-1", "--format=%ct")) || 0;

function fail(msg) {
  console.error(`\x1b[31m${msg}\x1b[0m`);
  process.exit(1);
}

/** 复制目录：按名字排序，拒绝符号链接 */
function copyTree(from, to) {
  fs.mkdirSync(to, { recursive: true });
  for (const n of fs.readdirSync(from).sort()) {
    const src = path.join(from, n), dst = path.join(to, n), st = fs.lstatSync(src);
    if (st.isSymbolicLink()) fail(`发布包里不允许符号链接：${src}`);
    if (st.isDirectory()) copyTree(src, dst);
    else fs.copyFileSync(src, dst);
  }
}

function walk(dir, out = []) {
  for (const n of fs.readdirSync(dir).sort()) {
    const p = path.join(dir, n), st = fs.lstatSync(p);
    if (st.isSymbolicLink()) fail(`发布包里不允许符号链接：${p}`);
    if (st.isDirectory()) walk(p, out);
    else out.push({ path: p, size: st.size });
  }
  return out;
}

// ---------------------------------------------------------------------------
// 1. 单文件
// ---------------------------------------------------------------------------

if (!/GNU tar/.test(spawnSync("tar", ["--version"], { encoding: "utf8" }).stdout ?? "")) fail("打包需要 GNU tar（Linux、WSL 或 CI）。");
fs.rmSync(stage, { recursive: true, force: true });
fs.mkdirSync(pkg, { recursive: true });

const { metafile } = await build({
  entryPoints: [path.join(root, "cli/pdf2word.ts")],
  outfile: path.join(pkg, "pdf2word.mjs"),
  bundle: true,
  platform: "node",
  format: "esm",
  target: "node20",
  minify: true,
  charset: "utf8",
  legalComments: "external",
  external: ["@napi-rs/canvas"],
  // 打进 ESM 单文件的 CommonJS 代码要用 require；别名避免与源码里的 createRequire 重名
  banner: { js: 'import { createRequire as __pdf2wordCreateRequire } from "node:module"; const require = __pdf2wordCreateRequire(import.meta.url);' },
  define: { __PDF2WORD_VERSION__: JSON.stringify(VERSION), __PDF2WORD_COMMIT__: JSON.stringify(COMMIT) },
  metafile: true,
  logLevel: "warning",
});
// esbuild 只收集 @license / @preserve / /*! 注释：pdf.js 的版权声明用的是 @licstart，jsdiff 没有版权注释，
// 收集结果是空的。按 metafile 把打进单文件的每个第三方包的 LICENSE 原文附在后面；找不到 LICENSE 时构建失败
const legal = path.join(pkg, "pdf2word.mjs.LEGAL.txt");
const bundled = [...new Set(Object.keys(metafile.inputs).map((f) => /node_modules\/((?:@[^/]+\/)?[^/]+)\//.exec(f)?.[1]).filter(Boolean))].sort();
const notices = bundled.map((p) => {
  const dir = path.join(root, "node_modules", p);
  const file = fs.readdirSync(dir).sort().find((n) => /^licen[cs]e/i.test(n));
  if (!file) fail(`打进单文件的 ${p} 没有 LICENSE 文件`);
  const { version, license } = JSON.parse(fs.readFileSync(path.join(dir, "package.json"), "utf8"));
  return `${"=".repeat(78)}\n${p}@${version}（${license}）\n${"=".repeat(78)}\n\n${fs.readFileSync(path.join(dir, file), "utf8").trim()}\n`;
});
const collected = fs.existsSync(legal) ? fs.readFileSync(legal, "utf8").trim() : "";
fs.writeFileSync(legal, [collected && `${collected}\n`, ...notices].filter(Boolean).join("\n"));

// ---------------------------------------------------------------------------
// 2. 数据文件
// ---------------------------------------------------------------------------

fs.copyFileSync(path.join(root, "src/services/pdfConvert/fontMetrics.json"), path.join(pkg, "fontMetrics.json"));
const pdfjsDir = path.dirname(createRequire(import.meta.url).resolve("pdfjs-dist/package.json"));
const pdfjsOut = path.join(pkg, "node_modules/pdfjs-dist");
fs.mkdirSync(pdfjsOut, { recursive: true });
for (const f of ["package.json", "LICENSE"]) fs.copyFileSync(path.join(pdfjsDir, f), path.join(pdfjsOut, f));
for (const d of ["cmaps", "standard_fonts"]) copyTree(path.join(pdfjsDir, d), path.join(pdfjsOut, d));

// 渲染组件：只带服务器沙箱的平台（linux-x64 glibc）。npm 只装构建机平台的原生模块，所以这里固定平台、缺了就失败，
// 不按构建机上碰巧装了哪些来带：同一提交在哪里构建，发布包的内容都一样
const CANVAS_PLATFORMS = ["linux-x64-gnu"];
const canvasDir = path.dirname(createRequire(import.meta.url).resolve("@napi-rs/canvas/package.json"));
const canvasVersion = JSON.parse(fs.readFileSync(path.join(canvasDir, "package.json"), "utf8")).version;
const canvasOut = path.join(pkg, "node_modules/@napi-rs/canvas");
fs.mkdirSync(canvasOut, { recursive: true });
for (const f of fs.readdirSync(canvasDir).sort()) if (/\.(js|json)$|^LICENSE$/.test(f)) fs.copyFileSync(path.join(canvasDir, f), path.join(canvasOut, f));
for (const platform of CANVAS_PLATFORMS) {
  const native = `@napi-rs/canvas-${platform}`;
  let dir;
  try {
    dir = path.dirname(createRequire(path.join(canvasDir, "package.json")).resolve(`${native}/package.json`));
  } catch {
    fail(`缺少 ${native}@${canvasVersion}：在 Linux x64 上 npm ci，或 npm i --no-save --os=linux --cpu=x64 --libc=glibc ${native}@${canvasVersion}`);
  }
  const { version } = JSON.parse(fs.readFileSync(path.join(dir, "package.json"), "utf8"));
  if (version !== canvasVersion) fail(`${native} ${version} 与 @napi-rs/canvas ${canvasVersion} 版本不一致`);
  const out = path.join(pkg, "node_modules", native);
  fs.mkdirSync(out, { recursive: true });
  for (const f of fs.readdirSync(dir).sort()) if (/\.(node|json)$/.test(f)) fs.copyFileSync(path.join(dir, f), path.join(out, f));
}

// ---------------------------------------------------------------------------
// 3. 可复现的压缩包
// ---------------------------------------------------------------------------

const files = walk(pkg);
const unpacked = files.reduce((a, f) => a + f.size, 0);
fs.rmSync(tgz, { force: true });
const tar = path.join(dist, `pdf2word-${VERSION}.tar`);
execFileSync("tar", ["--sort=name", `--mtime=@${MTIME}`, "--owner=0", "--group=0", "--numeric-owner", "--mode=a+rX,u+w,go-w", "--format=gnu", "-cf", tar, "-C", stage, "pdf2word"]);
execFileSync("gzip", ["-n", "-9", "-f", tar]);
fs.renameSync(`${tar}.gz`, tgz);
const sha = createHash("sha256").update(fs.readFileSync(tgz)).digest("hex");
fs.writeFileSync(`${tgz}.sha256`, `${sha}  ${name}\n`);

const mb = (n) => `${(n / 1024 / 1024).toFixed(2)} MB`;
console.log(`pdf2word ${VERSION} (doc-agent ${COMMIT})`);
console.log(`  ${path.relative(root, tgz)}：${mb(fs.statSync(tgz).size)}；解压后 ${mb(unpacked)}，${files.length} 个文件；单文件 ${mb(fs.statSync(path.join(pkg, "pdf2word.mjs")).size)}`);
console.log(`  sha256 ${sha}`);

if (!smoke) process.exit(0);

// ---------------------------------------------------------------------------
// 4. 冒烟测试：在只含解压结果的空目录里运行（没有 node_modules 可借，不传任何环境变量）
// ---------------------------------------------------------------------------

const fixtures = path.join(root, "test/fixtures");
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "pdf2word-smoke-"));
execFileSync("tar", ["-xzf", tgz, "-C", tmp, "--no-same-owner"]);
const entry = path.join(tmp, "pdf2word/pdf2word.mjs");
const work = path.join(tmp, "work");
fs.mkdirSync(work);

const results = [];
function check(label, fn) {
  try {
    const note = fn();
    results.push(true);
    console.log(`  \x1b[32m✓\x1b[0m ${label}${note ? `：${note}` : ""}`);
  } catch (e) {
    results.push(false);
    console.log(`  \x1b[31m✗ ${label}\x1b[0m\n    ${String(e?.message ?? e).split("\n").join("\n    ")}`);
  }
}
function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

function run(args, bin = entry) {
  const r = spawnSync(process.execPath, [bin, ...args], { cwd: work, env: {}, encoding: "utf8", timeout: 180_000, maxBuffer: 64 << 20 });
  return { status: r.status, stdout: r.stdout ?? "", stderr: r.stderr ?? "" };
}

/** stdout 恰好一行 JSON，以 \n 结尾；带着 tool / version / commit */
function parse(r) {
  const lines = r.stdout.split("\n");
  assert(r.stdout.endsWith("\n") && lines.length === 2, `stdout 应恰好一行（实际 ${lines.length - 1} 行）：${r.stdout.slice(0, 300)}\nstderr：${r.stderr.slice(-800)}`);
  const j = JSON.parse(lines[0]);
  assert(j.tool === "pdf2word" && j.version === VERSION && j.commit === COMMIT, `tool / version / commit 不对：${lines[0].slice(0, 200)}`);
  return j;
}

function convert(pdf, extra = [], bin = entry) {
  const out = path.join(work, `${path.basename(pdf, ".pdf")}-${results.length}.docx`);
  fs.rmSync(out, { force: true });
  const r = run(["convert", path.join(fixtures, pdf), out, ...extra], bin);
  return { r, j: parse(r), out };
}

/** 最小的 zip 读取：按中央目录取出一个条目 */
function zipEntry(buf, entryName) {
  const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  if (eocd < 0) return null;
  let p = buf.readUInt32LE(eocd + 16);
  for (let i = buf.readUInt16LE(eocd + 10); i > 0; i--) {
    if (buf.readUInt32LE(p) !== 0x02014b50) return null;
    const method = buf.readUInt16LE(p + 10), size = buf.readUInt32LE(p + 20), nlen = buf.readUInt16LE(p + 28);
    const skip = nlen + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32), local = buf.readUInt32LE(p + 42);
    if (buf.toString("utf8", p + 46, p + 46 + nlen) === entryName) {
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
      const data = buf.subarray(start, start + size);
      return method === 0 ? Buffer.from(data) : zlib.inflateRawSync(data);
    }
    p += 46 + skip;
  }
  return null;
}

function documentXml(file) {
  const buf = fs.readFileSync(file);
  assert(buf.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 0x03, 0x04])), "输出不是 zip");
  const xml = zipEntry(buf, "word/document.xml");
  assert(xml, "输出里没有 word/document.xml");
  return xml.toString("utf8");
}

const plain = (xml) => xml.replace(/<\/w:p>/g, "\n").replace(/<[^>]+>/g, "").replace(/\n+/g, "\n").trim();

/** 与完整开发环境（仓库的 node_modules，tsx 直接运行源码）转出的 word/document.xml 逐字节比较 */
function sameAsDev(pdf, xml) {
  const out = path.join(work, `${path.basename(pdf, ".pdf")}-dev.docx`);
  const r = spawnSync(process.execPath, ["--import", "tsx", path.join(root, "cli/pdf2word.ts"), "convert", path.join(fixtures, pdf), out], { cwd: root, encoding: "utf8", timeout: 180_000 });
  assert(r.status === 0, `开发环境转换失败：${r.stdout}${r.stderr.slice(-500)}`);
  const dev = documentXml(out);
  assert(dev === xml, `与开发环境的结果不一致（document.xml ${xml.length} / ${dev.length} 字节）`);
}

console.log(`冒烟测试（Node ${process.version}，${tmp}）`);
const t0 = Date.now();

check("--version：一行版本信息，退出码 0", () => {
  const r = run(["--version"]);
  assert(r.status === 0, `退出码 ${r.status}`);
  assert(r.stdout === `pdf2word ${VERSION} (doc-agent ${COMMIT})\n`, `输出：${JSON.stringify(r.stdout)}`);
});

check("contract-cn.pdf：转换成功，文字完整，输出是含 word/document.xml 的 zip", () => {
  const { r, j, out } = convert("contract-cn.pdf");
  assert(r.status === 0 && j.ok, `失败：${r.stdout}`);
  assert(j.mode === "flow", `默认模式应为 flow：${j.mode}`);
  assert(j.report.textTotal > 0 && j.report.textCoverage >= 0.995, `textTotal=${j.report.textTotal} textCoverage=${j.report.textCoverage}`);
  assert(j.bytes === fs.statSync(out).size, `bytes=${j.bytes} 与文件大小不符`);
  sameAsDev("contract-cn.pdf", documentXml(out));
  return `${j.report.pages} 页，${j.report.paragraphs} 段，${j.report.textTotal} 字，${j.report.ms} ms，与开发环境一致`;
});

check("cn-cidfont.pdf（未嵌入的 STSong-Light）：提取到文字，证明 cmaps 已打进包", () => {
  const { r, j, out } = convert("cn-cidfont.pdf");
  assert(r.status === 0 && j.ok, `失败：${r.stdout}`);
  assert(j.report.textTotal > 0, `textTotal=${j.report.textTotal}`);
  const xml = documentXml(out);
  assert(plain(xml).includes("保密信息"), "Word 里没有中文正文");
  sameAsDev("cn-cidfont.pdf", xml);
  return `${j.report.textTotal} 字，与开发环境一致`;
});

check("signature.pdf（标准字体未嵌入）：与开发环境的转换结果一致，证明 standard_fonts 已打进包", () => {
  const { r, j, out } = convert("signature.pdf");
  assert(r.status === 0 && j.ok, `失败：${r.stdout}`);
  sameAsDev("signature.pdf", documentXml(out));
  return `${j.report.paragraphs} 段，${j.report.textTotal} 字`;
});

check("scan-letter.pdf（扫描件）：转换成功，textTotal < 20 × 页数", () => {
  const { r, j } = convert("scan-letter.pdf");
  assert(r.status === 0 && j.ok, `失败：${r.stdout}`);
  assert(j.report.textTotal < 20 * j.report.pages, `textTotal=${j.report.textTotal} pages=${j.report.pages}`);
  return `textTotal=${j.report.textTotal}，pages=${j.report.pages}，images=${j.report.images}`;
});

check("--pages 1：只转第 1 页，报告写明页码范围与总页数", () => {
  const { r, j } = convert("contract-cn.pdf", ["--pages", "1"]);
  assert(r.status === 0 && j.ok, `失败：${r.stdout}`);
  assert(j.report.pages === 1 && j.report.totalPages === 3 && j.report.pageRange === "1", JSON.stringify(j.report).slice(0, 300));
});

check("--pages 超出页码：usage，退出码 64，不写输出", () => {
  const { r, j, out } = convert("contract-cn.pdf", ["--pages", "9-12"]);
  assert(r.status === 64 && !j.ok && j.code === "usage", r.stdout);
  assert(!fs.existsSync(out), "不应写出输出文件");
});

check("encrypted.pdf：encrypted，退出码 2，不写输出", () => {
  const { r, j, out } = convert("encrypted.pdf");
  assert(r.status === 2 && !j.ok && j.code === "encrypted", r.stdout);
  assert(!fs.existsSync(out), "不应写出输出文件");
});

check("不是 PDF：not_pdf，退出码 2", () => {
  const { r, j } = convert("sample.docx");
  assert(r.status === 2 && !j.ok && j.code === "not_pdf", r.stdout);
});

check("输入文件不存在：input_unreadable，退出码 2", () => {
  const r = run(["convert", path.join(work, "missing.pdf"), path.join(work, "x.docx")]);
  const j = parse(r);
  assert(r.status === 2 && j.code === "input_unreadable", r.stdout);
});

check("输出目录不存在：output_unwritable，退出码 2", () => {
  const r = run(["convert", path.join(fixtures, "cn-cidfont.pdf"), path.join(work, "no-such-dir/x.docx")]);
  const j = parse(r);
  assert(r.status === 2 && j.code === "output_unwritable", r.stdout);
});

check("缺参数 / 未知选项 / 错误的 mode：usage，退出码 64", () => {
  for (const args of [["convert", "a.pdf"], [], ["convert", "a.pdf", "b.docx", "--bogus"], ["convert", "a.pdf", "b.docx", "--mode", "x"], ["convert", "a.pdf", "b.docx", "--max-pages", "0"]]) {
    const r = run(args);
    const j = parse(r);
    assert(r.status === 64 && j.code === "usage", `${args.join(" ")} → ${r.status} ${r.stdout}`);
  }
});

for (const dir of ["cmaps", "standard_fonts"]) {
  check(`副本里删掉 ${dir}/：engine_incomplete，不读输入`, () => {
    const copy = fs.mkdtempSync(path.join(os.tmpdir(), "pdf2word-incomplete-"));
    try {
      copyTree(path.join(tmp, "pdf2word"), path.join(copy, "pdf2word"));
      fs.rmSync(path.join(copy, "pdf2word/node_modules/pdfjs-dist", dir), { recursive: true });
      const r = run(["convert", path.join(work, "missing.pdf"), path.join(work, "x.docx")], path.join(copy, "pdf2word/pdf2word.mjs"));
      const j = parse(r);
      assert(r.status === 2 && j.code === "engine_incomplete", r.stdout);
    } finally {
      fs.rmSync(copy, { recursive: true, force: true });
    }
  });
}

const canvasHere = process.platform === "linux" && process.arch === "x64" && !fs.readFileSync("/usr/bin/ldd", "utf8").includes("musl");
if (canvasHere) {
  check("outlined-table.pdf（文字被转成轮廓的表格）：渲染组件可用，表格整体渲染成图片贴回原位，与开发环境一致", () => {
    const { r, j, out } = convert("outlined-table.pdf");
    assert(r.status === 0 && j.ok, `失败：${r.stdout}`);
    assert(j.report.figures >= 1 && j.report.notes.some((n) => n.includes("渲染为 288 dpi")), `应渲染出图片：${JSON.stringify(j.report).slice(0, 400)}`);
    const buf = fs.readFileSync(out);
    assert(buf.includes(Buffer.from("word/media/")), "Word 里没有图片");
    sameAsDev("outlined-table.pdf", documentXml(out));
    return `${j.report.figures} 幅图，${j.report.ms} ms`;
  });

  check("scan-text.pdf（文字画在整页扫描图上面）：report.scanPages 分清文字页、整页成图的页和没读出来的页，与开发环境一致", () => {
    const { r, j, out } = convert("scan-text.pdf", ["--mode", "flow"]);
    assert(r.status === 0 && j.ok, `失败：${r.stdout}`);
    assert(JSON.stringify(j.report.scanPages) === JSON.stringify({ text: [1, 2, 3, 6], picture: [4], unread: [5] }), `scanPages：${JSON.stringify(j.report.scanPages)}`);
    assert(j.report.textCoverage === 1, `textCoverage=${j.report.textCoverage}`);
    const dev = path.join(work, "scan-text-dev.docx");
    const d = spawnSync(process.execPath, ["--import", "tsx", path.join(root, "cli/pdf2word.ts"), "convert", path.join(fixtures, "scan-text.pdf"), dev, "--mode", "flow"], { cwd: root, encoding: "utf8", timeout: 180_000 });
    assert(d.status === 0 && documentXml(dev) === documentXml(out), "与开发环境的结果不一致");
    return `${j.report.images} 张图，${j.report.ms} ms`;
  });

  check("副本里的渲染原生模块损坏：engine_incomplete，不静默丢掉矢量内容", () => {
    const copy = fs.mkdtempSync(path.join(os.tmpdir(), "pdf2word-canvas-"));
    try {
      copyTree(path.join(tmp, "pdf2word"), path.join(copy, "pdf2word"));
      fs.writeFileSync(path.join(copy, "pdf2word/node_modules/@napi-rs/canvas-linux-x64-gnu/skia.linux-x64-gnu.node"), "not a module");
      const r = run(["convert", path.join(work, "missing.pdf"), path.join(work, "x.docx")], path.join(copy, "pdf2word/pdf2word.mjs"));
      const j = parse(r);
      assert(r.status === 2 && j.code === "engine_incomplete", r.stdout);
    } finally {
      fs.rmSync(copy, { recursive: true, force: true });
    }
  });
} else {
  console.log(`  - 本机不是 linux-x64 glibc：跳过渲染的冒烟测试（发布包只带这个平台的渲染组件）`);
}

fs.rmSync(tmp, { recursive: true, force: true });
const failed = results.filter((x) => !x).length;
console.log(`${results.length - failed}/${results.length} 通过（${((Date.now() - t0) / 1000).toFixed(1)} 秒）`);
if (failed) process.exit(1);
