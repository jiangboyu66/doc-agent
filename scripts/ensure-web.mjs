#!/usr/bin/env node
/**
 * 启动前确认前端（web/dist）是用当前的 web/src 构建的；不是的话自动重新构建。
 *
 * web/dist 不在 git 里：以前 git pull 之后要记得手动 npm run build:web，否则浏览器里跑的还是旧前端
 * （例如上传框不接受新支持的 .doc）。现在 npm run dev / npm start 会先执行本脚本：
 *   - 前端源码没变：什么也不做（只比较一次哈希）；
 *   - 有变化：安装前端依赖（需要时）并重新构建，记录源码哈希到 web/dist/.source-hash；
 *   - 构建失败（例如离线装不了依赖）：提示原因，继续用现有的 web/dist 启动，不阻止服务。
 * 参数 --force：无论是否变化都重新构建（npm run build:web）。环境变量 DOC_AGENT_SKIP_WEB_BUILD=1 可跳过。
 */
import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const web = path.join(root, "web");
const dist = path.join(web, "dist");
const stampFile = path.join(dist, ".source-hash");
const force = process.argv.includes("--force");

if (process.env.DOC_AGENT_SKIP_WEB_BUILD && !force) process.exit(0);
if (!fs.existsSync(path.join(web, "package.json")) || !fs.existsSync(path.join(web, "src"))) process.exit(0);

function sourceHash() {
  const files = [];
  const walk = (dir) => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) walk(p);
      else if (e.isFile()) files.push(p);
    }
  };
  walk(path.join(web, "src"));
  for (const f of ["index.html", "package.json", "package-lock.json", "vite.config.ts", "tsconfig.json"]) {
    if (fs.existsSync(path.join(web, f))) files.push(path.join(web, f));
  }
  const h = createHash("sha256");
  for (const f of files.sort()) {
    h.update(path.relative(web, f).split(path.sep).join("/"));
    h.update("\0");
    // 换行符统一成 LF：Windows 上 git 自动转换 CRLF 不应触发重新构建
    h.update(fs.readFileSync(f).toString("utf8").replace(/\r\n/g, "\n"));
    h.update("\0");
  }
  return h.digest("hex");
}

const want = sourceHash();
const have = fs.existsSync(stampFile) ? fs.readFileSync(stampFile, "utf8").trim() : "";
const built = fs.existsSync(path.join(dist, "index.html"));
if (!force && want === have && built) process.exit(0);

const isWin = process.platform === "win32";
const npm = (args) => spawnSync(isWin ? "npm.cmd" : "npm", args, { cwd: root, stdio: "inherit", shell: isWin }).status === 0;
const depsInstalled = fs.existsSync(path.join(web, "node_modules", ".bin"));
const lockChanged = (() => {
  // 前端依赖有变化（package-lock 比已安装的新）时重新安装
  const lock = path.join(web, "package-lock.json"), installed = path.join(web, "node_modules", ".package-lock.json");
  try { return fs.statSync(lock).mtimeMs > fs.statSync(installed).mtimeMs; } catch { return false; }
})();

// 解压即用的发布包：带有构建好的前端但没有前端依赖，又不是强制构建——不联网安装，只提示
if (!force && built && !depsInstalled && have === "") process.exit(0);

console.log(built ? "\n前端代码有更新，正在重新构建（web/dist）……" : "\n首次构建前端（web/dist）……");
const ok = (depsInstalled && !lockChanged ? true : npm(["--prefix", "web", "install", "--no-audit", "--no-fund"])) && npm(["--prefix", "web", "run", "build"]);
if (!ok) {
  const msg = "前端构建失败。" + (built ? "继续使用现有的 web/dist（可能是旧版本）；" : "") + "可以手动运行 npm run build:web 查看原因。";
  console.warn(`\n⚠ ${msg}\n`);
  process.exit(force ? 1 : 0);
}
fs.writeFileSync(stampFile, want + "\n");
console.log("前端构建完成。\n");
