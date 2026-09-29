/**
 * 外部转换服务（LibreOffice / pandoc）
 *
 * 编辑本身从不经过这里——这里只负责"导出另一种格式的副本"与"PDF 版式预览"。
 * 每个外部进程都有超时（LibreOffice 在 Windows 上遇到配置锁会无限挂起，这是早期版本卡死的根因），
 * 每次调用 LibreOffice 使用独立的临时用户配置目录，避免与用户已打开的 Office 实例冲突。
 */

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import os from "node:os";
import fs from "node:fs/promises";
import { pathToFileURL } from "node:url";
import type { DocFormat } from "../documents/types.js";

const execFileAsync = promisify(execFile);
const PANDOC = process.env.PANDOC_PATH || "pandoc";
const SOFFICE = process.env.SOFFICE_PATH || "soffice";
const PANDOC_TIMEOUT = 60_000;
const SOFFICE_TIMEOUT = Number(process.env.SOFFICE_TIMEOUT_MS || 120_000);

export type ExportFormat = "docx" | "pdf" | "html" | "markdown";

async function run(cmd: string, args: string[], timeout: number): Promise<void> {
  try {
    await execFileAsync(cmd, args, { timeout, killSignal: "SIGKILL", windowsHide: true, maxBuffer: 64 * 1024 * 1024 });
  } catch (e: any) {
    if (e?.code === "ENOENT") {
      throw new Error(`未找到外部程序 "${cmd}"。${cmd === SOFFICE ? "请安装 LibreOffice 并确保 soffice 在 PATH 中，或在 .env 中设置 SOFFICE_PATH。" : "请安装 pandoc 并确保在 PATH 中，或在 .env 中设置 PANDOC_PATH。"}`);
    }
    if (e?.killed || e?.signal === "SIGKILL") {
      throw new Error(`外部程序 "${cmd}" 超过 ${timeout / 1000} 秒未完成，已终止。若本机开着 LibreOffice，请先关闭后重试。`);
    }
    throw new Error(`外部程序 "${cmd}" 执行失败：${(e?.stderr || e?.message || "").toString().slice(0, 500)}`);
  }
}

async function tmpDir(): Promise<string> {
  return fs.mkdtemp(path.join(os.tmpdir(), "doc-agent-"));
}

/** 用 LibreOffice 把文件转换成 PDF（得到与 Word 排版最接近的版式） */
export async function toPdfViaOffice(input: Buffer, ext: string): Promise<Buffer> {
  const dir = await tmpDir();
  const profile = path.join(dir, "lo-profile");
  try {
    const src = path.join(dir, `input.${ext}`);
    await fs.writeFile(src, input);
    await run(SOFFICE, ["--headless", "--norestore", "--nolockcheck", `-env:UserInstallation=${pathToFileURL(profile).href}`, "--convert-to", "pdf", "--outdir", dir, src], SOFFICE_TIMEOUT);
    return await fs.readFile(path.join(dir, "input.pdf"));
  } finally {
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

async function pandoc(input: Buffer, from: string, to: string, outExt: string, extra: string[] = []): Promise<Buffer> {
  const dir = await tmpDir();
  try {
    const src = path.join(dir, `input.${from === "gfm" ? "md" : from}`);
    const out = path.join(dir, `output.${outExt}`);
    await fs.writeFile(src, input);
    await run(PANDOC, [src, "-f", from, "-t", to, "--wrap=none", "-o", out, ...extra], PANDOC_TIMEOUT);
    return await fs.readFile(out);
  } finally {
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

export interface ExportOutcome {
  data: Buffer;
  ext: string;
  /** 这次导出是否可能损失版式细节（跨格式转换） */
  lossy: boolean;
  note: string;
}

/**
 * 把当前文档导出为目标格式。
 * 同格式导出 = 编辑后的原文件本身（完全保真）；跨格式导出必然是近似转换，会明确标注。
 */
export async function exportDocument(current: Buffer, from: DocFormat, to: ExportFormat): Promise<ExportOutcome> {
  const extOf: Record<ExportFormat, string> = { docx: "docx", pdf: "pdf", html: "html", markdown: "md" };
  if ((from === to) || (from === "markdown" && to === "markdown")) {
    return { data: current, ext: extOf[to], lossy: false, note: "与原文件同格式：导出的就是编辑后的原文件，未修改部分逐字节保持不变。" };
  }
  if (from === "pdf") throw new Error("PDF 为只读来源，不支持导出为其他格式。");
  if (to === "pdf") {
    const ext = from === "markdown" ? "md" : from;
    const source = from === "markdown" ? await pandoc(current, "gfm", "docx", "docx") : current;
    const pdf = await toPdfViaOffice(source, from === "markdown" ? "docx" : ext);
    return { data: pdf, ext: "pdf", lossy: from !== "docx", note: from === "docx" ? "由 LibreOffice 按 Word 版式渲染生成 PDF（字体缺失时可能有细微差异，Word 中“另存为 PDF”效果最准确）。" : "经格式转换后渲染为 PDF。" };
  }
  const fromArg = from === "markdown" ? "gfm" : from;
  const toArg = to === "markdown" ? "gfm" : to;
  const extra = to === "html" ? ["--standalone", "--metadata", "title=Document"] : [];
  const data = await pandoc(current, fromArg, toArg, extOf[to], extra);
  return { data, ext: extOf[to], lossy: true, note: `跨格式转换（${from} → ${to}）只能近似保留结构，版式细节可能丢失；需要完全保真请导出原格式。` };
}

export async function checkExternalTools(): Promise<{ pandoc: string | null; soffice: string | null }> {
  const probe = async (cmd: string) => {
    try {
      const { stdout } = await execFileAsync(cmd, ["--version"], { timeout: 20_000, windowsHide: true });
      return stdout.split(/\r?\n/)[0].trim();
    } catch {
      return null;
    }
  };
  const [p, s] = await Promise.all([probe(PANDOC), probe(SOFFICE)]);
  return { pandoc: p, soffice: s };
}
