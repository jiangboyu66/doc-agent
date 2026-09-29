#!/usr/bin/env node
/**
 * 命令行入口：与 Web 共用 QueryEngine、工具、权限、命令与会话存储。
 *
 *   npm run cli -- 论文.docx                    打开文档进入交互模式
 *   npm run cli -- --resume <会话ID>            恢复会话
 *   npm run cli -- 论文.docx -p "把摘要润色一下" --mode acceptEdits --output 结果.docx
 *                                              无人值守模式（headless）：执行后导出结果退出
 */

import "dotenv/config";
import { Command } from "commander";
import chalk from "chalk";
import prompts from "prompts";
import fs from "node:fs/promises";
import path from "node:path";
import readline from "node:readline";
import { diffWordsWithSpace } from "diff";
import { bootstrap, createClient, type Runtime } from "../bootstrap.js";
import { Session } from "../session/Session.js";
import { QueryEngine } from "../QueryEngine.js";
import { processUserInput } from "../commands.js";
import { exportDocument, type ExportFormat } from "../services/convert.js";
import type { EngineEvent, PermissionRequest, PermissionResponse } from "../bridge/protocol.js";
import type { PermissionMode } from "../config/settings.js";

function showDiff(before: string, after: string): string {
  if (!before) return chalk.green(after);
  if (!after) return chalk.red.strikethrough(before);
  return diffWordsWithSpace(before, after).map((p) => (p.added ? chalk.bgGreen.black(p.value) : p.removed ? chalk.bgRed.white(p.value) : chalk.gray(p.value))).join("");
}

async function askPermission(req: PermissionRequest, headless: boolean): Promise<PermissionResponse> {
  if (headless) return { decision: "deny", feedback: "无人值守模式下不能请求确认。请使用 --mode acceptEdits 或在配置中添加 allow 规则。" };
  console.log("\n" + chalk.yellow.bold(`⚠ 需要确认：${req.title}`) + (req.agentId !== "main" ? chalk.dim(`（来自子代理 ${req.agentId}）`) : ""));
  console.log(chalk.dim(req.reason));
  if (req.tool === "exit_plan_mode") console.log((req.input as any).plan);
  for (const p of (req.preview ?? []).slice(0, 10)) {
    console.log(chalk.cyan(`  ${p.ref}`));
    console.log("  " + showDiff(p.before, p.after).replace(/\n/g, "\n  "));
  }
  const { choice } = await prompts({
    type: "select",
    name: "choice",
    message: "如何处理？",
    choices: [
      { title: "允许", value: "allow" },
      { title: `本次会话都允许（${req.suggestedRule}）`, value: "session" },
      { title: `以后都允许（写入项目配置）`, value: "project" },
      { title: "拒绝并说明原因", value: "deny" },
    ],
  });
  if (choice === "allow") return { decision: "allow" };
  if (choice === "session" || choice === "project") return { decision: "allow", remember: choice };
  const { feedback } = await prompts({ type: "text", name: "feedback", message: "告诉 Agent 你希望怎么做（可留空）" });
  return { decision: "deny", feedback: feedback || undefined };
}

function render(e: EngineEvent, state: { streaming: boolean }) {
  const sub = (id: string) => (id !== "main" ? chalk.magenta(`[${id}] `) : "");
  switch (e.type) {
    case "text_delta":
      if (e.agentId !== "main") return;
      if (!state.streaming) { process.stdout.write("\n" + chalk.bold.blue("Agent › ")); state.streaming = true; }
      process.stdout.write(e.text);
      return;
    case "reasoning_delta":
      return;
    case "tool_use":
      state.streaming = false;
      console.log("\n" + sub(e.agentId) + chalk.dim(`● ${e.title}`));
      return;
    case "tool_result":
      console.log(sub(e.agentId) + (e.ok ? chalk.green("  ✓ ") : chalk.red("  ✗ ")) + chalk.dim(e.content.split("\n")[0].slice(0, 120)));
      return;
    case "doc_changed":
      console.log(chalk.green(`  ↳ 已保存为版本 v${e.version}：${e.label}`));
      return;
    case "todos":
      console.log(chalk.cyan("\n  待办：") + e.todos.map((t) => `${t.status === "completed" ? "✓" : t.status === "in_progress" ? "▶" : "○"} ${t.content}`).join("  "));
      return;
    case "subagent":
      console.log(chalk.magenta(`\n  ${e.status === "start" ? "▸ 启动" : "◂ 完成"}子代理 ${e.agentType}：${e.description}`));
      return;
    case "compact":
      console.log(chalk.dim(`\n  上下文已压缩（${e.kind}）：约 ${e.before} → ${e.after} tokens`));
      return;
    case "retry":
      console.log(chalk.yellow(`\n  ${e.reason} ${Math.round(e.delayMs / 1000)} 秒后第 ${e.attempt} 次重试…`));
      return;
    case "mode_changed":
      console.log(chalk.cyan(`\n  权限模式切换为 ${e.mode}`));
      return;
    case "hook":
      console.log(chalk.yellow(`\n  [Hook ${e.event}] ${e.message}`));
      return;
    case "command_output":
      console.log("\n" + e.text);
      return;
    case "error":
      console.log("\n" + chalk.red(`✗ ${e.message}`));
      return;
    case "usage":
      return;
    case "done":
      if (state.streaming) process.stdout.write("\n");
      state.streaming = false;
      return;
  }
}

async function runTurn(text: string, session: Session, runtime: Runtime, headless: boolean): Promise<void> {
  const abort = new AbortController();
  const onSigint = () => abort.abort();
  process.once("SIGINT", onSigint);
  let engine: QueryEngine | null = null;
  const getEngine = () => (engine ??= new QueryEngine(session, {
    client: createClient(runtime.settings),
    settings: runtime.settings,
    skills: runtime.skills,
    agents: runtime.agents,
    memory: runtime.memory,
    canUseTool: (req) => askPermission(req, headless),
  }));
  const state = { streaming: false };
  try {
    const input = await processUserInput(text, { session, runtime, compact: (ins) => getEngine().compactNow(abort.signal, ins) });
    if (input.kind === "local") return render({ type: "command_output", text: input.output }, state);
    for await (const e of getEngine().submitMessage(input.prompt, abort.signal)) render(e, state);
  } catch (e: any) {
    render({ type: "error", message: e?.message ?? String(e) }, state);
  } finally {
    process.removeListener("SIGINT", onSigint);
    await session.save();
  }
}

async function main() {
  const program = new Command();
  program
    .name("doc-agent")
    .description("文案 Agent：保真的文档编辑助手（DeepSeek）")
    .argument("[file]", "要编辑的文档（.docx / .md / .html / .pdf）")
    .option("--resume <id>", "恢复已有会话")
    .option("-p, --print <prompt>", "无人值守模式：执行这条指令后退出")
    .option("--output <file>", "无人值守模式结束后导出到该文件（按扩展名决定格式）")
    .option("--mode <mode>", "权限模式 default|acceptEdits|plan|bypassPermissions")
    .option("--track", "开启 Word 修订模式")
    .option("--model <model>", "模型，如 deepseek-flash / deepseek-v4-pro")
    .option("--author <name>", "修订与批注的作者名");
  program.parse();
  const opts = program.opts();
  const file = program.args[0];

  const runtime = await bootstrap({
    ...(opts.model ? { model: opts.model } : {}),
    ...(opts.mode ? { permissionMode: opts.mode as PermissionMode } : {}),
    ...(opts.track ? { trackChanges: true } : {}),
    ...(opts.author ? { author: opts.author } : {}),
  });

  let session: Session;
  if (opts.resume) session = await Session.load(opts.resume);
  else if (file) {
    const buf = await fs.readFile(path.resolve(file));
    session = await Session.create(buf, path.basename(file), { mode: runtime.settings.permissionMode, trackChanges: runtime.settings.trackChanges, author: runtime.settings.author });
  } else {
    const list = await Session.list();
    console.log("用法：doc-agent <文件> 或 doc-agent --resume <会话ID>\n\n最近的会话：");
    for (const m of list.slice(0, 10)) console.log(`  ${m.id}  ${m.filename}  v${m.currentVersion}  ${new Date(m.updatedAt).toLocaleString("zh-CN", { hour12: false })}`);
    return;
  }
  if (opts.mode) session.meta.mode = opts.mode;
  if (opts.track) session.meta.trackChanges = session.doc.capabilities.trackChanges;

  const s = session.doc.summary();
  console.log(chalk.bold(`\n文案 Agent · ${session.meta.filename}`) + chalk.dim(`（${s.format}，${s.blockCount} 段，约 ${s.charCount} 字）`));
  console.log(chalk.dim(`会话 ${session.id}｜模型 ${runtime.settings.model}｜模式 ${session.meta.mode}｜修订 ${session.meta.trackChanges ? "开" : "关"}｜输入 /help 查看命令`));
  if (s.template) console.log(chalk.cyan(`检测到模板：${s.template.name}`));

  if (opts.print) {
    await runTurn(opts.print, session, runtime, true);
    if (opts.output) {
      const ext = path.extname(opts.output).slice(1).toLowerCase();
      const fmt = (ext === "md" ? "markdown" : ext) as ExportFormat;
      const out = await exportDocument(session.current(), session.meta.format, fmt);
      await fs.writeFile(opts.output, out.data);
      console.log(chalk.green(`\n已导出到 ${opts.output}`) + chalk.dim(`（${out.note}）`));
    }
    return;
  }

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const ask = () => new Promise<string | null>((resolve) => {
    rl.question(chalk.bold.green("\n你 › "), resolve);
    rl.once("close", () => resolve(null));
  });
  for (;;) {
    const line = await ask();
    if (line === null || ["/exit", "/quit"].includes(line.trim())) break;
    if (!line.trim()) continue;
    rl.pause();
    await runTurn(line, session, runtime, false);
    rl.resume();
  }
  rl.close();
  console.log(chalk.dim(`\n会话已保存，可用 --resume ${session.id} 继续。`));
}

main().catch((e) => {
  console.error(chalk.red(e?.message ?? e));
  process.exit(1);
});
