/**
 * 上下文收集（对应 Claude Code 的 context.ts）
 *
 * - 长期记忆：项目 DOCAGENT.md 与用户 ~/.doc-agent/DOCAGENT.md（类似 CLAUDE.md），写入系统提示词；
 * - 当轮提醒：时间、权限模式、修订模式、文档版本、待办进度等每轮可能变化的状态，
 *   以 <system-reminder> 附在当轮用户消息上，保持历史前缀稳定以命中缓存。
 */

import fs from "node:fs/promises";
import path from "node:path";
import { PATHS } from "./config/settings.js";
import type { Session } from "./session/Session.js";

export async function loadMemory(): Promise<string> {
  const files = [path.join(PATHS.userDir, "DOCAGENT.md"), path.join(process.cwd(), "DOCAGENT.md")];
  const parts = await Promise.all(files.map((f) => fs.readFile(f, "utf8").catch(() => "")));
  return parts.filter((p) => p.trim()).join("\n\n");
}

export async function appendProjectMemory(text: string): Promise<string> {
  const file = path.join(process.cwd(), "DOCAGENT.md");
  let cur = "";
  try { cur = await fs.readFile(file, "utf8"); } catch { cur = "# 文案 Agent 长期偏好\n"; }
  const next = `${cur.trimEnd()}\n- ${text.trim()}\n`;
  await fs.writeFile(file, next, "utf8");
  return file;
}

const MODE_TEXT: Record<string, string> = {
  default: "标准模式：修改文档的操作需要用户逐一确认。",
  acceptEdits: "自动接受编辑模式：普通编辑自动执行，删除等破坏性操作仍需确认。",
  plan: "计划模式：只能读取与分析，不能修改文档。调研完成后调用 exit_plan_mode 提交完整修改计划，等待用户批准。",
  bypassPermissions: "免确认模式：所有操作直接执行，务必格外谨慎。",
};

export function buildTurnReminder(session: Session, extra?: string): string {
  const m = session.meta;
  const lines = [
    `当前时间：${new Date().toLocaleString("zh-CN", { hour12: false })}`,
    `权限模式：${MODE_TEXT[m.mode] ?? m.mode}`,
  ];
  if (session.doc.capabilities.trackChanges) {
    lines.push(m.trackChanges ? `修订模式：开启（作者"${m.author}"），所有修改会记录为 Word 修订，用户可逐条接受或拒绝。` : "修订模式：关闭，修改直接生效（每次修改都有版本快照可回滚）。");
  }
  lines.push(`文档版本：v${m.currentVersion}${m.currentVersion > 0 ? `（相对原始文件已有 ${m.versions.length - 1} 次修改）` : "（未修改）"}；结构版本 v${session.doc.structureVersion}`);
  if (m.todos.length) {
    lines.push("待办清单：" + m.todos.map((t) => `${t.status === "completed" ? "✓" : t.status === "in_progress" ? "▶" : "○"} ${t.content}`).join("；"));
  }
  if (m.assets?.length) lines.push(`已上传素材（可用 doc_insert_image 插入）：${m.assets.join("、")}`);
  if (extra) lines.push(extra);
  return `<system-reminder>\n${lines.join("\n")}\n</system-reminder>`;
}
