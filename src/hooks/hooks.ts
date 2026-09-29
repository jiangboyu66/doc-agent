/**
 * Hook 系统：在特定时机执行自定义逻辑（对应 Claude Code 的 hooks）
 *
 * 两类 Hook：
 *   - 内置 Hook（函数）：例如每次写操作后的"保真守卫"，发现文档结构损坏就自动回滚这次修改；
 *   - 配置 Hook（命令）：用户在 settings.json 中配置 shell 命令，事件数据以 JSON 从 stdin 传入。
 *     退出码 0 = 通过（stdout 可返回 JSON：{"decision":"block","reason":"…"} 或 {"additionalContext":"…"}）；
 *     退出码 2 = 阻止，stderr 内容反馈给模型；其它退出码 = 警告，不阻止。
 *
 * 事件：SessionStart / UserPromptSubmit / PreToolUse / PostToolUse / PreCompact / Stop
 */

import { spawn } from "node:child_process";
import type { ToolUseContext } from "../Tool.js";

export type HookEvent = "SessionStart" | "UserPromptSubmit" | "PreToolUse" | "PostToolUse" | "PreCompact" | "Stop";

export interface HookInput {
  event: HookEvent;
  sessionId: string;
  documentFormat: string;
  toolName?: string;
  toolInput?: unknown;
  toolResult?: { ok: boolean; content: string };
  prompt?: string;
}

export interface HookOutcome {
  block?: boolean;
  reason?: string;
  additionalContext?: string;
}

interface InternalHook {
  event: HookEvent;
  name: string;
  matcher?: RegExp;
  fn(input: HookInput, ctx: ToolUseContext): Promise<HookOutcome | void>;
}

const internalHooks: InternalHook[] = [];

export function registerInternalHook(h: InternalHook): void {
  internalHooks.push(h);
}

function runCommand(command: string, input: HookInput, timeoutMs: number): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((resolve) => {
    const child = spawn(command, { shell: true, windowsHide: true });
    let stdout = "", stderr = "";
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      resolve({ code: 1, stdout, stderr: `Hook 超时（${timeoutMs / 1000}s）` });
    }, timeoutMs);
    child.stdout.on("data", (d) => (stdout += d));
    child.stderr.on("data", (d) => (stderr += d));
    child.on("error", (e) => {
      clearTimeout(timer);
      resolve({ code: 1, stdout, stderr: e.message });
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      resolve({ code: code ?? 1, stdout, stderr });
    });
    child.stdin.end(JSON.stringify(input));
  });
}

export async function runHooks(event: HookEvent, input: Omit<HookInput, "event">, ctx: ToolUseContext, enabled: boolean): Promise<HookOutcome> {
  const full: HookInput = { event, ...input };
  const outcome: HookOutcome = {};
  const reasons: string[] = [];
  const contexts: string[] = [];
  const matches = (m: string | RegExp | undefined) => {
    if (!m) return true;
    if (!input.toolName) return true;
    return typeof m === "string" ? new RegExp(`^(?:${m})$`).test(input.toolName) : m.test(input.toolName);
  };

  for (const h of internalHooks.filter((x) => x.event === event && matches(x.matcher))) {
    const r = await h.fn(full, ctx);
    if (r?.block) { outcome.block = true; reasons.push(r.reason ?? h.name); }
    if (r?.additionalContext) contexts.push(r.additionalContext);
  }

  if (enabled) {
    for (const group of ctx.settings.hooks[event] ?? []) {
      if (!matches(group.matcher)) continue;
      for (const hk of group.hooks) {
        const res = await runCommand(hk.command, full, (hk.timeout ?? 30) * 1000);
        if (res.code === 2) {
          outcome.block = true;
          reasons.push(res.stderr.trim() || `Hook "${hk.command}" 阻止了此操作`);
        } else if (res.code === 0) {
          const out = res.stdout.trim();
          if (out.startsWith("{")) {
            try {
              const j = JSON.parse(out);
              if (j.decision === "block") { outcome.block = true; reasons.push(j.reason ?? hk.command); }
              if (j.additionalContext) contexts.push(String(j.additionalContext));
            } catch { /* 非 JSON 输出忽略 */ }
          } else if (out && event === "UserPromptSubmit") contexts.push(out);
        } else {
          ctx.emit({ type: "hook", event, message: `Hook "${hk.command}" 返回 ${res.code}：${res.stderr.trim().slice(0, 300)}` });
        }
      }
    }
  }
  if (reasons.length) outcome.reason = reasons.join("；");
  if (contexts.length) outcome.additionalContext = contexts.join("\n");
  return outcome;
}
