/**
 * 权限系统（对应 Claude Code 的 hooks/toolPermission 与 PermissionContext）
 *
 * 决策顺序（先命中先生效）：
 *   1. deny 规则         —— 任何模式下都拒绝
 *   2. ask 规则          —— 任何模式下都询问（哪怕是 bypassPermissions）
 *   3. 工具自身的检查    —— 例如一次删除大量段落、导出覆盖等，工具可以要求强制询问或直接拒绝
 *   4. bypassPermissions —— 全部放行
 *   5. 只读工具          —— 放行
 *   6. plan 模式         —— 拒绝一切写操作，并提示模型先用 exit_plan_mode 提交计划
 *   7. allow 规则        —— 放行（来源：会话内"本次会话都允许"、项目/用户配置）
 *   8. acceptEdits 模式  —— 非破坏性的编辑自动放行
 *   9. 其余              —— 询问用户（前端弹出修改预览）
 *
 * 规则语法：工具名 或 工具名(内容模式)，内容模式支持 * 通配，例如 doc_export(pdf)、doc_replace_text、doc_delete_blocks(*)
 */

import type { Tool, ToolUseContext } from "../Tool.js";
import type { PermissionMode } from "../config/settings.js";

export type Decision =
  | { behavior: "allow"; source: string }
  | { behavior: "deny"; message: string; source: string }
  | { behavior: "ask"; reason: string };

export function ruleFor(tool: Tool, input: unknown): string {
  const c = tool.permissionContent(input);
  return c ? `${tool.name}(${c})` : tool.name;
}

export function matchRule(rule: string, tool: Tool, input: unknown): boolean {
  const m = /^([A-Za-z0-9_]+)(?:\((.*)\))?$/.exec(rule.trim());
  if (!m) return false;
  if (m[1] !== tool.name) return false;
  if (m[2] === undefined || m[2] === "*") return true;
  const content = tool.permissionContent(input);
  const re = new RegExp("^" + m[2].split("*").map((s) => s.replace(/[.+?^${}()|[\]\\]/g, "\\$&")).join(".*") + "$");
  return re.test(content);
}

export async function hasPermissionsToUseTool(tool: Tool, input: any, ctx: ToolUseContext, mode: PermissionMode): Promise<Decision> {
  const rules = ctx.settings.permissions;
  const sessionAllow = ctx.session.meta.sessionRules;

  const deny = rules.deny.find((r) => matchRule(r, tool, input));
  if (deny) return { behavior: "deny", message: `被配置中的拒绝规则 "${deny}" 禁止。`, source: "deny-rule" };

  const ask = rules.ask.find((r) => matchRule(r, tool, input));
  if (ask) return { behavior: "ask", reason: `配置要求此类操作每次都需确认（规则 "${ask}"）。` };

  const own = await tool.checkPermissions(input, ctx);
  if (own.behavior === "deny") return { behavior: "deny", message: own.message, source: "tool" };
  if (own.behavior === "ask") return { behavior: "ask", reason: own.reason };
  if (own.behavior === "allow") return { behavior: "allow", source: "tool" };

  if (mode === "bypassPermissions") return { behavior: "allow", source: "bypass" };
  if (tool.isReadOnly(input)) return { behavior: "allow", source: "readonly" };

  if (mode === "plan") {
    return {
      behavior: "deny",
      message: "当前处于计划模式：不能修改文档。请先完成调研，然后调用 exit_plan_mode 把完整的修改计划提交给用户确认。",
      source: "plan-mode",
    };
  }

  const allow = [...sessionAllow, ...rules.allow].find((r) => matchRule(r, tool, input));
  if (allow) return { behavior: "allow", source: `allow-rule:${allow}` };

  if (mode === "acceptEdits" && tool.category === "edit" && !tool.isDestructive(input)) {
    return { behavior: "allow", source: "acceptEdits" };
  }
  return { behavior: "ask", reason: tool.isDestructive(input) ? "这是一个删除/批量修改类操作，需要你确认。" : "这个操作会修改文档，需要你确认。" };
}
