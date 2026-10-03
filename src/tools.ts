/**
 * 工具注册表（对应 Claude Code 的 src/tools.ts）
 *
 * 新增工具 = 在这里加一行。assembleToolPool 根据上下文过滤出"此刻可用"的工具：
 *   - isEnabled：文档格式不支持的能力（例如 PDF 的写操作）不出现；
 *   - 子代理白名单：审阅代理只拿到只读工具；
 *   - 子代理不能再启动子代理（避免无限递归）。
 */

import type { Tool, ToolUseContext } from "./Tool.js";
import { DOC_TOOLS } from "./tools/docTools.js";
import { META_TOOLS } from "./tools/metaTools.js";
import { EDITING_TOOLS } from "./tools/editingTools.js";

export const ALL_TOOLS: Tool[] = [...DOC_TOOLS, ...EDITING_TOOLS, ...META_TOOLS];

export function findTool(name: string): Tool | undefined {
  return ALL_TOOLS.find((t) => t.name === name);
}

export function assembleToolPool(ctx: ToolUseContext, allowList?: string[]): Tool[] {
  return ALL_TOOLS.filter((t) => {
    if (allowList && !allowList.includes(t.name)) return false;
    if (ctx.agentId !== "main" && t.name === "agent") return false;
    return t.isEnabled(ctx);
  });
}
