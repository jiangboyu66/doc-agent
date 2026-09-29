/**
 * 分层组装系统提示词（对应 Claude Code 的 buildEffectiveSystemPrompt）
 *
 * 优先级：override（完全替换）> 子代理提示词 > 用户自定义 > 默认
 * append 总是追加在末尾（override 时除外）。
 */

import { SYSTEM_PROMPT_DYNAMIC_BOUNDARY } from "../constants/prompts.js";

export function buildEffectiveSystemPrompt(p: {
  defaultPrompt: string[];
  sessionPrompt: string[];
  agentPrompt?: string;
  customPrompt?: string;
  appendPrompt?: string;
  overridePrompt?: string;
}): string {
  if (p.overridePrompt) return p.overridePrompt;
  const base = p.agentPrompt ? [p.agentPrompt] : p.customPrompt ? [p.customPrompt] : p.defaultPrompt;
  const sections = [...base, SYSTEM_PROMPT_DYNAMIC_BOUNDARY, ...p.sessionPrompt];
  if (p.appendPrompt) sections.push(p.appendPrompt);
  // 边界标记本身不发送给模型，它只用于约定"之前的内容跨会话稳定、之后的内容按会话变化"
  return sections.filter((s) => s !== SYSTEM_PROMPT_DYNAMIC_BOUNDARY).join("\n\n");
}
