/**
 * 上下文压缩（对应 Claude Code 的 services/compact）
 *
 * - 微压缩（microcompact）：较早的大段工具输出（例如读取整章的结果）折叠成一行说明，
 *   模型需要时可以重新读取。只折叠一次并写回历史，保证之后的请求前缀稳定、继续命中缓存。
 * - 自动压缩（autocompact）：估算的上下文超过预算的一定比例时，让模型把整段对话总结成结构化
 *   摘要，用摘要替换历史。/compact 命令可手动触发并附加"摘要时重点保留什么"的说明。
 */

import type { ApiMessage, ModelClient } from "./api/deepseek.js";

const FOLDED_MARK = "[旧的工具输出已折叠";

export function microCompact(history: ApiMessage[], keepRecent = 6, minChars = 1500): number {
  const toolIdx = history.map((m, i) => (m.role === "tool" ? i : -1)).filter((i) => i >= 0);
  const candidates = toolIdx.slice(0, Math.max(0, toolIdx.length - keepRecent));
  let saved = 0;
  for (const i of candidates) {
    const m = history[i] as Extract<ApiMessage, { role: "tool" }>;
    if (m.content.length < minChars || m.content.startsWith(FOLDED_MARK)) continue;
    const firstLine = m.content.split("\n")[0].slice(0, 120);
    const folded = `${FOLDED_MARK}以节省上下文（原 ${m.content.length} 字，开头：${firstLine}）。如需其中内容，请重新调用相应工具读取。]`;
    saved += m.content.length - folded.length;
    m.content = folded;
  }
  return saved;
}

function transcript(history: ApiMessage[]): string {
  const out: string[] = [];
  for (const m of history) {
    if (m.role === "user") out.push(`【用户】${m.content.replace(/<system-reminder>[\s\S]*?<\/system-reminder>/g, "").trim()}`);
    else if (m.role === "assistant") {
      if (m.content) out.push(`【助手】${m.content}`);
      for (const t of m.tool_calls ?? []) out.push(`【调用 ${t.function.name}】${t.function.arguments.slice(0, 400)}`);
    } else if (m.role === "tool") out.push(`【工具结果】${m.content.slice(0, 600)}`);
  }
  return out.join("\n");
}

const COMPACT_PROMPT = `请把上面的对话整理成一份供你自己继续工作的摘要。必须包含：
1. 用户的目标与所有明确要求（原话要点）、偏好与禁忌；
2. 已完成的修改：逐条列出段落引用（#开头的引用）、改了什么；
3. 已发现但尚未处理的问题、待办事项与下一步计划；
4. 文档的关键结构信息（重要章节及其引用、使用的样式）；
5. 用户拒绝过的操作及原因。
只输出摘要本身，使用简洁的中文条目。`;

export async function autoCompact(client: ModelClient, history: ApiMessage[], signal: AbortSignal, instructions?: string, midTurn = false): Promise<{ history: ApiMessage[]; summary: string }> {
  const res = await client.complete(
    {
      system: "你是一个对话摘要助手，负责为文档编辑代理压缩上下文。",
      messages: [{ role: "user", content: `${transcript(history)}\n\n---\n${COMPACT_PROMPT}${instructions ? `\n额外要求：${instructions}` : ""}` }],
      tools: [],
    },
    { signal }
  );
  const summary = res.content.trim();
  const head: ApiMessage = { role: "user", content: `<system-reminder>\n此前的对话已被压缩，以下是摘要：\n${summary}\n</system-reminder>\n${midTurn ? "请基于摘要继续完成当前任务。" : "以上是此前工作的摘要。"}` };
  // 轮中压缩：以用户消息结尾，模型直接继续；轮间压缩（/compact）：补一条助手确认，保持角色交替
  return { summary, history: midTurn ? [head] : [head, { role: "assistant", content: "好的，我已了解之前的进展。", reasoning_content: "" }] };
}
