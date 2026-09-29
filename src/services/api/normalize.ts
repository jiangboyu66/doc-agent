/**
 * normalizeMessagesForAPI：发送前的消息序列体检与修复（纵深防御）
 *
 * 早期版本出现过 400 错误 "Messages with role 'tool' must be a response to a preceding message with
 * 'tool_calls'"——根因是历史被错误地精简。现在历史总是完整保存，但在发送前仍然做一次结构校验：
 *   1. 每个 assistant.tool_calls 的 id 都必须紧跟对应的 tool 消息；缺失的补一条"未执行"的结果；
 *   2. 没有对应 tool_call 的孤立 tool 消息直接丢弃；
 *   3. 思考模式下每条 assistant 消息都带上 reasoning_content（DeepSeek 要求回传）。
 */

import type { ApiMessage } from "./deepseek.js";

export function normalizeMessagesForAPI(messages: ApiMessage[], thinking: boolean): ApiMessage[] {
  const out: ApiMessage[] = [];
  let pending: string[] = [];
  const flushPending = () => {
    for (const id of pending) out.push({ role: "tool", tool_call_id: id, content: "（该工具调用被中断，未执行）" });
    pending = [];
  };
  for (const m of messages) {
    if (m.role === "tool") {
      const idx = pending.indexOf(m.tool_call_id);
      if (idx === -1) continue; // 孤立的 tool 消息
      pending.splice(idx, 1);
      out.push({ role: "tool", tool_call_id: m.tool_call_id, content: m.content });
      continue;
    }
    flushPending();
    if (m.role === "assistant") {
      const a: ApiMessage = { role: "assistant", content: m.content ?? "" };
      if (m.tool_calls?.length) {
        a.tool_calls = m.tool_calls.map((t) => ({ id: t.id, type: "function", function: { name: t.function.name, arguments: t.function.arguments || "{}" } }));
        pending = a.tool_calls.map((t) => t.id);
      }
      if (thinking) a.reasoning_content = m.reasoning_content ?? "";
      out.push(a);
    } else {
      out.push({ role: m.role, content: m.content } as ApiMessage);
    }
  }
  flushPending();
  return out;
}

/** 粗略估算 token 数（中英混排按约 2.5 字符/词元估算，只用于触发压缩的阈值判断） */
export function estimateTokens(messages: ApiMessage[], system = ""): number {
  let chars = system.length;
  for (const m of messages) {
    chars += (m.content ?? "").length;
    if (m.role === "assistant") {
      chars += (m.reasoning_content ?? "").length;
      for (const t of m.tool_calls ?? []) chars += t.function.arguments.length + t.function.name.length;
    }
  }
  return Math.ceil(chars / 2.5);
}
