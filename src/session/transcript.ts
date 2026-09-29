/** 从完整消息历史还原给界面展示的对话记录（刷新页面 / 恢复会话时使用） */

import type { ApiMessage } from "../services/api/deepseek.js";
import { findTool } from "../tools.js";

export type TranscriptItem =
  | { kind: "user"; text: string }
  | { kind: "assistant"; text: string }
  | { kind: "tool"; id: string; name: string; title: string; ok: boolean; content: string };

export function buildTranscript(history: ApiMessage[]): TranscriptItem[] {
  const out: TranscriptItem[] = [];
  const pending = new Map<string, { name: string; title: string }>();
  for (const m of history) {
    if (m.role === "user") {
      const text = m.content.replace(/<system-reminder>[\s\S]*?<\/system-reminder>\s*/g, "").trim();
      if (text) out.push({ kind: "user", text: text.startsWith("用户调用了 /") ? text.split("\n")[0] : text });
    } else if (m.role === "assistant") {
      if (m.content?.trim()) out.push({ kind: "assistant", text: m.content });
      for (const t of m.tool_calls ?? []) {
        let title = t.function.name;
        try {
          const tool = findTool(t.function.name);
          if (tool) title = tool.userFacingName(JSON.parse(t.function.arguments || "{}"));
        } catch { /* 忽略 */ }
        pending.set(t.id, { name: t.function.name, title });
      }
    } else if (m.role === "tool") {
      const p = pending.get(m.tool_call_id);
      out.push({ kind: "tool", id: m.tool_call_id, name: p?.name ?? "tool", title: p?.title ?? "工具", ok: !m.content.startsWith("错误："), content: m.content });
    }
  }
  return out;
}
