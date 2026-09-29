import type { ChatItem, EngineEvent, PermissionRequest, TranscriptItem } from "./types";

let seq = 0;
const key = () => `k${++seq}`;

/** 从服务端保存的对话记录还原界面 */
export function fromTranscript(t: TranscriptItem[], pending: PermissionRequest[]): ChatItem[] {
  const items: ChatItem[] = t.map((x) => {
    if (x.kind === "user") return x.text.startsWith("/") || x.text.startsWith("用户调用了 /")
      ? { kind: "user", key: key(), text: x.text.replace(/^用户调用了 /, "") }
      : { kind: "user", key: key(), text: x.text };
    if (x.kind === "assistant") return { kind: "assistant", key: key(), text: x.text, reasoning: "", streaming: false, agentId: "main" };
    return { kind: "tool", key: key(), id: x.id, name: x.name, title: x.title, status: x.ok ? "ok" : "error", content: x.content, agentId: "main" };
  });
  for (const p of pending) items.push({ kind: "permission", key: key(), request: p, status: "pending" });
  return items;
}

/** 把一个引擎事件合并进对话列表（纯函数，便于 React 批量更新） */
export function applyEvent(items: ChatItem[], e: EngineEvent): ChatItem[] {
  switch (e.type) {
    case "text_delta":
    case "reasoning_delta": {
      if (e.agentId !== "main") return items;
      const last = items[items.length - 1];
      if (last?.kind === "assistant" && last.streaming && last.agentId === e.agentId) {
        const next = e.type === "text_delta" ? { ...last, text: last.text + e.text } : { ...last, reasoning: last.reasoning + e.text };
        return [...items.slice(0, -1), next];
      }
      return [...items, { kind: "assistant", key: key(), text: e.type === "text_delta" ? e.text : "", reasoning: e.type === "reasoning_delta" ? e.text : "", streaming: true, agentId: e.agentId }];
    }
    case "assistant_message": {
      if (e.agentId !== "main") return items;
      // 结束流式状态；若没有收到增量（极少数情况）则补上整段文字
      const idx = findLastIndex(items, (x) => x.kind === "assistant" && x.streaming);
      if (idx === -1) return [...items, { kind: "assistant", key: key(), text: e.text, reasoning: "", streaming: false, agentId: e.agentId }];
      const it = items[idx] as Extract<ChatItem, { kind: "assistant" }>;
      return replaceAt(items, idx, { ...it, text: it.text || e.text, streaming: false });
    }
    case "tool_use": {
      const closed = closeStreaming(items);
      return [...closed, { kind: "tool", key: key(), id: e.id, name: e.name, title: e.title, status: "running", content: "", agentId: e.agentId }];
    }
    case "tool_result": {
      const idx = findLastIndex(items, (x) => x.kind === "tool" && x.id === e.id);
      if (idx === -1) return items;
      const it = items[idx] as Extract<ChatItem, { kind: "tool" }>;
      return replaceAt(items, idx, { ...it, status: e.ok ? "ok" : "error", content: e.content, preview: e.preview, data: e.data });
    }
    case "permission_request":
      return [...closeStreaming(items), { kind: "permission", key: key(), request: e.request, status: "pending" }];
    case "permission_resolved": {
      const idx = findLastIndex(items, (x) => x.kind === "permission" && x.request.requestId === e.requestId);
      if (idx === -1) return items;
      return replaceAt(items, idx, { ...(items[idx] as Extract<ChatItem, { kind: "permission" }>), status: e.decision });
    }
    case "subagent": {
      if (e.status === "start") return [...closeStreaming(items), { kind: "subagent", key: key(), agentId: e.agentId, agentType: e.agentType, description: e.description, status: "running" }];
      const idx = findLastIndex(items, (x) => x.kind === "subagent" && x.agentId === e.agentId);
      if (idx === -1) return items;
      return replaceAt(items, idx, { ...(items[idx] as Extract<ChatItem, { kind: "subagent" }>), status: "done" });
    }
    case "doc_changed":
      return [...items, { kind: "notice", key: key(), tone: "success", text: `已保存为版本 v${e.version} · ${e.label}` }];
    case "compact":
      return [...items, { kind: "notice", key: key(), tone: "info", text: `上下文已压缩（${e.kind === "micro" ? "折叠旧输出" : e.kind === "auto" ? "自动摘要" : "手动摘要"}）：约 ${e.before.toLocaleString()} → ${e.after.toLocaleString()} tokens` }];
    case "hook":
      return [...items, { kind: "notice", key: key(), tone: "warn", text: `Hook（${e.event}）：${e.message}` }];
    case "mode_changed":
      return [...items, { kind: "notice", key: key(), tone: "info", text: `权限模式已切换为 ${e.mode}` }];
    case "retry":
      return [...items, { kind: "notice", key: key(), tone: "warn", text: `${e.reason} ${Math.round(e.delayMs / 1000)} 秒后第 ${e.attempt} 次重试…` }];
    case "command_output":
      return [...items, { kind: "command", key: key(), text: e.text }];
    case "error":
      return [...closeStreaming(items), { kind: "notice", key: key(), tone: "error", text: e.message }];
    case "done": {
      let out = closeStreaming(items);
      // 被中断时仍在运行的工具标记为失败
      out = out.map((x) => (x.kind === "tool" && x.status === "running" ? { ...x, status: "error", content: x.content || "未完成" } : x));
      if (e.reason === "interrupted") out = [...out, { kind: "notice", key: key(), tone: "warn", text: "已中断" }];
      return out;
    }
    default:
      return items;
  }
}

export function userItem(text: string): ChatItem {
  return { kind: "user", key: key(), text };
}

function closeStreaming(items: ChatItem[]): ChatItem[] {
  const idx = findLastIndex(items, (x) => x.kind === "assistant" && x.streaming);
  if (idx === -1) return items;
  return replaceAt(items, idx, { ...(items[idx] as Extract<ChatItem, { kind: "assistant" }>), streaming: false });
}

function findLastIndex<T>(arr: T[], fn: (x: T) => boolean): number {
  for (let i = arr.length - 1; i >= 0; i--) if (fn(arr[i])) return i;
  return -1;
}

function replaceAt<T>(arr: T[], i: number, v: T): T[] {
  const out = arr.slice();
  out[i] = v;
  return out;
}
