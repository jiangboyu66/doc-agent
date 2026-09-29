/**
 * 桥接协议：引擎 → 前端 的事件流（对应 Claude Code bridge/ 中 IDE 与 CLI 之间的 JSON 消息）
 *
 * QueryEngine 以异步生成器的方式产出这些事件；Web 端通过 SSE 接收，CLI 直接在进程内消费。
 * 两个前端共享同一套协议和同一个引擎，所以行为完全一致。
 */

export interface PermissionRequest {
  requestId: string;
  toolUseId: string;
  tool: string;
  title: string;
  input: unknown;
  reason: string;
  /** 干跑得到的修改预览 */
  preview?: Array<{ ref: string; before: string; after: string }>;
  /** 选择"以后都允许"时写入的规则 */
  suggestedRule: string;
  agentId: string;
}

export type PermissionResponse =
  | { decision: "allow"; remember?: "session" | "project" }
  | { decision: "deny"; feedback?: string };

export interface TodoItem {
  content: string;
  activeForm: string;
  status: "pending" | "in_progress" | "completed";
}

export type EngineEvent =
  | { type: "turn_start"; turnId: string }
  | { type: "text_delta"; text: string; agentId: string }
  | { type: "reasoning_delta"; text: string; agentId: string }
  | { type: "assistant_message"; text: string; agentId: string }
  | { type: "tool_use"; id: string; name: string; title: string; input: unknown; agentId: string }
  | { type: "tool_result"; id: string; name: string; ok: boolean; content: string; preview?: PermissionRequest["preview"]; data?: unknown; agentId: string }
  | { type: "permission_request"; request: PermissionRequest }
  | { type: "permission_resolved"; requestId: string; decision: "allow" | "deny" }
  | { type: "doc_changed"; version: number; label: string; changedRefs: string[]; structural: boolean }
  | { type: "todos"; todos: TodoItem[] }
  | { type: "usage"; promptTokens: number; cachedTokens: number; completionTokens: number; costUsd: number; totalCostUsd: number }
  | { type: "compact"; kind: "micro" | "auto" | "manual"; before: number; after: number }
  | { type: "hook"; event: string; message: string }
  | { type: "mode_changed"; mode: string }
  | { type: "subagent"; agentId: string; agentType: string; status: "start" | "done"; description: string }
  | { type: "command_output"; text: string }
  | { type: "retry"; attempt: number; delayMs: number; reason: string }
  | { type: "error"; message: string }
  | { type: "done"; reason: "completed" | "max_turns" | "interrupted" | "error" | "budget" | "command" };
