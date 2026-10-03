// 与服务端 src/bridge/protocol.ts 保持一致的事件协议（前端副本）

export interface PreviewItem { ref: string; before: string; after: string }

export interface PermissionRequest {
  requestId: string;
  toolUseId: string;
  tool: string;
  title: string;
  input: any;
  reason: string;
  preview?: PreviewItem[];
  suggestedRule: string;
  agentId: string;
}

export interface TodoItem { content: string; activeForm: string; status: "pending" | "in_progress" | "completed" }

export type EngineEvent =
  | { type: "turn_start"; turnId: string }
  | { type: "text_delta"; text: string; agentId: string }
  | { type: "reasoning_delta"; text: string; agentId: string }
  | { type: "assistant_message"; text: string; agentId: string }
  | { type: "tool_use"; id: string; name: string; title: string; input: unknown; agentId: string }
  | { type: "tool_result"; id: string; name: string; ok: boolean; content: string; preview?: PreviewItem[]; data?: any; agentId: string }
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
  | { type: "done"; reason: string };

export type PermissionMode = "default" | "acceptEdits" | "plan" | "bypassPermissions";
export type DocFormat = "docx" | "markdown" | "html" | "pdf";

export interface VersionInfo { v: number; label: string; at: number; changedRefs: string[]; format?: DocFormat }

export interface SessionMeta {
  id: string;
  filename: string;
  format: DocFormat;
  createdAt: number;
  updatedAt: number;
  mode: PermissionMode;
  trackChanges: boolean;
  author: string;
  todos: TodoItem[];
  cost: { requests: number; promptTokens: number; cachedTokens: number; completionTokens: number; costUsd: number; apiMs: number };
  versions: VersionInfo[];
  currentVersion: number;
  origin?: { fromSession: string; fromFile: string; engine: string; mode?: string; report: string; coverage: number };
  thinking?: boolean;
  sourceFormat?: DocFormat;
  sourceFile?: string;
}

export interface DocumentSummary {
  format: DocFormat;
  blockCount: number;
  charCount: number;
  headings: Array<{ ref: string; level: number; text: string; location: string }>;
  parts: string[];
  images: number;
  sections: number;
  trackedChanges: number;
  comments: number;
  template?: { id: string; name: string; evidence: string[] };
  notes: string[];
}

export type TranscriptItem =
  | { kind: "user"; text: string }
  | { kind: "assistant"; text: string; reasoning?: string }
  | { kind: "tool"; id: string; name: string; title: string; ok: boolean; content: string };

export interface SessionDetail {
  meta: SessionMeta;
  summary: DocumentSummary;
  transcript: TranscriptItem[];
  busy: boolean;
  pending: PermissionRequest[];
}

export interface RuntimeInfo {
  model: string;
  thinking: string;
  thinkingAvailable?: boolean;
  apiKey: boolean;
  external: { pandoc: string | null; soffice: string | null };
  pdfEngines?: { builtin: boolean; pdf2docx: boolean; libreoffice: boolean };
  features: Record<string, boolean>;
  skills: Array<{ name: string; description: string; source: string }>;
  agents: Array<{ type: string; description: string; readOnly: boolean }>;
  commands: Array<{ name: string; description: string; argumentHint?: string }>;
}

/** 对话区里的一条 */
export type ChatItem =
  | { kind: "user"; key: string; text: string }
  | { kind: "assistant"; key: string; text: string; reasoning: string; streaming: boolean; agentId: string }
  | { kind: "tool"; key: string; id: string; name: string; title: string; status: "running" | "ok" | "error"; content: string; preview?: PreviewItem[]; data?: any; agentId: string }
  | { kind: "permission"; key: string; request: PermissionRequest; status: "pending" | "allow" | "deny" }
  | { kind: "subagent"; key: string; agentId: string; agentType: string; description: string; status: "running" | "done" }
  | { kind: "notice"; key: string; tone: "info" | "warn" | "error" | "success"; text: string }
  | { kind: "command"; key: string; text: string };
