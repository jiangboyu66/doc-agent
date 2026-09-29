/**
 * 工具接口与 buildTool（对应 Claude Code 的 src/Tool.ts）
 *
 * 每个工具都是一块"乐高积木"：自带面向模型的说明书（description）、参数 schema（zod）、
 * 只读/并发安全/破坏性等元信息、输入校验、工具自身的权限判断、干跑预览和执行逻辑。
 * buildTool 为缺省字段提供"失败即安全"（fail-closed）的默认值：默认不是只读、默认不能并发、
 * 默认需要走权限检查——新工具忘了声明，也只会更保守而不会更危险。
 */

import { z } from "zod";
import type { Session } from "./session/Session.js";
import type { ResolvedSettings } from "./config/settings.js";
import type { EngineEvent, PermissionRequest } from "./bridge/protocol.js";
import type { EditOptions } from "./documents/types.js";
import type { SkillRegistry } from "./skills/loadSkills.js";
import type { AgentDefinition } from "./agents/builtInAgents.js";
import type { ApiMessage } from "./services/api/deepseek.js";

export type ToolCategory = "read" | "edit" | "export" | "meta" | "agent";

export interface ToolUseContext {
  session: Session;
  settings: ResolvedSettings;
  agentId: string;
  abortSignal: AbortSignal;
  emit(e: EngineEvent): void;
  editOptions(): EditOptions;
  skills: SkillRegistry;
  agents: AgentDefinition[];
  /** 依赖注入：由 QueryEngine 提供启动子代理的能力，AgentTool 不直接依赖引擎实现 */
  runAgent?(p: { agentType: string; description: string; prompt: string; parentMessages?: ApiMessage[] }): Promise<string>;
  /** 当前轮开始时的对话历史快照（fork 子代理继承上下文用） */
  parentMessages(): ApiMessage[];
}

export interface ToolResult {
  content: string;
  isError?: boolean;
  preview?: PermissionRequest["preview"];
  data?: unknown;
  /** 写工具修改了文档时填写，由引擎统一提交版本并通知前端 */
  docChange?: { label: string; changedRefs: string[]; structural: boolean };
}

export type PermissionResult =
  | { behavior: "allow" }
  | { behavior: "deny"; message: string }
  | { behavior: "ask"; reason: string }
  | { behavior: "passthrough" };

export type ValidationResult = { ok: true } | { ok: false; message: string };

export interface Tool<I = any> {
  name: string;
  category: ToolCategory;
  description(ctx: ToolUseContext): string;
  inputSchema: z.ZodType<I>;
  userFacingName(input: Partial<I>): string;
  isEnabled(ctx: ToolUseContext): boolean;
  isReadOnly(input: I): boolean;
  isConcurrencySafe(input: I): boolean;
  isDestructive(input: I): boolean;
  /** 权限规则匹配用的内容，例如 doc_export(pdf) 中的 "pdf" */
  permissionContent(input: I): string;
  validateInput(input: I, ctx: ToolUseContext): Promise<ValidationResult>;
  checkPermissions(input: I, ctx: ToolUseContext): Promise<PermissionResult>;
  preview?(input: I, ctx: ToolUseContext): Promise<ToolResult["preview"]>;
  call(input: I, ctx: ToolUseContext): Promise<ToolResult>;
  maxResultChars: number;
}

type Defaults = "isEnabled" | "isReadOnly" | "isConcurrencySafe" | "isDestructive" | "permissionContent" | "validateInput" | "checkPermissions" | "maxResultChars" | "userFacingName";
export type ToolDef<I> = Omit<Tool<I>, Defaults> & Partial<Pick<Tool<I>, Defaults>>;

export function buildTool<I>(def: ToolDef<I>): Tool<I> {
  // 显式传入的 undefined 不能覆盖默认值（否则 { checkPermissions: undefined } 会把安全默认值抹掉）
  const given = Object.fromEntries(Object.entries(def).filter(([, v]) => v !== undefined)) as ToolDef<I>;
  return {
    isEnabled: () => true,
    isReadOnly: () => false,
    isConcurrencySafe: () => false,
    isDestructive: () => false,
    permissionContent: () => "",
    validateInput: async () => ({ ok: true }),
    checkPermissions: async () => ({ behavior: "passthrough" }),
    maxResultChars: 30_000,
    userFacingName: () => def.name,
    ...given,
  };
}

/** 把 zod schema 转成 DeepSeek function calling 需要的 JSON Schema */
export function toolSchema(t: Tool, ctx: ToolUseContext) {
  const schema = z.toJSONSchema(t.inputSchema, { target: "draft-7" }) as Record<string, unknown>;
  delete schema.$schema;
  return { type: "function" as const, function: { name: t.name, description: t.description(ctx), parameters: schema } };
}
