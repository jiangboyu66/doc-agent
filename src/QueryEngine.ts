/**
 * QueryEngine —— 代理主循环（对应 Claude Code 的 src/QueryEngine.ts + query.ts）
 *
 * submitMessage() 是一个异步生成器：调用方 for await 即可拿到流式事件（文字增量、工具调用、
 * 权限请求、文档变更、用量……）。一轮对话的流程：
 *
 *   UserPromptSubmit Hook → 附加 <system-reminder> → [循环] 压缩检查 → 调用模型（流式）
 *   → 没有工具调用：Stop Hook → 结束
 *   → 有工具调用：按"并发安全"分批（只读工具并行、写工具串行）→ 每个调用：
 *        参数校验(zod) → validateInput → PreToolUse Hook → 权限决策（必要时干跑出预览并等待用户确认）
 *        → 执行 → PostToolUse Hook（含保真守卫，失败自动回滚）→ 提交版本快照
 *   → 工具结果回填，继续循环，直到模型给出最终答复 / 达到最大轮次 / 预算用尽 / 用户中断
 *
 * 权限确认期间主循环保持挂起（await 用户的决定），而不是把请求拆成"暂停/恢复"两段——
 * 消息历史从头到尾都是完整、合法的序列，从根源上避免了 tool_calls 配对错误。
 */

import { nanoid } from "nanoid";
import type { Session } from "./session/Session.js";
import type { ResolvedSettings } from "./config/settings.js";
import { persistProjectAllowRule } from "./config/settings.js";
import type { ApiMessage, ApiToolCall, ModelClient } from "./services/api/deepseek.js";
import { normalizeMessagesForAPI, estimateTokens } from "./services/api/normalize.js";
import { addUsage } from "./services/costTracker.js";
import { microCompact, autoCompact } from "./services/compact.js";
import type { EngineEvent, PermissionRequest, PermissionResponse } from "./bridge/protocol.js";
import { toolSchema, type Tool, type ToolUseContext } from "./Tool.js";
import { assembleToolPool } from "./tools.js";
import { hasPermissionsToUseTool, ruleFor } from "./permissions/permissions.js";
import { runHooks } from "./hooks/hooks.js";
import { feature } from "./config/features.js";
import type { SkillRegistry } from "./skills/loadSkills.js";
import type { AgentDefinition } from "./agents/builtInAgents.js";
import { getStaticSystemPrompt, getSessionSystemPrompt } from "./constants/prompts.js";
import { buildEffectiveSystemPrompt } from "./utils/systemPrompt.js";
import { buildTurnReminder } from "./context.js";
import { AsyncQueue } from "./utils/asyncQueue.js";
import { DocError } from "./documents/types.js";
import { revisionDate } from "./documents/docx/ooxml.js";

export interface EngineDeps {
  client: ModelClient;
  settings: ResolvedSettings;
  skills: SkillRegistry;
  agents: AgentDefinition[];
  memory: string;
  customSystemPrompt?: string;
  appendSystemPrompt?: string;
  /** 前端提供：展示权限请求并返回用户的决定 */
  canUseTool(req: PermissionRequest, signal: AbortSignal): Promise<PermissionResponse>;
}

interface EngineOptions {
  agentId?: string;
  agentDef?: AgentDefinition;
  initialMessages?: ApiMessage[];
  maxTurns?: number;
}

type Emit = (e: EngineEvent) => void;

export class QueryEngine {
  readonly agentId: string;
  private messages: ApiMessage[];
  private turnSnapshot: ApiMessage[] = [];

  constructor(private session: Session, private deps: EngineDeps, private opts: EngineOptions = {}) {
    this.agentId = opts.agentId ?? "main";
    // 主代理直接使用会话历史（同一个数组引用，保证持久化的就是真实发送的序列）
    this.messages = this.agentId === "main" ? session.history : [...(opts.initialMessages ?? [])];
  }

  get isMain() {
    return this.agentId === "main";
  }

  async *submitMessage(prompt: string, signal: AbortSignal): AsyncGenerator<EngineEvent> {
    const q = new AsyncQueue<EngineEvent>();
    const run = this.query(prompt, signal, (e) => q.push(e))
      .catch((err) => {
        q.push({ type: "error", message: err?.message ?? String(err) });
        q.push({ type: "done", reason: signal.aborted ? "interrupted" : "error" });
      })
      .finally(() => q.close());
    for await (const e of q) yield e;
    await run;
  }

  // -------------------------------------------------------------------------

  private makeContext(signal: AbortSignal, emit: Emit): ToolUseContext {
    const s = this.session;
    return {
      session: s,
      settings: this.deps.settings,
      agentId: this.agentId,
      abortSignal: signal,
      emit,
      editOptions: () => ({ track: s.meta.trackChanges && s.doc.capabilities.trackChanges, author: s.meta.author, date: revisionDate() }),
      skills: this.deps.skills,
      agents: this.deps.agents,
      runAgent: this.isMain ? (p) => this.runAgent(p, signal, emit) : undefined,
      parentMessages: () => this.turnSnapshot,
    };
  }

  private systemPrompt(): string {
    return buildEffectiveSystemPrompt({
      defaultPrompt: getStaticSystemPrompt(),
      sessionPrompt: getSessionSystemPrompt(this.session, this.deps.skills, this.deps.memory),
      agentPrompt: this.opts.agentDef?.prompt,
      customPrompt: this.deps.customSystemPrompt,
      appendPrompt: this.deps.appendSystemPrompt,
    });
  }

  /** 思考模式：会话级开关（界面上的"深度思考"按钮）优先，其次全局配置；特性开关关闭时一律不开 */
  private thinkingEnabled(): boolean {
    if (!feature("THINKING")) return false;
    return this.session.meta.thinking ?? this.deps.client.thinking;
  }

  private async persist() {
    if (this.isMain) await this.session.save();
  }

  private async query(prompt: string, signal: AbortSignal, emit: Emit): Promise<void> {
    const { client, settings } = this.deps;
    const ctx = this.makeContext(signal, emit);
    emit({ type: "turn_start", turnId: nanoid(8) });

    const hook = await runHooks("UserPromptSubmit", { sessionId: this.session.id, documentFormat: this.session.meta.format, prompt }, ctx, feature("HOOKS"));
    if (hook.block) {
      emit({ type: "error", message: `提交被 Hook 阻止：${hook.reason}` });
      emit({ type: "done", reason: "error" });
      return;
    }
    this.turnSnapshot = [...this.messages];
    this.messages.push({ role: "user", content: `${buildTurnReminder(this.session, hook.additionalContext)}\n\n${prompt}` });
    await this.persist();

    const maxTurns = this.opts.maxTurns ?? settings.maxTurns;
    for (let turn = 0; turn < maxTurns; turn++) {
      if (signal.aborted) {
        emit({ type: "done", reason: "interrupted" });
        return;
      }
      if (settings.maxBudgetUsd && this.session.meta.cost.costUsd >= settings.maxBudgetUsd) {
        emit({ type: "error", message: `已达到预算上限 $${settings.maxBudgetUsd}，停止执行。可在设置中调整 maxBudgetUsd。` });
        emit({ type: "done", reason: "budget" });
        return;
      }

      const system = this.systemPrompt();
      await this.maybeCompact(system, prompt, signal, emit, ctx);
      const tools = assembleToolPool(ctx, this.opts.agentDef?.tools);
      const thinking = this.thinkingEnabled();
      const t0 = Date.now();
      let res;
      try {
        res = await client.complete(
          { system, messages: normalizeMessagesForAPI(this.messages, thinking), tools: tools.map((t) => toolSchema(t, ctx)), thinking },
          {
            signal,
            onChunk: (c) => emit(c.kind === "text" ? { type: "text_delta", text: c.text, agentId: this.agentId } : { type: "reasoning_delta", text: c.text, agentId: this.agentId }),
            onRetry: (attempt, delayMs, reason) => emit({ type: "retry", attempt, delayMs, reason }),
          }
        );
      } catch (e: any) {
        if (signal.aborted) {
          emit({ type: "done", reason: "interrupted" });
          return;
        }
        throw e;
      }
      const cost = addUsage(this.session.meta.cost, client.model, res.usage, Date.now() - t0, settings.pricing);
      emit({ type: "usage", ...res.usage, costUsd: cost, totalCostUsd: this.session.meta.cost.costUsd });

      this.messages.push({
        role: "assistant",
        content: res.content || null,
        reasoning_content: res.reasoning,
        tool_calls: res.toolCalls.length ? res.toolCalls : undefined,
      });
      await this.persist();
      if (res.content.trim()) emit({ type: "assistant_message", text: res.content, agentId: this.agentId });

      if (!res.toolCalls.length) {
        const stop = await runHooks("Stop", { sessionId: this.session.id, documentFormat: this.session.meta.format }, ctx, feature("HOOKS"));
        if (stop.block && turn < maxTurns - 1) {
          this.messages.push({ role: "user", content: `<system-reminder>Stop Hook 要求继续：${stop.reason}</system-reminder>` });
          continue;
        }
        emit({ type: "done", reason: "completed" });
        return;
      }

      const results = await this.runTools(res.toolCalls, tools, ctx, emit);
      for (const r of results) this.messages.push({ role: "tool", tool_call_id: r.id, content: r.content });
      await this.persist();
    }
    emit({ type: "error", message: `已达到单轮最大步数（${maxTurns}），为避免失控已停止。可以回复"继续"让我接着完成。` });
    emit({ type: "done", reason: "max_turns" });
  }

  // -------------------------------------------------------------------------
  // 工具执行：并发安全的相邻调用并行，其余串行
  // -------------------------------------------------------------------------

  private async runTools(calls: ApiToolCall[], pool: Tool[], ctx: ToolUseContext, emit: Emit) {
    const parsed = calls.map((c) => {
      const tool = pool.find((t) => t.name === c.function.name);
      let input: unknown = undefined;
      let parseError: string | undefined;
      try {
        input = JSON.parse(c.function.arguments || "{}");
      } catch {
        parseError = "参数不是合法的 JSON，请检查引号与转义后重试。";
      }
      let safe = false;
      if (tool && !parseError) {
        const v = tool.inputSchema.safeParse(input);
        if (v.success) {
          input = v.data;
          safe = tool.isConcurrencySafe(v.data);
        }
      }
      return { call: c, tool, input, parseError, safe };
    });

    const batches: Array<typeof parsed> = [];
    for (const p of parsed) {
      const last = batches[batches.length - 1];
      if (p.safe && last && last[0].safe) last.push(p);
      else batches.push([p]);
    }

    const results: Array<{ id: string; content: string }> = [];
    for (const batch of batches) {
      if (ctx.abortSignal.aborted) {
        for (const p of batch) results.push({ id: p.call.id, content: "（用户中断，未执行）" });
        continue;
      }
      const out = await Promise.all(batch.map((p) => this.runToolUse(p, ctx, emit)));
      batch.forEach((p, i) => results.push({ id: p.call.id, content: out[i] }));
    }
    return results;
  }

  private async runToolUse(
    p: { call: ApiToolCall; tool?: Tool; input: unknown; parseError?: string },
    ctx: ToolUseContext,
    emit: Emit
  ): Promise<string> {
    const { call, tool } = p;
    const name = call.function.name;
    const fail = (content: string) => {
      emit({ type: "tool_result", id: call.id, name, ok: false, content, agentId: this.agentId });
      return `错误：${content}`;
    };
    if (!tool) return fail(`工具 ${name} 不存在或在当前文档/模式下不可用。`);
    if (p.parseError) return fail(p.parseError);
    const v = tool.inputSchema.safeParse(p.input);
    if (!v.success) {
      return fail(`参数不符合要求：${v.error.issues.map((i) => `${i.path.join(".") || "(根)"} ${i.message}`).join("；")}`);
    }
    const input = v.data;
    emit({ type: "tool_use", id: call.id, name, title: tool.userFacingName(input), input, agentId: this.agentId });

    const valid = await tool.validateInput(input, ctx);
    if (!valid.ok) return fail(valid.message);

    const pre = await runHooks("PreToolUse", { sessionId: this.session.id, documentFormat: this.session.meta.format, toolName: name, toolInput: input }, ctx, feature("HOOKS"));
    if (pre.block) return fail(`被 Hook 阻止：${pre.reason}`);

    const decision = await hasPermissionsToUseTool(tool, input, ctx, this.session.meta.mode);
    if (decision.behavior === "deny") return fail(decision.message);
    if (decision.behavior === "ask") {
      let preview: PermissionRequest["preview"];
      if (feature("DRY_RUN_PREVIEW") && tool.preview) {
        try {
          preview = await tool.preview(input, ctx);
        } catch (e: any) {
          // 干跑就失败（找不到原文、引用过期……）：直接把错误交给模型修正，不打扰用户
          return fail(e instanceof DocError ? e.message : `预览失败：${e.message}`);
        }
      }
      const request: PermissionRequest = {
        requestId: nanoid(10), toolUseId: call.id, tool: name, title: tool.userFacingName(input), input,
        reason: decision.reason, preview, suggestedRule: ruleFor(tool, input), agentId: this.agentId,
      };
      emit({ type: "permission_request", request });
      let resp: PermissionResponse;
      try {
        resp = await this.deps.canUseTool(request, ctx.abortSignal);
      } catch {
        resp = { decision: "deny", feedback: "确认过程被中断" };
      }
      emit({ type: "permission_resolved", requestId: request.requestId, decision: resp.decision });
      if (resp.decision === "deny") {
        return fail(`用户拒绝了这个操作。${resp.feedback ? `用户的说明：${resp.feedback}` : "请停下来询问用户希望如何处理，不要换一种方式重复同样的修改。"}`);
      }
      if (resp.remember === "session" && !this.session.meta.sessionRules.includes(request.suggestedRule)) {
        this.session.meta.sessionRules.push(request.suggestedRule);
      }
      if (resp.remember === "project") {
        await persistProjectAllowRule(request.suggestedRule);
        this.deps.settings.permissions.allow.push(request.suggestedRule);
      }
    }

    let result;
    try {
      result = await tool.call(input, ctx);
    } catch (e: any) {
      if (tool.category === "edit") await this.session.revertToCurrent();
      return fail(e instanceof DocError ? e.message : `工具执行出错：${e?.message ?? e}`);
    }

    const post = await runHooks(
      "PostToolUse",
      { sessionId: this.session.id, documentFormat: this.session.meta.format, toolName: name, toolInput: input, toolResult: { ok: !result.isError, content: result.content } },
      ctx,
      feature("HOOKS")
    );
    if (result.docChange) {
      if (post.block) {
        await this.session.revertToCurrent();
        return fail(`修改已自动撤销：${post.reason}`);
      }
      const version = await this.session.commit(result.docChange.label, result.docChange.changedRefs);
      emit({ type: "doc_changed", version, label: result.docChange.label, changedRefs: result.docChange.changedRefs, structural: result.docChange.structural });
    }

    let content = result.content;
    if (post.additionalContext) content += `\n\n<system-reminder>${post.additionalContext}</system-reminder>`;
    if (content.length > tool.maxResultChars) {
      content = content.slice(0, tool.maxResultChars) + `\n…（结果过长已截断，共 ${content.length} 字，请缩小范围后重试）`;
    }
    emit({ type: "tool_result", id: call.id, name, ok: !result.isError, content: result.content, preview: result.preview, data: result.data, agentId: this.agentId });
    return content;
  }

  // -------------------------------------------------------------------------
  // 子代理
  // -------------------------------------------------------------------------

  private async runAgent(
    p: { agentType: string; description: string; prompt: string; parentMessages?: ApiMessage[] },
    signal: AbortSignal,
    emit: Emit
  ): Promise<string> {
    const def: AgentDefinition | undefined = p.agentType === "fork"
      ? { agentType: "fork", description: "继承上下文的分叉代理", whenToUse: "", readOnly: false, prompt: undefined as unknown as string, source: "内置" }
      : this.deps.agents.find((a) => a.agentType === p.agentType);
    if (!def) throw new Error(`未知子代理 ${p.agentType}`);
    const agentId = `${p.agentType}-${nanoid(5)}`;
    emit({ type: "subagent", agentId, agentType: p.agentType, status: "start", description: p.description });
    const sub = new QueryEngine(this.session, this.deps, {
      agentId,
      agentDef: p.agentType === "fork" ? undefined : def,
      initialMessages: p.parentMessages ?? [],
      maxTurns: 30,
    });
    let final = "";
    for await (const e of sub.submitMessage(p.prompt, signal)) {
      if (e.type === "assistant_message") final = e.text;
      if (e.type === "text_delta" || e.type === "reasoning_delta" || e.type === "turn_start" || e.type === "done") continue;
      emit(e);
    }
    emit({ type: "subagent", agentId, agentType: p.agentType, status: "done", description: p.description });
    return final.trim() || "（子代理没有给出结论）";
  }

  // -------------------------------------------------------------------------
  // 上下文压缩
  // -------------------------------------------------------------------------

  private async maybeCompact(system: string, prompt: string, signal: AbortSignal, emit: Emit, ctx: ToolUseContext) {
    const budget = this.deps.settings.contextBudgetTokens;
    let est = estimateTokens(this.messages, system);
    if (feature("MICRO_COMPACT") && est > budget * 0.5) {
      const saved = microCompact(this.messages);
      if (saved > 0) {
        const after = estimateTokens(this.messages, system);
        emit({ type: "compact", kind: "micro", before: est, after });
        est = after;
      }
    }
    if (feature("AUTO_COMPACT") && est > budget * this.deps.settings.autoCompactRatio) {
      await runHooks("PreCompact", { sessionId: this.session.id, documentFormat: this.session.meta.format }, ctx, feature("HOOKS"));
      const r = await autoCompact(this.deps.client, this.messages, signal, `用户当前这一轮的请求原文是："${prompt.slice(0, 2000)}"，摘要中必须完整保留它以及它的完成进度。`, true);
      this.messages.splice(0, this.messages.length, ...r.history);
      const after = estimateTokens(this.messages, system);
      emit({ type: "compact", kind: "auto", before: est, after });
      await this.persist();
    }
  }

  /** /compact 命令手动压缩 */
  async compactNow(signal: AbortSignal, instructions?: string): Promise<{ before: number; after: number }> {
    const system = this.systemPrompt();
    const before = estimateTokens(this.messages, system);
    if (this.messages.length < 2) return { before, after: before };
    const r = await autoCompact(this.deps.client, this.messages, signal, instructions);
    this.messages.splice(0, this.messages.length, ...r.history);
    await this.persist();
    return { before, after: estimateTokens(this.messages, system) };
  }
}
