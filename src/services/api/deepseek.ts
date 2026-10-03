/**
 * DeepSeek API 服务层（对应 Claude Code 的 services/api/claude.ts）
 *
 * - OpenAI 兼容协议，流式输出，边收边解析 tool_calls 分片；
 * - 思考模式：DeepSeek 要求带工具的多轮对话必须把之前每轮 assistant 的 reasoning_content 原样回传，
 *   否则返回 400。这里把 reasoning_content 当作消息的一部分保存和回传；
 * - 错误分类 + 指数退避重试：429/5xx/网络错误重试；401/402/400 立即给出可读的中文原因；
 * - 前缀缓存：DeepSeek 自动按"请求前缀"命中缓存，本层保证 system 在最前且内容稳定。
 */

import OpenAI from "openai";
import type { ResolvedSettings } from "../../config/settings.js";

export interface ApiToolCall {
  id: string;
  type: "function";
  function: { name: string; arguments: string };
}

export type ApiMessage =
  | { role: "system"; content: string }
  | { role: "user"; content: string }
  | { role: "assistant"; content: string | null; reasoning_content?: string; tool_calls?: ApiToolCall[] }
  | { role: "tool"; tool_call_id: string; content: string };

export interface ApiToolSchema {
  type: "function";
  function: { name: string; description: string; parameters: Record<string, unknown> };
}

export interface StreamChunk {
  kind: "text" | "reasoning";
  text: string;
}

export interface CompletionResult {
  content: string;
  reasoning: string;
  toolCalls: ApiToolCall[];
  finishReason: string;
  usage: { promptTokens: number; cachedTokens: number; completionTokens: number };
}

export interface ModelClient {
  complete(
    req: { system: string; messages: ApiMessage[]; tools: ApiToolSchema[]; toolChoice?: "auto" | "none"; maxTokens?: number; /** 覆盖本次请求的思考模式 */ thinking?: boolean },
    opts: { signal: AbortSignal; onChunk?: (c: StreamChunk) => void; onRetry?: (attempt: number, delayMs: number, reason: string) => void }
  ): Promise<CompletionResult>;
  readonly model: string;
  readonly thinking: boolean;
}

export class ApiError extends Error {
  constructor(message: string, public readonly status: number | undefined, public readonly retryable: boolean) {
    super(message);
  }
}

/** 错误分类（对应 categorizeRetryableAPIError） */
export function categorize(e: any): ApiError {
  if (e instanceof ApiError) return e;
  if (e?.name === "AbortError" || e instanceof OpenAI.APIUserAbortError) return new ApiError("请求已被中断", undefined, false);
  const status: number | undefined = e?.status;
  const detail = e?.error?.message || e?.message || String(e);
  if (status === 401) return new ApiError("DeepSeek 认证失败：API Key 无效或已过期，请检查 .env 中的 DEEPSEEK_API_KEY。", status, false);
  if (status === 402) return new ApiError("DeepSeek 账户余额不足，请到 platform.deepseek.com 充值。", status, false);
  if (status === 400) return new ApiError(`DeepSeek 拒绝了请求（400）：${detail}`, status, false);
  if (status === 422) return new ApiError(`请求参数错误（422）：${detail}`, status, false);
  if (status === 429) return new ApiError("请求过于频繁（429），稍后自动重试。", status, true);
  if (status && status >= 500) return new ApiError(`DeepSeek 服务繁忙（${status}），稍后自动重试。`, status, true);
  if (e instanceof OpenAI.APIConnectionError || /ECONNRESET|ETIMEDOUT|ENOTFOUND|fetch failed|socket/i.test(detail)) {
    return new ApiError(`网络连接失败：${detail}`, undefined, true);
  }
  return new ApiError(detail, status, false);
}

const sleep = (ms: number, signal: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    const t = setTimeout(resolve, ms);
    signal.addEventListener("abort", () => {
      clearTimeout(t);
      reject(new ApiError("请求已被中断", undefined, false));
    }, { once: true });
  });

export class DeepSeekClient implements ModelClient {
  private client: OpenAI;
  readonly model: string;
  readonly thinking: boolean;
  private effort?: string;

  constructor(apiKey: string, settings: Pick<ResolvedSettings, "model" | "thinking" | "reasoningEffort">, baseURL = process.env.DEEPSEEK_BASE_URL || "https://api.deepseek.com") {
    this.client = new OpenAI({ apiKey, baseURL, timeout: 10 * 60 * 1000, maxRetries: 0 });
    this.model = settings.model;
    this.thinking = settings.thinking === "enabled";
    this.effort = settings.reasoningEffort;
  }

  async complete(
    req: Parameters<ModelClient["complete"]>[0],
    opts: Parameters<ModelClient["complete"]>[1]
  ): Promise<CompletionResult> {
    const MAX_ATTEMPTS = 5;
    for (let attempt = 1; ; attempt++) {
      let streamedAnything = false;
      try {
        return await this.once(req, opts, () => (streamedAnything = true));
      } catch (raw) {
        const err = categorize(raw);
        // 已经向界面输出过内容时不重试，避免重复文字
        if (!err.retryable || attempt >= MAX_ATTEMPTS || streamedAnything || opts.signal.aborted) throw err;
        const delay = Math.min(30_000, 1000 * 2 ** (attempt - 1)) + Math.floor(Math.random() * 500);
        opts.onRetry?.(attempt, delay, err.message);
        await sleep(delay, opts.signal);
      }
    }
  }

  private async once(
    req: Parameters<ModelClient["complete"]>[0],
    opts: Parameters<ModelClient["complete"]>[1],
    markStreamed: () => void
  ): Promise<CompletionResult> {
    const thinking = req.thinking ?? this.thinking;
    const body: Record<string, unknown> = {
      model: this.model,
      messages: [{ role: "system", content: req.system }, ...req.messages],
      stream: true,
      stream_options: { include_usage: true },
      thinking: { type: thinking ? "enabled" : "disabled" },
    };
    if (req.tools.length) {
      body.tools = req.tools;
      body.tool_choice = req.toolChoice ?? "auto";
    }
    if (thinking && this.effort) body.reasoning_effort = this.effort;
    if (req.maxTokens) body.max_tokens = req.maxTokens;
    if (!thinking) body.temperature = 0.3;

    const stream: any = await this.client.chat.completions.create(body as any, { signal: opts.signal });
    let content = "";
    let reasoning = "";
    let finishReason = "stop";
    const calls = new Map<number, ApiToolCall>();
    let usage = { promptTokens: 0, cachedTokens: 0, completionTokens: 0 };
    for await (const chunk of stream) {
      const choice = chunk.choices?.[0];
      const delta = choice?.delta ?? {};
      if (delta.reasoning_content) {
        reasoning += delta.reasoning_content;
        markStreamed();
        opts.onChunk?.({ kind: "reasoning", text: delta.reasoning_content });
      }
      if (delta.content) {
        content += delta.content;
        markStreamed();
        opts.onChunk?.({ kind: "text", text: delta.content });
      }
      for (const tc of delta.tool_calls ?? []) {
        const cur = calls.get(tc.index) ?? { id: "", type: "function" as const, function: { name: "", arguments: "" } };
        if (tc.id) cur.id = tc.id;
        if (tc.function?.name) cur.function.name += tc.function.name;
        if (tc.function?.arguments) cur.function.arguments += tc.function.arguments;
        calls.set(tc.index, cur);
      }
      if (choice?.finish_reason) finishReason = choice.finish_reason;
      if (chunk.usage) {
        usage = {
          promptTokens: chunk.usage.prompt_tokens ?? 0,
          cachedTokens: chunk.usage.prompt_cache_hit_tokens ?? chunk.usage.prompt_tokens_details?.cached_tokens ?? 0,
          completionTokens: chunk.usage.completion_tokens ?? 0,
        };
      }
    }
    const toolCalls = [...calls.entries()].sort((a, b) => a[0] - b[0]).map(([, v], i) => ({ ...v, id: v.id || `call_${Date.now()}_${i}` }));
    return { content, reasoning, toolCalls, finishReason, usage };
  }
}
