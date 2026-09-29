/**
 * 成本追踪（对应 Claude Code 的 cost-tracker.ts 与 /cost 命令）
 *
 * 价格来自 DeepSeek 官方定价页（美元 / 百万 tokens），区分缓存命中、未命中与输出，并区分高峰时段：
 * 周一至周五 UTC 01:00–04:00 与 06:00–10:00 为高峰（价格为非高峰的 2 倍）。价格可能调整，
 * 可在 settings.json 的 pricing 字段覆盖（填写非高峰价格）。
 */

export interface Price { cacheHit: number; cacheMiss: number; output: number }

export const DEFAULT_PRICING: Record<string, Price> = {
  "deepseek-flash": { cacheHit: 0.003, cacheMiss: 0.15, output: 0.6 },
  "deepseek-v4-pro": { cacheHit: 0.022, cacheMiss: 0.66, output: 1.98 },
};

export function isPeak(d = new Date()): boolean {
  const day = d.getUTCDay();
  if (day === 0 || day === 6) return false;
  const h = d.getUTCHours();
  return (h >= 1 && h < 4) || (h >= 6 && h < 10);
}

export interface CostState {
  requests: number;
  promptTokens: number;
  cachedTokens: number;
  completionTokens: number;
  costUsd: number;
  apiMs: number;
}

export function emptyCost(): CostState {
  return { requests: 0, promptTokens: 0, cachedTokens: 0, completionTokens: 0, costUsd: 0, apiMs: 0 };
}

export function priceFor(model: string, overrides?: Record<string, Price>): Price | undefined {
  return overrides?.[model] ?? DEFAULT_PRICING[model] ?? (model.includes("pro") ? DEFAULT_PRICING["deepseek-v4-pro"] : DEFAULT_PRICING["deepseek-flash"]);
}

export function addUsage(
  state: CostState,
  model: string,
  u: { promptTokens: number; cachedTokens: number; completionTokens: number },
  ms: number,
  overrides?: Record<string, Price>
): number {
  const p = priceFor(model, overrides);
  const mult = isPeak() ? 2 : 1;
  const cost = p
    ? ((u.cachedTokens * p.cacheHit + Math.max(0, u.promptTokens - u.cachedTokens) * p.cacheMiss + u.completionTokens * p.output) / 1e6) * mult
    : 0;
  state.requests++;
  state.promptTokens += u.promptTokens;
  state.cachedTokens += u.cachedTokens;
  state.completionTokens += u.completionTokens;
  state.costUsd += cost;
  state.apiMs += ms;
  return cost;
}

export function formatCost(s: CostState, model: string): string {
  const hitRate = s.promptTokens ? ((s.cachedTokens / s.promptTokens) * 100).toFixed(1) : "0";
  return [
    `模型：${model}`,
    `请求次数：${s.requests}，API 耗时合计：${(s.apiMs / 1000).toFixed(1)} 秒`,
    `输入：${s.promptTokens.toLocaleString()} tokens（其中缓存命中 ${s.cachedTokens.toLocaleString()}，命中率 ${hitRate}%）`,
    `输出：${s.completionTokens.toLocaleString()} tokens`,
    `估算费用：$${s.costUsd.toFixed(4)}（按官方价格估算，${isPeak() ? "当前为高峰时段" : "当前为非高峰时段"}）`,
  ].join("\n");
}
