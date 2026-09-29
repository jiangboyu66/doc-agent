/**
 * 特性开关（对应 Claude Code 用 bun:bundle 的 feature() 做条件编译）
 *
 * Node 没有编译期裁剪，这里用运行时开关实现同样的"一码多用"：同一份代码，通过配置打开/关闭子代理、
 * 自动压缩、Hook、插件等能力，方便灰度与排查问题。
 * 来源：settings.features < 环境变量 DOC_AGENT_FEATURES=a,b / DOC_AGENT_DISABLE=c,d
 */

export const FEATURE_DEFAULTS = {
  SUBAGENTS: true,
  SKILLS: true,
  PLUGINS: true,
  HOOKS: true,
  AUTO_COMPACT: true,
  MICRO_COMPACT: true,
  THINKING: true,
  COMMENTS: true,
  FIDELITY_GUARD: true,
  DRY_RUN_PREVIEW: true,
} as const;

export type FeatureName = keyof typeof FEATURE_DEFAULTS;

let overrides: Partial<Record<string, boolean>> = {};

export function configureFeatures(fromSettings: Record<string, boolean>): void {
  overrides = { ...fromSettings };
  for (const f of (process.env.DOC_AGENT_FEATURES ?? "").split(",").map((s) => s.trim()).filter(Boolean)) overrides[f] = true;
  for (const f of (process.env.DOC_AGENT_DISABLE ?? "").split(",").map((s) => s.trim()).filter(Boolean)) overrides[f] = false;
}

export function feature(name: FeatureName): boolean {
  return overrides[name] ?? FEATURE_DEFAULTS[name];
}

export function featureTable(): Record<string, boolean> {
  return Object.fromEntries(Object.keys(FEATURE_DEFAULTS).map((k) => [k, feature(k as FeatureName)]));
}
