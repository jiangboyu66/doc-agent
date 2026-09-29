/**
 * 分层配置（对应 Claude Code 的 settings.json 体系）
 *
 * 优先级从低到高：内置默认 < 用户级 ~/.doc-agent/settings.json < 项目级 .doc-agent/settings.json
 *               < 环境变量(.env) < 命令行参数
 * 权限规则（allow/deny/ask）会跨层合并，其余字段高优先级覆盖低优先级。
 */

import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { z } from "zod";

export const PermissionModeSchema = z.enum(["default", "acceptEdits", "plan", "bypassPermissions"]);
export type PermissionMode = z.infer<typeof PermissionModeSchema>;

const HookCommandSchema = z.object({ type: z.literal("command"), command: z.string(), timeout: z.number().optional() });
const HookMatcherSchema = z.object({ matcher: z.string().optional(), hooks: z.array(HookCommandSchema) });

export const SettingsSchema = z.object({
  model: z.string().optional(),
  thinking: z.enum(["enabled", "disabled"]).optional(),
  reasoningEffort: z.enum(["low", "high", "max"]).optional(),
  permissionMode: PermissionModeSchema.optional(),
  trackChanges: z.boolean().optional(),
  author: z.string().optional(),
  language: z.string().optional(),
  maxTurns: z.number().int().positive().optional(),
  contextBudgetTokens: z.number().int().positive().optional(),
  autoCompactRatio: z.number().min(0.2).max(0.95).optional(),
  maxBudgetUsd: z.number().positive().optional(),
  permissions: z.object({
    allow: z.array(z.string()).optional(),
    deny: z.array(z.string()).optional(),
    ask: z.array(z.string()).optional(),
  }).optional(),
  hooks: z.record(z.string(), z.array(HookMatcherSchema)).optional(),
  features: z.record(z.string(), z.boolean()).optional(),
  pricing: z.record(z.string(), z.object({ cacheHit: z.number(), cacheMiss: z.number(), output: z.number() })).optional(),
});
export type Settings = z.infer<typeof SettingsSchema>;

export interface ResolvedSettings {
  model: string;
  thinking: "enabled" | "disabled";
  reasoningEffort?: "low" | "high" | "max";
  permissionMode: PermissionMode;
  trackChanges: boolean;
  author: string;
  language: string;
  maxTurns: number;
  contextBudgetTokens: number;
  autoCompactRatio: number;
  maxBudgetUsd?: number;
  permissions: { allow: string[]; deny: string[]; ask: string[] };
  hooks: Record<string, Array<z.infer<typeof HookMatcherSchema>>>;
  features: Record<string, boolean>;
  pricing: Settings["pricing"];
  sources: string[];
}

export const PATHS = {
  userDir: process.env.DOC_AGENT_USER_DIR || path.join(os.homedir(), ".doc-agent"),
  projectDir: path.join(process.cwd(), ".doc-agent"),
  dataDir: process.env.DOC_AGENT_DATA_DIR || path.join(process.cwd(), "data"),
};

async function readJson(file: string): Promise<Settings | null> {
  try {
    const raw = await fs.readFile(file, "utf8");
    const parsed = SettingsSchema.safeParse(JSON.parse(raw));
    if (!parsed.success) {
      console.warn(`[settings] ${file} 格式有误，已忽略：${parsed.error.issues.map((i) => i.path.join(".") + " " + i.message).join("; ")}`);
      return null;
    }
    return parsed.data;
  } catch {
    return null;
  }
}

function fromEnv(): Settings {
  const e = process.env;
  const s: Settings = {};
  if (e.DEEPSEEK_MODEL) s.model = e.DEEPSEEK_MODEL;
  if (e.DEEPSEEK_THINKING === "enabled" || e.DEEPSEEK_THINKING === "disabled") s.thinking = e.DEEPSEEK_THINKING;
  if (e.DEEPSEEK_REASONING_EFFORT === "low" || e.DEEPSEEK_REASONING_EFFORT === "high" || e.DEEPSEEK_REASONING_EFFORT === "max") s.reasoningEffort = e.DEEPSEEK_REASONING_EFFORT;
  if (e.DOC_AGENT_AUTHOR) s.author = e.DOC_AGENT_AUTHOR;
  if (e.DOC_AGENT_PERMISSION_MODE) {
    const m = PermissionModeSchema.safeParse(e.DOC_AGENT_PERMISSION_MODE);
    if (m.success) s.permissionMode = m.data;
  }
  if (e.DOC_AGENT_TRACK_CHANGES) s.trackChanges = e.DOC_AGENT_TRACK_CHANGES === "true" || e.DOC_AGENT_TRACK_CHANGES === "1";
  if (e.DOC_AGENT_MAX_BUDGET_USD) s.maxBudgetUsd = Number(e.DOC_AGENT_MAX_BUDGET_USD);
  return s;
}

export async function loadSettings(cli: Settings = {}): Promise<ResolvedSettings> {
  const userFile = path.join(PATHS.userDir, "settings.json");
  const projectFile = path.join(PATHS.projectDir, "settings.json");
  const [user, project] = await Promise.all([readJson(userFile), readJson(projectFile)]);
  const layers: Array<[string, Settings | null]> = [["用户级", user], ["项目级", project], ["环境变量", fromEnv()], ["命令行", cli]];
  const merged: ResolvedSettings = {
    model: "deepseek-flash",
    thinking: "enabled",
    permissionMode: "default",
    trackChanges: false,
    author: "文案 Agent",
    language: "zh-CN",
    maxTurns: 40,
    contextBudgetTokens: 160_000,
    autoCompactRatio: 0.75,
    permissions: { allow: [], deny: [], ask: [] },
    hooks: {},
    features: {},
    pricing: undefined,
    sources: [],
  };
  for (const [name, s] of layers) {
    if (!s || !Object.keys(s).length) continue;
    merged.sources.push(name);
    const { permissions, hooks, features, pricing, ...rest } = s;
    for (const [k, v] of Object.entries(rest)) if (v !== undefined) (merged as any)[k] = v;
    if (permissions) {
      merged.permissions.allow.push(...(permissions.allow ?? []));
      merged.permissions.deny.push(...(permissions.deny ?? []));
      merged.permissions.ask.push(...(permissions.ask ?? []));
    }
    if (hooks) for (const [ev, list] of Object.entries(hooks)) (merged.hooks[ev] ??= []).push(...list);
    if (features) Object.assign(merged.features, features);
    if (pricing) merged.pricing = { ...(merged.pricing ?? {}), ...pricing };
  }
  return merged;
}

/** 把一条"永久允许"规则写入项目级配置（用户在确认框里选"以后都允许"时） */
export async function persistProjectAllowRule(rule: string): Promise<void> {
  const file = path.join(PATHS.projectDir, "settings.json");
  let cur: any = {};
  try {
    cur = JSON.parse(await fs.readFile(file, "utf8"));
  } catch { /* 新建 */ }
  cur.permissions ??= {};
  cur.permissions.allow ??= [];
  if (!cur.permissions.allow.includes(rule)) cur.permissions.allow.push(rule);
  await fs.mkdir(PATHS.projectDir, { recursive: true });
  await fs.writeFile(file, JSON.stringify(cur, null, 2), "utf8");
}
