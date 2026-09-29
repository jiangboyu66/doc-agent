/**
 * 插件系统（对应 Claude Code 的 plugins/）
 *
 * 插件 = 一个目录 + plugin.json 清单，可以一次性贡献：Skill、子代理、斜杠命令、Hook。
 * 加载位置：项目 plugins/、项目 .doc-agent/plugins/、用户 ~/.doc-agent/plugins/
 * 在 plugin.json 中设置 "enabled": false 可禁用。
 */

import fs from "node:fs/promises";
import path from "node:path";
import { z } from "zod";
import { PATHS } from "../config/settings.js";
import type { AgentDefinition } from "../agents/builtInAgents.js";

const ManifestSchema = z.object({
  name: z.string(),
  version: z.string().default("0.0.0"),
  description: z.string().default(""),
  enabled: z.boolean().default(true),
  skillsDir: z.string().optional(),
  agents: z.array(z.object({
    type: z.string(),
    description: z.string(),
    whenToUse: z.string().default(""),
    tools: z.array(z.string()).optional(),
    readOnly: z.boolean().default(true),
    prompt: z.string(),
  })).default([]),
  commands: z.array(z.object({ name: z.string(), description: z.string(), prompt: z.string() })).default([]),
  hooks: z.record(z.string(), z.array(z.object({ matcher: z.string().optional(), hooks: z.array(z.object({ type: z.literal("command"), command: z.string(), timeout: z.number().optional() })) }))).default({}),
});

export interface LoadedPlugin {
  name: string;
  version: string;
  description: string;
  dir: string;
  skillsDir?: string;
  agents: AgentDefinition[];
  commands: Array<{ name: string; description: string; prompt: string }>;
  hooks: z.infer<typeof ManifestSchema>["hooks"];
}

export async function loadPlugins(): Promise<{ plugins: LoadedPlugin[]; errors: string[] }> {
  const roots = [path.join(process.cwd(), "plugins"), path.join(PATHS.projectDir, "plugins"), path.join(PATHS.userDir, "plugins")];
  const plugins: LoadedPlugin[] = [];
  const errors: string[] = [];
  for (const root of roots) {
    let dirs: string[];
    try { dirs = await fs.readdir(root); } catch { continue; }
    for (const d of dirs) {
      const dir = path.join(root, d);
      const manifestPath = path.join(dir, "plugin.json");
      let raw: string;
      try { raw = await fs.readFile(manifestPath, "utf8"); } catch { continue; }
      const parsed = ManifestSchema.safeParse(JSON.parse(raw));
      if (!parsed.success) {
        errors.push(`${manifestPath}: ${parsed.error.issues.map((i) => `${i.path.join(".")} ${i.message}`).join("; ")}`);
        continue;
      }
      const m = parsed.data;
      if (!m.enabled || plugins.some((p) => p.name === m.name)) continue;
      plugins.push({
        name: m.name, version: m.version, description: m.description, dir,
        skillsDir: m.skillsDir ? path.join(dir, m.skillsDir) : undefined,
        agents: m.agents.map((a) => ({ agentType: a.type, description: a.description, whenToUse: a.whenToUse, tools: a.tools, readOnly: a.readOnly, prompt: a.prompt, source: `插件:${m.name}` })),
        commands: m.commands,
        hooks: m.hooks,
      });
    }
  }
  return { plugins, errors };
}
