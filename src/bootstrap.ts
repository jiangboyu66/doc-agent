/**
 * 启动装配（对应 Claude Code main.tsx 的并行预加载）
 *
 * 配置、插件、长期记忆、外部工具探测互不依赖 → Promise.all 并行加载；
 * Skill 依赖插件目录 → 插件就绪后加载；重型依赖（pdf.js、pandoc/LibreOffice 调用）按需懒加载。
 */

import { loadSettings, type ResolvedSettings, type Settings } from "./config/settings.js";
import { configureFeatures, feature } from "./config/features.js";
import { loadPlugins, type LoadedPlugin } from "./plugins/pluginLoader.js";
import { loadSkills, type SkillRegistry } from "./skills/loadSkills.js";
import { BUILT_IN_AGENTS, type AgentDefinition } from "./agents/builtInAgents.js";
import { loadMemory } from "./context.js";
import { checkExternalTools } from "./services/convert.js";
import { DeepSeekClient, type ModelClient } from "./services/api/deepseek.js";
import { registerInternalHook } from "./hooks/hooks.js";
import { findTool } from "./tools.js";

export interface Runtime {
  settings: ResolvedSettings;
  skills: SkillRegistry;
  agents: AgentDefinition[];
  plugins: LoadedPlugin[];
  pluginErrors: string[];
  memory: string;
  external: { pandoc: string | null; soffice: string | null };
  /** PDF 转 Word 可用的引擎 */
  pdfEngines: { builtin: boolean; pdf2docx: boolean; libreoffice: boolean };
  startupMs: number;
}

let builtinHooksRegistered = false;

function registerBuiltinHooks() {
  if (builtinHooksRegistered) return;
  builtinHooksRegistered = true;
  // 保真守卫：每次编辑后校验文档结构（XML 合法、部件完整、修订 ID 唯一），不通过就阻止提交并自动回滚
  registerInternalHook({
    event: "PostToolUse",
    name: "保真守卫",
    async fn(input, ctx) {
      if (!feature("FIDELITY_GUARD") || !input.toolResult?.ok) return;
      const tool = input.toolName ? findTool(input.toolName) : undefined;
      if (!tool || tool.category !== "edit") return;
      const rep = ctx.session.doc.fidelity(await ctx.session.original());
      if (!rep.ok) return { block: true, reason: `保真守卫检测到结构问题（${rep.problems.join("；")}），这次修改没有生效，请换一种修改方式。` };
    },
  });
}

export async function bootstrap(cli: Settings = {}): Promise<Runtime> {
  const t0 = Date.now();
  const settings = await loadSettings(cli);
  configureFeatures(settings.features);
  const [pluginRes, memory, external, hasPdf2docx] = await Promise.all([
    feature("PLUGINS") ? loadPlugins() : Promise.resolve({ plugins: [] as LoadedPlugin[], errors: [] as string[] }),
    loadMemory(),
    checkExternalTools(),
    import("./services/pdfConvert/index.js").then((m) => m.hasPdf2docx()).catch(() => false),
  ]);
  for (const p of pluginRes.plugins) {
    for (const [ev, list] of Object.entries(p.hooks)) (settings.hooks[ev] ??= []).push(...list);
  }
  const skills = await loadSkills(pluginRes.plugins.filter((p) => p.skillsDir).map((p) => ({ dir: p.skillsDir!, plugin: p.name })));
  const agents = [...BUILT_IN_AGENTS, ...pluginRes.plugins.flatMap((p) => p.agents)];
  registerBuiltinHooks();
  const pdfEngines = { builtin: true, pdf2docx: hasPdf2docx, libreoffice: !!external.soffice };
  return { settings, skills, agents, plugins: pluginRes.plugins, pluginErrors: pluginRes.errors, memory, external, pdfEngines, startupMs: Date.now() - t0 };
}

export function createClient(settings: ResolvedSettings): ModelClient {
  const key = process.env.DEEPSEEK_API_KEY;
  if (!key || key.includes("xxxx")) throw new Error("未配置 DEEPSEEK_API_KEY：请在项目根目录的 .env 文件中填写你的 DeepSeek API Key 后重启。");
  return new DeepSeekClient(key, { model: settings.model, thinking: feature("THINKING") ? settings.thinking : "disabled", reasoningEffort: settings.reasoningEffort });
}
