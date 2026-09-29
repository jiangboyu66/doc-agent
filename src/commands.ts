/**
 * 斜杠命令注册表（对应 Claude Code 的 src/commands.ts）
 *
 * 两类命令：
 *   - local：本地直接执行，不调用模型（/cost /undo /mode /export …），即时、零成本；
 *   - prompt：展开成一段提示词交给模型执行（/review，以及每个 Skill、插件命令都自动成为 /命令）。
 * Web 与 CLI 共用同一个注册表和同一个 processUserInput 入口。
 */

import type { Session } from "./session/Session.js";
import type { Runtime } from "./bootstrap.js";
import { formatCost } from "./services/costTracker.js";
import { exportDocument, type ExportFormat } from "./services/convert.js";
import { featureTable } from "./config/features.js";
import { appendProjectMemory } from "./context.js";
import { renderSkill } from "./skills/loadSkills.js";
import { PermissionModeSchema } from "./config/settings.js";
import { estimateTokens } from "./services/api/normalize.js";

export interface CommandContext {
  session: Session;
  runtime: Runtime;
  compact(instructions?: string): Promise<{ before: number; after: number }>;
}

export type Command =
  | { type: "local"; name: string; description: string; argumentHint?: string; run(args: string, ctx: CommandContext): Promise<string> }
  | { type: "prompt"; name: string; description: string; argumentHint?: string; source: string; getPrompt(args: string, ctx: CommandContext): string };

const BUILTIN: Command[] = [
  {
    type: "local", name: "help", description: "查看所有命令与技能",
    async run(_a, ctx) {
      const cmds = getCommands(ctx.runtime);
      return ["可用命令：", ...cmds.map((c) => `  /${c.name}${c.argumentHint ? ` ${c.argumentHint}` : ""} — ${c.description}${c.type === "prompt" ? `（${c.source}）` : ""}`)].join("\n");
    },
  },
  {
    type: "local", name: "mode", description: "切换权限模式", argumentHint: "default|acceptEdits|plan|bypassPermissions",
    async run(a, ctx) {
      if (!a.trim()) return `当前权限模式：${ctx.session.meta.mode}`;
      const m = PermissionModeSchema.safeParse(a.trim());
      if (!m.success) return "可选：default（逐一确认）/ acceptEdits（自动接受编辑）/ plan（只规划不修改）/ bypassPermissions（全部免确认）";
      ctx.session.meta.mode = m.data;
      await ctx.session.save();
      return `权限模式已切换为 ${m.data}`;
    },
  },
  {
    type: "local", name: "track", description: "开关 Word 修订模式（修改记录为可接受/拒绝的修订）", argumentHint: "on|off",
    async run(a, ctx) {
      if (!ctx.session.doc.capabilities.trackChanges) return "当前文档格式不支持修订模式（仅 Word 支持）。";
      const v = a.trim().toLowerCase();
      if (v !== "on" && v !== "off") return `修订模式当前为：${ctx.session.meta.trackChanges ? "开启" : "关闭"}`;
      ctx.session.meta.trackChanges = v === "on";
      await ctx.session.save();
      return `修订模式已${v === "on" ? "开启" : "关闭"}`;
    },
  },
  {
    type: "local", name: "author", description: "设置修订/批注的作者名", argumentHint: "名字",
    async run(a, ctx) {
      if (!a.trim()) return `当前作者：${ctx.session.meta.author}`;
      ctx.session.meta.author = a.trim();
      await ctx.session.save();
      return `作者已设为 ${a.trim()}`;
    },
  },
  {
    type: "local", name: "undo", description: "撤销上一次修改",
    async run(_a, ctx) {
      const vs = ctx.session.meta.versions;
      const idx = vs.findIndex((v) => v.v === ctx.session.meta.currentVersion);
      if (idx <= 0) return "已经是原始版本，没有可撤销的修改。";
      const target = vs[idx - 1];
      const nv = await ctx.session.rollback(target.v);
      return `已撤销"${vs[idx].label}"，文档回到 v${target.v} 的状态（记录为 v${nv}）。`;
    },
  },
  {
    type: "local", name: "history", description: "查看版本历史",
    async run(_a, ctx) {
      return ctx.session.meta.versions.map((v) => `${v.v === ctx.session.meta.currentVersion ? "▶" : " "} v${v.v}  ${new Date(v.at).toLocaleString("zh-CN", { hour12: false })}  ${v.label}`).join("\n");
    },
  },
  {
    type: "local", name: "rewind", description: "回滚到指定版本", argumentHint: "版本号",
    async run(a, ctx) {
      const v = Number(a.trim().replace(/^v/i, ""));
      if (!ctx.session.meta.versions.some((x) => x.v === v)) return "请提供有效版本号，可用 /history 查看。";
      const nv = await ctx.session.rollback(v);
      return `已回滚到 v${v}（记录为 v${nv}）。`;
    },
  },
  {
    type: "local", name: "verify", description: "保真校验：对比原始文件，报告改动范围",
    async run(_a, ctx) {
      const r = ctx.session.doc.fidelity(await ctx.session.original());
      return [
        r.ok ? "✓ 结构校验通过" : "✗ 发现问题",
        `部件 ${r.totalParts} 个，其中 ${r.identicalParts} 个与原文件逐字节一致`,
        ...r.modifiedParts.map((m) => `- 已修改 ${m.name}${m.changedBlocks.length ? `：${m.changedBlocks.slice(0, 20).join("、")}` : ""}`),
        ...r.addedParts.map((a) => `- 新增 ${a}`),
        ...r.problems.map((p) => `! ${p}`),
      ].join("\n");
    },
  },
  {
    type: "local", name: "export", description: "导出文件", argumentHint: "docx|pdf|html|markdown",
    async run(a, ctx) {
      const fmt = (a.trim() || ctx.session.meta.format) as ExportFormat;
      if (!["docx", "pdf", "html", "markdown"].includes(fmt)) return "格式可选：docx / pdf / html / markdown";
      const out = await exportDocument(ctx.session.current(), ctx.session.meta.format, fmt);
      const name = `${ctx.session.meta.filename.replace(/\.[^.]+$/, "")}-v${ctx.session.meta.currentVersion}.${out.ext}`;
      const file = await ctx.session.writeExport(name, out.data);
      return `已导出：${file}\n${out.note}`;
    },
  },
  {
    type: "local", name: "cost", description: "查看本会话的 token 用量与费用",
    async run(_a, ctx) {
      return formatCost(ctx.session.meta.cost, ctx.runtime.settings.model);
    },
  },
  {
    type: "local", name: "context", description: "查看当前上下文占用",
    async run(_a, ctx) {
      const est = estimateTokens(ctx.session.history);
      const budget = ctx.runtime.settings.contextBudgetTokens;
      return `对话历史约 ${est.toLocaleString()} tokens / 预算 ${budget.toLocaleString()}（${((est / budget) * 100).toFixed(1)}%），超过 ${(ctx.runtime.settings.autoCompactRatio * 100).toFixed(0)}% 时自动压缩。消息数：${ctx.session.history.length}`;
    },
  },
  {
    type: "local", name: "compact", description: "压缩对话上下文", argumentHint: "[摘要时重点保留的内容]",
    async run(a, ctx) {
      const r = await ctx.compact(a.trim() || undefined);
      return `上下文已压缩：约 ${r.before.toLocaleString()} → ${r.after.toLocaleString()} tokens`;
    },
  },
  {
    type: "local", name: "clear", description: "清空对话（文档与版本历史保留）",
    async run(_a, ctx) {
      ctx.session.history.length = 0;
      ctx.session.meta.todos = [];
      await ctx.session.save();
      return "对话已清空，文档内容和版本历史保持不变。";
    },
  },
  {
    type: "local", name: "todos", description: "查看待办清单",
    async run(_a, ctx) {
      const t = ctx.session.meta.todos;
      return t.length ? t.map((x) => `${x.status === "completed" ? "✓" : x.status === "in_progress" ? "▶" : "○"} ${x.content}`).join("\n") : "暂无待办";
    },
  },
  {
    type: "local", name: "skills", description: "列出可用技能",
    async run(_a, ctx) {
      const list = ctx.runtime.skills.list();
      return list.length ? list.map((s) => `/${s.name}${s.argumentHint ? ` ${s.argumentHint}` : ""} — ${s.description}（${s.source}）`).join("\n") : "暂无技能";
    },
  },
  {
    type: "local", name: "agents", description: "列出可用子代理",
    async run(_a, ctx) {
      return ctx.runtime.agents.map((a) => `${a.agentType}${a.readOnly ? "（只读）" : ""} — ${a.description}（${a.source}）`).join("\n");
    },
  },
  {
    type: "local", name: "memory", description: "查看长期偏好（DOCAGENT.md）",
    async run(_a, ctx) {
      return ctx.runtime.memory.trim() || "暂无长期偏好。可用 /remember 添加，例如 /remember 论文一律使用美式拼写";
    },
  },
  {
    type: "local", name: "remember", description: "记住一条长期偏好（写入项目 DOCAGENT.md）", argumentHint: "偏好内容",
    async run(a, ctx) {
      if (!a.trim()) return "用法：/remember 内容";
      const file = await appendProjectMemory(a);
      ctx.runtime.memory = `${ctx.runtime.memory}\n- ${a.trim()}`.trim();
      return `已记住，写入 ${file}（新会话生效；当前会话从下一轮起生效）。`;
    },
  },
  {
    type: "local", name: "doctor", description: "诊断运行环境",
    async run(_a, ctx) {
      const r = ctx.runtime;
      return [
        `DeepSeek API Key：${process.env.DEEPSEEK_API_KEY ? "已配置" : "未配置"}；模型：${r.settings.model}；思考模式：${r.settings.thinking}`,
        `pandoc：${r.external.pandoc ?? "未找到（导出 HTML/Markdown 需要）"}`,
        `LibreOffice：${r.external.soffice ?? "未找到（导出 PDF / 版式预览需要）"}`,
        `配置来源：${r.settings.sources.join(" < ") || "仅默认值"}`,
        `技能 ${r.skills.list().length} 个；子代理 ${r.agents.length} 个；插件 ${r.plugins.length} 个${r.pluginErrors.length ? `（加载失败：${r.pluginErrors.join("; ")}）` : ""}`,
        `特性开关：${Object.entries(featureTable()).map(([k, v]) => `${k}=${v ? "开" : "关"}`).join(" ")}`,
        `启动耗时：${r.startupMs} ms`,
      ].join("\n");
    },
  },
  {
    type: "prompt", name: "review", description: "系统性审阅文档（并行子代理），只给意见不直接修改", argumentHint: "[范围或关注点]", source: "内置",
    getPrompt: (a) => `请系统性审阅这份文档${a ? `，重点：${a}` : ""}。
步骤：1) 用 doc_outline 了解结构；2) 按章节划分范围，在同一轮中启动多个 reviewer 子代理并行审阅（每个负责 1–3 个章节，在 prompt 中写明章节标题与段落引用）；3) 汇总去重，按严重程度排序输出问题清单（带段落引用）。
不要修改文档；最后询问我希望修改哪些问题，或是否以批注形式标注在文档中。`,
  },
];

export function getCommands(runtime: Runtime): Command[] {
  const skillCmds: Command[] = runtime.skills.list().map((s) => ({
    type: "prompt", name: s.name, description: s.description, argumentHint: s.argumentHint, source: `技能·${s.source}`,
    getPrompt: (args) => `用户调用了 /${s.name} 技能。请严格按以下技能说明执行：\n\n${renderSkill(s, args)}`,
  }));
  const pluginCmds: Command[] = runtime.plugins.flatMap((p) => p.commands.map((c) => ({
    type: "prompt" as const, name: c.name, description: c.description, source: `插件·${p.name}`,
    getPrompt: (args: string) => c.prompt.replace(/\$ARGUMENTS/g, args),
  })));
  const all = [...BUILTIN, ...pluginCmds, ...skillCmds];
  const seen = new Set<string>();
  return all.filter((c) => (seen.has(c.name) ? false : (seen.add(c.name), true)));
}

export type ProcessedInput = { kind: "local"; output: string } | { kind: "prompt"; prompt: string; display: string };

/** 统一的用户输入入口（对应 Claude Code 的 processUserInput） */
export async function processUserInput(input: string, ctx: CommandContext): Promise<ProcessedInput> {
  const text = input.trim();
  const m = /^\/([A-Za-z0-9_-]+)(?:\s+([\s\S]*))?$/.exec(text);
  if (!m) return { kind: "prompt", prompt: text, display: text };
  const cmd = getCommands(ctx.runtime).find((c) => c.name === m[1]);
  if (!cmd) return { kind: "local", output: `未知命令 /${m[1]}。输入 /help 查看所有命令。` };
  const args = (m[2] ?? "").trim();
  if (cmd.type === "local") return { kind: "local", output: await cmd.run(args, ctx) };
  return { kind: "prompt", prompt: cmd.getPrompt(args, ctx), display: text };
}
