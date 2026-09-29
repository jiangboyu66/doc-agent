/**
 * 元工具：任务清单、计划模式出口、Skill 调用、子代理
 */

import { z } from "zod";
import { buildTool, type Tool } from "../Tool.js";
import { renderSkill } from "../skills/loadSkills.js";
import { feature } from "../config/features.js";

export const TodoWriteTool = buildTool({
  name: "todo_write",
  category: "meta",
  description: () => `维护本次任务的待办清单（用户在界面侧边栏实时看到进度）。

何时使用：任务包含 3 个及以上步骤、涉及多个章节、或用户一次提出多项要求时。简单的单步修改不需要。
规则：
- 每次调用都提交完整清单（会整体替换旧清单）；
- 同一时间只有一个 in_progress；完成一项立刻标记 completed，再开始下一项；
- content 用祈使句（"统一第三章术语"），activeForm 用进行时（"正在统一第三章术语"）；
- 最后一项通常是"核验修改结果"。`,
  inputSchema: z.object({
    todos: z.array(z.object({
      content: z.string().min(1),
      activeForm: z.string().min(1),
      status: z.enum(["pending", "in_progress", "completed"]),
    })),
  }),
  userFacingName: () => "更新任务清单",
  checkPermissions: async () => ({ behavior: "allow" }),
  async call(i, ctx) {
    ctx.session.meta.todos = i.todos;
    ctx.emit({ type: "todos", todos: i.todos });
    const done = i.todos.filter((t) => t.status === "completed").length;
    return { content: `任务清单已更新（${done}/${i.todos.length} 已完成）。请继续执行下一项。` };
  },
});

export const ExitPlanModeTool = buildTool({
  name: "exit_plan_mode",
  category: "meta",
  description: () => `仅在计划模式下使用：调研完成后，把完整的修改计划提交给用户审批。
计划应列出：要修改哪些段落（引用）、每处怎么改、使用哪些样式、预计影响范围。
用户批准后会自动退出计划模式，你再按计划执行修改；用户若提出调整，请修改计划后重新提交。`,
  inputSchema: z.object({ plan: z.string().min(1).describe("Markdown 格式的修改计划") }),
  userFacingName: () => "提交修改计划",
  isEnabled: (ctx) => ctx.session.meta.mode === "plan",
  checkPermissions: async () => ({ behavior: "ask", reason: "请审阅修改计划，批准后开始执行。" }),
  async call(_i, ctx) {
    ctx.session.meta.mode = "default";
    ctx.emit({ type: "mode_changed", mode: "default" });
    return { content: "用户已批准计划，已退出计划模式。现在按计划逐项执行修改（写操作仍会请求用户确认）。" };
  },
});

export const SkillTool = buildTool({
  name: "skill",
  category: "meta",
  description: (ctx) => {
    const list = ctx.skills.list().map((s) => `- ${s.name}：${s.description}${s.whenToUse ? `（适用：${s.whenToUse}）` : ""}`).join("\n");
    return `加载一个 Skill（预先编写好的专业工作流/规范说明），把它的完整内容注入当前对话，然后严格按其中的步骤执行。
当用户的任务与某个 Skill 的适用场景匹配时，先加载 Skill 再动手；不要凭记忆臆测 Skill 的内容。
可用的 Skill：
${list || "（暂无）"}`;
  },
  inputSchema: z.object({ name: z.string(), args: z.string().optional() }),
  userFacingName: (i) => `加载技能 ${i.name ?? ""}`,
  isEnabled: (ctx) => feature("SKILLS") && ctx.skills.list().length > 0,
  isReadOnly: () => true,
  isConcurrencySafe: () => false,
  async call(i, ctx) {
    const s = ctx.skills.get(i.name);
    if (!s) return { isError: true, content: `找不到 Skill "${i.name}"。可用：${ctx.skills.list().map((x) => x.name).join("、")}` };
    if (s.allowedTools?.length) {
      for (const t of s.allowedTools) if (!ctx.session.meta.sessionRules.includes(t)) ctx.session.meta.sessionRules.push(t);
    }
    return { content: `已加载 Skill "${s.name}"（来源：${s.source}）。请按以下内容执行：\n\n${renderSkill(s, i.args ?? "")}` };
  },
});

export const AgentTool = buildTool({
  name: "agent",
  category: "agent",
  description: (ctx) => {
    const list = ctx.agents.map((a) => `- ${a.agentType}${a.readOnly ? "（只读，可并行）" : ""}：${a.description}。适用：${a.whenToUse}`).join("\n");
    return `启动一个子代理独立完成子任务，子代理结束后只把最终结论返回给你（它的中间过程不占用你的上下文）。

可用类型：
${list}
- fork：继承你当前的完整对话上下文，适合"基于目前了解的情况，深入调研某个问题"。

何时使用：
- 长文档的系统性审阅：按章节同时启动多个 reviewer（在同一轮中发起多个 agent 调用即可并行）；
- 完成修改后，用 verifier 做独立核验（避免自己检查自己）；
- 边界清晰、步骤较多、会产生大量中间输出的子任务。
何时不要用：简单的读取或一两处修改——直接用文档工具更快。

写好子代理任务说明（prompt）：像给刚加入的同事交代工作——说明目标和原因、范围（哪些章节/段落引用，哪些不要碰）、已知信息、期望的输出格式。子代理看不到你的对话（fork 除外）。`;
  },
  inputSchema: z.object({
    subagent_type: z.string().describe("子代理类型，或 fork"),
    description: z.string().describe("3–8 个字的任务名，显示给用户"),
    prompt: z.string().min(1),
  }),
  userFacingName: (i) => `子代理 ${i.subagent_type ?? ""}：${i.description ?? ""}`,
  isEnabled: (ctx) => feature("SUBAGENTS") && ctx.agentId === "main",
  isReadOnly: (i) => i.subagent_type === "reviewer" || i.subagent_type === "verifier",
  isConcurrencySafe: (i) => i.subagent_type === "reviewer" || i.subagent_type === "verifier",
  checkPermissions: async () => ({ behavior: "allow" }),
  async validateInput(i, ctx) {
    if (i.subagent_type !== "fork" && !ctx.agents.some((a) => a.agentType === i.subagent_type)) {
      return { ok: false, message: `未知的子代理类型 "${i.subagent_type}"。可用：${[...ctx.agents.map((a) => a.agentType), "fork"].join("、")}` };
    }
    return { ok: true };
  },
  async call(i, ctx) {
    if (!ctx.runAgent) return { isError: true, content: "当前环境不支持子代理" };
    const result = await ctx.runAgent({
      agentType: i.subagent_type,
      description: i.description,
      prompt: i.prompt,
      parentMessages: i.subagent_type === "fork" ? ctx.parentMessages() : undefined,
    });
    return { content: `子代理「${i.description}」的结论：\n\n${result}` };
  },
});

export const META_TOOLS: Tool[] = [TodoWriteTool, ExitPlanModeTool, SkillTool, AgentTool] as Tool[];
