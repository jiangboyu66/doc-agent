/**
 * Skill 系统（对应 Claude Code 的 skills/ 与 SkillTool）
 *
 * Skill = 一份带元数据的 Markdown 工作流说明书。加载策略是"渐进式披露"：
 *   - 系统提示词里只放每个 Skill 的名称 + 一句话描述 + 适用场景（很短，不占上下文）；
 *   - 模型判断需要时调用 skill 工具，或用户输入 /技能名，才把完整正文注入对话。
 *
 * 来源（后加载的同名 Skill 覆盖先加载的）：内置 skills/ < 插件 < 用户 ~/.doc-agent/skills < 项目 .doc-agent/skills
 * SKILL.md frontmatter 字段：name, description, when_to_use, allowed_tools, argument_hint
 */

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import YAML from "yaml";
import { PATHS } from "../config/settings.js";

export interface Skill {
  name: string;
  description: string;
  whenToUse?: string;
  allowedTools?: string[];
  argumentHint?: string;
  body: string;
  source: string;
  baseDir: string;
}

export class SkillRegistry {
  private map = new Map<string, Skill>();
  add(s: Skill) { this.map.set(s.name, s); }
  get(name: string) { return this.map.get(name.replace(/^\//, "")); }
  list(): Skill[] { return [...this.map.values()].sort((a, b) => a.name.localeCompare(b.name)); }
}

export const BUNDLED_SKILLS_DIR = fileURLToPath(new URL("../../skills/", import.meta.url));

export function parseSkillFile(raw: string, fallbackName: string, source: string, baseDir: string): Skill | null {
  const m = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/.exec(raw);
  let fm: Record<string, unknown> = {};
  let body = raw;
  if (m) {
    try { fm = YAML.parse(m[1]) ?? {}; } catch { return null; }
    body = m[2];
  }
  const name = String(fm.name ?? fallbackName).trim();
  if (!name) return null;
  const tools = fm.allowed_tools ?? fm["allowed-tools"];
  return {
    name,
    description: String(fm.description ?? body.split("\n").find((l) => l.trim())?.replace(/^#+\s*/, "") ?? ""),
    whenToUse: fm.when_to_use ? String(fm.when_to_use) : fm.whenToUse ? String(fm.whenToUse) : undefined,
    allowedTools: Array.isArray(tools) ? tools.map(String) : typeof tools === "string" ? tools.split(/[,\s]+/).filter(Boolean) : undefined,
    argumentHint: fm.argument_hint ? String(fm.argument_hint) : undefined,
    body: body.trim(),
    source,
    baseDir,
  };
}

export async function loadSkillsFromDir(dir: string, source: string): Promise<Skill[]> {
  let entries: string[];
  try { entries = await fs.readdir(dir); } catch { return []; }
  const out: Skill[] = [];
  for (const e of entries) {
    const full = path.join(dir, e);
    try {
      const st = await fs.stat(full);
      if (st.isDirectory()) {
        const raw = await fs.readFile(path.join(full, "SKILL.md"), "utf8").catch(() => null);
        if (raw) {
          const s = parseSkillFile(raw, e, source, full);
          if (s) out.push(s);
        }
      } else if (e.endsWith(".md")) {
        const s = parseSkillFile(await fs.readFile(full, "utf8"), e.replace(/\.md$/, ""), source, dir);
        if (s) out.push(s);
      }
    } catch { /* 跳过无法读取的条目 */ }
  }
  return out;
}

export async function loadSkills(pluginSkillDirs: Array<{ dir: string; plugin: string }> = []): Promise<SkillRegistry> {
  const reg = new SkillRegistry();
  const layers: Array<Promise<Skill[]>> = [
    loadSkillsFromDir(BUNDLED_SKILLS_DIR, "内置"),
    ...pluginSkillDirs.map((p) => loadSkillsFromDir(p.dir, `插件:${p.plugin}`)),
    loadSkillsFromDir(path.join(PATHS.userDir, "skills"), "用户"),
    loadSkillsFromDir(path.join(PATHS.projectDir, "skills"), "项目"),
  ];
  for (const list of await Promise.all(layers)) for (const s of list) reg.add(s);
  return reg;
}

/** 把 Skill 正文展开为可执行的提示词（替换参数占位符） */
export function renderSkill(s: Skill, args: string): string {
  const body = s.body.replace(/\$ARGUMENTS/g, args || "（未提供参数）").replace(/\$\{SKILL_DIR\}/g, s.baseDir);
  return `<skill name="${s.name}">\n${body}\n</skill>${args ? `\n\n本次参数：${args}` : ""}`;
}
