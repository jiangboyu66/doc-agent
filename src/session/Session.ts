/**
 * 会话：一篇文档 + 一段对话 + 这段对话里的所有状态
 *
 * 持久化布局（data/sessions/<id>/）：
 *   session.json    元数据、模式、待办、成本、会话级权限规则、版本列表
 *   original.<ext>  上传的原始文件（永不修改，是保真校验的基准）
 *   versions/       每次写操作后的完整快照（支持 /undo、任意版本回滚）
 *   history.json    发给模型的完整消息序列（含 tool_calls / tool 结果 / reasoning_content）
 *   exports/        导出的文件
 */

import fs from "node:fs/promises";
import path from "node:path";
import { nanoid } from "nanoid";
import { loadDocument, detectFormat, type DocumentAdapter, type DocFormat } from "../documents/index.js";
import type { ApiMessage } from "../services/api/deepseek.js";
import type { TodoItem } from "../bridge/protocol.js";
import type { PermissionMode } from "../config/settings.js";
import { emptyCost, type CostState } from "../services/costTracker.js";
import { PATHS } from "../config/settings.js";

export interface VersionInfo {
  v: number;
  label: string;
  at: number;
  changedRefs: string[];
}

export interface SessionMeta {
  id: string;
  filename: string;
  format: DocFormat;
  createdAt: number;
  updatedAt: number;
  mode: PermissionMode;
  trackChanges: boolean;
  author: string;
  todos: TodoItem[];
  cost: CostState;
  sessionRules: string[];
  versions: VersionInfo[];
  currentVersion: number;
  /** 本会话中是否已经读取过文档（写工具的前置条件：先读后改） */
  hasRead: boolean;
  /** 文档结构版本：跨重载、回滚保持单调递增，保证旧的段落引用不会误指到别的段落 */
  structureVersion?: number;
  title?: string;
}

const EXT: Record<DocFormat, string> = { docx: "docx", markdown: "md", html: "html", pdf: "pdf" };

export class Session {
  history: ApiMessage[] = [];
  private constructor(public meta: SessionMeta, public doc: DocumentAdapter, private readonly dir: string) {}

  static sessionsDir(): string {
    return path.join(PATHS.dataDir, "sessions");
  }

  get id() { return this.meta.id; }
  get ext() { return EXT[this.meta.format]; }
  get exportsDir() { return path.join(this.dir, "exports"); }
  private versionFile(v: number) { return path.join(this.dir, "versions", `v${String(v).padStart(4, "0")}.${this.ext}`); }

  static async create(buf: Buffer, filename: string, defaults: { mode: PermissionMode; trackChanges: boolean; author: string }): Promise<Session> {
    const format = detectFormat(filename);
    const doc = await loadDocument(buf, format);
    const id = nanoid(12);
    const dir = path.join(Session.sessionsDir(), id);
    await fs.mkdir(path.join(dir, "versions"), { recursive: true });
    await fs.mkdir(path.join(dir, "exports"), { recursive: true });
    const meta: SessionMeta = {
      id, filename, format, createdAt: Date.now(), updatedAt: Date.now(),
      mode: defaults.mode, trackChanges: format === "docx" ? defaults.trackChanges : false, author: defaults.author,
      todos: [], cost: emptyCost(), sessionRules: [], versions: [{ v: 0, label: "原始文件", at: Date.now(), changedRefs: [] }],
      currentVersion: 0, hasRead: false,
    };
    const s = new Session(meta, doc, dir);
    await fs.writeFile(path.join(dir, `original.${s.ext}`), buf);
    await fs.writeFile(s.versionFile(0), buf);
    await s.save();
    return s;
  }

  static async load(id: string): Promise<Session> {
    if (!/^[A-Za-z0-9_-]{6,40}$/.test(id)) throw new Error("无效的会话 ID");
    const dir = path.join(Session.sessionsDir(), id);
    const meta: SessionMeta = JSON.parse(await fs.readFile(path.join(dir, "session.json"), "utf8"));
    const buf = await fs.readFile(path.join(dir, "versions", `v${String(meta.currentVersion).padStart(4, "0")}.${EXT[meta.format]}`));
    const s = new Session(meta, await loadDocument(buf, meta.format), dir);
    s.doc.structureVersion = meta.structureVersion ?? 0;
    try {
      s.history = JSON.parse(await fs.readFile(path.join(dir, "history.json"), "utf8"));
    } catch { /* 新会话 */ }
    return s;
  }

  static async list(): Promise<SessionMeta[]> {
    try {
      const ids = await fs.readdir(Session.sessionsDir());
      const metas = await Promise.all(ids.map(async (id) => {
        try { return JSON.parse(await fs.readFile(path.join(Session.sessionsDir(), id, "session.json"), "utf8")) as SessionMeta; } catch { return null; }
      }));
      return metas.filter((m): m is SessionMeta => !!m).sort((a, b) => b.updatedAt - a.updatedAt);
    } catch {
      return [];
    }
  }

  async save(): Promise<void> {
    this.meta.updatedAt = Date.now();
    this.meta.structureVersion = this.doc.structureVersion;
    await fs.writeFile(path.join(this.dir, "session.json"), JSON.stringify(this.meta, null, 2), "utf8");
    await fs.writeFile(path.join(this.dir, "history.json"), JSON.stringify(this.history), "utf8");
  }

  async original(): Promise<Buffer> {
    return fs.readFile(path.join(this.dir, `original.${this.ext}`));
  }

  current(): Buffer {
    return this.doc.serialize();
  }

  async versionBuffer(v: number): Promise<Buffer> {
    return fs.readFile(this.versionFile(v));
  }

  /** 写操作成功后提交一个新版本快照 */
  async commit(label: string, changedRefs: string[]): Promise<number> {
    const v = Math.max(...this.meta.versions.map((x) => x.v)) + 1;
    await fs.writeFile(this.versionFile(v), this.doc.serialize());
    this.meta.versions.push({ v, label, at: Date.now(), changedRefs });
    this.meta.currentVersion = v;
    await this.save();
    return v;
  }

  /** 回滚到任意历史版本（回滚本身也记录为一个新版本，历史不会丢） */
  async rollback(v: number): Promise<number> {
    const buf = await this.versionBuffer(v);
    const sv = this.doc.structureVersion;
    this.doc = await loadDocument(buf, this.meta.format);
    this.doc.structureVersion = sv + 1;
    return this.commit(`回滚到 v${v}`, []);
  }

  /** 丢弃未提交的内存修改（写工具失败或保真检查不通过时使用） */
  async revertToCurrent(): Promise<void> {
    const sv = this.doc.structureVersion;
    this.doc = await loadDocument(await this.versionBuffer(this.meta.currentVersion), this.meta.format);
    // 失败的写操作可能已递增结构版本；保留较大的值，宁可让模型重新读取也不让引用错位
    this.doc.structureVersion = sv;
  }

  /** 在文档副本上"干跑"，用于权限确认前生成修改预览，不影响真实文档 */
  async dryRun<T>(fn: (doc: DocumentAdapter) => T): Promise<T> {
    const clone = await loadDocument(this.doc.serialize(), this.meta.format);
    clone.structureVersion = this.doc.structureVersion;
    return fn(clone);
  }

  async writeExport(name: string, data: Buffer): Promise<string> {
    await fs.mkdir(this.exportsDir, { recursive: true });
    const file = path.join(this.exportsDir, name);
    await fs.writeFile(file, data);
    return file;
  }
}
