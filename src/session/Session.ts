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
  /** 该版本的文件格式（会话中途转换格式后，新旧版本格式不同；缺省为会话最初的格式） */
  format?: DocFormat;
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
  /** 由其他文件转换而来（例如 PDF → Word）：来源与转换报告 */
  origin?: { fromSession: string; fromFile: string; engine: string; mode?: string; report: string; coverage: number };
  title?: string;
  /** 本会话是否开启思考模式（未设置时跟随全局配置） */
  thinking?: boolean;
  /** 会话最初上传的格式与文件名（在会话内把 PDF 转换为 Word 后，用于下载原始文件、读取旧版本） */
  sourceFormat?: DocFormat;
  sourceFile?: string;
  /** 用户上传的素材（图片）文件名 */
  assets?: string[];
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
  /** 某个版本的格式：版本自身记录 → 会话最初的格式 → 当前格式 */
  versionFormat(v: number): DocFormat {
    return this.meta.versions.find((x) => x.v === v)?.format ?? this.meta.sourceFormat ?? this.meta.format;
  }
  private versionFile(v: number) { return path.join(this.dir, "versions", `v${String(v).padStart(4, "0")}.${EXT[this.versionFormat(v)]}`); }

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
    const fmt = meta.versions.find((x) => x.v === meta.currentVersion)?.format ?? meta.sourceFormat ?? meta.format;
    const buf = await fs.readFile(path.join(dir, "versions", `v${String(meta.currentVersion).padStart(4, "0")}.${EXT[fmt]}`));
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

  /** 保真校验的基准：当前格式的"原始文件"（会话内转换格式后，是转换得到的初始文件） */
  async original(): Promise<Buffer> {
    return fs.readFile(path.join(this.dir, `original.${this.ext}`));
  }

  /** 用户最初上传的文件（及其文件名、扩展名） */
  async uploaded(): Promise<{ data: Buffer; filename: string; ext: string }> {
    const fmt = this.meta.sourceFormat ?? this.meta.format;
    return { data: await fs.readFile(path.join(this.dir, `original.${EXT[fmt]}`)), filename: this.meta.sourceFile ?? this.meta.filename, ext: EXT[fmt] };
  }

  /**
   * 在本会话内切换到另一种格式的文档（例如 PDF → Word）：对话历史、待办、权限规则都保留，
   * 新文档成为后续编辑的对象与保真基准；旧版本仍可查看、回滚。调用方随后 commit() 生成新版本。
   */
  async switchDocument(buf: Buffer, format: DocFormat, filename: string, origin?: SessionMeta["origin"]): Promise<void> {
    const doc = await loadDocument(buf, format);
    if (!this.meta.sourceFormat) {
      this.meta.sourceFormat = this.meta.format;
      this.meta.sourceFile = this.meta.filename;
      // 旧版本补上格式标记
      for (const v of this.meta.versions) v.format ??= this.meta.format;
    }
    doc.structureVersion = this.doc.structureVersion + 1;
    this.doc = doc;
    this.meta.format = format;
    this.meta.filename = filename;
    if (origin) this.meta.origin = origin;
    this.meta.hasRead = false;
    if (format !== "docx") this.meta.trackChanges = false;
    await fs.writeFile(path.join(this.dir, `original.${EXT[format]}`), buf);
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
    this.meta.versions.push({ v, label, at: Date.now(), changedRefs, format: this.meta.format });
    await fs.writeFile(this.versionFile(v), this.doc.serialize());
    this.meta.currentVersion = v;
    await this.save();
    return v;
  }

  /** 回滚到任意历史版本（回滚本身也记录为一个新版本，历史不会丢） */
  async rollback(v: number): Promise<number> {
    const buf = await this.versionBuffer(v);
    const fmt = this.versionFormat(v);
    const sv = this.doc.structureVersion;
    this.doc = await loadDocument(buf, fmt);
    this.doc.structureVersion = sv + 1;
    if (fmt !== this.meta.format) {
      // 回滚到转换之前的版本：格式一并切回
      this.meta.format = fmt;
      const src = this.meta.sourceFile ?? this.meta.filename;
      this.meta.filename = fmt === this.meta.sourceFormat ? src : `${src.replace(/\.[^.]+$/, "")}.${EXT[fmt]}`;
      this.meta.hasRead = false;
    }
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

  // ---- 素材（用户上传的图片，供插入文档）----
  get assetsDir() { return path.join(this.dir, "assets"); }

  async saveAsset(name: string, data: Buffer): Promise<string> {
    await fs.mkdir(this.assetsDir, { recursive: true });
    const clean = path.basename(name).replace(/[\\/:*?"<>|\s]+/g, "_").slice(-80) || "image.png";
    let final = clean;
    for (let n = 2; ; n++) {
      try { await fs.access(path.join(this.assetsDir, final)); } catch { break; }
      final = clean.replace(/(\.[^.]+)?$/, `-${n}$1`);
    }
    await fs.writeFile(path.join(this.assetsDir, final), data);
    this.meta.assets = [...(this.meta.assets ?? []), final];
    await this.save();
    return final;
  }

  async listAssets(): Promise<Array<{ name: string; size: number }>> {
    try {
      const names = await fs.readdir(this.assetsDir);
      return Promise.all(names.map(async (n) => ({ name: n, size: (await fs.stat(path.join(this.assetsDir, n))).size })));
    } catch {
      return [];
    }
  }

  async readAsset(name: string): Promise<Buffer> {
    const safe = path.basename(name);
    try {
      return await fs.readFile(path.join(this.assetsDir, safe));
    } catch {
      const list = (await this.listAssets()).map((a) => a.name);
      throw new Error(`找不到素材"${name}"。${list.length ? `已上传的素材：${list.join("、")}` : "本会话还没有上传任何图片，请让用户先点击输入框上方的「图片」按钮上传。"}`);
    }
  }

  async writeExport(name: string, data: Buffer): Promise<string> {
    await fs.mkdir(this.exportsDir, { recursive: true });
    const file = path.join(this.exportsDir, name);
    await fs.writeFile(file, data);
    return file;
  }
}
