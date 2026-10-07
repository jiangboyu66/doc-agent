/**
 * Web 服务入口
 *
 * 一轮对话 = 一个 POST 请求，响应是 SSE 事件流（text/event-stream）。当模型需要用户确认某个修改时，
 * 事件流里出现 permission_request，引擎在服务端 await 用户的决定；前端通过另一个 POST 回传决定，
 * 引擎继续执行——整个过程在同一个事件流里完成，不存在"暂停-恢复"导致的状态不一致。
 */

import "dotenv/config";
import express, { type Request, type Response } from "express";
import multer from "multer";
import path from "node:path";
import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { bootstrap, createClient, type Runtime } from "../bootstrap.js";
import { Session } from "../session/Session.js";
import { QueryEngine } from "../QueryEngine.js";
import { processUserInput, getCommands } from "../commands.js";
import { buildTranscript } from "../session/transcript.js";
import { toPdfViaOffice, exportDocument, type ExportFormat } from "../services/convert.js";
import { featureTable } from "../config/features.js";
import type { EngineEvent, PermissionRequest, PermissionResponse } from "../bridge/protocol.js";
import { PermissionModeSchema } from "../config/settings.js";
import { convertPdfSession, convertPdfInPlace } from "../session/convertPdf.js";

interface Live {
  session: Session;
  abort: AbortController | null;
  pending: Map<string, { request: PermissionRequest; resolve: (r: PermissionResponse) => void }>;
}

const live = new Map<string, Live>();
let runtime: Runtime;

async function getLive(id: string): Promise<Live> {
  let l = live.get(id);
  if (!l) {
    l = { session: await Session.load(id), abort: null, pending: new Map() };
    live.set(id, l);
  }
  return l;
}

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 200 * 1024 * 1024 } });
const app = express();
app.use(express.json({ limit: "5mb" }));

const wrap = (fn: (req: Request, res: Response) => Promise<unknown>) => (req: Request, res: Response) =>
  fn(req, res).catch((e: any) => {
    if (!res.headersSent) res.status(e?.code === "ENOENT" ? 404 : 400).json({ error: e?.message ?? String(e) });
  });

app.get("/api/health", (_req, res) => res.json({ ok: true }));

app.get("/api/runtime", (_req, res) => {
  res.json({
    model: runtime.settings.model,
    thinking: runtime.settings.thinking,
    thinkingAvailable: featureTable().THINKING,
    apiKey: !!process.env.DEEPSEEK_API_KEY,
    external: runtime.external,
    pdfEngines: runtime.pdfEngines,
    features: featureTable(),
    skills: runtime.skills.list().map((s) => ({ name: s.name, description: s.description, source: s.source })),
    agents: runtime.agents.map((a) => ({ type: a.agentType, description: a.description, readOnly: a.readOnly })),
    commands: getCommands(runtime).map((c) => ({ name: c.name, description: c.description, argumentHint: c.argumentHint })),
  });
});

app.get("/api/sessions", wrap(async (_req, res) => res.json(await Session.list())));

app.post("/api/sessions", upload.single("file"), wrap(async (req, res) => {
  if (!req.file) throw new Error("请选择要上传的文件");
  // multer 按 latin1 解析文件名，中文文件名需要转回 UTF-8
  const filename = Buffer.from(req.file.originalname, "latin1").toString("utf8");
  const s = await Session.create(req.file.buffer, filename, {
    mode: runtime.settings.permissionMode, trackChanges: runtime.settings.trackChanges, author: runtime.settings.author,
  });
  live.set(s.id, { session: s, abort: null, pending: new Map() });
  res.json({ meta: s.meta, summary: s.doc.summary() });
}));

app.get("/api/sessions/:id", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  res.json({
    meta: l.session.meta,
    summary: l.session.doc.summary(),
    transcript: buildTranscript(l.session.history),
    busy: !!l.abort,
    pending: [...l.pending.values()].map((p) => p.request),
  });
}));

app.patch("/api/sessions/:id/settings", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const m = l.session.meta;
  if (req.body.mode !== undefined) m.mode = PermissionModeSchema.parse(req.body.mode);
  if (req.body.trackChanges !== undefined) m.trackChanges = !!req.body.trackChanges && l.session.doc.capabilities.trackChanges;
  if (typeof req.body.author === "string" && req.body.author.trim()) m.author = req.body.author.trim().slice(0, 60);
  if (typeof req.body.thinking === "boolean") m.thinking = req.body.thinking;
  await l.session.save();
  res.json({ meta: m });
}));

app.post("/api/sessions/:id/messages", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  if (l.abort) return res.status(409).json({ error: "上一轮还在进行中，请等待完成或先中断。" });
  const text = String(req.body.text ?? "").trim();
  if (!text) return res.status(400).json({ error: "消息不能为空" });

  res.writeHead(200, { "Content-Type": "text/event-stream; charset=utf-8", "Cache-Control": "no-cache", Connection: "keep-alive", "X-Accel-Buffering": "no" });
  const send = (e: EngineEvent) => res.write(`data: ${JSON.stringify(e)}\n\n`);
  const heartbeat = setInterval(() => res.write(": ping\n\n"), 15_000);
  const abort = new AbortController();
  l.abort = abort;
  res.on("close", () => {
    if (!res.writableEnded) abort.abort();
  });

  try {
    let engine: QueryEngine | null = null;
    const getEngine = () =>
      (engine ??= new QueryEngine(l.session, {
        client: createClient(runtime.settings),
        settings: runtime.settings,
        skills: runtime.skills,
        agents: runtime.agents,
        memory: runtime.memory,
        canUseTool: (request, signal) =>
          new Promise<PermissionResponse>((resolve) => {
            const timer = setTimeout(() => finish({ decision: "deny", feedback: "等待确认超时" }), 30 * 60 * 1000);
            const finish = (r: PermissionResponse) => {
              clearTimeout(timer);
              l.pending.delete(request.requestId);
              resolve(r);
            };
            l.pending.set(request.requestId, { request, resolve: finish });
            signal.addEventListener("abort", () => finish({ decision: "deny", feedback: "用户中断了操作" }), { once: true });
          }),
      }));
    const input = await processUserInput(text, { session: l.session, runtime, compact: (ins) => getEngine().compactNow(abort.signal, ins) });
    if (input.kind === "local") {
      send({ type: "command_output", text: input.output });
      send({ type: "done", reason: "command" });
    } else {
      for await (const e of getEngine().submitMessage(input.prompt, abort.signal)) send(e);
    }
  } catch (e: any) {
    send({ type: "error", message: e?.message ?? String(e) });
    send({ type: "done", reason: "error" });
  } finally {
    clearInterval(heartbeat);
    l.abort = null;
    for (const p of l.pending.values()) p.resolve({ decision: "deny", feedback: "会话已结束" });
    await l.session.save().catch(() => {});
    res.end();
  }
}));

app.post("/api/sessions/:id/permissions/:requestId", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const p = l.pending.get(String(req.params.requestId));
  if (!p) return res.status(404).json({ error: "该确认请求已失效" });
  const b = req.body ?? {};
  p.resolve(b.decision === "allow"
    ? { decision: "allow", remember: b.remember === "session" || b.remember === "project" ? b.remember : undefined }
    : { decision: "deny", feedback: typeof b.feedback === "string" ? b.feedback.slice(0, 2000) : undefined });
  res.json({ ok: true });
}));

app.post("/api/sessions/:id/interrupt", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  l.abort?.abort();
  res.json({ ok: true });
}));

app.post("/api/sessions/:id/rollback", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  if (l.abort) throw new Error("请等待当前这一轮完成后再回滚");
  const v = await l.session.rollback(Number(req.body.version));
  res.json({ meta: l.session.meta, version: v });
}));

const MIME: Record<string, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  md: "text/markdown; charset=utf-8", html: "text/html; charset=utf-8", pdf: "application/pdf",
};

app.get("/api/sessions/:id/document", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const v = req.query.version !== undefined ? Number(req.query.version) : l.session.meta.currentVersion;
  const up = req.query.original ? await l.session.uploaded() : null;
  const buf = up ? up.data : await l.session.versionBuffer(v);
  const ext = up ? up.ext : ({ docx: "docx", markdown: "md", html: "html", pdf: "pdf" } as const)[l.session.versionFormat(v)];
  const base = l.session.meta.filename.replace(/\.[^.]+$/, "");
  const name = up ? up.filename : `${base}-v${v}.${ext}`;
  res.setHeader("Content-Type", MIME[ext] ?? "application/octet-stream");
  res.setHeader("Content-Disposition", `${req.query.download ? "attachment" : "inline"}; filename*=UTF-8''${encodeURIComponent(name)}`);
  res.end(buf);
}));

const previewCache = new Map<string, Promise<Buffer>>();
app.get("/api/sessions/:id/preview.pdf", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const v = req.query.version !== undefined ? Number(req.query.version) : l.session.meta.currentVersion;
  const vf = l.session.versionFormat(v);
  if (vf === "pdf") {
    res.setHeader("Content-Type", "application/pdf");
    return res.end(await l.session.versionBuffer(v));
  }
  const key = `${l.session.id}:${v}`;
  if (!previewCache.has(key)) {
    const p = l.session.versionBuffer(v).then((b) => toPdfViaOffice(b, ({ docx: "docx", markdown: "md", html: "html", pdf: "pdf" } as const)[vf]));
    p.catch(() => previewCache.delete(key));
    previewCache.set(key, p);
  }
  res.setHeader("Content-Type", "application/pdf");
  res.end(await previewCache.get(key)!);
}));

/** 上传素材（图片），供 Agent 用 doc_insert_image 插入文档 */
const assetUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024, files: 10 } });
app.post("/api/sessions/:id/assets", assetUpload.array("files", 10), wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const files = (req.files as Express.Multer.File[] | undefined) ?? [];
  if (!files.length) throw new Error("请选择图片文件");
  const saved: string[] = [];
  for (const f of files) {
    const name = Buffer.from(f.originalname, "latin1").toString("utf8");
    if (!/\.(png|jpe?g|gif|bmp|svg|webp)$/i.test(name)) throw new Error(`不支持的图片格式：${name}（支持 PNG / JPG / GIF / BMP / SVG / WebP）`);
    saved.push(await l.session.saveAsset(name, f.buffer));
  }
  res.json({ assets: saved, all: l.session.meta.assets ?? [] });
}));

app.get("/api/sessions/:id/assets/:name", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const buf = await l.session.readAsset(String(req.params.name));
  const ext = path.extname(String(req.params.name)).slice(1).toLowerCase();
  res.setHeader("Content-Type", ({ png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", webp: "image/webp" } as Record<string, string>)[ext] ?? "application/octet-stream");
  res.end(buf);
}));

/** PDF → Word：生成一个新的可编辑会话（原 PDF 会话不变） */
app.post("/api/sessions/:id/convert", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const engine = ["builtin", "pdf2docx", "libreoffice"].includes(req.body?.engine) ? req.body.engine : "builtin";
  // 转换后用于编辑：默认流式（文字增减时自动重排）；明确要求时才逐页固定
  const mode = req.body?.mode === "exact" ? "exact" : "flow";
  if (req.body?.inPlace) {
    // 在当前会话内转换：对话保留，转换后直接继续编辑
    if (l.abort) throw new Error("Agent 正在运行，请等待完成或先中断。");
    const r = await convertPdfInPlace(l.session, { engine, mode });
    const version = await l.session.commit(r.label, []);
    return res.json({ meta: l.session.meta, report: r.report, text: r.text, version });
  }
  const { session, report, text } = await convertPdfSession(l.session, {
    engine, mode,
    defaults: { mode: runtime.settings.permissionMode, trackChanges: runtime.settings.trackChanges, author: runtime.settings.author },
  });
  live.set(session.id, { session, abort: null, pending: new Map() });
  res.json({ meta: session.meta, report, text });
}));

/** 直接导出（不经过模型）：同格式导出即编辑后的原文件，逐字节保真 */
app.post("/api/sessions/:id/export", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const fmt = String(req.body?.format ?? l.session.meta.format) as ExportFormat;
  if (!["docx", "pdf", "html", "markdown"].includes(fmt)) throw new Error("格式可选：docx / pdf / html / markdown");
  const pdfMode = req.body?.mode === "flow" ? "flow" : "exact";
  const pdfEngine = ["pdf2docx", "libreoffice"].includes(req.body?.engine) ? req.body.engine : "builtin";
  const out = await exportDocument(l.session.current(), l.session.meta.format, fmt, { pdfMode, pdfEngine });
  const base = l.session.meta.filename.replace(/\.[^.]+$/, "");
  // PDF 导出 Word：直接用原文件名（例如 表单.pdf → 表单.docx）
  const name = l.session.meta.format === "pdf" ? `${base}${pdfMode === "flow" ? "（可改写）" : ""}.${out.ext}` : `${base}-v${l.session.meta.currentVersion}.${out.ext}`;
  await l.session.writeExport(name, out.data);
  res.json({ name, url: `/api/sessions/${l.session.id}/exports/${encodeURIComponent(name)}`, note: out.note, lossy: out.lossy });
}));

app.delete("/api/sessions/:id", wrap(async (req, res) => {
  const id = String(req.params.id);
  const l = live.get(id);
  if (l?.abort) throw new Error("该会话正在运行，请先中断");
  if (!/^[A-Za-z0-9_-]{6,40}$/.test(id)) throw new Error("无效的会话 ID");
  live.delete(id);
  await fs.rm(path.join(Session.sessionsDir(), id), { recursive: true, force: true });
  res.json({ ok: true });
}));

app.get("/api/sessions/:id/exports/:name", wrap(async (req, res) => {
  const l = await getLive(String(req.params.id));
  const name = path.basename(String(req.params.name));
  const file = path.join(l.session.exportsDir, name);
  await fs.access(file);
  res.download(file, name);
}));

const webDist = fileURLToPath(new URL("../../web/dist/", import.meta.url));
// 前端资源：带哈希的 assets 长期缓存；index.html 每次都重新验证，保证更新后不会加载到旧版前端
app.use(express.static(webDist, {
  setHeaders(res, file) {
    if (file.endsWith("index.html")) res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
    else if (file.includes(`${path.sep}assets${path.sep}`)) res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
  },
}));
app.get(/^\/(?!api\/).*/, (_req, res) => {
  res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
  res.sendFile(path.join(webDist, "index.html"), (err) => {
    if (err) res.status(200).send("前端尚未构建：请先运行 npm run build:web");
  });
});

async function main() {
  runtime = await bootstrap();
  const port = Number(process.env.PORT || 3939);
  app.listen(port, () => {
    console.log(`\n  文案 Agent 已启动：http://localhost:${port}`);
    console.log(`  模型 ${runtime.settings.model}（思考模式 ${runtime.settings.thinking}）｜技能 ${runtime.skills.list().length} 个｜子代理 ${runtime.agents.length} 个｜插件 ${runtime.plugins.length} 个｜启动 ${runtime.startupMs} ms`);
    if (!process.env.DEEPSEEK_API_KEY) console.warn("  ⚠ 未配置 DEEPSEEK_API_KEY，请在 .env 中填写后重启");
    if (!runtime.external.soffice) console.warn("  ⚠ 未找到 LibreOffice（soffice），PDF 导出与版式预览不可用");
    if (!runtime.external.pandoc) console.warn("  ⚠ 未找到 pandoc，跨格式导出（HTML/Markdown）不可用");
  });
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
