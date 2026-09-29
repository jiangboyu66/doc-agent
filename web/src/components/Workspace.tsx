import { useCallback, useEffect, useRef, useState } from "react";
import { api, streamMessage } from "../api";
import { applyEvent, fromTranscript, userItem } from "../chatState";
import type { ChatItem, DocumentSummary, EngineEvent, PermissionMode, RuntimeInfo, SessionMeta, TodoItem } from "../types";
import { Chat } from "./Chat";
import { DocPreview } from "./DocPreview";
import { OutlinePanel, VersionsPanel } from "./Panels";

const MODES: Array<[PermissionMode, string, string]> = [
  ["default", "逐一确认", "每次修改前展示预览，由你确认"],
  ["acceptEdits", "自动接受编辑", "普通修改自动生效，删除等破坏性操作仍需确认"],
  ["plan", "只做计划", "只读调研并提交修改计划，确认后才开始修改"],
  ["bypassPermissions", "全部免确认", "所有操作自动执行（每次修改仍会保存版本，可回滚）"],
];

type Tab = "doc" | "outline" | "versions";

export function Workspace({ id, runtime, onBack }: { id: string; runtime: RuntimeInfo | null; onBack: () => void }) {
  const [meta, setMeta] = useState<SessionMeta | null>(null);
  const [summary, setSummary] = useState<DocumentSummary | null>(null);
  const [items, setItems] = useState<ChatItem[]>([]);
  const [todos, setTodos] = useState<TodoItem[]>([]);
  const [busy, setBusy] = useState(false);
  const [viewing, setViewing] = useState(0);
  const [tab, setTab] = useState<Tab>("doc");
  const [err, setErr] = useState("");
  const [exporting, setExporting] = useState<string | null>(null);
  const [exported, setExported] = useState<{ name: string; url: string; note: string; lossy: boolean } | null>(null);
  const [menu, setMenu] = useState(false);
  const followLatest = useRef(true);

  const refresh = useCallback(async () => {
    const d = await api.session(id);
    setMeta(d.meta);
    setSummary(d.summary);
    setTodos(d.meta.todos ?? []);
    if (followLatest.current) setViewing(d.meta.currentVersion);
    return d;
  }, [id]);

  useEffect(() => {
    refresh()
      .then((d) => {
        setItems(fromTranscript(d.transcript, d.pending));
        setViewing(d.meta.currentVersion);
        if (d.busy) setItems((x) => [...x, { kind: "notice", key: "busy", tone: "warn", text: "这个会话正在另一个窗口中运行。" }]);
      })
      .catch((e) => setErr(e.message));
  }, [refresh]);

  const onEvent = useCallback((e: EngineEvent) => {
    setItems((x) => applyEvent(x, e));
    if (e.type === "todos") setTodos(e.todos);
    if (e.type === "usage") setMeta((m) => (m ? { ...m, cost: { ...m.cost, costUsd: e.totalCostUsd } } : m));
    if (e.type === "mode_changed") setMeta((m) => (m ? { ...m, mode: e.mode as PermissionMode } : m));
    if (e.type === "doc_changed") {
      followLatest.current = true;
      refresh().catch(() => {});
    }
  }, [refresh]);

  const send = async (text: string) => {
    if (busy) return;
    setErr("");
    setItems((x) => [...x, userItem(text)]);
    setBusy(true);
    try {
      await streamMessage(id, text, onEvent);
    } catch (e: any) {
      setItems((x) => applyEvent(applyEvent(x, { type: "error", message: e.message }), { type: "done", reason: "error" }));
    } finally {
      setBusy(false);
      refresh().catch(() => {});
    }
  };

  const decide = async (requestId: string, body: { decision: "allow" | "deny"; remember?: "session" | "project"; feedback?: string }) => {
    try {
      await api.permission(id, requestId, body);
      setItems((x) => applyEvent(x, { type: "permission_resolved", requestId, decision: body.decision }));
    } catch (e: any) {
      setErr(e.message);
    }
  };

  const patch = async (p: Partial<Pick<SessionMeta, "mode" | "trackChanges" | "author">>) => {
    try {
      const r = await api.settings(id, p);
      setMeta(r.meta);
    } catch (e: any) {
      setErr(e.message);
    }
  };

  const rollback = async (v: number) => {
    if (!confirm(`回滚到 v${v}？回滚会生成一个新版本，之后仍可回到现在的状态。`)) return;
    try {
      await api.rollback(id, v);
      followLatest.current = true;
      await refresh();
      setItems((x) => [...x, { kind: "notice", key: `rb${Date.now()}`, tone: "info", text: `已回滚到 v${v} 的内容` }]);
    } catch (e: any) {
      setErr(e.message);
    }
  };

  const doExport = async (format: string) => {
    setMenu(false);
    setExporting(format);
    setExported(null);
    setErr("");
    try {
      setExported(await api.exportAs(id, format));
    } catch (e: any) {
      setErr(`导出失败：${e.message}`);
    } finally {
      setExporting(null);
    }
  };

  if (!meta || !summary) {
    return (
      <div className="center-screen">
        {err ? <div className="notice error">{err} <button className="btn small" onClick={onBack}>返回</button></div> : <span className="muted"><span className="spinner" aria-hidden /> 加载中…</span>}
      </div>
    );
  }

  const readOnly = meta.format === "pdf";
  const exportTargets = meta.format === "pdf" ? [] : (["docx", "pdf", "html", "markdown"] as const).filter((f) => f !== meta.format);

  return (
    <div className="workspace">
      <header className="topbar">
        <button className="btn ghost icon" onClick={onBack} aria-label="返回首页" title="返回首页">←</button>
        <div className="doc-title">
          <span className="badge">{{ docx: "Word", markdown: "Markdown", html: "HTML", pdf: "PDF" }[meta.format]}</span>
          <span className="name" title={meta.filename}>{meta.filename}</span>
        </div>

        <div className="controls">
          {!readOnly && (
            <label className="ctl">
              <span>权限</span>
              <select value={meta.mode} onChange={(e) => patch({ mode: e.target.value as PermissionMode })} title={MODES.find((m) => m[0] === meta.mode)?.[2]}>
                {MODES.map(([v, label, tip]) => <option key={v} value={v} title={tip}>{label}</option>)}
              </select>
            </label>
          )}
          {meta.format === "docx" && (
            <label className="ctl toggle" title="开启后，所有修改记录为 Word 修订（可在 Word 中逐条接受/拒绝）">
              <input type="checkbox" checked={meta.trackChanges} onChange={(e) => patch({ trackChanges: e.target.checked })} />
              <span>修订模式</span>
            </label>
          )}
          {meta.format === "docx" && (
            <label className="ctl">
              <span>作者</span>
              <input className="author" defaultValue={meta.author} onBlur={(e) => e.target.value.trim() && e.target.value !== meta.author && patch({ author: e.target.value })} />
            </label>
          )}
          <span className="cost" title={`${meta.cost.requests} 次请求 · 输入 ${meta.cost.promptTokens.toLocaleString()}（缓存命中 ${meta.cost.cachedTokens.toLocaleString()}）· 输出 ${meta.cost.completionTokens.toLocaleString()} tokens`}>
            ${meta.cost.costUsd.toFixed(4)}
          </span>
          <div className="menu-wrap">
            <button className="btn primary" onClick={() => setMenu(!menu)} aria-expanded={menu}>
              {exporting ? "导出中…" : "下载 / 导出"}
            </button>
            {menu && (
              <div className="menu" onMouseLeave={() => setMenu(false)}>
                <a href={api.documentUrl(id, { download: true })} download onClick={() => setMenu(false)}>
                  <b>当前版本 v{meta.currentVersion}</b>
                  <span>原格式，未修改部分逐字节保真</span>
                </a>
                <a href={api.documentUrl(id, { original: true, download: true })} download onClick={() => setMenu(false)}>
                  <b>原始文件</b>
                  <span>上传时的文件，从未被修改</span>
                </a>
                {exportTargets.map((f) => (
                  <button key={f} onClick={() => doExport(f)} disabled={!!exporting}>
                    <b>导出为 {f === "markdown" ? "Markdown" : f.toUpperCase()}</b>
                    <span>{f === "pdf" ? (meta.format === "docx" ? "LibreOffice 按 Word 版式渲染" : "转换后渲染") : "跨格式转换，版式可能有损"}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      </header>

      {(err || exported) && (
        <div className="banner-row">
          {err && <div className="notice error">{err} <button className="link" onClick={() => setErr("")}>关闭</button></div>}
          {exported && (
            <div className={`notice ${exported.lossy ? "warn" : "success"}`}>
              已导出 <a href={exported.url} download>{exported.name}</a> — {exported.note} <button className="link" onClick={() => setExported(null)}>关闭</button>
            </div>
          )}
        </div>
      )}

      <div className="split">
        <section className="left" aria-label="文档">
          <nav className="tabs" role="tablist">
            {([["doc", "文档"], ["outline", "大纲"], ["versions", `版本（${meta.versions.length}）`]] as Array<[Tab, string]>).map(([t, label]) => (
              <button key={t} role="tab" aria-selected={tab === t} className={tab === t ? "on" : ""} onClick={() => setTab(t)}>{label}</button>
            ))}
          </nav>
          <div className="left-body">
            {tab === "doc" && (
              <DocPreview sessionId={id} format={meta.format} version={viewing} isCurrent={viewing === meta.currentVersion} sofficeAvailable={!!runtime?.external.soffice} />
            )}
            {tab === "outline" && <OutlinePanel summary={summary} onAsk={send} />}
            {tab === "versions" && (
              <VersionsPanel
                meta={meta}
                viewing={viewing}
                busy={busy}
                onView={(v) => { followLatest.current = v === meta.currentVersion; setViewing(v); setTab("doc"); }}
                onRollback={rollback}
              />
            )}
          </div>
        </section>
        <Chat items={items} busy={busy} todos={todos} runtime={runtime} onSend={send} onInterrupt={() => api.interrupt(id)} onDecide={decide} />
      </div>
    </div>
  );
}
