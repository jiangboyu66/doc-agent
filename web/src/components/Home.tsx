import { useEffect, useRef, useState } from "react";
import { api } from "../api";
import type { RuntimeInfo, SessionMeta } from "../types";

const FORMAT_LABEL: Record<string, string> = { docx: "Word", markdown: "Markdown", html: "HTML", pdf: "PDF" };

export function Home({ runtime, onOpen }: { runtime: RuntimeInfo | null; onOpen: (id: string) => void }) {
  const [sessions, setSessions] = useState<SessionMeta[] | null>(null);
  const [drag, setDrag] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [err, setErr] = useState("");
  const input = useRef<HTMLInputElement>(null);

  const load = () => api.sessions().then(setSessions, (e) => setErr(e.message));
  useEffect(() => { load(); }, []);

  const upload = async (file: File | undefined) => {
    if (!file) return;
    setErr("");
    setUploading(true);
    try {
      const r = await api.upload(file);
      onOpen(r.meta.id);
    } catch (e: any) {
      setErr(e.message);
    } finally {
      setUploading(false);
    }
  };

  return (
    <main className="home">
      <header className="home-head">
        <h1 className="serif">文案 Agent</h1>
        <p className="lede">面向 Word、Markdown、HTML 的保真编辑助手。只改你要改的，其余每一个字节保持原样。</p>
      </header>

      {runtime && (!runtime.apiKey || !runtime.external.soffice) && (
        <div className="setup">
          {!runtime.apiKey && <div className="notice error">未配置 DEEPSEEK_API_KEY：在项目根目录的 .env 中填写后重启服务。本地命令（/export、/undo 等）不受影响。</div>}
          {!runtime.external.soffice && <div className="notice warn">未检测到 LibreOffice：PDF 导出与“精确版式”预览不可用。安装后可在 .env 中用 SOFFICE_PATH 指定路径。</div>}
          {!runtime.external.pandoc && <div className="notice warn">未检测到 pandoc：跨格式导出（如 Word → HTML）不可用；同格式编辑与导出不受影响。</div>}
        </div>
      )}

      <div
        className={`drop ${drag ? "drag" : ""}`}
        role="button"
        tabIndex={0}
        onClick={() => input.current?.click()}
        onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && input.current?.click()}
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => { e.preventDefault(); setDrag(false); upload(e.dataTransfer.files[0]); }}
      >
        <input ref={input} type="file" hidden accept=".docx,.md,.markdown,.html,.htm,.pdf" onChange={(e) => upload(e.target.files?.[0])} />
        <div className="drop-icon" aria-hidden>⬆</div>
        <div className="drop-title">{uploading ? "正在解析文档…" : "拖入文档，或点击选择文件"}</div>
        <div className="muted small">支持 .docx · .md · .html · .pdf（PDF 只读，用于审阅）</div>
      </div>
      {err && <div className="notice error">{err}</div>}

      <section className="recent">
        <h2>最近的会话</h2>
        {sessions === null ? (
          <div className="muted">加载中…</div>
        ) : sessions.length === 0 ? (
          <div className="muted">还没有会话。上传一份文档开始。</div>
        ) : (
          <ul>
            {sessions.map((s) => (
              <li key={s.id}>
                <button className="recent-main" onClick={() => onOpen(s.id)}>
                  <span className="badge">{FORMAT_LABEL[s.format]}</span>
                  <span className="recent-name">{s.filename}</span>
                  <span className="muted small">v{s.currentVersion} · {new Date(s.updatedAt).toLocaleString("zh-CN", { hour12: false })}</span>
                </button>
                <button
                  className="btn small ghost"
                  onClick={async () => {
                    if (!confirm(`删除会话“${s.filename}”？所有版本都会被删除。`)) return;
                    await api.remove(s.id).catch((e) => setErr(e.message));
                    load();
                  }}
                >
                  删除
                </button>
              </li>
            ))}
          </ul>
        )}
      </section>

      {runtime && (
        <footer className="home-foot muted small">
          模型 {runtime.model}（思考模式 {runtime.thinking}）· 技能 {runtime.skills.length} · 子代理 {runtime.agents.length}
        </footer>
      )}
    </main>
  );
}
