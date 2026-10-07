import { useState } from "react";
import { api } from "../api";
import type { RuntimeInfo } from "../types";

/** 触发浏览器下载（服务端已设置 Content-Disposition: attachment） */
function triggerDownload(url: string, name: string) {
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
}

const COLLAPSE_KEY = "docAgent.convertCard.collapsed";
/** 记住的折叠状态；没有记录时，小屏幕（窗口较矮或较窄）默认收起，把空间留给文档预览 */
const readCollapsed = () => {
  try {
    const v = localStorage.getItem(COLLAPSE_KEY);
    if (v === "1" || v === "0") return v === "1";
  } catch { /* 隐私模式等 */ }
  return typeof window !== "undefined" && (window.innerHeight < 860 || window.innerWidth < 960);
};

/**
 * PDF 会话顶部：PDF → Word（可折叠，折叠状态记在浏览器里）
 *  - 下载 Word 文件：直接得到 .docx（不改变当前会话）
 *  - 转换并继续编辑：在当前会话内转换为 Word，对话保留，之后可以直接让 Agent 修改
 */
export function ConvertCard({ sessionId, runtime, onDone }: { sessionId: string; runtime: RuntimeInfo | null; onDone: (text: string) => void }) {
  const [collapsed, setCollapsedState] = useState(readCollapsed);
  // 只记住用户主动的选择
  const setCollapsed = (v: boolean) => {
    setCollapsedState(v);
    try { localStorage.setItem(COLLAPSE_KEY, v ? "1" : "0"); } catch { /* 隐私模式等 */ }
  };
  // 自动：下载 Word 时逐页保留原排版（外观与 PDF 一致），转换后继续编辑时用可编辑的流式排版（扩写改写不乱版）
  const [mode, setMode] = useState<"auto" | "exact" | "flow">("auto");
  const [engine, setEngine] = useState("builtin");
  const [busy, setBusy] = useState<"" | "download" | "edit">("");
  const [err, setErr] = useState("");
  const [done, setDone] = useState<{ name: string; url: string; note: string } | null>(null);
  const engines = runtime?.pdfEngines ?? { builtin: true, pdf2docx: false, libreoffice: false };

  const download = async () => {
    setBusy("download");
    setErr("");
    setDone(null);
    try {
      const r = await api.exportAs(sessionId, "docx", { mode: mode === "flow" ? "flow" : "exact", engine });
      setDone(r);
      triggerDownload(r.url, r.name);
    } catch (e: any) {
      setErr(`导出失败：${e.message}`);
    } finally {
      setBusy("");
    }
  };

  const edit = async () => {
    setBusy("edit");
    setErr("");
    try {
      const r = await api.convert(sessionId, { engine, mode: mode === "exact" ? "exact" : "flow", inPlace: true });
      onDone(r.text);
    } catch (e: any) {
      setErr(`转换失败：${e.message}`);
      setBusy("");
    }
  };

  const actions = (
    <div className="convert-actions">
      <button className={`btn primary ${collapsed ? "small" : ""}`} onClick={download} disabled={!!busy}>{busy === "download" ? "正在生成…" : "下载 Word 文件"}</button>
      <button className={`btn ${collapsed ? "small" : ""}`} onClick={edit} disabled={!!busy} title="在当前会话内转换为 Word，对话保留，可以直接让 Agent 修改">{busy === "edit" ? "正在转换…" : "转换并继续编辑"}</button>
    </div>
  );

  if (collapsed) {
    return (
      <div className="convert-card collapsed">
        <button className="convert-toggle" onClick={() => setCollapsed(false)} aria-expanded={false} title="展开 PDF → Word 选项">
          <span className="chev" aria-hidden>▸</span>
          <span className="convert-title">PDF → Word</span>
          <span className="muted small">{mode === "auto" ? "自动选择排版" : mode === "exact" ? "逐页保留原排版" : "可编辑排版"}</span>
        </button>
        {actions}
        {err && <div className="notice error">{err}</div>}
      </div>
    );
  }

  return (
    <div className="convert-card">
      <div className="convert-head">
        <button className="convert-toggle" onClick={() => setCollapsed(true)} aria-expanded title="收起，让文档预览占满空间">
          <span className="chev" aria-hidden>▾</span>
          <span>
            <span className="convert-title">PDF → Word</span>
            <span className="muted small convert-desc">把这份 PDF 导出为可编辑的 Word 文件（.docx）。原 PDF 保持不变，完成后会报告文字完整性。</span>
          </span>
        </button>
        {actions}
      </div>
      <div className="convert-options" role="radiogroup" aria-label="排版方式">
        <label className={`opt ${mode === "auto" ? "on" : ""}`}>
          <input type="radio" name="mode" checked={mode === "auto"} onChange={() => setMode("auto")} />
          <span>
            <b>自动（推荐）</b>
            <span className="muted small">下载 Word 文件时逐页保留原排版；转换并继续编辑时用可编辑排版，扩写、改写后版面自动重排。</span>
          </span>
        </label>
        <label className={`opt ${mode === "exact" ? "on" : ""}`}>
          <input type="radio" name="mode" checked={mode === "exact"} onChange={() => setMode("exact")} />
          <span>
            <b>逐页保留原排版</b>
            <span className="muted small">每一行、每一页都与 PDF 一致（每行硬换行、逐页固定）。只适合改个别字词，增加文字会挤乱版面。</span>
          </span>
        </label>
        <label className={`opt ${mode === "flow" ? "on" : ""}`}>
          <input type="radio" name="mode" checked={mode === "flow"} onChange={() => setMode("flow")} />
          <span>
            <b>可编辑排版</b>
            <span className="muted small">字体、字号、分栏、页眉页脚、公式与图片都与原文一致；段落自动换行、跨栏跨页连续，扩写改写不乱版。</span>
          </span>
        </label>
      </div>
      {(engines.pdf2docx || engines.libreoffice) && (
        <label className="ctl engine">
          <span>转换引擎</span>
          <select value={engine} onChange={(e) => setEngine(e.target.value)}>
            <option value="builtin">内置引擎（推荐）</option>
            {engines.pdf2docx && <option value="pdf2docx">pdf2docx（Python）</option>}
            {engines.libreoffice && <option value="libreoffice">LibreOffice（文本框式）</option>}
          </select>
        </label>
      )}
      {done && (
        <div className="notice success">
          已生成 <a href={done.url} download={done.name}>{done.name}</a>（没有自动开始下载时请点击文件名）。{"\n"}
          {done.note}
        </div>
      )}
      {err && <div className="notice error">{err}</div>}
    </div>
  );
}
