import { useEffect, useRef, useState } from "react";
import { api } from "../api";
import type { DocFormat } from "../types";
import { Markdown } from "./Markdown";

type View = "fast" | "layout" | "source";

/**
 * 文档预览
 * - Word：默认用 docx-preview 在浏览器里渲染（快，显示修订和批注）；"精确版式"用 LibreOffice 渲染成 PDF，
 *   分页、字体度量与 Word 更接近。
 * - Markdown：渲染 / 源码；HTML：沙箱 iframe（禁用脚本）；PDF：浏览器内置阅读器。
 */
export function DocPreview({ sessionId, format, version, isCurrent, sofficeAvailable }: { sessionId: string; format: DocFormat; version: number; isCurrent: boolean; sofficeAvailable: boolean }) {
  const [view, setView] = useState<View>("fast");
  const [zoom, setZoom] = useState<number | "fit">("fit");
  const url = api.documentUrl(sessionId, { version });

  const views: Array<[View, string, string]> =
    format === "docx"
      ? [["fast", "快速预览", "浏览器内渲染，显示修订与批注"], ["layout", "精确版式", sofficeAvailable ? "LibreOffice 渲染的 PDF，分页与 Word 接近" : "需要安装 LibreOffice"]]
      : format === "markdown"
        ? [["fast", "渲染", ""], ["source", "源码", ""]]
        : format === "html"
          ? [["fast", "页面", ""], ["source", "源码", ""]]
          : [];

  return (
    <div className="preview">
      <div className="preview-bar">
        <div className="seg">
          {views.map(([v, label, tip]) => (
            <button key={v} className={view === v ? "on" : ""} title={tip} disabled={v === "layout" && !sofficeAvailable} onClick={() => setView(v)}>{label}</button>
          ))}
        </div>
        <span className="muted small">{isCurrent ? `当前版本 v${version}` : `正在查看历史版本 v${version}`}</span>
        {format === "docx" && view === "fast" && (
          <div className="seg right">
            <button onClick={() => setZoom((z) => Math.max(0.4, (z === "fit" ? 1 : z) - 0.1))} aria-label="缩小">−</button>
            <button className={zoom === "fit" ? "on" : ""} onClick={() => setZoom("fit")}>适应</button>
            <button onClick={() => setZoom((z) => Math.min(2, (z === "fit" ? 1 : z) + 0.1))} aria-label="放大">＋</button>
          </div>
        )}
      </div>
      <div className="preview-body">
        {format === "docx" && view === "fast" && <DocxView url={url} zoom={zoom} />}
        {format === "docx" && view === "layout" && <iframe key={version} className="frame" title="版式预览" src={api.previewPdfUrl(sessionId, version)} />}
        {format === "pdf" && <iframe className="frame" title="PDF" src={url} />}
        {format === "html" && view === "fast" && <iframe key={url} className="frame white" title="HTML 预览" sandbox="" src={url} />}
        {(format === "markdown" || format === "html") && view !== "fast" && <SourceView url={url} />}
        {format === "markdown" && view === "fast" && <MarkdownView url={url} />}
      </div>
    </div>
  );
}

function useText(url: string) {
  const [text, setText] = useState<string | null>(null);
  const [err, setErr] = useState("");
  useEffect(() => {
    let alive = true;
    setErr("");
    fetch(url).then((r) => (r.ok ? r.text() : Promise.reject(new Error(`加载失败（${r.status}）`)))).then((t) => alive && setText(t), (e) => alive && setErr(e.message));
    return () => { alive = false; };
  }, [url]);
  return { text, err };
}

function SourceView({ url }: { url: string }) {
  const { text, err } = useText(url);
  if (err) return <div className="notice error">{err}</div>;
  return <pre className="source">{text ?? "加载中…"}</pre>;
}

function MarkdownView({ url }: { url: string }) {
  const { text, err } = useText(url);
  if (err) return <div className="notice error">{err}</div>;
  const body = (text ?? "").replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, "");
  return <div className="paper"><Markdown className="doc-md" text={body} /></div>;
}

function DocxView({ url, zoom }: { url: string; zoom: number | "fit" }) {
  const host = useRef<HTMLDivElement>(null);
  const outer = useRef<HTMLDivElement>(null);
  const [state, setState] = useState<"loading" | "ok" | "error">("loading");
  const [err, setErr] = useState("");
  const [scale, setScale] = useState(1);

  useEffect(() => {
    let alive = true;
    setState("loading");
    (async () => {
      try {
        const [{ renderAsync }, res] = await Promise.all([import("docx-preview"), fetch(url)]);
        if (!res.ok) throw new Error(`加载失败（${res.status}）`);
        const blob = await res.blob();
        if (!alive || !host.current) return;
        host.current.innerHTML = "";
        await renderAsync(blob, host.current, undefined, {
          className: "docx",
          inWrapper: true,
          breakPages: true,
          ignoreLastRenderedPageBreak: false,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
          renderComments: true,
          renderChanges: true,
          experimental: true,
          useBase64URL: true,
        } as any);
        if (alive) setState("ok");
      } catch (e: any) {
        if (alive) {
          setErr(e?.message ?? String(e));
          setState("error");
        }
      }
    })();
    return () => { alive = false; };
  }, [url]);

  // 适应宽度：按最宽页面与容器宽度计算缩放
  useEffect(() => {
    if (zoom !== "fit") { setScale(zoom); return; }
    const el = outer.current;
    if (!el) return;
    const fit = () => {
      const pages = host.current?.querySelectorAll<HTMLElement>("section.docx");
      const w = Math.max(0, ...Array.from(pages ?? []).map((p) => p.offsetWidth));
      if (w) setScale(Math.min(1.25, (el.clientWidth - 32) / w));
    };
    fit();
    const ro = new ResizeObserver(fit);
    ro.observe(el);
    return () => ro.disconnect();
  }, [zoom, state]);

  return (
    <div className="docx-outer" ref={outer}>
      {state === "loading" && <div className="loading"><span className="spinner" aria-hidden /> 正在渲染文档…</div>}
      {state === "error" && <div className="notice error">预览失败：{err}。可切换到“精确版式”或下载后用 Word 打开。</div>}
      <div className="docx-host" ref={host} style={{ zoom: scale } as any} />
    </div>
  );
}
