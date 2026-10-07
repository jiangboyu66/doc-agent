import { useCallback, useEffect, useRef, useState } from "react";
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
  const [pageInfo, setPageInfo] = useState<{ current: number; total: number } | null>(null);
  const onPages = useCallback((p: { current: number; total: number }) => setPageInfo((old) => (old && old.current === p.current && old.total === p.total ? old : p)), []);
  const url = api.documentUrl(sessionId, { version });
  useEffect(() => setPageInfo(null), [url]);

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
        {format === "docx" && view === "fast" && pageInfo && (
          <span className="page-indicator" title="当前页 / 总页数">第 {pageInfo.current} / {pageInfo.total} 页</span>
        )}
        {format === "docx" && view === "fast" && (
          <div className="seg right">
            <button onClick={() => setZoom((z) => Math.max(0.4, (z === "fit" ? 1 : z) - 0.1))} aria-label="缩小">−</button>
            <button className={zoom === "fit" ? "on" : ""} onClick={() => setZoom("fit")}>适应</button>
            <button onClick={() => setZoom((z) => Math.min(2, (z === "fit" ? 1 : z) + 0.1))} aria-label="放大">＋</button>
          </div>
        )}
      </div>
      <div className="preview-body">
        {format === "docx" && view === "fast" && <DocxView url={url} zoom={zoom} onPages={onPages} />}
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

function DocxView({ url, zoom, onPages }: { url: string; zoom: number | "fit"; onPages?: (p: { current: number; total: number }) => void }) {
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
        const [{ renderAsync }, { paginate, prepareDocx, numberPages, fixAtLeastSpacing }, res] = await Promise.all([import("docx-preview"), import("../paginate"), fetch(url)]);
        if (!res.ok) throw new Error(`加载失败（${res.status}）`);
        const blob = await res.blob();
        if (!alive || !host.current) return;
        const opts = {
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
        } as any;
        // 在屏幕外渲染两份：正文（去掉文档记录的旧分页位置，由 paginate 按 Word 规则重新分页）
        // 与原样渲染（取各页的页眉页脚；分页失败时直接使用）
        const stage = document.createElement("div");
        stage.style.cssText = "position:absolute;left:-100000px;top:0;visibility:hidden;pointer-events:none;";
        const plain = stage.cloneNode() as HTMLDivElement;
        document.body.append(stage, plain);
        try {
          const prep = await prepareDocx(blob).catch(() => ({ blob, plain: blob, balance: true, pageStart: 1 }));
          await renderAsync(prep.blob, stage, undefined, opts);
          await renderAsync(prep.plain, plain, undefined, opts);
          fixAtLeastSpacing(stage);
          fixAtLeastSpacing(plain);
          let out: HTMLElement = stage;
          try {
            await (document as any).fonts?.ready;
            paginate(stage, plain, { balance: prep.balance });
          } catch (e) {
            console.warn("快速预览分页失败，使用原始分页", e);
            out = plain;
          }
          if (!alive || !host.current) return;
          host.current.innerHTML = "";
          host.current.append(...Array.from(out.childNodes));
          // 页眉页脚填入实际页码，每页下方标注"第 n 页 / 共 N 页"
          const total = numberPages(host.current, prep.pageStart);
          onPages?.({ current: 1, total });
        } finally {
          stage.remove();
          plain.remove();
        }
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

  // 滚动时报告当前页（视口上三分之一处所在的页）
  useEffect(() => {
    if (state !== "ok" || !onPages) return;
    const scroller = outer.current?.parentElement;
    if (!scroller) return;
    let raf = 0;
    const update = () => {
      raf = 0;
      const pages = Array.from(host.current?.querySelectorAll<HTMLElement>("section.docx[data-page]") ?? []);
      if (!pages.length) return;
      const box = scroller.getBoundingClientRect();
      const probe = box.top + box.height / 3;
      let cur = 1;
      for (const p of pages) { if (p.getBoundingClientRect().top <= probe) cur = Number(p.dataset.page); else break; }
      onPages({ current: cur, total: pages.length });
    };
    const onScroll = () => { if (!raf) raf = requestAnimationFrame(update); };
    scroller.addEventListener("scroll", onScroll, { passive: true });
    update();
    return () => { scroller.removeEventListener("scroll", onScroll); if (raf) cancelAnimationFrame(raf); };
  }, [state, scale, onPages]);

  return (
    <div className="docx-outer" ref={outer}>
      {state === "loading" && <div className="loading"><span className="spinner" aria-hidden /> 正在渲染文档…</div>}
      {state === "error" && <div className="notice error">预览失败：{err}。可切换到“精确版式”或下载后用 Word 打开。</div>}
      <div className="docx-host" ref={host} style={{ zoom: scale } as any} />
    </div>
  );
}
