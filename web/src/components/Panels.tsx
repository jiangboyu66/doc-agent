import type { DocumentSummary, SessionMeta } from "../types";
import { api } from "../api";

const time = (t: number) => new Date(t).toLocaleString("zh-CN", { hour12: false, month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });

export function VersionsPanel({ meta, viewing, onView, onRollback, busy }: { meta: SessionMeta; viewing: number; onView: (v: number) => void; onRollback: (v: number) => void; busy: boolean }) {
  const list = [...meta.versions].reverse();
  return (
    <div className="panel-list">
      {list.map((v) => (
        <div key={v.v} className={`version ${v.v === viewing ? "on" : ""}`}>
          <button className="version-main" onClick={() => onView(v.v)}>
            <span className="vtag">v{v.v}</span>
            <span className="vlabel">{v.label}</span>
            <span className="muted small">{time(v.at)}{v.changedRefs.length ? ` · ${v.changedRefs.length} 处` : ""}</span>
          </button>
          <div className="version-actions">
            {v.v === meta.currentVersion ? (
              <span className="badge green">当前</span>
            ) : (
              <button className="btn small ghost" disabled={busy} onClick={() => onRollback(v.v)} title="回滚本身会生成一个新版本，历史不会丢失">回滚到此</button>
            )}
            <a className="btn small ghost" href={api.documentUrl(meta.id, { version: v.v, download: true })} download>下载</a>
          </div>
        </div>
      ))}
    </div>
  );
}

export function OutlinePanel({ summary, onAsk }: { summary: DocumentSummary; onAsk: (text: string) => void }) {
  return (
    <div className="panel-list">
      <div className="facts">
        <div><b>{summary.blockCount}</b><span>段落</span></div>
        <div><b>{summary.charCount.toLocaleString()}</b><span>字符</span></div>
        {summary.format === "docx" && <div><b>{summary.images}</b><span>图片</span></div>}
        {summary.format === "docx" && <div><b>{summary.trackedChanges}</b><span>修订</span></div>}
        {summary.format === "docx" && <div><b>{summary.comments}</b><span>批注</span></div>}
      </div>
      {summary.template && (
        <div className="notice info">检测到模板：{summary.template.name}。Agent 会自动加载相应的写作与排版规范。</div>
      )}
      {summary.notes.map((n, i) => <div key={i} className="notice warn">{n}</div>)}
      {summary.headings.length === 0 && <div className="muted small pad">未识别到标题</div>}
      <ul className="outline">
        {summary.headings.map((h) => (
          <li key={h.ref} style={{ paddingLeft: (h.level - 1) * 14 }}>
            <button onClick={() => onAsk(`请阅读“${h.text}”这一节（${h.ref}），指出可以改进的地方`)} title="让 Agent 审阅这一节">
              {h.text || "（空标题）"}
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
}
