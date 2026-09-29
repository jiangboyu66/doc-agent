import { useState } from "react";
import type { PermissionRequest } from "../types";
import { PreviewList } from "./DiffView";
import { Markdown } from "./Markdown";

type Decide = (body: { decision: "allow" | "deny"; remember?: "session" | "project"; feedback?: string }) => Promise<void>;

export function PermissionCard({ request, status, onDecide }: { request: PermissionRequest; status: "pending" | "allow" | "deny"; onDecide: Decide }) {
  const [denying, setDenying] = useState(false);
  const [feedback, setFeedback] = useState("");
  const [busy, setBusy] = useState(false);
  const isPlan = request.tool === "exit_plan_mode";

  const act = async (b: Parameters<Decide>[0]) => {
    setBusy(true);
    try {
      await onDecide(b);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={`perm-card ${status}`}>
      <div className="perm-head">
        <span className="perm-icon" aria-hidden>{isPlan ? "☰" : "✎"}</span>
        <div>
          <div className="perm-title">{isPlan ? "修改计划待确认" : request.title}</div>
          <div className="muted small">
            {request.reason}
            {request.agentId !== "main" && <> · 来自子代理 {request.agentId}</>}
          </div>
        </div>
      </div>

      {isPlan && typeof request.input?.plan === "string" && <Markdown className="perm-plan" text={request.input.plan} />}
      {!!request.preview?.length && <PreviewList items={request.preview} />}
      {!isPlan && !request.preview?.length && (
        <pre className="perm-input">{JSON.stringify(request.input, null, 2)}</pre>
      )}

      {status === "pending" ? (
        denying ? (
          <div className="perm-deny">
            <textarea
              autoFocus
              rows={2}
              placeholder="告诉 Agent 你希望怎么改（可留空）"
              value={feedback}
              onChange={(e) => setFeedback(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) act({ decision: "deny", feedback: feedback.trim() || undefined });
              }}
            />
            <div className="row gap">
              <button className="btn danger" disabled={busy} onClick={() => act({ decision: "deny", feedback: feedback.trim() || undefined })}>拒绝并反馈</button>
              <button className="btn ghost" disabled={busy} onClick={() => setDenying(false)}>返回</button>
            </div>
          </div>
        ) : (
          <div className="perm-actions">
            <button className="btn primary" disabled={busy} onClick={() => act({ decision: "allow" })}>{isPlan ? "批准计划" : "允许"}</button>
            {!isPlan && (
              <>
                <button className="btn" disabled={busy} onClick={() => act({ decision: "allow", remember: "session" })} title={`本会话内自动允许：${request.suggestedRule}`}>本会话都允许</button>
                <button className="btn" disabled={busy} onClick={() => act({ decision: "allow", remember: "project" })} title={`写入项目配置 .doc-agent/settings.json：${request.suggestedRule}`}>以后都允许</button>
              </>
            )}
            <button className="btn ghost" disabled={busy} onClick={() => setDenying(true)}>{isPlan ? "修改计划" : "拒绝"}</button>
          </div>
        )
      ) : (
        <div className={`perm-status ${status}`}>{status === "allow" ? "已允许" : "已拒绝"}</div>
      )}
    </div>
  );
}
