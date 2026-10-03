import type { EngineEvent, RuntimeInfo, SessionDetail, SessionMeta, DocumentSummary } from "./types";

async function json<T>(res: Response): Promise<T> {
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error((body as any).error || `请求失败（${res.status}）`);
  return body as T;
}

export const api = {
  runtime: () => fetch("/api/runtime").then((r) => json<RuntimeInfo>(r)),
  sessions: () => fetch("/api/sessions").then((r) => json<SessionMeta[]>(r)),
  session: (id: string) => fetch(`/api/sessions/${id}`).then((r) => json<SessionDetail>(r)),
  upload(file: File) {
    const fd = new FormData();
    fd.append("file", file);
    return fetch("/api/sessions", { method: "POST", body: fd }).then((r) => json<{ meta: SessionMeta; summary: DocumentSummary }>(r));
  },
  uploadAssets(id: string, files: File[]) {
    const fd = new FormData();
    for (const f of files) fd.append("files", f);
    return fetch(`/api/sessions/${id}/assets`, { method: "POST", body: fd }).then((r) => json<{ assets: string[]; all: string[] }>(r));
  },
  remove: (id: string) => fetch(`/api/sessions/${id}`, { method: "DELETE" }).then((r) => json<{ ok: true }>(r)),
  settings: (id: string, patch: Partial<Pick<SessionMeta, "mode" | "trackChanges" | "author" | "thinking">>) =>
    fetch(`/api/sessions/${id}/settings`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(patch) }).then((r) => json<{ meta: SessionMeta }>(r)),
  permission: (id: string, requestId: string, body: { decision: "allow" | "deny"; remember?: "session" | "project"; feedback?: string }) =>
    fetch(`/api/sessions/${id}/permissions/${requestId}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }).then((r) => json<{ ok: true }>(r)),
  interrupt: (id: string) => fetch(`/api/sessions/${id}/interrupt`, { method: "POST" }).then((r) => json<{ ok: true }>(r)),
  rollback: (id: string, version: number) =>
    fetch(`/api/sessions/${id}/rollback`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ version }) }).then((r) => json<{ meta: SessionMeta; version: number }>(r)),
  exportAs: (id: string, format: string, pdf?: { mode: "exact" | "flow"; engine: string }) =>
    fetch(`/api/sessions/${id}/export`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ format, ...(pdf ?? {}) }) }).then((r) =>
      json<{ name: string; url: string; note: string; lossy: boolean }>(r)
    ),
  convert: (id: string, body: { engine: string; mode: "exact" | "flow"; inPlace?: boolean }) =>
    fetch(`/api/sessions/${id}/convert`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }).then((r) =>
      json<{ meta: SessionMeta; text: string }>(r)
    ),
  documentUrl: (id: string, opts: { version?: number; original?: boolean; download?: boolean } = {}) => {
    const q = new URLSearchParams();
    if (opts.version !== undefined) q.set("version", String(opts.version));
    if (opts.original) q.set("original", "1");
    if (opts.download) q.set("download", "1");
    return `/api/sessions/${id}/document?${q}`;
  },
  previewPdfUrl: (id: string, version: number) => `/api/sessions/${id}/preview.pdf?version=${version}`,
};

/**
 * 发送一条消息并以 SSE 流读取引擎事件。
 * 用 fetch + ReadableStream 而不是 EventSource：EventSource 只支持 GET，而消息需要 POST。
 */
export async function streamMessage(id: string, text: string, onEvent: (e: EngineEvent) => void, signal?: AbortSignal): Promise<void> {
  const res = await fetch(`/api/sessions/${id}/messages`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text }),
    signal,
  });
  if (!res.ok || !res.body) {
    const body = await res.json().catch(() => ({}));
    throw new Error((body as any).error || `请求失败（${res.status}）`);
  }
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buf = "";
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    let idx: number;
    while ((idx = buf.indexOf("\n\n")) !== -1) {
      const frame = buf.slice(0, idx);
      buf = buf.slice(idx + 2);
      const data = frame.split("\n").filter((l) => l.startsWith("data: ")).map((l) => l.slice(6)).join("\n");
      if (!data) continue; // 心跳
      try {
        onEvent(JSON.parse(data));
      } catch {
        /* 忽略损坏的帧 */
      }
    }
  }
}
