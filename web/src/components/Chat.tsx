import { useEffect, useMemo, useRef, useState } from "react";
import type { ChatItem, RuntimeInfo, TodoItem } from "../types";
import { Markdown } from "./Markdown";
import { PermissionCard } from "./PermissionCard";
import { PreviewList } from "./DiffView";

type Tool = Extract<ChatItem, { kind: "tool" }>;

function ToolCard({ item, hidePreview }: { item: Tool; hidePreview?: boolean }) {
  const [open, setOpen] = useState(false);
  const firstLine = item.content.split("\n")[0];
  return (
    <div className={`tool-card ${item.status}`}>
      <button className="tool-head" onClick={() => setOpen(!open)} aria-expanded={open}>
        <span className={`dot ${item.status}`} aria-hidden />
        <span className="tool-title">{item.title}</span>
        {item.status !== "running" && firstLine && <span className="tool-summary">{firstLine.replace(/^错误：/, "")}</span>}
        <span className="chev" aria-hidden>{open ? "▾" : "▸"}</span>
      </button>
      {item.status === "ok" && item.preview?.length && !open && !hidePreview ? <PreviewList items={item.preview} max={3} /> : null}
      {item.data?.url && (
        <a className="btn small primary download" href={item.data.url} download>
          下载 {item.data.name}
        </a>
      )}
      {open && <pre className="tool-body">{item.content || "（执行中…）"}</pre>}
    </div>
  );
}

function SubagentBlock({ item, children }: { item: Extract<ChatItem, { kind: "subagent" }>; children: ChatItem[] }) {
  const [open, setOpen] = useState(false);
  return (
    <div className={`subagent ${item.status}`}>
      <button className="subagent-head" onClick={() => setOpen(!open)}>
        <span className={`dot ${item.status === "running" ? "running" : "ok"}`} aria-hidden />
        <span className="badge">{item.agentType}</span>
        <span>{item.description}</span>
        <span className="muted small">{children.length} 步</span>
        <span className="chev" aria-hidden>{open ? "▾" : "▸"}</span>
      </button>
      {open && (
        <div className="subagent-body">
          {children.map((c) => (c.kind === "tool" ? <ToolCard key={c.key} item={c} /> : null))}
        </div>
      )}
    </div>
  );
}

function Todos({ todos }: { todos: TodoItem[] }) {
  if (!todos.length) return null;
  const done = todos.filter((t) => t.status === "completed").length;
  return (
    <details className="todos" open>
      <summary>
        待办 <span className="muted">{done}/{todos.length}</span>
      </summary>
      <ul>
        {todos.map((t, i) => (
          <li key={i} className={t.status}>
            <span aria-hidden>{t.status === "completed" ? "✓" : t.status === "in_progress" ? "▶" : "○"}</span>
            {t.status === "in_progress" ? t.activeForm : t.content}
          </li>
        ))}
      </ul>
    </details>
  );
}

export function Chat(props: {
  items: ChatItem[];
  busy: boolean;
  todos: TodoItem[];
  runtime: RuntimeInfo | null;
  onSend: (text: string) => void;
  onInterrupt: () => void;
  onDecide: (requestId: string, b: { decision: "allow" | "deny"; remember?: "session" | "project"; feedback?: string }) => Promise<void>;
}) {
  const { items, busy, onSend } = props;
  const [text, setText] = useState("");
  const [sel, setSel] = useState(0);
  const scroller = useRef<HTMLDivElement>(null);
  const stick = useRef(true);

  useEffect(() => {
    const el = scroller.current;
    if (el && stick.current) el.scrollTop = el.scrollHeight;
  }, [items]);

  // 子代理的工具调用收进子代理卡片里
  const grouped = useMemo(() => {
    const out: Array<{ item: ChatItem; children?: ChatItem[] }> = [];
    const bySub = new Map<string, ChatItem[]>();
    for (const it of items) {
      if (it.kind === "subagent") {
        const children: ChatItem[] = [];
        bySub.set(it.agentId, children);
        out.push({ item: it, children });
      } else if (it.kind === "tool" && it.agentId !== "main" && bySub.has(it.agentId)) {
        bySub.get(it.agentId)!.push(it);
      } else out.push({ item: it });
    }
    return out;
  }, [items]);

  const suggestions = useMemo(() => {
    const m = /^\/(\S*)$/.exec(text);
    if (!m || !props.runtime) return [];
    return props.runtime.commands.filter((c) => c.name.startsWith(m[1])).slice(0, 8);
  }, [text, props.runtime]);

  const submit = () => {
    const t = text.trim();
    if (!t || busy) return;
    onSend(t);
    setText("");
    stick.current = true;
  };

  const pending = items.some((i) => i.kind === "permission" && i.status === "pending");
  // 已经在确认卡片里展示过预览的修改，工具卡片里不再重复
  const confirmed = useMemo(() => new Set(items.flatMap((i) => (i.kind === "permission" ? [i.request.toolUseId] : []))), [items]);

  return (
    <section className="chat" aria-label="对话">
      <div
        className="chat-scroll"
        ref={scroller}
        onScroll={(e) => {
          const el = e.currentTarget;
          stick.current = el.scrollHeight - el.scrollTop - el.clientHeight < 80;
        }}
      >
        {!items.length && (
          <div className="chat-empty">
            <p className="serif">告诉我你想怎么改这份文档。</p>
            <div className="chips">
              {["通读全文，列出语病与错别字", "把摘要润色得更学术", "检查格式是否统一", "/review"].map((s) => (
                <button key={s} className="chip" onClick={() => onSend(s)}>{s}</button>
              ))}
            </div>
            <p className="muted small">所有修改都直接作用在原文件上：未改动的内容逐字节保持不变，每次修改都会生成可回滚的版本。</p>
          </div>
        )}
        {grouped.map(({ item, children }) => {
          switch (item.kind) {
            case "user":
              return <div key={item.key} className="msg user"><div className="bubble">{item.text}</div></div>;
            case "assistant":
              return (
                <div key={item.key} className="msg assistant">
                  {item.reasoning && (
                    <details className="reasoning">
                      <summary>{item.streaming && !item.text ? "思考中…" : "思考过程"}</summary>
                      <div>{item.reasoning}</div>
                    </details>
                  )}
                  {item.text && <Markdown text={item.text} />}
                  {item.streaming && <span className="caret" aria-hidden />}
                </div>
              );
            case "tool":
              return <ToolCard key={item.key} item={item} hidePreview={confirmed.has(item.id)} />;
            case "subagent":
              return <SubagentBlock key={item.key} item={item} children={children ?? []} />;
            case "permission":
              return (
                <PermissionCard key={item.key} request={item.request} status={item.status} onDecide={(b) => props.onDecide(item.request.requestId, b)} />
              );
            case "notice":
              return <div key={item.key} className={`notice ${item.tone}`}>{item.text}</div>;
            case "command":
              return <pre key={item.key} className="command-output">{item.text}</pre>;
          }
        })}
        {busy && !pending && <div className="working"><span className="spinner" aria-hidden /> Agent 正在处理…</div>}
      </div>

      <Todos todos={props.todos} />

      <div className="composer">
        {suggestions.length > 0 && (
          <ul className="suggest" role="listbox">
            {suggestions.map((c, i) => (
              <li
                key={c.name}
                role="option"
                aria-selected={i === sel}
                className={i === sel ? "active" : ""}
                onMouseDown={(e) => {
                  e.preventDefault();
                  setText(`/${c.name} `);
                }}
              >
                <b>/{c.name}</b> {c.argumentHint && <span className="muted">{c.argumentHint}</span>} <span className="muted">— {c.description}</span>
              </li>
            ))}
          </ul>
        )}
        <textarea
          rows={3}
          value={text}
          placeholder={busy ? "Agent 正在处理，可点击“中断”…" : "输入修改要求，Enter 发送，Shift+Enter 换行；输入 / 查看命令"}
          onChange={(e) => {
            setText(e.target.value);
            setSel(0);
          }}
          onKeyDown={(e) => {
            if (suggestions.length) {
              if (e.key === "ArrowDown") { e.preventDefault(); setSel((sel + 1) % suggestions.length); return; }
              if (e.key === "ArrowUp") { e.preventDefault(); setSel((sel - 1 + suggestions.length) % suggestions.length); return; }
              if (e.key === "Tab" || (e.key === "Enter" && !e.shiftKey && text !== `/${suggestions[sel].name}`)) {
                e.preventDefault();
                setText(`/${suggestions[sel].name} `);
                return;
              }
            }
            if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
              e.preventDefault();
              submit();
            }
          }}
        />
        <div className="composer-actions">
          {busy ? (
            <button className="btn danger" onClick={props.onInterrupt}>中断</button>
          ) : (
            <button className="btn primary" disabled={!text.trim()} onClick={submit}>发送</button>
          )}
        </div>
      </div>
    </section>
  );
}
