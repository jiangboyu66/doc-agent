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

/**
 * 思考过程：受控的展开/折叠（思考进行中自动展开，答复开始后自动收起；用户手动点过之后以用户为准）
 */
function Reasoning({ text, streaming, answering }: { text: string; streaming: boolean; answering: boolean }) {
  const [userOpen, setUserOpen] = useState<boolean | null>(null);
  const thinking = streaming && !answering;
  const open = userOpen ?? thinking;
  const body = text.trim();
  if (!body) return null;
  return (
    <div className={`reasoning ${open ? "open" : ""}`}>
      <button type="button" className="reasoning-head" aria-expanded={open} onClick={() => setUserOpen(!open)}>
        <span className="chev" aria-hidden>{open ? "▾" : "▸"}</span>
        {thinking ? <span>思考中…</span> : <span>思考过程</span>}
        <span className="muted small">{body.length} 字</span>
      </button>
      {open && <div className="reasoning-body">{body}</div>}
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
  /** 思考模式：当前状态与切换（不可用时为 undefined，不显示按钮） */
  thinking?: boolean;
  onToggleThinking?: () => void;
  /** 上传图片等素材（供 Agent 插入文档） */
  onAttach?: (files: File[]) => void;
  attaching?: boolean;
  format?: string;
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
              {(props.format === "pdf"
                ? ["转换成 Word 后帮我润色摘要", "总结这篇文档的要点", "检查参考文献格式", "/review"]
                : props.format === "docx"
                  ? ["通读全文，列出语病与错别字", "把摘要润色得更学术", "统一正文格式：宋体小四、1.5 倍行距、首行缩进 2 字", "在页脚加上「第 X 页 共 Y 页」", "在标题后插入目录", "/review"]
                  : ["通读全文，列出语病与错别字", "把摘要润色得更学术", "检查格式是否统一", "/review"]
              ).map((s) => (
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
                  <Reasoning text={item.reasoning} streaming={item.streaming} answering={!!item.text} />
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
        <div className="composer-tools">
          {props.onToggleThinking && (
            <button
              type="button"
              className={`pill ${props.thinking ? "on" : ""}`}
              aria-pressed={!!props.thinking}
              title={props.thinking ? "已开启：模型先思考再回答，复杂修改更可靠，但更慢、消耗更多 tokens" : "已关闭：直接回答，更快更省"}
              onClick={props.onToggleThinking}
            >
              <span aria-hidden>{props.thinking ? "●" : "○"}</span> 深度思考
            </button>
          )}
          {props.onAttach && (
            <label className={`pill ${props.attaching ? "busy" : ""}`} title="上传图片（PNG/JPG/GIF/SVG），之后可以让 Agent 把它插入文档">
              <input
                type="file"
                accept="image/png,image/jpeg,image/gif,image/svg+xml,image/bmp,image/webp"
                multiple
                hidden
                onChange={(e) => {
                  const files = [...(e.target.files ?? [])];
                  e.target.value = "";
                  if (files.length) props.onAttach!(files);
                }}
              />
              <span aria-hidden>＋</span> {props.attaching ? "上传中…" : "图片"}
            </label>
          )}
        </div>
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
