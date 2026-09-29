import { diffChars, diffWordsWithSpace, type Change } from "diff";
import { useMemo } from "react";
import type { PreviewItem } from "../types";

const CJK = /[　-鿿＀-￯]/g;

/** 中文按字、西文按词比较，避免整句被标红 */
function diff(a: string, b: string): Change[] {
  const cjk = ((a + b).match(CJK) ?? []).length;
  return cjk > (a.length + b.length) * 0.2 ? diffChars(a, b) : diffWordsWithSpace(a, b);
}

/** 较长的未改动片段折叠成省略号，只展示修改附近的上下文 */
function condense(parts: Change[], ctx = 40): Array<Change | { skipped: number }> {
  return parts.map((p, i) => {
    if (p.added || p.removed || p.value.length <= ctx * 2 + 10) return p;
    const head = i === 0 ? "" : p.value.slice(0, ctx);
    const tail = i === parts.length - 1 ? "" : p.value.slice(-ctx);
    const skipped = p.value.length - head.length - tail.length;
    return [{ ...p, value: head }, { skipped }, { ...p, value: tail }] as any;
  }).flat();
}

export function InlineDiff({ before, after }: { before: string; after: string }) {
  const parts = useMemo(() => condense(diff(before, after)), [before, after]);
  if (!before) return <span className="diff-add">{after}</span>;
  if (!after) return <span className="diff-del">{before}</span>;
  return (
    <span>
      {parts.map((p, i) =>
        "skipped" in p ? (
          <span key={i} className="diff-skip">…</span>
        ) : p.added ? (
          <ins key={i} className="diff-add">{p.value}</ins>
        ) : p.removed ? (
          <del key={i} className="diff-del">{p.value}</del>
        ) : (
          <span key={i}>{p.value}</span>
        )
      )}
    </span>
  );
}

export function PreviewList({ items, max = 8 }: { items: PreviewItem[]; max?: number }) {
  const shown = items.slice(0, max);
  return (
    <div className="preview-list">
      {shown.map((p, i) => (
        <div key={i} className="preview-row">
          <span className="ref-chip">{p.ref}</span>
          <div className="preview-text">
            <InlineDiff before={p.before} after={p.after} />
          </div>
        </div>
      ))}
      {items.length > max && <div className="muted small">另有 {items.length - max} 处修改未显示</div>}
    </div>
  );
}
