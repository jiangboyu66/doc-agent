/**
 * 文本定位与最小差异计算（各格式适配器共用）
 *
 * 1. 定位：模型给出的 old_text 常常和文档字符有细微差别（弯引号 vs 直引号、全角破折号、不间断空格……）。
 *    先精确匹配，找不到再做"归一化匹配"，并把匹配位置映射回原文字符下标。
 * 2. 最小差异：模型给出整句 old/new，我们只修改真正变化的词，未变化的字符保持原来的 run 和格式。
 *    归一化后相等的字符（如原文弯引号、模型写成直引号）视为未变化，保留原文字符——这是保真的关键。
 */

import { diffArrays } from "diff";

const NORM_MAP: Record<string, string> = {
  "‘": "'", "’": "'", "‚": "'", "‛": "'", "′": "'",
  "“": '"', "”": '"', "„": '"', "‟": '"', "″": '"',
  "‐": "-", "‑": "-", "‒": "-", "–": "-", "—": "-", "―": "-", "−": "-",
  " ": " ", " ": " ", " ": " ", "　": " ", "\t": " ",
};

export function normChar(c: string): string {
  return NORM_MAP[c] ?? c;
}

/** 归一化并返回每个归一化字符对应的原文下标 */
function normalizeWithMap(s: string): { text: string; map: number[] } {
  let text = "";
  const map: number[] = [];
  let lastWasSpace = false;
  for (let i = 0; i < s.length; i++) {
    const c = normChar(s[i]);
    const isSpace = c === " " || c === "\n" || c === "\r";
    if (isSpace) {
      if (lastWasSpace) continue;
      text += " ";
      map.push(i);
      lastWasSpace = true;
    } else {
      text += c;
      map.push(i);
      lastWasSpace = false;
    }
  }
  return { text, map };
}

export interface Match {
  start: number;
  end: number;
  fuzzy: boolean;
}

/** 在 haystack 中查找 needle 的全部出现位置（先精确，后归一化） */
export function findAll(haystack: string, needle: string, caseSensitive = true): Match[] {
  if (!needle) return [];
  const exact: Match[] = [];
  const h = caseSensitive ? haystack : haystack.toLowerCase();
  const n = caseSensitive ? needle : needle.toLowerCase();
  let i = 0;
  while ((i = h.indexOf(n, i)) !== -1) {
    exact.push({ start: i, end: i + n.length, fuzzy: false });
    i += n.length;
  }
  if (exact.length) return exact;

  const H = normalizeWithMap(h);
  const N = normalizeWithMap(n).text.trim();
  if (!N) return [];
  const out: Match[] = [];
  let k = 0;
  while ((k = H.text.indexOf(N, k)) !== -1) {
    const start = H.map[k];
    const lastIdx = H.map[k + N.length - 1];
    out.push({ start, end: lastIdx + 1, fuzzy: true });
    k += N.length;
  }
  return out;
}

// ---------------- 分词 ----------------

const TOKEN_RE = /⟨[^⟩]*⟩|[㐀-鿿豈-﫿぀-ヿ가-힯]|[A-Za-z0-9À-ɏͰ-ϿЀ-ӿ]+|\s+|[\s\S]/gu;

export function tokenize(s: string): string[] {
  return s.match(TOKEN_RE) ?? [];
}

function normToken(t: string): string {
  if (/^\s+$/.test(t)) return " ";
  let r = "";
  for (const c of t) r += normChar(c);
  return r;
}

export interface Hunk {
  /** 在旧文本中的字符区间 */
  start: number;
  end: number;
  /** 替换成的新文本 */
  text: string;
}

/**
 * 计算把 oldText 改成 newText 所需的最小修改块。
 * 归一化后相等的 token 视为未修改（保留原文字符）。修改过于零碎时合并成一个整体修改，避免
 * 产生大量碎片 run。
 */
export function computeHunks(oldText: string, newText: string): Hunk[] {
  if (oldText === newText) return [];
  const a = tokenize(oldText);
  const b = tokenize(newText);
  const parts = diffArrays(a, b, { comparator: (x: string, y: string) => normToken(x) === normToken(y) });

  const hunks: Hunk[] = [];
  let pos = 0;
  let ai = 0; // 旧 token 下标：位置必须按旧文本的真实 token 计算（公共部分 jsdiff 返回的是新文本的 token）
  let cur: Hunk | null = null;
  let unchangedTokensInside = 0;
  const oldLen = (n: number) => {
    let len = 0;
    for (let k = 0; k < n; k++) len += a[ai + k].length;
    ai += n;
    return len;
  };
  for (const p of parts) {
    const count = p.count ?? p.value.length;
    if (!p.added && !p.removed) {
      if (cur) {
        hunks.push(cur);
        cur = null;
      }
      pos += oldLen(count);
      unchangedTokensInside += count;
      continue;
    }
    if (!cur) cur = { start: pos, end: pos, text: "" };
    if (p.removed) {
      pos += oldLen(count);
      cur.end = pos;
    } else {
      cur.text += p.value.join("");
    }
  }
  if (cur) hunks.push(cur);

  if (hunks.length > 1) {
    const span = hunks[hunks.length - 1].end - hunks[0].start;
    const changed = hunks.reduce((s, h) => s + (h.end - h.start) + h.text.length, 0);
    // 太零碎：合并为一个从第一个修改到最后一个修改的整体替换
    if (hunks.length > 8 || (span > 0 && changed / (span * 2) > 0.6)) {
      const start = hunks[0].start;
      const end = hunks[hunks.length - 1].end;
      const prefixLen = start;
      const suffixLen = oldText.length - end;
      const text = newText.slice(prefixLen, newText.length - suffixLen);
      // 前后缀在归一化下相等但字符可能不同，因此用原文前后缀长度切分新文本只在长度一致时成立；
      // 不一致时退化为整段替换
      const prefixOk = normToken(oldText.slice(0, start)) === normToken(newText.slice(0, prefixLen));
      const suffixOk = normToken(oldText.slice(end)) === normToken(newText.slice(newText.length - suffixLen));
      if (prefixOk && suffixOk) return [trimHunk(oldText, { start, end, text })];
      return [trimHunk(oldText, { start: 0, end: oldText.length, text: newText })];
    }
  }
  void unchangedTokensInside;
  return hunks.map((h) => trimHunk(oldText, h));
}

/**
 * 对单个修改块再做一次字符级的前后缀裁剪：例如 "Author1" → "Smith1" 只改 "Author"，
 * 末尾那个上标 "1" 所在的 run 完全不动，从而保住它的上标格式。
 */
function trimHunk(oldText: string, h: Hunk): Hunk {
  let { start, end, text } = h;
  while (start < end && text.length && normChar(oldText[start]) === normChar(text[0])) {
    start++;
    text = text.slice(1);
  }
  while (end > start && text.length && normChar(oldText[end - 1]) === normChar(text[text.length - 1])) {
    end--;
    text = text.slice(0, -1);
  }
  return { start, end, text };
}

/** 截断显示用 */
export function clip(s: string, n = 120): string {
  const oneLine = s.replace(/\n/g, "⏎");
  return oneLine.length > n ? oneLine.slice(0, n) + "…" : oneLine;
}
