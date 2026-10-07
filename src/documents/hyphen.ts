/**
 * 行尾连字符的判断：PDF 排版时在行尾断开单词（"sharp-" + "ness"）与本身带连字符的复合词在行尾断开
 * （"noise-" + "dominated"）外观一样。重新排版（段落自动换行）时，前者要去掉连字符，后者要保留。
 *
 * 依据（按优先级）：文档中别处出现过带连字符的写法 → 保留；出现过连写的写法 → 去掉；
 * 前半部分是常见的复合词前缀 → 保留；其余（音节断词）→ 去掉。
 */

const KEEP_PREFIX = new Set([
  "multi", "cross", "self", "non", "pre", "post", "co", "sub", "semi", "anti", "inter", "intra", "over", "under", "well",
  "low", "high", "two", "three", "four", "single", "double", "full", "half", "state", "deep", "end", "real", "long", "short",
  "large", "small", "fine", "coarse", "patch", "pixel", "slice", "noise", "edge", "data", "task", "domain", "channel",
  "spatial", "frequency", "image", "feature", "level", "scale", "window", "time", "user", "open", "closed", "top", "bottom",
  "left", "right", "first", "second", "third", "zero", "one", "dual", "quasi", "pseudo",
]);

export function collectWords(text: string): Set<string> {
  const out = new Set<string>();
  for (const m of text.matchAll(/[A-Za-z]+(?:-[A-Za-z]+)*/g)) {
    const w = m[0].toLowerCase();
    out.add(w);
    if (w.includes("-")) out.add(w.replace(/-/g, ""));
  }
  return out;
}

/** 每个词（小写）在全文中独立出现的次数 */
export function wordCounts(text: string): Map<string, number> {
  const out = new Map<string, number>();
  for (const m of text.matchAll(/[A-Za-z]+/g)) {
    const w = m[0].toLowerCase();
    out.set(w, (out.get(w) ?? 0) + 1);
  }
  return out;
}

/** a：连字符前的字母片段；b：下一行的第一个词。返回 true 表示这是排版断词，应去掉连字符 */
export function isSoftHyphen(a: string, b: string, words: Set<string>, hyphenated?: Set<string>, counts?: Map<string, number>): boolean {
  const A = (/[A-Za-z]+$/.exec(a)?.[0] ?? "").toLowerCase();
  const B = (/^[A-Za-z]+/.exec(b)?.[0] ?? "");
  if (!A || !B) return false;
  // 多段复合词的中间（image-to-image、state-of-the-art）
  if (/-[A-Za-z]+$/.test(a)) return false;
  if (/^[A-Z]/.test(B) && !/^[A-Z]+$/.test(B)) return false; // 下一段以大写开头：多为专有名词复合（Swin-Transformer）
  const lowB = B.toLowerCase();
  if (hyphenated?.has(`${A}-${lowB}`)) return false;
  if (words.has(`${A}-${lowB}`)) return false;
  if (words.has(`${A}${lowB}`)) return true;
  if (KEEP_PREFIX.has(A)) return false;
  // 缩写词 + 连字符（NDCT-consistent、CT-Mamba）是复合词
  const rawA = /[A-Za-z]+$/.exec(a)?.[0] ?? "";
  if (rawA.length >= 2 && rawA === rawA.toUpperCase()) return false;
  // 两半在文中都作为独立的词出现过（除了这一次之外至少还有一次）：复合词（attention-gated）
  if (counts && A.length >= 3 && (counts.get(A) ?? 0) >= 2 && (counts.get(lowB) ?? 0) >= 2) return false;
  return A.length >= 2 && lowB.length >= 2;
}

/** 文档中真正带连字符的写法（不在行尾断开处的） */
export function collectHyphenated(text: string): Set<string> {
  const out = new Set<string>();
  for (const m of text.matchAll(/[A-Za-z]+(?:-[A-Za-z]+)+/g)) out.add(m[0].toLowerCase());
  return out;
}
