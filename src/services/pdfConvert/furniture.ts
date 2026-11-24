/**
 * 流式模式：哪些行作为页眉页脚离开正文
 *
 * Word 的一节只有一套页眉页脚（另有首页、偶数页两套），各页之间的差别只能靠页码域表达。
 * 所以一行文字只有在 Word 的页眉页脚能在它所在的那一页原样再现时，才写进页眉页脚：
 *   - 行的身份按数字逐位比较：只有逐页等于印刷页码的那一位（页码）可以不同，其余数字是文字的一部分。
 *     "第4条"与"第7条"、"附件1"与"附件2"、同一模板的两份函件的函号，都是不同的行；
 *   - 只在一页出现的行回到正文（栏底标题、恰好排在页首的条款标题、每页一份的附件标题）；
 *   - 在多页出现、确是页眉页脚、但 Word 写不出来的行：另一节的页眉、逐页递增却不是本文页码的编号（Bates 号、
 *     另起的页码）。它们不回正文（放回去会截断跨页的段落），在 notes 里如实说明没有写进去；
 *   - 页码域只放在真正的页码位上：页眉里的版本号、合同编号即使恰好等于某一页的页码，也保持原文。
 */

import type { Line } from "./layout.js";

export interface FurniturePick {
  /** 第一页的印刷页码（印刷页码 = 页序号 + pageNumStart）；没有可靠页码时为 undefined */
  pageNumStart?: number;
  /** 默认（奇数页）页眉页脚取自哪一页 */
  defPage: number;
  evenRep?: number;
  evenOdd: boolean;
  /** 第一页单独一套（与其余页不同） */
  titlePg: boolean;
}

export interface Furniture {
  /** 写进页眉页脚的行（逐页） */
  keep: Line[][];
  /** 是页眉页脚、但 Word 写不出来的行（不回正文，见 notes） */
  hidden: Line[][];
  /** 含页码位的行：只有这些行里的页码写成 PAGE 域 */
  pageNumberLines: Set<Line>;
  pick: FurniturePick;
  notes: string[];
}

/** 行的文字：相邻片段之间有可见间距才加空格（"ABC000" + "1234" 是一个编号 ABC0001234，与写进 Word 的段落一致） */
const lineText = (l: Line) => {
  const spans = [...l.spans].sort((a, b) => a.x0 - b.x0);
  let out = "", lastX = -Infinity;
  for (const sp of spans) {
    if (out && sp.x0 - lastX > 0.15 * sp.size) out += " ";
    out += sp.text;
    lastX = sp.x1;
  }
  return out.replace(/\s+/g, " ").trim();
};
/** 看起来像页码的行（数字归一后）：带"页 / Page / p."，或只有数字与装饰符号 */
const PAGEISH = /页|page|p\.|^[#\s\-–—_·•|/／()（）\[\]]*$/i;

type Item = { l: Line; i: number; s: string; runs: Array<{ v: string; at: number }>; kind: Array<"P" | "C" | undefined> };

/** 页码：某一位数字减去页序号在多数页上相同（与 buildHeaders 原来的判据一致） */
function pageOffset(lines: Line[][], n: number): number | undefined {
  const off = new Map<number, Set<number>>();
  lines.forEach((ls, i) => {
    for (const l of ls) for (const m of lineText(l).matchAll(/(?<![\d.,])\d{1,4}(?![\d.,]?\d)/g)) {
      const k = Number(m[0]) - i;
      if (!off.has(k)) off.set(k, new Set());
      off.get(k)!.add(i);
    }
  });
  const best = [...off.entries()].sort((a, b) => b[1].size - a[1].size)[0];
  return best && best[1].size >= Math.max(2, Math.ceil(n * 0.5)) ? best[0] : undefined;
}

/** 页码区间的写法："3、5–7" */
function ranges(nums: number[]): string {
  const s = [...new Set(nums)].sort((a, b) => a - b);
  const out: string[] = [];
  for (let k = 0; k < s.length; ) {
    let e = k;
    while (e + 1 < s.length && s[e + 1] === s[e] + 1) e++;
    out.push(e > k ? `${s[k]}–${s[e]}` : `${s[k]}`);
    k = e + 1;
  }
  return out.join("、");
}

/**
 * @param floating 规则判出的页眉页脚候选行（逐页）
 * @param mids 每页的半高：上半页的是页眉，下半页的是页脚
 * @param pageNos 每页在原 PDF 中的页码（写 notes 用）
 */
export function settleFurniture(floating: Line[][], mids: number[], pageNos: number[]): Furniture {
  const n = floating.length;
  const start = pageOffset(floating, n);
  const items: Item[] = [];
  floating.forEach((ls, i) => {
    for (const l of ls) {
      const s = lineText(l);
      const runs = [...s.matchAll(/\d+/g)].map((m) => ({ v: m[0], at: m.index! }));
      items.push({ l, i, s, runs, kind: runs.map(() => undefined) });
    }
  });
  const template = (it: Item) => it.s.replace(/\d+/g, "#").replace(/\s+/g, "");
  const groups = new Map<string, Item[]>();
  for (const it of items) {
    const k = template(it);
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k)!.push(it);
  }
  for (const [tpl, g] of groups) {
    for (let j = 0; j < g[0].runs.length; j++) {
      // 页码位：逐页等于印刷页码（不是小数、千分位的一部分，没有前导零），并且出现在足够多的页上
      if (start !== undefined) {
        const hits = g.filter((it) => {
          const { v, at } = it.runs[j];
          return v === String(it.i + start) && v.length <= 4 && !/[.,]$/.test(it.s.slice(0, at)) && !/^[.,]\d/.test(it.s.slice(at + v.length));
        });
        if (new Set(hits.map((h) => h.i)).size >= Math.max(2, Math.ceil(0.4 * n))) for (const h of hits) h.kind[j] = "P";
      }
      // 其余逐页递增的编号：连续三页以上、每页加一；只认页码样式的行（另起的页码）和 Bates 号这样的长编号，
      // 条款号、附件号碰巧逐页递增也不算
      if (!PAGEISH.test(tpl) && !g.every((it) => it.runs[j].v.length >= 5)) continue;
      const rest = g.filter((it) => !it.kind[j]).sort((a, b) => a.i - b.i);
      let seg: Item[] = [];
      const flush = () => {
        if (new Set(seg.map((x) => x.i)).size >= 3) for (const x of seg) x.kind[j] = "C";
        seg = [];
      };
      for (const it of rest) {
        const last = seg[seg.length - 1];
        if (last && (it.i - last.i > 2 || it.i === last.i || Number(it.runs[j].v) - it.i !== Number(last.runs[j].v) - last.i)) flush();
        seg.push(it);
      }
      flush();
    }
  }
  const key = (it: Item) => {
    let k = "", at = 0;
    it.runs.forEach((r, j) => {
      k += it.s.slice(at, r.at) + (it.kind[j] === "P" ? "#P" : it.kind[j] === "C" ? "#C" : r.v);
      at = r.at + r.v.length;
    });
    return (k + it.s.slice(at)).replace(/\s+/g, "");
  };
  const keys = new Map<Item, string>(items.map((it) => [it, key(it)]));
  const pagesOf = new Map<string, Set<number>>();
  for (const it of items) {
    const k = keys.get(it)!;
    if (!pagesOf.has(k)) pagesOf.set(k, new Set());
    pagesOf.get(k)!.add(it.i);
  }

  const keep: Line[][] = floating.map(() => []);
  const hidden: Line[][] = floating.map(() => []);
  const counters: Item[] = [];
  const candidates: Item[] = [];
  for (const it of items) {
    if (pagesOf.get(keys.get(it)!)!.size < 2) continue; // 只在一页出现：回到正文
    if (it.kind.includes("C")) counters.push(it);
    else candidates.push(it);
  }

  // 每页一套页眉页脚的"签名"；默认、偶数页、首页各取出现最多的那一套（与 buildHeaders 的选法相同，只是按文字身份比较）
  const printed = (i: number) => i + (start ?? 1);
  const byPage = floating.map(() => [] as Item[]);
  for (const it of candidates) byPage[it.i].push(it);
  const sig = (i: number) => {
    const part = (head: boolean) => byPage[i].filter((it) => (it.l.top < mids[i]) === head).sort((a, b) => a.l.top - b.l.top || a.l.x0 - b.l.x0).map((it) => keys.get(it)!).join("\n");
    return part(true) + "||" + part(false);
  };
  const sigs = floating.map((_, i) => sig(i));
  const major = (ids: number[]) => {
    const count = new Map<string, number>();
    let best: number | undefined, bestN = 0;
    for (const i of ids) {
      const c = (count.get(sigs[i]) ?? 0) + 1;
      count.set(sigs[i], c);
      if (c > bestN) { bestN = c; best = ids.find((x) => sigs[x] === sigs[i]); }
    }
    return best;
  };
  const rest = [...Array(n).keys()].slice(1);
  const oddRep = major(rest.filter((i) => printed(i) % 2 === 1)), evenRep = major(rest.filter((i) => printed(i) % 2 === 0));
  const evenOdd = oddRep !== undefined && evenRep !== undefined && sigs[oddRep] !== sigs[evenRep];
  const defPage = evenOdd ? oddRep! : major(rest) ?? 0;
  const rep0 = evenOdd && printed(0) % 2 === 0 ? evenRep! : defPage;
  const titlePg = n > 1 && sigs[0] !== sigs[rep0];
  const keysOf = (i: number) => new Set(byPage[i].map((it) => keys.get(it)!));
  const repKeys = new Map<number, Set<string>>();
  const allowed = (i: number) => {
    const rep = i === 0 && titlePg ? 0 : evenOdd ? (printed(i) % 2 === 0 ? evenRep! : oddRep!) : defPage;
    if (!repKeys.has(rep)) repKeys.set(rep, keysOf(rep));
    return repKeys.get(rep)!;
  };
  const others: Item[] = [];
  for (const it of candidates) {
    if (allowed(it.i).has(keys.get(it)!)) keep[it.i].push(it.l);
    // 写不进去的：至少三页都有，才认定是另一节的页眉页脚（不回正文）；只有两页的证据不足，宁可留在正文里
    else if (pagesOf.get(keys.get(it)!)!.size >= 3) others.push(it);
  }
  for (const it of [...counters, ...others]) hidden[it.i].push(it.l);

  const notes: string[] = [];
  const describe = (list: Item[], say: (where: string, pages: string, sample: string) => string) => {
    const byKey = new Map<string, Item[]>();
    for (const it of list) {
      const k = (it.l.top < mids[it.i] ? "页眉" : "页脚") + "\u0000" + template(it);
      if (!byKey.has(k)) byKey.set(k, []);
      byKey.get(k)!.push(it);
    }
    for (const [k, g] of [...byKey.entries()].slice(0, 3)) {
      const sample = g[0].s.replace(/\s+/g, " ").trim();
      notes.push(say(k.split("\u0000")[0], ranges(g.map((it) => pageNos[it.i])), sample.length > 30 ? `${sample.slice(0, 30)}…` : sample));
    }
  };
  describe(counters, (where, pages, sample) => `第 ${pages} 页${where}中逐页递增的编号（如"${sample}"）不是本文的页码，Word 的页码域表示不了，没有写进${where}。`);
  describe(others, (where, pages, sample) => `第 ${pages} 页的${where}（如"${sample}"）与其余页不同，Word 一节只能有一套${where}，这几页的${where}没有写进去。`);

  return {
    keep, hidden,
    pageNumberLines: new Set(items.filter((it) => it.kind.includes("P")).map((it) => it.l)),
    pick: { pageNumStart: start, defPage, evenRep: evenOdd ? evenRep : undefined, evenOdd, titlePg },
    notes,
  };
}
