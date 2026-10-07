/**
 * 重新转换后保留修改（润色、改写）
 *
 * 场景：PDF 转成 Word 后做过润色，之后发现转换本身有问题（公式、表格），需要用改进后的转换器重新转换。
 * 重新转换得到的是原文，这里把之前的修改逐段搬到新文档上：
 *   - 有"修改前的转换结果"（base）时：只搬 base → edited 中真正改过的段落；
 *   - 没有 base 时：按文字相似度把新旧段落对应起来，文字不同的段落用修改后的文字。
 * 段落按先后顺序对齐（单调匹配），相似度不够的不搬，并在报告中列出。
 * 写入用 DocxDocument.replaceText：只改真正变化的词，段落格式、字体字号保持新转换的结果。
 */

import { diffArrays } from "diff";
import { DocxDocument } from "../documents/docx/DocxDocument.js";

export interface CarryReport { applied: number; skipped: Array<{ text: string; reason: string }>; candidates: number }

const words = (s: string) => s.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [];
const norm = (s: string) => s.replace(/⟨[^⟩]*⟩/g, "").replace(/\s+/g, " ").trim();
/** 只比较"词"：去掉组合重音、标点、空白（转换器改进带来的 F̂ / 上下标顺序差异不算修改） */
const core = (s: string) => norm(s).normalize("NFD").replace(/[\u0300-\u036f\u20d0-\u20ff]/g, "").replace(/[^\p{L}\p{N}]+/gu, " ").trim().toLowerCase();
/** 段首编号（FIGURE 9. / TABLE 7. / D. / 2) / IV.） */
const label = (s: string) => /^\s*((fig(ure)?\.?|table)\s*[\dIVXLC]+|[A-Z]\.|\d+\)|[IVXLC]+\.)/i.exec(s)?.[0].replace(/\s+/g, " ").toUpperCase() ?? null;

/** 词序列相似度（0–1） */
function similarity(a: string, b: string): number {
  const A = words(a), B = words(b);
  if (!A.length || !B.length) return 0;
  let same = 0;
  for (const part of diffArrays(A, B)) if (!part.added && !part.removed) same += part.count ?? part.value.length;
  return (2 * same) / (A.length + B.length);
}

type Para = { ref: string; text: string };

function bodyParas(d: DocxDocument): Para[] {
  return d.listBlocks().filter((b: any) => (b.location === "正文" || /表格/.test(b.location ?? "")) && norm(b.text).length >= 2).map((b: any) => ({ ref: b.ref, text: b.text }));
}

/** 单调对齐：对每个 from 段落，在 to 中向后的窗口里找最相似的段落 */
function alignParas(from: Para[], to: Para[], min: number): Array<[number, number, number]> {
  const out: Array<[number, number, number]> = [];
  let ptr = 0;
  for (let i = 0; i < from.length; i++) {
    let best = -1, bestS = 0;
    for (let j = ptr; j < Math.min(to.length, ptr + 25); j++) {
      const s = similarity(from[i].text, to[j].text);
      if (s > bestS) { bestS = s; best = j; }
      if (s > 0.98) break;
    }
    if (best >= 0 && bestS >= min) { out.push([i, best, bestS]); ptr = best + 1; }
  }
  return out;
}

export function carryOverEdits(fresh: Buffer, edited: Buffer, base?: Buffer): { data: Buffer; report: CarryReport } {
  const target = new DocxDocument(fresh);
  const E = bodyParas(new DocxDocument(edited));
  const report: CarryReport = { applied: 0, skipped: [], candidates: 0 };
  // 需要搬的修改：edited 段落 → 它在"原文"中的样子
  let edits: Array<{ before: string; after: string }>;
  if (base) {
    const B = bodyParas(new DocxDocument(base));
    edits = [];
    for (const [bi, ei] of alignParas(B, E, 0.4)) if (norm(B[bi].text) !== norm(E[ei].text)) edits.push({ before: B[bi].text, after: E[ei].text });
  } else edits = E.map((e) => ({ before: e.text, after: e.text }));

  const T = bodyParas(target);
  const pairs = alignParas(edits.map((e) => ({ ref: "", text: e.before })), T, base ? 0.6 : 0.55);
  for (const [ei, ti, sim] of pairs) {
    const want = edits[ei].after, have = T[ti].text;
    if (norm(want) === norm(have) || core(want) === core(have)) continue;
    // 只有公式记号层面的差异（上下标 i、th、单个字母、数字的拆并）：是转换器的差异，不是润色
    {
      const A = core(have).split(" "), B = core(want).split(" ");
      const mathish = (t: string) => t.length <= 2 || /\d/.test(t) || /^[a-z]{1,2}(th|st|nd|rd)$/.test(t);
      const changed = diffArrays(A, B).filter((x) => x.added || x.removed).flatMap((x) => x.value);
      if (changed.every(mathish)) continue;
    }
    // 编号不同（另一张图 / 另一个小节）：对错了段落
    const lw = label(want), lh = label(have);
    if (lw && lh && lw !== lh) continue;
    // 修改后的文字只是新段落的一部分（旧转换把段落切碎了）：不是修改
    if (core(have).includes(core(want))) continue;
    // 短段落（标题、题注）要求更高的相似度
    if (have.length < 150 && sim < 0.85) continue;
    report.candidates++;
    // 新段落中有图片、公式等对象：只替换文字部分（对象不能被跨越）
    if (/⟨[^⟩]*⟩/.test(have) || /⟨[^⟩]*⟩/.test(want)) { report.skipped.push({ text: want.slice(0, 60), reason: "段落中含图片 / 公式等对象" }); continue; }
    // 修改后的文字里混入了旧转换的碎片（与新段落差异过大）：不搬，避免把错误带回来
    if (!base && sim < 0.7 && Math.abs(want.length - have.length) > 0.5 * have.length) { report.skipped.push({ text: want.slice(0, 60), reason: `与新段落相似度只有 ${Math.round(sim * 100)}%` }); continue; }
    try {
      target.replaceText({ ref: T[ti].ref, oldText: have, newText: want }, { track: false, author: "文案 Agent", date: new Date().toISOString() } as any);
      report.applied++;
    } catch (e: any) {
      report.skipped.push({ text: want.slice(0, 60), reason: String(e?.message ?? e).slice(0, 80) });
    }
  }
  return { data: target.serialize(), report };
}
