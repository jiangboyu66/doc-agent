/**
 * PDF 会话 → 新的 Word 会话
 *
 * PDF 本身只读；转换后生成一个独立的 Word 会话（原 PDF 会话保持不变），转换报告写入新会话的 origin，
 * 之后的所有编辑都在 Word 文档上进行，享受完整的保真编辑、修订、版本回滚能力。
 */

import { Session } from "./Session.js";
import { pdfToDocx, describeReport, type PdfEngine, type ConversionReport } from "../services/pdfConvert/index.js";
import type { LayoutMode } from "../services/pdfConvert/layout.js";

export async function convertPdfSession(
  src: Session,
  opts: { engine?: PdfEngine; mode?: LayoutMode; defaults: { mode: Session["meta"]["mode"]; trackChanges: boolean; author: string } }
): Promise<{ session: Session; report: ConversionReport; text: string }> {
  if (src.meta.format !== "pdf") throw new Error("只有 PDF 文档需要转换。");
  const pdf = await src.original();
  const base = src.meta.filename.replace(/\.pdf$/i, "");
  const { data, report } = await pdfToDocx(pdf, { engine: opts.engine ?? "builtin", mode: opts.mode ?? "flow", title: base });
  const session = await Session.create(data, `${base}.docx`, opts.defaults);
  const text = describeReport(report);
  session.meta.origin = {
    fromSession: src.id,
    fromFile: src.meta.filename,
    engine: report.engine,
    mode: opts.engine === "builtin" || !opts.engine ? opts.mode ?? "flow" : undefined,
    report: text,
    coverage: report.textCoverage,
  };
  session.meta.versions[0].label = `由 PDF 转换（${src.meta.filename}）`;
  await session.save();
  return { session, report, text };
}

/**
 * 在当前会话内把 PDF 转换为 Word（对话历史保留，Agent 可以接着修改）：
 * 转换结果成为新的工作文档与保真基准，原 PDF 仍是 v0，可随时查看或回滚。调用方负责 commit()。
 * 已经转换过的会话（原文件是 PDF）可以换一种版式模式重新转换（例如逐页一致 → 流式，以便扩写）。
 * 默认流式：字体、字号、分栏与原文一致，文字增减时自动重排，适合继续编辑。
 */
export async function convertPdfInPlace(
  s: Session,
  opts: { engine?: PdfEngine; mode?: LayoutMode }
): Promise<{ report: ConversionReport; text: string; label: string }> {
  const reconvert = s.meta.format !== "pdf" && s.meta.sourceFormat === "pdf";
  if (s.meta.format !== "pdf" && !reconvert) throw new Error("当前文档不是 PDF，无需转换。");
  const pdf = reconvert ? (await s.uploaded()).data : await s.original();
  const base = (reconvert ? s.meta.sourceFile ?? s.meta.filename : s.meta.filename).replace(/\.(pdf|docx)$/i, "");
  const engine = opts.engine ?? "builtin";
  const mode = opts.mode ?? "flow";
  let { data, report } = await pdfToDocx(pdf, { engine, mode, title: base });
  let text = describeReport(report);
  // 重新转换：把转换之后做过的修改（润色、改写）搬到新文档上
  if (reconvert && s.meta.currentVersion > 0) {
    try {
      const { carryOverEdits } = await import("../services/carryOver.js");
      const prevBase = await s.original();
      const r = carryOverEdits(data, s.current(), prevBase);
      if (r.report.candidates) {
        data = r.data;
        text += `\n已把之前的 ${r.report.applied} 处段落修改搬到新文档上` + (r.report.skipped.length ? `；${r.report.skipped.length} 处没能自动搬运（${r.report.skipped.slice(0, 3).map((x) => `「${x.text.slice(0, 20)}…」${x.reason}`).join("；")}），需要重新修改。` : "。");
      }
    } catch (e: any) {
      text += `\n之前的修改没能自动搬运（${String(e?.message ?? e).slice(0, 80)}），需要重新修改。`;
    }
  }
  await s.switchDocument(data, "docx", `${base}.docx`, {
    fromSession: s.id,
    fromFile: reconvert ? s.meta.sourceFile ?? s.meta.filename : s.meta.filename,
    engine: report.engine,
    mode: engine === "builtin" ? mode : undefined,
    report: text,
    coverage: report.textCoverage,
  });
  const how = engine === "builtin" ? (mode === "exact" ? "逐页保留原排版" : "可编辑流式排版") : engine;
  return { report, text, label: `${reconvert ? "从原 PDF 重新转换" : "由 PDF 转换为 Word"}（${how}）` };
}
