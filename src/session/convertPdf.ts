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
  const { data, report } = await pdfToDocx(pdf, { engine: opts.engine ?? "builtin", mode: opts.mode ?? "exact", title: base });
  const session = await Session.create(data, `${base}.docx`, opts.defaults);
  const text = describeReport(report);
  session.meta.origin = {
    fromSession: src.id,
    fromFile: src.meta.filename,
    engine: report.engine,
    mode: opts.engine === "builtin" || !opts.engine ? opts.mode ?? "exact" : undefined,
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
 */
export async function convertPdfInPlace(
  s: Session,
  opts: { engine?: PdfEngine; mode?: LayoutMode }
): Promise<{ report: ConversionReport; text: string; label: string }> {
  if (s.meta.format !== "pdf") throw new Error("当前文档不是 PDF，无需转换。");
  const pdf = await s.original();
  const base = s.meta.filename.replace(/\.pdf$/i, "");
  const engine = opts.engine ?? "builtin";
  const mode = opts.mode ?? "exact";
  const { data, report } = await pdfToDocx(pdf, { engine, mode, title: base });
  const text = describeReport(report);
  await s.switchDocument(data, "docx", `${base}.docx`, {
    fromSession: s.id,
    fromFile: s.meta.filename,
    engine: report.engine,
    mode: engine === "builtin" ? mode : undefined,
    report: text,
    coverage: report.textCoverage,
  });
  return { report, text, label: `由 PDF 转换为 Word（${engine === "builtin" ? (mode === "exact" ? "保留原排版" : "便于改写") : engine}）` };
}
