/**
 * 文档适配器接口（策略模式）
 *
 * 不同格式的文档（Word / Markdown / HTML / PDF）各自实现这个接口。工具层只面向接口编程，
 * 通过 capabilities 判断当前文档支持哪些操作——对应 Claude Code 里工具的 isEnabled()：
 * 文档不支持的能力，对应工具不会出现在模型可用的工具列表里，从源头避免模型尝试做不到的事。
 *
 * 核心原则：所有编辑都直接作用于"源格式"本身，不做格式往返转换，因此未被修改的内容保持原样。
 */

import type { ShapeSpec, ChartSpec } from "./docx/build.js";

export type DocFormat = "docx" | "markdown" | "html" | "pdf";

export interface DocumentCapabilities {
  replaceText: boolean;
  formatText: boolean;
  paragraphProps: boolean;
  insertBlocks: boolean;
  deleteBlocks: boolean;
  tables: boolean;
  comments: boolean;
  trackChanges: boolean;
  styles: boolean;
  /** 插入图片 / 矢量图形 / 图表 */
  media?: boolean;
  /** 分页符分节符、页面设置、页眉页脚、目录 */
  layout?: boolean;
  /** 项目符号与编号 */
  lists?: boolean;
  /** 脚注尾注 */
  notes?: boolean;
  /** 超链接 */
  links?: boolean;
  /** 修改 / 新建样式 */
  styleEdit?: boolean;
  /** 接受 / 拒绝修订 */
  revisions?: boolean;
  /** 新建表格与表格结构编辑 */
  tableEdit?: boolean;
  /** 移动段落 */
  move?: boolean;
  /** 公式 */
  equations?: boolean;
  /** 可以插入原始标记（Markdown / HTML 源码），表格、图片等由工具层生成标记后插入 */
  rawInsert?: boolean;
}

export interface BlockInfo {
  /** 稳定引用，工具调用时用它定位段落 */
  ref: string;
  /** 所在位置描述：正文 / 表格1 第2行第3列 / 页眉(header1) / 脚注2 … */
  location: string;
  kind: "paragraph" | "heading" | "title" | "list" | "table-cell" | "code" | "quote" | "other";
  level?: number;
  style?: string;
  /** 渲染文本：不可编辑对象以 ⟨图片⟩ ⟨公式⟩ ⟨脚注1⟩ 等占位符表示 */
  text: string;
  flags?: string[];
}

export interface SearchHit {
  ref: string;
  location: string;
  text: string;
  index: number;
  match: string;
}

export interface OutlineItem {
  ref: string;
  level: number;
  text: string;
  location: string;
}

export interface DocumentSummary {
  format: DocFormat;
  blockCount: number;
  charCount: number;
  headings: OutlineItem[];
  parts: string[];
  tables: number;
  images: number;
  sections: number;
  trackedChanges: number;
  comments: number;
  template?: { id: string; name: string; evidence: string[] };
  notes: string[];
}

export interface EditOptions {
  track: boolean;
  author: string;
  date: string;
}

export interface EditResult {
  changedRefs: string[];
  /** 给模型看的简短说明 */
  summary: string;
  /** 给用户看的前后对比 */
  preview: Array<{ ref: string; before: string; after: string }>;
  structural: boolean;
}

export interface RunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: "single" | "double" | "none";
  strike?: boolean;
  color?: string;
  highlight?: string;
  size_pt?: number;
  font?: string;
  vertical?: "superscript" | "subscript" | "baseline";
  small_caps?: boolean;
  all_caps?: boolean;
  char_style?: string;
}

export interface ParagraphFormat {
  style?: string;
  alignment?: "left" | "center" | "right" | "justify";
  space_before_pt?: number;
  space_after_pt?: number;
  line_spacing?: number;
  indent_left_pt?: number;
  indent_right_pt?: number;
  first_line_pt?: number;
  hanging_pt?: number;
  keep_with_next?: boolean;
  page_break_before?: boolean;
}

export interface NewBlock {
  text: string;
  style?: string;
  /** 以哪个已有段落为格式模板（默认用锚点段落） */
  like?: string;
}

export interface FidelityReport {
  ok: boolean;
  totalParts: number;
  identicalParts: number;
  modifiedParts: Array<{ name: string; changedBlocks: string[]; note?: string }>;
  addedParts: string[];
  problems: string[];
}

export interface DocumentAdapter {
  readonly format: DocFormat;
  readonly capabilities: DocumentCapabilities;
  /** 结构版本号：插入/删除段落后递增，用于检测过期的段落引用 */
  structureVersion: number;

  summary(): DocumentSummary;
  listBlocks(opts?: { part?: string }): BlockInfo[];
  getBlock(ref: string): BlockInfo;
  /** 某段落的格式详情（样式、对齐、run 级格式） */
  describeFormat(ref: string): string;
  search(query: string, opts: { regex?: boolean; caseSensitive?: boolean }): SearchHit[];
  listStyles(): Array<{ id: string; name: string; type: string; inUse: number }>;

  replaceText(p: { ref?: string; oldText: string; newText: string; replaceAll?: boolean }, o: EditOptions): EditResult;
  formatText(p: { ref: string; text?: string; format: RunFormat }, o: EditOptions): EditResult;
  setParagraph(p: { refs: string[]; format: ParagraphFormat }, o: EditOptions): EditResult;
  insertBlocks(p: { anchor: string; position: "before" | "after"; blocks: NewBlock[] }, o: EditOptions): EditResult;
  deleteBlocks(p: { refs: string[] }, o: EditOptions): EditResult;
  insertTableRow(p: { ref: string; position: "before" | "after"; cells: string[] }, o: EditOptions): EditResult;
  addComment(p: { ref: string; text?: string; comment: string }, o: EditOptions): EditResult;

  serialize(): Buffer;
  fidelity(original: Buffer): FidelityReport;

  /** 把旧版本的临时引用换算成当前引用（段落仍在时） */
  translateRef?(ref: string): string;
  /** 由本程序从 PDF 转换来的 Word：转换时的版式模式（exact 逐页一致、flow 流式）；其他文档为 undefined */
  convertedLayout?(): "exact" | "flow" | undefined;
  /** 整体替换文档内容（全局排版工具使用） */
  reload?(buf: Buffer): void;

  // ---- 扩展能力（可选，由 capabilities 声明是否支持）----
  insertRaw?(p: { anchor: string; position: "before" | "after"; markup: string; label?: string }, o: EditOptions): EditResult;
  insertTable?(p: { anchor: string; position: "before" | "after"; rows: string[][]; header?: boolean; style?: "grid" | "three_line" | "plain" | "banded"; align?: "left" | "center" | "right"; widths?: number[]; width_pct?: number; font_size_pt?: number; caption?: string; caption_position?: "above" | "below" }, o: EditOptions): EditResult;
  editTable?(p: any, o: EditOptions): EditResult;
  insertImage?(p: { anchor: string; position: "before" | "after"; data: Buffer; width_pt?: number; height_pt?: number; align?: "left" | "center" | "right"; caption?: string; alt?: string; name?: string }, o: EditOptions): EditResult;
  insertShapes?(p: { anchor: string; position: "before" | "after"; width_pt: number; height_pt: number; shapes: ShapeSpec[]; caption?: string; alt?: string; align?: "left" | "center" | "right"; fallbackPng?: Buffer }, o: EditOptions): EditResult;
  insertChart?(p: { anchor: string; position: "before" | "after"; chart: ChartSpec; width_pt?: number; height_pt?: number; caption?: string; fallbackPng?: Buffer }, o: EditOptions): EditResult;
  insertEquation?(p: { anchor?: string; position?: "before" | "after"; ref?: string; after_text?: string; latex: string; display?: boolean; number?: string }, o: EditOptions): EditResult;
  insertBreak?(p: { anchor: string; position: "before" | "after"; kind: "page" | "column" | "section_next_page" | "section_continuous" | "section_odd_page" | "section_even_page" }, o: EditOptions): EditResult;
  pageSetup?(p: any, o: EditOptions): EditResult;
  setHeaderFooter?(p: { kind: "header" | "footer"; text: string; align?: "left" | "center" | "right"; ref?: string; first_page?: boolean }, o: EditOptions): EditResult;
  setList?(p: { refs: string[]; kind: "bullet" | "number" | "chinese" | "outline" | "none"; level?: number; restart?: boolean }, o: EditOptions): EditResult;
  insertNote?(p: { ref: string; after_text?: string; text: string; kind?: "footnote" | "endnote" }, o: EditOptions): EditResult;
  insertLink?(p: { ref: string; text: string; url: string }, o: EditOptions): EditResult;
  insertToc?(p: { anchor: string; position: "before" | "after"; levels?: number; title?: string }, o: EditOptions): EditResult;
  modifyStyle?(p: any, o: EditOptions): EditResult;
  reviewChanges?(p: { action: "accept" | "reject"; refs?: string[]; author?: string }, o: EditOptions): EditResult;
  listComments?(): Array<{ id: string; author: string; date: string; text: string; ref?: string; anchor: string }>;
  deleteComments?(p: { ids?: string[]; all?: boolean }): EditResult;
  moveBlocks?(p: { refs: string[]; anchor: string; position: "before" | "after" }, o: EditOptions): EditResult;
}

export class DocError extends Error {
  constructor(
    message: string,
    public readonly code: "not_found" | "ambiguous" | "stale_ref" | "unsupported" | "protected" | "invalid" = "invalid"
  ) {
    super(message);
  }
}
