/**
 * 文档适配器接口（策略模式）
 *
 * 不同格式的文档（Word / Markdown / HTML / PDF）各自实现这个接口。工具层只面向接口编程，
 * 通过 capabilities 判断当前文档支持哪些操作——对应 Claude Code 里工具的 isEnabled()：
 * 文档不支持的能力，对应工具不会出现在模型可用的工具列表里，从源头避免模型尝试做不到的事。
 *
 * 核心原则：所有编辑都直接作用于"源格式"本身，不做格式往返转换，因此未被修改的内容保持原样。
 */

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
}

export class DocError extends Error {
  constructor(
    message: string,
    public readonly code: "not_found" | "ambiguous" | "stale_ref" | "unsupported" | "protected" | "invalid" = "invalid"
  ) {
    super(message);
  }
}
