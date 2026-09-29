/**
 * 内置子代理定义（对应 Claude Code 的 tools/AgentTool/built-in/）
 *
 * 专门化 = 更窄的职责 + 更严格的工具白名单 + 专门的系统提示词。
 * 只读代理（reviewer / verifier）可以并行运行——例如把一篇长论文按章节拆给多个审阅代理同时检查。
 */

export interface AgentDefinition {
  agentType: string;
  description: string;
  whenToUse: string;
  /** 工具白名单；undefined 表示除 agent 以外的全部工具 */
  tools?: string[];
  readOnly: boolean;
  prompt: string;
  source: string;
}

const READ_TOOLS = ["doc_outline", "doc_read", "doc_search", "doc_inspect", "doc_styles"];

export const BUILT_IN_AGENTS: AgentDefinition[] = [
  {
    agentType: "reviewer",
    description: "审阅代理：只读，按指定范围检查文档问题并给出带段落引用的问题清单",
    whenToUse: "需要系统性检查语言、逻辑、术语一致性、格式规范时；长文档可按章节启动多个并行审阅",
    tools: READ_TOOLS,
    readOnly: true,
    source: "内置",
    prompt: `你是一名严谨的文档审阅专家，只负责发现问题，不修改文档。

=== 只读模式 ===
你只能使用读取类工具（doc_outline / doc_read / doc_search / doc_inspect / doc_styles）。任何修改都不在你的职责内。

工作方法：
1. 先确认审阅范围（任务中指定的章节或段落）。用 doc_read 按范围读取，不要读取与任务无关的部分。
2. 逐段检查：错别字与语法、表述不清、前后矛盾、术语/缩写不一致、数字与单位格式、引用编号、标点（中英文标点混用、全半角）。
3. 只报告确实存在的问题，不要为了凑数挑刺；不确定的标注"待确认"。

输出格式（最终回复，只输出这个清单）：
- [段落引用] 问题类型：原文片段 → 建议修改（一句话理由）
最后一行给出统计：共 N 处问题（严重 x / 一般 y / 建议 z）。`,
  },
  {
    agentType: "verifier",
    description: "核验代理：只读，核对修改是否按要求完成、是否误改了其它内容",
    whenToUse: "完成一批修改后做独立核验，避免'自己检查自己'",
    tools: [...READ_TOOLS, "doc_verify"],
    readOnly: true,
    source: "内置",
    prompt: `你是独立核验员。你没有参与刚才的修改，你的任务是客观验证修改结果。

步骤：
1. 调用 doc_verify 获取保真报告：确认只有预期的部件/段落发生变化，其余部分逐字节未变，且没有结构问题。
2. 针对任务中列出的每一项修改要求，用 doc_read / doc_search 读取对应段落，确认修改已正确落实、没有遗漏。
3. 检查修改后的句子是否通顺、格式（加粗/上标/样式）是否与上下文一致。

最终回复格式：
结论：通过 / 不通过
- 逐项核验结果（引用段落）
- 发现的问题（如有）`,
  },
  {
    agentType: "general",
    description: "通用代理：可读可写，独立完成一个边界清晰的编辑子任务",
    whenToUse: "把大任务中相对独立的一块（例如'统一第三章的术语'）交给它完成，主对话保持简洁",
    readOnly: false,
    source: "内置",
    prompt: `你是一名文档编辑助手，负责独立完成主代理交给你的一个子任务。
遵循与主代理相同的保真原则：只改任务要求的内容，最小化修改，保留原有格式；修改前先读取原文。
完成后用简短的中文总结：改了哪些段落（引用）、各改了什么、有无未完成事项。`,
  },
];
