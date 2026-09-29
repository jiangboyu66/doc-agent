# 文案 Agent v2

基于 DeepSeek 的文档编辑 Agent，支持 Word（.docx）、Markdown、HTML，PDF 只读审阅。
提供 Web 界面和命令行（CLI）两种用法，二者共用同一个引擎。

**核心目标：100% 保真。** 只改你要求改的地方，文档其余部分逐字节保持不变。
架构按 [cc-src-learning](https://github.com/mylxsw/cc-src-learning) 对 Claude Code 源码的分析逐项实现，对应关系见下文。

---

## 快速开始

需要 Node.js ≥ 20。

```bash
npm install
cp .env.example .env        # Windows: copy .env.example .env，然后填写 DEEPSEEK_API_KEY
npm run dev                 # 打开 http://localhost:3939
```

压缩包里已经带有构建好的前端（`web/dist`）。如果修改了前端代码，运行 `npm run build:web` 重新构建。

可选的外部工具：

| 工具 | 用途 | 没有它时 |
| --- | --- | --- |
| LibreOffice（soffice） | 导出 PDF、Word 的“精确版式”预览 | 这两项不可用，其余功能正常 |
| pandoc | 跨格式导出（如 Word → HTML） | 跨格式导出不可用；同格式编辑与导出不受影响 |

命令不在 PATH 中时，在 `.env` 里设置 `SOFFICE_PATH` 或 `PANDOC_PATH`。

### 命令行

```bash
npm run cli -- 论文.docx                       # 交互模式
npm run cli -- --resume <会话ID>               # 恢复会话（Web 与 CLI 的会话互通）
npm run cli -- 论文.docx -p "统一术语：智能代理→文档代理" --mode acceptEdits --track --output 结果.docx
                                               # 无人值守：执行后导出退出
```

### 测试

```bash
npm test                                            # 引擎保真测试 + 模拟模型跑完整代理流程（不需要 API Key）
DOC_AGENT_TEST_DOCX=你的文档.docx npm test          # 额外用你的真实文档做往返与修订测试
```

---

## 保真是怎么保证的

v1 先把 Word 转成 Markdown 再转回去，每转一次都会丢格式。v2 彻底放弃了中间格式，直接在 OOXML 上做外科手术式的修改。

1. **ZIP 原样拷贝**（`src/documents/docx/zip.ts`）
   只有被修改的部件会重新压缩，图片、字体、样式表、主题等其他部件逐字节拷贝原始数据。
2. **XML 原样输出**（`src/documents/docx/xml.ts`）
   每个节点都记着自己在源文件中的位置。没有改动的节点直接输出原始文本，属性顺序、命名空间声明、空白和实体写法都不会变。
3. **最小差分**（`src/documents/textMatch.ts`）
   模型给出新句子后，引擎在字符和词的粒度上计算差异，只替换真正变化的片段。
   例如 “Author1” 改成 “Smith1” 时，上标的 “1” 仍然保留在原来的 run 里。
   匹配时容忍直引号和弯引号、各种破折号、不换行空格的差异。
4. **不可编辑对象受保护**
   图片、公式、文本框、脚注引用、域代码、批注边界等显示为 ⟨图片⟩ ⟨公式⟩ ⟨脚注1⟩ 之类的占位符。修改范围一旦跨过它们就会被拒绝。
5. **新内容沿用已有格式**
   新段落复制同样式段落的段落属性，并使用文字最多的 run 的格式，不会继承首字的加粗、颜色等临时格式。
   属性按 OOXML Schema 规定的顺序插入。
6. **Word 修订模式**
   所有修改都可以记录为 `w:ins` / `w:del` / `w:rPrChange` / `w:pPrChange` 修订，在 Word 里逐条接受或拒绝。批注写入 `comments.xml`。
7. **保真守卫**
   每次写操作之后，内置 Hook 会校验 XML 合法性、部件完整性和修订 ID 唯一性。校验不通过时，这次修改自动回滚，并告诉模型换一种改法。
8. **版本快照**
   每次修改都保存完整的版本。可以回滚到任意版本，回滚本身也是一个新版本，历史不会丢失。
   原始文件永远不改，作为保真对比的基准。

Markdown 和 HTML 采用同样的思路：直接在源文本上按位置拼接，保留原有的换行符（CRLF/LF）、实体写法、属性和注释。

实测结果（IEEE Access 模板，10 类修改，覆盖普通模式和修订模式）：

- 通过 OOXML Schema 校验和修订完整性校验；
- LibreOffice 渲染后，除修改处外版式完全一致；
- 未修改时往返输出字节一致；
- 38 个部件中 34 个字节一致。

输入 `/verify` 可以随时查看当前文档的保真报告。

---

## 架构：cc-src-learning 设计理念对照

| 设计理念 | Claude Code 中 | 本项目 |
| --- | --- | --- |
| 代理主循环（异步生成器，流式事件） | `QueryEngine.ts` / `query.ts` | `src/QueryEngine.ts` |
| 权限确认在循环内 await，不拆成“暂停/恢复” | `canUseTool` | `QueryEngine.runToolUse` + `server.ts` 的 pending 表 |
| 工具接口 + `buildTool` 安全默认值（fail-closed） | `Tool.ts` | `src/Tool.ts` |
| 工具池按上下文组装（能力不支持的工具不给模型） | `tools.ts` / `assembleToolPool` | `src/tools.ts`，`isEnabled` 读文档 capabilities |
| 只读工具并行、写工具串行 | `isConcurrencySafe` 分批 | `QueryEngine.runTools` |
| 先读后改 | Read-before-Edit 校验 | `docTools.ts` 的 `requireRead` |
| 权限模式 default / acceptEdits / plan / bypassPermissions | `permissions/` | `src/permissions/permissions.ts` |
| allow / deny / ask 规则，会话级与项目级记忆 | settings 权限规则 | `permissions` 配置 + 「本会话都允许 / 以后都允许」 |
| 计划模式与 ExitPlanMode | `ExitPlanModeTool` | `metaTools.ts` 的 `exit_plan_mode` |
| 确认前的修改预览 | Edit diff 预览 | `session.dryRun`：在副本上干跑，生成前后对比 |
| Hooks（PreToolUse / PostToolUse / UserPromptSubmit / Stop…） | `hooks/` | `src/hooks/hooks.ts`，退出码 2 表示阻止；内置保真守卫 |
| 分层系统提示词 + 动态边界 | `constants/prompts.ts` | `src/constants/prompts.ts`，每轮状态放进 `<system-reminder>`，保持前缀稳定以命中 DeepSeek 缓存 |
| 系统提示词优先级（override > agent > custom > default） | `buildEffectiveSystemPrompt` | `src/utils/systemPrompt.ts` |
| Skills：SKILL.md + frontmatter，渐进披露 | `skills/` | `src/skills/loadSkills.ts`、`skills/`（IEEE 论文、校对、学术润色、格式统一、保格式翻译） |
| 子代理（只读审校并行、fork 继承上下文） | `AgentTool` | `metaTools.ts` 的 `agent` + `src/agents/builtInAgents.ts` |
| 插件：同时贡献技能、命令、子代理、Hook | `plugins/` | `src/plugins/pluginLoader.ts` + 示例插件 `plugins/zh-typography` |
| 斜杠命令：local（本地执行，不调模型）/ prompt（展开为提示词） | `commands/` | `src/commands.ts`（/undo /verify /export /cost /compact /review …） |
| 上下文压缩：微压缩 + 自动摘要 + /compact | `services/compact` | `src/services/compact.ts` |
| 消息规范化（修复 tool_call 配对） | `normalizeMessagesForAPI` | `src/services/api/normalize.ts` |
| API 错误分类 + 指数退避重试 | `services/api` | `src/services/api/deepseek.ts` |
| 成本追踪（含缓存命中、峰谷价） | `cost-tracker.ts` | `src/services/costTracker.ts`，`/cost` |
| TodoWrite 任务清单 | `TodoWriteTool` | `metaTools.ts` 的 `todo_write`，Web 端实时显示 |
| 分层配置（用户 < 项目 < 环境变量 < 命令行） | settings | `src/config/settings.ts` |
| Feature flags | `feature()` | `src/config/features.ts` |
| 启动并行预加载 | `main.tsx` | `src/bootstrap.ts` |
| 长期记忆（CLAUDE.md） | memory | `DOCAGENT.md`（用户级 + 项目级），`/remember` 写入 |
| 桥接协议（同一引擎服务多个前端） | `bridge/` | `src/bridge/protocol.ts`：Web 通过 SSE 接收，CLI 在进程内消费 |

### 目录

```
src/
  QueryEngine.ts          代理主循环
  Tool.ts  tools.ts       工具接口与工具池
  tools/docTools.ts       文档工具（读取/搜索/修改/格式/插入/删除/表格/批注/导出/校验）
  tools/metaTools.ts      todo_write / exit_plan_mode / skill / agent
  documents/              文档适配器（策略模式）
    docx/                 原生 OOXML 引擎（zip / xml / ooxml / DocxDocument）
    markdown/ html/ pdf/
  permissions/ hooks/ skills/ agents/ plugins/ config/ services/ session/
  entrypoints/server.ts   Web 服务（Express + SSE）
  entrypoints/cli.ts      命令行
web/                      React 前端（Vite）
skills/                   内置技能
plugins/                  插件
examples/                 配置与 Hook 示例
test/                     自动化测试
```

---

## 配置

配置分层合并，优先级从低到高：

1. 内置默认
2. `~/.doc-agent/settings.json`
3. `.doc-agent/settings.json`
4. `.env`
5. 命令行参数

完整示例（权限规则、Hook、功能开关）见 `examples/settings.example.json`，复制到 `.doc-agent/settings.json` 即可生效。

权限规则写法：
- `doc_replace_text`：该工具全部允许。
- `doc_delete_blocks`：配合 `ask`，每次都确认。
- 规则支持 `工具名(内容*)` 通配。

长期记忆：在项目根目录写 `DOCAGENT.md`，例如“术语统一用‘文档代理’”“参考文献用 GB/T 7714 格式”。Agent 每次会话都会读取。

### 模型

| 模型 | 说明 |
| --- | --- |
| `deepseek-flash` | 默认，快且便宜 |
| `deepseek-v4-pro` | 更强，适合复杂改写 |

旧的 `deepseek-chat` / `deepseek-reasoner` 已于 2026-07-24 下线。

思考模式由 `DEEPSEEK_THINKING` 控制。开启时，引擎会按 DeepSeek 的要求在多轮工具调用中回传 `reasoning_content`。

---

## Windows 说明

- 路径统一用 `fileURLToPath` / `pathToFileURL` 处理，中文路径和文件名可正常使用。
- LibreOffice 每次都用独立的临时用户配置启动，不会被已打开的 LibreOffice 窗口卡住。转换有超时保护。
- 如果 `soffice` 不在 PATH 中：`SOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe`。

## 已知限制

- **PDF 只读。** PDF 没有可编辑的段落结构，只能阅读、搜索和审阅。需要修改请提供原始 Word 文件。
- **不支持 .doc。** 请先在 Word 中另存为 .docx。
- **跨格式导出（Word → HTML/Markdown）必然有损。** 需要完全保真请导出原格式。
- **“快速预览”由浏览器渲染，与 Word 在分页和字体度量上可能略有差异。** 以“精确版式”（LibreOffice）或 Word 为准。下载的文件本身不受影响。
- 修改跨越图片、公式、域代码等对象时会被拒绝，这是保护机制。Agent 会改为分段修改。
