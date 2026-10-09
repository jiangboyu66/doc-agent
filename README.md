# 文案 Agent v2

基于 DeepSeek 的文档编辑 Agent，支持 Word（.docx）、Markdown、HTML；PDF 可在会话内一键转换为 Word 后继续编辑。
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

## PDF 转 Word

PDF 是只读的最终版式格式，不能直接改写。上传 PDF 后，页面上方会出现「PDF → Word」卡片：

- **下载 Word 文件**：直接得到 `原文件名.docx`；
- **转换并继续编辑**：在**当前会话内**转换为 Word（对话保留），之后直接让 Agent 修改；原 PDF 保留为版本 v0，可随时回滚。默认使用**可编辑排版**：扩写、改写后版面自动重排，不会错乱。

卡片可以点标题左侧的 ▾ 收起成一行，让文档预览占满空间（折叠状态会记住）。在对话里直接说"把摘要润色一下"也可以：Agent 会先调用 `doc_convert_to_word` 在会话内转换，再接着修改，不需要你手动切换文档。

也可以在对话中直接说"把这个 PDF 导出成 Word"（Agent 调用 `doc_export`），或输入 `/export docx`（`/export docx flow` 为便于改写版），或在右上角「下载 / 导出」菜单中选择「导出为 Word（.docx）」。

内置转换引擎（`src/services/pdfConvert/`，纯 Node 实现，不需要安装任何额外程序）按"字在哪里、线在哪里"重建 Word 文档：

- **文字**：字体（含子集字体名解析；LaTeX 的 Nimbus / URW / TeX Gyre / Computer Modern / Latin Modern 字体、LibreOffice/Linux 度量兼容字体映射回 Times New Roman / Arial / Calibri / Cambria Math 等，并按 Word 字体的实际字宽做不折行校验）、字号、粗斜体、颜色、上下标、下划线、删除线、Symbol 字体符号（Φ → ×）、连字（ﬁ → fi）。
- **段落**：按基线成行、按折行规律成段；对齐方式由实测字宽判断（两端对齐的行会被拉宽或压缩）；首行缩进、悬挂缩进、列表符号、制表位（原位置）、段前距与行距（按原基线位置计算）；样式中的字符紧缩/加宽也会被测出并还原。
- **表格**：由边框线还原网格，含合并单元格、每条边的有无/粗细/颜色、底纹、单元格垂直对齐；学术论文常见的"三线表"（只有横线）也能识别。
- **版面**：双栏检测（两栏基线不对齐也能识别）→ Word 分栏；页眉页脚、页面装饰按页面位置放置。
- **矢量图、行间公式、首字下沉**：流程图、曲线图、TeX 公式（分式、根号、大括号、求和号）无法用段落还原，会被自动识别，按 288 dpi 从 PDF 原样渲染成透明背景的高清图片；图中/公式中的文字写入图片的替代文字，并计入完整性校验。公式和独立成块的图是**行内图片**（独占一段），随文字流移动——前面的文字变多变少时不会与文字重叠，快速预览中位置也正确；带 (a)(b) 标注的位图、多张照片排成的组图连同子图编号整体渲染成一张图（论文里的示意图常被出版社切成十来张相邻的位图切片，这类“位图拼图”也会合并成一张图；图题第一行里的 “(a) …” 不会被误当成子图编号而把图题并进来）；曲线图、柱状图的网格线不会被误当作表格，多个并排的子图面板合成一幅图。渲染依赖 `@napi-rs/canvas`（`npm install` 时自动安装预编译版本，Windows / macOS / Linux 通用）；加载失败时自动退回文字重建。设置 `PDF_RASTER=0` 可关闭此功能。
- **没有文字编码的公式与符号（Type 3 字体）**：老式 TeX / DVI 生成的 PDF（如 2000 年前后的 IEEE 论文）常把数学符号画成 Type 3 位图字形——它们既不是文字也不是路径，以前会整片消失（公式只剩编号、行内的 $M$、$t^{(m)}$ 变成空白），带乱码编码的 Type 3 字形则以 Word 中不存在的 "Type3" 字体出现、显示成另一种字体。现在：
  - 字形（图像蒙版）与紧贴的分数线、向量下划线按位置聚成墨迹块；同一栏里排满整行（或从栏左边缘起、段落续行、成段文字）的正文行中的墨迹块是**行内公式**，作为行内图片插回原行、与原基线对齐，必要时在原位置把文字拆开；其余是**独立公式**，上下相邻的合成一块，连同其中的 "if"、"otherwise"、公式编号整体成图（成段的正文绝不并进公式图片）；
  - 求和号的上下限、重音、分子分母跟随紧贴的字形归行，不会被拆到相邻的行；
  - 行内公式图片放进固定行距的一行（只调整升降，略高的等比缩小），不撑高行、也不被 Word 的固定行距裁掉；逐行一致模式下一页的内容与原文一一对应，不会把栏底的行挤到下一页；
  - Type 3 字体的文字：只用于公式时按原样渲染成图片，正文也用 Type 3 时保留编码正常的文字（换成正文字体），只把乱码字符成图；
  - 冗余：位图字形不依赖画布，由蒙版直接合成 PNG；画布不可用或渲染失败时公式与符号仍以图片保留（只有 Type 3 矢量字形需要画布）。渲染时提供 pdf.js 的标准字体与 CMap 数据，PDF 没有嵌入的 Times / Helvetica、中日韩编码也能正确画进图里。
- **出版社处理过的 PDF**：表格文字被转成矢量轮廓（没有可提取的文字）时，整张表按原样渲染成图片，内容不丢失；MathTime 等出版社数学字体（MTSYN / RMTMI / BLEX）排的公式、变量用正文斜体的公式都能识别成整块公式；行内公式里与字母叠放的上下标、重音（F̂）并回正文行，不再单独占一行撑大行距；Times LT Std、Formata 等定制字体映射到 Word 必有的同类字体；表题与表格、图片与图题保持同页。
- **重新转换保留修改**：转换后润色过、又需要用改进后的转换器重新转换时，之前的段落修改会自动搬到新文档上（只搬真正的文字修改，转换器本身带来的差异不会被带回）。
- **首字下沉、标题、破折号**：首字母的基线常与第二行正文对齐、被并进同一行（"T rapid disintegration…"），现在也识别为首字下沉；多行居中的标题以正文为准判断居中（版心被页码撑宽也不会误判成两端对齐）；参考文献里画成短横线的 "——" 还原为文字；固定行距的段落里有比行距高的字时改用"最小值"行距，避免 Word 裁掉字的顶部。
- **图文并排**：照片框、表格旁边并排有文字时（如作者简介），表格改为浮动定位，旁边文字保持原位。文字环绕的图片（作者照片、首字下沉）会按原 PDF 中图片与旁边文字的实测距离写入环绕间距（右 / 左 / 下），而不是 Word 默认的 2pt，文字不再贴着照片。
- **标题**：按视觉样式识别标题层级，转换后的文档有大纲，Agent 可以按章节定位。

两种排版方式（卡片默认"自动"：下载 Word 文件用逐页保留，转换并继续编辑用可编辑排版）：

| 方式 | 适合 | 说明 |
| --- | --- | --- |
| 逐页保留原排版（exact） | 只下载、只改个别字词、需要与原件逐行对照 | 每一行以换行结束，每页固定分页分栏，行、页与 PDF 完全对应；逐行做"不折行"校验。**增加文字会挤乱版面**，不适合扩写 |
| 可编辑排版（flow） | 扩写、改写、增删段落（转换并继续编辑、Agent 自动转换的默认方式） | 字体、字号、行距、缩进、分栏与原文一致；段落自动换行，跨栏、跨页的段落合并为一段（行尾排版断词的连字符去掉，开启自动断词）；页眉页脚写成真正的页眉页脚（页码为 PAGE 域，首页 / 奇偶页不同自动识别），首页栏底的作者单位等注释放进首页页脚；正文中的横线变为段落边框。文字增减时整篇自动重排，不会错乱 |

已经用"逐页保留"方式转换、又需要扩写时，在对话里说一声即可：Agent 会调用 `doc_convert_to_word(layout="flow")` 从原 PDF 重新转换（之前做过的润色、改写会逐段自动搬到新文档上，没能搬运的会列出来；旧版本可回滚）。对于下载后再上传的逐页保留版 Word，Agent 会提示你上传原 PDF。

"快速预览"中每页下方标注"第 n 页 / 共 N 页"，工具栏显示当前滚动到的页；页眉页脚里的页码域（PAGE / NUMPAGES）显示为实际页码。"快速预览"会按 Word 的规则重新分页：每页正文高度固定，两栏先排满左栏再排右栏（栏底对齐），节的最后一页两栏平衡，放不下的段落按行拆到下一页，页眉页脚按首页 / 奇偶页取对应的一份。扩写、删改后预览立刻按新内容重新分页，与 Word / "精确版式"基本一致（个别行尾可能因浏览器字宽差异不同）。

快速预览用 docx-preview 渲染，该库有几处与 Word 不一致的行为，预览端做了修正（`web/src/paginate.ts`）：

- **“最小值”行距**（`lineRule="atLeast"`）：docx-preview 把它渲染成“字号 + N pt”，行距约为 Word 的 1.9 倍，与固定行距、单倍行距的段落排在一起时松紧不一。现在改写成 `max(N pt, 1.15em)`，即“至少 N 磅，字体自然行高更大时取自然行高”。
- **文字环绕图片的间距**：docx-preview 把环绕图片渲染成浮动图片，却忽略 `distR / distB`，文字紧贴图片。现在按文档里写的间距加上；旧版转换出的文档只有默认的 2pt 时，预览至少留 6pt。
- **按页面坐标定位的图片 / 文本框 / 图形**：docx-preview 不支持（图片画在所在行内，文本框和直线、矩形根本不画）。逐页保留原排版的文档里首字下沉、页眉页脚文字、插图、分隔线都是这样放的，以前在预览里会错位或消失。现在渲染前把它们换成 docx-preview 能渲染的形式并做标记，渲染后按原坐标绝对定位到所在页面上，不参与分页。
- **以手动换行结束的两端对齐行**：浏览器不对齐，Word 默认会对齐；逐页保留原排版的文档每行都以换行结束，以前整页行尾参差不齐。现在这些行两端对齐（段落末行仍左对齐），文档打开了 `doNotExpandShiftReturn` 兼容选项时不处理。
- **段前分页、与下段同页、段中不分页**：docx-preview 只认段落样式里的段前分页，段落上直接设置的（Agent 调整分页时改的就是这种）被忽略，"与下段同页""段中不分页"完全不支持，以前 Agent 调整分页后快速预览不变、只有精确版式正确。现在渲染前给这些段落做标记（段落属性优先于样式，含 basedOn 继承），重新分页时按 Word 的规则处理：段前分页从新页开始；整段移到下一页时，前面紧挨着的"与下段同页"段落随它一起走（整页都是这类段落时不移）；"段中不分页"的段落不拆开。
- **固定行距的行被撑高**：CSS 中段落字号形成的"支柱"与大字号文字、上下标会把行框撑得比固定行距高，逐行累积后正文越往下越比 PDF 低。现在固定行距的段落按段内最大字号对齐、上下标不参与撑高，各行位置与 Word 一致。

每次转换都会做**文字完整性校验**（逐字符比对 PDF 与生成的 Word），结果显示在新会话的开头。实测（用 LibreOffice 渲染后与原 PDF 逐像素比对墨迹重合度）：含流程图与 17 个公式的 IEEE Access 双栏论文 12 页 → 12 页、平均 96%；IEEE 双栏论文模板 9 页 → 9 页、平均 96%；中文表单 96%；双栏测试文档 99%；文字完整性均为 100%。

也可以选择其他引擎：`pdf2docx`（需要 `pip install pdf2docx`，有时对无框线表格更好）、LibreOffice（每行一个文本框，版面接近但不便编辑）。命令行直接转换：`npm run cli -- 表单.pdf --output 表单.docx`。

Markdown 和 HTML 采用同样的思路：直接在源文本上按位置拼接，保留原有的换行符（CRLF/LF）、实体写法、属性和注释。

实测结果（IEEE Access 模板，10 类修改，覆盖普通模式和修订模式）：

- 通过 OOXML Schema 校验和修订完整性校验；
- LibreOffice 渲染后，除修改处外版式完全一致；
- 未修改时往返输出字节一致；
- 38 个部件中 34 个字节一致。

输入 `/verify` 可以随时查看当前文档的保真报告。

---

## 编辑、排版与绘图工具

除了改文字、改格式、插入/删除段落、表格加行、批注之外，Agent 还可以使用下面这些工具（只在当前文档支持时出现）。Word 中全部生成**原生对象**，在 Word 里可以继续编辑；Markdown / HTML 生成等价的源码片段。

| 类别 | 工具 | Word 中的效果 | Markdown / HTML |
| --- | --- | --- | --- |
| 表格 | `doc_insert_table` 新建表格（网格线 / 三线表 / 无边框 / 隔行底纹、题注、列宽） | 原生表格，表头跨页重复 | 管道表格 / `<table>` |
| | `doc_edit_table` 删除行列、插入列、合并单元格、底纹、对齐、边框、宽度、删除表格 | 原生表格结构 | — |
| 图片 | `doc_insert_image` 插入你上传的图片（输入框上方「＋ 图片」按钮） | 嵌入图片 + 图题 | 内嵌 data URI |
| 绘图 | `doc_draw` 矩形、圆角矩形、椭圆、菱形、三角形、平行四边形、六边形、圆柱、云形、直线、箭头、文字框任意组合 | 可编辑的形状组合 | 内嵌 SVG |
| | `doc_insert_diagram` 给出节点和连线，自动分层排版成流程图 / 结构图 / 架构图 | 可编辑的形状组合 | 内嵌 SVG |
| 图表 | `doc_insert_chart` 柱状、条形、折线、饼图、环形、面积、散点（可堆叠、数据标签） | 原生图表，数据内嵌 Excel，可右键"编辑数据" | 内嵌 SVG |
| 公式 | `doc_insert_equation` LaTeX → Word 公式（分式、根式、上下标、求和积分、括号、希腊字母…），行间公式可带编号且编号与公式同行 | 原生公式（OMML） | `$$…$$` / `\[…\]` |
| 版面 | `doc_insert_break` 分页符 / 分栏符 / 分节符；`doc_page_setup` 纸张、横向、边距、分栏、行号 | 原生分节 | 分页符 |
| | `doc_header_footer` 页眉页脚，`{PAGE}` `{NUMPAGES}` 自动页码；`doc_insert_toc` 自动目录 | 页码域、TOC 域 | 目录（Markdown） |
| 样式 | `doc_modify_style` 修改 / 新建样式（统一全文格式最规范的方式）；`doc_set_list` 项目符号与多级编号 | 样式表、编号定义 | — |
| 引用 | `doc_insert_footnote` 脚注 / 尾注；`doc_insert_link` 超链接 | 原生脚注、超链接 | Markdown 链接 |
| 审阅 | `doc_review_changes` 接受 / 拒绝修订；`doc_comments` 查看 / 删除批注；`doc_move_blocks` 移动段落 | — | 移动块（Markdown） |
| 期刊排版 | `doc_journal_layout` 通栏图表放页顶、图表不跨页、页面不留白、两栏逐行对齐（见下节） | 分节、段距、浮动体位置 | — |

形状组合与图表另外带一张 PNG 后备图（`mc:AlternateContent`）：Word / WPS / LibreOffice 显示可编辑对象，浏览器内的"快速预览"显示后备图。修订模式下插入的表格、图片、图表、公式、目录都记录为修订，可以在 Word 中整体拒绝。

## 学术期刊排版（journal-layout 技能 + doc_journal_layout 工具）

对 Agent 说"按期刊格式排版""图表不要跨页""页面不要留白""两栏底部对齐"，它会加载 `journal-layout` 技能并调用 `doc_journal_layout`，按 LaTeX 期刊模板的版面规则整体调整（只改版面，不改文字、字体、字号）：

| 规则 | 做法 |
| --- | --- |
| 大图表占满整行 | 宽于一栏的图、表改为通栏（单栏分节、居中），放在页面**顶部**；不放第一页 |
| 栏内图表 | 当前栏剩余空间放得下就紧跟引用段落，放不下移到**下一栏顶部**，文字继续排满当前栏 |
| 不在页首页尾被打断 | 表格每行不跨页、整表与表题同页（`cantSplit` + 与下段同页），图与图题不分离 |
| 不留大面积空白 | 浮动体按实际分页放到页顶 / 栏顶；连续分节前不再平衡分栏（`noColumnBalance`），页顶通栏图表之前的那一页排满 |
| 顶部底部平齐 | 基线网格：正文固定行距；段距、标题、公式、栏内图表高度凑成行距整数倍，左右栏逐行对齐；关闭孤行控制 |

定位依据 LibreOffice 的实际分页：把文档渲染成 PDF，按文字顺序把每个词对应回文档位置，找到"下一页顶部""下一栏顶部"；那里恰在段落中间时把段落拆成两段（前半段末行两端对齐、后半段不缩进，看起来与连续段落一致）。每个图表渲染一次（约 3 秒），最后再检查一遍，报告仍有留白、两栏未对齐的页。没有 LibreOffice 时只做通栏、不跨页与网格对齐。

修改内容后可以**反复运行**：每次先还原上一次的排版（拆开的段落合并、插入的分节去掉、图表回到引用位置之后），再重新排。导出 PDF 与排版使用同一渲染引擎，版面一致；Word 中因字体度量不同，个别图表可能相差几行。快速预览也遵循同样的规则（通栏图表放不下时推迟到下一页顶部，文字继续排满本页）。

题注须以 `Fig. 1:` / `Figure 1.` / `TABLE I` / `图 1` / `表 1` 开头才会被识别为图表。早期转换结果中并进正文段落末尾的表题会被自动拆成独立段落。

## 深度思考

输入框上方的「深度思考」按钮控制本会话是否开启 DeepSeek 思考模式（也可以输入 `/think on|off`）：开启时模型先思考再回答，复杂修改更可靠，但更慢、消耗更多 tokens。思考过程显示在每条回复上方，点击即可展开（思考进行中自动展开），刷新页面后仍然保留。

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
  tools/editingTools.ts   编辑 / 排版 / 绘图工具（表格、图片、形状、流程图、图表、公式、分节、页面、页眉页脚、目录、编号、脚注、链接、样式、修订、批注、移动）
  tools/markup.ts         Markdown / HTML 的表格、图片、SVG 图形与图表标记
  tools/metaTools.ts      todo_write / exit_plan_mode / skill / agent
  services/fallbackImages.ts  形状与图表的 PNG 后备图（快速预览用）
  services/pdfConvert/    PDF → Word 版面重建（extract 提取 / fonts 字体映射 / layout 重建 / raster 矢量图与公式识别 / render 区域渲染 / docxWriter 生成）
  documents/              文档适配器（策略模式）
    docx/                 原生 OOXML 引擎（zip / xml / ooxml / build 构件生成 / DocxDocument）
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
- 使用 pdf2docx 引擎时，若 Python 不在 PATH 中：`PYTHON_PATH=C:\Python312\python.exe`（内置引擎不需要 Python）。

## 已知限制

- **PDF 需先转换为 Word 才能修改**（会话内一键完成）。可编辑排版（流式）下由 Word 重新断行，个别行尾、分页位置会与 PDF 略有不同；页面上的装饰色块只保留首页的。转换是版面重建：文字 100% 保留、版面高度还原；矢量图与行间公式以高清图片还原（不是可编辑对象），旋转文字不转换；扫描件（整页图片）需要先做 OCR。有原始 Word 文件时，直接上传 Word 效果最好。
- **公式工具支持常用 LaTeX 子集。** 矩阵、多行对齐（align）等复杂环境暂不支持，可拆成多个公式。
- **不支持 .doc。** 请先在 Word 中另存为 .docx。
- **跨格式导出（Word → HTML/Markdown）必然有损。** 需要完全保真请导出原格式。
- **“快速预览”由浏览器渲染，与 Word 在分页和字体度量上可能略有差异。** 以“精确版式”（LibreOffice）或 Word 为准。下载的文件本身不受影响。快速预览不能显示原生图表与形状本身，显示的是它们的后备图片；公式在快速预览中不显示。
- **PDF 转换中的公式是图片。** 没有文字编码的公式与符号（Type 3 字体）、行间公式以图片还原，可以移动、缩放，但不能逐字修改；Type 3 矢量字形需要画布（`@napi-rs/canvas`）才能渲染，画布不可用时这类字形会缺失（位图字形不受影响）。
- 修改跨越图片、公式、域代码等对象时会被拒绝，这是保护机制。Agent 会改为分段修改。
