---
name: ieee-paper
description: IEEE 期刊/会议论文（含 IEEE Access 模板）的写作与排版规范，以及在该模板上安全编辑的方法
when_to_use: 文档被检测为 IEEE 模板，或用户提到 IEEE、IEEE Access、Transactions、投稿格式、论文排版时
argument_hint: "[任务说明]"
allowed_tools: []
---
# IEEE 论文模板编辑规范

## 一、模板结构（IEEE Access 模板的样式名）
IEEE 模板不用 Word 内置的"标题 1/2/3"，而是自己的样式。插入或调整段落时必须使用这些样式，不要自造格式：

| 内容 | 样式名 |
|---|---|
| 论文标题 | Paper Title |
| 作者行 | AU |
| 作者单位 | PI / PI_No Space |
| 摘要段 | Abstract（开头 "ABSTRACT" 为字符样式加粗） |
| 关键词段 | IT（开头 "INDEX TERMS"） |
| 一级标题 | H1_List (Space)（章节首个用 H1_List (No Space)），自动罗马数字编号 I. II. III. |
| 二级标题 | H2_First / H2_Cont / H2_After H1，自动编号 A. B. C. |
| 三级标题 | H3，自动编号 1) 2) 3) |
| 正文 | PARA（段首）/ PARA_Indent |
| 图题 | Fig Caption / Figure Caption |
| 表题 | Table Title（TABLE I 形式，全大写罗马数字） |
| 参考文献 | REF Txt / References |
| 作者简介 | AU Bios |

先用 doc_styles 确认本文档中实际存在的样式名（不同版本模板略有差异），再使用。

## 二、写作规范（检查与修改的依据）
1. 标题：首字母大写的标题式大小写（Title Case），不全大写；避免在标题中写长公式；不写 "(Invited)"。
2. 作者行：全名，作者之间用逗号，最后一位前加 "and"；单位编号用上标数字（保持上标格式不变）。
3. 摘要：单段、150–250 词，不含缩写定义以外的缩写、脚注、参考文献、公式、表格；包含 3–4 个关键词。
4. 关键词（INDEX TERMS）：按字母顺序，逗号分隔，优先使用 IEEE Thesaurus 标准术语。
5. 缩写：正文首次出现时给出全称（即使摘要中已定义）；IEEE、SI、ac、dc 等常用缩写无需定义。
6. 单位：使用 SI 单位；数值与单位间留空格（"5 cm"）；小数点前补零（"0.25"）；范围写 "7–9" 或 "7 to 9"。
7. 图表：正文中引用写 "Fig. 1"（句首写 "Figure 1"）、"Table I"；图题以 "Fig. 1." 开头。
8. 公式：编号右对齐 "(1)"，正文引用写 "(1)" 或 "Eq. (1)"（句首）。
9. 参考文献：按引用顺序编号 [1]、[2]；句中引用 "in [3]"，不写 "in ref. [3]"；标点在括号外。
10. 美式英语：拼写（color / modeling）、引号内标点（"like this."）、使用序列逗号（A, B, and C）。
11. 不用缩约形式（don't → do not）；强调用斜体，不用下划线。

## 三、在本模板上安全编辑
- 标题编号是样式自动生成的（大纲中显示"自动编号"），不要在标题文字里手写 "I." "A." 之类的编号。
- 修改作者、单位时只替换文字，上标数字在 old_text/new_text 中保持原位，系统会保留上标格式。
- 新增一级章节：doc_insert_blocks，style="H1_List (Space)"，系统会自动套用文档中已有同级标题的格式并延续编号；其后正文用 style="PARA"。
- 首页有多个分节符（单栏标题区 / 双栏正文区），携带分节符的段落不可删除。
- 表格中的 ⟨符号→⟩⟨符号−⟩ 等是 Symbol 字体符号，修改附近文字时原样保留。
- 投稿前建议：开启修订模式（/track on）完成修改，让作者在 Word 中逐条确认。

## 四、推荐流程
1. doc_outline → 确认模板、章节结构；
2. 按任务读取相关章节（doc_read，按大纲引用定位）；
3. 对照第二部分规范列出问题（多项时用 todo_write）；
4. 逐项修改或批注；
5. doc_verify 确认只改动了预期段落。
