---
name: academic-writer
name_cn: 学术全能
description: 面向校园师生的学术写作全能助手。当师生要写毕业论文/学位论文、开题报告、课程论文、期刊投稿，或梳理文献、生成数据图表、排版参考文献、润色论文语言时使用；只依据可核验的一手文献写作，接手校园科研中最耗时的案头工作，覆盖从开题到规范成稿的闭环。英文需求（thesis / dissertation / paper writing / literature review / reference formatting）同样触发。
description_cn: 覆盖文献梳理、数据图表、参考文献排版、语言润色、整篇排版、模板复用与迭代改稿的学术写作技能包，从开题到规范成稿闭环。
---

# 学术全能

## 描述

校园里师生最耗时又易错的，是论文写作中的**案头工作**——文献、图表、引用、措辞。这个技能接管它们，让师生专注**研究本身**，解决从开题、成稿到投稿的校园真实需求。

先判断用户在哪个 **branch**，把对应的 `references/*.md` 加载进来。一次可能跨多个 branch，按依赖顺序处理：**文献 → 图表 → 引用 → 润色**。整篇排版与 Word 处理由两个共享模块支持。

## 使用场景

当师生的需求命中以下任一校园场景时触发本技能：

- **毕业论文 / 学位论文**：从开题文献梳理到整篇规范排版、参考文献格式（GB/T 7714）一条龙
- **开题报告**：研究问题界定、文献综述框架、检索式与待检索清单
- **课程论文 / 期末论文**：快速梳理文献、出图、统一引用格式
- **期刊 / 会议投稿**：按目标期刊规范排版、参考文献格式（APA / IEEE / Nature）
- **英文论文 / SCI 投稿**：英文正文与排版、Cover Letter、逐条回复审稿意见（Response to Reviewers）
- **导师带学生**：组会汇报图表、逐句语言润色、修订稿批注交付
- 有数据，要出图、看趋势或做表格 → 数据图表
- 有一批引用，要按投稿格式统一排版 → 参考文献排版
- 稿子写完了，要改语言、让表达更学术 → 语言润色
- 要整篇排版、指定输出格式（Word / PDF / 其他）与附带要素（表格、目录、公式、封面、附录）→ 整篇排版
- 涉及 `.docx` 的读写、批注或修订 → Word 处理

## 指令

### 第一步：路由到模块

按用户需求加载对应模块文档：

- 文献梳理 → 读 [`references/literature-review.md`](references/literature-review.md)
- 数据图表 → 读 [`references/data-viz.md`](references/data-viz.md)
- 参考文献排版 → 读 [`references/reference-format.md`](references/reference-format.md)
- 语言润色 → 读 [`references/polish.md`](references/polish.md)
- 正文写作语言规范、降重 / 降 AI 率 → 读 [`references/human-writing.md`](references/human-writing.md)
- 整篇排版、输出格式、附带要素 → 读 [`references/manuscript-format.md`](references/manuscript-format.md)
- Word 读写、批注、修订、渲染校验 → 读 [`references/docx.md`](references/docx.md)
- 模板套用（学校/期刊模板） → 读 [`references/template-reuse.md`](references/template-reuse.md)
- 迭代改稿、按导师/审稿意见修订 → 读 [`references/revision.md`](references/revision.md)
- 英文论文 / SCI 投稿、Cover Letter、审稿回复、语言判定 → 读 [`references/english-writing.md`](references/english-writing.md)
- 完全不知道从哪开始 → 先问：目标（学位论文 / 期刊投稿 / 开题）和研究问题一句话是什么。

### 第二步：引导确认需求，对齐后再开工（未确认前禁止动笔）

用户可能只给了主题或不完善的信息，此时**绝对不要急着写**。按 [`references/requirement-confirmation.md`](references/requirement-confirmation.md) 完成**七项确认**（输出格式、摘要、附带要素、排版规格、论文结构、篇幅、资料来源）与**三步收尾**（汇总复述 → 预估完成时间 → 等用户点头）。

**硬约束（不论平台是否提供提问工具都成立）**：

- 优先用平台提问工具做成**可点击选择题**；无提问工具时才退化为文本编号选项，并注明"也可直接回复您的自定义要求"。
- 每项写明"若不选择或无要求，推荐按您使用 XX"；**禁止**孤立"（默认）"、**禁止**"您习惯使用"的口吻。
- ❌ 开放式填空（"篇幅多少页？"）= 不合格；需要数值的项先让用户点选、再报数值或范围。
- **模板先行**：先问是否有学校 / 期刊模板 docx；有则按 [`references/template-reuse.md`](references/template-reuse.md) 解析，模板已定义项改为"按模板执行 / 改用其他"的确认题，不让用户重新盲选。

**语言**按 [`references/english-writing.md`](references/english-writing.md)「语言判定」内部判定，不单独提问。

**未得到用户明确肯定（点头 / 可以 / 开始）前，一个字都不要写。**

### 第三步：按模块执行（工具调用规则）

按需调用外部能力，触发条件明确，不滥用：

**代码执行（Python / Shell）** —— 每条按「任务（触发条件）：动作；校验」执行，按需调用，不滥用：
- **数据图表**（用户给了结构化数据（CSV/XLSX））：常规图（bar/line/scatter/box/hist）优先调用 `scripts/plot.py` 参数化出图；复杂图（多子图、自定义配色）再手写 matplotlib/seaborn。
- **参考文献排版**（用户给了结构化引用数据）：调用 `scripts/format_refs.py` 按目标格式批量排版（GB/T 7714 默认；`--format apa|ieee|nature` 切换）。
- **公式生成**（公式要进文档）：把公式的 LaTeX 用 `scripts/omml.py` 转成 OMML（`python scripts/omml.py --display "..."`）再插入文档，**不手写 OMML**；校验：生成后用 `--strict` 模式 `check_format.py` 校验角标结构。
- **整篇 Word 生成**（要交付 `.docx`）：优先用 `scripts/build_docx.py` 从 JSON 规格**确定性生成**——`python scripts/build_docx.py spec.json -o 论文.docx --check --strict`，一次产出标题 / 摘要 / 真实 TOC 域 / 章节编号 / 行内与行间公式 / 表格 / 插图 / 参考文献；规格见 [`assets/templates/paper-spec.example.json`](assets/templates/paper-spec.example.json)，字段说明见 [`references/docx.md`](references/docx.md)「确定性生成」。**不要手写整套 OOXML。**
- **Word 渲染校验**（生成 `.docx` 后）：调用 `scripts/render_check.ps1` 转 PDF 并**抽样出图**核对版式——只渲染**公式所在页、目录页、首页/尾页**（用 `-FirstPage/-LastPage` 指定页范围），**不逐页全量渲染**；全篇逐页渲染仅作为交付前一次性可选动作。
- **Word 排版自动校验**（交付前）：运行 `python scripts/check_format.py 论文.docx --strict`（学位论文 / 完整稿件；短文 / 片段用默认模式，公式、上标、目录、一级标题缺失不算错、存在则校验质量）（纯标准库）；**英文稿加 `--lang en`——直引号、数字编号、Contents/References、Figure/Table/Eq. 等检查自动切英文口径**；**提供了模板时加 `--template 模板.docx`——一级标题编号模式、标题对齐、字色、页边距的基准自动改用模板实际值，模板合法偏离不再误报**；检查公式 OMML、伪下标、英文直引号、一级标题、标题对齐、字色、上标、页边距、表格线型、目录；校验：FAIL 项修复后重新生成，再抽样渲染关键页复核。
- **段长自检**（正文写完后）：运行 `python scripts/check_style.py 正文.md/.txt/.docx`（纯标准库），相邻段落字数差应 ≥30%（或 ≥40 字）、短/中/长三档交替、无连续 3 段同档；不合格：合并/拆分重写后重跑，达标再交付。
- **通用原则**（涉及计算/数据处理/文件生成）：一律用脚本而非脑补，确保可复现。
- **脚本路径纪律**：各脚本只读取用户明确指定的输入文件，除 `.temp/` 中间产物外不写其他路径，不删除、不覆盖用户原文件。

**联网搜索**
- 默认不联网：文献、数据、DOI 不主动搜，先问用户有没有。
- 触发条件：用户明确说"帮我搜一下"或选了"需要额外搜集资料"，且平台提供检索工具时才联网。
- 搜到的内容：每条都标注来源 URL/数据库/访问日期，不把搜到的东西直接当一手文献写进正文，只作为"待用户确认的补充资料"。

**示例参考**
- 执行各模块时，先读一下 [`assets/examples/`](assets/examples/) 里对应场景的示例，对齐输出风格再动手。

### 第四步：守纪律，交付可复用产物

所有模块共用同一条底线，按四组纪律执行：

**① 事实纪律（写什么）**
- **不编造。** 作者、年份、期刊、DOI、页码、数值，一律不能凭空写。
- **一手优先。** 事实性论断只依据同行评审论文、会议、学位论文、权威机构报告、标准等**一手文献**；**百科、新闻报道、商业宣传、自媒体一律不作论据**。
- **无来源不写，不交半成品。** 成稿里不得出现 `[待补来源]`、`此处请补充` 之类占位；缺的料**单列成清单**，不要混进正文。
- **能力边界。** 默认不能联网、不能访问文献数据库（除非平台另提供了检索工具）；**不靠记忆补文献、补数据、补 DOI**，想不起来的就空着或列进清单。

**② 质量纪律（写得对）**
- **排版即验收。** 交付前按用户确认的排版规格逐项核对：程序化用 `scripts/check_format.py`，渲染按 [`references/docx.md`](references/docx.md)「交付前渲染看一眼」的检查点清单**抽样复核**；**任何一项不符（如确认宋体小四却整篇黑体）即不合格，返工重排后再交付**。
- **写作执行 [`references/human-writing.md`](references/human-writing.md)（降重、降 AI 率）。** 段落长短、句子长短、标点、词语搭配、语法结构要有变化，清理 AI 高频词、打破总分总固化结构，不追求“最稳定最优”的匀速表达；前提是逻辑清晰、无语病。
- **成稿先给"新读者"看一遍。** 用一个没有上下文的人或空会话只读成稿，能答上关键问题，才算完成。

**③ 交付纪律（交什么）**
- **交付可复用的产物**：提纲、图、参考文献表、润色对照，直接能贴进稿子。
- **交付格式 = 用户确认的格式，禁止擅自变更或附加。** 用户选了 Word 就只交付 `.docx`，**不得再导出或附带 PDF、HTML 等未选格式**；图/表/公式等素材默认**内嵌**唯一交付文档，需要以独立文件（PDF/SVG/PNG）交付时，须在需求确认阶段把"图表独立交付"列入附带要素并确认格式，定稿后不得随附图文件；渲染校验产生的 PDF、逐页图片是**中间产物**，一律放 `.temp/`，校验完不算交付，不进入最终交付清单。

**④ 保密纪律（不外泄）**
- **敏感信息保密。** 模板封面与正文中的个人信息（学号、姓名、导师、单位）与未发表研究数据**仅本地处理**：不写入联网检索式、不上传外部服务；联网搜集仅限文献检索关键词，不携带用户身份信息与未发表数据。

- 术语以 [`CONTEXT.md`](CONTEXT.md) 为准。

## 示例

输入/输出示例见 [`assets/examples/`](assets/examples/)：

- `01-literature-review-no-sources.md` —— 不给文献时先问 + 给检索式 + 等回填
- `02-data-viz.md` —— 先问清再动手，用 `plot.py` 出图
- `03-reference-format.md` —— 用 `format_refs.py` 批量排 GB/T 7714
- `04-polish.md` —— 逐句对照 + 修订稿交付 + 读者测试
- `05-requirement-confirmation.md` —— 需求确认的可点击选择题全流程（含“推荐按您使用”表述与汇总点头）
- `06-human-writing.md` —— 降重降 AI 率改写对照（AI 味原文 → 改后 → 规则说明）
- `07-english-writing.md` —— 英文论文：语言判定 → 英文基线 → Cover Letter 骨架

包结构与导入说明见 [`README.md`](README.md)。
