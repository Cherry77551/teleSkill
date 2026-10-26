# 示例 07 · 英文论文（语言判定 → 英文基线 → 投稿文书）

> 场景：用户说"帮我写一篇投 IEEE Transactions on Smart Grid 的论文"，未明说语言。

## 第一步：语言判定（不单独发问）

- 信号：指明目标国际期刊（IEEE Trans）→ 判**英文**（信号表见 `references/english-writing.md`）。
- 不弹"用中文还是英文"的语言题；判定结果进汇总复述。

## 第二步：需求确认要点（英文模式）

- 其余六题照常（输出格式 / 摘要 / 附带要素 / 结构选方案 F（IMRaD）/ 篇幅 / 资料来源）；
- 排版规格推荐值切换英文基线：Times New Roman 12pt、数字编号 1/1.1、行距按期刊；
- 引用格式推荐 IEEE（`format_refs.py --format ieee`）。

## 汇总复述（点头后才动笔）

> 已确认：Word 文档；仅英文摘要；含表格（三线表）、公式；**本次按英文写作执行**（依据：检测到目标期刊 IEEE Trans）；结构采用方案 F（IMRaD）；篇幅约 8 页双栏；参考文献 IEEE 格式，由我联网搜集并逐条标注来源。预计完成时间约 2 小时。以上还有什么需要修改的地方吗？没有的话请点个头，我将开始写。

## 交付前校验

`python scripts/check_format.py paper.docx --lang en --strict` → 直引号合法、`1 Introduction` 数字编号、References 一条一行；公式 OMML、伪下标、纯黑照常校验。

## 附带产出 · Cover Letter（骨架示例）

> Dear Editor-in-Chief,
> We are submitting our manuscript titled "Deep Learning for Power Scheduling" (Article) for consideration in IEEE Transactions on Smart Grid. The paper proposes … and reports a 12% cost reduction on two public benchmarks. We believe it fits the journal's scope on …
> We confirm that this work is original, has not been published elsewhere, and is not under consideration by another journal. All authors have approved the submission and declare no conflict of interest.
> Sincerely, … (通讯作者)

主编姓名未提供 → 留 `[待补]`，不编造。
