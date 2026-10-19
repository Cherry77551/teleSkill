# 学术全能 · academic-writer

面向科研写作的 TeleAgent Skill。把文献梳理、数据图表、参考文献排版、语言润色这四类**案头工作**接管过来，让科研人员专注**创新**本身，实现从灵感激发到规范成稿的闭环。

适用学科：通用理科 / 工科（不绑定具体领域）。

## 技能结构

```
academic-writer/
├── SKILL.md          # 必需：元数据 + 路由 + 需求确认（选择题交互）+ 交付纪律
├── CONTEXT.md        # 术语表
├── references/       # 按需加载
│   ├── literature-review.md    文献梳理
│   ├── data-viz.md             数据分析与可视化
│   ├── reference-format.md     参考文献排版（GB/T 7714、上标引用）
│   ├── polish.md               语言润色
│   ├── human-writing.md        降重降 AI 率语言规范（写正文与润色都执行）
│   ├── manuscript-format.md    排版基线 / 公式与符号规范 / 交付前校验
│   ├── template-reuse.md       学校 / 期刊模板套用
│   ├── revision.md             迭代改稿（修订闭环）
│   ├── english-writing.md      英文论文：语言判定、英文基线、降 AI 率英文规则、Cover Letter / 审稿回复
│   └── docx.md                 Word 读写 / OMML 公式 / 渲染校验 / 效率纪律
├── scripts/          # 可执行脚本
│   ├── check_format.py         .docx 排版自动校验（纯标准库；--strict 完整模式；--lang en 英文口径；--template 模板基准）
│   ├── check_style.py          段长错落自检（纯标准库）
│   ├── eval_check.py           自动化回归检查（结构完整性 + 四种引用格式冒烟 + 语言口径冒烟；--docx/--text 可挂校验）
│   ├── format_refs.py          GB/T 7714 / APA / IEEE / Nature 批量排版（--format 切换）
│   ├── plot.py                 参数化绘图（bar/line/scatter/box/hist）
│   ├── omml.py                 LaTeX → OMML 公式生成（角标结构确定性正确）
│   ├── render_check.ps1        docx 渲染成图逐页校验（中间产物；-UseWord 走 Word COM）
│   └── refs.example.csv
├── assets/{templates,checklists,examples}   # examples 含需求确认与降重改写示例
└── evals/            # 回归用例（24 条）
    └── eval-cases.md
```

## 怎么用

1. 把整个 `academic-writer/` 打包为 `.zip`，通过**技能广场**导入 → 安全审核 → 自动加载。
2. 直接说需求，例如："帮我把这 20 篇 PDF 梳理成一篇综述框架。"
3. 技能自己判断该走哪个模块（文献 / 图表 / 引用 / 排版 / 润色）；不确定时先问目标和研究问题。

## 设计要点

- **共享纪律**：不编造、一手优先、无来源不写、先对齐再动手。
- **需求确认**：可点击选择题逐项确认（格式 / 摘要 / 要素 / 规格 / 结构 / 篇幅 / 来源），汇总复述 → 预估 → 等用户点头；确认的格式即唯一交付格式。
- **排版基线**：定义见 `references/manuscript-format.md`（宋体小四 / 黑体标题 / 纯黑 / 标题居中 / 目录独占一页 / 上标引用 / OMML 公式 / 真实下标 / 默认全线表格）。
- **降重降 AI 率**：正文与润色均执行 `human-writing.md`，段长句长量化错落。
- **成稿校验**：`check_format.py` + `check_style.py` 程序化校验，渲染抽样复核；润色与审阅用修订 / 批注交付。
- **可评测**：`evals/eval-cases.md` 二十四条回归用例；可脚本验证的部分（结构完整性、四种引用格式冒烟、语言口径冒烟，`--docx`/`--text` 可挂排版与段长校验）由 `scripts/eval_check.py` 一键回归。
- **中英双语**：语言判定不设语言题（信号判英文/默认中文），英文基线见 `references/english-writing.md`，英文口径校验 `check_format.py --lang en`；英文投稿文书（Cover Letter / 审稿回复）一并覆盖。

## 交付

- `队伍名+作品名.zip`（技能包）
- 说明文档与演示视频按需提供
