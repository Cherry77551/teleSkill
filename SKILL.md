---
name: academic-writer
description: 学术写作全能助手。当用户要梳理文献、生成数据图表、排版参考文献或润色论文语言时使用；接管科研中最耗时的案头工作，覆盖从灵感激发到规范成稿的闭环。
---

# 学术全能

科研写作里耗时又易错的是**案头工作**——文献、图表、引用、措辞。这个技能接管它们，让人专注**创新**本身。

先判断用户在哪个 **branch**，把对应的 `references/*.md` 加载进来。一次可能跨多个 branch，按依赖顺序处理：**文献 → 图表 → 引用 → 润色**。

## 路由

- 手里有一堆文献，要理出头绪 → 读 [`references/literature-review.md`](references/literature-review.md)
- 有数据，要出图或看趋势 → 读 [`references/data-viz.md`](references/data-viz.md)
- 有引用，要接投稿格式 → 读 [`references/reference-format.md`](references/reference-format.md)
- 稿子写完了，要改语言 → 读 [`references/polish.md`](references/polish.md)
- 完全不知道从哪开始 → 先问：目标（学位论文 / 期刊投稿 / 开题）和研究问题一句话是什么。

## 通用纪律

四个模块共用同一条底线：

- **不编造。** 文献、数据、引用一律可溯源；查不到就标 `[待补来源]`，绝不虚构。
- **先对齐再动手。** 用途、学科、目标格式不确定时，一次问一个问题。
- **交付可复用的产物**：提纲、图、参考文献表、润色对照，直接能贴进稿子。
- 术语以 [`CONTEXT.md`](CONTEXT.md) 为准。
