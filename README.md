# 学术全能 · academic-writer

面向科研写作的 TeleAgent Skill。把文献梳理、数据图表、参考文献排版、语言润色这四类**案头工作**接管过来，让科研人员专注**创新**本身，实现从灵感激发到规范成稿的闭环。

适用学科：通用理科 / 工科（不绑定具体领域）。

## 技能结构

```
academic-writer/
├── SKILL.md          # 必需：元数据 + 路由 + 通用纪律
├── CONTEXT.md        # 术语表
├── references/       # 四个模块，按需加载
│   ├── literature-review.md    文献梳理
│   ├── data-viz.md             数据分析与可视化
│   ├── reference-format.md     参考文献排版
│   └── polish.md               语言润色
├── scripts/          # 可执行脚本
│   ├── format_refs.py          GB/T 7714 批量排版
│   └── refs.example.csv
└── assets/{templates,checklists}
```

## 怎么用

1. 把整个 `academic-writer/` 打包为 `.zip`，通过**技能广场**导入 → 安全审核 → 自动加载。
2. 直接说需求，例如："帮我把这 20 篇 PDF 梳理成一篇综述框架。"
3. 技能自己判断该走哪个模块（文献 / 图表 / 引用 / 润色）；不确定时先问目标和研究问题。

## 共享纪律

四个模块共用同一条底线：**不编造**、**可溯源**、**先对齐再动手**。模块之间用**转交**打通成闭环：`literature-review → reference-format → polish`。

## 交付

- `队伍名+作品名.zip`（技能包）
- `队伍名+作品名.docx`（技能说明文档）
- `队伍名+作品名.mp4`（演示视频）
