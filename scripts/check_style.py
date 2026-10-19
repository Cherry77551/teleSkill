#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_style.py —— 文稿段落长度错落检查（纯 Python 标准库，零依赖）

用法:
    python check_style.py 文稿.txt|.md|.docx [--min-ratio 0.3] [--min-abs 40] [--short 120] [--mid 300]

检查项（对应 references/human-writing.md 第 4 条量化标准）:
  1. 相邻两段字数差 >= max(较短段 30% 或比例, 绝对差 40 字)，不足时标记"等长对"
  2. 段长三档分布: 短段(<=120) / 中段(120~300) / 长段(>300)；不允许全部挤在中段
  3. 不允许出现连续 3 段落在同一档

说明: .md 输入自动跳过 Markdown 标题（# 级），代码块（```）与引用行（>），避免干扰段长统计。

退出码: 0 = 段长错落达标；1 = 存在不合格；2 = 文件无法读取
"""

import argparse
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def extract_paragraphs(path):
    """按段落提取文本。.txt/.md 按空行分段；.docx 按 <w:p> 分段。"""
    if path.lower().endswith(".docx"):
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        paras = []
        for elem in root.iter("{%s}p" % W):
            text = "".join(t.text or "" for t in elem.iter("{%s}t" % W))
            paras.append(text.strip())
        return [p for p in paras if p]
    with open(path, "r", encoding="utf-8") as fh:
        text = fh.read()
    kept, in_code = [], False
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        if s.startswith("#") or s.startswith("---") or s.startswith(">"):
            continue
        kept.append(line)
    blocks = re.split(r"\n\s*\n", "\n".join(kept))
    paras = []
    for b in blocks:
        lines = [line.strip() for line in b.splitlines() if line.strip()]
        if lines:
            paras.append(" ".join(lines))
    return paras


def char_count(text):
    return len(re.sub(r"\s", "", text))


def main():
    ap = argparse.ArgumentParser(description="文稿段长错落检查（纯标准库）")
    ap.add_argument("input", help="文稿路径（.txt / .md / .docx）")
    ap.add_argument("--min-ratio", type=float, default=0.3, help="相邻段字数差最小比例，默认 0.3（30%%）")
    ap.add_argument("--min-abs", type=int, default=40, help="相邻段字数差最小绝对值，默认 40")
    ap.add_argument("--short", type=int, default=120, help="短段上限字数，默认 120")
    ap.add_argument("--mid", type=int, default=300, help="中段上限字数，默认 300")
    args = ap.parse_args()

    try:
        paras = extract_paragraphs(args.input)
    except Exception as e:
        print("错误: 无法读取文件 %s：%s" % (args.input, e))
        sys.exit(2)

    sizes = [char_count(p) for p in paras]
    if not sizes:
        print("提示: 未提取到段落")
        sys.exit(0)

    print("段落数: %d" % len(sizes))
    print("每段字数: %s" % sizes)
    print("最短 %d / 最长 %d / 平均 %d" % (min(sizes), max(sizes), sum(sizes) // len(sizes)))

    issues = []
    for i in range(len(sizes) - 1):
        a, b = sizes[i], sizes[i + 1]
        need = max(args.min_abs, int(args.min_ratio * max(a, b)))
        if abs(a - b) < need:
            issues.append("第 %d、%d 段等长：%d 与 %d（差 %d < %d）" % (i + 1, i + 2, a, b, abs(a - b), need))

    short = [i for i, s in enumerate(sizes) if s <= args.short]
    mid = [i for i, s in enumerate(sizes) if args.short < s <= args.mid]
    long_ = [i for i, s in enumerate(sizes) if s > args.mid]
    if mid and len(mid) == len(sizes):
        issues.append("全部 %d 段都落在中段区间（%d~%d 字），需要短段/长段错落" % (len(mid), args.short + 1, args.mid))
    elif not short and not long_:
        issues.append("没有短段也没有长段，长度无三档变化")

    for k in range(len(sizes) - 2):
        seg = sizes[k:k + 3]
        if all(s <= args.short for s in seg) or all(args.short < s <= args.mid for s in seg) or all(s > args.mid for s in seg):
            issues.append("第 %d~%d 段连续 3 段同档（%s）" % (k + 1, k + 3, seg))
            break

    if len(sizes) < 5:
        print("提示: 段落数较少（%d 段），结果仅供参考" % len(sizes))

    if issues:
        print("不合格 %d 处:" % len(issues))
        for it in issues:
            print("  - " + it)
        print("-" * 56)
        print("结果: 段长错落不达标，请合并/拆分相关段落（或调整句长）后重跑")
        sys.exit(1)

    print("-" * 56)
    print("结果: 段长错落达标")
    sys.exit(0)


if __name__ == "__main__":
    main()