#!/usr/bin/env python3
"""参数化绘图：读 CSV，按指定图类型出图（矢量优先）。

用法:
    python plot.py data.csv --x 列名 --y 列名 --kind bar -o out.pdf

图类型 kind（对应要传达的关系）:
    bar     比较：x 为类别、y 为数值
    line    趋势：x、y 为数值
    scatter 相关：x、y 为数值
    box     分布：仅需 --y（一列数值）
    hist    分布：仅需 --y（一列数值）

说明: 确定性出图，不截断纵轴、不用 3D/彩虹色阶、不改数据。
     中文标签自动尝试 Noto Sans CJK SC / WenQuanYi / 微软雅黑。
"""
import argparse
import csv

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = [
    "Noto Sans CJK SC", "WenQuanYi Micro Hei",
    "Microsoft YaHei", "SimHei", "DejaVu Sans",
]
plt.rcParams["axes.unicode_minus"] = False


def read_rows(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def to_float(v):
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return None


def main():
    p = argparse.ArgumentParser(description="参数化出图")
    p.add_argument("csv")
    p.add_argument("--x", help="x 轴列名（box/hist 不需要）")
    p.add_argument("--y", required=True, help="y 轴列名")
    p.add_argument("--kind", required=True,
                   choices=["bar", "line", "scatter", "box", "hist"])
    p.add_argument("-o", "--out", default="figure.pdf")
    args = p.parse_args()

    if args.kind in ("bar", "line", "scatter") and not args.x:
        p.error("--kind %s 需要提供 --x 列名" % args.kind)

    rows = read_rows(args.csv)
    fig, ax = plt.subplots()

    if args.kind in ("box", "hist"):
        ys = [v for v in (to_float(r.get(args.y)) for r in rows) if v is not None]
        if args.kind == "box":
            ax.boxplot(ys)
            ax.set_xticklabels([args.y])
        else:
            ax.hist(ys, bins="auto")
        ax.set_ylabel(args.y)
    else:
        xs_raw = [r.get(args.x) for r in rows]
        ys = [to_float(r.get(args.y)) for r in rows]
        xs = [str(x) for x in xs_raw] if args.kind == "bar" else [to_float(x) for x in xs_raw]
        pairs = [(x, y) for x, y in zip(xs, ys) if y is not None and x is not None]
        x, y = [p[0] for p in pairs], [p[1] for p in pairs]

        if args.kind == "bar":
            ax.bar(x, y)
        elif args.kind == "line":
            ax.plot(x, y)
        else:
            ax.scatter(x, y)
        ax.set_xlabel(args.x)
        ax.set_ylabel(args.y)

    fig.tight_layout()
    fig.savefig(args.out)
    print(f"已生成: {args.out}")


if __name__ == "__main__":
    main()
