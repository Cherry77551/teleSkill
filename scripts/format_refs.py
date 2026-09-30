#!/usr/bin/env python3
"""把结构化参考文献字段格式化为 GB/T 7714 著录条目。

用法:
    python format_refs.py refs.csv

输入 CSV 表头（顺序不限，缺失列按空处理）:
    type,authors,title,source,year,volume,issue,pages,doi
    - type:   J(期刊) / M(专著) / D(学位论文) / C(会议)
    - authors: 用 ';' 分隔，例如 "张三;李四;王五;赵六"

输出: 每行一条 GB/T 7714 参考文献，按输入顺序编号。
说明: 本脚本只做确定性排版，不补全字段；缺失即留空，绝不编造。
"""
import csv
import sys

MARK = {"J": "[J]", "M": "[M]", "D": "[D]", "C": "[C]"}


def fmt_authors(authors):
    names = [a.strip() for a in authors.split(";") if a.strip()]
    if len(names) > 3:
        names = names[:3] + ["等"]
    return ", ".join(names)


def fmt_entry(i, row):
    t = row.get("type", "").strip().upper()
    mark = MARK.get(t, "")
    authors = fmt_authors(row.get("authors", ""))
    title = row.get("title", "").strip()
    source = row.get("source", "").strip()
    year = row.get("year", "").strip()
    volume = row.get("volume", "").strip()
    issue = row.get("issue", "").strip()
    pages = row.get("pages", "").strip()
    doi = row.get("doi", "").strip()

    head = f"[{i}] {authors}. {title}{mark}."

    pub = source
    if year:
        pub += (", " if pub else "") + year
    vol_issue = volume
    if issue:
        vol_issue += f"({issue})"
    if vol_issue:
        pub += (", " if pub else "") + vol_issue
    if pages:
        pub += ": " + pages

    parts = [head]
    if pub:
        parts.append(pub + ".")
    if doi:
        parts.append(f"DOI: {doi}.")
    return " ".join(parts)


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "refs.csv"
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    for i, row in enumerate(rows, 1):
        print(fmt_entry(i, row))


if __name__ == "__main__":
    main()
