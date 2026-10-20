#!/usr/bin/env python3
"""把结构化参考文献字段格式化为 GB/T 7714-2015 / APA 7 / IEEE / Nature 著录条目。

用法:
    python format_refs.py refs.csv
    python format_refs.py refs.csv --format apa
    python format_refs.py refs.bib --dedupe
    python format_refs.py refs.bib --sort year

输入:
    1) CSV（表头顺序不限，缺失列按空处理）:
       type,authors,title,source,year,volume,issue,pages,address,doi,lang
       - type:    J(期刊) / M(专著) / D(学位论文) / C(会议)
       - authors: 用 ';' 分隔，例如 "张三;李四;王五;赵六"
       - address: 出版地/保存地/会议地点（M、D 类型必填；J 类型留空）
       - lang:    可选。zh=中文、en=西文、auto(默认)=按作者字符自动判断，
                  决定作者超过 3 人时补"等"还是"et al."。
    2. BibTeX（.bib）: @article / @book / @inproceedings / @phdthesis / @mastersthesis
       type 映射: article→J、book→M、inproceedings/conference→C、phdthesis/mastersthesis→D
       author 字段内的 " and " 自动转为 ';' 分隔。
       字段值支持花括号（含嵌套）、双引号或裸值（裸写如 year = 2024）三种写法。

选项:
    --format   输出格式: gbt（GB/T 7714，默认）/ apa（APA 7）/ ieee / nature
    --dedupe   按题名去重（保留信息更全的一条）
    --sort     排序方式: year（按年份升序）；默认按输入顺序

输出: 每行一条参考文献，按输入顺序；编号随格式——gbt/ieee 用 [N]、nature 用 N.、
      apa 为作者-年制不编号。作者名一律按输入原样排版，不改写为缩写。
说明: 本脚本只做确定性排版，不补全字段；缺失即留空，绝不编造。
"""
import argparse
import csv
import re
import sys

try:
    from _paths import PathError, resolve_input
except ImportError:  # 作为包导入时补 scripts/ 到 sys.path
    import os as _os
    sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
    from _paths import PathError, resolve_input

MARK = {"J": "[J]", "M": "[M]", "D": "[D]", "C": "[C]"}
BIB_TYPE = {"article": "J", "book": "M", "inproceedings": "C", "conference": "C",
            "phdthesis": "D", "mastersthesis": "D"}


def has_cjk(text):
    """是否含 CJK 汉字（含扩展 A 区），用于自动判断文献语种。"""
    for ch in text:
        o = ord(ch)
        if 0x3400 <= o <= 0x4DBF or 0x4E00 <= o <= 0x9FFF:
            return True
    return False


def detect_lang(authors, lang):
    """返回 'zh' 或 'en'：显式 lang 优先，否则按作者字符判断。"""
    lang = (lang or "").strip().lower()
    if lang in ("zh", "en"):
        return lang
    return "zh" if has_cjk(authors) else "en"


def fmt_authors(authors, lang):
    """作者 >3 时列前 3 名，中文补"等"、西文补"et al"。"""
    names = [a.strip() for a in authors.split(";") if a.strip()]
    if not names:
        return ""
    if len(names) > 3:
        names = names[:3] + ["等" if lang == "zh" else "et al"]
    return ", ".join(names)


def fmt_gbt(i, row):
    t = (row.get("type") or "").strip().upper()
    mark = MARK.get(t, "")
    authors_raw = row.get("authors") or ""
    lang = detect_lang(authors_raw, row.get("lang"))
    authors = fmt_authors(authors_raw, lang)

    title = (row.get("title") or "").strip()
    source = (row.get("source") or "").strip()   # 刊名 / 书名 / 会议名 / 保存单位
    place = (row.get("address") or "").strip()   # 出版地 / 保存地
    year = (row.get("year") or "").strip()
    volume = (row.get("volume") or "").strip()
    issue = (row.get("issue") or "").strip()
    pages = (row.get("pages") or "").strip()
    doi = (row.get("doi") or "").strip()

    head = "[%d] %s. %s%s." % (i, authors, title, mark)

    if t == "J":
        body = source
        if year:
            body += (", " if body else "") + year
        vol_issue = volume + ("(%s)" % issue if issue else "")
        if vol_issue:
            body += (", " if body else "") + vol_issue
        if pages:
            body += ": " + pages
    elif t == "M":
        body = (place + ": " if place else "") + source
        if year:
            body += (", " if body else "") + year
        if pages:
            body += ": " + pages
    elif t == "D":
        body = (place + ": " if place else "") + source
        if year:
            body += (", " if body else "") + year
    elif t == "C":
        body = source
        if place:
            body += (", " if body else "") + place
        if year:
            body += (", " if body else "") + year
        if pages:
            body += ": " + pages
    else:
        body = source
        if year:
            body += (", " if body else "") + year

    parts = [head]
    if body:
        parts.append(body + ".")
    if doi:
        parts.append("DOI: %s." % doi)
    return " ".join(parts)


def split_names(authors_raw):
    return [a.strip() for a in (authors_raw or "").split(";") if a.strip()]


def join_apa(names):
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    return ", ".join(names[:-1]) + ", & " + names[-1]


def join_ieee(names):
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    return ", ".join(names[:-1]) + " and " + names[-1]


def join_nature(names):
    if not names:
        return ""
    if len(names) > 4:
        return names[0] + " et al."
    return ", ".join(names)


def _ay_head(authors, year):
    """作者-年制头部：'Authors (Year). '；缺哪项少哪项，不留悬挂空格。"""
    parts = []
    if authors:
        parts.append(authors)
    if year:
        parts.append("(%s)" % year)
    if not parts:
        return ""
    return " ".join(parts) + ". "


def fmt_apa(i, row):
    """APA 7（作者-年制，不编号）。"""
    t = (row.get("type") or "").strip().upper()
    authors = join_apa(split_names(row.get("authors") or ""))
    title = (row.get("title") or "").strip()
    source = (row.get("source") or "").strip()
    year = (row.get("year") or "").strip()
    volume = (row.get("volume") or "").strip()
    issue = (row.get("issue") or "").strip()
    pages = (row.get("pages") or "").strip()
    doi = (row.get("doi") or "").strip()
    head = _ay_head(authors, year)

    if t == "J":
        s = head + "%s. %s" % (title, source)
        if volume:
            s += ", " + volume + ("(%s)" % issue if issue else "")
        if pages:
            s += ", " + pages
        s += "."
    elif t == "M":
        s = head + "%s. %s." % (title, source)
    elif t == "D":
        s = head + "%s [Thesis]. %s." % (title, source)
    elif t == "C":
        s = head + "%s. In %s" % (title, source)
        if pages:
            s += " (pp. %s)" % pages
        s += "."
    else:
        s = head + "%s. %s." % (title, source)
    if doi:
        s += " https://doi.org/" + doi
    return s.strip()


def fmt_ieee(i, row):
    """IEEE（数字制，[N] 编号）。"""
    t = (row.get("type") or "").strip().upper()
    authors = join_ieee(split_names(row.get("authors") or ""))
    title = (row.get("title") or "").strip()
    source = (row.get("source") or "").strip()
    place = (row.get("address") or "").strip()
    year = (row.get("year") or "").strip()
    volume = (row.get("volume") or "").strip()
    issue = (row.get("issue") or "").strip()
    pages = (row.get("pages") or "").strip()

    seg = "[%d]" % i
    if authors:
        seg += " " + authors + ","
    tail = []
    if t == "J":
        if title:
            tail.append('"%s",' % title)
        if source:
            tail.append(source)
        if volume:
            tail.append("vol. " + volume)
        if issue:
            tail.append("no. " + issue)
        if pages:
            tail.append("pp. " + pages)
        if year:
            tail.append(year)
    elif t == "M":
        if title:
            tail.append(title + ",")
        tail.append((place + ": " if place else "") + source)
        if year:
            tail.append(year)
    elif t == "D":
        if title:
            tail.append('"%s",' % title)
        tail.append("Thesis")
        if source:
            tail.append(source)
        if year:
            tail.append(year)
    elif t == "C":
        if title:
            tail.append('"%s",' % title)
        if source:
            tail.append("in " + source)
        if year:
            tail.append(year)
        if pages:
            tail.append("pp. " + pages)
    else:
        tail.append((place + ": " if place else "") + source)
        if year:
            tail.append(year)
    items = [seg.rstrip(",")] + [x.rstrip(",") for x in tail if x]
    return ", ".join(items) + "."


def fmt_nature(i, row):
    """Nature（数字制，N. 编号；作者 >4 取第一作者 et al.）。"""
    t = (row.get("type") or "").strip().upper()
    authors = join_nature(split_names(row.get("authors") or ""))
    title = (row.get("title") or "").strip()
    source = (row.get("source") or "").strip()
    place = (row.get("address") or "").strip()
    year = (row.get("year") or "").strip()
    volume = (row.get("volume") or "").strip()
    pages = (row.get("pages") or "").strip()
    doi = (row.get("doi") or "").strip()

    seg = "%d." % i
    if authors:
        seg += " " + authors
    if t == "J":
        s = seg
        if title:
            s += " " + title + "."
        if source:
            s += " " + source
        if volume:
            s += " " + volume
        if pages:
            s += ", " + pages
        if year:
            s += " (%s)" % year
        s += "."
        if doi:
            s += " https://doi.org/" + doi
    elif t == "M":
        s = seg
        if title:
            s += " " + title + "."
        pub = (place + ": " if place else "") + source
        if pub:
            s += " (" + pub + (", " + year if year else "") + ")."
    elif t == "D":
        s = seg
        if title:
            s += " " + title + "."
        if source:
            s += " Thesis, " + source
        if year:
            s += " (%s)." % year
    elif t == "C":
        s = seg
        if title:
            s += " " + title + "."
        if source:
            s += " In " + source
        if pages:
            s += ", " + pages
        if year:
            s += " (%s)." % year
    else:
        s = seg
        if title:
            s += " " + title + "."
        if source:
            s += " " + source
        if year:
            s += " (%s)." % year
    return s


def fmt_entry(i, row, style="gbt"):
    """按目标格式排版第 i 条参考文献（1 起）。缺失字段留空，不编造。"""
    if style == "apa":
        return fmt_apa(i, row)
    if style == "ieee":
        return fmt_ieee(i, row)
    if style == "nature":
        return fmt_nature(i, row)
    return fmt_gbt(i, row)


def parse_bib(path):
    """解析 .bib 文件为行字典列表（含大括号嵌套的健壮处理）。"""
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    rows = []
    for m in re.finditer(r"@(\w+)\s*\{\s*([^,]+),", text):
        btype = m.group(1).lower()
        start = m.end()
        depth = 1
        i = start
        while i < len(text) and depth:
            if text[i] == "{":
                depth += 1
            elif text[i] == "}":
                depth -= 1
            i += 1
        body = text[start:i - 1]
        fields = {}
        for fm in re.finditer(r"(\w+)\s*=\s*", body):
            key = fm.group(1).lower()
            vs = fm.end()
            if vs >= len(body):
                continue
            ch = body[vs]
            if ch == "{":
                depth, i = 1, vs + 1
                while i < len(body) and depth:
                    if body[i] == "{":
                        depth += 1
                    elif body[i] == "}":
                        depth -= 1
                    i += 1
                value = body[vs + 1:i - 1]
            elif ch == '"':
                end = body.find('"', vs + 1)
                value = body[vs + 1:end] if end != -1 else body[vs + 1:]
            else:
                m = re.match(r"[^,\s}]+", body[vs:])
                value = m.group(0) if m else ""
            fields[key] = value.strip()
        rows.append({
            "type": BIB_TYPE.get(btype, "J"),
            "authors": fields.get("author", "").replace(" and ", ";"),
            "title": fields.get("title", ""),
            "source": fields.get("journal") or fields.get("booktitle")
                      or fields.get("publisher") or fields.get("school") or "",
            "year": fields.get("year", ""),
            "volume": fields.get("volume", ""),
            "issue": fields.get("number", ""),
            "pages": fields.get("pages", ""),
            "doi": fields.get("doi", ""),
            "address": fields.get("address", ""),
            "lang": "auto",
        })
    return rows


def main():
    ap = argparse.ArgumentParser(description="参考文献格式化（GB/T 7714 / APA 7 / IEEE / Nature；支持 CSV 与 BibTeX）")
    ap.add_argument("input", help="输入文件（.csv 或 .bib）")
    ap.add_argument("--format", choices=["gbt", "apa", "ieee", "nature"], default="gbt",
                    help="输出格式，默认 gbt（GB/T 7714）")
    ap.add_argument("--dedupe", action="store_true", help="按题名去重（保留信息更全的一条）")
    ap.add_argument("--sort", choices=["year"], default=None, help="排序方式，默认按输入顺序")
    args = ap.parse_args()

    try:
        src = resolve_input(args.input)
    except PathError as e:
        print("错误: 路径校验失败：%s" % e, file=sys.stderr)
        sys.exit(2)

    if src.lower().endswith(".bib"):
        rows = parse_bib(src)
    else:
        with open(src, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))

    if args.dedupe:
        seen = {}
        for row in rows:
            key = (row.get("title") or "").strip().lower()
            if not key:
                continue
            if key not in seen or len(str(row)) > len(str(seen[key])):
                seen[key] = row
        rows = list(seen.values())

    if args.sort == "year":
        rows.sort(key=lambda r: (r.get("year") or "9999"))

    for i, row in enumerate(rows, 1):
        print(fmt_entry(i, row, args.format))


if __name__ == "__main__":
    main()