#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_format.py —— .docx 排版自动校验脚本（纯 Python 标准库，零第三方依赖）

用法:
    python check_format.py 论文.docx                        # 默认条件模式：短文/片段
    python check_format.py 论文.docx --strict               # 完整论文模式：公式/上标/目录/一级标题必须存在
    python check_format.py 论文.docx --lang en              # 英文口径：直引号合法、数字编号、Contents/References
    python check_format.py 论文.docx --template 模板.docx   # 模板基准模式：版式类基准改用模板实际值

检查项（与 references/manuscript-format.md 排版基线一一对应）:
  1. 公式         OMML 公式对象（<m:oMath>）数量与结构（角标 sSub/sSup/sSubSup 完整、无字面 _ ^）；
                  无公式时默认 PASS，--strict 下 FAIL
  2. 伪下标        正文中 x_1 / ω_i 式下划线伪下标（含希腊字母，应为 0）
  3. 英文直引号    中文正文中出现 ASCII " '（应为 0）
  4. 一级标题     "一、二、三……" 中文编号；无则默认 PASS，--strict 下 FAIL
  5. 标题对齐      题目/一级标题（Heading1、一、二、三）须居中；二级/三级（Heading2/3、1.1）须左对齐
  6. 全黑字色      除 auto 外是否出现非纯黑 (000000) 字色
  7. 上标引用      是否存在上标（w:vertAlign=superscript）；无则默认 PASS，--strict 下 FAIL
  8. 页边距        是否设置 w:pgMar
  9. 表格          w:tbl 数量与线型（三线表 / 全线表格）
  10. 目录          是否含"目录"标题或 TOC 域；无则默认 PASS，--strict 下 FAIL

退出码: 0 = 全部通过；1 = 存在不合格项；2 = 无法解析

模板基准模式（--template 模板.docx）:
    版式类基准从模板 docx 实际提取——标题编号模式（第X章 / 一、）、标题对齐
    （以模板实际占位段落为准，样式定义兑底）、模板字色加入允许色、页边距与模板一致性；
    学术底线项不受模板影响，仍按规范校验（公式 OMML、伪下标、引号、上标、交叉引用、参考文献结构）。

语言口径（--lang auto|zh|en，默认 auto 按正文 CJK 占比检测）:
    en 英文口径：英文直引号合法、一级标题按数字编号（1 / 1.1 / I.）、目录认 Contents、
    参考文献认 References 且不校验 [J]/[M] 载体标识、交叉引用认 Figure/Table/Eq.、
    正文引用不强制上标（APA/IEEE 行内引用合法）；公式 OMML、伪下标、字色、页边距、
    表格线型等学术底线项两种口径均校验。zh 口径与历史行为完全一致。
"""

import argparse
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

try:
    from _paths import PathError, resolve_input
except ImportError:  # 作为包导入时补 scripts/ 到 sys.path
    import os as _os
    sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
    from _paths import PathError, resolve_input

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# GB/T 7714 载体标识：J/M/D/C/N/R/S/P/G/Z/A 及联机 /OL 变体
CARRIER_RE = re.compile(
    r"\[(?:[JMDCNRSPGZA]|(?:J|M|C|N|R|S|P|G|Z|EB|DB|CP|A)/OL)\]"
)


def qw(tag):
    return "{%s}%s" % (W, tag)


def qm(tag):
    return "{%s}%s" % (M, tag)


def text_of(node):
    return "".join(t.text or "" for t in node.iter(qw("t")))


def parse_document(docx_path):
    with zipfile.ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    paras = []
    for p in root.iter(qw("p")):
        t = text_of(p).strip()
        if t:
            paras.append(t)
    return root, "".join(paras), paras


def pstyle_val(p):
    ppr = p.find(qw("pPr"))
    if ppr is None:
        return None
    ps = ppr.find(qw("pStyle"))
    return ps.get(qw("val")) if ps is not None else None


def pjc(p):
    ppr = p.find(qw("pPr"))
    if ppr is None:
        return None
    jc = ppr.find(qw("jc"))
    return jc.get(qw("val")) if jc is not None else None


def parse_template(path):
    """解析模板 docx，提取版式基准：页边距、标题对齐（实际占位优先）、允许字色、编号模式。"""
    with zipfile.ZipFile(path) as zf:
        doc = ET.fromstring(zf.read("word/document.xml"))
        styles = None
        try:
            styles = ET.fromstring(zf.read("word/styles.xml"))
        except KeyError:
            pass
    tpl = {"margins": None, "h1_jc": None, "h2_jc": None,
           "allowed_colors": set(), "heading_pattern": "cn_ordinal"}

    pg = doc.find(".//" + qw("pgMar"))
    if pg is not None:
        tpl["margins"] = {k: pg.get(qw(k)) for k in ("top", "right", "bottom", "left")}

    for src in (doc, styles):
        if src is None:
            continue
        for c in src.iter(qw("color")):
            v = (c.get(qw("val")) or "auto").lower()
            if v not in ("auto", "000000", "black"):
                tpl["allowed_colors"].add(v)

    # 样式表里 heading 1/2 的对齐定义（styleId → (级别, jc)）
    hmap = {}
    if styles is not None:
        for st in styles.iter(qw("style")):
            if (st.get(qw("type")) or "") != "paragraph":
                continue
            nm = st.find(qw("name"))
            name = ((nm.get(qw("val")) or "").strip().lower() if nm is not None else "")
            if name not in ("heading 1", "heading 2"):
                continue
            ppr = st.find(qw("pPr"))
            jc = None
            if ppr is not None:
                j = ppr.find(qw("jc"))
                jc = j.get(qw("val")) if j is not None else None
            hmap[st.get(qw("styleId"))] = (1 if name == "heading 1" else 2, jc)

    # 实际占位段落优先：以 document.xml 中首个使用该样式、有文字的段落为准
    for p in doc.iter(qw("p")):
        sid = pstyle_val(p)
        if sid in hmap and text_of(p).strip():
            key = "h1_jc" if hmap[sid][0] == 1 else "h2_jc"
            if tpl[key] is None:
                tpl[key] = pjc(p) or hmap[sid][1]
    for sid, (lvl, jc) in hmap.items():
        key = "h1_jc" if lvl == 1 else "h2_jc"
        if tpl[key] is None and jc:
            tpl[key] = jc

    # 标题编号模式：统计模板占位文本（第X章 / 一、/ 数字制）
    texts = [text_of(p).strip() for p in doc.iter(qw("p"))]
    cn = sum(1 for t in texts if re.match(r"^[一二三四五六七八九十]+、", t))
    ch = sum(1 for t in texts if re.match(r"^第[一二三四五六七八九十百\d]+章", t))
    rx_num = _h1_num_re()
    nm = sum(1 for t in texts if rx_num.match(t))
    if ch and ch >= cn:
        tpl["heading_pattern"] = "chapter"
    elif nm and nm > cn:
        tpl["heading_pattern"] = "num"
    return tpl


def _math_has_content(el):
    """判断 m:e / m:sub / m:sup 等参数位是否有实质内容。"""
    for t in el.iter(qm("t")):
        if (t.text or "").strip():
            return True
    for tag in ("f", "rad", "d", "nary", "sSub", "sSup", "sSubSup", "acc", "bar", "func"):
        if el.find(".//" + qm(tag)) is not None:
            return True
    return False


def _omml_issues(root):
    """检测 OMML 结构缺陷：角标缺基/缺子/缺上、公式内字面 _ ^。"""
    issues = []
    for om in root.iter(qm("oMath")):
        for t in om.iter(qm("t")):
            s = t.text or ""
            if "_" in s or "^" in s:
                issues.append("公式内出现字面下划线/上标符: %s" % s[:20])
        for tag, req in (("sSub", ("e", "sub")), ("sSup", ("e", "sup")),
                         ("sSubSup", ("e", "sub", "sup"))):
            for node in om.iter(qm(tag)):
                for child in req:
                    el = node.find(qm(child))
                    if el is None or not _math_has_content(el):
                        issues.append("%s 结构不完整（缺 %s）" % (tag, child))
    return issues


def check_omml(root, strict=False):
    n = len(list(root.iter(qm("oMath"))))
    if n == 0:
        if strict:
            return ("文档无公式（strict 模式要求存在）", False, "OMML 公式数=0")
        return ("未检测到公式（若文档含公式则必须为 OMML）", True, "OMML 公式数=0")
    issues = _omml_issues(root)
    if issues:
        return ("公式结构有误（角标/命名空间）", False, "OMML 公式数=%d；%s" % (n, issues[:5]))
    return ("公式为 OMML 原生对象且结构完整", True, "OMML 公式数=%d" % n)


def check_pseudo_subscript(full_text):
    hits = re.findall(r"[\u0370-\u03FFa-zA-Z\u4E00-\u9FFF]_[A-Za-z0-9\u0370-\u03FF]", full_text)
    ok = len(hits) == 0
    return ("无伪下标（含希腊字母）" if ok else "存在伪下标，须改为真实下标格式", ok, "伪下标 %d 处，示例=%s" % (len(hits), hits[:8]))


def check_ascii_quotes(full_text, lang="zh"):
    if lang == "en":
        return ("英文口径：英文直引号合法", True, "跳过")
    hits = re.findall(r'[\u4e00-\u9fff]["\']|["\'][\u4e00-\u9fff]', full_text)
    ok = len(hits) == 0
    return ("正文无英文直引号" if ok else "中文句中出现英文直引号，须用中文引号“”", ok, "直引号 %d 处" % len(hits))


def check_cn_headings(full_text, strict=False, tpl=None, lang="zh", paras=None):
    pattern = (tpl or {}).get("heading_pattern")
    if not pattern:
        pattern = "num" if lang == "en" else "cn_ordinal"
    if pattern == "chapter":
        found = len(re.findall(r"第[一二三四五六七八九十百\d]+章", full_text))
        label = "一级标题为模板编号模式（第X章）"
    elif pattern == "num":
        rx = _h1_num_re()
        found = sum(1 for t in (paras or []) if rx.match(t.strip()))
        label = "一级标题为数字编号（1 / 1.1 / I.）"
    else:
        found = len(re.findall(r"[一二三四五六七八九十]+、", full_text))
        label = "一级标题为中文编号“一、二、三……”"
    if found:
        return (label, True, "命中 %d 个" % found)
    if strict:
        return ("未检测到一级标题编号（strict 模式要求存在；基准模式=%s）" % pattern, False, "命中 0 个")
    return ("未检测到一级标题编号", True, "命中 0 个")


def _jc_name(jc):
    return {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}.get(jc, jc or "未设置")


def detect_lang(full_text):
    """按正文 CJK 占比判定语言口径：zh / en；空文本归 zh（保持中文口径零变化）。"""
    cjk = sum(1 for ch in full_text if "\u4e00" <= ch <= "\u9fff")
    letters = sum(1 for ch in full_text if ch.isalpha())
    if letters == 0 and cjk == 0:
        return "zh"
    return "en" if cjk <= letters * 0.15 else "zh"


def _h1_num_re():
    """数字制一级标题：'1 Introduction' / '1. Introduction' / 'I. INTRODUCTION'（不匹配 1.1 多级）。"""
    return re.compile(r"^(?:\d+\.?|[IVXLC]+\.)\s+\S")


_TITLE_EXCLUDE = {"摘要", "abstract", "关键词", "keywords", "目录", "contents"}


def _looks_like_bare_title(text, tlen):
    """无 Title 样式时，把首个短、无编号、无句读的段落视为论文题目。"""
    if not text or len(text) > tlen:
        return False
    if text.strip().lower() in _TITLE_EXCLUDE:
        return False
    if re.match(r"^[一二三四五六七八九十]+、", text):
        return False
    if re.match(r"^\d+(?:\.\d+)*[\.\s]", text):
        return False
    if re.search(r"[。，、；：！？,;:]", text):
        return False
    return True


def check_headings_align(root, tpl=None, lang="zh"):
    """题目与一级标题对齐基准默认居中、二三级默认左对齐；--template 按模板实际值，--lang en 一级按数字编号识别。"""
    pattern = (tpl or {}).get("heading_pattern")
    if not pattern:
        pattern = "num" if lang == "en" else "cn_ordinal"
    h1_expect = (tpl or {}).get("h1_jc") or "center"
    h2_expect = (tpl or {}).get("h2_jc")  # None → 二三级不得居中
    rx_num = _h1_num_re()
    rx_title_style = re.compile(r"^title\d*$", re.I)
    tlen = 60 if lang == "en" else 40
    bad_title, bad_center, bad_left = [], [], []
    first_seen = False
    for p in root.iter(qw("p")):
        t = text_of(p).strip()
        if not t:
            continue
        style = pstyle_val(p) or ""
        align = pjc(p)

        is_title = bool(rx_title_style.match(style))
        if not first_seen:
            first_seen = True
            if not is_title and _looks_like_bare_title(t, tlen):
                is_title = True
        if is_title:
            if align != "center":
                bad_title.append(t[:20])
            continue

        if len(t) > tlen:
            continue
        is_h1 = ("Heading1" in style) or bool(re.match(r"^[一二三四五六七八九十]+、", t))
        if pattern == "chapter" and re.match(r"^第[一二三四五六七八九十百\d]+章", t):
            is_h1 = True
        if pattern == "num" and rx_num.match(t):
            is_h1 = True
        # 二级/三级：Heading2/3，或“1.1 …”式编号；排除带句读的正文（如“3.14 是圆周率。”）
        is_sub = ("Heading2" in style or "Heading3" in style) or (
            bool(re.match(r"^\d+(?:\.\d+)+\s+\S", t)) and not re.search(r"[。，、；：！？,;]", t)
        )
        if is_h1 and align != h1_expect:
            bad_center.append(t[:20])
        if is_sub and ((h2_expect is None and align == "center") or
                       (h2_expect is not None and align != h2_expect)):
            bad_left.append(t[:20])
    if bad_title or bad_center or bad_left:
        msg = "标题对齐错误（基准：题目居中 / 一级 %s / 二三级 %s）" % (
            _jc_name(h1_expect), _jc_name(h2_expect) if h2_expect else "左对齐")
        if bad_title:
            msg += "；题目应居中未居中: %s" % bad_title[:5]
        if bad_center:
            msg += "；应 %s 未 %s: %s" % (_jc_name(h1_expect), _jc_name(h1_expect), bad_center[:5])
        if bad_left:
            msg += "；二三级对齐与基准不符: %s" % bad_left[:5]
        return (msg, False, "")
    return ("标题对齐正确（题目居中，一级 %s，二三级 %s）" % (
        _jc_name(h1_expect), _jc_name(h2_expect) if h2_expect else "左对齐"), True, "")


def check_black_color(root, tpl=None):
    allowed = {"auto", "000000", "black"}
    extra = (tpl or {}).get("allowed_colors") or set()
    bad = []
    for c in root.iter(qw("color")):
        val = (c.get(qw("val")) or "auto").lower()
        if val not in allowed and val not in extra:
            bad.append(val)
    ok = len(bad) == 0
    label = "字色合规（纯黑 + 模板色 %d 种）" % len(extra) if extra else "全文纯黑字色"
    return (label if ok else "存在非黑且非模板色 %s" % sorted(set(bad)), ok, "问题色 %d 处" % len(bad))


def check_superscript(root, strict=False, lang="zh"):
    vals = [v.get(qw("val")) for v in root.iter(qw("vertAlign"))]
    sup = [v for v in vals if v == "superscript"]
    if sup:
        return ("存在上标（引用/标注）", True, "上标 %d 处" % len(sup))
    if strict and lang != "en":
        return ("文档无上标（strict 模式要求存在）", False, "上标 0 处")
    if lang == "en":
        return ("英文口径：行内引用 (Author, 2020) / [1] 合法，不强制上标", True, "上标 0 处")
    return ("未检测到上标（若正文有引用须上标）", True, "上标 0 处")


def check_page_setup(root, tpl=None):
    pg = root.find(".//" + qw("pgMar"))
    if pg is None:
        return ("未检测到页边距设置", False, "")
    if tpl and tpl.get("margins"):
        diffs = []
        for k, v in tpl["margins"].items():
            if v and pg.get(qw(k)) != v:
                diffs.append("%s: 文档=%s 模板=%s" % (k, pg.get(qw(k)), v))
        if diffs:
            return ("页边距与模板不一致", False, "; ".join(diffs))
        return ("页边距与模板一致", True, "")
    return ("页边距已设置", True, "")


def check_tables(root):
    tbls = list(root.iter(qw("tbl")))
    if not tbls:
        return ("无表格", True, "0 个表格")
    styles = []
    bad = []
    for i, t in enumerate(tbls, 1):
        tpr = t.find(qw("tblPr"))
        borders = tpr.find(qw("tblBorders")) if tpr is not None else None
        if borders is None:
            styles.append("默认边框")
            continue
        has_l = borders.find(qw("left")) is not None
        has_r = borders.find(qw("right")) is not None
        has_v = borders.find(qw("insideV")) is not None
        if not has_l and not has_r and not has_v:
            # 三线表：只有顶线/栏目线/底线三条线，行间不得有横线
            ih = borders.find(qw("insideH"))
            ih_bad = ih is not None and (ih.get(qw("val")) or "single") != "none"
            top = borders.find(qw("top"))
            bot = borders.find(qw("bottom"))
            tb_missing = (top is None or (top.get(qw("val")) or "single") == "none"
                          or bot is None or (bot.get(qw("val")) or "single") == "none")
            reasons = []
            if ih_bad:
                reasons.append("行间有横线(insideH)")
            if tb_missing:
                reasons.append("缺顶线/底线")
            if reasons:
                bad.append("表%d %s" % (i, "、".join(reasons)))
                styles.append("三线表样式(异常)")
            else:
                styles.append("三线表样式")
        else:
            styles.append("全线表格样式")
    if bad:
        return ("表格线型: " + " | ".join(styles), False,
                "%d 个表格; %s" % (len(tbls), "; ".join(bad[:3])))
    return ("表格线型: " + " | ".join(styles), True, "%d 个表格" % len(tbls))


def check_toc(full_text, root, strict=False, lang="zh"):
    if lang == "en":
        has_text = ("contents" in full_text.lower()) or ("目录" in full_text)
    else:
        has_text = "目录" in full_text
    has_field = any("TOC" in (t.text or "") for t in root.iter(qw("instrText")))
    if not has_field:
        has_field = any("TOC" in (f.get(qw("instr")) or "") for f in root.iter(qw("fldSimple")))
    detail = "文本=%s 域=%s" % (has_text, has_field)
    if has_field:
        return ("含目录（TOC 域，可更新页码）", True, detail)
    if has_text:
        return ("目录为静态文本，须改为 TOC 域（自动生成、可更新页码）", False, detail)
    if strict:
        return ("文档无目录（strict 模式要求存在）", False, detail)
    return ("未检测到目录", True, detail)


def check_crossref(full_text, paras, lang="zh"):
    if lang == "en":
        fig_t, tab_t = r"^\s*(?:Figure|Fig\.)\s*(\d+)", r"^\s*Table\s+(\d+)"
        fig_i, tab_i, eq_i = r"(?:Figure|Fig\.)\s*(\d+)", r"Table\s+(\d+)", r"(?:Eqs?\.|Equations?)\s*\((\d+)\)"
    else:
        fig_t, tab_t = r"^\s*图\s*(\d+)", r"^\s*表\s*(\d+)"
        fig_i, tab_i, eq_i = r"图\s*(\d+)", r"表\s*(\d+)", r"式\s*\((\d+)\)"
    fig_titles = set()
    tab_titles = set()
    for t in paras:
        m = re.match(fig_t, t)
        if m:
            fig_titles.add(int(m.group(1)))
        m = re.match(tab_t, t)
        if m:
            tab_titles.add(int(m.group(1)))
    eq_all = {int(m.group(1)) for m in re.finditer(r"\((\d+)\)", full_text)}
    issues = []
    for t in paras:
        if re.match(fig_t, t) or re.match(tab_t, t):
            continue
        for m in re.finditer(fig_i, t):
            n = int(m.group(1))
            if n not in fig_titles:
                issues.append("引用不存在的图 %d" % n)
        for m in re.finditer(tab_i, t):
            n = int(m.group(1))
            if n not in tab_titles:
                issues.append("引用不存在的表 %d" % n)
        for m in re.finditer(eq_i, t):
            n = int(m.group(1))
            if n not in eq_all:
                issues.append("引用不存在的式(%d)" % n)
    ok = len(issues) == 0
    return ("交叉引用编号一致" if ok else "交叉引用不一致", ok, "图题 %d / 表题 %d，问题 %s" % (len(fig_titles), len(tab_titles), issues[:5]))


def check_references(paras, lang="zh"):
    """参考文献检查：zh 口径条目以 [N] 开头、含 [J]/[M]/[D]/[C] 载体标识；en 口径认 References 章节、不校验 GB/T 组成。"""
    section = "References" if lang == "en" else "参考文献"
    idx = None
    for i, t in enumerate(paras):
        if t.strip() == section:
            idx = i
            break
    if idx is None:
        return ("未检测到参考文献章节（若含参考文献则须有）", True, "跳过")
    entries = [t for t in paras[idx + 1:] if t.strip()]
    if not entries:
        return ("参考文献章节无条目", False, "")
    if lang == "en":
        return ("参考文献条目存在（英文口径不校验 [J]/[M] 载体标识，组成由 format_refs.py 保证）", True, "共 %d 条" % len(entries))
    bad = []
    for t in entries:
        if not re.match(r"^\[\d+\]\s+", t):
            bad.append("条目未以 [N] 开头: %s" % t[:20])
        elif not CARRIER_RE.search(t):
            bad.append("条目缺载体标识（[J/M/D/C/N/R/S/P/G/Z/A] 或 /OL）: %s" % t[:20])
    ok = len(bad) == 0
    return ("参考文献条目结构正确" if ok else "参考文献条目结构错误", ok,
            "共 %d 条；问题 %d 处: %s" % (len(entries), len(bad), bad[:5]))


def check_title_len(paras, lang="zh"):
    """标题长度提示：中文约 30 字、英文约 90 字符（一行放下）。"""
    limit = 90 if lang == "en" else 30
    long_titles = []
    rx_num = _h1_num_re()
    for t in paras:
        s = t.strip()
        is_cn = bool(re.match(r"^[一二三四五六七八九十]+、", s))
        is_num = bool(re.match(r"^\d+(\.\d+)*[\.\s]", s)) or bool(rx_num.match(s))
        if (is_cn or is_num) and len(s) > limit:
            long_titles.append(s[:60])
    if not long_titles:
        return ("标题长度正常（≤%d %s）" % (limit, "字符" if lang == "en" else "字"), True, "")
    return ("标题过长（建议 ≤%d、一行放下）: %s" % (limit, long_titles[:5]), True, "提示")


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=".docx 排版自动校验（纯标准库）")
    ap.add_argument("docx", help="待校验的 Word 文档路径")
    ap.add_argument("--strict", action="store_true",
                    help="完整论文模式：公式/上标/目录/一级标题必须存在")
    ap.add_argument("--template", default=None,
                    help="模板 docx：版式类基准（编号/对齐/字色/页边距）改用模板实际值，学术底线项仍按规范")
    ap.add_argument("--lang", choices=["auto", "zh", "en"], default="auto",
                    help="语言口径：auto 自动检测（默认）；en 英文口径（直引号合法、数字编号、Contents/References、不强制上标）")
    args = ap.parse_args()

    try:
        docx_path = resolve_input(args.docx)
        root, full_text, paras = parse_document(docx_path)
    except PathError as e:
        print("[ERROR] 路径校验失败: %s" % e)
        sys.exit(2)
    except Exception as e:
        print("[ERROR] 无法解析文档: %s" % e)
        sys.exit(2)

    lang = args.lang
    if lang == "auto":
        lang = detect_lang(full_text)
    print("[语言] %s口径%s" % ("英文" if lang == "en" else "中文", "（--lang 指定）" if args.lang != "auto" else "（自动检测）"))

    tpl = None
    if args.template:
        try:
            tpl = parse_template(resolve_input(args.template))
        except PathError as e:
            print("[ERROR] 路径校验失败: %s" % e)
            sys.exit(2)
        except Exception as e:
            print("[ERROR] 无法解析模板: %s" % e)
            sys.exit(2)
        print("[模板基准] 编号模式=%s 一级对齐=%s 二级对齐=%s 模板色=%s 页边距=%s"
              % (tpl["heading_pattern"], tpl["h1_jc"] or "默认(居中)", tpl["h2_jc"] or "默认(左)",
                 ", ".join(sorted(tpl["allowed_colors"])) or "无", tpl["margins"] or "未设置"))

    checks = [
        ("公式(OMML)", check_omml(root, args.strict)),
        ("伪下标", check_pseudo_subscript(full_text)),
        ("中文引号", check_ascii_quotes(full_text, lang)),
        ("一级标题编号", check_cn_headings(full_text, args.strict, tpl, lang, paras)),
        ("标题对齐", check_headings_align(root, tpl, lang)),
        ("字色", check_black_color(root, tpl)),
        ("引用上标", check_superscript(root, args.strict, lang)),
        ("页边距", check_page_setup(root, tpl)),
        ("表格线型", check_tables(root)),
        ("目录", check_toc(full_text, root, args.strict, lang)),
        ("交叉引用", check_crossref(full_text, paras, lang)),
        ("参考文献", check_references(paras, lang)),
        ("标题长度", check_title_len(paras, lang)),
    ]

    failed = 0
    for name, (msg, ok, detail) in checks:
        tag = "PASS" if ok else "FAIL"
        if not ok:
            failed += 1
        print("[%s] %-16s %s" % (tag, name, (msg if ok else msg + "  <-- " + detail)))

    print("-" * 60)
    print("结果: %d 项不合格 / 共 %d 项" % (failed, len(checks)))
    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()