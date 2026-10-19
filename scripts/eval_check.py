#!/usr/bin/env python3
"""技能包自动化回归检查（evals/eval-cases.md 中可脚本验证的部分）。

用法:
    python scripts/eval_check.py                    # 结构完整性 + 脚本语法 + 引用格式冒烟
    python scripts/eval_check.py --docx 论文.docx   # 追加 check_format.py 排版校验（E16）
    python scripts/eval_check.py --text 正文.md     # 追加 check_style.py 段长自检（E15）

覆盖:
    - 包结构完整性：frontmatter 字段齐全、SKILL.md 引用 README（软禁止豁免）、
      references/ 与 assets/examples/ 每个文件都被 SKILL.md 链接
    - scripts/*.py 全部语法编译通过
    - format_refs.py 四种格式冒烟（gbt/apa/ieee/nature）：退出码 0、
      每条独占一行、编号形式正确 → 对应 E4
    - check_format.py 语言口径冒烟（英文夹具 --lang en / 自动检测；中文夹具回归）→ E23
    - 可选 --docx / --text 挂接 check_format.py / check_style.py → E16 / E15
    行为类用例（E1-E3、E5-E14、E17-E22、E24）仍需按 eval-cases.md 人工跑。

退出码: 全部 PASS 为 0，出现 FAIL 为 1。纯标准库。
"""
import argparse
import csv
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # scripts/ 上一级
ENV = dict(os.environ, PYTHONIOENCODING="utf-8")


def read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def check_frontmatter(results):
    """SKILL.md frontmatter 必含 name/name_cn/description/description_cn 且非空。"""
    lines = read(os.path.join(ROOT, "SKILL.md")).split("\n")
    fm = {}
    if lines and lines[0].strip() == "---":
        for ln in lines[1:]:
            if ln.strip() == "---":
                break
            if ":" in ln:
                k, v = ln.split(":", 1)
                fm[k.strip()] = v.strip()
    for key in ("name", "name_cn", "description", "description_cn"):
        results.append(("frontmatter.%s 非空" % key, bool(fm.get(key)),
                        fm.get(key)[:30] + "..." if len(fm.get(key) or "") > 30 else (fm.get(key) or "缺失"), "-"))


def check_links(results):
    """SKILL.md 引用 README.md；references/ 与 assets/examples/ 每个文件都被 SKILL.md 链接。"""
    skill_text = read(os.path.join(ROOT, "SKILL.md"))
    results.append(("SKILL.md 引用 README.md（软禁止豁免）",
                    "README.md" in skill_text, "-", "-"))
    for sub in ("references", os.path.join("assets", "examples")):
        d = os.path.join(ROOT, sub)
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            if fn.endswith(".md"):
                results.append(("%s/%s 被 SKILL.md 引用" % (sub, fn),
                                fn in skill_text, "-", "-"))


def check_scripts_compile(results):
    """scripts/*.py 全部语法编译通过。"""
    sdir = os.path.join(ROOT, "scripts")
    for fn in sorted(os.listdir(sdir)):
        if fn.endswith(".py"):
            path = os.path.join(sdir, fn)
            try:
                compile(read(path), path, "exec")
                ok, detail = True, "-"
            except SyntaxError as e:
                ok, detail = False, "line %s: %s" % (e.lineno, e.msg)
            results.append(("scripts/%s 语法编译" % fn, ok, detail, "-"))


def run(cmd):
    p = subprocess.run(cmd, capture_output=True, env=ENV)
    return p.returncode, p.stdout.decode("utf-8", "replace")


def check_format_refs(results):
    """对 refs.example.csv 跑四种格式：退出码 0、每条一行、编号形式正确（E4）。"""
    fr = os.path.join(ROOT, "scripts", "format_refs.py")
    sample = os.path.join(ROOT, "scripts", "refs.example.csv")
    with open(sample, newline="", encoding="utf-8-sig") as f:
        n = sum(1 for r in csv.DictReader(f) if any((v or "").strip() for v in r.values()))
    for fmt in ("gbt", "apa", "ieee", "nature"):
        code, out = run([sys.executable, fr, sample, "--format", fmt])
        rows = [ln for ln in out.splitlines() if ln.strip()]
        ok = code == 0 and len(rows) == n
        detail = "exit=%d rows=%d expect=%d" % (code, len(rows), n)
        if ok and fmt in ("gbt", "ieee") and not all(r.startswith("[") for r in rows):
            ok, detail = False, detail + "；编号未以 [ 开头"
        if ok and fmt == "nature" and not all(r[:1].isdigit() for r in rows):
            ok, detail = False, detail + "；编号未以数字开头"
        if ok and fmt == "apa" and any(r.startswith("[") for r in rows):
            ok, detail = False, detail + "；作者-年制不应带 [N] 编号"
        results.append(("format_refs --format %s 冒烟" % fmt, ok, detail, "E4"))


def check_docx(results, path):
    """追加：check_format.py 排版校验（E16），退出码即结果。"""
    code, out = run([sys.executable, os.path.join(ROOT, "scripts", "check_format.py"), path])
    tail = (out.strip().splitlines() or ["-"])[-1]
    results.append(("check_format.py %s" % os.path.basename(path), code == 0, tail, "E16"))


_fx_w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _fx_p(text, jc=None):
    jcx = '<w:jc w:val="%s"/>' % jc if jc else ""
    ppr = "<w:pPr>%s</w:pPr>" % jcx if jcx else ""
    return '<w:p>%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr, text)


def _fx_docx(path, paras):
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document %s><w:body>%s'
           '<w:sectPr><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/></w:sectPr>'
           '</w:body></w:document>') % (_fx_w, "".join(paras))
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", doc)


def check_format_lang(results):
    """语言口径冒烟（E23）：英文夹具走英文口径全 PASS、自动检测正确；中文夹具回归零变化。"""
    cf = os.path.join(ROOT, "scripts", "check_format.py")
    tmp = tempfile.mkdtemp(prefix="aw_eval_")
    try:
        en_doc = os.path.join(tmp, "en.docx")
        zh_doc = os.path.join(tmp, "zh.docx")
        _fx_docx(en_doc, [
            _fx_p("Deep Learning for Power Scheduling", "center"),
            _fx_p("1 Introduction", "center"),
            _fx_p("1.1 Background", "left"),
            _fx_p('The "smart grid" integrates renewables into modern power systems.'),
            _fx_p("Contents"),
            _fx_p("As shown in Figure 1, the framework has three modules."),
            _fx_p("Figure 1 Overview of the proposed framework."),
            _fx_p("References"),
            _fx_p("[1] Zhang S, Li M. A scheduling method. IEEE Trans. Power Syst., 2023, 38(2): 1-9."),
        ])
        _fx_docx(zh_doc, [
            _fx_p("面向智能电网的调度研究", "center"),
            _fx_p("一、引言", "center"),
            _fx_p("1.1 研究背景", "left"),
            _fx_p("本文研究智能电网的调度问题。"),
            _fx_p("目录"),
            _fx_p("参考文献"),
            _fx_p("[1] 张三, 李四. 一种调度方法[J]. 电力学报, 2023, 12(3): 45-58."),
        ])

        code, out = run([sys.executable, cf, en_doc, "--lang", "en"])
        ok = code == 0 and "英文口径：英文直引号合法" in out and "数字编号" in out
        results.append(("check_format --lang en 英文口径冒烟", ok, "exit=%d" % code, "E23"))

        code, out = run([sys.executable, cf, en_doc])
        ok = code == 0 and "[语言] 英文口径（自动检测）" in out
        results.append(("check_format 自动检测英文", ok, "exit=%d" % code, "E23"))

        code, out = run([sys.executable, cf, zh_doc])
        ok = code == 0 and "[语言] 中文口径（自动检测）" in out and "中文编号" in out
        results.append(("check_format 自动检测中文回归", ok, "exit=%d" % code, "E23"))

        code, out = run([sys.executable, cf, zh_doc, "--lang", "zh"])
        results.append(("check_format --lang zh 中文口径回归", code == 0, "exit=%d" % code, "E23"))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def check_text(results, path):
    """追加：check_style.py 段长自检（E15），退出码即结果。"""
    code, out = run([sys.executable, os.path.join(ROOT, "scripts", "check_style.py"), path])
    tail = (out.strip().splitlines() or ["-"])[-1]
    results.append(("check_style.py %s" % os.path.basename(path), code == 0, tail, "E15"))


def main():
    ap = argparse.ArgumentParser(description="技能包自动化回归检查（evals 可脚本部分）")
    ap.add_argument("--docx", default=None, help="可选：追加 check_format.py 排版校验（E16）")
    ap.add_argument("--text", default=None, help="可选：追加 check_style.py 段长自检（E15）")
    args = ap.parse_args()

    results = []
    check_frontmatter(results)
    check_links(results)
    check_scripts_compile(results)
    check_format_refs(results)
    check_format_lang(results)
    if args.docx:
        check_docx(results, args.docx)
    if args.text:
        check_text(results, args.text)

    fails = [r for r in results if not r[1]]
    print("%-42s %-5s %-34s %s" % ("检查项", "结果", "说明", "关联用例"))
    print("-" * 92)
    for name, ok, detail, ev in results:
        print("%-42s %-5s %-34s %s" % (name, "PASS" if ok else "FAIL", detail, ev))
    print("-" * 92)
    print("共 %d 项：PASS %d，FAIL %d" % (len(results), len(results) - len(fails), len(fails)))
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
