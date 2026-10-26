#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
omml.py —— LaTeX 子集 → OMML（Word 原生公式）转换器（纯标准库，零第三方依赖）

用途：把论文里的公式确定性地转成结构正确的 OMML，避免手写 XML 导致角标（上下标）
结构错误。生成后交给 python-docx 的 parse_xml 插入段落，或用 Word 打开核对。

支持的结构：上下标 sSub / sSup / sSubSup、分式 \\frac、根式 \\sqrt[?]、求和求积积分
\\sum \\prod \\int \\oint、希腊字母、常用运算符。

用法:
    python omml.py "x_i^2"                     # 行内公式，打印 OMML 片段
    python omml.py --display "\\frac{a}{b}"    # 行间公式（oMathPara）
    python omml.py --dump "\\sum_{i=1}^{n} x_i"  # 同时打印解析结构

    # 代码里（依赖 python-docx）:
    from omml import insert_omml
    insert_omml(paragraph, r"\\frac{a}{b}", display=True)
"""

import argparse
import re
import sys

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ϵ", "zeta": "ζ", "eta": "η", "theta": "θ", "iota": "ι",
    "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π",
    "rho": "ρ", "sigma": "σ", "tau": "τ", "upsilon": "υ", "phi": "φ",
    "varphi": "ϕ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}

SYMBOLS = {
    "times": "×", "cdot": "·", "div": "÷", "pm": "±", "mp": "∓",
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥", "neq": "≠", "ne": "≠",
    "approx": "≈", "equiv": "≡", "propto": "∝", "sim": "∼",
    "rightarrow": "→", "to": "→", "leftarrow": "←", "Rightarrow": "⇒",
    "leftrightarrow": "↔", "partial": "∂", "nabla": "∇", "infty": "∞",
    "in": "∈", "notin": "∉", "subset": "⊂", "supset": "⊃",
    "cup": "∪", "cap": "∩", "forall": "∀", "exists": "∃",
    "prime": "′", "ell": "ℓ", "degree": "°",
}

FUNCS = {"sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos",
         "arctan", "log", "ln", "lg", "exp", "lim", "max", "min", "det",
         "gcd", "mod", "sup", "inf"}

NARY = {"sum": "∑", "prod": "∏", "int": "∫", "oint": "∮",
        "bigcup": "⋃", "bigcap": "⋂"}

ACCENTS = {"bar": "bar", "hat": "^", "widehat": "^", "vec": "→",
           "dot": "˙", "ddot": "¨", "tilde": "~", "widetilde": "~"}


def esc(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def tokenize(s):
    toks = []
    i = 0
    while i < len(s):
        c = s[i]
        if c == "\\":
            m = re.match(r"\\([A-Za-z]+)", s[i:])
            if m:
                toks.append(("cmd", m.group(1)))
                i += m.end()
            else:
                toks.append(("sym", s[i + 1] if i + 1 < len(s) else ""))
                i += 2
        elif c in "{}^_":
            toks.append((c, c))
            i += 1
        elif c.isspace():
            i += 1
        elif c.isdigit():
            m = re.match(r"\d+(\.\d+)?", s[i:])
            toks.append(("num", m.group(0)))
            i += m.end()
        elif c.isalpha():
            m = re.match(r"[A-Za-z]+", s[i:])
            toks.append(("var", m.group(0)))
            i += m.end()
        else:
            toks.append(("op", c))
            i += 1
    return toks


class Parser:
    def __init__(self, toks):
        self.t = toks
        self.i = 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def next(self):
        tok = self.peek()
        self.i += 1
        return tok

    def parse_seq(self, stop=None):
        nodes = []
        while True:
            k, v = self.peek()
            if k is None or k == stop or v == stop:
                break
            nodes.append(self.parse_scripts(self.parse_atom()))
        return nodes

    def parse_scripts(self, atom):
        sub = sup = None
        while True:
            k, _v = self.peek()
            if k == "_":
                self.next()
                sub = self.parse_arg()
            elif k == "^":
                self.next()
                sup = self.parse_arg()
            else:
                break
        if sub is None and sup is None:
            return atom
        if atom["type"] == "nary":
            atom["sub"] = sub
            atom["sup"] = sup
            return atom
        if sub is not None and sup is not None:
            return {"type": "subsup", "base": atom, "sub": sub, "sup": sup}
        if sub is not None:
            return {"type": "sub", "base": atom, "sub": sub}
        return {"type": "sup", "base": atom, "sup": sup}

    def parse_arg(self):
        k, _v = self.peek()
        if k == "{":
            self.next()
            items = self.parse_seq(stop="}")
            if self.peek()[0] == "}":
                self.next()
            return {"type": "group", "items": items}
        return self.parse_atom()

    def parse_atom(self):
        k, v = self.peek()
        if k == "{":
            self.next()
            items = self.parse_seq(stop="}")
            if self.peek()[0] == "}":
                self.next()
            return {"type": "group", "items": items}
        if k == "cmd":
            self.next()
            return self.parse_cmd(v)
        if k == "num":
            self.next()
            return {"type": "run", "text": v, "italic": False}
        if k == "var":
            self.next()
            return {"type": "run", "text": v, "italic": len(v) == 1}
        if k in ("op", "sym"):
            self.next()
            return {"type": "run", "text": v, "italic": False}
        self.next()
        return {"type": "run", "text": "" if v is None else str(v), "italic": False}

    def parse_cmd(self, name):
        if name == "frac":
            return {"type": "frac", "num": self.parse_arg(), "den": self.parse_arg()}
        if name == "sqrt":
            deg = None
            k, v = self.peek()
            if k == "op" and v == "[":
                self.next()
                items = self.parse_seq(stop="]")
                if self.peek()[0] == "op" and self.peek()[1] == "]":
                    self.next()
                deg = {"type": "group", "items": items}
            return {"type": "rad", "deg": deg, "base": self.parse_arg()}
        if name in NARY:
            return {"type": "nary", "chr": NARY[name], "sub": None, "sup": None, "body": None}
        if name in ACCENTS:
            return {"type": "acc", "kind": ACCENTS[name], "base": self.parse_arg()}
        if name in ("left", "right"):
            return {"type": "group", "items": []}
        if name in FUNCS:
            return {"type": "run", "text": name, "italic": False}
        if name in GREEK:
            return {"type": "run", "text": GREEK[name], "italic": True}
        if name in SYMBOLS:
            return {"type": "run", "text": SYMBOLS[name], "italic": False}
        return {"type": "run", "text": name, "italic": True}


def run_xml(text, italic):
    rpr = "" if italic else '<m:rPr><m:sty m:val="p"/></m:rPr>'
    return '<m:r>%s<m:t xml:space="preserve">%s</m:t></m:r>' % (rpr, esc(text))


def gen(node):
    t = node["type"]
    if t == "run":
        return run_xml(node["text"], node["italic"])
    if t == "group":
        return "".join(gen(n) for n in node["items"])
    if t == "frac":
        return "<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>" % (
            gen(node["num"]), gen(node["den"]))
    if t == "rad":
        if node["deg"] is None:
            return ('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
                    '<m:deg/><m:e>%s</m:e></m:rad>') % gen(node["base"])
        return '<m:rad><m:deg>%s</m:deg><m:e>%s</m:e></m:rad>' % (
            gen(node["deg"]), gen(node["base"]))
    if t == "acc":
        if node["kind"] == "bar":
            return ('<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>'
                    '<m:e>%s</m:e></m:bar>') % gen(node["base"])
        return '<m:acc><m:accPr><m:chr m:val="%s"/></m:accPr><m:e>%s</m:e></m:acc>' % (
            esc(node["kind"]), gen(node["base"]))
    if t == "sub":
        return "<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>" % (
            gen(node["base"]), gen(node["sub"]))
    if t == "sup":
        return "<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>" % (
            gen(node["base"]), gen(node["sup"]))
    if t == "subsup":
        return ("<m:sSubSup><m:e>%s</m:e><m:sub>%s</m:sub><m:sup>%s</m:sup>"
                "</m:sSubSup>") % (gen(node["base"]), gen(node["sub"]), gen(node["sup"]))
    if t == "nary":
        sub = gen(node["sub"]) if node["sub"] else ""
        sup = gen(node["sup"]) if node["sup"] else ""
        body = gen(node["body"]) if node["body"] else ""
        return ('<m:nary><m:naryPr><m:chr m:val="%s"/><m:limLoc m:val="undOvr"/>'
                '</m:naryPr><m:sub>%s</m:sub><m:sup>%s</m:sup><m:e>%s</m:e>'
                '</m:nary>') % (esc(node["chr"]), sub, sup, body)
    return ""


def parse(latex):
    return Parser(tokenize(latex)).parse_seq()


def latex_to_omml(latex, display=False):
    """返回可被 python-docx parse_xml 解析的 OMML 片段（自带命名空间声明）。"""
    inner = "".join(gen(n) for n in parse(latex))
    ns = 'xmlns:m="%s" xmlns:w="%s"' % (M, W)
    if display:
        return "<m:oMathPara %s><m:oMath>%s</m:oMath></m:oMathPara>" % (ns, inner)
    return "<m:oMath %s>%s</m:oMath>" % (ns, inner)


def insert_omml(paragraph, latex, display=False):
    """把 LaTeX 公式作为 OMML 插入 python-docx 段落。依赖 python-docx。"""
    from docx.oxml import parse_xml
    paragraph._p.append(parse_xml(latex_to_omml(latex, display)))


def _dump(node, depth=0):
    pad = "  " * depth
    t = node["type"]
    if t == "run":
        return ["%srun %r italic=%s" % (pad, node["text"], node["italic"])]
    lines = ["%s%s" % (pad, t)]
    for key in ("base", "num", "den", "deg", "sub", "sup", "body"):
        if node.get(key):
            lines += _dump(node[key], depth + 1)
    for child in node.get("items", []):
        lines += _dump(child, depth + 1)
    return lines


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="LaTeX 子集 → OMML（纯标准库）")
    ap.add_argument("latex", nargs="*", help="LaTeX 公式，可多个；省略则从 stdin 逐行读")
    ap.add_argument("--display", action="store_true", help="行间公式（oMathPara）")
    ap.add_argument("--dump", action="store_true", help="同时打印解析结构")
    args = ap.parse_args()

    formulas = args.latex if args.latex else [ln.strip() for ln in sys.stdin if ln.strip()]
    for f in formulas:
        if args.dump:
            for line in _dump({"type": "group", "items": parse(f)}):
                print(line)
            print("-" * 40)
        print(latex_to_omml(f, args.display))


if __name__ == "__main__":
    main()
