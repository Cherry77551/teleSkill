#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_docx.py —— 确定性 .docx 生成脚手架（纯标准库 + 同目录 omml.py，零第三方依赖）

把「论文规格 JSON」确定性地生成可交付的 .docx：标题/作者、摘要、真实 TOC 域、
自动编号章节、正文（支持行内公式 $...$ 与上标引用 [n]）、行间公式（居中+编号靠右）、
表格（全线 / 三线）、插图（PNG/JPEG 内嵌）、参考文献（悬挂缩进）。

用法:
    python build_docx.py spec.json -o 论文.docx
    python build_docx.py spec.json -o 论文.docx --check --strict
    python build_docx.py spec.json -o 论文.docx --check --lang en

规格 JSON 示例见 assets/templates/paper-spec.example.json。字段：
    lang         "zh" | "en"（默认 zh；决定自动编号与引用上标）
    title        论文题目（Title 样式，居中）
    author/date  可选，居中
    toc          是否插入目录（默认 true；插入真实 TOC 域）
    abstract     {"zh": "...", "en": "..."} 或字符串
    keywords     可选，关键词串
    sections     [{heading, level, numbered, paragraphs[], equations[], tables[], figures[]}]
    references   ["[1] 作者. 题名[J]. 刊名, 年, 卷(期): 页.", ...]

退出码: 0 = 生成成功（--check 时以 check_format 结果为准）；2 = 输入/路径错误。
"""

import argparse
import json
import os
import re
import struct
import subprocess
import sys
import zipfile
from xml.sax.saxutils import escape

try:
    from omml import latex_to_omml
except ImportError:  # 作为包导入时补 scripts/ 到 sys.path
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from omml import latex_to_omml

try:
    from _paths import PathError, resolve_input, resolve_output
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from _paths import PathError, resolve_input, resolve_output

# ---- 页面几何（A4，四边 2.5cm；单位 twip = 1/20 pt）----
A4_W, A4_H = 11906, 16838
MARGIN = 1417
CONTENT_W = A4_W - 2 * MARGIN
CENTER_TAB = CONTENT_W // 2
RIGHT_TAB = CONTENT_W
EMU_PER_CM = 360000

CN_DIGITS = "零一二三四五六七八九"


def e(text):
    return escape(str(text))


def a(text):
    return (str(text).replace("&", "&amp;").replace('"', "&quot;")
            .replace("<", "&lt;").replace(">", "&gt;"))


def cn_number(n):
    if n <= 0:
        return str(n)
    if n <= 10:
        return "十" if n == 10 else CN_DIGITS[n]
    if n < 20:
        return "十" + CN_DIGITS[n - 10]
    if n < 100:
        tens, ones = divmod(n, 10)
        return CN_DIGITS[tens] + "十" + (CN_DIGITS[ones] if ones else "")
    return str(n)


# ---- 包内固定部件 ----
CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:updateFields w:val="true"/>
  <w:compat/>
</w:settings>'''

APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
       '<Application>academic-writer</Application></Properties>')

DOC_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')

SECTPR = ('<w:sectPr><w:pgSz w:w="%d" w:h="%d"/>'
          '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
          'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
          % (A4_W, A4_H, MARGIN, MARGIN, MARGIN, MARGIN))

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:before="240" w:after="240" w:line="360" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>
      <w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="180" w:after="60" w:line="360" w:lineRule="auto"/>
      <w:jc w:val="left"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="120" w:after="60" w:line="360" w:lineRule="auto"/>
      <w:jc w:val="left"/><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
</w:styles>'''


def text_run(text, size=24, cn="宋体", en="Times New Roman",
             bold=False, italic=False, sup=False):
    rpr = ['<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
           % (a(en), a(en), a(cn), a(en)),
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size, size)]
    if bold:
        rpr.append("<w:b/><w:bCs/>")
    if italic:
        rpr.append("<w:i/><w:iCs/>")
    if sup:
        rpr.append('<w:vertAlign w:val="superscript"/>')
    return '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (
        "".join(rpr), e(text))


def para(inner, style=None, jc=None, first_line_chars=200, line=360,
         left=None, hanging=None, tabs=None, outline=None,
         keep_next=False, page_break=False):
    parts = []
    if style:
        parts.append('<w:pStyle w:val="%s"/>' % a(style))
    if keep_next:
        parts.append("<w:keepNext/>")
    if page_break:
        parts.append("<w:pageBreakBefore/>")
    if tabs:
        parts.append(tabs)
    parts.append('<w:spacing w:line="%d" w:lineRule="auto"/>' % line)
    ind = []
    if left is not None:
        ind.append('w:left="%d"' % left)
    if hanging is not None:
        ind.append('w:hanging="%d"' % hanging)
    if first_line_chars:
        ind.append('w:firstLineChars="%d"' % first_line_chars)
    if ind:
        parts.append('<w:ind %s/>' % " ".join(ind))
    if jc:
        parts.append('<w:jc w:val="%s"/>' % a(jc))
    if outline is not None:
        parts.append('<w:outlineLvl w:val="%d"/>' % outline)
    return "<w:p><w:pPr>%s</w:pPr>%s</w:p>" % ("".join(parts), inner)


def page_break():
    return "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"


def image_size_from_bytes(data):
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", data[16:24])
        return w, h, "png"
    if data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return w, h, "jpeg"
            if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
                i += 2
                continue
            seglen = struct.unpack(">H", data[i + 2:i + 4])[0]
            i += 2 + seglen
    raise ValueError("不支持的图片格式（仅 png/jpeg）")


def drawing_xml(rid, pic_id, name, cx, cy):
    return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="Picture %d"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
            % (cx, cy, pic_id, pic_id, pic_id, a(name), rid, cx, cy))


class Builder:
    def __init__(self, spec, base_dir="."):
        self.spec = spec
        self.base_dir = base_dir
        self.lang = (spec.get("lang") or "zh").lower()
        self.body = []
        self.images = []
        self.rels = []
        self.h = [0, 0, 0]
        self.eq_no = 0
        self.pic_id = 0
        self.sup_citations = (self.lang != "en")

    def numbered_heading(self, heading, level):
        if (re.match(r"^[一二三四五六七八九十]+、", heading)
                or re.match(r"^\d+(?:\.\d+)*[\.\s]", heading)):
            return heading
        if level == 1:
            self.h = [self.h[0] + 1, 0, 0]
            if self.lang == "en":
                return "%d %s" % (self.h[0], heading)
            return "%s、%s" % (cn_number(self.h[0]), heading)
        if level == 2:
            self.h[1] += 1
            self.h[2] = 0
            return "%d.%d %s" % (self.h[0], self.h[1], heading)
        self.h[2] += 1
        return "%d.%d.%d %s" % (self.h[0], self.h[1], self.h[2], heading)

    def heading(self, text, level, numbered=True):
        if numbered:
            text = self.numbered_heading(text, level)
        return para(text_run(text, size={1: 32, 2: 28, 3: 24}[level],
                             cn="黑体", bold=True),
                    style="Heading%d" % level,
                    jc="center" if level == 1 else "left",
                    first_line_chars=0, outline=level - 1)

    def inline(self, text, size=24):
        out, pos = [], 0
        for m in re.finditer(r"(\$[^$]+\$)|(\[\d+(?:[-,]\d+)*\])", text):
            if m.start() > pos:
                out.append(text_run(text[pos:m.start()], size=size))
            if m.group(1):
                out.append(latex_to_omml(m.group(1)[1:-1], display=False))
            else:
                out.append(text_run(m.group(2), size=size, sup=self.sup_citations))
            pos = m.end()
        if pos < len(text):
            out.append(text_run(text[pos:], size=size))
        return "".join(out)

    def equation(self, eq):
        if isinstance(eq, str):
            eq = {"latex": eq}
        self.eq_no += 1
        num = eq.get("number", self.eq_no)
        omml = latex_to_omml(eq["latex"], display=False)
        tabs = ('<w:tabs><w:tab w:val="center" w:pos="%d"/>'
                '<w:tab w:val="right" w:pos="%d"/></w:tabs>' % (CENTER_TAB, RIGHT_TAB))
        inner = ("<w:r><w:tab/></w:r>" + omml
                 + "<w:r><w:tab/></w:r>" + text_run("(%s)" % num))
        return para(inner, first_line_chars=0, tabs=tabs)

    def _borders(self, kind):
        if kind == "three-line":
            # 三线表：表级只保留顶线/底线；栏目线由表头单元格底边画，行间无横线
            return ('<w:tblBorders>'
                    '<w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>'
                    '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/>'
                    '</w:tblBorders>')
        return ('<w:tblBorders>'
                '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '</w:tblBorders>')

    def _row(self, cells, header=False, sep=False):
        tcs = []
        for c in cells:
            inner = (text_run(str(c), size=21, bold=True) if header
                     else self.inline(str(c), size=21))
            tcb = ('<w:tcBorders><w:bottom w:val="single" w:sz="4" '
                   'w:space="0" w:color="auto"/></w:tcBorders>') if sep else ''
            tcs.append('<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>%s</w:tcPr>'
                       '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/>'
                       '<w:jc w:val="center"/></w:pPr>%s</w:p></w:tc>' % (tcb, inner))
        return "<w:tr>%s</w:tr>" % "".join(tcs)

    def table(self, spec):
        parts = []
        cap = spec.get("caption")
        if cap:
            parts.append(para(text_run(cap, size=21, bold=True),
                              jc="center", first_line_chars=0, keep_next=True))
        kind = spec.get("type", "grid")
        parts.append('<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>%s'
                     '<w:tblLayout w:type="autofit"/></w:tblPr>' % self._borders(kind))
        header = spec.get("header") or []
        if header:
            parts.append(self._row(header, header=True, sep=(kind == "three-line")))
        for row in spec.get("rows", []) or []:
            parts.append(self._row(row, header=False))
        parts.append("</w:tbl>")
        parts.append(para("", first_line_chars=0))
        return parts

    def _add_image(self, path, width_cm):
        raw = path if os.path.isabs(path) else os.path.join(self.base_dir, path)
        with open(raw, "rb") as fh:
            data = fh.read()
        w, h, ext = image_size_from_bytes(data)
        idx = len(self.images) + 1
        name = "image%d.%s" % (idx, ext)
        rid = "rIdImg%d" % idx
        self.images.append((name, data))
        self.rels.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
                         'officeDocument/2006/relationships/image" Target="media/%s"/>' % (rid, name))
        cx = int(float(width_cm) * EMU_PER_CM)
        cy = int(cx * h / w) if w else int(7 * EMU_PER_CM)
        return rid, name, cx, cy

    def figure(self, spec):
        parts = []
        if spec.get("path"):
            rid, name, cx, cy = self._add_image(spec["path"], spec.get("width_cm", 14))
            self.pic_id += 1
            parts.append(para(drawing_xml(rid, self.pic_id, name, cx, cy),
                              jc="center", first_line_chars=0))
        if spec.get("caption"):
            parts.append(para(text_run(spec["caption"], size=21),
                              jc="center", first_line_chars=0))
        return parts

    def title_block(self):
        s = self.spec
        if s.get("title"):
            self.body.append(para(text_run(s["title"], size=44, cn="黑体", bold=True),
                                  style="Title", jc="center", first_line_chars=0))
        for key in ("author", "date"):
            if s.get(key):
                self.body.append(para(text_run(s[key], size=24),
                                      jc="center", first_line_chars=0))

    def abstract_block(self):
        ab = self.spec.get("abstract")
        if not ab:
            return
        if isinstance(ab, str):
            ab = {"zh": ab}
        if self.lang == "en":
            if ab.get("en"):
                self.body.append(self.heading("Abstract", 1, numbered=False))
                self.body.append(para(self.inline(ab["en"]), jc="both", first_line_chars=0))
            if ab.get("zh"):
                self.body.append(self.heading("摘要", 1, numbered=False))
                self.body.append(para(self.inline(ab["zh"]), jc="both", first_line_chars=0))
        else:
            self.body.append(self.heading("摘要", 1, numbered=False))
            if ab.get("zh"):
                self.body.append(para(self.inline(ab["zh"]), jc="both", first_line_chars=200))
            if ab.get("en"):
                self.body.append(self.heading("Abstract", 1, numbered=False))
                self.body.append(para(self.inline(ab["en"]), jc="both", first_line_chars=0))
        if self.spec.get("keywords"):
            label = "Keywords: " if self.lang == "en" else "关键词："
            self.body.append(para(text_run(label, size=24, bold=True)
                                  + self.inline(self.spec["keywords"]),
                                  jc="both", first_line_chars=0))

    def toc_block(self):
        if not self.spec.get("toc", True):
            return
        label = "Contents" if self.lang == "en" else "目录"
        hint = ("Right-click and Update Field to build the TOC."
                if self.lang == "en" else "在 Word 中右键“更新域”生成目录")
        self.body.append(page_break())
        self.body.append(para(text_run(label, size=32, cn="黑体", bold=True),
                              jc="center", first_line_chars=0))
        self.body.append(para(
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>' % e(hint),
            jc="left", first_line_chars=0))
        self.body.append(page_break())

    def sections(self):
        for sec in self.spec.get("sections", []) or []:
            level = int(sec.get("level", 1))
            if sec.get("heading"):
                self.body.append(self.heading(sec["heading"], level,
                                              numbered=sec.get("numbered", True)))
            for ptext in sec.get("paragraphs", []) or []:
                self.body.append(para(self.inline(ptext), jc="both", first_line_chars=200))
            for eq in sec.get("equations", []) or []:
                self.body.append(self.equation(eq))
            for tbl in sec.get("tables", []) or []:
                self.body.extend(self.table(tbl))
            for fig in sec.get("figures", []) or []:
                self.body.extend(self.figure(fig))

    def references(self):
        refs = self.spec.get("references") or []
        if not refs:
            return
        self.body.append(self.heading("References" if self.lang == "en" else "参考文献",
                                      1, numbered=False))
        for r in refs:
            if isinstance(r, dict):
                r = r.get("text", "")
            self.body.append(para(self.inline(str(r)), jc="left",
                                  first_line_chars=0, left=480, hanging=480))

    def build(self, out_path):
        self.title_block()
        self.abstract_block()
        self.toc_block()
        self.sections()
        self.references()
        document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<w:document %s><w:body>%s%s</w:body></w:document>'
                    % (DOC_NS, "".join(self.body), SECTPR))
        doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                    '%s</Relationships>' % "".join(self.rels))
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CONTENT_TYPES)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/document.xml", document)
            z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", SETTINGS)
            z.writestr("word/_rels/document.xml.rels", doc_rels)
            z.writestr("docProps/core.xml",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                       'xmlns:dc="http://purl.org/dc/elements/1.1/">'
                       '<dc:title>%s</dc:title><dc:creator>%s</dc:creator></cp:coreProperties>'
                       % (e(self.spec.get("title", "")), e(self.spec.get("author", ""))))
            z.writestr("docProps/app.xml", APP)
            for name, data in self.images:
                z.writestr("word/media/%s" % name, data)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="论文规格 JSON → 确定性 .docx（纯标准库）")
    ap.add_argument("spec", help="JSON 规格文件")
    ap.add_argument("-o", "--out", default=None, help="输出 .docx（默认 spec 同名 .docx）")
    ap.add_argument("--check", action="store_true", help="生成后立即运行 check_format.py")
    ap.add_argument("--strict", action="store_true", help="--check 时用 strict 完整模式")
    ap.add_argument("--lang", choices=["auto", "zh", "en"], default=None,
                    help="--check 时传给 check_format（默认取 spec.lang）")
    ap.add_argument("--template", default=None, help="--check 时传给 check_format 的模板 docx")
    args = ap.parse_args()

    try:
        spec_path = resolve_input(args.spec)
        with open(spec_path, encoding="utf-8") as fh:
            spec = json.load(fh)
    except PathError as exc:
        print("[ERROR] 路径校验失败: %s" % exc)
        sys.exit(2)
    except Exception as exc:
        print("[ERROR] 无法读取 spec: %s" % exc)
        sys.exit(2)

    out = args.out or (os.path.splitext(spec_path)[0] + ".docx")
    try:
        out = resolve_output(out)
    except PathError as exc:
        print("[ERROR] 路径校验失败: %s" % exc)
        sys.exit(2)

    Builder(spec, os.path.dirname(spec_path)).build(out)
    print("已生成: %s" % out)

    if args.check:
        cmd = [sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                            "check_format.py"), out]
        if args.strict:
            cmd.append("--strict")
        lang = args.lang or (spec.get("lang") if spec.get("lang") in ("zh", "en") else None)
        if lang:
            cmd += ["--lang", lang]
        if args.template:
            cmd += ["--template", args.template]
        sys.exit(subprocess.call(cmd))


if __name__ == "__main__":
    main()
