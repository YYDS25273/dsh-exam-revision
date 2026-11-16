# -*- coding: utf-8 -*-
"""把题库 JSON 排版成 Word：题面 / 答案 / 解析 / 考点 / 易错 五色分区。"""

import datetime
import json
import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from omml_helper import omath

SONG = "宋体"
HEI = "黑体"
LATIN = "Times New Roman"

# 题库数据（由 JSON 载入）与封面文案
CHAPTERS = []
COVER = {
    "text1": "继电保护",
    "text2": "复 习 题 详 解",
    "subtitle": "逐题详解　·　分步推导　·　考点归纳　·　易错提醒",
    "footer": "",
}
RED = RGBColor(0xC0, 0x00, 0x00)
GRAY = RGBColor(0x59, 0x59, 0x59)
DARK = RGBColor(0x1F, 0x33, 0x1F)

# 各板块配色（十六进制用于底纹/边框，RGBColor 用于文字）
C_STEM_BG, C_STEM_BAR, C_STEM_TXT = "EAF1F8", "2E75B6", RGBColor(0x1F, 0x4E, 0x79)
C_ANS_BG, C_ANS_BAR, C_ANS_TXT = "EDF7EF", "2E7D32", RGBColor(0x1B, 0x6E, 0x2F)
C_STEP_BG, C_STEP_BAR, C_STEP_TXT = "E4F3F5", "00838F", RGBColor(0x00, 0x65, 0x70)
C_PT_BG, C_PT_BAR, C_PT_TXT = "F4EFFA", "6A4C93", RGBColor(0x4B, 0x2E, 0x7A)
C_TRAP_BG, C_TRAP_BAR, C_TRAP_TXT = "FFF4E5", "E65100", RGBColor(0xB2, 0x4A, 0x00)
C_IDX_BG, C_IDX_BAR = "F5F5F5", "9E9E9E"

OUT = "继电保护复习题详解.docx"

_SPLIT_MATH = re.compile(r"(\$[^$]*\$)")
_SPLIT_BOLD = re.compile(r"(\*\*[^*]+\*\*)")


def set_run_font(run, ea=SONG, latin=LATIN, size=None, bold=None, color=None):
    run.font.name = latin
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), ea)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


# ── 段落视觉元素：底纹 / 边框 ──────────────────────────────────────────────
# w:pPr 子元素有固定顺序，插入时必须按序，否则 Word 可能报错
_PPR_ORDER = [
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
]


def _ppr_put(p, el):
    """按 schema 顺序把 el 插入段落的 pPr。"""
    pPr = p._p.get_or_add_pPr()
    tag = el.tag
    idx = next((i for i, t in enumerate(_PPR_ORDER) if qn(t) == tag), len(_PPR_ORDER))
    for child in list(pPr):
        cidx = next((i for i, t in enumerate(_PPR_ORDER) if qn(t) == child.tag),
                    len(_PPR_ORDER))
        if cidx > idx:
            child.addprevious(el)
            return el
    pPr.append(el)
    return el


def shade(p, fill):
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:color"), "auto")
    el.set(qn("w:fill"), fill)
    _ppr_put(p, el)


def border(p, sides, color, sz=8, space=4):
    """给段落加边框；sides 为 'left' / 'right' / 'top' / 'bottom' 的可迭代集合。"""
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        _ppr_put(p, pBdr)
    for side in sides:
        el = OxmlElement("w:" + side)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), str(space))
        el.set(qn("w:color"), color)
        pBdr.append(el)


def add_rich(paragraph, text, ea=SONG, size=10.5, color=None, bold=False):
    """支持 **加粗** 与 $公式$ 的行内富文本。"""
    for chunk in _SPLIT_MATH.split(text):
        if not chunk:
            continue
        if chunk.startswith("$") and chunk.endswith("$") and len(chunk) > 2:
            paragraph._p.append(omath(chunk[1:-1]))
            continue
        for piece in _SPLIT_BOLD.split(chunk):
            if not piece:
                continue
            is_bold = piece.startswith("**") and piece.endswith("**")
            seg = piece[2:-2] if is_bold else piece
            run = paragraph.add_run(seg)
            set_run_font(run, ea=ea, size=size, bold=(bold or is_bold), color=color)


def setup(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.1)
    sec.left_margin = Cm(2.4)
    sec.right_margin = Cm(2.1)
    sec.footer_distance = Cm(1.2)

    normal = doc.styles["Normal"]
    normal.font.name = LATIN
    normal.font.size = Pt(10.5)
    normal.element.get_or_add_rPr()
    rfonts = normal.element.rPr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        normal.element.rPr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), SONG)
    rfonts.set(qn("w:ascii"), LATIN)
    rfonts.set(qn("w:hAnsi"), LATIN)
    pf = normal.paragraph_format
    pf.line_spacing = 1.38
    pf.space_before = Pt(0)
    pf.space_after = Pt(3)

    heads = {"Heading 1": (17, 16, 10), "Heading 2": (13, 14, 8),
             "Heading 3": (11, 12, 4)}
    for name, (size, before, after) in heads.items():
        st = doc.styles[name]
        st.font.name = LATIN
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        rf.set(qn("w:eastAsia"), HEI)
        rf.set(qn("w:ascii"), HEI)
        rf.set(qn("w:hAnsi"), HEI)
        hpf = st.paragraph_format
        hpf.space_before = Pt(before)
        hpf.space_after = Pt(after)
        hpf.line_spacing = 1.2
        hpf.keep_with_next = True

    # 页脚页码
    footer_p = sec.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run()
    set_run_font(run, ea=SONG, size=9, color=GRAY)
    for tag, attr in (("w:fldChar", "begin"), ("w:instrText", "PAGE"),
                      ("w:fldChar", "separate"), ("w:t", "1"),
                      ("w:fldChar", "end")):
        el = OxmlElement(tag)
        if tag == "w:fldChar":
            el.set(qn("w:fldCharType"), attr)
        else:
            if tag == "w:instrText":
                el.set(qn("xml:space"), "preserve")
                el.text = " PAGE "
            else:
                el.text = attr
        run._r.append(el)

    # 打开时自动更新域（目录）
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    settings.append(upd)

    # 提升到 Word 2013+ 模式，去掉标题栏的「兼容性模式」
    for cs in settings.iter(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            cs.set(qn("w:val"), "15")
    for el in settings.iter(qn("w:themeFontLang")):
        el.set(qn("w:eastAsia"), "zh-CN")


def add_toc(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-2" \\h \\z \\u '
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "（目录：打开文档后按 Ctrl+A 再按 F9 可刷新）"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, txt, end):
        run._r.append(el)


def cover(doc, total):
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(COVER.get("text1", "继电保护"))
    set_run_font(r, ea=HEI, latin=HEI, size=34, bold=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run(COVER.get("text2", "复 习 题 详 解"))
    set_run_font(r, ea=HEI, latin=HEI, size=34, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(COVER.get("subtitle",
                            "逐题详解　·　分步推导　·　考点归纳　·　易错提醒"))
    set_run_font(r, ea=SONG, size=12, color=GRAY)
    p.paragraph_format.space_after = Pt(60)

    box = doc.add_paragraph()
    box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = box.add_run(
        "本册收录：%d 题\n整理日期：%s" % (total, datetime.date.today().strftime("%Y 年 %m 月 %d 日")))
    set_run_font(r, ea=SONG, size=10.5, color=GRAY)
    box.paragraph_format.line_spacing = 1.8


def heading_chapter(doc, text, first=False):
    p = doc.add_heading(text, level=1)
    if not first:
        p.paragraph_format.page_break_before = True
    border(p, ("bottom",), C_STEM_BAR, sz=12, space=6)
    p.paragraph_format.space_after = Pt(12)
    return p


def heading_section(doc, text):
    p = doc.add_heading(text, level=2)
    shade(p, "EDF2F7")
    border(p, ("bottom",), "A9BFD6", sz=6, space=4)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(0.15)
    return p


LABELS = {
    "single": "【解析】",
    "multi": "【解析】",
    "judge": "【判断依据】",
    "short": "【答题要点】",
    "blank": "【解析】",
}


def question_block(doc, q):
    kind = q.get("kind", "single")
    options = q.get("options") or {}

    # ── 题干：浅蓝底纹 + 左侧蓝色竖条 ──
    p = doc.add_paragraph(style="Heading 3")
    shade(p, C_STEM_BG)
    border(p, ("left",), C_STEM_BAR, sz=24, space=6)
    pf = p.paragraph_format
    pf.space_before = Pt(15)
    pf.space_after = Pt(3)
    pf.left_indent = Cm(0.30)
    pf.right_indent = Cm(0.15)
    pf.keep_with_next = True
    num = p.add_run("%d．" % q["id"])
    set_run_font(num, ea=HEI, latin=HEI, size=11, bold=True, color=C_STEM_TXT)
    add_rich(p, q["stem"], ea=HEI, size=11, bold=True, color=C_STEM_TXT)

    # ── 选项：统一缩进，不加装饰 ──
    if kind == "judge":
        op = doc.add_paragraph()
        op.paragraph_format.left_indent = Cm(1.10)
        op.paragraph_format.space_before = Pt(2)
        op.paragraph_format.space_after = Pt(2)
        op.paragraph_format.keep_with_next = True
        r = op.add_run("A．对　　　B．错")
        set_run_font(r, ea=SONG, size=10.5)
    elif kind not in ("short", "blank"):
        for key in "ABCD":
            text = options.get(key)
            if not text:
                continue
            op = doc.add_paragraph()
            op.paragraph_format.left_indent = Cm(1.15)
            op.paragraph_format.first_line_indent = Cm(-0.55)
            op.paragraph_format.space_after = Pt(1)
            op.paragraph_format.line_spacing = 1.28
            op.paragraph_format.keep_with_next = True
            r = op.add_run("%s．" % key)
            set_run_font(r, ea=SONG, size=10.5)
            add_rich(op, text, size=10.5)

    # ── 答案：红色实线框 + 浅红底纹 ──
    ans = q["answer"]
    ap = doc.add_paragraph()
    shade(ap, C_ANS_BG)
    border(ap, ("top", "bottom", "left", "right"), C_ANS_BAR, sz=6, space=3)
    apf = ap.paragraph_format
    apf.space_before = Pt(7)
    apf.space_after = Pt(6)
    apf.left_indent = Cm(0.30)
    apf.right_indent = Cm(0.15)
    apf.keep_with_next = True
    r = ap.add_run("【答案】")
    set_run_font(r, ea=HEI, latin=HEI, size=10.5, bold=True, color=C_ANS_TXT)
    if kind == "blank":
        r = ap.add_run(ans)
        set_run_font(r, ea=SONG, size=10.5, bold=True, color=C_ANS_TXT)
    elif kind == "short":
        r = ap.add_run(ans)
        set_run_font(r, ea=SONG, size=10.5, bold=False, color=C_ANS_TXT)
    else:
        r = ap.add_run(ans)
        set_run_font(r, ea=LATIN, latin=LATIN, size=12.5, bold=True, color=C_ANS_TXT)
        if kind == "judge":
            r = ap.add_run("（%s）" % options.get(ans, ""))
            set_run_font(r, ea=SONG, size=10.5, bold=True, color=C_ANS_TXT)
        elif len(ans) == 1 and options.get(ans):
            r = ap.add_run("　" + options[ans])
            set_run_font(r, ea=SONG, size=10.5, bold=True, color=C_ANS_TXT)

    # ── 解析：青色底纹 + 左侧青色竖条，整块连通 ──
    last_step = len(q["steps"]) - 1
    for i, para in enumerate(q["steps"]):
        sp = doc.add_paragraph()
        shade(sp, C_STEP_BG)
        border(sp, ("left",), C_STEP_BAR, sz=24, space=6)
        spf = sp.paragraph_format
        spf.left_indent = Cm(0.30)
        spf.right_indent = Cm(0.15)
        spf.space_before = Pt(5 if i == 0 else 0)
        spf.space_after = Pt(7 if i == last_step else 0)
        spf.line_spacing = 1.42
        spf.first_line_indent = Cm(0)
        if i == 0:
            r = sp.add_run(LABELS.get(kind, "【解析】"))
            set_run_font(r, ea=HEI, latin=HEI, size=10.5, bold=True, color=C_STEP_TXT)
        add_rich(sp, para, size=10.5)

    # ── 考点 / 易错：整块底纹 + 左侧彩色竖条 ──
    def labelled(label, items, mark, bg, bar, txt, gap_before=9):
        if not items:
            return
        lp = doc.add_paragraph()
        shade(lp, bg)
        border(lp, ("left",), bar, sz=24, space=6)
        lpf = lp.paragraph_format
        lpf.space_before = Pt(gap_before)
        lpf.space_after = Pt(0)
        lpf.left_indent = Cm(0.30)
        lpf.right_indent = Cm(0.15)
        lpf.keep_with_next = True
        r = lp.add_run(label)
        set_run_font(r, ea=HEI, latin=HEI, size=10, bold=True, color=txt)

        last = len(items) - 1
        for i, it in enumerate(items):
            ip = doc.add_paragraph()
            shade(ip, bg)
            border(ip, ("left",), bar, sz=24, space=6)
            ipf = ip.paragraph_format
            ipf.left_indent = Cm(0.30)
            ipf.right_indent = Cm(0.15)
            ipf.space_before = Pt(0)
            ipf.space_after = Pt(7 if i == last else 0)
            ipf.line_spacing = 1.32
            r = ip.add_run(mark + " ")
            set_run_font(r, ea=SONG, size=10, color=txt)
            add_rich(ip, it, size=10)

    labelled("【考点】", q.get("points"), "▪", C_PT_BG, C_PT_BAR, C_PT_TXT)
    labelled("【易错·应试提醒】", q.get("traps"), "▲",
             C_TRAP_BG, C_TRAP_BAR, C_TRAP_TXT, gap_before=8)


def _cell(cell, text, width_cm, center=True, size=9):
    cell.width = Cm(width_cm)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.05
    r = p.add_run(text)
    set_run_font(r, ea=SONG, size=size)


def answer_index(doc):
    """全书末尾的答案速查表。"""
    p = doc.add_heading("附录　答案速查表", level=1)
    p.paragraph_format.page_break_before = True
    note = doc.add_paragraph()
    note.paragraph_format.space_after = Pt(4)
    r = note.add_run("本表只列答案，便于自测后快速核对；每题的分步解析、考点与易错提醒见正文对应题号。")
    set_run_font(r, ea=SONG, size=9, color=GRAY)

    for ch in CHAPTERS:
        doc.add_heading(ch["title"], level=2)
        for sec in ch["sections"]:
            qs = sec["questions"]
            kind = qs[0].get("kind", "single")

            lp = doc.add_paragraph()
            lp.paragraph_format.space_before = Pt(7)
            lp.paragraph_format.space_after = Pt(2)
            lp.paragraph_format.keep_with_next = True
            r = lp.add_run("%s（%d 题）" % (sec["title"].split("（")[0], len(qs)))
            set_run_font(r, ea=HEI, latin=HEI, size=10, bold=True)

            if kind == "short":
                t = doc.add_table(rows=0, cols=2)
                t.style = "Table Grid"
                t.autofit = False
                t.alignment = WD_TABLE_ALIGNMENT.CENTER
                for q in qs:
                    cells = t.add_row().cells
                    _cell(cells[0], str(q["id"]), 1.2, size=8.5)
                    _cell(cells[1], q["answer"], 15.3, center=False, size=8.5)
                for col, w in zip(t.columns, (1.2, 15.3)):
                    col.width = Cm(w)
            else:
                per = 5
                n = len(qs)
                rows = (n + per - 1) // per
                t = doc.add_table(rows=0, cols=per * 2)
                t.style = "Table Grid"
                t.autofit = False
                t.alignment = WD_TABLE_ALIGNMENT.CENTER
                for i in range(rows):
                    cells = t.add_row().cells
                    for j in range(per):
                        idx = i * per + j
                        if idx < n:
                            q = qs[idx]
                            ans = q["answer"]
                            if q.get("kind") == "judge":
                                ans = q["options"].get(ans, ans)
                            _cell(cells[j * 2], str(q["id"]), 1.2)
                            _cell(cells[j * 2 + 1], ans, 2.1)
                        else:
                            _cell(cells[j * 2], "", 1.2)
                            _cell(cells[j * 2 + 1], "", 2.1)
                for i, col in enumerate(t.columns):
                    col.width = Cm(1.2 if i % 2 == 0 else 2.1)


KINDS = ("single", "multi", "judge", "blank", "short")


def validate(data):
    """尽早报错，并指出是哪一道题缺字段，方便调用方修数据。"""
    chapters = data.get("chapters")
    if not isinstance(chapters, list) or not chapters:
        raise ValueError("题库 JSON 缺少 chapters 数组")
    total = 0
    for ci, ch in enumerate(chapters, 1):
        if "title" not in ch:
            raise ValueError("第 %d 章缺少 title" % ci)
        sections = ch.get("sections")
        if not isinstance(sections, list) or not sections:
            raise ValueError("第 %d 章（%s）缺少 sections" % (ci, ch["title"]))
        for si, sec in enumerate(sections, 1):
            if "title" not in sec:
                raise ValueError("%s 第 %d 节缺少 title" % (ch["title"], si))
            qs = sec.get("questions")
            if not isinstance(qs, list) or not qs:
                raise ValueError("%s / %s 没有题目" % (ch["title"], sec["title"]))
            for q in qs:
                qid = q.get("id", "?")
                kind = q.get("kind", "single")
                if kind not in KINDS:
                    raise ValueError("题 %s 的 kind 非法：%s（可选 %s）"
                                     % (qid, kind, "/".join(KINDS)))
                for field in ("id", "stem", "answer"):
                    if q.get(field) is None:
                        raise ValueError("题 %s 缺少 %s" % (qid, field))
                if not q.get("steps"):
                    raise ValueError("题 %s 缺少 steps（解析段落）" % qid)
                if kind in ("single", "multi", "judge") and not q.get("options"):
                    raise ValueError("题 %s 为选择题但缺少 options" % qid)
                total += 1
    return total


def build(data, out):
    global CHAPTERS, COVER
    CHAPTERS = data["chapters"]
    COVER.update(data.get("cover") or {})
    total = validate(data)

    doc = Document()
    setup(doc)
    cover(doc, total)

    doc.add_page_break()
    p = doc.add_heading("目　录", level=1)
    p.paragraph_format.space_before = Pt(0)
    add_toc(doc)

    for ch in CHAPTERS:
        heading_chapter(doc, ch["title"], first=False)
        n_ch = sum(len(s["questions"]) for s in ch["sections"])
        mp = doc.add_paragraph()
        mp.paragraph_format.space_after = Pt(2)
        meta = ch.get("meta", "").rstrip("。")
        r = mp.add_run("（%s　·　已收录 %d 题）" % (meta, n_ch))
        set_run_font(r, ea=SONG, size=9, color=GRAY)

        for sec in ch["sections"]:
            heading_section(doc, sec["title"])
            for q in sec["questions"]:
                question_block(doc, q)

    answer_index(doc)

    doc.save(out)
    return out, total


def main(argv):
    if len(argv) < 3:
        print("用法: build_docx.py <题库.json> <输出.docx>", flush=True)
        return 2
    bank_path, out_path = argv[1], argv[2]
    with open(bank_path, encoding="utf-8-sig") as f:
        data = json.load(f)
    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    path, n = build(data, out_path)
    print("saved: %s  questions: %d" % (path, n), flush=True)
    return 0


if __name__ == "__main__":
    import sys

    sys.exit(main(sys.argv))
