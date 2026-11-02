# -*- coding: utf-8 -*-
"""第 2-16 次 Workshop 共用建库工具。

视觉体系沿用 build_w2.py（第 2 次课的原始设计），但按 TRIZa-Lecture-SKILL 修订：
  - 正文 >=20pt；只有 kicker 14pt、页脚 12pt、页码 12pt、封面署名 12pt 可更小
  - 页脚统一为「TRIZ 创新构思法实战研讨班  ·  第 N 次 Workshop」
  - 一页只回答一个问题；标题写成一句话；每页 <=6 个要点
"""
import math

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from w_fig import draw_fig, caption
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml import parse_xml

FONT = "微软雅黑"
INK = RGBColor(0x1A, 0x1D, 0x23)
INK2 = RGBColor(0x56, 0x5C, 0x66)
INK3 = RGBColor(0x8A, 0x90, 0x9A)
PAPER = RGBColor(0xFA, 0xF9, 0xF6)
RED = RGBColor(0xB0, 0x3A, 0x2B)
BLUE = RGBColor(0x2C, 0x5F, 0x8A)
AMBER = RGBColor(0xA9, 0x6E, 0x24)
GREEN = RGBColor(0x35, 0x69, 0x4A)
LINE = RGBColor(0xDD, 0xD8, 0xCE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
Z0 = RGBColor(0xF2, 0xEF, 0xE8)
Z1 = RGBColor(0xFC, 0xFB, 0xF8)

ACCENTS = {"red": RED, "blue": BLUE, "amber": AMBER, "green": GREEN}

SW, SH = 13.333, 7.5
ML, MR = 0.85, 0.85
CW = SW - ML - MR
FOOT = "TRIZ 创新构思法实战研讨班  ·  第 %d 次 Workshop"

SZ_KICKER = 14
SZ_TITLE = 30
SZ_BODY = 21
SZ_BODY_S = 20
SZ_FOOT = 12
SZ_PAGE = 12
SZ_COVER_META = 12
SZ_BOXLABEL = 20

TOTAL_PAGES = 12


def _set_ea(run, name):
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = parse_xml('<%s %s typeface="%s"/>' % (tag, nsdecls("a"), name))
            rPr.append(el)
        else:
            el.set("typeface", name)


def rich(p, text, size, color=INK, bold=False, accent=RED, spc=None):
    for i, seg in enumerate(text.split("**")):
        if seg == "":
            continue
        r = p.add_run()
        r.text = seg
        r.font.name = FONT
        r.font.size = Pt(size)
        if i % 2 == 1:
            r.font.bold = True
            r.font.color.rgb = accent
        else:
            r.font.bold = bold
            r.font.color.rgb = color
        if spc:
            r._r.get_or_add_rPr().set("spc", str(spc))
        _set_ea(r, FONT)
    return p


def textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    return tf


def para(tf, text, size, color=INK, bold=False, accent=RED, space_before=0,
         space_after=6, align=PP_ALIGN.LEFT, first=False, line=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    if line:
        p.line_spacing = line
    rich(p, text, size, color, bold, accent)
    return p


def bullet(tf, text, size=SZ_BODY, color=INK, accent=RED, marker="\u25b8", first=False,
           space_after=14, line=1.26, indent=0.32):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_after = Pt(space_after)
    p.line_spacing = line
    p.left_indent = Inches(indent)
    p.first_line_indent = Inches(-indent)
    r = p.add_run()
    r.text = marker + "  "
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.color.rgb = accent
    r.font.bold = True
    _set_ea(r, FONT)
    rich(p, text, size, color, accent=accent)
    return p


def rect(slide, left, top, width, height, fill, line_col=None,
         shape=MSO_SHAPE.RECTANGLE, lw=1.0):
    s = slide.shapes.add_shape(shape, Inches(left), Inches(top), Inches(width), Inches(height))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line_col is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line_col
        s.line.width = Pt(lw)
    tf = s.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.20)
    tf.margin_right = Inches(0.20)
    tf.margin_top = Inches(0.12)
    tf.margin_bottom = Inches(0.12)
    return s


def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = PAPER
    return s


def page(prs, kicker, title, accent=RED, notes=""):
    s = blank(prs)
    tf = textbox(s, ML, 0.50, CW, 0.30)
    rich(tf.paragraphs[0], kicker, SZ_KICKER, INK3, bold=True, accent=accent, spc=240)
    tf = textbox(s, ML, 0.86, CW, 0.86)
    para(tf, title, SZ_TITLE, INK, bold=True, accent=accent, first=True,
         space_after=0, line=1.06)
    rect(s, ML, 1.82, 1.05, 0.045, accent)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def foot(s, idx, n, accent=RED):
    rect(s, ML, 6.88, CW, 0.008, LINE)
    tf = textbox(s, ML, 6.98, 9.4, 0.28)
    rich(tf.paragraphs[0], FOOT % n, SZ_FOOT, INK3, accent=accent, spc=100)
    tf2 = textbox(s, SW - MR - 0.9, 6.98, 0.9, 0.28)
    para(tf2, "%02d / %02d" % (idx, TOTAL_PAGES), SZ_PAGE, INK3, first=True,
         align=PP_ALIGN.RIGHT)


def band(s, text, top, accent=RED, size=SZ_BODY_S, height=0.66):
    r = rect(s, ML + 0.14, top, CW - 0.14, height, None)
    r.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    r.text_frame.margin_left = Inches(0.10)
    rich(r.text_frame.paragraphs[0], text, size, INK, accent=accent)
    rect(s, ML, top, 0.055, height, accent)
    rect(s, ML + 0.14, top, CW - 0.14, 0.007, LINE)
    rect(s, ML + 0.14, top + height, CW - 0.14, 0.007, LINE)


def box(s, left, top, width, height, label, lines, accent=RED,
        size=SZ_BODY_S, line=1.16, gap=5):
    b = rect(s, left, top, width, height, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.26)
    b.text_frame.margin_top = Inches(0.22)
    tf = b.text_frame
    rich(tf.paragraphs[0], label, SZ_BOXLABEL, INK3, bold=True, accent=accent, spc=120)
    for t in lines:
        para(tf, t, size, INK, accent=accent, space_before=gap, line=line)
    return b


def table(s, rows, top, col_w, size=SZ_BODY_S, row_h=0.78, head_h=0.52,
          accent=RED, head_fill=INK, left=ML):
    nrow, ncol = len(rows), len(rows[0])
    width = sum(col_w)
    shp = s.shapes.add_table(nrow, ncol, Inches(left), Inches(top),
                             Inches(width), Inches(head_h + row_h * (nrow - 1)))
    tbl = shp.table
    tbl.first_row = True
    tbl.horz_banding = False
    tbl.last_row = False
    tbl.first_col = False
    tbl.last_col = False
    for i, w in enumerate(col_w):
        tbl.columns[i].width = Inches(w)
    tbl.rows[0].height = Inches(head_h)
    for i in range(1, nrow):
        tbl.rows[i].height = Inches(row_h)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = tbl.cell(ri, ci)
            c.margin_left = Inches(0.12)
            c.margin_right = Inches(0.08)
            c.margin_top = Inches(0.04)
            c.margin_bottom = Inches(0.04)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            c.fill.fore_color.rgb = head_fill if ri == 0 else (Z0 if (ri - 1) % 2 == 0 else Z1)
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.line_spacing = 1.04
            if ri == 0:
                rich(p, val, size, WHITE, bold=True, accent=accent)
            else:
                rich(p, val, size, INK, bold=(ci == 0), accent=accent)
    return tbl


def cover(prs, kicker, title, subtitle, meta, accent=RED, notes=""):
    s = blank(prs)
    rect(s, 0, 0, 0.16, SH, accent)
    tf = textbox(s, 1.15, 1.58, 10.4, 0.30)
    rich(tf.paragraphs[0], kicker, SZ_KICKER, INK3, bold=True, accent=accent, spc=280)
    tf = textbox(s, 1.15, 2.02, 10.9, 1.62)
    para(tf, title, 42, INK, bold=True, accent=accent, first=True, line=1.10, space_after=0)
    rect(s, 1.18, 3.98, 1.5, 0.05, accent)
    tf = textbox(s, 1.15, 4.28, 10.4, 1.10)
    para(tf, subtitle, SZ_BODY, INK2, first=True, line=1.45, space_after=0, accent=accent)
    tf = textbox(s, 1.15, 6.10, 10.4, 0.85)
    para(tf, meta, SZ_COVER_META, INK3, first=True, line=1.5, space_after=0)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def build_session(sess):
    """按固定 12 页骨架生成一份课件。sess 定义见 w_content_*.py。"""
    n = sess["n"]
    acc = ACCENTS[sess["accent"]]
    nt = sess["notes"]
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    idx = 0

    def nxt():
        nonlocal idx
        idx += 1
        return idx

    i = nxt()
    cover(prs, "TRIZ 创新构思法实战研讨班 · 第 %d 次 WORKSHOP" % n,
          sess["cover_title"], sess["cover_sub"], sess["cover_meta"], acc,
          nt.get("cover", ""))

    i = nxt()
    d = sess["dilemma"]
    s = page(prs, "00  ·  先说困境", d["title"], acc, nt.get("dilemma", ""))
    tf = textbox(s, ML, 2.14, CW, 3.2)
    for k, t in enumerate(d["items"]):
        bullet(tf, t, SZ_BODY, INK, acc, first=(k == 0), space_after=20)
    band(s, d["band"], 5.40, acc, SZ_BODY_S, 0.68)
    foot(s, i, n, acc)

    i = nxt()
    dfn = sess["define"]
    s = page(prs, "01  ·  定义", dfn["title"], acc, nt.get("define", ""))
    tf = textbox(s, ML, 2.12, 6.25, 3.9)
    for k, t in enumerate(dfn["points"]):
        bullet(tf, t, SZ_BODY_S, INK, acc, first=(k == 0), space_after=16)
    box(s, 7.35, 2.10, 5.0, 4.15, dfn["box_title"], dfn["box"], acc, SZ_BODY_S, line=1.16, gap=5)
    foot(s, i, n, acc)

    i = nxt()
    wy = sess["why"]
    s = page(prs, "02  ·  为什么需要它", wy["title"], acc, nt.get("why", ""))
    tf = textbox(s, ML, 2.12, CW, 3.3)
    for k, t in enumerate(wy["items"]):
        bullet(tf, t, SZ_BODY_S, INK, acc, first=(k == 0), space_after=18)
    band(s, wy["band"], 5.50, acc, SZ_BODY_S, 0.66)
    foot(s, i, n, acc)

    i = nxt()
    fl = sess["flow"]
    s = page(prs, "03  ·  怎么做", fl["title"], acc, nt.get("flow", ""))
    tf = textbox(s, ML, 2.08, 7.35, 4.22)
    for k, ht in enumerate(fl["steps"]):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(6)
        p.line_spacing = 1.10
        rich(p, ht[0] + "　", SZ_BODY_S, INK, bold=True, accent=acc)
        rich(p, ht[1], SZ_BODY_S, INK2, accent=acc)
    box(s, 8.55, 2.08, 3.8, 3.95, fl["box_title"], fl["box"], acc, SZ_BODY_S, line=1.16, gap=5)
    foot(s, i, n, acc)

    i = nxt()
    tb = sess["table"]
    s = page(prs, "04  ·  一张表", tb["title"], acc, nt.get("table", ""))
    table(s, tb["rows"], 2.12, tb["col_w"], SZ_BODY_S, row_h=tb.get("row_h", 0.78),
          head_h=tb.get("head_h", 0.52), accent=acc)
    band(s, tb["band"], tb.get("band_top", 5.90), acc, SZ_BODY_S, 0.60)
    foot(s, i, n, acc)

    i = nxt()
    e1 = sess["ex1"]
    s = page(prs, "05  ·  例子一", e1["title"], acc, nt.get("ex1", ""))
    tf = textbox(s, ML, 2.10, 6.1, 4.05)
    for k, t in enumerate(e1["left"]):
        bullet(tf, t, SZ_BODY_S, INK, acc, first=(k == 0), space_after=15)
    box(s, 7.2, 2.10, 5.15, 4.10, e1["box_title"], e1["box"], acc, SZ_BODY_S, line=1.16, gap=5)
    foot(s, i, n, acc)

    i = nxt()
    e2 = sess["ex2"]
    s = page(prs, "06  ·  例子二", e2["title"], acc, nt.get("ex2", ""))
    tf = textbox(s, ML, 2.10, 6.1, 4.05)
    for k, t in enumerate(e2["left"]):
        bullet(tf, t, SZ_BODY_S, INK, acc, first=(k == 0), space_after=15)
    if e2.get("fig"):
        draw_fig(s, e2["fig"], 7.15, 2.02, 5.20, 3.52, acc)
        caption(s, 7.18, 5.62, 5.20, 0.62, e2["fig"]["cap"], SZ_BODY_S, acc)
    else:
        box(s, 7.2, 2.10, 5.15, 4.10, e2["box_title"], e2["box"], acc, SZ_BODY_S, line=1.16, gap=5)
    foot(s, i, n, acc)

    i = nxt()
    wn = sess["warn"]
    s = page(prs, "07  ·  警告", wn["title"], acc, nt.get("warn", ""))
    tf = textbox(s, ML, 2.10, 6.3, 4.00)
    for k, t in enumerate(wn["items"]):
        bullet(tf, t, SZ_BODY_S, INK, acc, first=(k == 0), space_after=18)
    box(s, 7.5, 2.10, 4.85, 4.00, "一个自检问题", wn["selfcheck"], acc, SZ_BODY_S, line=1.16, gap=5)
    foot(s, i, n, acc)

    i = nxt()
    pr = sess["prompt"]
    s = page(prs, "08  ·  TRIZa 实操", pr["title"], acc, nt.get("prompt", ""))
    b = rect(s, ML, 2.10, CW, 3.45, Z1, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.30)
    b.text_frame.margin_top = Inches(0.20)
    tf = b.text_frame
    rich(tf.paragraphs[0], "把这段话直接粘进 TRIZa 发明工作台", SZ_BOXLABEL, INK3, bold=True,
         accent=acc, spc=100)
    for t in pr["lines"]:
        para(tf, t, SZ_BODY_S, INK, accent=acc, space_before=3, line=1.22)
    _h = 0.62 if len(pr["tip"]) <= 38 else 0.95
    band(s, pr["tip"], 6.62 - _h, acc, SZ_BODY_S, _h)
    foot(s, i, n, acc)

    i = nxt()
    tk = sess["task"]
    s = page(prs, "09  ·  课堂练习", tk["title"], acc, nt.get("task", ""))
    tf = textbox(s, ML, 2.08, 7.2, 4.15)
    marks = ["\u2460", "\u2461", "\u2462", "\u2463", "\u2464"]
    for k, t in enumerate(tk["steps"]):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(14)
        p.line_spacing = 1.24
        rich(p, marks[k] + "　", SZ_BODY_S, acc, bold=True, accent=acc)
        rich(p, t, SZ_BODY_S, INK, accent=acc)
    box(s, 8.40, 2.08, 3.95, 4.38, "提交要求 · " + tk["time"], tk["submit"], acc,
        SZ_BODY_S, line=1.5, gap=11)
    foot(s, i, n, acc)

    i = nxt()
    sm = sess["summary"]
    s = page(prs, "10  ·  小结", "三句话，带回去用", acc, nt.get("summary", ""))
    tf = textbox(s, ML, 2.14, CW, 3.3)
    for k, t in enumerate(sm["points"]):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(24)
        p.line_spacing = 1.3
        rich(p, "0%d　" % (k + 1), 20, acc, bold=True, accent=acc, spc=100)
        rich(p, t, SZ_BODY, INK, accent=acc)
    _sh = 0.78 if len(sm["next"]) <= 44 else 0.96
    foot(s, i, n, acc)

    return prs


def audit(path):
    prs = Presentation(str(path))
    n = len(prs.slides._sldIdLst)
    minf, notes, warns, dense = 99.0, [], 0, []
    for si, s in enumerate(prs.slides, 1):
        nt = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
        notes.append(len(nt))
        body_chars = 0
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            w_in = sh.width / 914400.0
            h_in = sh.height / 914400.0
            need = 0.0
            for p in sh.text_frame.paragraphs:
                txt = "".join(r.text for r in p.runs)
                if not txt.strip():
                    continue
                body_chars += len(txt)
                sz = max([r.font.size.pt for r in p.runs if r.font.size] or [20])
                minf = min(minf, sz)
                cpl = max(1.0, (w_in - 0.44) * 72.0 / (sz * 1.02))
                need += math.ceil(len(txt) / cpl) * sz * 1.46 / 72.0
            if need > h_in * 1.18 and h_in > 0.4:
                warns += 1
                print("   ! 可能溢出 p%-2d 需 %.2f 框 %.2f :: %s"
                      % (si, need, h_in,
                         "".join(r.text for r in sh.text_frame.paragraphs[0].runs)[:26]))
        if si > 1 and body_chars > 270:
            dense.append((si, body_chars))
    est = sum(notes) / 235.0
    return dict(slides=n, min_font=minf, notes=sum(notes), est_min=est,
                warns=warns, dense=dense, note_list=notes)