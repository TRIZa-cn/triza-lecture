# -*- coding: utf-8 -*-
"""只重建《16讲讲稿集》，不动录音稿（录音稿的段号要与他已录的 16 个文件保持对应）。"""
import sys, os
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16")
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from pptx import Presentation
import _l16_spec_a as A, _l16_spec_a2 as A2, _l16_spec_b as Bs, _l16_spec_c as C

P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
A_N = [s.notes_slide.notes_text_frame.text.strip() for s in Presentation(os.path.join(P, "TRIZ_W2A_矛盾解剖_15min_优化版.pptx")).slides]
B_N = [s.notes_slide.notes_text_frame.text.strip() for s in Presentation(os.path.join(P, "TRIZ_W2B_盒子与矛盾论_15min.pptx")).slides]
SPECS = {s["no"]: s for s in A.SPECS + A2.SPECS + Bs.SPECS + C.SPECS}
RED = RGBColor(0xB0, 0x3A, 0x2B); GREY = RGBColor(0x55, 0x5F, 0x6B)

g = Document(); gs = g.sections[0]
gs.page_height, gs.page_width = Cm(29.7), Cm(21.0)
gs.left_margin = gs.right_margin = Cm(2.4); gs.top_margin = gs.bottom_margin = Cm(2.2)

def setf(r, size, bold=False, color=None):
    r.font.size = Pt(size); r.bold = bold; r.font.name = "微软雅黑"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    if color: r.font.color.rgb = color

def _sanitize(x):
    return ''.join(c if (ord(c) >= 32 or c in '\n\t') else ' ' for c in x)


def gp(t, size=12, bold=False, before=0, after=4, color=None, spacing=1.45):
    t = _sanitize(t)
    p = g.add_paragraph()
    p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    setf(p.add_run(t), size, bold, color)

h = g.add_paragraph(); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
setf(h.add_run("TRIZ 创新构思法实战研讨班 · 16 讲讲稿集（v1.1）"), 20, True)
gp("每讲八页骨架：封面、困境、定义、流程、例子、表、边界、收束。正文即口播稿。"
   "本版对应 A 讲 13 页（新增第 5 页《同一个问题，三种写法》）。", 10.5, color=GREY, after=10, spacing=1.35)

def sec(title, pages):
    gp(title, 15, True, before=14, after=6, color=RED)
    tot = sum(len(x) for x in pages)
    gp(f"共 {len(pages)} 页 · {tot} 字 · 约 {tot/270:.1f} 分钟", 9.5, color=GREY, after=6, spacing=1.2)
    for i, (hd, bd) in enumerate(pages, 1):
        gp(f"{i:02d}　{hd}", 11, True, before=6, after=2)
        gp(bd, 11.5, after=3)

gp("讲 1　矛盾解剖与约束观（A 段 13 页 + B 段 13 页）", 15, True, before=8, after=6, color=RED)
gp("A 段", 12, True, before=6, after=4)
for i, n in enumerate(A_N, 1):
    gp(f"A{i:02d}　第 {i} 页", 11, True, before=4, after=2); gp(n, 11.5, after=3)
gp("B 段", 12, True, before=8, after=4)
for i, n in enumerate(B_N, 1):
    gp(f"B{i:02d}　第 {i} 页", 11, True, before=4, after=2); gp(n, 11.5, after=3)
for no in range(2, 17):
    sp = SPECS[no]
    sec(f"讲 {no:02d}　{sp['file']}", [(p["title"], p.get("notes", "")) for p in sp["pages"]])

OUT = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-16讲讲稿集-v1.docx"
g.save(OUT)
print("讲稿集 ok:", round(os.path.getsize(OUT)/1024, 1), "KB | A 段", len(A_N), "页 | B 段", len(B_N), "页")