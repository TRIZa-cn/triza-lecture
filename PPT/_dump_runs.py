# -*- coding: utf-8 -*-
from pptx import Presentation
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(p)
def runs(sh, label):
    print("##", label)
    for pi, para in enumerate(sh.text_frame.paragraphs):
        for ri, r in enumerate(para.runs):
            c = r.font.color.rgb if (r.font.color and r.font.color.type is not None) else None
            print(f"  p{pi} r{ri} sz={r.font.size.pt if r.font.size else None} b={r.font.bold} col={c} :: {r.text}")
s6 = prs.slides[5]; s7 = prs.slides[6]; s10 = prs.slides[9]
for sh in s6.shapes:
    if sh.has_text_frame and sh.name in ("TextBox 4",): runs(sh, "s6 TextBox 4")
for sh in s7.shapes:
    if sh.has_text_frame and sh.name in ("TextBox 4",): runs(sh, "s7 TextBox 4")
for sh in s10.shapes:
    if sh.has_text_frame and sh.name == "Rectangle 5": runs(sh, "s10 Rectangle 5")
    if sh.has_table:
        print("## s10 table cell runs")
        for ri, row in enumerate(sh.table.rows):
            for ci, c in enumerate(row.cells):
                nr = [len(pp.runs) for pp in c.text_frame.paragraphs]
                print(f"  r{ri}c{ci} runs/para={nr} :: {c.text_frame.text}")
print("notes s7 frame paras:", len(s7.notes_slide.notes_text_frame.paragraphs), "runs:", [len(pp.runs) for pp in s7.notes_slide.notes_text_frame.paragraphs])
print("notes s10 frame paras:", len(s10.notes_slide.notes_text_frame.paragraphs), "runs:", [len(pp.runs) for pp in s10.notes_slide.notes_text_frame.paragraphs])