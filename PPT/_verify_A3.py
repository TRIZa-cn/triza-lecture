# -*- coding: utf-8 -*-
import glob, os
from pptx import Presentation
D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
p = os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版.pptx")
prs = Presentation(p)
for idx in (6, 7, 10):
    s = prs.slides[idx-1]
    print("="*18, "SLIDE", idx)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            print(f"-- {sh.name} T{round(sh.top/914400,2)} H{round(sh.height/914400,2)}")
            for para in sh.text_frame.paragraphs:
                sz = sorted({r.font.size.pt for r in para.runs if r.font.size})
                print(f"   {sz} :: {para.text}")
        if sh.has_table:
            t = sh.table
            print("   cols:", [round(c.width/914400,2) for c in t.columns])
            for row in t.rows:
                print("   |", " | ".join(c.text_frame.text for c in row.cells))
print()
print("NOTES 6/7/10 len:", [len(prs.slides[i].notes_slide.notes_text_frame.text) for i in (5,6,9)])
print("total notes:", sum(len(s.notes_slide.notes_text_frame.text) for s in prs.slides))
print()
print("=== footer check across W2 decks ===")
for f in ["TRIZ_W2A_矛盾解剖_15min.pptx","TRIZ_W2A_矛盾解剖_15min_优化版.pptx",
          "TRIZ_W2A_矛盾解剖_15min_优化版_自动解说.pptx","TRIZ_W2A_矛盾解剖_15min_自动解说.pptx",
          "TRIZ_W2B_盒子与矛盾论_15min.pptx","TRIZ_W2B_盒子与矛盾论_15min_自动解说.pptx"]:
    q = os.path.join(D, f)
    pr = Presentation(q)
    txt = []
    for s in pr.slides:
        for sh in s.shapes:
            if sh.has_text_frame and "创新构思法" in sh.text_frame.text:
                txt.append(sh.text_frame.text.strip()); break
    bad = [t for t in set(txt) if "创新构思法研讨班" in t]
    print(f"{f}: samples={len(set(txt))} 残留旧名={len(bad)}  {sorted(set(txt))[:1]}")