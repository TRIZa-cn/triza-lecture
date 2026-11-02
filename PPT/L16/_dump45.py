# -*- coding: utf-8 -*-
from pptx import Presentation
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(p)
for idx in (4, 5):
    s = prs.slides[idx-1]
    print("=" * 16, "SLIDE", idx)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            print(f"-- {sh.name} T{round(sh.top/914400,2)} H{round(sh.height/914400,2)} W{round(sh.width/914400,2)}")
            for pi, para in enumerate(sh.text_frame.paragraphs):
                sz = sorted({r.font.size.pt for r in para.runs if r.font.size})
                b = [r.font.bold for r in para.runs]
                print(f"   p{pi} {sz} b={b} :: {para.text}")
        if sh.has_table:
            t = sh.table
            print(f"-- TABLE {sh.name} T{round(sh.top/914400,2)} rows={len(t.rows)} cols={[round(c.width/914400,2) for c in t.columns]}")
            for r in t.rows:
                print("   |", " | ".join(c.text_frame.text for c in r.cells))
    print("NOTES:", s.notes_slide.notes_text_frame.text)