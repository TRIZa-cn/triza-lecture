# -*- coding: utf-8 -*-
import re
from pptx import Presentation
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2B_盒子与矛盾论_15min.pptx"
prs = Presentation(p)
print("B deck slides:", len(prs.slides))
for i, s in enumerate(prs.slides, 1):
    texts = []
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            texts.append(sh.text_frame.text.strip().replace("\n", " / "))
        elif sh.has_table:
            texts.append("[TABLE]")
    head = " || ".join(texts[:3])[:160]
    print(f"--- s{i}: {head}")
print()
print("notes B:")
for i, s in enumerate(prs.slides, 1):
    n = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
    print(f"[{i}] {len(n)}字 {n[:90]}")