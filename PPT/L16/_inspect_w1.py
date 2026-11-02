# -*- coding: utf-8 -*-
import os
from pptx import Presentation
from pptx.util import Pt
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
print("exists:", os.path.isfile(P), "| size:", round(os.path.getsize(P)/1024,1) if os.path.isfile(P) else "-", "KB")
prs = Presentation(P)
print("slides:", len(prs.slides), "| size:", round(prs.slide_width/914400,2), "x", round(prs.slide_height/914400,2))
for i, s in enumerate(prs.slides, 1):
    print(f"\n===== 第 {i} 页 =====")
    for sh in s.shapes:
        info = f"  [{sh.shape_type}] {sh.name} L{sh.left/914400:.2f} T{sh.top/914400:.2f} W{sh.width/914400:.2f} H{sh.height/914400:.2f}"
        print(info)
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                txt = "".join(r.text for r in p.runs)
                if not txt.strip(): continue
                sizes = sorted({(r.font.size.pt if r.font.size else None) for r in p.runs}, key=lambda x: (x is None, x))
                print(f"      {sizes} | {txt[:80]}")
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    print(f"  NOTES({len(n)}): {n[:150]}{'...' if len(n)>150 else ''}")