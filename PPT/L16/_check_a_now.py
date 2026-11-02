# -*- coding: utf-8 -*-
import hashlib, os
from pptx import Presentation
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(P)
h = hashlib.sha256(open(P, "rb").read()).hexdigest()[:16]
print("页数:", len(prs.slides), "| 大小:", round(os.path.getsize(P)/1024,1), "KB | sha256:", h)
print("notes 总字数:", sum(len(s.notes_slide.notes_text_frame.text) for s in prs.slides))
titles = []
for s in prs.slides:
    t = [sh.text_frame.text.strip() for sh in s.shapes if sh.has_text_frame and sh.name == "TextBox 2"]
    titles.append(t[0] if t else "(封面)")
for i, t in enumerate(titles, 1):
    print(f"  {i:2d} {t}")