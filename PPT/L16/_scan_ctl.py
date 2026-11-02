# -*- coding: utf-8 -*-
import re, os, sys
from pptx import Presentation
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
for f in ["TRIZ_W2A_矛盾解剖_15min_优化版.pptx", "TRIZ_W2B_盒子与矛盾论_15min.pptx"]:
    prs = Presentation(os.path.join(P, f))
    for i, s in enumerate(prs.slides, 1):
        n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
        bad = [hex(ord(c)) for c in n if ord(c) < 32 and c not in "\n\t"]
        if bad:
            print(f"{f} 第{i}页 含控制字符: {set(bad)}")
print("扫描完成")