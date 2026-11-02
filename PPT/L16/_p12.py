# -*- coding: utf-8 -*-
from pptx import Presentation
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
prs = Presentation(P)
s = prs.slides[11]
print("== 第 12 页 ==  slide %.2f x %.2f in" % (prs.slide_width/914400, prs.slide_height/914400))
for sh in s.shapes:
    print("  [%s] %s L%.2f T%.2f W%.2f H%.2f" % (sh.shape_type, sh.name,
          sh.left/914400, sh.top/914400, sh.width/914400, sh.height/914400))
    if sh.has_text_frame:
        for line in sh.text_frame.text.strip().splitlines():
            print("       · " + line)
    if sh.has_table:
        for r in sh.table.rows:
            print("       | " + " | ".join(c.text for c in r.cells))
print("--- NOTES ---")
print(s.notes_slide.notes_text_frame.text)