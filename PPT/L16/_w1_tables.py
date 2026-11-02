# -*- coding: utf-8 -*-
from pptx import Presentation
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
prs = Presentation(P)
for i, s in enumerate(prs.slides, 1):
    for sh in s.shapes:
        if sh.has_table:
            t = sh.table
            print(f"\n=== 第 {i} 页 表格 {sh.name}  {len(t.rows)}行 x {len(t.columns)}列  W{sh.width/914400:.2f} H{sh.height/914400:.2f}")
            for r_i, row in enumerate(t.rows):
                cells = [c.text.replace("\n", "⏎") for c in row.cells]
                h = row.height/914400 if row.height else None
                print(f"  r{r_i} h={h and round(h,2)} | " + " ｜ ".join(cells))