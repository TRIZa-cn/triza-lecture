# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Emu
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(p)
def inch(v): return round((v or 0)/914400.0, 2)
for idx in (5, 6, 7, 10):
    s = prs.slides[idx-1]
    print("=" * 20, "SLIDE", idx, "layout:", s.slide_layout.name)
    for sh in s.shapes:
        kind = "TABLE" if sh.has_table else ("TEXT" if sh.has_text_frame else sh.shape_type)
        print(f"-- [{sh.shape_id}] {sh.name} {kind} L{inch(sh.left)} T{inch(sh.top)} W{inch(sh.width)} H{inch(sh.height)}")
        if sh.has_text_frame:
            for pi, para in enumerate(sh.text_frame.paragraphs):
                sizes = sorted({r.font.size.pt for r in para.runs if r.font.size})
                bold = any(r.font.bold for r in para.runs)
                print(f"     p{pi} sz={sizes} b={bold} :: {para.text}")
        if sh.has_table:
            t = sh.table
            print("     cols:", [inch(c.width) for c in t.columns], "rows h:", [inch(r.height) for r in t.rows])
            for ri, row in enumerate(t.rows):
                cells = []
                for c in row.cells:
                    szs = sorted({r.font.size.pt for pp in c.text_frame.paragraphs for r in pp.runs if r.font.size})
                    cells.append(f"{c.text_frame.text}{szs}")
                print(f"     r{ri}: " + " | ".join(cells))
print()
print("=" * 20, "NOTES (all)")
for i, s in enumerate(prs.slides, 1):
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    print(f"--- notes {i} ({len(n)}字)\n{n}")