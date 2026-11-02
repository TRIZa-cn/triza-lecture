import sys
from pptx import Presentation
from pptx.util import Emu
p = r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_20min_20261022\FOS_科普_20min_20261022\FOS_科普_15min.pptx"
prs = Presentation(p)
print("SIZE", prs.slide_width, prs.slide_height, prs.slide_width/914400, prs.slide_height/914400)
print("LAYOUTS", [l.name for l in prs.slide_masters[0].slide_layouts])
print("N", len(prs.slides))
for i, s in enumerate(prs.slides, 1):
    print("="*70)
    print("SLIDE", i, "layout:", s.slide_layout.name)
    for sh in s.shapes:
        t = ""
        if sh.has_text_frame:
            t = sh.text_frame.text.replace("\n"," | ")
        print(f"  [{sh.shape_type}] name={sh.name!r} pos=({sh.left},{sh.top}) size=({sh.width},{sh.height}) txt={t[:300]!r}")
        if sh.has_table:
            for r in sh.table.rows:
                print("     ROW:", [c.text for c in r.cells])
    if s.has_notes_slide:
        nt = s.notes_slide.notes_text_frame.text
        print("  NOTES:", nt.replace("\n"," / ")[:2000])
