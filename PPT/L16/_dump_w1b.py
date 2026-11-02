# -*- coding: utf-8 -*-
from pptx import Presentation
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
prs = Presentation(P)
print("slides:", len(prs.slides))
for i, s in enumerate(prs.slides, 1):
    print("\n" + "=" * 76)
    print("== 第 %d 页 ==" % i)
    for sh in s.shapes:
        if sh.has_table:
            t = sh.table
            print("  [TABLE] %dx%d  L%.2f T%.2f W%.2f" % (len(t.rows), len(t.columns), sh.left/914400, sh.top/914400, sh.width/914400))
            for r in t.rows:
                print("     | " + " | ".join(c.text.replace("\n","⏎")[:36] for c in r.cells))
            continue
        if not sh.has_text_frame: continue
        txt = sh.text_frame.text.strip()
        if not txt: continue
        sizes = set()
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size: sizes.add(r.font.size.pt)
        print("  [%s] %s L%.2f T%.2f W%.2f H%.2f 字号%s" % (
            sh.name, sh.shape_type, sh.left/914400, sh.top/914400, sh.width/914400, sh.height/914400, sorted(sizes)))
        for line in txt.splitlines():
            print("      · " + line)
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    print("  --- NOTES(%d) ---" % len(n))
    print(n)