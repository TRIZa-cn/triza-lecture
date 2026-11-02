# -*- coding: utf-8 -*-
from pathlib import Path
from pptx import Presentation
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
prs = Presentation(str(D / "TRIZa_V8_AI4Science_15min.pptx"))
for i, s in enumerate(prs.slides, 1):
    kinds = []
    for sh in s.shapes:
        if sh.shape_type == 13:
            kinds.append("PIC %.2fx%.2f" % (sh.width / 914400, sh.height / 914400))
    if kinds:
        print("S%d %s" % (i, " ".join(kinds)))
for idx in (0, 2, 5, 7, 9):
    s = prs.slides[idx]
    txt = [sh.text_frame.text.replace("\n", " / ") for sh in s.shapes
           if sh.has_text_frame and sh.text_frame.text.strip()]
    print("--- S%d: %s" % (idx + 1, " | ".join(txt)[:420]))
print("NOTES0:", prs.slides[0].notes_slide.notes_text_frame.text[:120])
