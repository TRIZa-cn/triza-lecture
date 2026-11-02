# -*- coding: utf-8 -*-
import json, re
from pathlib import Path
from pptx import Presentation

P = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZa_AI4Science_15min.pptx")
prs = Presentation(str(P))
rows = []
for i, s in enumerate(prs.slides, 1):
    t = ""
    if s.has_notes_slide:
        t = s.notes_slide.notes_text_frame.text.strip()
    rows.append(t)
words = sum(len(re.findall(r"[A-Za-z][A-Za-z'-]*", t)) for t in rows)
print("slides", len(rows), "words", words)
out = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_notes_new.json")
out.write_text(json.dumps({"file": str(P), "slides": len(rows), "words": words, "notes": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
for i, t in enumerate(rows, 1):
    print("=== %02d (%d chars) ===" % (i, len(t)))
    print(t)
