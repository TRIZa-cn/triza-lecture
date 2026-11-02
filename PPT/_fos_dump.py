# -*- coding: utf-8 -*-
import json, sys
from pathlib import Path
from pptx import Presentation

src = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\FOS_科普_15min_优化版+闰秒案例.pptx")
print("exists", src.exists(), "size", src.stat().st_size if src.exists() else 0)
prs = Presentation(str(src))
notes = []
for i, s in enumerate(prs.slides, 1):
    t = ""
    if s.has_notes_slide:
        t = s.notes_slide.notes_text_frame.text.strip()
    notes.append(t)
chars = [len(n) for n in notes]
print("slides", len(notes))
print("empty", sum(1 for n in notes if not n))
print("chars_total", sum(chars))
print("chars", chars)
out = {"file": str(src), "slides": len(notes), "notes": notes,
       "chars": sum(chars)}
Path(r"D:\doc\ZGCA\Lectures\TRIZ\_fos_notes.json").write_text(
    json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
for i, n in enumerate(notes, 1):
    print("s%02d (%d): %s" % (i, len(n), n[:60].replace("\n", " ")))
