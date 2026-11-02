# -*- coding: utf-8 -*-
import json, sys
from pathlib import Path
from pptx import Presentation
src = sys.argv[1]; out = Path(sys.argv[2])
prs = Presentation(src)
notes = [s.notes_slide.notes_text_frame.text.strip() for s in prs.slides]
data = {"file": src, "slides": len(notes), "notes": notes,
        "words": sum(len(n.split()) for n in notes)}
out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
print("SLIDES", len(notes), "WORDS", data["words"], "EMPTY", sum(1 for n in notes if not n))
