# -*- coding: utf-8 -*-
import json, re
from pathlib import Path
from pypdf import PdfReader
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
V2 = D / "TRIZa_AI4Science_15min_v2.pptx"
SRC = D / "TRIZa_AI4Science_15min.pptx"
PDF = D / "TRIZa_AI4Science_15min_v2.pdf"

prs = Presentation(str(V2))
notes = [s.notes_slide.notes_text_frame.text for s in prs.slides]
rd = PdfReader(str(PDF))
wavs = sorted((D / "_edge_v2").glob("s*.wav"))
import wave
with wave.open(str(D / "_edge_v2" / "s02.wav")) as w:
    s02 = round(w.getnframes() / w.getframerate(), 2)

out = {
    "russian_fix": ("Russian acronym" in notes[1]) and ("Latin acronym" not in notes[1]),
    "audio_files": len(wavs),
    "s02_sec": s02,
    "pdf_pages": len(rd.pages),
    "pdf_bytes": PDF.stat().st_size,
    "v2_slides": len(prs.slides),
    "original_unchanged": SRC.stat().st_mtime < (D / "_notes_v2.json").stat().st_mtime,
    "original_bytes": SRC.stat().st_size,
    "pdf_s2_snippet": " ".join((rd.pages[1].extract_text() or "").split())[:60],
}
print(json.dumps(out, ensure_ascii=False))
