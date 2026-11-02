# -*- coding: utf-8 -*-
import json, sys
from pathlib import Path
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
P = D / "TRIZa_V8_AI4Science_15min.pptx"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"
SRC = [r"D:\doc\ZGCLab\triza\paper\triza\TRIZa_V8_8页正式稿改动与发布记录_20261012.md",
       r"D:\doc\ZGCLab\triza\paper\triza\TRIZa_Supplementary_Appendix_20261012.pdf",
       r"D:\doc\ZGCLab\triza\paper\triza\TRIZa_AI4Science_ACM_Overleaf_V8_20261012.zip",
       r"D:\doc\ZGCA\Lectures\TRIZ\_triza_v8_src\main.tex",
       r"D:\doc\ZGCA\Lectures\TRIZ\_triza_v8_src\figures\triza_edit_framework.png"]

try:
    prs = Presentation(str(P))
except Exception as e:
    print(json.dumps({"pptx_ok": False, "err": str(e)}))
    sys.exit(0)

SW, SH = prs.slide_width, prs.slide_height
minf, viol, oob, empty_slides, mathn, v8hits = 999.0, [], 0, [], 0, 0
notes_words, empty_notes = 0, 0
KEY = ["2.87", "4.23", "0.15", "0.01584", "3.2277", "2,040"]
corpus = []

for i, s in enumerate(prs.slides, 1):
    texts = 0
    for sh in s.shapes:
        mathn += len(sh._element.findall(".//" + M))
        try:
            if (sh.left < -10000 or sh.top < -10000
                    or sh.left + sh.width > SW + 10000 or sh.top + sh.height > SH + 10000):
                oob += 1
        except TypeError:
            pass
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text
        if t.strip():
            texts += 1
            corpus.append(t)
            if "V8" in t:
                v8hits += 1
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                sz = r.font.size.pt if r.font.size else None
                if sz is not None:
                    if sz < minf:
                        minf = sz
                    if sz < 20:
                        viol.append([i, r.text[:18], sz])
    if texts == 0:
        empty_slides.append(i)
    note = s.notes_slide.notes_text_frame.text.strip()
    notes_words += len(note.split())
    if len(note.split()) < 20:
        empty_notes += 1
    corpus.append(note)

blob = "\n".join(corpus)
found = [k for k in KEY if k in blob]
missing = [k for k in KEY if k not in blob]
src_mtime = max(Path(f).stat().st_mtime for f in SRC)

out = {
    "pptx_ok": True,
    "slides": len(prs.slides),
    "min_font": minf,
    "font_viol_count": len(viol),
    "font_viol": viol[:6],
    "oob_shapes": oob,
    "empty_slides": empty_slides,
    "math_objects": mathn,
    "v8_hits": v8hits,
    "empty_notes": empty_notes,
    "notes_words": notes_words,
    "notes_min_at135wpm": round(notes_words / 135.0, 1),
    "reminder_hits": sum(blob.lower().count(w) for w in
                         ["speaker note", "say this", "timing:", "pause here",
                          "read aloud", "cue:", "backup slide"]),
    "key_numbers_found": len(found),
    "key_numbers_missing": missing,
    "source_unchanged": src_mtime < (D / "build_v8b.py").stat().st_mtime,
    "pptx_bytes": P.stat().st_size,
}
print(json.dumps(out, ensure_ascii=False))
