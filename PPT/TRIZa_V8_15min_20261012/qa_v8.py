# -*- coding: utf-8 -*-
import json
from pathlib import Path
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
P = D / "TRIZa_V8_AI4Science_15min.pptx"
SRC = [Path(r"D:\doc\ZGCLab\triza\paper\triza\TRIZa_V8_8页正式稿改动与发布记录_20261012.md"),
       Path(r"D:\doc\ZGCLab\triza\paper\triza\TRIZa_Supplementary_Appendix_20261012.pdf"),
       Path(r"D:\doc\ZGCLab\triza\paper\triza\TRIZa_AI4Science_ACM_Overleaf_V8_20261012.zip"),
       Path(r"D:\doc\ZGCA\Lectures\TRIZ\_triza_v8_src\main.tex"),
       Path(r"D:\doc\ZGCA\Lectures\TRIZ\_triza_v8_src\supplement.tex")]

PATS = ["speaker note", "notes:", "say this", "timing", "pause", "read aloud",
        "cue:", "optional", "backup", "slide", "slide:", "ppt", "deck",
        "\u5907\u7528", "\u63d0\u793a", "\u8bb2\u7a3f", "\u5e7b\u706f\u7247", "\u505c\u987f"]

res = {"pptx_ok": False}
try:
    prs = Presentation(str(P))
    res["pptx_ok"] = True
except Exception as e:
    print(json.dumps({"pptx_ok": False, "err": str(e)}))
    raise SystemExit

sl = list(prs.slides)
res["slides"] = len(sl)
minf, viol = 999.0, []
page_text, notes = [], []
oob = 0
for i, s in enumerate(sl, 1):
    for sh in s.shapes:
        try:
            if sh.left is None:
                continue
            if (sh.left < -10000 or sh.top < -10000 or
                    sh.left + sh.width > prs.slide_width + 20000 or
                    sh.top + sh.height > prs.slide_height + 20000):
                oob += 1
        except Exception:
            pass
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                sz = r.font.size.pt if r.font.size is not None else None
                if sz is None:
                    viol.append([i, r.text[:28], None])
                    continue
                minf = min(minf, sz)
                if sz < 20:
                    viol.append([i, r.text[:28], sz])
    page_text.append("\n".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame))
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    notes.append(n)

res["min_font"] = minf
res["font_viol_count"] = len(viol)
res["font_viol"] = viol[:12]
res["oob_shapes"] = oob
res["empty_notes"] = sum(1 for n in notes if not n.strip())
res["notes_words"] = sum(len(n.split()) for n in notes)
res["notes_min_at135wpm"] = round(sum(len(n.split()) for n in notes) / 135.0, 1)
hits = []
for i, n in enumerate(notes, 1):
    low = n.lower()
    for pat in PATS:
        if pat in low:
            hits.append([i, pat])
res["reminder_hits"] = len(hits)
res["reminder_detail"] = hits[:10]
blob = "\n".join(page_text) + "\n" + "\n".join(notes)
keys = ["2.87", "4.23", "0.15", "0.01584", "3.2277", "2,040"]
res["key_numbers_found"] = sum(1 for k in keys if k in blob)
res["key_numbers_missing"] = [k for k in keys if k not in blob]
res["empty_slides"] = sum(1 for t in page_text if not t.strip())
res["pages_short"] = [i + 1 for i, t in enumerate(page_text) if len(t.strip()) < 40]
src_max = max(p.stat().st_mtime for p in SRC if p.exists())
res["source_unchanged"] = src_max < (D / "build_v8.py").stat().st_mtime
res["pptx_bytes"] = P.stat().st_size
print(json.dumps(res, ensure_ascii=False))
