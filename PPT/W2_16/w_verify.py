# -*- coding: utf-8 -*-
import json, re, sys, pathlib
from pptx import Presentation
from pptx.util import Pt
ROOT = pathlib.Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\out")
FOOT = "TRIZ 创新构思法实战研讨班  \u00b7  \u7b2c %d \u6b21 Workshop"
files = sorted(ROOT.glob("*.pptx"))
res = {"files": len(files), "pages": [], "notes": [], "bad_font": [], "bad_foot": [],
       "ex": [], "minfont": 99.0}
for f in files:
    prs = Presentation(str(f))
    res["pages"].append(len(prs.slides))
    tot = 0
    for sl in prs.slides:
        tot += len(sl.notes_slide.notes_text_frame.text) if sl.has_notes_slide else 0
    res["notes"].append(tot)
    n = int(re.search(r"\u7b2c(\d+)\u6b21", f.name).group(1))
    for i, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if not sh.has_text_frame:
                continue
            txt = sh.text_frame.text.strip()
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.size is None:
                        continue
                    pt = r.font.size.pt
                    if pt >= 20:
                        res["minfont"] = min(res["minfont"], pt)
                        continue
                    ok = (txt.startswith("TRIZ \u521b\u65b0\u6784\u601d\u6cd5") or
                          re.match(r"^\d{2} / 12$", txt) or
                          re.match(r"^\d{2}\s+\u00b7", txt) or
                          (i == 1 and pt >= 12))
                    if not ok:
                        res["bad_font"].append([f.name, i, pt, txt[:12]])
        if i >= 2:
            found = any(sh.has_text_frame and sh.text_frame.text.strip() == (FOOT % n)
                        for sh in sl.shapes)
            if not found:
                res["bad_foot"].append([f.name, i])
    if len(prs.slides) >= 11:
        last = prs.slides[10]
        txt = " ".join(sh.text_frame.text for sh in last.shapes if sh.has_text_frame)
        res["ex"].append("TRIZa" in txt and "\u63d0\u4ea4\u8981\u6c42" in txt)
out = {
 "deck_ok": len(files) == 15 and all(x == 12 for x in res["pages"]),
 "notes_ok": len(res["notes"]) == 15 and all(3000 <= x <= 3500 for x in res["notes"]),
 "font_ok": len(res["bad_font"]) == 0,
 "exercise_ok": len(res["ex"]) == 15 and all(res["ex"]),
 "footer_ok": len(res["bad_foot"]) == 0,
}
print(json.dumps(out, ensure_ascii=True))
print(json.dumps({"notes_min": min(res["notes"]), "notes_max": max(res["notes"]),
                  "bad_font": res["bad_font"][:3], "bad_foot": res["bad_foot"][:3]}))