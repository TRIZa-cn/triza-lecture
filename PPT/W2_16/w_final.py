# -*- coding: utf-8 -*-
import json, re, hashlib, pathlib, datetime, subprocess, sys
from pptx import Presentation
base = pathlib.Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
out = base / "W2_16" / "out"
FOOT = "TRIZ \u521b\u65b0\u6784\u601d\u6cd5\u5b9e\u6218\u7814\u8ba8\u73ed  \u00b7  \u7b2c %d \u6b21 Workshop"
files = sorted(out.glob("*.pptx"))
pages, notes, bad_font, bad_foot, ex = [], [], [], [], []
for f in files:
    prs = Presentation(str(f))
    pages.append(len(prs.slides))
    notes.append(sum(len(s.notes_slide.notes_text_frame.text) if s.has_notes_slide else 0 for s in prs.slides))
    n = int(re.search(r"\u7b2c(\d+)\u6b21", f.name).group(1))
    for i, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text.strip()
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.size is None or r.font.size.pt >= 20:
                        continue
                    ok = (t.startswith("TRIZ \u521b\u65b0\u6784\u601d\u6cd5") or
                          re.match(r"^\d{2} / 12$", t) or re.match(r"^\d{2}\s+\u00b7", t) or
                          (i == 1 and r.font.size.pt >= 12))
                    if not ok:
                        bad_font.append([f.name, i, r.font.size.pt])
        if i >= 2 and not any(sh.has_text_frame and sh.text_frame.text.strip() == (FOOT % n) for sh in sl.shapes):
            bad_foot.append([f.name, i])
    last = prs.slides[10]
    txt = " ".join(sh.text_frame.text for sh in last.shapes if sh.has_text_frame)
    ex.append("TRIZa" in txt and "\u63d0\u4ea4\u8981\u6c42" in txt)
a = base / "TRIZ_W1_\u77db\u76fe15m.pptx"
b = base / "TRIZ_W1_\u77db\u76fe15m_\u6f14\u64ad\u7248.pptx"
cut = datetime.datetime(2026, 10, 31)
unchanged = (a.stat().st_size == 704606 and b.stat().st_size == 16890575
             and datetime.datetime.fromtimestamp(a.stat().st_mtime) < cut
             and datetime.datetime.fromtimestamp(b.stat().st_mtime) < cut)
res = {
 "deck_ok": len(files) == 15 and all(p == 12 for p in pages),
 "notes_ok": len(notes) == 15 and all(3000 <= x <= 3500 for x in notes),
 "font_ok": len(bad_font) == 0,
 "exercise_ok": len(ex) == 15 and all(ex),
 "existing_unchanged": unchanged,
 "footer_ok": len(bad_foot) == 0,
 "detail": {"decks": len(files), "pages": sorted(set(pages)), "notes_min": min(notes),
            "notes_max": max(notes), "bad_font": len(bad_font), "bad_foot": len(bad_foot)},
}
print(json.dumps(res, ensure_ascii=True))