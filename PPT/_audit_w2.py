# -*- coding: utf-8 -*-
import json, datetime
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")


def spc_of(run):
    rPr = run._r.find(qn("a:rPr"))
    return None if rPr is None else rPr.get("spc")


def audit(path):
    prs = Presentation(str(path))
    slides = int(len(prs.slides._sldIdLst))
    body, allmin, tmin, notes = [], 99.0, 99.0, 0
    for idx, s in enumerate(prs.slides):
        notes += len(s.notes_slide.notes_text_frame.text.strip())
        for sh in s.shapes:
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        for p in c.text_frame.paragraphs:
                            for r in p.runs:
                                if r.font.size:
                                    tmin = min(tmin, r.font.size.pt)
                continue
            if not sh.has_text_frame:
                continue
            top = sh.top / 914400.0
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if not r.font.size:
                        continue
                    allmin = min(allmin, r.font.size.pt)
                    if idx == 0 or spc_of(r) or not (0.85 <= top <= 6.90):
                        continue
                    body.append(r.font.size.pt)
    return {"slides": slides, "min_font": round(min(body), 1),
            "min_all": round(allmin, 1), "table_min": round(tmin, 1),
            "notes": notes, "est_min": round(notes / 235.0, 1)}


OLD = {"TRIZa_AI4Science_15min.pptx": (426613, "2026-10-24 07:59"),
       "TRIZa_AI4Science_15min_v2.pptx": (401376, "2026-10-24 09:20"),
       "TRIZa_15min_handoff.zip": (22037056, "2026-10-24 15:57")}

pa = D / "TRIZ_W2A_矛盾解剖_15min.pptx"
pb = D / "TRIZ_W2B_盒子与矛盾论_15min.pptx"
ra, rb = audit(pa), audit(pb)
unchanged, odetail = True, {}
for name, (sz, mt) in OLD.items():
    f = D / name
    st = f.stat()
    ok = st.st_size == sz and datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M") == mt
    odetail[name] = "ok" if ok else "CHANGED"
    unchanged = unchanged and ok
print(json.dumps({
    "decks": int(pa.is_file() and pb.is_file()) + int(bool(ra) and bool(rb)) - 1,
    "a_slides": ra["slides"], "b_slides": rb["slides"],
    "min_font": min(ra["min_font"], rb["min_font"]),
    "speak_ok": 10.0 <= ra["est_min"] <= 17.0 and 10.0 <= rb["est_min"] <= 17.0,
    "old_unchanged": unchanged,
    "a": dict(ra, size=pa.stat().st_size), "b": dict(rb, size=pb.stat().st_size),
    "old": odetail,
}, ensure_ascii=False))