# -*- coding: utf-8 -*-
import hashlib, json, pathlib, re
from pptx import Presentation
OUT = pathlib.Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\out")
PPT = pathlib.Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
FOOT = "TRIZ 创新构思法实战研讨班  \u00b7  \u7b2c %d \u6b21 Workshop"
files = sorted(OUT.glob("*.pptx"))
pages, notes, bad_font, bad_foot, ex = [], [], [], [], []
figs = {}
for f in files:
    prs = Presentation(str(f))
    pages.append(len(prs.slides))
    tot = 0
    for sl in prs.slides:
        tot += len(sl.notes_slide.notes_text_frame.text) if sl.has_notes_slide else 0
    notes.append(tot)
    n = int(re.search(r"\u7b2c(\d+)\u6b21", f.name).group(1))
    figs[n] = 0
    for i, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if sh.name.startswith("fig-"):
                figs[n] += 1
                continue
            if not sh.has_text_frame:
                continue
            txt = sh.text_frame.text.strip()
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.size is None:
                        continue
                    pt = r.font.size.pt
                    if pt >= 20:
                        continue
                    ok = (txt.startswith("TRIZ \u521b\u65b0\u6784\u601d\u6cd5") or
                          re.match(r"^\d{2} / 12$", txt) or
                          re.match(r"^\d{2}\s+\u00b7", txt) or
                          (i == 1 and pt >= 12))
                    if not ok:
                        bad_font.append([f.name, i, pt, txt[:12]])
        if i >= 2:
            if not any(sh.has_text_frame and sh.text_frame.text.strip() == (FOOT % n) for sh in sl.shapes):
                bad_foot.append([f.name, i])
    if len(prs.slides) >= 11:
        txt = " ".join(sh.text_frame.text for sh in prs.slides[10].shapes if sh.has_text_frame)
        ex.append("TRIZa" in txt and "\u63d0\u4ea4\u8981\u6c42" in txt)

deck_ok = len(files) == 15 and all(x == 12 for x in pages)
notes_ok = len(notes) == 15 and all(3000 <= x <= 3500 for x in notes)
font_ok = not bad_font
exercise_ok = len(ex) == 15 and all(ex)
footer_ok = not bad_foot

fig_ok = figs.get(2, 0) >= 6 and figs.get(3, 0) >= 6 and (figs.get(2, 0) + figs.get(3, 0)) >= 14

d4 = [f for f in files if "\u7b2c04\u6b21" in f.name][0]
prs4 = Presentation(str(d4))
texts, nrow = [], None
for sl in prs4.slides:
    for sh in sl.shapes:
        if sh.has_text_frame:
            texts.append(sh.text_frame.text)
        if getattr(sh, "has_table", False) and sh.has_table:
            nrow = len(sh.table.rows)
    if sl.has_notes_slide:
        texts.append(sl.notes_slide.notes_text_frame.text)
blob = "\n".join(texts)
six = blob.count("\u516d\u6761")
five = blob.count("\u4e94\u6761")
consistency_ok = (six > 0) and (five == 0) and (nrow == 7)

def sha(p):
    h = hashlib.sha256()
    h.update(p.read_bytes())
    return h.hexdigest()[:16]

w1 = PPT / "TRIZ_W1_\u77db\u76fe15m.pptx"
pl = PPT / "TRIZ_W1_\u77db\u76fe15m_\u6f14\u64ad\u7248.pptx"
existing_unchanged = (w1.stat().st_size == 704606 and sha(w1) == "b7d5cf1da957f33f"
                      and pl.stat().st_size == 16890575 and sha(pl) == "f8a2c6ab2e2c1e56")
recheck_ok = deck_ok and notes_ok and font_ok and exercise_ok and footer_ok

print(json.dumps({
    "deck_ok": deck_ok, "notes_ok": notes_ok, "font_ok": font_ok,
    "exercise_ok": exercise_ok, "footer_ok": footer_ok,
    "existing_unchanged": existing_unchanged,
    "fig_ok": fig_ok, "consistency_ok": consistency_ok, "recheck_ok": recheck_ok,
    "detail": {"decks": len(files), "pages": sorted(set(pages)),
               "notes_min": min(notes), "notes_max": max(notes),
               "bad_font": len(bad_font), "bad_foot": len(bad_foot),
               "fig_shapes": figs, "six": six, "five": five, "table_rows": nrow,
               "bad_font_sample": bad_font[:3]},
}, ensure_ascii=False))