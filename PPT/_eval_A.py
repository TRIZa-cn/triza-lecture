# -*- coding: utf-8 -*-
import json, re
from pathlib import Path
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
F = D / "TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(str(F))

def fonts(run):
    s = run.font.size
    return s.pt if s else None

rows = []
tot_chars = 0
for i, s in enumerate(prs.slides, 1):
    body, tbl, notes, sizes = [], 0, "", []
    for sh in s.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t:
                body.append(t)
                for para in sh.text_frame.paragraphs:
                    for r in para.runs:
                        f = fonts(r)
                        if f and r.text.strip():
                            sizes.append(f)
        elif sh.has_table:
            tbl += 1
            for row in sh.table.rows:
                for c in row.cells:
                    for para in c.text_frame.paragraphs:
                        for r in para.runs:
                            f = fonts(r)
                            if f and r.text.strip():
                                sizes.append(f)
    notes = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
    txt = "\n".join(body)
    chars = len(re.sub(r"\s", "", txt))
    tot_chars += chars
    sents = [x for x in re.split(r"[。！？；\n]", txt) if x.strip()]
    lens = [len(re.sub(r"\s", "", x)) for x in sents]
    rows.append({
        "slide": i, "shapes": len(s.shapes), "chars": chars, "tables": tbl,
        "min_font": min(sizes) if sizes else None,
        "notes_chars": len(notes), "notes_sec": round(len(notes) / 4.31, 1),
        "sents": len(sents), "avg_sent": round(sum(lens) / len(lens), 1) if lens else 0,
        "max_sent": max(lens) if lens else 0,
        "long_sents": sum(1 for l in lens if l > 40),
        "numbers": len(re.findall(r"\d", txt)),
        "bullets": txt.count("▸"),
    })

all_notes = "".join(s.notes_slide.notes_text_frame.text for s in prs.slides)
nl = [len(x) for x in re.split(r"[。！？；]", all_notes) if x.strip()]
out = {
    "slides": len(prs.slides),
    "slide_chars_total": tot_chars,
    "min_font_body": min(r["min_font"] for r in rows if r["min_font"]),
    "notes_total": len(all_notes),
    "notes_avg_sent": round(sum(nl) / len(nl), 1),
    "notes_max_sent": max(nl),
    "notes_long_sents": sum(1 for l in nl if l > 45),
    "notes_questions": all_notes.count("？"),
    "per_slide": rows,
}
print(json.dumps(out, ensure_ascii=False))