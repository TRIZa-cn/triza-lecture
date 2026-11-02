# -*- coding: utf-8 -*-
"""TRIZa-Lecture-audit.py — 课件可读性与时长审计。

统计项自动计算；LV 分级与"是否已解释"由人工术语表提供。
不输出合成分数：按维度给 A/B/C/D 档，总体取最差维度（木桶原则）。

用法:
  python TRIZa-Lecture-audit.py deck.pptx [--cpm 260] [--target-min 15] [--terms terms.json]
"""
import argparse, json, re
from pathlib import Path
from pptx import Presentation

FOOTER_TOP = 6.9  # 英寸


def slide_text(slide):
    parts = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            parts.append(sh.text_frame.text.strip())
        elif sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    if c.text_frame.text.strip():
                        parts.append(c.text_frame.text.strip())
    return "\n".join(parts)


def collect_fonts(slide):
    body, table = [], []
    for sh in slide.shapes:
        top = (sh.top or 0) / 914400.0
        if sh.has_text_frame:
            if top >= FOOTER_TOP:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.text.strip() and r.font.size:
                        body.append(r.font.size.pt)
        elif sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    for p in c.text_frame.paragraphs:
                        for r in p.runs:
                            if r.text.strip() and r.font.size:
                                table.append(r.font.size.pt)
    return body, table


def grade(value, cuts):
    """cuts: [(上限, 档)] 由好到差。"""
    for lim, g in cuts:
        if value <= lim:
            return g
    return cuts[-1][1]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--cpm", type=float, default=260.0)
    ap.add_argument("--target-min", type=float, default=0.0)
    ap.add_argument("--terms", default="")
    a = ap.parse_args()

    prs = Presentation(a.deck)
    rows, all_notes = [], []
    for i, s in enumerate(prs.slides, 1):
        txt = slide_text(s)
        body, table = collect_fonts(s)
        notes = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
        all_notes.append(notes)
        sents = [x for x in re.split(r"[。！？；\n]", txt) if x.strip()]
        lens = [len(re.sub(r"\s", "", x)) for x in sents]
        rows.append({
            "slide": i,
            "chars": len(re.sub(r"\s", "", txt)),
            "bullets": txt.count("\u25b8"),
            "table_rows": sum(len(sh.table.rows) for sh in s.shapes if sh.has_table),
            "min_font_body": min(body) if body else None,
            "min_font_table": min(table) if table else None,
            "avg_sent": round(sum(lens) / len(lens), 1) if lens else 0,
            "max_sent": max(lens) if lens else 0,
            "notes_chars": len(notes),
            "notes_sec": round(len(notes) / a.cpm * 60, 1),
        })

    notes_all = "".join(all_notes)
    nl = [len(x) for x in re.split(r"[。！？；]", notes_all) if x.strip()]
    deck_text = "\n".join(slide_text(s) for s in prs.slides)
    est_min = round(len(notes_all) / a.cpm, 2)
    avg = round(sum(nl) / len(nl), 1) if nl else 0

    out = {
        "deck": str(a.deck),
        "slides": len(prs.slides),
        "chars_slides": sum(r["chars"] for r in rows),
        "notes_chars": len(notes_all),
        "notes_avg_sent": avg,
        "notes_max_sent": max(nl) if nl else 0,
        "notes_long_sents": sum(1 for x in nl if x > 45),
        "notes_questions": notes_all.count("？"),
        "est_minutes": est_min,
        "cpm": a.cpm,
        "min_font_body": min([r["min_font_body"] for r in rows if r["min_font_body"]] or [0]),
        "min_font_table": min([r["min_font_table"] for r in rows if r["min_font_table"]] or [0]),
        "per_slide": rows,
    }

    grades = {}
    grades["sentence"] = grade(avg, [(24, "A"), (30, "B"), (36, "C"), (999, "D")])
    over = [r["slide"] for r in rows if r["chars"] > 250 or r["bullets"] > 6]
    grades["density"] = "A" if not over else ("B" if len(over) <= 2 else "C")
    out["density_over_slides"] = over

    if a.target_min:
        d = abs(est_min - a.target_min) / a.target_min
        out["target_min"] = a.target_min
        out["delta_min"] = round(est_min - a.target_min, 2)
        grades["time"] = "A" if d <= .10 else "B" if d <= .25 else "C" if d <= .40 else "D"

    if a.terms:
        spec = json.loads(Path(a.terms).read_text(encoding="utf-8"))
        found, missing = [], []
        for t in spec["terms"]:
            al = t.get("aliases", [t["name"]])
            if not any(k in deck_text for k in al):
                missing.append(t["name"]); continue
            first = next((r["slide"] for r in rows
                          if any(k in slide_text(prs.slides[r["slide"] - 1]) for k in al)), None)
            found.append({**t, "first_slide": first})
        need = [t for t in found if t["lv"] >= 2]
        unexp = [t for t in need if not t["explained"]]
        U = len(unexp) / len(need) if need else None
        out["terms"] = {
            "n_found": len(found), "missing": missing,
            "n_need_explain": len(need), "n_unexplained": len(unexp),
            "U": round(U, 3) if U is not None else None,
            "unexplained": [{"name": t["name"], "slide": t["first_slide"], "why": t.get("where", "")}
                            for t in unexp],
        }
        grades["terms"] = "n/a" if U is None else grade(U, [(.10, "A"), (.20, "B"), (.35, "C"), (9, "D")])

    order = {"A": 0, "B": 1, "C": 2, "D": 3}
    worst = max((g for g in grades.values() if g in order), key=lambda g: order[g], default=None)
    out["grades"] = grades
    out["overall"] = worst
    print(json.dumps(out, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()