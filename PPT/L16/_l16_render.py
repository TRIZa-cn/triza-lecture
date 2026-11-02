# -*- coding: utf-8 -*-
"""L16 renderer: reuse build_w2 visual helpers, render 8-page decks from specs."""
import sys, os
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
import build_w2 as B
from pptx import Presentation
from pptx.util import Inches

OUT = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16"
LABEL = "TRIZ 创新构思法实战研讨班  ·  第 %02d 讲  ·  共 16 讲"


def build_one(spec):
    no = spec["no"]
    prs = Presentation()
    prs.slide_width = Inches(B.SW)
    prs.slide_height = Inches(B.SH)
    label = LABEL % no
    for i, pg in enumerate(spec["pages"], 1):
        k = pg["kind"]
        if k == "cover":
            s = B.cover(prs, pg["kicker"], pg["title"], pg["subtitle"], pg.get("meta", ""),
                        accent=pg.get("accent", B.RED), notes=pg.get("notes", ""))
        else:
            s = B.page(prs, pg["kicker"], pg["title"], notes=pg.get("notes", ""),
                       accent=pg.get("accent", B.RED))
            if k == "page":
                tf = B.textbox(s, B.ML, 2.12, B.CW, 3.55)
                for j, t in enumerate(pg["bullets"]):
                    B.bullet(tf, t, size=pg.get("size", 19), first=(j == 0))
            elif k == "table":
                B.table(s, pg["rows"], B.ML, pg.get("top", 2.12), B.CW, pg["col_w"],
                        size=pg.get("size", 15), row_h=pg.get("row_h", 0.44))
        if pg.get("band"):
            B.band(s, pg["band"], pg.get("band_top", 6.05), size=pg.get("band_size", 15),
                   accent=pg.get("accent", B.RED))
        B.footer(s, i, label, accent=pg.get("accent", B.RED))
    name = "%02d_%s.pptx" % (no, spec["file"])
    path = os.path.join(OUT, name)
    prs.save(path)
    return path, len(spec["pages"])


def build_all(specs):
    res = []
    for sp in specs:
        p, n = build_one(sp)
        chars = sum(len(pg.get("notes", "")) for pg in sp["pages"])
        res.append({"no": sp["no"], "file": os.path.basename(p), "pages": n, "notes": chars,
                    "min": round(chars / 270, 1)})
    return res


if __name__ == "__main__":
    import _l16_spec_a as A
    import _l16_spec_b as C
    out = build_all(A.SPECS + C.SPECS)
    for r in out:
        print(r)
    print("total decks:", len(out), "pages:", sum(r["pages"] for r in out))