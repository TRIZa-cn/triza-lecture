# -*- coding: utf-8 -*-
import math
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
PA = D / "TRIZ_W2A_矛盾解剖_15min.pptx"
PB = D / "TRIZ_W2B_盒子与矛盾论_15min.pptx"

# (slide_index_0based, top_in, old_h_in, new_h_in)
FIX_A = [(6, 2.14, 3.50, 4.05), (8, 2.14, 3.28, 4.32)]
FIX_B = [(4, 2.14, 1.52, 1.64), (4, 3.89, 1.52, 1.64),
         (10, 2.14, 3.30, 4.32), (11, 2.14, 2.55, 2.78)]


def apply(path, fixes):
    prs = Presentation(str(path))
    for si, s in enumerate(prs.slides):
        for sh in s.shapes:
            top = sh.top / 914400.0
            h = sh.height / 914400.0
            for fsi, ftop, fh, nh in fixes:
                if si == fsi and abs(top - ftop) < 0.02 and abs(h - fh) < 0.03:
                    sh.height = Inches(nh)
                    print("  fixed s%d %.2f -> %.2f" % (si + 1, h, nh))
    prs.save(str(path))


def fit(path):
    prs = Presentation(str(path))
    print("==", path.name, "==")
    bad = 0
    for idx, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_table or not sh.has_text_frame:
                continue
            tf = sh.text_frame
            txtall = tf.text.strip()
            if not txtall:
                continue
            w = sh.width / 914400.0
            h = sh.height / 914400.0
            ml = (tf.margin_left or 0) / 914400.0
            mr = (tf.margin_right or 0) / 914400.0
            mt = (tf.margin_top or 0) / 914400.0
            mb = (tf.margin_bottom or 0) / 914400.0
            need = mt + mb
            for p in tf.paragraphs:
                txt = "".join(r.text for r in p.runs)
                if not txt.strip():
                    continue
                sz = max([r.font.size.pt for r in p.runs if r.font.size] or [18])
                cpl = max(1.0, (w - ml - mr) * 72.0 / (sz * 0.96))
                n = sum(len(seg) for seg in txt.split("**"))
                lines = math.ceil(n / cpl)
                ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
                need += lines * sz * ls * 1.18 / 72.0
                need += ((p.space_before or 0) + (p.space_after or 0)) / 12700.0 / 72.0
            if need > h * 1.02:
                bad += 1
                print("  s%-2d need %5.2f > h %5.2f y=%4.2f %s"
                      % (idx, need, h, sh.top / 914400.0, txtall[:26].replace("\n", " ")))
    print("  flagged:", bad)


if __name__ == "__main__":
    print("deck A"); apply(PA, FIX_A)
    print("deck B"); apply(PB, FIX_B)
    fit(PA); fit(PB)