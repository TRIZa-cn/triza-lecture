# -*- coding: utf-8 -*-
import math
from pathlib import Path
from pptx import Presentation
from pptx.util import Emu

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")


def fit(path):
    prs = Presentation(str(path))
    print("==", path.name, "==")
    bad = 0
    for idx, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_table or not sh.has_text_frame:
                continue
            tf = sh.text_frame
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
                pw = w - ml - mr
                unit = sz * 0.96
                cpl = max(1.0, pw * 72.0 / unit)
                n = 0
                for seg in txt.split("**"):
                    n += len(seg)
                lines = math.ceil(n / cpl)
                ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
                need += lines * sz * ls * 1.18 / 72.0
                need += ((p.space_before or 0) + (p.space_after or 0)) / 12700.0 / 72.0
            if need > h * 1.02:
                bad += 1
                print("  s%-2d need %5.2f > h %5.2f  y=%4.2f  %s"
                      % (idx, need, h, sh.top / 914400.0,
                         "".join(r.text for r in tf.paragraphs[0].runs)[:30]))
    print("  flagged:", bad)


fit(D / "TRIZ_W2A_矛盾解剖_15min.pptx")
fit(D / "TRIZ_W2B_盒子与矛盾论_15min.pptx")