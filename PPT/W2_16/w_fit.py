# -*- coding: utf-8 -*-
"""用真实字体度量文本所需高度，检查是否溢出形状。"""
import sys, pathlib, json
from pptx import Presentation
from PIL import ImageFont

FONTS = [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyhbd.ttc",
         r"C:\Windows\Fonts\simhei.ttf"]
_fc = {}


def font(size, bold=False):
    key = (size, bold)
    if key not in _fc:
        path = FONTS[1] if bold else FONTS[0]
        _fc[key] = ImageFont.truetype(path, max(1, int(round(size))))
    return _fc[key]


def wrap_lines(txt, size, bold, avail_pt):
    if not txt:
        return 1
    f = font(size, bold)
    lines, cur = 1, ""
    for ch in txt:
        trial = cur + ch
        if f.getlength(trial) <= avail_pt or not cur:
            cur = trial
        else:
            lines += 1
            cur = ch
    return lines


def check(path):
    prs = Presentation(str(path))
    W, H = prs.slide_width / 914400.0, prs.slide_height / 914400.0
    bad, small = [], []
    for si, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            l = sh.left / 914400.0; t = sh.top / 914400.0
            w = sh.width / 914400.0; h = sh.height / 914400.0
            if l + w > W + 0.02 or t + h > H + 0.02 or l < -0.02 or t < -0.02:
                bad.append((si, "越界", round(l, 2), round(t, 2), round(w, 2), round(h, 2)))
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            ml = (tf.margin_left or 0) / 914400.0
            mr = (tf.margin_right or 0) / 914400.0
            mt = (tf.margin_top or 0) / 914400.0
            mb = (tf.margin_bottom or 0) / 914400.0
            avail_w = (w - ml - mr) * 72.0
            avail_h = (h - mt - mb) * 72.0
            need = 0.0
            first = ""
            for p in tf.paragraphs:
                txt = "".join(r.text for r in p.runs)
                if not first:
                    first = txt[:18]
                if not txt.strip():
                    need += 8
                    continue
                size = max([r.font.size.pt for r in p.runs if r.font.size] or [18])
                bold = any(r.font.bold for r in p.runs)
                if size < 20:
                    small.append((si, size, txt[:16]))
                ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
                sb = (p.space_before.pt if p.space_before else 0)
                sa = (p.space_after.pt if p.space_after else 0)
                n = wrap_lines(txt, size, bold, avail_w)
                need += n * size * 1.2 * ls + sb + sa
            if avail_h > 0 and need > avail_h * 1.02:
                bad.append((si, "溢出", first, round(need, 1), round(avail_h, 1)))
    return dict(file=path.name, slides=len(prs.slides._sldIdLst),
                overflow=len(bad), small=len(small),
                detail=bad[:8], small_detail=small[:6])


if __name__ == "__main__":
    root = pathlib.Path(sys.argv[1])
    out = []
    for f in sorted(root.glob("*.pptx")):
        r = check(f)
        out.append(r)
        print("%-26s pages=%2d overflow=%d small=%d"
              % (r["file"], r["slides"], r["overflow"], r["small"]))
        for d in r["detail"]:
            print("    ", d)
        for d in r["small_detail"]:
            print("     small", d)
    print(json.dumps(out, ensure_ascii=False)[:400])