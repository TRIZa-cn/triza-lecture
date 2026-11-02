# -*- coding: utf-8 -*-
from pathlib import Path

from pptx import Presentation

OUT = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
PPTX = OUT / "FOS_科普_15min_优化版+闰秒案例.pptx"
MD = OUT / "讲稿_逐字稿_闰秒案例版.md"

prs = Presentation(str(PPTX))
notes = [s.notes_slide.notes_text_frame.text for s in prs.slides]

heads = []
for sl in prs.slides:
    best = (0, "")
    for shp in sl.shapes:
        if not shp.has_text_frame:
            continue
        for p in shp.text_frame.paragraphs:
            for r in p.runs:
                sz = r.font.size.pt if r.font.size else 0
                if sz > best[0] and r.text.strip():
                    best = (sz, r.text.strip())
    heads.append(best[1])

main = sum(len(n) for n in notes[:30])
app = sum(len(n) for n in notes[30:])
lines = []
lines.append("# FOS 科普 15 分钟 · 逐字口播稿（闰秒案例版）")
lines.append("")
lines.append("> 共 %d 页。正文 30 页 %d 字，按每分钟 260 字约 %.1f 分钟；附录 5 页 %d 字。"
             % (len(notes), main, main / 260.0, app))
lines.append("> 每页下方文字即该页 notes，可直接照读。")
lines.append("")
for i, (h, n) in enumerate(zip(heads, notes), 1):
    tag = "" if i <= 30 else "（附录）"
    lines.append("## %02d%s · %s" % (i, tag, h))
    lines.append("")
    lines.append(n)
    lines.append("")
MD.write_text("\n".join(lines), encoding="utf-8")
print("MD", MD.name, MD.stat().st_size, "chars", len("\n".join(lines)))
print("MAIN", main, "APP", app, "MIN", round(main / 260.0, 1))
