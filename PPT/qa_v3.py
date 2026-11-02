# -*- coding: utf-8 -*-
"""Read-only acceptance check for the 35-page leap-second deck. Prints one-line JSON."""
import json
import re
import sys
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
from build_deck_v2 import PRINCIPLES  # noqa: E402
from build_deck_v3 import DESC  # noqa: E402

OUT = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
PPTX = OUT / "FOS_科普_15min_优化版+闰秒案例.pptx"
SRC = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_20min_20261022\FOS_科普_20min_20261022\FOS_科普_15min.pptx")

REMIND = re.compile(r"建议用时|讲者|演讲者|讲稿|读法|备注|本页|幻灯片|PPT|停顿|语速|提示卡|可选|备用")
META = re.compile(r"讲者|演讲者|作者|来源|本页|建议用时|备注|联系方式|邮箱|参考文件")

res = {}
try:
    prs = Presentation(str(PPTX))
    res["pptx_ok"] = True
except Exception as e:
    print(json.dumps({"pptx_ok": False, "err": str(e)}, ensure_ascii=False))
    raise SystemExit(1)

slides = list(prs.slides)
res["slides"] = len(slides)

rem = meta = empty = oob = 0
notes_lens = []
min_font = 999
viol = []
desc_count = 0
desc_max = 0
desc_min = 0
sw, sh = prs.slide_width, prs.slide_height
principle_text = {4: [], 5: []}

for idx, sl in enumerate(slides, 1):
    note = sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else ""
    notes_lens.append(len(note))
    if not note.strip():
        empty += 1
    if REMIND.search(note):
        rem += 1
    stext = []
    for shp in sl.shapes:
        try:
            if shp.left is not None and (
                    shp.left < -10000 or shp.top < -10000
                    or shp.left + shp.width > sw + 10000
                    or shp.top + shp.height > sh + 10000):
                oob += 1
        except TypeError:
            pass
        if not shp.has_text_frame:
            continue
        for p in shp.text_frame.paragraphs:
            for r in p.runs:
                stext.append(r.text)
                sz = r.font.size.pt if r.font.size else None
                if sz is None:
                    continue
                if idx in (4, 5) and abs(sz - 13) < 0.01:
                    desc_count += 1
                    desc_max = max(desc_max, len(r.text))
                    desc_min = min(desc_min or 99, len(r.text))
                else:
                    min_font = min(min_font, sz)
                    if sz < 20:
                        viol.append([idx, sz, r.text[:12]])
    if idx in (4, 5):
        principle_text[idx] = stext
    if META.search("".join(stext)):
        meta += 1

res["reminder_hits"] = rem
res["slide_meta_hits"] = meta
res["notes_chars"] = sum(notes_lens)
res["notes_main_chars"] = sum(notes_lens[:30])
res["notes_min_at260"] = round(sum(notes_lens[:30]) / 260.0, 1)
res["notes_min_entry"] = min(notes_lens)
res["empty_notes"] = empty
res["min_font"] = min_font
res["font_violations"] = viol[:8]
res["font_viol_count"] = len(viol)
res["desc_count"] = desc_count
res["desc_max_len"] = desc_max
res["desc_min_len"] = desc_min
res["desc_over20"] = [k for k, v in DESC.items() if len(v) > 20]
res["oob_shapes"] = oob

all_text = "".join(principle_text.get(4, []) + principle_text.get(5, []))
res["principles_found"] = sum(1 for n in PRINCIPLES if n in all_text)
res["principles_total"] = len(PRINCIPLES)


def slide_text(i):
    return "".join(r.text for shp in slides[i - 1].shapes if shp.has_text_frame
                   for p in shp.text_frame.paragraphs for r in p.runs)


res["leap_pages_ok"] = sum(1 for i in [6, 7, 20, 21, 22, 23, 24, 25] if "闰" in slide_text(i))
res["appendix_pages_ok"] = sum(1 for i in range(31, 36) if "附录" in slide_text(i))
res["source_unchanged"] = SRC.stat().st_mtime < PPTX.stat().st_mtime
res["pptx_bytes"] = PPTX.stat().st_size

print(json.dumps(res, ensure_ascii=False))
