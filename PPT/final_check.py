import json, sys
from pathlib import Path
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
from pptx import Presentation
import build_deck_v2 as B

OUT = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
PPTX = OUT / "FOS_科普_15min_优化版.pptx"
SRC = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_20min_20261022\FOS_科普_20min_20261022\FOS_科普_15min.pptx")

REMIND = ["建议用时", "讲者", "来源", "读法", "讲稿", "备注", "本页", "幻灯片", "PPT", "停顿", "语速"]
META = ["建议用时", "讲者", "来源", "读法", "讲稿", "备注", "本页", "幻灯片", "PPT", "作者"]

prs = Presentation(str(PPTX))
res = {"pptx_ok": True, "slides": len(prs.slides)}
rem = 0; meta = 0; total = 0; mins = []
for s in prs.slides:
    bt = "\n".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
    for k in META:
        if k in bt:
            meta += 1
    nt = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    total += len(nt)
    mins.append(len(nt))
    for k in REMIND:
        if k in nt:
            rem += 1
res["reminder_hits"] = rem
res["slide_meta_hits"] = meta
res["notes_chars"] = total
res["notes_min_entry"] = min(mins)
res["notes_min_at260"] = round(total / 260.0, 1)

grid = "\n".join(sh.text_frame.text for s in list(prs.slides)[3:5]
                 for sh in s.shapes if sh.has_text_frame)
found = sum(1 for name in B.PRINCIPLES if name in grid)
res["principles_found"] = found
res["principles_total"] = len(B.PRINCIPLES)
res["source_unchanged"] = SRC.stat().st_mtime < PPTX.stat().st_mtime
print(json.dumps(res, ensure_ascii=False))
