import json, sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Emu

OUT = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
PPTX = OUT / "FOS_科普_15min_优化版.pptx"
SRC = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_20min_20261022\FOS_科普_20min_20261022\FOS_科普_15min.pptx")

REMIND = ["建议用时", "讲者", "讲者要点", "来源", "读法", "讲稿", "备注", "本页",
          "幻灯片", "PPT", "停顿", "语速", "讲的时候"]
META = ["建议用时", "讲者", "来源", "读法", "讲稿", "备注", "本页", "幻灯片", "PPT", "作者"]

res = {}
try:
    prs = Presentation(str(PPTX))
    res["pptx_ok"] = True
except Exception as e:
    print(json.dumps({"pptx_ok": False, "err": str(e)}, ensure_ascii=False)); sys.exit(0)

SW, SH = prs.slide_width, prs.slide_height
slides = list(prs.slides)
res["slides"] = len(slides)
rem, meta, oob = [], [], []
total = 0
empty = []
for i, s in enumerate(slides, 1):
    body = []
    for sh in s.shapes:
        if sh.has_text_frame:
            body.append(sh.text_frame.text)
        try:
            if sh.left is not None and sh.top is not None and sh.width and sh.height:
                if sh.left + sh.width > SW + 2000:
                    oob.append([i, "right", round(Emu(sh.left + sh.width).inches, 2)])
                if sh.top + sh.height > SH + 2000:
                    oob.append([i, "bottom", round(Emu(sh.top + sh.height).inches, 2)])
                if sh.left < -2000 or sh.top < -2000:
                    oob.append([i, "neg", round(Emu(sh.left).inches, 2), round(Emu(sh.top).inches, 2)])
        except Exception:
            pass
    bt = "\n".join(body)
    for k in META:
        if k in bt:
            meta.append([i, k])
    nt = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    total += len(nt)
    if not nt.strip():
        empty.append(i)
    for k in REMIND:
        if k in nt:
            rem.append([i, k])

res["reminder_hits"] = len(rem)
res["reminder_detail"] = rem[:15]
res["slide_meta_hits"] = len(meta)
res["meta_detail"] = meta[:15]
res["notes_chars"] = total
res["notes_min_at260"] = round(total / 260.0, 1)
res["empty_notes"] = empty
res["oob_shapes"] = len(oob)
res["oob_detail"] = oob[:12]
res["source_unchanged"] = SRC.stat().st_mtime < PPTX.stat().st_mtime
print(json.dumps(res, ensure_ascii=False))
