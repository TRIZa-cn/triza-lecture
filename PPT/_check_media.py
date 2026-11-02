# -*- coding: utf-8 -*-
import json, os
from pptx import Presentation
D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
NAR = os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版_自动解说.pptx")
print("size MB:", round(os.path.getsize(NAR)/1048576, 2))
pr = Presentation(NAR)
ns_p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
audio = 0; adv = []; sample = ""
for i, s in enumerate(pr.slides, 1):
    rels = [r.reltype.rsplit("/", 1)[-1] for r in s.part.rels.values()]
    media = [r for r in rels if r in ("audio", "video", "media")]
    has_timing = s._element.find(ns_p + "timing") is not None
    tr = s._element.find(ns_p + "transition")
    advtm = None
    if tr is not None:
        advtm = tr.get("advTm") or tr.get("advClick")
    audio += int(bool(media))
    adv.append(advtm)
    if i == 1:
        sample = f"rels={rels} timing={has_timing} transition={tr is not None} advTm={advtm}"
print("slides with media rel:", audio, "/", len(pr.slides))
print("advance times:", adv)
print("slide1:", sample)
n10 = pr.slides[9].notes_slide.notes_text_frame.text
bar = [sh.text_frame.text for sh in pr.slides[9].shapes if sh.name == "Rectangle 5"][0]
print("notes10 has 主体:", "主体" in n10, "| bar has 承担主体:", "承担主体" in bar)
print("notes10 len:", len(n10))