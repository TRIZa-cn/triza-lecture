# -*- coding: utf-8 -*-
import json, os, re, zipfile
from pptx import Presentation
D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
f = os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版_原声解说.pptx")
meta = json.load(open(os.path.join(D, "LWY录音", "_slide_audio.json"), encoding="utf-8"))
prs = Presentation(f)
z = zipfile.ZipFile(f)
names = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
               key=lambda s: int(re.search(r"(\d+)\.xml", s).group(1)))
adv, media = [], 0
for n in names:
    x = z.read(n).decode("utf-8")
    m = re.findall(r'advTm="(\d+)"', x)
    adv.append(int(m[0])/1000 if m else None)
    media += int("<p:audio" in x or "audioFile" in x)
z.close()
durs = [r["dur"] for r in meta["rows"]]
print("slides:", len(prs.slides), "| media:", media, "| adv:", sum(1 for a in adv if a),
      "| 匹配:", all(a and abs(a-(d+0.6))<0.25 for a, d in zip(adv, durs)),
      "| 总长:", round(sum(durs)/60, 2), "min | 体积:", round(os.path.getsize(f)/1048576, 2), "MB")
s7 = prs.slides[6]
print("s7 正文含散热片:", any("散热片" in (sh.text_frame.text if sh.has_text_frame else "") for sh in s7.shapes))
print("s7 正文含鳍片:", any("鳍片" in (sh.text_frame.text if sh.has_text_frame else "") for sh in s7.shapes))
print("s7 讲稿:", s7.notes_slide.notes_text_frame.text[95:150])