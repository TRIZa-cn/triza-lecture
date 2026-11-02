# -*- coding: utf-8 -*-
import json, os, re, zipfile
from pptx import Presentation
D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
f = os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版_原声解说.pptx")
meta = json.load(open(os.path.join(D, "LWY录音", "_slide_audio.json"), encoding="utf-8"))
prs = Presentation(f)
print("slides:", len(prs.slides), "| size MB:", round(os.path.getsize(f)/1048576, 2))
z = zipfile.ZipFile(f)
names = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
               key=lambda s: int(re.search(r"(\d+)\.xml", s).group(1)))
adv = []; media = 0
for n in names:
    x = z.read(n).decode("utf-8")
    m = re.findall(r'advTm="(\d+)"', x)
    adv.append(int(m[0])/1000 if m else None)
    media += int("audio" in x.lower() or "<p:audio" in x)
z.close()
durs = [r["dur"] for r in meta["rows"]]
ok = all(a and abs(a - (d + 0.6)) < 0.25 for a, d in zip(adv, durs))
print("media:", media, "| adv:", sum(1 for a in adv if a), "| advTm 与音轨匹配:", ok)
print("逐页 音轨秒 / 自动翻页秒:")
for i, (d, a) in enumerate(zip(durs, adv), 1):
    print(f"   s{i:02d}  {d:6.2f}  {a}")
print("总时长:", round(sum(durs), 1), "s =", round(sum(durs)/60, 2), "min")
print("原 TTS 版:", round(os.path.getsize(os.path.join(D,'TRIZ_W2A_矛盾解剖_15min_优化版_自动解说.pptx'))/1048576, 2), "MB")
print("原声版:", round(os.path.getsize(f)/1048576, 2), "MB")