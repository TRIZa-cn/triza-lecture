# -*- coding: utf-8 -*-
import json, re, datetime
from pathlib import Path
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
OPT = D / "TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
NAR = D / "TRIZ_W2A_矛盾解剖_15min_优化版_自动解说.pptx"
stats = json.loads((D / "_A_stats.json").read_text(encoding="utf-8"))

def shape_starting(slide, head):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip().startswith(head):
            return sh
    return None

prs = Presentation(str(NAR))
audio, adv, tot_ms, slides, gaps, chars = 0, 0, 0, 0, [], 0
for i, s in enumerate(prs.slides, 1):
    slides += 1
    x = s.element.xml
    n = x.count("audioFile")
    audio += n
    m = re.search(r'advTm="(\d+)"', x)
    ms = int(m.group(1)) if m else 0
    if ms:
        adv += 1
        tot_ms += ms
    if n != 1 or not ms or "p:audio" not in x:
        gaps.append(i)
    chars += len(s.notes_slide.notes_text_frame.text)

audio_s = tot_ms / 1000.0 - 0.6 * slides
cpm = round(chars / audio_s * 60, 1)

opt = Presentation(str(OPT))
s4 = shape_starting(opt.slides[3], "技术矛盾给你方向").text_frame.text
s10 = shape_starting(opt.slides[9], "创新等级不取决于投入多少").text_frame.text
notes_all = [s.notes_slide.notes_text_frame.text for s in opt.slides]
c1 = ("停在第一行" in s4) and ("同一个对象" in s4)
c2 = ("一次性消失" in s10) and ("换了原理" in s10)
c3 = not any("ARIZ" in t.upper() for t in notes_all)
c4 = not any("---" in t for t in notes_all)
c5 = any(("技术矛盾" in t and ("翻译成" in t or "转成" in t)) for t in notes_all)
text_ok = c1 and c2 and c3 and c4 and c5
SUB = {"s4_explains": c1, "s10_explains": c2, "no_ariz": c3,
       "no_marker": c4, "mgmt_consistent": c5}

b_ok = True
detail = {}
for name in ["TRIZ_W2B_盒子与矛盾论_15min.pptx", "TRIZa_AI4Science_15min.pptx",
             "TRIZa_AI4Science_15min_v2.pptx", "TRIZa_15min_handoff.zip"]:
    f = D / name
    exp_sz, exp_mt = stats[name]
    ok = f.is_file() and f.stat().st_size == exp_sz and \
        datetime.datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M") == exp_mt
    detail[name] = "ok" if ok else "CHANGED"
    b_ok = b_ok and ok

a = D / "TRIZ_W2A_矛盾解剖_15min.pptx"
exp_sz, exp_mt = stats["TRIZ_W2A_矛盾解剖_15min.pptx"]
a_ok = a.stat().st_size == exp_sz and \
    datetime.datetime.fromtimestamp(a.stat().st_mtime).strftime("%Y-%m-%d %H:%M") == exp_mt

out = {
    "optimized_files": int(OPT.is_file()) + int(NAR.is_file()),
    "a_audio": audio, "a_slides": slides, "a_adv": adv, "gaps": gaps,
    "adv_ok": adv == slides and not gaps,
    "chars_per_min": cpm, "audio_sec": round(audio_s, 1), "chars": chars,
    "text_ok": text_ok, "sub": SUB,
    "b_unchanged": b_ok, "old": detail,
    "orig_a_unchanged": a_ok,
    "size_opt": OPT.stat().st_size, "size_nar": NAR.stat().st_size,
}
print(json.dumps(out, ensure_ascii=False))