# -*- coding: utf-8 -*-
import json, re, datetime
from pathlib import Path
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
PAIRS = [("a", D / "TRIZ_W2A_矛盾解剖_15min_自动解说.pptx"),
         ("b", D / "TRIZ_W2B_盒子与矛盾论_15min_自动解说.pptx")]
OLDS = {"TRIZa_AI4Science_15min.pptx": (426613, "2026-10-24 07:59"),
        "TRIZa_AI4Science_15min_v2.pptx": (401376, "2026-10-24 09:20"),
        "TRIZ_W2A_矛盾解剖_15min.pptx": (75741, "2026-10-28 19:38"),
        "TRIZ_W2B_盒子与矛盾论_15min.pptx": (80795, "2026-10-28 19:38"),
        "TRIZa_15min_handoff.zip": (22037056, "2026-10-24 15:57")}

res = {"narrated_decks": 0, "old": {}}
for tag, f in PAIRS:
    prs = Presentation(str(f))
    audio, adv, tot_ms, slides, gaps = 0, 0, 0, 0, []
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
    res[tag] = {"slides": slides, "audio": audio, "adv": adv,
                "total_s": round(tot_ms / 1000.0, 1), "gaps": gaps,
                "size": f.stat().st_size}

unch = True
for name, (sz, mt) in OLDS.items():
    g = D / name
    ok = g.is_file() and g.stat().st_size == sz and \
        datetime.datetime.fromtimestamp(g.stat().st_mtime).strftime("%Y-%m-%d %H:%M") == mt
    res["old"][name] = "ok" if ok else "CHANGED"
    unch = unch and ok

out = {
    "narrated_decks": sum(1 for _, f in PAIRS if f.is_file()),
    "a_audio": res["a"]["audio"], "b_audio": res["b"]["audio"],
    "a_slides": res["a"]["slides"], "b_slides": res["b"]["slides"],
    "a_adv": res["a"]["adv"], "b_adv": res["b"]["adv"],
    "a_total_s": res["a"]["total_s"], "b_total_s": res["b"]["total_s"],
    "a_gaps": res["a"]["gaps"], "b_gaps": res["b"]["gaps"],
    "adv_ok": res["a"]["adv"] == res["a"]["slides"] and res["b"]["adv"] == res["b"]["slides"],
    "runtime_ok": 600 <= res["a"]["total_s"] <= 1020 and 600 <= res["b"]["total_s"] <= 1020,
    "old_unchanged": unch,
    "sizes": {"a": res["a"]["size"], "b": res["b"]["size"]},
    "old": res["old"],
}
print(json.dumps(out, ensure_ascii=False))