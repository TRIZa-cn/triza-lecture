# -*- coding: utf-8 -*-
import json, wave
from pathlib import Path
BASE = Path(r"D:\doc\ZGCA\Lectures\TRIZ")
d = json.loads((BASE / "_notes_tts.json").read_text(encoding="utf-8"))
print("S02 TEXT:", d["notes"][1][:220])
durs = []
for p in sorted((BASE / "_kang2").glob("s*.wav")):
    with wave.open(str(p), "rb") as w:
        durs.append(round(w.getnframes() / w.getframerate(), 2))
tot = sum(durs)
print("durations", durs)
print("total_min", round(tot / 60, 2), "wpm", round(d["words"] / (tot / 60), 1))
