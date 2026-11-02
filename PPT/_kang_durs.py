# -*- coding: utf-8 -*-
import json, wave
from pathlib import Path
BASE = Path(r"D:\doc\ZGCA\Lectures\TRIZ")
durs = []
for p in sorted((BASE / "_kang").glob("s*.wav")):
    with wave.open(str(p), "rb") as w:
        durs.append(round(w.getnframes() / w.getframerate(), 2))
words = json.loads((BASE / "_notes_dump.json").read_text(encoding="utf-8"))["words"]
tot = sum(durs)
(BASE / "_kang_durs.json").write_text(json.dumps({"durations": durs, "total": round(tot, 1), "slides": len(durs)}, indent=1), encoding="utf-8")
print("durations", durs)
print("total_sec", round(tot, 1), "total_min", round(tot / 60, 2), "wpm", round(words / (tot / 60), 1))
