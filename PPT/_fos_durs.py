# -*- coding: utf-8 -*-
import json, wave
from pathlib import Path
BASE = Path(r"D:\doc\ZGCA\Lectures\TRIZ")
durs = []
for p in sorted((BASE / "_fos").glob("s*.wav")):
    with wave.open(str(p), "rb") as w:
        durs.append(round(w.getnframes() / w.getframerate(), 2))
tot = sum(durs)
print("files", len(durs))
print("durations", durs)
print("total_sec", round(tot, 1), "total_min", round(tot / 60, 2))
(BASE / "_fos_durs.json").write_text(json.dumps({"durations": durs, "total": round(tot, 1)}), encoding="utf-8")
