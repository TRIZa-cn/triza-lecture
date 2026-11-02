# -*- coding: utf-8 -*-
import json, wave
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
T = D / "TRIZa_AI4Science_15min_v2_Andrew_8pct.wav"
V2 = D / "TRIZa_AI4Science_15min_v2.pptx"
cue = json.loads((D / "_cue_v2.json").read_text(encoding="utf-8"))
with wave.open(str(T)) as w:
    sec = w.getnframes() / w.getframerate()
print(json.dumps({
    "track_exists": T.exists(),
    "track_sec": round(sec, 2),
    "track_mb": round(T.stat().st_size / 1048576, 1),
    "cues": len(cue["cues"]),
    "cues_match_slides": [c["slide"] for c in cue["cues"]] == list(range(1, 14)),
    "v2_unchanged": V2.stat().st_mtime < T.stat().st_mtime,
    "v2_bytes": V2.stat().st_size,
}, ensure_ascii=False))
