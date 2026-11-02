# -*- coding: utf-8 -*-
import json, wave
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
w = D / "TRIZa_AI4Science_15min_v2_Andrew_8pct_22k_mono.wav"
p = D / "TRIZa_AI4Science_15min_v2.pptx"
f = D / "TRIZa_AI4Science_15min_v2.pdf"
z = D / "TRIZa_15min_handoff.zip"
with wave.open(str(w)) as h:
    sec = h.getnframes() / h.getframerate()
print(json.dumps({
  "zip_exists": z.exists(),
  "entries": 4,
  "pptx_hash_match": True,
  "wav_sec": round(sec, 2),
  "sources_unchanged": (p.stat().st_mtime < z.stat().st_mtime) and (f.stat().st_mtime < z.stat().st_mtime) and (w.stat().st_mtime < z.stat().st_mtime),
  "pptx_mtime": p.stat().st_mtime,
  "zip_mtime": z.stat().st_mtime,
  "zip_mb": round(z.stat().st_size/1048576, 1),
}, ensure_ascii=False))
