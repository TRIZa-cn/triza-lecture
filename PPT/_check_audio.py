# -*- coding: utf-8 -*-
import json, wave
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
log = (D / "_audio_only.log").read_text(encoding="utf-8")
w = D / "TRIZa_AI4Science_15min_v2_Andrew_8pct.wav"
with wave.open(str(w)) as h:
    sec = h.getnframes() / h.getframerate()
lines = [l for l in log.splitlines() if "elapsed=" in l]
el = float(lines[-1].split("elapsed=")[1].split("s")[0]) if lines else 0.0
print(json.dumps({
  "play_state": "done" if "play done" in log else "unknown",
  "elapsed_sec": round(el, 1),
  "wav_sec": round(sec, 2),
  "powerpoint_running": False,
}, ensure_ascii=False))
