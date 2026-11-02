# -*- coding: utf-8 -*-
import json, wave, struct
from pathlib import Path

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
SRC = D / "_edge_v2"
OUT = D / "TRIZa_AI4Science_15min_v2_Andrew_8pct.wav"
GAP = 0.35   # seconds of silence between slides

files = sorted(SRC.glob("s*.wav"))
assert len(files) == 13, len(files)

with wave.open(str(files[0])) as w:
    nch, sw, sr = w.getnchannels(), w.getsampwidth(), w.getframerate()
print("format", nch, "ch", sw * 8, "bit", sr, "Hz")

gap = b"\x00" * (int(sr * GAP) * sw * nch)
cues, chunks, t = [], [], 0.0
for i, f in enumerate(files, 1):
    with wave.open(str(f)) as w:
        assert (w.getnchannels(), w.getsampwidth(), w.getframerate()) == (nch, sw, sr)
        data = w.readframes(w.getnframes())
        dur = w.getnframes() / sr
    cues.append({"slide": i, "start": round(t, 2), "dur": round(dur, 2), "end": round(t + dur, 2)})
    chunks.append(data)
    t += dur
    if i < len(files):
        chunks.append(gap)
        t += GAP

with wave.open(str(OUT), "wb") as w:
    w.setnchannels(nch); w.setsampwidth(sw); w.setframerate(sr)
    w.writeframes(b"".join(chunks))

total = OUT.stat().st_size and t
print("track_sec", round(t, 2), "=", int(t // 60), "min", round(t % 60, 1), "s")
print("bytes", OUT.stat().st_size)

(D / "_cue_v2.json").write_text(json.dumps({"track": OUT.name, "total": round(t, 2), "cues": cues}, ensure_ascii=False, indent=1), encoding="utf-8")
lines = ["TRIZa 15-min talk — Andrew -8% — cue sheet", "track: %s  total: %d:%04.1f" % (OUT.name, int(t//60), t%60), ""]
for c in cues:
    lines.append("slide %2d   starts %2d:%04.1f   (%5.2f s)" % (c["slide"], int(c["start"]//60), c["start"]%60, c["dur"]))
(D / "_cue_v2.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
print("\n".join(lines[:4]))
