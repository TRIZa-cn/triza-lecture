# -*- coding: utf-8 -*-
import json, wave
import numpy as np
from pathlib import Path

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
SRC = D / "TRIZa_AI4Science_15min_v2_Andrew_8pct.wav"
OUT = D / "TRIZa_AI4Science_15min_v2_Andrew_8pct_22k_mono.wav"

with wave.open(str(SRC)) as w:
    nch, sw, sr = w.getnchannels(), w.getsampwidth(), w.getframerate()
    a = np.frombuffer(w.readframes(w.getnframes()), dtype=np.int16)
a = a.reshape(-1, nch).astype(np.int32)
mono = a.mean(axis=1)                       # stereo -> mono
mono = (mono[: len(mono) // 2 * 2].reshape(-1, 2).mean(axis=1))  # 44100 -> 22050 (box filter)
out = np.clip(np.round(mono), -32768, 32767).astype(np.int16)

with wave.open(str(OUT), "wb") as w:
    w.setnchannels(1); w.setsampwidth(2); w.setframerate(22050)
    w.writeframes(out.tobytes())

with wave.open(str(OUT)) as w:
    dur = w.getnframes() / w.getframerate()
    print(json.dumps({
        "mono_sec": round(dur, 2),
        "mono_hz": w.getframerate(),
        "mono_bytes": OUT.stat().st_size,
        "mono_mb": round(OUT.stat().st_size / 1048576, 1),
        "full_bytes": SRC.stat().st_size,
        "full_mb": round(SRC.stat().st_size / 1048576, 1),
    }, ensure_ascii=False))
