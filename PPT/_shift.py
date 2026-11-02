# -*- coding: utf-8 -*-
import json, wave, math
import numpy as np
from pathlib import Path

BASE = Path(r"D:\doc\ZGCA\Lectures\TRIZ")
RAW = BASE / "_raw"
OUT = BASE / "_male"
F = 0.72
N = 1024
HA = 512
HS = int(round(HA * F))

def read(p):
    with wave.open(str(p), "rb") as w:
        sr = w.getframerate(); ch = w.getnchannels()
        raw = w.readframes(w.getnframes())
    x = np.frombuffer(raw, dtype="<i2").astype(np.float32) / 32768.0
    if ch > 1:
        x = x.reshape(-1, ch).mean(axis=1)
    return x, sr

def write(p, x, sr):
    y = np.clip(x, -1.0, 1.0)
    with wave.open(str(p), "wb") as w:
        w.setnchannels(1); w.setsampwidth(2); w.setframerate(sr)
        w.writeframes((y * 32767.0).astype("<i2").tobytes())

def resample(x, factor):
    n = int(len(x) / factor)
    idx = np.arange(n, dtype=np.float64) * factor
    i0 = np.floor(idx).astype(np.int64)
    np.clip(i0, 0, len(x) - 2, out=i0)
    frac = (idx - i0).astype(np.float32)
    return x[i0] * (1.0 - frac) + x[i0 + 1] * frac

def ola_compress(x, factor):
    win = np.hanning(N).astype(np.float32)
    out_len = int(len(x) * factor) + N + 1
    out = np.zeros(out_len, dtype=np.float32)
    norm = np.zeros(out_len, dtype=np.float32)
    pos_a = 0; pos_s = 0
    while pos_a + N <= len(x):
        m = min(N, out_len - pos_s)
        if m <= 0:
            break
        out[pos_s:pos_s+m] += x[pos_a:pos_a+m] * win[:m]
        norm[pos_s:pos_s+m] += win[:m]
        pos_a += HA; pos_s += HS
    norm[norm < 1e-6] = 1.0
    return out / norm

def pitch_down(x, f):
    y = resample(x, f)
    y = ola_compress(y, f)
    peak = float(np.max(np.abs(y))) or 1.0
    y = y * (0.92 / peak)
    fade = int(0.01 * 16000)
    if len(y) > 2 * fade:
        y[:fade] *= np.linspace(0.0, 1.0, fade)
        y[-fade:] *= np.linspace(1.0, 0.0, fade)
    return y

meta = json.loads((BASE / "_notes_dump.json").read_text(encoding="utf-8"))
files = sorted(RAW.glob("s*.wav"))
durs = []
combined = []
sr_out = 16000
for i, p in enumerate(files):
    x, sr = read(p)
    sr_out = sr
    y = pitch_down(x, F)
    write(OUT / p.name, y, sr)
    d = len(y) / sr
    durs.append(round(d, 2))
    combined.append(y)
    gap = np.zeros(int(0.9 * sr), dtype=np.float32)
    combined.append(gap)

cat = np.concatenate(combined) if combined else np.zeros(1, dtype=np.float32)
write(BASE / "TRIZa_talk_male.wav", cat, sr_out)
(BASE / "_male_durs.json").write_text(json.dumps({"factor": F, "durations": durs,
    "total": round(len(cat)/sr_out, 1), "slides": len(files)}, indent=1), encoding="utf-8")
print("files", len(files), "total_min", round(len(cat)/sr_out/60, 2))
print("durations", durs)
