# -*- coding: utf-8 -*-
import os, json, wave, array, subprocess, re
import imageio_ffmpeg
FF = imageio_ffmpeg.get_ffmpeg_exe()
SRC = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音"
WAV = os.path.join(SRC, "wav"); OUT = os.path.join(SRC, "slide_mp3")
os.makedirs(OUT, exist_ok=True)

def read(p):
    w = wave.open(p, "rb"); n, sr, ch = w.getnframes(), w.getframerate(), w.getnchannels()
    a = array.array("h"); a.frombytes(w.readframes(n)); w.close()
    if ch > 1: a = array.array("h", [sum(a[i*ch:(i+1)*ch])//ch for i in range(len(a)//ch)])
    return a, sr

def span(a, sr, rel=0.04, frame=0.02):
    f = max(1, int(sr*frame)); n = len(a)//f; peak = 0; env = []
    for i in range(n):
        seg = a[i*f:(i+1)*f]; v = max(abs(x) for x in seg) if seg else 0
        env.append(v); peak = max(peak, v)
    thr = max(300, int(peak*rel)); idx = [i for i, v in enumerate(env) if v > thr]
    return (idx[0]*frame, (idx[-1]+1)*frame) if idx else (0.0, len(a)/sr)

rows = []
for i in range(1, 13):          # S01..S12 -> slide 1..12
    src = os.path.join(WAV, "voice_S%02d.wav" % i)
    a, sr = read(src); st, en = span(a, sr)
    start = max(0.0, st - 0.20); end = min(len(a)/sr, en + 0.45)
    dst = os.path.join(OUT, "A%02d.mp3" % i)
    subprocess.run([FF, "-y", "-loglevel", "error", "-i", src, "-ss", "%.3f" % start,
                    "-to", "%.3f" % end, "-ac", "1", "-ar", "44100",
                    "-c:a", "libmp3lame", "-b:a", "96k", dst], check=True)
    rc = subprocess.run([FF, "-hide_banner", "-i", dst], capture_output=True, text=True, encoding="utf-8", errors="ignore")
    m = re.search(r"Duration: (\d+):(\d+):(\d+\.\d+)", rc.stderr)
    d = int(m.group(1))*3600 + int(m.group(2))*60 + float(m.group(3))
    rows.append({"slide": i, "file": dst, "dur": round(d, 2),
                 "trim": [round(start, 2), round(end, 2)], "bytes": os.path.getsize(dst)})
    print(f"slide {i:2d}  {d:6.2f}s  trim {start:5.2f}~{end:6.2f}  {os.path.getsize(dst)/1024:6.0f} KB")

json.dump({"rows": rows, "total": round(sum(r["dur"] for r in rows), 2),
           "min": round(sum(r["dur"] for r in rows)/60, 2)},
          open(os.path.join(SRC, "_slide_audio.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("total", round(sum(r['dur'] for r in rows), 1), "s =", round(sum(r['dur'] for r in rows)/60, 2), "min")