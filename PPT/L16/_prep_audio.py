# -*- coding: utf-8 -*-
import os, re, json, subprocess, shutil
import imageio_ffmpeg
FF = imageio_ffmpeg.get_ffmpeg_exe()
SRC = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音"
WAV = os.path.join(SRC, "wav"); MP3 = os.path.join(SRC, "mp3")
os.makedirs(WAV, exist_ok=True); os.makedirs(MP3, exist_ok=True)

files = sorted([f for f in os.listdir(SRC) if f.lower().endswith((".m4a", ".wav", ".mp3")) and os.path.isfile(os.path.join(SRC, f))])
print("found:", len(files))

# 1) rename  ->  voice_Snn.<ext>
mp = {}
for i, f in enumerate(files):
    ext = os.path.splitext(f)[1].lower()
    new = "voice_S%02d%s" % (i, ext)
    if f != new:
        os.rename(os.path.join(SRC, f), os.path.join(SRC, new))
    mp[new] = f
json.dump(mp, open(os.path.join(SRC, "_renamed_map.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("renamed ->", sorted(mp)[:3], "...", sorted(mp)[-1])

def run(args):
    r = subprocess.run(args, capture_output=True, text=True, encoding="utf-8", errors="ignore")
    return r.returncode, (r.stderr or "")

def dur_of(p):
    rc, err = run([FF, "-hide_banner", "-i", p])
    m = re.search(r"Duration: (\d+):(\d+):(\d+\.\d+)", err)
    if not m: return None, err[-300:]
    h, mi, s = int(m.group(1)), int(m.group(2)), float(m.group(3))
    return h*3600 + mi*60 + s, err

meta = {}
for i in range(len(files)):
    src = os.path.join(SRC, "voice_S%02d.m4a" % i)
    if not os.path.isfile(src):
        cand = [f for f in os.listdir(SRC) if f.startswith("voice_S%02d." % i)]
        src = os.path.join(SRC, cand[0])
    d, err = dur_of(src)
    rc1, e1 = run([FF, "-y", "-loglevel", "error", "-i", src, "-ac", "1", "-ar", "44100", "-c:a", "pcm_s16le",
                   os.path.join(WAV, "voice_S%02d.wav" % i)])
    rc2, e2 = run([FF, "-y", "-loglevel", "error", "-i", src, "-ac", "1", "-ar", "44100", "-c:a", "libmp3lame", "-b:a", "128k",
                   os.path.join(MP3, "voice_S%02d.mp3" % i)])
    meta.setdefault("src", {})["S%02d" % i] = {"file": os.path.basename(src), "dur": round(d or 0, 2),
                                               "wav": rc1 == 0, "mp3": rc2 == 0}
    print(f"S{i:02d} {os.path.basename(src)}  {d:7.2f}s  wav={rc1==0} mp3={rc2==0}" + ("" if rc1 == 0 else " " + e1[-160:]))

# 2) cpm vs 讲稿
raw = open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_录音稿.txt", encoding="utf-8").read()
segs = {}
for blk in raw.split("### ")[1:]:
    head, _, body = blk.partition("\n")
    m = re.match(r"(S\d+)", head)
    if m: segs[m.group(1)] = len(re.sub(r"\s", "", body))
print("\n 段   讲稿字数  录音秒  实测字/分")
rows = []
for i in range(len(files)):
    k = "S%02d" % i
    d = meta["src"][k]["dur"]; c = segs.get(k, 0)
    cpm = c / d * 60 if d else 0
    rows.append((k, c, d, round(cpm, 1)))
    print(f" {k}  {c:8d}  {d:7.2f}  {cpm:7.1f}")
tc = sum(r[1] for r in rows); td = sum(r[2] for r in rows)
print(f"\n合计 {tc} 字 / {td/60:.2f} 分钟 → {tc/td*60:.1f} 字/分钟")
med = sorted(r[3] for r in rows)[len(rows)//2]
print("中位:", med)
print("偏离 >8% 的段:", [(r[0], r[3]) for r in rows if abs(r[3]-med)/med > 0.08])
json.dump(meta, open(os.path.join(SRC, "_audio_meta.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)