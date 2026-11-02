# -*- coding: utf-8 -*-
import os, glob, subprocess, json, re, csv
import imageio_ffmpeg
FF = imageio_ffmpeg.get_ffmpeg_exe()

SRC = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音20261030"
WAV = os.path.join(SRC, "_训练WAV"); MP3 = os.path.join(SRC, "_嵌入MP3")
os.makedirs(WAV, exist_ok=True); os.makedirs(MP3, exist_ok=True)

files = sorted(glob.glob(os.path.join(SRC, "*.m4a")))
assert len(files) == 18, len(files)

def probe(p):
    r = subprocess.run([FF, "-hide_banner", "-i", p], capture_output=True, text=True, errors="ignore")
    s = r.stderr
    m = re.search(r"Audio:\s*([\w.]+)[^,]*,\s*(\d+)\s*Hz,\s*([^,]+),", s)
    d = re.search(r"Duration:\s*(\d+):(\d+):([\d.]+)", s)
    dur = 0
    if d: dur = int(d.group(1))*3600 + int(d.group(2))*60 + float(d.group(3))
    return (m.group(1) if m else "?", int(m.group(2)) if m else 0,
            (m.group(3).strip() if m else "?"), dur, len(open(p,"rb").read()))

def run(args):
    r = subprocess.run(args, capture_output=True, text=True, errors="ignore")
    if r.returncode != 0:
        raise RuntimeError(r.stderr[-800:])
    return r

rows = []
print("序 源文件                 编码      采样率  声道  时长      WAV时长   MP3")
for n, p in enumerate(files, 1):
    codec, sr, ch, dur, size = probe(p)
    base = "p%02d" % n
    w = os.path.join(WAV, base + ".wav")
    m = os.path.join(MP3, base + ".mp3")
    # 训练用：32k 单声道 16bit，仅掐头去尾静音（-50dB）
    run([FF, "-y", "-hide_banner", "-loglevel", "error", "-i", p,
         "-ac", "1", "-ar", "32000", "-c:a", "pcm_s16le",
         "-af", "silenceremove=start_periods=1:start_threshold=-50dB:detection=peak,"
                "areverse,silenceremove=start_periods=1:start_threshold=-50dB:detection=peak,areverse",
         w])
    # 演示用：44.1k 单声道 128k mp3
    run([FF, "-y", "-hide_banner", "-loglevel", "error", "-i", p,
         "-ac", "1", "-ar", "44100", "-c:a", "libmp3lame", "-b:a", "128k", m])
    _, _, _, wdur, _ = probe(w)
    rows.append((n, os.path.basename(p), base, dur, wdur))
    print("%2d %-22s %-9s %5d  %-5s %6.1f   %6.1f  %6.0fKB" % (
        n, os.path.basename(p), codec, sr, ch, dur, wdur, os.path.getsize(m)/1024))

tw = sum(r[3] for r in rows); twa = sum(r[4] for r in rows)
print("\n源总时长 %.1fs (%.2f 分)｜WAV 总时长 %.1fs (%.2f 分)" % (tw, tw/60, twa, twa/60))
viol = [r for r in rows if abs(r[3]-r[4]) > 0.8]
print("掐静音后时长偏差 >0.8s 的条目:", viol if viol else "无")
print("WAV 目录大小 %.1f MB｜MP3 目录大小 %.1f MB" % (
    sum(os.path.getsize(f) for f in glob.glob(WAV+"/*.wav"))/1048576,
    sum(os.path.getsize(f) for f in glob.glob(MP3+"/*.mp3"))/1048576))
with open(os.path.join(SRC, "_音频清单.csv"), "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["页", "源文件", "输出编号", "源时长秒", "WAV时长秒"])
    for r in rows: w.writerow([r[0], r[1], r[2], round(r[3],2), round(r[4],2)])
print("清单已写: _音频清单.csv")