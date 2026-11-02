# -*- coding: utf-8 -*-
import subprocess, re, glob, os
import imageio_ffmpeg
FF = imageio_ffmpeg.get_ffmpeg_exe()
D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音20261030\_训练WAV"
files = sorted(glob.glob(D + r"\p*.wav"))
print("文件  mean_dB  max_dB  削波(>0dB采样)  提示")
bad = []
for p in files:
    r = subprocess.run([FF, "-hide_banner", "-i", p, "-af", "volumedetect", "-f", "null", "-"],
                       capture_output=True, text=True, errors="ignore")
    mean = re.search(r"mean_volume:\s*(-?[\d.]+) dB", r.stderr)
    mx = re.search(r"max_volume:\s*(-?[\d.]+) dB", r.stderr)
    m = float(mean.group(1)) if mean else 0
    x = float(mx.group(1)) if mx else 0
    note = ""
    if x >= -0.1: note = "顶到 0dB，可能削波"
    elif m < -30: note = "偏小"
    elif m > -12: note = "偏大"
    if note: bad.append((os.path.basename(p), note))
    print("%-9s %7.1f  %7.1f  %s" % (os.path.basename(p), m, x, note or "正常"))

# 整体响度
r = subprocess.run([FF, "-hide_banner", "-i", files[0], "-af", "loudnorm=print_format=summary", "-f", "null", "-"],
                   capture_output=True, text=True, errors="ignore")
print("\n第 1 页 loudnorm 摘要:")
for line in r.stderr.splitlines():
    if any(k in line for k in ("Input Integrated", "Input True Peak", "Input LRA", "Input Threshold")):
        print("  " + line.strip())

# 声道是否真立体声
r = subprocess.run([FF, "-hide_banner", "-i", r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音20261030\20261030_135841.m4a",
                    "-af", "astats=metadata=1:reset=0", "-f", "null", "-"],
                   capture_output=True, text=True, errors="ignore")
ch = re.findall(r"Channel:\s*(\d+)", r.stderr)
print("\n源文件声道:", sorted(set(ch)) or "未解析")
print("\n需要关注的条目:", bad if bad else "无")