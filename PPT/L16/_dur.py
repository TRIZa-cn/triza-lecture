# -*- coding: utf-8 -*-
import struct, os, glob
def mp4_duration(path):
    with open(path, "rb") as f:
        data = f.read(2_000_000)
    # 找 moov/mvhd
    i = data.find(b"mvhd")
    if i < 0: return None
    ver = data[i+4]
    if ver == 0:
        ts, dur = struct.unpack(">II", data[i+16:i+24])
    else:
        ts, dur = struct.unpack(">IQ", data[i+24:i+36])
    return dur / ts if ts else None

D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音20261030"
fs = sorted(glob.glob(os.path.join(D, "*.m4a")))
tot = 0
print("序号  文件                              时长     累计")
for n, p in enumerate(fs, 1):
    d = mp4_duration(p) or 0
    tot += d
    print("%2d   %-32s %5.1fs  %6.1fs" % (n, os.path.basename(p), d, tot))
print("\n合计 %.1f 秒 = %.2f 分钟" % (tot, tot/60))
print("平均每页 %.1f 秒" % (tot/len(fs)))