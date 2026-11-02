# -*- coding: utf-8 -*-
import os, re, glob
import numpy as np
import miniaudio

SRC = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音"
files = sorted([f for f in os.listdir(SRC) if f.lower().endswith((".wav", ".mp3", ".m4a", ".flac"))],
               key=lambda s: s)
print("files:", len(files))
for f in files:
    p = os.path.join(SRC, f)
    d = miniaudio.decode_file(p)
    dur = len(d.samples) / (d.nchannels * d.sample_rate)
    print(f"  {f:<40} {os.path.getsize(p)/1048576:6.2f} MB  {dur:7.2f}s  {d.sample_rate} Hz  {d.nchannels}ch")