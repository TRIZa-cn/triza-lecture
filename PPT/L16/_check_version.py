# -*- coding: utf-8 -*-
import os, re, json
import numpy as np
import miniaudio

SRC = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音\wav"
raw = open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_录音稿.txt", encoding="utf-8").read()
segs = {}
for blk in raw.split("### ")[1:]:
    head, _, body = blk.partition("\n")
    m = re.match(r"(S\d+)", head)
    if m: segs[m.group(1)] = len(re.sub(r"\s", "", body))

def span(a, sr, rel=0.04, frame=0.02):
    f = max(1, int(sr*frame)); n = len(a)//f
    e = np.abs(a[:n*f].reshape(n, f)).max(axis=1).astype(float)
    thr = max(300.0, e.max()*rel); idx = np.where(e > thr)[0]
    return (idx[0]*frame, (idx[-1]+1)*frame) if len(idx) else (0.0, len(a)/sr)

rows = {}
for i in range(16):
    p = os.path.join(SRC, "voice_S%02d.wav" % i)
    d = miniaudio.decode_file(p)
    a = np.asarray(d.samples, dtype=np.int16)
    if d.nchannels > 1: a = a.reshape(-1, d.nchannels).mean(axis=1).astype(np.int16)
    st, en = span(a, d.sample_rate)
    rows["S%02d" % i] = {"full": len(a)/d.sample_rate, "speech": en-st, "lead": st, "tail": len(a)/d.sample_rate-en}

UNCHANGED = ["S01","S03","S06","S07","S08","S09","S10","S11","S12"]
cpm_u = [segs[k]/rows[k]["speech"]*60 for k in UNCHANGED]
base = float(np.median(cpm_u))
print(f"未改动段的语音语速中位: {base:.1f} 字/分  (范围 {min(cpm_u):.1f}~{max(cpm_u):.1f})")

OLD = {"S02": 271, "S04": 219, "S05": 208}   # 三页改前的字数
print("\n段   实测语音秒  按旧稿预测  按新稿预测   实测字/分  判定")
for k in ["S02", "S04", "S05"]:
    sp = rows[k]["speech"]
    pred_old = OLD[k]/base*60
    pred_new = segs[k]/base*60
    verdict = "旧稿" if abs(sp-pred_old) < abs(sp-pred_new) else "新稿"
    print(f"{k}  {sp:9.2f}  {pred_old:9.2f}  {pred_new:9.2f}  {segs[k]/sp*60:9.1f}  {verdict}"
          f"  (偏差 旧{sp-pred_old:+.1f}s 新{sp-pred_new:+.1f}s)")

print("\n全表（语音口径）：")
tc = sum(segs[k] for k in rows); ts = sum(rows[k]["speech"] for k in rows)
for k in sorted(rows):
    print(f"  {k}  {segs[k]:4d}字  {rows[k]['speech']:6.2f}s  首{rows[k]['lead']:4.2f} 尾{rows[k]['tail']:4.2f}  {segs[k]/rows[k]['speech']*60:6.1f} 字/分")
print(f"\n合计 {tc} 字 / 语音 {ts/60:.2f} 分钟 → {tc/ts*60:.1f} 字/分钟（含整段 {sum(v['full'] for v in rows.values())/60:.2f} 分钟）")