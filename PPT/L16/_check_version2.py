# -*- coding: utf-8 -*-
import os, re, wave, array
SRC = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音\wav"
raw = open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_录音稿.txt", encoding="utf-8").read()
segs = {}
for blk in raw.split("### ")[1:]:
    head, _, body = blk.partition("\n")
    m = re.match(r"(S\d+)", head)
    if m: segs[m.group(1)] = len(re.sub(r"\s", "", body))

def read(p):
    w = wave.open(p, "rb")
    n, sr, ch = w.getnframes(), w.getframerate(), w.getnchannels()
    a = array.array("h"); a.frombytes(w.readframes(n)); w.close()
    if ch > 1:
        a = array.array("h", [sum(a[i*ch:(i+1)*ch])//ch for i in range(len(a)//ch)])
    return a, sr

def span(a, sr, rel=0.04, frame=0.02):
    f = max(1, int(sr*frame)); n = len(a)//f
    peak = 0; env = []
    for i in range(n):
        seg = a[i*f:(i+1)*f]
        v = max(abs(x) for x in seg) if seg else 0
        env.append(v); peak = max(peak, v)
    thr = max(300, int(peak*rel))
    idx = [i for i, v in enumerate(env) if v > thr]
    if not idx: return 0.0, len(a)/sr
    return idx[0]*frame, (idx[-1]+1)*frame

rows = {}
for i in range(16):
    a, sr = read(os.path.join(SRC, "voice_S%02d.wav" % i))
    st, en = span(a, sr)
    full = len(a)/sr
    rows["S%02d" % i] = {"full": full, "speech": en-st, "lead": st, "tail": full-en, "sr": sr}

UN = ["S01","S03","S06","S07","S08","S09","S10","S11","S12"]
base = sorted(segs[k]/rows[k]["speech"]*60 for k in UN)
base_med = base[len(base)//2]
print(f"未改动 9 段语音语速：{base_med:.1f} 字/分  (min {base[0]:.1f}, max {base[-1]:.1f})")
OLD = {"S02": 271, "S04": 219, "S05": 208}
print("\n段   语音秒   旧稿预测   新稿预测   实测   判定（离谁近）")
for k in ["S02", "S04", "S05"]:
    sp = rows[k]["speech"]
    po, pn = OLD[k]/base_med*60, segs[k]/base_med*60
    v = "旧稿" if abs(sp-po) < abs(sp-pn) else "新稿"
    print(f"{k}  {sp:7.2f}  {po:9.2f}  {pn:9.2f}  {sp:7.2f}   {v}  (Δ旧{sp-po:+.1f}s Δ新{sp-pn:+.1f}s)")

print("\n全表：")
tc = sum(segs[k] for k in rows); ts = sum(rows[k]["speech"] for k in rows)
for k in sorted(rows):
    r = rows[k]
    print(f"  {k}  {segs[k]:4d}字  {r['speech']:6.2f}s  首{r['lead']:4.2f} 尾{r['tail']:4.2f}  "
          f"{segs[k]/r['speech']*60:6.1f}字/分  sr={r['sr']}")
print(f"\n合计 {tc} 字 / 语音 {ts/60:.2f} 分钟 → {tc/ts*60:.1f} 字/分钟")
print(f"整段合计 {sum(v['full'] for v in rows.values())/60:.2f} 分钟")