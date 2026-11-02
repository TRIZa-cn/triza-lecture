# -*- coding: utf-8 -*-
import os, re, json, wave, array, subprocess, shutil, zipfile
import imageio_ffmpeg
FF = imageio_ffmpeg.get_ffmpeg_exe()
BASE = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音训练交接包"
SRC = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音\wav"
if os.path.isdir(BASE): shutil.rmtree(BASE)
for d in ["01_音频", "02_对照文字", "03_训练步骤", "04_校验"]:
    os.makedirs(os.path.join(BASE, d))

raw = open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_录音稿.txt", encoding="utf-8").read()
TEXT, TITLE = {}, {}
for blk in raw.split("### ")[1:]:
    head, _, body = blk.partition("\n")
    m = re.match(r"(S\d+)\s+(.*?)\s+\(", head)
    if m:
        TEXT[m.group(1)] = re.sub(r"\s+", " ", body).strip()
        TITLE[m.group(1)] = m.group(2)

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
for i in range(16):
    sid = "S%02d" % i
    src = os.path.join(SRC, "voice_" + sid + ".wav")
    a, sr = read(src)
    st, en = span(a, sr)
    s, e = max(0.0, st-0.15), min(len(a)/sr, en+0.30)
    dst = os.path.join(BASE, "01_音频", sid + ".wav")
    subprocess.run([FF, "-y", "-loglevel", "error", "-i", src, "-ss", "%.3f" % s, "-to", "%.3f" % e,
                    "-ac", "1", "-ar", "44100", "-c:a", "pcm_s16le", dst], check=True)
    d = e - s
    chars = len(re.sub(r"\s", "", TEXT[sid]))
    rows.append({"id": sid, "title": TITLE[sid], "chars": chars, "sec": round(d, 2),
                 "cpm": round(chars/d*60, 1), "bytes": os.path.getsize(dst)})

json.dump(rows, open(os.path.join(BASE, "04_校验", "_stats.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

# ---------- list.txt / labels.csv / 逐段文字 ----------
with open(os.path.join(BASE, "02_对照文字", "list.txt"), "w", encoding="utf-8") as f:
    for r in rows:
        f.write("01_音频/%s.wav|LWY|zh|%s\n" % (r["id"], TEXT[r["id"]]))
with open(os.path.join(BASE, "02_对照文字", "labels.csv"), "w", encoding="utf-8-sig") as f:
    f.write("文件名,秒,字数,字每分,标题,文字\n")
    for r in rows:
        t = TEXT[r["id"]].replace('"', '""')
        f.write('"%s.wav",%.2f,%d,%.1f,"%s","%s"\n' % (r["id"], r["sec"], r["chars"], r["cpm"], r["title"], t))
with open(os.path.join(BASE, "02_对照文字", "逐段文字.md"), "w", encoding="utf-8") as f:
    f.write("# 逐段对照文字（16 段，共 %d 字）\n\n" % sum(r["chars"] for r in rows))
    f.write("录音总时长 %.2f 分钟，实测语速 %.1f 字/分。校正 ASR 时以此为准。\n\n" % (
        sum(r["sec"] for r in rows)/60, sum(r["chars"] for r in rows)/sum(r["sec"] for r in rows)*60))
    for r in rows:
        f.write("## %s　%s\n\n" % (r["id"], r["title"]))
        f.write("`%s.wav` · %.1f 秒 · %d 字 · %.0f 字/分\n\n" % (r["id"], r["sec"], r["chars"], r["cpm"]))
        f.write(TEXT[r["id"]] + "\n\n")
print("音频 16 段, 文字与清单已生成")
for r in rows: print(" ", r["id"], r["sec"], "s", r["chars"], "字", r["cpm"], "字/分")