# -*- coding: utf-8 -*-
"""TRIZa-VoiceSpeed.py —— 录音语速测量（按段号对齐讲稿）"""
import argparse, json, os, re, sys
import numpy as np
import miniaudio

EXT = (".wav", ".mp3", ".m4a", ".flac", ".aac", ".ogg")


def parse_script(path):
    txt = open(path, encoding="utf-8").read()
    segs, cur = [], None
    for line in txt.splitlines():
        m = re.match(r"^###\s+(S\d+)\s+(.*?)\s+\((\d+)字,\s*([^)]+)\)\s*$", line)
        if m:
            if cur:
                segs.append(cur)
            cur = {"id": m.group(1), "title": m.group(2), "chars": int(m.group(3)),
                   "file": m.group(4).strip(), "text": ""}
        elif cur is not None:
            cur["text"] += line + "\n"
    if cur:
        segs.append(cur)
    for s in segs:
        s["chars"] = len(re.sub(r"\s", "", s["text"])) or s["chars"]
    return segs


def find_audio(folder, name):
    stem = os.path.splitext(name)[0]
    for e in EXT:
        p = os.path.join(folder, stem + e)
        if os.path.isfile(p):
            return p
    for f in os.listdir(folder):
        if os.path.splitext(f)[0].lower() == stem.lower() and f.lower().endswith(EXT):
            return os.path.join(folder, f)
    return None


def load(path):
    d = miniaudio.decode_file(path, output_format=miniaudio.SampleFormat.SIGNED16)
    a = np.asarray(d.samples, dtype=np.int16)
    if d.nchannels > 1:
        a = a.reshape(-1, d.nchannels).mean(axis=1).astype(np.int16)
    return a, d.sample_rate


def speech_span(a, sr, rel=0.03, frame=0.02):
    """返回 (start_sec, end_sec)：首尾超过阈值的语音区间。"""
    f = max(1, int(sr * frame))
    n = len(a) // f
    if n < 3:
        return 0.0, len(a) / sr
    e = np.abs(a[:n * f].reshape(n, f)).max(axis=1).astype(np.float64)
    thr = max(300.0, e.max() * rel)
    idx = np.where(e > thr)[0]
    if len(idx) == 0:
        return 0.0, len(a) / sr
    return idx[0] * frame, (idx[-1] + 1) * frame


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True, help="录音文件夹")
    ap.add_argument("--script", required=True, help="_录音稿.txt")
    ap.add_argument("--raw", action="store_true", help="按整段时长算（含首尾静音）")
    a = ap.parse_args()

    segs = parse_script(a.script)
    rows, miss = [], []
    tot_c = tot_s = tot_c2 = tot_s2 = 0.0
    for s in segs:
        p = find_audio(a.dir, s["file"])
        if not p:
            miss.append(s["id"])
            continue
        data, sr = load(p)
        dur = len(data) / sr
        st, en = speech_span(data, sr)
        eff = max(0.05, en - st)
        use = dur if a.raw else eff
        cpm = s["chars"] / use * 60
        rows.append({"id": s["id"], "chars": s["chars"], "sec": round(use, 2),
                     "dur": round(dur, 2), "lead": round(st, 2), "tail": round(dur - en, 2),
                     "cpm": round(cpm, 1), "file": os.path.basename(p)})
        tot_c += s["chars"]; tot_s += use
        tot_c2 += s["chars"]; tot_s2 += dur

    if not rows:
        print(json.dumps({"error": "没有找到任何音频", "missing": miss}, ensure_ascii=False))
        return 2

    cpms = [r["cpm"] for r in rows]
    out = {
        "segments_found": len(rows), "segments_missing": miss,
        "chars": int(tot_c),
        "speech_sec": round(tot_s, 1), "total_sec": round(tot_s2, 1),
        "cpm": round(tot_c / tot_s * 60, 1),
        "cpm_with_silence": round(tot_c2 / tot_s2 * 60, 1) if tot_s2 else None,
        "cpm_min": min(cpms), "cpm_max": max(cpms),
        "cpm_median": round(float(np.median(cpms)), 1),
        "target": [265, 275],
        "verdict": "达标" if 265 <= tot_c / tot_s * 60 <= 275 else
                   ("偏慢" if tot_c / tot_s * 60 < 265 else "偏快"),
    }
    print(json.dumps(out, ensure_ascii=False, indent=1))
    print()
    print(f"{'段':<5}{'字数':>5}{'语音秒':>8}{'整段秒':>8}{'首静音':>7}{'尾静音':>7}{'字/分':>8}")
    for r in rows:
        print(f"{r['id']:<5}{r['chars']:>5}{r['sec']:>8.1f}{r['dur']:>8.1f}"
              f"{r['lead']:>7.2f}{r['tail']:>7.2f}{r['cpm']:>8.1f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())