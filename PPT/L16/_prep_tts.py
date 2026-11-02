# -*- coding: utf-8 -*-
import os, re, zipfile
from pptx import Presentation

P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
OUT = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\配音稿_W1"
os.makedirs(OUT, exist_ok=True)

# 用已交付的讲稿反推真实语速
zf = zipfile.ZipFile(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZa_声音模型与试讲_20261029.zip")
txt = zf.read("delivery/lectures/02_功能分析.txt").decode("utf-8")
chars = len(re.sub(r"\s", "", txt))
cpm_real = chars / (280.65 / 60)
print("交付实测量：02_功能分析 %d 字 / 280.65 秒 = %.1f 字/分" % (chars, cpm_real))

prs = Presentation(P)
rows = []
for i, s in enumerate(prs.slides, 1):
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    n = n.replace("\v", " ").replace("\x0b", " ")
    warn = ""
    if "--以下用新叙事替换---" in n:
        n = n.split("--以下用新叙事替换---")[0].strip()
        warn = "含旧编辑标记，已截断"
    title = ""
    for sh in s.shapes:
        if sh.has_text_frame and sh.name == "TextBox 2" and len(sh.text_frame.text) > 3:
            title = sh.text_frame.text.strip(); break
    if i == 1: title = "封面"
    safe = re.sub(r'[\\/:*?"<>|· 　]', "", title)[:12] or "页"
    fn = "%02d_%s.txt" % (i, safe)
    open(os.path.join(OUT, fn), "w", encoding="utf-8", newline="\n").write(n.strip() + "\n")
    rows.append((i, title, len(n.strip()), warn))

print("\n页  字数  预计秒  标题")
tot = 0
for i, t, c, w in rows:
    tot += c
    print(f"{i:2d}  {c:4d}  {c/cpm_real*60:6.1f}  {t[:26]}{'  ← ' + w if w else ''}")
print("\n合计 %d 字 → 预计 %.1f 分钟（按交付实测 %.1f 字/分）" % (tot, tot/cpm_real, cpm_real))

# 跑批脚本（在 Linux 机器上执行）
sh = """#!/usr/bin/env bash
# 在已配置好的 GPT-SoVITS 工作目录执行：
#   bash /media/wislab/Datasets/hyh_water/audio/tts.sh --input-dir <本目录> --output <输出目录>
# 生成 16 段，文件名与 txt 一致。
set -e
WORK=/media/wislab/Datasets/hyh_water/audio
bash "$WORK/tts.sh" --input-dir "<改成配音稿_W1的路径>" --output "$WORK/output/W1_矛盾15m" --speed 1.0
"""
open(os.path.join(OUT, "_run_batch.sh"), "w", encoding="utf-8", newline="\n").write(sh)
print("\n已写:", OUT)
for f in sorted(os.listdir(OUT))[:3]:
    print("  ", f)
print("    ... 共", len([f for f in os.listdir(OUT) if f.endswith('.txt')]), "个 txt")