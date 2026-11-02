# -*- coding: utf-8 -*-
import os
from pptx import Presentation
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
prs = Presentation(P)
FOOT = "TRIZ 创新构思法实战研讨班  ·  第 1 次 Workshop"
small, tot_notes, foot_ok = [], 0, 0
for i, s in enumerate(prs.slides, 1):
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    tot_notes += len(n)
    has_foot = False
    for sh in s.shapes:
        frames = []
        if sh.has_text_frame: frames.append(sh.text_frame)
        if sh.has_table:
            for row in sh.table.rows:
                for c in row.cells: frames.append(c.text_frame)
        for tf in frames:
            for p in tf.paragraphs:
                for r in p.runs:
                    if not r.text.strip(): continue
                    sz = r.font.size.pt if r.font.size else None
                    if sz is None or sz < 20:
                        small.append((i, sz, r.text[:40]))
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == FOOT:
            has_foot = True
    foot_ok += has_foot
    print(f"  第{i:2d}页 页脚={'✓' if has_foot else '✗'} notes={len(n):3d}字")
print("\n页脚正确页数:", foot_ok, "/", len(prs.slides))
print("notes 合计:", tot_notes, "字 →", round(tot_notes/270, 2), "分钟 @270 |", round(tot_notes/250, 2), "分钟 @250")
print("\n<20pt 的文字（应只有 kicker 14 / 页脚 12 / 页码 12 / 封面署名 12）:")
from collections import Counter
c = Counter()
for i, sz, t in small:
    c[(sz, t[:14])] += 1
for (sz, t), k in sorted(c.items(), key=lambda x: (x[0][0] is None, x[0][0])):
    print(f"  {sz}pt x{k}  「{t}」")