# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
import build_w2 as B
from pptx import Presentation
from pptx.util import Inches

SRC = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
OUT = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m_矩阵图版.pptx"
PNG = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\矛盾矩阵_示意.png"

prs = Presentation(SRC)
s = prs.slides[11]

card = [sh for sh in s.shapes if sh.name == "Rectangle 5"][0]
card._element.getparent().remove(card._element)      # 卡片下移为通栏

IW = 3.95
s.shapes.add_picture(PNG, Inches(8.55), Inches(2.30), width=Inches(IW),
                     height=Inches(round(IW * 1180 / 1680, 3)))
B.band(s, "矩阵给的是方向，不是答案——它不给公式、不给参数、不给阈值。", 5.42, size=20, height=0.72)
prs.save(OUT)

chk = Presentation(OUT)
print("已保存:", OUT, "|", round(os.path.getsize(OUT)/1024), "KB")
print("总页数:", len(chk.slides))
s12 = chk.slides[11]
print("\n第 12 页：")
for sh in s12.shapes:
    print("  [%s] %s L%.2f T%.2f W%.2f H%.2f" % (sh.shape_type, sh.name,
          sh.left/914400, sh.top/914400, sh.width/914400, sh.height/914400))
    if sh.has_text_frame and sh.text_frame.text.strip():
        print("        「" + sh.text_frame.text.strip().splitlines()[0] + "」")

# 原有内容是否完整
print("\n--- 完整性核对 ---")
print("notes 第12页字数:", len(s12.notes_slide.notes_text_frame.text))
print("原文件未被改动:",
      os.path.getsize(SRC), "bytes /", __import__("datetime").datetime.fromtimestamp(os.path.getmtime(SRC)).strftime("%m-%d %H:%M"))