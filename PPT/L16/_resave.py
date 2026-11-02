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
[sh for sh in s.shapes if sh.name == "Rectangle 5"][0]._element.getparent().remove(
    [sh for sh in s.shapes if sh.name == "Rectangle 5"][0]._element)
IW = 3.95
s.shapes.add_picture(PNG, Inches(8.55), Inches(2.30), width=Inches(IW), height=Inches(round(IW*1180/1680,3)))
B.band(s, "矩阵给的是方向，不是答案——它不给公式、不给参数、不给阈值。", 5.42, size=20, height=0.72)
prs.save(OUT)
c = Presentation(OUT)
print("重存完成:", round(os.path.getsize(OUT)/1024), "KB |", len(c.slides), "页 | 第12页图片:",
      [sh.name for sh in c.slides[11].shapes if sh.shape_type == 13])