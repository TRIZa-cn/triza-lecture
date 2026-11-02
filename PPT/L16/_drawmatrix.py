# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import os

PAPER = (250, 249, 246); LINE = (214, 210, 200); INK = (26, 29, 35)
GRAY  = (107, 114, 128); RED  = (176, 58, 43); WHITE = (255, 255, 255)
YA  = r"C:\Windows\Fonts\msyh.ttc"
YAB = r"C:\Windows\Fonts\msyhbd.ttc"
def F(sz, bold=False): return ImageFont.truetype(YAB if bold else YA, sz)

W, H = 1680, 1180
img = Image.new("RGB", (W, H), PAPER)
d = ImageDraw.Draw(img)

N, CELL = 39, 22
x0, y0 = 200, 170
x1, y1 = x0 + N*CELL, y0 + N*CELL          # 200..1058 / 170..1028

# 网格
for i in range(N+1):
    c = LINE
    d.line([(x0 + i*CELL, y0), (x0 + i*CELL, y1)], fill=c, width=1)
    d.line([(x0, y0 + i*CELL), (x1, y0 + i*CELL)], fill=c, width=1)
d.rectangle([x0, y0, x1, y1], outline=(150, 145, 135), width=3)

# 高亮格 + 圈
col, row = 8, 13
rx, ry = x0 + col*CELL, y0 + row*CELL
d.rectangle([rx, ry, rx+CELL, ry+CELL], fill=RED)
cx, cy = rx + CELL//2, ry + CELL//2
d.ellipse([cx-36, cy-36, cx+36, cy+36], outline=RED, width=5)

# 引线
d.line([(cx+36, cy), (1105, cy)], fill=RED, width=4)
d.ellipse([1105-7, cy-7, 1105+7, cy+7], fill=RED)

# 右侧标注卡
bx0, by0, bx1, by1 = 1112, cy-165, 1655, cy+165
d.rounded_rectangle([bx0, by0, bx1, by1], radius=14, fill=WHITE, outline=(210, 205, 195), width=3)
d.text((bx0+34, by0+28), "改善参数 9", font=F(42, True), fill=INK)
d.text((bx0+34, by0+82), "×  恶化参数 14", font=F(42, True), fill=INK)
d.line([(bx0+34, by0+150), (bx1-34, by0+150)], fill=(220, 216, 208), width=2)
d.text((bx0+34, by0+168), "→  原理", font=F(40), fill=GRAY)
d.text((bx0+34, by0+222), "35 · 10 · 34", font=F(52, True), fill=RED)

# 轴标
d.text((x0 - 4, y0 - 66), "改善参数（横）→", font=F(40), fill=GRAY)
lbl = Image.new("RGBA", (430, 66), (0, 0, 0, 0))
ImageDraw.Draw(lbl).text((0, 0), "← 恶化参数（纵）", font=F(40), fill=GRAY)
img.paste(lbl.rotate(90, expand=True), (58, (y0+y1)//2 - 215), lbl.rotate(90, expand=True))

# 角标
d.text((x1 - 92, y1 + 26), "39 × 39", font=F(38, True), fill=GRAY)
d.text((x0, y1 + 26), "每个格子给 3–4 条原理编号", font=F(38), fill=GRAY)

out = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\矛盾矩阵_示意.png"
os.makedirs(os.path.dirname(out), exist_ok=True)
img.save(out, "PNG")
print("saved:", out, img.size)

# ---------- 装进第 12 页 ----------
import sys, shutil
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
import build_w2 as B
from pptx import Presentation
from pptx.util import Inches, Pt

P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
BAK = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m_备份-加矩阵图前.pptx"
if not os.path.exists(BAK): shutil.copyfile(P, BAK)
prs = Presentation(P)
s = prs.slides[11]

card = [sh for sh in s.shapes if sh.name == "Rectangle 5"][0]
card._element.getparent().remove(card._element)          # 卡片下移成通栏

IW = 3.95
s.shapes.add_picture(out, Inches(8.55), Inches(2.30), width=Inches(IW),
                     height=Inches(round(IW*1180/1680, 3)))
B.band(s, "矩阵给的是方向，不是答案——它不给公式、不给参数、不给阈值。", 5.42, size=20, height=0.72)
prs.save(P)

chk = Presentation(P)
s12 = chk.slides[11]
print("\n第 12 页现有形状：")
for sh in s12.shapes:
    print("  [%s] %s L%.2f T%.2f W%.2f H%.2f" % (sh.shape_type, sh.name,
          sh.left/914400, sh.top/914400, sh.width/914400, sh.height/914400))
print("\n总页数:", len(chk.slides))