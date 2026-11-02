# -*- coding: utf-8 -*-
import json
from PIL import Image, ImageDraw, ImageFont

D = json.load(open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\contradiction_matrix_raw.json", encoding="utf-8"))
CELLS = {tuple(map(int, k.split(","))): v for k, v in D["cells"].items()}

CN = {1:"运动物体重量",2:"静止物体重量",3:"运动物体长度",4:"静止物体长度",5:"运动物体面积",6:"静止物体面积",
7:"运动物体体积",8:"静止物体体积",9:"速度",10:"力",11:"应力或压力",12:"形状",13:"结构稳定性",14:"强度",
15:"运动物体作用时间",16:"静止物体作用时间",17:"温度",18:"光照强度",19:"运动物体能量",20:"静止物体能量",
21:"功率",22:"能量损失",23:"物质损失",24:"信息损失",25:"时间损失",26:"物质的量",27:"可靠性",28:"测量精度",
29:"制造精度",30:"作用于物体的有害因素",31:"物体产生的有害因素",32:"可制造性",33:"操作流程方便性",34:"可维修性",
35:"适应性、通用性",36:"系统复杂性",37:"检测难度",38:"自动化程度",39:"生产率"}

PAPER=(250,249,246); INK=(26,29,35); RED=(176,58,43); GRAY=(115,110,102)
GRID=(206,201,189); DIAG=(230,226,215); WHITE=(255,255,255)
def F(s,b=False):
    return ImageFont.truetype(r"C:\Windows\Fonts\msyhbd.ttc" if b else r"C:\Windows\Fonts\msyh.ttc", s)

N, C = 39, 62
LW, TH = 205, 230
X0, Y0 = LW, TH
W = LW + N*C + 26
H = TH + N*C + 76

img = Image.new("RGB", (W, H), PAPER)
d = ImageDraw.Draw(img)
fnum, fcell, fname = F(15, True), F(14), F(15)

# 底与对角灰
d.rectangle([X0, Y0, X0+N*C, Y0+N*C], fill=WHITE)
for i in range(N):
    d.rectangle([X0+i*C, Y0+i*C, X0+(i+1)*C, Y0+(i+1)*C], fill=DIAG)
for i in range(N+1):
    d.line([(X0+i*C, Y0), (X0+i*C, Y0+N*C)], fill=GRID, width=1)
    d.line([(X0, Y0+i*C), (X0+N*C, Y0+i*C)], fill=GRID, width=1)
d.rectangle([X0, Y0, X0+N*C, Y0+N*C], outline=(135,129,117), width=3)

# 顶部列号 + 竖排名（紧贴网格）
for c in range(1, N+1):
    cx = X0 + (c-1)*C + C//2
    t = str(c); d.text((cx - d.textlength(t, font=fnum)/2, TH - 214), t, font=fnum, fill=GRAY)
    nm = CN[c]
    tmp = Image.new("RGBA", (int(d.textlength(nm, font=fname))+6, 21), (0,0,0,0))
    ImageDraw.Draw(tmp).text((2, 0), nm, font=fname, fill=INK)
    rot = tmp.rotate(90, expand=True)
    img.paste(rot, (cx - rot.width//2, TH - 6 - rot.height), rot)

# 左侧行号 + 名（右对齐）
for r in range(1, N+1):
    cy = Y0 + (r-1)*C + C//2
    d.text((10, cy - 12), "%2d" % r, font=fnum, fill=GRAY)
    nm = CN[r]
    d.text((LW - 12 - d.textlength(nm, font=fname), cy - 12), nm, font=fname, fill=INK)

# 单元格：固定槽位，左对齐
SX = (8, 34)          # 两列槽位的 x 偏移
for (r, c), plist in CELLS.items():
    x = X0 + (c-1)*C; y = Y0 + (r-1)*C
    items = [str(p) for p in plist[:4]]
    if len(items) <= 2:
        ty = y + C//2 - 10
        for i, t in enumerate(items):
            d.text((x + SX[i], ty), t, font=fcell, fill=RED if i == 0 else INK)
    else:
        for i, t in enumerate(items):
            d.text((x + SX[i % 2], y + 9 + (i // 2)*18), t, font=fcell, fill=RED if i == 0 else INK)

d.text((12, TH - 210), "列 = 被恶化", font=F(16, True), fill=GRAY)
d.text((12, TH - 186), "行 = 要改善", font=F(16, True), fill=GRAY)
d.text((LW, H - 50),
       "Altshuller 39×39 矛盾矩阵　格内为发明原理编号，红色为第一推荐　灰色为对角格（同一参数，按惯例留空）",
       font=F(22), fill=GRAY)

out = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\矛盾矩阵_39x39_真表.png"
img.save(out, "PNG")
print("已生成:", out, img.size, round(__import__("os").path.getsize(out)/1024), "KB")
img.crop((0, 0, 200+C*6, 230+C*7)).save(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\_crop_tl.png")
img.crop((205+C*8, 230+C*12, 205+C*16, 230+C*19)).save(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\_crop_mid.png")
print("crop ok")