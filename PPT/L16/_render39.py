# -*- coding: utf-8 -*-
import json, re
from PIL import Image, ImageDraw, ImageFont

D = json.load(open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\contradiction_matrix_raw.json", encoding="utf-8"))
CELLS = {tuple(map(int, k.split(","))): v for k, v in D["cells"].items()}

CN = {1:"运动物体重量",2:"静止物体重量",3:"运动物体长度",4:"静止物体长度",5:"运动物体面积",6:"静止物体面积",
7:"运动物体体积",8:"静止物体体积",9:"速度",10:"力",11:"应力或压力",12:"形状",13:"结构稳定性",14:"强度",
15:"运动物体作用时间",16:"静止物体作用时间",17:"温度",18:"光照强度",19:"运动物体能量",20:"静止物体能量",
21:"功率",22:"能量损失",23:"物质损失",24:"信息损失",25:"时间损失",26:"物质的量",27:"可靠性",28:"测量精度",
29:"制造精度",30:"作用于物体的有害因素",31:"物体产生的有害因素",32:"可制造性",33:"操作流程方便性",34:"可维修性",
35:"适应性、通用性",36:"系统复杂性",37:"检测难度",38:"自动化程度",39:"生产率"}

PAPER=(250,249,246); INK=(26,29,35); RED=(176,58,43); GRAY=(120,116,108)
GRID=(201,196,184); DIAG=(232,228,218); WHITE=(255,255,255)
YA=r"C:\Windows\Fonts\msyh.ttc"; YAB=r"C:\Windows\Fonts\msyhbd.ttc"
def F(s,b=False): return ImageFont.truetype(YAB if b else YA, s)

N, C = 39, 56
LW, TH = 200, 215
X0, Y0 = LW, TH
W = LW + N*C + 24
H = TH + N*C + 72

img = Image.new("RGB", (W, H), PAPER)
d = ImageDraw.Draw(img)

# 网格底
d.rectangle([X0, Y0, X0+N*C, Y0+N*C], fill=WHITE, outline=GRID, width=2)
for r in range(1, N+1):
    for c in range(1, N+1):
        x = X0 + (c-1)*C; y = Y0 + (r-1)*C
        if r == c:
            d.rectangle([x, y, x+C, y+C], fill=DIAG)
# 网格线
for i in range(N+1):
    d.line([(X0+i*C, Y0), (X0+i*C, Y0+N*C)], fill=GRID, width=1)
    d.line([(X0, Y0+i*C), (X0+N*C, Y0+i*C)], fill=GRID, width=1)
d.rectangle([X0, Y0, X0+N*C, Y0+N*C], outline=(140,134,122), width=3)

fnum, fcell, fname = F(14, True), F(13), F(14)

# 顶部：列号 + 竖排参数名
for c in range(1, N+1):
    cx = X0 + (c-1)*C + C//2
    t = str(c)
    tw = d.textlength(t, font=fnum)
    d.text((cx - tw/2, TH - 200), t, font=fnum, fill=GRAY)
    nm = CN[c]
    tmp = Image.new("RGBA", (int(d.textlength(nm, font=fname))+8, 22), (0,0,0,0))
    ImageDraw.Draw(tmp).text((2, 0), nm, font=fname, fill=INK)
    rot = tmp.rotate(90, expand=True)
    img.paste(rot, (cx - rot.width//2, TH - 178), rot)

# 左侧：行号 + 参数名（右对齐）
for r in range(1, N+1):
    cy = Y0 + (r-1)*C + C//2
    d.text((8, cy - 11), "%2d" % r, font=fnum, fill=GRAY)
    nm = CN[r]
    tw = d.textlength(nm, font=fname)
    d.text((LW - 14 - tw, cy - 11), nm, font=fname, fill=INK)

# 单元格内容
for (r, c), plist in CELLS.items():
    x = X0 + (c-1)*C; y = Y0 + (r-1)*C
    items = [str(p) for p in plist[:4]]
    lines = [items[:2], items[2:]] if len(items) > 2 else ([items] if items else [])
    lines = [l for l in lines if l]
    if not lines: continue
    lh = 16
    top = y + C//2 - (len(lines)*lh)//2
    for li, line in enumerate(lines):
        parts = []
        for i, t in enumerate(line):
            parts.append((t, RED if (li == 0 and i == 0) else INK))
        sep = " "
        total = sum(d.textlength(t, font=fcell) for t, _ in parts) + d.textlength(sep, font=fcell)*(len(parts)-1)
        px = x + C/2 - total/2
        for i, (t, col) in enumerate(parts):
            d.text((px, top + li*lh), t, font=fcell, fill=col)
            px += d.textlength(t, font=fcell) + d.textlength(sep, font=fcell)

# 角落说明
d.text((10, TH - 196), "列 = 被恶化", font=F(15, True), fill=GRAY)
d.text((10, TH - 172), "行 = 要改善", font=F(15, True), fill=GRAY)

# 底部
cap = "Altshuller 39×39 矛盾矩阵　格内为发明原理编号，红色为第一推荐　灰格为对角（同一参数，按惯例留空）"
d.text((LW, H - 52), cap, font=F(21), fill=GRAY)

out = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\矛盾矩阵_39x39_真表.png"
img.save(out, "PNG")
print("已生成:", out, img.size)
print("有值格:", len(CELLS), "| 对角线格被误填:", sum(1 for (r,c) in CELLS if r==c))
mx = max(len(v) for v in CELLS.values()); print("单格最多原理数:", mx)
print("抽查 1,3:", CELLS.get((1,3)), " 14,1:", CELLS.get((14,1)), " 9,14:", CELLS.get((9,14)))