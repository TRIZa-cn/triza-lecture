# -*- coding: utf-8 -*-
# 只改底部说明：把编号标成"示意"，避免被当成真实查表结果
from PIL import Image, ImageDraw, ImageFont
PAPER=(250,249,246); GRAY=(107,114,128)
img = Image.open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\矛盾矩阵_示意.png").convert("RGB")
d = ImageDraw.Draw(img)
F = lambda s: ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", s)
d.rectangle([180, 1050, 1300, 1105], fill=PAPER)     # 盖住旧说明
d.text((200, 1056), "格子里的编号是示意，实际以 39 × 39 表为准", font=F(38), fill=GRAY)
img.save(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\矛盾矩阵_示意.png", "PNG")
print("已改注：", img.size)