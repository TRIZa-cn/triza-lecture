from PIL import Image
im = Image.open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\矛盾矩阵_39x39_真表.png")
print("full:", im.size)
# 左上角区域：表头 + 前 8x8 格
im.crop((0, 0, 200+56*9, 215+56*9)).save(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\_crop_tl.png")
# 中部：第 14-22 行 / 10-18 列
im.crop((200+56*9, 215+56*13, 200+56*19, 215+56*22)).save(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\_crop_mid.png")
print("crops ok")