# -*- coding: utf-8 -*-
from pptx import Presentation
A = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
OLD = "我要把“矛盾”这个词，从一句口号，变成你能写在纸上、还能被别人反驳的一句话。因为大部分人卡住，不是因为不聪明，而是因为从来没有把问题写成一句可以被推翻的话。"
NEW = ("我要把“矛盾”这个词，从一句口号，变成你能写在纸上的一句话。写出来的好处很直接：别人一眼就能看出你卡在哪，也就能一句话帮到你；写不出来，他只能点头。"
       "而且这件事很便宜——一句话在纸上被问倒，你只花五分钟改；同一个想法在项目里被撞见，要花三个月。"
       "大部分人卡住，不是因为不聪明，而是因为问题一直留在脑子里，从来没被写成一句别人看得懂、也检验得了的话。")
prs = Presentation(A)
n1 = prs.slides[0].notes_slide.notes_text_frame.text
assert OLD in n1, "未找到原句"
head = n1.split(OLD)[0]
prs.slides[0].notes_slide.notes_text_frame.text = head + NEW
prs.save(A)

chk = Presentation(A)
t = chk.slides[0].notes_slide.notes_text_frame.text
print("第 1 页讲稿 %d 字（原 169）\n约 %.0f 秒 @250字/分\n" % (len(t), len(t)/250*60))
print(t)
print("\n全篇 notes:", sum(len(x.notes_slide.notes_text_frame.text) for x in chk.slides), "字")