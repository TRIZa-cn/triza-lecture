# -*- coding: utf-8 -*-
import sys, os, shutil
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
import build_w2 as B
from pptx import Presentation

A = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
BAK = A.replace(".pptx", "_插页前备份.pptx")
if not os.path.exists(BAK):
    shutil.copyfile(A, BAK)
prs = Presentation(A)
print("before:", len(prs.slides), "slides")

# 取第 4 页页脚原文，保证新页页脚与其他页一致
page4 = prs.slides[3]
LABEL = None
for sh in page4.shapes:
    if sh.has_text_frame and sh.top / 914400 > 6.9 and not sh.text_frame.text.strip().isdigit():
        LABEL = sh.text_frame.text.strip()
        break
print("label:", LABEL)

NOTE = ("这一页是本讲的分水岭。同一件事，三种写法，能用到的工具完全不同。我拿智能体的审批说。"
        "第一种，抱怨：审批太多，大家烦。这种写法没有参数，所以没有工具，你只能加人、加钱、加流程，问题原样回来。"
        "第二种，技术矛盾：审批覆盖率每提高一点，流程时延就涨一点。这个写法能用矛盾矩阵，但矩阵只给方向，最后多半落到折中——"
        "覆盖率提到八成，时延压到两秒。"
        "第三种，物理矛盾：同一条授权判据，既要粗到智能体能自主调用，又要细到不漏掉 read 权限。"
        "到这一层，同一个对象出来了：那双授权判据——就是那条判据本身。有了对象，才有分离，才可能有真解。"
        "所以我们练的不是解题，是把话写深一层：写法决定你能用哪把工具。")
NOTE = NOTE.replace("那双授权判据——就是那条判据本身", "那条授权判据本身")

NEW_TITLE = "同一个问题，三种写法"
s = B.page(prs, "02b  ·  三种写法", NEW_TITLE, notes=NOTE)

ROWS = [
    ["写法", "审批这件事长什么样", "你只能用"],
    ["① 抱怨", "审批太多，大家烦", "没有工具，只能加人加钱"],
    ["② 技术矛盾", "审批覆盖率 ↑ ↔ 流程时延 ↑", "矩阵 → 大多只能折中"],
    ["③ 物理矛盾", "**同一条授权判据**，既要粗又要细", "分离原理 → 可能真解"],
]
B.table(s, ROWS, B.ML, 2.12, B.CW, [2.0, 5.9, 3.73], size=15, row_h=0.78)
B.band(s, "写法决定能用哪把工具。写浅一层，不是解不出来，是工具都用不上。", 5.45, size=15)
B.footer(s, 5, LABEL)

# 把新页移到第 5 位（索引 4）
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
new = ids[-1]
sldIdLst.remove(new)
sldIdLst.insert(4, new)

# 旧 05..12 页脚号后移一位（降序避免撞号）
cnt = 0
for n in range(12, 4, -1):
    old, newno = "%02d" % n, "%02d" % (n + 1)
    for sh in prs.slides[n - 1].shapes:
        if sh.has_text_frame and sh.top / 914400 > 6.9:
            t = sh.text_frame.text.strip()
            if t == old:
                sh.text_frame.paragraphs[0].runs[0].text = newno
                cnt += 1
print("renumbered footers:", cnt)

# 第 4 页讲稿补一句交接
n4 = prs.slides[3].notes_slide.notes_text_frame.text
ADD = "下一页，我用你们天天遇到的智能体审批，把这三行一次走完。"
if ADD not in n4:
    prs.slides[3].notes_slide.notes_text_frame.text = n4.rstrip() + ADD
print("page4 notes:", len(prs.slides[3].notes_slide.notes_text_frame.text))

prs.save(A)

chk = Presentation(A)
print("after:", len(chk.slides), "slides")
for i, sl in enumerate(chk.slides, 1):
    ti = [sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.name == "TextBox 2"]
    ki = [sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.name == "TextBox 1"]
    fn = [sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.top / 914400 > 6.9 and sh.text_frame.text.strip().isdigit()]
    print(f"  {i:2d} [{ki[0] if ki else '-'}] {(ti[0] if ti else '(封面)')[:30]:<32} 页脚={fn[0] if fn else '-'}")
tot = sum(len(x.notes_slide.notes_text_frame.text) for x in chk.slides)
print("notes total:", tot, "≈", round(tot / 270, 2), "min")