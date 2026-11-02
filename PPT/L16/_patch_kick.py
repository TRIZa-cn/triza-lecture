# -*- coding: utf-8 -*-
from pptx import Presentation
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(P)
s2, s4, s5 = prs.slides[1], prs.slides[3], prs.slides[4]

# ---------- 1) 第 2 页：搭桥之后加一句 ----------
n2 = s2.notes_slide.notes_text_frame.text
ADD2 = "顺带问一句：你手上那个约束，是谁给的？是物理，是标准，还是别人定的？这个问题先放着，等会儿要用。"
anchor = "对不对？"
if ADD2 not in n2:
    assert anchor in n2
    n2 = n2.replace(anchor, anchor + ADD2, 1)
    s2.notes_slide.notes_text_frame.text = n2
print("s2 notes:", len(n2))

# ---------- 2) 第 4 页：标题 + 讲稿 ----------
t = [sh for sh in s4.shapes if sh.name == "TextBox 2"][0]
runs = t.text_frame.paragraphs[0].runs
NEW_TITLE = "约束有三种来源，对应三种矛盾"
runs[0].text = NEW_TITLE
for r in runs[1:]:
    r._r.getparent().remove(r._r)
print("s4 title ->", t.text_frame.text)

N4 = ("矛盾有三种形态，但三种形态不是三种难度，是约束落在三个不同的地方。"
      "约束落在两个参数之间，就是技术矛盾：改善 A，恶化 B，用矛盾矩阵加四十条原理。"
      "约束落在同一个对象上，就是物理矛盾：同一个参数既要 A 又要非 A，用分离原理。"
      "约束落在两个主体之间，就是管理矛盾：两个部门的目标冲突，TRIZ 没有专用工具，做法是先翻译成技术矛盾，再进矩阵——"
      "比如“交付要快”和“质量要稳”，翻译过来就是周期和缺陷率这两个参数。"
      "现在看最后那句：为什么停在第一行就平庸？因为技术矛盾里只有两个参数在拉扯，没有一个“同一个对象”，分离根本无处下手，"
      "你只能把额度从这边挪到那边——那就是折中。")
s4.notes_slide.notes_text_frame.text = N4
print("s4 notes:", len(N4))

# ---------- 3) 第 5 页：半句 + 讲稿 ----------
tb5 = [sh for sh in s5.shapes if sh.name == "TextBox 5"][0]
p0 = tb5.text_frame.paragraphs[0]
p0.runs[0].text = "追问五次“为什么必须二选一”" 
if len(p0.runs) == 1:
    r = p0.add_run()
    r.text = "——把约束追到对象上"
    r.font.size = p0.runs[0].font.size
    r.font.bold = True
    r.font.name = p0.runs[0].font.name
else:
    p0.runs[1].text = "——把约束追到对象上"
print("s5 p0 ->", p0.text)

N5 = ("这一步是分水岭。技术矛盾是症状，物理矛盾是病灶。怎么走？把约束往下追，一直追到它落在同一个对象上——"
      "具体做法是五问法，连续问“为什么必须二选一”，问到不能再用“因为”回答为止。"
      "看机翼：航程不够，要多装油，油装进机翼，机翼变厚，阻力上升。追问到底，落在一句话上——机翼那个位置，既要厚，又要薄。"
      "这就是把约束追到了对象上。然后呢？让机翼内部直接成为油箱，外形不变，容积变大。这就是整体油箱。"
      "只治症状，方案就只能是折中。注意，五问法最容易走偏：问成“谁的责任”，就变成追责了；要一直问“为什么必须”。")
s5.notes_slide.notes_text_frame.text = N5
print("s5 notes:", len(N5))

prs.save(P)
tot = sum(len(s.notes_slide.notes_text_frame.text) for s in Presentation(P).slides)
print("A notes total:", tot, "≈", round(tot/270, 2), "min @270")