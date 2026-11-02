# -*- coding: utf-8 -*-
import sys, os, shutil
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
import build_w2 as B
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
BAK = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m_备份-加三步法前.pptx"
if not os.path.exists(BAK): shutil.copyfile(P, BAK)

prs = Presentation(P)
print("before:", len(prs.slides))

def fill(shape, blocks):
    """blocks = [(text, size, bold, color), ...]"""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.clear()
    first = True
    for text, size, bold, color in blocks:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(7)
        p.line_spacing = 1.2
        r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = bold
        r.font.name = "微软雅黑"
        r.font.color.rgb = color
        r._element.get_or_add_rPr()

def set_notes(s, text):
    ns = s.notes_slide
    tf = ns.notes_text_frame
    if tf is None:
        import copy
        from pptx.enum.shapes import PP_PLACEHOLDER
        src = prs.slides[1].notes_slide
        for shp in src.shapes:
            if shp.is_placeholder and shp.placeholder_format.type == PP_PLACEHOLDER.BODY:
                ns.shapes._spTree.append(copy.deepcopy(shp._element)); break
        tf = ns.notes_text_frame
    tf.text = text

# ============ 第 7 页：填内容 ============
s7 = prs.slides[6]
shapes = {sh.name: sh for sh in s7.shapes}
fill(shapes["TextBox 2"], [("（David Epstein）《盒子之内》：没有矛盾，就造一个", 30, False, B.INK)])
fill(shapes["Rectangle 4"], [
    ("常识：跳出盒子", 20, True, B.INK3),
    ("发散思维，想到更多选项，", 20, False, B.INK),
    ("追求更多自由。", 20, False, B.INK),
])
fill(shapes["Rectangle 5"], [
    ("这本书：钻进盒子", 20, True, B.RED),
    ("主动给自己设约束，", 20, False, B.INK),
    ("把可能性砍掉 90%。", 20, False, B.INK),
])
fill(shapes["Rectangle 6"], [
    ("约束落到两个要求上，就成了矛盾——所以没有矛盾时，你可以亲手造一个。", 20, False, B.INK),
])
set_notes(s7,
 "那么，如果一个问题里本来就没有矛盾呢？我给你一个反直觉的办法：没有矛盾，就自己造一个。"
 "这句话不是 TRIZ 里的，是我自己的说法。而且它也不是让你没苦硬吃，它来自一本书的启发。"
 "作者是写《成长的边界》的大卫·爱泼斯坦——那本书讲未来属于通才，跟过去一直推的专精路线正好相反。"
 "他今年五月出的新书叫《盒子之内：为什么约束让我们变得更好》。"
 "我们从小听到的建议是「跳出盒子思考」，他倒过来：不是跳出盒子，而是钻进盒子；不是摆脱约束，而是主动给自己造约束。"
 "哪怕本来没有盒子，你也应该先动手造一个。"
 "而约束落到两个要求上，就是矛盾——所以没有矛盾的时候，你可以造一个。")

# ============ 新增页：造盒子的三步法（插到第 11 位） ============
s = B.page(prs, "03f  ·  三步法", "造盒子的三步法：切断、配对、别绕道")
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip() == "03f  ·  三步法":
        for p in sh.text_frame.paragraphs:
            for r in p.runs: r.font.size = Pt(14)

B.table(s, [["步", "做什么", "水杯那个例子"],
            ["① 切断", "找到并堵死你最省力的那条路", "不设计「圆柱 + 把手」这个现成答案"],
            ["② 配对", "从四个维度挑一到几条，成对设", "没有把手 + 老人稳端热水；只用天然材料"],
            ["③ 别绕道", "钻进盒子，时时自检", "手一痒想加个小把手——那是放弃，不是创新"]],
        B.ML, 2.14, B.CW, [1.55, 4.40, 5.68], size=20, head_size=20, row_h=0.80)

tf = B.textbox(s, B.ML, 5.10, B.CW, 0.75)
B.bullet(tf, "门捷列夫拿掉「一卷讲完」，就没有周期表；苏斯博士拿掉「只用 50 个词」，就没有《绿鸡蛋和火腿》。",
         20, first=True, space_after=6)
B.band(s, "约束不是牢笼，是脚手架——你自己搭的盒子，才是把你逼上去的那道梯子。", 5.98, size=20, height=0.72)

tf = B.textbox(s, B.ML, 6.98, B.CW - 1.3, 0.32)
B.rich(tf.paragraphs[0], "TRIZ 创新构思法实战研讨班  ·  第 1 次 Workshop", 12, B.INK3, accent=B.RED, spc=100)
tf2 = B.textbox(s, B.SW - B.MR - 0.7, 6.98, 0.7, 0.32)
B.para(tf2, "11", 12, B.INK3, align=PP_ALIGN.RIGHT, first=True)

set_notes(s,
 "好，四个维度回答的是「往哪个方向设」。具体到操作，可以缩成三步。咱们不讲抽象规则，直接拿一个最普通的任务走一遍：设计一款水杯。"
 "先问自己，如果现在动手，我会设计出什么？大概率是圆柱杯身加一个把手，再换个颜色、印个图案——因为几百年来所有人都这么做，这是最现成、最不费力的答案。"
 "而这条路，恰恰是第一步要堵死的：识别最省力路径，主动切断。"
 "第二步，挑几个维度，设配对约束。先给一条主体约束：这个杯子不许有把手，还要让手上没劲的老人也能稳稳端起一杯热水。"
 "这一条砍掉了「加把手」，同时逼你死磕一个新问题：没有把手，怎么解决烫和拿不稳？再叠一层资源约束：不许用双层不锈钢、塑料隔热套，只能用天然材料。"
 "两条一配对，盒子就有了。常规设计被砍光，剩下的可能里，那种贴合掌心的天然异形杯，恰恰是别人没走过的路。"
 "第三步，钻进盒子，别想着绕道。最大的诱惑不是放弃，而是画着画着，手又痒了，「要不还是加个小把手吧」——所以得时时回头自检。"
 "同一个三步，用在一首诗上也成立：写故乡，不写我怎么想念，只写故乡怎么一点点消失，而且全诗不许出现一个情绪词——那是目标约束加方法约束。"
 "门捷列夫拿掉「一卷讲完」，就没有周期表；苏斯博士拿掉「五十个词」，就没有《绿鸡蛋和火腿》。约束真的在起作用，创造力才会被逼到极限。")

# 移动到第 11 位（索引 10）
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
new = ids[-1]
sldIdLst.remove(new)
sldIdLst.insert(10, new)
print("after insert:", len(prs.slides))

# ============ 页脚按位置重排 + kicker 顺号 ============
KICKS = {7: "03b  ·  盒子", 8: "03c  ·  配对约束", 9: "03d  ·  往哪设",
         10: "03e  ·  两种用法", 11: "03f  ·  三步法"}
fixed = 0
for idx, sl in enumerate(prs.slides, 1):
    if idx >= 2:
        want = "%02d" % idx
        for sh in sl.shapes:
            if sh.has_text_frame and sh.top/914400 > 6.9 and sh.text_frame.text.strip().isdigit():
                if sh.text_frame.text.strip() != want:
                    sh.text_frame.paragraphs[0].runs[0].text = want; fixed += 1
    if idx in KICKS:
        for sh in sl.shapes:
            if sh.has_text_frame and sh.top/914400 < 0.6 and sh.left/914400 < 1.0:
                t = sh.text_frame.text.strip()
                if "·" in t and len(t) < 24:
                    sh.text_frame.paragraphs[0].runs[0].text = KICKS[idx]
print("页脚修正:", fixed)

prs.save(P)
chk = Presentation(P)
print("\n最终:", len(chk.slides), "页 | notes 合计",
      sum(len(x.notes_slide.notes_text_frame.text) for x in chk.slides), "字")
for i, sl in enumerate(chk.slides, 1):
    ti = [sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.name == "TextBox 2"]
    ki = [sh.text_frame.text.strip() for sh in sl.shapes
          if sh.has_text_frame and sh.top/914400 < 0.6 and sh.left/914400 < 1.0 and "·" in sh.text_frame.text]
    fn = [sh.text_frame.text.strip() for sh in sl.shapes
          if sh.has_text_frame and sh.top/914400 > 6.9 and sh.text_frame.text.strip().isdigit()]
    print(f"  {i:2d} {(ki[0] if ki else '-'):<18} {(ti[0] if ti else '(封面)')[:40]:<42} 页脚={fn[0] if fn else '-'}")