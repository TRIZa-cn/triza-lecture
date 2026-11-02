# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os

OUT = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-Workshop-会议通知-模板.docx"
GREY = RGBColor(0x55, 0x5F, 0x6B)
RED = RGBColor(0xB0, 0x3A, 0x2B)
d = Document(); s = d.sections[0]
s.page_height, s.page_width = Cm(29.7), Cm(21.0)
s.left_margin = s.right_margin = Cm(2.4); s.top_margin = s.bottom_margin = Cm(2.2)

def setf(r, size, bold=False, color=None, name="微软雅黑"):
    r.font.size = Pt(size); r.bold = bold; r.font.name = name
    r._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if color: r.font.color.rgb = color

def P(t="", size=11.5, bold=False, color=None, before=0, after=5, spacing=1.45, indent=0):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    if indent: p.paragraph_format.left_indent = Cm(indent)
    if t: setf(p.add_run(t), size, bold, color)
    return p

def H(t, size=15):
    P(t, size, True, RED, before=16, after=6)

h = d.add_paragraph(); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
setf(h.add_run("TRIZa Workshop · 创新构思实战研讨会"), 20, True)
h2 = d.add_paragraph(); h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
setf(h2.add_run("会议通知（v2 · 每周只改【】里的内容）"), 11, False, GREY)

# ---------------- 一、群消息 ----------------
H("一、群消息版（约 300 字，直接贴）")
P("【TRIZa Workshop · 本周五】")
P("时间：10月30日（周五）16:00–17:30　地点：【　　　　】", after=8)
P("自愿参加，不用报名，中途来、提前走都行。")
P("来的时候最好带一个真问题：你自己卡了两周以上、又一直没找到好办法的那个。科研、工程、管理都行，不用准备 PPT，一句话就够——", after=4)
P("　　我想让 ______ 变好，但现在 ______ 会变差。", bold=True, after=8)
P("现场这么走：先亮题（讲清卡在哪、试过什么）；我用 15 分钟讲一个小专题；然后把题扔进 TRIZa 现场跑，一起看它的输出、逐条判真假；最后收敛出一条可以验证的下一步。", after=8)
P("本周专题：怎么把一句抱怨，写成一句能判定的矛盾。", bold=True, after=3)
P("听完带走三样东西：一个判据（能不能指出“同一个对象”）、一个动作（把抱怨改写成“提高 X，同时不牺牲 Y”）、一次查表（现场拿某位的题查矛盾矩阵）。", after=8)
P("不保证当场出方案。但每轮都会留一条可验证的下一步。")

# ---------------- 二、完整版 ----------------
H("二、邮件 / 公告版")
P("各位同学、同事：", after=8)
P("每周五 16:00–17:30，我们在【地点】开一个 TRIZa Workshop 创新构思实战研讨会，自愿参加。")
P("这个活动只做一件事：把真实难题当场往前推一步。所以顺序和听讲座相反——不是先讲理论再举例，而是先有人把真问题摆出来，理论只在他需要的时候出现。", after=8)

P("现场怎么走", 12, True, before=6, after=4)
P("16:00–16:10　亮题。愿意的同学讲清自己那个难题：卡在哪、试过什么、为什么不行。", indent=0.5, after=3)
P("16:10–16:25　小专题。15 分钟讲一个 TRIZ 知识点，只讲这次要用的那部分。", indent=0.5, after=3)
P("16:25–17:10　现场跑。题目输入 TRIZa，一起看它给的矛盾拆解、原理建议和候选方向，逐条判断哪些成立、哪些是空话。", indent=0.5, after=3)
P("17:10–17:30　收敛。定下一条可验证的下一步：要做什么实验、看什么指标、什么条件下放弃。", indent=0.5, after=8)

P("带什么来", 12, True, before=6, after=4)
P("一句话的题目模板：我想让 ____ 变好，但现在 ____ 会变差。", bold=True, after=3)
P("填得出来，说明这是个能被讨论的问题；填不出来，我们一起把它填出来——这个过程本身往往就是收获。", after=3)
P("有实验记录、数据、代码片段的带上更好，没有也行。想带题的同学可以提前在群里说一声，方便把它排在前面。", after=8)

P("本周专题（可整块替换）", 12, True, before=6, after=4)
P("矛盾的定义与矛盾矩阵：把问题从“哪里不好”，推进到“哪两个要求必须二选一”。", after=4)
P("这一讲的核心动作只有一个——把同一件事写出三种深度：", after=3)
P("① 抱怨：审批太多，大家烦。→ 没有参数，也就没有工具。", indent=0.4, after=2)
P("② 技术矛盾：审批覆盖率 ↑ ↔ 流程时延 ↑。→ 能用矩阵，但只能折中。", indent=0.4, after=2)
P("③ 物理矛盾：同一条授权判据，既要粗又要细。→ 指认出“同一个对象”，才谈得上分离。", indent=0.4, after=4)
P("演示用智能体的审批。解法不在这一讲讲，留到后面的裁剪、ARIZ 和功能导向搜索。", color=GREY, after=8)

P("几句实话", 12, True, before=6, after=4)
P("· TRIZa 给的是候选方向和结构，不是结论。判断由出题人自己做。", indent=0.4, after=3)
P("· 不保证当场出方案，但每轮都留一条可验证的下一步。", indent=0.4, after=3)
P("· 涉及未公开数据、涉密项目或合作方材料的，请先脱敏，或只讲结构不讲数据。出的题归出题人，不会拿去做项目或对外分享。", indent=0.4, after=8)
P("联系方式：【　　　　】　　召集人：【　　　　】")

# ---------------- 三、替换清单 ----------------
H("三、每周替换清单")
for t in ["① 日期（本周五的具体日期）", "② 地点 / 线上会议室号", "③ 本周专题那一段（三行即可）",
          "④ 若有指定主讲或特邀出题人，补姓名", "⑤ 联系方式与召集人"]:
    P(t, 11, indent=0.4, after=3)

# ---------------- 四、两种写法 ----------------
H("四、15 分钟专题：两种写法，别混用")
P("大纲版（写进教学大纲，覆盖考核范围）：", 12, True, before=6, after=3)
P("「矛盾定义与矛盾矩阵，区分技术矛盾（两个参数间的权衡）与物理矛盾（同一参数的相反需求）。学习使用 39（或现代 TRIZ 中的 48 个）个通用工程参数和矛盾矩阵，将具体问题抽象化，并匹配 40 条发明原理。」", color=GREY, after=6)
P("现场版（照这个讲，15 分钟）：", 12, True, before=6, after=3)
P("· 1 分钟　钩子：你带来的难题，八成不是“想不出办法”，而是两个都正确的要求在打架。")
P("· 4 分钟　一个判据：能不能指出“同一个对象”。能 → 物理矛盾；不能 → 技术矛盾。")
P("· 6 分钟　一个动作：把抱怨改写成“提高 X，同时不牺牲 Y”，再翻成两个可测量的参数。用某位同学的题现场走一遍。")
P("· 4 分钟　一次查表：拿这个参数对查矩阵，得到 3–4 条原理编号，逐条问“在我这儿对应什么动作”。")
P("· 30 秒　一句边界：管理矛盾 TRIZ 没有专用工具，要先翻译成技术矛盾。", after=6)
P("原因：15 分钟里，听众能记住的只有一个判据加一个动作；39/48/40 三组数字同时出现，反而一个都记不住。数字留在页面，嘴上只讲动作。", color=GREY)

# ---------------- 五、本次优化 ----------------
H("五、这一版比上一版改了什么")
rows = [
 ("开头先说“不用报名、中途能走”", "自愿活动最大的阻力是怕被绑住，先把出口给出来"),
 ("把“带题模板”提到流程之前", "参与者只需要准备一句话，这是唯一的准备动作"),
 ("专题从“讲知识点”改成“带走三样东西”", "听众按可带走的东西记，比按知识框架记牢"),
 ("明确 16:00–17:30 的时间盒", "四个环节各有归属，现场不容易散"),
 ("加“想带题可提前说一声”", "避免现场抢麦，也让你能提前安排顺序"),
 ("边界压到三条", "原来的四条里有两条是同一件事，合并后更容易读完"),
]
tb = d.add_table(rows=len(rows) + 1, cols=2); tb.style = "Table Grid"
for i, (a, b) in enumerate([("改动", "原因")] + rows):
    for j, v in enumerate((a, b)):
        c = tb.cell(i, j); c.text = ""
        setf(c.paragraphs[0].add_run(v), 10.5, i == 0)
d.save(OUT)
print("saved:", OUT, round(os.path.getsize(OUT)/1024, 1), "KB")