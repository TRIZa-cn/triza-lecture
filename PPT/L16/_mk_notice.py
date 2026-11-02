# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

OUT = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-Workshop-会议通知-模板.docx"
GREY = RGBColor(0x55, 0x5F, 0x6B)
RED = RGBColor(0xB0, 0x3A, 0x2B)
d = Document()
s = d.sections[0]
s.page_height, s.page_width = Cm(29.7), Cm(21.0)
s.left_margin = s.right_margin = Cm(2.4); s.top_margin = s.bottom_margin = Cm(2.2)

def setf(run, size, bold=False, color=None, name="微软雅黑"):
    run.font.size = Pt(size); run.bold = bold; run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if color: run.font.color.rgb = color

def P(text="", size=11.5, bold=False, color=None, before=0, after=5, spacing=1.45, indent=0):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    if indent: p.paragraph_format.left_indent = Cm(indent)
    if text: setf(p.add_run(text), size, bold, color)
    return p

def H(text, size=15):
    P(text, size, True, RED, before=14, after=6)

h = d.add_paragraph(); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
setf(h.add_run("TRIZa Workshop · 创新构思实战研讨会"), 20, True)
h2 = d.add_paragraph(); h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
setf(h2.add_run("会议通知（模板，每周只改【】里的内容）"), 11, False, GREY)

H("一、群消息短版（约 220 字，直接贴群）")
P("【TRIZa Workshop】本周五（【10月30日】）16:00，【地点】。")
P("自愿参加，不用报名，来就行。唯一的要求：带一个真问题——你自己这段时间确实卡住、又一直没有好办法的那个，科研、工程、管理都行。")
P("现场这么走：先用 15 分钟讲一个 TRIZ 小专题（本周：矛盾的定义与矛盾矩阵），然后把你带来的问题丢给 TRIZa 跑，我们看着它的输出一起挑、一起骂、一起改。目标是当场给你一条明确有启发的路，运气好能给出方案甚至代码。")
P("本周专题：怎么把一句抱怨，写成一句可测量、可被推翻的矛盾。")
P("带题模板（一句话就够）：我想让【　　】变好，但现在【　　】会变差。")

H("二、完整版通知（发邮件或公众号）")
P("各位同学、同事：", after=8)
P("每周五 16:00 起，我们在【地点】开一个 TRIZa Workshop 创新构思实战研讨会，自愿参加。")
P("这个活动只有一个目的：把真实难题当场往前推一步。所以顺序和我们平时听讲座相反——不是先讲理论再举例，而是先有人把真问题摆出来，理论只在他需要的时候出现。")
P("时间：每周五 16:00–17:30（可中途加入，随时退出）")
P("地点：【　　　　】")
P("参加方式：自愿，无需报名；有事来不了也不用请假")
P("", after=2)

P("现场流程", 12, True, before=6, after=4)
P("16:00–16:10　亮题。愿意的同学讲清自己那个难题：卡在哪、试过什么、为什么不行。", indent=0.5, after=3)
P("16:10–16:25　小专题。15 分钟讲一个 TRIZ 知识点，只讲这次要用的那部分。", indent=0.5, after=3)
P("16:25–17:10　现场跑。把题目输入 TRIZa，一起看它给的矛盾拆解、原理建议和候选方向，逐条判断哪些成立、哪些是空话。", indent=0.5, after=3)
P("17:10–17:30　收敛。定下一条可验证的下一步：要做什么实验、看什么指标、什么条件下放弃。", indent=0.5, after=6)

P("带什么来", 12, True, before=6, after=4)
P("一句话的题目模板：", after=2)
P("　　我想让 ____ 变好，但现在 ____ 会变差。", bold=True, after=4)
P("填得出来，就说明这是一个能被讨论的问题；填不出来，我们一起把它填出来——这个过程本身往往就是收获。")
P("如果有之前的实验记录、数据、代码片段，带上更好；没有也行。", after=6)

P("本周专题", 12, True, before=6, after=4)
P("矛盾的定义与矛盾矩阵：把问题从“哪里不好”推进到“哪两个要求必须二选一”。", after=3)
P("这 15 分钟你只需要带走三样东西：", after=2)
P("① 一个判据——能不能指出“同一个对象”？能，是物理矛盾；不能，是技术矛盾。", indent=0.4, after=2)
P("② 一个动作——把抱怨改写成“提高 X，同时不牺牲 Y”，再翻成两个可测量的参数。", indent=0.4, after=2)
P("③ 一次查表——现场拿一位同学的题目查矛盾矩阵，看它给的方向有没有启发。", indent=0.4, after=6)

P("几句实话（先说清，省得后面误会）", 12, True, before=6, after=4)
P("· TRIZa 给的是候选方向和结构，不是结论。所有判断由出题人自己做，我们只负责把它逼到必须选择的地方。", indent=0.4, after=3)
P("· 不保证当场出方案。有些题目一轮就有路，有些要几周；但每轮都应该留下一条可验证的下一步。", indent=0.4, after=3)
P("· 涉及未公开数据、涉密项目或合作方材料的，请先脱敏，或只讲结构不讲数据。不能上传的，我们用口头描述走一遍。", indent=0.4, after=3)
P("· 出的题归出题人。我们不会拿别人的题去写论文、做项目或对外分享；需要引用会先在群里问。", indent=0.4, after=8)

P("联系方式：【　　　　】　　召集人：【　　　　】", after=4)

H("三、每周替换清单")
for t in ["① 日期（本周五的具体日期）", "② 地点 / 线上会议室号", "③ 本期专题名称与三样带走的东西",
          "④ 若有指定主讲或特邀出题人，补姓名", "⑤ 若有材料或前置阅读，补链接"]:
    P(t, 11, indent=0.4, after=3)

H("四、关于那 15 分钟专题简介：两种写法，别混用")
P("你现在这段（大纲版）：", 12, True, before=6, after=3)
P("「矛盾定义与矛盾矩阵，区分技术矛盾（两个参数间的权衡）与物理矛盾（同一参数的相反需求）。学习使用 39（或现代 TRIZ 中的 48 个）个通用工程参数和矛盾矩阵，将具体问题抽象化，并匹配 40 条发明原理。」", color=GREY, after=4)
P("这段话适合写进教学大纲——它要覆盖考核范围，必须完整。但它不是 15 分钟能讲完的东西，也不该照它讲。", after=6)

P("现场版（建议改按这个讲）：", 12, True, before=6, after=3)
P("· 1 分钟　钩子：你带来那个难题，八成不是“想不出办法”，而是两个都正确的要求在打架。")
P("· 4 分钟　一个判据：能不能指出“同一个对象”。能 → 物理矛盾（同一把锁，既要断电可用，又要只认云端）；不能 → 技术矛盾（散热好 ↔ 机身薄）。")
P("· 6 分钟　一个动作：把抱怨改写成“提高 X，同时不牺牲 Y”，再从参数表里挑出两个可测量的量。现场用某位同学的题走一遍。")
P("· 4 分钟　一次查表：拿这个参数对查矩阵，得到 3–4 条原理编号，逐条问“在我的场景里对应什么动作”。")
P("· 30 秒　一句边界：管理矛盾 TRIZ 没有专用工具，要先翻译成技术矛盾。", after=6)
P("为什么这么改：15 分钟里，听众记住的只能是一个判据加一个动作；39/48/40 这三组数字同时出现，反而让人记不住任何一个。数字留在页面和讲义里，嘴上只讲动作。", color=GREY)
d.save(OUT)
print("saved:", OUT, round(__import__("os").path.getsize(OUT)/1024, 1), "KB")