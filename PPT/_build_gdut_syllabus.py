# -*- coding: utf-8 -*-
import shutil
from docx import Document
from docx.shared import Pt

SRC = r"D:\doc\ZGCA\Lectures\TRIZ\北京中关村学院教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx"
DST = r"D:\doc\ZGCA\Lectures\TRIZ\广东工业大学教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx"
shutil.copyfile(SRC, DST)

d = Document(DST)
P = d.paragraphs

def mk(before_para, text, bold=False):
    p = before_para.insert_paragraph_before("", style="Normal")
    r = p.add_run(text)
    r.font.size = Pt(12)
    r.bold = bold
    return p

# 1) 学校名称
assert P[0].text.strip() == "北京中关村学院课程大纲"
P[0].runs[0].text = "广东工业大学课程大纲"

# 2) 课程简介（中文）末尾补一段：TRIZa 作为教学载体
p016 = P[16]
mk(P[17], "本课程的实践环节统一在 TRIZa 平台（TRIZa.cn）上完成。它把 TRIZ 的方法与工具做成有输入、有输出、有记录的“工作台”，学生每一步判断都留下依据，课后能接着做，助教也能核对；课堂讨论因此不会停在“聊得很热闹，散会什么都没留下”。平台是辅助工具，用来降低使用门槛，不替代 TRIZ 的判断，也不替代学生自己的思考。")

# 3) 英文简介末尾补一句
p021 = P[21]
r = p021.add_run(" All hands-on sessions run on TRIZa (TRIZa.cn): its workbenches keep every TRIZ judgement and its evidence on record, and AKBank accumulates citable domain knowledge instead of unsourced recollection.")
r.font.size = Pt(12)

# 4) 教学内容安排末尾：一句衔接
mk(P[72], "以上各章的方法与工具，均在 TRIZa 的发明工作台上配套实现。课堂按“讲理论—看演示—做实战”三步走，每次实战都在同一套记录里推进，学期末可完整复盘。")

# 5) 新增一节：TRIZa 平台与配套工具
anchor = None
for p in d.paragraphs:
    if p.text.strip() == "作业和实验安排":
        anchor = p
        break
assert anchor is not None

BLOCK = [
    ("TRIZa 平台：把 TRIZ 落到可操作的工具上", True),
    ("TRIZ 的概念和工具有很多，学生最常卡住的地方不是听不懂，而是“知道有这么个工具，却不知道下一步该做什么”。TRIZa 是本课程配套的实践平台，它把 TRIZ 的每一步做成有输入、有输出、有记录的工作台，使课堂演练从口头讨论变成可复盘的工程过程。本课程把五个工作台按“想—证—写—讲”的顺序串起来用。", False),
    ("一、发明工作台", True),
    ("功能：把一个模糊的科研或管理难题，按 TRIZ 的顺序拆成可检查的几步——写清目标与不可牺牲的约束，指出是哪两个参数在互相拉扯（技术矛盾），再追问“为什么非此即彼”（物理矛盾），给出候选方向，并为每个候选写清它要验证什么。", False),
    ("效果：讨论过程留痕，每一步都有依据和出处；课后能接着做，助教能核对，避免“聊得很热闹，散会什么都没留下”。", False),
    ("使用方法：①用一句话写目标，再列 2–3 条不可牺牲的约束；②描述现状与冲突，让平台按 TRIZ 术语复述，学生确认或纠正；③从候选方向里挑一个，写清验证方式与放弃条件；④导出本次记录，作为课程设计的一节。", False),
    ("二、AKBank：把领域知识沉淀成可引用的“知识原子”", True),
    ("它是什么：AKBank 是配套的知识库，每条记录由三部分组成——公开证据的题录（文献、专利、标准）、许可信息、以及使用者自己写的一句话要点。三样齐全，才算一条可用的记录。", False),
    ("为什么重要：TRIZ 的跨行业类比最怕“我记得好像有人这么做过”。AKBank 要求每条知识都带出处、许可和适用边界，避免把类比当成证据。", False),
    ("怎么沉淀：每次研讨结束时，把这次真正起作用的那几条外部知识登记进去（来源、许可、要点、适用边界）；下次遇到同类问题先检索，命中就引用，没命中就明确写“未检索到”，不允许编。一个学期下来，全班共同维护一个课程 AKBank，它既是复习材料，也是下一届的起点。", False),
    ("三、论文工作台", True),
    ("功能：把发明工作台里定下来的问题、矛盾、方案与验证结果，整理成论文结构（摘要、方法、实验、结论），并标出哪一部分还缺证据。", False),
    ("效果：写论文不必从空白页开始；更有用的是它会指出“这一句没有数据支撑”，逼着学生补实验，而不是补形容词。", False),
    ("使用方法：选定目标期刊或会议的结构模板 → 导入发明工作台的项目记录 → 逐节补齐 → 对照“缺证据清单”回实验台补数据。", False),
    ("四、实验工作台", True),
    ("功能：把“要验证什么”写成可执行的实验方案——指标怎么算、和谁比、样本多少、什么条件下停；跑完只记录真实指标。", False),
    ("效果：结论可复核，能区分“我觉得更好”和“测出来更好”；同一方案换参数重跑时，两次结果可以直接比。", False),
    ("使用方法：先冻结评价口径，再跑实验；跑完只填实测数字，不填期望值；结论里写清这次实验不能证明什么。", False),
    ("五、Presentation 工作台", True),
    ("功能：把一学期的工作整理成汇报或讲课材料——每页一个论点，配口播稿，并核对页数、字号、术语首次出现是否解释、讲稿时长是否与课时匹配。", False),
    ("效果：汇报不靠临场发挥；学生交上来的不只是幻灯片，还有可核对的讲稿与时长。", False),
    ("使用方法：选模板 → 导入项目记录 → 生成页面与讲稿 → 按核验清单逐项过（术语、时长、字号）→ 试讲一次再定稿。", False),
    ("六、使用边界（必须讲清）", True),
    ("平台给的是候选和结构，不是结论，判断由学生自己负责。检索不到就写“未检索到”，不允许用模型编造文献、数据或专利号。涉及未公开的科研数据，按学校和项目组的保密要求决定是否上传；不能上传的，用脱敏后的描述替代。评分看的是判断依据与过程记录，不看平台用得多不多。", False),
]
for text, bold in BLOCK:
    mk(anchor, text, bold)
mk(anchor, "", False)  # 与下一节之间留一个空行

# 6) 作业和实验安排：补课程设计的落地要求
for p in d.paragraphs:
    if p.text.strip() == "助教需求":
        anchor2 = p
        break
assert anchor2 is not None
for text in [
    "课程设计统一在 TRIZa 平台上完成，过程记录随报告一并提交：发明工作台的项目记录（问题、矛盾、候选方向与依据）、实验工作台的方案与实测指标、论文工作台的成稿、Presentation 工作台的汇报材料，四样合起来才算一次完整的课程设计。",
    "建议节奏：第 1–3 次课用发明工作台把问题写清并定下矛盾；第 4–5 次课做实验工作台的方案与首轮实测；第 6–7 次课用论文工作台成稿；第 8 次课用 Presentation 工作台做 10 分钟汇报。",
]:
    mk(anchor2, text)
mk(anchor2, "", False)

# 7) 考核方式：补过程记录要求
for p in d.paragraphs:
    if p.text.strip().startswith("最终成绩=课堂积极参与讨论贡献问题或建议20%"):
        r = p.add_run("；课程设计报告须附 TRIZa 过程记录，过程记录缺失的，按未完成计。")
        r.font.size = Pt(12)
        break

# 8) 教师简介：说明五个工作台
for p in d.paragraphs:
    if "开发了基于TRIZ理论的端到端创新构思和开发平台TRIZa.cn" in p.text:
        r = p.add_run("该平台由发明工作台、论文工作台、实验工作台与 Presentation 工作台组成，并以 AKBank 沉淀可引用的领域知识。")
        r.font.size = Pt(12)
        break

d.save(DST)

# 校验
chk = Document(DST)
txt = "\n".join(x.text for x in chk.paragraphs)
print("paras", len(chk.paragraphs))
print("title", chk.paragraphs[0].text)
print("北京中关村学院 occurrences:", txt.count("北京中关村学院"))
print("中关村 occurrences:", txt.count("中关村"))
print("广东工业大学 occurrences:", txt.count("广东工业大学"))
for k in ["发明工作台", "AKBank", "论文工作台", "实验工作台", "Presentation 工作台", "使用边界"]:
    print(k, txt.count(k))