# -*- coding: utf-8 -*-
import json, datetime
from copy import deepcopy
from pathlib import Path
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
SRC = D / "TRIZ_W2A_矛盾解剖_15min.pptx"
OUT = D / "TRIZ_W2A_矛盾解剖_15min_优化版.pptx"

NOTES = [
 "好，我们开始。先问一句：你有没有遇到过这种时候——明明很努力，方案也出了一堆，可就是没有一个能用？今天这十五分钟，我只讲一件事：TRIZ 怎么看待一个难题。我要把“矛盾”这个词，从一句口号，变成你能写在纸上、还能被别人反驳的一句话。",
 "先看三个真事。电池，要跑一整天，要十五分钟充满，还不能起火。入侵检测模型，要把攻击抓全，又几乎不许误报。智能体，要自动把事做完，又要每一步都能追责。你发现没有？这三件事都不是“想不出来”，而是两个都正确的要求在打架。方案从来不缺，缺的是把这两个要求摆在同一张纸上。",
 "那什么叫矛盾？TRIZ 的定义很窄：改善一个参数，会恶化另一个参数。注意，是两个可测量的参数，不是一个模糊的不满。给你一句判据：你能把它写成“提高 X，同时不牺牲 Y”吗？写得出来，就是矛盾；写不出来，多半还只是抱怨。抱怨没法讨论，结构才能讨论。",
 "矛盾有三种形态。技术矛盾，改善 A 恶化 B，用矛盾矩阵加四十条原理。物理矛盾，同一个参数既要 A 又要非 A，用分离原理。管理矛盾，两个部门目标冲突，TRIZ 没有专用工具，做法是先翻译成技术矛盾，再进矩阵。为什么停在第一行就平庸？因为技术矛盾里只有两个参数在拉扯，没有一个“同一个对象”，分离根本无处下手，你只能把额度从这边挪到那边——那就是折中。",
 "这一步是分水岭。技术矛盾是症状，物理矛盾是病灶。怎么走？用五问法，连续问“为什么必须二选一”，问到不能再用“因为”回答为止。看机翼：航程不够，要多装油，油装进机翼，机翼变厚，阻力上升。追问到底，落在一句话上——机翼那个位置，既要厚，又要薄。然后呢？让机翼内部直接成为油箱，外形不变，容积变大。这就是整体油箱。",
 "矩阵是个查表工具。横轴是你要改善的参数，纵轴是会被恶化的参数，格子里是几条原理的编号。四步：写成参数对、查格子、追问语义、收敛。第三步最容易偷懒——很多人拿到编号就去搜“分割是什么意思”，然后套模板。第四步是纪律：只挑一两条做可验证的方案，其余归档。查不到，就说查不到。它给的是方向，不是答案。",
 "第一个例子。笔记本散热要好，风扇要大、风道要厚；可用户要薄。这是技术矛盾。查矩阵，拿到分割、抽取、动态化、中介物。真正落地的是热管：不在 CPU 原地散热，把热搬到机身边缘的鳍片上再吹走。注意，矛盾没有被折中，它是被换了位置。这是空间分离的一次具体使用。",
 "分离是对付物理矛盾的总办法：让“既要又要”在不同条件下分别成立，而不是同时成立。四把刀：时间、空间、条件、尺度。雨伞只在雨天有用，是时间；刀刃硬刀背韧，是空间；气囊只在碰撞瞬间充气，是条件；羽绒被宏观蓬松、微观锁气，是尺度。这四把刀，比背四十条原理更管用。",
 "第二个例子来自我们自己的论文。入侵检测要提高召回，误报就上升，这是技术矛盾。写成物理矛盾就是：同一个阈值，既要低到抓得住，又要高到不冤枉。解法是条件分离——按时段、按业务线、按置信度分层，各用各的阈值。同一批两千零四十个决策上，三种策略就是三种分离方式。关键细节：取舍写在实验之前，不是事后挑一个好看的。",
 "这一页是本讲的答案。阿奇舒勒按发明程度把专利分五级：一级不考虑矛盾，二级解决技术矛盾，三级解决物理矛盾，四级让矛盾一次性消失，五级是全新发现。比例是 32、45、18、4、不到 1。四级那句“一次性消失”什么意思？不是压住，也不是绕开，而是换了原理——原来那两个要求不再落在同一个对象上，冲突的前提没了。比如白炽灯要亮就得发热，LED 换了发光机制，亮度不再依赖高温，矛盾就不成立了。但四级只有百分之四，而且多半是事后才看得清。",
 "最后给一个警告：折中不是解决矛盾。双方各让一步，那个“既要又要”还在，只是被压住了，条件一变就原样回来，而且常常更贵。TRIZ 的立场很硬：折中是没解出矛盾时交的税。怎么判断手里的方案是解还是折中？就问一句——落地以后，那个“又要…又要…”还在不在。还在，就是折中。",
 "收一下，三句话。第一，矛盾不是障碍，它是问题被你理解到位的证据。第二，技术矛盾给方向，物理矛盾给刀刃，分离给解法，缺一件，方案就退回折中。第三，你能把矛盾写多准，创新就能走多远。下一讲，我们把它和另外两种约束观放在一起看，包括《矛盾论》。",
]

S4 = ["技术矛盾给你方向，物理矛盾给你刀刃。",
      "停在第一行＝只知道“哪两个参数在打架”，没问“为什么非此即彼”；没有“同一个对象”，分离就无处下手，只剩折中。"]
S10 = ["创新等级不取决于投入多少，取决于你把矛盾处理到哪一层。停在折中，就永远在一级和二级。",
       "“一次性消失”＝换了原理，原来那两个要求不再落在同一个对象上；不是压住，也不是绕开。"]


def fill(tf, texts):
    p0 = tf.paragraphs[0]
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    runs = list(p0.runs)
    if runs:
        runs[0].text = texts[0]
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.add_run().text = texts[0]
    r0 = p0.runs[0]
    for t in texts[1:]:
        np = tf.add_paragraph()
        if p0._p.pPr is not None:
            np._p.insert(0, deepcopy(p0._p.pPr))
        r = np.add_run()
        r.text = t
        if r0._r.rPr is not None:
            r._r.insert(0, deepcopy(r0._r.rPr))


def shape_starting(slide, head):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip().startswith(head):
            return sh
    return None


stats = {}
for name in ["TRIZ_W2A_矛盾解剖_15min.pptx", "TRIZ_W2B_盒子与矛盾论_15min.pptx",
             "TRIZa_AI4Science_15min.pptx", "TRIZa_AI4Science_15min_v2.pptx",
             "TRIZa_15min_handoff.zip"]:
    f = D / name
    st = f.stat()
    stats[name] = [st.st_size, datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")]
(D / "_A_stats.json").write_text(json.dumps(stats, ensure_ascii=False), encoding="utf-8")

prs = Presentation(str(SRC))
s4 = prs.slides[3]
sh = shape_starting(s4, "技术矛盾给你方向")
fill(sh.text_frame, S4)
sh.height = int(0.95 * 914400)

s10 = prs.slides[9]
sh10 = shape_starting(s10, "创新等级不取决于投入多少")
fill(sh10.text_frame, S10)
sh10.height = int(0.80 * 914400)

for i, s in enumerate(prs.slides):
    s.notes_slide.notes_text_frame.text = NOTES[i]

prs.save(str(OUT))

chk = Presentation(str(OUT))
t4 = shape_starting(chk.slides[3], "技术矛盾给你方向").text_frame.text
t10 = shape_starting(chk.slides[9], "创新等级不取决于投入多少").text_frame.text
nl = [len(s.notes_slide.notes_text_frame.text) for s in chk.slides]
total = sum(nl)
print(json.dumps({
    "out": OUT.name, "size": OUT.stat().st_size, "slides": len(chk.slides),
    "s4_text": t4, "s10_text": t10,
    "note_chars": nl, "note_total": total,
    "est_min_at_260": round(total / 260.0, 2),
    "has_ariz": any("ARIZ" in s.notes_slide.notes_text_frame.text.upper() for s in chk.slides),
    "has_marker": any(("---" in s.notes_slide.notes_text_frame.text) or ("请" in s.notes_slide.notes_text_frame.text and "解释" in s.notes_slide.notes_text_frame.text) for s in chk.slides),
}, ensure_ascii=False))