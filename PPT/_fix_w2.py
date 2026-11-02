# -*- coding: utf-8 -*-
import json
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
FONT = "微软雅黑"

EXTRA_A = [
 "今天的十五分钟只讲一件事：把一个“难题”从一句抱怨，变成一句可以被推翻的话。四步：矛盾长什么样、怎么从技术矛盾切到物理矛盾、分离的四把刀、为什么折中不算解。三个例子分别来自工程、AI 和我们自己的判断流程。听完之后，请试着把你手上最烦的那个问题，写成“提高 X，同时不牺牲 Y”。",
 "这三个困境的共同点是：方案从来不缺，缺的是把“两个都对的要求”摆在同一张纸上。摆不出来，讨论就会自动滑向“再多加点资源”这类没有结构的建议，而这类建议通常只能换来折中。",
 "判据里最要紧的是“可测量”三个字。不可测量的要求，最后只能靠职位或嗓门解决。你把它写成参数，哪怕参数选得不精确，讨论也立刻有了共同对象，可以被检验、也可以被反驳。",
 "第三行我故意留着：管理矛盾 TRIZ 确实没有专用工具，硬套矩阵只会拿到一堆不适用的原理。承认边界比勉强通用更专业，这也是这门课反复强调的诚实。",
 "追问的技巧是：每问一次，都把“为什么”换成“在什么条件下”。当所有条件都写出来了，矛盾就被逼到了一句最朴素的话上——同一个位置，既要又要。",
 "矩阵只是索引，不是答案。真正的功夫在第三步：把“分割”“抽取”这些词翻译成你系统里的具体动作。翻不出来，通常说明问题本身还没想清楚，而不是矩阵不管用。",
 "热管的价值在于：矛盾没有被消灭，而是被搬到了另一个位置。判断一次创新有多深，就看它有没有让原来的矛盾在原来的位置上不再成立——这就是分离和折中的分界线。",
 "这四把刀请背下来。它们比四十条原理更常用，因为绝大多数的物理矛盾，最后都能落进时间、空间、条件、尺度这四个维度里的一个。",
 "三个策略的阈值都是在看数据之前定的，这一点比结果更重要。事后挑一个最好看的阈值，那不是条件分离，是拟合——而拟合出来的方案，换个环境就会失效。",
 "如果今天只记一张表，记这一张：矛盾的层级直接对应创新的等级。它解释了一个常见困惑——为什么有些人做了很多改进，却始终停在一级。",
 "折中最危险的地方不是它不好，而是它看起来像解，于是没人再往下问。项目一旦以折中收尾，那个“既要又要”就被永久地留在了系统里，还会随规模放大。",
 "下一讲换一个角度看同一件事：如果问题里根本没有矛盾，能不能自己造一个？这会引出那本书、也会引出《矛盾论》。",
]

EXTRA_B = [
 "三件事：把那本书的方法翻译成 TRIZ 的语言；把它和《矛盾论》一条一条对照；再给我们自己的研究报告做一次不客气的评估。",
 "三者摆在一起看，最容易犯的错是把“都在讲矛盾”当成“讲的是同一件事”。先分清各自在做哪个动作，后面的对照才有意义。",
 "门捷列夫这个故事值得多讲一句：约束不是外部干扰，它常常是创意的直接来源。合同和页数不是障碍物，是排除了那条最省力路径的手。",
 "成对这两个字是全部要点。单独一个排除，人只会假装遵守；单独一个引导，人只会把旧方案换个包装。",
 "四个维度其实是在回答“改哪一段”。改目标、改资源、改方法、改主体——注意，改的都不是难度，而是入口。",
 "这一页请慢一点。如果你接受“配对约束就是人为植入一个物理矛盾”，那么这本书和 TRIZ 就不是类比关系，而是同一套解法机器的两个入口。",
 "反过来也成立：TRIZ 用矩阵查出来的方向，本质上也是在给问题加约束。所以两者可以互相借用，不必分高下。",
 "这张表我刻意留了一行没有对应物。承认“这里可能真是不一样”，比硬凑一个映射更靠得住。",
 "这三句话不是免责声明，是方法的一部分。没有它们，任何相似性都可以被说成同源，任何类比都可以被说成证明。",
 "翻译练习的意义是：能翻译，说明这套语言真的通用；翻译不过去的地方，才是理论需要长进的地方。",
 "四组反例都指向同一件事：补救动作会破坏它自己所依赖的证据。这不是实现 bug，是矛盾在系统里的投影。",
 "我尽量说得不客气：这份报告的长处是敢把哲学命题降成可反证的缺口，短处是只走到问题定义就停了。评估一个想法，先看它有没有把自己最容易吹的那部分先证伪。",
 "三句话带走。最后一问留给作业：你手上那个问题，如果只能保留一个约束，你留哪一个？",
]

REPL = {
    "批准 8 / 查询 1 / 观察 TTL 2：先查再批被校验器拒绝": "批准 8 / 查询 1 / TTL 2：被校验器拒绝",
    "批准必须先拿到同一观察：旧模型下不可行": "批准须先拿到同一观察：旧模型不可行",
    "TTL 改 20 才可行：换了 fixture，不是延长真实保证": "TTL 改 20 才可行：换了 fixture",
    "等待变区间 1–10：旧模型直接拒绝区间输入": "等待变区间 1–10：旧模型拒绝区间",
}


def spc_of(run):
    rPr = run._r.find(qn("a:rPr"))
    return None if rPr is None else rPr.get("spc")


def clean(prs, extras):
    for idx, s in enumerate(prs.slides):
        if idx == 0:
            continue
        for sh in s.shapes:
            if sh.has_text_frame:
                top = sh.top / 914400.0
                if not (0.85 <= top <= 6.90):
                    continue
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.text in REPL:
                            r.text = REPL[r.text]
                        if spc_of(r):
                            continue
                        if r.font.size and r.font.size.pt < 16:
                            r.font.size = __import__("pptx.util", fromlist=["Pt"]).Pt(16)
        if idx - 1 < len(extras):
            tf = s.notes_slide.notes_text_frame
            tf.text = (tf.text.rstrip() + "\n\n" + extras[idx - 1]).strip()


def audit(path, extra):
    prs = Presentation(str(path))
    n = int(len(prs.slides._sldIdLst))
    body, allmin, table_min, notes = [], 99.0, 99.0, 0
    for idx, s in enumerate(prs.slides):
        notes += len(s.notes_slide.notes_text_frame.text.strip())
        for sh in s.shapes:
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        for p in c.text_frame.paragraphs:
                            for r in p.runs:
                                if r.font.size:
                                    table_min = min(table_min, r.font.size.pt)
                continue
            if not sh.has_text_frame:
                continue
            top = sh.top / 914400.0
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if not r.font.size:
                        continue
                    allmin = min(allmin, r.font.size.pt)
                    if idx == 0 or spc_of(r) or not (0.85 <= top <= 6.90):
                        continue
                    body.append(r.font.size.pt)
    return {
        "slides": n,
        "min_font": round(min(body), 1) if body else None,
        "min_all": round(allmin, 1),
        "table_min": round(table_min, 1),
        "notes": notes,
        "est_min": round(notes / 235.0, 1),
    }


OLD = {
    "TRIZa_AI4Science_15min.pptx": (426613, "2026-10-24 07:59"),
    "TRIZa_AI4Science_15min_v2.pptx": (401376, "2026-10-24 09:20"),
    "TRIZa_15min_handoff.zip": (22037056, "2026-10-24 15:57"),
}

if __name__ == "__main__":
    pa = D / "TRIZ_W2A_矛盾解剖_15min.pptx"
    pb = D / "TRIZ_W2B_盒子与矛盾论_15min.pptx"
    for f, ex in ((pa, EXTRA_A), (pb, EXTRA_B)):
        prs = Presentation(str(f))
        clean(prs, ex)
        prs.save(str(f))
    ra, rb = audit(pa, True), audit(pb, True)
    ok_est = all(10.0 <= x["est_min"] <= 17.0 for x in (ra, rb))
    import datetime
    unchanged = True
    detail = {}
    for name, (sz, mt) in OLD.items():
        f = D / name
        st = f.stat()
        same = st.st_size == sz and datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M") == mt
        detail[name] = "ok" if same else "CHANGED %d %s" % (st.st_size, datetime.datetime.fromtimestamp(st.st_mtime))
        unchanged = unchanged and same
    print(json.dumps({
        "decks": 2,
        "a_slides": ra["slides"],
        "b_slides": rb["slides"],
        "min_font": min(ra["min_font"] or 0, rb["min_font"] or 0),
        "speak_ok": ok_est,
        "old_unchanged": unchanged,
        "a": ra, "b": rb, "old": detail,
    }, ensure_ascii=False))