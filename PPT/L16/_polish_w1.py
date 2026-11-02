# -*- coding: utf-8 -*-
from pptx import Presentation
import os
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
prs = Presentation(P)

N = {
8: "书里用大量案例说明一件事：没有规则和约束，人反而发挥不出最大的创造力。很多时候，限制倒逼出自由，而表面的自由反而是枷锁。"
   "而书里对我们最有用的，是怎么造一个能激发创造力的新盒子，这个方法叫配对约束。"
   "它其实是认知心理学家帕特里夏·斯托克斯（Patricia D. Stokes）提出的，两个动作，缺一不可。"
   "第一个是排除：主动禁掉你最擅长、最顺手的那条路，切断最省力路径。第二个是引导：同时给一条具体、可执行的新方向。"
   "那为什么必须成对？因为只排除，人会无所适从——就像很多安全规定不接地气，人本能就会逃避、绕过，或者假装遵守；"
   "而只引导，旧习惯会不断把你拉回老路。所以你看，排除是刹车，引导是方向盘，缺一个就开不动。",
9: "那往哪设约束？有四个维度，说的都是「改哪一段」。不是改难度，而是改入口。"
   "第一，目标约束，改「做什么」：莫奈不还原客观世界，改成画我这一刻看到的光。"
   "第二，资源约束，改「用什么」：布伦屋只给小预算，逼导演在叙事张力上下功夫。"
   "第三，方法约束，改「怎么做」：村上春树写《且听风吟》，先用英语写，因为词汇不够，只能写短句，再译回日语。"
   "第四，主体约束，改「核心对象」：戴森给风扇下了一条硬规矩，不许有扇叶。"
   "所以记住，这四个维度动的是入口，不是难度。",
10: "刚才讲约束有利于创新，这不算跑题，但是我们现在要回到 TRIZ 主线，对一下表。"
    "TRIZ 的假设是：矛盾本来就在系统里，你的任务是把它提取出来。而这本书的主张正好相反：如果没有矛盾，你就造一个。"
    "听起来是两个立场，其实是同一个机制的两种用法。"
    "看苏斯博士跟出版人打的那个赌：只用五十个词写一本书，结果有了《绿鸡蛋和火腿》，成了儿童文学里最有名的作品之一——"
    "因为他被迫把所有精力都花在节奏、韵律和排列组合上，每个词都被用到了极致。"
    "用 TRIZ 的语言重述一遍：方法约束排除了长句和生僻词，而目标约束又要求故事仍然完整有趣。"
    "两条合起来就是一个物理矛盾：同一本书既要极简，又要完整。而解法，是把复杂度转移到韵律和重复结构上。"
    "所以这就是分离——不是巧合，是同一个机制。",
12: "有了矛盾以后还没完，接下来要找解矛盾的方法。方法之一就是矩阵，它是个查表工具。"
    "横轴是你想改善的参数，纵轴是会被恶化的参数，格子里是几条原理的编号。"
    "先说清楚什么叫通用工程参数——就是任何工程问题都能落到的那几十个可以测量的量：重量、温度、速度、功耗、噪声、可靠性。"
    "所以它不是你嘴上说的「用户体验差」，而是「响应时间从两秒降到两百毫秒」。"
    "四步：写成参数对、查格子、追问语义、收敛。但是这里有个必须说清的边界：矩阵给的是方向，不是答案。"
    "它告诉你该往哪一类改动上试，却不给公式、不给参数、不给阈值。",
}
for idx, text in N.items():
    prs.slides[idx - 1].notes_slide.notes_text_frame.text = text
prs.save(P)
print("notes 已更新:", sorted(N))

# ================= 校验 =================
chk = Presentation(P)
FOOT = "TRIZ 创新构思法实战研讨班  ·  第 1 次 Workshop"
small, foot_ok, tot = [], 0, 0
for i, s in enumerate(chk.slides, 1):
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    tot += len(n)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == FOOT: foot_ok += 1
        frames = [sh.text_frame] if sh.has_text_frame else []
        if sh.has_table:
            frames += [c.text_frame for r in sh.table.rows for c in r.cells]
        for tf in frames:
            for p in tf.paragraphs:
                for r in p.runs:
                    if not r.text.strip(): continue
                    sz = r.font.size.pt if r.font.size else None
                    if sz is None or sz < 20: small.append((i, sz, r.text[:26]))
print("\n页数:", len(chk.slides), "| 页脚正确:", foot_ok, "| notes:", tot, "字 →",
      round(tot/247.6, 2), "分钟 @247.6字/分")
from collections import Counter
c = Counter((sz, t) for _, sz, t in small)
print("\n<20pt 的文字：")
for (sz, t), k in sorted(c.items(), key=lambda x: (x[0][0] is None, x[0][0])):
    print(f"  {sz}pt x{k}  「{t}」")
print("\n超出页面下边界的形状（top+height > 7.45）：")
bad = 0
for i, s in enumerate(chk.slides, 1):
    for sh in s.shapes:
        if sh.height is None or sh.top is None: continue
        bot = (sh.top + sh.height) / 914400
        if bot > 7.45:
            print(f"  第{i}页 {sh.name} 底边 {bot:.2f}"); bad += 1
print("  无" if bad == 0 else f"  共 {bad} 处")