# -*- coding: utf-8 -*-
import copy, glob, os
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.text.text import _Paragraph

D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
OPT = os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版.pptx")
RED = RGBColor(0xB0, 0x3A, 0x2B)
INK = RGBColor(0x1A, 0x1D, 0x23)
GREY = RGBColor(0x56, 0x5C, 0x66)

prs = Presentation(OPT)
s6, s7, s10 = prs.slides[5], prs.slides[6], prs.slides[9]
by = lambda sl, nm: [x for x in sl.shapes if x.name == nm][0]

# ---------- s6: ① 补"什么叫通用工程参数" ----------
tb = by(s6, "TextBox 4")
tb.text_frame.paragraphs[0].runs[1].text = \
    "把问题翻译成 39（现代 48）个通用工程参数里的两个——参数就是可测量的量：重量、温度、速度、功耗、噪声"

# ---------- s7: 四个原理各加一句动作 ----------
tb7 = by(s7, "TextBox 4")
tf = tb7.text_frame
txBody = tf._txBody
tpl_b = copy.deepcopy(txBody.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p')[0])
tpl_d = copy.deepcopy(txBody.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p')[1])
for p in txBody.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
    txBody.remove(p)

LINES = [
    (tpl_b, [("▸  技术矛盾", True, RED, 19), ("：散热能力 ↑ ↔ 机身厚度 ↓", False, INK, 19)]),
    (tpl_b, [("▸  ", True, RED, 19), ("查矩阵 → 分割 / 抽取 / 动态化 / 中介物", False, INK, 19)]),
    (tpl_d, [("分割＝拆成互不影响的块　抽取＝把起作用的部件单独拿走", False, GREY, 16)]),
    (tpl_d, [("动态化＝让部件能变能适应　中介物＝中间加个传递载体", False, GREY, 16)]),
    (tpl_b, [("▸  ", True, RED, 19), ("落地：", False, INK, 19), ("热管", True, RED, 19),
             (" 把热从 CPU “搬”到边缘鳍片再吹走", False, INK, 19)]),
    (tpl_b, [("▸  ", True, RED, 19), ("本质：不在原地散热，而是", False, INK, 19),
             ("换一个位置", True, RED, 19), ("散热", False, INK, 19)]),
]
for tpl, runs in LINES:
    el = copy.deepcopy(tpl)
    for r in el.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}r'):
        el.remove(r)
    for erp in el.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr'):
        el.remove(erp)
    txBody.append(el)
    para = _Paragraph(el, tf)
    for text, bold, color, size in runs:
        run = para.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color

# ---------- s10: 五级表减负 ----------
tbl_shape = [x for x in s10.shapes if x.has_table][0]
tbl = tbl_shape.table
for i, w in enumerate([2.10, 4.00, 4.35, 1.16]):
    tbl.columns[i].width = Inches(w)
CELLS = {(1, 1): "本领域常规改进", (2, 1): "本行业已知方法", (3, 1): "引入行业外方法",
         (4, 1): "已公开原理首用", (5, 1): "新科学、新产业", (4, 2): "换主体，矛盾作废"}
for (ri, ci), txt in CELLS.items():
    tbl.cell(ri, ci).text_frame.paragraphs[0].runs[0].text = txt

# ---------- s10: 底栏第二句改成"换主体"口径 ----------
bar = by(s10, "Rectangle 5")
bar.text_frame.paragraphs[1].runs[0].text = \
    "“一次性消失”＝需求不变，换承担主体；老约束只属于老主体，它派生的一串矛盾一起作废。"

# ---------- 讲稿（notes）----------
N6 = ("矩阵是个查表工具。横轴是你要改善的参数，纵轴是会被恶化的参数，格子里是几条原理的编号。"
      "先说清楚什么叫通用工程参数——就是任何工程问题都能落到的那几十个可以测量的量：重量、温度、速度、功耗、噪声、可靠性。"
      "它不是“用户体验差”这种话，而是“响应时间从两秒降到两百毫秒”。四步：写成参数对、查格子、追问语义、收敛。"
      "第三步最容易偷懒——很多人拿到编号就去搜“分割是什么意思”，然后套模板。第四步是纪律：只挑一两条做可验证的方案，其余归档。"
      "还有一种情况：新领域参数根本填不进去，那就说查不到，不要硬凑。记住，它给的是方向，不是答案；不给公式，也不给阈值。")
N7 = ("第一个例子。笔记本散热要好，风扇要大、风道要厚；可用户要薄。这是技术矛盾。"
      "查矩阵，拿到四个原理，四个都翻译成动作：分割，把整体拆成互不影响的块；抽取，把真正起作用的那个部件单独拿出来；"
      "动态化，让部件能变、能适应工况；中介物，中间加一个传递的载体。"
      "真正落地的是抽取加中介物——热管：不在 CPU 原地散热，把热搬到机身边缘的鳍片上再吹走。"
      "鳍片就是散热片，很多片薄金属，为的是让吹过去的空气吃到更大的面积。"
      "注意，矛盾没有被折中，它是被换了位置。这是空间分离的一次具体使用。"
      "对比一下折中：风扇小一点、厚度加一点，再塞一层石墨片——噪音上来了，矛盾还在，只是被压住了。")
N10 = ("这一页是本讲的答案。阿奇舒勒按发明程度把专利分五级：一级不考虑矛盾，二级解决技术矛盾，三级解决物理矛盾，"
       "四级让矛盾一次性消失，五级是全新发现。比例是 32、45、18、4、不到 1。"
       "四级为什么叫一次性？关键在这句：约束不挂在需求上，挂在“谁来承担这个需求”上。"
       "白炽灯要亮，就得靠高温发热来发光——“亮必须热”这条约束，属于热辐射发光这个主体。"
       "LED 换了主体，用电致发光，亮度不再依赖高温，这条约束就整个作废。"
       "约束一走，由它派生的一串矛盾——效率、寿命、发热、体积——同时失效。这就是一次性：不是解掉一个，是一批不再成立。"
       "所以判断四级，问三句话：需求还是原来那个需求吗？承担需求的还是同一个对象吗？原来那条“因为什么所以什么”还成立吗？"
       "三句都答变了，才是四级。再补一句：主体换了，矛盾不是消失，是搬家。新主体有新约束，新矛盾马上来。所以一次性说的是这一批，不是从此没有矛盾。")
for sl, txt in ((s6, N6), (s7, N7), (s10, N10)):
    sl.notes_slide.notes_text_frame.text = txt

prs.save(OPT)
print("A 优化版 patched; notes:", len(N6), len(N7), len(N10))

# ---------- 页脚：研讨班 → 实战研讨班 ----------
TARGETS = [
    "TRIZ_W2A_矛盾解剖_15min.pptx",
    "TRIZ_W2A_矛盾解剖_15min_优化版.pptx",
    "TRIZ_W2A_矛盾解剖_15min_自动解说.pptx",
    "TRIZ_W2B_盒子与矛盾论_15min.pptx",
    "TRIZ_W2B_盒子与矛盾论_15min_自动解说.pptx",
]
OLD, NEW = "创新构思法研讨班", "创新构思法实战研讨班"
for name in TARGETS:
    p = os.path.join(D, name)
    prs2 = Presentation(p)
    n = 0
    for sl in prs2.slides:
        frames = [sh.text_frame for sh in sl.shapes if sh.has_text_frame]
        for sh in sl.shapes:
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        frames.append(c.text_frame)
        for f in frames:
            for para in f.paragraphs:
                for run in para.runs:
                    if OLD in run.text:
                        run.text = run.text.replace(OLD, NEW)
                        n += 1
    prs2.save(p)
    print(f"{name}: replaced {n} runs")