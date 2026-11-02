# -*- coding: utf-8 -*-
import os, shutil
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZa-Lecture-SKILL.md"
t = open(p, encoding="utf-8").read()
shutil.copyfile(p, p + ".bak-20261029")

OLD1 = "硬约束：正文 ≥16pt（投影到 4 米外仍可读），页脚与表格可到 13.5pt 但要说明是辅助元素；一页不超过 6 个要点；表格不超过 6 行；标题 30pt 左右，kicker（\"03 · 关键一步\"）11pt。"
NEW1 = ("硬约束：**正文 ≥20pt**（投影到 4 米外仍可读）。放不下就删内容或合并，**不许缩小字号硬塞**。"
        "只有三类版式标记可以低于 20pt：kicker 14pt、页脚 12pt、页码 12pt，每页合计不超过 4 处。"
        "一页不超过 6 个要点；表格不超过 6 行；标题 30pt 左右。")
assert OLD1 in t, "规则 101 行未命中"
t = t.replace(OLD1, NEW1)

OLD2 = "| 7 | 正文最小字号 | ≥16pt |"
NEW2 = "| 7 | 正文最小字号 | ≥20pt（版式标记除外：kicker 14 / 页脚 12 / 页码 12）|"
assert OLD2 in t
t = t.replace(OLD2, NEW2)

SECTION = """

---

## 校准补充（2026-10 第二次修订）

### 1. 字号下限 20pt，放不下就删

- **规则**：正文（要点、表格、判据框、说明文字）**≥20pt**。
- **违反时的正确动作**：删掉一条、合并两条、拆成两页——**而不是缩到 18pt**。字号是被内容驱动的，不是被版面驱动的。
- **豁免清单**（仅此四类，每页合计 ≤4 处）：kicker 14pt、页脚 12pt、页码 12pt、封面署名 12pt。
- **理由**：16–18pt 的中文在 1080p 录制与投影上都会糊；实测某讲 16pt 说明文字，回放时需放大才能读。
- **副作用要说清**：≥20pt 之后单页信息量下降约三成，必须靠**页数**补回来。13 页 20pt 的信息量约等于 16–17 页 16pt。改字号不是零成本，是拿页数换可读性。

### 2. 讲稿要有连接词

**规则**：每 100 字至少 3 个连接词。念出来不别扭，靠的就是它们。

| 作用 | 常用词 |
|---|---|
| 顺承 | 所以、于是、因此、这样一来 |
| 转折 | 但是、可是、而、不过、反过来 |
| 因果 | 因为、原因是、这就解释了 |
| 递进 | 而且、更重要的是、再说、退一步说 |
| 举例 | 比如、你看、举个例子 |
| 收束 | 所以记住、一句话、说白了 |

反例（书面语，念出来发干）：
> 改善 A 会恶化 B。矛盾需要被形式化。形式化后可查矩阵。

正例：
> 改善 A 就会恶化 B。**所以**矛盾得先写成一句话。**而**写成一句话之后，你才能查矩阵——**但是**注意，矩阵只给方向。

硬要求：连续三个短句都没有连接词，就说明这份讲稿是写给人看的，不是给人说的。

### 3. 页脚统一格式

全课程统一为：

```
TRIZ 创新构思法实战研讨班  ·  第 N 次 Workshop
```

- `·` 两侧各两个空格；
- **不再写「A 讲 / B 讲」**——场次信息由页脚的「第 N 次」与标题共同承担；
- 封面不写页脚，用顶部 kicker 标识课程与场次；
- 新增课件一律沿用，不另起一套。
"""
if "校准补充（2026-10 第二次修订）" not in t:
    t = t.rstrip() + SECTION
open(p, "w", encoding="utf-8").write(t)
print("SKILL updated:", round(len(t.encode("utf-8"))/1024, 1), "KB")

# ---------- 播讲稿 docx ----------
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from pptx import Presentation

prs = Presentation(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx")
pages = []
for i, s in enumerate(prs.slides, 1):
    n = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
    title = ""
    for sh in s.shapes:
        if sh.has_text_frame and sh.name == "TextBox 2" and len(sh.text_frame.text) > 3:
            title = sh.text_frame.text.strip(); break
    if not title and i == 1:
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip().startswith("矛盾的解剖学"):
                title = sh.text_frame.text.strip(); break
    pages.append((i, title, n))
total = sum(len(x[2]) for x in pages)

g = Document(); gs = g.sections[0]
gs.page_height, gs.page_width = Cm(29.7), Cm(21.0)
gs.left_margin = gs.right_margin = Cm(2.4); gs.top_margin = gs.bottom_margin = Cm(2.2)
GREY = RGBColor(0x55, 0x5F, 0x6B); RED = RGBColor(0xB0, 0x3A, 0x2B)

def setf(r, size, bold=False, color=None):
    r.font.size = Pt(size); r.bold = bold; r.font.name = "微软雅黑"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    if color: r.font.color.rgb = color

def P(txt, size=12, bold=False, color=None, before=0, after=5, spacing=1.5, align=None):
    p = g.add_paragraph()
    p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    if align is not None: p.alignment = align
    setf(p.add_run(txt), size, bold, color)
    return p

h = g.add_paragraph(); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
setf(h.add_run("TRIZ 创新构思法实战研讨班 · 第 1 次 Workshop"), 20, True)
h2 = g.add_paragraph(); h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
setf(h2.add_run("《矛盾的解剖学》播讲稿"), 15, True, RED)
P("16 页 · 全篇 %d 字 · 按 250 字/分约 %.1f 分钟，按 270 字/分约 %.1f 分钟" % (
    total, total/250, total/270), 10.5, color=GREY, after=4, spacing=1.3)
P("念法：按页顺序念，页与页之间停 1–2 秒。每页第一句都是承上句，不要跳过——它是给听众的落脚点。",
  10.5, color=GREY, after=14, spacing=1.3)

for i, title, note in pages:
    P("%02d　%s" % (i, title or "封面"), 13, True, RED, before=12, after=3)
    P(note, 12, after=4)

OUT = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-W1-播讲稿.docx"
g.save(OUT)
print("docx:", OUT, round(os.path.getsize(OUT)/1024, 1), "KB |", len(pages), "页 |", total, "字")