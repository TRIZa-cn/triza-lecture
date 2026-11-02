# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16")
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from pptx import Presentation
import _l16_spec_a as A, _l16_spec_a2 as A2, _l16_spec_b as B, _l16_spec_c as C
from _l16_openers import OPENERS

P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
A_DECK = os.path.join(P, "TRIZ_W2A_矛盾解剖_15min_优化版.pptx")
B_DECK = os.path.join(P, "TRIZ_W2B_盒子与矛盾论_15min.pptx")
GREY = RGBColor(0x55, 0x5F, 0x6B)
RED = RGBColor(0xB0, 0x3A, 0x2B)

def notes_of(path):
    prs = Presentation(path)
    return [s.notes_slide.notes_text_frame.text.strip() for s in prs.slides]

A_N = notes_of(A_DECK)
B_N = notes_of(B_DECK)
SPECS = {s["no"]: s for s in A.SPECS + A2.SPECS + B.SPECS + C.SPECS}

WARM = ("好，先做一段热身。这段不用管内容，照着读就行。TRIZ，读作 treez，中文也翻译成萃智，或者萃思；"
        "TRIZa，读作 treeza，就是在 TRIZ 后面加一个字母 a。常见的缩写还有 ARIZ、IFR、MPV、FOS，遇到了就按字母念。"
        "数字这边我们经常要说：39 个通用工程参数，48 个现代参数，40 条发明原理，76 条标准解；"
        "还有 32%、45%、18%、4%，以及不到 1%。做实验的时候还会说：检测延迟要低于 200 毫秒，误报率不超过 3%，"
        "样本量是 2040 条决策，训练集和测试集按 7 比 3 划分。"
        "最后一句长一点的：如果读者看完这段描述，仍然说不清你要解决的是哪两个参数在打架，那说明问题还没有被写成矛盾。")

SEGS = [("S00", "热身：术语、缩写、数字、长句", WARM)]
for i, n in enumerate(A_N, 1):
    SEGS.append((f"S{i:02d}", f"讲 1 · A 段 · 第 {i} 页", n))
SEGS.append(("S13", "讲 1 · B 段 · 第 1–2 页", B_N[0] + "\n\n" + B_N[1]))
SEGS.append(("S14", "讲 1 · B 段 · 第 3–4 页", B_N[2] + "\n\n" + B_N[3]))
groups = [(2, 3, 4), (5, 6, 7), (8, 9, 10), (11, 12, 13), (14, 15, 16)]
for gi, g in enumerate(groups, 15):
    txt = "\n\n".join(OPENERS[k] for k in g)
    SEGS.append((f"S{gi:02d}", f"讲 {g[0]}—{g[-1]} 开场白", txt))

def setfont(run, size, bold=False, color=None, name="微软雅黑"):
    run.font.size = Pt(size); run.bold = bold; run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if color: run.font.color.rgb = color

# ---------------- 1) 录音讲稿 ----------------
d = Document()
s = d.sections[0]
s.page_height, s.page_width = Cm(29.7), Cm(21.0)
s.left_margin = s.right_margin = Cm(2.2); s.top_margin = s.bottom_margin = Cm(2.0)

def para(doc, text, size=18, bold=False, before=0, after=6, color=None, spacing=1.6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    setfont(p.add_run(text), size, bold, color)
    return p

head = d.add_paragraph(); head.alignment = WD_ALIGN_PARAGRAPH.CENTER
setfont(head.add_run("《TRIZ 创新构思法实战研讨班》声音模型训练 · 录音讲稿（20 分钟版）"), 19, True)

total = sum(len(x[2]) for x in SEGS)
para(d, f"共 {len(SEGS)} 段，约 {total} 字，按 270 字/分钟约 {total/270:.1f} 分钟。"
        "内容＝A 讲全文 + B 讲前四页 + 讲 2–16 的开场白，录完既能当课件配音，也能当少样本语音克隆的微调素材。",
     11, color=GREY, after=12, spacing=1.3)

para(d, "一、录音前先看这 6 条", 14, True, before=8, after=6, color=RED)
for t in ["1. 环境：小房间，关空调、风扇、加湿器；窗帘、衣柜、床边的吸音效果通常比大厅好。手机静音。",
          "2. 麦克风：USB 电容麦最好；笔记本内置麦也能用，但机器要放在软垫上，别让风扇声进话筒。",
          "3. 口距：15～20 厘米，稍微偏一点对着嘴角，不要正对气流，避免“噗”“啪”的喷气声。",
          "4. 一次读完一段，段内不改口、不剪接。读错就整段重来——剪接缝会被模型学进去。",
          "5. 每段结束停 2 秒再停录，方便我自动切分；同一批录音不要换位置、换姿势。",
          "6. 全程不要降噪、不要压缩、不要混响、不要开自动增益。原始干声最好，后期我来做。"]:
    para(d, t, 11, spacing=1.4, after=3)

para(d, "二、录音参数", 14, True, before=8, after=6, color=RED)
rows = [("采样率 / 位深", "48 kHz / 24 bit（最低 44.1 kHz / 16 bit）"),
        ("声道", "单声道（Mono）"),
        ("格式", "WAV（不要 MP3、不要 m4a）"),
        ("电平峰值", "−6 dBFS 左右，不要爆到 0 dB"),
        ("语速", "每分钟 265～275 字；比上次的 258 略快一点，正常讲课节奏"),
        ("总时长", "约 20 分钟有效语音（微调建议 10 分钟以上，20 分钟更稳）")]
tb = d.add_table(rows=len(rows), cols=2); tb.style = "Table Grid"; tb.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(rows):
    for j, val in enumerate((k, v)):
        c = tb.cell(i, j); c.text = ""
        setfont(c.paragraphs[0].add_run(val), 11, (j == 0))

para(d, "三、命名与交付", 14, True, before=8, after=6, color=RED)
para(d, f"按下面段号命名，一段一个文件，共 {len(SEGS)} 个：voice_S00.wav … voice_S{len(SEGS)-1:02d}.wav", 11, spacing=1.4, after=3)
para(d, "全部放进同一个文件夹发我。段号不要改，我按段号对应课件页与讲次。", 11, spacing=1.4, after=3)

para(d, "四、朗读要求", 14, True, before=8, after=6, color=RED)
for t in ["像给同事讲，不要播音腔；句末自然降调。",
          "问句要真的升调——“对不对”“是不是”是真问题，不是口头禅。",
          "破折号停半拍，逗号几乎不停，句号停够。",
          "允许少量自然口语（吸气、轻微的“嗯”），但不要刻意加；不要笑场、不要清嗓子。",
          "关键句不要刻意加重。重音由内容决定，模型会学走你的刻意重音。",
          "TRIZ、TRIZa、ARIZ 的读法在 S00 里，先读一遍热身，再正式录。"]:
    para(d, t, 11, spacing=1.4, after=3)

d.add_page_break()
para(d, f"五、{len(SEGS)} 段讲稿（段号即文件名）", 14, True, after=10, color=RED)
lines = []
for sid, title, text in SEGS:
    n = len(text)
    para(d, f"{sid}　{title}", 12, True, before=8, after=2, color=GREY)
    para(d, f"约 {n} 字 · 约 {n/270*60:.0f} 秒 · voice_{sid}.wav", 9.5, color=GREY, after=4, spacing=1.2)
    for chunk in text.split("\n\n"):
        para(d, chunk, 17, after=8, spacing=1.7)
    lines.append(f"### {sid} {title} ({n}字, voice_{sid}.wav)\n{text}\n")
d.save(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音模型训练-录音讲稿.docx")
open(os.path.join(P, "_录音稿.txt"), "w", encoding="utf-8").write("\n".join(lines))

# ---------------- 2) 16 讲讲稿集 ----------------
g = Document()
gs = g.sections[0]
gs.page_height, gs.page_width = Cm(29.7), Cm(21.0)
gs.left_margin = gs.right_margin = Cm(2.4); gs.top_margin = gs.bottom_margin = Cm(2.2)

def gp(text, size=12, bold=False, before=0, after=4, color=None, spacing=1.45):
    p = g.add_paragraph()
    p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    setfont(p.add_run(text), size, bold, color)
    return p

h = g.add_paragraph(); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
setfont(h.add_run("TRIZ 创新构思法实战研讨班 · 16 讲讲稿集（v1）"), 20, True)
gp("每讲一页八页骨架：封面、困境、定义、流程、例子、表、边界、收束。讲稿写在每页下方，正文即口播稿。"
   "v1 每讲约 900–1100 字（3.5–4 分钟核心叙述），配合现场演示与学生难题研讨使用；若要撑满 15 分钟纯讲述，需按讲补厚到 3500–4000 字。",
   10.5, color=GREY, after=10, spacing=1.35)

def sec(title, pages, unit="页"):
    gp(title, 15, True, before=14, after=6, color=RED)
    tot = sum(len(x) for x in pages)
    gp(f"共 {len(pages)} {unit} · {tot} 字 · 约 {tot/270:.1f} 分钟", 9.5, color=GREY, after=6, spacing=1.2)
    for i, (head_, body) in enumerate(pages, 1):
        gp(f"{i:02d}　{head_}", 11, True, before=6, after=2)
        gp(body, 11.5, after=3)

gp("讲 1　矛盾解剖与约束观（A 段 + B 段，2 学时内容）", 15, True, before=8, after=6, color=RED)
gp("A 段（12 页）", 12, True, before=6, after=4)
for i, n in enumerate(A_N, 1):
    gp(f"A{i:02d}　第 {i} 页", 11, True, before=4, after=2)
    gp(n, 11.5, after=3)
gp("B 段（13 页）", 12, True, before=8, after=4)
for i, n in enumerate(B_N, 1):
    gp(f"B{i:02d}　第 {i} 页", 11, True, before=4, after=2)
    gp(n, 11.5, after=3)

for no in range(2, 17):
    sp = SPECS[no]
    pgs = [(p["title"], p.get("notes", "")) for p in sp["pages"]]
    sec(f"讲 {no:02d}　{sp['file']}", pgs)

g.save(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-16讲讲稿集-v1.docx")
print("voice docx:", total, "chars", len(SEGS), "segs", round(total/270, 2), "min")
print("讲稿集 ok")

# ---------------- 3) 语速参数 -13% ----------------
q = os.path.join(P, "_tts_opt.py")
t = open(q, encoding="utf-8").read()
t2 = t.replace('RATE = "-16%"', 'RATE = "-13%"')
if t2 != t:
    open(q, "w", encoding="utf-8").write(t2)
print("TTS rate ->", [l for l in t2.splitlines() if l.startswith("RATE")])