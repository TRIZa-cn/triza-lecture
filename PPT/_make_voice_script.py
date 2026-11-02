# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from pptx import Presentation

DECK = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
OUT_DOCX = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音模型训练-录音讲稿.docx"
OUT_TXT = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_录音稿.txt"

prs = Presentation(DECK)
notes = [s.notes_slide.notes_text_frame.text.strip() for s in prs.slides]

WARM = ("好，先做一段热身。这段不用管内容，照着读就行。TRIZ，读作 treez，中文也翻译成萃智，或者萃思；"
        "TRIZa，读作 treeza，就是在 TRIZ 后面加一个字母 a。常见的缩写还有 ARIZ、IFR、MPV、FOS，遇到了就按字母念。"
        "数字这边我们经常要说：39 个通用工程参数，48 个现代参数，40 条发明原理，76 条标准解；"
        "还有 32%、45%、18%、4%，以及不到 1%。做实验的时候还会说：检测延迟要低于 200 毫秒，误报率不超过 3%，"
        "样本量是 2040 条决策，训练集和测试集按 7 比 3 划分。"
        "最后一句长一点的：如果读者看完这段描述，仍然说不清你要解决的是哪两个参数在打架，那说明问题还没有被写成矛盾。")

SEGS = [("S00", "热身：术语、缩写、数字、长句", WARM, "voice_S00.wav")]
for i, n in enumerate(notes, 1):
    SEGS.append((f"S{i:02d}", f"课件第 {i} 页讲稿", n, f"voice_S{i:02d}.wav"))

d = Document()
sec = d.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
sec.left_margin = sec.right_margin = Cm(2.2)
sec.top_margin = sec.bottom_margin = Cm(2.0)

def para(text, size=18, bold=False, before=0, after=6, color=None, indent=0, spacing=1.6):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.font.name = "微软雅黑"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    if color:
        r.font.color.rgb = color
    return p

GREY = RGBColor(0x55, 0x5F, 0x6B)
RED = RGBColor(0xB0, 0x3A, 0x2B)

t = d.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("《TRIZ 创新构思法实战研讨班》声音模型训练 · 录音讲稿")
r.font.size = Pt(20); r.bold = True; r.font.name = "微软雅黑"
r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

total_chars = sum(len(s[2]) for s in SEGS)
para(f"共 {len(SEGS)} 段，约 {total_chars} 字，按 260 字/分钟约 {total_chars/260:.1f} 分钟。"
     "全文就是 A 讲课件 12 页的解说稿，录音可直接当作课件配音，也可作为少样本语音克隆（GPT-SoVITS 等）的微调素材。",
     size=11, color=GREY, after=14, spacing=1.3)

para("一、录音前先看这 6 条", size=14, bold=True, before=10, after=6, color=RED)
for s in [
    "1. 环境：小房间，关空调、风扇、加湿器；不要在大厅或瓷砖墙面前录。衣柜前、挂毯前、床边的效果通常更好。手机静音，门关好。",
    "2. 麦克风：USB 电容麦最好；笔记本内置麦也能用，但要把整机放在软垫上，别让风扇声进话筒。",
    "3. 口距：15～20 厘米，稍微偏一点对着嘴角，不要正对气流，避免“噗”“啪”的喷气声。",
    "4. 一次读完一段，段内不改口、不剪接。读错就整段重来——剪接过的音频会让模型学到拼接缝。",
    "5. 每段结束停 2 秒再按停止，方便后面自动切分；不要在同一次录音里换位置、换姿势。",
    "6. 全程不要降噪、不要压缩、不要混响、不要开自动增益。原始干声最好，后期处理由我来做。",
]:
    para(s, size=11, spacing=1.4, after=3)

para("二、录音参数", size=14, bold=True, before=10, after=6, color=RED)
rows = [("采样率 / 位深", "48 kHz / 24 bit（最低 44.1 kHz / 16 bit）"),
        ("声道", "单声道（Mono）"),
        ("格式", "WAV（不要 MP3、不要 m4a）"),
        ("电平峰值", "−6 dBFS 左右，不要爆到 0 dB（爆音不可修复）"),
        ("语速", "每分钟 250～270 字，正常讲课语速，不要刻意放慢"),
        ("总时长", "有效语音 ≥ 10 分钟（1 分钟只够零样本，做微调建议 10 分钟以上）")]
tb = d.add_table(rows=len(rows), cols=2)
tb.style = "Table Grid"
tb.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(rows):
    for j, val in enumerate((k, v)):
        c = tb.cell(i, j)
        c.text = ""
        rr = c.paragraphs[0].add_run(val)
        rr.font.size = Pt(11)
        rr.font.name = "微软雅黑"
        rr._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        if j == 0:
            rr.bold = True

para("三、命名与交付", size=14, bold=True, before=10, after=6, color=RED)
para("按下面的段号命名，一段一个文件，全部放进同一个文件夹发我。段号不要改，我按段号对应课件页做替换。",
     size=11, spacing=1.4, after=3)
para("文件名：voice_S00.wav … voice_S12.wav（共 13 个）", size=11, bold=True, spacing=1.4, after=3)

para("四、朗读要求", size=14, bold=True, before=10, after=6, color=RED)
for s in [
    "像给同事讲，不要播音腔。这是课堂口播，不是新闻联播。句末自然降调。",
    "问句要真的升调（“你有没有遇到过这种时候？”——它是真问题，不是修辞）。",
    "破折号处停半拍，不是停一整句；“，”处几乎不停，“。”处停够。",
    "允许少量自然口语（吸气、轻微的“嗯”），但不要刻意加；不要笑场、不要清嗓子。",
    "关键句不要刻意加重。重音由内容决定，模型会学到你刻意的重音，以后每句都会重。",
    "读不准的词（TRIZ、ARIZ、MPV）：S00 已经给出发音，先读一遍热身，再正式录。",
]:
    para(s, size=11, spacing=1.4, after=3)

d.add_page_break()
para("五、13 段讲稿（照此朗读，段号即文件名序号）", size=14, bold=True, before=0, after=12, color=RED)

txt_lines = []
for sid, title, text, fname in SEGS:
    n = len(text)
    sec_s = n / 260 * 60
    para(f"{sid}　{title}", size=12, bold=True, before=10, after=2, color=GREY)
    para(f"约 {n} 字 · 约 {sec_s:.0f} 秒 · {fname}", size=9.5, color=GREY, after=4, spacing=1.2)
    para(text, size=18, after=14, spacing=1.7)
    txt_lines.append(f"### {sid} {title} ({n}字, {fname})\n{text}\n")

d.save(OUT_DOCX)
open(OUT_TXT, "w", encoding="utf-8").write("\n".join(txt_lines))
print("docx ok", total_chars, "chars,", len(SEGS), "segs,", round(total_chars/260, 2), "min")