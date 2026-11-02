# -*- coding: utf-8 -*-
import os, glob, json
D = r"D:\doc\ZGCA\Lectures\TRIZ"
SK = os.path.join(D, "PPT", "TRIZa-Lecture-SKILL.md")
t = open(SK, encoding="utf-8").read()

ADD = """

## 校准补充（2026-10 修订，据实战数据）

**1. 目标语速 270 字/分钟 → rate ≈ −13%。**
同一音色（zh-CN-YunyangNeural）实测：−16% → 258.6 字/分；base ≈ 310。rate 与语速近似线性，故 270 字/分对应 rate ≈ −13%。仍按老规矩：先合成 150 字样本、测时长、算字/分，再定 rate，不要直接抄这张表。

**2. 跨行业例子必须配一句搭桥。**
电池、热管这类硬约束例子对软件/安全听众会掉线。讲完例子立刻用一句把结构接回听众经验：
> "电池这个例子是经典的 TRIZ 案例，虽然离软件的场景远，但问题结构是一样的——你看，做漏洞扫描，追得深就跑得慢，对不对？"
搭桥句写在**讲稿**里，不占页面。它不改变例子，只改变听众的接入点。

**3. 交付时必须标明是"骨架版"还是"定稿版"。**
8 页骨架 + 约 1000 字讲稿 ≈ 4 分钟核心叙述，只够研讨班"15 分钟理论段"的一半（15 分钟纯讲述要 3500–4000 字）。写交付说明时直接写"v1 骨架，3.5–4 分钟"，不要让 1000 字冒充 15 分钟——这是本项目已经踩过的坑。

**4. 表格 ≤6 行是硬约束，条目多就改双栏。**
需要列 8–10 条时，不要突破行数，改成两组并列（"编号 / 内容 / 编号 / 内容"），行数减半、内容不删。实测 10 行表在 15pt 下仍可读，但违反信息预算，且与全课其他页不一致。
"""
if "校准补充（2026-10 修订" not in t:
    t = t.rstrip() + ADD
    open(SK, "w", encoding="utf-8").write(t)
    print("SKILL patched,", len(t), "bytes")
else:
    print("SKILL already patched")

# 清理测试残留
tmp = os.path.join(D, "PPT", "_tmp_voicetest")
if os.path.isdir(tmp):
    for f in glob.glob(os.path.join(tmp, "*")):
        os.remove(f)
    os.rmdir(tmp)
    print("tmp cleaned")

# 清单
files = sorted(glob.glob(os.path.join(D, "PPT", "L16", "*.pptx")))
inv = {"decks": len(files), "pages": 0, "bytes": 0}
from pptx import Presentation
for f in files:
    p = Presentation(f); inv["pages"] += len(p.slides); inv["bytes"] += os.path.getsize(f)
print(json.dumps(inv, ensure_ascii=False))
for f in files:
    print("  ", os.path.basename(f), round(os.path.getsize(f)/1024, 1), "KB")
print("\ndocs:")
for f in ["TRIZa-16讲讲稿集-v1.docx", "TRIZa-声音模型训练-录音讲稿.docx", "TRIZa-VoiceSpeed.py"]:
    p = os.path.join(D, f)
    print("  ", f, round(os.path.getsize(p)/1024, 1), "KB" if os.path.isfile(p) else "MISSING")