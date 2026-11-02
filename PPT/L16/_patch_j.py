# -*- coding: utf-8 -*-
import shutil
p = r"D:\doc\ZGCLab\triza\project\TRIZa_网安参数与矛盾库_设计稿.md"
t = open(p, encoding="utf-8").read()
shutil.copyfile(p, p + ".bak-J-20261029")

ANCHOR = "| `design/akbank/seed/contradictions_technical.seed.json` | 技术矛盾种子（实测生成；含机器遗忘空格对与 4 格 proposed_fill） | **41** |"
NEW_ROWS = ANCHOR + "\n" + "\n".join([
 "| `design/akbank/schema/constraint_candidate.schema.json` | 约束对（CC）记录模式；六个字段与 `contradiction.schema.json` 的 `physical` 同构，可字段拷贝提升为 PC | — |",
 "| `design/akbank/seed/constraint_pairs.seed.json` | 约束对种子（CC-001 已对齐 PC-01；CC-002 待落库；另收 3 条 R1/R2/R4 反例） | **2** |",
])
assert t.count(ANCHOR) == 1, f"锚点匹配 {t.count(ANCHOR)} 次"
t = t.replace(ANCHOR, NEW_ROWS)

NOTE_ANCHOR = "3. 任何参数定义/原理集修改**必须新增版本**，不得原地覆盖。"
NOTE_NEW = NOTE_ANCHOR + "\n4. 约束对提升为物理矛盾时（`CC → PC`），`promote` **只做字段拷贝、不做再翻译**；且必须校验 `pcvt.verdict == valid_physical`，否则只能落成技术矛盾。契约见 `design/akbank/TRIZa_约束生成器_输入输出设计.md`。"
assert t.count(NOTE_ANCHOR) == 1
t = t.replace(NOTE_ANCHOR, NOTE_NEW)

open(p, "w", encoding="utf-8").write(t)
print("saved, bytes:", len(t.encode("utf-8")))

lines = t.splitlines()
start = next(i for i, l in enumerate(lines) if l.startswith("## J"))
end = next((i for i, l in enumerate(lines) if i > start and l.startswith("## ")), len(lines))
print("---- J 段（%d–%d 行）----" % (start + 1, end))
for i in range(start, end):
    print(f"{i+1:4d}| {lines[i]}")