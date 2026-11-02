# -*- coding: utf-8 -*-
import json
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
raw = (D / "_audit_A.json").read_bytes()
if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
    txt = raw.decode("utf-16")
else:
    txt = raw.decode("utf-8-sig")
d = json.loads(txt)
g = d["grades"]; t = d["terms"]

rows = "\n".join(
    f"| {r['slide']} | {r['chars']} | {r['bullets']} | {r['table_rows'] or '-'} | "
    f"{r['notes_chars']} | {r['notes_sec']:.0f} | {r['avg_sent']} |"
    for r in d["per_slide"])

un = "\n".join(f"| {x['name']} | {x['slide']} | {x['why']} |" for x in t["unexplained"])

md = f"""# A 讲优化版 · 课件评估报告

对象：`TRIZ_W2A_矛盾解剖_15min_优化版.pptx`（{d['slides']} 页）
口径：`TRIZa-Lecture-SKILL.md` 的五档判定，按维度分开给，总体取最差维度。
复算命令：`python TRIZa-Lecture-audit.py 该文件 --cpm 258.6 --target-min 15 --terms _terms_A.json`

## 一句话结论

**总体 D，但不是"做得差"，是"时长差得远"。** 结构与语言都在 A/B 档；真正要动的是两处：8 个专业名词没解释（U={t['U']:.0%}，C 档），以及 {d['est_minutes']:.2f} 分钟的讲稿要撑 15 分钟（偏差 {d['delta_min']:+.2f} 分钟，D 档）。

## 四个维度

| 维度 | 档位 | 实测 | 说明 |
|---|---|---|---|
| 术语（U） | {g['terms']} | U={t['U']:.0%}（{t['n_unexplained']}/{t['n_need_explain']}） | LV≥2 共 {t['n_need_explain']} 个，{t['n_unexplained']} 个首次出现时没解释 |
| 句子 | {g['sentence']} | 平均 {d['notes_avg_sent']} 字，最长 {d['notes_max_sent']} 字 | 口播节奏正常；最长那句在备注里，念出来会憋气 |
| 密度 | {g['density']} | 每页 78–244 字，要点 ≤4，表格 ≤6 行 | 整体偏轻，不是偏重；第 9 页最重 |
| 时长 | {g['time']} | {d['est_minutes']:.2f} min vs 15 min | 缺 {abs(d['delta_min']):.1f} 分钟，需加内容或改标称时长 |

## 逐页数据

| 页 | 正文字数 | 要点 | 表格行 | 讲稿字数 | 约合秒 | 平均句长 |
|---|---|---|---|---|---|---|
{rows}

合计：正文 {d['chars_slides']} 字，讲稿 {d['notes_chars']} 字，问句 {d['notes_questions']} 个，>45 字长句 {d['notes_long_sents']} 句，最小字号 {d['min_font_body']:.0f}pt（页眉 kicker），表格 {d['min_font_table']:.1f}pt。

## 术语账：8 个没解释的词

{t['n_need_explain']} 个 LV≥2 词里，16 个已在首次出现处给了操作性解释（技术矛盾、物理矛盾、管理矛盾、矛盾矩阵、四类分离、五问法、发明等级、整体油箱、热管、召回、误报、阈值），8 个没有：

| 词 | 首次出现 | 为什么算没解释 |
|---|---|---|
{un}

这 8 个里有 6 个是"名词当已知量用"：鳍片、风道、抽取、分割、动态化、中介物。它们不影响主线，但会让跨学科听众在某一句上卡住半秒。第 9 页的置信度同理。

## 修改建议（按性价比排序）

1. **第 7 页给"分割/抽取/动态化/中介物"各加四个字的动作**（如"分割＝把整体拆成互不影响的块"）。四个词一次补齐，U 从 33% 降到 17%，档位 C→B。
2. **第 6 页补一句"什么叫通用工程参数"**（"把任何工程问题都归到 39 个可测量的量上，比如重量、温度、速度"）。这是矩阵那页唯一的门槛。
3. **第 9 页把 6 行数据块压成一句结论**（"同一批 2040 个决策，三种策略的取舍事先声明，不是排名"），细节留讲稿。
4. **时长二选一**：要么加 6 分钟内容（建议加在第 5 页追问链、第 8 页四把刀的对比、第 10 页四级的例子），要么把标称改成"9 分钟 + 讨论"。
5. 第 12 页"下一讲 B"那句是唯一跨讲依赖，若 B 不讲了要删。

## 局限

- 没有试讲，没有听众测试，以上全部是文本代理指标，不等于"听众听懂了"。
- 术语分级是人工判定，脚本只做统计与定位；换一批听众，U 会变。
- 脚本无法区分正文与页眉字号，11pt 是页眉 kicker 的量测值，正文实际字号需目视确认。
- 未评估内容正确性、引用准确性与讲者现场表达。
"""
out = D / "TRIZa-Lecture-评估-A讲优化版.md"
out.write_text(md, encoding="utf-8")
print("report", out.stat().st_size)