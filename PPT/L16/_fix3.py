# -*- coding: utf-8 -*-
import json, os, shutil
from pptx import Presentation

# ---------- 1) A 讲：按位置重排页脚 ----------
A = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(A)
fixed = 0
for idx, sl in enumerate(prs.slides, 1):
    if idx == 1:
        continue
    want = "%02d" % idx
    for sh in sl.shapes:
        if sh.has_text_frame and sh.top / 914400 > 6.9 and sh.text_frame.text.strip().isdigit():
            if sh.text_frame.text.strip() != want:
                sh.text_frame.paragraphs[0].runs[0].text = want
                fixed += 1
prs.save(A)
print("footers fixed:", fixed)

chk = Presentation(A)
print("A 讲:", len(chk.slides), "页 |", sum(len(x.notes_slide.notes_text_frame.text) for x in chk.slides), "字")
for i, sl in enumerate(chk.slides, 1):
    ti = [sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.name == "TextBox 2"]
    fn = [sh.text_frame.text.strip() for sh in sl.shapes
          if sh.has_text_frame and sh.top / 914400 > 6.9 and sh.text_frame.text.strip().isdigit()]
    print(f"  {i:2d} 页脚={fn[0] if fn else '-':<3} {(ti[0] if ti else '(封面)')[:34]}")

# ---------- 2) B 讲：开场补一句衔接 ----------
BP = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2B_盒子与矛盾论_15min.pptx"
bp = Presentation(BP)
n1 = bp.slides[0].notes_slide.notes_text_frame.text
ADD = "另外提前说一句：这一讲后面要讲的「补救动作毁掉它依赖的证据」，其实就是 A 讲那个审批问题的下游——同一套系统，同一个矛盾换了个面。"
if ADD not in n1:
    bp.slides[0].notes_slide.notes_text_frame.text = n1.rstrip() + ADD
    bp.save(BP)
print("B 讲 notes1:", len(n1), "->", len(Presentation(BP).slides[0].notes_slide.notes_text_frame.text))

# ---------- 3) AKBank 种子：新增 SEC-P25 ----------
SEED = r"D:\doc\ZGCLab\triza\project\design\akbank\seed\parameters.seed.json"
shutil.copyfile(SEED, SEED + ".bak-20261029")
d = json.load(open(SEED, encoding="utf-8"))
print("\nseed 改前 count:", d.get("count"), "| items:", len(d["items"]))
print("claim_boundary 类型:", type(d.get("claim_boundary")).__name__)
print("claim_boundary:", json.dumps(d.get("claim_boundary"), ensure_ascii=False)[:300])

P25 = {
    "id": "SEC-P25",
    "name": "操作可逆性",
    "domain": "security",
    "status": "proposed",
    "definition": "一次操作造成的影响能否在承诺时限内被完整撤回",
    "metric": {
        "method": "对受控操作注入一次错误动作，测量撤回成功率与撤回时延；可逆 = 撤回成功率≥承诺阈值 且 时延≤承诺上限",
        "unit": "布尔 + 秒",
        "direction": "increase_good",
        "range": [0, None]
    },
    "classic_mapping": {
        "params": [],
        "note": "经典39与Mann均无对应；最接近的 #34（抛弃与再生）语义偏离"
    },
    "matrix_addressability": {
        "semantic_boundary": True,
        "gradable": True,
        "has_partner": True,
        "verdict": "addressable",
        "partners": "SEC-P13, SEC-P14, SEC-P05"
    },
    "evidence_level": "D",
    "sources": ["A 讲第一讲三层写法讨论（2026-10-29）", "本稿 §B.2 参数清单缺口说明"]
}
if not any(x["id"] == "SEC-P25" for x in d["items"]):
    d["items"].append(P25)
    d["count"] = len(d["items"])
    json.dump(d, open(SEED, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("added SEC-P25; count ->", d["count"])
else:
    print("SEC-P25 已存在")

back = json.load(open(SEED, encoding="utf-8"))
print("回读校验: count =", back["count"], "| items =", len(back["items"]),
      "| 末条 =", back["items"][-1]["id"], back["items"][-1]["name"])
missing = [k for k in ["id","name","domain","status","definition","metric","classic_mapping",
                       "matrix_addressability","evidence_level"] if k not in back["items"][-1]]
print("必填字段缺失:", missing or "无")