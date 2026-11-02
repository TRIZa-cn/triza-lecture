# -*- coding: utf-8 -*-
import json, os
D = r"D:\doc\ZGCLab\triza\project\design\akbank"
sch = json.load(open(os.path.join(D, "schema", "constraint_candidate.schema.json"), encoding="utf-8"))
cs  = json.load(open(os.path.join(D, "schema", "contradiction.schema.json"), encoding="utf-8"))
sd  = json.load(open(os.path.join(D, "seed", "constraint_pairs.seed.json"), encoding="utf-8"))
ps  = json.load(open(os.path.join(D, "seed", "parameters.seed.json"), encoding="utf-8"))
pids = {x["id"] for x in ps["items"]}
phys = set(cs["definitions"]["physical"]["required"])
print("CC required 覆盖 physical required:", phys <= set(sch["required"]), "| 缺:", phys - set(sch["required"]))
print("CC 多余字段(生成器元数据):", sorted(set(sch["required"]) - phys))
bad = [x["id"] for x in sd["items"] if x["parameter"] not in pids]
print("CC 引用的参数是否都在 25 条里:", "全部存在" if not bad else bad)
for x in sd["items"]:
    ok = all(k in x for k in sch["required"])
    v = x["pcvt"]["verdict"]
    print(f"  {x['id']} dim={x['dimension']} param={x['parameter']} verdict={v} 必填齐={ok} "
          f"measurable={'有' if x['guidance']['measurable'] else '空'} abandon={'有' if x['abandon_condition'] else '空'}")