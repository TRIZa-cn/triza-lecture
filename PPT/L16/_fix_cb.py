# -*- coding: utf-8 -*-
import json, os
SEED = r"D:\doc\ZGCLab\triza\project\design\akbank\seed\parameters.seed.json"
d = json.load(open(SEED, encoding="utf-8"))
cb = d["claim_boundary"]
if "SEC-P25" not in cb["first_of_kind"]:
    cb["first_of_kind"].append("SEC-P25")
    json.dump(d, open(SEED, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("claim_boundary.first_of_kind += SEC-P25")

d2 = json.load(open(SEED, encoding="utf-8"))
cb2 = d2["claim_boundary"]
ids = [x["id"] for x in d2["items"]]
cov = set(cb2["first_of_kind"]) | set(cb2["quantified_refinement"])
print("count:", d2["count"], "| items:", len(ids))
print("first_of_kind:", len(cb2["first_of_kind"]), "| quantified:", len(cb2["quantified_refinement"]))
print("未归类的参数:", [i for i in ids if i not in cov] or "无")
print("边界里不存在的参数:", [i for i in cov if i not in ids] or "无")
print("无经典映射但不在 first_of_kind:",
      [x["id"] for x in d2["items"] if not x["classic_mapping"].get("params") and x["id"] not in cb2["first_of_kind"]] or "无")