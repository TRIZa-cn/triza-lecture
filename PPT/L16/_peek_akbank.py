# -*- coding: utf-8 -*-
import json, os
for p in [r"D:\doc\ZGCLab\triza\project\design\akbank\seed\parameters.seed.json",
          r"D:\doc\ZGCLab\triza\project\design\akbank\schema\parameter.schema.json"]:
    print("=" * 20, os.path.basename(p), os.path.getsize(p) if os.path.isfile(p) else "MISSING")
    if not os.path.isfile(p): continue
    d = json.load(open(p, encoding="utf-8"))
    if "schema" in os.path.basename(p):
        print(json.dumps(d, ensure_ascii=False, indent=1)[:1500])
    else:
        print("type:", type(d).__name__, "| keys:", list(d)[:8] if isinstance(d, dict) else len(d))
        items = d if isinstance(d, list) else d.get("parameters") or d.get("items") or []
        print("count:", len(items))
        print("first:", json.dumps(items[0], ensure_ascii=False, indent=1)[:900])
        print("last id:", items[-1].get("id"))