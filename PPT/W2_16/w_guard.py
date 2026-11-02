# -*- coding: utf-8 -*-
import json, hashlib, pathlib, datetime
base = pathlib.Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
out = base / "W2_16" / "out"
a = base / "TRIZ_W1_\u77db\u76fe15m.pptx"
b = base / "TRIZ_W1_\u77db\u76fe15m_\u6f14\u64ad\u7248.pptx"
def info(p):
    st = p.stat()
    return {"size": st.st_size,
            "mtime": datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"),
            "sha256": hashlib.sha256(p.read_bytes()).hexdigest()[:16]}
ia, ib = info(a), info(b)
cutoff = datetime.datetime(2026, 10, 31)
unchanged = (ia["size"] == 704606 and ib["size"] == 16890575
             and datetime.datetime.fromtimestamp(a.stat().st_mtime) < cutoff
             and datetime.datetime.fromtimestamp(b.stat().st_mtime) < cutoff)
print(json.dumps({"existing_unchanged": unchanged, "w1": ia, "play": ib,
                  "decks": len(list(out.glob("*.pptx")))}, ensure_ascii=True))