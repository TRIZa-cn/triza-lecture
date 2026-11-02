import io, json, re
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\w_s0204.py"
t = io.open(p, encoding="utf-8").read()
i = t.find('"n": 4,')
seg = t[i:i+9000]
m = re.search(r'"rows":\s*(\[\[.*?\]\])\s*,\s*"col_w"', seg, re.S)
rows = json.loads(m.group(1)) if m else []
io.open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\_s4_table.txt","w",encoding="utf-8").write(
    "rows=" + str(len(rows)) + "\n" + "\n".join(" | ".join(r) for r in rows))
print("rows", len(rows))