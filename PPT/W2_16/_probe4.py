import io, re, sys
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\w_s0204.py"
t = io.open(p, encoding="utf-8").read()
i = t.find('"n": 4,')
seg = t[i:i+9000]
for m in re.finditer(r'[^"\n]{0,40}(?:五条|六条|40\s*条|四十条|条原理|个原理|原理)[^"\n]{0,40}', seg):
    print(repr(m.group(0)))
