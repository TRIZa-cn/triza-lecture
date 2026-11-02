# -*- coding: utf-8 -*-
from pathlib import Path
p = Path(r"D:\doc\ZGCA\Lectures\TRIZ\_shift.py")
s = p.read_text(encoding="utf-8")
old = """    while pos_a + N <= len(x):
        out[pos_s:pos_s+N] += x[pos_a:pos_a+N] * win
        norm[pos_s:pos_s+N] += win
        pos_a += HA; pos_s += HS"""
new = """    while pos_a + N <= len(x):
        m = min(N, out_len - pos_s)
        if m <= 0:
            break
        out[pos_s:pos_s+m] += x[pos_a:pos_a+m] * win[:m]
        norm[pos_s:pos_s+m] += win[:m]
        pos_a += HA; pos_s += HS"""
assert old in s, "anchor not found"
p.write_text(s.replace(old, new), encoding="utf-8")
print("patched")
