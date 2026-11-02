import io
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\w_s0204.py"
t = io.open(p, encoding="utf-8").read()
i = t.find('"n": 4,')
seg = t[i:i+12000]
a = seg.find('"table"')
b = seg.find('"ex1"', a)
io.open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\_s4_table.txt","w",encoding="utf-8").write(seg[a:b])
print(seg[a:b])