# -*- coding: utf-8 -*-
import zipfile, os
Z = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZa_声音模型与试讲_20261029.zip"
print("zip size:", round(os.path.getsize(Z)/1048576, 1), "MB")
zf = zipfile.ZipFile(Z)
infos = zf.infolist()
print("entries:", len(infos))
total = sum(i.file_size for i in infos)
print("uncompressed:", round(total/1048576, 1), "MB")
print("\n---- 按目录汇总 ----")
from collections import defaultdict
agg = defaultdict(lambda: [0, 0])
for i in infos:
    top = i.filename.split("/")[0] if "/" in i.filename else "(root)"
    agg[top][0] += 1; agg[top][1] += i.file_size
for k, (n, b) in sorted(agg.items(), key=lambda x: -x[1][1])[:25]:
    print(f"  {k[:60]:<62} {n:>6} files  {b/1048576:>9.1f} MB")
print("\n---- 非模型文件（可能含说明/环境）----")
for i in infos:
    n = i.filename
    if any(n.lower().endswith(x) for x in (".md", ".txt", ".bat", ".sh", ".yaml", ".yml", ".json", ".py", ".exe")):
        print(f"  {n}  ({i.file_size/1024:.1f} KB)")