# -*- coding: utf-8 -*-
import zipfile, os
z = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音训练交接包-20261029.zip"
zf = zipfile.ZipFile(z)
bad = zf.testzip()
print("zip:", round(os.path.getsize(z)/1048576, 1), "MB | 完整性:", "OK" if bad is None else bad)
print("条目:", len(zf.namelist()))
print("\n---- README 前 30 行 ----")
t = zf.read("TRIZa-声音训练交接包/README.md").decode("utf-8")
print("\n".join(t.splitlines()[:30]))
print("\n---- 训练步骤 目录结构 ----")
s = zf.read("TRIZa-声音训练交接包/03_训练步骤/训练步骤-6GB显卡.md").decode("utf-8")
for line in s.splitlines():
    if line.startswith("## "): print("  ", line[3:])
print("\n---- list.txt 首行 ----")
print(zf.read("TRIZa-声音训练交接包/02_对照文字/list.txt").decode("utf-8").splitlines()[0][:120])
zf.close()