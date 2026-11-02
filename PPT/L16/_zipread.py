# -*- coding: utf-8 -*-
import zipfile, json
Z = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZa_声音模型与试讲_20261029.zip"
zf = zipfile.ZipFile(Z)
for name in ["README.md", "delivery/使用说明.txt", "delivery/model_manifest.json",
             "delivery/tts_infer.yaml", "delivery/reference.txt", "requirements-lock.txt"]:
    print("=" * 70)
    print("##", name)
    try:
        print(zf.read(name).decode("utf-8")[:2200])
    except Exception as e:
        print("读取失败:", e)
print("=" * 70)
print("## 大文件（模型权重/音频）")
for i in zf.infolist():
    if i.file_size > 1_000_000:
        print(f"  {i.filename}  {i.file_size/1048576:.1f} MB")