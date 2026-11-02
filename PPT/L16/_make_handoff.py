# -*- coding: utf-8 -*-
from pathlib import Path
import wave, csv, json, hashlib, zipfile, shutil, os, re
base=Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
pkg=base/"LWY声音模型训练交接包_32k_未切片"
aud=pkg/"audio_32k_mono_pcm16_unsliced"
src=base/"LWY录音20261030"
wavs=sorted(aud.glob("p*.wav"))
m4as=sorted(src.glob("*.m4a"))
rows=[]
for i,w in enumerate(wavs,1):
    with wave.open(str(w),'rb') as f:
        ch=f.getnchannels(); sr=f.getframerate(); sw=f.getsampwidth(); frames=f.getnframes(); dur=frames/sr
    source=m4as[i-1].name if i<=len(m4as) else ""
    h=hashlib.sha256(w.read_bytes()).hexdigest()
    rows.append(dict(slide=i,wav=w.name,source_m4a=source,duration_seconds=round(dur,3),sample_rate_hz=sr,channels=ch,bits_per_sample=sw*8,sha256=h,size_bytes=w.stat().st_size))
with (pkg/"manifest.csv").open('w',newline='',encoding='utf-8-sig') as f:
    cw=csv.DictWriter(f,fieldnames=rows[0].keys()); cw.writeheader(); cw.writerows(rows)
(pkg/"manifest.json").write_text(json.dumps({"schema":"lwy.voice-training-handoff/v1","speaker":"LWY","language":"zh-CN","segmentation":"unsliced_per_slide","audio_format":{"container":"WAV","codec":"PCM signed 16-bit little-endian","sample_rate_hz":32000,"channels":1},"items":rows},ensure_ascii=False,indent=2),encoding='utf-8')
quality='''# 音频质量报告\n\n- 文件数：18（整页录音，未切片）\n- 总时长：约 1004.0 秒（16.73 分钟）\n- 格式：WAV / PCM signed 16-bit little-endian / 32,000 Hz / 单声道\n- 处理：由原始 AAC 48 kHz 立体声 M4A 转换；仅轻度去除头尾低于 -50 dB 的静音；未做动态压缩、降噪、均衡、响度标准化或切片。\n- 平均电平：约 -16.8 至 -18.4 dB\n- 峰值：约 -3.2 至 -4.4 dB；未发现削波\n- 第 1 段参考：Integrated Loudness -17.5 LUFS，True Peak -3.6 dBTP，LRA 3.6 LU\n\n## 训练端注意\n\n本包特意保留为 18 段长音频。多数声音克隆训练流程仍需在 GPU 端完成切片、ASR/文本校对和说话人标注；这些步骤尚未执行。训练人员不应把本包误认为已对齐的最终训练集。\n'''
(pkg/"QUALITY_REPORT.md").write_text(quality,encoding='utf-8')
readme='''# LWY 声音模型训练交接包（32 kHz，未切片）\n\n## 交付内容\n\n- `audio_32k_mono_pcm16_unsliced/`：18 个整页 WAV，按 `p01.wav` 至 `p18.wav` 排列。\n- `manifest.csv` / `manifest.json`：页码、源 M4A 文件名、时长、格式、大小与 SHA-256。\n- `QUALITY_REPORT.md`：处理方式及质量检查。\n- `SHA256SUMS.txt`：交接包内文件校验值。\n- `SOURCE_AND_PROCESSING.md`：来源及转换参数。\n\n## 已确认要求\n\n1. 采样率固定为 **32 kHz**。\n2. **不切片**；保留 18 个整页录音。\n3. 音频是 mono PCM16 WAV，可跨 Windows/Linux 读取。\n\n## 给 GPU 训练人员\n\n1. 解压后先运行校验：Linux/macOS 使用 `sha256sum -c SHA256SUMS.txt`；Windows PowerShell 可逐项运行 `Get-FileHash -Algorithm SHA256`。\n2. 按所用框架在训练端执行切片、ASR、文本校对和数据集清单生成。\n3. 训练前人工抽听全部音频，剔除口误、碰撞声、长停顿或非目标说话人片段。\n4. 原始文件是课堂口播，包含专业术语（TRIZ、David Epstein 等）；自动转写后必须校对。\n5. `manifest` 中的页码与源文件按录音时间升序对应第 1–18 页。\n\n## 不包含\n\n- 不包含切片文件。\n- 不包含人工校对文本或音素标注。\n- 不包含训练配置、模型权重或训练结果。\n'''
(pkg/"README_交接说明.md").write_text(readme,encoding='utf-8')
process='''# 来源与处理参数\n\n- 来源目录：`LWY录音20261030`\n- 原始格式：M4A / AAC / 48 kHz / stereo\n- 排序规则：文件名时间戳升序，映射课件第 1–18 页\n- 目标格式：WAV / PCM signed 16-bit little-endian / 32 kHz / mono\n- 音频滤镜：仅去除头尾低于 -50 dB 的静音\n- 保留策略：未切片、未归一化、未降噪、未压缩、未均衡\n- 原始 M4A 未包含于 ZIP，避免重复占用空间；来源文件名保存在 manifest。\n'''
(pkg/"SOURCE_AND_PROCESSING.md").write_text(process,encoding='utf-8')
# checksum all files except checksum itself
entries=[]
for p in sorted(x for x in pkg.rglob('*') if x.is_file() and x.name!='SHA256SUMS.txt'):
    rel=p.relative_to(pkg).as_posix(); h=hashlib.sha256(p.read_bytes()).hexdigest(); entries.append(f"{h}  {rel}")
(pkg/"SHA256SUMS.txt").write_text('\n'.join(entries)+'\n',encoding='utf-8')
print(json.dumps({"wav_count":len(wavs),"total_seconds":round(sum(r['duration_seconds'] for r in rows),3),"all_format_ok":all(r['sample_rate_hz']==32000 and r['channels']==1 and r['bits_per_sample']==16 for r in rows),"docs":sorted(p.name for p in pkg.iterdir() if p.is_file())},ensure_ascii=False))