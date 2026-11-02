# -*- coding: utf-8 -*-
from pathlib import Path
import zipfile,wave,hashlib,json,re
base=Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT"); pkg=base/"LWY声音模型训练交接包_32k_未切片"; zp=Path(str(pkg)+".zip")
aud=pkg/"audio_32k_mono_pcm16_unsliced"; wavs=sorted(aud.glob('*.wav'))
formats=[]
for p in wavs:
 with wave.open(str(p),'rb') as w: formats.append((w.getframerate(),w.getnchannels(),w.getsampwidth()*8,w.getnframes()/w.getframerate()))
required={"README_交接说明.md","QUALITY_REPORT.md","SOURCE_AND_PROCESSING.md","manifest.csv","manifest.json","SHA256SUMS.txt"}; roots={p.name for p in pkg.iterdir() if p.is_file()}
lines=(pkg/'SHA256SUMS.txt').read_text(encoding='utf-8').splitlines(); checksum_ok=True
for line in lines:
 h,rel=line.split('  ',1); checksum_ok &= hashlib.sha256((pkg/rel).read_bytes()).hexdigest()==h
non_sum=[p for p in pkg.rglob('*') if p.is_file() and p.name!='SHA256SUMS.txt']; checksum_ok &= len(lines)==len(non_sum)
with zipfile.ZipFile(zp) as z:
 names=z.namelist(); zipwav=[n for n in names if n.lower().endswith('.wav')]; zip_ok=z.testzip() is None and len(zipwav)==18 and required.issubset({Path(n).name for n in names})
expected={f'p{i:02}.wav' for i in range(1,19)}; audio_names={p.name for p in wavs}; unsliced=audio_names==expected and len(audio_names)==18
ppt_unchanged=(base/'TRIZ_W1_矛盾15m.pptx').stat().st_size==704606 and (base/'TRIZ_W1_矛盾15m_演播版.pptx').stat().st_size>16000000
result={"wav_count":len(wavs),"docs_ok":required.issubset(roots) and checksum_ok,"zip_ok":zip_ok,"unsliced_only":unsliced,"ppt_unchanged":ppt_unchanged,"all_audio_format_ok":all(x[:3]==(32000,1,16) for x in formats),"checksum_entries":len(lines),"checksum_ok":checksum_ok,"total_seconds":round(sum(x[3] for x in formats),3),"zip_sha256":hashlib.sha256(zp.read_bytes()).hexdigest(),"zip_bytes":zp.stat().st_size}
print(json.dumps(result,ensure_ascii=False,separators=(',',':')))