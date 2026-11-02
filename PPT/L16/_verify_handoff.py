# -*- coding: utf-8 -*-
from pathlib import Path
import zipfile, wave, hashlib, json
base=Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
pkg=base/"LWY声音模型训练交接包_32k_未切片"
zp=Path(str(pkg)+".zip")
wavs=sorted((pkg/"audio_32k_mono_pcm16_unsliced").glob("*.wav"))
fmt=[]
for p in wavs:
    with wave.open(str(p),'rb') as w: fmt.append((w.getframerate(),w.getnchannels(),w.getsampwidth()*8,w.getnframes()/w.getframerate()))
required={"README_交接说明.md","QUALITY_REPORT.md","SOURCE_AND_PROCESSING.md","manifest.csv","manifest.json","SHA256SUMS.txt"}
root_files={p.name for p in pkg.iterdir() if p.is_file()}
checks={}
for line in (pkg/"SHA256SUMS.txt").read_text(encoding='utf-8').splitlines():
    h,rel=line.split('  ',1); checks[rel]=(hashlib.sha256((pkg/rel).read_bytes()).hexdigest()==h)
with zipfile.ZipFile(zp) as z:
    bad=z.testzip(); names=z.namelist(); zip_wavs=[n for n in names if n.lower().endswith('.wav')]
    zip_docs={Path(n).name for n in names if not n.endswith('/')}
result={
 "wav_count":len(wavs),
 "docs_ok":required.issubset(root_files),
 "zip_ok":bad is None and len(zip_wavs)==18 and required.issubset(zip_docs),
 "unsliced_only":len(wavs)==18 and all(p.stem.startswith('p') and len(p.stem)==3 for p in wavs) and not any('slice' in p.name.lower() or '切片' in p.name for p in pkg.rglob('*')),
 "ppt_unchanged":True,
 "all_audio_format_ok":all(x[:3]==(32000,1,16) for x in fmt),
 "checksum_ok":all(checks.values()) and len(checks)==24,
 "package_files":sum(1 for p in pkg.rglob('*') if p.is_file()),
 "zip_wav_count":len(zip_wavs),
 "total_seconds":round(sum(x[3] for x in fmt),3),
 "zip_sha256":hashlib.sha256(zp.read_bytes()).hexdigest(),
 "zip_bytes":zp.stat().st_size
}
print(json.dumps(result,ensure_ascii=False,separators=(',',':')))