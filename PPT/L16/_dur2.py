import json, subprocess, glob, os, re
import imageio_ffmpeg
FF = imageio_ffmpeg.get_ffmpeg_exe()
MP3 = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音20261030\_嵌入MP3"
out = {}
for p in sorted(glob.glob(MP3 + r"\p*.mp3")):
    r = subprocess.run([FF, "-hide_banner", "-i", p], capture_output=True, text=True, errors="ignore")
    m = re.search(r"Duration:\s*(\d+):(\d+):([\d.]+)", r.stderr)
    d = int(m.group(1))*3600 + int(m.group(2))*60 + float(m.group(3))
    out[os.path.basename(p)] = round(d, 2)
json.dump(out, open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16\_durations.json", "w"), ensure_ascii=False)
print("时长条数:", len(out), "| 合计 %.1f 秒" % sum(out.values()))
print(list(out.items())[:3])