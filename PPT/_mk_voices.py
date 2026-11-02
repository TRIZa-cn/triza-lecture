# -*- coding: utf-8 -*-
import asyncio, json, wave
from pathlib import Path
import edge_tts

BASE = Path(r"D:\doc\ZGCA\Lectures\TRIZ")
OUT = BASE / "_voices"
OUT.mkdir(exist_ok=True)

d = json.loads((BASE / "_notes_tts.json").read_text(encoding="utf-8"))
t = d["notes"][0]
# first two sentences
cut = t.find(". ", 60)
text = t[: cut + 1] if cut > 0 else t[:200]

VOICES = {
    "andrew": "en-US-AndrewNeural",
    "guy": "en-US-GuyNeural",
    "brian": "en-US-BrianNeural",
    "ryan_gb": "en-GB-RyanNeural",
}

async def main():
    for name, v in VOICES.items():
        mp3 = OUT / (name + ".mp3")
        await edge_tts.Communicate(text, v).save(str(mp3))
        print(name, v, mp3.stat().st_size, "bytes")

asyncio.run(main())

import miniaudio
for name in VOICES:
    mp3 = OUT / (name + ".mp3")
    dec = miniaudio.decode_file(str(mp3), output_format=miniaudio.SampleFormat.SIGNED16)
    with wave.open(str(OUT / (name + ".wav")), "wb") as w:
        w.setnchannels(dec.nchannels)
        w.setsampwidth(2)
        w.setframerate(dec.sample_rate)
        w.writeframes(dec.samples.tobytes())
    print("  %-8s sr=%d ch=%d dur=%.2fs" % (name, dec.sample_rate, dec.nchannels, dec.num_frames / dec.sample_rate))

print("TEXT:", text)
