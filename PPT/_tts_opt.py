# -*- coding: utf-8 -*-
import asyncio, json, re
from pathlib import Path
import edge_tts, miniaudio
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
DECK = D / "TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
OUT = D / "narration_opt"
VOICE = "zh-CN-YunyangNeural"
RATE = "-13%"


def clean(t):
    t = t.replace("**", "").replace("\u3000", " ")
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def dur(path):
    snd = miniaudio.decode_file(str(path))
    return len(snd.samples) / (snd.nchannels * snd.sample_rate)


async def main():
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation(str(DECK))
    rows, total_c, total_s = [], 0, 0.0
    for i, s in enumerate(prs.slides, 1):
        txt = clean(s.notes_slide.notes_text_frame.text)
        f = OUT / ("A%02d.mp3" % i)
        await edge_tts.Communicate(txt, VOICE, rate=RATE).save(str(f))
        d = dur(f)
        rows.append({"slide": i, "chars": len(txt), "dur": round(d, 2), "file": f.name})
        total_c += len(txt)
        total_s += d
    json.dump({"deck": DECK.name, "voice": VOICE, "rate": RATE, "rows": rows,
               "total_chars": total_c, "total_sec": round(total_s, 2),
               "chars_per_min": round(total_c / total_s * 60, 1)},
              open(D / "_narration_opt.json", "w", encoding="utf-8"), ensure_ascii=False)
    print(json.dumps({"total_chars": total_c, "total_sec": round(total_s, 2),
                      "min": round(total_s / 60, 2),
                      "chars_per_min": round(total_c / total_s * 60, 1),
                      "per": [(r["slide"], r["chars"], r["dur"]) for r in rows]},
                     ensure_ascii=False))

asyncio.run(main())