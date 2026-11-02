# -*- coding: utf-8 -*-
import io, json, os, re, subprocess

R = r"D:\doc\ZGCA\Lectures\TRIZ"

def git(*a):
    p = subprocess.run(["git", "-C", R] + list(a), capture_output=True, text=True, encoding="utf-8", errors="replace")
    return p.returncode, (p.stdout or "").strip(), (p.stderr or "").strip()

rc_log, log, _ = git("log", "--oneline")
commits = len([x for x in log.splitlines() if x.strip()])
rc_head, head, _ = git("rev-parse", "HEAD")
rc_rem, rem, _ = git("rev-parse", "origin/main")
rc_tag, tag, _ = git("tag", "--list", "v0.1")
rc_files, files, _ = git("ls-files")
fl = [x for x in files.splitlines() if x.strip()]
media = [x for x in fl if re.search(r"\.(wav|mp3|m4a|zip|7z|rar|pdf|epub|png|jpg|html|css)$", x, re.I)]
lock = [x for x in fl if os.path.basename(x).startswith("~$")]
sizes = []
for x in fl:
    p = os.path.join(R, x.replace("/", os.sep))
    if os.path.isfile(p):
        sizes.append((os.path.getsize(p), x))
maxb, maxf = (max(sizes) if sizes else (0, ""))
decks = [x for x in fl if x.startswith("PPT/W2_16/out/") and x.endswith(".pptx")]

repo_ok = (rc_log == 0 and commits >= 3 and bool(rem) and head == rem and bool(tag))
ignore_ok = (len(media) == 0 and len(lock) == 0 and maxb < 2 * 1024 * 1024)

md = os.path.join(R, "docs", "\u8bb2\u8bfe\u98ce\u683c\u6863\u6848.md")
js = os.path.join(R, "style", "style_profile.json")
style_md = io.open(md, encoding="utf-8").read() if os.path.exists(md) else ""
prof = json.loads(io.open(js, encoding="utf-8").read()) if os.path.exists(js) else {}
nrules = len(prof.get("rules", []))
style_ok = (len(style_md) >= 800) and (nrules >= 8)

sk = io.open(os.path.join(R, "PPT", "TRIZa-Lecture-SKILL.md"), encoding="utf-8").read()
skill_ok = ('version: "1.2.0"' in sk) and ("## \u4f8b\u56fe\u89c4\u8303" in sk) and ("## \u6761\u6570\u4e00\u81f4\u6027" in sk)

vp = os.path.join(R, "docs", "\u58f0\u97f3\u6a21\u578b\u6f14\u64ad\u7248\u65b9\u6848.md")
vd = io.open(vp, encoding="utf-8").read() if os.path.exists(vp) else ""
kw_plan = {"48000": "48000" in vd, "80kbps": ("80 kbps" in vd or "80kbps" in vd),
           "sync_log": "sync_log" in vd, "launcher": "\u542f\u52a8\u5668" in vd,
           "mp3": "MP3" in vd or "mp3" in vd}
plan_ok = sum(kw_plan.values()) >= 4
kw_gap = ["\u6743\u91cd", "venv", "torch", "ffmpeg", "\u53c2\u8003", "\u5207\u7247", "\u6388\u6743"]
gap_hits = sum(1 for k in kw_gap if k in vd)
honest_ok = (gap_hits >= 5) and ("\u5df2\u5b8c\u6210\u8bad\u7ec3" not in vd)

print(json.dumps({
    "repo_ok": repo_ok, "ignore_ok": ignore_ok,
    "style_ok": style_ok, "skill_ok": skill_ok,
    "plan_ok": plan_ok, "honest_ok": honest_ok,
    "detail": {"commits": commits, "head_eq_remote": head == rem, "tag": tag,
               "tracked": len(fl), "media": media[:5], "locks": lock[:5],
               "max_bytes": maxb, "max_file": maxf, "decks": len(decks),
               "style_rules": nrules, "style_md_bytes": len(style_md),
               "skill_bytes": len(sk), "plan_kw": kw_plan, "gap_hits": gap_hits,
               "doc_bytes": len(vd)},
}, ensure_ascii=False))