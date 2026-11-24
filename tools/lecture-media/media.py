"""Install Windows media dependencies or run the lecture PPTX/video tools with them."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

HERE = Path(__file__).resolve().parent
FILES = ["package.json", "package-lock.json", "requirements.txt"]


def main():
    if os.name != "nt" or len(sys.argv) < 2 or sys.argv[1] not in {"install", "pptx", "render"}:
        raise SystemExit("Usage on Windows: python media.py install | pptx PROJECT [OPTIONS] | render PROJECT [OPTIONS]")
    local = Path(os.environ["LOCALAPPDATA"])
    digest = hashlib.sha256(b"".join((HERE/name).read_bytes() for name in FILES)).hexdigest()[:16]
    parent = local/"TRIZa"/"LectureMedia"
    runtime = parent/(sys.implementation.cache_tag+"-"+digest)
    if not runtime.is_dir():
        parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=parent) as tmp:
            stage = Path(tmp)/"runtime"
            stage.mkdir()
            for name in FILES:
                shutil.copy2(HERE/name, stage)
            npm = shutil.which("npm.cmd")
            if not npm:
                raise SystemExit("Node.js with npm is required for lecture media setup.")
            subprocess.run([npm,"ci","--ignore-scripts","--no-audit","--no-fund"],cwd=stage,check=True)
            subprocess.run([sys.executable,"-m","pip","install","--disable-pip-version-check","--only-binary=:all:","--target",str(stage/"pydeps"),"-r",str(stage/"requirements.txt")],check=True)
            try:
                stage.rename(runtime)
            except FileExistsError:
                if not runtime.is_dir():
                    raise
    env = os.environ.copy()
    env["PYTHONPATH"] = str(runtime/"pydeps")+os.pathsep+env.get("PYTHONPATH","")
    if sys.argv[1] == "install":
        subprocess.run([sys.executable,"-c","import pptx, yaml, PIL, imageio_ffmpeg; print(imageio_ffmpeg.get_ffmpeg_version())"],env=env,check=True)
        print(runtime)
        return
    if sys.argv[1] == "pptx":
        exporter = HERE.parents[1]/"skills/ppt-master/scripts/svg_to_pptx.py"
        subprocess.run([sys.executable,str(exporter),*sys.argv[2:]],env=env,check=True)
        return
    node = local/"TRIZa Desktop/resources/node/node.exe"
    cli = runtime/"node_modules/hyperframes/bin/hyperframes.mjs"
    ffmpeg = subprocess.check_output([sys.executable,"-c","import imageio_ffmpeg;print(imageio_ffmpeg.get_ffmpeg_exe())"],env=env,text=True).strip()
    env.setdefault("HYPERFRAMES_FFMPEG_PATH",ffmpeg)
    env.setdefault("HYPERFRAMES_FFPROBE_PATH",str(runtime/"node_modules/@ffprobe-installer/win32-x64/ffprobe.exe"))
    env.setdefault("HYPERFRAMES_NO_TELEMETRY","1")
    for base, relative in ((os.environ.get("ProgramFiles(x86)"),"Microsoft/Edge/Application/msedge.exe"),(os.environ.get("ProgramFiles"),"Google/Chrome/Application/chrome.exe")):
        if base and (Path(base)/relative).is_file():
            env.setdefault("HYPERFRAMES_BROWSER_PATH",str(Path(base)/relative))
            break
    subprocess.run([str(node) if node.is_file() else "node",str(cli),"render",*sys.argv[2:]],env=env,check=True)


if __name__ == "__main__":
    main()
