#!/usr/bin/env bash
# 在已配置好的 GPT-SoVITS 工作目录执行：
#   bash /media/wislab/Datasets/hyh_water/audio/tts.sh --input-dir <本目录> --output <输出目录>
# 生成 16 段，文件名与 txt 一致。
set -e
WORK=/media/wislab/Datasets/hyh_water/audio
bash "$WORK/tts.sh" --input-dir "<改成配音稿_W1的路径>" --output "$WORK/output/W1_矛盾15m" --speed 1.0
