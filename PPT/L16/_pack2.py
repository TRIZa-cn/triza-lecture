# -*- coding: utf-8 -*-
import os, json, zipfile
BASE = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音训练交接包"
st = json.load(open(os.path.join(BASE, "04_校验", "_stats.json"), encoding="utf-8"))
tot_s = sum(r["sec"] for r in st); tot_c = sum(r["chars"] for r in st)
cpm = tot_c / tot_s * 60

def W(sub, name, text):
    p = os.path.join(BASE, sub, name)
    open(p, "w", encoding="utf-8").write(text)
    return p

# ---------------- README ----------------
readme = f"""# TRIZa 声音模型训练交接包

给「用刘文印本人声音训练 TTS 模型」这件事准备的完整素材与步骤。
打包时间：2026-10-29。素材来源：16 段手机录音（m4a/AAC），已转 44.1 kHz 单声道 WAV 并去掉首尾静音。

## 先看结论：6 GB 显存够不够？

**够，但只能走 V2 / V2Pro / V2ProPlus 这条线，不能碰 V3 / V4。**

| 版本线 | 微调显存 | 6 GB 能否 |
|---|---|---|
| V2 / V2Pro / V2ProPlus | 推理 6 GB 起步；WebUI 默认 batch = 显存GB÷2 → 6 GB 给 **2** | **能**，必要时把 batch 降到 1 |
| V3 / V4 | 全量微调 14 GB；开梯度检查点 12 GB；LoRA **8 GB** | **不能**，LoRA 都超 |

出处：NexGPU 的 GPT-SoVITS 部署页（第三方整理，非官方文档，但与社区经验一致）：
<https://nexgpu.net/models/gpt-sovits/> 。官方权重在 HuggingFace `lj1995/GPT-SoVITS`，整仓约 5.3 GB。
**上机前请用官方 README 复核一遍显存口径**，第三方页面的数字可能随版本变化。

另外两条硬规则（官方代码里的校验）：
- 推理/微调用的**参考音频必须在 3～10 秒**，超了直接报「参考音频在3~10秒范围外，请更换！」。
- `config.py` 里显存低于 4 GB 会直接回落 CPU；Pascal（P40/1080）和 GTX 16 系会被强制 float32，速度减半。

## 包内清单

```
01_音频/           16 个 WAV（44.1 kHz 单声道，已裁首尾静音）  约 {sum(r['bytes'] for r in st)/1048576:.0f} MB
02_对照文字/
   list.txt        GPT-SoVITS 标注格式：路径|说话人|语言|文本
   labels.csv      Excel 可读的对照表（文件名/秒/字数/语速/文本）
   逐段文字.md      人读版，校对 ASR 用
03_训练步骤/
   训练步骤-6GB显卡.md   从零到出模型的完整流程（主文件）
   参数速查.md          6/8/12/24 GB 的 batch 与 epoch 建议
04_校验/
   素材体检.md          16 段的时长、字数、语速、可用性判断
```

## 最短路径

1. 读 `03_训练步骤/训练步骤-6GB显卡.md`（15 分钟）。
2. 按里面第 0–3 步装环境、下权重（约 5.3 GB，国内建议加 `--source HF-Mirror`）。
3. 第 4–8 步处理素材：切分 → FunASR 打标 → **用 `02_对照文字/逐段文字.md` 校对** → 生成 list。
4. 第 9–10 步训练：SoVITS 先训，GPT 后训，batch 都设 2。
5. 第 11 步拿 `S04` 或 `S10` 里一句完整的话做参考音频，试听。

## 这份素材的真实水平（别高估）

- 手机录制，不是录音棚；房间混响和底噪都有一点，达不到 V4 那条线「48 kHz 母带」的要求，所以**只该走 V2Pro**。
- 总时长 **{tot_s/60:.1f} 分钟**、**{tot_c} 字**、**{cpm:.1f} 字/分钟**。官方说 1 分钟能训，那是最低线；17 分钟属于「够用、不出错」的量，不是「越多越好」。
- 内容是一门 TRIZ 课的讲稿：术语密度高（TRIZ、ARIZ、MPV、矩阵、裁剪、物场…），日常口语少。**训出来的模型会擅长念讲义，不擅长闲聊**——正好符合用途。

## 训完之后

产出的两个权重（`SoVITS_weights_v2Pro/*.pth`、`GPT_weights_v2Pro/*.ckpt`）拷回本机，告诉我路径，我把剩下 15 讲的解说批量生成。第 7 页 `散热片` 那处旧词也能一并修掉。
"""
W("", "README.md", readme)

# ---------------- 训练步骤 ----------------
step = f"""# 训练步骤（6 GB 显存档）

对象：GPT-SoVITS V2Pro / V2ProPlus，素材为本包 `01_音频/`。
以下步骤在 6 GB 显存上验证过路线的可行性依据是社区口径，**本机未实测**；时间均为估算，不是实测值。

---

## 0. 环境

- Windows 10/11 + NVIDIA 驱动（CUDA 12.x 运行时随 PyTorch 装，不用单独装 CUDA Toolkit）
- **Python 3.10 或 3.11**（3.12 一般也行；**3.13/3.14 不要用**，torch 轮子不全）
- ffmpeg（WebUI 的 UVR5 与切片都依赖它）
- 磁盘：权重 5.3 GB + 素材 1 GB + 训练产物 2 GB，留 **15 GB** 余量
- 内存 16 GB 起；Windows 虚拟内存设到 20 GB 以上（训练会大量用页面文件）

装环境（国内加镜像源）：

```bash
git clone https://github.com/RVC-Boss/GPT-SoVITS.git
cd GPT-SoVITS
bash install.sh --device CU128 --source HF-Mirror
```

Windows 下如果 `install.sh` 跑不了，用 `conda` 建 3.10 环境后手动 `pip install -r requirements.txt`，再补 `pip install torch --index-url https://download.pytorch.org/whl/cu121`。

## 1. 权重

```bash
huggingface-cli download lj1995/GPT-SoVITS --local-dir GPT_SoVITS/pretrained_models
```

跑 **V2Pro** 必须额外确认这个文件在位（容易漏，漏了要到推理才报错）：

```
GPT_SoVITS/pretrained_models/sv/pretrained_eres2netv2w24s4ep4.ckpt   # 108 MB
```

## 2. 启动 WebUI

```bash
python webui.py zh_CN
```

- 主界面：<http://127.0.0.1:9874>
- UVR5 人声分离：9873（本包素材用不上，见第 3 步）

## 3. UVR5 去伴奏：跳过

手机录的是干声，没有人声伴奏。**不要习惯性地跑 UVR5**——分离模型会削掉一部分高频，手机的齿音本来就不多，削完更闷。
只有一种情况要处理：如果你听下来底噪明显（嘶嘶声），可以用 WebUI 里的轻度降噪，但**做完必须 A/B 对比原声**，降噪过度会让模型学到"发闷"的音色。

## 4. 切分（0-切分音频）

- 目标片长：**3～10 秒**（这是硬约束，不是建议）
- `01_音频/` 共 {tot_s:.0f} 秒 → 预计切出 **180～220 片**
- 参数：`min_length=3000`、`max_length=10000`、`min_interval=300`、`max_sil_kept=500`
- 切完听 3～5 片：有没有从句子中间切开、有没有把两个字切成一片。切坏的片直接删，不要留。

## 5. ASR 打标（1-ASR）

- 引擎选 **FunASR**（中文最准），有字幕的可用 `faster-whisper`
- 语言代码只认 `zh` / `ja` / `en` / `ko` / `yue`，中文写 `zh`
- 输出是 `xxx.list`，每行格式：`音频路径|说话人名|语言|文本`

## 6. 校对（最关键的一步）

FunASR 会把专名念错，而**错字会直接污染模型**。用 `02_对照文字/逐段文字.md` 逐条比对，重点抓这些词：

| 容易被打成 | 正确 |
|---|---|
| 吹子 / 脆子 / T R I Z | TRIZ |
| 阿奇舒勒 / 阿奇苏勒 | 阿奇舒勒 |
| 萃智 / 脆智 | 萃智 |
| 矛盾举证 / 矛盾巨阵 | 矛盾矩阵 |
| 物场模型 | 物场模型 |
| 功能导向搜索 | 功能导向搜索 |
| 置信度 / 致信度 | 置信度 |

其余规则：
- **删掉所有标点以外的符号**（括号、破折号、书名号、引号），否则会被念出来
- 数字统一成阿拉伯数字或统一成汉字，**不要混**
- 每条文本长度控制在 30 字以内（太长说明切分不对，回去切）
- 空文本、纯标点的行直接删

## 7. 格式化（一键三连里的「格式化」）

生成 `xxx.list` 与训练用的 `2-name2text.txt`。检查：条数 = 切片数，没有空行。

## 8. 训练前最后一次检查

- [ ] 切片全部 3～10 秒（超出会被丢）
- [ ] 每条都有文本，且没有乱码
- [ ] 说话人列**全部相同**（本包用 `LWY`）
- [ ] 采样率统一（WebUI 会重采样到 32 kHz，不用你手动转）

---

## 9. 训练 SoVITS（先训这个）

| 参数 | 6 GB 建议 | 说明 |
|---|---|---|
| batch_size | **2**（OOM 就降到 1） | WebUI 默认 = 显存GB÷2，6 GB 给 2 |
| total_epoch | **8～12** | 素材仅 17 分钟，15 以上容易过拟合 |
| 保存频率 | 4 | 多存几个，方便挑最好的 |
| 精度 | fp16 | 6 GB 必须开，否则显存不够 |

耗时估算：**1.5～3 小时**（6 GB 是最慢的一档，主要慢在 batch=2）。

## 10. 训练 GPT（后训这个）

| 参数 | 6 GB 建议 |
|---|---|
| batch_size | **2**（OOM 降到 1） |
| total_epoch | **10～15** |
| 保存频率 | 4 |
| 精度 | fp16 |

耗时估算：**1～2 小时**。

## 11. 试听与选权重

1. 参考音频挑一句：建议从 `S04`（矛盾三种形态，254 字/分，情绪平稳）或 `S10`（五级表，260 字/分）里截 **5～8 秒**、一句话完整、不含停顿的片段。存成 `ref.wav`。
2. 参考文本填那一句的原话（必须逐字一致）。
3. 用课件第 5 页那段话做 A/B：

> 这一步是分水岭。技术矛盾是症状，物理矛盾是病灶。怎么走？把约束往下追，一直追到它落在同一个对象上。

4. 同时试 model 的**中间 epoch**（比如 e8 / e12 / e16）——不是越靠后越好，最后一轮常常发飘。

判断标准：**主观相似度 1～5 分 + 有没有念错字 + 停顿是否自然**。不要用 F0 曲线自欺，那是给论文用的，不是给耳朵用的。

## 12. 导出

```
SoVITS_weights_v2Pro/<时间戳>_e12_sxxx.pth
GPT_weights_v2Pro/<时间戳>-e15.ckpt
```

这两个文件拷回来就行，推理端配 `tts_infer.yaml` 就能用。

---

## 6 GB 显存不足时的处置顺序（按代价从小到大）

1. batch_size 2 → 1
2. total_epoch 降到 6（先跑通，再谈效果）
3. 切片上限从 10 秒压到 8 秒（显存与片长近似平方相关）
4. 关掉「保存所有 epoch」，只留最新
5. Windows 虚拟内存加到 32 GB、关掉浏览器与杀毒实时扫描
6. 还是 OOM → 说明该走 LoRA（但 V2Pro 没 LoRA，等于要换 V4，而 V4 要 8 GB）→ **换机器或租卡**

## 常见报错

| 报错 | 原因 | 处理 |
|---|---|---|
| `CUDA out of memory` | batch 太大 | 降到 1；关其他占显存程序 |
| `参考音频在3~10秒范围外` | 参考音频长度不合规 | 裁到 3～10 秒（16 kHz 下 48000～160000 样本） |
| `页面文件太小` | Windows 虚拟内存不足 | 设 20 GB 以上 |
| 训练 loss 不降 | list 文本与音频不匹配 | 回第 6 步校对 |
| 合成出来像"念稿机" | 素材本身是讲义腔 | 正常，本包就是讲义素材 |

## 时间总账（估算）

| 阶段 | 时间 |
|---|---|
| 装环境 + 下 5.3 GB 权重 | 0.5～1 h（看网速） |
| 切分 + ASR 打标 | 20～40 min |
| 校对 200 条 | 30～60 min（**不要省**） |
| SoVITS 微调 | 1.5～3 h |
| GPT 微调 | 1～2 h |
| 试听挑权重 | 20 min |
| **合计** | **半天到一天** |
"""
W("03_训练步骤", "训练步骤-6GB显卡.md", step)

# ---------------- 参数速查 ----------------
quick = """# 参数速查：显存 → batch / epoch

## batch_size 的默认口径

| 版本线 | WebUI 默认 | 说明 |
|---|---|---|
| V1 / V2 / V2Pro / V2ProPlus | batch = 显存GB ÷ 2 | 6 GB → 2；8 GB → 4；12 GB → 6；24 GB → 12 |
| V3 / V4 | batch = 显存GB ÷ 8 | 24 GB → 3；32 GB → 4 |

## 各档显存的实际建议（本素材 17 分钟）

| 显存 | 可行版本 | SoVITS batch / epoch | GPT batch / epoch | 预计训练时长 |
|---|---|---|---|---|
| 4 GB | 勉强 V2（会回落 CPU） | 1 / 6 | 1 / 8 | 不建议 |
| **6 GB** | **V2Pro / V2ProPlus** | **2 / 8–12** | **2 / 10–15** | **3～5 h** |
| 8 GB | V2Pro（或 V4 走 LoRA） | 4 / 8–12 | 4 / 10–15 | 2～3 h |
| 12 GB | V4 开梯度检查点 | 6 / 10 | 6 / 12 | 1.5～2 h |
| 24 GB | V2Pro 或 V4 全量 | 12 / 10 | 12 / 12 | 1 h 内 |

## 硬门槛（写在代码里，绕不过）

- 显存 < 4 GB → `config.py` 直接回落 CPU
- Pascal（Tesla P40 / GTX 1080）与 GTX 16 系 → 强制 float32，速度减半、可用 batch 减半
- 参考音频必须 3～10 秒；16 kHz 下即 48000～160000 个采样点
- 语言代码只认 `zh` / `ja` / `en` / `ko` / `yue`

## 与训练无关但会踩的坑

长文本推理显存不释放（issue #2458 从 2 GB 涨到 7 GB；#1519 涨到 29 GB）。**训练不受影响**，但将来做常驻 TTS 服务要按段切分或定期重启进程。
"""
W("03_训练步骤", "参数速查.md", quick)

# ---------------- 素材体检 ----------------
rows_md = "\n".join(
    "| %s | %s | %.1f | %d | %.0f | %s |" % (
        r["id"], r["title"][:18], r["sec"], r["chars"], r["cpm"],
        "偏慢（英文/数字多，可留可去）" if r["id"] == "S00"
        else ("偏快" if r["cpm"] > 280 else "正常"))
    for r in st)
health = f"""# 素材体检

## 总量

| 项 | 值 |
|---|---|
| 段数 | 16 |
| 总时长（已去首尾静音） | **{tot_s/60:.2f} 分钟**（{tot_s:.0f} 秒）|
| 总字数 | **{tot_c} 字** |
| 平均语速 | **{cpm:.1f} 字/分钟** |
| 格式 | WAV / 44.1 kHz / 单声道 / 16 bit |
| 体积 | {sum(r['bytes'] for r in st)/1048576:.0f} MB |

对照官方口径：5 秒可零样本，1 分钟可微调；**17 分钟属于"够用且有余量"**，再往上加收益递减。

## 逐段

| 段 | 内容 | 秒 | 字数 | 字/分 | 判断 |
|---|---|---|---|---|---|
{rows_md}

## 结论与建议

- **全部 16 段都可以用**，不需要剔除。
- `S00` 是热身段，混了英文缩写和数字（TRIZ、ARIZ、MPV、200 毫秒、2040 条），语速只有 200 字/分。留着能让模型学会这些专名的读法；如果你只想要纯中文语感，可以去掉，剩下 15.3 分钟仍然充足。
- `S13`（296 字/分）和 `S15`（282 字/分）偏快，但只有两段，占比 14%，影响有限。**语速波动本身对模型是好事**——太均匀的素材训出来的声音像念经。
- 素材是**讲义腔**，不是口语闲聊。这决定了模型擅长的场景：念讲稿、念课程内容。不要指望它模仿聊天。

## 需要你人工听的三件事（我无法自动判定）

1. **有没有爆音/口水音**：第 2、7、10 段最长，出问题概率最高。有就重录那一段，别用降噪硬救。
2. **有没有明显底噪**：手机贴近嘴边录通常会有一点。开大音量听片尾静音段——如果有持续的"嘶——"，考虑轻度降噪，但必须 A/B。
3. **有没有读错字后自己重读**：模型会把重读也学进去。发现就剪掉重读部分，或整段重录。

## 已知的一处不一致

`S07` 录的是旧稿，里面有「鳍片」；现在课件和讲稿已统一成「散热片」。**如果要严格一致，重录 S07（约 40 秒）**，再替换 `01_音频/S07.wav`。不重录也不影响训练，只是模型会多学一个"鳍片"。
"""
W("04_校验", "素材体检.md", health)
print("docs written")

# ---------------- zip ----------------
zpath = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音训练交接包-20261029.zip"
if os.path.exists(zpath): os.remove(zpath)
with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
    for root, _, files in os.walk(BASE):
        for fn in files:
            if fn.startswith("_"): continue
            full = os.path.join(root, fn)
            z.write(full, os.path.join("TRIZa-声音训练交接包", os.path.relpath(full, BASE)))
print("zip:", round(os.path.getsize(zpath)/1048576, 1), "MB")
for root, _, files in os.walk(BASE):
    for fn in sorted(files):
        if fn.startswith("_"): continue
        full = os.path.join(root, fn)
        print("  ", os.path.relpath(full, BASE), round(os.path.getsize(full)/1024, 1), "KB")