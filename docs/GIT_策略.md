# GIT 策略 · TRIZ 讲义仓

## 1. 仓库定位

仓库根 = `D:\doc\ZGCA\Lectures\TRIZ`，只对「讲义生产线」做版本管理：

- 生成课件的脚本（`PPT/W2_16/*.py`、`PPT/build_w2.py`、审计脚本）
- 课件成品（`PPT/W2_16/out/*.pptx`，单份约 75 KB）
- 规范与文档（`PPT/TRIZa-Lecture-SKILL.md`、`docs/`、`style/`）
- 教学大纲与讲稿集（`.docx`，小体积）

录音、抓取网页、PDF 文献、TTS 中间音频**不入库**，它们留在磁盘上，靠 `.gitignore` 挡住。

## 2. 分支与标签

| 用途 | 名称 | 说明 |
| --- | --- | --- |
| 主线 | `main` | 始终可重建、可校验通过 |
| 单讲改动 | `deck/s03-causal` | 一次只改一讲，合并前必须跑 `w_final.py` |
| 批量重构 | `refactor/kit` | 动 `w_kit.py` 骨架时用 |
| 里程碑 | `v0.1` `v1.0` | 交付给学生的版本打标签 |

标签语义：`v0.x` = 内部迭代；`v1.0` = 16 讲齐备、五项校验全绿、可用于开课。

## 3. 提交信息

```
<类型>: <一句话>

类型：deck / kit / notes / skill / docs / git
例：
deck: 第4次课补齐反向作用，原理条数统一为六条
kit: 表格行高改 0.50，容纳七行
notes: 第7次课讲稿补 60 字到 3041
skill: 新增例图规范与条数一致性规则
```

一次提交只做一件事。改课件就只提课件，别把 `w_kit.py` 的重构混进来。

## 4. 哪些文件可以推

| 类型 | 推 | 理由 |
| --- | --- | --- |
| `.py` `.ps1` `.md` `.json`（<100KB） | ✅ | 文本，diff 可读，是资产本体 |
| `.pptx`（<2MB） | ✅ | 成品就是交付物，75KB 完全可接受 |
| `.docx` 大纲/讲稿集（<2MB） | ✅ | 教学依据 |
| `.wav` `.mp3` `.m4a` | ❌ | 832 MB，且可由脚本重建 |
| `.zip` 交接包/训练包 | ❌ | 449 MB，走网盘或 Release |
| `.pdf` `.epub` 文献 | ❌ | 271 MB，且非本仓产物 |
| `.html` `.css` 抓取网页 | ❌ | 噪音，随时可重抓 |
| `*_备份-*.pptx` | ❌ | git 已承担版本职责，文件级备份冗余 |
| `TRIZ_W1_矛盾15m_演播版.pptx`（16.9 MB） | ❌ | 超阈值，走 Release 附件 |

超过 2 MB 的成品一律走 GitHub/Gitee 的 Release 附件，仓库里只留一行 `artifacts/README.md` 记录 SHA-256 与下载位置。

如果以后确实要把大课件纳入版本管理，再启 Git LFS（`git lfs track "*.pptx"`），不要用普通提交硬塞。

## 5. 远端

已配置本地裸仓，可离线演练完整推送：

```
origin  D:/doc/ZGCA/Lectures/_remotes/triz-lecture.git
```

换成真实托管（二选一，先建空仓，不要勾选 README）：

```bash
# GitHub（私有仓）
git remote set-url origin git@github.com:<账号>/triz-lecture.git
git push -u origin main --tags

# Gitee（国内更快）
git remote set-url origin git@gitee.com:<账号>/triz-lecture.git
git push -u origin main --tags
```

推之前确认 `git status` 里没有媒体文件：

```bash
git ls-files -s | Measure-Object -Line   # 文件数
git ls-files | ForEach-Object { Get-Item $_ } | Sort-Object Length -Desc | Select-Object -First 5
```

## 6. 日常流程

```powershell
# 1. 改内容模块
# 2. 重建
python PPT\W2_16\w_build.py
# 3. 版式检查
python PPT\W2_16\w_fit.py PPT\W2_16\out
# 4. 五项校验
python PPT\W2_16\w_final.py
# 5. 只有全绿才提交
git add PPT/W2_16
git commit -m "deck: ..."
git push
```

校验不通过就提交，等于把坏版本写进历史。第 4 步是门禁。

## 7. 排障

- `detected dubious ownership`：D 盘不记录属主，已加 `git config --global --add safe.directory`。
- 中文文件名乱码：已设 `core.quotepath false`。
- 换行符：`.gitattributes` 固定 `*.py eol=lf`、`*.ps1 eol=crlf`，避免整文件 diff。