# TRIZa 讲义仓 · 只放最终演播版

本仓是**白名单仓**：默认忽略一切，只跟踪 `playablePPT/` 下的最终演播版课件。

## 为什么只放演播版

课件的源码、生成脚本、讲稿与校验工具留在本地工作盘，不对外发布。
对外只发布可以直接播放的**最终演播版 PPT**，由作者本人手工放入 `playablePPT/` 后提交。

## 目录

| 路径 | 内容 |
| --- | --- |
| `playablePPT/*.pptx` | 最终演播版课件（唯一允许入库的成品） |
| `.gitignore` | 白名单规则 |
| `LICENSE` | MIT |

## 怎么加一份演播版

```powershell
# 1. 自己把最终演播版拖进 playablePPT/
git add playablePPT/xxx_演播版.pptx
git commit -m "feat: 新增第 N 次课演播版"
git push
```

## 许可

MIT，见 `LICENSE`。