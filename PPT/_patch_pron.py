# -*- coding: utf-8 -*-
from pathlib import Path
p = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZa-Lecture-SKILL.md")
t = p.read_text(encoding="utf-8")
anchor = "## 讲稿：写给人说的，不是给人看的"
assert anchor in t
block = """### 专名读音：先定下来，再写稿

拉丁缩写是口播事故的高发区：合成语音常逐字母念（"T-R-I-Z"），或按英文拼读规则念错。讲稿定稿前先把专名读音写死，写进术语表。

| 专名 | 读音 | 近似读法 | 常见错读 |
|---|---|---|---|
| TRIZ | /triːz/ | "treez"，尾音是 z 不是 s | 逐字母 T-R-I-Z；读成 /traɪz/（"try-z"） |
| TRIZa | /triːza/ | "treeza"，末尾多一个轻 ə | 逐字母 T-R-I-Z-A；读成 /traɪzə/ |

实操三条：

1. 讲稿首次出现时写成"TRIZ（读 treez）"或直接写中文名"萃智"，并确认两者指同一件事。中文语境里"萃智"是 TRIZ 的通行译名，但首次出现要给出对应关系，不要默认听众知道。
2. 合成后必须试听这一处。TTS 对同一缩写在不同上下文可能念得不一样，只有听才算核验。
3. 其他缩写（FOS、ARIZ、V2 之类）按同样办法先定读音；定不下来就改用中文全称，不要赌合成引擎。

"""
t = t.replace(anchor, block + anchor, 1)
t = t.replace("| 9 | 数字有来源 | 每个数字可追到出处 |",
              "| 9 | 数字有来源 | 每个数字可追到出处 |\n| 11 | 专名读音已定 | TRIZ=/triːz/、TRIZa=/triːza/，且已试听 |", 1)
p.write_text(t, encoding="utf-8")
print("skill patched", len(t))