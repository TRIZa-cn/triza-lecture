---
name: pptx-shape-diagrams
description: 用 python-pptx 原生形状画示意图，并让动画在课堂版点击触发、在配音演播版按语音时间自动播放。当你需要把讲稿里的"功能图/关系图"变成幻灯片上可动画的真实形状，或要让 pptx 的入场动画跟旁白对齐时使用。
---

# 用 pptx 形状画示意图 + 配音同步动画

python-pptx 没有动画 API，但 PowerPoint 的动画就是 slide XML 里的 `<p:timing>`。
所以做法是：**形状用 python-pptx 画，动画手写 XML 注入**。这样每页都能两种活法——
课堂版点一下走一格，演播版自己不点也会按旁白的节奏走。

## 为什么不直接出 PNG

PNG 在这种场景下有三个硬伤：
1. 一张图只能整体出现，做不到"最后那条线飞进来"；
2. 1280 宽以上的投影下会糊，改一个字要重新出图；
3. 配色和字体跟别的页对不上，一眼看出是外挂的。

原生形状的代价是位置要手算。手算一次，之后改文案不用动图。

## 第 1 步：先让图服从讲稿，不是反过来

**只画讲稿里讲成"功能句"的线和动词。**

讲稿里出现一个组件名，不等于图上该出现一条线。判断标准：
这句话有没有说出"谁 对 谁 做了 什么"。

反例（真实踩过）：一份图的规格列了 8 条线，但录音里明确讲成功能句的只有 5 条。
把剩下 3 个动词补上去，观众讲到第 5 条时会开始找那个还没讲的东西，
注意力就从"缺了一条"跑到"这是什么"上了。宁可留白。

同理：图里不出现数字和百分比，除非讲稿里就有。
示意图像统计图，是非常难收回来的一种误导。

## 图的出场顺序就是叙事顺序

一张图分几次出现，观众就被引导着走几段。两条规矩：

- **先破后立。** 先给"错的那半张"，停一下让人看到它不够，再补上对的那一笔。
  实测里 A 页 就这么做：五个组依次入场，讲完"这些都画齐了"，最后那条缺的线才飞进来。
- **关键的那一笔必须最后落。** 那条线要是提前出现，"图上根本就不存在"这句话就废了。
  写稿时先定哪一笔是最后一笔，再倒推每一组对应哪一句台词。

反过来：如果一页里所有东西本来就该同时出现，**就别加动画**。动画是用来控制注意力顺序的，
不是用来"显得高级"的。同一页点五次才显示完，观众会开始烦。

## 第 2 步：先读现有版面，再决定动谁

不要从空白页开始想坐标。先读目标页已经有什么：

```python
from pptx import Presentation
from pptx.util import Inches
prs = Presentation(DECK)
s = prs.slides[N - 1]
for sh in s.shapes:
    print(sh.shape_id, sh.shape_type, sh.name,
          round(sh.left / 914400, 2), round(sh.top / 914400, 2),
          round(sh.width / 914400, 2), round(sh.height / 914400, 2),
          (sh.text_frame.text[:20] if sh.has_text_frame else ""))
```

十四页的课件，标题区、页脚、分隔线通常已经有一套固定位置。
图应该长在**原本放文字的那块地**上，而不是挤进空白里。
两种常见改法：

- **整块替换**：那一页左边本来是一列步骤文字，删掉，原位画图（讲稿不变，字没了但话说得更多）。
- **收窄让位**：大字标题保持，把它的 `width` 改小，右边腾出来的地方画图。

## 第 3 步：画。只用三种原语

```python
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn
from lxml import etree

def node(grp, x, y, w, h, label, size=18, color=INK, border=INK, dash=False, fill=WHITE):
    b = grp.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             Inches(x), Inches(y), Inches(w), Inches(h))
    b.fill.solid(); b.fill.fore_color.rgb = fill
    b.line.color.rgb = border; b.line.width = Pt(1.25)
    if dash:
        b.line.dash_style = MSO_LINE.DASH
    b.shadow.inherit = False          # 不关掉会有难看的默认阴影
    tf = b.text_frame
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = label; _font(r, size, color, True)
    return b

def arrow(grp, x1, y1, x2, y2, color=INK, width=1.75, dash=False, head=True):
    c = grp.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                 Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color; c.line.width = Pt(width)
    if dash:
        c.line.dash_style = MSO_LINE.DASH
    if head:                          # python-pptx 没有箭头 API，手写 a:tailEnd
        ln = c.line._get_or_add_ln()
        t = etree.SubElement(ln, qn("a:tailEnd"))
        t.set("type", "triangle"); t.set("w", "med"); t.set("len", "med")
    return c
```

**分组是必须的**，不是整理癖。动画的单位是组：一条"从 a 到 b 的线 + 线上那个动词 + 端点的小叉"
必须是一个组，否则它们会各飞各的。

```python
g = slide.shapes.add_group_shape()
node(g, ...); arrow(g, ...)       # 组内坐标 = 幻灯片坐标，不用换算
g.shape_id                        # 动画要用的 id
```

中文字体必须同时设 `a:ea` 和 `a:cs`，否则 PowerPoint 会给中文回退到宋体：

```python
def _font(run, size, color, bold=False, name="微软雅黑"):
    f = run.font
    f.size = Pt(size); f.bold = bold; f.color.rgb = color; f.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag)) or etree.SubElement(rPr, qn(tag))
        el.set("typeface", name)
```

## 第 4 步：动画 XML

一个入场效果在 XML 里长这样（`presetID` 决定效果，`nodeType` 决定触发方式）：

| 想要 | presetID | 备注 |
|---|---|---|
| 淡入 | 10 | 最通用 |
| 飞入 | 2 | `presetSubtype` 2=右 8=左 4=下 |
| 擦除 | 22 | "线画出来"的感觉，适合箭头 |

每个效果内部 = 一个 `p:set`（把 `style.visibility` 置为 visible）+ 位移/透明度动画：

```xml
<p:par><p:cTn id="N" presetID="2" presetClass="entr" presetSubtype="2" fill="hold"
              grpId="0" nodeType="clickEffect">
  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
  <p:childTnLst>
    <p:set><p:cBhvr><p:cTn id="N+1" dur="1" fill="hold">
      <p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
      <p:tgtEl><p:spTgt spid="12"/></p:tgtEl>
      <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
    </p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
    <p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">
      <p:cTn id="N+2" dur="800" fill="hold"/>
      <p:tgtEl><p:spTgt spid="12"/></p:tgtEl>
      <p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>
      <p:tavLst>
        <p:tav tm="0"><p:val><p:strVal val="1+#ppt_w/2"/></p:val></p:tav>
        <p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>
      </p:tavLst></p:anim>
  </p:childTnLst>
</p:cTn></p:par>
```

外面还要套一层完整骨架，而且 `p:cTn id` 全局唯一，别和 `p:spTgt spid` 搞混：

```xml
<p:timing><p:tnLst><p:par>
  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
    <p:seq concurrent="1" nextAc="seek">
      <p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
        <!-- 每个效果一层 par，外面套 stCondLst delay="indefinite" 就是"等一次点击" -->
      </p:childTnLst></p:cTn>
      <p:prevCondLst>...</p:prevCondLst><p:nextCondLst>...</p:nextCondLst>
    </p:seq>
  </p:childTnLst></p:cTn>
</p:par></p:tnLst>
<p:bldLst><p:bldP spid="12" grpId="0" animBg="1"/></p:bldLst>
</p:timing>
```

插入位置：`</p:extLst>` 之前，没有 `p:extLst` 就 append 到 `p:sld`。
**注入前先删旧的** `sld.findall(qn("p:timing"))`，否则脚本跑第二遍就两份动画。

## 第 5 步：让同一套动画在演播版里自己走

课堂版：`nodeType="clickEffect"`，外层 `delay="indefinite"` → 一次点击一个效果。

演播版：`nodeType="withEffect"`，`delay` 给毫秒绝对值 → 音频一开始就排好了队。

`delay` 怎么算？**不要去猜秒数，用旁白里的锚点字符串**：

```python
# 锚点 = 这句话里的一句片段，复制粘贴，不要手打
pos = narration_text.find(anchor)
delay_ms = int(page_audio_ms * pos / len(narration_text))
```

假设语速在页内基本均匀——这是唯一的一个假设，要写进注释。
中文 TTS 和真人朗读在一页之内的字速率波动通常在可接受范围，
但如果某页有大段停顿，就得改成按句切音频、按句累计。

注入点：演播版的 timing 由"播放音频"那个 `p:cMediaNode` 撑起来，
找到它的收尾 `</p:cmd></p:childTnLst></p:cTn></p:par>`，把效果 par 追加在后面，
再把 `p:bldLst` 补到 `</p:tnLst>` 之后。这样音频和形状在同一个 `mainSeq` 里，
PowerPoint 才会认为"这些是自动的"，放映时不出现点击提示。

一个可验证的判据：用 COM 读回来，音频效果应该是 `TriggerType=3`（上一动画之后），
形状效果应该是 `TriggerType=2`（与上一动画同时）且 `TriggerDelayTime` 各不相同。
如果形状也是 3，说明被排成了串行，放起来会一卡一卡。

## 第 6 步：验证。不能只看图

**能看图的模型**：导出 PNG 看。PowerPoint COM 一行就够：

```powershell
$p.Slides.Item(8).Export("$env:TEMP\s8.png","PNG",1600,900)
```

**不能看图的模型**（Web GUI 里常见）：用几何数值代替眼睛。三个必查：

1. **越界**：`left + width > 幻灯片宽`。13.333in 的版面就是 960×540 pt。
2. **重叠**：新画的东西有没有压到结论框、页脚、分隔线。
   注意组的边界框和组内实际形状是两回事——左右两组超系统组件的边界框可能叠在一起，
   但它们的成员分别贴在页面两侧，视觉上并不冲突。判重叠要看到子项。
3. **效果链**：`MainSequence.Count`，以及每个效果的 `Shape.Name` 和 `TriggerDelayTime`。
   这一步能抓出"组名对上了但 id 错位"的问题。

```powershell
$seq = $p.Slides.Item(8).TimeLine.MainSequence
foreach ($k in 1..$seq.Count) {
  $e = $seq.Item($k)
  "{0} {1} trig={2} delay={3}" -f $k, $e.Shape.Name, $e.Timing.TriggerType, $e.Timing.TriggerDelayTime
}
```

动画数对不上、或者 `Shape.Name` 全是 `Group N` 而你分不清谁是谁时，
给组起名字：`g.name = "A 页 超系统"`。

## 会踩的坑

- **脚本要幂等，而且从干净副本起步。** 第一次运行前把原始 pptx 备份一份，
  之后每次都从备份重建，而不是从上次的产物上叠加。否则重跑一次就是两套形状。
- **目标文件被 PowerPoint 打开时写不进去**（`PermissionError`）。
  这不是代码错。加个环境变量让脚本能输出到别处先验证，正式文件等文件关掉再写：
  `DST = os.environ.get("OUT") or 默认路径`。
- **exFAT 盘上不能用需要硬链接的原子写**（某些工具会报 `EISDIR`）。
  写到本地临时目录再 `Copy-Item` 过去。
- **`p:cTn id` 必须全局唯一**，注入时从一个明显的高位起编（比如 100），别从 1 开始。
- **`p:spTgt spid` 用的是 `shape_id`，不是 shapes 列表下标。**
  `shapes.add_group_shape()` 返回的组，`shape_id` 要等画完子形状后再读也来得及，
  但动画 XML 一定要用最终这个 id。
- **虚线要配箭头时**，`dash_style` 设在 `a:ln` 上，`tailEnd` 也是 `a:ln` 的子元素，
  两者可以共存，顺序是 `noFill/solidFill` → `prstDash` → `headEnd/tailEnd`。

## 交付时该一起交的东西

1. 画图的脚本本身（别人改文案时要能重跑）。
2. 锚点与算出来的毫秒（一个 json），出了问题能立刻看出是"锚点找错了"还是"语速假设崩了"。
3. 一句诚实的说明：图形是否经过人眼确认。数值检查能证明没越界、动画链对，
   但证明不了"好不好看"。没看过就别写"已确认版式无误"。
