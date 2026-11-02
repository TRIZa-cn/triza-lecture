# -*- coding: utf-8 -*-
import sys, os, glob, re
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16")
import _l16_render as R
import _l16_spec_a as A, _l16_spec_a2 as A2, _l16_spec_b as B, _l16_spec_c as C

FIX = {
 4: {"col_w": [0.8, 2.0, 3.0, 0.8, 2.0, 3.03], "row_h": 0.6, "rows": [
    ["编号", "原理", "一句话动作", "编号", "原理", "一句话动作"],
    ["01", "分割", "拆成互不影响的块", "02", "抽取", "把起作用的部件单独拿走"],
    ["03", "局部质量", "各部分按需各不相同", "10", "预先作用", "把动作提前做完"],
    ["15", "动态化", "能变、能适应工况", "24", "中介物", "中间加一个传递载体"],
    ["25", "自服务", "系统自己维护自己", "28", "机械系统替代", "用光、声、电磁替代机械动作"],
    ["35", "参数变化", "改状态、浓度、温度、柔性", "", "", ""]]},
 7: {"col_w": [2.2, 3.6, 2.2, 3.63], "row_h": 0.6, "rows": [
    ["法则", "一句话", "法则", "一句话"],
    ["系统完备性", "缺的零件早晚补齐", "能量传递", "能量流顺畅，少转换"],
    ["协调性", "尺寸、节奏、材料要匹配", "向理想化", "功能变强，代价变小"],
    ["子系统不均衡", "最弱的那个先被改", "向微观级", "从机械走到分子与场"],
    ["向超系统", "干不动就并进去", "动态化", "从固定到可调"]]},
 11: {"col_w": [1.6, 4.2, 1.6, 4.23], "row_h": 0.66, "rows": [
    ["类别", "包含什么", "类别", "包含什么"],
    ["物质", "材料、废弃物、下脚料", "场（能量）", "热、振动、电磁、声、光"],
    ["时间", "等待、空转、维护窗口", "空间", "空隙、边缘、表面、地下"],
    ["信息", "日志、标签、历史数据", "功能", "系统已有的动作能力"]]},
 15: {"col_w": [3.2, 2.6, 3.2, 2.63], "row_h": 0.66, "rows": [
    ["假设", "怎么验", "假设", "怎么验"],
    ["时间是对的", "拨时钟", "权限是最新的", "改一次看延迟"],
    ["缓存是一致的", "并发写后立刻读", "网络是通的", "断网 30 秒"],
    ["日志是完整的", "两处日志对账", "配置是生效的", "改一条无害配置"],
    ["重启是无害的", "随机重启一次", "", ""]]},
}

ALL = A.SPECS + A2.SPECS + B.SPECS + C.SPECS
for sp in ALL:
    if sp["no"] in FIX:
        for pg in sp["pages"]:
            if pg["kind"] == "table":
                pg.update(FIX[sp["no"]])
                print("patched table in lecture", sp["no"])
        R.build_one(sp)

print("\nre-check table rows:")
for f in sorted(glob.glob(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16\*.pptx")):
    prs = R.Presentation(f)
    mx = max((len(sh.table.rows) for s in prs.slides for sh in s.shapes if sh.has_table), default=0)
    print(f"  {os.path.basename(f):<38} 最多表格行 {mx}")