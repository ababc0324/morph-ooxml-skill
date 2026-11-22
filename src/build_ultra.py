# -*- coding: utf-8 -*-
"""
Morph ULTRA：在 12 页 PRO 版基础上继续加码到 24 页 / 20 个演员 / 9 类新效果。

复用 build_pro 的前 12 帧关键帧，扩充对象池后新增：
  爆炸聚合、透视翻转（Z 轴错觉）、数字滚动 + 刻度盘 + 进度条、
  整页调色板渐变、双遮罩对切、4 页快闪卡点、单字放大成背景、
  套娃回环（首尾同帧可无限循环）。
"""
import copy
import html
import math
import re
import zipfile
from lxml import etree

import build_pro as bp

# 需要自备一个已含 9 页基础内容的 BASE pptx（本仓库不含示例交付物）。
BASE = "./平滑切换演示_Morph版.pptx"
DST = "./平滑切换演示_MorphULTRA版.pptx"

W, H = bp.W, bp.H
st, T, S, mod, esc = bp.st, bp.T, bp.S, bp.mod, bp.esc
C = dict(bp.C)
C.update({"orange": "FB923C", "pink": "F472B6", "amber": "FBBF24",
          "rose": "FB7185", "lime": "A3E635"})

P = bp.P
MC, P14, P159, A, R = bp.MC, bp.P14, bp.P159, bp.A, bp.R
CT_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# 20 个跨页恒定演员
ORDER = ["bg", "grid", "orb", "halo", "curtain", "curtain2", "ring", "dial", "bar",
         "card1", "card2", "card3", "card4",
         "spark1", "spark2", "spark3", "title", "sub", "num", "tag"]
IDS = {n: i + 2 for i, n in enumerate(ORDER)}

OLD_FILES = [1, 7, 4, 3, 8, 5, 11, 10, 12]
NEW_FILES = list(range(30, 30 + 15))          # 15 个新增物理页
SLIDE_FILES = OLD_FILES + NEW_FILES           # 24 页


def extra():
    """6 个新演员的默认待命态（画布外 / 全透明，不干扰既有画面）。"""
    return {
        "curtain2": st(-2 * W, 0, W, H, fill=C["indigo"], alpha=0),
        "dial":     st(-3000000, 2000000, 1600000, 1600000, fill=C["blue"],
                       alpha=0, geom="ellipse", text="", sz=1500, align="ctr", bold=True),
        "bar":      st(900000, -1000000, 1000000, 120000, fill=C["teal"],
                       alpha=0, geom="roundRect", text="", sz=1000),
        "spark1":   st(-2000000, 3000000, 260000, 260000, fill=C["white"],
                       alpha=0, geom="ellipse", text="", sz=800),
        "spark2":   st(-2000000, 3400000, 180000, 180000, fill=C["blue"],
                       alpha=0, geom="ellipse", text="", sz=800),
        "spark3":   st(-2000000, 3800000, 220000, 220000, fill=C["pink"],
                       alpha=0, geom="ellipse", text="", sz=800),
    }


# ---------- 继承 PRO 版 12 帧，补齐新演员 ----------
FR = []
for f in bp.FR:
    g = copy.deepcopy(f)
    g.update(extra())
    FR.append(g)


def cap(i, title, sub, label="MORPH / ULTRA"):
    return dict(
        title=T(900000, 620000, title, sz=3400, bold=True),
        sub=S(920000, 1480000, sub, sz=1750),
        num=dict(text=f"{i:02d} / 24", x=920000, y=5950000, cx=4200000),
        tag=dict(text=label, x=9900000, y=380000, cx=1900000, cy=350000,
                 color=C["dim"], sz=1100, align="r"),
    )


def numbered(frame, i, title, sub, label="MORPH / ULTRA", **kw):
    f = mod(frame, **cap(i, title, sub, label))
    for k, v in kw.items():
        f[k].update(v)
    return f


# 前 12 页页码 / 角标统一到 ULTRA 版
for i, f in enumerate(FR, 1):
    f["num"].update(text=f"{i:02d} / 24")
    if f["tag"].get("text") == "FLOW MOTION":
        f["tag"].update(text="MORPH / ULTRA")
FR[0]["title"].update(text="平滑切换 ULTRA", sz=5600)
FR[0]["sub"].update(text="MORPH · 24 页 / 20 对象 / 9 类新增效果")

# ---------- 页 12 改写：单字放大成背景 + 聚合待命（为页13爆炸） ----------
center = W // 2, H // 2
FR[11] = numbered(
    mod(FR[11],
        grid=dict(x=1200000, y=-600000, cx=10000000, cy=8000000,
                  text="顺", sz=52000, color="1B3E6B", align="ctr", bold=True,
                  alpha=0, geom="rect"),
        orb=dict(x=3800000, y=1100000, cx=4600000, cy=4600000, alpha=30000),
        halo=dict(x=6400000, y=2700000, cx=3600000, cy=3600000, alpha=24000),
        ring=dict(x=4700000, y=1800000, cx=2800000, cy=2800000, alpha=16000),
        card1=dict(x=center[0] - 200000, y=center[1] - 150000,
                   cx=400000, cy=300000, rot=0, alpha=0, geom="roundRect", text=""),
        card2=dict(x=center[0] - 200000, y=center[1] - 150000,
                   cx=400000, cy=300000, rot=0, alpha=0, sz=1400, text=""),
        card3=dict(x=center[0] - 200000, y=center[1] - 150000,
                   cx=400000, cy=300000, rot=0, alpha=0, text=""),
        card4=dict(x=center[0] - 200000, y=center[1] - 150000,
                   cx=400000, cy=300000, rot=0, alpha=0, text=""),
        spark1=dict(x=center[0], y=center[1], alpha=0),
        spark2=dict(x=center[0], y=center[1], alpha=0),
        spark3=dict(x=center[0], y=center[1], alpha=0),
        title=T(900000, 2400000, "⑩ 把一个字放大成背景", sz=4400, cy=1000000),
        sub=S(920000, 3600000, "byChar · 字形本身也能跨页补间", sz=1900),
        tag=dict(text="MORPH / CHAR"),
    ), 12, "⑩ 把一个字放大成背景", "byChar · 字形本身也能跨页补间", "MORPH / CHAR",
)
# grid 是纯文字框：给一点可见度（渲染器对无填充文字用文字自身透明度）
FR[11]["grid"].update(alpha=100000)

# ---------- 页 13 【新】爆炸聚合：中心一点 → 全屏散开 ----------
FR.append(numbered(
    mod(FR[-1],
        grid=dict(x=-16000000, y=-12000000, alpha=0, text=""),
        card1=dict(x=900000, y=2100000, cx=2400000, cy=1600000, rot=-14,
                   alpha=90000, fill=C["blue"], text="爆", sz=2200),
        card2=dict(x=4200000, y=1700000, cx=3000000, cy=1900000, rot=8,
                   alpha=90000, fill=C["purple"], text="炸", sz=2400),
        card3=dict(x=7900000, y=2400000, cx=2400000, cy=1600000, rot=-8,
                   alpha=88000, fill=C["teal"], text="散", sz=2200),
        card4=dict(x=4500000, y=4400000, cx=2800000, cy=1400000, rot=6,
                   alpha=85000, fill=C["indigo"], text="开", sz=2200),
        spark1=dict(x=2200000, y=1200000, cx=320000, cy=320000, alpha=95000),
        spark2=dict(x=8500000, y=1500000, cx=240000, cy=240000, alpha=90000),
        spark3=dict(x=9800000, y=4800000, cx=300000, cy=300000, alpha=90000),
        orb=dict(x=-1600000, y=3600000, cx=3400000, cy=3400000, alpha=22000),
        halo=dict(x=9200000, y=-1000000, cx=3000000, cy=3000000, alpha=24000),
        ring=dict(x=4000000, y=1200000, cx=4200000, cy=4200000, alpha=14000),
    ), 13, "⑪ 爆炸：同一批对象从一点散开",
    "前页全部叠在中心 · 这页只改坐标、尺寸和透明度", "MORPH / EXPLODE"))

# ---------- 页 14 【新】透视翻转：近大远小的 Z 轴队列 ----------
persp = [(2300000, 900000, 5400000, 3000000, 0, 95000),
         (3000000, 2200000, 4200000, 2300000, 0, 78000),
         (3600000, 3300000, 3200000, 1700000, 0, 60000),
         (4100000, 4200000, 2400000, 1200000, 0, 46000)]
FR.append(numbered(
    mod(FR[-1],
        card1=dict(x=persp[0][0], y=persp[0][1], cx=persp[0][2], cy=persp[0][3],
                   rot=persp[0][4], alpha=persp[0][5], geom="trapezoid",
                   sz=1800, text="近"),
        card2=dict(x=persp[1][0], y=persp[1][1], cx=persp[1][2], cy=persp[1][3],
                   rot=persp[1][4], alpha=persp[1][5], geom="trapezoid",
                   sz=1700, text="中"),
        card3=dict(x=persp[2][0], y=persp[2][1], cx=persp[2][2], cy=persp[2][3],
                   rot=persp[2][4], alpha=persp[2][5], geom="trapezoid",
                   sz=1500, text="远"),
        card4=dict(x=persp[3][0], y=persp[3][1], cx=persp[3][2], cy=persp[3][3],
                   rot=persp[3][4], alpha=persp[3][5], geom="trapezoid",
                   sz=1400, text="深"),
        spark1=dict(x=1600000, y=1600000, alpha=0),
        spark2=dict(x=10600000, y=1400000, alpha=0),
        spark3=dict(x=10800000, y=4900000, alpha=0),
        orb=dict(x=4200000, y=-2600000, cx=3800000, cy=3800000, alpha=18000),
        halo=dict(x=8800000, y=4200000, cx=3000000, cy=3000000, alpha=22000),
    ), 14, "⑫ 透视：缩放差就是 Z 轴",
    "四层梯形递缩 + 透明度衰减 · 不用 3D 也有纵深", "MORPH / DEPTH"))

# ---------- 页 15~17 【新】数字滚动 + 刻度盘 + 进度条 ----------
def gauge(pct, rot, dial_text, i, title, sub, label="MORPH / DATA"):
    return numbered(
        mod(FR[-1],
            card1=dict(x=1400000, y=2500000, cx=2100000, cy=1400000, rot=0,
                       alpha=80000, geom="roundRect", fill=C["blue"], sz=1600, text="位移"),
            card2=dict(x=4100000, y=2200000, cx=2100000, cy=1700000, rot=0,
                       alpha=88000, geom="roundRect", fill=C["purple"], sz=1600, text="缩放"),
            card3=dict(x=6800000, y=2500000, cx=2100000, cy=1400000, rot=0,
                       alpha=80000, geom="roundRect", fill=C["teal"], sz=1600, text="旋转"),
            card4=dict(x=9500000, y=2200000, cx=2100000, cy=1700000, rot=0,
                       alpha=85000, geom="roundRect", fill=C["indigo"], sz=1600, text="颜色"),
            dial=dict(x=900000, y=3900000, cx=1500000, cy=1500000, rot=rot,
                      alpha=90000, geom="donut", fill=C["amber"],
                      text=dial_text, sz=1700, align="ctr", bold=True),
            bar=dict(x=900000, y=5700000, cx=int(10000000 * pct), cy=130000,
                     alpha=95000, geom="roundRect", fill=C["teal"], text="", sz=1000),
            ring=dict(x=8400000, y=3600000, cx=3000000, cy=3000000, alpha=16000,
                      geom="ellipse", fill=C["blue"]),
            orb=dict(x=-1400000, y=600000, cx=3000000, cy=3000000, alpha=20000),
            halo=dict(x=10000000, y=-1200000, cx=2800000, cy=2800000, alpha=22000),
            num=dict(text=f"{i:02d} / 24", x=920000, y=5950000, cx=4200000),
        ), i, title, sub, label)

FR.append(gauge(0.18, 0,   "18%", 15, "⑬ 数字自己会滚",
                "文本、角度、宽度一起插值 · 18% → 55% → 100%"))
FR.append(gauge(0.55, 45,  "55%", 16, "⑭ 刻度盘转起来",
                "同一个 !!dial 对象 · 旋转 + 数值 + 环形填充同时变"))
FR.append(gauge(1.00, 90, "100%", 17, "⑮ 进度条长满",
                "!!bar 只改 cx · 三页连起来就是一段数据动画"))
# 数字位也跟着滚（num 文本保持页码，另用 dial 承担数字）

# ---------- 页 18 【新】整页调色板渐变 + 双遮罩对切 ----------
warm = {
    "orb": C["orange"], "halo": C["rose"], "ring": C["amber"],
    "card1": C["orange"], "card2": C["rose"], "card3": C["lime"],
    "card4": C["pink"], "dial": C["amber"], "bar": C["orange"],
    "spark1": C["amber"], "spark2": C["rose"], "spark3": C["lime"],
}
FR.append(numbered(
    mod(FR[-1],
        curtain=dict(x=-W, y=0, cx=W, cy=H, alpha=52000, fill=C["rose"]),
        curtain2=dict(x=W + 200000, y=0, cx=W, cy=H, alpha=52000, fill=C["orange"]),
        card1=dict(x=1100000, y=2300000, cx=2400000, cy=1700000, rot=-8,
                   fill=C["orange"], alpha=90000, geom="roundRect", text="冷", sz=2000),
        card2=dict(x=4200000, y=2000000, cx=3200000, cy=2000000, rot=0,
                   fill=C["rose"], alpha=92000, geom="roundRect", text="→", sz=2600),
        card3=dict(x=8100000, y=2300000, cx=2400000, cy=1700000, rot=8,
                   fill=C["lime"], alpha=90000, geom="roundRect", text="暖", sz=2000),
        card4=dict(x=4700000, y=4500000, cx=2600000, cy=1200000, rot=0,
                   fill=C["pink"], alpha=85000, geom="roundRect", text="COLOR", sz=1700),
        dial=dict(x=900000, y=3900000, cx=1500000, cy=1500000, rot=0,
                  alpha=90000, text="100%"),
        bar=dict(x=900000, y=5700000, cx=10000000, cy=130000, alpha=95000),
        orb=dict(x=3600000, y=800000, cx=4400000, cy=4400000, alpha=30000),
        halo=dict(x=8200000, y=3400000, cx=3400000, cy=3400000, alpha=28000),
        ring=dict(x=4200000, y=1400000, cx=4200000, cy=4200000, alpha=20000),
        spark1=dict(x=2200000, y=1200000, alpha=90000, cx=300000, cy=300000),
        spark2=dict(x=9800000, y=1200000, alpha=85000, cx=260000, cy=260000),
        spark3=dict(x=11000000, y=4800000, alpha=85000, cx=300000, cy=300000),
    ), 18, "⑯ 整页换色，也是补间",
    "所有对象 fill 一起从冷蓝紫渐变到暖橙粉 · 双遮罩对切", "MORPH / PALETTE"))
for _k, _v in warm.items():
    FR[-1][_k]["fill"] = _v
# 让两条遮罩在第18页保持可见但退到画布边缘外，下一页再扫回来
FR[-1]["curtain"].update(x=-W - 200000, alpha=60000)
FR[-1]["curtain2"].update(x=W + 200000, alpha=60000)

# ---------- 页 19~22 【新】快闪卡点：4 页 0.3s 平滑 + 0.5s 停留 ----------
flash_pos = [
    # 四角炸开
    [(700000, 1500000, 2000000, 1300000, -18), (9500000, 1500000, 2000000, 1300000, 18),
     (700000, 4300000, 2000000, 1300000, 12), (9500000, 4300000, 2000000, 1300000, -12)],
    # 收成一条
    [(1100000, 3100000, 2200000, 1300000, 0), (3800000, 3100000, 2200000, 1300000, 0),
     (6500000, 3100000, 2200000, 1300000, 0), (9200000, 3100000, 2200000, 1300000, 0)],
    # 纵向堆叠
    [(4800000, 900000, 2600000, 1000000, 0), (4800000, 2100000, 2600000, 1000000, 0),
     (4800000, 3300000, 2600000, 1000000, 0), (4800000, 4500000, 2600000, 1000000, 0)],
    # 回到中心聚合
    [(5100000, 2600000, 2000000, 1400000, 0), (5100000, 2600000, 2000000, 1400000, 0),
     (5100000, 2600000, 2000000, 1400000, 0), (5100000, 2600000, 2000000, 1400000, 0)],
]
for j, pos in enumerate(flash_pos):
    i = 19 + j
    FR.append(numbered(
        mod(FR[-1],
            **{f"card{k+1}": dict(x=pos[k][0], y=pos[k][1], cx=pos[k][2], cy=pos[k][3],
                                  rot=pos[k][4], alpha=90000 if j < 3 else 70000,
                                  geom="roundRect", sz=1500,
                                  text=["01", "02", "03", "04"][k])
               for k in range(4)},
            curtain=dict(x=-W - 200000, alpha=0),
            curtain2=dict(x=W + 200000, alpha=0),
            spark1=dict(x=1500000 + j * 2200000, y=1000000, alpha=95000),
            spark2=dict(x=3500000 + j * 1800000, y=5200000, alpha=90000),
            spark3=dict(x=8000000 - j * 900000, y=1600000, alpha=90000),
            orb=dict(x=3000000 + j * 1400000, y=1200000, cx=3600000, cy=3600000, alpha=22000),
            halo=dict(x=7600000 - j * 900000, y=3200000, cx=3000000, cy=3000000, alpha=24000),
            ring=dict(x=4400000, y=1500000, cx=3300000, cy=3300000, alpha=14000),
            dial=dict(x=-3000000, alpha=0),
            bar=dict(x=900000, y=5700000, cx=10000000 - j * 2200000, cy=130000, alpha=85000),
            title=T(900000, 620000, f"⑰ 快闪 {j+1}/4 · 0.3 秒一帧", sz=3400),
            sub=S(920000, 1480000, "把平滑时长压到 300ms，连续四页就是卡点节奏", sz=1750),
        ), i, f"⑰ 快闪 {j+1}/4 · 0.3 秒一帧",
        "把平滑时长压到 300ms，连续四页就是卡点节奏", "MORPH / FLASH"))

# ---------- 页 23 【新】套娃回环：四层同心缩放 ----------
FR.append(numbered(
    mod(FR[-1],
        card1=dict(x=900000, y=900000, cx=10400000, cy=5400000, rot=0,
                   alpha=22000, fill=C["blue"], geom="roundRect", text="", sz=1000),
        card2=dict(x=2100000, y=1700000, cx=8000000, cy=4000000, rot=0,
                   alpha=28000, fill=C["purple"], geom="roundRect", text="", sz=1000),
        card3=dict(x=3300000, y=2400000, cx=5600000, cy=2800000, rot=0,
                   alpha=34000, fill=C["teal"], geom="roundRect", text="", sz=1000),
        card4=dict(x=4500000, y=3000000, cx=3200000, cy=1600000, rot=0,
                   alpha=42000, fill=C["indigo"], geom="roundRect", text="", sz=1000),
        ring=dict(x=3800000, y=1400000, cx=4600000, cy=4600000, alpha=18000, geom="ellipse"),
        orb=dict(x=3400000, y=1000000, cx=5400000, cy=5400000, alpha=20000),
        halo=dict(x=6600000, y=2600000, cx=3600000, cy=3600000, alpha=22000),
        spark1=dict(alpha=0), spark2=dict(alpha=0), spark3=dict(alpha=0),
        dial=dict(alpha=0), bar=dict(alpha=0),
        title=T(900000, 620000, "⑱ 套娃：一层套一层", sz=3400),
        sub=S(920000, 1480000, "四层同心放大 · 最后一页回到封面，整支片可无缝循环", sz=1750),
    ), 23, "⑱ 套娃：一层套一层",
    "四层同心放大 · 最后一页回到封面，整支片可无缝循环", "MORPH / LOOP"))

# ---------- 页 24 与封面完全同帧（video 可 loop；PPT 停在封面） ----------
FR.append(copy.deepcopy(FR[0]))

assert len(FR) == 24

# 每页 [morph 毫秒, 停留 毫秒]；最后页不自动换页
TIMINGS = [[900, 2100] for _ in range(24)]
for i in range(18, 22):          # 页 19~22 快闪
    TIMINGS[i] = [300, 500]
TIMINGS[22] = [800, 1800]
TIMINGS[23] = [900, None]

OPTIONS = ["byObject"] * 24
OPTIONS[10] = "byWord"           # 页 11 逐词
OPTIONS[11] = "byChar"           # 页 12 逐字
OPTIONS[19] = "byChar"           # 快闪段补一次逐字
OPTIONS[23] = "byObject"

FLASH = [18]                     # 快闪段开始处闪白

NOTES = [
    "ULTRA 版：24 页、20 个同名对象、九类新增效果，按 F5 自动连播约一分钟。",
    "原理：同一批 !! 对象在相邻两页的属性差，由 PowerPoint 自动补成中间帧。",
    "① 对象同名才配对，四张卡从画布外飞入。",
    "② 复制页面再改坐标，不要重新绘制。",
    "③ 单卡放大就是推镜。",
    "④ !! 强制异形配对，圆角矩形变六边形、三角、星形、菱形。",
    "⑤ 轮盘：公转与自转同时插值。",
    "⑥⑦ 连续两层无限缩放，镜头穿进画面。",
    "⑧ 多层视差 + 遮罩横扫揭示。",
    "⑨ byWord 按词补间。",
    "⑩ byChar，把一个字放大成整页背景，同时把对象收拢到中心待命。",
    "⑪ 爆炸：上一页全叠在中心，这页只改坐标与尺寸就炸开。",
    "⑫ 透视：四层梯形递缩加透明度衰减，做出 Z 轴纵深。",
    "⑬ 数字滚动开始：18%，文本、角度、宽度一起变。",
    "⑭ 刻度盘转到 45 度、数值 55%。",
    "⑮ 进度条长满 100%，三页连起来就是一段数据动画。",
    "⑯ 整页调色板从冷蓝紫渐变到暖橙粉，双遮罩对切。",
    "⑰ 快闪第一帧：平滑时长压到 300 毫秒。",
    "⑰ 快闪第二帧：收成一条。",
    "⑰ 快闪第三帧：纵向堆叠。",
    "⑰ 快闪第四帧：回到中心聚合。",
    "⑱ 套娃：四层同心缩放。",
    "结尾与封面同帧，整支视频可以无缝循环。",
]


# ---------------- OOXML 生成 ----------------
def shape_xml(name, s, sid):
    rot = int(round(float(s.get("rot", 0)) * 60000))
    x, y, cx, cy = (int(s[k]) for k in ("x", "y", "cx", "cy"))
    geom = s.get("geom", "rect")
    fill, alpha = s.get("fill"), int(s.get("alpha", 100000))
    algn = s.get("align", "l")
    if algn not in ("l", "ctr", "r"):
        algn = "l"
    fill_xml = "<a:noFill/>"
    if fill:
        fill_xml = (f'<a:solidFill><a:srgbClr val="{fill}">'
                    f'<a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>')
    av = ('<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>'
          if geom == "donut" else "<a:avLst/>")
    tx = ""
    if s.get("text"):
        tx = (f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" '
              f'anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
              + bp.run_xml(s["text"], s.get("sz", 1800), s.get("color", "FFFFFF"),
                           s.get("bold", False), algn) + "</p:txBody>")
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="!!{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm rot="{rot}"><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}">{av}</a:prstGeom>'
        f'{fill_xml}<a:ln><a:noFill/></a:ln></p:spPr>{tx}</p:sp>'
    )


def transition_xml(idx, option, dur, adv):
    adv_attr = ' advClick="1"' if adv is None else f' advClick="1" advTm="{adv}"'
    choice = (f'<p:transition spd="slow" p14:dur="{dur}"{adv_attr}>'
              f'<p159:morph option="{option}"/></p:transition>')
    fallback = f'<p:transition spd="slow"{adv_attr}><p:fade/></p:transition>'
    return ('<mc:AlternateContent xmlns:mc="' + MC + '">'
            '<mc:Choice xmlns:p159="' + P159 + '" Requires="p159">' + choice +
            '</mc:Choice><mc:Fallback>' + fallback + '</mc:Fallback>'
            '</mc:AlternateContent>')


def slide_xml(frame, idx):
    sps = "".join(shape_xml(n, frame[n], IDS[n]) for n in ORDER)
    dur, adv = TIMINGS[idx]
    root = (
        f'<p:sld xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}" '
        f'xmlns:mc="{MC}" xmlns:p14="{P14}"><p:cSld><p:spTree>'
        f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{sps}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        f'{transition_xml(idx, OPTIONS[idx], dur, adv)}</p:sld>'
    )
    return ('<?xml version="1.0" encoding="UTF-8"?>' + root).encode("utf-8")


def main():
    zin = zipfile.ZipFile(BASE)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    orig = dict((i.filename, d) for i, d in items)
    patched = {}

    for idx, n in enumerate(SLIDE_FILES):
        x = slide_xml(FR[idx], idx)
        etree.fromstring(x)
        patched[f"ppt/slides/slide{n}.xml"] = x

    # 新增页的 rels：只保留 slideLayout 关系
    lay = re.search(r'<Relationship[^>]*Type="[^"]*slideLayout"[^>]*/>',
                    orig["ppt/slides/_rels/slide1.xml.rels"].decode("utf-8"))
    layout_rel = lay.group(0) if lay else ""
    for n in NEW_FILES:
        patched[f"ppt/slides/_rels/slide{n}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<Relationships xmlns="{CT_REL}">{layout_rel}</Relationships>'
        ).encode("utf-8")

    # 只追加新增页，保留原有页序节点的位置与既有 rId / sldId。
    # CT_Presentation 的子元素顺序是严格固定的：sldIdLst 必须排在 sldSz 之前，
    # 放到末尾会被 PowerPoint / WPS / 在线导入的 schema 校验判为损坏。
    prels = orig["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    rid0 = max(int(m) for m in re.findall(r'Id="rId(\d+)"', prels)) + 1
    rels_add = "".join(
        f'<Relationship Id="rId{rid0 + j}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        f'relationships/slide" Target="slides/slide{n}.xml"/>'
        for j, n in enumerate(NEW_FILES))
    patched["ppt/_rels/presentation.xml.rels"] = prels.replace(
        "</Relationships>", rels_add + "</Relationships>").encode("utf-8")

    ct = orig["[Content_Types].xml"].decode("utf-8")
    ov = "".join(
        f'<Override PartName="/ppt/slides/slide{n}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'presentationml.slide+xml"/>' for n in NEW_FILES)
    patched["[Content_Types].xml"] = ct.replace("</Types>", ov + "</Types>").encode("utf-8")

    pres = orig["ppt/presentation.xml"].decode("utf-8")
    assert "<p:sldIdLst>" in pres and "</p:sldIdLst>" in pres
    sid0 = max(int(m) for m in re.findall(r'<p:sldId id="(\d+)"', pres)) + 1
    sids = "".join(f'<p:sldId id="{sid0 + j}" r:id="rId{rid0 + j}"/>'
                   for j in range(len(NEW_FILES)))
    patched["ppt/presentation.xml"] = pres.replace(
        "</p:sldIdLst>", sids + "</p:sldIdLst>").encode("utf-8")

    # 前 9 页旧备注换成新讲稿（新增页无备注，不阻断）
    for idx, n in enumerate(OLD_FILES):
        rel_name = f"ppt/slides/_rels/slide{n}.xml.rels"
        m = re.search(r'Target="\.\./notesSlides/(notesSlide\d+\.xml)"',
                      orig.get(rel_name, b"").decode("utf-8"))
        if not m:
            continue
        key = "ppt/notesSlides/" + m.group(1)
        if key not in orig:
            continue
        s = orig[key].decode("utf-8")
        ts = re.findall(r"<a:t>(.*?)</a:t>", s, flags=re.S)
        if ts:
            s = s.replace(f"<a:t>{ts[0]}</a:t>", f"<a:t>{esc(NOTES[idx])}</a:t>", 1)
            for extra_t in ts[1:]:
                s = s.replace(f"<a:t>{extra_t}</a:t>", "<a:t></a:t>", 1)
        patched[key] = s.encode("utf-8")

    zout = zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED)
    for info, data in items:
        zout.writestr(info, patched.get(info.filename, data))
    for name, data in patched.items():
        if name not in orig:
            zout.writestr(name, data)
    zout.close()
    zin.close()

    # 自检
    chk = zipfile.ZipFile(DST)
    assert chk.testzip() is None
    for name in patched:
        if name.endswith(".xml") or name.endswith(".rels"):
            etree.fromstring(chk.read(name))
    for idx, n in enumerate(SLIDE_FILES):
        d = chk.read(f"ppt/slides/slide{n}.xml").decode("utf-8")
        r = etree.fromstring(d.encode("utf-8"))
        names = [e.get("name") for e in r.findall(f".//{{{P}}}cNvPr")][1:]
        assert names == ["!!" + k for k in ORDER], (n, names)
        assert f'<p159:morph option="{OPTIONS[idx]}"/>' in d
        assert f'p14:dur="{TIMINGS[idx][0]}"' in d
    pres_out = chk.read("ppt/presentation.xml").decode("utf-8")
    rels_out = chk.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    ct_out = chk.read("[Content_Types].xml").decode("utf-8")
    assert pres_out.count("<p:sldId ") == 24
    # schema 顺序：sldIdLst 必须在 sldSz 之前
    assert pres_out.find("<p:sldIdLst>") < pres_out.find("<p:sldSz"), "sldIdLst 位置错误"
    assert pres_out.find("</p:sldIdLst>") < pres_out.find("<p:notesSz")
    # 每个页序 id 都能解析到存在的 slide 部件，且被声明 content type
    relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="(slides/slide\d+\.xml)"', rels_out))
    order_ids = re.findall(r'<p:sldId id="\d+" r:id="(rId\d+)"/>', pres_out)
    assert len(order_ids) == 24, order_ids
    for rid in order_ids:
        target = relmap.get(rid)
        assert target, f"页序引用了不存在的关系 {rid}"
        part = "ppt/" + target
        assert part in chk.namelist(), f"缺少部件 {part}"
        assert f'PartName="/{part}"' in ct_out, f"未声明 content type: {part}"
    # rId 与 sldId 全表唯一
    assert len(set(order_ids)) == 24
    sid_list = [int(m) for m in re.findall(r'<p:sldId id="(\d+)"', pres_out)]
    assert len(set(sid_list)) == 24 and min(sid_list) >= 256
    chk.close()
    print("  schema 自检通过：sldIdLst 位置正确 / 24 页关系与部件一一对应 / id 唯一")

    print("OK ->", DST)
    print("  24 页 / 20 个 !! 同名对象 / 跨页 shape id 一致")
    print("  快闪 4 页 300ms · 常规 900ms · byWord×1 byChar×2")
    print("  首尾同帧可循环 · presProps useTimings=1")


if __name__ == "__main__":
    main()
