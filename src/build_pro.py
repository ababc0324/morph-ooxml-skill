# -*- coding: utf-8 -*-
"""
Morph PRO：在 9 页基础上扩到 12 页，加入 6 种进阶平滑效果。

新增手法：
  A 异形变换（!! 强制圆→六边形→三角→星形→菱形配对）
  B 轮盘旋转（多卡绕中心公转 + 自转）
  C 无限缩放定位（连续两页穿入放大，制造 infinite zoom）
  D 多层视差 + 遮罩横扫揭示
  E byWord 逐词补间
  F byChar 逐字补间
"""
import copy
import html
import os
import re
import zipfile
from lxml import etree

# 需要自备一个已含 9 页基础内容的 BASE pptx（本仓库不含示例交付物）。
BASE = "./平滑切换演示_Morph版.pptx"
DST = "./平滑切换演示_MorphPRO版.pptx"

W, H = 12192000, 6858000

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
CT_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

C = {
    "bg": "050B1A", "blue": "25A7FF", "purple": "7C3CFF",
    "teal": "5EEAD4", "indigo": "3B82F6", "white": "FFFFFF",
    "muted": "9FB6D9", "dim": "6E86A8", "pink": "F472B6",
    "amber": "FBBF24",
}


def st(x, y, cx, cy, rot=0, fill=None, alpha=100000, geom="rect",
       text="", sz=1800, color="FFFFFF", align="l", bold=False):
    return dict(x=x, y=y, cx=cx, cy=cy, rot=rot, fill=fill, alpha=alpha,
                geom=geom, text=text, sz=sz, color=color, align=align, bold=bold)


# 每页固定 14 个演员；跨页 shape id 与 !! 名称完全一致
ORDER = ["bg", "grid", "orb", "halo", "curtain", "ring",
         "card1", "card2", "card3", "card4",
         "title", "sub", "num", "tag"]
IDS = {n: i + 2 for i, n in enumerate(ORDER)}

# 每页切换的 Morph 匹配粒度
OPTIONS = ["byObject"] * 10 + ["byWord", "byChar"]


def T(x, y, text, sz=3800, color="FFFFFF", align="l",
      cx=10000000, cy=800000, bold=True):
    return dict(x=x, y=y, cx=cx, cy=cy, text=text, sz=sz, color=color,
                align=align, bold=bold)


def S(x, y, text, sz=1800, color="9FB6D9", cx=10000000, cy=500000, bold=False):
    return T(x, y, text, sz=sz, color=color, cx=cx, cy=cy, bold=bold)


def mod(prev, **kw):
    f = copy.deepcopy(prev)
    for k, v in kw.items():
        f[k].update(v)
    return f


def caption(i, title, sub, label="FLOW MOTION"):
    return dict(
        title=T(900000, 620000, title, sz=3500, bold=True),
        sub=S(920000, 1480000, sub, sz=1800),
        num=dict(text=f"{i:02d} / 12", x=920000, y=5900000, cx=4000000),
        tag=dict(text=label, x=10100000, y=380000, cx=1700000, cy=350000,
                 color=C["dim"], sz=1200, align="r"),
    )


def with_caption(prev, i, title, sub, label="FLOW MOTION"):
    return mod(prev, **caption(i, title, sub, label))


# ---------- 页 1 封面 ----------
f1 = {
    "bg":      st(0, 0, W, H, fill=C["bg"]),
    "grid":    st(0, 0, W, H, fill=None, text="", alpha=0),
    "orb":     st(7200000, 700000, 4600000, 4600000, fill=C["blue"], alpha=42000, geom="ellipse"),
    "halo":    st(9000000, 3700000, 2800000, 2800000, fill=C["purple"], alpha=32000, geom="ellipse"),
    "curtain": st(-W, 0, W, H, fill=C["indigo"], alpha=0),
    "ring":    st(5100000, 1900000, 2000000, 2000000, fill=None, alpha=0, geom="ellipse"),
    "card1":   st(-3200000, 2400000, 2400000, 1700000, rot=-12, fill=C["blue"], alpha=85000,
                  geom="roundRect", text="转场", sz=1700, align="ctr", bold=True),
    "card2":   st(W + 800000, 1200000, 2400000, 1700000, rot=12, fill=C["purple"], alpha=85000,
                  geom="roundRect", text="入场", sz=1700, align="ctr", bold=True),
    "card3":   st(-3200000, 5200000, 2400000, 1700000, rot=10, fill=C["teal"], alpha=80000,
                  geom="roundRect", text="节奏", sz=1700, align="ctr", bold=True),
    "card4":   st(W + 800000, 5200000, 2400000, 1700000, rot=-10, fill=C["indigo"], alpha=80000,
                  geom="roundRect", text="强调", sz=1700, align="ctr", bold=True),
    "title":   st(900000, 2250000, 9500000, 1000000, text="平滑切换 PRO", sz=6000,
                  color=C["white"], bold=True),
    "sub":     st(920000, 3400000, 9500000, 600000,
                  text="MORPH · 进阶效果合集 / 12 页自动连播", sz=2000, color=C["muted"]),
    "num":     st(920000, 4400000, 4000000, 400000, text="01 / 12", sz=1500, color=C["blue"]),
    "tag":     st(10100000, 400000, 1700000, 350000, text="FLOW MOTION", sz=1200,
                  color=C["dim"], align="r"),
}
FR = [f1]

# ---------- 页 2 原理：镜头拉开 ----------
f2 = mod(f1,
    orb=dict(x=9800000, y=-900000, cx=3000000, cy=3000000, alpha=38000),
    halo=dict(x=-700000, y=4700000, cx=2400000, cy=2400000, alpha=30000),
    card1=dict(x=-2600000), card2=dict(x=W + 300000),
    card3=dict(x=-2600000), card4=dict(x=W + 300000),
    title=T(900000, 1900000, "补间的维度，不止位置", sz=4600, bold=True, cy=900000),
    sub=S(920000, 3000000, "复制同一页 → 改属性 → 切换选「平滑」", sz=1900),
    num=dict(text="02 / 12", y=3900000),
)
FR.append(f2)

# ---------- 页 3 ① 飞入成行 ----------
row = [(846000 + i * 2700000, 2500000, 2400000, 1700000) for i in range(4)]
f3 = mod(f2,
    card1=dict(x=row[0][0], y=row[0][1], cx=row[0][2], cy=row[0][3], rot=0),
    card2=dict(x=row[1][0], y=row[1][1], cx=row[1][2], cy=row[1][3], rot=0),
    card3=dict(x=row[2][0], y=row[2][1], cx=row[2][2], cy=row[2][3], rot=0),
    card4=dict(x=row[3][0], y=row[3][1], cx=row[3][2], cy=row[3][3], rot=0),
    orb=dict(x=4700000, y=-2400000, cx=3200000, cy=3200000, alpha=22000),
    halo=dict(x=9200000, y=4800000, cx=2200000, cy=2200000, alpha=24000),
)
FR.append(with_caption(f3, 3, "① 对象同名，才会补间", "!! 前缀强制配对 · 跨页保持同一批形状"))

# ---------- 页 4 ② 田字格 + 旋转 ----------
grid = [(2200000, 2100000, 3200000, 1500000, -6),
        (6300000, 2100000, 3200000, 1500000, 5),
        (2200000, 4000000, 3200000, 1500000, 5),
        (6300000, 4000000, 3200000, 1500000, -5)]
f4 = mod(FR[-1],
    card1=dict(x=grid[0][0], y=grid[0][1], cx=grid[0][2], cy=grid[0][3], rot=grid[0][4]),
    card2=dict(x=grid[1][0], y=grid[1][1], cx=grid[1][2], cy=grid[1][3], rot=grid[1][4]),
    card3=dict(x=grid[2][0], y=grid[2][1], cx=grid[2][2], cy=grid[2][3], rot=grid[2][4]),
    card4=dict(x=grid[3][0], y=grid[3][1], cx=grid[3][2], cy=grid[3][3], rot=grid[3][4]),
    orb=dict(x=-1200000, y=4200000, cx=2600000, cy=2600000, alpha=26000),
    halo=dict(x=10000000, y=-800000, cx=2400000, cy=2400000, alpha=26000),
)
FR.append(with_caption(f4, 4, "② 复制页，不要重新画", "只改坐标与角度，中间帧由 PowerPoint 自己算"))

# ---------- 页 5 ③ 推镜 ----------
f5 = mod(FR[-1],
    card1=dict(x=500000, y=5400000, cx=1400000, cy=900000, rot=0, alpha=45000),
    card2=dict(x=2600000, y=2100000, cx=7000000, cy=3600000, rot=0, alpha=92000, sz=2600),
    card3=dict(x=9800000, y=5400000, cx=1400000, cy=900000, rot=0, alpha=45000),
    card4=dict(x=10400000, y=2100000, cx=1400000, cy=900000, rot=0, alpha=45000),
    orb=dict(x=4300000, y=1400000, cx=4200000, cy=4200000, alpha=26000),
    halo=dict(x=7600000, y=3600000, cx=3000000, cy=3000000, alpha=22000),
)
FR.append(with_caption(mod(f5, num=dict(text="05 / 12", x=2100000, y=5950000, cx=2500000)),
                       5, "③ 缩放定位，就是一次推镜", "把一张卡放大，镜头自然推近"))

# ---------- 页 6 【新 A】异形变换：!! 强制不同几何配对 ----------
shapes = ["hexagon", "triangle", "star5", "diamond"]
hexpos = [(1500000, 2200000, 3000000, 1900000, -8),
          (6000000, 1900000, 2700000, 2000000, 8),
          (2600000, 4200000, 2700000, 1700000, 6),
          (7000000, 4300000, 3000000, 1800000, -6)]
f6 = mod(FR[-1],
    card1=dict(x=hexpos[0][0], y=hexpos[0][1], cx=hexpos[0][2], cy=hexpos[0][3],
               rot=hexpos[0][4], geom=shapes[0], alpha=90000, fill=C["blue"]),
    card2=dict(x=hexpos[1][0], y=hexpos[1][1], cx=hexpos[1][2], cy=hexpos[1][3],
               rot=hexpos[1][4], geom=shapes[1], alpha=90000, fill=C["purple"], sz=1700),
    card3=dict(x=hexpos[2][0], y=hexpos[2][1], cx=hexpos[2][2], cy=hexpos[2][3],
               rot=hexpos[2][4], geom=shapes[2], alpha=88000, fill=C["teal"]),
    card4=dict(x=hexpos[3][0], y=hexpos[3][1], cx=hexpos[3][2], cy=hexpos[3][3],
               rot=hexpos[3][4], geom=shapes[3], alpha=88000, fill=C["indigo"]),
    ring=dict(x=4300000, y=1600000, cx=3600000, cy=3600000, alpha=18000,
              fill=C["teal"], geom="ellipse"),
    orb=dict(x=-1400000, y=3900000, cx=3000000, cy=3000000, alpha=22000),
    halo=dict(x=9800000, y=-900000, cx=2600000, cy=2600000, alpha=24000),
    num=dict(text="06 / 12", x=920000, y=5900000, cx=4000000),
)
FR.append(with_caption(f6, 6, "④ 异形变换：圆也能平滑成星",
                       "!! 同名强制配对 · roundRect → hexagon / triangle / star5 / diamond",
                       "MORPH / SHAPE"))

# ---------- 页 7 【新 B】轮盘旋转：公转 + 自转 ----------
import math
cx0, cy0 = W / 2, H / 2 + 200000
R0 = 3100000
wheel = []
for k in range(4):
    ang = math.radians(-90 + k * 90 + 72)
    wheel.append((int(cx0 + R0 * math.cos(ang) - 1100000),
                  int(cy0 + R0 * math.sin(ang) - 800000), 2200000, 1600000,
                  -90 + k * 90 + 72))
f7 = mod(FR[-1],
    card1=dict(x=wheel[0][0], y=wheel[0][1], cx=wheel[0][2], cy=wheel[0][3],
               rot=wheel[0][4], geom="roundRect", alpha=90000, sz=1400),
    card2=dict(x=wheel[1][0], y=wheel[1][1], cx=wheel[1][2], cy=wheel[1][3],
               rot=wheel[1][4], geom="roundRect", alpha=90000, sz=1400),
    card3=dict(x=wheel[2][0], y=wheel[2][1], cx=wheel[2][2], cy=wheel[2][3],
               rot=wheel[2][4], geom="roundRect", alpha=88000, sz=1400),
    card4=dict(x=wheel[3][0], y=wheel[3][1], cx=wheel[3][2], cy=wheel[3][3],
               rot=wheel[3][4], geom="roundRect", alpha=88000, sz=1400),
    ring=dict(x=int(cx0 - 1500000), y=int(cy0 - 1500000), cx=3000000, cy=3000000,
              alpha=26000, fill=C["blue"], geom="ellipse"),
    orb=dict(x=int(cx0 - 2100000), y=int(cy0 - 2100000), cx=4200000, cy=4200000, alpha=20000),
    halo=dict(x=9000000, y=4200000, cx=2600000, cy=2600000, alpha=26000),
)
FR.append(with_caption(f7, 7, "⑤ 轮盘：公转与自转同时插值",
                       "四张卡绕中心旋转 72° · 各自角度也一起变", "MORPH / ROTATE"))

# ---------- 页 8 【新 C】无限缩放 1：穿入 ----------
f8 = mod(FR[-1],
    card1=dict(x=-2600000, y=2500000, cx=1800000, cy=1200000, rot=-20, alpha=35000),
    card2=dict(x=1300000, y=900000, cx=9600000, cy=5600000, rot=0, alpha=96000,
               sz=3400, geom="roundRect", fill=C["purple"], text="入场"),
    card3=dict(x=W + 600000, y=4200000, cx=1800000, cy=1200000, rot=18, alpha=35000),
    card4=dict(x=W + 600000, y=900000, cx=1800000, cy=1200000, rot=-18, alpha=35000),
    ring=dict(x=2500000, y=1500000, cx=7200000, cy=4400000, alpha=12000,
              fill=C["teal"], geom="ellipse"),
    orb=dict(x=3600000, y=1400000, cx=5000000, cy=5000000, alpha=18000),
    halo=dict(x=7000000, y=3000000, cx=3600000, cy=3600000, alpha=18000),
)
FR.append(with_caption(mod(f8, num=dict(text="08 / 12", x=920000, y=6000000, cx=4000000)),
                       8, "⑥ 连续放大，就是无限缩放",
                       "同一对象逐页放大 2 倍 · 穿进画面，不用任何视频插件",
                       "MORPH / ZOOM"))

# ---------- 页 9 【新 C】无限缩放 2：再穿一层 ----------
f9 = mod(FR[-1],
    curtain=dict(x=-W, y=0, cx=W, cy=H, fill=C["indigo"], alpha=62000),
    card1=dict(x=-3600000, y=2800000, alpha=0),
    card2=dict(x=-2200000, y=-1300000, cx=16000000, cy=9600000, rot=0,
               alpha=97000, sz=4600, text=""),
    card3=dict(x=W + 1600000, y=4500000, alpha=0),
    card4=dict(x=W + 1600000, y=800000, alpha=0),
    ring=dict(x=1600000, y=600000, cx=9000000, cy=5600000, alpha=16000,
              fill=C["blue"], geom="ellipse"),
    orb=dict(x=2400000, y=700000, cx=7400000, cy=5600000, alpha=22000),
    halo=dict(x=6400000, y=2400000, cx=4600000, cy=4200000, alpha=20000),
)
FR.append(with_caption(mod(f9, num=dict(text="09 / 12", x=920000, y=6000000, cx=4000000)),
                       9, "⑦ 再放大一层，镜头继续往里",
                       "关键是两页用同一个 !! 对象，只改尺寸与坐标", "MORPH / ZOOM"))

# ---------- 页 10 【新 D】三层视差 + 遮罩横扫 ----------
f10 = mod(FR[-1],
    card1=dict(x=900000, y=2400000, cx=2600000, cy=1700000, rot=0, alpha=88000,
               geom="roundRect", fill=C["blue"], text="转场", sz=1700),
    card2=dict(x=4000000, y=2200000, cx=4200000, cy=2200000, rot=0, alpha=90000,
               geom="roundRect", fill=C["purple"], text="入场", sz=2100),
    card3=dict(x=8600000, y=2400000, cx=2600000, cy=1700000, rot=0, alpha=88000,
               geom="roundRect", fill=C["teal"], text="节奏", sz=1700),
    card4=dict(x=4700000, y=4700000, cx=2800000, cy=1200000, rot=0, alpha=80000,
               geom="roundRect", fill=C["indigo"], text="强调", sz=1700),
    curtain=dict(x=-W, y=0, cx=W, cy=H, fill=C["blue"], alpha=0),
    ring=dict(x=3600000, y=1000000, cx=5000000, cy=5000000, alpha=18000,
              fill=C["teal"], geom="ellipse"),
    orb=dict(x=-900000, y=500000, cx=3600000, cy=3600000, alpha=24000),
    halo=dict(x=9500000, y=3800000, cx=3200000, cy=3200000, alpha=26000),
)
FR.append(with_caption(f10, 10, "⑧ 多层视差 + 遮罩揭示",
                       "前景 / 中景 / 背景位移幅度不同，速度差就是景深", "MORPH / PARALLAX"))
# 遮罩：页9 在左侧画布外待命，页10 落到右侧画布外；补间时整条色带横扫画面
FR[-1]["curtain"].update(dict(x=W + 100000, alpha=62000))

# ---------- 页 11 【新 E】byWord 逐词补间 ----------
f11 = mod(FR[-1],
    curtain=dict(x=W + 200000, alpha=0),
    card1=dict(x=1500000, y=4200000, cx=2000000, cy=1300000, alpha=70000,
               text="逐", sz=2600, fill=C["blue"]),
    card2=dict(x=4300000, y=3000000, cx=2600000, cy=1800000, alpha=88000,
               text="词", sz=3200, fill=C["purple"]),
    card3=dict(x=7500000, y=4200000, cx=2000000, cy=1300000, alpha=70000,
               text="补", sz=2600, fill=C["teal"]),
    card4=dict(x=10100000, y=2800000, cx=1600000, cy=1300000, alpha=70000,
               text="间", sz=2400, fill=C["indigo"]),
    orb=dict(x=4200000, y=800000, cx=3800000, cy=3800000, alpha=24000),
    halo=dict(x=8200000, y=3600000, cx=3000000, cy=3000000, alpha=22000),
    ring=dict(x=3300000, y=1400000, cx=5600000, cy=4400000, alpha=14000,
              fill=C["blue"], geom="ellipse"),
)
FR.append(with_caption(f11, 11, "⑨ 切到「按词」，文字自己走位",
                       "p159:morph option=\"byWord\" · 相同词在两页间独立补间",
                       "MORPH / WORD"))

# ---------- 页 12 【新 F】byChar 逐字补间 + 收尾 ----------
f12 = mod(FR[-1],
    card1=dict(x=-3000000, y=4400000, alpha=0),
    card2=dict(x=W + 800000, y=2800000, alpha=0),
    card3=dict(x=-3000000, y=5200000, alpha=0),
    card4=dict(x=W + 800000, y=4800000, alpha=0),
    curtain=dict(x=-W, alpha=0),
    orb=dict(x=3800000, y=1100000, cx=4600000, cy=4600000, alpha=40000, fill=C["blue"]),
    halo=dict(x=6400000, y=2700000, cx=3600000, cy=3600000, alpha=28000, fill=C["purple"]),
    ring=dict(x=4700000, y=1800000, cx=2800000, cy=2800000, alpha=18000,
              fill=C["teal"], geom="ellipse"),
    title=T(900000, 2400000, "顺滑，是被感觉到的", sz=5200, bold=True, cy=1000000),
    sub=S(920000, 3600000, "MORPH PRO / 12 页连播 · Thank you", sz=2000),
    num=dict(text="12 / 12", y=4500000),
    tag=dict(text="END"),
)
FR.append(f12)

assert len(FR) == 12

NOTES = [
    "开场：PRO 版把基础手法和六种进阶效果串成一支 12 页短片，全程自动连播，按 F5 即可。",
    "原理回顾：平滑切换对比相邻两页的同一批对象，把位置、大小、旋转、颜色、透明度的差值补成中间帧。",
    "① 配对：跨页对象用 !! 前缀同名，PowerPoint 才会认成同一个实体。四张卡从画布外飞入。",
    "② 做法：永远复制页面再改属性，别重新画。一行变田字格只改了四个坐标和角度。",
    "③ 推镜：把一张卡放大，镜头自然推近，其余缩小退到背景。",
    "④ 异形变换：!! 同名可以强制不同几何互相平滑，圆角矩形变成六边形、三角、星形和菱形。",
    "⑤ 轮盘：四张卡绕中心公转 72 度，同时各自角度变化，位移与旋转一起插值。",
    "⑥ 无限缩放第一层：同一张卡放大到接近满屏，制造穿入画面的错觉。",
    "⑦ 无限缩放第二层：继续放大到溢出画布，中间帧由 PowerPoint 自动算，不需要任何视频插件。",
    "⑧ 视差：前景卡片、中景光球、背景光晕位移幅度不同，速度差就产生了景深；遮罩横扫做揭示。",
    "⑨ 按词：切换效果选项改成「按词」，morph 的 option 写 byWord，相同词会在两页之间独立补间。",
    "结尾：按字补间 byChar 收束。12 页每页 1 秒平滑、停留 3 秒，文件里每页都写了原生 morph。",
]


# ---------------- OOXML ----------------
def esc(s):
    return html.escape(str(s), quote=False)


def run_xml(text, sz, color, bold, align):
    out = []
    for p in str(text).split("\n"):
        b = ' b="1"' if bold else ""
        out.append(
            f'<a:p><a:pPr algn="{align}"/><a:r>'
            f'<a:rPr lang="zh-CN" sz="{int(sz)}"{b} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/>'
            f'</a:rPr><a:t>{esc(p)}</a:t></a:r></a:p>'
        )
    return "".join(out)


def shape_xml(name, s, sid):
    rot = int(round(float(s.get("rot", 0)) * 60000))
    x, y, cx, cy = (int(s[k]) for k in ("x", "y", "cx", "cy"))
    geom = s.get("geom", "rect")
    fill = s.get("fill")
    alpha = int(s.get("alpha", 100000))
    algn = s.get("align", "l")
    if algn not in ("l", "ctr", "r"):
        algn = "l"
    fill_xml = "<a:noFill/>"
    if fill:
        fill_xml = (f'<a:solidFill><a:srgbClr val="{fill}">'
                    f'<a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>')
    tx = ""
    if s.get("text"):
        tx = (f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" '
              f'anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
              + run_xml(s["text"], s.get("sz", 1800), s.get("color", "FFFFFF"),
                        s.get("bold", False), algn) + "</p:txBody>")
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="!!{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm rot="{rot}"><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
        f'{fill_xml}<a:ln><a:noFill/></a:ln></p:spPr>{tx}</p:sp>'
    )


def transition_xml(idx, option, dur=1000, adv=3000):
    last = idx == len(FR) - 1
    adv_attr = ' advClick="1"' if last else f' advClick="1" advTm="{adv}"'
    choice = (f'<p:transition spd="slow" p14:dur="{dur}"{adv_attr}>'
              f'<p159:morph option="{option}"/></p:transition>')
    fallback = f'<p:transition spd="slow"{adv_attr}><p:fade/></p:transition>'
    return ('<mc:AlternateContent xmlns:mc="' + MC + '">'
            '<mc:Choice xmlns:p159="' + P159 + '" Requires="p159">' + choice +
            '</mc:Choice><mc:Fallback>' + fallback + '</mc:Fallback>'
            '</mc:AlternateContent>')


def slide_xml(frame, idx):
    sps = "".join(shape_xml(n, frame[n], IDS[n]) for n in ORDER)
    root = (
        f'<p:sld xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}" '
        f'xmlns:mc="{MC}" xmlns:p14="{P14}"><p:cSld><p:spTree>'
        f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{sps}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        f'{transition_xml(idx, OPTIONS[idx])}</p:sld>'
    )
    return ('<?xml version="1.0" encoding="UTF-8"?>' + root).encode("utf-8")


def ensure_base():
    """无外部 BASE 时，用 build_base 从零生成基础稿（agent 即开即用）。"""
    if not os.path.exists(BASE):
        from build_base import build_base
        build_base()
        print(f"[build_pro] BASE 不存在，已自动生成基础稿 -> {BASE}")


def main():
    ensure_base()
    zin = zipfile.ZipFile(BASE)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    orig = dict((i.filename, d) for i, d in items)
    patched = {}

    # 原有 9 页，按物理文件顺序写入 1~9 帧
    old = [1, 7, 4, 3, 8, 5, 11, 10, 12]
    for idx, n in enumerate(old):
        name = f"ppt/slides/slide{n}.xml"
        x = slide_xml(FR[idx], idx)
        etree.fromstring(x)
        patched[name] = x

    # 新增 3 页：slide20/21/22，补齐 content types / presentation rels / sldIdLst
    new_names = ["ppt/slides/slide20.xml", "ppt/slides/slide21.xml",
                 "ppt/slides/slide22.xml"]
    rels_src = orig["ppt/slides/_rels/slide1.xml.rels"].decode("utf-8")
    lay = re.search(r'<Relationship[^>]*Type="[^"]*slideLayout"[^>]*/>', rels_src)
    layout_rel = lay.group(0) if lay else ""

    prels = orig["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', prels)]
    nxt = max(used) + 1
    ct = orig["[Content_Types].xml"].decode("utf-8")
    pres = orig["ppt/presentation.xml"].decode("utf-8")
    sid_used = [int(m) for m in re.findall(r'<p:sldId id="(\d+)"', pres)]
    sid_next = max(sid_used) + 1

    for j, name in enumerate(new_names):
        idx = 9 + j
        x = slide_xml(FR[idx], idx)
        etree.fromstring(x)
        patched[name] = x
        relname = name.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        patched[relname] = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<Relationships xmlns="{CT_REL}">{layout_rel}</Relationships>'
        ).encode("utf-8")
        rid = f"rId{nxt + j}"
        prels = prels.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            f'relationships/slide" Target="slides/slide{20 + j}.xml"/></Relationships>'
        )
        ct = ct.replace(
            "</Types>",
            f'<Override PartName="/ppt/slides/slide{20 + j}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'presentationml.slide+xml"/></Types>'
        )
        pres = pres.replace("</p:sldIdLst>",
                            f'<p:sldId id="{sid_next + j}" r:id="{rid}"/></p:sldIdLst>')
    patched["ppt/_rels/presentation.xml.rels"] = prels.encode("utf-8")
    patched["[Content_Types].xml"] = ct.encode("utf-8")
    patched["ppt/presentation.xml"] = pres.encode("utf-8")

    # 口播备注：原 9 页
    for idx, n in enumerate(old[:9]):
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
            for extra in ts[1:]:
                s = s.replace(f"<a:t>{extra}</a:t>", "<a:t></a:t>", 1)
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
    files = [f"ppt/slides/slide{n}.xml" for n in old] + new_names
    for idx, name in enumerate(files):
        d = chk.read(name).decode("utf-8")
        r = etree.fromstring(d.encode("utf-8"))
        names = [e.get("name") for e in r.findall(f".//{{{P}}}cNvPr")][1:]
        assert names == ["!!" + k for k in ORDER], (name, names)
        assert f'<p159:morph option="{OPTIONS[idx]}"/>' in d
    assert chk.read("ppt/presentation.xml").decode().count("<p:sldId ") == 12
    chk.close()

    print("OK ->", DST)
    print("  12 页 / 每页 14 个 !! 同名对象 / shape id 跨页一致")
    print("  byObject ×10 + byWord ×1 + byChar ×1 · 1.0s 平滑 · 3.0s 停留")


if __name__ == "__main__":
    main()
