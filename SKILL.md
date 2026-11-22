---
name: morph-ooxml
description: >
  Generate native PowerPoint Morph (smooth) transition decks by hand-writing OOXML.
  Use when the user wants a PPTX with PowerPoint Morph/smooth transitions, or to
  build/verify the self-contained morph generator. Pure text asset, no binaries.
---

# Morph（平滑切换）PPT 构建 · OOXML 规则与生成器

本技能让 agent **零外部输入**即可产出带原生 Morph 平滑切换的 PPTX。
核心是「兼容性规则层」与「可复用页面生成器骨架」，纯文本资产，MIT。

## 前置条件

- Python 3.10+
- 依赖：`pip install -r requirements.txt`（`lxml>=4.9`）

## 三步使用

### 1) 验证链路（推荐先跑，无需 PowerPoint）

```bash
python src/selftest.py
```

会依次：生成 9 页基础稿 → 12 页 PRO → 24 页 ULTRA，并跑全部结构自检
（zip 完整性、XML 可解析、`!!` 同名对象、`p159:morph` 声明、`sldIdLst` 页数、
部件与页序一一对应）。**退出码 0 即通过。**

### 2) 生成成品

```bash
# 12 页进阶效果版
python src/build_pro.py
# -> 平滑切换演示_MorphPRO版.pptx

# 24 页 ULTRA 版（复用 build_pro 前 12 帧，对象池扩到 20）
python src/build_ultra.py
# -> 平滑切换演示_MorphULTRA版.pptx
```

两个脚本**不需要外部 BASE**：首次运行会自动调用 `build_base` 从零生成
`平滑切换演示_Morph版.pptx`（9 页基础稿），再在其上追加新页。
产物落在当前工作目录（默认仓库根）。

### 3) 真机验证 Morph 是否生效

自检只能证明「文件没坏」，**不能证明 Morph 真生效**——`mc:AlternateContent`
的 Choice 分支在 LibreOffice 一律走 Fallback。**必须**在 PowerPoint
2019 / 2021 / 2024 / 365 或 WPS 中按 F5 播放实测。

## 技术要点（改页面前必读）

- **对象配对**：相邻页对象同名才被 Morph 识别；`!!` 前缀可强制配对两个
  完全不同形状（圆角矩形 → 六边形也能长过去）。
- **跨页 shape id 一致**：每页固定 N 个对象、id 一一对应，新增页不重建编号。
- **切换声明**：每页 `p159:morph option="byObject|byWord|byChar"`，外层用
  `mc:AlternateContent` 包 `fade` 兜底，旧版退化成淡入而非打不开。
- **自动连播**：`p14:dur` 控时长、`advTm` 控停留、`p14:advClick="0"` 免点击，
  配合 `presProps.xml` 的 `useTimings="1"`。
- **效果本质**：只改同一批对象的坐标/尺寸/旋转/填充/文字，让 Morph 算中间帧；
  退场对象挪出画布而非删除；异形靠 `!!` + 不同 `prstGeom`。

## 失败排查

若生成的文件 PowerPoint 打不开或导入失败，按 `OOXML_LESSONS.md` 的
**10 个坑位 / 6 层验证 / 硬约束清单 / 决策树**逐项核对。常见：
`sldIdLst` 位置、MCE 命名空间、`p:pic`/`p:graphicFrame` 前缀、
`[Content_Types].xml` 缺失、关系与部件不对应、`am3d:model3d` 等新命名空间
Choice 分支（**无真机样本 XML 不要写**）。

## 约定

- 纯文本资产：不提交 `.pptx`、图片、视频、二进制（`.gitignore` 已排除）。
- 每个 `.py` 脚本尾部自带 L1–L4 自检；改完必跑 `python src/selftest.py`。
- 详见 `README.md`（技术说明）与 `USAGE.md`（人读使用说明）。
