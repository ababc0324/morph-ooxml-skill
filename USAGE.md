# 使用说明（USAGE）

本仓库是「手写 OOXML 生成 PowerPoint 原生 Morph 平滑切换」的最小开源资产。
**纯文本、无二进制**，且**自包含**：不需要外部 BASE，`python src/*.py` 即可产出成品。

---

## 1. 环境

- **Python 3.10+**
- 依赖：`pip install -r requirements.txt`（仅 `lxml>=4.9`）

## 2. 快速开始

```bash
# ① 验证整条链路（无需 PowerPoint；退出码 0 = 通过）
python src/selftest.py

# ② 生成 12 页进阶效果版
python src/build_pro.py        # -> 平滑切换演示_MorphPRO版.pptx

# ③ 生成 24 页 ULTRA 版
python src/build_ultra.py      # -> 平滑切换演示_MorphULTRA版.pptx
```

- **无需 BASE**：首次运行会自动从零生成 9 页基础稿 `平滑切换演示_Morph版.pptx`
  （`build_base.py`），再在其上扩展。
- 产物生成在当前工作目录（建议在仓库根运行）。
- 两个生成器都带**内置结构自检**：`zipfile.testzip()` 为 `None`、全部
  `.xml`/`.rels` 可解析、`!!` 同名对象跨页一致、`p159:morph` 声明在位、
  `sldIdLst` 页数与部件一一对应、shape id 唯一。

## 3. 各文件职责

| 文件 | 作用 |
|---|---|
| `src/build_base.py` | 从零生成 9 页基础稿（包结构脚手架），供两个生成器复用 |
| `src/build_pro.py` | 12 页 PRO 版（14 对象，byObject×10 / byWord / byChar） |
| `src/build_ultra.py` | 24 页 ULTRA 版（20 对象，9 类新效果，更严自检） |
| `src/selftest.py` | 一键跑通 基础稿 → PRO → ULTRA + 全部结构自检 |
| `OOXML_LESSONS.md` | 手写 OOXML 导入失败的 10 坑、6 层校验、硬约束、决策树 |
| `SKILL.md` | agent 可读的技能入口（即开即用说明） |
| `README.md` | 技术要点与效果原理 |

## 4. 验证 Morph 是否真生效（重要）

> 内置自检只能证明「文件没坏」，**不能证明 Morph 真生效**。凡涉及
> `mc:AlternateContent` 的 Choice 分支，LibreOffice 一律走 Fallback。

必须在 **PowerPoint 2019 / 2021 / 2024 / 365 或 WPS** 中打开产物，**按 F5**
播放；正常应看到对象位置/大小/旋转/颜色被平滑补间。若看到的是淡入淡出
而非 Morph，说明当前渲染器不支持 `p159:morph`，属预期降级。

## 5. 常见问题

- **PowerPoint 打不开 / 导入失败**：按 `OOXML_LESSONS.md` 的 10 坑位与决策树排查，
  常见为 `sldIdLst` 位置、MCE 命名空间、`[Content_Types].xml` 缺失、关系错位。
- **想改页面**：读 `build_pro.py` 的 `FR`（关键帧列表）——每页是同一批对象的
  属性表，改坐标/尺寸/旋转/填充/文字即可；**永远复制前一页再改，不要重画**。
- **新增命名空间元素**（如 `am3d:model3d`、`sectionZm`）：没有真机生成的样本
  XML **不要写**，否则一旦被 PowerPoint 选中会触发严格 schema 校验，本地无法验证。

## 6. 许可

MIT（见 `LICENSE`）。代码与文档可自由使用、修改、再分发。
