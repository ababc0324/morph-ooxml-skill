# 平滑切换（Morph）PPT 构建 · OOXML 规则与生成器

一份关于「用 Python 手写 OOXML、在 PowerPoint 里做原生平滑切换（Morph）」的最小开源资产：
核心是**兼容性规则层**与**可复用的页面生成器骨架**。

> 本仓库只保留纯文本资产（文档 + 脚本骨架），不附带示例 pptx / 视频 / 图片 / 模型。
> 这样做一是控制体积、便于直接纳入版本库，二是避免交付物二进制里残留的编辑器/工具元数据。

---

## 这个仓库能给你什么

- **`OOXML_LESSONS.md`** ★ 最核心：把「手写 OOXML 为什么会导入失败」拆成 10 个坑位、6 层验证模型、一份硬约束清单和一张决策树。想在 PPTX 里加原生特性前，先读它。
- **`src/build_pro.py`** 基础 12 页 Morph 生成器：展示了 `!!` 同名对象配对、跨页 shape id 一致、`p159:morph` + `mc:AlternateContent` fade 兜底的完整写法。
- **`src/build_ultra.py`** 进阶生成器：复用 `build_pro` 的前 12 帧，把对象池扩到 20 个、叠出 9 类新效果，且自带严格的 schema 自检（sldIdLst 位置、页序关系与部件一一对应、id 唯一）。

---

## 重要说明：脚本是「骨架」，不是一键可跑

`build_pro.py` / `build_ultra.py` 遵循的是原工程的**增量链**模式：脚本读取一个已含基础页的 `BASE` pptx，在其尾部追加新页并改写内容类型与页序清单。

- 默认 `BASE = "./平滑切换演示_Morph版.pptx"`，**本仓库不含该文件**（示例交付物已移除）。
- 因此直接运行会卡在「找不到 BASE」。脚本的价值在于**展示 OOXML 正确写法**，作为你自己的生成流程的起点，而不是开箱即用的批处理工具。
- `build_ultra.py` 通过 `import build_pro as bp` 复用底层常量与帧源，两个文件需放在同一目录、用 Python 3.10+ 运行。

---

## 目录结构

```
.
├── README.md                  本文件
├── LICENSE                    MIT
├── requirements.txt           运行依赖（仅 lxml）
├── .gitignore                忽略二进制，防止误传
├── OOXML_LESSONS.md          ★ 兼容性复盘与硬约束
└── src/
    ├── build_pro.py          Morph 基础版生成器（12 页 / 14 对象）
    └── build_ultra.py        ULTRA 进阶生成器（24 页 / 20 对象）
```

---

## 核心技术要点

### Morph 平滑切换的三要素

1. **对象配对**：相邻页的对象同名即可被 Morph 识别；用 `!!` 前缀可强制配对两个完全不同的形状（圆角矩形 → 六边形也能长过去）。
2. **跨页 shape id 一致**：每页固定 N 个对象、id 一一对应，新增页不重建编号。
3. **切换声明**：每页 `p159:morph option="byObject|byWord|byChar"`，外层用 `mc:AlternateContent` 包 fade 兜底，旧版退化成淡入而不是打不开。

自动连播：`p14:dur` 控制切换时长、`advTm` 控制停留、`p14:advClick="0"` 免点击，配合 `presProps.xml` 的 `useTimings="1"`。常规 900ms、快闪 150–300ms、终章 1400ms 停留。

### 效果是怎么「堆」出来的

所有效果本质都是**只改同一批对象的坐标 / 尺寸 / 旋转 / 填充 / 文字**，让 Morph 算中间帧：

- 退场对象不是删掉，而是挪到画布外（补间会把它 fly 出去）
- 异形变换靠 `!!` 强制配对 + 不同 `prstGeom`
- 缩放定位、立方体、时钟、百叶窗、万花筒、正弦波、螺旋、多米诺……全是坐标与角度的组合，不含任何自绘动画时间线

### 原生特性要谨慎

手写 `p14:sectionLst`（节）、`a:hlinkClick`（标准超链接跳转）这类**标准元素**是安全的；但 `sectionZm` / `sldZm` / `am3d:model3d` 这类**新命名空间 Choice 分支**，一旦被 PowerPoint 选中就会触发严格 schema 校验，本地无法验证。**没有真机生成的样本 XML 就不要写**。详见 `OOXML_LESSONS.md`。

---

## 本地自检（改脚本后必跑）

```bash
python -m py_compile src/*.py

# 若已自备 BASE pptx：
python src/build_pro.py
python src/build_ultra.py
```

两个脚本尾部都内置严格自检（L1–L4 级）：`zipfile.testzip()` 为 `None`、全部 `.xml`/`.rels` 可解析、页序 rId 唯一且能解析到真实部件、`[Content_Types].xml` 全部声明、`<p:sldIdLst>` 位于 `<p:sldSz>` 之前。

> **自检只能证明「文件没坏」，不能证明 Morph 真生效。** 凡涉及 `mc:AlternateContent` 的 Choice 分支，LibreOffice 一律走 Fallback，必须真机按 F5 才能验证。请务必在 PowerPoint 2019 / 2021 / 2024 / 365 或 WPS 中实测。

---

## 许可

[MIT](./LICENSE)。代码与文档可自由使用、修改、再分发。
