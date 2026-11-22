# OOXML 兼容性复盘：两次「内容有问题」的完整归因

> 这份文档是整包最该先读的部分。它把两次真机失败拆成了可复用的硬约束，
> 也是「以后还想往 PPTX 里加原生特性」时的判断依据。

---

## 结论先行

**两句可直接当规则用：**

1. **能打开 ≠ 结构合规。** 本地 XML 解析全绿、LibreOffice 页数正确，
   都只证明「Fallback 分支没问题」；一旦 PowerPoint 认识某个命名空间，
   它会选中 `mc:Choice` 并做严格 schema 校验，此时手写扩展的任何一处错字都会让整包被判损坏。
2. **无法真机验证的扩展，不要写进 Choice。** 用标准元素（`p:pic`、`a:hlinkClick`、
   `p14:sectionLst`）实现 90% 的观感，比手写 `sectionZm` / `am3d:model3d` 安全一个数量级。
   真要上原生特性，唯一稳妥路径是**让 PowerPoint 自己插入一次并存盘，再以那份文件为模板**。

---

## 两次故障对照

| | 故障一（ULTRA 初版） | 故障二（NATIVE 版） |
|---|---|---|
| 现象 | 导入被拒 | PowerPoint 弹「发现内容有问题，可尝试修复」 |
| 触发范围 | 整包 | 整包 |
| 根因 | `presentation.xml` 里 `<p:sldIdLst>` 被放到了 `<p:sldSz>` 之后 | 手写的 Zoom / 3D `mc:Choice` 分支存在多处 schema 不合规 |
| 为什么本地没发现 | 本地只做 XML 可解析 + 部件存在性检查，没查元素顺序 | LibreOffice 不认 `p16sz/p16sl/am3d`，直接走 Fallback；本地渲染 72 页全绿 |
| 修法 | 不重建页序清单，只在 `</p:sldIdLst>` 前追加新页 | 移除全部手写 Choice；跨节跳转改标准超链接；3D 旋转改 `!!` 同名 + Morph |
| 操作建议 | 换修复后的文件 | **点「取消」，不要点「修复」**（修复会删掉不认识的部分，可能连 Morph 一起清掉） |

**共同点**：两次都是「文件结构本身能解析，但 PowerPoint 的严格校验不认」。
**差别**：第一次是标准元素的位置错；第二次是新命名空间里的内容错，且被 Fallback 掩盖。

---

## 坑位清单（按命中顺序）

### 坑 1 · `sldIdLst` 的位置是硬顺序

`presentation.xml` 中 `<p:sldIdLst>` 必须排在 `<p:sldSz>` 等节点**之前**。
扩页时重建整个 `presentation.xml` 很容易把它挪到末尾，PowerPoint / WPS / 在线导入会直接拒收。

```python
# 正确：在已验证文件的页序清单尾部追加，绝不重建节点位置
assert pres.find("<p:sldIdLst>") < pres.find("<p:sldSz")
pres = pres.replace("</p:sldIdLst>", new_sldIds + "</p:sldIdLst>")
```

新增页必须**三处同步**：`presentation.xml.rels`（加 slide 关系）、
`[Content_Types].xml`（加 slide Override）、`ppt/presentation.xml`（加 `p:sldId`）。
新增页 rId / sldId 从现有最大值 +1 开始，不动既有编号。

### 坑 2 · MCE 的验证盲区（关键教训）

```
mc:AlternateContent
├── Choice  Requires="am3d"   ← PowerPoint 认识 → 选中 → 严格校验
└── Fallback  p:pic           ← LibreOffice / 旧版走这里 → 渲染正常
```

**Fallback 做得越完整，越容易掩盖 Choice 的错误。** 本地所有静态检查、PDF 导出、
抽帧目视都只覆盖 Fallback。因此：

> 任何带 `Choice` 的扩展，本地结论只能写「包完整、Fallback 正常」，
> 绝不能写「特性已验证」。必须真机 F5，或干脆不写 Choice。

### 坑 3 · `p:pic` 的填充和形状属性是 `p:` 前缀，不是 `a:`

```xml
<!-- 错 -->
<p:pic>...<a:blipFill>...</a:blipFill><a:spPr>...</a:spPr></p:pic>
<!-- 对 -->
<p:pic>...<p:blipFill><a:blip r:embed="rIdX"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm>...</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>
```

`p:nvPicPr / p:blipFill / p:spPr` 属于 PresentationML；
内部的 `a:blip / a:xfrm / a:prstGeom` 才属于 DrawingML。混用前缀是最容易犯的低级错。

### 坑 4 · `p:graphicFrame` 的变换是 `p:xfrm`

图表、3D 模型等图形帧用 `<p:xfrm>`（不是 `<a:xfrm>`）承载位置尺寸：

```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>...</p:nvGraphicFramePr>
  <p:xfrm><a:off x=".." y=".."/><a:ext cx=".." cy=".."/></p:xfrm>
  <a:graphic><a:graphicData uri="...">...</a:graphicData></a:graphic>
</p:graphicFrame>
```

### 坑 5 · `am3d:model3d` 的元素名与必需子节点

依据公开的 MS-ODRAWXML 2.31.3.x：

| 写错 | 正确 | 说明 |
|---|---|---|
| `am3d:persp` | `am3d:perspective` | 透视投影元素名，`fov` 取值须在 (0,180) |
| `am3d:transform` | `am3d:trans` | 变换容器名 |
| 只写 `rot` | `meterPerModelUnit` + `preTrans` + `scale` + `rot` + `postTrans` | `preTrans` / `postTrans` 规范上必需 |
| 子元素顺序随意 | `spPr → camera → trans → raster → extLst? → objViewport → ambientLight → 灯光组` | 严格 sequence |

`camera` 子节点为 `pos / up / lookAt / perspective`（或 `orthographic`）。
这些名字只能从规范推，**没有真机样本时不要凭直觉拼**——这正是失败的直接原因。

### 坑 6 · Zoom 对象的结构与引用完整性

依据 MS-PPTX 2.2.15 / 2.8.3.1 / 2.9.3.2 / 2.10.3.2：

- `sectionZm` 属 `.../2016/sectionzoom`，`sldZm` 属 `.../2016/slidezoom`
- `zmPr` 子元素顺序固定：`blipFill` → `spPr`（`CT_ZoomObjectProperties`）
- `sectionId` 必须等于 `p14:sectionLst` 里某个 `p14:section` 的 GUID
- `sldZm` 的 `sldId` 必须是 `presentation.xml` 里真实存在的页序 id
- Fallback 必须是标准 `p:pic`；`summaryZm` 的 Fallback 是 `p:grpSp`

**这些约束互相牵连，一处 GUID 对不上就整包报错。**

### 坑 7 · `[Content_Types].xml` 的默认类型插入位置

新增二进制类型（如 `.glb` → `model/gltf-binary`）时要插在 `<Types ...>` 开标签之后：

```python
m = re.search(r"<Types\b[^>]*>", ct)
ct = ct[:m.end()] + '<Default Extension="glb" ContentType="model/gltf-binary"/>' + ct[m.end():]
```

不要用 `ct.replace(">", ...)` —— 会先命中 XML 声明的 `?>`，把声明打断，
产生 `Extra content at the end of the document`。

### 坑 8 · 关系与部件零悬空

所有 `.rels` 的 `Target` 都要能解析到包内真实部件；不能有重复部件名。
新增页最容易漏的是「rels 里写了图片，但媒体部件没写进 zip」——
本地若不遍历全部 rels 校验，LibreOffice 可能只是静默丢图，PowerPoint 会报损坏。

### 坑 9 · 透明对象上的残留文字

`alpha=0` 的形状若仍带 `<a:t>` 文本，PowerPoint 会照常把字画出来。
做「对象退场」时要同时清空 `text`，而不只是设透明。

### 坑 10 · 图片与图框比例

图片会被拉伸填充到图框。3:1 的章节卡要配 3:1 的框，正方形 poster 配正方形框，
否则卡片变形、文字被压扁。半透明装饰要画在独立 RGBA 图层再合成，
直接 `ImageDraw` 到 RGB 底会丢 alpha 变成实心色块。

---

## 验证分层模型（建议直接固化进流程）

| 层 | 能证明什么 | 不能证明什么 |
|---|---|---|
| L1 `zipfile.testzip()` + 无重复部件 | 包没坏 | 任何 XML 语义 |
| L2 全部 `.xml`/`.rels` 可解析 | 格式合法 | 元素顺序、命名空间正确性 |
| L3 页序 / rels / Content_Types 一一对应，零悬空 | 引用完整 | PowerPoint 的 schema 校验 |
| L4 根元素与已验证可导入版本逐字对比 | 没引入新命名空间 | Choice 内部内容 |
| L5 LibreOffice 无头导出 PDF + 抽帧 | **Fallback 能渲染、页数正确** | Morph 补间、任何 Choice 分支 |
| L6 **真机 PowerPoint F5** | Morph / Zoom / 3D 是否真生效 | — |

**L1–L5 全绿不等于能导入**：NATIVE 版就是 L1–L5 全绿、L6 失败。
反过来，只要 L1–L4 用了「与已验证基线逐字对比」，风险就极低——
最终修复版正是这么做的：幻灯片骨架直接复用可导入版本的生成函数，根命名空间一个字不改。

---

## 硬约束清单（可直接抄作规则）

- [ ] 扩页时**只在** `</p:sldIdLst>` 前追加，断言其位置在 `<p:sldSz>` 之前；不重建 `presentation.xml`
- [ ] 新增页同步三处：presentation rels / Content_Types Override / sldId
- [ ] 新 rId / sldId 从现有最大值 +1，既有编号一律不动
- [ ] 遍历全部 `.rels` 校验 Target 存在；检查重复部件名
- [ ] 新增二进制类型用正则定位 `<Types ...>` 开标签后插入 Default
- [ ] `p:pic` 用 `p:blipFill` / `p:spPr`；`p:graphicFrame` 用 `p:xfrm`
- [ ] `!!` 同名对象跨页 id 一致；每页 `p159:morph` + `mc:AlternateContent` fade 兜底
- [ ] 透明对象同时清空 `text`
- [ ] 图片比例与图框比例一致
- [ ] **`Choice` 分支无真机样本则不写**；宁可用标准元素 + 超链接替代
- [ ] 交付说明里必须写明「本地只验到 L5，Choice 行为未验证」
- [ ] 出片门禁：`presProps.xml` 含 `useTimings="1"`；LibreOffice 页数与预期一致

---

## 决策树：要不要手写原生扩展？

```
需要新特性
  ├─ 能用「标准形状改坐标 + Morph」表达？ ── 是 ──▶ 做，零风险
  ├─ 能用标准元素近似表达（图片 / 超链接 / 节分组）？ ── 是 ──▶ 做，说明清楚少了什么
  └─ 必须写新命名空间的 Choice 分支
        ├─ 手上有 PowerPoint 真机生成的样本 XML？ ── 是 ──▶ 逐字照抄结构
        └─ 没有 ─────────────────────────────────▶ 不写。
                                                     用 PowerPoint 插入一次并存盘，
                                                     以那份文件为模板再做合并。
```

NATIVE 版就卡在最后那个「没有」上：规范只给了元素清单，没给可直接照抄的样本，
凭推断拼出来的 `am3d` 子树被 PowerPoint 判为不合规。**这个判断不该靠赌。**
