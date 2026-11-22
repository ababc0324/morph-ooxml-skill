# -*- coding: utf-8 -*-
"""
build_base：从零生成一个合法的“9 页基础稿”PPTX（平滑切换演示_Morph版.pptx）。

用途：
  让 build_pro.py / build_ultra.py 无需外部 BASE 即可运行（agent 即开即用）。
  本模块只负责生成“包结构”脚手架 —— 9 张基础页内容会被 build_pro 完全改写，
  这里只保证包合法（content types / rels / master / layout / theme / 空 slide）。

产出：
  ./平滑切换演示_Morph版.pptx   （9 张空基础页，slide 编号与 build_pro 的 old 一致）
"""
import os
import zipfile

BASE = "./平滑切换演示_Morph版.pptx"

# 与 build_pro.OLD 完全一致的物理页号（build_pro 会逐张改写）
SLIDES = [1, 7, 4, 3, 8, 5, 11, 10, 12]

NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

T_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
T_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
T_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
T_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
T_PRES = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
T_PRESPROPS = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"
T_VIEWPROPS = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"
T_TABLESTYLES = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"
T_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
T_APP = "application/vnd.openxmlformats-officedocument.extended-properties+xml"

R_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_SLIDE = R_OFFICE + "/slide"
R_MASTER = R_OFFICE + "/slideMaster"
R_LAYOUT = R_OFFICE + "/slideLayout"
R_THEME = R_OFFICE + "/theme"
R_PRES = R_OFFICE + "/officeDocument"
R_PRESPROPS = R_OFFICE + "/presProps"
R_VIEWPROPS = R_OFFICE + "/viewProps"
R_TABLESTYLES = R_OFFICE + "/tableStyles"
R_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
R_APP = R_OFFICE + "/extended-properties"


def _ct():
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="{T_SLIDE}"/>'
        for n in SLIDES)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{NS_CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/ppt/presentation.xml" ContentType="{T_PRES}"/>'
        f'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="{T_MASTER}"/>'
        f'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="{T_LAYOUT}"/>'
        f'<Override PartName="/ppt/theme/theme1.xml" ContentType="{T_THEME}"/>'
        f'<Override PartName="/ppt/presProps.xml" ContentType="{T_PRESPROPS}"/>'
        f'<Override PartName="/ppt/viewProps.xml" ContentType="{T_VIEWPROPS}"/>'
        f'<Override PartName="/ppt/tableStyles.xml" ContentType="{T_TABLESTYLES}"/>'
        f'<Override PartName="/docProps/core.xml" ContentType="{T_CORE}"/>'
        f'<Override PartName="/docProps/app.xml" ContentType="{T_APP}"/>'
        + overrides +
        '</Types>').encode("utf-8")


def _root_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}">'
        f'<Relationship Id="rId1" Type="{R_PRES}" Target="ppt/presentation.xml"/>'
        f'<Relationship Id="rId2" Type="{R_CORE}" Target="docProps/core.xml"/>'
        f'<Relationship Id="rId3" Type="{R_APP}" Target="docProps/app.xml"/>'
        '</Relationships>').encode("utf-8")


def _presentation():
    sid = "".join(
        f'<p:sldId id="{256 + i}" r:id="rId{2 + i}"/>'
        for i in range(len(SLIDES)))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:presentation xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{sid}</p:sldIdLst>'
        '<p:sldSz cx="12192000" cy="6858000"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        '<p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr>'
        '<a:lvl1pPr algn="l"/><a:lvl2pPr algn="l"/><a:lvl3pPr algn="l"/>'
        '<a:lvl4pPr algn="l"/><a:lvl5pPr algn="l"/><a:lvl6pPr algn="l"/>'
        '<a:lvl7pPr algn="l"/><a:lvl8pPr algn="l"/><a:lvl9pPr algn="l"/></p:defaultTextStyle>'
        '</p:presentation>').encode("utf-8")


def _presentation_rels():
    rels = "".join(
        f'<Relationship Id="rId{2 + i}" Type="{R_SLIDE}" Target="slides/slide{n}.xml"/>'
        for i, n in enumerate(SLIDES))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}">'
        f'<Relationship Id="rId1" Type="{R_MASTER}" Target="slideMasters/slideMaster1.xml"/>'
        + rels +
        f'<Relationship Id="rId11" Type="{R_PRESPROPS}" Target="presProps.xml"/>'
        f'<Relationship Id="rId12" Type="{R_VIEWPROPS}" Target="viewProps.xml"/>'
        f'<Relationship Id="rId13" Type="{R_TABLESTYLES}" Target="tableStyles.xml"/>'
        '</Relationships>').encode("utf-8")


def _presprops():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:presentationPr xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
        '<p:showPr useTimings="1"/>'
        '</p:presentationPr>').encode("utf-8")


def _viewprops():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:viewPr xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
        '<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>'
        '<p:slideViewPr><p:cSldViewPr snapToGrid="0"><p:cViewPr varScale="1">'
        '<p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale>'
        '<p:origin x="0" y="0"/></p:cViewPr></p:cSldViewPr></p:slideViewPr>'
        '<p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/>'
        '<a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr>'
        '<p:gridSpacing cx="72008" cy="72008"/>'
        '</p:viewPr>').encode("utf-8")


def _tablestyles():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<a:tblStyleLst xmlns:a="{NS_A}" def="{R_OFFICE}/tableStyles/mediumStyle2Accent1"/>'
        ).encode("utf-8")


def _theme():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<a:theme xmlns:a="{NS_A}" name="MorphTheme">'
        '<a:themeElements>'
        '<a:clrScheme name="Morph">'
        '<a:dk1><a:sysClr val="windowText" lastClr="050B1A"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="1F2A44"/></a:dk2>'
        '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
        '<a:accent1><a:srgbClr val="25A7FF"/></a:accent1>'
        '<a:accent2><a:srgbClr val="7C3CFF"/></a:accent2>'
        '<a:accent3><a:srgbClr val="5EEAD4"/></a:accent3>'
        '<a:accent4><a:srgbClr val="3B82F6"/></a:accent4>'
        '<a:accent5><a:srgbClr val="F472B6"/></a:accent5>'
        '<a:accent6><a:srgbClr val="FBBF24"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
        '<a:hlinkFollowed><a:srgbClr val="954F72"/></a:hlinkFollowed>'
        '</a:clrScheme>'
        '<a:fontScheme name="Morph">'
        '<a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/>'
        '<a:cs typeface=""/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/>'
        '<a:cs typeface=""/></a:minorFont>'
        '</a:fontScheme>'
        '<a:fmtScheme name="Morph">'
        '<a:fillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs>'
        '</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>'
        '</a:fillStyleLst>'
        '<a:lnStyleLst>'
        '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
        '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
        '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
        '</a:lnStyleLst>'
        '<a:effectStyleLst>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '</a:effectStyleLst>'
        '<a:bgFillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '</a:bgFillStyleLst>'
        '</a:fmtScheme>'
        '</a:themeElements>'
        '<a:objectDefaults/>'
        '<a:extraClrSchemeLst/>'
        '</a:theme>').encode("utf-8")


def _theme_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}"/>').encode("utf-8")


def _master():
    # 空 spTree + 1 个布局 + 文本/填充/线样式（最简但合法）
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sldMaster xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
        '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
        'accent6="accent6" hlink="hlink" hlinkFollowed="hlinkFollowed"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        '<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400"/></a:lvl1pPr>'
        '</p:titleStyle><p:bodyStyle><a:lvl1pPr algn="l"><a:defRPr sz="1800"/>'
        '</a:lvl1pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="zh-CN"/>'
        '</a:defPPr></p:otherStyle></p:txStyles>'
        '</p:sldMaster>').encode("utf-8")


def _master_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}">'
        f'<Relationship Id="rId1" Type="{R_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/>'
        f'<Relationship Id="rId2" Type="{R_THEME}" Target="../theme/theme1.xml"/>'
        '</Relationships>').encode("utf-8")


def _layout():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sldLayout xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}" '
        'type="blank" preserve="1">'
        '<p:cSld name="空白"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sldLayout>').encode("utf-8")


def _layout_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}">'
        f'<Relationship Id="rId1" Type="{R_MASTER}" Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>').encode("utf-8")


def _slide(n):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
        f'<p:cSld name="slide{n}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sld>').encode("utf-8")


def _slide_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}">'
        f'<Relationship Id="rId1" Type="{R_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>').encode("utf-8")


def _core():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>平滑切换演示 · Morph 基础稿</dc:title>'
        '<dc:creator>morph-ooxml-skill</dc:creator>'
        '<cp:lastModifiedBy>morph-ooxml-skill</cp:lastModifiedBy>'
        '</cp:coreProperties>').encode("utf-8")


def _app():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>morph-ooxml-skill</Application>'
        '<Slides>9</Slides>'
        '</Properties>').encode("utf-8")


def build_base(dst=BASE):
    """从零写出 9 页基础稿。返回 dst 路径。"""
    parts = {}
    parts["[Content_Types].xml"] = _ct()
    parts["_rels/.rels"] = _root_rels()
    parts["docProps/core.xml"] = _core()
    parts["docProps/app.xml"] = _app()
    parts["ppt/presentation.xml"] = _presentation()
    parts["ppt/_rels/presentation.xml.rels"] = _presentation_rels()
    parts["ppt/presProps.xml"] = _presprops()
    parts["ppt/viewProps.xml"] = _viewprops()
    parts["ppt/tableStyles.xml"] = _tablestyles()
    parts["ppt/theme/theme1.xml"] = _theme()
    parts["ppt/theme/_rels/theme1.xml.rels"] = _theme_rels()
    parts["ppt/slideMasters/slideMaster1.xml"] = _master()
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = _master_rels()
    parts["ppt/slideLayouts/slideLayout1.xml"] = _layout()
    parts["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = _layout_rels()
    for n in SLIDES:
        parts[f"ppt/slides/slide{n}.xml"] = _slide(n)
        parts[f"ppt/slides/_rels/slide{n}.xml.rels"] = _slide_rels()

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return dst


def selftest_base(dst=BASE):
    """基础稿结构自检：zip 完整 + 全部 xml/rels 可解析 + 关键部件存在。"""
    from lxml import etree
    import zipfile as _z
    with _z.ZipFile(dst) as z:
        assert z.testzip() is None, "base zip corrupt"
        need = ["[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml",
                "ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml",
                "ppt/theme/theme1.xml"]
        for n in need:
            assert n in z.namelist(), f"missing {n}"
        for n in z.namelist():
            if n.endswith(".xml") or n.endswith(".rels"):
                etree.fromstring(z.read(n))
        pres = z.read("ppt/presentation.xml").decode("utf-8")
        assert pres.count("<p:sldId ") == len(SLIDES), "sldIdLst count"
    return True


if __name__ == "__main__":
    build_base()
    selftest_base()
    print("OK ->", BASE)
    print(f"  {len(SLIDES)} 张基础页 / slide 编号 {SLIDES}")
