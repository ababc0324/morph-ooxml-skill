# -*- coding: utf-8 -*-
"""
selftest：无外部输入、无 PowerPoint，验证整条生成链路。

流程：
  1. build_base   -> 从零生成 9 页基础稿 + 结构自检
  2. build_pro    -> 12 页 PRO 版（含 L1-L4 自检）
  3. build_ultra  -> 24 页 ULTRA 版（含更严自检）
  4. 确认两个成品文件存在

用法（在仓库根目录）：
  python src/selftest.py
退出码 0 表示全部通过；否则非 0。
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build_base          # noqa: E402
import build_pro           # noqa: E402
import build_ultra         # noqa: E402


def main():
    # 1) 基础稿
    base = os.path.abspath(build_base.BASE)
    build_base.build_base(base)
    build_base.selftest_base(base)
    print(f"[1/4] base OK -> {base}")

    # 2) PRO
    build_pro.ensure_base()
    pro_out = os.path.abspath(build_pro.DST)
    if os.path.exists(pro_out):
        os.remove(pro_out)
    build_pro.main()
    assert os.path.exists(pro_out), "PRO 未生成"
    assert zipfile.ZipFile(pro_out).testzip() is None, "PRO zip corrupt"
    print(f"[2/4] pro OK -> {pro_out}")

    # 3) ULTRA
    if not os.path.exists(os.path.abspath(build_ultra.BASE)):
        from build_base import build_base as _bb
        _bb(os.path.abspath(build_ultra.BASE))
    ultra_out = os.path.abspath(build_ultra.DST)
    if os.path.exists(ultra_out):
        os.remove(ultra_out)
    build_ultra.main()
    assert os.path.exists(ultra_out), "ULTRA 未生成"
    assert zipfile.ZipFile(ultra_out).testzip() is None, "ULTRA zip corrupt"
    print(f"[3/4] ultra OK -> {ultra_out}")

    # 4) 内容类型与页数一致性（最终确认）
    for out, pages in ((pro_out, 12), (ultra_out, 24)):
        with zipfile.ZipFile(out) as z:
            pres = z.read("ppt/presentation.xml").decode("utf-8")
            assert pres.count("<p:sldId ") == pages, f"{out}: sldId count != {pages}"
            n_slides = sum(1 for n in z.namelist() if n.startswith("ppt/slides/slide")
                           and n.endswith(".xml"))
            assert n_slides == pages, f"{out}: slide parts {n_slides} != {pages}"
    print("[4/4] package consistency OK (PRO=12, ULTRA=24)")

    print("\nALL CHECKS PASSED")
    print("  PRO   ->", pro_out)
    print("  ULTRA ->", ultra_out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
