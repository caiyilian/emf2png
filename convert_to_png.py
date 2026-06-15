#!/usr/bin/env python3
"""
convert_to_png.py — 右键菜单调用：将 EMF/PPT 转为 PNG（裁剪白边）。

用法:
    uv run python convert_to_png.py <文件路径>

EMF → PNG:  1x, DPI 300, 裁剪白边
PPT → PNG:  1x, DPI 100, 多页加后缀 _001, _002...
"""

import sys
import time
from pathlib import Path

# 将项目根目录和 src/ 加入模块搜索路径
_root = Path(__file__).parent.resolve()
_src = str(_root / "src")
if _src not in sys.path:
    sys.path.insert(0, _src)
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))


def convert_emf_to_png(emf_path: str) -> list[str]:
    """EMF → PNG（裁剪白边，1x，300DPI）。"""
    from emf_to_png import convert
    from trim_whitespace import trim_white_borders

    out = Path(emf_path).with_suffix(".png")
    convert(emf_path, str(out), scale=1.0, dpi=300)
    trim_white_borders(str(out))
    return [str(out)]


def convert_ppt_to_png(ppt_path: str) -> list[str]:
    """PPT → PNG（裁剪白边，1x，100DPI，多页加后缀）。"""
    from ppt_to_emf import ppt_to_emf
    from emf_to_png import batch_convert
    from trim_whitespace import batch_trim_white_borders

    ppt = Path(ppt_path)
    out_dir = ppt.parent

    # PPT → EMF
    emf_files = ppt_to_emf(
        ppt_path=str(ppt.resolve()),
        output_dir=str(out_dir),
        start=1,
        end=None,
    )

    # EMF → PNG（1x, 100DPI）
    png_files = batch_convert(
        emf_files=emf_files,
        output_dir=str(out_dir),
        scale=1.0,
        dpi=100,
        keep_emf=False,
    )

    # 裁剪白边
    batch_trim_white_borders(png_files)

    # 如果只有一页，去掉 _001 后缀
    if len(png_files) == 1:
        single = out_dir / f"{ppt.stem}.png"
        Path(png_files[0]).rename(single)
        return [str(single)]

    return png_files


def main():
    if len(sys.argv) < 2:
        print("用法: python convert_to_png.py <文件路径>")
        sys.exit(1)

    path = sys.argv[1]
    suffix = Path(path).suffix.lower()

    start = time.time()
    try:
        if suffix == ".emf":
            results = convert_emf_to_png(path)
        elif suffix in (".ppt", ".pptx", ".pptm"):
            results = convert_ppt_to_png(path)
        else:
            print(f"不支持的文件格式: {suffix}")
            sys.exit(1)

        elapsed = time.time() - start
        for r in results:
            size_kb = Path(r).stat().st_size / 1024
            print(f"[OK] {r} ({size_kb:.0f} KB)")
        print(f"耗时: {elapsed:.1f}s")

    except Exception as e:
        print(f"[ERR] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
