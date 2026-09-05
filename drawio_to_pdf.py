#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
drawio_to_pdf.py — Draw.io 文件转单页 PDF并裁剪白边。

流程:
    .drawio → draw.io CLI 导出临时 PDF → 宽松像素边界裁剪 → 最终 PDF

用法:
    python drawio_to_pdf.py diagram.drawio
    python drawio_to_pdf.py diagram.drawio -o diagram_trimmed.pdf
    python drawio_to_pdf.py diagram.drawio --zoom 3 --dpi 300

draw.io 路径优先级:
    --drawio-exe → DRAWIO_EXE 环境变量 → 常见安装路径 → PATH
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import tempfile
from pathlib import Path


DEFAULT_TIMEOUT = 120


def _find_crop_bounds(arr, strict: bool = False):
    """按 drawio2pdf_trim.py 的规则找到内容边界。

    宽松模式把 RGB 全部不小于 248 的像素视为空白，并允许每行/列有
    0.5% 的非空白像素。这一点很重要：Draw.io 导出的页面背景可能是
    接近白色但不是纯白色，不能先用“是否纯白底”把整页裁剪跳过。
    """
    import numpy as np

    alpha = arr[:, :, 3]
    rgb = arr[:, :, :3]

    if strict:
        white_mask = (
            (rgb[:, :, 0] == 255)
            & (rgb[:, :, 1] == 255)
            & (rgb[:, :, 2] == 255)
        )
    else:
        white_mask = (
            (rgb[:, :, 0] >= 248)
            & (rgb[:, :, 1] >= 248)
            & (rgb[:, :, 2] >= 248)
        )
    white_mask |= alpha == 0

    if strict:
        rows_all_white = np.all(white_mask, axis=1)
        cols_all_white = np.all(white_mask, axis=0)
    else:
        rows_all_white = np.mean(white_mask, axis=1) >= 0.995
        cols_all_white = np.mean(white_mask, axis=0) >= 0.995

    # 整页为空白时不生成零尺寸页面。
    if np.all(rows_all_white) or np.all(cols_all_white):
        return None

    top = int(np.argmax(~rows_all_white))
    bottom = int(len(rows_all_white) - np.argmax(~rows_all_white[::-1]))
    left = int(np.argmax(~cols_all_white))
    right = int(len(cols_all_white) - np.argmax(~cols_all_white[::-1]))
    return (left, top, right, bottom)


def _trim_drawio_pdf(
    src: Path,
    out: Path,
    zoom: float = 2.5,
    dpi: int = 300,
    strict: bool = False,
) -> Path:
    """使用 drawio2pdf_trim.py 的宽松算法裁剪单页 PDF。"""
    import img2pdf
    import numpy as np
    import pymupdf
    from PIL import Image

    Image.MAX_IMAGE_PIXELS = 500_000_000

    with pymupdf.open(src) as doc:
        if doc.page_count != 1:
            raise RuntimeError(f"仅支持单页 PDF，该文件共 {doc.page_count} 页")

        page = doc[0]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("RGBA")

    print(f"[1/3] 渲染完成: {pix.width}x{pix.height}px (zoom={zoom})")

    arr = np.array(img)
    bounds = _find_crop_bounds(arr, strict=strict)
    if bounds is None:
        print("[2/3] 整页纯白，跳过裁剪")
        bounds = (0, 0, img.width, img.height)
    else:
        print(
            f"[2/3] 裁剪白边: {img.width}x{img.height}px -> "
            f"{bounds[2] - bounds[0]}x{bounds[3] - bounds[1]}px"
        )

    cropped = img.crop(bounds).convert("RGB")
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="pdf_trim_") as tmp:
        tmp_png = Path(tmp) / "trimmed.png"
        cropped.save(tmp_png, dpi=(dpi, dpi))
        pdf_bytes = img2pdf.convert(
            str(tmp_png),
            layout_fun=img2pdf.get_layout_fun(None),
        )
        out.write_bytes(pdf_bytes)

    print(f"[3/3] 输出: {out} ({out.stat().st_size:,} bytes)")
    return out


def _candidate_drawio_paths(explicit: str | None = None) -> list[str]:
    """返回按优先级排列的 draw.io 可执行文件候选路径。"""
    candidates: list[str] = []

    if explicit:
        candidates.append(explicit)

    env_path = os.environ.get("DRAWIO_EXE")
    if env_path:
        candidates.append(env_path)

    candidates.extend(
        [
            r"E:\draw.io-30.0.4-windows\draw.io.exe",
            r"C:\Program Files\draw.io\draw.io.exe",
            r"C:\Program Files (x86)\draw.io\draw.io.exe",
            r"C:\Program Files\diagrams.net\diagrams.net.exe",
            r"C:\Program Files (x86)\diagrams.net\diagrams.net.exe",
        ]
    )

    for env_name in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(env_name)
        if base:
            candidates.extend(
                [
                    str(Path(base) / "draw.io" / "draw.io.exe"),
                    str(Path(base) / "diagrams.net" / "diagrams.net.exe"),
                ]
            )

    for command in ("drawio.exe", "draw.io.exe", "diagrams.net.exe"):
        found = shutil.which(command)
        if found:
            candidates.append(found)

    return candidates


def find_drawio(explicit: str | None = None) -> str:
    """查找 draw.io 可执行文件并返回绝对路径。"""
    seen: set[str] = set()
    for candidate in _candidate_drawio_paths(explicit):
        candidate = candidate.strip().strip('"')
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)

        path = Path(candidate).expanduser()
        if path.is_file():
            return str(path.resolve())

    raise FileNotFoundError(
        "找不到 draw.io 可执行文件。请安装 draw.io，或通过 --drawio-exe "
        "参数 / DRAWIO_EXE 环境变量指定路径。"
    )


def export_drawio_to_pdf(
    drawio: str,
    source: str | Path,
    output: str | Path,
    border: str = "0",
    timeout: int = DEFAULT_TIMEOUT,
) -> Path:
    """调用 draw.io CLI 将一个 .drawio 文件导出为 PDF。"""
    source_path = Path(source).resolve()
    output_path = Path(output).resolve()

    if not source_path.is_file():
        raise FileNotFoundError(f"Draw.io 文件不存在: {source_path}")
    if not source_path.name.lower().endswith((".drawio", ".drawio.xml")):
        raise ValueError(f"不支持的 Draw.io 文件格式: {source_path.suffix}")
    if timeout <= 0:
        raise ValueError("timeout 必须大于 0")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    command = [
        drawio,
        "--export",
        "--format",
        "pdf",
        "--crop",
        "--border",
        str(border),
        "--output",
        str(output_path),
        str(source_path),
    ]

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            f"draw.io 导出超时 ({timeout}s): {source_path.name}"
        ) from exc
    except OSError as exc:
        raise RuntimeError(f"无法启动 draw.io: {exc}") from exc

    if result.returncode != 0:
        details = "\n".join(part for part in (result.stdout, result.stderr) if part)
        raise RuntimeError(f"draw.io 导出失败 (exit={result.returncode}):\n{details}")

    if not output_path.is_file() or output_path.stat().st_size < 1000:
        raise RuntimeError("draw.io 导出完成但 PDF 缺失或文件过小")

    return output_path


def drawio_to_pdf(
    source: str | Path,
    output_pdf: str | Path | None = None,
    *,
    drawio_exe: str | None = None,
    border: str = "0",
    zoom: float = 2.5,
    dpi: int = 300,
    strict: bool = False,
    timeout: int = DEFAULT_TIMEOUT,
) -> str:
    """将 Draw.io 文件导出为单页 PDF，并裁剪四周白边。

    draw.io 的原始导出先写入临时目录；只有 PDF 成功渲染并完成裁剪后，
    才会写入 output_pdf，因此不会把中间未裁剪文件暴露给调用方。
    """
    source_path = Path(source).resolve()
    output_path = (
        Path(output_pdf).resolve()
        if output_pdf is not None
        else source_path.with_suffix(".pdf")
    )

    if zoom <= 0:
        raise ValueError("zoom 必须大于 0")
    if dpi <= 0:
        raise ValueError("dpi 必须大于 0")

    drawio_path = find_drawio(drawio_exe)
    with tempfile.TemporaryDirectory(prefix="drawio2pdf_") as temp_dir:
        raw_pdf = Path(temp_dir) / "drawio_export.pdf"
        trimmed_pdf = Path(temp_dir) / "drawio_trimmed.pdf"

        export_drawio_to_pdf(
            drawio=drawio_path,
            source=source_path,
            output=raw_pdf,
            border=border,
            timeout=timeout,
        )
        _trim_drawio_pdf(
            raw_pdf,
            trimmed_pdf,
            zoom=zoom,
            dpi=dpi,
            strict=strict,
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(trimmed_pdf, output_path)

    return str(output_path.resolve())


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="将 Draw.io 文件导出为单页 PDF，并裁剪四周白边"
    )
    parser.add_argument("input", type=Path, help="输入的 .drawio 文件路径")
    parser.add_argument(
        "output_positional",
        nargs="?",
        type=Path,
        help="输出 PDF 路径（可选，默认与输入同名）",
    )
    parser.add_argument("-o", "--output", type=Path, help="输出 PDF 路径")
    parser.add_argument(
        "--border",
        default="0",
        help="draw.io 导出时的边框宽度（默认: 0）",
    )
    parser.add_argument(
        "--zoom",
        type=float,
        default=2.5,
        help="白边检测渲染倍率，越高越清晰（默认: 2.5）",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="裁剪后 PDF 的 DPI（默认: 300）",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="严格模式：只把纯白 #FFFFFF 视为白色，默认允许抗锯齿误差",
    )
    parser.add_argument(
        "--drawio-exe",
        help="draw.io 可执行文件路径（也可设置 DRAWIO_EXE）",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
        help=f"draw.io 导出超时时间（秒，默认: {DEFAULT_TIMEOUT}）",
    )
    args = parser.parse_args(argv)

    if args.output and args.output_positional:
        parser.error("不能同时使用位置输出路径和 --output")

    output = args.output or args.output_positional
    try:
        drawio_path = find_drawio(args.drawio_exe)
        print(f"draw.io: {drawio_path}")
        result = drawio_to_pdf(
            args.input,
            output,
            drawio_exe=drawio_path,
            border=args.border,
            zoom=args.zoom,
            dpi=args.dpi,
            strict=args.strict,
            timeout=args.timeout,
        )
        size_kb = Path(result).stat().st_size / 1024
        print(f"OK: {result} ({size_kb:.0f} KB) - 已裁剪白边")
        return 0
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"错误: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
