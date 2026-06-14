"""
PNG → PDF 合并模块。

将多张 PNG 图片按页码顺序合并为一个 PDF 文件。
使用 img2pdf 直接嵌入原图数据，不重新编码，保持最高质量。
"""

from __future__ import annotations

from pathlib import Path

import img2pdf


def images_to_pdf(png_files: list[str], output_pdf: str) -> str:
    """
    将 PNG 图片列表合并为 PDF。

    每张 PNG 作为 PDF 的一页，保持原始分辨率。
    使用 img2pdf 直接将 PNG 数据嵌入 PDF，不经过 Pillow 重新编码，
    保证图片质量无损。

    Args:
        png_files: PNG 文件路径列表（按页码顺序）
        output_pdf: 输出 PDF 路径

    Returns:
        生成的 PDF 文件路径

    Raises:
        FileNotFoundError: 任一 PNG 文件不存在
        RuntimeError: PDF 生成失败
    """
    # 校验文件存在
    for f in png_files:
        if not Path(f).exists():
            raise FileNotFoundError(f"PNG 文件不存在: {f}")

    if not png_files:
        raise RuntimeError("没有可合并的 PNG 文件")

    pdf_path = Path(output_pdf)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        # img2pdf.convert 返回 PDF 字节数据
        pdf_bytes = img2pdf.convert(
            png_files,
            # 每页独立，不拼接多图到一页
            layout_fun=img2pdf.get_layout_fun(None),
        )
        pdf_path.write_bytes(pdf_bytes)
    except Exception as e:
        raise RuntimeError(f"PDF 生成失败: {e}")

    if not pdf_path.exists():
        raise RuntimeError(f"PDF 文件未生成: {output_pdf}")

    return str(pdf_path.resolve())
