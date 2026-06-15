#!/usr/bin/env python3
"""
ppt_to_pdf.py — PowerPoint VBA 插件辅助脚本。

供 PowerPoint 宏调用，将当前 PPT 一键转换为 PDF（附带裁剪白边）。

策略:
  1. 用 PowerPoint COM 直接导出为 PDF（原生质量，速度快，文件小）
  2. 从第一页 EMF 检测白边裁剪量
  3. 用 pikepdf 裁剪所有 PDF 页面

输出:
    与 PPT 同目录下的同名 PDF 文件（如 presentation.pptx → presentation.pdf）
    与 PPT 同目录下的同名 .log 文件（调试用）
"""

import sys
import time
import logging
from pathlib import Path

# 将项目根目录和 src/ 加入模块搜索路径
_root = Path(__file__).parent.resolve()
_src = str(_root / "src")
if _src not in sys.path:
    sys.path.insert(0, _src)
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))


def _get_crop_bounds_from_png(png_path: str) -> tuple[tuple[float, float, float, float], int, int]:
    """
    从 PNG 图片检测白边裁剪量。

    Returns:
        (left, top, right, bottom) 裁剪像素数, 图片宽度, 图片高度
    """
    from PIL import Image
    import numpy as np

    img = np.array(Image.open(png_path).convert("RGBA"))
    h, w = img.shape[:2]
    rgb = img[:, :, :3]
    alpha = img[:, :, 3]

    white_mask = (rgb[:, :, 0] == 255) & (rgb[:, :, 1] == 255) & (rgb[:, :, 2] == 255)
    white_mask |= (alpha == 0)

    rows_all_white = np.all(white_mask, axis=1)
    cols_all_white = np.all(white_mask, axis=0)

    if np.all(rows_all_white):
        return (0, 0, 0, 0), w, h

    top = int(np.argmax(~rows_all_white))
    bottom = int(len(rows_all_white) - np.argmax(~rows_all_white[::-1]))
    left = int(np.argmax(~cols_all_white))
    right = int(len(cols_all_white) - np.argmax(~cols_all_white[::-1]))

    try:
        Path(png_path).unlink()
    except Exception:
        pass

    return (left, top, w - right, h - bottom), w, h


def _crop_pdf_pages(pdf_path: str, crop_px: tuple[float, float, float, float], page_width: int, page_height: int, img_width: int, img_height: int):
    """
    用 pikepdf 裁剪 PDF 页面白边。

    crop_px: (left, top, right, bottom) 裁剪像素数（来自 PNG 检测）
    page_width, page_height: PDF 页面原始尺寸（点）
    img_width, img_height: PNG 图片尺寸（像素），用于缩放比例
    """
    import pikepdf
    import tempfile

    left, top, right, bottom = crop_px

    # 将像素裁剪量缩放到 PDF 点坐标
    scale_x = page_width / img_width
    scale_y = page_height / img_height

    crop_left = left * scale_x
    crop_top = top * scale_y
    crop_right = right * scale_x
    crop_bottom = bottom * scale_y

    with pikepdf.open(pdf_path) as pdf:
        for page in pdf.pages:
            w = float(page_width)
            h = float(page_height)

            # 设置 CropBox 裁剪白边
            new_left = crop_left
            new_bottom = crop_bottom
            new_right = w - crop_right
            new_top = h - crop_top

            page.trimbox = pikepdf.Rectangle(new_left, new_bottom, new_right, new_top)
            page.cropbox = pikepdf.Rectangle(new_left, new_bottom, new_right, new_top)

        # 保存到临时文件再替换，避免 overwrite 错误
        fd, tmp = tempfile.mkstemp(suffix=".pdf")
        import os
        os.close(fd)
        pdf.save(tmp, linearize=True)

    # 替换原文件
    import shutil
    shutil.move(tmp, pdf_path)


def ppt_to_pdf(ppt_path: str) -> str:
    """
    将 PPT 转换为 PDF，附带裁剪白边。

    策略:
      1. PowerPoint COM 直接导出 PDF（原生质量）
      2. 从第一页 EMF→PNG 检测裁剪量
      3. pikepdf 裁剪所有页面

    Args:
        ppt_path: PPT/PPTX/PPTM 文件路径

    Returns:
        生成的 PDF 文件路径
    """
    import win32com.client
    import pythoncom

    ppt = Path(ppt_path)
    if not ppt.exists():
        raise FileNotFoundError(f"PPT 文件不存在: {ppt_path}")

    suffix = ppt.suffix.lower()
    if suffix not in (".ppt", ".pptx", ".pptm", ".ppsm", ".potx", ".potm"):
        raise RuntimeError(f"不支持的文件格式: {suffix}")

    output_dir = ppt.parent
    pdf_path = output_dir / f"{ppt.stem}.pdf"

    pythoncom.CoInitialize()

    app = None
    pres = None

    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.DisplayAlerts = False

        pres = app.Presentations.Open(str(ppt.resolve()), WithWindow=False)

        # ── 步骤1: 获取页面尺寸 ──
        slide_width = pres.PageSetup.SlideWidth
        slide_height = pres.PageSetup.SlideHeight
        logging.info(f"  页面尺寸: {slide_width}x{slide_height}pt")

        # ── 步骤2: 导出第一页为 PNG（用于裁剪检测，不经过 EMF） ──
        png_temp = str(output_dir / f"~{ppt.stem}_crop.png")
        pres.Slides(1).Export(png_temp, "PNG", slide_width, slide_height)

        # ── 步骤3: PowerPoint 直接导出 PDF ──
        logging.info("PowerPoint 导出 PDF...")
        # SaveAs with ppSaveAsPDF (32) 比 ExportAsFixedFormat 更稳定
        pres.SaveAs(str(pdf_path), 32)
        logging.info(f"  PDF 已生成: {pdf_path}")

    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app = None
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # ── 步骤4: 从 PNG 检测裁剪量（PPT 已关闭，不冲突） ──
    logging.info("检测白边裁剪量...")
    crop_px, img_w, img_h = _get_crop_bounds_from_png(png_temp)
    logging.info(f"  裁剪: L={crop_px[0]} R={crop_px[2]} T={crop_px[1]} B={crop_px[3]}px")

    # 清理临时 PNG
    try:
        Path(png_temp).unlink()
    except Exception:
        pass

    # ── 步骤5: 裁剪 PDF 白边 ──
    if any(c > 0 for c in crop_px):
        logging.info("裁剪 PDF 白边...")
        _crop_pdf_pages(str(pdf_path), crop_px, slide_width, slide_height, img_w, img_h)
        logging.info("  裁剪完成")
    else:
        logging.info("  无需裁剪")

    if not pdf_path.exists():
        raise RuntimeError(f"PDF 文件未生成: {pdf_path}")

    return str(pdf_path)


def main():
    if len(sys.argv) < 2:
        print("用法: python ppt_to_pdf.py <PPT文件路径>")
        sys.exit(1)

    ppt_path = sys.argv[1]

    ppt = Path(ppt_path)
    log_path = ppt.with_suffix(".log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.FileHandler(str(log_path), encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
    logging.info(f"开始转换: {ppt_path}")

    try:
        start = time.time()
        result = ppt_to_pdf(ppt_path)
        elapsed = time.time() - start
        logging.info(f"[OK] PDF 已生成: {result} (耗时 {elapsed:.1f}s)")
    except Exception as e:
        logging.error(f"转换失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
