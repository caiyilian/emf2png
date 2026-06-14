"""
裁剪 PNG 纯白边 (#FFFFFF) 模块。

四边扫描法，严格匹配纯白像素。
非白底背景自动跳过，保持原样。
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image
import numpy as np


# Pillow 默认图片尺寸限制较小，对大图（如 4x 高清导出）会报警
# 这里调高限制避免 DecompressionBombWarning
Image.MAX_IMAGE_PIXELS = 500_000_000


def _is_white_background(arr: np.ndarray) -> bool:
    """
    判断图片是否为白底背景。

    检查四边（首行、末行、首列、末列）是否全部为纯白像素。
    透明像素 (alpha=0) 也视为白。

    Args:
        arr: RGBA numpy 数组，形状 (H, W, 4)

    Returns:
        四边全白 → True（认为是白底）
        任意边框存在非白像素 → False（非白底，跳过裁剪）
    """
    alpha = arr[:, :, 3]
    rgb = arr[:, :, :3]

    white_mask = (rgb[:, :, 0] == 255) & (rgb[:, :, 1] == 255) & (rgb[:, :, 2] == 255)
    white_mask |= (alpha == 0)  # 透明像素视为白

    rows_all_white = np.all(white_mask, axis=1)  # 每行是否全白
    cols_all_white = np.all(white_mask, axis=0)  # 每列是否全白

    # 四边都全白 → 白底
    return bool(rows_all_white[0] and rows_all_white[-1] and cols_all_white[0] and cols_all_white[-1])


def _find_crop_bounds(arr: np.ndarray) -> tuple[int, int, int, int] | None:
    """
    找到内容区域的边界（跳过纯白行列）。

    Returns:
        (left, top, right, bottom) 裁剪边界
        如果图片纯白（无内容），返回 None
    """
    alpha = arr[:, :, 3]
    rgb = arr[:, :, :3]

    white_mask = (rgb[:, :, 0] == 255) & (rgb[:, :, 1] == 255) & (rgb[:, :, 2] == 255)
    white_mask |= (alpha == 0)

    rows_all_white = np.all(white_mask, axis=1)
    cols_all_white = np.all(white_mask, axis=0)

    # 如果所有行列都全白（纯白图片），不裁剪
    if np.all(rows_all_white) or np.all(cols_all_white):
        return None

    top = int(np.argmax(~rows_all_white))
    bottom = int(len(rows_all_white) - np.argmax(~rows_all_white[::-1]))
    left = int(np.argmax(~cols_all_white))
    right = int(len(cols_all_white) - np.argmax(~cols_all_white[::-1]))

    return (left, top, right, bottom)


def trim_white_borders(png_path: str) -> str:
    """
    裁剪单张 PNG 的纯白边 (#FFFFFF)。

    检测逻辑：
      1. 图片四边（首行、末行、首列、末列）全部为纯白或透明
         → 认为白底，执行裁剪
      2. 否则 → 非白底背景，跳过不处理

    裁剪方式：
      逐行列扫描，找到第一个非白像素的边界，矩形裁切。
      透明像素 (alpha=0) 也视为白边一并裁掉。

    Args:
        png_path: PNG 文件路径

    Returns:
        处理后的 PNG 文件路径（原地覆盖）
    """
    img = Image.open(png_path).convert("RGBA")
    arr = np.array(img)

    # 判断是否为白底
    if not _is_white_background(arr):
        return png_path  # 非白底，跳过

    # 找到裁剪边界
    bounds = _find_crop_bounds(arr)
    if bounds is None:
        return png_path  # 纯白图，跳过

    left, top, right, bottom = bounds

    # 如果裁剪后尺寸不变，跳过
    if left == 0 and top == 0 and right == img.width and bottom == img.height:
        return png_path

    cropped = img.crop((left, top, right, bottom))
    cropped.save(png_path)

    return png_path


def batch_trim_white_borders(
    png_files: list[str],
    progress_callback=None,
) -> list[str]:
    """批量裁剪白边，就地覆盖。"""
    for f in png_files:
        trim_white_borders(f)
        if progress_callback:
            progress_callback(1, len(png_files), Path(f).name)
    return png_files
