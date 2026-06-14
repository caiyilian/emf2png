"""
裁剪 PNG 纯白边 (#FFFFFF) 模块。

四边扫描法，严格匹配纯白像素。
非白底背景自动跳过，保持原样。
"""

from pathlib import Path


def trim_white_borders(png_path: str) -> str:
    """
    裁剪单张 PNG 的纯白边。

    若图片非白底（四边不全白），跳过不处理。

    Args:
        png_path: PNG 文件路径

    Returns:
        处理后的 PNG 文件路径（原地覆盖）
    """
    # TODO: 阶段4 用 Pillow + numpy 实现
    raise NotImplementedError("裁剪白边模块尚未实现（阶段4）")


def batch_trim_white_borders(png_files: list[str]) -> list[str]:
    """批量裁剪白边，就地覆盖。"""
    for f in png_files:
        trim_white_borders(f)
    return png_files
