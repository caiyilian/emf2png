"""
PPT → EMF 转换模块。

使用 PowerPoint COM 自动化，后台将每页幻灯片导出为 EMF 文件。
仅在 Windows + Microsoft Office 环境下可用。
"""

from pathlib import Path


def ppt_to_emf(
    ppt_path: str,
    output_dir: str,
    start: int = 1,
    end: int | None = None,
    width: int = 1920,
    height: int = 1080,
    progress_callback=None,
) -> list[str]:
    """
    将 PPT 每页导出为 EMF 文件。

    Args:
        ppt_path: PPT/PPTX 文件路径
        output_dir: EMF 输出目录
        start: 起始页码 (1-based)
        end: 结束页码 (None 表示全部)
        width: 导出宽度
        height: 导出高度
        progress_callback: 可选进度回调 fn(current, total, filename)

    Returns:
        EMF 文件路径列表
    """
    # TODO: 阶段3 实现 win32com 调用
    raise NotImplementedError(
        "PPT→EMF 模块尚未实现。请先安装 pywin32: pip install pywin32\n"
        "该功能需要 Windows + Microsoft PowerPoint。"
    )
