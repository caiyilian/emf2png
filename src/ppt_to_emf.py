"""
PPT → EMF 转换模块。

使用 PowerPoint COM 自动化，后台将每页幻灯片导出为 EMF 文件。
仅在 Windows + Microsoft Office 环境下可用。
"""

from __future__ import annotations

import os
import sys
import time
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

    使用 PowerPoint COM 自动化，后台打开 PPT 并将每页幻灯片导出为 EMF。
    导出完成后自动关闭 PowerPoint，不留进程残留。

    Args:
        ppt_path: PPT/PPTX 文件路径（绝对路径或相对路径）
        output_dir: EMF 输出目录
        start: 起始页码 (1-based)
        end: 结束页码 (None 表示全部)
        width: 导出宽度（像素）
        height: 导出高度（像素）
        progress_callback: 可选进度回调 fn(current, total, filename)

    Returns:
        EMF 文件路径列表（按页码顺序）

    Raises:
        FileNotFoundError: PPT 文件不存在
        RuntimeError: Office 未安装、文件损坏、导出失败等
    """
    ppt = Path(ppt_path)
    if not ppt.exists():
        raise FileNotFoundError(f"PPT 文件不存在: {ppt_path}")

    ppt_abs = str(ppt.resolve())

    out = Path(output_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    # 尝试导入 pywin32
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()  # 初始化 COM 线程模型
    except ImportError:
        raise RuntimeError(
            "缺少 pywin32 模块，请执行: uv pip install pywin32\n"
            "该功能需要 Windows + Microsoft PowerPoint。"
        )

    app = None
    pres = None
    emf_files: list[str] = []

    try:
        # 检测 Office 是否安装
        try:
            app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception:
            raise RuntimeError(
                "无法启动 PowerPoint，请确保已安装 Microsoft Office。"
            )

        # 后台运行（某些系统不允许隐藏窗口，所以不强制设置 Visible）
        # 通过 Presentations.Open 的 WithWindow=False 参数保持窗口不显示
        app.DisplayAlerts = False

        # 打开 PPT 文件
        try:
            pres = app.Presentations.Open(ppt_abs, WithWindow=False)
        except Exception as e:
            raise RuntimeError(f"无法打开 PPT 文件: {e}")

        total_slides = pres.Slides.Count
        if total_slides == 0:
            raise RuntimeError("PPT 文件中没有任何幻灯片。")

        # 校验页码范围
        start = max(1, min(start, total_slides))
        if end is None or end > total_slides:
            end = total_slides
        end = max(1, end)

        if start > end:
            start, end = end, start

        selected_range = range(start, end + 1)
        slide_count = len(selected_range)

        print(f"  → PPT 共 {total_slides} 页，导出第 {start}~{end} 页 ({slide_count} 页)")

        # 逐页导出
        errors: list[str] = []
        for idx, slide_num in enumerate(selected_range, start=1):
            slide = pres.Slides(slide_num)
            # 输出文件: slide_001.emf, slide_002.emf, ...
            emf_name = f"slide_{slide_num:03d}.emf"
            emf_path = str(out / emf_name)

            if progress_callback:
                progress_callback(idx, slide_count, f"第{slide_num}页")

            try:
                slide.Export(emf_path, "EMF", width, height)
                # 验证导出文件是否生成
                if not Path(emf_path).exists():
                    raise RuntimeError(f"导出完成但文件未生成: {emf_path}")
                emf_files.append(emf_path)
            except Exception as e:
                err_msg = f"第{slide_num}页导出失败: {e}"
                errors.append(err_msg)
                print(f"  [!] {err_msg}")

    finally:
        # 清理：关闭 PPT，退出 PowerPoint
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    if not emf_files:
        raise RuntimeError("未成功导出任何页面。" + (f" 错误: {'; '.join(errors)}" if errors else ""))

    if errors:
        print(f"  [!] 部分页面导出失败 ({len(errors)}/{slide_count} 页):")
        for e in errors:
            print(f"       {e}")

    return emf_files
