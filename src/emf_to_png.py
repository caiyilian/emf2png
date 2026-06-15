"""
EMF → PNG 转换模块。

调用 emf_to_png.exe（Windows GDI 原生渲染）实现高保真转换，
再通过 Pillow 进行缩放和 DPI 设置。
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PIL import Image


# Pillow 大图限制
Image.MAX_IMAGE_PIXELS = 500_000_000


def get_exe_path() -> str:
    """返回 emf_to_png.exe 的绝对路径。

    支持两种运行模式：
    - 源码运行: 相对于 __file__ 寻找 ../bin/emf_to_png.exe
    - PyInstaller 打包: 相对于 sys._MEIPASS 寻找 bin/emf_to_png.exe
    """
    # PyInstaller 打包模式下，资源文件在 sys._MEIPASS 目录
    base = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(__file__).parent.parent
    exe = base / "bin" / "emf_to_png.exe"
    if not exe.exists():
        raise FileNotFoundError(
            f"找不到 emf_to_png.exe (预期位置: {exe})\n"
            f"请确保该文件存在于 bin/ 目录下"
        )
    return str(exe.resolve())


def _resize_and_set_dpi(
    png_path: str,
    scale: float,
    dpi: int,
) -> str:
    """
    对已生成的 PNG 进行缩放并设置 DPI 元数据。

    Args:
        png_path: 已生成的 PNG 文件路径
        scale: 缩放倍率 (1.0=原大)
        dpi: 输出 DPI (写入 PNG metadata)

    Returns:
        处理后的 PNG 文件路径（原地覆盖）
    """
    img = Image.open(png_path)

    # 缩放
    if scale != 1.0:
        new_size = (int(img.width * scale), int(img.height * scale))
        img = img.resize(new_size, Image.Resampling.LANCZOS)

    # 设置 DPI
    img.save(png_path, dpi=(dpi, dpi))

    return png_path


def convert(
    emf_path: str,
    png_path: str,
    scale: float = 1.0,
    dpi: int = 300,
) -> str:
    """
    将单个 EMF 文件转换为 PNG，可选缩放和 DPI 设置。

    流程: exe 渲染 → Pillow 缩放 (scale≠1) → 写入 DPI 元数据

    Args:
        emf_path: 输入 EMF 文件路径
        png_path: 输出 PNG 文件路径
        scale: 缩放倍率 (1.0=原大, 2.0=2x)
        dpi: 输出 PNG 的 DPI

    Returns:
        生成的 PNG 文件路径

    Raises:
        FileNotFoundError: emf_to_png.exe 或输入 EMF 不存在
        RuntimeError: 转换失败
    """
    emf = Path(emf_path)
    if not emf.exists():
        raise FileNotFoundError(f"输入 EMF 文件不存在: {emf_path}")

    png = Path(png_path)
    png.parent.mkdir(parents=True, exist_ok=True)

    exe = get_exe_path()

    # --------------------------------------------------
    # 步骤 A: exe 渲染 (原生分辨率)
    # --------------------------------------------------
    cmd = [exe, str(emf.resolve()), str(png.resolve())]

    # 不使用 capture_output=True 避免 Windows pipe 死锁
    try:
        result = subprocess.run(cmd, timeout=120)
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"EMF→PNG 转换超时 (120s): {emf.name}\n"
            f"  文件过大或 exe 卡死，请尝试缩小 scale 值"
        )
    except OSError as e:
        raise RuntimeError(
            f"EMF→PNG 转换进程启动失败: {e}\n"
            f"  请检查 emf_to_png.exe 是否可执行"
        )

    if result.returncode != 0:
        raise RuntimeError(
            f"EMF→PNG 转换失败 (exit={result.returncode}): {emf.name}\n"
            f"  请检查 EMF 文件是否损坏"
        )

    if not png.exists():
        raise RuntimeError(f"转换完成但输出文件未生成: {png_path}")

    # --------------------------------------------------
    # 步骤 B: 缩放 + DPI (Pillow 后处理)
    # --------------------------------------------------
    try:
        _resize_and_set_dpi(str(png.resolve()), scale, dpi)
    except OSError as e:
        # 磁盘空间不足 / 权限错误
        raise RuntimeError(
            f"PNG 后处理失败 (磁盘空间不足?): {e}\n"
            f"  请检查磁盘剩余空间"
        )
    except Exception as e:
        raise RuntimeError(f"PNG 后处理失败: {e}")

    return str(png.resolve())


def batch_convert(
    emf_files: list[str],
    output_dir: str,
    scale: float = 1.0,
    dpi: int = 300,
    keep_emf: bool = False,
    progress_callback=None,
) -> list[str]:
    """
    批量将 EMF 文件转换为 PNG。

    Args:
        emf_files: EMF 文件路径列表
        output_dir: PNG 输出目录
        scale: 缩放倍率
        dpi: 输出 DPI
        keep_emf: 是否保留 EMF 源文件（否则删除）
        progress_callback: 可选进度回调 fn(current, total, filename)

    Returns:
        生成的 PNG 文件路径列表
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    png_files: list[str] = []
    total = len(emf_files)

    for i, emf_path in enumerate(emf_files):
        emf = Path(emf_path)
        png_name = emf.stem + ".png"
        png_path = str(out_dir / png_name)

        if progress_callback:
            progress_callback(i + 1, total, emf.name)

        convert(emf_path, png_path, scale, dpi)
        png_files.append(png_path)

        # 不保留 EMF 则删除
        if not keep_emf:
            try:
                emf.unlink()
            except OSError:
                pass

    return png_files
