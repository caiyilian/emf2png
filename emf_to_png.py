"""
EMF → PNG 转换模块。

调用 emf_to_png.exe（Windows GDI 原生渲染）实现高保真转换。
exe 来自参考项目 emf-pro-converter_with_exe。
"""

import subprocess
import sys
from pathlib import Path


def get_exe_path() -> str:
    """返回 emf_to_png.exe 的绝对路径。"""
    exe = Path(__file__).parent / "emf_to_png.exe"
    if not exe.exists():
        raise FileNotFoundError(
            f"找不到 emf_to_png.exe，请确保该文件存在于: {exe}"
        )
    return str(exe.resolve())


def convert(emf_path: str, png_path: str, scale: float = 1.0) -> str:
    """
    将单个 EMF 文件转换为 PNG。

    Args:
        emf_path: 输入 EMF 文件路径
        png_path: 输出 PNG 文件路径
        scale: 缩放倍率（传递给 exe）

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

    # 构建命令行参数
    # emf_to_png.exe <input_emf> [output_png]
    cmd = [exe, str(emf.resolve()), str(png.resolve())]

    # 注意: 不使用 capture_output=True，因为 emf_to_png.exe
    # 输出内容较多（PIL DecompressionBombWarning 等），
    # 用管道捕获容易在 Windows 下造成缓冲区死锁。
    # 直接让子进程的输出流向父进程的控制台。
    result = subprocess.run(
        cmd,
        timeout=60,
    )

    if result.returncode != 0:
        raise RuntimeError(f"EMF→PNG 转换失败 (exit={result.returncode})")

    if not png.exists():
        raise RuntimeError(f"转换完成但输出文件未生成: {png_path}")

    return str(png.resolve())


def batch_convert(
    emf_files: list[str],
    output_dir: str,
    scale: float = 1.0,
    keep_emf: bool = False,
    progress_callback=None,
) -> list[str]:
    """
    批量将 EMF 文件转换为 PNG。

    Args:
        emf_files: EMF 文件路径列表
        output_dir: PNG 输出目录
        scale: 缩放倍率
        keep_emf: 是否保留 EMF 源文件（否则删除）
        progress_callback: 可选进度回调 fn(current, total, filename)

    Returns:
        生成的 PNG 文件路径列表
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    png_files = []
    total = len(emf_files)

    for i, emf_path in enumerate(emf_files):
        emf = Path(emf_path)
        png_name = emf.stem + ".png"
        png_path = str(out_dir / png_name)

        if progress_callback:
            progress_callback(i + 1, total, emf.name)

        convert(emf_path, png_path, scale)
        png_files.append(png_path)

        # 不保留 EMF 则删除
        if not keep_emf:
            try:
                emf.unlink()
            except OSError:
                pass

    return png_files
