#!/usr/bin/env python3
"""
convert_to_png.py — 右键菜单调用：将 EMF/PPT 转为 PNG（裁剪白边）。

用法:
    uv run python convert_to_png.py <文件路径>

EMF → PNG:  1x, DPI 300, 裁剪白边（先用临时名，完成后再重命名）
PPT → PNG:  1x, DPI 100, 多页加后缀 _001, _002...
"""

import sys
import time
import uuid
from pathlib import Path

_root = Path(__file__).parent.resolve()
_src = str(_root / "src")
if _src not in sys.path:
    sys.path.insert(0, _src)
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))


def convert_emf_to_png(emf_path: str) -> list[str]:
    """EMF → PNG（裁剪白边，1x，300DPI）。

    先用临时文件名转换+裁剪，最后 rename 到目标名，
    避免用户看到未裁剪的中间文件。
    """
    from emf_to_png import convert
    from trim_whitespace import trim_white_borders

    out = Path(emf_path).with_suffix(".png")
    if out.exists() and not _confirm_overwrite(str(out)):
        return []
    # 临时后缀，转换完成再改名
    tmp = out.with_suffix(".tmp.png")
    convert(emf_path, str(tmp), scale=1.0, dpi=300)
    trim_white_borders(str(tmp))
    if out.exists():
        out.unlink()
    tmp.rename(out)
    return [str(out)]


def convert_ppt_to_png(ppt_path: str) -> list[str]:
    """PPT → PNG（裁剪白边，1x，100DPI，多页加后缀）。

    中间 EMF/PNG 文件使用唯一临时前缀，避免与已存在的文件冲突。
    """
    from ppt_to_emf import ppt_to_emf
    from emf_to_png import batch_convert
    from trim_whitespace import batch_trim_white_borders

    ppt = Path(ppt_path)
    out_dir = ppt.parent

    # 先检查目标文件是否存在
    single_out = out_dir / f"{ppt.stem}.png"
    if single_out.exists() and not _confirm_overwrite(str(single_out)):
        return []

    # 用唯一前缀避免与现有文件冲突
    tag = uuid.uuid4().hex[:8]

    # PPT → EMF（用临时前缀）
    from ppt_to_emf import ppt_to_emf
    # 修改幻灯片导出时的文件命名，不能用前缀参数，只能用临时目录
    # 改为导出到临时目录再移回
    import tempfile
    tmp_dir = Path(tempfile.mkdtemp())

    emf_files = ppt_to_emf(
        ppt_path=str(ppt.resolve()),
        output_dir=str(tmp_dir),
        start=1,
        end=None,
    )

    # EMF → PNG（仍在临时目录）
    png_files = batch_convert(
        emf_files=emf_files,
        output_dir=str(tmp_dir),
        scale=1.0,
        dpi=100,
        keep_emf=True,
    )

    # 裁剪白边（临时目录）
    batch_trim_white_borders(png_files)

    # 移到输出目录并重命名
    results = []
    import shutil
    for i, png_path in enumerate(png_files, start=1):
        src = Path(png_path)
        if len(png_files) == 1:
            dst_name = f"{ppt.stem}.png"
        else:
            dst_name = f"{ppt.stem}_{i:03d}.png"
        dst = out_dir / dst_name
        shutil.move(str(src), str(dst))
        results.append(str(dst))

    # 清理临时目录
    shutil.rmtree(tmp_dir, ignore_errors=True)

    return results


def main():
    if len(sys.argv) < 2:
        _show_message("用法", "python convert_to_png.py <文件路径>")
        sys.exit(1)

    path = sys.argv[1]
    suffix = Path(path).suffix.lower()
    if suffix not in (".emf", ".ppt", ".pptx", ".pptm"):
        _show_message("错误", f"不支持的文件格式: {suffix}")
        sys.exit(1)

    # 弹出一个进度窗口（类似 PPT 的"正在发布"）
    import tkinter as tk
    from tkinter import ttk

    win = tk.Tk()
    win.title("正在转换...")
    win.geometry("360x100")
    win.resizable(False, False)
    # 居中
    win.update_idletasks()
    x = (win.winfo_screenwidth() - 360) // 2
    y = (win.winfo_screenheight() - 100) // 2
    win.geometry(f"+{x}+{y}")

    tk.Label(win, text=f"正在转换: {Path(path).name}", font=("", 10)).pack(pady=(10, 4))
    progress = ttk.Progressbar(win, length=320, mode="indeterminate")
    progress.pack(pady=4)
    progress.start()
    tk.Label(win, text="请稍候...", font=("", 9), fg="gray").pack()

    # 在后台线程执行转换
    results = []
    error = [None]

    def run():
        nonlocal results
        try:
            start = time.time()
            if suffix == ".emf":
                results = convert_emf_to_png(path)
            else:
                results = convert_ppt_to_png(path)
            elapsed = time.time() - start
            win.after(0, _on_done, results, elapsed)
        except Exception as e:
            error[0] = e
            win.after(0, _on_error, str(e))

    def _on_done(res, elapsed):
        progress.stop()
        win.destroy()

    def _on_error(msg):
        progress.stop()
        win.destroy()
        _show_message("转换失败", str(msg))

    import threading
    t = threading.Thread(target=run, daemon=True)
    t.start()
    win.mainloop()


def _show_message(title: str, msg: str):
    """弹出提示对话框（兼容 pythonw.exe 无控制台模式）。"""
    try:
        import tkinter
        from tkinter import messagebox
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(title, msg)
        root.destroy()
    except Exception:
        pass


def _confirm_overwrite(path: str) -> bool:
    """询问用户是否覆盖已有文件。返回 True=覆盖，False=跳过。"""
    try:
        import tkinter
        from tkinter import messagebox
        root = tkinter.Tk()
        root.withdraw()
        result = messagebox.askyesno(
            "文件已存在",
            f"文件已存在:\n{path}\n\n是否覆盖？",
        )
        root.destroy()
        return result
    except Exception:
        return True  # 弹窗失败则默认覆盖


if __name__ == "__main__":
    main()
