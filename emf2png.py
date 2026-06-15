#!/usr/bin/env python3
"""
emf2png — 将 PowerPoint 幻灯片一键转换为高清 PNG 图片。

工作流: PPT → EMF → PNG (→ 裁剪白边 → 合并 PDF)

用法:
    python emf2png.py 产品介绍.pptx
    python emf2png.py 产品介绍.pptx --trim -s 4
    python emf2png.py 产品介绍.pptx --start 3 --end 10 --merge-pdf
    emf2png.exe 产品介绍.pptx -o ./output
"""

import sys
from pathlib import Path

# 将 src/ 加入模块搜索路径，使 src.xxx 可直接 import
# 源码模式: 相对于 __file__ 的 src/
# PyInstaller 模式: 相对于 sys._MEIPASS 的 src/
if hasattr(sys, "_MEIPASS"):
    _base = Path(sys._MEIPASS)
else:
    _base = Path(__file__).parent
_src = str(_base / "src")
if _src not in sys.path:
    sys.path.insert(0, _src)


import argparse
import signal
import time
from functools import partial

from tqdm import tqdm


# ============================================================
#  全局清理
# ============================================================

_cleanup_actions: list[callable] = []


def register_cleanup(fn: callable):
    """注册清理函数，在退出或异常时执行。"""
    _cleanup_actions.append(fn)


def _do_cleanup():
    """执行所有注册的清理动作。"""
    for fn in reversed(_cleanup_actions):
        try:
            fn()
        except Exception:
            pass


def _handle_exception(exc_type, exc_value, exc_traceback):
    """全局异常钩子：退出前执行清理。"""
    _do_cleanup()
    if exc_type is KeyboardInterrupt:
        tqdm.write("\n[!] 用户中断，正在清理...")
        sys.exit(130)
    else:
        # 非 SystemExit 的异常才调用默认 excepthook
        sys.__excepthook__(exc_type, exc_value, exc_traceback)


# 注册异常钩子
sys.excepthook = _handle_exception


# ============================================================
#  进度与日志工具
# ============================================================

def make_progress_bar(total: int, desc: str, unit: str = "项"):
    """
    创建 tqdm 进度条，返回 (update_callback, bar) 元组。

    用法:
        cb, bar = make_progress_bar(len(files), "转换中")
        for f in files:
            cb(1, len(files), f.name)
        bar.close()
    """
    if total <= 0:
        total = 1
    bar = tqdm(
        total=total,
        desc=desc,
        unit=unit,
        leave=False,
        bar_format="{desc}: {n}/{total} [{bar:20}] {percentage:3.0f}% {unit}",
    )

    def callback(current: int, total: int, filename: str = ""):
        bar.update(1)

    return callback, bar


def print_simple_progress(current: int, total: int, filename: str = ""):
    """简单的文本进度回调（用于 PPT→EMF 步骤，因为总数是运行时才知道）。"""
    pass


def print_step(step_label: str):
    """打印步骤标题分隔线。"""
    tqdm.write(f"\n[{step_label}]")


def print_summary(png_files: list[str], elapsed: float, output_dir: str):
    """打印完成汇总。"""
    total_size = sum(Path(f).stat().st_size for f in png_files)
    tqdm.write(f"\n{'='*50}")
    tqdm.write(f"[OK] 完成! 共生成 {len(png_files)} 个 PNG 文件")
    tqdm.write(f"   耗时: {elapsed:.1f}s")
    tqdm.write(f"   总大小: {total_size / 1024:.0f} KB")
    tqdm.write(f"   输出目录: {Path(output_dir).resolve()}")
    tqdm.write(f"{'='*50}")


# ============================================================
#  参数解析
# ============================================================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="将 PowerPoint / EMF 文件转换为 PNG 图片 (PPT → EMF → PNG)",
        epilog="示例:\n"
               "  %(prog)s 产品介绍.pptx\n"
               "  %(prog)s 产品介绍.pptx --trim -s 4\n"
               "  %(prog)s 图表.emf --trim -s 4\n"
               "  %(prog)s 产品介绍.pptx --start 3 --end 10 --merge-pdf\n"
               "  %(prog)s 产品介绍.pptx --keep-emf -o ./output",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "input",
        help="输入的 .ppt / .pptx / .emf 文件路径",
    )
    parser.add_argument(
        "-o", "--output",
        default="./output",
        help="输出目录 (默认: ./output)",
    )
    parser.add_argument(
        "-s", "--scale",
        type=float,
        default=2.0,
        help="PNG 缩放倍率，值越大越清晰 (默认: 2.0)",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="输出 PNG 的 DPI (默认: 300)",
    )
    parser.add_argument(
        "--keep-emf",
        action="store_true",
        help="保留中间 EMF 文件，默认转换后自动删除",
    )
    parser.add_argument(
        "--merge-pdf",
        action="store_true",
        help="将所有 PNG 合并为一个 PDF 文件",
    )
    parser.add_argument(
        "--trim",
        action="store_true",
        help="裁剪纯白边 (#FFFFFF)，非白底背景自动跳过不处理",
    )
    parser.add_argument(
        "--start",
        type=int,
        default=1,
        help="起始页码，从 1 开始 (默认: 1)",
    )
    parser.add_argument(
        "--end",
        type=int,
        default=None,
        help="结束页码，包含该页 (默认: 全部)",
    )

    return parser


# ============================================================
#  主流程
# ============================================================

def main():
    try:
        _main()
    except FileNotFoundError as e:
        tqdm.write(f"\n[ERR] 文件错误: {e}")
        _do_cleanup()
        sys.exit(1)
    except RuntimeError as e:
        tqdm.write(f"\n[ERR] {e}")
        _do_cleanup()
        sys.exit(1)
    except Exception:
        _do_cleanup()
        raise


def _main():
    parser = build_parser()
    args = parser.parse_args()

    # 校验输入
    input_path = Path(args.input)
    if not input_path.exists():
        tqdm.write(f"\n[ERR] 文件不存在: {input_path}")
        tqdm.write(f"       请检查路径是否正确")
        sys.exit(1)

    suffix = input_path.suffix.lower()
    if suffix not in (".ppt", ".pptx", ".emf"):
        tqdm.write(f"\n[ERR] 不支持的文件格式: {suffix}")
        tqdm.write(f"       仅支持 .ppt、.pptx 和 .emf 文件")
        sys.exit(1)
    if suffix == ".ppt":
        tqdm.write(f"[!] 提示: .ppt 是旧版格式，建议另存为 .pptx 以获得更好的兼容性")

    # 确保输出目录存在
    try:
        output_dir = Path(args.output)
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        tqdm.write(f"\n[ERR] 无法创建输出目录: {output_dir}")
        tqdm.write(f"       {e}")
        sys.exit(1)

    start_time = time.time()

    # ------------------------------------------------
    # 步骤1: PPT → EMF（仅对 PPT 输入）
    # ------------------------------------------------
    if suffix == ".emf":
        # EMF 直接输入，跳过步骤1
        emf_files = [str(input_path.resolve())]
        tqdm.write(f"  -> 直接处理 EMF 文件: {input_path.name}")
    else:
        print_step("1/4  PPT 导出为 EMF")
        from ppt_to_emf import ppt_to_emf

        emf_files = ppt_to_emf(
            ppt_path=str(input_path.resolve()),
            output_dir=str(output_dir),
            start=args.start,
            end=args.end,
            progress_callback=print_simple_progress,
        )
        tqdm.write(f"  -> 生成 {len(emf_files)} 个 EMF 文件")

    # ------------------------------------------------
    # 步骤2: EMF → PNG
    # ------------------------------------------------
    print_step("2/4  EMF 转换为 PNG")
    from emf_to_png import batch_convert

    cb2, bar2 = make_progress_bar(len(emf_files), "  EMF->PNG", "页")
    png_files = batch_convert(
        emf_files=emf_files,
        output_dir=str(output_dir),
        scale=args.scale,
        dpi=args.dpi,
        keep_emf=args.keep_emf,
        progress_callback=cb2,
    )
    bar2.close()
    tqdm.write(f"  -> 生成 {len(png_files)} 个 PNG 文件")

    # ------------------------------------------------
    # 步骤 2b: 裁剪白边
    # ------------------------------------------------
    if args.trim:
        print_step("3/4  裁剪纯白边")
        from trim_whitespace import batch_trim_white_borders

        cb3, bar3 = make_progress_bar(len(png_files), "  裁剪白边", "张")
        png_files = batch_trim_white_borders(
            png_files,
            progress_callback=cb3,
        )
        bar3.close()
        tqdm.write(f"  -> 白边裁剪完成")

    # ------------------------------------------------
    # 步骤3: 合并 PDF
    # ------------------------------------------------
    if args.merge_pdf:
        print_step("4/4  合并为 PDF")
        from merge_pdf import images_to_pdf

        pdf_path = str(output_dir / "output.pdf")
        images_to_pdf(png_files, pdf_path)
        tqdm.write(f"  -> PDF 已生成: {pdf_path}")

    # ------------------------------------------------
    # 完成
    # ------------------------------------------------
    elapsed = time.time() - start_time
    print_summary(png_files, elapsed, str(output_dir))


if __name__ == "__main__":
    main()
