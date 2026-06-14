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
_src = str(Path(__file__).parent / "src")
if _src not in sys.path:
    sys.path.insert(0, _src)


import argparse
import time


# ============================================================
#  进度与日志工具
# ============================================================

def print_step(step: str, message: str):
    """打印带编号的步骤信息。"""
    print(f"\n[{step}] {message}")


def print_progress(current: int, total: int, filename: str):
    """进度回调，用于 tqdm 集成（阶段7 将替换为真实进度条）。"""
    print(f"  [{current}/{total}] {filename}")


def print_summary(png_files: list[str], elapsed: float, output_dir: str):
    """打印完成汇总。"""
    total_size = sum(Path(f).stat().st_size for f in png_files)
    print(f"\n{'='*50}")
    print(f"[OK] 完成! 共生成 {len(png_files)} 个 PNG 文件")
    print(f"   耗时: {elapsed:.1f}s")
    print(f"   总大小: {total_size / 1024:.0f} KB")
    print(f"   输出目录: {Path(output_dir).resolve()}")
    print(f"{'='*50}")


# ============================================================
#  参数解析
# ============================================================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="将 PowerPoint 文件转换为 PNG 图片 (PPT → EMF → PNG)",
        epilog="示例:\n"
               "  %(prog)s 产品介绍.pptx\n"
               "  %(prog)s 产品介绍.pptx --trim -s 4\n"
               "  %(prog)s 产品介绍.pptx --start 3 --end 10 --merge-pdf\n"
               "  %(prog)s 产品介绍.pptx --keep-emf -o ./output",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "input",
        help="输入的 .ppt 或 .pptx 文件路径",
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
    parser = build_parser()
    args = parser.parse_args()

    # 校验输入
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERR] 错误: 文件不存在 — {input_path}")
        sys.exit(1)
    if input_path.suffix.lower() not in (".ppt", ".pptx"):
        print(f"[ERR] 错误: 不支持的文件格式 — {input_path.suffix} (仅支持 .ppt / .pptx)")
        sys.exit(1)

    # 确保输出目录存在
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    start_time = time.time()

    # ------------------------------------------------
    # 步骤1: PPT → EMF
    # ------------------------------------------------
    print_step("1/3", f"正在将 PPT 导出为 EMF: {input_path.name}")
    from ppt_to_emf import ppt_to_emf

    emf_files = ppt_to_emf(
        ppt_path=str(input_path.resolve()),
        output_dir=str(output_dir),
        start=args.start,
        end=args.end,
        progress_callback=print_progress,
    )
    print(f"  → 生成 {len(emf_files)} 个 EMF 文件")

    # ------------------------------------------------
    # 步骤2: EMF → PNG
    # ------------------------------------------------
    print_step("2/3", f"正在将 EMF 转换为 PNG (scale={args.scale})...")
    from emf_to_png import batch_convert

    png_files = batch_convert(
        emf_files=emf_files,
        output_dir=str(output_dir),
        scale=args.scale,
        dpi=args.dpi,
        keep_emf=args.keep_emf,
        progress_callback=print_progress,
    )
    print(f"  → 生成 {len(png_files)} 个 PNG 文件")

    # ------------------------------------------------
    # 步骤 2.5: 裁剪白边
    # ------------------------------------------------
    if args.trim:
        print_step("2.5/3", "正在裁剪纯白边...")
        from trim_whitespace import batch_trim_white_borders
        png_files = batch_trim_white_borders(png_files)
        print(f"  → 白边裁剪完成")

    # ------------------------------------------------
    # 步骤3: 合并 PDF
    # ------------------------------------------------
    if args.merge_pdf:
        print_step("3/3", "正在合并为 PDF...")
        try:
            from merge_pdf import images_to_pdf
            pdf_path = str(output_dir / "output.pdf")
            images_to_pdf(png_files, pdf_path)
            print(f"  → PDF 已生成: {pdf_path}")
        except NotImplementedError as e:
            print(f"  [!] {e}")

    # ------------------------------------------------
    # 完成
    # ------------------------------------------------
    elapsed = time.time() - start_time
    print_summary(png_files, elapsed, str(output_dir))


if __name__ == "__main__":
    main()
