#!/usr/bin/env python3
"""
emf2png — 图形界面入口。

使用 CustomTkinter 构建的现代桌面 GUI，封装了 CLI 的所有功能。
"""

import sys
from pathlib import Path

# 将 src/ 加入模块搜索路径，使 src.xxx 可直接 import
_src = str(Path(__file__).parent.parent / "src")
if _src not in sys.path:
    sys.path.insert(0, _src)

import customtkinter as ctk
from tkinter import filedialog, messagebox


# ──────────────────────────────────────────────
#  主题与常量
# ──────────────────────────────────────────────

APP_TITLE = "emf2png — PPT 转 PNG 素材提取工具"
APP_ICON = None  # 阶段14 添加图标

WINDOW_WIDTH = 640
WINDOW_HEIGHT = 700
WINDOW_MIN_WIDTH = 500
WINDOW_MIN_HEIGHT = 600

# 颜色
COLOR_ACCENT = "#4cc9f0"   # 青蓝强调色


# ──────────────────────────────────────────────
#  主窗口
# ──────────────────────────────────────────────

class App(ctk.CTk):
    """emf2png 主窗口。"""

    def __init__(self):
        super().__init__()

        # ── 窗口基础设置 ──
        self.title(APP_TITLE)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        ctk.set_appearance_mode("Dark")

        # ── 网格布局: 3 行 ──
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_params_area()
        self._build_footer()

    # ═════════════════════════════════════════
    #  顶部：标题栏
    # ═════════════════════════════════════════

    def _build_header(self):
        frame = ctk.CTkFrame(self, corner_radius=0)
        frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)

        title = ctk.CTkLabel(
            frame,
            text="emf2png",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLOR_ACCENT,
        )
        title.pack(pady=(16, 0))

        subtitle = ctk.CTkLabel(
            frame,
            text="PowerPoint 幻灯片 → 高清 PNG 素材",
            font=ctk.CTkFont(size=13),
            text_color=("gray40", "gray60"),
        )
        subtitle.pack(pady=(2, 12))

    # ═════════════════════════════════════════
    #  中间：参数配置区
    # ═════════════════════════════════════════

    def _build_params_area(self):
        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            label_text=" 转换参数",
            label_font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.scroll_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
        self.scroll_frame.grid_columnconfigure(0, weight=1)

        # ── 输入文件选择 ──
        self._build_file_row()

        # ── 输出目录选择 ──
        self._build_output_dir_row()

        # ── 输出设置分组 ──
        self._build_output_settings()

        # ── 附加选项分组 ──
        self._build_extra_options()

    # ─────────────────────────────────────
    #  输入文件选择
    # ─────────────────────────────────────

    def _build_file_row(self):
        group = ctk.CTkFrame(self.scroll_frame)
        group.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        group.grid_columnconfigure(1, weight=1)

        label = ctk.CTkLabel(
            group,
            text="PPT 文件",
            font=ctk.CTkFont(size=13, weight="bold"),
            width=80,
            anchor="w",
        )
        label.grid(row=0, column=0, padx=(10, 4), pady=10, sticky="w")

        self.entry_ppt = ctk.CTkEntry(
            group,
            placeholder_text="选择或拖入 .ppt / .pptx 文件",
        )
        self.entry_ppt.grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=10)

        self.btn_browse = ctk.CTkButton(
            group,
            text="浏览...",
            width=80,
            command=self._on_browse_ppt,
        )
        self.btn_browse.grid(row=0, column=2, padx=(0, 10), pady=10)

    def _on_browse_ppt(self):
        path = filedialog.askopenfilename(
            title="选择 PowerPoint 文件",
            filetypes=[
                ("PowerPoint 文件", "*.pptx *.ppt"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.entry_ppt.delete(0, "end")
            self.entry_ppt.insert(0, path)
            self._auto_set_output_dir(path)

    # ─────────────────────────────────────
    #  输出目录选择
    # ─────────────────────────────────────

    def _build_output_dir_row(self):
        group = ctk.CTkFrame(self.scroll_frame)
        group.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        group.grid_columnconfigure(1, weight=1)

        label = ctk.CTkLabel(
            group,
            text="输出目录",
            font=ctk.CTkFont(size=13, weight="bold"),
            width=80,
            anchor="w",
        )
        label.grid(row=0, column=0, padx=(10, 4), pady=10, sticky="w")

        self.entry_output = ctk.CTkEntry(
            group,
            placeholder_text="默认: ./output/幻灯片名称/",
        )
        self.entry_output.grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=10)

        self.btn_output = ctk.CTkButton(
            group,
            text="选择目录...",
            width=90,
            command=self._on_browse_output,
        )
        self.btn_output.grid(row=0, column=2, padx=(0, 10), pady=10)

    def _on_browse_output(self):
        path = filedialog.askdirectory(
            title="选择输出目录",
            mustexist=False,
        )
        if path:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, path)

    def _auto_set_output_dir(self, ppt_path: str):
        """选择 PPT 后自动填充输出目录: ./output/<幻灯片名>/"""
        ppt_name = Path(ppt_path).stem
        default_dir = f"./output/{ppt_name}/"
        current = self.entry_output.get().strip()
        if not current:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, default_dir)

    # ═════════════════════════════════════════
    #  输出设置：缩放、DPI、页码范围
    # ═════════════════════════════════════════

    def _build_output_settings(self):
        """缩放倍率、DPI、页码范围。"""
        group = ctk.CTkFrame(self.scroll_frame)
        group.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        group.grid_columnconfigure(1, weight=1)

        # ── 分组标题 ──
        title = ctk.CTkLabel(
            group,
            text="输出设置",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, columnspan=3, sticky="ew", padx=10, pady=(8, 2))

        # ── 缩放倍率（Slider） ──
        label_scale = ctk.CTkLabel(
            group, text="缩放倍率", width=80, anchor="w",
        )
        label_scale.grid(row=1, column=0, padx=(10, 4), pady=6, sticky="w")

        self.scale_var = ctk.DoubleVar(value=2.0)
        self.scale_slider = ctk.CTkSlider(
            group,
            from_=0.5, to=8.0,
            number_of_steps=15,
            variable=self.scale_var,
            command=self._on_scale_change,
        )
        self.scale_slider.grid(row=1, column=1, sticky="ew", padx=(0, 4), pady=6)

        self.scale_label = ctk.CTkLabel(
            group, text="2.0x", width=50, anchor="w",
        )
        self.scale_label.grid(row=1, column=2, padx=(0, 10), pady=6, sticky="w")

        # ── DPI ──
        label_dpi = ctk.CTkLabel(
            group, text="DPI", width=80, anchor="w",
        )
        label_dpi.grid(row=2, column=0, padx=(10, 4), pady=6, sticky="w")

        self.dpi_var = ctk.StringVar(value="300")
        self.dpi_entry = ctk.CTkEntry(
            group,
            textvariable=self.dpi_var,
            width=100,
        )
        self.dpi_entry.grid(row=2, column=1, sticky="w", padx=(0, 4), pady=6)

        # ── 页码范围 ──
        label_range = ctk.CTkLabel(
            group, text="页码范围", width=80, anchor="w",
        )
        label_range.grid(row=3, column=0, padx=(10, 4), pady=6, sticky="w")

        range_frame = ctk.CTkFrame(group, fg_color="transparent")
        range_frame.grid(row=3, column=1, sticky="w", padx=(0, 4), pady=6)
        range_frame.grid_columnconfigure(1, weight=0)

        self.start_var = ctk.StringVar(value="1")
        self.start_entry = ctk.CTkEntry(
            range_frame, textvariable=self.start_var, width=60,
            placeholder_text="起始",
        )
        self.start_entry.grid(row=0, column=0, padx=(0, 4))

        label_to = ctk.CTkLabel(range_frame, text="至")
        label_to.grid(row=0, column=1, padx=2)

        self.end_var = ctk.StringVar(value="")
        self.end_entry = ctk.CTkEntry(
            range_frame, textvariable=self.end_var, width=60,
            placeholder_text="全部",
        )
        self.end_entry.grid(row=0, column=2, padx=(4, 0))

    def _on_scale_change(self, value):
        """滑块拖动时更新显示文字。"""
        self.scale_label.configure(text=f"{value:.1f}x")

    # ═════════════════════════════════════════
    #  附加选项：复选框
    # ═════════════════════════════════════════

    def _build_extra_options(self):
        """裁剪白边、保留 EMF、合并 PDF 复选框。"""
        group = ctk.CTkFrame(self.scroll_frame)
        group.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        title = ctk.CTkLabel(
            group,
            text="附加选项",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, columnspan=3, sticky="ew", padx=10, pady=(8, 2))

        # 复选框行
        cb_frame = ctk.CTkFrame(group, fg_color="transparent")
        cb_frame.grid(row=1, column=0, sticky="w", padx=10, pady=(4, 10))

        self.trim_var = ctk.BooleanVar(value=True)
        cb_trim = ctk.CTkCheckBox(
            cb_frame, text="裁剪白边", variable=self.trim_var,
        )
        cb_trim.grid(row=0, column=0, padx=(0, 16))

        self.keep_emf_var = ctk.BooleanVar(value=False)
        cb_keep = ctk.CTkCheckBox(
            cb_frame, text="保留 EMF", variable=self.keep_emf_var,
        )
        cb_keep.grid(row=0, column=1, padx=(0, 16))

        self.merge_pdf_var = ctk.BooleanVar(value=False)
        cb_pdf = ctk.CTkCheckBox(
            cb_frame, text="合并为 PDF", variable=self.merge_pdf_var,
        )
        cb_pdf.grid(row=0, column=2, padx=(0, 16))

    # ═════════════════════════════════════════
    #  底部：按钮栏（阶段12 填充）
    # ═════════════════════════════════════════

    def _build_footer(self):
        frame = ctk.CTkFrame(self, corner_radius=0)
        frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

        placeholder = ctk.CTkLabel(
            frame,
            text="转换按钮将在后续阶段添加",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray40"),
        )
        placeholder.grid(row=0, column=0, columnspan=2, pady=16)


# ──────────────────────────────────────────────
#  启动
# ──────────────────────────────────────────────

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
