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
COLOR_BG = "#1a1a2e"      # 深色背景
COLOR_ACCENT = "#4cc9f0"   # 青蓝强调色
COLOR_SUCCESS = "#2ecc71"  # 绿色成功
COLOR_ERROR = "#e74c3c"    # 红色错误


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
        ctk.set_appearance_mode("Dark")  # 默认暗色主题

        # ── 网格布局: 3 行 (顶部标题, 中间参数, 底部按钮) ──
        self.grid_rowconfigure(0, weight=0)  # 顶部标题栏，固定高度
        self.grid_rowconfigure(1, weight=1)  # 中间参数区，可伸缩
        self.grid_rowconfigure(2, weight=0)  # 底部按钮栏，固定高度
        self.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_params_area()
        self._build_footer()

    # ─────────────────────────────────────
    #  顶部：标题栏
    # ─────────────────────────────────────

    def _build_header(self):
        """应用标题 + 简短说明。"""
        frame = ctk.CTkFrame(self, corner_radius=0)
        frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)

        # 应用名称
        title = ctk.CTkLabel(
            frame,
            text="emf2png",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLOR_ACCENT,
        )
        title.pack(pady=(16, 0))

        # 副标题
        subtitle = ctk.CTkLabel(
            frame,
            text="PowerPoint 幻灯片 → 高清 PNG 素材",
            font=ctk.CTkFont(size=13),
            text_color=("gray40", "gray60"),
        )
        subtitle.pack(pady=(2, 12))

    # ─────────────────────────────────────
    #  中间：参数配置区（阶段10-11 填充）
    # ─────────────────────────────────────

    def _build_params_area(self):
        """可滚动的参数配置面板。"""
        # 可滚动容器
        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            label_text=" 转换参数",
            label_font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.scroll_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))

        # ── 占位提示（后续阶段替换为实际组件） ──
        placeholder = ctk.CTkLabel(
            self.scroll_frame,
            text="参数面板将在后续阶段添加",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray40"),
        )
        placeholder.pack(expand=True, pady=80)

    # ─────────────────────────────────────
    #  底部：按钮栏（阶段12-13 填充）
    # ─────────────────────────────────────

    def _build_footer(self):
        """底部按钮区域。"""
        frame = ctk.CTkFrame(self, corner_radius=0)
        frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

        # 占位提示
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
