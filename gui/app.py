#!/usr/bin/env python3
"""
emf2png — 图形界面入口。

使用 CustomTkinter 构建的现代桌面 GUI，封装了 CLI 的所有功能。
"""

import sys
import time
import json
import os
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

WINDOW_WIDTH = 640
WINDOW_HEIGHT = 700
WINDOW_MIN_WIDTH = 500
WINDOW_MIN_HEIGHT = 600

# 颜色
COLOR_ACCENT = "#4cc9f0"   # 青蓝强调色

# 配置文件路径 (%LOCALAPPDATA%/emf2png/config.json)
CONFIG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "emf2png"
CONFIG_PATH = CONFIG_DIR / "config.json"

THEMES = ["Dark", "Light", "System"]


# ──────────────────────────────────────────────
#  主窗口
# ──────────────────────────────────────────────

class App(ctk.CTk):
    """emf2png 主窗口。"""

    def __init__(self):
        super().__init__()

        # ── 窗口基础设置 ──
        self.title(APP_TITLE)
        self.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        ctk.set_appearance_mode(self._load_appearance_mode())

        # ── 状态 ──
        self._converting = False
        self._last_output_dir = None
        self._theme_idx = THEMES.index(ctk.get_appearance_mode())

        # ── 恢复窗口位置 ──
        self._load_window_geometry()

        # ── 全局异常钩子 ──
        self.report_callback_exception = self._on_tk_exception

        # ── 关闭确认 ──
        self.protocol("WM_DELETE_WINDOW", self._on_close)

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
        frame.grid_columnconfigure(0, weight=1)

        # 标题
        title_frame = ctk.CTkFrame(frame, fg_color="transparent")
        title_frame.grid(row=0, column=0, pady=(16, 0))

        title = ctk.CTkLabel(
            title_frame,
            text="emf2png",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLOR_ACCENT,
        )
        title.pack(side="left")

        # 主题切换按钮
        self.btn_theme = ctk.CTkButton(
            title_frame,
            text="🌙",
            width=32, height=32,
            corner_radius=16,
            command=self._toggle_theme,
        )
        self.btn_theme.pack(side="left", padx=(10, 0))
        self._update_theme_button_text()

        # 副标题
        subtitle = ctk.CTkLabel(
            frame,
            text="PowerPoint 幻灯片 → 高清 PNG 素材",
            font=ctk.CTkFont(size=13),
            text_color=("gray40", "gray60"),
        )
        subtitle.grid(row=1, column=0, pady=(2, 12))

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
    #  底部：进度条 + 日志 + 转换按钮
    # ═════════════════════════════════════════

    def _build_footer(self):
        frame = ctk.CTkFrame(self, corner_radius=0)
        frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        frame.grid_columnconfigure(0, weight=1)

        # ── 进度条 ──
        self.progress_bar = ctk.CTkProgressBar(frame, height=6)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 0))
        self.progress_bar.set(0)

        # ── 日志区域 ──
        self.log_text = ctk.CTkTextbox(
            frame, height=120, state="disabled",
            font=ctk.CTkFont(family="Consolas", size=11),
        )
        self.log_text.grid(row=1, column=0, sticky="ew", padx=12, pady=(4, 6))

        # ── 按钮行 ──
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=(0, 8))

        self.btn_open_output = ctk.CTkButton(
            btn_frame,
            text="打开输出目录",
            width=140,
            command=self._on_open_output,
        )
        self.btn_open_output.pack(side="left", padx=(0, 8))
        self.btn_open_output.pack_forget()  # 默认隐藏，转换完成后显示

        self.btn_convert = ctk.CTkButton(
            btn_frame,
            text="开始转换",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=36,
            width=180,
            command=self._on_start_convert,
        )
        self.btn_convert.pack(side="left")

    # ═════════════════════════════════════════
    #  转换执行
    # ═════════════════════════════════════════

    def _on_start_convert(self):
        """读取参数，启动后台转换线程。"""
        import threading

        # 读取所有参数
        ppt_path = self.entry_ppt.get().strip()
        if not ppt_path or not Path(ppt_path).exists():
            self._log("[ERR] 请先选择有效的 PPT 文件")
            return

        output_dir = self.entry_output.get().strip() or f"./output/{Path(ppt_path).stem}/"
        scale = self.scale_var.get()
        try:
            dpi = int(self.dpi_var.get())
        except ValueError:
            dpi = 300
        try:
            start = int(self.start_var.get())
        except ValueError:
            start = 1
        end_str = self.end_var.get().strip()
        end = int(end_str) if end_str else None
        trim = self.trim_var.get()
        keep_emf = self.keep_emf_var.get()
        merge_pdf = self.merge_pdf_var.get()

        # 状态
        self._converting = True
        self._convert_start_time = time.time()
        self._last_output_dir = output_dir

        # 禁用 UI
        self._set_ui_enabled(False)
        self.progress_bar.set(0)
        self._clear_log()
        self.btn_open_output.pack_forget()  # 隐藏"打开目录"按钮

        # 在后台线程运行
        args = (ppt_path, output_dir, scale, dpi, start, end, trim, keep_emf, merge_pdf)
        t = threading.Thread(target=self._run_conversion, args=args, daemon=True)
        t.start()

    def _log(self, msg: str):
        """向日志区域追加一行（线程安全）。"""
        self.after(0, self._append_log, msg)

    def _append_log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("0.0", "end")
        self.log_text.configure(state="disabled")

    def _set_ui_enabled(self, enabled: bool):
        """启用/禁用所有交互控件。"""
        state = "normal" if enabled else "disabled"
        self.btn_browse.configure(state=state)
        self.btn_output.configure(state=state)
        self.btn_convert.configure(
            state=state,
            text="开始转换" if enabled else "转换中...",
        )
        self.entry_ppt.configure(state=state)
        self.entry_output.configure(state=state)
        self.dpi_entry.configure(state=state)
        self.start_entry.configure(state=state)
        self.end_entry.configure(state=state)
        self.scale_slider.configure(state=state)

    def _make_progress_callback(self):
        """创建进度回调闭包，更新进度条和日志。"""
        step_map = {
            1: "PPT 导出为 EMF",
            2: "EMF 转换为 PNG",
            3: "裁剪白边",
            4: "合并 PDF",
        }
        current_step = [0]  # mutable closure

        def on_step(step_num: int, message: str):
            current_step[0] = step_num
            total_steps = 4
            base_progress = (step_num - 1) / total_steps
            self.progress_bar.set(base_progress)
            self._log(f"\n[{step_num}/{total_steps}] {message}")

        def on_progress(current: int, total: int, filename: str = ""):
            if total > 0:
                step = current_step[0]
                total_steps = 4
                step_progress = current / total / total_steps
                overall = (step - 1) / total_steps + step_progress
                self.progress_bar.set(min(overall, 1.0))
            file_msg = f"  [{current}/{total}] {filename}" if filename else f"  [{current}/{total}]"
            self._log(file_msg)

        return on_step, on_progress

    def _run_conversion(
        self,
        ppt_path: str,
        output_dir: str,
        scale: float,
        dpi: int,
        start: int,
        end: int | None,
        trim: bool,
        keep_emf: bool,
        merge_pdf: bool,
    ):
        """在后台线程中执行转换流程。"""
        on_step, on_progress = self._make_progress_callback()

        try:
            # 确保输出目录
            Path(output_dir).mkdir(parents=True, exist_ok=True)

            # ── 步骤1: PPT → EMF ──
            on_step(1, "PPT 导出为 EMF")
            from src.ppt_to_emf import ppt_to_emf
            emf_files = ppt_to_emf(
                ppt_path=ppt_path,
                output_dir=output_dir,
                start=start,
                end=end,
                progress_callback=on_progress,
            )
            self._log(f"  -> 生成 {len(emf_files)} 个 EMF 文件")

            # ── 步骤2: EMF → PNG ──
            on_step(2, "EMF 转换为 PNG")
            from src.emf_to_png import batch_convert
            png_files = batch_convert(
                emf_files=emf_files,
                output_dir=output_dir,
                scale=scale,
                dpi=dpi,
                keep_emf=keep_emf,
                progress_callback=on_progress,
            )
            self._log(f"  -> 生成 {len(png_files)} 个 PNG 文件")

            # ── 步骤3: 裁剪白边 ──
            if trim:
                on_step(3, "裁剪纯白边")
                from src.trim_whitespace import batch_trim_white_borders
                batch_trim_white_borders(
                    png_files,
                    progress_callback=on_progress,
                )
                self._log(f"  -> 白边裁剪完成")

            # ── 步骤4: 合并 PDF ──
            if merge_pdf:
                on_step(4, "合并为 PDF")
                from src.merge_pdf import images_to_pdf
                pdf_path = str(Path(output_dir) / "output.pdf")
                images_to_pdf(png_files, pdf_path)
                self._log(f"  -> PDF 已生成: {pdf_path}")

            # ── 完成 ──
            self.after(0, self._on_convert_done, png_files, output_dir)

        except Exception as e:
            self.after(0, self._on_convert_error, str(e))

    def _on_convert_done(self, png_files: list, output_dir: str):
        """转换完成：显示汇总，恢复 UI，显示"打开目录"按钮。"""
        elapsed = time.time() - self._convert_start_time
        total_size = sum(Path(f).stat().st_size for f in png_files)
        self.progress_bar.set(1.0)
        self._log(f"\n{'='*40}")
        self._log(f"[OK] 完成! 共生成 {len(png_files)} 个 PNG 文件")
        self._log(f"    耗时: {elapsed:.1f}s")
        self._log(f"    总大小: {total_size / 1024:.0f} KB")
        self._log(f"    输出目录: {Path(output_dir).resolve()}")
        self._log(f"{'='*40}")
        self._converting = False
        self._set_ui_enabled(True)
        self.btn_convert.configure(text="再次转换")
        self.btn_open_output.pack(side="left", padx=(0, 8), before=self.btn_convert)

    def _on_convert_error(self, msg: str):
        """转换出错：显示错误弹窗，恢复 UI。"""
        self._converting = False
        self._log(f"\n[ERR] {msg}")
        self._set_ui_enabled(True)
        self.btn_convert.configure(text="重新转换")
        self.after(100, self._show_error_dialog, msg)

    def _show_error_dialog(self, msg: str):
        """弹出错误对话框。"""
        messagebox.showerror(
            title="转换失败",
            message=f"转换过程中出现错误:\n\n{msg}\n\n请检查参数或文件后重试。",
        )

    def _on_open_output(self):
        """在资源管理器中打开输出目录。"""
        if self._last_output_dir:
            import os
            path = Path(self._last_output_dir).resolve()
            if path.exists():
                os.startfile(str(path))

    def _on_close(self):
        """关闭窗口确认 + 保存窗口位置。"""
        if self._converting:
            result = messagebox.askyesno(
                title="确认退出",
                message="正在转换中，确定要退出吗？\n\n未完成的转换将被中断。",
            )
            if not result:
                return
        self._save_window_geometry()
        self._save_appearance_mode()
        self.destroy()

    # ═════════════════════════════════════════
    #  主题切换
    # ═════════════════════════════════════════

    def _toggle_theme(self):
        """在 Dark / Light / System 间循环切换。"""
        self._theme_idx = (self._theme_idx + 1) % len(THEMES)
        mode = THEMES[self._theme_idx]
        ctk.set_appearance_mode(mode)
        self._update_theme_button_text()

    def _update_theme_button_text(self):
        """更新主题按钮的文字。"""
        icons = {"Dark": "Dark", "Light": "Light", "System": "Auto"}
        current = ctk.get_appearance_mode()
        self.btn_theme.configure(text=icons.get(current, "Dark"))

    # ═════════════════════════════════════════
    #  窗口位置记忆
    # ═════════════════════════════════════════

    def _load_window_geometry(self):
        """从配置文件恢复窗口位置和大小。"""
        try:
            if CONFIG_PATH.exists():
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                geo = data.get("window_geometry", "")
                if geo:
                    self.geometry(geo)
                    return
        except Exception:
            pass
        # 默认尺寸
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

    def _save_window_geometry(self):
        """保存窗口位置和大小到配置文件。"""
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            data = {"window_geometry": self.geometry()}
            if CONFIG_PATH.exists():
                existing = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                existing.update(data)
                data = existing
            CONFIG_PATH.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    # ═════════════════════════════════════════
    #  主题记忆
    # ═════════════════════════════════════════

    def _load_appearance_mode(self) -> str:
        """从配置文件恢复主题。"""
        try:
            if CONFIG_PATH.exists():
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                return data.get("appearance_mode", "Dark")
        except Exception:
            pass
        return "Dark"

    def _save_appearance_mode(self):
        """保存当前主题到配置文件。"""
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            data = {"appearance_mode": ctk.get_appearance_mode()}
            if CONFIG_PATH.exists():
                existing = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                existing.update(data)
                data = existing
            CONFIG_PATH.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    # ═════════════════════════════════════════
    #  全局异常处理
    # ═════════════════════════════════════════

    @staticmethod
    def _on_tk_exception(exc_type, exc_value, exc_traceback):
        """捕获 tkinter 回调中的未处理异常，弹窗提示而非崩溃。"""
        import traceback
        err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        try:
            messagebox.showerror(
                title="意外错误",
                message=f"程序遇到意外错误:\n\n{exc_value}\n\n"
                        f"详细信息已打印到控制台。",
            )
        except Exception:
            pass
        # 仍然打印到控制台供调试
        print(err_msg, file=sys.stderr)


# ──────────────────────────────────────────────
#  启动
# ──────────────────────────────────────────────

def main():
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        import traceback
        err_msg = "".join(traceback.format_exc())
        try:
            messagebox.showerror(
                title="启动失败",
                message=f"程序启动时出现错误:\n\n{e}",
            )
        except Exception:
            pass
        print(err_msg, file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
