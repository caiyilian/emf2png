# PPT → EMF → PNG 转换工具 · 详细开发方案

> 基于 `emf-pro-converter_with_exe` 参考项目（EMF→PNG Web 版）设计，去掉 Web 层，改为桌面/命令行工具，并加入 PPT→EMF 链路。

---

## 一、整体链路

```
.ppt / .pptx
      │
      ▼  [步骤1]  PowerPoint COM 自动化
      │  将每页幻灯片导出为 .emf
      │
   .emf 文件 (每页一个)
      │
      ▼  [步骤2]  emf_to_png.exe
      │  将每个 .emf 转换为 .png
      │
   .png 文件 (每页一个, 含白边/白底)
      │
      ▼  [步骤2.5]  裁剪纯白边 (--trim)
      │  检测图片四周边界，若为纯白 (#FFFFFF) 则裁掉
      │  若幻灯片背景非白色则跳过，保持原样
      │
   .png 文件 (裁剪后, 内容紧凑)
      │
      ▼  [步骤3]  可选：合并 PDF / 批量重命名
      │
   最终输出
```

---

## 二、技术选型

### 2.1 PPT → EMF

| 方案 | 优点 | 缺点 | 推荐度 |
|------|------|------|--------|
| **PowerPoint COM (win32com)** | 保真度最高，支持动画/图表/公式 | 仅 Windows，需安装 Office | ⭐⭐⭐ 首选 |
| python-pptx + 虚拟打印 | 跨平台 | 复杂布局易失真 | ⭐⭐ 备选 |
| LibreOffice UNO API | 免费跨平台 | 安装大，EMF 质量一般 | ⭐ 备选 |

### 2.2 EMF → PNG

直接复用 `emf_to_png.exe`（来自参考项目，PyInstaller 打包），Windows GDI 原生渲染。

### 2.3 裁剪白边

Python + Pillow + numpy，四边扫描法。纯白 `#FFFFFF` 严格匹配，非白底自动跳过。

### 2.4 GUI 框架

| 方案 | 优点 | 缺点 | 推荐度 |
|------|------|------|--------|
| **CustomTkinter** | 现代外观（圆角/暗色主题），API 简洁，轻量 | 仅桌面，无高级组件 | ⭐⭐⭐ 首选 |
| PySide6 | 最成熟，有 Designer 可视化设计 | 体积大 (~100MB)，学习曲线陡 | ⭐⭐ 备选 |
| Tkinter 标准库 | 无需安装 | 外观老旧，代码繁琐，不好用 | ⭐ 不推荐 |
| Flet | 现代 Flutter 风格，可打包 Web | 太重，API 变化快 | ⭐ 不推荐 |

**推荐方案：CustomTkinter**

```bash
uv pip install customtkinter
```

优势：
- 一行代码实现暗色/亮色主题切换
- 组件自带圆角、现代样式（CTkButton, CTkEntry, CTkSlider 等）
- API 与 Tkinter 相似，迁移成本低
- 单文件即可启动，适合快速迭代

---

## 三、环境与架构

### 3.0 使用 uv（推荐）

项目使用 [uv](https://docs.astral.sh/uv/) 管理虚拟环境和依赖（Python 3.12），已配置好 `.venv/` 在项目根目录：

```bash
# 安装依赖（首次）
uv pip install -r requirements.txt

# 运行
uv run emf2png.py 产品介绍.pptx --trim -s 4

# 或激活环境后直接运行
.venv\Scripts\activate
emf2png.py 产品介绍.pptx --trim -s 4
```

### 3.1 架构设计

```
emf2png/
├── emf2png.py              # 主入口 CLI
├── src/                    # 核心模块（CLI 核心逻辑）
│   ├── __init__.py
│   ├── emf_to_png.py       # EMF → PNG 模块
│   ├── ppt_to_emf.py       # PPT → EMF 模块
│   ├── trim_whitespace.py  # 裁剪纯白边
│   └── merge_pdf.py        # 可选，PNG 合并为 PDF
├── gui/                    # 图形界面
│   ├── __init__.py
│   └── app.py              # CustomTkinter 主窗口
├── bin/
│   └── emf_to_png.exe      # 预编译转换器 (来自参考项目)
├── docs/
│   └── 开发方案_PPT转EMF转PNG.md
├── .venv/                  # uv 虚拟环境
├── requirements.txt
├── .gitignore
└── README.md
```

### 3.1 主入口 CLI

```python
# emf2png.py — 结构概览
import argparse
from pathlib import Path
from ppt_to_emf import ppt_to_emf
from emf_to_png import batch_emf_to_png

def main():
    parser = argparse.ArgumentParser(
        description="将 PowerPoint 文件转换为 PNG 图片 (PPT → EMF → PNG)"
    )
    parser.add_argument("input", help="输入的 .ppt 或 .pptx 文件路径")
    parser.add_argument("-o", "--output", default="./output",
                        help="输出目录 (默认: ./output)")
    parser.add_argument("-s", "--scale", type=float, default=2.0,
                        help="PNG 缩放倍率 (默认: 2.0)")
    parser.add_argument("--dpi", type=int, default=300,
                        help="输出 DPI (默认: 300)")
    parser.add_argument("--keep-emf", action="store_true",
                        help="保留中间 EMF 文件")
    parser.add_argument("--merge-pdf", action="store_true",
                        help="将所有 PNG 合并为一个 PDF")
    parser.add_argument("--trim", action="store_true",
                        help="裁剪纯白边 (#FFFFFF)，非白底背景自动跳过")
    parser.add_argument("--start", type=int, default=1,
                        help="起始页码 (默认: 1)")
    parser.add_argument("--end", type=int, default=None,
                        help="结束页码 (默认: 全部)")
    args = parser.parse_args()
    # ...
```

### 3.2 参数总表

| 参数 | 说明 | 默认 |
|------|------|------|
| `input` | PPT/PPTX 文件路径 | **必填** |
| `-o / --output` | 输出目录 | `./output` |
| `-s / --scale` | PNG 缩放倍率 | `2.0` |
| `--dpi` | 输出 DPI | `300` |
| `--keep-emf` | 保留中间 EMF 文件 | `False` |
| `--merge-pdf` | 合成为 PDF | `False` |
| `--trim` | 裁剪纯白边 | `False` |
| `--start` | 起始页码 | `1` |
| `--end` | 结束页码 | 全部 |

---

## 四、分阶段详细开发计划

### 阶段 0：项目初始化

| # | 任务 | 说明 |
|---|------|------|
| 0.1 | 创建 `.gitignore` | 忽略 test.pptx、`__pycache__/`、`.env`、`output/`、`*.emf` |
| 0.2 | 创建 `requirements.txt` | 写入 pywin32、Pillow、numpy、tqdm、img2pdf |
| 0.3 | 初始化 Git 仓库 | `git init`，首次 commit |
| 0.4 | 创建 GitHub 仓库 | `gh repo create emf2png --public --push` |
| 0.5 | 创建 `README.md` | 项目介绍、用法、参数表 |

### 阶段 1：CLI 骨架

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 1.1 | 创建项目目录结构 | `emf2png/` 各模块占位文件 | 目录完整 |
| 1.2 | 实现参数解析 | `emf2png.py` 中 argparse 完整定义 | `python emf2png.py --help` 输出正确 |
| 1.3 | 实现输出目录自动创建 | `Path(output).mkdir(parents=True, exist_ok=True)` | 运行后目录存在 |
| 1.4 | 实现日志/进度输出框架 | print + tqdm 占位 | 运行可见进度 |

### 阶段 2：EMF → PNG 模块

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 2.1 | 复制 `emf_to_png.exe` 到项目 | `emf2png/emf_to_png.exe` | 文件存在 |
| 2.2 | 实现 `emf_to_png.py` — 单文件转换 | `convert(emf_path, png_path, scale)` | 用 test.emf 转换成功 |
| 2.3 | 实现子进程调用 exe | `subprocess.run([exe, input, output])` | 捕获 stdout/stderr |
| 2.4 | 实现错误处理 | exe 不存在、转换失败、超时 | 模拟缺失 exe 报错明确 |
| 2.5 | 实现 `batch_emf_to_png()` | 批量转换列表 | 3 个以上 emf 批量成功 |
| 2.6 | 实现 `--keep-emf` 控制 | 转换后删除/保留 emf | 开关行为正确 |
| 2.7 | 写入主流程步骤2调用 | emf2png.py 中串联 | 完整链路 EMF→PNG |

### 阶段 3：PPT → EMF 模块

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 3.1 | 探测 Office 安装 | `win32com.client.Dispatch("PowerPoint.Application")` 异常处理 | 未装 Office 报友好提示 |
| 3.2 | 实现打开 PPT 文件 | `app.Presentations.Open(ppt_path, WithWindow=False)` | 后台打开成功 |
| 3.3 | 实现幻灯片遍历 + EMF 导出 | `slide.Export(emf_path, "EMF", width, height)` | 导出 emf 文件存在 |
| 3.4 | 实现页码范围控制 | `--start` / `--end` 参数过滤 | 只导出指定页 |
| 3.5 | 实现输出路径组织 | emf 统一存到 `output/` | 目录整洁 |
| 3.6 | 实现关闭和清理 | `pres.Close()` / `app.Quit()` | 无 PowerPoint 进程残留 |
| 3.7 | 实现错误处理 | 文件损坏、页数超限、COM 错误 | 异常页跳过，不中断 |
| 3.8 | 写入主流程步骤1调用 | emf2png.py 中串联 PPT→EMF→PNG | `python emf2png.py test.pptx` 完整跑通 |

### 阶段 4：裁剪白边模块

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 4.1 | 实现单图裁剪 `trim_white_borders()` | Pillow 打开 → numpy 扫描 → crop → 覆盖保存 | 白边图裁剪后尺寸正确 |
| 4.2 | 实现白底检测逻辑 | 首行/末行/首列/末列全白才裁 | 非白底图跳过 |
| 4.3 | 实现透明像素处理 | alpha=0 视为白 | 半透明边缘不误裁 |
| 4.4 | 实现 `batch_trim_white_borders()` | 批量处理列表 | 批量验证 |
| 4.5 | 写入主流程步骤2.5 | emf2png.py 中 `--trim` 分支 | `--trim` 生效 |

### 阶段 5：缩放与 DPI 控制

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 5.1 | 将 `--scale` 参数传给 `emf_to_png.exe` | exe 支持 scale 参数 | scale=2 比 scale=1 像素翻倍 |
| 5.2 | 将 `--dpi` 参数写入 PNG metadata | Pillow 设置 dpi 信息 | `identify -verbose` 验证 |
| 5.3 | 测试不同缩放组合 | 1x/2x/4x/8x | 清晰度递增 |

### 阶段 6：PDF 合并模块

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 6.1 | 实现 `images_to_pdf()` | 用 img2pdf 或 reportlab 合并 | 生成 PDF 可正常打开 |
| 6.2 | 实现页码顺序 | 按 slide_001 → slide_N 顺序 | PDF 页码正确 |
| 6.3 | 写入主流程步骤3 | emf2png.py 中 `--merge-pdf` 分支 | PDF 文件生成 |

### 阶段 7：进度显示与用户体验

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 7.1 | 集成 tqdm 进度条 | PPT→EMF 进度、EMF→PNG 进度 | 进度条动起来 |
| 7.2 | 优化彩色日志输出 | 步骤编号 [1/3] [2/3] 清晰 | 可读性检查 |
| 7.3 | 统计汇总输出 | 完成时显示耗时、页数、总大小 | 验证正确性 |

### 阶段 8：错误处理与健壮性

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 8.1 | Office 未安装检测 | 友好提示 + exit code 1 | 在无 Office 机器测试 |
| 8.2 | PPT 文件不存在/损坏 | 明确报错 | 传不存在的文件 |
| 8.3 | EM→PNG exe 缺失 | 自动 fallback 到 Python 实现 | 删除 exe 后测试 |
| 8.4 | 输出磁盘空间不足 | 捕获写入异常 | 模拟满盘 |
| 8.5 | 超时保护（PPT 卡死） | 子进程 timeout | 处理大文件 |

### 阶段 9：GUI — 技术选型与窗口骨架

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 9.1 | 安装 customtkinter | `uv pip install customtkinter` | import 成功 |
| 9.2 | 创建 `gui/app.py` 主文件 | `gui/__init__.py` + `gui/app.py` | 文件存在 |
| 9.3 | 实现主窗口框架 | `CTk` 窗口，设置标题/大小/最小尺寸 | 窗口能打开 |
| 9.4 | 实现三栏布局 | 顶部标题栏 / 中间参数区 / 底部按钮栏 | `grid` 布局正确 |
| 9.5 | 实现菜单栏或标题 | "文件"/"帮助" 菜单或醒目标题文字 | 界面元素可见 |

### 阶段 10：GUI — 文件选择与输出目录

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 10.1 | "选择 PPT 文件"按钮 + 路径显示 | `CTkButton` + `CTkEntry` | 点击弹出文件对话框 |
| 10.2 | 文件类型过滤 | `.ppt` `.pptx` 过滤器 | 只显示 PPT 文件 |
| 10.3 | 拖拽文件支持 | 拖入 .pptx 自动填充路径 | 拖拽文件到窗口 |
| 10.4 | "输出目录"选择组件 | `CTkButton` + `CTkEntry` | 弹出目录选择器 |
| 10.5 | 默认输出目录逻辑 | 未选择时默认 `./output/幻灯片名/` | 自动填充 |

### 阶段 11：GUI — 参数配置面板

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 11.1 | 缩放倍率 (Scale) | `CTkSlider` + 标签显示数值 | 拖动滑块数值联动 |
| 11.2 | DPI 输入 | `CTkEntry`（数字输入框） | 输入/修改 DPI 值 |
| 11.3 | 页码范围（起始/结束） | 两个 `CTkEntry` 或 `CTkSpinbox` | 输入页码值 |
| 11.4 | 复选框组 | 三个 `CTkCheckBox`: 裁剪白边 / 保留 EMF / 合并 PDF | 勾选状态正确 |
| 11.5 | 参数面板布局美化 | 分组标题、间距、对齐 | 视觉整齐 |

### 阶段 12：GUI — 转换执行与进度

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 12.1 | "开始转换"主按钮 | 醒目 `CTkButton` | 点击触发转换 |
| 12.2 | 后台线程执行 | `threading.Thread` 运行转换，不阻塞界面 | 界面不卡死 |
| 12.3 | 进度条组件 | `CTkProgressBar`，0~100% | 进度实时更新 |
| 12.4 | 日志输出区域 | `CTkTextbox`（只读），显示步骤信息 | 日志逐行追加 |
| 12.5 | 回调桥接 | 将 CLI 的 `progress_callback` 连接到 GUI 更新 | 进度条 + 日志同时更新 |
| 12.6 | 按钮状态管理 | 转换中禁用按钮 / 完成恢复 | 不可重复点击 |

### 阶段 13：GUI — 结果展示与功能完善

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 13.1 | 完成汇总展示 | 显示页数、耗时、文件大小 | 信息完整 |
| 13.2 | "打开输出目录"按钮 | 调用 `os.startdir()` 打开文件夹 | 资源管理器弹出 |
| 13.3 | 错误对话框 | `CTkMessagebox` 或自定义弹窗 | 错误时弹出 |
| 13.4 | 再次转换（不清空日志） | 支持连续转换不关闭窗口 | 多次点击正常 |
| 13.5 | 窗口关闭确认 | 转换中关闭时弹出确认 | 防止误关 |

### 阶段 14：GUI — 美化与细节打磨

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 14.1 | 暗色/亮色主题切换 | `set_appearance_mode()` 默认暗色 | 主题统一 |
| 14.2 | 色彩方案 | 主题色、按钮色、背景色统一 | 视觉一致 |
| 14.3 | 应用图标 | `.ico` 文件 + `iconbitmap()` | 标题栏有图标 |
| 14.4 | 窗口大小与位置记忆 | `tkinter.Tk.geometry()` 保存/恢复 | 关闭重开位置不变 |
| 14.5 | 错误边界（GUI 级） | 全局 `try/except` + 弹窗提示 | 异常不崩溃 |

### 阶段 15：PyInstaller 打包

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 15.1 | 编写 `build_exe.spec` | 参考 `emf_to_png.spec`，包含 gui/ 和 src/ | 编译通过 |
| 15.2 | 打包 CLI 版 | `pyinstaller build_exe.spec` CLI 入口 | 生成 `dist/emf2png.exe` |
| 15.3 | 打包 GUI 版 | 单独 spec 或入口切换，打包 `gui/app.py` | 生成 `dist/emf2png-gui.exe` |
| 15.4 | 嵌入 `emf_to_png.exe` | 作为数据文件打包进 exe | exe 能找到内部资源 |
| 15.5 | 裸机测试（CLI） | 在无 Python 环境运行 | 完整链路可用 |
| 15.6 | 裸机测试（GUI） | 在无 Python 环境运行 | 窗口正常打开 |
| 15.7 | GitHub Release 发布 | 上传 exe 到 Release | 可下载 |

### 阶段 16：GitHub 集成与 CI

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 16.1 | 配置 `.gitignore` 最终版 | 排除所有不应上传的文件 | `git status` 整洁 |
| 16.2 | GitHub Actions：Python lint | flake8 / ruff 检查 | CI 绿 |
| 16.3 | GitHub Actions：自动打包 CLI | push tag 自动 PyInstaller + Release | Release 自动创建 |
| 16.4 | GitHub Actions：自动打包 GUI | push tag 自动 PyInstaller + Release | Release 自动创建 |
| 16.5 | GitHub Actions：测试 | 基本功能测试（无 Office 部分） | CI 绿 |

### 阶段 17：文档完善

| # | 任务 | 产出 | 验证方法 |
|---|------|------|----------|
| 17.1 | `README.md` 最终版 | 安装、用法、示例、参数、GUI 截图完整 | 检查无误 |
| 17.2 | GUI 界面截图 | README 中展示主窗口截图 | 美观 |
| 17.3 | 添加 `--help` 输出截图 | README 中展示 | 美观 |
| 17.4 | 添加常见问题 FAQ | Office 要求、字体、性能等 | 覆盖常见疑问 |

### 阶段 18：可选高级功能

| # | 任务 | 说明 | 优先级 |
|---|------|------|--------|
| 18.1 | 批量文件夹处理 | 一次处理多个 PPT | P2 |
| 18.2 | 水印/页眉页脚 | 导出前叠加 | P3 |
| 18.3 | 对比模式（原图 vs PNG） | 质量控制 | P3 |
| 18.4 | 保留原始 EMF 矢量嵌入 | 输出 EMF + PNG 双份 | P3 |

---

## 五、依赖清单

```txt
# requirements.txt
pywin32>=306        # PowerPoint COM 自动化
Pillow>=10.0.0      # 图片处理（EMF→PNG 底层 + 裁剪白边）
numpy>=1.24.0       # 裁剪白边算法（像素矩阵运算）
tqdm>=4.65.0        # 进度条（CLI 模式）
customtkinter>=5.2.0 # GUI 界面（可选）
img2pdf>=0.4.0      # PNG→PDF 合并（可选）
```

打包时额外：

```txt
# build requirements
pyinstaller>=6.0.0
```

---

## 六、关键技术难点

### 6.1 EMF 渲染质量

**问题**：EMF 是矢量格式，渲染为 PNG 时精度取决于渲染引擎。

**对策**：
- 使用 Windows GDI 原生渲染（参考项目的 exe 方案）
- `--scale` 参数控制渲染分辨率（推荐 2x ~ 4x）
- 输出 PNG 保持透明背景（EMF 支持透明）

### 6.2 PowerPoint COM 自动化限制

**问题**：
- 需安装 Microsoft Office（PowerPoint）
- 后台模式无法导出某些元素（视频、ActiveX）
- 并发限制（PowerPoint 单实例）

**对策**：
- 检测 Office 安装，未安装时给出清晰指引
- 串行处理，避免并发
- 捕获 COM 异常并跳过故障页

### 6.3 中文字体兼容

**问题**：PPT 中的中文字体在渲染时可能缺失。

**对策**：
- 依赖本地系统字体，无需额外配置
- 导出时 PowerPoint 自动嵌入字体（需 PPT 设置）
- 提示用户如需跨机器运行，建议 PPT 中勾选"嵌入字体"

### 6.4 裁剪白边算法

**核心逻辑**：四边扫描法，严格匹配 `#FFFFFF`。

```python
from PIL import Image
import numpy as np

def trim_white_borders(png_path: str) -> str:
    """
    裁剪 PNG 的纯白边 (#FFFFFF)。
    若图片非白底（任何边框行/列含有非白像素），则跳过，返回原路径。
    """
    img = Image.open(png_path).convert("RGBA")
    arr = np.array(img)

    alpha = arr[:, :, 3]
    rgb = arr[:, :, :3]

    # 白色掩码：R=255, G=255, B=255 或 alpha=0（透明视为白）
    white_mask = (rgb[:, :, 0] == 255) & (rgb[:, :, 1] == 255) & (rgb[:, :, 2] == 255)
    white_mask |= (alpha == 0)

    rows = np.all(white_mask, axis=1)   # 每行是否全白
    cols = np.all(white_mask, axis=0)   # 每列是否全白

    # 判断是否为白底：首行、末行、首列、末列都全白 → 是
    if not (rows[0] and rows[-1] and cols[0] and cols[-1]):
        return png_path  # 非白底，跳过

    # 找到第一个非白行/列
    top = np.argmax(~rows) if not np.all(rows) else 0
    bottom = len(rows) - np.argmax(~rows[::-1]) if not np.all(rows) else len(rows)
    left = np.argmax(~cols) if not np.all(cols) else 0
    right = len(cols) - np.argmax(~cols[::-1]) if not np.all(cols) else len(cols)

    cropped = img.crop((left, top, right, bottom))
    cropped.save(png_path)
    return png_path

def batch_trim_white_borders(png_files: list[str]) -> list[str]:
    for f in png_files:
        trim_white_borders(f)
    return png_files
```

---

## 七、参考项目结构分析

取自 `E:\projects\emf-pro-converter_with_exe`：

```
emf-pro-converter_with_exe/
├── server/
│   ├── index.js              # Express 服务器（接收上传 → 调用 exe → 返回 PNG）
│   ├── bin/
│   │   └── emf_to_png.exe    # ★ 核心：EMF→PNG 转换器（PyInstaller 打包）
│   └── scripts/
│       └── emf_to_png.py     # ★ 核心源码（我们复用其逻辑）
├── services/
│   ├── emfProcessor.ts       # 前端 EMF 处理服务（调用后端 API）
│   └── geminiService.ts      # Gemini AI 图片分析（我们不需要）
├── App.tsx                   # React 前端（我们不需要）
├── components/               # React 组件（我们不需要）
└── emf_to_png.spec           # PyInstaller 打包配置
```

**可复用**：`emf_to_png.exe`、`emf_to_png.py`（源码参考）、`emf_to_png.spec`（打包参考）

**不需要**：所有 Web 前端（React/Express）、Gemini AI 服务

---

## 八、使用示例

```bash
# 1. 基本用法
python emf2png.py 产品介绍.pptx

# 2. 高清输出
python emf2png.py 产品介绍.pptx -o ./output --scale 4 --dpi 600

# 3. 裁剪白边 + 高清（提取素材标准用法）
python emf2png.py 产品介绍.pptx --trim -s 4

# 4. 指定页码 + 合并 PDF
python emf2png.py 产品介绍.pptx --start 3 --end 10 --merge-pdf

# 5. 保留中间 EMF
python emf2png.py 产品介绍.pptx --keep-emf

# 6. 打包版
emf2png.exe 产品介绍.pptx --trim -s 4
```

---

## 九、输出目录结构

```
output/
├── slide_001.png       # 第1页
├── slide_002.png       # 第2页
├── ...
├── slide_020.png       # 第20页
├── output.pdf          # (可选) 合并后的 PDF
└── .emf/               # (仅 --keep-emf 时)
    ├── slide_001.emf
    ├── slide_002.emf
    └── ...
```

---

## 十、进度总览

```
阶段 0: 项目初始化        ██████████ 100%  ✅
阶段 1: CLI 骨架          ██████████ 100%  ✅
阶段 2: EMF→PNG 模块     ██████████ 100%  ✅
阶段 3: PPT→EMF 模块     ██████████ 100%  ✅
阶段 4: 裁剪白边         ██████████ 100%  ✅
阶段 5: 缩放与 DPI       ██████████ 100%  ✅
阶段 6: PDF 合并          ██████████ 100%  ✅
阶段 7: 进度与 UX         ██████████ 100%  ✅
阶段 8: 错误处理          ██████████ 100%  ✅
─────────────────────────────────────────────
阶段 9: GUI 窗口骨架      ░░░░░░░░░░   0%  ◀ 当前
阶段10: GUI 文件选择      ░░░░░░░░░░   0%
阶段11: GUI 参数面板      ░░░░░░░░░░   0%
阶段12: GUI 转换与进度    ░░░░░░░░░░   0%
阶段13: GUI 结果展示      ░░░░░░░░░░   0%
阶段14: GUI 美化细节      ░░░░░░░░░░   0%
阶段15: PyInstaller 打包  ░░░░░░░░░░   0%
阶段16: GitHub CI         ░░░░░░░░░░   0%
阶段17: 文档完善          ░░░░░░░░░░   0%
```

---

*文档生成日期: 2025-07-17*
*参考项目: emf-pro-converter_with_exe (EMF→PNG Web 版)*
