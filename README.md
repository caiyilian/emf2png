# emf2png

> 一键将 PowerPoint 幻灯片导出为高清 PNG 图片，或将 Draw.io 图表导出为裁剪白边的单页 PDF。

```
PPT/PPTX  →  EMF  →  PNG (可选裁剪白边)  →  PDF (可选合并)
.drawio   →  draw.io CLI → PDF → 单页白边裁剪
```

---

## 场景

从 PPT 中提取**图表、流程图、架构图、图标、插画**等素材：

```bash
emf2png.exe 产品介绍.pptx --trim -s 4
# 输出: slide_001.png, slide_002.png, ...（已裁白边，直接用于设计工具）
```

---

## 快速开始

### 方式一：使用打包好的 exe（无需 Python）

从 [Releases](https://github.com/caiyilian/emf2png/releases) 下载 `emf2png.exe`，然后：

```bash
emf2png.exe 产品介绍.pptx -o ./output
emf2png.exe 产品介绍.pptx --trim -s 2          # 裁剪白边 + 2x 高清
emf2png.exe 产品介绍.pptx --start 3 --end 10    # 指定页码范围
emf2png.exe 架构图.drawio -o ./output            # Draw.io → 单页 PDF + 裁剪白边
```

### 方式二：使用 uv（推荐）

项目使用 [uv](https://docs.astral.sh/uv/) 管理 Python 环境和依赖。

```bash
# 创建虚拟环境并安装依赖（已有 .venv 则跳过）
uv venv
uv pip install -r requirements.txt

# 运行
uv run emf2png.py 产品介绍.pptx --trim -s 4
```

### 方式三：Python 源码运行（传统 pip）

```bash
pip install -r requirements.txt
python emf2png.py 产品介绍.pptx --trim -s 4
```

Draw.io 文件也可以直接转换：

```bash
python drawio_to_pdf.py 架构图.drawio
python drawio_to_pdf.py 架构图.drawio -o ./output/架构图.pdf
# 如需只裁剪纯白边，可显式启用严格模式
python drawio_to_pdf.py 架构图.drawio --strict
```

---

## 参数说明

| 参数 | 说明 | 默认 |
|------|------|------|
| `input` | PPT/PPTX/EMF/Draw.io 文件路径 | **必填** |
| `-o, --output` | 输出目录 | `./output` |
| `-s, --scale` | PNG 缩放倍率（越高越清晰） | `2.0` |
| `--dpi` | 输出 DPI | `300` |
| `--trim` | 裁剪纯白边 (#FFFFFF)，非白底自动跳过 | `False` |
| `--no-trim-strict` | 宽松模式：使用 >=248 阈值，解决抗锯齿白边问题 | 严格模式(默认) |
| `--keep-emf` | 保留中间 EMF 文件 | `False` |
| `--merge-pdf` | 合并为 PDF | `False` |
| `--start` | 起始页码 | `1` |
| `--end` | 结束页码 | 全部 |
| `--drawio-border` | Draw.io 导出边框宽度 | `0` |
| `--pdf-zoom` | Draw.io PDF 白边检测渲染倍率 | `2.5` |
| `--drawio-strict` | Draw.io 只把纯白 `#FFFFFF` 视为空白 | `False` |

---

## 典型用法

```bash
# 素材提取（推荐）
emf2png.exe 产品介绍.pptx --trim -s 4

# 整本导出为 PDF
emf2png.exe 产品介绍.pptx --merge-pdf

# 指定范围和缩放
emf2png.exe 产品介绍.pptx --start 5 --end 15 -s 3 --trim --merge-pdf

# Draw.io 导出单页 PDF 并裁剪白边
emf2png.exe 架构图.drawio -o ./output
python drawio_to_pdf.py 架构图.drawio --zoom 3 --dpi 300
```

---

## 系统要求

- **操作系统**: Windows 10/11
- **Office**: Microsoft PowerPoint（用于 PPT→EMF 导出）
- **Draw.io**: draw.io Desktop（用于 .drawio→PDF 导出，可通过 `DRAWIO_EXE` 指定路径）
- **Python**（仅源码运行需要）: 3.10+（推荐使用项目的 uv 环境，Python 3.12）

---

## 工作流程

```
.ppt/.pptx
     │
     ▼  PowerPoint COM 自动化
     │  后台打开 PowerPoint，每页导出为 EMF
     │
  .emf 文件 (每页一个)
     │
     ▼  emf_to_png.exe
     │  使用 Windows GDI 原生渲染，保持矢量精度
     │
  .png 文件 (含白底)
     │
     ▼  裁剪白边 (--trim)
     │  四边扫描法，严格匹配 #FFFFFF
     │  非白底背景自动跳过，保留原样
     │
  .png 文件 (内容紧凑)
     │
     ▼  合并 PDF (--merge-pdf)
     │
  output.pdf
```

Draw.io 文件使用独立链路，不依赖 PowerPoint：

```
.drawio
     │
     ▼  draw.io CLI（--export --format pdf --crop --border 0）
     │
  临时 PDF
     │
     ▼  drawio_to_pdf.py（渲染、按 >=248 阈值和 99.5% 行列容差寻找边界）
     │
  单页 PDF（最终输出）
```

---

## 项目结构

```
emf2png/
├── emf2png.py              # 主入口 CLI
├── drawio_to_pdf.py        # Draw.io → 单页 PDF + 白边裁剪
├── pdf_trim.py             # 单页 PDF 白边裁剪工具
├── src/                    # 核心模块
│   ├── __init__.py
│   ├── emf_to_png.py       # EMF → PNG 模块
│   ├── ppt_to_emf.py       # PPT → EMF 模块
│   ├── trim_whitespace.py  # 白边裁剪模块
│   └── merge_pdf.py        # PDF 合并模块
├── bin/
│   └── emf_to_png.exe      # 预编译转换器
├── docs/
│   └── 开发方案_PPT转EMF转PNG.md
├── .venv/                  # uv 虚拟环境
├── requirements.txt
├── .gitignore
└── README.md
```

---

## 许可证

MIT
