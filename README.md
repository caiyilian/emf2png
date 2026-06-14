# emf2png

> 一键将 PowerPoint 幻灯片导出为高清 PNG 图片，自动裁剪白边，提取矢量级素材。

```
PPT/PPTX  →  EMF  →  PNG (可选裁剪白边)  →  PDF (可选合并)
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

---

## 参数说明

| 参数 | 说明 | 默认 |
|------|------|------|
| `input` | PPT/PPTX 文件路径 | **必填** |
| `-o, --output` | 输出目录 | `./output` |
| `-s, --scale` | PNG 缩放倍率（越高越清晰） | `2.0` |
| `--dpi` | 输出 DPI | `300` |
| `--trim` | 裁剪纯白边 (#FFFFFF)，非白底自动跳过 | `False` |
| `--keep-emf` | 保留中间 EMF 文件 | `False` |
| `--merge-pdf` | 合并为 PDF | `False` |
| `--start` | 起始页码 | `1` |
| `--end` | 结束页码 | 全部 |

---

## 典型用法

```bash
# 素材提取（推荐）
emf2png.exe 产品介绍.pptx --trim -s 4

# 整本导出为 PDF
emf2png.exe 产品介绍.pptx --merge-pdf

# 指定范围和缩放
emf2png.exe 产品介绍.pptx --start 5 --end 15 -s 3 --trim --merge-pdf
```

---

## 系统要求

- **操作系统**: Windows 10/11
- **Office**: Microsoft PowerPoint（用于 PPT→EMF 导出）
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

---

## 项目结构

```
emf2png/
├── emf2png.py              # 主入口 CLI
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
