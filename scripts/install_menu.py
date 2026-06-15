#!/usr/bin/env python3
"""
emf2png 右键菜单安装器。

双击运行即可安装/卸载。
本脚本会自动以正确编码写入注册表，无乱码问题。
"""

import sys
import os
import winreg
from pathlib import Path

ROOT = Path(__file__).parent.parent.resolve()
PYTHONW = ROOT / ".venv" / "Scripts" / "pythonw.exe"
PDF_SCRIPT = ROOT / "ppt_to_pdf.py"
PNG_SCRIPT = ROOT / "convert_to_png.py"

# 注册表路径模板：每个菜单项使用独立的键名
# 用不同键名而非子键，避免级联菜单兼容性问题
KEY_TPL = r"Software\Classes\SystemFileAssociations\{ext}\shell\emf2png-{action}"

MENU_ITEMS = [
    # (扩展名, 动作名, 显示文字, 脚本路径)
    (".pptx", "pdf", "导出 PDF(裁剪白边)", PDF_SCRIPT),
    (".pptm", "pdf", "导出 PDF(裁剪白边)", PDF_SCRIPT),
    (".ppt", "pdf", "导出 PDF(裁剪白边)", PDF_SCRIPT),
    (".emf", "png", "转为 PNG(裁剪白边)", PNG_SCRIPT),
    (".pptx", "png", "导出 PNG(裁剪白边)", PNG_SCRIPT),
    (".pptm", "png", "导出 PNG(裁剪白边)", PNG_SCRIPT),
    (".ppt", "png", "导出 PNG(裁剪白边)", PNG_SCRIPT),
]


def is_installed() -> bool:
    """检查是否已安装（任一菜单项存在即可）。"""
    first = MENU_ITEMS[0]
    key_path = KEY_TPL.format(ext=first[0], action=first[1])
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_QUERY_VALUE)
        winreg.QueryValueEx(key, "")
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False


def install():
    """安装右键菜单。"""
    print("安装右键菜单...")

    # 先清理所有旧格式（兼容 emf2png 直接键、emf2png\pdf 子键等）
    for ext in (".pptx", ".pptm", ".ppt", ".emf"):
        base = r"Software\Classes\SystemFileAssociations\{ext}\shell\emf2png".format(ext=ext)
        for suffix in ["", r"\command", r"\pdf\command", r"\pdf", r"\png\command", r"\png"]:
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, base + suffix)
            except FileNotFoundError:
                pass
            except Exception:
                pass

    for ext, action, label, script in MENU_ITEMS:
        key_path = KEY_TPL.format(ext=ext, action=action)
        try:
            # 创建主键（带默认值）
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, label)
            winreg.CloseKey(key)

            # 创建 command 子键
            cmd_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r"\command")
            cmd = f'"{PYTHONW}" "{script}" "%1"'
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, cmd)
            winreg.CloseKey(cmd_key)
            print(f"  [OK] {ext} [{action}] -> {label}")
        except Exception as e:
            print(f"  [ERR] {ext} [{action}]: {e}")

    print(f"\n[OK] 安装完成！右键文件试试。")


def uninstall():
    """卸载右键菜单。"""
    print("卸载右键菜单...")
    # 收集所有唯一的键路径
    keys = set()
    for ext, action, _, _ in MENU_ITEMS:
        keys.add(KEY_TPL.format(ext=ext, action=action))

    # 先删 command 子键，再删父键
    for key_path in keys:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path + r"\command")
        except FileNotFoundError:
            pass
    for key_path in keys:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
            print(f"  [OK] 已卸载: {key_path}")
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"  [ERR] {key_path}: {e}")

    print("\n[OK] 已卸载。")


def main():
    print("=" * 40)
    print("  emf2png 右键菜单安装器")
    print("=" * 40)
    print()

    if not PYTHONW.exists():
        print(f"[ERR] 找不到 pythonw.exe: {PYTHONW}")
        print("请确认虚拟环境存在: uv venv")
        input("\n按 Enter 退出...")
        sys.exit(1)

    if not PDF_SCRIPT.exists():
        print(f"[ERR] 找不到 {PDF_SCRIPT}")
        input("\n按 Enter 退出...")
        sys.exit(1)
    if not PNG_SCRIPT.exists():
        print(f"[ERR] 找不到 {PNG_SCRIPT}")
        input("\n按 Enter 退出...")
        sys.exit(1)

    if len(sys.argv) > 1:
        # 命令行模式
        if sys.argv[1] == "install":
            install()
        elif sys.argv[1] == "uninstall":
            uninstall()
        return

    # 交互模式
    if is_installed():
        print("右键菜单已安装。")
        choice = input("按 1 卸载, 按 2 重新安装, 按 Enter 退出: ").strip()
        if choice == "1":
            uninstall()
        elif choice == "2":
            uninstall()
            print()
            install()
    else:
        print("右键菜单未安装。")
        choice = input("按 1 安装, 按 Enter 退出: ").strip()
        if choice == "1":
            install()

    input("\n按 Enter 退出...")


if __name__ == "__main__":
    main()
