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

ROOT = Path(__file__).parent.resolve()
PYTHONW = ROOT / ".venv" / "Scripts" / "pythonw.exe"
SCRIPT = ROOT / "ppt_to_pdf.py"

# 注册表路径模板
KEY_TPL = r"Software\Classes\SystemFileAssociations\{ext}\shell\emf2png"

MENU_ITEMS = {
    ".pptx": "导出 PDF(裁剪白边)",
    ".pptm": "导出 PDF(裁剪白边)",
    ".ppt": "导出 PDF(裁剪白边)",
    ".emf": "转为 PNG(裁剪白边)",
}


def is_installed() -> bool:
    """检查是否已安装。"""
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY_TPL.format(ext=".pptx"), 0, winreg.KEY_QUERY_VALUE)
        winreg.QueryValueEx(key, "")
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False


def install():
    """安装右键菜单。"""
    print("安装右键菜单...")
    for ext, label in MENU_ITEMS.items():
        key_path = KEY_TPL.format(ext=ext)
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, label)
            winreg.CloseKey(key)

            cmd_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r"\command")
            cmd = f'"{PYTHONW}" "{SCRIPT}" "%1"'
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, cmd)
            winreg.CloseKey(cmd_key)
            print(f"  [OK] {ext} -> {label}")
        except Exception as e:
            print(f"  [ERR] {ext}: {e}")

    print(f"\n[OK] 安装完成！右键 .pptx 文件试试。")


def uninstall():
    """卸载右键菜单。"""
    print("卸载右键菜单...")
    for ext in MENU_ITEMS:
        key_path = KEY_TPL.format(ext=ext)
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path + r"\command")
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
            print(f"  [OK] {ext} 已删除")
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"  [ERR] {ext}: {e}")
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

    if not SCRIPT.exists():
        print(f"[ERR] 找不到 {SCRIPT}")
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
