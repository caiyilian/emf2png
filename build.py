#!/usr/bin/env python3
"""Build emf2png executables with PyInstaller.

Usage:
    uv run python build.py          # 打包 CLI + GUI
    uv run python build.py cli      # 只打包 CLI
    uv run python build.py gui      # 只打包 GUI
"""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent
DIST = ROOT / "dist"


def _pyinstaller(args: list[str]):
    """Run PyInstaller with given args."""
    cmd = [sys.executable, "-m", "PyInstaller", "--clean", "--noconfirm"] + args
    print(f"[BUILD] {' '.join(str(a) for a in cmd)}")
    subprocess.run(cmd, check=True, cwd=ROOT)


def build_cli():
    """打包 CLI 版: emf2png.exe"""
    print("=" * 60)
    print("Building CLI version: emf2png.exe")
    print("=" * 60)
    _pyinstaller([
        "--onefile",
        "--name", "emf2png",
        "--add-data", "src;src",
        "--add-data", "bin;bin",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL._tkinter_finder",
        "--hidden-import", "numpy",
        "--hidden-import", "img2pdf",
        "emf2png.py",
    ])
    exe = DIST / "emf2png.exe"
    if exe.exists():
        size_mb = exe.stat().st_size / 1024 / 1024
        print(f"[OK] CLI built: {exe} ({size_mb:.1f} MB)")


def build_gui():
    """打包 GUI 版: emf2png-gui.exe"""
    print("=" * 60)
    print("Building GUI version: emf2png-gui.exe")
    print("=" * 60)
    _pyinstaller([
        "--onefile",
        "--windowed",
        "--name", "emf2png-gui",
        "--add-data", "src;src",
        "--add-data", "bin;bin",
        "--add-data", "gui;gui",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL._tkinter_finder",
        "--hidden-import", "numpy",
        "--hidden-import", "img2pdf",
        "--hidden-import", "customtkinter",
        "gui/app.py",
    ])
    exe = DIST / "emf2png-gui.exe"
    if exe.exists():
        size_mb = exe.stat().st_size / 1024 / 1024
        print(f"[OK] GUI built: {exe} ({size_mb:.1f} MB)")


def main():
    targets = sys.argv[1:] if len(sys.argv) > 1 else ["cli", "gui"]
    for t in targets:
        if t == "cli":
            build_cli()
        elif t == "gui":
            build_gui()
        else:
            print(f"[!] Unknown target: {t}")

    print("\nDone! Files in dist/")
    for f in sorted(DIST.iterdir()):
        size = f.stat().st_size / 1024 / 1024
        print(f"  {f.name}: {size:.1f} MB")


if __name__ == "__main__":
    main()
