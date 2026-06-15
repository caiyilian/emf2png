@echo off
chcp 65001 >nul
title emf2png - 卸载右键菜单
cd /d "%~dp0"
.venv\Scripts\python.exe install_menu.py uninstall
echo.
pause
