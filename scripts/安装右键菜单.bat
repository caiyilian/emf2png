@echo off
chcp 65001 >nul
title emf2png - 安装右键菜单
cd /d "%~dp0.."
.venv\Scripts\python.exe scripts\install_menu.py install
echo.
pause
