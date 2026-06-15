@echo off
chcp 65001 >nul
title emf2png - PPT 转 PDF（裁剪白边）
echo ============================================
echo   emf2png - PPT 转 PDF（自动裁剪白边）
echo ============================================
echo.
echo  把 PPT 文件拖到这个窗口上，松手即可转换
echo.
echo  按 Ctrl+C 退出
echo ============================================
echo.

:loop
set "file="
set /p file=拖入 PPT 文件路径: 

if "%file%"=="" goto loop

:: 去掉可能的引号
set file=%file:"=%

if not exist "%file%" (
    echo [错误] 文件不存在
    goto loop
)

echo.
echo 正在转换: %file%
echo.

cd /d "E:\projects\emf2png"
E:\projects\emf2png\.venv\Scripts\python.exe ppt_to_pdf.py "%file%"

if %errorlevel% equ 0 (
    echo.
    echo [完成] PDF 已生成！
) else (
    echo.
    echo [失败] 请查看上面的错误信息
)

echo.
goto loop
