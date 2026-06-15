#!/usr/bin/env python3
"""
install_addin.py — 将 emf2png VBA 宏安装为 PowerPoint 全局加载项。

用法:
    uv run python install_addin.py

效果:
    一次运行，所有 PPT 文件都能在"加载项"选项卡中找到"导出PDF(裁剪白边)"按钮。
    不需要在每个 PPT 文件中粘贴 VBA 代码。
"""

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).parent.resolve()


def main():
    print("emf2png — 安装 PowerPoint 全局加载项")
    print("=" * 50)

    # VBA 宏文件
    bas_path = ROOT / "docs" / "vba_macro.bas"
    if not bas_path.exists():
        print(f"[ERR] 找不到宏文件: {bas_path}")
        sys.exit(1)

    vba_code = bas_path.read_text(encoding="utf-8")
    print(f"[OK] 读取宏文件: {bas_path}")

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("[ERR] 缺少 pywin32，请执行: uv pip install pywin32")
        sys.exit(1)

    pythoncom.CoInitialize()
    app = None

    try:
        print("正在启动 PowerPoint ...")
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True

        # 检查是否可访问 VBA
        try:
            test_pres = app.Presentations.Add()
            _ = test_pres.VBProject
            test_pres.Close()
        except Exception:
            print("\n[!] PowerPoint 安全设置阻止了 VBA 自动安装。")
            print("   请在 PowerPoint 中设置:")
            print("   文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置")
            print("   → 勾选'信任对 VBA 工程对象模型的访问'\n")
            print("   设置后重新运行本脚本。\n")
            print("或者使用手动安装：")
            _print_manual_instructions(bas_path)
            sys.exit(1)

        # 临时打开一个空白演示文稿
        pres = app.Presentations.Add()

        # 直接用代码写入 VBA 模块（避免 .bas 文件导入编码问题）
        vb_project = pres.VBProject
        vb_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vb_module.Name = "emf2png"

        # 写入临时 .bas 文件（纯 ASCII，无编码问题）
        vba_bas = ROOT / "~temp_vba.bas"
        vba_code = """Attribute VB_Name = "emf2png"
Public Sub ConvertToPDF()
    Dim pptPath As String
    Dim scriptPath As String
    Dim pythonExe As String
    Dim command As String
    Dim shellObj As Object
    Dim pdfPath As String
    Dim fso As Object
    pptPath = ActivePresentation.FullName
    scriptPath = "E:\projects\emf2png\ppt_to_pdf.py"
    pythonExe = "E:\projects\emf2png\.venv\Scripts\python.exe"
    command = pythonExe & " " & Chr(34) & scriptPath & Chr(34) & " " & Chr(34) & pptPath & Chr(34)
    Set shellObj = CreateObject("WScript.Shell")
    shellObj.Run command, 0, True
    pdfPath = Left(pptPath, InStrRev(pptPath, ".")) & "pdf"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pdfPath) Then
        MsgBox "Done! PDF saved to: " & pdfPath, vbInformation, "emf2png"
    Else
        MsgBox "Conversion failed." & vbCrLf & command, vbExclamation, "emf2png"
    End If
End Sub"""
        vba_bas.write_text(vba_code, encoding="ascii")
        vb_module = vb_project.VBComponents.Import(str(vba_bas))
        vba_bas.unlink()
        print("[OK] VBA 代码写入成功")

        # 保存为临时 PPTM（明确指定宏启用格式）
        temp_ppt = ROOT / "~temp_addin.pptm"
        pres.SaveAs(str(temp_ppt), 25)  # 25 = ppSaveAsOpenXMLPresentationMacroEnabled
        pres.Close()

        # 复制到 AddIns 目录并注册为加载项
        addin_dir = Path.home() / "AppData" / "Roaming" / "Microsoft" / "AddIns"
        addin_dir.mkdir(parents=True, exist_ok=True)
        addin_path = addin_dir / "emf2png.pptm"
        import shutil
        shutil.copy2(str(temp_ppt), str(addin_path))

        # 用 AddIns.Add 注册
        try:
            registered = app.AddIns.Add(str(addin_path))
            registered.Loaded = True
            print(f"[OK] 加载项已注册并启用")
        except Exception as e:
            print(f"[!] 注册失败（可手动加载）：{e}")

        # 清理临时文件
        try:
            temp_ppt.unlink()
        except Exception:
            pass

        print(f"\n[OK] 安装完成！")
        print(f"     加载项: {addin_path}")
        print(f"\n下一步：关闭所有 PowerPoint 后重新打开任意 PPT 即可")
        _print_uninstall(addin_path)

    except Exception as e:
        print(f"\n[ERR] 安装失败: {e}")
        _print_manual_instructions(bas_path)
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _print_manual_instructions(bas_path: Path):
    print(f"\n手动安装方法：")
    print(f"   1. 打开 PowerPoint → Alt+F11 打开 VBA 编辑器")
    print(f"   2. 菜单: 文件 → 导入文件 → 选择：")
    print(f"      {bas_path}")
    print(f"   3. 菜单: 文件 → 另存为 → 选择文件类型:")
    print(f"      'PowerPoint 加载项 (*.ppam)'")
    print(f"      保存到: %APPDATA%\\Microsoft\\AddIns\\emf2png.ppam")
    print(f"   4. 关闭 PowerPoint 重新打开即可")


def _print_uninstall(addin_path: Path):
    print(f"\n卸载：")
    print(f"   删除文件: {addin_path}")
    print(f"   或在 PowerPoint 中: 文件 → 选项 → 加载项 → 管理:PowerPoint加载项 → 找到 emf2png → 删除")


if __name__ == "__main__":
    main()
