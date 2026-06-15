Sub ConvertToPDF()
    Dim pptPath As String
    Dim scriptPath As String
    Dim pythonExe As String
    Dim command As String
    Dim shellObj As Object
    Dim waitUntil As Date
    
    ' 获取当前 PPT 完整路径
    pptPath = ActivePresentation.FullName
    
    ' ===== 以下路径请根据你的实际安装位置修改 =====
    
    ' ppt_to_pdf.py 的路径
    scriptPath = "E:\projects\emf2png\ppt_to_pdf.py"
    
    ' Python 可执行文件路径（用项目虚拟环境的 Python，避免 PATH 问题）
    pythonExe = "E:\projects\emf2png\.venv\Scripts\python.exe"
    
    ' ============================================
    
    ' 构建完整命令
    command = pythonExe & " """ & scriptPath & """ """ & pptPath & """"
    
    ' 用 WScript.Shell 执行并等待完成
    Set shellObj = CreateObject("WScript.Shell")
    shellObj.Run command, 0, True  ' 0=隐藏窗口, True=等待完成
    
    ' 检查 PDF 是否生成
    Dim pdfPath As String
    pdfPath = Left(pptPath, InStrRev(pptPath, ".")) & "pdf"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(pdfPath) Then
        MsgBox "转换完成！" & vbCrLf & vbCrLf & _
               "PDF 已保存至: " & pdfPath, _
               vbInformation, "emf2png"
    Else
        MsgBox "转换似乎失败了。" & vbCrLf & vbCrLf & _
               "请尝试在命令行中运行以下命令查看错误信息：" & vbCrLf & _
               pythonExe & " """ & scriptPath & """ """ & pptPath & """", _
               vbExclamation, "emf2png"
    End If
End Sub
