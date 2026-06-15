Attribute VB_Name = "emf2png"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Sub ConvertToPDF()
    Dim pptPath As String
    Dim scriptPath As String
    Dim pythonExe As String
    Dim command As String
    Dim shellObj As Object
    Dim pdfPath As String
    Dim fso As Object
    
    pptPath = ActivePresentation.FullName
    
    ' === EDIT THESE PATHS ===
    scriptPath = "E:\projects\emf2png\ppt_to_pdf.py"
    pythonExe = "E:\projects\emf2png\.venv\Scripts\python.exe"
    ' ========================
    
    command = pythonExe & " """ & scriptPath & """ """ & pptPath & """"
    
    Set shellObj = CreateObject("WScript.Shell")
    shellObj.Run command, 0, True
    
    pdfPath = Left(pptPath, InStrRev(pptPath, ".")) & "pdf"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(pdfPath) Then
        MsgBox "Done! PDF saved to: " & pdfPath, vbInformation, "emf2png"
    Else
        MsgBox "Conversion failed. Try running:" & vbCrLf & vbCrLf & _
               command, vbExclamation, "emf2png"
    End If
End Sub
