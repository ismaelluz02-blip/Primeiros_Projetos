Option Explicit

Dim shell, fso, baseDir, appDir, appScript, pyCmd, cmd
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
appDir = baseDir
appScript = fso.BuildPath(appDir, "sistema_faturamento.py")

If Not fso.FileExists(appScript) Then
    MsgBox "Arquivo nao encontrado: " & appScript, vbExclamation, "Sistema de Faturamento"
    WScript.Quit 1
End If

shell.CurrentDirectory = appDir
pyCmd = ""

If fso.FileExists(fso.BuildPath(appDir, ".venv\Scripts\pythonw.exe")) Then
    pyCmd = Quote(fso.BuildPath(appDir, ".venv\Scripts\pythonw.exe"))
ElseIf fso.FileExists(fso.BuildPath(appDir, ".venv\Scripts\python.exe")) Then
    pyCmd = Quote(fso.BuildPath(appDir, ".venv\Scripts\python.exe"))
ElseIf ComandoDisponivel("pythonw") Then
    pyCmd = "pythonw"
ElseIf ComandoDisponivel("pyw") Then
    pyCmd = "pyw"
ElseIf ComandoDisponivel("py") Then
    pyCmd = "py"
ElseIf ComandoDisponivel("python") Then
    pyCmd = "python"
Else
    MsgBox "Nao foi encontrado Python para iniciar o sistema." & vbCrLf & _
           "Instale o Python ou ajuste o launcher.", vbExclamation, "Sistema de Faturamento"
    WScript.Quit 2
End If

cmd = pyCmd & " " & Quote(appScript)
shell.Run cmd, 0, False
WScript.Quit 0

Function Quote(s)
    Quote = Chr(34) & s & Chr(34)
End Function

Function ComandoDisponivel(nome)
    On Error Resume Next
    Dim exec
    Set exec = shell.Exec("cmd /c where " & nome & " >nul 2>nul")
    If Err.Number <> 0 Then
        ComandoDisponivel = False
        Err.Clear
        Exit Function
    End If

    Do While exec.Status = 0
        WScript.Sleep 50
    Loop

    ComandoDisponivel = (exec.ExitCode = 0)
End Function
