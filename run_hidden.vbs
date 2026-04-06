' run_hidden.vbs
' Lanza converter_ui.py en segundo plano sin ventana de consola visible.
' Este script es llamado por la entrada de registro de inicio automatico.

Dim shell
Set shell = CreateObject("WScript.Shell")

' Obtener la carpeta donde esta este .vbs (misma que el proyecto)
Dim scriptDir
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Construir el comando: busca primero el Python del venv, luego el del sistema
Dim pythonExe
Dim venvPython
venvPython = scriptDir & "\venv\Scripts\pythonw.exe"

If CreateObject("Scripting.FileSystemObject").FileExists(venvPython) Then
    pythonExe = """" & venvPython & """"
Else
    pythonExe = "pythonw.exe"
End If

Dim targetScript
targetScript = """" & scriptDir & "\converter_ui.py" & """"

' 0 = ventana oculta, False = no esperar a que termine
shell.Run pythonExe & " " & targetScript, 0, False

Set shell = Nothing
