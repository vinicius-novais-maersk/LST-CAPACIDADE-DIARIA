Dim shell
Dim fso
Dim baseDir
Dim psScript
Dim command

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
psScript = baseDir & "\run_lst_capacidade_diaria.ps1"
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psScript & """"

shell.Run command, 0, True
