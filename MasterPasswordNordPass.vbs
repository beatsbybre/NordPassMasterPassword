Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "C:\Users\%username%\AppData\Local\Programs\nordpass\NordPass.exe"
Wscript.Sleep 2000
objShell.SendKeys "{CLEAR}" 
objShell.SendKeys "<MasterPassword>"
objShell.SendKeys "{ENTER}"
WScript.Quit
