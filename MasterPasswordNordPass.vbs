Dim objShell
Set objShell = CreateObject("WScript.Shell")
strUserName = objShell.ExpandEnvironmentStrings( "%USERNAME%" )
objShell.Run "C:\Users\" & strUserName & "\AppData\Local\Programs\nordpass\NordPass.exe"
Wscript.Sleep 2000
objShell.SendKeys "{CLEAR}" 
objShell.SendKeys "<MasterPassword>"
objShell.SendKeys "{ENTER}"
WScript.Quit
