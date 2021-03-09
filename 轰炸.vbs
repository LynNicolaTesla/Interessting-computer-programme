Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate"…Ú‹À"
for i = 1 to 20
WScript.Sleep 99
WshShell.SendKeys"^v"
WshShell.SendKeys i
WshShell.SendKeys"%s"
next
