Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Program Files\operator_automatization-main\1"
WshShell.Run Chr(34) & "C:\Program Files\operator_automatization-main\1\start_automation.bat" & Chr(34), 0, False
Set WshShell = Nothing
