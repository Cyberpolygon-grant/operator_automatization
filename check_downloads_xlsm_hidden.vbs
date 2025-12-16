Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Получаем путь к батнику (в той же папке, где находится VBS)
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = scriptPath & "\check_downloads_xlsm.bat"

' Запускаем батник скрыто
WshShell.Run """" & batPath & """", 0, False

Set WshShell = Nothing
Set fso = Nothing
