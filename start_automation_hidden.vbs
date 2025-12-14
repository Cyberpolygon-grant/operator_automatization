' VBS скрипт для скрытого запуска автоматизации оператора ДБО
' Запускает bat-файл в скрытом режиме с автоматическим перезапуском

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Получаем путь к директории скрипта
ScriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
BatFile = ScriptDir & "\start_automation.bat"

' Проверяем наличие bat-файла
If Not fso.FileExists(BatFile) Then
    ' В скрытом режиме не показываем сообщение, просто выходим
    WScript.Quit
End If

' Счетчик перезапусков
RestartCount = 0

' Бесконечный цикл с перезапуском
Do
    RestartCount = RestartCount + 1
    
    ' Запускаем bat-файл в скрытом режиме (0 = скрыто) с параметром HIDDEN
    ' Используем Run с параметром WaitOnReturn = True, чтобы ждать завершения
    ProcessID = WshShell.Run("cmd.exe /c """ & BatFile & """ HIDDEN", 0, True)
    
    ' Процесс завершен - перезапускаем через 3 секунды
    WScript.Sleep 3000
Loop

' Освобождаем объекты
Set WshShell = Nothing
Set fso = Nothing
