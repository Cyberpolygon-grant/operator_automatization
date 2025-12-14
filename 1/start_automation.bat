@echo off
REM Скрипт-обертка для запуска Python скрипта с автоперезапуском
REM При завершении скрипта автоматически перезапускается
REM Консоль скрыта от пользователя

REM Переходим в директорию скрипта
cd /d "C:\Program Files\operator_automatization-main\1"

REM Скрываем окно консоли
if not "%1"=="HIDDEN" (
    REM Запускаем себя в скрытом режиме через VBS
    set "VBS_TEMP=%TEMP%\start_hidden_temp.vbs"
    (
        echo Set WshShell = CreateObject("WScript.Shell"^)
        echo WshShell.CurrentDirectory = "C:\Program Files\operator_automatization-main\1"
        echo WshShell.Run Chr^(34^) ^& "C:\Program Files\operator_automatization-main\1\start_automation.bat" ^& Chr^(34^) ^& " HIDDEN", 0, False
        echo Set WshShell = Nothing
    ) > "%VBS_TEMP%"
    cscript //nologo "%VBS_TEMP%"
    del "%VBS_TEMP%" >nul 2>&1
    exit /b
)

REM Бесконечный цикл с автоперезапуском (скрытый режим)
:LOOP
    REM Запускаем Python скрипт скрыто (pythonw.exe или через start /B)
    REM Пробуем pythonw.exe (не показывает консоль)
    where pythonw.exe >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        pythonw.exe dbo_automation.py
    ) else (
        REM Если pythonw нет, используем python через start /B (скрытый режим)
        start /B "" python dbo_automation.py
        REM Ждем завершения процесса
        :WAIT_LOOP
        tasklist /FI "IMAGENAME eq python.exe" /FI "COMMANDLINE eq *dbo_automation.py*" 2>nul | find /I "python.exe" >nul
        if %ERRORLEVEL% EQU 0 (
            timeout /t 1 /nobreak >nul
            goto WAIT_LOOP
        )
    )
    
    REM Если скрипт завершился, ждем 3 секунды и перезапускаем
    timeout /t 3 /nobreak >nul
    goto LOOP
