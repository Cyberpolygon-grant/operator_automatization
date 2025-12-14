@echo off
REM Скрипт для остановки автоматизации оператора ДБО

echo ========================================
echo Остановка автоматизации оператора ДБО
echo ========================================
echo.

REM Останавливаем все процессы Python, связанные с dbo_automation
echo Остановка процессов Python (dbo_automation)...
taskkill /F /FI "WINDOWTITLE eq *dbo_automation*" /IM python.exe >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq *Автоматизация оператора ДБО*" /IM python.exe >nul 2>&1
taskkill /F /IM python.exe /FI "WINDOWTITLE eq *dbo_automation*" >nul 2>&1

REM Останавливаем скрытые bat-скрипты
echo Остановка скрытых скриптов...
taskkill /F /FI "WINDOWTITLE eq *start_automation_hidden*" /IM cmd.exe >nul 2>&1

REM Останавливаем bat-файлы
echo Остановка bat-скриптов...
taskkill /F /FI "WINDOWTITLE eq *Автоматизация оператора ДБО*" /IM cmd.exe >nul 2>&1

REM Останавливаем все процессы cmd.exe, которые запускают start_automation.bat
for /f "tokens=2" %%i in ('tasklist /FI "IMAGENAME eq cmd.exe" /FO LIST ^| findstr /I "PID"') do (
    wmic process where "ProcessId=%%i" get CommandLine 2>nul | findstr /I "start_automation" >nul
    if !errorlevel! equ 0 (
        taskkill /F /PID %%i >nul 2>&1
    )
)

echo.
echo ========================================
echo Автоматизация остановлена
echo ========================================
echo.
echo Примечание: Если скрипт в автозагрузке, он запустится снова
echo при следующем входе в систему
echo.
pause
