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

REM Останавливаем скрытые bat-скрипты
echo Остановка скрытых скриптов...
taskkill /F /FI "WINDOWTITLE eq *start_automation_hidden*" /IM cmd.exe >nul 2>&1

REM Останавливаем bat-файлы
echo Остановка bat-скриптов...
taskkill /F /FI "WINDOWTITLE eq *Автоматизация оператора ДБО*" /IM cmd.exe >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq *Охранник автоматизации ДБО*" /IM cmd.exe >nul 2>&1

echo.
echo ========================================
echo Автоматизация остановлена
echo ========================================
echo.
echo Примечание: Если скрипт в автозагрузке, он запустится снова
echo при следующем входе в систему
echo.
pause
