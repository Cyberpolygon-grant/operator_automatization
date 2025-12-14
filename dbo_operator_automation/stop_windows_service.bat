@echo off
REM Остановка Windows сервиса автоматизации ДБО

echo Остановка сервиса автоматизации ДБО...
echo.

REM Поиск и завершение процессов
for /f "tokens=2" %%a in ('tasklist /FI "IMAGENAME eq pythonw.exe" /FO LIST ^| findstr /I "PID"') do (
    echo Проверка процесса PID: %%a
    wmic process where "ProcessId=%%a" get CommandLine | findstr /I "dbo_windows_service" >nul
    if not errorlevel 1 (
        echo Завершение процесса PID: %%a
        taskkill /PID %%a /F >nul 2>&1
    )
)

for /f "tokens=2" %%a in ('tasklist /FI "IMAGENAME eq python.exe" /FO LIST ^| findstr /I "PID"') do (
    echo Проверка процесса PID: %%a
    wmic process where "ProcessId=%%a" get CommandLine | findstr /I "dbo_windows_service" >nul
    if not errorlevel 1 (
        echo Завершение процесса PID: %%a
        taskkill /PID %%a /F >nul 2>&1
    )
)

echo.
echo Сервис остановлен
pause

