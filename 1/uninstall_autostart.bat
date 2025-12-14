@echo off
REM Скрипт для удаления автозапуска

echo ========================================
echo Удаление автозапуска
echo ========================================
echo.

set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "STARTUP_VBS=%STARTUP_DIR%\dbo_automation_startup.vbs"
set "STARTUP_LNK=%STARTUP_DIR%\dbo_automation_startup.lnk"

REM Останавливаем процессы
echo Остановка процессов...
taskkill /F /FI "WINDOWTITLE eq *start_automation*" /IM cmd.exe >nul 2>&1
taskkill /F /FI "IMAGENAME eq python.exe" /FI "COMMANDLINE eq *dbo_automation.py*" >nul 2>&1

REM Удаляем VBS файл из папки автозагрузки
if exist "%STARTUP_VBS%" (
    del /F /Q "%STARTUP_VBS%"
    if %ERRORLEVEL% EQU 0 (
        echo [OK] Удален файл: dbo_automation_startup.vbs
    ) else (
        echo [ERROR] Не удалось удалить файл: dbo_automation_startup.vbs
    )
) else (
    echo [INFO] Файл dbo_automation_startup.vbs не найден
)

REM Удаляем ярлык из папки автозагрузки
if exist "%STARTUP_LNK%" (
    del /F /Q "%STARTUP_LNK%"
    if %ERRORLEVEL% EQU 0 (
        echo [OK] Удален ярлык: dbo_automation_startup.lnk
    ) else (
        echo [ERROR] Не удалось удалить ярлык: dbo_automation_startup.lnk
    )
) else (
    echo [INFO] Ярлык dbo_automation_startup.lnk не найден
)

echo.
echo ========================================
echo Автозапуск удален
echo ========================================
echo.
pause
