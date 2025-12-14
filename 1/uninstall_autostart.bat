@echo off
setlocal enabledelayedexpansion
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
for /f "tokens=2" %%i in ('tasklist /FI "IMAGENAME eq python.exe" /FO LIST ^| findstr /I "PID"') do (
    wmic process where "ProcessId=%%i" get CommandLine 2>nul | findstr /I "dbo_automation" >nul
    if !errorlevel! equ 0 (
        taskkill /F /PID %%i >nul 2>&1
    )
)
timeout /t 2 /nobreak >nul

REM Удаляем VBS файл из папки автозагрузки
if exist "%STARTUP_VBS%" (
    attrib -R "%STARTUP_VBS%" >nul 2>&1
    del /F /Q "%STARTUP_VBS%" >nul 2>&1
    if exist "%STARTUP_VBS%" (
        echo [WARNING] Не удалось удалить файл: dbo_automation_startup.vbs
        echo Попробуйте удалить вручную: %STARTUP_VBS%
    ) else (
        echo [OK] Удален файл: dbo_automation_startup.vbs
    )
) else (
    echo [INFO] Файл dbo_automation_startup.vbs не найден
)

REM Удаляем ярлык из папки автозагрузки
if exist "%STARTUP_LNK%" (
    attrib -R "%STARTUP_LNK%" >nul 2>&1
    del /F /Q "%STARTUP_LNK%" >nul 2>&1
    if exist "%STARTUP_LNK%" (
        REM Пробуем через PowerShell
        powershell -Command "Remove-Item -Path '%STARTUP_LNK%' -Force -ErrorAction SilentlyContinue" >nul 2>&1
        if exist "%STARTUP_LNK%" (
            echo [WARNING] Не удалось удалить ярлык: dbo_automation_startup.lnk
            echo Попробуйте удалить вручную: %STARTUP_LNK%
        ) else (
            echo [OK] Удален ярлык: dbo_automation_startup.lnk
        )
    ) else (
        echo [OK] Удален ярлык: dbo_automation_startup.lnk
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
