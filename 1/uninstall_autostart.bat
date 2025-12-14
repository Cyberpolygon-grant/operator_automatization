@echo off
REM Скрипт для удаления автозапуска

echo ========================================
echo Удаление автозапуска
echo ========================================
echo.

set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "STARTUP_BAT=%STARTUP_DIR%\start_automation_hidden.bat"

REM Останавливаем все процессы автоматизации
echo Остановка процессов автоматизации...
call "%~dp0stop_automation.bat" >nul 2>&1

REM Удаляем файл из папки автозагрузки
if exist "%STARTUP_BAT%" (
    del /F /Q "%STARTUP_BAT%"
    if %ERRORLEVEL% EQU 0 (
        echo [SUCCESS] Файл удален из папки автозагрузки
    ) else (
        echo [ERROR] Не удалось удалить файл из папки автозагрузки
    )
) else (
    echo [INFO] Файл не найден в папке автозагрузки
)

echo.
echo ========================================
echo Автозапуск удален
echo ========================================
echo.
pause
