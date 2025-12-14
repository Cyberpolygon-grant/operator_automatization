@echo off
REM Простой скрипт для установки автозапуска Python скрипта
REM Создает bat-файл в папке автозагрузки, который запускает dbo_automation.py

echo ========================================
echo Установка автозапуска
echo ========================================
echo.

REM Путь к директории скриптов в Program Files
set "SCRIPT_DIR=C:\Program Files\operator_automatization-main\1"
set "SCRIPT_PATH=%SCRIPT_DIR%\dbo_automation.py"
set "START_SCRIPT=%SCRIPT_DIR%\start_automation.bat"
set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "STARTUP_BAT=%STARTUP_DIR%\dbo_automation_startup.bat"

REM Проверяем наличие Python скрипта
if not exist "%SCRIPT_PATH%" (
    echo [ERROR] Файл не найден: %SCRIPT_PATH%
    echo.
    echo Убедитесь, что файл dbo_automation.py находится в:
    echo %SCRIPT_DIR%
    pause
    exit /b 1
)

REM Проверяем наличие start_automation.bat
if not exist "%START_SCRIPT%" (
    echo [ERROR] Файл не найден: %START_SCRIPT%
    echo.
    echo Убедитесь, что файл start_automation.bat находится в:
    echo %SCRIPT_DIR%
    pause
    exit /b 1
)

REM Создаем папку автозагрузки, если не существует
if not exist "%STARTUP_DIR%" (
    mkdir "%STARTUP_DIR%"
)

REM Создаем bat-файл в папке автозагрузки, который запускает start_automation.bat
(
    echo @echo off
    echo cd /d "C:\Program Files\operator_automatization-main\1"
    echo start "" "C:\Program Files\operator_automatization-main\1\start_automation.bat"
) > "%STARTUP_BAT%"

if %ERRORLEVEL% EQU 0 (
    echo [SUCCESS] Автозапуск успешно установлен!
    echo.
    echo Файл создан: %STARTUP_BAT%
    echo.
    echo Python скрипт будет запускаться при входе в систему
    echo Скрипт автоматически перезапускается при завершении
    echo.
    echo Для удаления автозапуска:
    echo 1. Нажмите Win+R
    echo 2. Введите: shell:startup
    echo 3. Удалите файл dbo_automation_startup.bat
    echo.
    echo Для остановки процесса:
    echo - Завершите процесс cmd.exe (start_automation.bat) в Диспетчере задач
    echo.
) else (
    echo [ERROR] Не удалось создать файл автозапуска
    echo.
)

pause
