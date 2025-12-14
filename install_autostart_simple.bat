@echo off
REM Простой скрипт для установки автозапуска через папку автозагрузки
REM Не требует прав администратора

echo ========================================
echo Установка автозапуска (простой метод)
echo ========================================
echo.

REM Получаем путь к директории скрипта
set SCRIPT_DIR=%~dp0
set BAT_FILE=%SCRIPT_DIR%start_automation.bat
set STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

REM Проверяем наличие bat-файла
if not exist "%BAT_FILE%" (
    echo [ERROR] Файл не найден: %BAT_FILE%
    pause
    exit /b 1
)

REM Создаем папку автозагрузки, если не существует
if not exist "%STARTUP_DIR%" (
    mkdir "%STARTUP_DIR%"
)

REM Копируем bat-файл в папку автозагрузки
copy "%BAT_FILE%" "%STARTUP_DIR%\" >nul

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo [SUCCESS] Автозапуск успешно установлен!
    echo ========================================
    echo.
    echo Файл скопирован в папку автозагрузки:
    echo %STARTUP_DIR%
    echo.
    echo Скрипт будет запускаться при входе в систему
    echo.
    echo Для удаления автозапуска:
    echo 1. Нажмите Win+R
    echo 2. Введите: shell:startup
    echo 3. Удалите файл start_automation.bat
    echo.
) else (
    echo.
    echo [ERROR] Не удалось скопировать файл в папку автозагрузки
    echo.
)

pause
