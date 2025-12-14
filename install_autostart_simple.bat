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

REM Копируем VBS скрипт (скрытый запуск) в папку автозагрузки
set VBS_FILE=%SCRIPT_DIR%start_automation_hidden.vbs
if exist "%VBS_FILE%" (
    copy "%VBS_FILE%" "%STARTUP_DIR%\" >nul
    set INSTALLED_FILE=start_automation_hidden.vbs
) else (
    REM Если VBS нет, копируем bat-файл
    copy "%BAT_FILE%" "%STARTUP_DIR%\" >nul
    set INSTALLED_FILE=start_automation.bat
)

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo [SUCCESS] Автозапуск успешно установлен!
    echo ========================================
    echo.
    echo Файл скопирован в папку автозагрузки:
    echo %STARTUP_DIR%
    echo Файл: %INSTALLED_FILE%
    echo.
    echo Скрипт будет запускаться СКРЫТО при входе в систему
    echo Окно не будет видно, но скрипт будет работать
    echo.
    echo Для удаления автозапуска:
    echo 1. Нажмите Win+R
    echo 2. Введите: shell:startup
    echo 3. Удалите файл %INSTALLED_FILE%
    echo.
) else (
    echo.
    echo [ERROR] Не удалось скопировать файл в папку автозагрузки
    echo.
)

pause
