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

REM Создаем специальный bat-файл для автозагрузки с абсолютными путями
set HIDDEN_BAT=%SCRIPT_DIR%start_automation_hidden.bat
set STARTUP_BAT=%STARTUP_DIR%\start_automation_hidden.bat

REM Получаем абсолютный путь (убираем завершающий слэш)
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

REM Получаем абсолютный путь (убираем завершающий слэш, если есть)
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

REM Создаем bat-файл в папке автозагрузки с абсолютным путем к скрипту
(
    echo @echo off
    echo REM Автозапуск автоматизации оператора ДБО
    echo REM Создан автоматически установщиком
    echo.
    echo REM Устанавливаем кодировку UTF-8
    echo chcp 65001 ^>nul 2^>^&1
    echo.
    echo REM Абсолютный путь к директории скрипта
    echo set "SCRIPT_DIR=%SCRIPT_DIR%"
    echo.
    echo REM Переходим в директорию скрипта
    echo cd /d "%%SCRIPT_DIR%%"
    echo.
    echo REM Проверяем наличие файла
    echo if not exist "%%SCRIPT_DIR%%\start_automation.bat" ^(
    echo     exit /b 1
    echo ^)
    echo.
    echo REM Запускаем основной скрипт в скрытом режиме
    echo start "" /MIN "%%SCRIPT_DIR%%\start_automation.bat" HIDDEN
    echo.
    echo REM Выходим сразу
    echo exit /b 0
) > "%STARTUP_BAT%"

set INSTALLED_FILE=start_automation_hidden.bat

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
    echo Окно будет свернуто в трей, скрипт будет работать в фоне
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
