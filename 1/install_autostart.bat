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
set "HIDDEN_SCRIPT=%SCRIPT_DIR%\start_automation_hidden.vbs"
set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "STARTUP_VBS=%STARTUP_DIR%\dbo_automation_startup.vbs"

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

REM Проверяем наличие start_automation_hidden.vbs
if not exist "%HIDDEN_SCRIPT%" (
    echo [ERROR] Файл не найден: %HIDDEN_SCRIPT%
    echo.
    echo Убедитесь, что файл start_automation_hidden.vbs находится в:
    echo %SCRIPT_DIR%
    pause
    exit /b 1
)

REM Создаем папку автозагрузки, если не существует
if not exist "%STARTUP_DIR%" (
    mkdir "%STARTUP_DIR%"
)

REM Создаем ярлык для запуска VBS скрипта (чтобы избежать запроса о доверии)
set "STARTUP_LNK=%STARTUP_DIR%\dbo_automation_startup.lnk"
set "VBS_PATH=%SCRIPT_DIR%\start_automation_hidden.vbs"

REM Используем PowerShell для создания ярлыка
powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%STARTUP_LNK%'); $Shortcut.TargetPath = 'wscript.exe'; $Shortcut.Arguments = '\"\"%VBS_PATH%\"\"'; $Shortcut.WorkingDirectory = '%SCRIPT_DIR%'; $Shortcut.WindowStyle = 7; $Shortcut.Save()"

if %ERRORLEVEL% EQU 0 (
    echo [SUCCESS] Автозапуск успешно установлен!
    echo.
    echo Ярлык создан: %STARTUP_LNK%
    echo.
    echo Python скрипт будет запускаться СКРЫТО при входе в систему
    echo Скрипт автоматически перезапускается при завершении
    echo Консоль не будет видна пользователю
    echo Система не будет спрашивать о доверии (используется ярлык)
    echo.
    echo Для удаления автозапуска:
    echo - Запустите: uninstall_autostart.bat
    echo Или вручную:
    echo 1. Нажмите Win+R
    echo 2. Введите: shell:startup
    echo 3. Удалите файл dbo_automation_startup.lnk
    echo.
    echo Для остановки процесса:
    echo - Завершите процесс cmd.exe (start_automation.bat) в Диспетчере задач
    echo.
) else (
    echo [ERROR] Не удалось создать ярлык автозапуска
    echo.
)

pause
