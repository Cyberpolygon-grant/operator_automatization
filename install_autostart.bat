@echo off
REM Скрипт для установки автоматизации оператора ДБО в автозагрузку Windows

echo ========================================
echo Установка автозапуска автоматизации ДБО
echo ========================================
echo.

REM Проверяем права администратора
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Требуются права администратора!
    echo Запустите этот файл от имени администратора
    echo.
    pause
    exit /b 1
)

REM Получаем путь к директории скрипта
set SCRIPT_DIR=%~dp0
set BAT_FILE=%SCRIPT_DIR%start_automation.bat
set TASK_NAME=DBOOperatorAutomation

REM Проверяем наличие bat-файла
if not exist "%BAT_FILE%" (
    echo [ERROR] Файл не найден: %BAT_FILE%
    pause
    exit /b 1
)

echo Удаляем существующую задачу (если есть)...
schtasks /Delete /TN "%TASK_NAME%" /F >nul 2>&1

echo Создаем задачу в планировщике Windows...
echo.

REM Получаем имя текущего пользователя для запуска от его имени
for /f "tokens=2" %%a in ('whoami /user /fo list ^| findstr /i "SID"') do set USER_SID=%%a
for /f "tokens=1" %%a in ('whoami') do set CURRENT_USER=%%a

REM Создаем задачу в планировщике (запуск от имени текущего пользователя)
schtasks /Create /TN "%TASK_NAME%" /TR "\"%BAT_FILE%\"" /SC ONLOGON /RL HIGHEST /F /RU "%CURRENT_USER%"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo [SUCCESS] Автозапуск успешно установлен!
    echo ========================================
    echo.
    echo Задача создана: %TASK_NAME%
    echo Запуск: При входе пользователя в систему
    echo Пользователь: %CURRENT_USER%
    echo Приоритет: Высший
    echo.
    echo Для удаления автозапуска запустите: uninstall_autostart.bat
    echo.
    echo Проверка задачи:
    schtasks /Query /TN "%TASK_NAME%" /FO LIST /V | findstr /i "TaskName Status"
    echo.
) else (
    echo.
    echo [ERROR] Не удалось создать задачу в планировщике
    echo Код ошибки: %ERRORLEVEL%
    echo.
    echo Попробуйте альтернативный метод:
    echo 1. Нажмите Win+R
    echo 2. Введите: shell:startup
    echo 3. Скопируйте туда файл: start_automation.bat
    echo.
    echo Или создайте ярлык на start_automation.bat в папке автозагрузки
    echo.
)

pause
