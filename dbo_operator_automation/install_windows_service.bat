@echo off
REM Установка Windows сервиса через планировщик задач
REM Требует прав администратора

echo ========================================
echo УСТАНОВКА СЛУЖБЫ АВТОМАТИЗАЦИИ ДБО
echo ========================================
echo.

REM Проверка прав администратора
net session >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Требуются права администратора!
    echo Запустите этот скрипт от имени администратора
    pause
    exit /b 1
)

REM Получаем путь к скрипту
set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%SCRIPT_DIR%start_windows_service.bat

REM Создаем задачу в планировщике задач
echo Создание задачи в планировщике задач...
schtasks /Create /TN "DBO Operator Automation" /TR "\"%SCRIPT_PATH%\"" /SC ONLOGON /RU SYSTEM /F >nul 2>&1

if errorlevel 1 (
    echo ОШИБКА: Не удалось создать задачу
    pause
    exit /b 1
)

echo.
echo Задача успешно создана!
echo.
echo Задача будет запускаться при входе пользователя в систему
echo.
echo Для управления задачей используйте:
echo   schtasks /Query /TN "DBO Operator Automation"  - просмотр задачи
echo   schtasks /Run /TN "DBO Operator Automation"   - запуск задачи
echo   schtasks /Delete /TN "DBO Operator Automation" /F  - удаление задачи
echo.
pause

