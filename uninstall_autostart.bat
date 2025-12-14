@echo off
REM Скрипт для удаления автозапуска автоматизации оператора ДБО

echo ========================================
echo Удаление автозапуска автоматизации ДБО
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

set TASK_NAME=DBOOperatorAutomation

echo Удаляем задачу из планировщика Windows...
schtasks /Delete /TN "%TASK_NAME%" /F

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo [SUCCESS] Автозапуск успешно удален!
    echo ========================================
    echo.
) else (
    echo.
    echo [WARNING] Задача не найдена или уже удалена
    echo.
)

pause
