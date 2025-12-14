@echo off
REM Запуск Windows сервиса автоматизации ДБО
REM Запускается скрыто и защищен от завершения

REM Переход в директорию скрипта
cd /d "%~dp0"

REM Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не найден!
    echo Установите Python 3.11 или выше
    pause
    exit /b 1
)

REM Проверка наличия psutil
python -c "import psutil" >nul 2>&1
if errorlevel 1 (
    echo Установка psutil...
    pip install psutil
)

REM Запуск сервиса скрыто (без окна консоли)
REM Используем pythonw.exe если доступен, иначе python.exe
where pythonw.exe >nul 2>&1
if errorlevel 1 (
    REM pythonw.exe не найден, используем python.exe с запуском в фоне
    start /B "" python dbo_windows_service.py
) else (
    REM Используем pythonw.exe для полностью скрытого запуска
    pythonw dbo_windows_service.py
)

echo.
echo Сервис автоматизации ДБО запущен в фоновом режиме
echo Логи сохраняются в dbo_windows_service.log
echo.
echo Для остановки используйте stop_windows_service.bat
echo или завершите процесс через Диспетчер задач
echo.

