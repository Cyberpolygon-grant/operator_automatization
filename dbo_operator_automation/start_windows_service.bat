@echo off
chcp 65001 >nul
REM Запуск Windows сервиса автоматизации ДБО
REM Запускается скрыто и защищен от завершения

echo ========================================
echo АВТОМАТИЗАЦИЯ ОПЕРАТОРА ДБО
echo ========================================
echo.

REM Переход в директорию скрипта
cd /d "%~dp0"
echo Текущая директория: %CD%
echo.

REM Проверка наличия Python
echo Проверка Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не найден!
    echo Установите Python 3.11 или выше
    pause
    exit /b 1
)
python --version
echo.

REM Проверка наличия psutil
echo Проверка psutil...
python -c "import psutil" >nul 2>&1
if errorlevel 1 (
    echo Установка psutil...
    pip install psutil
    echo.
) else (
    echo psutil установлен
    echo.
)

echo Запуск сервиса в фоновом режиме...
echo.

REM Запуск сервиса скрыто (без окна консоли)
REM Используем pythonw.exe если доступен, иначе python.exe
where pythonw.exe >nul 2>&1
if errorlevel 1 (
    REM pythonw.exe не найден, используем python.exe с запуском в фоне
    echo Используется python.exe
    start /B "" python dbo_windows_service.py
) else (
    REM Используем pythonw.exe для полностью скрытого запуска
    echo Используется pythonw.exe
    start /B "" pythonw dbo_windows_service.py
)

timeout /t 2 /nobreak >nul

echo ========================================
echo Сервис автоматизации ДБО запущен!
echo ========================================
echo.
echo Сервис работает в фоновом режиме
echo Логи сохраняются в: dbo_windows_service.log
echo Файлы скачиваются в: Downloads\
echo.
echo Для остановки используйте: stop_windows_service.bat
echo или завершите процесс через Диспетчер задач
echo.
pause

