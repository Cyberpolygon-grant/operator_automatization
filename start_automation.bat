@echo off
REM Скрипт автозапуска автоматизации оператора ДБО
REM Автоматически перезапускается при ошибках и защищен от завершения
REM Запускается в скрытом режиме через VBS скрипт

REM Если запущен в скрытом режиме
if "%1"=="HIDDEN" (
    REM Скрытый режим - убираем заголовок
    title 
    REM Перенаправляем вывод в файл лога вместо консоли
    set LOG_FILE=%~dp0automation.log
) else (
    title Автоматизация оператора ДБО
    color 0A
    set LOG_FILE=CON
)

REM Получаем путь к директории скрипта
cd /d "%~dp0"

REM Путь к Python скрипту (используем абсолютный путь)
set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%SCRIPT_DIR%dbo_automation.py
set PYTHON_CMD=python

REM Если python не найден, пробуем python3
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    where python3 >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        set PYTHON_CMD=python3
    )
)

REM Проверяем наличие Python
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python не найден в PATH!
    echo Установите Python или укажите полный путь к python.exe
    pause
    exit /b 1
)

REM Проверяем наличие скрипта
if not exist "%SCRIPT_PATH%" (
    echo [ERROR] Скрипт не найден: %SCRIPT_PATH%
    pause
    exit /b 1
)

echo ========================================
echo Автоматизация оператора ДБО
echo Автозапуск с защитой от завершения
echo ========================================
echo.
echo Скрипт: %SCRIPT_PATH%
echo Python: %PYTHON_CMD%
echo.
echo Для остановки закройте это окно
echo ========================================
echo.

REM Счетчик перезапусков
set RESTART_COUNT=0

REM Бесконечный цикл с автоматическим перезапуском
:LOOP
    set /a RESTART_COUNT+=1
    
    REM В скрытом режиме перенаправляем вывод в файл
    if "%1"=="HIDDEN" (
        echo [%DATE% %TIME%] Запуск скрипта автоматизации (попытка #%RESTART_COUNT%)... >> "%LOG_FILE%"
        echo. >> "%LOG_FILE%"
        
        REM Запускаем скрипт и ждем его завершения (вывод в файл)
        %PYTHON_CMD% "%SCRIPT_PATH%" >> "%LOG_FILE%" 2>&1
        set EXIT_CODE=%ERRORLEVEL%
        
        echo. >> "%LOG_FILE%"
        echo [%DATE% %TIME%] Скрипт завершился с кодом: %EXIT_CODE% >> "%LOG_FILE%"
        
        REM В скрытом режиме всегда перезапускаем (даже при коде 0)
        REM Это нужно для защиты от закрытия окна
        echo [Скрытый режим] Перезапуск через 3 секунды... >> "%LOG_FILE%"
        timeout /t 3 /nobreak >nul
        goto LOOP
    ) else (
        REM Видимый режим - вывод в консоль
        echo [%DATE% %TIME%] Запуск скрипта автоматизации (попытка #%RESTART_COUNT%)...
        echo.
        
        REM Запускаем скрипт и ждем его завершения
        %PYTHON_CMD% "%SCRIPT_PATH%"
        set EXIT_CODE=%ERRORLEVEL%
        
        echo.
        echo ========================================
        echo [%DATE% %TIME%] Скрипт завершился с кодом: %EXIT_CODE%
        
        REM Если код выхода 0 - нормальное завершение (Ctrl+C)
        if %EXIT_CODE% EQU 0 (
            echo Нормальное завершение работы
            echo Всего перезапусков: %RESTART_COUNT%
            echo.
            pause
            exit /b 0
        )
        
        REM Иначе - ошибка, перезапускаем
        echo Ошибка! Автоматический перезапуск через 5 секунд...
        echo Всего перезапусков: %RESTART_COUNT%
        echo Нажмите Ctrl+C для остановки
        echo ========================================
        echo.
        
        REM Ждем 5 секунд перед перезапуском (нельзя прервать)
        timeout /t 5 /nobreak >nul
        
        REM Очищаем экран и перезапускаем
        cls
        goto LOOP
    )
