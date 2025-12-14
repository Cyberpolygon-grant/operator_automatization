@echo off
REM Скрипт для скрытого запуска автоматизации оператора ДБО
REM Запускается через автозагрузку без запроса о доверии

REM Устанавливаем кодировку UTF-8
chcp 65001 >nul 2>&1

REM Если в переменной SCRIPT_DIR уже установлен путь (из установщика), используем его
if not defined SCRIPT_DIR (
    REM Получаем путь к директории скрипта (где находится этот bat-файл)
    set "SCRIPT_DIR=%~dp0"
    
    REM Если скрипт запущен из папки автозагрузки, ищем оригинальную директорию
    REM Проверяем, есть ли start_automation.bat в текущей директории
    if not exist "%SCRIPT_DIR%start_automation.bat" (
        REM Если файла нет, пробуем найти его в стандартных местах
        if exist "%USERPROFILE%\Downloads\mail\dbo_operator_automation\start_automation.bat" (
            set "SCRIPT_DIR=%USERPROFILE%\Downloads\mail\dbo_operator_automation\"
        ) else if exist "C:\Program Files\operator_automatization-main\1\start_automation.bat" (
            set "SCRIPT_DIR=C:\Program Files\operator_automatization-main\1\"
        ) else if exist "C:\Users\%USERNAME%\Downloads\mail\dbo_operator_automation\start_automation.bat" (
            set "SCRIPT_DIR=C:\Users\%USERNAME%\Downloads\mail\dbo_operator_automation\"
        ) else (
            REM Если не нашли, выходим без ошибки (чтобы не показывать окно)
            exit /b 0
        )
    )
)

REM Убираем завершающий слэш, если есть
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

REM Переходим в директорию скрипта
cd /d "%SCRIPT_DIR%"

REM Проверяем наличие файла перед запуском
if not exist "%SCRIPT_DIR%\start_automation.bat" (
    exit /b 1
)

REM Запускаем основной скрипт в скрытом режиме (окно свернуто)
REM Используем start /MIN для минимизации окна
start "" /MIN "%SCRIPT_DIR%\start_automation.bat" HIDDEN

REM Выходим сразу
exit /b 0
