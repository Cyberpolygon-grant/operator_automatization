@echo off
REM Скрипт для скрытого запуска автоматизации оператора ДБО
REM Запускается через автозагрузку без запроса о доверии (в отличие от VBS)
REM Все скрипты находятся в C:\Program Files\operator_automatization-main\1\

REM Устанавливаем кодировку UTF-8
chcp 65001 >nul 2>&1

REM Если в переменной SCRIPT_DIR уже установлен путь (из установщика), используем его
if not defined SCRIPT_DIR (
    REM Используем стандартный путь в Program Files
    set "SCRIPT_DIR=C:\Program Files\operator_automatization-main\1"
)

REM Переходим в директорию скрипта
cd /d "%SCRIPT_DIR%"

REM Проверяем наличие файла перед запуском
if not exist "%SCRIPT_DIR%\start_automation.bat" (
    REM Если файла нет в Program Files, выходим без ошибки
    exit /b 1
)

REM Запускаем основной скрипт в скрытом режиме (окно свернуто)
REM Используем start /MIN для минимизации окна
start "" /MIN "%SCRIPT_DIR%\start_automation.bat" HIDDEN

REM Выходим сразу
exit /b 0
