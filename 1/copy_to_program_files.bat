@echo off
REM Скрипт для копирования всех файлов в Program Files
REM Требует прав администратора

echo ========================================
echo Копирование файлов в Program Files
echo ========================================
echo.

REM Проверяем права администратора
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Этот скрипт требует прав администратора!
    echo.
    echo Запустите этот файл от имени администратора:
    echo Правой кнопкой мыши - Запуск от имени администратора
    echo.
    pause
    exit /b 1
)

REM Путь назначения
set "TARGET_DIR=C:\Program Files\operator_automatization-main\1"

REM Создаем целевую директорию, если не существует
if not exist "%TARGET_DIR%" (
    echo Создание директории: %TARGET_DIR%
    mkdir "%TARGET_DIR%"
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Не удалось создать директорию!
        pause
        exit /b 1
    )
)

REM Копируем все необходимые файлы
echo.
echo Копирование файлов...
echo.

REM Копируем Python скрипт
if exist "dbo_automation.py" (
    copy /Y "dbo_automation.py" "%TARGET_DIR%\" >nul
    if %ERRORLEVEL% EQU 0 (
        echo [OK] dbo_automation.py
    ) else (
        echo [ERROR] Не удалось скопировать dbo_automation.py
    )
) else (
    echo [WARNING] dbo_automation.py не найден в текущей папке
)

REM Копируем bat-файлы
for %%f in (start_automation.bat start_automation_hidden.bat install_autostart_simple.bat stop_automation.bat uninstall_autostart.bat) do (
    if exist "%%f" (
        copy /Y "%%f" "%TARGET_DIR%\" >nul
        if %ERRORLEVEL% EQU 0 (
            echo [OK] %%f
        ) else (
            echo [ERROR] Не удалось скопировать %%f
        )
    ) else (
        echo [WARNING] %%f не найден в текущей папке
    )
)

REM Копируем инструкцию, если есть
if exist "ИНСТРУКЦИЯ_АВТОЗАПУСК.txt" (
    copy /Y "ИНСТРУКЦИЯ_АВТОЗАПУСК.txt" "%TARGET_DIR%\" >nul
    if %ERRORLEVEL% EQU 0 (
        echo [OK] ИНСТРУКЦИЯ_АВТОЗАПУСК.txt
    )
)

echo.
echo ========================================
echo [SUCCESS] Файлы скопированы!
echo ========================================
echo.
echo Все файлы находятся в:
echo %TARGET_DIR%
echo.
echo Теперь запустите install_autostart_simple.bat из этой папки
echo для установки автозапуска.
echo.
pause
