@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

REM Путь к папке Downloads
set "DOWNLOADS_PATH=%USERPROFILE%\Downloads"

REM Файл для хранения списка уже открытых файлов
set "TRACK_FILE=%TEMP%\xlsm_opened_files.txt"

REM Создаём файл отслеживания, если его нет
if not exist "%TRACK_FILE%" (
    echo. > "%TRACK_FILE%" 2>nul
)

:loop
    REM Проверяем существование папки Downloads
    if not exist "%DOWNLOADS_PATH%" (
        timeout /t 30 /nobreak >nul 2>&1
        goto loop
    )
    
    REM Ищем все .xlsm файлы в папке Downloads
    set "found=0"
    for %%F in ("%DOWNLOADS_PATH%\*.xlsm") do (
        set "file_path=%%F"
        set "file_name=%%~nxF"
        
        REM Проверяем, не открывали ли мы уже этот файл
        findstr /C:"!file_path!" "%TRACK_FILE%" >nul 2>&1
        if errorlevel 1 (
            REM Файл ещё не открывался - открываем его
            REM Открываем файл через start (скрыто)
            start "" "!file_path!" >nul 2>&1
            
            REM Добавляем файл в список открытых
            echo !file_path! >> "%TRACK_FILE%" 2>nul
            
            set "found=1"
        )
    )
    
    REM Очищаем список от несуществующих файлов (каждые 10 итераций)
    set /a "iterations+=1"
    if !iterations! geq 10 (
        set "iterations=0"
        (
            for /f "usebackq delims=" %%F in ("%TRACK_FILE%") do (
                if exist "%%F" echo %%F
            )
        ) > "%TRACK_FILE%.tmp" 2>nul
        move /y "%TRACK_FILE%.tmp" "%TRACK_FILE%" >nul 2>&1
    )
    
    timeout /t 30 /nobreak >nul 2>&1
    
goto loop
