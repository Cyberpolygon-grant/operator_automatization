@echo off
REM Скрипт-обертка для запуска Python скрипта с автоперезапуском
REM При завершении скрипта автоматически перезапускается

REM Переходим в директорию скрипта
cd /d "C:\Program Files\operator_automatization-main\1"

REM Бесконечный цикл с автоперезапуском
:LOOP
    REM Запускаем Python скрипт и ждем его завершения
    python dbo_automation.py
    
    REM Если скрипт завершился, ждем 3 секунды и перезапускаем
    timeout /t 3 /nobreak >nul
    goto LOOP
