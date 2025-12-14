# PowerShell скрипт для автозапуска автоматизации оператора ДБО
# Более надежный вариант с лучшей обработкой ошибок

# Получаем путь к директории скрипта
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "dbo_automation.py"

# Проверяем наличие Python
$PythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $PythonCmd) {
    Write-Host "[ERROR] Python не найден в PATH!" -ForegroundColor Red
    Write-Host "Установите Python или укажите полный путь к python.exe"
    Read-Host "Нажмите Enter для выхода"
    exit 1
}

# Проверяем наличие скрипта
if (-not (Test-Path $PythonScript)) {
    Write-Host "[ERROR] Скрипт не найден: $PythonScript" -ForegroundColor Red
    Read-Host "Нажмите Enter для выхода"
    exit 1
}

Write-Host "========================================" -ForegroundColor Green
Write-Host "Автоматизация оператора ДБО" -ForegroundColor Green
Write-Host "Автозапуск с защитой от завершения" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Скрипт: $PythonScript" -ForegroundColor Cyan
Write-Host "Python: $($PythonCmd.Source)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Для остановки закройте это окно или нажмите Ctrl+C" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Бесконечный цикл с автоматическим перезапуском
$RestartDelay = 5
$RestartCount = 0

while ($true) {
    $RestartCount++
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    Write-Host "[$Timestamp] Запуск скрипта автоматизации (попытка #$RestartCount)..." -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Запускаем скрипт и ждем его завершения
        $Process = Start-Process -FilePath $PythonCmd.Source -ArgumentList "`"$PythonScript`"" -Wait -NoNewWindow -PassThru
        
        $ExitCode = $Process.ExitCode
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "[$Timestamp] Скрипт завершился с кодом: $ExitCode" -ForegroundColor Yellow
        
        # Если код выхода 0 - нормальное завершение (Ctrl+C)
        if ($ExitCode -eq 0) {
            Write-Host "Нормальное завершение работы" -ForegroundColor Green
            Write-Host ""
            Read-Host "Нажмите Enter для выхода"
            break
        }
        
        # Иначе - ошибка, перезапускаем
        Write-Host "Ошибка! Автоматический перезапуск через $RestartDelay секунд..." -ForegroundColor Red
        Write-Host "Нажмите Ctrl+C для остановки" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host ""
        
        # Ждем перед перезапуском
        Start-Sleep -Seconds $RestartDelay
        
        # Очищаем экран
        Clear-Host
        
    } catch {
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host ""
        Write-Host "[$Timestamp] Критическая ошибка: $_" -ForegroundColor Red
        Write-Host "Перезапуск через $RestartDelay секунд..." -ForegroundColor Yellow
        Write-Host ""
        
        Start-Sleep -Seconds $RestartDelay
        Clear-Host
    }
}
