#!/usr/bin/env python3
"""
Windows-специфичный скрипт с защитой от завершения
Запускается скрыто и автоматически перезапускается при сбое
"""

import sys
import os
import time
import subprocess
import psutil
from pathlib import Path
import logging

# Константа для скрытого запуска в Windows
if sys.platform == 'win32':
    CREATE_NO_WINDOW = 0x08000000
else:
    CREATE_NO_WINDOW = 0

# Настройка логирования
log_dir = Path(__file__).parent
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / 'dbo_windows_service.log'),
    ]
)

logger = logging.getLogger(__name__)

# Добавляем текущую директорию в путь
sys.path.insert(0, str(Path(__file__).parent))

from dbo_operator_automation import DBOOperatorAutomation
import dbo_operator_config as config


class WindowsService:
    """Windows сервис с защитой от завершения"""
    
    def __init__(self):
        self.script_path = Path(__file__).parent / "start_dbo_automation.py"
        self.pythonw_path = self.find_pythonw()
        self.process = None
        self.restart_delay = 5  # Задержка перед перезапуском (секунды)
        self.max_restarts = 10  # Максимум перезапусков подряд
        self.restart_count = 0
        
    def find_pythonw(self):
        """Поиск pythonw.exe для скрытого запуска"""
        # Пробуем найти pythonw.exe
        python_exe = sys.executable
        if python_exe.endswith('python.exe'):
            pythonw_exe = python_exe.replace('python.exe', 'pythonw.exe')
            if os.path.exists(pythonw_exe):
                return pythonw_exe
        
        # Если не нашли, используем python.exe
        return python_exe
    
    def is_process_running(self):
        """Проверка, запущен ли процесс"""
        if self.process is None:
            return False
        
        try:
            # Проверяем, существует ли процесс
            process = psutil.Process(self.process.pid)
            return process.is_running()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return False
    
    def start_process(self):
        """Запуск процесса скрипта"""
        try:
            logger.info(f"Запуск процесса: {self.pythonw_path} {self.script_path}")
            
            # Запускаем скрипт скрыто (без окна консоли)
            if self.pythonw_path.endswith('pythonw.exe'):
                # pythonw.exe запускает скрипт без окна
                self.process = subprocess.Popen(
                    [self.pythonw_path, str(self.script_path)],
                    creationflags=CREATE_NO_WINDOW,
                    cwd=str(Path(__file__).parent),
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            else:
                # Если pythonw.exe нет, используем python.exe с CREATE_NO_WINDOW
                self.process = subprocess.Popen(
                    [self.pythonw_path, str(self.script_path)],
                    creationflags=CREATE_NO_WINDOW,
                    cwd=str(Path(__file__).parent),
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            
            logger.info(f"Процесс запущен с PID: {self.process.pid}")
            self.restart_count = 0  # Сбрасываем счетчик перезапусков
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при запуске процесса: {e}")
            return False
    
    def stop_process(self):
        """Остановка процесса"""
        if self.process:
            try:
                logger.info(f"Остановка процесса PID: {self.process.pid}")
                self.process.terminate()
                try:
                    self.process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    self.process.kill()
                logger.info("Процесс остановлен")
            except Exception as e:
                logger.error(f"Ошибка при остановке процесса: {e}")
            finally:
                self.process = None
    
    def run(self):
        """Основной цикл работы сервиса"""
        logger.info("Запуск Windows сервиса автоматизации ДБО")
        logger.info(f"Python: {self.pythonw_path}")
        logger.info(f"Скрипт: {self.script_path}")
        
        try:
            while True:
                # Проверяем, запущен ли процесс
                if not self.is_process_running():
                    logger.warning("Процесс не запущен или завершился")
                    
                    # Проверяем лимит перезапусков
                    if self.restart_count >= self.max_restarts:
                        logger.error(f"Достигнут лимит перезапусков ({self.max_restarts}). Остановка сервиса.")
                        break
                    
                    # Перезапускаем процесс
                    logger.info(f"Перезапуск процесса (попытка {self.restart_count + 1}/{self.max_restarts})...")
                    time.sleep(self.restart_delay)
                    
                    if self.start_process():
                        logger.info("Процесс успешно перезапущен")
                    else:
                        self.restart_count += 1
                        logger.error(f"Не удалось перезапустить процесс. Попытка {self.restart_count}/{self.max_restarts}")
                
                # Ждем перед следующей проверкой
                time.sleep(10)  # Проверяем каждые 10 секунд
                
        except KeyboardInterrupt:
            logger.info("Остановка сервиса по запросу пользователя")
        except Exception as e:
            logger.error(f"Критическая ошибка сервиса: {e}")
        finally:
            self.stop_process()
            logger.info("Сервис остановлен")


def main():
    """Главная функция"""
    try:
        service = WindowsService()
        service.run()
    except Exception as e:
        logger.error(f"Фатальная ошибка: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

