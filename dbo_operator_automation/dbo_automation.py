#!/usr/bin/env python3
"""
Автоматизация оператора ДБО через Docker-контейнер
Скачивает файлы из контейнера phishing-demo и автоматически открывает Excel файлы
Логи выводятся в консоль

ВСЁ В ОДНОМ ФАЙЛЕ - просто запустите: python dbo_automation.py
"""

import os
import time
import subprocess
import platform
import json
import shutil
from pathlib import Path
import logging
from datetime import datetime

# ============================================================================
# КОНФИГУРАЦИЯ - ИЗМЕНИТЕ ПОД СВОИ НАСТРОЙКИ
# ============================================================================

# Путь к директории с файлами из Docker-контейнера
# По умолчанию: ./sent_attachments (относительно скрипта)
CONTAINER_ATTACHMENTS_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "sent_attachments"
)

# Директория для сохранения скачанных файлов
USER_HOME = Path.home()
DOWNLOAD_DIR = str(USER_HOME / "Downloads")

# Интервал проверки новых файлов (в секундах)
CHECK_INTERVAL = 5

# Автоматически открывать Excel файлы
AUTO_OPEN_EXCEL = True

# ============================================================================
# НАСТРОЙКА ЛОГИРОВАНИЯ
# ============================================================================

log_format = '%(asctime)s [%(levelname)-8s] %(message)s'
date_format = '%Y-%m-%d %H:%M:%S'

logging.basicConfig(
    level=logging.INFO,
    format=log_format,
    datefmt=date_format,
    handlers=[
        logging.StreamHandler()  # Только консоль
    ]
)

logger = logging.getLogger(__name__)

# ============================================================================
# КЛАСС АВТОМАТИЗАЦИИ
# ============================================================================

class DBOOperatorAutomation:
    """Автоматизация работы оператора ДБО через Docker-контейнер"""
    
    def __init__(self, container_dir, download_dir="downloaded_attachments"):
        """Инициализация автоматизации"""
        self.container_dir = Path(container_dir)
        self.download_dir = Path(download_dir)
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.processed_files = set()
        
        logger.info(f"Инициализация автоматизации")
        logger.info(f"Директория контейнера: {self.container_dir}")
        logger.info(f"Директория загрузки: {self.download_dir}")
    
    def check_container_directory(self):
        """Проверка существования директории контейнера"""
        if not self.container_dir.exists():
            logger.warning(f"⚠ Директория контейнера не найдена: {self.container_dir}")
            logger.info(f"   Убедитесь, что Docker-контейнер запущен и volume настроен")
            return False
        return True
    
    def get_new_metadata_files(self):
        """Получение списка новых JSON файлов с метаданными"""
        try:
            if not self.container_dir.exists():
                return []
            
            metadata_files = []
            for file_path in self.container_dir.glob("*_metadata.json"):
                if str(file_path) not in self.processed_files:
                    metadata_files.append(file_path)
            
            return sorted(metadata_files)
        except Exception as e:
            logger.error(f"❌ Ошибка при получении списка файлов: {e}")
            return []
    
    def load_email_metadata(self, metadata_file):
        """Загрузка метаданных письма из JSON файла"""
        try:
            with open(metadata_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"❌ Ошибка при загрузке метаданных {metadata_file}: {e}")
            return None
    
    def copy_attachment(self, source_file, target_filename):
        """Копирование файла из контейнера в директорию загрузки"""
        try:
            target_path = self.download_dir / target_filename
            
            # Если файл уже существует, добавляем номер
            counter = 1
            original_path = target_path
            while target_path.exists():
                stem = original_path.stem
                suffix = original_path.suffix
                target_path = self.download_dir / f"{stem}_{counter}{suffix}"
                counter += 1
            
            shutil.copy2(source_file, target_path)
            logger.info(f"   Файл скопирован: {target_path.name}")
            return target_path
        except Exception as e:
            logger.error(f"❌ Ошибка при копировании файла {source_file}: {e}")
            return None
    
    def open_excel_file(self, file_path):
        """Открытие Excel файла для запуска VBA макросов"""
        try:
            if not file_path.exists():
                logger.error(f"❌ Файл не найден: {file_path}")
                return False
            
            logger.info(f"📂 Открытие Excel файла: {file_path.name}")
            
            if platform.system() == "Windows":
                subprocess.Popen(
                    ['start', '', str(file_path)],
                    shell=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            else:
                opener = 'xdg-open' if platform.system() == "Linux" else 'open'
                subprocess.Popen(
                    [opener, str(file_path)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            
            logger.info(f"✓ Excel файл открыт: {file_path.name}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка при открытии файла {file_path}: {e}")
            return False
    
    def process_email_metadata(self, metadata_file, auto_open=True):
        """Обработка метаданных письма и копирование файлов"""
        try:
            metadata = self.load_email_metadata(metadata_file)
            if not metadata:
                return
            
            email_type = metadata.get('type', 'unknown')
            sender = metadata.get('from', 'Unknown')
            subject = metadata.get('subject', 'No Subject')
            company = metadata.get('company', 'Unknown Company')
            attachments = metadata.get('attachments', [])
            
            logger.info(f"📧 Обработка письма: {metadata_file.name}")
            logger.info(f"   Тип: {email_type}")
            logger.info(f"   От: {sender}")
            logger.info(f"   Тема: {subject}")
            logger.info(f"   Компания: {company}")
            logger.info(f"   Вложений: {len(attachments)}")
            
            downloaded_files = []
            
            for attachment_info in attachments:
                saved_as = attachment_info.get('saved_as')
                original_filename = attachment_info.get('filename', saved_as)
                
                if not saved_as:
                    continue
                
                source_file = self.container_dir / saved_as
                
                if not source_file.exists():
                    logger.warning(f"   ⚠ Файл не найден: {saved_as}")
                    continue
                
                logger.info(f"📎 Копирование вложения: {original_filename}")
                
                target_path = self.copy_attachment(source_file, original_filename)
                if target_path:
                    downloaded_files.append(target_path)
            
            if downloaded_files:
                logger.info(f"✓ Скачано файлов: {len(downloaded_files)}")
                
                if auto_open:
                    for file_path in downloaded_files:
                        if file_path.suffix.lower() in ['.xls', '.xlsx', '.xlsm']:
                            self.open_excel_file(file_path)
                
                # Помечаем метаданные как обработанные
                self.processed_files.add(str(metadata_file))
            else:
                logger.info("   Вложений не найдено")
            
        except Exception as e:
            logger.error(f"❌ Ошибка при обработке письма {metadata_file}: {e}")
    
    def process_new_emails(self, auto_open=True):
        """Обработка новых писем"""
        try:
            if not self.check_container_directory():
                return
            
            metadata_files = self.get_new_metadata_files()
            
            if not metadata_files:
                logger.info("📭 Новых писем нет")
                return
            
            logger.info(f"📬 Найдено новых писем: {len(metadata_files)}")
            
            for metadata_file in metadata_files:
                self.process_email_metadata(metadata_file, auto_open=auto_open)
            
        except Exception as e:
            logger.error(f"❌ Ошибка при обработке писем: {e}")
    
    def run_continuous(self, check_interval=5, auto_open=True):
        """Непрерывная проверка новых файлов"""
        logger.info("=" * 60)
        logger.info("ЗАПУСК НЕПРЕРЫВНОЙ ПРОВЕРКИ ФАЙЛОВ ИЗ КОНТЕЙНЕРА")
        logger.info(f"Интервал проверки: {check_interval} сек")
        logger.info("=" * 60)
        logger.info("")
        
        if not self.check_container_directory():
            logger.error("❌ Директория контейнера не найдена. Завершение работы.")
            logger.info("   Убедитесь, что:")
            logger.info("   1. Docker-контейнер phishing-demo запущен")
            logger.info("   2. Volume настроен в docker-compose.yml")
            logger.info("   3. Путь к директории правильный")
            return
        
        try:
            while True:
                logger.info(f"\n{'=' * 60}")
                logger.info(f"Проверка файлов: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                logger.info(f"{'=' * 60}")
                
                self.process_new_emails(auto_open=auto_open)
                
                logger.info(f"\nОжидание {check_interval} сек до следующей проверки...")
                time.sleep(check_interval)
                
        except KeyboardInterrupt:
            logger.info("\n\nОстановка по запросу пользователя (Ctrl+C)")
        except Exception as e:
            logger.error(f"\n❌ Критическая ошибка: {e}")
            raise

# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    """Запуск автоматизации"""
    
    print("=" * 60)
    print("АВТОМАТИЗАЦИЯ ОПЕРАТОРА ДБО (через Docker-контейнер)")
    print("=" * 60)
    print(f"Директория контейнера: {CONTAINER_ATTACHMENTS_DIR}")
    print(f"Директория загрузки: {DOWNLOAD_DIR}")
    print(f"Интервал проверки: {CHECK_INTERVAL} сек")
    print(f"Автооткрытие Excel: {'Да' if AUTO_OPEN_EXCEL else 'Нет'}")
    print("=" * 60)
    print()
    print("Логи выводятся в консоль")
    print("Для остановки нажмите Ctrl+C")
    print()
    
    # Проверяем, что папка Downloads существует
    download_path = Path(DOWNLOAD_DIR)
    if not download_path.exists():
        download_path.mkdir(parents=True, exist_ok=True)
        print(f"✓ Папка Downloads создана: {download_path}")
    else:
        print(f"✓ Папка Downloads найдена: {download_path}")
    print()
    
    # Проверяем директорию контейнера
    container_path = Path(CONTAINER_ATTACHMENTS_DIR)
    if not container_path.exists():
        print(f"⚠ Директория контейнера не найдена: {container_path}")
        print(f"   Создаем директорию...")
        container_path.mkdir(parents=True, exist_ok=True)
        print(f"✓ Директория создана: {container_path}")
    else:
        print(f"✓ Директория контейнера найдена: {container_path}")
    print()
    
    # Создаем экземпляр автоматизации
    automation = DBOOperatorAutomation(
        container_dir=CONTAINER_ATTACHMENTS_DIR,
        download_dir=DOWNLOAD_DIR
    )
    
    # Запускаем непрерывную проверку
    try:
        automation.run_continuous(
            check_interval=CHECK_INTERVAL,
            auto_open=AUTO_OPEN_EXCEL
        )
    except KeyboardInterrupt:
        print("\n\nОстановка по запросу пользователя (Ctrl+C)")
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        raise


if __name__ == "__main__":
    main()
