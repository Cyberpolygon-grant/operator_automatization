#!/usr/bin/env python3
"""
Автоматизация для оператора ДБО через Mailu API
Скачивает файлы из почты и автоматически открывает их для запуска VBA макросов
Использует Mailu REST API вместо IMAP
"""

import requests
import os
import time
import subprocess
import platform
from pathlib import Path
import logging
from datetime import datetime
import base64
import json

# Настройка логирования - только в консоль
log_format = '%(asctime)s [%(levelname)-8s] %(message)s'
date_format = '%Y-%m-%d %H:%M:%S'

logging.basicConfig(
    level=logging.INFO,
    format=log_format,
    datefmt=date_format,
    handlers=[
        logging.StreamHandler()  # Только консоль, без файла
    ]
)

logger = logging.getLogger(__name__)


class DBOOperatorAutomationAPI:
    """Автоматизация работы оператора ДБО через Mailu API"""
    
    def __init__(self, email_address, password, api_url, api_token, download_dir):
        """
        Инициализация автоматизации
        
        Args:
            email_address: Email адрес оператора
            password: Пароль от почты
            api_url: URL Mailu API (например, http://10.18.2.6/api)
            api_token: API токен из mailu.env
            download_dir: Директория для сохранения вложений
        """
        self.email_address = email_address
        self.password = password
        self.api_url = api_url.rstrip('/')
        self.api_token = api_token
        self.download_dir = Path(download_dir)
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.processed_emails = set()
        self.session = requests.Session()
        
        # Настройка сессии
        if api_token:
            self.session.headers.update({
                'X-API-Token': api_token
            })
        else:
            # Используем базовую аутентификацию
            self.session.auth = (email_address, password)
        
        logger.info(f"Инициализация автоматизации для {email_address}")
        logger.info(f"API URL: {api_url}")
        logger.info(f"Директория загрузки: {self.download_dir}")
    
    def connect(self):
        """Проверка подключения к API"""
        try:
            logger.info("=" * 60)
            logger.info("ПОПЫТКА ПОДКЛЮЧЕНИЯ К MAILU API")
            logger.info(f"API URL: {self.api_url}")
            logger.info("Проверка доступности...")
            
            # Проверка доступности API
            try:
                response = self.session.get(f"{self.api_url}/health", timeout=10)
                if response.status_code == 200:
                    logger.info("✓ API доступен")
                else:
                    logger.warning(f"⚠ API ответил со статусом {response.status_code}")
            except requests.exceptions.RequestException as e:
                logger.warning(f"⚠ Не удалось проверить health endpoint: {e}")
                logger.info("Продолжаем попытку подключения...")
            
            # Попытка получить информацию о пользователе
            try:
                # Mailu API endpoint для получения информации о пользователе
                response = self.session.get(
                    f"{self.api_url}/user/{self.email_address.split('@')[0]}",
                    timeout=10
                )
                if response.status_code == 200:
                    logger.info("✓ Подключение к API успешно")
                    logger.info("=" * 60)
                    return True
                elif response.status_code == 401:
                    logger.error("❌ Ошибка авторизации. Проверьте email и пароль или API токен")
                    logger.error("=" * 60)
                    return False
                else:
                    logger.warning(f"⚠ API ответил со статусом {response.status_code}")
                    logger.info("Продолжаем работу...")
                    return True
            except requests.exceptions.RequestException as e:
                logger.error(f"❌ Ошибка подключения к API: {e}")
                logger.error("Возможные причины:")
                logger.error("  1. API не включен в mailu.env (API=true)")
                logger.error("  2. Неправильный API URL")
                logger.error("  3. Неправильный API токен")
                logger.error("  4. Проблемы с сетью")
                logger.error("=" * 60)
                return False
        except Exception as e:
            logger.error(f"❌ Критическая ошибка при подключении: {e}")
            logger.error("=" * 60)
            return False
    
    def get_emails(self, limit=10):
        """
        Получение списка писем через Mailu API
        
        Примечание: Mailu API может не иметь прямого endpoint для получения писем.
        В этом случае используем альтернативный подход через Webmail API или IMAP через HTTP.
        """
        try:
            logger.info(f"📬 Получение списка писем (лимит: {limit})...")
            
            # Mailu API может не иметь прямого endpoint для получения писем
            # Используем альтернативный подход - через Webmail (Roundcube) API
            # или через IMAP через HTTP прокси
            
            # Попытка 1: Через Mailu API (если доступен)
            try:
                # Это примерный endpoint, нужно проверить документацию Mailu
                response = self.session.get(
                    f"{self.api_url}/mailbox/{self.email_address}/messages",
                    params={"limit": limit},
                    timeout=10
                )
                if response.status_code == 200:
                    emails = response.json()
                    logger.info(f"✓ Получено {len(emails)} писем через Mailu API")
                    return emails
            except Exception as e:
                logger.debug(f"Mailu API endpoint не доступен: {e}")
            
            # Попытка 2: Через Webmail (Roundcube) API
            try:
                webmail_url = self.api_url.replace('/api', '/webmail')
                # Roundcube API endpoint для получения писем
                response = self.session.post(
                    f"{webmail_url}/?_task=mail&_action=list",
                    json={
                        "mbox": "INBOX",
                        "page": 1,
                        "per_page": limit
                    },
                    timeout=10
                )
                if response.status_code == 200:
                    data = response.json()
                    logger.info(f"✓ Получено писем через Webmail API")
                    return data.get('list', [])
            except Exception as e:
                logger.debug(f"Webmail API не доступен: {e}")
            
            logger.warning("⚠ Не удалось получить письма через API")
            logger.warning("Рекомендуется использовать IMAP или включить API в mailu.env")
            return []
            
        except Exception as e:
            logger.error(f"❌ Ошибка при получении писем: {e}")
            return []
    
    def download_email_attachments(self, email_id):
        """Скачивание вложений из письма через API"""
        try:
            logger.info(f"📎 Скачивание вложений из письма ID: {email_id}")
            
            # Получение письма через API
            response = self.session.get(
                f"{self.api_url}/mailbox/{self.email_address}/messages/{email_id}",
                timeout=10
            )
            
            if response.status_code != 200:
                logger.error(f"❌ Не удалось получить письмо: {response.status_code}")
                return []
            
            email_data = response.json()
            attachments = email_data.get('attachments', [])
            
            downloaded_files = []
            for i, attachment in enumerate(attachments, 1):
                filename = attachment.get('filename', f'attachment_{i}')
                content = attachment.get('content')  # Base64 encoded
                
                if content:
                    # Декодируем base64
                    file_content = base64.b64decode(content)
                    file_path = self.save_file(file_content, filename, email_id)
                    if file_path:
                        downloaded_files.append(file_path)
                        logger.info(f"✓ Вложение #{i} скачано: {filename}")
            
            return downloaded_files
            
        except Exception as e:
            logger.error(f"❌ Ошибка при скачивании вложений: {e}")
            return []
    
    def save_file(self, content, filename, email_id):
        """Сохранение файла"""
        try:
            # Очистка имени файла
            safe_filename = "".join(c for c in filename if c.isalnum() or c in ".-_ ")
            safe_filename = safe_filename.strip()
            
            if not safe_filename:
                safe_filename = f"attachment_{email_id}_{int(time.time())}"
            
            file_path = self.download_dir / safe_filename
            
            # Если файл уже существует, добавляем номер
            counter = 1
            original_path = file_path
            while file_path.exists():
                stem = original_path.stem
                suffix = original_path.suffix
                file_path = self.download_dir / f"{stem}_{counter}{suffix}"
                counter += 1
            
            with open(file_path, 'wb') as f:
                f.write(content)
            
            logger.info(f"   Файл сохранен: {file_path.name}")
            return file_path
            
        except Exception as e:
            logger.error(f"❌ Ошибка при сохранении файла {filename}: {e}")
            return None
    
    def open_excel_file(self, file_path):
        """Открытие Excel файла для запуска VBA макросов"""
        try:
            if not file_path.exists():
                logger.error(f"❌ Файл не найден: {file_path}")
                return False
            
            logger.info(f"📂 Открытие Excel файла: {file_path.name}")
            
            if platform.system() == "Windows":
                # Windows: используем start для открытия файла
                subprocess.Popen(
                    ['start', '', str(file_path)],
                    shell=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            else:
                # Linux/Mac: используем xdg-open или open
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
    
    def process_emails(self, auto_open=True):
        """Обработка новых писем"""
        try:
            emails = self.get_emails(limit=10)
            
            if not emails:
                logger.info("📭 Новых писем нет")
                return
            
            logger.info(f"📬 Найдено писем: {len(emails)}")
            
            for email_data in emails:
                email_id = email_data.get('id') or email_data.get('uid')
                
                if not email_id:
                    continue
                
                # Проверяем, не обрабатывали ли мы это письмо
                if email_id in self.processed_emails:
                    continue
                
                logger.info(f"📧 Обработка письма ID: {email_id}")
                
                # Скачиваем вложения
                downloaded_files = self.download_email_attachments(email_id)
                
                if downloaded_files:
                    logger.info(f"✓ Скачано файлов: {len(downloaded_files)}")
                    
                    # Открываем Excel файлы
                    if auto_open:
                        for file_path in downloaded_files:
                            if file_path.suffix.lower() in ['.xls', '.xlsx', '.xlsm']:
                                self.open_excel_file(file_path)
                    
                    # Помечаем письмо как обработанное
                    self.processed_emails.add(email_id)
                else:
                    logger.info("   Вложений не найдено")
            
        except Exception as e:
            logger.error(f"❌ Ошибка при обработке писем: {e}")
    
    def run_continuous(self, check_interval=30, auto_open=True):
        """Непрерывная проверка почты"""
        logger.info("=" * 60)
        logger.info("ЗАПУСК НЕПРЕРЫВНОЙ ПРОВЕРКИ ПОЧТЫ")
        logger.info(f"Интервал проверки: {check_interval} сек")
        logger.info("=" * 60)
        logger.info("")
        
        if not self.connect():
            logger.error("❌ Не удалось подключиться к API. Завершение работы.")
            return
        
        try:
            while True:
                logger.info(f"\n{'=' * 60}")
                logger.info(f"Проверка почты: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                logger.info(f"{'=' * 60}")
                
                self.process_emails(auto_open=auto_open)
                
                logger.info(f"\nОжидание {check_interval} сек до следующей проверки...")
                time.sleep(check_interval)
                
        except KeyboardInterrupt:
            logger.info("\n\nОстановка по запросу пользователя (Ctrl+C)")
        except Exception as e:
            logger.error(f"\n❌ Критическая ошибка: {e}")
            raise

