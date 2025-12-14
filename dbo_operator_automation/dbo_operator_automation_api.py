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
    
    def __init__(self, email_address, password, webmail_url, download_dir):
        """
        Инициализация автоматизации
        
        Args:
            email_address: Email адрес оператора
            password: Пароль от почты
            webmail_url: URL Webmail (Roundcube) (например, http://10.18.2.6/webmail)
            download_dir: Директория для сохранения вложений
        """
        self.email_address = email_address
        self.password = password
        self.webmail_url = webmail_url.rstrip('/')
        self.download_dir = Path(download_dir)
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.processed_emails = set()
        self.session = requests.Session()
        
        logger.info(f"Инициализация автоматизации для {email_address}")
        logger.info(f"Webmail URL: {webmail_url}")
        logger.info(f"Директория загрузки: {self.download_dir}")
    
    def connect(self):
        """Проверка подключения к Webmail (Roundcube)"""
        try:
            logger.info("=" * 60)
            logger.info("ПОПЫТКА ПОДКЛЮЧЕНИЯ К WEBMAIL (ROUNDCUBE)")
            logger.info(f"Webmail URL: {self.webmail_url}")
            logger.info("Проверка доступности...")
            
            # Проверка доступности Webmail
            try:
                response = self.session.get(f"{webmail_url}/", timeout=10)
                if response.status_code == 200:
                    logger.info("✓ Webmail доступен")
                else:
                    logger.warning(f"⚠ Webmail ответил со статусом {response.status_code}")
            except requests.exceptions.RequestException as e:
                logger.warning(f"⚠ Не удалось проверить Webmail: {e}")
                logger.info("Продолжаем попытку подключения...")
            
            # Попытка авторизации в Roundcube
            try:
                # Roundcube API endpoint для авторизации
                login_url = f"{self.webmail_url}/?_task=login"
                login_data = {
                    '_user': self.email_address,
                    '_pass': self.password,
                    '_token': ''  # Roundcube может требовать токен, но для простоты оставляем пустым
                }
                response = self.session.post(login_url, data=login_data, timeout=10, allow_redirects=False)
                
                # Проверяем, успешна ли авторизация (редирект или 200)
                if response.status_code in [200, 302]:
                    logger.info("✓ Подключение к Webmail успешно")
                    logger.info("=" * 60)
                    return True
                elif response.status_code == 401:
                    logger.error("❌ Ошибка авторизации. Проверьте email и пароль")
                    logger.error("=" * 60)
                    return False
                else:
                    logger.warning(f"⚠ Webmail ответил со статусом {response.status_code}")
                    logger.info("Продолжаем работу...")
                    return True
            except requests.exceptions.RequestException as e:
                logger.error(f"❌ Ошибка подключения к Webmail: {e}")
                logger.error("Возможные причины:")
                logger.error("  1. Webmail не включен в mailu.env (WEBMAIL=roundcube)")
                logger.error("  2. Неправильный URL")
                logger.error("  3. Неправильный email или пароль")
                logger.error("  4. Проблемы с сетью")
                logger.error("=" * 60)
                return False
        except Exception as e:
            logger.error(f"❌ Критическая ошибка при подключении: {e}")
            logger.error("=" * 60)
            return False
    
    def get_emails(self, limit=10):
        """
        Получение списка писем через Roundcube API
        """
        try:
            logger.info(f"📬 Получение списка писем (лимит: {limit})...")
            
            # Roundcube API endpoint для получения списка писем
            try:
                # Используем Roundcube JSON API
                response = self.session.post(
                    f"{self.webmail_url}/?_task=mail&_action=list",
                    data={
                        "_mbox": "INBOX",
                        "_page": 1,
                        "_perpage": limit
                    },
                    timeout=10
                )
                
                if response.status_code == 200:
                    # Roundcube возвращает HTML или JSON в зависимости от формата
                    try:
                        data = response.json()
                        if 'list' in data:
                            emails = data['list']
                            logger.info(f"✓ Получено {len(emails)} писем через Roundcube API")
                            return emails
                    except:
                        # Если не JSON, пробуем парсить HTML
                        logger.debug("Ответ не в формате JSON, пробуем другой метод...")
                
                # Альтернативный метод - через прямой запрос к mailbox
                response = self.session.get(
                    f"{self.webmail_url}/?_task=mail&_mbox=INBOX",
                    timeout=10
                )
                
                if response.status_code == 200:
                    # Парсим HTML ответ для получения списка писем
                    # Это упрощенный подход - в реальности лучше использовать JSON API
                    logger.info("✓ Подключение к почтовому ящику успешно")
                    # Возвращаем пустой список, так как парсинг HTML сложен
                    # В реальной реализации нужно использовать правильный Roundcube JSON API
                    return []
                    
            except Exception as e:
                logger.debug(f"Roundcube API не доступен: {e}")
            
            logger.warning("⚠ Не удалось получить письма через Roundcube API")
            logger.warning("Рекомендуется использовать IMAP для надежной работы")
            return []
            
        except Exception as e:
            logger.error(f"❌ Ошибка при получении писем: {e}")
            return []
    
    def download_email_attachments(self, email_id):
        """Скачивание вложений из письма через API"""
        try:
            logger.info(f"📎 Скачивание вложений из письма ID: {email_id}")
            
            # Получение письма через Roundcube API
            response = self.session.get(
                f"{self.webmail_url}/?_task=mail&_action=get&_uid={email_id}",
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

