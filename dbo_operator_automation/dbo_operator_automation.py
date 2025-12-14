#!/usr/bin/env python3
"""
Автоматизация для оператора ДБО
Скачивает файлы из почты и автоматически открывает их для запуска VBA макросов
"""

import imaplib
import email
import os
import time
import subprocess
import platform
from email.header import decode_header
from pathlib import Path
import logging
from datetime import datetime

# Настройка логирования
log_dir = Path(__file__).parent
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / 'dbo_operator_automation.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


class DBOOperatorAutomation:
    """Автоматизация работы оператора ДБО"""
    
    def __init__(self, email_address, password, imap_server, imap_port=993, download_dir="downloaded_attachments"):
        """
        Инициализация автоматизации
        
        Args:
            email_address: Email адрес оператора
            password: Пароль от почты
            imap_server: IMAP сервер
            imap_port: IMAP порт (по умолчанию 993 для SSL)
            download_dir: Директория для сохранения вложений
        """
        self.email_address = email_address
        self.password = password
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.download_dir = Path(download_dir)
        self.download_dir.mkdir(exist_ok=True)
        
        # Создаем поддиректории для разных типов файлов
        self.excel_dir = self.download_dir / "excel_files"
        self.other_dir = self.download_dir / "other_files"
        self.excel_dir.mkdir(exist_ok=True)
        self.other_dir.mkdir(exist_ok=True)
        
        self.imap = None
        self.processed_emails = set()  # Для отслеживания обработанных писем
        
    def connect(self):
        """Подключение к IMAP серверу"""
        try:
            logger.info(f"Подключение к IMAP серверу {self.imap_server}:{self.imap_port}")
            self.imap = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            self.imap.login(self.email_address, self.password)
            logger.info("Успешное подключение к почте")
            return True
        except Exception as e:
            logger.error(f"Ошибка подключения к почте: {e}")
            return False
    
    def disconnect(self):
        """Отключение от IMAP сервера"""
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
                logger.info("Отключение от почты")
            except Exception as e:
                logger.error(f"Ошибка при отключении: {e}")
    
    def decode_mime_words(self, s):
        """Декодирование MIME заголовков"""
        decoded = decode_header(s)
        return ''.join(
            word.decode(encoding or 'utf-8') if isinstance(word, bytes) else word
            for word, encoding in decoded
        )
    
    def download_attachments(self, msg, email_id):
        """Скачивание вложений из письма"""
        downloaded_files = []
        
        try:
            # Проверяем, есть ли вложения
            if msg.is_multipart():
                for part in msg.walk():
                    content_disposition = str(part.get("Content-Disposition", ""))
                    
                    # Ищем вложения
                    if "attachment" in content_disposition or "filename" in content_disposition:
                        # Получаем имя файла
                        filename = part.get_filename()
                        if filename:
                            filename = self.decode_mime_words(filename)
                            
                            # Сохраняем файл
                            file_path = self.save_attachment(part, filename, email_id)
                            if file_path:
                                downloaded_files.append(file_path)
                                logger.info(f"Скачан файл: {filename}")
            
            return downloaded_files
        except Exception as e:
            logger.error(f"Ошибка при скачивании вложений: {e}")
            return []
    
    def save_attachment(self, part, filename, email_id):
        """Сохранение вложения на диск"""
        try:
            # Определяем директорию по расширению файла
            file_ext = Path(filename).suffix.lower()
            if file_ext in ['.xlsm', '.xlsx', '.xls']:
                save_dir = self.excel_dir
            else:
                save_dir = self.other_dir
            
            # Добавляем timestamp и email_id для уникальности
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_filename = f"{timestamp}_{email_id}_{filename}"
            file_path = save_dir / safe_filename
            
            # Сохраняем файл
            with open(file_path, 'wb') as f:
                f.write(part.get_payload(decode=True))
            
            logger.info(f"Файл сохранен: {file_path}")
            return file_path
        except Exception as e:
            logger.error(f"Ошибка при сохранении файла {filename}: {e}")
            return None
    
    def open_excel_file(self, file_path):
        """Автоматическое открытие Excel файла"""
        try:
            file_path = Path(file_path)
            if not file_path.exists():
                logger.error(f"Файл не найден: {file_path}")
                return False
            
            system = platform.system()
            
            if system == "Windows":
                # Windows: используем start для открытия файла
                subprocess.Popen(['start', '', str(file_path)], shell=True)
                logger.info(f"Открыт файл в Excel: {file_path}")
            elif system == "Darwin":  # macOS
                subprocess.Popen(['open', str(file_path)])
                logger.info(f"Открыт файл в Excel: {file_path}")
            elif system == "Linux":
                # Linux: пробуем разные способы
                try:
                    subprocess.Popen(['xdg-open', str(file_path)])
                    logger.info(f"Открыт файл в Excel: {file_path}")
                except:
                    # Пробуем через libreoffice
                    subprocess.Popen(['libreoffice', '--calc', str(file_path)])
                    logger.info(f"Открыт файл через LibreOffice: {file_path}")
            else:
                logger.warning(f"Неизвестная ОС: {system}")
                return False
            
            return True
        except Exception as e:
            logger.error(f"Ошибка при открытии файла {file_path}: {e}")
            return False
    
    def process_new_emails(self, auto_open=True):
        """Обработка новых писем"""
        try:
            # Выбираем папку INBOX
            self.imap.select("INBOX")
            
            # Ищем непрочитанные письма
            status, messages = self.imap.search(None, 'UNSEEN')
            
            if status != 'OK':
                logger.warning("Не удалось найти письма")
                return []
            
            email_ids = messages[0].split()
            logger.info(f"Найдено {len(email_ids)} новых писем")
            
            processed_files = []
            
            for email_id in email_ids:
                try:
                    email_id_str = email_id.decode('utf-8')
                    
                    # Пропускаем уже обработанные письма
                    if email_id_str in self.processed_emails:
                        continue
                    
                    # Получаем письмо
                    status, msg_data = self.imap.fetch(email_id, '(RFC822)')
                    
                    if status != 'OK':
                        continue
                    
                    # Парсим письмо
                    email_body = msg_data[0][1]
                    msg = email.message_from_bytes(email_body)
                    
                    # Получаем информацию о письме
                    subject = self.decode_mime_words(msg["Subject"] or "Без темы")
                    sender = self.decode_mime_words(msg["From"] or "Неизвестный")
                    
                    logger.info(f"Обработка письма: {subject} от {sender}")
                    
                    # Скачиваем вложения
                    downloaded_files = self.download_attachments(msg, email_id_str)
                    
                    # Открываем Excel файлы автоматически
                    for file_path in downloaded_files:
                        file_ext = Path(file_path).suffix.lower()
                        if file_ext in ['.xlsm', '.xlsx', '.xls']:
                            if auto_open:
                                logger.info(f"Автоматическое открытие Excel файла: {file_path}")
                                self.open_excel_file(file_path)
                                # Небольшая задержка между открытием файлов
                                time.sleep(2)
                            processed_files.append(file_path)
                    
                    # Помечаем письмо как обработанное
                    self.processed_emails.add(email_id_str)
                    
                    # Помечаем письмо как прочитанное (опционально)
                    # self.imap.store(email_id, '+FLAGS', '\\Seen')
                    
                except Exception as e:
                    logger.error(f"Ошибка при обработке письма {email_id}: {e}")
                    continue
            
            return processed_files
            
        except Exception as e:
            logger.error(f"Ошибка при обработке писем: {e}")
            return []
    
    def run_continuous(self, check_interval=30, auto_open=True):
        """
        Непрерывная работа: проверка почты каждые N секунд
        
        Args:
            check_interval: Интервал проверки почты в секундах (по умолчанию 30)
            auto_open: Автоматически открывать Excel файлы (по умолчанию True)
        """
        logger.info(f"Запуск непрерывной проверки почты (интервал: {check_interval} сек)")
        logger.info(f"Автоматическое открытие Excel файлов: {'Включено' if auto_open else 'Выключено'}")
        
        if not self.connect():
            logger.error("Не удалось подключиться к почте")
            return
        
        try:
            while True:
                logger.info("Проверка новых писем...")
                processed_files = self.process_new_emails(auto_open=auto_open)
                
                if processed_files:
                    logger.info(f"Обработано файлов: {len(processed_files)}")
                else:
                    logger.info("Новых писем не найдено")
                
                logger.info(f"Ожидание {check_interval} секунд до следующей проверки...")
                time.sleep(check_interval)
                
        except KeyboardInterrupt:
            logger.info("Остановка по запросу пользователя")
        except Exception as e:
            logger.error(f"Критическая ошибка: {e}")
        finally:
            self.disconnect()
    
    def run_once(self, auto_open=True):
        """
        Однократная проверка почты
        
        Args:
            auto_open: Автоматически открывать Excel файлы (по умолчанию True)
        """
        logger.info("Однократная проверка почты")
        
        if not self.connect():
            logger.error("Не удалось подключиться к почте")
            return []
        
        try:
            processed_files = self.process_new_emails(auto_open=auto_open)
            return processed_files
        finally:
            self.disconnect()


def main():
    """Основная функция"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Автоматизация оператора ДБО - скачивание и открытие файлов из почты')
    parser.add_argument('--email', type=str, required=True, help='Email адрес оператора')
    parser.add_argument('--password', type=str, required=True, help='Пароль от почты')
    parser.add_argument('--imap-server', type=str, default='localhost', help='IMAP сервер (по умолчанию: localhost)')
    parser.add_argument('--imap-port', type=int, default=993, help='IMAP порт (по умолчанию: 993)')
    parser.add_argument('--download-dir', type=str, default='downloaded_attachments', help='Директория для сохранения файлов')
    parser.add_argument('--interval', type=int, default=30, help='Интервал проверки почты в секундах (по умолчанию: 30)')
    parser.add_argument('--once', action='store_true', help='Выполнить однократную проверку вместо непрерывной')
    parser.add_argument('--no-auto-open', action='store_true', help='Не открывать Excel файлы автоматически')
    
    args = parser.parse_args()
    
    # Создаем экземпляр автоматизации
    automation = DBOOperatorAutomation(
        email_address=args.email,
        password=args.password,
        imap_server=args.imap_server,
        imap_port=args.imap_port,
        download_dir=args.download_dir
    )
    
    # Запускаем в зависимости от режима
    if args.once:
        files = automation.run_once(auto_open=not args.no_auto_open)
        print(f"\nОбработано файлов: {len(files)}")
        for f in files:
            print(f"  - {f}")
    else:
        automation.run_continuous(
            check_interval=args.interval,
            auto_open=not args.no_auto_open
        )


if __name__ == "__main__":
    main()

