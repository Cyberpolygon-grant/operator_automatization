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
import socket
from email.header import decode_header
from pathlib import Path
import logging
from datetime import datetime

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


class DBOOperatorAutomation:
    """Автоматизация работы оператора ДБО"""
    
    def __init__(self, email_address, password, imap_server, imap_port=993, download_dir="downloaded_attachments", use_ssl=True):
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
        self.use_ssl = use_ssl
        self.download_dir = Path(download_dir)
        # Не создаем поддиректории - сохраняем все файлы прямо в Downloads
        self.download_dir.mkdir(exist_ok=True)
        
        # Используем саму папку Downloads для всех файлов
        self.excel_dir = self.download_dir
        self.other_dir = self.download_dir
        
        self.imap = None
        self.processed_emails = set()  # Для отслеживания обработанных писем
        
    def connect(self):
        """Подключение к IMAP серверу"""
        try:
            logger.info("=" * 60)
            logger.info("ПОПЫТКА ПОДКЛЮЧЕНИЯ К ПОЧТЕ")
            logger.info(f"IMAP сервер: {self.imap_server}:{self.imap_port}")
            logger.info(f"Email: {self.email_address}")
            logger.info("Подключение...")
            
            # Проверка доступности сервера
            logger.info(f"Проверка доступности сервера {self.imap_server}:{self.imap_port}...")
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(5)
                result = sock.connect_ex((self.imap_server, self.imap_port))
                sock.close()
                if result != 0:
                    logger.error(f"❌ Сервер {self.imap_server}:{self.imap_port} недоступен!")
                    logger.error("Возможные причины:")
                    logger.error("  1. Mailu не запущен (запустите: docker compose up -d)")
                    logger.error("  2. Неправильный IP адрес или порт")
                    logger.error("  3. Порт заблокирован файрволом")
                    logger.error("  4. Проблемы с сетью")
                    logger.error("")
                    logger.error("Для Mailu попробуйте:")
                    logger.error("  - IMAP_SERVER = 'localhost' или '127.0.0.1'")
                    logger.error("  - IMAP_PORT = 143 (без SSL) или 993 (с SSL)")
                    logger.error("  - USE_SSL = False (если TLS_FLAVOR=notls)")
                    logger.error("=" * 60)
                    return False
                logger.info("✓ Сервер доступен")
            except socket.gaierror:
                logger.error(f"❌ Не удалось разрешить имя сервера: {self.imap_server}")
                logger.error("Проверьте правильность IP адреса или доменного имени")
                logger.error("Для Mailu используйте 'localhost' или '127.0.0.1'")
                logger.error("=" * 60)
                return False
            except Exception as e:
                logger.warning(f"⚠ Не удалось проверить доступность сервера: {e}")
                logger.info("Продолжаем попытку подключения...")
            
            # Подключение к IMAP
            if self.use_ssl:
                logger.info("Установка SSL соединения...")
                # Для Mailu с TLS_FLAVOR=notls может потребоваться отключить проверку сертификата
                import ssl
                context = ssl.create_default_context()
                context.check_hostname = False
                context.verify_mode = ssl.CERT_NONE
                self.imap = imaplib.IMAP4_SSL(self.imap_server, self.imap_port, timeout=10, ssl_context=context)
                logger.info("✓ SSL соединение установлено")
            else:
                logger.info("Установка обычного соединения (без SSL)...")
                self.imap = imaplib.IMAP4(self.imap_server, self.imap_port)
                logger.info("✓ Соединение установлено")
            
            logger.info("Авторизация...")
            self.imap.login(self.email_address, self.password)
            logger.info("✓ Авторизация успешна")
            logger.info("✓ ПОДКЛЮЧЕНИЕ К ПОЧТЕ УСТАНОВЛЕНО")
            logger.info("=" * 60)
            return True
        except imaplib.IMAP4.error as e:
            logger.error("=" * 60)
            logger.error("ОШИБКА АВТОРИЗАЦИИ")
            logger.error(f"Сервер: {self.imap_server}:{self.imap_port}")
            logger.error(f"Email: {self.email_address}")
            logger.error(f"Ошибка: {e}")
            logger.error("Возможные причины:")
            logger.error("  1. Неправильный email или пароль")
            logger.error("  2. Учетная запись заблокирована")
            logger.error("=" * 60)
            return False
        except ConnectionRefusedError:
            logger.error("=" * 60)
            logger.error("ОШИБКА ПОДКЛЮЧЕНИЯ К ПОЧТЕ")
            logger.error(f"Сервер: {self.imap_server}:{self.imap_port}")
            logger.error(f"Email: {self.email_address}")
            logger.error("Ошибка: Подключение отклонено сервером")
            logger.error("")
            logger.error("Возможные причины:")
            logger.error("  1. Mailu не запущен на удаленной машине")
            logger.error("  2. Неправильный порт (попробуйте 143 вместо 993 или наоборот)")
            logger.error("  3. Порт заблокирован файрволом на удаленной машине")
            logger.error("  4. Неправильный IP адрес")
            logger.error("  5. Проблемы с сетью между машинами")
            logger.error("")
            logger.error("Проверьте:")
            logger.error(f"  - Доступность сервера: ping {self.imap_server}")
            logger.error(f"  - Открыт ли порт: telnet {self.imap_server} {self.imap_port}")
            logger.error(f"  - Запущен ли Mailu: docker compose ps (на удаленной машине)")
            logger.error("")
            logger.error("Для Mailu с TLS_FLAVOR=notls:")
            logger.error("  - Используйте порт 143 с USE_SSL = False")
            logger.error("  - Или порт 993 с USE_SSL = False (может работать)")
            logger.error("=" * 60)
            return False
        except Exception as e:
            logger.error("=" * 60)
            logger.error("ОШИБКА ПОДКЛЮЧЕНИЯ К ПОЧТЕ")
            logger.error(f"Сервер: {self.imap_server}:{self.imap_port}")
            logger.error(f"Email: {self.email_address}")
            logger.error(f"Ошибка: {e}")
            logger.error("")
            logger.error("Возможные причины:")
            if "10061" in str(e) or "Connection refused" in str(e):
                logger.error("  - Сервер отклоняет подключение")
                logger.error("  - Проверьте, запущен ли IMAP сервер")
                logger.error("  - Проверьте правильность порта")
            elif "timed out" in str(e).lower():
                logger.error("  - Превышено время ожидания")
                logger.error("  - Сервер не отвечает")
                logger.error("  - Проверьте доступность сервера")
            else:
                logger.error("  - Неизвестная ошибка подключения")
            logger.error("=" * 60)
            return False
    
    def disconnect(self):
        """Отключение от IMAP сервера"""
        if self.imap:
            try:
                logger.info("Отключение от почты...")
                self.imap.close()
                self.imap.logout()
                logger.info("✓ Отключение от почты выполнено")
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
            logger.info(f"📎 Поиск вложений в письме ID: {email_id}")
            
            # Проверяем, есть ли вложения
            if msg.is_multipart():
                attachment_count = 0
                for part in msg.walk():
                    content_disposition = str(part.get("Content-Disposition", ""))
                    
                    # Ищем вложения
                    if "attachment" in content_disposition or "filename" in content_disposition:
                        attachment_count += 1
                        # Получаем имя файла
                        filename = part.get_filename()
                        if filename:
                            filename = self.decode_mime_words(filename)
                            logger.info(f"   Найдено вложение #{attachment_count}: {filename}")
                            
                            # Сохраняем файл
                            file_path = self.save_attachment(part, filename, email_id)
                            if file_path:
                                downloaded_files.append(file_path)
                                logger.info(f"✓ Вложение #{attachment_count} успешно скачано")
                
                if attachment_count == 0:
                    logger.info("   Вложений не найдено")
                else:
                    logger.info(f"✓ Всего вложений обработано: {len(downloaded_files)}/{attachment_count}")
            else:
                logger.info("   Письмо не содержит вложений (не multipart)")
            
            return downloaded_files
        except Exception as e:
            logger.error(f"❌ ОШИБКА при скачивании вложений: {e}")
            return []
    
    def save_attachment(self, part, filename, email_id):
        """Сохранение вложения на диск"""
        try:
            # Определяем директорию по расширению файла
            file_ext = Path(filename).suffix.lower()
            if file_ext in ['.xlsm', '.xlsx', '.xls']:
                save_dir = self.excel_dir
                file_type = "Excel"
            else:
                save_dir = self.other_dir
                file_type = "Другой"
            
            # Добавляем timestamp и email_id для уникальности
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_filename = f"{timestamp}_{email_id}_{filename}"
            file_path = save_dir / safe_filename
            
            # Получаем размер файла
            file_data = part.get_payload(decode=True)
            file_size = len(file_data)
            file_size_mb = file_size / (1024 * 1024)
            
            logger.info(f"📥 СКАЧИВАНИЕ ФАЙЛА")
            logger.info(f"   Имя: {filename}")
            logger.info(f"   Тип: {file_type}")
            logger.info(f"   Размер: {file_size_mb:.2f} МБ ({file_size} байт)")
            logger.info(f"   Путь сохранения: {file_path}")
            
            # Сохраняем файл
            with open(file_path, 'wb') as f:
                f.write(file_data)
            
            logger.info(f"✓ Файл успешно сохранен: {file_path}")
            return file_path
        except Exception as e:
            logger.error(f"❌ ОШИБКА при сохранении файла {filename}: {e}")
            return None
    
    def open_excel_file(self, file_path):
        """Автоматическое открытие Excel файла"""
        try:
            file_path = Path(file_path)
            if not file_path.exists():
                logger.error(f"❌ Файл не найден: {file_path}")
                return False
            
            logger.info(f"🚀 ОТКРЫТИЕ EXCEL ФАЙЛА")
            logger.info(f"   Файл: {file_path.name}")
            logger.info(f"   Полный путь: {file_path}")
            
            system = platform.system()
            logger.info(f"   ОС: {system}")
            
            if system == "Windows":
                # Windows: используем start для открытия файла
                logger.info("   Команда: start (Windows)")
                process = subprocess.Popen(['start', '', str(file_path)], shell=True)
                logger.info(f"✓ Файл отправлен на открытие в Excel (PID: {process.pid})")
            elif system == "Darwin":  # macOS
                logger.info("   Команда: open (macOS)")
                process = subprocess.Popen(['open', str(file_path)])
                logger.info(f"✓ Файл отправлен на открытие в Excel (PID: {process.pid})")
            elif system == "Linux":
                # Linux: пробуем разные способы
                try:
                    logger.info("   Команда: xdg-open (Linux)")
                    process = subprocess.Popen(['xdg-open', str(file_path)])
                    logger.info(f"✓ Файл отправлен на открытие в Excel (PID: {process.pid})")
                except:
                    # Пробуем через libreoffice
                    logger.info("   Команда: libreoffice (Linux)")
                    process = subprocess.Popen(['libreoffice', '--calc', str(file_path)])
                    logger.info(f"✓ Файл отправлен на открытие через LibreOffice (PID: {process.pid})")
            else:
                logger.warning(f"❌ Неизвестная ОС: {system}")
                return False
            
            logger.info(f"✓ Excel файл успешно запущен: {file_path.name}")
            return True
        except Exception as e:
            logger.error(f"❌ ОШИБКА при открытии файла {file_path}: {e}")
            return False
    
    def process_new_emails(self, auto_open=True):
        """Обработка новых писем"""
        try:
            logger.info("")
            logger.info("=" * 60)
            logger.info("ПРОВЕРКА НОВЫХ ПИСЕМ")
            logger.info("=" * 60)
            
            # Выбираем папку INBOX
            logger.info("Выбор папки INBOX...")
            status, data = self.imap.select("INBOX")
            if status == 'OK':
                logger.info(f"✓ Папка INBOX выбрана (всего писем: {data[0].decode()})")
            else:
                logger.error(f"❌ Не удалось выбрать папку INBOX")
                return []
            
            # Ищем непрочитанные письма
            logger.info("Поиск непрочитанных писем...")
            status, messages = self.imap.search(None, 'UNSEEN')
            
            if status != 'OK':
                logger.warning("❌ Не удалось выполнить поиск писем")
                return []
            
            email_ids = messages[0].split()
            logger.info(f"✓ Найдено новых писем: {len(email_ids)}")
            
            if len(email_ids) == 0:
                logger.info("Новых писем не найдено")
                logger.info("=" * 60)
                return []
            
            processed_files = []
            
            for idx, email_id in enumerate(email_ids, 1):
                try:
                    email_id_str = email_id.decode('utf-8')
                    
                    logger.info("")
                    logger.info(f"📧 ОБРАБОТКА ПИСЬМА #{idx}/{len(email_ids)}")
                    logger.info(f"   ID письма: {email_id_str}")
                    
                    # Пропускаем уже обработанные письма
                    if email_id_str in self.processed_emails:
                        logger.info(f"   ⚠ Письмо уже обработано ранее, пропускаем")
                        continue
                    
                    # Получаем письмо
                    logger.info("   Загрузка письма...")
                    status, msg_data = self.imap.fetch(email_id, '(RFC822)')
                    
                    if status != 'OK':
                        logger.error(f"   ❌ Не удалось загрузить письмо")
                        continue
                    
                    # Парсим письмо
                    email_body = msg_data[0][1]
                    msg = email.message_from_bytes(email_body)
                    
                    # Получаем информацию о письме
                    subject = self.decode_mime_words(msg["Subject"] or "Без темы")
                    sender = self.decode_mime_words(msg["From"] or "Неизвестный")
                    date = msg.get("Date", "Неизвестно")
                    
                    logger.info(f"   От: {sender}")
                    logger.info(f"   Тема: {subject}")
                    logger.info(f"   Дата: {date}")
                    
                    # Скачиваем вложения
                    downloaded_files = self.download_attachments(msg, email_id_str)
                    
                    # Открываем Excel файлы автоматически
                    excel_files_count = 0
                    for file_path in downloaded_files:
                        file_ext = Path(file_path).suffix.lower()
                        if file_ext in ['.xlsm', '.xlsx', '.xls']:
                            excel_files_count += 1
                            if auto_open:
                                logger.info("")
                                self.open_excel_file(file_path)
                                # Небольшая задержка между открытием файлов
                                time.sleep(2)
                            processed_files.append(file_path)
                    
                    if excel_files_count > 0:
                        logger.info(f"✓ Excel файлов обработано: {excel_files_count}")
                    
                    # Помечаем письмо как обработанное
                    self.processed_emails.add(email_id_str)
                    logger.info(f"✓ Письмо #{idx} успешно обработано")
                    
                    # Помечаем письмо как прочитанное (опционально)
                    # self.imap.store(email_id, '+FLAGS', '\\Seen')
                    
                except Exception as e:
                    logger.error(f"❌ ОШИБКА при обработке письма {email_id}: {e}")
                    continue
            
            logger.info("")
            logger.info("=" * 60)
            logger.info(f"ИТОГО ОБРАБОТАНО: {len(processed_files)} файлов из {len(email_ids)} писем")
            logger.info("=" * 60)
            logger.info("")
            
            return processed_files
            
        except Exception as e:
            logger.error(f"❌ ОШИБКА при обработке писем: {e}")
            return []
    
    def run_continuous(self, check_interval=30, auto_open=True):
        """
        Непрерывная работа: проверка почты каждые N секунд
        
        Args:
            check_interval: Интервал проверки почты в секундах (по умолчанию 30)
            auto_open: Автоматически открывать Excel файлы (по умолчанию True)
        """
        logger.info("")
        logger.info("=" * 60)
        logger.info("ЗАПУСК АВТОМАТИЗАЦИИ ОПЕРАТОРА ДБО")
        logger.info("=" * 60)
        logger.info(f"Режим: Непрерывная проверка")
        logger.info(f"Интервал проверки: {check_interval} секунд")
        logger.info(f"Автооткрытие Excel: {'✓ Включено' if auto_open else '✗ Выключено'}")
        logger.info(f"Директория загрузки: {self.download_dir}")
        logger.info(f"Все файлы сохраняются в: {self.download_dir}")
        logger.info("=" * 60)
        logger.info("")
        
        if not self.connect():
            logger.error("❌ Не удалось подключиться к почте. Завершение работы.")
            return
        
        check_count = 0
        try:
            while True:
                check_count += 1
                logger.info("")
                logger.info(f"🔄 ПРОВЕРКА #{check_count} - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                
                processed_files = self.process_new_emails(auto_open=auto_open)
                
                if processed_files:
                    logger.info(f"✓ В этой проверке обработано файлов: {len(processed_files)}")
                else:
                    logger.info("ℹ Новых писем не найдено")
                
                logger.info(f"⏳ Ожидание {check_interval} секунд до следующей проверки...")
                logger.info("")
                time.sleep(check_interval)
                
        except KeyboardInterrupt:
            logger.info("")
            logger.info("=" * 60)
            logger.info("ОСТАНОВКА ПО ЗАПРОСУ ПОЛЬЗОВАТЕЛЯ")
            logger.info(f"Всего проверок выполнено: {check_count}")
            logger.info("=" * 60)
        except Exception as e:
            logger.error("")
            logger.error("=" * 60)
            logger.error(f"❌ КРИТИЧЕСКАЯ ОШИБКА: {e}")
            logger.error(f"Всего проверок выполнено: {check_count}")
            logger.error("=" * 60)
        finally:
            self.disconnect()
            logger.info("")
            logger.info("Автоматизация остановлена")
    
    def run_once(self, auto_open=True):
        """
        Однократная проверка почты
        
        Args:
            auto_open: Автоматически открывать Excel файлы (по умолчанию True)
        """
        logger.info("")
        logger.info("=" * 60)
        logger.info("ОДНОКРАТНАЯ ПРОВЕРКА ПОЧТЫ")
        logger.info("=" * 60)
        logger.info(f"Автооткрытие Excel: {'✓ Включено' if auto_open else '✗ Выключено'}")
        logger.info(f"Директория загрузки: {self.download_dir}")
        logger.info("=" * 60)
        logger.info("")
        
        if not self.connect():
            logger.error("❌ Не удалось подключиться к почте")
            return []
        
        try:
            processed_files = self.process_new_emails(auto_open=auto_open)
            logger.info("")
            logger.info("=" * 60)
            logger.info("ПРОВЕРКА ЗАВЕРШЕНА")
            logger.info(f"Обработано файлов: {len(processed_files)}")
            logger.info("=" * 60)
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

