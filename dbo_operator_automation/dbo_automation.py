#!/usr/bin/env python3
"""
Автоматизация оператора ДБО через IMAP с SSH туннелем
Скачивает файлы из почты и автоматически открывает Excel файлы
Логи выводятся в консоль

ВСЁ В ОДНОМ ФАЙЛЕ - просто запустите: python dbo_automation.py
"""

import imaplib
import email
import os
import time
import subprocess
import platform
import socket
import threading
from email.header import decode_header
from pathlib import Path
import logging
from datetime import datetime

try:
    import paramiko
    PARAMIKO_AVAILABLE = True
except ImportError:
    PARAMIKO_AVAILABLE = False

# ============================================================================
# КОНФИГУРАЦИЯ - ИЗМЕНИТЕ ПОД СВОИ НАСТРОЙКИ
# ============================================================================

# Настройки почты оператора ДБО
EMAIL_ADDRESS = "operator1@financepro.ru"
EMAIL_PASSWORD = "1q2w#E$R"

# Настройки SSH туннеля
SSH_HOST = "10.18.2.6"  # IP адрес удаленной машины
SSH_USER = "iux"  # Пользователь для SSH
SSH_PASSWORD = "InfoTecs1830"  # Пароль для SSH
SSH_PORT = 22  # SSH порт

# Настройки IMAP
USE_SSL = False  # TLS отключен в Mailu (TLS_FLAVOR=notls)

# Директория для сохранения скачанных файлов
USER_HOME = Path.home()
DOWNLOAD_DIR = str(USER_HOME / "Downloads")

# Интервал проверки почты (в секундах)
CHECK_INTERVAL = 30

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
# КЛАСС SSH ТУННЕЛЯ
# ============================================================================

class SSHTunnel:
    """Управление SSH туннелем"""
    
    def __init__(self, ssh_host, ssh_user, ssh_password=None, ssh_port=22, local_port=1430, remote_host='localhost', remote_port=143):
        """Инициализация SSH туннеля"""
        self.ssh_host = ssh_host
        self.ssh_user = ssh_user
        self.ssh_password = ssh_password
        self.ssh_port = ssh_port
        self.local_port = local_port
        self.remote_host = remote_host
        self.remote_port = remote_port
        self.process = None
        self.transport = None
        self.client = None
        self.local_server = None
        self.is_running = False
    
    def start(self):
        """Запуск SSH туннеля"""
        try:
            logger.info(f"🔗 Создание SSH туннеля...")
            logger.info(f"   SSH: {self.ssh_user}@{self.ssh_host}:{self.ssh_port}")
            logger.info(f"   Туннель: localhost:{self.local_port} -> {self.remote_host}:{self.remote_port}")
            
            if self.ssh_password and PARAMIKO_AVAILABLE:
                return self._start_with_paramiko()
            else:
                return self._start_with_ssh_command()
                
        except Exception as e:
            logger.error(f"❌ Ошибка при создании SSH туннеля: {e}")
            return False
    
    def _start_with_paramiko(self):
        """Запуск SSH туннеля через paramiko (с паролем)"""
        try:
            logger.info("   Использование paramiko для SSH туннеля...")
            
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            client.connect(
                hostname=self.ssh_host,
                port=self.ssh_port,
                username=self.ssh_user,
                password=self.ssh_password,
                timeout=10,
                look_for_keys=False,
                allow_agent=False
            )
            
            transport = client.get_transport()
            
            local_server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            local_server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            local_server.bind(('127.0.0.1', self.local_port))
            local_server.listen(1)
            
            def tunnel_handler():
                while self.is_running:
                    try:
                        local_sock, _ = local_server.accept()
                        remote_sock = transport.open_channel(
                            'direct-tcpip',
                            (self.remote_host, self.remote_port),
                            ('127.0.0.1', self.local_port)
                        )
                        threading.Thread(
                            target=self._forward,
                            args=(local_sock, remote_sock),
                            daemon=True
                        ).start()
                        threading.Thread(
                            target=self._forward,
                            args=(remote_sock, local_sock),
                            daemon=True
                        ).start()
                    except Exception as e:
                        if self.is_running:
                            logger.debug(f"Ошибка в туннеле: {e}")
            
            tunnel_thread = threading.Thread(target=tunnel_handler, daemon=True)
            tunnel_thread.start()
            
            self.transport = transport
            self.client = client
            self.local_server = local_server
            self.is_running = True
            
            time.sleep(1)
            
            logger.info(f"✓ SSH туннель создан через paramiko: localhost:{self.local_port}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка при создании SSH туннеля через paramiko: {e}")
            if "Authentication failed" in str(e):
                logger.error("   Проверьте правильность пароля")
            return False
    
    def _forward(self, source, dest):
        """Пересылка данных между сокетами"""
        try:
            while True:
                data = source.recv(4096)
                if not data:
                    break
                dest.sendall(data)
        except:
            pass
        finally:
            source.close()
            dest.close()
    
    def _start_with_ssh_command(self):
        """Запуск SSH туннеля через команду ssh (без пароля, используя ключи)"""
        try:
            known_hosts_file = 'NUL' if platform.system() == 'Windows' else '/dev/null'
            ssh_cmd = [
                'ssh',
                '-N',
                '-f',
                '-L', f'{self.local_port}:{self.remote_host}:{self.remote_port}',
                '-o', 'StrictHostKeyChecking=no',
                '-o', f'UserKnownHostsFile={known_hosts_file}',
                '-o', 'LogLevel=ERROR',
                '-o', 'ServerAliveInterval=60',
                '-o', 'ServerAliveCountMax=3',
                '-o', 'ConnectTimeout=10',
                '-p', str(self.ssh_port),
                f'{self.ssh_user}@{self.ssh_host}'
            ]
            
            if platform.system() == 'Windows':
                logger.info("   Запуск SSH туннеля (Windows)...")
                self.process = subprocess.Popen(
                    ssh_cmd[:-1],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                )
            else:
                logger.info("   Запуск SSH туннеля (Linux/Mac)...")
                self.process = subprocess.Popen(
                    ssh_cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
            
            time.sleep(2)
            
            if self.process.poll() is None:
                self.is_running = True
                logger.info(f"✓ SSH туннель создан: localhost:{self.local_port}")
                return True
            else:
                stderr = self.process.stderr.read().decode('utf-8', errors='ignore')
                logger.error(f"❌ Не удалось создать SSH туннель")
                logger.error(f"   Ошибка: {stderr}")
                if self.ssh_password:
                    logger.error("   Попробуйте установить paramiko: pip install paramiko")
                return False
                
        except FileNotFoundError:
            logger.error("❌ SSH не найден. Установите OpenSSH клиент.")
            return False
        except Exception as e:
            logger.error(f"❌ Ошибка при создании SSH туннеля: {e}")
            return False
    
    def stop(self):
        """Остановка SSH туннеля"""
        self.is_running = False
        
        if self.transport:
            try:
                self.transport.close()
                self.client.close()
                self.local_server.close()
                logger.info("✓ SSH туннель остановлен (paramiko)")
            except:
                pass
            self.transport = None
            self.client = None
            self.local_server = None
        
        if self.process:
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
                logger.info("✓ SSH туннель остановлен (ssh command)")
            except:
                try:
                    self.process.kill()
                except:
                    pass
            self.process = None
    
    def check(self):
        """Проверка, работает ли туннель"""
        if not self.is_running:
            return False
        
        if self.transport:
            if not self.transport.is_active():
                self.is_running = False
                return False
        
        if self.process:
            if self.process.poll() is not None:
                self.is_running = False
                return False
        
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            result = sock.connect_ex(('localhost', self.local_port))
            sock.close()
            return result == 0
        except:
            return False

# ============================================================================
# КЛАСС АВТОМАТИЗАЦИИ
# ============================================================================

class DBOOperatorAutomation:
    """Автоматизация работы оператора ДБО через IMAP с SSH туннелем"""
    
    def __init__(self, email_address, password, ssh_host, ssh_user, ssh_password=None, ssh_port=22, 
                 download_dir="downloaded_attachments", use_ssl=False):
        """Инициализация автоматизации"""
        self.email_address = email_address
        self.password = password
        self.ssh_host = ssh_host
        self.ssh_user = ssh_user
        self.ssh_port = ssh_port
        self.download_dir = Path(download_dir)
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.use_ssl = use_ssl
        self.processed_emails = set()
        self.imap = None
        
        self.ssh_tunnel = SSHTunnel(
            ssh_host=ssh_host,
            ssh_user=ssh_user,
            ssh_password=ssh_password,
            ssh_port=ssh_port,
            local_port=1430,
            remote_host='localhost',
            remote_port=143
        )
        
        logger.info(f"Инициализация автоматизации для {email_address}")
        logger.info(f"SSH: {ssh_user}@{ssh_host}:{ssh_port}")
        logger.info(f"Директория загрузки: {self.download_dir}")
    
    def connect(self):
        """Подключение к IMAP серверу через SSH туннель"""
        try:
            logger.info("=" * 60)
            logger.info("ПОПЫТКА ПОДКЛЮЧЕНИЯ К ПОЧТЕ ЧЕРЕЗ SSH ТУННЕЛЬ")
            
            if not self.ssh_tunnel.is_running:
                if not self.ssh_tunnel.start():
                    logger.error("❌ Не удалось создать SSH туннель")
                    logger.error("=" * 60)
                    return False
            else:
                if not self.ssh_tunnel.check():
                    logger.warning("⚠ SSH туннель не работает, пересоздаем...")
                    self.ssh_tunnel.stop()
                    if not self.ssh_tunnel.start():
                        logger.error("❌ Не удалось пересоздать SSH туннель")
                        logger.error("=" * 60)
                        return False
            
            logger.info(f"IMAP через туннель: localhost:1430")
            logger.info(f"Email: {self.email_address}")
            logger.info("Подключение...")
            
            if self.use_ssl:
                logger.info("Установка SSL соединения...")
                import ssl
                context = ssl.create_default_context()
                context.check_hostname = False
                context.verify_mode = ssl.CERT_NONE
                self.imap = imaplib.IMAP4_SSL('localhost', 1430, timeout=10, ssl_context=context)
                logger.info("✓ SSL соединение установлено")
            else:
                logger.info("Установка обычного соединения (без SSL)...")
                self.imap = imaplib.IMAP4('localhost', 1430, timeout=10)
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
            logger.error(f"Email: {self.email_address}")
            logger.error(f"Ошибка: {e}")
            logger.error("=" * 60)
            return False
        except Exception as e:
            logger.error("=" * 60)
            logger.error("ОШИБКА ПОДКЛЮЧЕНИЯ К ПОЧТЕ")
            logger.error(f"Ошибка: {e}")
            logger.error("=" * 60)
            return False
    
    def disconnect(self):
        """Отключение от IMAP и остановка SSH туннеля"""
        try:
            if self.imap:
                self.imap.logout()
                self.imap = None
        except:
            pass
        self.ssh_tunnel.stop()
    
    def get_unread_emails(self):
        """Получение списка непрочитанных писем"""
        try:
            self.imap.select('INBOX')
            status, messages = self.imap.search(None, 'UNSEEN')
            
            if status != 'OK':
                logger.warning("⚠ Не удалось выполнить поиск писем")
                return []
            
            email_ids = messages[0].split()
            return email_ids
            
        except Exception as e:
            logger.error(f"❌ Ошибка при получении списка писем: {e}")
            return []
    
    def download_attachments(self, msg, email_id):
        """Скачивание вложений из письма"""
        downloaded_files = []
        
        try:
            logger.info(f"📎 Поиск вложений в письме ID: {email_id}")
            
            if msg.is_multipart():
                attachment_count = 0
                for part in msg.walk():
                    content_disposition = str(part.get("Content-Disposition", ""))
                    
                    if "attachment" in content_disposition or "filename" in content_disposition:
                        attachment_count += 1
                        filename = part.get_filename()
                        if filename:
                            filename = self.decode_mime_words(filename)
                            logger.info(f"   Найдено вложение #{attachment_count}: {filename}")
                            
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
        """Сохранение вложения"""
        try:
            safe_filename = "".join(c for c in filename if c.isalnum() or c in ".-_ ")
            safe_filename = safe_filename.strip()
            
            if not safe_filename:
                safe_filename = f"attachment_{email_id.decode()}_{int(time.time())}"
            
            file_path = self.download_dir / safe_filename
            
            counter = 1
            original_path = file_path
            while file_path.exists():
                stem = original_path.stem
                suffix = original_path.suffix
                file_path = self.download_dir / f"{stem}_{counter}{suffix}"
                counter += 1
            
            payload = part.get_payload(decode=True)
            if payload:
                with open(file_path, 'wb') as f:
                    f.write(payload)
                logger.info(f"   Файл сохранен: {file_path.name}")
                return file_path
            
        except Exception as e:
            logger.error(f"❌ Ошибка при сохранении файла {filename}: {e}")
            return None
    
    def decode_mime_words(self, s):
        """Декодирование MIME заголовков"""
        decoded = decode_header(s)
        parts = []
        for part, encoding in decoded:
            if isinstance(part, bytes):
                if encoding:
                    parts.append(part.decode(encoding))
                else:
                    parts.append(part.decode('utf-8', errors='ignore'))
            else:
                parts.append(part)
        return ''.join(parts)
    
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
    
    def process_emails(self, auto_open=True):
        """Обработка новых писем"""
        try:
            email_ids = self.get_unread_emails()
            
            if not email_ids:
                logger.info("📭 Новых писем нет")
                return
            
            logger.info(f"📬 Найдено новых писем: {len(email_ids)}")
            
            for email_id in email_ids:
                if email_id in self.processed_emails:
                    continue
                
                logger.info(f"📧 Обработка письма ID: {email_id.decode()}")
                
                status, msg_data = self.imap.fetch(email_id, '(RFC822)')
                if status != 'OK':
                    continue
                
                email_body = msg_data[0][1]
                msg = email.message_from_bytes(email_body)
                
                subject = self.decode_mime_words(msg['Subject'] or '')
                from_addr = msg['From'] or 'Unknown'
                
                logger.info(f"   От: {from_addr}")
                logger.info(f"   Тема: {subject}")
                
                downloaded_files = self.download_attachments(msg, email_id)
                
                if downloaded_files:
                    logger.info(f"✓ Скачано файлов: {len(downloaded_files)}")
                    
                    if auto_open:
                        for file_path in downloaded_files:
                            if file_path.suffix.lower() in ['.xls', '.xlsx', '.xlsm']:
                                self.open_excel_file(file_path)
                    
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
            logger.error("❌ Не удалось подключиться к почте. Завершение работы.")
            return
        
        try:
            while True:
                logger.info(f"\n{'=' * 60}")
                logger.info(f"Проверка почты: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                logger.info(f"{'=' * 60}")
                
                if not self.ssh_tunnel.check():
                    logger.warning("⚠ SSH туннель не работает, переподключаемся...")
                    if self.imap:
                        try:
                            self.imap.logout()
                        except:
                            pass
                        self.imap = None
                    
                    if not self.connect():
                        logger.error("❌ Не удалось переподключиться. Ожидание...")
                        time.sleep(check_interval)
                        continue
                
                self.process_emails(auto_open=auto_open)
                
                logger.info(f"\nОжидание {check_interval} сек до следующей проверки...")
                time.sleep(check_interval)
                
        except KeyboardInterrupt:
            logger.info("\n\nОстановка по запросу пользователя (Ctrl+C)")
        except Exception as e:
            logger.error(f"\n❌ Критическая ошибка: {e}")
            raise
        finally:
            self.disconnect()

# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    """Запуск автоматизации"""
    
    print("=" * 60)
    print("АВТОМАТИЗАЦИЯ ОПЕРАТОРА ДБО (через SSH туннель)")
    print("=" * 60)
    print(f"Email: {EMAIL_ADDRESS}")
    print(f"SSH: {SSH_USER}@{SSH_HOST}:{SSH_PORT}")
    if SSH_PASSWORD:
        print(f"SSH пароль: {'*' * len(SSH_PASSWORD)}")
    print(f"IMAP через туннель: localhost:1430")
    print(f"Директория загрузки: {DOWNLOAD_DIR}")
    print(f"Интервал проверки: {CHECK_INTERVAL} сек")
    print(f"Автооткрытие Excel: {'Да' if AUTO_OPEN_EXCEL else 'Нет'}")
    print("=" * 60)
    print()
    print("Логи выводятся в консоль")
    print("Для остановки нажмите Ctrl+C")
    print()
    
    # Проверка paramiko
    try:
        import paramiko
        print("✓ paramiko установлен - будет использоваться парольная аутентификация")
    except ImportError:
        print("⚠ paramiko не установлен")
        print("   Для использования пароля установите: pip install paramiko")
        print("   Или настройте SSH ключи для автоматической авторизации")
    print()
    
    # Проверяем, что папка Downloads существует
    download_path = Path(DOWNLOAD_DIR)
    if not download_path.exists():
        download_path.mkdir(parents=True, exist_ok=True)
        print(f"✓ Папка Downloads создана: {download_path}")
    else:
        print(f"✓ Папка Downloads найдена: {download_path}")
    print()
    
    # Создаем экземпляр автоматизации
    automation = DBOOperatorAutomation(
        email_address=EMAIL_ADDRESS,
        password=EMAIL_PASSWORD,
        ssh_host=SSH_HOST,
        ssh_user=SSH_USER,
        ssh_password=SSH_PASSWORD,
        ssh_port=SSH_PORT,
        download_dir=DOWNLOAD_DIR,
        use_ssl=USE_SSL
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

