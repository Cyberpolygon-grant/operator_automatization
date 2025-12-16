#!/usr/bin/env python3
"""
Автоматизация оператора ДБО через Docker-контейнер
Скачивает файлы из контейнера phishing-demo и автоматически открывает только .xlsm файлы
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
import threading

try:
    import paramiko
    PARAMIKO_AVAILABLE = True
except ImportError:
    PARAMIKO_AVAILABLE = False

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

# ============================================================================
# КОНФИГУРАЦИЯ - ИЗМЕНИТЕ ПОД СВОИ НАСТРОЙКИ
# ============================================================================

# ============================================================================
# SSH НАСТРОЙКИ (для подключения к удаленному серверу)
# ============================================================================

# Использовать SSH для подключения к удаленному серверу
USE_SSH = True

# Настройки SSH подключения
SSH_HOST = "10.18.2.6"  # IP адрес удаленного сервера
SSH_USER = "iux"  # Пользователь для SSH
SSH_PASSWORD = "InfoTecs1830"  # Пароль для SSH (или None для ключей)
SSH_PORT = 22  # SSH порт

# Путь к директории с файлами на удаленном сервере
# Укажите полный путь к директории sent_attachments на удаленном сервере
REMOTE_ATTACHMENTS_DIR = "/home/iux/mail/sent_attachments"  # Путь на удаленном сервере

# ============================================================================
# ЛОКАЛЬНЫЕ НАСТРОЙКИ (если USE_SSH = False)
# ============================================================================

# Путь к директории с файлами из Docker-контейнера (локально)
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

# Автоматически открывать только .xlsm файлы (файлы с макросами)
AUTO_OPEN_EXCEL = True

# Время до автоматического закрытия Excel файла (в секундах)
EXCEL_CLOSE_DELAY = 7

# Обрабатывать все файлы заново (игнорировать список обработанных)
PROCESS_ALL_FILES = False

# Время жизни скачанных файлов в минутах (после этого они удаляются)
FILE_LIFETIME_MINUTES = 10

# ============================================================================
# НАСТРОЙКА ЛОГИРОВАНИЯ
# ============================================================================

log_format = '%(asctime)s [%(levelname)-8s] %(message)s'
date_format = '%Y-%m-%d %H:%M:%S'

logging.basicConfig(
    level=logging.INFO,  # Можно изменить на DEBUG для детального логирования
    format=log_format,
    datefmt=date_format,
    handlers=[
        logging.StreamHandler()  # Только консоль
    ]
)

logger = logging.getLogger(__name__)

# ============================================================================
# КЛАСС SSH ПОДКЛЮЧЕНИЯ
# ============================================================================

class SSHConnection:
    """Управление SSH подключением и SFTP"""
    
    def __init__(self, host, user, password=None, port=22):
        """Инициализация SSH подключения"""
        self.host = host
        self.user = user
        self.password = password
        self.port = port
        self.client = None
        self.sftp = None
        self.is_connected = False
    
    def connect(self):
        """Подключение к SSH серверу"""
        try:
            if not PARAMIKO_AVAILABLE:
                logger.error("❌ paramiko не установлен. Установите: pip install paramiko")
                return False
            
            logger.info(f"🔗 Подключение к SSH серверу...")
            logger.info(f"   SSH: {self.user}@{self.host}:{self.port}")
            
            self.client = paramiko.SSHClient()
            self.client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            self.client.connect(
                hostname=self.host,
                port=self.port,
                username=self.user,
                password=self.password,
                timeout=10,
                look_for_keys=False,
                allow_agent=False
            )
            
            self.sftp = self.client.open_sftp()
            self.is_connected = True
            
            logger.info(f"✓ SSH подключение установлено")
            return True
            
        except paramiko.AuthenticationException:
            logger.error(f"❌ Ошибка аутентификации SSH")
            logger.error(f"   Проверьте правильность пароля")
            return False
        except Exception as e:
            logger.error(f"❌ Ошибка подключения к SSH: {e}")
            return False
    
    def disconnect(self):
        """Отключение от SSH сервера"""
        try:
            if self.sftp:
                self.sftp.close()
                self.sftp = None
            if self.client:
                self.client.close()
                self.client = None
            self.is_connected = False
            logger.info("✓ SSH подключение закрыто")
        except:
            pass
    
    def list_files(self, remote_dir):
        """Получение списка файлов в удаленной директории"""
        try:
            if not self.is_connected:
                return []
            
            files = []
            try:
                for item in self.sftp.listdir_attr(remote_dir):
                    files.append({
                        'name': item.filename,
                        'size': item.st_size,
                        'mtime': item.st_mtime
                    })
            except FileNotFoundError:
                logger.warning(f"⚠ Директория не найдена на удаленном сервере: {remote_dir}")
                return []
            
            return files
        except Exception as e:
            logger.error(f"❌ Ошибка при получении списка файлов: {e}")
            return []
    
    def download_file(self, remote_path, local_path):
        """Скачивание файла с удаленного сервера"""
        try:
            if not self.is_connected:
                return False
            
            self.sftp.get(remote_path, local_path)
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка при скачивании файла {remote_path}: {e}")
            return False
    
    def read_file(self, remote_path):
        """Чтение содержимого файла с удаленного сервера"""
        try:
            if not self.is_connected:
                return None
            
            with self.sftp.open(remote_path, 'r') as f:
                return f.read()
        except Exception as e:
            logger.error(f"❌ Ошибка при чтении файла {remote_path}: {e}")
            return None

# ============================================================================
# КЛАСС АВТОМАТИЗАЦИИ
# ============================================================================

class DBOOperatorAutomation:
    """Автоматизация работы оператора ДБО через Docker-контейнер"""
    
    def __init__(self, container_dir=None, download_dir="downloaded_attachments", 
                 process_all=False, use_ssh=False, ssh_host=None, ssh_user=None, 
                 ssh_password=None, ssh_port=22, remote_dir=None):
        """Инициализация автоматизации"""
        self.use_ssh = use_ssh
        self.download_dir = Path(download_dir)
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.process_all = process_all
        self.processed_files = set()
        self.start_time = datetime.now()  # Время запуска скрипта для фильтрации старых файлов
        self.downloaded_files_times = {}  # {file_path: download_time} для отслеживания времени скачивания
        
        if use_ssh:
            self.ssh = SSHConnection(ssh_host, ssh_user, ssh_password, ssh_port)
            self.remote_dir = remote_dir
            self.container_dir = None
            logger.info(f"Инициализация автоматизации (SSH режим)")
            logger.info(f"SSH сервер: {ssh_user}@{ssh_host}:{ssh_port}")
            logger.info(f"Удаленная директория: {remote_dir}")
        else:
            self.ssh = None
            self.remote_dir = None
            self.container_dir = Path(container_dir) if container_dir else None
            logger.info(f"Инициализация автоматизации (локальный режим)")
            if self.container_dir:
                logger.info(f"Директория контейнера: {self.container_dir}")
            else:
                logger.warning(f"⚠ Директория контейнера не указана")
        
        logger.info(f"Директория загрузки: {self.download_dir}")
        if process_all:
            logger.info(f"⚠ Режим обработки всех файлов (игнорируется список обработанных)")
    
    def check_container_directory(self):
        """Проверка существования директории контейнера"""
        if self.use_ssh:
            if not self.ssh.is_connected:
                if not self.ssh.connect():
                    return False
            # Проверяем доступность удаленной директории
            try:
                self.ssh.sftp.listdir(self.remote_dir)
                return True
            except:
                logger.warning(f"⚠ Удаленная директория не найдена: {self.remote_dir}")
                return False
        else:
            if not self.container_dir or not self.container_dir.exists():
                logger.warning(f"⚠ Директория контейнера не найдена: {self.container_dir}")
                logger.info(f"   Убедитесь, что Docker-контейнер запущен и volume настроен")
                return False
            return True
    
    def get_new_metadata_files(self):
        """Получение списка новых JSON файлов с метаданными"""
        try:
            if self.use_ssh:
                if not self.ssh.is_connected:
                    return []
                
                # Получаем список файлов через SSH
                all_files_info = self.ssh.list_files(self.remote_dir)
                all_files = [f['name'] for f in all_files_info]
                
                logger.debug(f"   Всего файлов в директории: {len(all_files)}")
                if all_files:
                    logger.debug(f"   Примеры файлов: {all_files[:5]}")
                
                metadata_files = []
                for file_info in all_files_info:
                    filename = file_info['name']
                    if filename.endswith('_metadata.json'):
                        # Проверяем время модификации файла - только файлы после запуска
                        file_mtime = datetime.fromtimestamp(file_info['mtime'])
                        if file_mtime < self.start_time:
                            logger.debug(f"   Пропущен старый файл: {filename} (создан: {file_mtime.strftime('%Y-%m-%d %H:%M:%S')})")
                            continue
                        
                        file_key = f"{self.remote_dir}/{filename}"
                        if self.process_all or file_key not in self.processed_files:
                            metadata_files.append({
                                'name': filename,
                                'path': f"{self.remote_dir}/{filename}",
                                'remote': True,
                                'mtime': file_mtime
                            })
                            logger.debug(f"   Найден новый файл метаданных: {filename} (создан: {file_mtime.strftime('%Y-%m-%d %H:%M:%S')})")
                        else:
                            logger.debug(f"   Файл уже обработан: {filename}")
                
                # Если нет метаданных, но есть другие файлы, показываем предупреждение
                if not metadata_files and all_files:
                    non_metadata = [f for f in all_files if not f.endswith('_metadata.json')]
                    if non_metadata:
                        logger.warning(f"   ⚠ Найдены файлы без метаданных: {len(non_metadata)} файл(ов)")
                        logger.info(f"   Убедитесь, что контейнер создает файлы *_metadata.json")
                
                return sorted(metadata_files, key=lambda x: x['name'])
            else:
                if not self.container_dir or not self.container_dir.exists():
                    return []
                
                # Показываем все файлы в директории для отладки
                all_files = list(self.container_dir.iterdir())
                logger.debug(f"   Всего файлов в директории: {len(all_files)}")
                if all_files:
                    logger.debug(f"   Примеры файлов: {[f.name for f in all_files[:5]]}")
                
                metadata_files = []
                for file_path in self.container_dir.glob("*_metadata.json"):
                    # Проверяем время модификации файла - только файлы после запуска
                    file_mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
                    if file_mtime < self.start_time:
                        logger.debug(f"   Пропущен старый файл: {file_path.name} (создан: {file_mtime.strftime('%Y-%m-%d %H:%M:%S')})")
                        continue
                    
                    file_str = str(file_path)
                    if self.process_all or file_str not in self.processed_files:
                        metadata_files.append({
                            'name': file_path.name,
                            'path': str(file_path),
                            'remote': False,
                            'mtime': file_mtime
                        })
                        logger.debug(f"   Найден новый файл метаданных: {file_path.name} (создан: {file_mtime.strftime('%Y-%m-%d %H:%M:%S')})")
                    else:
                        logger.debug(f"   Файл уже обработан: {file_path.name}")
                
                # Если нет метаданных, но есть другие файлы, показываем предупреждение
                if not metadata_files and all_files:
                    non_metadata = [f for f in all_files if not f.name.endswith('_metadata.json')]
                    if non_metadata:
                        logger.warning(f"   ⚠ Найдены файлы без метаданных: {len(non_metadata)} файл(ов)")
                        logger.info(f"   Убедитесь, что контейнер создает файлы *_metadata.json")
                
                return sorted(metadata_files, key=lambda x: x['name'])
        except Exception as e:
            logger.error(f"❌ Ошибка при получении списка файлов: {e}")
            return []
    
    def load_email_metadata(self, metadata_file_info):
        """Загрузка метаданных письма из JSON файла"""
        try:
            if metadata_file_info['remote']:
                # Читаем с удаленного сервера
                content = self.ssh.read_file(metadata_file_info['path'])
                if content:
                    return json.loads(content.decode('utf-8'))
                return None
            else:
                # Читаем локально
                with open(metadata_file_info['path'], 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            logger.error(f"❌ Ошибка при загрузке метаданных {metadata_file_info.get('name', 'unknown')}: {e}")
            return None
    
    def copy_attachment(self, source_file, target_filename, is_remote=False):
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
            
            if is_remote:
                # Скачиваем с удаленного сервера через SFTP
                if self.ssh.download_file(source_file, str(target_path)):
                    logger.info(f"   Файл скачан: {target_path.name}")
                    # Сохраняем время скачивания файла для последующего удаления
                    self.downloaded_files_times[str(target_path)] = datetime.now()
                    return target_path
                else:
                    return None
            else:
                # Копируем локально
                shutil.copy2(source_file, target_path)
                logger.info(f"   Файл скопирован: {target_path.name}")
                
                # Сохраняем время скачивания файла для последующего удаления
                self.downloaded_files_times[str(target_path)] = datetime.now()
                
                return target_path
        except Exception as e:
            logger.error(f"❌ Ошибка при копировании файла {source_file}: {e}")
            return None
    
    def close_excel_file(self, file_path, delay_seconds=7):
        """Закрытие Excel файла через заданное время"""
        def close_after_delay():
            time.sleep(delay_seconds)
            try:
                if platform.system() == "Windows":
                    # Пытаемся использовать COM объект Excel (если доступен)
                    if WIN32COM_AVAILABLE:
                        try:
                            excel = win32com.client.GetActiveObject("Excel.Application")
                            for workbook in excel.Workbooks:
                                try:
                                    if workbook.FullName.lower() == str(file_path.resolve()).lower():
                                        workbook.Close(SaveChanges=False)
                                        logger.info(f"✓ .xlsm файл закрыт: {file_path.name}")
                                        return
                                except:
                                    continue
                        except Exception:
                            pass  # Excel не запущен или другая ошибка
                    
                    # Альтернативный метод - закрываем все процессы Excel
                    try:
                        subprocess.run(
                            ['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            timeout=5
                        )
                        logger.info(f"✓ Excel закрыт: {file_path.name}")
                    except Exception as e:
                        logger.warning(f"⚠ Не удалось закрыть Excel: {file_path.name} ({e})")
                else:
                    # Для Linux/Mac используем pkill
                    try:
                        subprocess.run(
                            ['pkill', '-f', file_path.name],
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            timeout=5
                        )
                        logger.info(f"✓ .xlsm файл закрыт: {file_path.name}")
                    except Exception as e:
                        logger.warning(f"⚠ Не удалось закрыть файл: {file_path.name} ({e})")
            except Exception as e:
                logger.debug(f"   Ошибка при закрытии файла: {e}")
        
        # Запускаем закрытие в отдельном потоке
        thread = threading.Thread(target=close_after_delay, daemon=True)
        thread.start()
    
    def open_excel_file(self, file_path, close_delay=7):
        """Открытие .xlsm файла для запуска VBA макросов через батник"""
        try:
            if not file_path.exists():
                logger.error(f"❌ Файл не найден: {file_path}")
                return False
            
            logger.info(f"📂 Открытие .xlsm файла: {file_path.name}")
            
            if platform.system() == "Windows":
                # Создаём временный батник для открытия файла
                file_path_abs = str(file_path.resolve())
                # Экранируем кавычки в пути для батника
                file_path_escaped = file_path_abs.replace('"', '""')
                bat_content = f'@echo off\ncd /d "{os.path.dirname(file_path_abs)}"\nstart "" "{file_path_escaped}"\n'
                
                bat_file = self.download_dir / f"open_{file_path.stem}_{int(time.time())}.bat"
                
                try:
                    with open(bat_file, 'w', encoding='cp866') as f:
                        f.write(bat_content)
                    
                    # Запускаем батник через cmd
                    subprocess.Popen(
                        ['cmd.exe', '/c', str(bat_file)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        cwd=str(self.download_dir)
                    )
                    
                    # Удаляем батник через небольшую задержку
                    def cleanup_bat():
                        time.sleep(3)
                        try:
                            if bat_file.exists():
                                bat_file.unlink()
                        except:
                            pass
                    
                    threading.Thread(target=cleanup_bat, daemon=True).start()
                    
                except Exception as e:
                    logger.warning(f"⚠ Ошибка создания батника, открываем напрямую: {e}")
                    # Fallback на прямое открытие
                    subprocess.Popen(
                        ['cmd.exe', '/c', 'start', '', str(file_path)],
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
            
            logger.info(f"✓ .xlsm файл открыт: {file_path.name}")
            logger.info(f"   ⏰ Автоматическое закрытие через {close_delay} сек...")
            
            # Запускаем автоматическое закрытие
            self.close_excel_file(file_path, close_delay)
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка при открытии файла {file_path}: {e}")
            return False
    
    def process_email_metadata(self, metadata_file_info, auto_open=True):
        """Обработка метаданных письма и копирование файлов"""
        try:
            metadata = self.load_email_metadata(metadata_file_info)
            if not metadata:
                return
            
            email_type = metadata.get('type', 'unknown')
            sender = metadata.get('from', 'Unknown')
            subject = metadata.get('subject', 'No Subject')
            company = metadata.get('company', 'Unknown Company')
            attachments = metadata.get('attachments', [])
            
            logger.info(f"📧 Обработка письма: {metadata_file_info['name']}")
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
                
                if self.use_ssh:
                    source_file = f"{self.remote_dir}/{saved_as}"
                    is_remote = True
                else:
                    source_file = self.container_dir / saved_as
                    is_remote = False
                
                # Проверяем существование файла
                if is_remote:
                    # Для удаленных файлов проверяем через SSH
                    try:
                        self.ssh.sftp.stat(f"{self.remote_dir}/{saved_as}")
                    except:
                        logger.warning(f"   ⚠ Файл не найден на удаленном сервере: {saved_as}")
                        continue
                else:
                    if not source_file.exists():
                        logger.warning(f"   ⚠ Файл не найден: {saved_as}")
                        continue
                
                logger.info(f"📎 Копирование вложения: {original_filename}")
                
                target_path = self.copy_attachment(source_file, original_filename, is_remote=is_remote)
                if target_path:
                    downloaded_files.append(target_path)
            
            if downloaded_files:
                logger.info(f"✓ Скачано файлов: {len(downloaded_files)}")
                
                if auto_open:
                    for file_path in downloaded_files:
                        if file_path.suffix.lower() == '.xlsm':
                            self.open_excel_file(file_path, close_delay=EXCEL_CLOSE_DELAY)
                
                # Помечаем метаданные как обработанные
                file_key = metadata_file_info['path']
                self.processed_files.add(file_key)
            else:
                logger.info("   Вложений не найдено")
            
        except Exception as e:
            logger.error(f"❌ Ошибка при обработке письма {metadata_file_info.get('name', 'unknown')}: {e}")
    
    def process_file_directly(self, file_path, auto_open=True):
        """Обработка файла напрямую без метаданных"""
        try:
            if not file_path.exists():
                return False
            
            # Копируем файл
            target_path = self.copy_attachment(file_path, file_path.name)
            if not target_path:
                return False
            
            # Открываем только .xlsm файлы
            if auto_open and target_path.suffix.lower() == '.xlsm':
                self.open_excel_file(target_path, close_delay=EXCEL_CLOSE_DELAY)
            
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка при обработке файла {file_path}: {e}")
            return False
    
    def cleanup_old_files(self, lifetime_minutes=10):
        """Удаление файлов, скачанных более указанного времени назад"""
        try:
            current_time = datetime.now()
            files_to_delete = []
            
            # Проверяем все отслеживаемые файлы
            for file_path_str, download_time in list(self.downloaded_files_times.items()):
                file_path = Path(file_path_str)
                age_minutes = (current_time - download_time).total_seconds() / 60
                
                if age_minutes >= lifetime_minutes:
                    if file_path.exists():
                        files_to_delete.append((file_path, age_minutes))
            
            # Удаляем старые файлы
            for file_path, age_minutes in files_to_delete:
                try:
                    file_path.unlink()
                    del self.downloaded_files_times[str(file_path)]
                    logger.info(f"🗑️  Удален старый файл: {file_path.name} (возраст: {age_minutes:.1f} мин)")
                except Exception as e:
                    logger.warning(f"⚠ Не удалось удалить файл {file_path.name}: {e}")
            
            if files_to_delete:
                logger.info(f"✓ Удалено старых файлов: {len(files_to_delete)}")
            
        except Exception as e:
            logger.error(f"❌ Ошибка при очистке старых файлов: {e}")
    
    def process_new_emails(self, auto_open=True):
        """Обработка новых писем"""
        try:
            if not self.check_container_directory():
                return
            
            # Показываем информацию о директории
            if self.use_ssh:
                logger.debug(f"   Проверка удаленной директории: {self.remote_dir}")
            else:
                logger.debug(f"   Проверка директории: {self.container_dir}")
                if self.container_dir:
                    logger.debug(f"   Абсолютный путь: {self.container_dir.resolve()}")
            
            metadata_files = self.get_new_metadata_files()
            
            if not metadata_files:
                # Показываем более детальную информацию
                if self.use_ssh:
                    all_files_info = self.ssh.list_files(self.remote_dir)
                    all_files = [f['name'] for f in all_files_info]
                    json_files = [f for f in all_files if f.endswith('_metadata.json')]
                    other_files = [f for f in all_files if not f.endswith('_metadata.json')]
                else:
                    all_files = list(self.container_dir.iterdir()) if self.container_dir and self.container_dir.exists() else []
                    json_files = [f.name for f in all_files if f.suffix == '.json' and f.name.endswith('_metadata.json')]
                    other_files = [f.name for f in all_files if not f.name.endswith('_metadata.json')]
                
                if all_files:
                    logger.info(f"📭 Новых писем с метаданными нет")
                    logger.info(f"   Найдено: {len(json_files)} файлов метаданных, {len(other_files)} других файлов")
                    
                    if json_files:
                        logger.info(f"   JSON файлы метаданных: {json_files[:3]}")
                        logger.info(f"   Возможно, все файлы уже обработаны")
                else:
                    logger.info("📭 Новых писем нет (директория пуста)")
                return
            
            logger.info(f"📬 Найдено новых писем: {len(metadata_files)}")
            
            for metadata_file in metadata_files:
                self.process_email_metadata(metadata_file, auto_open=auto_open)
            
        except Exception as e:
            logger.error(f"❌ Ошибка при обработке писем: {e}")
            import traceback
            logger.debug(traceback.format_exc())
    
    def run_continuous(self, check_interval=5, auto_open=True):
        """Непрерывная проверка новых файлов"""
        logger.info("=" * 60)
        logger.info("ЗАПУСК НЕПРЕРЫВНОЙ ПРОВЕРКИ ФАЙЛОВ ИЗ КОНТЕЙНЕРА")
        logger.info(f"Интервал проверки: {check_interval} сек")
        logger.info("=" * 60)
        logger.info("")
        
        if not self.check_container_directory():
            logger.error("❌ Не удалось подключиться к директории с файлами. Завершение работы.")
            if self.use_ssh:
                logger.info("   Убедитесь, что:")
                logger.info("   1. SSH сервер доступен")
                logger.info("   2. Правильные учетные данные SSH")
                logger.info("   3. Путь к удаленной директории правильный")
            else:
                logger.info("   Убедитесь, что:")
                logger.info("   1. Docker-контейнер phishing-demo запущен")
                logger.info("   2. Volume настроен в docker-compose.yml")
                logger.info("   3. Путь к директории правильный")
            return
        
        try:
            cleanup_interval_minutes = FILE_LIFETIME_MINUTES
            last_cleanup_time = datetime.now()
            
            while True:
                logger.info(f"\n{'=' * 60}")
                logger.info(f"Проверка файлов: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                logger.info(f"{'=' * 60}")
                
                self.process_new_emails(auto_open=auto_open)
                
                # Периодически очищаем старые файлы (каждые 5 минут или после скачивания)
                current_time = datetime.now()
                time_since_cleanup = (current_time - last_cleanup_time).total_seconds() / 60
                
                if time_since_cleanup >= 5:  # Проверяем каждые 5 минут
                    self.cleanup_old_files(lifetime_minutes=FILE_LIFETIME_MINUTES)
                    last_cleanup_time = current_time
                
                logger.info(f"\nОжидание {check_interval} сек до следующей проверки...")
                time.sleep(check_interval)
                
        except KeyboardInterrupt:
            logger.info("\n\nОстановка по запросу пользователя (Ctrl+C)")
        except Exception as e:
            logger.error(f"\n❌ Критическая ошибка: {e}")
            raise
        finally:
            if self.use_ssh and self.ssh:
                self.ssh.disconnect()

# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    """Запуск автоматизации"""
    
    print("=" * 60)
    print("АВТОМАТИЗАЦИЯ ОПЕРАТОРА ДБО")
    if USE_SSH:
        print("(через SSH подключение к удаленному серверу)")
    else:
        print("(через Docker-контейнер)")
    print("=" * 60)
    if USE_SSH:
        print(f"SSH сервер: {SSH_USER}@{SSH_HOST}:{SSH_PORT}")
        print(f"Удаленная директория: {REMOTE_ATTACHMENTS_DIR}")
    else:
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
    
    if not USE_SSH:
        # Проверяем локальную директорию контейнера
        container_path = Path(CONTAINER_ATTACHMENTS_DIR)
        print(f"📁 Путь к директории контейнера: {container_path}")
        print(f"   Абсолютный путь: {container_path.resolve()}")
        
        if not container_path.exists():
            print(f"⚠ Директория контейнера не найдена: {container_path}")
            print(f"   Создаем директорию...")
            container_path.mkdir(parents=True, exist_ok=True)
            print(f"✓ Директория создана: {container_path}")
        else:
            print(f"✓ Директория контейнера найдена: {container_path}")
            # Показываем содержимое
            files = list(container_path.iterdir())
            if files:
                print(f"   Найдено файлов в директории: {len(files)}")
                json_files = [f for f in files if f.name.endswith('_metadata.json')]
                other_files = [f for f in files if not f.name.endswith('_metadata.json')]
                print(f"   - Файлов метаданных: {len(json_files)}")
                print(f"   - Других файлов: {len(other_files)}")
                if json_files:
                    print(f"   Примеры метаданных: {[f.name for f in json_files[:3]]}")
                if other_files:
                    print(f"   Примеры других файлов: {[f.name for f in other_files[:3]]}")
            else:
                print(f"   Директория пуста")
        print()
    else:
        print(f"📁 Подключение к удаленному серверу через SSH...")
        print()
    
    # Создаем экземпляр автоматизации
    if USE_SSH:
        if not PARAMIKO_AVAILABLE:
            print("❌ paramiko не установлен!")
            print("   Установите: pip install paramiko")
            print("   Или установите из requirements: pip install -r requirements_ssh.txt")
            return
        
        automation = DBOOperatorAutomation(
            download_dir=DOWNLOAD_DIR,
            process_all=PROCESS_ALL_FILES,
            use_ssh=True,
            ssh_host=SSH_HOST,
            ssh_user=SSH_USER,
            ssh_password=SSH_PASSWORD,
            ssh_port=SSH_PORT,
            remote_dir=REMOTE_ATTACHMENTS_DIR
        )
    else:
        automation = DBOOperatorAutomation(
            container_dir=CONTAINER_ATTACHMENTS_DIR,
            download_dir=DOWNLOAD_DIR,
            process_all=PROCESS_ALL_FILES,
            use_ssh=False
        )
    
    # Запускаем непрерывную проверку
    try:
        automation.run_continuous(
            check_interval=CHECK_INTERVAL,
            auto_open=AUTO_OPEN_EXCEL
        )
    except KeyboardInterrupt:
        print("\n\nОстановка по запросу пользователя (Ctrl+C)")
        logger.info("Остановка по запросу пользователя (Ctrl+C)")
    except Exception as e:
        import traceback
        error_msg = f"\n❌ Критическая ошибка: {e}\n"
        error_msg += f"Тип ошибки: {type(e).__name__}\n"
        error_msg += f"Детали:\n{traceback.format_exc()}"
        print(error_msg)
        logger.critical(f"Критическая ошибка: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
