#!/usr/bin/env python3
"""
Конфигурация для автоматизации оператора ДБО
Измените настройки под вашу среду
"""

import os
from pathlib import Path

# Настройки почты оператора ДБО
EMAIL_ADDRESS = "operator1@financepro.ru"
EMAIL_PASSWORD = "1q2w#E$R"

# Настройки IMAP сервера
IMAP_SERVER = "10.18.2.6"  # IP адрес IMAP сервера
IMAP_PORT = 993  # 993 для SSL, 143 для обычного подключения

# Директория для сохранения скачанных файлов
# Используем папку Downloads пользователя
USER_HOME = Path.home()
DOWNLOAD_DIR = str(USER_HOME / "Downloads")

# Интервал проверки почты (в секундах)
CHECK_INTERVAL = 30  # Проверка каждые 30 секунд

# Автоматически открывать Excel файлы
AUTO_OPEN_EXCEL = True

# Режим работы: 'continuous' (непрерывный) или 'once' (однократный)
MODE = 'continuous'

