#!/usr/bin/env python3
"""
Конфигурация для автоматизации оператора ДБО через IMAP с SSH туннелем
"""

import os
from pathlib import Path

# Настройки почты оператора ДБО
EMAIL_ADDRESS = "operator1@financepro.ru"
EMAIL_PASSWORD = "1q2w#E$R"

# Настройки SSH туннеля
# SSH подключение к удаленной машине с Mailu
SSH_HOST = "10.18.2.6"  # IP адрес удаленной машины
SSH_USER = "iux"  # Пользователь для SSH (измените на свой)
SSH_PORT = 22  # SSH порт (по умолчанию 22)

# Настройки IMAP
# IMAP будет доступен через SSH туннель на localhost:1430
# На удаленной машине IMAP доступен на localhost:143
USE_SSL = False  # TLS отключен в Mailu (TLS_FLAVOR=notls)

# Директория для сохранения скачанных файлов
USER_HOME = Path.home()
DOWNLOAD_DIR = str(USER_HOME / "Downloads")

# Интервал проверки почты (в секундах)
CHECK_INTERVAL = 30  # Проверка каждые 30 секунд

# Автоматически открывать Excel файлы
AUTO_OPEN_EXCEL = True

