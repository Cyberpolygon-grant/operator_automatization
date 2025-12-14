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
# Mailu доступен на удаленной машине (10.18.2.6)
# Front сервис Mailu пробрасывает порты 143 и 993
# TLS отключен (TLS_FLAVOR=notls в mailu.env), используем порт 143 без SSL
IMAP_SERVER = "10.18.2.6"  # IP адрес удаленной машины с Mailu
IMAP_PORT = 143  # 143 для обычного подключения (993 для SSL, но TLS отключен в Mailu)
USE_SSL = False  # TLS отключен в Mailu (TLS_FLAVOR=notls)
# Если порт 143 не работает, попробуйте 993 с USE_SSL = False

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

