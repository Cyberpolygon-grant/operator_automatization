#!/usr/bin/env python3
"""
Конфигурация для автоматизации оператора ДБО через Mailu API
"""

import os
from pathlib import Path

# Настройки почты оператора ДБО
EMAIL_ADDRESS = "operator1@financepro.ru"
EMAIL_PASSWORD = "1q2w#E$R"

# Настройки Mailu API
# Mailu доступен на удаленной машине (10.18.2.6)
# API работает через HTTP/HTTPS (порты 80/443), которые уже открыты
API_URL = "http://10.18.2.6/api"  # URL Mailu API
API_TOKEN = "admin"  # API токен из mailu.env (для тестирования)
# Если API_TOKEN пустой, будет использоваться базовая аутентификация (email/password)

# Директория для сохранения скачанных файлов
USER_HOME = Path.home()
DOWNLOAD_DIR = str(USER_HOME / "Downloads")

# Интервал проверки почты (в секундах)
CHECK_INTERVAL = 30  # Проверка каждые 30 секунд

# Автоматически открывать Excel файлы
AUTO_OPEN_EXCEL = True

