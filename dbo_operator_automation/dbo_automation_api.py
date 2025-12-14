#!/usr/bin/env python3
"""
Автоматизация оператора ДБО через Mailu API
Скачивает файлы из почты и автоматически открывает Excel файлы
Логи выводятся в консоль
"""

import sys
import os
from pathlib import Path

# Добавляем текущую директорию в путь для импорта
sys.path.insert(0, str(Path(__file__).parent))

from dbo_operator_automation_api import DBOOperatorAutomationAPI
import dbo_operator_config_api as config

def main():
    """Запуск автоматизации"""
    
    print("=" * 60)
    print("АВТОМАТИЗАЦИЯ ОПЕРАТОРА ДБО (через Webmail)")
    print("=" * 60)
    print(f"Email: {config.EMAIL_ADDRESS}")
    print(f"Webmail URL: {config.WEBMAIL_URL}")
    print(f"Директория загрузки: {config.DOWNLOAD_DIR}")
    print(f"Интервал проверки: {config.CHECK_INTERVAL} сек")
    print(f"Автооткрытие Excel: {'Да' if config.AUTO_OPEN_EXCEL else 'Нет'}")
    print("=" * 60)
    print()
    print("Логи выводятся в консоль")
    print("Для остановки нажмите Ctrl+C")
    print()
    
    # Проверяем, что папка Downloads существует
    download_path = Path(config.DOWNLOAD_DIR)
    if not download_path.exists():
        download_path.mkdir(parents=True, exist_ok=True)
        print(f"✓ Папка Downloads создана: {download_path}")
    else:
        print(f"✓ Папка Downloads найдена: {download_path}")
    print()
    
    # Создаем экземпляр автоматизации
    automation = DBOOperatorAutomationAPI(
        email_address=config.EMAIL_ADDRESS,
        password=config.EMAIL_PASSWORD,
        webmail_url=config.WEBMAIL_URL,
        download_dir=config.DOWNLOAD_DIR
    )
    
    # Запускаем непрерывную проверку
    try:
        automation.run_continuous(
            check_interval=config.CHECK_INTERVAL,
            auto_open=config.AUTO_OPEN_EXCEL
        )
    except KeyboardInterrupt:
        print("\n\nОстановка по запросу пользователя (Ctrl+C)")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

