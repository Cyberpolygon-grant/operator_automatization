#!/usr/bin/env python3
"""
Упрощенный запуск автоматизации оператора ДБО
Использует настройки из dbo_operator_config.py
"""

import sys
import os
from pathlib import Path

# Добавляем текущую директорию в путь для импорта
sys.path.insert(0, str(Path(__file__).parent))

from dbo_operator_automation import DBOOperatorAutomation
import dbo_operator_config as config

def main():
    """Запуск автоматизации с настройками из конфига"""
    
    print("=" * 60)
    print("АВТОМАТИЗАЦИЯ ОПЕРАТОРА ДБО")
    print("=" * 60)
    print(f"Email: {config.EMAIL_ADDRESS}")
    print(f"IMAP сервер: {config.IMAP_SERVER}:{config.IMAP_PORT}")
    print(f"Директория загрузки: {config.DOWNLOAD_DIR}")
    print(f"Интервал проверки: {config.CHECK_INTERVAL} сек")
    print(f"Автооткрытие Excel: {'Да' if config.AUTO_OPEN_EXCEL else 'Нет'}")
    print(f"Режим: {config.MODE}")
    print("=" * 60)
    print()
    
    # Создаем директорию, если её нет
    download_path = Path(config.DOWNLOAD_DIR)
    download_path.mkdir(parents=True, exist_ok=True)
    print(f"✓ Директория загрузки создана: {download_path}")
    print()
    
    # Создаем экземпляр автоматизации
    automation = DBOOperatorAutomation(
        email_address=config.EMAIL_ADDRESS,
        password=config.EMAIL_PASSWORD,
        imap_server=config.IMAP_SERVER,
        imap_port=config.IMAP_PORT,
        download_dir=config.DOWNLOAD_DIR
    )
    
    # Запускаем в зависимости от режима
    if config.MODE == 'once':
        print("Выполнение однократной проверки почты...")
        files = automation.run_once(auto_open=config.AUTO_OPEN_EXCEL)
        print(f"\n✓ Обработано файлов: {len(files)}")
        if files:
            print("\nСкачанные файлы:")
            for f in files:
                print(f"  - {f}")
    else:
        print("Запуск непрерывной проверки почты...")
        print("Нажмите Ctrl+C для остановки\n")
        automation.run_continuous(
            check_interval=config.CHECK_INTERVAL,
            auto_open=config.AUTO_OPEN_EXCEL
        )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nОстановка по запросу пользователя")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        sys.exit(1)

