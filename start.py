#!/usr/bin/env python3
"""
Скрипт для простого запуска приложения
Просто запустите этот файл двойным кликом!
"""

import os
import sys
import subprocess
import platform

def check_python():
    """Проверяет версию Python"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("❌ Требуется Python 3.8 или выше")
        print(f"   Ваша версия: {version.major}.{version.minor}.{version.micro}")
        print("\nСкачайте Python с https://www.python.org/downloads/")
        input("\nНажмите Enter для выхода...")
        sys.exit(1)
    print(f"✅ Python {version.major}.{version.minor}.{version.micro} найден")

def install_dependencies():
    """Устанавливает зависимости"""
    print("\n📦 Проверка зависимостей...")

    try:
        import flask
        import openpyxl
        import docx
        import webview
        print("✅ Все зависимости установлены")
        return True
    except ImportError:
        print("\n⚠️  Необходимо установить зависимости")
        response = input("Установить сейчас? (y/n): ").lower()

        if response == 'y' or response == 'yes' or response == 'д' or response == 'да':
            print("\n📥 Установка зависимостей...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
                print("\n✅ Зависимости успешно установлены!")
                return True
            except subprocess.CalledProcessError:
                print("\n❌ Ошибка при установке зависимостей")
                print("Попробуйте установить вручную: pip3 install -r requirements.txt")
                input("\nНажмите Enter для выхода...")
                return False
        else:
            print("\n❌ Невозможно запустить без зависимостей")
            input("\nНажмите Enter для выхода...")
            return False

def start_app():
    """Запускает приложение"""
    print("\n🚀 Запуск приложения...")
    print("   Откроется окно с приложением")
    print("   Чтобы остановить приложение, закройте окно приложения\n")

    try:
        # Запускаем gui_app.py
        subprocess.run([sys.executable, "gui_app.py"])
    except KeyboardInterrupt:
        print("\n\n👋 Приложение остановлено")
    except Exception as e:
        print(f"\n❌ Ошибка при запуске: {e}")
        input("\nНажмите Enter для выхода...")

def main():
    print("=" * 50)
    print("  Генератор актов Wildberries")
    print("=" * 50)

    # Меняем рабочую директорию на директорию скрипта
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # Проверяем Python
    check_python()

    # Устанавливаем зависимости если нужно
    if not install_dependencies():
        return

    # Запускаем приложение
    start_app()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Непредвиденная ошибка: {e}")
        input("\nНажмите Enter для выхода...")
