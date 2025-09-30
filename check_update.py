#!/usr/bin/env python3
"""
Скрипт проверки обновлений на GitHub
"""

import requests
import json
import os
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import webbrowser


def check_for_updates():
    """Проверяет наличие обновлений на GitHub"""
    repo_owner = "Hashmaster82"  # ЗАМЕНИТЕ на ваш GitHub username
    repo_name = "registrum"  # ЗАМЕНИТЕ на название репозитория

    current_version = "1.0.0"  # Текущая версия программы

    print(f"Текущая версия: {current_version}")
    print("Проверка обновлений...")

    try:
        # Получаем информацию о последнем релизе
        url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest"
        response = requests.get(url, timeout=10)

        if response.status_code == 200:
            release_info = response.json()
            latest_version = release_info['tag_name'].lstrip('v')
            release_url = release_info['html_url']

            print(f"Последняя версия на GitHub: {latest_version}")

            # Сравниваем версии (простое сравнение строк)
            if latest_version > current_version:
                print("⚠ Доступно обновление!")

                # Показываем диалог об обновлении
                root = tk.Tk()
                root.withdraw()  # Скрываем основное окно

                update_info = f"""
Доступна новая версия программы!

Текущая версия: {current_version}
Новая версия: {latest_version}

Хотите перейти на страницу загрузки?
                """.strip()

                result = messagebox.askyesno(
                    "Доступно обновление",
                    update_info,
                    icon='info'
                )

                if result:
                    webbrowser.open(release_url)
                    print("Открыта страница загрузки обновления")
                else:
                    print("Обновление отклонено пользователем")

            else:
                print("✓ У вас актуальная версия программы")

        else:
            print("ℹ Не удалось проверить обновления (проблемы с подключением)")

    except requests.exceptions.RequestException as e:
        print(f"ℹ Не удалось проверить обновления: {e}")
    except Exception as e:
        print(f"ℹ Ошибка при проверке обновлений: {e}")


def check_dependencies():
    """Проверяет установлены ли необходимые зависимости"""
    try:
        import requests
        return True
    except ImportError:
        print("⚠ Библиотека 'requests' не установлена")
        print("Установка необходимых зависимостей...")

        try:
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", "requests"])
            print("✓ Зависимости установлены")
            return True
        except Exception as e:
            print(f"✗ Ошибка установки зависимостей: {e}")
            return False


if __name__ == "__main__":
    import sys

    # Проверяем зависимости
    if not check_dependencies():
        print("Не удалось установить зависимости для проверки обновлений")
        sys.exit(1)

    # Проверяем обновления
    check_for_updates()