#!/usr/bin/env python3
"""
Вспомогательный скрипт для загрузки шрифта Chakra Petch
"""

import os
import requests
from pathlib import Path


def download_chakra_font():
    """Скачивает шрифт Chakra Petch если он отсутствует"""
    assets_dir = Path(__file__).parent / "assets"
    assets_dir.mkdir(exist_ok=True)

    font_url = "https://github.com/google/fonts/raw/main/ofl/chakrapetch/ChakraPetch-Regular.ttf"
    font_path = assets_dir / "ChakraPetch-Regular.ttf"

    if font_path.exists():
        print("✓ Шрифт ChakraPetch уже установлен")
        return True

    print("Загрузка шрифта ChakraPetch...")
    try:
        response = requests.get(font_url)
        response.raise_for_status()

        with open(font_path, 'wb') as f:
            f.write(response.content)

        print("✓ Шрифт успешно загружен")
        return True
    except Exception as e:
        print(f"✗ Ошибка загрузки шрифта: {e}")
        print("Программа будет использовать стандартные шрифты")
        return False


if __name__ == "__main__":
    download_chakra_font()