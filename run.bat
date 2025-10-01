@echo off
chcp 65001 >nul
title Registrum - Запуск приложения

echo ========================================
echo        REGISTRUM - ЗАПУСК ПРОГРАММЫ
echo ========================================
echo.

:: Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не установлен или не добавлен в PATH
    echo Установите Python с официального сайта:
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✓ Python обнаружен

:: Проверка виртуального окружения
if exist venv\ (
    echo Активация виртуального окружения...
    call venv\Scripts\activate.bat
) else (
    echo Создание виртуального окружения...
    python -m venv venv
    call venv\Scripts\activate.bat
    echo Установка зависимостей...
    pip install -r requirements.txt
)

:: Проверка обновлений на GitHub
echo.
echo Проверка обновлений на GitHub...
python check_update.py

:: Проверка шрифта
if not exist assets\ChakraPetch-Regular.ttf (
    echo.
    echo Установка шрифта ChakraPetch...
    python install_font.py
)

:: Запуск приложения
echo.
echo ========================================
echo          ЗАПУСК ПРОГРАММЫ
echo ========================================
echo.
python app.py

:: Пауза после закрытия программы
echo.
echo Программа завершена.
pause