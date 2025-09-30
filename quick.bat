@echo off
chcp 65001 >nul
title Registrum

echo Запуск Registrum...
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)
python app.py