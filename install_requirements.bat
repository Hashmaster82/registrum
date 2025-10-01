@echo off
setlocal

echo Установка зависимостей из requirements.txt...

:: Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Ошибка: Python не найден. Убедитесь, что Python установлен и добавлен в PATH.
    pause
    exit /b 1
)

:: Проверяем наличие requirements.txt
if not exist "requirements.txt" (
    echo Ошибка: Файл requirements.txt не найден в текущей папке.
    pause
    exit /b 1
)

:: Обновляем pip и устанавливаем зависимости
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo Все зависимости успешно установлены!
pause