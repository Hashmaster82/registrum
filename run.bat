@echo off
setlocal

:: Путь к текущей папке
set "SCRIPT_DIR=%~dp0"

:: Проверяем, инициализирован ли git-репозиторий
if not exist ".git" (
    echo Репозиторий не найден. Обновление невозможно.
    echo Запуск приложения...
    python app.py
    goto :eof
)

echo Проверка обновлений в репозитории...

:: Получаем изменения из удалённого репозитория
git fetch origin

:: Сравниваем локальный и удалённый HEAD
git diff --quiet HEAD origin/main
if errorlevel 1 (
    echo.
    echo Обнаружены обновления!
    echo.
    set /p "choice=Хотите обновиться сейчас? (y/n): "
    if /i "%choice%"=="y" (
        echo Выполняется обновление...
        git pull origin main
        if errorlevel 1 (
            echo Ошибка при обновлении. Приложение запущено без обновления.
        ) else (
            echo Обновление завершено.
        )
    ) else (
        echo Обновление отменено.
    )
) else (
    echo Обновлений нет.
)

echo.
echo Запуск Registrum...
python app.py

pause