@echo off
chcp 65001 >nul
title Registrum App Launcher

echo ========================================
echo      Registrum - Запуск приложения
echo ========================================
echo.

set "REPO_URL=https://github.com/Hashmaster82/registrum.git"
set "SCRIPT_NAME=registrum_app.py"
set "TEMP_DIR=%TEMP%\registrum_update"

:: Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не установлен или не добавлен в PATH
    echo Установите Python и повторите попытку
    pause
    exit /b 1
)

:: Создаем временную папку для проверки обновлений
if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

:: Проверяем наличие Git
git --version >nul 2>&1
if errorlevel 1 (
    echo ВНИМАНИЕ: Git не установлен. Пропускаем проверку обновлений.
    echo Для автоматических обновлений установите Git
    goto :run_app
)

echo Проверяем наличие обновлений...
cd /d "%TEMP_DIR%"

:: Клонируем или обновляем репозиторий для проверки версии
if exist "registrum" (
    cd registrum
    git fetch origin >nul 2>&1
) else (
    git clone --depth 1 %REPO_URL% registrum >nul 2>&1
    if errorlevel 1 (
        echo Не удалось проверить обновления
        goto :run_app
    )
    cd registrum
)

:: Получаем информацию о последнем коммите
for /f "tokens=1" %%i in ('git log -1 --format^=%%h') do set "LATEST_COMMIT=%%i"
for /f "tokens=1,2" %%i in ('git log -1 --format^=%%ci') do set "LATEST_DATE=%%i %%j"

cd /d "%~dp0"

:: Проверяем текущую версию (если есть файл версии)
set "CURRENT_COMMIT="
set "CURRENT_DATE="
if exist "version.txt" (
    for /f "tokens=1,2" %%i in (version.txt) do (
        set "CURRENT_COMMIT=%%i"
        set "CURRENT_DATE=%%j"
    )
)

if "%CURRENT_COMMIT%"=="" (
    echo Информация о версии не найдена
) else (
    echo Текущая версия: %CURRENT_COMMIT% (%CURRENT_DATE%)
)

echo Последняя версия: %LATEST_COMMIT% (%LATEST_DATE%)

:: Сравниваем версии
if "%CURRENT_COMMIT%"=="%LATEST_COMMIT%" (
    echo У вас установлена последняя версия
) else (
    echo.
    echo Доступно обновление!
    echo.
    choice /c YN /n /m "Хотите обновиться? (Y-да, N-нет)"
    if errorlevel 2 goto :run_app

    echo.
    echo Выполняем обновление...

    :: Создаем резервную копию текущей версии
    if not exist "backup" mkdir "backup"
    set "BACKUP_FOLDER=backup\registrum_backup_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%"
    set "BACKUP_FOLDER=%BACKUP_FOLDER: =0%"

    echo Создаем резервную копию в %BACKUP_FOLDER%
    xcopy * "%BACKUP_FOLDER%" /E /I /Y >nul 2>&1

    :: Обновляем файлы
    cd /d "%TEMP_DIR%\registrum"
    xcopy * "%~dp0" /E /Y /Q >nul 2>&1

    :: Сохраняем информацию о новой версии
    cd /d "%~dp0"
    echo %LATEST_COMMIT% %LATEST_DATE% > version.txt

    echo Обновление завершено!
    echo Резервная копия сохранена в: %BACKUP_FOLDER%
)

:run_app
echo.
echo Запуск приложения...
echo.

:: Проверяем наличие основного скрипта
if not exist "%SCRIPT_NAME%" (
    echo ОШИБКА: Основной файл %SCRIPT_NAME% не найден
    echo.
    echo Попытка найти файл Python...
    dir *.py /b
    echo.
    echo Укажите правильное имя файла в переменной SCRIPT_NAME
    pause
    exit /b 1
)

:: Проверяем зависимости
echo Проверяем зависимости Python...
python -c "import tkinter, json, pathlib, datetime, re, shutil, reportlab, openpyxl, matplotlib" 2>nul
if errorlevel 1 (
    echo.
    echo Устанавливаем необходимые зависимости...
    pip install -r requirements.txt 2>nul
    if errorlevel 1 (
        echo Не удалось установить зависимости автоматически
        echo Устанавливаем вручную...
        pip install reportlab openpyxl matplotlib
    )
)

:: Запускаем приложение
echo.
echo ========================================
echo       Запуск Registrum...
echo ========================================
python "%SCRIPT_NAME%"

:: Если приложение завершилось с ошибкой
if errorlevel 1 (
    echo.
    echo Приложение завершилось с ошибкой (код: %errorlevel%)
    echo.
    pause
)

exit /b 0