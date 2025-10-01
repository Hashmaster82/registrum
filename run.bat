@echo off
chcp 65001 >nul
title Registrum App Launcher

echo ========================================
echo      Registrum - Запуск приложения
echo ========================================
echo.

set "REPO_URL=https://github.com/Hashmaster82/registrum.git"
set "SCRIPT_NAME=app.py"
set "TEMP_DIR=%TEMP%\registrum_update"

:: Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не установлен или не добавлен в PATH
    echo Установите Python и повторите попытку
    pause
    exit /b 1
)

:: Создаем временную папку
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

if exist "registrum" (
    cd registrum
    git fetch origin master --quiet
    if errorlevel 1 (
        echo Не удалось получить данные из репозитория
        cd /d "%~dp0"
        goto :run_app
    )
) else (
    git clone --depth 1 --quiet %REPO_URL% registrum
    if errorlevel 1 (
        echo Не удалось клонировать репозиторий
        cd /d "%~dp0"
        goto :run_app
    )
    cd registrum
)

:: Получаем хеш последнего коммита в ветке master
for /f "delims=" %%i in ('git rev-parse --short origin/master 2^>nul') do set "LATEST_COMMIT=%%i"

cd /d "%~dp0"

:: Читаем текущую версию (только хеш)
set "CURRENT_COMMIT="
if exist "version.txt" (
    set /p CURRENT_COMMIT=<version.txt
)

if defined CURRENT_COMMIT (
    echo Текущая версия: %CURRENT_COMMIT%
) else (
    echo Информация о версии не найдена
)

if defined LATEST_COMMIT (
    echo Последняя версия: %LATEST_COMMIT%
) else (
    echo Не удалось определить последнюю версию
    goto :run_app
)

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

    :: Получаем временную метку в формате YYYYMMDD_HHMM (без кириллицы!)
    for /f "delims=" %%i in ('powershell -command "Get-Date -Format 'yyyyMMdd_HHmm'"') do set "TIMESTAMP=%%i"
    set "BACKUP_FOLDER=backup\registrum_backup_%TIMESTAMP%"

    :: Создаем резервную копию
    if not exist "backup" mkdir "backup"
    echo Создаем резервную копию в %BACKUP_FOLDER%
    xcopy * "%BACKUP_FOLDER%" /E /I /Y /Q >nul 2>&1

    :: Копируем новые файлы из временного репозитория
    xcopy "%TEMP_DIR%\registrum\*" "%~dp0" /E /Y /Q >nul 2>&1

    :: Сохраняем только хеш коммита как текущую версию
    echo %LATEST_COMMIT% > version.txt

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
    echo Содержимое папки:
    dir /b *.py
    echo.
    pause
    exit /b 1
)

:: Проверяем зависимости
echo Проверяем зависимости Python...
python -c "import reportlab, openpyxl, matplotlib" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Устанавливаем необходимые зависимости...
    pip install reportlab openpyxl matplotlib --quiet
    if errorlevel 1 (
        echo Не удалось установить зависимости. Попробуйте вручную:
        echo pip install reportlab openpyxl matplotlib
        pause
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