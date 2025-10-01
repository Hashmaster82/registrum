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
        goto :run_app
    )
) else (
    git clone --depth 1 --quiet %REPO_URL% registrum
    if errorlevel 1 (
        echo Не удалось клонировать репозиторий
        goto :run_app
    )
    cd registrum
)

:: Получаем хеш последнего коммита в master
for /f "delims=" %%i in ('git rev-parse --short origin/master 2^>nul') do set "LATEST_COMMIT=%%i"

cd /d "%~dp0"

:: Читаем текущую версию
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

:: Сравниваем
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

    :: Резервная копия
    set "BACKUP_FOLDER=backup\registrum_backup_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%"
    set "BACKUP_FOLDER=%BACKUP_FOLDER: =0%"
    if not exist "backup" mkdir "backup"
    echo Создаем резервную копию в %BACKUP_FOLDER%
    xcopy * "%BACKUP_FOLDER%" /E /I /Y /Q >nul 2>&1

    :: Копируем новые файлы
    xcopy "%TEMP_DIR%\registrum\*" "%~dp0" /E /Y /Q >nul 2>&1

    :: Сохраняем новую версию
    echo %LATEST_COMMIT% > version.txt

    echo Обновление завершено!
    echo Резервная копия: %BACKUP_FOLDER%
)

:run_app
echo.
echo Запуск приложения...
echo.

if not exist "%SCRIPT_NAME%" (
    echo ОШИБКА: Файл %SCRIPT_NAME% не найден
    pause
    exit /b 1
)

:: Проверка зависимостей
echo Проверяем зависимости...
python -c "import reportlab, openpyxl, matplotlib" >nul 2>&1
if errorlevel 1 (
    echo Устанавливаем зависимости...
    pip install reportlab openpyxl matplotlib --quiet
)

echo.
echo ========================================
echo       Запуск Registrum...
echo ========================================
python "%SCRIPT_NAME%"

if errorlevel 1 (
    echo.
    echo Ошибка при запуске приложения.
    pause
)

exit /b 0