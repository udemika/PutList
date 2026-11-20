@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo ============================================================================
echo   СБОРКА ОБНОВЛЕННОГО EXE ДЛЯ WINDOWS 11: Про.Путевые Листы.exe
echo   (Легковые авто Форма № 3 + Трактора МТЗ-82 и Экскаваторы 412-АПК / ЭСМ-2)
echo ============================================================================
echo.

REM 1. Поиск Python (python или py)
set "PY_CMD="
where python >nul 2>&1
if %errorlevel%==0 (
    set "PY_CMD=python"
) else (
    where py >nul 2>&1
    if %errorlevel%==0 (
        set "PY_CMD=py -3"
    )
)

if "%PY_CMD%"=="" (
    echo [ОШИБКА] Python не найден в системе!
    echo Установите Python 3.11 или 3.12 с сайта https://www.python.org/downloads/windows/
    echo При установке обязательно поставьте галочку "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

echo [1/4] Найден интерпретатор: %PY_CMD%
%PY_CMD% --version
echo.

echo [2/4] Установка и проверка библиотек (requirements.txt + pyinstaller)...
%PY_CMD% -m pip install --upgrade pip
%PY_CMD% -m pip install -r requirements.txt pyinstaller
if %errorlevel% neq 0 (
    echo.
    echo [ОШИБКА] Не удалось установить зависимости. Проверьте подключение к интернету.
    pause
    exit /b 1
)
echo.

echo [3/4] Сборка нового Про.Путевые Листы.exe через PyInstaller...
if exist "build" rmdir /s /q "build"
if exist "dist\Про.Путевые Листы" rmdir /s /q "dist\Про.Путевые Листы"

%PY_CMD% -m PyInstaller --noconfirm --clean --windowed --onedir ^
    --name "Про.Путевые Листы" ^
    --collect-all customtkinter ^
    --collect-all tkintermapview ^
    --collect-all tkcalendar ^
    --hidden-import openpyxl ^
    --hidden-import PIL._tkinter_finder ^
    --add-data "templates;templates" ^
    main.py

if %errorlevel% neq 0 (
    echo.
    echo [ОШИБКА] Сборка EXE завершилась с ошибкой. Смотрите сообщения выше.
    pause
    exit /b 1
)
echo.

echo [4/4] Копирование баз данных, настроек и бланков рядом с новым EXE...
set "OUT_DIR=dist\Про.Путевые Листы"
if not exist "%OUT_DIR%\templates" mkdir "%OUT_DIR%\templates"
xcopy /e /i /y /q "templates" "%OUT_DIR%\templates" >nul

if exist "waybills.db" copy /y "waybills.db" "%OUT_DIR%\" >nul
if exist "tractors.db" copy /y "tractors.db" "%OUT_DIR%\" >nul
if exist "memory.json" copy /y "memory.json" "%OUT_DIR%\" >nul
if exist "memory_tractors.json" copy /y "memory_tractors.json" "%OUT_DIR%\" >nul

REM Обновляем EXE и _internal также в корневой папке проекта
taskkill /f /im "Про.Путевые Листы.exe" >nul 2>&1
copy /y "%OUT_DIR%\Про.Путевые Листы.exe" "Про.Путевые Листы.exe" >nul
if exist "_internal" rmdir /s /q "_internal"
xcopy /e /i /y /q "%OUT_DIR%\_internal" "_internal" >nul

REM Удаляем временные файлы сборки
if exist "build" rmdir /s /q "build"
if exist "Про.Путевые Листы.spec" del /f /q "Про.Путевые Листы.spec"

echo.
echo ============================================================================
echo   ГОТОВО! Обновленный файл "Про.Путевые Листы.exe" успешно собран!
echo.
echo   1) В текущей папке обновлен "Про.Путевые Листы.exe" — можно запускать двойным кликом.
echo   2) В папке "%OUT_DIR%\" лежит готовая переносимая сборка для других ПК Windows 11.
echo ============================================================================
echo.
pause
