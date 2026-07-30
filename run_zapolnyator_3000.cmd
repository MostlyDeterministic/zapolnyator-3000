@echo off
setlocal
chcp 65001 >nul

title Заполнятор 3000
set "SCRIPT=%~dp0run_fill_signals_ods.ps1"

if not exist "%SCRIPT%" (
    echo [ОШИБКА] Не найден файл: "%SCRIPT%"
    echo.
    echo Нажмите любую клавишу для закрытия окна...
    pause >nul
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -NoPause
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo Заполнятор 3000 завершился с ошибкой. Код: %EXIT_CODE%
)

echo Нажмите любую клавишу для закрытия окна...
pause >nul
exit /b %EXIT_CODE%
