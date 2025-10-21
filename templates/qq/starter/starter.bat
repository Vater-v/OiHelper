@echo PODOJDITE POJALUYSTA
@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

REM Путь к EXE рядом с батником
set "EXE=%~dp0Starter.exe"
if not exist "%EXE%" (
  echo [ERROR] Не найден EXE: "%EXE%"
  exit /b 1
)

REM ===== Значения по умолчанию (можно переопределять извне) =====
if not defined ACCOUNT  set "ACCOUNT=TEST"
if not defined DURATION set "DURATION=10"
if not defined TITLE    set "TITLE=Проверка"
if not defined MESSAGE  set "MESSAGE=Идёт тест 10 сек"
REM опционально:
REM set "FLAGFILE=C:\Temp\ready_%ACCOUNT%.flag"
REM set "ICON=C:\Path\icon.ico"
REM set "TOPMOST=1"  (0 — не ставить поверх)

REM ===== Запуск =====
"%EXE%" --account "%ACCOUNT%" --duration "%DURATION%" ^
  %IFDEF_TITLE% --title "%TITLE%" ^
  %IFDEF_MESSAGE% --message "%MESSAGE%" ^
  %IFDEF_FLAGFILE% --flag-file "%FLAGFILE%" ^
  %IFDEF_ICON% --icon "%ICON%" ^
  %IFDEF_TOPMOST% --topmost

endlocal
