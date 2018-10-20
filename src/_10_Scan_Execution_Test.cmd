@echo off
SetLocal EnableExtensions

if "%~1"=="" goto DontAsk

set "ch="
set /p "ch=Execute HJT scan test (y/n) ? "
if /i "%ch%" neq "y" exit /b

:DontAsk

2>NUL del "HiJackThis.log"
2>NUL del "HiJackThis_debug.log"

start "" /w "HiJackThis.exe" /accepteula /silentautolog /default /skipIgnoreList /timeout:57 /debugtofile

echo.

if not exist "HiJackThis.log" (
  echo Critical error!
  echo No HiJackThis.log file.
  pause>NUL
)

echo.

if not exist "HiJackThis_debug.log" (
  echo Critical error!
  echo No HiJackThis_debug.log file.
  pause>NUL
)

if exist "HiJackThis_debug.log" start "" "HiJackThis_debug.log"
if exist "HiJackThis.log" start "" "HiJackThis.log"

