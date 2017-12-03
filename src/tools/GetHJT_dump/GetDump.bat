@echo off
SetLocal EnableExtensions

cd /d "%~dps0" || cd /d "%~dp0"

if not exist "procdump.exe" (
  echo You need to unpack this archive first
  echo.
  echo Вам нужно сперва распаковать архив
  pause >NUL
  exit /b
)

if not exist "HiJackThis.exe" (
  echo You need to copy HiJackThis.exe to this folder first
  echo.
  echo Вам нужно сперва скопировать программу HiJackThis.exe ^(из папки ...\Autologger\HiJackThis\ ^) в эту папку
  pause >NUL
  exit /b
)

net session >NUL 2>NUL || (
  echo Run this file as Administrator !!!
  echo.
  echo Запускайте этот файл правой кнопкой мыши "От имени администратора" !!!
  pause >NUL
  exit /b
)

procdump.exe -accepteula -ma -l -o -e -w -x . HiJackThis.exe /silentautolog /debug

ping 127.1 -n 5

2>NUL md helper

call :VerCheck "%~dps0" && copy /y HiJackThis.exe helper\Jack.exe

ren helper\Jack.exe _poly.exe

helper\_poly.exe /silentautolog /debug

ren helper\_poly.exe Jack.exe

ren helper\HiJackThis*.log HiJackThis*.poly.log 

7za.exe a -ssw -mx5 -y Dumps.zip HiJackThis*.dmp HiJackThis*.log helper\HiJackThis*.log

exit /b

:VerCheck [file]
  :: не ниже 2.6.4.23 [24.04.2017]
  For /f "tokens=1-3 delims=. " %%a in ('dir "%~1\*Jack*.exe"^| find "Jack"') do set YYYY=%%c& set MM=%%b& set DD=%%a
  if not defined YYYY For /f "tokens=1-3 delims=/ " %%a in ('dir "%~1\*Jack*.exe"^| find "Jack"') do set YYYY=%%c& set MM=%%a& set DD=%%b
  if %YYYY% LSS 2017 exit /b 1
  if %MM% LSS 4 exit /b 1
  if %DD% LSS 24 exit /b 1
exit /b 0
