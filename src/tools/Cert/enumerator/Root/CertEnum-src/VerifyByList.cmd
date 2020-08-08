@echo off
SetLocal EnableExtensions
cd /d "%~dp0"

if not exist apps.txt (

  echo Creating List. Please wait...

  dir /b /a-d /s "%SystemDrive%\*.exe" "%SystemRoot%\*.sys" > apps.txt
  echo.
  set /p "=Ready! Total lines = " <NUL

  find /c /v "" < apps.txt

  echo.
  echo.

  pause
)


:: Clear
if exist report_sign_signed.txt del report_sign_signed.txt
if exist report_sign_bad.txt del report_sign_bad.txt
if exist report_sign_error.txt del report_sign_error.txt
if exist report_sign_not_signed.txt del report_sign_not_signed.txt

For /F "delims=" %%a in ('find /c /v "" ^< apps.txt') do set Total=%%a


echo.
echo.
echo Begin digital signature checking...
echo.
echo.

set t0=%time%

For /F "UseBackQ delims=" %%a in ("apps.txt") do call :CheckOne "%%~a"

echo.
echo.
echo All Done!
set t1=%time%
call :difftime
echo.
echo.
pause

goto :eof



:CheckOne
  set /a n+=1
  echo %n% / %Total%

  SignVer.exe "%~1"

  if %errorlevel%==1000 (
    echo "%~1">> report_sign_signed.txt
  ) else (
  if %errorlevel%==1001 (
    echo "%~1">> report_sign_not_signed.txt
  ) else (
  if %errorlevel%==1002 (
    echo "%~1">> report_sign_bad.txt
  ) else (
  if %errorlevel%==1 (
    echo "%~1">> report_sign_error.txt
  ))))

exit /B

:difftime
  for /F "tokens=1-8 delims=:.," %%a in ("%t0: =0%:%t1: =0%") do set /a "a=(((1%%e-1%%a)*60)+1%%f-1%%b)*6000+1%%g%%h-1%%c%%d, a+=(a>>31) & 8640000, a/=100"
  echo Time spent: %a% seconds.
goto :eof