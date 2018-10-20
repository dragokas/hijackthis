@echo off
SetLocal EnableExtensions
cd /d "%~dp0"

echo -----------------------------------------
echo New Certificates Checker by Alex Dragokas
echo -----------------------------------------
echo.

if "%~1" neq "Admin" (
  call :GetPrivileges || exit /b
)

set "certold=tools\cert\ms"
set "certnew=tools\cert\ms-new"
set "logfile=%temp%\cert.log"

2>NUL md "%certnew%"
2>NUL del "%logfile%"

echo This tool will download and check for new certificates of such types:
echo.
echo 1. TRUSTED ROOT
echo 2. DISALLOWED
echo.
echo Downloading to "%certnew%" ...
certutil -syncWithWU "%certnew%"

echo Generating cert from container...
certutil -generateSSTFromWU "%certnew%\WURoots.sst"
echo.

if not exist "%certnew%\disallowedcert.sst" (
  echo FAILURE! Cannot download disallowedcert.sst && echo. & pause>NUL
  exit /b
)

if not exist "%certnew%\WURoots.sst" (
  echo FAILURE! Cannot download WURoots.sst && echo. & pause>NUL
  exit /b
)

set "bRefreshRequired="
>NUL fc /b "%certold%\disallowedcert.sst" "%certnew%\disallowedcert.sst" && echo Disallowed certs are up-to-date. || (
  set bRefreshRequired=true
  echo NEW DISALLOWED CERT FOUND !!! ^(see: disallowedcert.sst^)
  echo It's mean you need manually add them to source code in 'modMain.bas' colSafeCert collection.
  call :dlg "Open .sst container file? (Y/N)" && start "" "%certnew%\disallowedcert.sst"
)
echo.
>NUL fc /b "%certold%\WURoots.sst" "%certnew%\WURoots.sst" && echo Trusted certs are up-to-date. || (
  set bRefreshRequired=true
  echo NEW TRUSTED CERT FOUND !!! ^(see: WURoots.sst^)
  echo It's mean you need manually check are there any new Microsoft Root cert. If so, please add it to source code 'modVerifyDigiSign.bas' IsMicrosoftCertHash^(^).
  call :dlg "Open .sst container file? (Y/N)" && start "" "%certnew%\WURoots.sst"
)

echo.
echo List of new certificates:
echo.

set n=0
for %%a in ("%certnew%\*.crt") do if not exist "%certold%\%%~nxa" (
  echo Fingerprint: %%~na
  echo Fingerprint: %%~na>> "%logfile%"
  certutil -dump "%%a" | find /i "CN=" | find /n /v "" | find "[1]"
  certutil -dump "%%a" | find /i "CN=" | find /n /v "" | find "[1]" >> "%logfile%"
  echo -
  echo ->> "%logfile%"
  set /a n+=1
)
if %n%==0 (
  echo No new cert was found.
) else (
  echo WARNING: this list contains both trusted root and disallowed if found.
  echo WARNING: this list contains both trusted root and disallowed if found.>>"%logfile%"
  echo So, you need manually check is it 'Trusted root' or 'Disallowed' certificate !!!
  echo So, you need manually check is it 'Trusted root' or 'Disallowed' certificate !!!>>"%logfile%"
)

echo.
if exist "%logfile%" call :dlg "Open log-file? (Y/N) " && start "" "%logfile%"

echo.
if defined bRefreshRequired call :dlg "Refresh certificates in './tools/cert' storage? (Y/N) " && (
  copy /y "%certnew%\*" "%certold%\" && echo Successfully copied to "%certold%"
)

rd /s /q "%certnew%" && echo Temp. certs are removed.
echo.

pause
goto :eof

:dlg [msg]
  set "ch="
  set /p "ch=%~1"
  if /i "%ch%"=="Y" exit /b 0
exit /b 1

:GetPrivileges
  net session >NUL 2>NUL || (
    echo.
    echo Administrative privileges required.
    echo.
    mshta "vbscript:CreateObject("Shell.Application").ShellExecute("%~fs0", "Admin", "", "runas", 1) & Close()"
    exit /B 1
  )
exit /B