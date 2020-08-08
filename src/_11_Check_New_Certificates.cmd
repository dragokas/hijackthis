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

echo Press Y to download new certificates now.
echo Press N if you want to manually copy fresh certificates to folder: "%certnew%"
echo (useful, when you are offline, or you are expecting some problems with automatic downloading)
echo.
set "ch="
set /p "ch=Your choice (Y/N): "
echo.
if /i "%ch%"=="Y" (

  echo Downloading to "%certnew%" ...
  certutil -syncWithWU "%certnew%"

  echo Generating cert from container...
  certutil -generateSSTFromWU "%certnew%\WURoots.sst"
  echo.
)

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
  echo It's mean you need manually add them to source code in 'modMain.bas' colDisallowedCert collection.
  echo To simplify procedure:
  echo.
  echo 1. Open disallowedcert.sst, select all, and export them to file.
  echo 2. Import that storage to your system in "Disallowed" section. Use certmgr.msc to verify.
  echo 3. Extract cert info using tools\Cert\enumerator\Disallowed\DisallowedCertEnumerator.exe
  echo 4. Update hjt.txt with HJT source code of "colDisallowedCert" collection contents.
  echo 5. Run "Compare-cert.exe" and receive Hashes.txt to append HJT source code with.
  echo.
  call :dlg "Open .sst container file? (Y/N)" && start "" "%certnew%\disallowedcert.sst"
)
echo.
>NUL fc /b "%certold%\WURoots.sst" "%certnew%\WURoots.sst" && echo Trusted certs are up-to-date. || (
  set bRefreshRequired=true
  echo NEW TRUSTED CERT FOUND !!! ^(see: WURoots.sst^)
  echo It's mean you need manually check are there any new Microsoft Root cert. If so, please add it to source code 'modVerifyDigiSign.bas' IsMicrosoftCertHash^(^).
  echo To simplify procedure:
  echo.
  echo 1. Open WURoots.sst, select all, and export them to file.
  echo 2. Import that storage to your system in "Trusted Root" section. Use certmgr.msc to verify. 
  echo 3. Extract cert info using tools\Cert\enumerator\Root\CertEnumerator.exe
  echo 4. Compare HJT source code with Hashes.csv and append with new certs.
  echo.
  echo Also, this script is about to create delta, you can use it to found new root certificates.
  echo.
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