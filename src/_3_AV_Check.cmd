@rem For stand-alone updated version, refer to:
@rem https://www.safezone.cc/resources/virustotal-console-checker.233/
@rem https://github.com/VirusTotal/vt-cli/releases
@echo off
SetLocal EnableExtensions
pushd "%~dp0"
::title VirusTotal Console Checker v1.3
echo VirusTotal check ... & echo.

set vt_cli=tools\VirusTotal\vt.exe
set "filepath=%~f1"
if "%filepath%"=="" set "filepath=%~DP0HiJackThis.exe"
set "hash="

if not exist "%vt_cli%" (
  echo vt.exe doesn't found!
  echo Download it at: https://github.com/VirusTotal/vt-cli/releases
  pause & goto :eof
)
if not exist "%UserProfile%\.vt.toml" %vt_cli% init

:enter_path
  if not defined filepath (set /p "filepath=Enter file path to check on VT: " & goto enter_path)
  if not exist "%filepath%" (echo Cannot found file: "%filepath%" & set "filepath=" & goto enter_path)
  echo.
  echo Checking file: "%filepath%" ...

:: if the file already verified, vt.exe file returns detailed data
:: so we can use hash+date submission directly in call to vt.exe analysis
call :gethash "%filepath%" hash
"%vt_cli%" file %hash% 1> details.log 2> err.log
<err.log find /i "not found" && goto scan

echo Found result
for /f "tokens=1,2 delims=: " %%a in (details.log) do if "%%~a"=="last_analysis_date" set "date_analysis=%%b"
set id=f-%hash%-%date_analysis%
goto analysis

:scan
  del details.log
  echo Send ...
  echo.
  for /f "delims=" %%a in ('%vt_cli% scan file "%filepath%"') do set "result=%%a"
:getid
  for /f "tokens=1* delims= " %%a in ("%result%") do set "result=%%b" & if not defined result (set id=%%a) else (goto getid)
  echo Got file ID: "%id%"

::in case file already scanned
:analysis
  del err.log
  ::echo id=%id%
  %vt_cli% analysis %id% > vt.log
  < vt.log find /i "completed" && goto parse
  < vt.log find /i "status:"
  echo Waiting for analysis ...

timeout /t 20 >NUL

:wait_queue
  timeout /t 5 >NUL
  %vt_cli% analysis %id% > vt.log
  < vt.log find /i "queued" && goto wait_queue

:parse
set "is_malicious="
set "is_suspicious="
set "is_failure="
set "is_timeout="
set "is_stats="
echo.
for /f tokens^=1^,2*^ delims^=^"^  %%a in (vt.log) do call :t_line "%%a" "%%b"
del vt.log

:sel
  echo.
  echo 1. Open logfile with extended info
  echo 2. Open VirusTotal link to file
  echo 3. Exit
  choice /C 123 /N
  if %errorlevel%==1 call :details
  if %errorlevel%==2 call :vt_open
  if %errorlevel%==3 (popd & goto :eof)
goto sel

:t_line
  if defined is_malicious (
	if "%~1"=="engine_name:" set "engine_name=%~2"
	if "%~1"=="engine_update:" set "engine_update=%~2"
	if "%~1"=="engine_version:" set "engine_version=%~2"
	if "%~1"=="result:" (
	  echo ! MALICIOUS: %engine_name%: "%~2"   ^(%engine_update%, %engine_version%^)
	  set "is_malicious="
	)
  )
  if defined is_suspicious (
    if "%~1"=="engine_name:" set "engine_name=%~2"
	if "%~1"=="engine_update:" set "engine_update=%~2"
	if "%~1"=="engine_version:" set "engine_version=%~2"
	if "%~1"=="result:" (
	  echo ! SUSPICIOUS: %engine_name%: "%~2"   ^(%engine_update%, %engine_version%^)
	  set "is_suspicious="
	)
  )
  if defined is_timeout (
    if "%~1"=="engine_name:" (echo %~2 ^(timeout^)& set "is_timeout=")
  )
  if defined is_failure (
    if "%~1"=="engine_name:" (echo %~2 ^(failure^)& set "is_failure=")
  )
  if defined is_stats echo %~1 %~2
  if "%~1"=="category:" if "%~2"=="malicious" set "is_malicious=true"
  if "%~1"=="category:" if "%~2"=="suspicious" set "is_suspicious=true"
  if "%~1"=="category:" if "%~2"=="failure" set "is_failure=true"
  if "%~1"=="category:" if "%~2"=="timeout" set "is_timeout=true"
  if "%~1"=="category:" if "%~2"=="confirmed-timeout" set "is_timeout=true"
  if "%~1"=="stats:" (set "is_stats=true" & echo.)
exit /b

:details
  if exist details.log goto show_details
  call :gethash "%filepath%" hash
  "%vt_cli%" file %hash% > details.log
:show_details
  start "" details.log
  echo details.log is created and opened.
exit /b

:vt_open
  if not defined hash call :gethash "%filepath%" hash
  set vt_link=https://www.virustotal.com/gui/file/%hash%/detection
  echo VirusTotal link: %vt_link%
  start "" "%vt_link%"
exit /b

:gethash [file] [out_var]
  for /f "delims=" %%a in ('certutil -hashfile "%~1" SHA256 ^| find /v "hash"') do set "t=%%a"
  set "t=%t: =%"
  ::echo Checksum SHA256: "%t%"
  set "%~2=%t%"
exit /b