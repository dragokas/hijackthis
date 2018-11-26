@echo off
SetLocal enableExtensions
echo.
::echo Checking for non-screened STOP statements ...
set "detected="

For %%a in (*.bas *.frm *.cls) do (
  set module=%%a
  for /F "tokens=1* delims=[] " %%b in ('find /i /n "stop" ^< "%%~a" ^| find /i /v "if inIDE Then Stop" ^|findstr /r /i /c:"^]stop" /c:" stop" /c:"	stop" ^| findstr /r /i /c:"stop$" /c:"stop:" /c:"stop " /c:"stop	"') do (
    set n=%%b
    set "line=%%c"
    call :check || set detected=true
  )
)
::echo Non-screened STOP statements doesn't detected.
if defined detected exit /b 1
goto :eof

:check
  SetLocal EnableDelayedExpansion
  if "!line:~0,1!" neq "'" call :SkipKeyWords "NET.exe STOP" "Critical Stop" || (
    echo.
    echo ^^^!^^^!^^^! Attention ^^^!^^^!^^^! Non-screened STOP statement has been detected.
    echo.
    echo Line ü !n!
    echo Module: %Module%
    echo.
    echo !line!
    pause
    exit /b 1
  )
  EndLocal
exit /b

:SkipKeyWords
  for %%a in (%*) do (
    if "!line!" neq "!line:%%~a=!" exit /b 0
  )
exit /b 1