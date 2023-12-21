@echo off
SetLocal EnableExtensions
cd /d "%~DP0"

::call :compile TrickVB6LauncherTypeLibrary.idl
::call :compile IRegexp.odl
::call :compile MLang.Idl

pause

goto :eof

:compile
  :: DON'T CHANGE ANYTHING
  :: midl doesn't like spaces, RU characters in path etc
  set "out=%TEMP%"
  echo.---
  echo Compiling %~1
  echo.---
  call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars32.bat"
  ::del "%~n1.tlb"
  midl.exe /win32 /mktyplib203 "%~1" /out "%out%"
  move /y "%out%\%~n1.tlb" "%~dp0"
exit /b
