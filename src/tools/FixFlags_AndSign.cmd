@echo off
SetLocal EnableExtensions

:: This batch script is required to fix PE flags in order to support 2k/XP when VB6 compiler uses modern linker.

:: set the path to your signing script here
set "SignScript_1=h:\_AVZ\Наши разработки\_Dragokas\DigiSign\SignME.cmd"

copy /y "%~dp0TSAwarePatch\TSAwarePatch.exe" "%~dp0TSAwarePatch.tmp.exe" >NUL

if "%~1"=="" (
  call :doAction "Align4byte\Align4byte.exe"
  call :doAction "ChangeIcon\IC.exe"
  call :doAction "RegTLib\RegTLib.exe"
  call :doAction "RemoveSign\RemSign.exe"
  call :doAction "TSAwarePatch\TSAwarePatch.exe"
  call :doAction "VersionPatcher\VersionPatcher.exe"
) else (
  call :doAction "%~1"
)

del "%~dp0TSAwarePatch.tmp.exe"
if "%~1"=="" pause
goto :eof

:doAction
  if exist "%~1" (
    echo.
    echo Apply patch to "%~1" ...
    "%~dp0TSAwarePatch.tmp.exe" "%~1"
    if exist "%SignScript_1%" call "%SignScript_1%" "%~1" /silent
  )
exit /b
