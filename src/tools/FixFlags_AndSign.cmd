@echo off
SetLocal EnableExtensions

:: This batch script is required to fix PE flags in order to support 2k/XP when VB6 compiler uses modern linker.

set "SignScript_1=h:\_AVZ\Наши разработки\_Dragokas\DigiSign\SignME.cmd"

copy /y TSAwarePatch\TSAwarePatch.exe TSAwarePatch.tmp.exe

call :doAction "Align4byte\Align4byte.exe"
call :doAction "ChangeIcon\IC.exe"
call :doAction "RegTLib\RegTLib.exe"
call :doAction "RemoveSign\RemSign.exe"
call :doAction "TSAwarePatch\TSAwarePatch.exe"
call :doAction "VersionPatcher\VersionPatcher.exe"

del TSAwarePatch.tmp.exe
pause
goto :eof

:doAction
  if exist "%~1" (
    echo.
    echo Apply patch to "%~1" ...
    TSAwarePatch.tmp.exe "%~1"
    if exist "%SignScript_1%" call "%SignScript_1%" "%~1" /silent
  )
exit /b
