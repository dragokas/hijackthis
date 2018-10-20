@echo off

rem -------------------------------------------------------------
rem  Build package for chocolatey.
rem -------------------------------------------------------------

SetLocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

rem  Get package name.
cd ..\
for %%a in (".") do set CURRENT_DIR_NAME=%%~na

echo ===== Test (install form local source) "%CURRENT_DIR_NAME%" package ====

if "%~1"=="" (
    call :GetPrivileges
    exit /b
)

set "PACKAGE_NANE=%~1"
cd ./build/%PACKAGE_NANE%

call cinst -fvyd %PACKAGE_NANE% -s . --pre

if not "%1" == "1" (
	pause
)

endlocal
goto :eof

:GetPrivileges
  net session >NUL 2>NUL || (
    echo.
    echo Administrative privileges required.
    echo.
    mshta "vbscript:CreateObject("Shell.Application").ShellExecute("%~fs0", "%CURRENT_DIR_NAME%", "", "runas", 1) & Close()"
    exit /B 1
  )
exit /B