@echo off
SetLocal EnableExtensions
cd /d "%~dp0"

:: File name of project
set "ProjFile=_HijackThis.vbp"

:: Name of the task to create / run
set "TaskName=Run HJT Project"

echo.
echo ---  Loading HiJackthis Fork Project ...
echo.
echo.

Call :GetOSBitness OSBitness
if "%OSBitness%"=="x32" (set "PF=%ProgramFiles%") else (set "PF=%ProgramFiles(x86)%")

if not exist "%PF%\Microsoft Visual Studio\VB98\vb6.exe" (echo VB6 IDE either not installed or located in unknown folder! & pause >NUL & exit /B)

set "compiler=%PF%\Microsoft Visual Studio\VB98\vb6.exe"
set "flags_patch=Tools\FixFlags_AndSign.cmd"

if not exist "%compiler%" (
  echo Warning! Compiler is not found.
  echo.
  pause
)

2>NUL del HiJackThis.pdb

::XP ?
ver |>NUL find " 5." && (start "" "%PF%\Microsoft Visual Studio\VB98\vb6.exe" "%~dp0%ProjFile%" & exit /b)

call :TaskExist

if defined TaskExist (
  call :RunProject NoCheck
) else (
  if "%~1" neq "Admin" (
    call :GetPrivileges
  ) else (
    call :BuildCustomProject "tools\RegTLib\Project1.vbp" "tools\RegTLib\RegTLib.exe"
    if exist "tools\RegTLib\REGTLIB.EXE" (
      tools\RegTLib\REGTLIB.EXE %SystemRoot%\System32\msdatsrc.tlb /admin
      if "%OSBitness%"=="x64" (
        tools\RegTLib\REGTLIB.EXE %SystemRoot%\SysWow64\msdatsrc.tlb /admin
      )
    )
    regsvr32.exe /s MSCOMCTL.OCX
    call :CreateTask
    call :RunProject
  )
)

goto :eof

:CreateTask
  schtasks.exe /create /tn "%TaskName%" /SC ONCE /ST 00:00 /F /RL HIGHEST /tr "\"%PF%\Microsoft Visual Studio\VB98\vb6.exe\" \"%~dp0%ProjFile%\""
exit /b

:CheckErrorHandler
  set "ShouldTweak="
  for /f "tokens=3" %%a in ('reg query "HKCU\Software\Microsoft\VBA\Microsoft Visual Basic" /v "BreakOnAllErrors"') do set BreakOnAllErrors=%%a
  for /f "tokens=3" %%a in ('reg query "HKCU\Software\Microsoft\VBA\Microsoft Visual Basic" /v "BreakOnServerErrors"') do set BreakOnServerErrors=%%a
  
  if "%BreakOnAllErrors%" neq "0x0" set ShouldTweak=true
  if "%BreakOnServerErrors%" neq "0x0" set ShouldTweak=true

  if not defined ShouldTweak exit /b
  
  echo.
  echo Note: it is required that we make small adjustments of VB6 IDE
  echo       in the registry for HiJackThis project to work properly.
  echo.
  echo       Do you allow to change 'Error Trapping' option:
  echo       - to "Break on Unhandled Errors"
  echo.
  set "ch="
  set /p "ch=       ? (Y/N) "
  if /i "%ch%" neq "N" (
    reg add "HKCU\Software\Microsoft\VBA\Microsoft Visual Basic" /v "BreakOnAllErrors" /t REG_DWORD /f /d 0
    reg add "HKCU\Software\Microsoft\VBA\Microsoft Visual Basic" /v "BreakOnServerErrors" /t REG_DWORD /f /d 0
  )
exit /b

:RunProject
  call :CheckErrorHandler
  if "%~1"=="NoCheck" (
    rem if project already run
    schtasks.exe /query /FO LIST /tn "%TaskName%" | findstr /i /C:"Running" /C:"Выполняется" && (
      echo.&echo Project already run !
      pause >NUL
    ) || (
      schtasks.exe /run /tn "%TaskName%" || start "" "%ProjFile%"
    )
  ) else (
    rem Task exists ?
    schtasks.exe /query /FO LIST /tn "%TaskName%" | find /i "%TaskName%" && (
      schtasks.exe /run /tn "%TaskName%"
    ) || (
      start "" "%ProjFile%"
    )
  )
exit /B

:TaskExist
  set "TaskExist="
  schtasks.exe /query /FO LIST | find /i "%TaskName%" && set "TaskExist=1"
exit /B

:GetPrivileges
  net session >NUL 2>NUL || (
    echo.
    echo Administrative privileges required.
    echo.
    mshta "vbscript:CreateObject("Shell.Application").ShellExecute("%~fs0", "Admin", "", "runas", 1) & Close()"
    exit /B 1
  )
exit /B

:GetOSBitness
  :: Определение битности ОС
  set xOS=x64& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 set xOS=x32
  set "%~1=%xOS%"
Exit /B

:BuildCustomProject [prj] [exe]
  set "prj=%~1"
  set "exe=%~2"
  echo.
  for %%a in ("%prj%") do set "fld=%%~dpa"
  <NUL set /p "x=%prj% - "
  if not exist "%exe%" ("%compiler%" /m "%prj%" /outdir "%fld%" && call "%flags_patch%" "%exe%"&& echo OK || echo FAILED !!!) else (echo Exist)
exit /b
