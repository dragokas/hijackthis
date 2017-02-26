@echo off
SetLocal EnableExtensions EnableDelayedExpansion

:: help info on http://www.vbaccelerator.com/home/VB/Code/Libraries/Resources/Using_RC_EXE/article.asp

echo.
echo :: Creating resource file
echo.

Call :GetOSBitness OSBitness
if "%OSBitness%"=="x32" (set "PF=%ProgramFiles%") else (set "PF=%ProgramFiles(x86)%")
cd /d "%~dp0"

set Res[1]=1 #24 manifest.txt
set Res[2]=101 CUSTOM TasksWhite.csv
set Res[3]=102 CUSTOM MSCOMCTL.OCX.bak
set Res[4]=103 CUSTOM readme - History.txt
set Res[5]=201 CUSTOM _Lang_EN.lng
set Res[6]=202 CUSTOM _Lang_RU.lng

2>NUL del /f /a 1.RC

For /L %%C in (1 1 10) do (
  if defined Res[%%C] (
    for /f "tokens=1-2*" %%a in ("!Res[%%C]!") do (
      set "ID=%%a"
      set "type=%%b"
      set "file=%%c"
      >NUL copy /y "!file!" "!file!.tmp" || (
        echo Error occured during creation resource from: "!file!.tmp"
        pause
      )
      echo !ID! !type! LOADONCALL DISCARDABLE "!file!.tmp">> 1.RC
    )
  )
)

2>nul del /f /a RESOURCE.res

"%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo RESOURCE.res 1.RC && (
    echo.& echo -------   SUCCESS
) || (
    echo Error occured during creation resource from: 1.RC
    pause
)

:: Clear
For /L %%C in (1 1 10) do (
  if defined Res[%%C] (
    for /f "tokens=1-2*" %%a in ("!Res[%%C]!") do (
      >NUL del "%%c.tmp"
    )
  )
)
2>NUL del /f /a 1.RC

exit /b

:GetOSBitness
  set "xOS=x64"& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 set "xOS=x32"
  set "%~1=%xOS%"
Exit /B
