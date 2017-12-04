@echo off
SetLocal EnableExtensions EnableDelayedExpansion

:: help info on:
:: http://www.vbaccelerator.com/home/VB/Code/Libraries/Resources/Using_RC_EXE/article.asp
:: http://www.thevbzone.com/l_res.htm

set "TaskName=Run HJT Project"

schtasks.exe /query /FO LIST /tn "%TaskName%" | findstr /i /C:"Running" /C:"Выполняется" && (
  echo.&echo Project already run !& echo Please, close it first.
  pause >NUL
  exit /b
)

echo.
echo :: Creating resource file
echo.

Call :GetOSBitness OSBitness
if "%OSBitness%"=="x32" (set "PF=%ProgramFiles%") else (set "PF=%ProgramFiles(x86)%")
cd /d "%~dp0"
set ResCnt=0

::Note: type + if file requires 4-byte alignment

call :AddResource + 1 #24 manifest.txt
call :AddResource + 101 CUSTOM TasksWhite.csv
call :AddResource - 102 CUSTOM MSCOMCTL.OCX.bak
call :AddResource + 103 CUSTOM readme - History.txt
call :AddResource + 201 CUSTOM _Lang_EN.lng
call :AddResource + 202 CUSTOM _Lang_RU.lng
call :AddResource - 301 CUSTOM tools\PCRE2\pcre2-16.dll
call :AddResource - 302 CUSTOM tools\ABR\abr.exe
call :AddResource - 303 CUSTOM tools\ABR\restore.exe
call :AddResource - 304 CUSTOM tools\ABR\restore_x64.exe
call :AddResource - ADSSPY       BITMAP ico\main\menu\ADSSpy.bmp
call :AddResource - CROSS_RED    BITMAP ico\main\menu\cross.bmp
call :AddResource - CROSS_BLACK  BITMAP ico\main\menu\Remove.bmp
call :AddResource - GLOBE        BITMAP ico\main\menu\globe.bmp
call :AddResource - HOSTS        BITMAP ico\main\menu\hosts.bmp
call :AddResource - IE           BITMAP ico\main\menu\ie.bmp
call :AddResource - KEY          BITMAP ico\main\menu\Key.bmp
call :AddResource - PROCMAN      BITMAP ico\main\menu\ProcMan.bmp
call :AddResource - SETTINGS     BITMAP ico\main\menu\settings2.bmp
call :AddResource - SIGNATURE    BITMAP ico\main\menu\signature.bmp
call :AddResource - STARTUPLIST  BITMAP ico\main\menu\StartupList.bmp
call :AddResource - UNINSTALLER  BITMAP ico\main\menu\Uninstaller.bmp
call :AddResource - INSTALL      BITMAP ico\main\menu\install.bmp
call :AddResource - UPDATE       BITMAP ico\main\menu\update.bmp

2>NUL del /f /a 1.RC

For /L %%C in (1 1 %ResCnt%) do (
  if defined Res[%%C] (
    for /f "tokens=1-3*" %%a in ("!Res[%%C]!") do (
      set "Align=%%a"
      set "ID=%%b"
      set "type=%%c"
      set "file=%%~d"
      if !Align!==+ Tools\Align4byte\Align4byte.exe "!file!"
      set "file=!file:\=\\!"
      >NUL copy /y "!file!" "!file!.tmp" || (
        echo Error occured during creation resource from: "!file!.tmp"
        pause
      )
      echo !ID! !type! LOADONCALL DISCARDABLE "!file!.tmp">> 1.RC
    )
  )
)

:: Adding string table
set StrN=0
set Label=STRINGS
for /f "delims=[]" %%a in ('^< "%~f0" find /n ":%Label%"') do set StrN=%%a
if "%StrN%" neq "0" more +%StrN% < "%~f0" >> 1.RC

2>nul del /f /a RESOURCE.res

"%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo RESOURCE.res 1.RC && (
    echo.& echo -------   SUCCESS
) || (
    echo Error occured during creation resource from: 1.RC
    pause
)

:: Clear
For /L %%C in (1 1 %ResCnt%) do (
  if defined Res[%%C] (
    for /f "tokens=1-3*" %%a in ("!Res[%%C]!") do (
      >NUL del "%%d.tmp"
    )
  )
)
2>NUL del /f /a 1.RC

exit /b

:AddResource
  set /a ResCnt+=1
  set Res[%ResCnt%]=%*
Exit /b

:GetOSBitness
  set "xOS=x64"& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 set "xOS=x32"
  set "%~1=%xOS%"
Exit /B

:STRINGS

STRINGTABLE
BEGIN
100, 	"Error loading constants"
101, 	"Error loading project"
102, 	"Error copying file - "
103, 	"Win32 error."
104, 	"Error execute line - "
105, 	"Error running the executable file"
200, 	"PROJECT"
300, 	"kernel32"
301, 	"VirtualAlloc"
302, 	"VirtualProtect"
303, 	"VirtualFree"
304, 	"RtlMoveMemory"
305, 	"RtlFillMemory"
306, 	"lstrcpynW"
307, 	"LoadLibraryA"
308, 	"GetProcAddress"
309, 	"ExitProcess"
310, 	"HeapAlloc"
311, 	"HeapFree"
312, 	"GetProcessHeap"
313, 	"GetCurrentProcess"
350, 	"ntdll"
351, 	"NtQueryInformationProcess"
400, 	"user32"
401, 	"MessageBoxW"
500, 	"Success"
501, 	"Unable to get NT headers from EXE file"
502, 	"Invalid data directory"
503, 	"Unable to allocate memory"
504, 	"Unable to protect memory"
505, 	"LoadLibrary failed"
506, 	"Process information not found"
END 
