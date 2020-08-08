@echo off
SetLocal EnableExtensions EnableDelayedExpansion

:: help info on:
:: http://www.vbaccelerator.com/home/VB/Code/Libraries/Resources/Using_RC_EXE/article.asp
:: http://www.thevbzone.com/l_res.htm

set "TaskName=Run HJT Project"

:CheckRun
tasklist /v /FI "IMAGENAME eq VB6.exe" 2>NUL|>NUL find /i "HiJackThis" && (
  echo.&echo Project already run !& echo Please, close it first.
  pause >NUL
  goto CheckRun
)

echo.
echo :: Creating resource file
echo.

Call :GetOSBitness OSBitness
if "%OSBitness%"=="x32" (set "PF=%ProgramFiles%") else (set "PF=%ProgramFiles(x86)%")
cd /d "%~dp0"
set ResCnt=0

::Note: type "+" if file requires 4-byte alignment

call :AddResource + 1 #24 manifest.txt
call :AddResource + 101 CUSTOM TasksWhite.csv
call :AddResource - 102 CUSTOM MSCOMCTL.OCX.bak
call :AddResource + 103 CUSTOM _ChangeLog_en.txt
call :AddResource + 104 CUSTOM _ChangeLog_ru.txt
call :AddResource + 201 CUSTOM _Lang_EN.lng
call :AddResource + 202 CUSTOM _Lang_RU.lng
call :AddResource + 203 CUSTOM _Lang_UA.lng
call :AddResource + 204 CUSTOM _Lang_FR.lng
call :AddResource - 301 CUSTOM tools\PCRE2\pcre2-16.dll
call :AddResource - 302 CUSTOM tools\ABR\abr.exe
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
call :AddResource - LNKCHECK     BITMAP ico\main\menu\LnkChecker.bmp
call :AddResource - LNKCLEAN     BITMAP ico\main\menu\LnkCleaner.bmp

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
if "%StrN%" neq "0" more /E +%StrN% < "%~f0" >> 1.RC

:: preparing multilingual VerInfo section
call :Make_VerInfo_res

::2>nul del /f /a RESOURCE.res
"%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo 1.res 1.RC && (
    echo.& echo -------   SUCCESS
) || (
    echo Error occured during creation resource from: 1.RC
    pause
)

copy /b /y 1.res + VerInfo.res RESOURCE.res

:: Clear
For /L %%C in (1 1 %ResCnt%) do (
  if defined Res[%%C] (
    for /f "tokens=1-3*" %%a in ("!Res[%%C]!") do (
      >NUL del "%%d.tmp"
    )
  )
)

2>NUL del /f /a VerInfo.res
2>NUL del /f /a 1.res
2>NUL del /f /a 1.RC

exit /b

:AddResource [4-byte alignment required (+/-)] [ID name] [Resource type] [path]
  set /a ResCnt+=1
  set Res[%ResCnt%]=%*
Exit /b

:GetOSBitness
  set "xOS=x64"& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 set "xOS=x32"
  set "%~1=%xOS%"
Exit /B

:Make_VerInfo_res
  echo Prepairing VerInfo_*.rc
  if not Defined ProjFile For %%a in (*.vbp) do set ProjFile=%%a
  
  :: Reading verion from VBP
  For /F "UseBackQ tokens=1* delims==" %%a in ("%ProjFile%") do (
    if /i "%%a"=="MajorVer" set "Major=%%b"
    if /i "%%a"=="MinorVer" set "Minor=%%b"
    if /i "%%a"=="BuildVer" set "Build=%%b"
    if /i "%%a"=="RevisionVer" set "Revision=%%b"
  )
  if not defined Build set Build=0

  ren tools\ReplaceByRegular\Regular.txt Regular.txt.bak
  copy /y VerInfo_DE.rc VerInfo_DE.rc.bak
  copy /y VerInfo_FR.rc VerInfo_FR.rc.bak
  copy /y VerInfo_RU.rc VerInfo_RU.rc.bak
  copy /y VerInfo_UA.rc VerInfo_UA.rc.bak
  (
    echo word1=1\.2\.3\.4
    echo word2=%Major%.%Minor%.%Build%.%Revision%
    echo word1=1,2,3,4
    echo word2=%Major%,%Minor%,%Build%,%Revision%
  ) > tools\ReplaceByRegular\Regular.txt

  echo RC Version patch
  cscript.exe //nologo tools\ReplaceByRegular\ReplaceByRegular.vbs "%~dp0VerInfo_DE.rc"
  cscript.exe //nologo tools\ReplaceByRegular\ReplaceByRegular.vbs "%~dp0VerInfo_FR.rc"
  cscript.exe //nologo tools\ReplaceByRegular\ReplaceByRegular.vbs "%~dp0VerInfo_RU.rc"
  cscript.exe //nologo tools\ReplaceByRegular\ReplaceByRegular.vbs "%~dp0VerInfo_UA.rc"

  del "tools\ReplaceByRegular\Replace - log.log"

  echo Creating multilingual VerInfo.res
  2>NUL del VerInfo_RU.res
  "%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo VerInfo_RU.res /l 0x419 VerInfo_RU.rc && (
    echo.& echo -------   SUCCESS
  ) || (
    echo Error occured during creation resource from: VerInfo_RU.rc
    pause
  )
  2>NUL del VerInfo_UA.res
  "%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo VerInfo_UA.res /l 0x422 VerInfo_UA.rc && (
    echo.& echo -------   SUCCESS
  ) || (
    echo Error occured during creation resource from: VerInfo_UA.rc
    pause
  )
  2>NUL del VerInfo_FR.res
  "%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo VerInfo_FR.res /l 0x40C VerInfo_FR.rc && (
    echo.& echo -------   SUCCESS
  ) || (
    echo Error occured during creation resource from: VerInfo_FR.rc
    pause
  )
  2>NUL del VerInfo_DE.res
  "%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo VerInfo_DE.res /l 0x407 VerInfo_DE.rc && (
    echo.& echo -------   SUCCESS
  ) || (
    echo Error occured during creation resource from: VerInfo_DE.rc
    pause
  )
  
  :: Special note: you shouldn't add here 'English' ResInfo section (0x409), because it's been build automatically.
  :: Trying to append such .res file cause your project failed to compile.
  :: To do such a job you need to patch .exe file or fill required info in .vbp file.

  copy /b /y VerInfo_RU.res + VerInfo_UA.res + VerInfo_FR.res + VerInfo_DE.res VerInfo.res

  del /f /a VerInfo_*.res
  move /y VerInfo_DE.rc.bak VerInfo_DE.rc
  move /y VerInfo_FR.rc.bak VerInfo_FR.rc
  move /y VerInfo_RU.rc.bak VerInfo_RU.rc
  move /y VerInfo_UA.rc.bak VerInfo_UA.rc
  del tools\ReplaceByRegular\Regular.txt
  ren tools\ReplaceByRegular\Regular.txt.bak Regular.txt
exit /b

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
600,	"Ссылки"
601,	"(Нет)"
602,	"Не можу обрати цю мову!\nСпершу Вам необхідно обрати мову для програм, що не підтримують Юнікод, - Українську\nчерез Панель керування -> Регіональні стандарти."
603,	"Не могу выбрать этот язык!\nСперва Вам необходимо выставить язык для программ, не поддерживающих Юникод, на Русский\nчерез Панель управления -> Региональные стандарты."
604,	"Корпорация Майкрософт"
605,	"Компьютер\"
606,	"Рабочий стол"
607,	"Руководство пользователя"
END 
