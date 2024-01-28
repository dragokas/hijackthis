@echo off
SetLocal EnableExtensions EnableDelayedExpansion

:: help info on:
:: http://www.vbaccelerator.com/home/VB/Code/Libraries/Resources/Using_RC_EXE/article.asp
:: http://www.thevbzone.com/l_res.htm

if "%~1"=="" (
  cmd /c "%~f0" 1
  goto :eof
)

:: To encrypt text resources
set ResEncoder=tools\Caesar_Encoder\ResEncoder\ResEncoder.exe

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

call :AddResource a 1 #24 manifest.txt
call :AddResource ae 101 CUSTOM database\TasksWhite.csv
::call :AddResource - 102 CUSTOM MSCOMCTL.OCX.res
call :AddResource ae 103 CUSTOM _ChangeLog_en.txt
call :AddResource ae 104 CUSTOM _ChangeLog_ru.txt
call :AddResource ae 105 CUSTOM database\hosts_xp
call :AddResource ae 106 CUSTOM database\hosts_vista
call :AddResource ae 107 CUSTOM database\hosts_7-11
call :AddResource ae 108 CUSTOM database\R_Section.txt
call :AddResource ae 109 CUSTOM database\SafeDomains.txt
call :AddResource ae 110 CUSTOM database\SafeProtocols.txt
call :AddResource ae 111 CUSTOM database\SafeFilters.txt
call :AddResource ae 112 CUSTOM database\SafeDNS.txt
call :AddResource ae 113 CUSTOM database\DisallowedCert.txt
call :AddResource ae 114 CUSTOM database\LoLBin.txt
call :AddResource ae 115 CUSTOM database\ServicePath.txt
call :AddResource ae 116 CUSTOM database\ServiceFilename.txt
call :AddResource ae 117 CUSTOM database\DriverMapped.txt
call :AddResource ae 118 CUSTOM database\LoLBin_Protect.txt
call :AddResource ae 119 CUSTOM database\CriticalRm.txt
call :AddResource ae 201 CUSTOM _Lang_EN.lng
call :AddResource ae 202 CUSTOM _Lang_RU.lng
call :AddResource ae 203 CUSTOM _Lang_UA.lng
call :AddResource ae 204 CUSTOM _Lang_FR.lng
call :AddResource ae 205 CUSTOM _Lang_SP.lng
call :AddResource e 301 CUSTOM tools\PCRE2\pcre2-16.dll
call :AddResource e 302 CUSTOM tools\ABR\abr.exe
call :AddResource e 303 CUSTOM tools\ABR\restore.exe
call :AddResource e 304 CUSTOM tools\ABR\restore_x64.exe
call :AddResource - BUTTON CUSTOM ico\Themes\Button_1.bmp
call :AddResource - ADSSPY       BITMAP ico\main\menu\ADSSpy.bmp
call :AddResource - CROSS_RED    BITMAP ico\main\menu\cross.bmp
call :AddResource - CROSS_BLACK  BITMAP ico\main\menu\Remove.bmp
call :AddResource - GLOBE        BITMAP ico\main\menu\globe.bmp
call :AddResource - HOSTS        BITMAP ico\main\menu\hosts.bmp
call :AddResource - IE           BITMAP ico\main\menu\ie.bmp
call :AddResource - KEY          BITMAP ico\main\menu\Key.bmp
call :AddResource - REGTYPE      BITMAP ico\main\menu\RegType.bmp
call :AddResource - PROCMAN      BITMAP ico\main\menu\ProcMan.bmp
call :AddResource - SETTINGS     BITMAP ico\main\menu\settings2.bmp
call :AddResource - SIGNATURE    BITMAP ico\main\menu\signature.bmp
call :AddResource - STARTUPLIST  BITMAP ico\main\menu\StartupList.bmp
call :AddResource - UNINSTALLER  BITMAP ico\main\menu\Uninstaller.bmp
call :AddResource - INSTALL      BITMAP ico\main\menu\install.bmp
call :AddResource - UPDATE       BITMAP ico\main\menu\update.bmp
call :AddResource - LNKCHECK     BITMAP ico\main\menu\LnkChecker.bmp
call :AddResource - LNKCLEAN     BITMAP ico\main\menu\LnkCleaner.bmp

:: Prepare "apps" folder
copy /y tools\ABR\abr.exe apps\

2>NUL del /f /a 1.RC

For /L %%C in (1 1 %ResCnt%) do (
  if defined Res[%%C] (
    for /f "tokens=1-3*" %%a in ("!Res[%%C]!") do (
	  set Align=
	  set Encrypt=
	  set "Flags=%%a"
	  if "!Flags!"=="a" set Align=+
	  if "!Flags!"=="e" set Encrypt=+
	  if "!Flags!"=="ae" (
	    set Align=+
		set Encrypt=+
	  )
      set "ID=%%b"
      set "type=%%c"
      set "file=%%~d"
      >NUL copy /y "!file!" "!file!.tmp" || (
        echo Error occured during creation resource from: "!file!.tmp"
        pause
      )
	  rem for /f "delims=" %%Z in ("!file!") do echo [RES] Adding resource !Res[%%C]! ^(size=%%~zZ^)
	  
	  if !Align!==+ Tools\Align4byte\Align4byte.exe "!file!.tmp"
	  if !Encrypt!==+ call :encrypt_res binary "!file!.tmp"
	  set "file=!file:\=\\!"
      echo !ID! !type! LOADONCALL DISCARDABLE "!file!.tmp">> 1.RC
    )
  )
)

:PrepStrTable
:: Adding string table
call :GetFileSha1 "tools\PCRE2\pcre2-16.dll" ShaPCRE
call :GetFileSha1 "apps\abr.exe" ShaABR
call :GetFileSha1 "apps\VBCCR17.OCX" ShaOCX
set StrN=0
set Label=STRINGS
for /f "delims=[]" %%a in ('^< "%~f0" find /n ":%Label%"') do set StrN=%%a
if "%StrN%" neq "0" more /E +%StrN% < "%~f0" >> 1.RC
echo 700, 	"%ShaPCRE%">> 1.RC
echo 701, 	"%ShaABR%">> 1.RC
echo 702, 	"%ShaOCX%">> 1.RC
echo END>>1.RC

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

::@Flags
::"a" - if the file requires 4-byte alignment
::"e" - if the file requires encryption
::"ae" - if the file requires 4-byte alignment and encryption
:AddResource [Flags] [ID name] [Resource type] [path] 
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
  copy /y VerInfo_SP.rc VerInfo_SP.rc.bak
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
  cscript.exe //nologo tools\ReplaceByRegular\ReplaceByRegular.vbs "%~dp0VerInfo_SP.rc"

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
  2>NUL del VerInfo_SP.res
  "%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo VerInfo_SP.res /l 0xC0A VerInfo_SP.rc && (
    echo.& echo -------   SUCCESS
  ) || (
    echo Error occured during creation resource from: VerInfo_SP.rc
    pause
  )
  
  :: Special note: you shouldn't add here 'English' ResInfo section (0x409), because it's been build automatically.
  :: Trying to append such .res file cause your project failed to compile.
  :: To do such a job you need to patch .exe file or fill required info in .vbp file.

  copy /b /y VerInfo_RU.res + VerInfo_UA.res + VerInfo_FR.res + VerInfo_DE.res + VerInfo_SP.res VerInfo.res

  del /f /a VerInfo_*.res
  move /y VerInfo_DE.rc.bak VerInfo_DE.rc
  move /y VerInfo_FR.rc.bak VerInfo_FR.rc
  move /y VerInfo_RU.rc.bak VerInfo_RU.rc
  move /y VerInfo_UA.rc.bak VerInfo_UA.rc
  move /y VerInfo_SP.rc.bak VerInfo_SP.rc
  del tools\ReplaceByRegular\Regular.txt
  ren tools\ReplaceByRegular\Regular.txt.bak Regular.txt
exit /b

:GetFileSha1 [file] [out_var]
  for /f "delims=" %%a in ('certutil -hashfile "%~1" SHA1 ^| find /v "hash"') do set "t=%%a"
  set "t=%t: =%"
  set "%~2=%t%"
  if "%t%"=="" (echo Error in calculation Sha1 of file "%~1" & pause)
exit /b

:encrypt_res [text/binary] [file]
  "%ResEncoder%" encrypt %~1 "%~2" "%~2.enc"
  if errorlevel 1 (
    echo Error in encryption resource file: "%~2"
    pause
  )
  ::verify can we get decrypted file to match with original hash
  "%ResEncoder%" decrypt %~1 "%~2.enc" "%~2.dec" >nul
  call :GetFileSha1 "%~2" HashOrig
  call :GetFileSha1 "%~2.dec" HashDecrypted
  if "%HashOrig%" neq "%HashDecrypted%" (
    echo Mismatched hash in attempt to encrypt resource file: "%~2"
	pause
  )
  del "%~2.dec"
  move /y "%~2.enc" "%~2" >NUL
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
601,	"(Нет)"
604,	"Корпорация Майкрософт"
605,	"Компьютер\"
606,	"Рабочий стол"
607,	"Руководство пользователя"
608,	"Руководство пользователя (дополнение)"
//700 PCRE2 Sha1
//701 ABR Sha1
//See details at :PrepStrTable label