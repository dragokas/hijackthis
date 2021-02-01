
:: Project builder by Alex Dragokas

:: This is visual Basic 6 project builder                               [ ver. 1.10 private ]
:: which provide backup system with local version management

:: Script contains third party software:
:: - Icon Changer by anny05
:: - UPX by Markus Oberhumer, Laszlo Molnar & John Reiser
:: - 7Zip by Igor Pavlov

:: This script doing the following:
:: - check whether admin permissions are required (on first run you should confirm UAC in order to register libraries)
:: - check forgotten 'stop' statements in the source code
:: - check for incorrent conditional compilation options
:: - clear unnecessary source files
:: - ask for new version to assign (I added ability to change 'build' version; VB6 IDE cannot do it)
:: - concatenate external files to 1 resource (calling separate script)
:: - compile all dependency projects
:: - compile basic project using VB6 compiler (release and debug versions)
:: - add icon with 256x256 size (VB6 IDE don't support it)
:: - add manifest (not used here, cos already integrated in HJT)
:: - adds TSAware, ASLR and DEP protection
:: - upx it (disabled)
:: - add digital signature (external script) (if exist)
:: - check is PDB (debug) file is correct
:: - zip it to backup storage (folder 'Archive')
:: - create chocolatey package
:: - all operations (compile / zip / sign / upload) is under checksum/error validation.
:: - modify paths and some version of components references in VBP file to support source building on wide range of OS.

@echo off
SetLocal EnableExtensions
cd /d "%~dp0"

:: -------------------- S E T T I N G S -----------------------

:: Choosing the field for version autoincrementation. Allowed values are: Major, Minor, Build or Revision
set IncField=Revision

:: Project's filename, for instance - Test.vbp
:: if the project is single, this field can be empty.
set ProjFile=

:: Name of resulting EXE file, for instance - Test.exe
:: This field can be empty
set AppName=

:: Do not use UPX (false - use it)
set NoUPX=true

:: List of file extensions and additional folders in project's directory to include in zip-backup
set arcList=*.vbp *.vbw *.rc *.res *.exe *.frm *.frx *.lvw *.cmd *.csi *.csv *.txt *.log *.PDM *.SCC *.lng *.pdb *.tlb *.ocx *.dll *.md *.gitignore *.bak *.odl *.idl Tools Ico

:: Folder for backup of archive
set ArcFolder=Archive

:: Icon file (for projects without form or projects with non-standart icons). Leave this field empty, if the icon has been already defined in form or if you don't need the icon.
set icoFile=ico\main\HJT_2.ico

:: Location of EXE of the program 'Manifest by The Trick'
set ManifestEXE=
::.\ManifestByTheTrick\Manifested.exe

:: Location of manifest file
set Manifest=
::.\ManifestByTheTrick\manifest_asInvoker.txt

:: Location of script(s) for adding digital signature
set SignScript_1=h:\_AVZ\Наши разработки\_Dragokas\DigiSign\SignME.cmd
set SignScript_2=c:\DigiSign\SignME.cmd

:: Version Patcher EXE (support for 'build' field of PE EXE version)
set VerPatcher=Tools\VersionPatcher\VersionPatcher.exe

:: Dll characteristics flags patcher => ASLR + DEP + Terminal Services aware
set FlagsPatcher=Tools\TSAwarePatch\TSAwarePatch.exe

:: Flags patch & signing script
set flags_patch=Tools\FixFlags_AndSign.cmd

:: Check debug symbols for matching? (true / false)
set CheckPDB=true
set CheckPDB_tool=tools\ChkMatch\ChkMatch.exe

:: -------------------------------------------------------------------------------------------

if "%~1"=="Fast" set bFast=true

echo [ HiJackThis Fork project builder ]
echo.

:: > XP/2003 ?
ver |>NUL find " 5." || if not exist "Registration_Marker.txt" call :GetPrivileges || exit /b

echo ===^> PLEASE, don't forget to translate new Changelog into English. Thanks :^)
echo.

echo Searching forgotten 'stop' statements ...

:: Searching for non-screened STOP statements
call _5_Check_Stop_Statements.cmd && echo OK.

:: Checking for non-default conditional constant set and forgot
< frmEULA.frm find /i "#Const" | find /i "= True" && (echo.& echo Non-default constant has been detected !!!& echo.& pause)

:: Cleaning logs etc.
call _4_Clear.cmd

:: Searching project's file name
if not Defined ProjFile For %%a in (*.vbp) do set ProjFile=%%a

:: Reading verion from VBP
:: Reading list of linked files
For /F "UseBackQ tokens=1* delims==" %%a in ("%ProjFile%") do (
  if /i "%%a"=="MajorVer" set "Major=%%b"
  if /i "%%a"=="MinorVer" set "Minor=%%b"
  if /i "%%a"=="BuildVer" set "Build=%%b"
  if /i "%%a"=="RevisionVer" set "Revision=%%b"
  if /i "%%a"=="Title" set "ProjTitle=%%~b"
  if /i "%%a"=="Module" call :AddToArcList "%%~b"
  if /i "%%a"=="Class" call :AddToArcList "%%~b"
  if /i "%%a"=="Form" call :AddToArcList "%%~b"
  if /i "%%a"=="ExeName32" if not Defined AppName set AppName=%%~b
)
if not defined MajorVer set MajorVer=0
if not defined MinorVer set MinorVer=0
if not defined Revision set Revision=0
if not defined Build (
  set Build=0
  echo [Private Section]>>"%ProjFile%"
  echo BuildVer>>"%ProjFile%"
)

set OldVersion=%Major%.%Minor%.%Build%.%Revision%

:: If file doesn't contain name of EXE for compilation, we get it the same as project name
::if not Defined AppName For %%a in ("%ProjFile%") do set AppName=%%~na.exe
set AppName=HiJackThis.exe

echo.
echo Compilation of -==  %ProjTitle% -^> %AppName%  ==-
echo.

:: Checking requirements

:: UPX exist ?

if "%NoUPX%"=="false" if not exist "tools\upx\upx.exe" (
  echo.
  echo WARNING:
  echo In order, first, you should download UPX to tools\upx\upx.exe
  echo You can continue, but EXE will be not packed.
  echo.
  pause
)

:: 7zip exist ?

if not exist "tools\7zip\7za.exe" (
  echo.
  echo WARNING:
  echo In order, first, you should download console version of 7zip to tools\7zip\7za.exe
  echo You can continue, but builder will not be able to create a backup!!!
  echo.
  pause
)

:: Is project's window closed ?

call :CheckOpenIDE

echo.
echo.
echo ENTER - to autoincrement %IncField%.
echo - (dash) - to leave old version.
echo.
echo Current version is: %OldVersion%

set newVersion=
set /p newVersion=".      New version: "
::set newVersion=-

:: bringing version line to the standard
if "%newVersion%"=="-" (set "newVersion=%OldVersion%"& echo .      New version: %OldVersion%)
if Defined newVersion call :SplitVersionLine "%newVersion%"
if "%Revision%"=="" set Revision=0
if Defined newVersion set newVersion=%Major%.%Minor%.%Build%.%Revision%
:: if not entered - use autoincrement
if not defined newVersion (
  if /i "%IncField%"=="Major" set /a Major+=1
  if /i "%IncField%"=="Minor" set /a Minor+=1
  if /i "%IncField%"=="Build" set /a Build+=1
  if /i "%IncField%"=="Revision" set /a Revision+=1
)
if not defined newVersion echo .      New version: %Major%.%Minor%.%Build%.%Revision%
:: Updating version line based on autoincrement
if defined newVersion (call :SplitVersionLine "%newVersion%") else (set newVersion=%Major%.%Minor%.%Build%.%Revision%)
call :UpdateProject
echo.

:: searching compiler
Call :GetOSBitness OSBitness
if "%OSBitness%"=="x32" (set "PF=%ProgramFiles%") else (set "PF=%ProgramFiles(x86)%")

set "compiler=%PF%\Microsoft Visual Studio\VB98\vb6.exe"

echo [Building support tools]

call :BuildCustomProject "tools\TSAwarePatch\TSAware_c.vbp" "tools\TSAwarePatch\TSAwarePatch.exe"
call :BuildCustomProject "tools\Align4byte\Project1.vbp" "tools\Align4byte\Align4byte.exe"
call :BuildCustomProject "tools\ChangeIcon\Project1.vbp" "tools\ChangeIcon\IC.exe"
call :BuildCustomProject "tools\RegTLib\Project1.vbp" "tools\RegTLib\RegTLib.exe"
call :BuildCustomProject "tools\RemoveSign\Project1.vbp" "tools\RemoveSign\RemSign.exe"

:: updating resources (it allows prepare and concatenate several resource files: currently for HJT it is a manifest file and whitelists + language files (in future))
call "%~dp0_1_Update_Resource.cmd"

set "arc=%ArcFolder%\%ProjTitle%_%newVersion%"

:: adding support for "Build" version (we should reset it first to default 1.1.1.1)
::call :SplitVersionLine 1.1.1.1
::call :UpdateProject
::::::::::::::::::::::::::::::

for %%a in ("%AppName%") do set "ExeName=%%~na"
:: !!! required for v14.14 linker !!!
del %ExeName%.pdb 2>NUL

echo Starting compilation ...
echo.
echo [Building basic project]
echo.

if exist "%AppName%" del "%AppName%"
"%compiler%" /m "%ProjFile%" /outdir "%~dp0" && echo Compilation is successfull. || (
  echo Could not compile !!!
  echo.
  echo Press ENTER
  echo to roll back version to old: %OldVersion%
  echo and to open project.
  pause >NUL
  call :SplitVersionLine "%OldVersion%"
  call :UpdateProject
  start "" "%ProjFile%"
  Exit /B
)

>NUL copy NUL "Registration_Marker.txt"

:: injecting "Build" field of version
::call :SplitVersionLine %newVersion%
::call :UpdateProject
::"%VerPatcher%" "%cd%\%AppName%" %newVersion%
:::::::::::::::::::::::::::::::

:: TS aware + ASLR + DEP
echo.
echo. Applying ASLR, DEP, TSAWARE, Subsystem version fix ...
echo.
"%FlagsPatcher%" "%cd%\%AppName%" || (echo.& pause)

:: for update checker (in future)
> "%cd%\HiJackThis-update.txt" set /p "=%newVersion%"<NUL
> "%cd%\HiJackThis-update-test.txt" set /p "=%newVersion%"<NUL

:: Adding high-quality icon
if Defined icoFile if exist "%icoFile%" (Tools\ChangeIcon\IC.exe "%cd%\%AppName%" "%icoFile%") else (echo Icon file isn't found !!! & echo. & pause)

if "%NoUPX%"=="true" goto :No_UPX
::Tools\upx\upx.exe --best -f --compress-exports=0 --compress-icons=0 --compress-resources=0 --strip-relocs=0 -o "%AppName%_upx" "%AppName%"
::Tools\upx\upx.exe -9 -f -o "%AppName%_upx" "%AppName%"
Tools\upx\upx.exe --best -f -o "%AppName%_upx" "%AppName%"
if %errorlevel%==0 (
  del "%AppName%"
  ren "%AppName%_upx" "%AppName%"
  ping -n 3 127.1 >NUL
)
:No_UPX

:: Other actions with EXE after compilation

:: manifest
::"%ManifestEXE%" "%cd%\%AppName%" "%Manifest%" -silent

:: Adding digital signature
ping -n 2 127.1 >NUL
if exist "%SignScript_1%" call "%SignScript_1%" "%cd%\%AppName%" /silent
if exist "%SignScript_2%" call "%SignScript_2%" "%cd%\%AppName%" /silent

For /F "delims=" %%a in ("%AppName%") do set "AppTitle=%%~na"

:: Checking debug. symbols for matching the image
if /i "%CheckPDB%"=="true" (
  echo.
  echo Checking debug. symbols ...
  "%CheckPDB_tool%" -c "%cd%\%AppTitle%.exe" "%cd%\%AppTitle%.pdb" | find /i "Result: Matched" || (
    "%CheckPDB_tool%" -c "%cd%\%AppTitle%.exe" "%cd%\%AppTitle%.pdb"
    echo.
    pause
    echo.
  )
)

if defined bFast goto skipBackup

echo.
echo Updating Chocolatey package
echo.
set nuspec=tools\chocolatey-packages\hijackthis\src\hijackthis.nuspec
2>NUL del %nuspec%_2
for /f "tokens=1* delims=[]" %%a in ('^< %nuspec% find /n /v ""') do (
  if "%%a"=="4" (
    echo     ^<version^>%newVersion%^</version^>>>%nuspec%_2
    goto nuspec_exit
  )
  >>%nuspec%_2 echo %%b
)
:nuspec_exit
< %nuspec% more /E +4 >>%nuspec%_2
move /y %nuspec%_2 %nuspec%

for %%a in ("cpack.exe") do if "%%~$PATH:a"=="" (echo Choco is not installed. Skip.& goto Skip_Choco)
call tools\chocolatey-packages\hijackthis\src\1-build-package.bat 1
cd /d "%~dp0"

:Skip_Choco

:: Creating backup
echo.
echo Making backup...
2>NUL md "%ArcFolder%"
2>NUL del "%ArcFolder%\%ProjTitle%_%newVersion%.zip"
Tools\7zip\7za.exe a -mx9 -y -o"%ArcFolder%" "%ArcFolder%\%ProjTitle%_%newVersion%.zip" %arcList% >NUL 2>&1 && echo Backup is success. || (
Tools\7zip\7za.exe a -mx9 -y -o"%ArcFolder%" "%ArcFolder%\%ProjTitle%_%newVersion%.zip" %arcList%
  echo.
  echo Error has occured during creation of backup !!!
  echo.
  pause
)

:: For server uploading

:: Delete old
if exist "%AppTitle%.zip" del /f "%AppTitle%.zip"
:: Pack
Tools\7zip\7za.exe a -mx9 -y -o"%cd%" "%AppTitle%.zip" "%cd%\%AppTitle%.exe"
:: Test
Tools\7zip\7za.exe t "%cd%\%AppTitle%.zip"
:: If there was errors
if %errorlevel% neq 0 (pause & exit /B)

:: For debug purposes
copy /y "%cd%\%AppTitle%.exe" "%AppTitle%_dbg.exe"
Tools\7zip\7za.exe a -mx9 -y -o"%cd%" "%AppTitle%_dbg.zip" "%AppTitle%_dbg.exe"
copy /y "%AppTitle%_dbg.zip" "%AppTitle%_dbg_test.zip"

:: Pseudo-polymorph
copy /y "%cd%\%AppTitle%.exe" "HJT_poly.pif"
:: remove signature
Tools\RemoveSign\RemSign.exe "%cd%\HJT_poly.pif"
:: pack
Tools\7zip\7za.exe a -mx9 -y -o"%cd%" "%AppTitle%_poly.zip" "HJT_poly.pif"

:: For Vir Labs
copy /y "%cd%\%AppTitle%.exe" %AppTitle%.ex_

if exist "_%AppTitle%_pass_infected.zip" del /f "_%AppTitle%_pass_infected.zip"
if exist "_%AppTitle%_pass_virus.zip" del /f "_%AppTitle%_pass_virus.zip"
if exist "_%AppTitle%_pass_clean.zip" del /f "_%AppTitle%_pass_clean.zip"
:: Pack
Tools\7zip\7za.exe a -mx1 -pinfected -y -o"%cd%" "_%AppTitle%_pass_infected.zip" "%cd%\%AppTitle%.ex_"
Tools\7zip\7za.exe a -mx1 -pvirus -y -o"%cd%" "_%AppTitle%_pass_virus.zip" "%cd%\%AppTitle%.ex_"
Tools\7zip\7za.exe a -mx1 -pclean -y -o"%cd%" "_%AppTitle%_pass_clean.zip" "%cd%\%AppTitle%.ex_"
del "%cd%\%AppTitle%.ex_"
:: Test
Tools\7zip\7za.exe t -pinfected "%cd%\_%AppTitle%_pass_infected.zip"
if %errorlevel% neq 0 (pause & exit /B)
Tools\7zip\7za.exe t -pvirus "%cd%\_%AppTitle%_pass_virus.zip"
if %errorlevel% neq 0 (pause & exit /B)
Tools\7zip\7za.exe t -pclean "%cd%\_%AppTitle%_pass_clean.zip"
if %errorlevel% neq 0 (pause & exit /B)

:skipBackup

::copy /y MSCOMCTL.OCX.bak MSCOMCTL.OCX
copy /y HiJackThis.zip HiJackThis_test.zip
del /f /a /q *.tmp 2>NUL

if defined bFast goto skipVT

echo.
echo "%cd%\%AppName%" | clip
echo Path to HiJackThis.exe has been copied to clipboard (for VT check).
echo.

:skipVT
if defined bFast goto skipAskHotUpdate

echo.
set "ch="
set /p "ch=Would you like to write hot-update.txt ? (Y/N)"
if /i "%ch%" neq "n" (
  start "" hot-changelog.txt
  start "" _ChangeLog_en.txt
)

:skipAskHotUpdate

:: test running HJT scan from Autologger (2 logs should be created - HiJackThis.log and HiJackThis_debug.log)
if not defined bFast call _10_Scan_Execution_Test.cmd Ask

exit /B

:BuildCustomProject [prj] [exe]
  set "prj=%~1"
  set "exe=%~2"
  echo.
  for %%a in ("%prj%") do set "fld=%%~dpa"
  <NUL set /p "x=%prj% - "
  if not exist "%exe%" ("%compiler%" /m "%prj%" /outdir "%fld%" && call "%flags_patch%" "%exe%"&& echo OK || echo FAILED !!!) else (echo Exist)
exit /b

:GetOSBitness
  :: Determination of OS bitness
  set xOS=x64& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 set "xOS=x32"
  set "%~1=%xOS%"
Exit /B

:SplitVersionLine %1-Line
  :: Split version string into components Major, Minor, Build, Revision
  For /F "tokens=1-4 delims=." %%a in ("%~1") do (
    set Major=%%a
    set Minor=%%b
    set Build=%%c
    set Revision=%%d
  )
Exit /B

:UpdateProject
  :: Modify and rewrite VBP project file
  2>NUL del "%ProjFile%_"
  (
  For /F "UseBackQ tokens=1* delims==" %%a in ("%ProjFile%") do (
    if /i "%%a"=="MajorVer" (
      echo MajorVer=%Major%
    ) else (
    if /i "%%a"=="MinorVer" (
      echo MinorVer=%Minor%
    ) else (
    if /i "%%a"=="RevisionVer" (
      echo RevisionVer=%Revision%
    ) else (
    if /i "%%a"=="BuildVer" (
      echo BuildVer=%Build%
    ) else (
    if /i "%%a"=="Path32" (
      rem Do nothing... delete this line (because it can cause unexpected conflicts)
    ) else (
      if "%%b"=="" (echo %%a) else (echo %%a=%%b)
    )))))
  )) >> "%ProjFile%_"
  move /y "%ProjFile%_" "%ProjFile%" >NUL
  call :Normalize_VBP_References
  call :RemoveMSComctlVer
Exit /B

:CheckOpenIDE
  :: Checking, whether the project is not opened
  tasklist /FI "IMAGENAME eq VB6.exe" /V /FO CSV /NH 2>NUL |>NUL find /i "HiJackThis" && (
    echo I should close IDE window of this project !!!
    rem Sending request signal for closing (without force)
    taskkill /im VB6.exe
    echo.
    pause
    goto CheckOpenIDE
  )
Exit /B

:AddToArcList
  :: Add linked file to the list of backup files for next archivation
  For /F "tokens=1* delims=; " %%a in ("%~1") do (
    if "%%~b" neq "" (
      set arcList=%arcList% "%%~b"
    ) else (
      set arcList=%arcList% "%%~a"
    )
  )
Exit /B

:Normalize_VBP_References
  :: remove relative path
  :: substitute correct System path based on current OS bitness

  :: Examples:
  :: 1.
  :: Reference=*\G{C88FCAC2-DE90-11D3-9876-8517F6B99C68}#1.6#0#..\..\..\CHECKB~1\olelib2.tlb#Edanmo's OLE interfaces for Implements v1.51
  :: will be replaced by:
  :: Reference=*\G{C88FCAC2-DE90-11D3-9876-8517F6B99C68}#1.6#0#olelib2.tlb#Edanmo's OLE interfaces for Implements v1.51
  :: 2.
  :: Reference=*\G{E34CB9F1-C7F7-424C-BE29-027DCC09363A}#1.0#0#C:\Windows\SysWOW64\taskschd.dll#TaskScheduler
  :: will be replaced by:
  :: Reference=*\G{E34CB9F1-C7F7-424C-BE29-027DCC09363A}#1.0#0#C:\Windows\System32\taskschd.dll#TaskScheduler
  :: (if OS bitness is x32)

  :: Split results:
  :: token 1 (a): Reference=*\G{C88FCAC2-DE90-11D3-9876-8517F6B99C68}
  :: token 2 (b): 1.6
  :: token 3 (c): 0
  :: token 4 (d): ..\..\..\CHECKB~1\olelib2.tlb
  :: token 5 (e): Edanmo's OLE interfaces for Implements v1.51

  :: Modify and rewrite VBP project file
  2>NUL del "%ProjFile%_"
  (
  echo Type=Exe
  For /F "tokens=1,2,3,4,5 delims=#" %%a in ('^< "%ProjFile%" findstr /IRC:"^Reference="') do (
    call :IsBeginWith "%%~d" ".." && (
      rem remove relative path
      echo %%a#%%b#%%c#%%~nxd#%%e
    ) || (
    call :IsBeginWith "%%~d" "%SystemRoot%\System32" "%SystemRoot%\SysWOW64" && (
      rem substitute correct bitness
      rem if "%OSBitness%"=="x32" (
      rem  echo %%a#%%b#%%c#%SystemRoot%\System32\%%~nxd#%%e
      rem ) else (
      rem  echo %%a#%%b#%%c#%SystemRoot%\SysWOW64\%%~nxd#%%e
      rem )
      echo %%a#%%b#%%c#%%~nxd#%%e
    ) || (
      echo %%a#%%b#%%c#%%d#%%e
    ))
  )) >> "%ProjFile%_"
  :: skip 1-st line (Type=Exe) and References lines
  < "%ProjFile%" more /E +1 | findstr /IVRC:"^Reference=" >> "%ProjFile%_"
  
  move /y "%ProjFile%_" "%ProjFile%" >NUL
  
exit /B

:RemoveMSComctlVer
  2>NUL del "%ProjFile%_"
  For /F "UseBackQ delims=" %%a in ("%ProjFile%") do (
    For /F "tokens=1-2 delims==#" %%b in ("%%a") do (
      if "%%c"=="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" (
        echo Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}##0; MSCOMCTL.OCX
      ) else (
        echo %%a
      )
    )
  ) >> "%ProjFile%_"
  move /y "%ProjFile%_" "%ProjFile%" >NUL
Exit /b

:IsBeginWith [in_source] [paramarray_search term]
  :: return code: 0 - success, 1 - failure.
  set "s_src=%~1"
  call :len_of_var "%~2" "len"
  setlocal EnableDelayedExpansion
  if /i "!s_src:~,%len%!"=="%~2" (endlocal & exit /B 0)
  endlocal
if "%~3"=="" (exit /B 1) else (shift /2 & goto IsBeginWith)

:len_of_var [in_Text] [out_Len.of.Text]
  set "_var=%~1"& set "_count=0"
  :_count_loop
  set "_var=%_var:~1%"
  set /a _count+=1
  if not defined _var (set "%~2=%_count%"& exit /b) else (goto _count_loop)
Exit /B

:GetPrivileges
  net session >NUL 2>NUL || (
    echo.
    echo Administrative privileges required.
    echo.
    mshta "vbscript:CreateObject("Shell.Application").ShellExecute("%~fs0", "Admin", "", "runas", 1) & Close()"
    exit /B 1
  )
exit /B
