
:: Local SVN Subsystem & Project builder by Alex Dragokas

:: This is visual Basic 6 project builder                               [ ver. 1.4 private ]
:: with half-auto managing of subversion

:: Script contains third party software:
:: - Icon Changer by anny05
:: - UPX by Markus Oberhumer, Laszlo Molnar & John Reiser
:: - 7Zip by Igor Pavlov

:: This script doing the following:
:: - check current version & project compilation options
:: - ask for new version to assign (I added ability to change 'build' version; VB6 IDE cannot do it)
:: - compile it using VB6 compiler
:: - add icon with 256x256 size (VB6 IDE don't support it)
:: - add manifest (not used here, cos already integrated)
:: - upx it
:: - add digital signature (external script)
:: - send file for virus checking (using either PhrozenSoft VirusTotal Uploader or Aitotal by Alex1983; separate script)
:: - zip it to storage (folder 'Archive')
:: - all operations (compile / zip / sign / upload) is under checksum/error validation.

@echo off
SetLocal EnableExtensions
cd /d "%~dp0"

:: -------------------- S E T T I N G S -----------------------

:: Choosing the field for version incrementation. Allowed values: Major, Minor, Build or Revision
set IncField=Revision

:: Project's filename, for instance - Test.vbp
:: if the project is single, this fiels can be empty.
set ProjFile=

:: Name of resulting EXE file, for instance - Test.exe
:: This field can be empty
set AppName=

:: Do not use UPX (false - use it)
set NoUPX=false

:: List of file extensions and additional folders in project's directory to include in zip-backup
set arcList=*.vbp *.vbw *.res *.exe *.frm *.frx *.lvw *.cmd *.csi *.csv *.txt *.log *.PDM *.SCC *.lng *.pdb

:: Folder for backup of archive
set ArcFolder=Archive

:: Icon file (for projects without form or projects with non-standart icons). Leave this field empty, if the icon has been already defined in form or if you don't need the icon.
set icoFile=ico\HJT.ico

:: Location of EXE of the program 'Manifest by The Trick'
set ManifestEXE=
::.\ManifestByTheTrick\Manifested.exe

:: Location of manifest file
set Manifest=
::.\ManifestByTheTrick\manifest_asInvoker.txt

:: Location of script(s) for adding digital signature
::set SignScript_1=h:\_AVZ\Наши разработки\_Dragokas\DigiSign\SignME.cmd
::set SignScript_2=d:\Наши проекты\Цифровая подпись\SignME.cmd

:: Version Patcher EXE (support for 'build' field of version)
set VerPatcher=Tools\VersionPatcher\VersionPatcher.exe



:: -------------------------------------------------------------------------------------------

:: Определяем имя файла проекта
if not Defined ProjFile For %%a in (*.vbp) do set ProjFile=%%a

:: Чтение версии из VBP
:: Чтение списка слинкованных файлов
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

:: Если в файле проекта не указано имя EXE для компиляции, берем такое же имя как и у проекта
::if not Defined AppName For %%a in ("%ProjFile%") do set AppName=%%~na.exe
set AppName=HiJackThis.exe

echo.
echo Compilation of -==  %ProjTitle% -^> %AppName%  ==-
echo.
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

:: Привожу введенную версию к стандарту
if "%newVersion%"=="-" (set "newVersion=%OldVersion%"& echo .      New version: %OldVersion%)
if Defined newVersion call :SplitVersionLine "%newVersion%"
if "%Revision%"=="" set Revision=0
if Defined newVersion set newVersion=%Major%.%Minor%.%Build%.%Revision%
:: Если не введена - автоинкремент
if not defined newVersion (
  if /i "%IncField%"=="Major" set /a Major+=1
  if /i "%IncField%"=="Minor" set /a Minor+=1
  if /i "%IncField%"=="Build" set /a Build+=1
  if /i "%IncField%"=="Revision" set /a Revision+=1
)
if not defined newVersion echo .      New version: %Major%.%Minor%.%Build%.%Revision%
:: Обновление строки версии с учетом автоинкремента
if defined newVersion (call :SplitVersionLine "%newVersion%") else (set newVersion=%Major%.%Minor%.%Build%.%Revision%)
call :UpdateProject
echo.

:: updating resources
call "%~dp0_1_Update_Resource.cmd"

:: поиск компилятора
Call :GetOSBitness OSBitness
if "%OSBitness%"=="x32" (set "PF=%ProgramFiles%") else (set "PF=%ProgramFiles(x86)%")

set "arc=%ArcFolder%\%ProjTitle%_%newVersion%"

:: support for "Build" version
call :SplitVersionLine 1.1.1.1
call :UpdateProject
::::::::::::::::::::::::::::::

if exist "%AppName%" del "%AppName%"
"%PF%\Microsoft Visual Studio\VB98\vb6.exe" /m "%ProjFile%" && echo Compilation is successfull. || (
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

:: support for "Build" version
call :SplitVersionLine %newVersion%
call :UpdateProject
"%VerPatcher%" "%cd%\%AppName%" %newVersion%
:::::::::::::::::::::::::::::::

:: for update checker
> "%cd%\HiJackThis-update.txt" set /p "=%newVersion%"<NUL

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

:: Sign
ping -n 2 127.1 >NUL
if exist "%SignScript_1%" call "%SignScript_1%" "%cd%\%AppName%" /silent
if exist "%SignScript_2%" call "%SignScript_2%" "%cd%\%AppName%" /silent

echo.
2>NUL md "%ArcFolder%"
2>NUL del "%ArcFolder%\%ProjTitle%_%newVersion%.zip"
Tools\7zip\7za.exe a -mx9 -y -o"%ArcFolder%" "%ArcFolder%\%ProjTitle%_%newVersion%.zip" %arcList% >NUL 2>&1 && echo Backup is success. || (
Tools\7zip\7za.exe a -mx9 -y -o"%ArcFolder%" "%ArcFolder%\%ProjTitle%_%newVersion%.zip" %arcList%
  echo.
  echo Error has occured during creation of backup !!!
  echo.
  pause
)

For /F "delims=" %%a in ("%AppName%") do set "AppTitle=%%~na"

:: For server uploading

:: Delete old
if exist "%AppTitle%.zip" del /f "%AppTitle%.zip"
:: Pack
Tools\7zip\7za.exe a -mx9 -y -o"%cd%" "%AppTitle%.zip" "%cd%\%AppTitle%.exe"
:: Test
Tools\7zip\7za.exe t "%cd%\%AppTitle%.zip"
:: If there was errors
if %errorlevel% neq 0 (pause & exit /B)

:: For Vir Labs
del %AppTitle%.ex_ 2>NUL
ren "%cd%\%AppTitle%.exe" %AppTitle%.ex_

if exist "_%AppTitle%_pass_infected.zip" del /f "_%AppTitle%_pass_infected.zip"
if exist "_%AppTitle%_pass_virus.zip" del /f "_%AppTitle%_pass_virus.zip"
if exist "_%AppTitle%_pass_clean.zip" del /f "_%AppTitle%_pass_clean.zip"
:: Pack
Tools\7zip\7za.exe a -mx9 -pinfected -y -o"%cd%" "_%AppTitle%_pass_infected.zip" "%cd%\%AppTitle%.ex_"
Tools\7zip\7za.exe a -mx9 -pvirus -y -o"%cd%" "_%AppTitle%_pass_virus.zip" "%cd%\%AppTitle%.ex_"
Tools\7zip\7za.exe a -mx9 -pclean -y -o"%cd%" "_%AppTitle%_pass_clean.zip" "%cd%\%AppTitle%.ex_"
ren "%cd%\%AppTitle%.ex_" %AppTitle%.exe
:: Test
Tools\7zip\7za.exe t -pinfected "%cd%\_%AppTitle%_pass_infected.zip"
:: If there was errors
if %errorlevel% neq 0 (pause & exit /B)
Tools\7zip\7za.exe t -pvirus "%cd%\_%AppTitle%_pass_virus.zip"
:: If there was errors
if %errorlevel% neq 0 (pause & exit /B)
Tools\7zip\7za.exe t -pclean "%cd%\_%AppTitle%_pass_clean.zip"
:: If there was errors
if %errorlevel% neq 0 (pause & exit /B)

ping -n 2 127.1 >NUL

exit /B

:GetOSBitness
  :: Определение битности ОС
  set xOS=x64& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 set xOS=x32
  set "%~1=%xOS%"
Exit /B

:SplitVersionLine %1-Line
  :: Разбить строку версии на компоненты Major, Minor, Revision
  For /F "tokens=1-4 delims=." %%a in ("%~1") do (
    set Major=%%a
    set Minor=%%b
    set Build=%%c
    set Revision=%%d
  )
Exit /B

:UpdateProject
  :: Перезаписать файл проекта VBP
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
      rem Do nothing... delete this line
    ) else (
      if "%%b"=="" (echo %%a) else (echo %%a=%%b)
    )))))
  )) >> "%ProjFile%_"
  move /y "%ProjFile%_" "%ProjFile%" >NUL
Exit /B

:CheckOpenIDE
  :: Проверка, не открыт ли данный проект
  tasklist /FI "IMAGENAME eq VB6.exe" /V /FO CSV /NH |>NUL find /i "%ProjTitle%" && (
    echo I should close IDE window of this project !!!
    rem Отправка сигнала о завершении (без форсирования)
    taskkill /im VB6.exe
    echo.
    pause
    goto CheckOpenIDE
  )
Exit /B

:AddToArcList
  :: Добавить слинкованный файл к списку архивации бекапа
  For /F "tokens=1* delims=; " %%a in ("%~1") do (
    if "%%~b" neq "" (
      set arcList=%arcList% "%%~b"
    ) else (
      set arcList=%arcList% "%%~a"
    )
  )
Exit /B