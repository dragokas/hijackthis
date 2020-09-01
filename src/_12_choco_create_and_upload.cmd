@echo off
SetLocal EnableExtensions

<HiJackThis-update.txt set /p newVersion=

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

set "ch="
set /p "ch=Do you want to upload package to Chocolatey? (Y/N)"
if /i "%ch%" neq "Y" goto :eof

cd "tools\chocolatey-packages"
call _9_upload_git.cmd

for %a in ("cpack.exe") do if "%~$PATH:a"=="" (echo Choco is not installed. Skip.& goto Skip_Choco)
cd "hijackthis/src"
call 3-push-package.bat

:Skip_Choco
pause
