@echo off
SetLocal EnableExtensions

echo.
echo --------------------------------
echo GitHub Uploader by Alex Dragokas
echo --------------------------------
echo.

set /p Ver=< HiJackThis-update.txt

echo New HiJackThis version is: %Ver%
echo.

echo.
set "ch="
set /p "ch=Would you like to write hot-update.txt ? (Y/N)"
if /i "%ch%" neq "n" (
  start "" hot-changelog.txt
  start "" ChangeLog\_TODO_HiJackThis.txt
)

"C:\Program Files\Git\bin\sh.exe" --login -i -- "upload.sh"

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