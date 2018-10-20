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
  start "" ChangeLog\ChangeLog.txt
)

"C:\Program Files\Git\bin\sh.exe" --login -i -- "upload.sh"
pause