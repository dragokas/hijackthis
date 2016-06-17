@echo off
SetLocal EnableExtensions
cd /d "%~dp0"
set "csv=%cd%\TasksWhite.csv"

set "NewCSV=%csv%.tmp"

copy /y "%csv%" "%NewCSV%"

:: align to 4-byte boundary
Tools\Align4byte\Align4byte.exe "%NewCSV%"

:: Encrypt
::Tools\enCryptFile\XOR_Encryptor.exe "%newcsv%" 100

::set "crypted=%NewCSV%.crypt"
set "crypted=%NewCSV%"

:: RC doesn't accept pathes with Russian characters
For /F "delims=" %%a in ("%crypted%") do set "cryptedName=%%~nxa"

Call :GetOSBitness OSBitness
if "%OSBitness%"=="x32" (set "PF=%ProgramFiles%") else (set "PF=%ProgramFiles(x86)%")

echo.
echo :: Creating resource file
:: http://www.vbaccelerator.com/home/VB/Code/Libraries/Resources/Using_RC_EXE/article.asp
del "TasksWhite.RES" 2>NUL
echo 101 CUSTOM LOADONCALL DISCARDABLE "%cryptedName%"> 1.RC
"%PF%\Microsoft Visual Studio\VB98\Wizards\rc.exe" /r /v /fo "TasksWhite.RES" 1.RC

:: Updating resource of project
copy /b /y "manifest_backup.RES" + "TasksWhite.RES" RESOURCE.res && echo -------   SUCCESS   -------

:: Clearing
del "TasksWhite.RES"
del 1.RC
del "%NewCSV%" 2>NUL
del "%crypted%" 2>NUL
ping -n 2 127.1 >NUL
exit /B

:GetOSBitness
  set "xOS=x64"& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 set "xOS=x32"
  set "%~1=%xOS%"
Exit /B
