@REM This is a generated file.
@echo off
setlocal
SET srcdir="C:\tmp\pcre2-10.23"
SET pcre2test="C:\tmp\pcre2-10.23\vstudio\DEBUG\pcre2test.exe"
if not [%CMAKE_CONFIG_TYPE%]==[] SET pcre2test="C:\tmp\pcre2-10.23\vstudio\%CMAKE_CONFIG_TYPE%\pcre2test.exe"
call %srcdir%\RunTest.Bat
if errorlevel 1 exit /b 1
echo RunTest.bat tests successfully completed
