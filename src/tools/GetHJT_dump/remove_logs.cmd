@echo off
SetLocal EnableExtensions

cd /d "%~dp0"

del HiJackThis.exe
del /f /a *.dmp
del /f /a *.log
del Dumps.zip
