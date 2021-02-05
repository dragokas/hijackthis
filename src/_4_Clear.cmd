@echo off
SetLocal EnableExtensions
cd /d "%~dp0"

(
del /Q _9_*.txt
del _HiJackThis_pass_infected.zip
del _HiJackThis_pass_virus.zip
del _HiJackThis_pass_clean.zip
rem del HiJackThis.zip
rem del /q HiJackThis*.zip
del hiJackthis.log
del HiJackThis_2.log
del HiJackThis_debug.log
del startuplist.html
del startuplist.txt
del _HiJackThis.csi
del FixReg.log
del *.pdb
del _HJT_src.zip
del /q *.csi
del streams.txt
del "Check Browsers LNK.exe"
del ClearLNK.exe
del /q *.vbw
del DigiSign.log
del processlist.txt
del uninstall_list.txt
del DigiSign.csv
del Errors.log
) 2>NUL