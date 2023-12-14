@echo off
SetLocal EnableExtensions
cd /d "%~dp0"

(
del /Q _9_*.txt
del _HiJackThis_pass_infected.zip
del _HiJackThis_pass_virus.zip
del _HiJackThis_pass_clean.zip
del _HiJackThis_pass_infected.rar
del HiJackThis.zip
del HiJackThis_dbg.zip
del HiJackThis_dbg_test.zip
del HiJackThis_poly.zip
del HiJackThis_test.zip
rem del HiJackThis.zip
rem del /q HiJackThis*.zip
del hiJackthis.log
del HiJackThis_2.log
del HiJackThis_debug.log
del startuplist.html
del startuplist.txt
del RegKeyType.csv
del RegKeyType.log
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
del _Concat.txt
del _Extracted_lang.txt
del _Func.txt
del _Dll.txt
del Tasks.csv
del HJ-Fixlog*.log
del FixFile.log
) 2>NUL
