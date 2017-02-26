@echo off
SetLocal enableExtensions

(
  For %%a in (*.bas *.frm *.cls) do findstr /i /c:"translate(" < "%%~a"
) > _Extracted_lang.txt

pause