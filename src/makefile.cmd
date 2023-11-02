@echo off
setlocal EnableExtensions
call "%~dp0_0_Open Project Elevated  - !!! - .cmd"
echo.
echo Please close IDE... then, press ENTER.
pause >NUL
@call "%~dp0_2_Make_UPX_Sign.cmd"