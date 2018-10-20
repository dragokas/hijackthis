To build HJT project:

1. You need:
 - all files in root folder
 - 2 folders: Tools, ico.
 - VB6 SP6 IDE.

2. Make sure you have the last Service Pack 6 of VB6 IDE: https://www.microsoft.com/en-us/download/details.aspx?id=5721

3. Go to Tools\Align4byte
Open project (.vbp file)
Compile it to Align4byte.exe (File -> Compile...)

4. Go to Tools\ChangeIcon
Compile it to IC.exe

5. Go to Tools\VersionPatcher
Compile it to VersionPatcher.exe

6. Go to Tools\upx
Place here UPX.exe
You can find it at: https://upx.github.io/

7. Go to Tools\7zip.
Place here console version of 7zip. Rename it into 7za.exe
You can find it at: http://www.7-zip.org/download.html

8. Run _0_Open Project Elevated  - !!! - .cmd
to make sure your OS created all library references on developer machine.

9. (optional) Digital signature:
Do this step if you want automatically sign EXE after building.
Edit file "_2_Make_&_UPX_&_Sign.cmd" and specify path to your signing bat-file at line: set SignScript_1=

10. Run makefile.
Press - (dash) and ENTER.

You'll get:
 - HiJackThis.exe
 - HiJackThis.zip
 - _HiJackThis_pass_clean.zip (for AV labs in case of false positives)
 - _HiJackThis_pass_infected.zip (for AV labs in case of false positives)
 - _HiJackThis_pass_virus.zip (for AV labs in case of false positives)

