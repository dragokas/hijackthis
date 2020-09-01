To build HJT project:

1. You need:
 - Windows XP / 7 / 8.1 or 10.
 - all source files in root folder
 - 2 folders: Tools, ico.
 - Visual Basic 6 IDE with SP6 update installed on your Windows machine.

2. Make sure you have the last Service Pack 6 of VB6 IDE: https://www.microsoft.com/en-us/download/details.aspx?id=5721

3. Go to Tools\upx
Place here UPX.exe
You can find it at: https://upx.github.io/

4. Go to Tools\ABR
Download http://dsrt.dyndns.org/files/abr.zip and unpack exe-files to Tools\ABR folder.

5. Go to Tools\7zip.
Place here standalone console version of 7zip.
You can find it at: http://www.7-zip.org/download.html

6. (optional) Digital signature:
Do this step if you want automatically sign EXE after building.
Edit file "_2_Make_UPX_Sign.cmd" and specify path to your signing bat-file at line: set SignScript_1=

7. Run "_0_Open Project Elevated  - !!! - .cmd"
Press Y and ENTER.
This will open the project and update internal version number of libraries.
Close and save the project by request.

8. Run makefile.
Press - (dash) and ENTER.

You'll get:
 - HiJackThis.exe
 - HiJackThis.zip
 - HiJackThis_dbg.exe
 - HiJackThis_dbg.zip
 - _HiJackThis_pass_clean.zip (for AV labs in case of false positives)
 - _HiJackThis_pass_infected.zip (for AV labs in case of false positives)
 - _HiJackThis_pass_virus.zip (for AV labs in case of false positives)
