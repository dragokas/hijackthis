To build HJT project:

1. You need all root files and 2 folders: Tools, ico.

2. Make sure you have VB6 SP6 IDE.

3. Go to Tools\Align4byte
Compile it to Align4byte.exe

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

9. Run makefile.
