To build HiJackThis+ project:

* You need:
 - Windows XP / 7 / 8.1 / 10 or 11.
 - all source files in the root folder and 2 folders in it: "tools" and "ico".
 - Visual Basic 6 IDE with Service Pack 6 update installed.

* Make sure you have the latest Service Pack 6 of VB6 IDE: https://www.microsoft.com/en-us/download/details.aspx?id=5721

* (optional) to use \src\_11_Check_New_Certificates.cmd functionality which is updating built-in certificates list, you must compile:
 - \src\tools\Cert\enumerator\Disallowed\compare-cert-src\Project1.vbp as Compare-cert.exe
 - \src\tools\Cert\enumerator\Disallowed\CertEnumerator-src\Project1.vbp as DisallowedCertEnumerator.exe
 - \src\tools\Cert\enumerator\Root\CertEnum-src\Project1.vbp as CertEnumerator.exe
and launch it if requested by above script.

* Go to \src\Tools\ABR\
Download http://dsrt.dyndns.org/files/abr.zip and unpack exe-files to Tools\ABR folder.
Files must be named as abr.exe, restore.exe and restore_x64.exe

* Go to \src\Tools\7zip\
Place here standalone console version of 7zip as 7za.exe file.
You can find it at: http://www.7-zip.org/download.html

* Go to \src\
Place here MSCOMCTL.OCX Microsoft component file.
It must be dustributed along with Visual Basic 6 installation files.

* Go to \src\apps\
Place here VBCCR17.OCX compiled from \src\tools\VBCCR\ActiveX Control Version\VBCCR17.vbp

* Go to \src\tools\PCRE2\
Place here pcre2-16.dll compiled from this project: https://github.com/tannerhelland/PCRE2-VB6-DLL

* (optional) Digital signature:
Do this step if you want automatically sign EXE after building.
Edit file "\src\_2_Make_UPX_Sign.cmd" and specify path to your signing bat-file at line: set SignScript_1=

* Run "\src\_0_Open Project Elevated  - !!! - .cmd"
Press Y and ENTER.
This will open the project and update internal version number of libraries.
Close and save the project by request.

* Run makefile.cmd
Press - (dash) and ENTER.

You'll get:
 - HiJackThis.exe
 - HiJackThis.zip
 - HiJackThis_dbg.exe
 - HiJackThis_dbg.zip
 - _HiJackThis_pass_clean.zip (for AV labs in case of false positives)
 - _HiJackThis_pass_infected.zip (for AV labs in case of false positives)
 - _HiJackThis_pass_virus.zip (for AV labs in case of false positives)
