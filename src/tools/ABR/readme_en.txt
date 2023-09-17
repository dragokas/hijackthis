ABR - Autobackup registry v1.10 Copyright(c) 2017-2022 by D.Kuznetsov
Supported OS: Win2k-Win10 x86/x64
Freeware

web     : http://dsrt.dyndns.org:8888
e-mail  : demkd@mail.ru

Syntax abr.exe:
abr.exe [options] backup_path
backup_path - save system registry to the specific folder (default: %SystemRoot%\ABR)
options:
/i - install as a service
/u - uninstall service
/days:n - delete backup folders older than n days (default: 15)

Syntax restore.exe:
(restore_x64.exe for Windows x64)
restore.exe [options] system_drive_letter
system_drive_letter - for offline systems only
options:
/nr - no restart
(!) if you start restore.exe w/o params saved in the current folder registry will be restored for active system.
(!) for inactive systems you must set system_drive_letter parameter (drive letter where target system located)

Use defrag.exe to defragment and repair the saved copy of the registry.
Syntax defrag.exe:
(defrag_x64.exe for Windows x64)
defrag.exe
