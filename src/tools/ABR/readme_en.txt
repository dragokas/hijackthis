ABR - Autobackup registry v1.20 Copyright(c) 2017-2024 by D.Kuznetsov
Supported OS: Win2k-Win11 x86/x64
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
restore.exe [options] windows_drive_letter
windows_drive_letter - for inactive systems, including when booting into the command line
options:
/nr - no restart
(!) if you start restore.exe w/o params saved in the current folder registry will be restored for active system.
(!) for inactive systems you must set system_drive_letter parameter (drive letter where target system located)

Use defrag.exe to defragment and repair the saved copy of the registry.
Syntax defrag.exe: defrag.exe
(use defrag_x64.exe for Windows x64)

Useful utilities:
bootrec /rebuildbcd (creating a new BCD, the old BCD must first be deleted)
shutdown /r /o /t 0 (reboot to boot menu mode, from this menu you can boot into the command line)