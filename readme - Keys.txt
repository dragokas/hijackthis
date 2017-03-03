     Command-line parameters:
     --------------------------------

     1. HiJackThis
     2. StartupList
     3. How to use?

     
     
     1. HiJackThis
     ---------------
     
       /startupscan	- automatically scan the system (the same as button "Do a system scan only")
       /autolog	- automatically scan the system, save a logfile and open it
       /silentautolog	- the same as /autolog, except with no required user intervention
       /StartupList	- run scan by 'StartupList' module
       /ihatewhitelists - ignore all internal whitelists
       /md5		- calculate md5 hash of files
       /uninstall	- remove all HiJackThis Registry entries, backups and quit
       /deleteonreboot "c:\file.sys" - delete the specified file after system rebooting using PendingFileRenameOperations mechanism
       /accepteula 	- accept the agreement. It will not be displayed when program start
       /noGUI		- do not show program window during the scan
       /SysTray	- run program minimized in notification area (system tray)
       /LangEN	- force use English language for user interface
       /LangRU	- force use Russian language for user interface
       /debug		- debug mode (run tracing). Renaming application file to HiJackThis_debug.exe will be the same.
     
     
     2. StartupList
     ----------------
     
     Next keys affects on StartupList module settings only.
     It can be launched manually from the section "Config" -> "Misc Tools" -> "StartupList Scan" 
     or via the key /StartupList
     
       /showempty	- Show empty sections
       /showcmts	- Show comments in .bat files
       /noshowclsids	- Hide class IDs
       /noshowprivate	- Hide usernames and computer name
       /noshowusers	- Hide entries from other users
       /noshowhardware	- Hide entries from other hardware configurations
       /showlargehosts	- Show hosts file even when more than 1000 lines are in it
       /showlargezones	- Show Zones even when more than 1000 domains are in them
       /autosave	- Run hidden, automatically save a report and quit
       /autosavepath:"c:\scan.log" - Specify where to save log, when using /autosave
     
     
     3. How to use command line keys?
     -----------------------------------------
     
     Create text file run.txt near the program HiJackThis.exe. Rename it into extension .bat
     Right mouse click on file run.bat. Choose "Edit".
     
     Example of .bat file code:
     
     cd /d "%~dps0"
     HiJackThis.exe /autolog /ihatewhitelists
     
     Launch file run.bat.
     Alternatively, you can create a shortcut (.lnk) to HiJackThis.exe, open shortcut properties window and append key(s) to the end of line in the box 'Target'.
     