* Trend Micro HiJackThis v2.0.6 Fork by Alex Dragokas ver. 1.0 *


1. Description
--------------

Trend Micro HiJackThis is a free utility that generates an in depth report of registry and file settings from your computer. 
HiJackThis makes no separation between safe and unsafe settings in its scan results giving you the ability to selectively remove items from your machine. 
In addition to this scan and remove capability HiJackThis comes with several tools useful in manually removing malware from a computer.

IMPORTANT: HiJackThis does not determine what is good or bad. Do not make any changes to your computer settings unless you are an expert computer user.

Advanced users can use HiJackThis to remove unwanted settings or files.


2. System Requirements
----------------------

Operating System
    * Microsoft™ Windows™ XP
    * Microsoft™ Windows™ 2000
    * Microsoft™ Windows™ Me
    * Microsoft™ Windows™ 98
    * Microsoft™ Windows™ Vista
    * Microsoft™ Windows™ 7
    * Microsoft™ Windows™ 8
    * Microsoft™ Windows™ 8.1
    * Microsoft™ Windows™ 10

Software
    * Microsoft Internet Explorer 6.0 or newer
    * Mozilla™ Firefox™ 1.5 or newer


3. Sections
------------------

The different sections of hijacking possibilities have been separated into the following groups.
You can get more detailed information about an item by selecting it from the list of found items OR highlighting the relevant line below, and clicking 'Info on selected item'.

 R - Registry, StartPage/SearchPage changes
    R0 - Changed registry value
    R1 - Created registry value
    R2 - Created registry key
    R3 - Created extra registry value where only one should be
 F - IniFiles, autoloading entries
    F0 - Changed inifile value
    F1 - Created inifile value
    F2 - Changed inifile value, mapped to Registry
    F3 - Created inifile value, mapped to Registry
 N - Netscape/Mozilla StartPage/SearchPage changes
    N1 - Change in prefs.js of Netscape 4.x
    N2 - Change in prefs.js of Netscape 6
    N3 - Change in prefs.js of Netscape 7
    N4 - Change in prefs.js of Mozilla
 O - Other, several sections which represent:
    O1 - Hijack of auto.search.msn.com with Hosts file
    O2 - Enumeration of existing MSIE BHO's
    O3 - Enumeration of existing MSIE toolbars
    O4 - Enumeration of suspicious autoloading Registry entries
    O5 - Blocking of loading Internet Options in Control Panel
    O6 - Disabling of 'Internet Options' Main tab with Policies
    O7 - Disabling of Regedit with Policies
    O8 - Extra MSIE context menu items
    O9 - Extra 'Tools' menuitems and buttons
    O10 - Breaking of Internet access by New.Net or WebHancer
    O11 - Extra options in MSIE 'Advanced' settings tab
    O12 - MSIE plugins for file extensions or MIME types
    O13 - Hijack of default URL prefixes
    O14 - Changing of IERESET.INF
    O15 - Trusted Zone Autoadd
    O16 - Download Program Files item
    O17 - Domain hijack
    O18 - Enumeration of existing protocols and filters
    O19 - User stylesheet hijack
    O20 - AppInit_DLLs autorun Registry value, Winlogon Notify Registry keys
    O21 - ShellServiceObjectDelayLoad (SSODL) autorun Registry key
    O22 - SharedTaskScheduler autorun Registry key
    O23 - Enumeration of NT Services
    O24 - Enumeration of ActiveX Desktop Components


4. Command-line parameters
------------------

/startupscan - automatically scan the system (the same as button "Do a system scan only")
/autolog - automatically scan the system, save a logfile and open it
/silentautolog - the same as /autolog, except with no required user intervention
/ihatewhitelists - ignore all internal whitelists
/uninstall - remove all HiJackThis Registry entries, backups and quit
/deleteonreboot "c:\file.sys" - delete the file specified after system rebooting, using mechanism PendingFileRenameOperations

Keys for module StartupList:

Next keys affects on StartupList module settings only.
It can be launched manually from the section Config -> "Misc Tools" -> Generate StartupList Log.

/full - show some rarely important sections: Stub Paths, Explorer Check, Config.sys, Dosstart.bat, Superhidden Extensions, Regedit.exe Check, WinNT Services, Win9x VxD Services.
/complete - to include empty sections and unsuspicious data
/full     - to include several rarely-important sections
/verbose  - to add additional info on each section
/force9x  - to include Win9x-only startups even if running on WinNT
/forcent  - to include WinNT-only startups even if running on Win9x
/forceall - to include all Win9x and WinNT startups, regardless of platform
/history - StartupList module version changelog
/html - for a report in HTML format

How to use command line keys?

Create text file run.txt near the program HiJackThis.exe. Rename it into extension .bat
Right mouse click on file run.bat. Choose "Edit".

Example of .bat file code:

cd /d "%~dps0"
HiJackThis.exe /autolog /ihatewhitelists
 
Launch file run.bat.
If the version of operation system is Windows Vista/7/8/8.1 or Later, you must launch this .bat file by right mouse click and choose "Run as Administrator".


5. HJT version history
------------------

[v2.0.6 Beta (r35)]
* Added O4-64 - Autorun on Wow64 registry keys
* Added cheking of Opera version
* Added horizontal Scroll Bar to results screen
* Fixed modUtils to get Chrome Version XP/Win7
* determine the correct Windows version
* Changed URL where crash window refer to
* Removed ADS Spy
* Removed URL check when clicking 'Analyze this'
* Removed code of version checker
* Removed code of SpyBot and AdAware version checking

[v2.0.5 Beta (r21)]
* Fixed "No internet connection available" when pressing the button Analyze This
* Fixed the link of update website, now send you to sourceforge.net projects
* Fixed left-right scrollbar when in safe mode or low screen resolution
* 'default' restored hosts file didn't include ipv6 address entry	 
* support newest version of FireFox

[v2.0.4 (r10)]
* Fixed parser issues on winlogon notify
* Fixed issues to handle certain environment variables
* Rename HJT generates complete scan log

[v2.00.0]
* AnalyzeThis added for log file statistics
* Recognizes Windows Vista and IE7
* Fixed a few bugs in the O23 method
* Fixed a bug in the O22 method (SharedTaskScheduler)
* Did a few tweaks on the log format
* Fixed and improved ADS Spy
* Improved Itty Bitty Procman (processes are frozen before they are killed)
* Added listing of O4 autoruns from other users
* Added listing of the Policies Run items in O4 method, used by SmitFraud trojan
* Added /silentautolog parameter for system admins
* Added /deleteonreboot [file] parameter for system admins
* Added O24 - ActiveX Desktop Components enumeration
* Added Enhanced Security Confirguration (ESC) Zones to O15 Trusted Sites check

[v1.99.1]
* Added Winlogon Notify keys to O20 listing
* Fixed crashing bug on certain Win2000 and WinXP systems at O23 listing
* Fixed lots and lots of 'unexpected error' bugs
* Fixed lots of inproper functioning bugs (i.e. stuff that didn't work)
* Added 'Delete NT Service' function in Misc Tools section
* Added ProtocolDefaults to O15 listing
* Fixed MD5 hashing not working
* Fixed 'ISTSVC' autorun entries with garbage data not being fixed
* Fixed HijackThis uninstall entry not being updated/created on new versions
* Added Uninstall Manager in Misc Tools to manage 'Add/Remove Software' list
* Added option to scan the system at startup, then show results or quit if nothing found

[v1.99]
* Added O23 (NT Services) in light of newer trojans
* Integrated ADS Spy into Misc Tools section
* Added 'Action taken' to info in 'More info on this item'

[v1.98]
* Definitive support for Japanese/Chinese/Korean systems
* Added O20 (AppInit_DLLs) in light of newer trojans
* Added O21 (ShellServiceObjectDelayLoad, SSODL) in light of newer trojans
* Added O22 (SharedTaskScheduler) in light of newer trojans
* Backups of fixed items are now saved in separate folder
* HijackThis now checks if it was started from a temp folder
* Added a small process manager (Misc Tools section)

[v1.96]
* Lots of bugfixes and small enhancements! Among others:
* Fix for Japanese IE toolbars
* Fix for searchwww.com fake CLSID trick in IE toolbars and BHO's
* Attributes on Hosts file will now be restored when scanning/fixing/restoring it.
* Added several files to the LSP whitelist
* Fixed some issues with incorrectly re-encrypting data, making R0/R1 go undetected until a restart
* All sites in the Trusted Zone are now shown, with the exception of those on the nonstandard but safe domain list

[v1.95]
* Added a new regval to check for from Whazit hijack (Start Page_bak).
* Excluded IE logo change tweak from toolbar detection (BrandBitmap and SmBrandBitmap).
* New in logfile: Running processes at time of scan.
* Checkmarks for running StartupList with /full and /complete in HijackThis UI.
* New O19 method to check for Datanotary hijack of user stylesheet.
* Google.com IP added to whitelist for Hosts file check.

[v1.94]
* Fixed a bug in the Check for Updates function that could cause corrupt downloads on certain systems.
* Fixed a bug in enumeration of toolbars (Lop toolbars are now listed!).
* Added imon.dll, drwhook.dll and wspirda.dll to LSP safelist.
* Fixed a bug where DPF could not be deleted.
* Fixed a stupid bug in enumeration of autostarting shortcuts.
* Fixed info on Netscape 6/7 and Mozilla saying '%shitbrowser%' (oops).
* Fixed bug where logfile would not auto-open on systems that don't have .log filetype registered.
* Added support for backing up F0 and F1 items (d'oh!).

[v1.93]
* Added mclsp.dll (McAfee), WPS.DLL (Sygate Firewall), zklspr.dll (Zero Knowledge) and mxavlsp.dll (OnTrack) to LSP safelist.
* Fixed a bug in LSP routine for Win95.
* Made taborder nicer.
* Fixed a bug in backup/restore of IE plugins.
* Added UltimateSearch hijack in O17 method (I think).
* Fixed a bug with detecting/removing BHO's disabled by BHODemon.
* Also fixed a bug in StartupList (now version 1.52.1).

[v1.92]
* Fixed two stupid bugs in backup restore function.
* Added DiamondCS file to LSP files safelist.
* Added a few more items to the protocol safelist.
* Log is now opened immediately after saving.
* Removed rd.yahoo.com from NSBSD list (spammers are starting to use this, no doubt spyware authors will follow).
* Updated integrated StartupList to v1.52.
* In light of SpywareNuker/BPS Spyware Remover, any strings relevant to reverse-engineers are now encrypted.
* Rudimentary proxy support for the Check for Updates function.

[v1.91]
* Added rd.yahoo.com to the Nonstandard But Safe Domains list.
* Added 8 new protocols to the protocol check safelist, as well as showing the file that handles the protocol in the log (O18).
* Added listing of programs/links in Startup folders (O4).
* Fixed 'Check for Update' not detecting new versions.

[v1.9]
* Added check for Lop.com 'Domain' hijack (O17).
* Bugfix in URLSearchHook (R3) fix.
* Improved O1 (Hosts file) check.
* Rewrote code to delete BHO's, fixing a really nasty bug with orphaned BHO keys.
* Added AutoConfigURL and proxyserver checks (R1).
* IE Extensions (Button/Tools menuitem) in HKEY_CURRENT_USER are now also detected.
* Added check for extra protocols (O18).

[v1.81]
* Added 'ignore non-standard but safe domains' option.
* Improved Winsock LSP hijackers detection.
* Integrated StartupList updated to v1.4.

[v1.8]
* Fixed a few bugs.
* Adds detecting of free.aol.com in Trusted Zone.
* Adds checking of URLSearchHooks key, which should have only one value.
* Adds listing/deleting of Download Program Files.
* Integrated StartupList into the new 'Misc Tools' section of the Config screen!

[v1.71]
* Improves detecting of O6.
* Some internal changes/improvements.

[v1.7]
* Adds backup function! Yay!
* Added check for default URL prefix
* Added check for changing of IERESET.INF
* Added check for changing of Netscape/Mozilla homepage and default search engine.

[v1.61]
* Fixes Runtime Error when Hosts file is empty.

[v1.6]
* Added enumerating of MSIE plugins
* Added check for extra options in 'Advanced' tab of 'Internet Options'.

[v1.5]
* Adds 'Uninstall & Exit' and 'Check for update online' functions.
* Expands enumeration of autoloading Registry entries (now also scans for .vbs, .js, .dll, rundll32 and service)

[v1.4]
* Adds repairing of broken Internet access (aka Winsock or LSP fix) by New.Net/WebHancer
* A few bugfixes/enhancements

[v1.3]
* Adds detecting of extra MSIE context menu items
* Added detecting of extra 'Tools' menu items and extra buttons
* Added 'Confirm deleting/ignoring items' checkbox

[v1.2]
* Adds 'Ignorelist' and 'Info' functions

[v1.1]
* Supports BHO's, some default URL changes

[v1.0]
* Original release



6. StartupList module version history
------------------

v1.52
* Fixed stupid 'Bad filename or number' error at startup (hopefully)
* Fixed two bugs in function that reads settings from .ini files
* Added two more files to LSP files safelist (MS Firewall and
  DiamondCS)
* Fixed not detecting modified Shell line in XP (among others, this
  BIG bug affected two sections)
* Added listing of values in ShellServiceObjectDelayLoad regkey

v1.51
* Added switch: /full, which will show some rarely important
  sections that otherwise remain hidden:  Stub Paths, Explorer Check, 
  Config.sys, Dosstart.bat, Superhidden Extensions, 
  Regedit.exe Check, WinNT Services, Win9x VxD Services
* Lines in BAT files with both 'ECHO' and '>' are now shown
* Windows NT Logon/logoff scripts are now listed (new section)
* Rudimentary check for PendingFileRenameOperations in NT, located
  in above section. Also moved BootExecute check to this section

v1.5
* Added more files to safe list of LSP files
* REM/ECHO line in .bat files only listed with /complete switch
* Check for Policies\System\Shell= at SYSTEM.INI check
* Added enumeration of Windows NT/2000/XP services (only
  with /full switch)
* Also lists Windows 9x Vxd services (only with /full switch)

v1.4
* Added listing of Winsock LSP providers
* Fixed a NT bug with Load key

v1.35
* Fixed a few items not appearing in NT/2000/XP.
* Made Regedit check even more supple.

v1.34
* Added listing of drivers= line from system.ini
* Some more sections are now hidden if nothing interesting is there
* Enumeration of Stub Paths now shows disabled items
* Fixed a few bugs
* Workaround for Atguard 'From:' bug :)

v1.33
* Fixed some erroneous errors.
* Added listing of MSIE version.

v1.32
* Fixed a few bugs. That's basically it. :)

v1.31
* Finally added alternative (and better) method for listing processes
  in Windows NT/2000/XP (PSAPI.DLL needed for NT4)
* Improved filename extracting from shortcuts - StartupList should
  not be able to extract filenames with a 100% success rate
* Creation date is now displayed for Wininit.ini and Wininit.bak
* Added Regedit check
* Added listing of BHO's
* Added listing of Task Scheduler jobs
* Added listing of 'Download Program Files' (aka ActiveX Objects)

v1.3
* Added /html parameter, for a report in HTML format
* Lots of performance enhancements, more readble code (like you care :)
* Also some small upgrades/tweaks

v1.23
* Now also lists WININIT.BAK (the last WININIT.INI)

v1.22
* Made System.ini check platform independant (was Win9x only)
* The target file & path is now extracted from enumerated shortcuts
* Fixed MAJOR bug - GetWindowsVersion wasn't remembered, WinNT was
  assumed

v1.21
* Fixed some WinNT bugs
* Slightly improved Explorer.exe check in WinNT

v1.2
* Added WinNT-only startups
* Added Windows version check
* Added command line parameters /verbose, /complete,
  /force9x, /forcent and /forceall

v1.1
* Added RunOnceEx listing

v1.0
* Initial release
