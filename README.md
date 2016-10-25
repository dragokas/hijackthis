# Trend Micro HiJackThis v3

## Overview

HiJackThis is a free utility that generates an in depth report of registry and file settings of your computer. 

HiJackThis makes no separation between safe and unsafe settings in its scan results giving you the ability to selectively remove items from your machine. 

Also HiJackThis comes with several tools useful in manually remove malware from a computer:
 * StartupList 2
 * Process Manager
 * Uninstall manager
 * Hosts file manager
 * Alternative Data Spy
 * Delete file / service staff
 * Digital Signature Checker
 * Registry key unlocker

## Features

 * Lists the contents of key areas of the Registry and hard drive
 * Generate reports and presents them in an organized fashion
 * Does not target specific programs and URLs
 * Detects only the methods used by hijackers to force you onto their sites
 * The possibility of adding to autostart scanning at system boot

## Log analysis

IMPORTANT: HiJackThis does not make value based calls between what is considered good or bad.
It is important to exercise caution and avoid making changes to your computer settings, unless you have expert knowledge.

Unless you are expert, we recommend you to submit your case to online helper forums such as:
- English: www.bleepingcomputer.com ; http://www.cnet.com/forums/
- French: http://forum.malekal.com/
- Russian: http://safezone.cc/pravila/

## System Requirements

Operating System
  * Microsoft™ Windows™ 10 / 8.1 / 8 / 7 / Vista / XP

## Recent updates

###### 2.6.1.25
 - O22 - Tasks. Added recursive scanning (in depth) based on whitelists.
 - O25 - bug fixes.
 - O4 - added recognize of altering autostart shortcuts into PE EXE.
 - O4 - reworked (refactoring), fixed bug with lines duplicating, added new keys for checking.
 - improved interaction with ignorelist, fixed bugs. Now log contains number of entries in ignorelist if exist.
 - added context menu to scan results window: "Fix checked", "Info on selected", "Add to ignore list", "Search on Google", "ReScan".
 - unicode characters in report file (earlier it was '?' characters).

###### 2.6.2.0
 - O1 - Hosts: been improved function for reading Hosts on systems with active write protection.
 - O7 - IPSec subsection added (it's IP Security policies which allow fine tuning of IP packets filter).
 - O25 - WMI Events: simplified and trimmed to provide output of actual malware only; added whitelist.
 - O22 - Tasks: whitelist for OS Win XP/Vista/7/8/8.1/10 have been updated.
 - O4 - removed false entries that could apply to disable autorun items on Win 8+
 - O4 - added subsections:

..\StartupApproved\Run

..\StartupApproved\Run32

..\StartupApproved\StartupFolder
It's an analogue of MSConfig for Win 8+ (disabled autorun items).
 - O4 - added checking of keys HKCU/HKLM/HKU for:

..\Software\Microsoft\Windows\CurrentVersion\Run-

..\Software\Microsoft\Windows\CurrentVersion\RunServices-

..\Software\Microsoft\Windows\CurrentVersion\RunOnce-

..\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce-

..\Software\Microsoft\Command Processor -> AutoRun
 - O4 - Startup other users: (new subsection) - checking of folder "Startup" of other users.
 - O4 - MSConfig: renamed to MSConfig\startupreg.
 - O4 - MSConfig\startupfolder: added (disabled items of folder "Startup" on WinXP).
 - Removed flickering of desktop during the scan on some OS (bug in v1.19).
 - F3 Fix: not worked (bug in v.1.20)
 - ALT+TAB is now switches to the HJT active tool window instead of the main window.
 - Added tool for batch digital signature checking and whether file is Windows Protected.
 - Fixed an issue where the digital signature verification led to the connection to Internet.
 - Progress bar made over the entire width.
 - Added missing icons of program.
 - Removed info about OS Product Type and OS Suite Mask.
 - All mention of HKUS hive replaced by HKU.
 - Removed prefixes -64. Added prefixes -32 (meaning: key is under redirection, i.e. 32-bit key on 64-bit OS).
 - List of backups sorted in reverse order (beginning from the most fresh entry).
 - Fix O2, Fix O3: supplemented by removing the HKCU/HKLM keys for:

..\Software\Microsoft\Internet Explorer\Extension Compatibility\{CLSID}

..\Software\Microsoft\Windows\CurrentVersion\Ext\Stats\{CLSID}

..\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\{CLSID}

..\Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved\{CLSID}

..\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration\{CLSID}

..\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration{CLSID}

..\Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID -> {CLSID}