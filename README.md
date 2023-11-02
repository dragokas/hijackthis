## Download
[![](https://dragokas.com/tools/img/hjt/Icon_mini.png)](https://dragokas.com/tools/HiJackThis_test.zip)
[Latest build](https://dragokas.com/tools/HiJackThis_test.zip) [v3.x Alpha] - test version

[![](https://dragokas.com/tools/img/hjt/Icon_mini.png)](https://dragokas.com/tools/HiJackThis_stable.zip)
[Stable build](https://dragokas.com/tools/HiJackThis_stable.zip) [v2 outdated] - not updated anymore

![](https://dragokas.com/tools/img/hjt/main_menu2.png)

_(this is alpha-version - major changes are in progress; some functions may work unstable; the plus+ means smooth introduction of detection of modern malware and step-by-step moving to replace some modules on C++ to access 64-bit processes in normal way and multi-threading; it is also a differentiation among other possible forks around)_

# HijackThis+

**HijackThis+ (Plus)** (previously called: HiJackThis Fork v3) is a fork and a continuation of the original [Trend Micro HijackThis by Merijn Bellekom](https://sourceforge.net/projects/hjt/) development, once a well-known tool.

At the moment, it is a step-by-step 100% rewritten source code of the original engine, aimed to provide a full compatiblity with the most recent Windows OS and a balance beetween compiling very fast results in logfile and combatting with the most popular malware, inluding the one not known to other antiviruses.

It is made by Alex Dragokas - a lawyer, security observer and malware researcher from Ukraine ([Chernobyl](https://en.wikipedia.org/wiki/Chernobyl_disaster), [Na'Vi](https://en.wikipedia.org/wiki/Natus_Vincere), [Щедрик](https://www.youtube.com/watch?v=ZZEMvVcf5-Q), Male slavery, nazism, fascism, [concentration camp](https://www.youtube.com/watch?v=3ASp9tr_-DQ), War of money, authority and blackmail, USA Bio-labs, [radioactive contamination from British](https://www.youtube.com/watch?v=ajY4qcc4OWc), Bombs, drones, whores and crazy people, Colony of USA). Yankee go home! F\*ck Russia, F\*ck Ukraine, F\*ck USA, F\*ck all the world giving weapon for killing the people, imposing sanctions against Ukrainian people in Crimea. F\*cking ukrainophobs, rusophobs and castrated negrophobes, provoking genocide of Ukrainian people with hundred tons of "free" cluster and nuclear weapon, exploding at our land like your own polygon.

## Overview

HijackThis+ is a free utility for Microsoft Windows that scans your computer for settings changed by adware, spyware, malware and other unwanted programs. Shortly, consider it like Sysinternals [Autoruns](https://learn.microsoft.com/en-us/sysinternals/downloads/autoruns).

The difference from classical antiviruses is the ability to function without constant database updates, because HijackThis+ primarily detects **hijacking methods** rather than comparing items against a pre-built database (signatures). This allows it to detect new or previously unknown malware - but it also makes **no distinction** between safe and unsafe items. Users are expected to research all scanned items manually, and only remove items from their PC when absolutely appropriate.

Therefore, FALSE POSITIVES ARE LIKELY. If you are ever unsure, you should consult with a knowledgeable expert BEFORE deleting anything.

HijackThis+ is not a replacement of a classical antivirus. It doesn't provide a real-time protection, because it is a passive scanner only. Consider it as an addition. However, you can use it in form of boot-up automatical scanner in the following way: 
 * Run the scanning by clicking "Do a system scan only"
 * Add all items in the ignore-list
 * Set up boot-up scan in menu "File" - "Settings" - "Add HijackThis to startup"
 * Next time when user logged in, HijackThis will silently scan your OS and display UI if only new records in your system were found.

## Tutorial

 * Please, refer to the [List of tutorials](https://github.com/dragokas/hijackthis/wiki/HJT:-Tutorial)

## Features

 * Lists non-default settings in the registry, hard drive and memory related to autostart
 * Generates organized, easily readable reports
 * Does not use a database of specific malware, adware, etc
 * Detects potential *methods* used by hijackers
 * Can be configured to automatically scan at system boot up
 
## Advantages

 * Short logs
 * Fast scans
 * Not necessarily to create fixing scripts manually
 * No need for internet access or recurring database updates
 * Already familiar to many people
 * Portable

## New in version 2.6+

 * Detects several new hijacking methods
 * Fully supports new versions of OS Windows 
 * New and updated supplementary tools
 * Improved interface, security and backups

HijackThis+ also comes with several modules useful for specific analysis and removing malware from a computer:
 * StartupList 2 **(\*new\*)**
 * Process Manager
 * Uninstall Manager
 * Hosts File Manager
 * Alternative Data Spy
 * Services Removing Tool
 * Batch Digital Signature Checker **(\*new\*)**
 * Registry Key Type Analyzer **(\*new\*)**
 * Registry Key Unlocker **(\*new\*)**
 * Files Unlocker **(\*new\*)**
 * Check Browsers' LNK & ClearLNK (as downloadable components) **(\*new\*)**

## Log analysis

**IMPORTANT**: HijackThis+ does not make value-based calls on what is considered good or bad.
You must exercise caution when using this tool. Avoid making changes to your computer settings without thoroughly studying the consequences of each change.

If you are not already an expert, we recommend submitting your case to an online help forum. Here are some suggestions:
- English: [Our GitHub](https://github.com/dragokas/hijackthis/wiki/How-to-make-a-request-for-help-in-the-PC-cure-section%3F) ; [GeeksToGo](http://www.geekstogo.com/forum/topic/2852-malware-and-spyware-cleaning-guide/) ;  [BleepingComputer](https://www.bleepingcomputer.com/forums/t/34773/preparation-guide-for-use-before-using-malware-removal-tools-and-requesting-help/)
- Russian: [SafeZone](https://safezone.cc/pravila/) ; [CyberForum](https://www.cyberforum.ru/viruses/thread49792.html) ; [OSZone](http://forum.oszone.net/thread-98169.html) ; [SoftBoard](https://softboard.ru/topic/51343-правила-подраздела/) ; [THG](http://www.thg.ru/forum/showthread.php?t=92236) ; [VirusInfo](https://virusinfo.info/showthread.php?t=1235) ; [KasperskyClub](https://forum.kasperskyclub.ru/index.php?showtopic=43640)

> Note: currently, only [VIRUSNET association](https://github.com/VIRUSNET-Association) can provide direct analysis of HijackThis+ logs in [our github 'Issues' section](https://github.com/dragokas/hijackthis/wiki/How-to-make-a-request-for-help-in-the-PC-cure-section%3F). Please feel free to ask help there (English/Russian only).

## Technical support

 * [Actual short User's manual](https://dragokas.com/tools/help/hjt_tutorial.html) (in English)
 * [Actual complete User's manual](https://regist.safezone.cc/hijackthis_help/hijackthis.html) (in Russian)
 * [Recent updates by the author](https://safezone.cc/threads/27470/) (in Russian)
 * [Additional instructions on Wiki-pages](https://github.com/dragokas/hijackthis/wiki)
 * Discussion and news are in [this topic](https://safezone.cc/threads/hijackthis-fork-i-voprosy-k-razrabotchikam.28770/) (in Russian) or on [GeeksToGo](https://www.geekstogo.com/forum/topic/361755-hijackthisfork-improvement-development-bug-reports/) (in English; access restricted to experts only) or on our [GitHub page](https://github.com/dragokas/hijackthis/discussions/137) (for everybody).
 * You can also freely ask questions, report bugs, or propose improvements by [creating an issue on GitHub](https://github.com/dragokas/hijackthis/issues)

## System requirements

Operating system
  * Microsoft™ Windows™ 11 / 10 / 8.1 / 8 / 7 / Vista / XP (32/64-bit desktop and server)

## Copyrights

 * **Alex Dragokas** { [@dragokas](https://github.com/dragokas) } - author of fork (major v3 and all post-v2.0.6 updates), refactoring, additions, tools integration
 * **Merijn Bellekom** { [@mrbellek](https://github.com/mrbellek) } - original author, author of the new [StartupList v2](https://github.com/mrbellek/StartupList2) and [ADS Spy](https://github.com/mrbellek/ADSspy)
 * **Trend Micro** { [@trendmicro](https://github.com/trendmicro) } - owner of the [original version](https://sourceforge.net/projects/hjt/) (2.0.5)
### Thanks to:
 * **regist** (VIRUSNET) { [@regist](https://forum.kasperskyclub.ru/index.php?showuser=44533) } - for the valuable tips and ideas, user's manual, database updates, closed and beta-testing
 * **Sandor** (VIRUSNET) { [@Sandor-Helper](https://github.com/Sandor-Helper) } - for the beta-testing, lot of reports, PC treatment on GitHub and forums of association
 * **akok** (VIRUSNET) { [@akokSZ](https://github.com/akokSZ) } - for product promotion, providing a platform for tests and discussion, help with resolving conflicts with antiviruses
 * **SafeZone.cc team** (general [VIRUSNET](https://github.com/VIRUSNET-Association/VIRUSNET) community) - for promotion and support, feedback and bug reports, PC treatment on forums of association
 * **Fernando Mercês** { [@merces](https://github.com/merces) } (Trend Micro) - coordinator of original HJT, for the tips, suggestions and promotion
 * **Loucif Kharouni** { [@loucifkharouni](https://github.com/loucifkharouni) } (Trend Micro) - coordinator of original HJT, for the tips & suggestions

HijackThis+ by Alex Dragokas is a continuation of Trend Micro HijackThis development, based on [v.2.0.6](https://sourceforge.net/p/hjt/code/HEAD/tree/beta/2.0.6/) branch and 100% rewritten at the moment. HijackThis+ was initially supported by Trend Micro, but they have since refused support and closed its GitHub repository.
HijackThis+ is distributed under the initial [GPLv2 license](https://github.com/dragokas/hijackthis/blob/devel/LICENSE.md). It also includes several tools and plugins available as freeware.

## Reviews & Mirrors
(clickable)

[![](https://dragokas.com/tools/img/hjt/softpedia-reward.png)](https://www.softpedia.com/get/Security/Security-Related/HiJackThis-Fork.shtml) [![](https://dragokas.com/tools/img/hjt/mg_certified.gif)](https://www.majorgeeks.com/files/details/hijackthis_fork.html) [![](https://dragokas.com/tools/img/hjt/comss_one.png)](https://www.comss.ru/page.php?id=6749)
[![](https://dragokas.com/tools/img/hjt/chocolatey_badge2.png)](https://chocolatey.org/packages/hijackthis)

**Note:** These mirrors belong to other companies. They are non-official.

### More references:
[Wikipedia (EN)](https://en.wikipedia.org/wiki/HijackThis)
[Wikipedia (RU)](https://ru.wikipedia.org/wiki/HijackThis)
[LabRats - Intro to HijackThis (Video)](https://www.youtube.com/watch?v=oZU6mTkGgGY)

// TODO... Please, report more if you know :)

## Other projects

You may also find my other programs useful:
- [Check Browsers' LNK](https://toolslib.net/downloads/viewdownload/80-check-browsers-lnk/) & [ClearLNK](https://toolslib.net/downloads/viewdownload/81-clearlnk/) to cure shortcuts
- [Different tools](https://github.com/SafeZone-cc) at SafeZone repository.
- [My articles, tutorials and research](https://www.cyberforum.ru/blogs/218284/blog3628.html) (in Russian)
