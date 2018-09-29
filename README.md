# HiJackThis Fork v3

Hi, I am Polshyn Stanislav from Ukraine ([colony of USA](https://youtu.be/puLPOPKypEg?t=7218), [Chernobyl](https://en.wikipedia.org/wiki/Chernobyl_disaster), [Klitschko](https://en.wikipedia.org/wiki/Wladimir_Klitschko), [Na'Vi](https://en.wikipedia.org/wiki/Natus_Vincere)). 
I am happy to present a continuation of Trend Micro HiJackThis development.

## Overview

HiJackThis Fork is a free utility for Microsoft Windows that scans your computer for settings changed by adware, spyware, malware and other unwanted programs.

HiJackThis Fork primarily detects **hijacking methods** rather than comparing items against a pre-built database.  This allows it to detect new or previously unknown malware - but it also makes **no distinction** between safe and unsafe items.  Users are expected to research all scanned items, and only remove items from their PC when absolutely appropriate.

Therefore, FALSE POSITIVES ARE LIKELY. If you are ever unsure, you should consult with a knowledgeable expert BEFORE deleting anything.


## Download
[![](https://dragokas.com/tools/img/hjt/Icon_mini.png)](https://dragokas.com/tools/HiJackThis.zip)
[Pre-built binary (release version) for Windows](https://dragokas.com/tools/HiJackThis.zip)

[Nightly build (private test version) for Windows](https://dragokas.com/tools/HiJackThis_test.zip)

![](https://dragokas.com/tools/img/hjt/Scanning.png)

## Features

 * Lists non-default settings in the registry, hard drive and memory related to autostart
 * Generates organized, easily readable reports
 * Does not use a database of specific malware, adware, etc
 * Detects potential *methods* used by hijackers
 * Can be configured to automatically scan at system boot up
 
## Advantages

 * Short logs
 * Fast scans
 * No need to manually create fixing scripts
 * No need for Internet access or recurring database updates
 * Already familiar to many people
 * Portable

## New in version 3

 * Detects several new hijacking methods
 * Fully supports new Windows versions
 * New and updated supplementary tools
 * Improved interface, security and backups

HiJackThis also comes with several useful tools for manually removing malware from a computer:
 * StartupList 2 **(\*new\*)**
 * Process Manager
 * Uninstall manager
 * Hosts file manager
 * Alternative Data Spy
 * Delete file / service staff
 * Digital Signature Checker **(\*new\*)**
 * Registry key unlocker **(\*new\*)**
 * Check Browsers' LNK & ClearLNK (as downloadable component) **(\*new\*)**

## Log analysis

**IMPORTANT**: HiJackThis Fork does not make value-based calls on what is considered good or bad.
You must exercise caution when using this tool. Avoid making changes to your computer settings without thoroughly studying the consequences of each change.

If you are not already an expert, we recommend submitting your case to an online help forum. Here are some suggestions:
- English: [Our GitHub](https://github.com/dragokas/hijackthis/wiki/How-to-make-a-request-for-help-in-the-PC-cure-section%3F) ; [GeeksToGo](http://www.geekstogo.com/forum/topic/2852-malware-and-spyware-cleaning-guide/) ;  [BleepingComputer](https://www.bleepingcomputer.com/forums/t/34773/preparation-guide-for-use-before-using-malware-removal-tools-and-requesting-help/)
- Russian: [SafeZone](http://safezone.cc/pravila/) ; [CyberForum](http://www.cyberforum.ru/viruses/thread49792.html) ; [OSZone](http://forum.oszone.net/thread-98169.html) ; [SoftBoard](https://softboard.ru/topic/51343-правила-подраздела/) ; [THG](http://www.thg.ru/forum/showthread.php?t=92236) ; [VirusInfo](https://virusinfo.info/showthread.php?t=1235) ; [KasperskyClub](https://forum.kasperskyclub.ru/index.php?showtopic=43640)

> Note: currently, only Russian-speaking anti-malware supporting team (e.g., [VIRUSNET association](https://github.com/VIRUSNET-Association)) can provide direct analysis of HiJackThis logs in [our github 'Issues' section](https://github.com/dragokas/hijackthis/wiki/How-to-make-a-request-for-help-in-the-PC-cure-section%3F). Please feel free to ask help there.

## Technical support

 * [Actual short User's manual](http://dragokas.com/tools/help/hjt_tutorial.html) (in English)
 * [Actual complete User's manual](https://regist.safezone.cc/hijackthis_help/hijackthis.html) (in Russian)
 * [Recent updates by the author](https://safezone.cc/threads/27470/) (in Russian)
 * [Additional instructions on Wiki-pages](https://github.com/dragokas/hijackthis/wiki)
 * Discussion and news are in [this topic](https://safezone.cc/threads/hijackthis-fork-i-voprosy-k-razrabotchikam.28770/) (in Russian) or on [GeeksToGo](http://www.geekstogo.com/forum/topic/361755-hijackthisfork-improvement-development-bug-reports/) (in English; access restricted to experts only) or on our [GitHub page](https://github.com/dragokas/hijackthis/issues/4) (for everybody).
 * You can also freely ask questions, report bugs, or propose improvements by [creating an issue on GitHub](https://github.com/dragokas/hijackthis/issues)

## System Requirements

Operating System
  * Microsoft™ Windows™ 10 / 8.1 / 8 / 7 / Vista / XP / 2000 (32/64-bit desktop and server)

## Copyrights

 * **Polshyn Stanislav** { [@dragokas](https://github.com/dragokas) } - author of fork (major v3 and all post-v2.0.6 updates), refactoring, additions, tools integration
 * **Merijn Bellekom** { [@mrbellek](https://github.com/mrbellek) } - original author, author of the new [StartupList v2](https://github.com/mrbellek/StartupList2) and [ADS Spy](https://github.com/mrbellek/ADSspy)
 * **Trend Micro** { [@trendmicro](https://github.com/trendmicro) } - owner of the [original version](https://sourceforge.net/projects/hjt/) (2.0.5)
### Thanks to:
 * **regist** (VIRUSNET) { [@regist](https://forum.kasperskyclub.ru/index.php?showuser=44533) } - for the valuable tips and ideas, user's manual, database updates, closed and beta-testing
 * **Sandor** (VIRUSNET) { [@Sandor-Helper](https://github.com/Sandor-Helper) } - for the beta-testing, PC treatment on GitHub and forums of association
 * **SafeZone.cc team** (general [VIRUSNET](https://safezone.cc/threads/specialisty-i-internet-resursy-nekommercheskoj-dobrovolnoj-associacii-virusnet.22945/) community) - for promotion and support, PC treatment on forums of association
 * **Fernando Mercês** { [@merces](https://github.com/merces) } (Trend Micro) - coordinator of original HJT, for the tips, suggestions and promotion
 * **Loucif Kharouni** { [@loucifkharouni](https://github.com/loucifkharouni) } (Trend Micro) - coordinator of original HJT, for the tips & suggestions

HiJackThis Fork by Alex Dragokas is a continuation of Trend Micro HiJackThis development, based on [v.2.0.6](https://sourceforge.net/p/hjt/code/HEAD/tree/beta/2.0.6/). It was initially supported by Trend Micro, but they have since refused support and closed the GitHub repository.
HiJackThis Fork is distributed under the [GPLv2 license](https://github.com/dragokas/hijackthis/blob/devel/LICENSE.md). It also includes several tools and plugins available as freeware.

## Donate

For more than three years, I have maintained this project in my free time.
If you find it useful, you can support me by donating any amount to:
 * BTC: [17hkU3eKPngHrG3P9uqXwMLE3ztmtfGDZ4](https://dragokas.com/tools/img/BTC_QR.png)
 * [Yandex.Money](https://money.yandex.ru/to/410011191892975)
 * WebMoney:
 - Z389963582741
 - R963062285529

## Other projects

You may also find my other programs useful:
- [Check Browsers' LNK](https://toolslib.net/downloads/viewdownload/80-check-browsers-lnk/) & [ClearLNK](https://toolslib.net/downloads/viewdownload/81-clearlnk/) to cure shortcuts
- [Different tools](https://github.com/SafeZone-cc) at SafeZone repository.
- [My articles, tutorials and research](http://www.cyberforum.ru/blogs/218284/blog3628.html)
