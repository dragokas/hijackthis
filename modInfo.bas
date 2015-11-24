Attribute VB_Name = "modInfo"
Option Explicit

Public Sub GetInfo(ByVal sItem$)
    Dim sMsg$
    On Error GoTo Error:
    If InStr(sItem, vbCrLf) > 0 Then sItem = Left(sItem, InStr(sItem, vbCrLf) - 1)
    Select Case Trim(Left(sItem, 3))
        Case "R0"
            sMsg = "A Registry value that has been changed " & _
            "from the default, resulting in a changed " & _
            "IE Search Page, Start Page, Search Bar Page " & _
            "or Search Assistant." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is restored to preset URL.)"
        Case "R1"
            sMsg = "A Registry value that has been created " & _
            "and is not present in a default Windows " & _
            "install nor needed, possibly resulting in a " & _
            "changed IE Search Page, Start Page, Search Bar " & _
            "Page or Search Assistant." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted.)"
        Case "R2"
            sMsg = "A Registry key that has been created " & _
            "and is not present in a default Windows " & _
            "install nor needed, possibly resulting in a " & _
            "changed IE Search Page, Start Page, Search Bar " & _
            "Page or Search Assistant." & vbCrLf & vbCrLf & _
            "(Action taken: Registry key is deleted, with everything in it.)"
        Case "R3"
            sMsg = "A Registry value that has been created " & _
            "in a key where only one value should be. Only " & _
            "is used for the URLSearchHooks regkey." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted, default URLSearchHook " & _
            "value is restored.)"
        Case "F0"
            sMsg = "An inifile value that has been changed " & _
            "from the default value, possibly resulting in " & _
            "program(s) loading at Windows startup. Often " & _
            "used to autostart a program that is even " & _
            "harder to disable." & vbCrLf & vbCrLf & _
            "Default: Shell=explorer.exe" & vbCrLf & _
            "Infected example: Shell=explorer.exe,openme.exe" & vbCrLf & vbCrLf & _
            "(Action taken: Default inifile value is restored.)"
        Case "F1"
            sMsg = "An inifile value that has been created " & _
            "and is not present in a default Windows " & _
            "install nor needed, possibly resulting in " & _
            "program(s) loading at Windows startup. Often " & _
            "used to autostart program(s) that are hard " & _
            "to disable." & vbCrLf & vbCrLf & _
            "Default: run= OR load=" & vbCrLf & _
            "Infected example: run=dialer.exe" & vbCrLf & vbCrLf & _
            "(Action taken: Inifile value is deleted.)"
        Case "N1"
            sMsg = "Netscape 4.x stores the browsers homepage " & _
            "the prefs.js file located in the user's Netscape " & _
            "directory. LOP.com has been known to change this " & _
            "value." & vbCrLf & vbCrLf & _
            "(Action taken: Setting is restored to preset URL.)"
        Case "N2", "N3", "N4"
            sMsg = "%SHITBROWSER% stores the browser's homepage in " & _
            "prefs.js file located deep in the 'Application Data' " & _
            "folder. The default search engine is also stored " & _
            "in this file. LOP.com has been known to change the " & _
            "homepage URL." & vbCrLf & vbCrLf & _
            "(Action taken: Setting is restored to preset URL.)"
            
            If Trim(Left(sItem, 3)) = "N2" Then sMsg = Replace(sMsg, "%SHITBROWSER%", "Netscape 6")
            If Trim(Left(sItem, 3)) = "N3" Then sMsg = Replace(sMsg, "%SHITBROWSER%", "Netscape 7")
            If Trim(Left(sItem, 3)) = "N4" Then sMsg = Replace(sMsg, "%SHITBROWSER%", "Mozilla")
        Case "O1"
            sMsg = "A change in the 'Hosts' system file " & _
            "Windows uses to lookup domain names before " & _
            "quering internet DNS servers, effectively " & _
            "making Windows believe that 'auto.search.msn" & _
            ".com' has a different IP than it really has " & _
            "and thus making IE open the wrong page when" & _
            "ever you enter an invalid domain name in the " & _
            "IE Address Bar." & vbCrLf & vbCrLf & _
            "Infected example: 213.67.109.7" & vbTab & "auto.search.msn.com" & vbCrLf & vbCrLf & _
            "(Action taken: Line is deleted from hosts file.)"
        Case "O2"
            sMsg = "A BHO (Browser Helper Object) is a specially " & _
            "crafted program that integrates into IE, and " & _
            "has virtually unlimited access rights on your " & _
            "system. Though BHO's can be helpful (like the " & _
            "Google Toolbar), hijackers often use them for " & _
            "malicious purposes such as tracking your " & _
            "online behaviour, displaying popup ads etc." & vbCrLf & vbCrLf & _
            "(Action taken: Registry key and CLSID key are deleted, BHO dll file is deleted.)"
        Case "O3"
            sMsg = "IE Toolbars are part of BHO's (Browser Helper " & _
            "Objects) like the Google Toolbar that are " & _
            "helpful, but can also be annoying and malicious " & _
            "by tracking your behaviour and displaying " & _
            "popup ads." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted.)"
        Case "O4"
            sMsg = "This part of the scan checks for several " & _
            "suspicious entries that autoload when Windows " & _
            "starts. Autoloading entries can load " & _
            "a Registry script, VB script or JavaScript" & _
            "file, possibly causing the IE Start Page, " & _
            "Search Page, Search Bar and Search Assistant " & _
            "to revert back to a hijacker's page after a " & _
            "system reboot. Also, a DLL file can be loaded " & _
            "that can hook into several parts of your system." & vbCrLf & vbCrLf & _
            "Infected examples:" & vbCrLf & vbCrLf & _
            "regedit c:\windows\system\sp.tmp /s" & vbCrLf & _
            "KERNEL32.VBS" & vbCrLf & _
            "c:\windows\temp\install.js" & vbCrLf & _
            "rundll32 C:\Program Files\NewDotNet\newdotnet4_5.dll,NewDotNetStartup" & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted.)"
        Case "O5"
            sMsg = "Modifying CONTROL.INI can cause Windows " & _
            "to hide certain icons in the Control Panel. " & _
            "Though originally meant to speed up loading of " & _
            "Control Panel and reducing clutter, it can be " & _
            "used by a hijacker to prevent access to the " & _
            "'Internet Options' window." & vbCrLf & vbCrLf & _
            "Infected example:" & vbCrLf & "[don't load]" & _
            vbCrLf & "inetcpl.cpl=yes OR inetcpl.cpl=no" & vbCrLf & vbCrLf & _
            "(Action taken: Line is deleted from Control.ini file.)"
        Case "O6"
            sMsg = "Disabling of the 'Internet Options' menu " & _
            "menu entry in the 'Tools' menu of IE is done " & _
            "by using Windows Policies. Normally used by " & _
            "administrators to restrict their users, it can " & _
            "be used by hijackers to prevent access to the " & _
            "'Internet Options' window." & vbCrLf & vbCrLf & _
            "StartPage Guard also uses Policies to restrict " & _
            "homepage changes, done by hijackers." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted.)"
        Case "O7"
            sMsg = "Disabling of Regedit is done by using " & _
            "Windows Policies. Normally used by administrators " & _
            "to restrict their users, it can be used by " & _
            "hijackers to prevent access to the Registry editor." & _
            " This results in a message saying that your " & _
            "administrator has not given you privilege to use " & _
            "Regedit when running it." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted.)"
        Case "O8"
            sMsg = "Extra items in the context (right-click) menu " & _
            "can prove helpful or annoying. Some recent hijackers " & _
            "add an item to the context menu. The MSIE PowerTweaks " & _
            "Web Accessory adds several useful items, among which " & _
            """Highlight"", ""Zoom In/Out"", ""Links list"", """ & _
            "Images list"" and ""Web Search""." & vbCrLf & vbCrLf & _
            "(Action taken: Registry key is deleted.)"
            
        Case "O9"
            sMsg = "Extra items in the MSIE 'Tools' menu and extra " & _
            "buttons in the main toolbar are usally present as " & _
            "branding (Dell Home button) or after system updates " & _
            "(MSN Messenger button) and rarely by hijackers. The " & _
            "MSIE PowerTweaks Web Accessory adds two menu items, " & _
            "being ""Add site to Trusted Zone"" and ""Add site to " & _
            "Restricted Zone""." & vbCrLf & vbCrLf & _
            "(Action taken: Registry key is deleted.)"
            
        Case "O10"
            sMsg = "The Windows Socket system (Winsock) uses a list of " & _
            "providers for resolving DNS names (i.e. translating www." & _
            "microsoft.com into an IP address). This is called the Layered " & _
            "Service Provider (LSP). A few programs are capable of " & _
            "injecting their own (spyware) providers in the LSP. If files " & _
            "referenced by the LSP are " & _
            "missing or the 'chain' of providers is broken, none of the " & _
            "programs on your system can access the Internet. Removing " & _
            "references to missing files and repairing the chain will " & _
            "restore your Internet access." & vbCrLf & "So far, only a few " & _
            "programs use a Winsock hook." & vbCrLf & vbCrLf & _
            "Note: This is a risky procedure. If it should fail, " & _
            "get LSPFix from http://www.cexx.org/lspfix.htm to repair the " & _
            "Winsock stack." & vbCrLf & vbCrLf & _
            "(Action taken: none. Use LSPFix to modify the Winsock stack.)"
            '"(Action taken: Registry key is deleted from chain, gaps are fixed.)"
            
        Case "O11" 'MSIE options group
            sMsg = "The options in the 'Advanced' tab of MSIE options " & _
            "are stored in the Registry, and extra options can be " & _
            "added easily by creating extra Registry keys. Very " & _
            "rarely, spyware/hijackers add their own options there " & _
            "which are hard to remove. E.g. CommonName adds a section " & _
            "'CommonName' with a few options." & vbCrLf & vbCrLf & _
            "(Action taken: Registry key is deleted, with everything in it.)"
            
        Case "O12" 'MSIE plugins
            sMsg = "Plugins handle filetypes that aren't supported " & _
            "natively by MSIE. Common plugins handle Macromedia " & _
            "Flash, Acrobat PDF documents and Windows Media formats, " & _
            "enabling the browser to open these itself instead of " & _
            "launching a separate program. When hijackers or spyware " & _
            "add plugins for their filetypes, the danger exists that " & _
            "they get reinstalled if everything except the plugin has " & _
            "been removed, and the browser opens such a file." & vbCrLf & vbCrLf & _
            "(Action taken: Registry key is deleted, with everything in it.)"
        
        Case "O13" 'DefaultPrefix
            sMsg = "When you type an URL into MSIE's Address bar without " & _
            "the prefix (http://), it is automatically added when you " & _
            "hit Enter. This prefix is stored in the Registry, together " & _
            "with the default prefixes for FTP, Gopher and a few other " & _
            "protocols. When a hijacker changes these to the URL of his " & _
            "server, you always get redirected there when you forget to " & _
            "type the prefix. Prolivation uses this hijack." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is restored to default data.)"
            
        Case "O14" 'IERESET.INF
            sMsg = "When you hit 'Reset Web Settings' on the 'Programs' tab " & _
            "of the MSIE Options dialog, your homepage, search page and a " & _
            "few other sites get reset to their defaults. These defaults are " & _
            "stored in C:\Windows\Inf\Iereset.inf. When a hijacker changes these " & _
            "to his own URLs, you get (re)infected rather than cured when you " & _
            "click 'Reset Web Settings'. SearchALot uses this hijack." & vbCrLf & vbCrLf & _
            "(Action taken: Value in Inf file is restore to default data.)"
        
        Case "O15" 'Trusted Zone Autoadd
            sMsg = "Websites in the Trusted Zone (see Internet Options," & _
            "Security, Trusted Zone, Sites) are allowed to use normally " & _
            "dangerous scripts and ActiveX objects normal sites aren't " & _
            "allowed to use. Some programs will " & _
            "automatically add a site to the Trusted Zone without you " & _
            "knowing. Only a very few legitimate programs are known to do this " & _
            "(Netscape 6 is one of them) and a lot of browser hijackers" & _
            "add sites with ActiveX content to them." & vbCrLf & vbCrLf & _
            "(Action taken: Registry key is deleted, with everything in it.)"
            
        Case "O16" 'Downloaded Program Files
            sMsg = "The Download Program Files (DPF) folder in your " & _
            "Windows base folder holds various types of programs " & _
            "that were downloaded from the Internet. These programs " & _
            "are loaded whenever Internet Explorer is active." & _
            "Legitimate examples are the Java VM, Microsoft XML " & _
            "Parser and the Google Toolbar." & vbCrLf & _
            "Unfortunately, due to the lack security of IE, malicious " & _
            "sites let IE automatically download porn dialers, " & _
            "bogus plugins, ActiveX Objects etc to this folder, " & _
            "which haunt you with popups, huge phone bills, random " & _
            "crashes, browser hijackings and whatnot."
        
        Case "O17" 'Domain hijack
            sMsg = "Windows uses several registry values as a help " & _
            "to resolve domain names into IP addresses. Hijacking " & _
            "these values can cause all programs that use the Internet " & _
            "to be redirected to other pages for seemingly unknown " & _
            "reasons." & vbCrLf & _
            "New versions of Lop.com use this method, together with a " & _
            "(huge) list of cryptic domains." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted.)"
        
        Case "O18" 'Protocol & Filter
            sMsg = "A protocol is a 'language' Windows uses to 'talk' " & _
            "to programs, servers or itself. Webservers use the " & _
            "'http:' protocol, FTP servers use the 'ftp:' protocol, " & _
            "Windows Explorer uses the 'file:' protocol. Introducing " & _
            "a new protocol to Windows or changing an existing one " & _
            "can burrow deep into how Windows handles files." & vbCrLf & _
            "CommonName and Lop.com both register a new protocol " & _
            "when installed (cn: and ayb:)." & vbCrLf & _
            vbCrLf & _
            "The filters are content types accepted by Internet Explorer " & _
            "(and internally by Windows). If a filter exists for a content " & _
            "type, it passes through the file handling that content type " & _
            "first. Several variants of the CWS trojan add a text/html " & _
            "and text/plain filters, allowing them to hook all of the webpage " & _
            "content passed through Internet Explorer." & vbCrLf & vbCrLf & _
            "(Action taken: Registry key is deleted, with everything in it.)"
            
        Case "O19" 'User stylesheet
            sMsg = "IE has an option to use a user-defined stylesheet " & _
            "for all pages instead of the default one, to enable " & _
            "handicapped users to better view the pages." & vbCrLf & _
            "An especially vile hijacking method made by Datanotary " & _
            "has surfaced, which overwrites any stylesheet the user has " & _
            "setup and replaces it with one that causes popups, as well " & _
            "a system slowdown when typing or loading pages with many " & _
            "pictures." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted.)"
        
        Case "O20"  'AppInit_DLLs + WinLogon Notify subkeys
            sMsg = "Files specified in the AppInit_DLLs Registry value " & _
            "are loaded very early in Windows startup and stay in memory " & _
            "until system shutdown. This way of loading a .dll is hardly " & _
            "ever used, except by trojans." & vbCrLf & _
            "The WinLogon Notify Registry subkeys load dll files into memory " & _
            "at about the same point in the boot process, keeping them " & _
            "loaded into memory until the session ends. Apart from several " & _
            "Windows system components, the programs VX2, ABetterInternet " & _
            "and Look2Me use this Registry key." & vbCrLf & _
            "Since both methods ensure the dll file stays loaded in " & _
            "memory the entire time, fixing this won't help if the dll " & _
            "puts back the Registry value or key immediately. In such cases, " & _
            "the use of the 'Delete file on reboot' function or KillBox is " & _
            "recommended to first delete the file." & vbCrLf & vbCrLf & _
            "(Action taken for AppInit_DLLs: Registry value is cleared, but not deleted.)" & vbCrLf & _
            "(Action taken for Winlogon Notify: Registry key is deleted."
            
        Case "O21"  'ShellServiceObjectDelayLoad
            sMsg = "This is an undocumented Registry key that contains a list " & _
            "of references to CLSIDs, which in turn reference .dll files " & _
            "that are then loaded by Explorer.exe at system startup. " & _
            "The .dll files stay in memory until Explorer.exe quits, which is " & _
            "achieved either by shutting down the system or killing the shell " & _
            "process." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted, CLSID key is deleted.)"
            
        Case "O22"  'SharedTaskScheduler
            sMsg = "This is an undocumented Registry key that contains a list " & _
            "of CLSIDs, which in turn reference .dll files that are loaded " & _
            "by Explorer.exe at system startup. The .dll files stay in memory " & _
            "until Explorer.exe quits, which is achieved either by shutting " & _
            "down the system or killing the shell process." & vbCrLf & vbCrLf & _
            "(Action taken: Registry value is deleted, CLSID key is deleted.)"
            
        Case "O23" 'NT Services
            sMsg = "The 'Services' in Windows NT4, Windows 2000, Windows XP and " & _
            "Windows 2003 are a special type of programs that are essential to " & _
            "the system and are required for proper functioning of the system. " & _
            "Service processes are started before the user logs in and are " & _
            "protected by Windows. They can only be stopped " & _
            "from the services dialog in the Administrative Tools window." & vbCrLf & _
            "Malware that registers itself as a service is subsequently also harder " & _
            "to kill." & vbCrLf & vbCrLf & _
            "(Action taken: services is disabled and stopped. Reboot needed.)"
        
        Case "O24"
            sMsg = "Desktop Components are ActiveX objects that can be made " & _
            "part of the desktop whenever Active Desktop is enabled (introduced " & _
            "in Windows 98), where it runs as a (small) website widget." & _
            vbCrLf & "Malware misuses this feature by setting the desktop " & _
            "background to a local HTML file with a large, bogus warning." & _
            vbCrLf & vbCrLf & _
            "(Action taken: ActiveX object is deleted from Registry.)"
        Case Else
            Exit Sub
    End Select
    sMsg = "Detailed information on item " & Trim(Left(sItem, 3)) & ":" & vbCrLf & vbCrLf & sMsg
    MsgBox sItem & vbCrLf & vbCrLf & sMsg, vbInformation
    Exit Sub
    
Error:
    ErrorMsg "modInfo_GetInfo", Err.Number, Err.Description, "sItem=" & sItem
End Sub
