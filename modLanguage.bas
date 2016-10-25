Attribute VB_Name = "modLanguage"
Option Explicit

Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDefaultUILanguage Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32.dll" () As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32.dll" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoW" (ByVal LCID As Long, ByVal LCTYPE As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long

Private Const LOCALE_SENGLANGUAGE = &H1001&

Private sLines$(), bDontPrompt As Boolean

Public Sub LoadLanguageSettings()
    With OSVer
        .LangDisplayCode = GetUserDefaultUILanguage Mod &H10000
        .LangDisplayName = GetLangNameByCultureCode(.LangDisplayCode)
    
        .LangSystemCode = GetSystemDefaultUILanguage Mod &H10000
        .LangSystemName = GetLangNameByCultureCode(.LangSystemCode)
    
        .LangNonUnicodeCode = GetSystemDefaultLCID Mod &H10000
        .LangNonUnicodeName = GetLangNameByCultureCode(.LangNonUnicodeCode)
    End With
End Sub

Public Sub LoadLanguage(LCode As Long, Force As Boolean)
    Dim IsSlavian As Boolean
    
    LoadLanguageSettings
    
    'If the language for programs that do not support Unicode controls set such
    'that does not contain Cyrillic, we need to use the English localization
    IsSlavian = IsSlavianCultureCode(OSVer.LangNonUnicodeCode)
    
    ' Force choosing of language: no checks for non-Unicode language settings
    If Force Then
        Select Case LCode
        Case &H419&, &H422&, &H423& 'Russian, Ukrainian, Belarusian
            LangRU
        Case &H409& 'English
            LoadDefaultLanguage
        Case Else
            LoadDefaultLanguage
        End Select
    Else
        Select Case LCode 'OSVer.LangDisplayCode
        Case &H419&, &H422&, &H423& 'Russian, Ukrainian, Belarusian
            If IsSlavian Then
                LangRU
            Else
                If Not bAutoLog Then MsgBoxW "Cannot set Russian language!" & vbCrLf & _
                    "First, you must set language for non-Unicode programs to Russian" & vbCrLf & _
                    "trought the Control panel -> system language settings.", vbCritical
                LoadDefaultLanguage
            End If
        Case &H409& 'English
            LoadDefaultLanguage
        Case Else
            LoadDefaultLanguage
        End Select
    End If
End Sub

Function IsSlavianCultureCode(CultureCode As Long) As Boolean ' languages with Cyrillic alphabet
    Select Case CultureCode
        Case &H419&, &H422&, &H423&, &H402&
            IsSlavianCultureCode = True
    End Select
End Function

Public Function IsRussianLangCode(CultureCode As Long) As Boolean
    Select Case CultureCode
        Case &H419&, &H422&, &H423&
            IsRussianLangCode = True
    End Select
End Function

Function GetLangNameByCultureCode(LCID As Long) As String
    Dim buf As String
    Dim lr  As Long
    buf = Space$(1000)
    lr = GetLocaleInfo(LCID, LOCALE_SENGLANGUAGE, StrPtr(buf), ByVal 1000&)
    If lr Then
        GetLangNameByCultureCode = Left$(buf, lr - 1)
    End If
End Function

Public Sub LoadLanguageFile(sFile$, Optional bSilent As Boolean = False)
    Dim i&, j&, ff%
    If sFile = vbNullString Then Exit Sub
    If Not FileExists(BuildPath(AppPath(), sFile)) Then Exit Sub
    ff = FreeFile()
    Open BuildPath(AppPath(), sFile) For Input As #ff
        sLines = Split(Input(FileLen(BuildPath(AppPath(), sFile)), #ff), vbCrLf)
    Close #ff
    
    If bSilent = True Then bDontPrompt = True
    
    For i = 0 To UBound(sLines)
        For j = 0 To UBound(sLines)
            If Len(sLines(i)) >= 3 And Len(sLines(j)) >= 3 Then
                If Val(Left(sLines(i), 3)) > 0 And i <> j Then
                    If Left(sLines(i), 3) = Left(sLines(j), 3) Then
                        MsgBoxW "The language file '" & sFile & "' " & _
                               "is invalid (ambiguous id numbers)." & _
                               vbCrLf & vbCrLf & _
                               sLines(i) & vbCrLf & sLines(j), vbExclamation
                        Exit Sub
                    End If
                End If
            End If
        Next j
    Next i
    
    ReloadLanguage
End Sub

Private Sub ReloadLanguage()
    Dim i%, Translation$
    On Error GoTo Error:
    
    With frmMain
        For i = 0 To UBound(sLines)
            If Len(sLines(i)) >= 3 Then
                sLines(i) = Replace(sLines(i), "\n", vbCrLf)
                Translation = Mid$(sLines(i), 5)
                Select Case Left$(sLines(i), 3)
                    Case "// ":
                    Case "000"
                        If Not bDontPrompt Then
                            'If MsgBoxW("Load file for language '" & Mid(sLines(i), 5) & "'?", vbYesNo + vbQuestion) = vbNo Then
                            '    Exit Sub
                            'End If
                        End If
                        bDontPrompt = False
                    
                    Case "001": .lblInfo(0).Caption = Translation
                    Case "004": .lblInfo(1).Caption = Translation
                    Case "009": .cmdMainMenu.Caption = Translation
                    Case "010": .fraScan.Caption = Translation
                    Case "011": .cmdScan.Caption = Translation
                    Case "012":
                    Case "013": .cmdFix.Caption = Translation
                    Case "014": .cmdInfo.Caption = Translation
                    Case "015": .fraOther.Caption = Translation
                    Case "016": .cmdHelp.Caption = Translation
                    Case "017":
                    Case "018": If Not .fraConfig.Visible Then .cmdConfig.Caption = Translation
                    Case "019": If .fraConfig.Visible Then .cmdConfig.Caption = Translation
                    Case "020": .cmdSaveDef.Caption = Translation
                    
                    Case "030": .fraHelp.Caption = Translation
                    
                    Case "040": .fraConfig.Caption = Translation
                    Case "041": .chkConfigTabs(0).Caption = Translation
                    Case "042": .chkConfigTabs(1).Caption = Translation
                    Case "043": .chkConfigTabs(2).Caption = Translation
                    Case "044": .chkConfigTabs(3).Caption = Translation
                    
                    Case "050": .chkAutoMark.Caption = Translation
                    Case "051": .chkBackup.Caption = Translation
                    Case "052": .chkConfirm.Caption = Translation
                    Case "053": .chkIgnoreSafe.Caption = Translation
                    Case "054": .chkLogProcesses.Caption = Translation
                    Case "055": .chkShowN00bFrame.Caption = Translation
                    Case "056": .chkConfigStartupScan.Caption = Translation
                    
                    Case "060": .lblConfigInfo(3).Caption = Translation
                    Case "061": .lblConfigInfo(0).Caption = Translation
                    Case "062": .lblConfigInfo(1).Caption = Translation
                    Case "063": .lblConfigInfo(2).Caption = Translation
                    Case "064": .lblConfigInfo(4).Caption = Translation
                    
                    Case "070": .lblConfigInfo(5).Caption = Translation
                    Case "071": .cmdConfigIgnoreDelSel.Caption = Translation
                    Case "072": .cmdConfigIgnoreDelAll.Caption = Translation
                    
                    Case "080": .lblConfigInfo(6).Caption = Translation
                    Case "081": .cmdConfigBackupRestore.Caption = Translation
                    Case "082": .cmdConfigBackupDelete.Caption = Translation
                    Case "083": .cmdConfigBackupDeleteAll.Caption = Translation
                    
                    Case "090": .lblConfigInfo(7).Caption = Translation
                    Case "091": .cmdStartupList.Caption = Translation
                    Case "092": .chkStartupListFull.Caption = Translation
                    Case "093": .chkStartupListComplete.Caption = Translation
                    
                    Case "100": .lblConfigInfo(16).Caption = Translation
                    Case "101": .cmdProcessManager.Caption = Translation
                    Case "102": .lblConfigInfo(12).Caption = Translation
                    Case "103": .cmdHostsManager.Caption = Translation
                    Case "104": .lblConfigInfo(13).Caption = Translation
                    Case "105": .cmdDelOnReboot.Caption = Translation
                    Case "106": .lblInfo(2).Caption = Translation
                    Case "107": .cmdDeleteService.Caption = Translation
                    Case "108": .lblInfo(6).Caption = Translation
                    'Case "109": .lblInfo(3).Caption = Translation
                    'Case "110": .cmdARSMan.Caption = Translation
                    Case "111": .lblInfo(7).Caption = Translation
                    
                    Case "120": .lblConfigInfo(17).Caption = Translation
                    Case "121": .chkDoMD5.Caption = Translation
                    Case "122": .chkAdvLogEnvVars.Caption = Translation
                    
                    Case "130": .lblConfigInfo(22).Caption = Translation
                    Case "131": .cmdLangLoad.Caption = Translation
                    Case "132": .cmdLangReset.Caption = Translation
                    
                    Case "140": .lblConfigInfo(18).Caption = Translation
                    Case "141": .cmdCheckUpdate.Caption = Translation
                    'Case "142": .lblConfigInfo(10).Caption = Translation
                    Case "143": .lblConfigInfo(11).Caption = Translation
                    
                    ''Case "150": .lblConfigInfo(20).Caption = Translation
                    ''Case "151": .cmdUninstall.Caption = Translation
                    ''Case "152": .lblConfigInfo(9).Caption = Translation
                    
                    Case "160": .fraN00b.Caption = Translation
                    Case "161": .lblInfo(4).Caption = Translation
                    Case "162": .cmdN00bLog.Caption = Translation
                    Case "163": .cmdN00bScan.Caption = Translation
                    Case "164": .cmdN00bBackups.Caption = Translation
                    Case "165": .cmdN00bTools.Caption = Translation
                    Case "166": .cmdN00bHJTQuickStart.Caption = Translation
                    'Case "167": .lblInfo(5).Caption = Translation
                    Case "168": .cmdN00bClose.Caption = Translation
                    Case "169": .chkShowN00b.Caption = Translation
                    Case "183": .lblInfo(9).Caption = Translation
                    
                    Case "170": .fraProcessManager.Caption = Translation
                    Case "171": .lblConfigInfo(8).Caption = Translation
                    Case "172": .chkProcManShowDLLs.Caption = Translation
                    Case "173": .cmdProcManKill.Caption = Translation
                    Case "174": .cmdProcManRefresh.Caption = Translation
                    Case "175": .cmdProcManRun.Caption = Translation
                    Case "176": .cmdProcManBack.Caption = Translation
                    Case "177": .lblProcManDblClick.Caption = Translation
                    
                                      
                    Case "210": .fraUninstMan.Caption = Translation
                    Case "211": .lblInfo(11).Caption = Translation
                    Case "212": .lblInfo(8).Caption = Translation
                    Case "213": .lblInfo(10).Caption = Translation
                    Case "214": .cmdUninstManDelete.Caption = Translation
                    Case "215": .cmdUninstManEdit.Caption = Translation
                    Case "216": .cmdUninstManOpen.Caption = Translation
                    Case "217": .cmdUninstManRefresh.Caption = Translation
                    Case "218": .cmdUninstManSave.Caption = Translation
                    Case "219": .cmdUninstManBack.Caption = Translation
                    
                    Case "270": .fraHostsMan.Caption = Translation
                    Case "271": .lblConfigInfo(14).Caption = Translation
                    Case "272": .cmdHostsManDel.Caption = Translation
                    Case "273": .cmdHostsManToggle.Caption = Translation
                    Case "274": .cmdHostsManOpen.Caption = Translation
                    Case "275": .cmdHostsManBack.Caption = Translation
                    Case "276": .lblConfigInfo(15).Caption = Translation
                    
                    Case "999": SetCharSet CInt(Translation)
                End Select
            End If
        Next i
    End With
    Exit Sub
Error:
    If MsgBoxW("Invalid language file" & _
              IIf(Err.Number > 0, " (" & Err.Description & ") ", "") & _
              ". Reset to default (English)?", vbYesNo + vbExclamation) = vbYes Then
        LoadDefaultLanguage
        ReloadLanguage
    End If
    If inIDE Then Stop: Resume Next
End Sub

Public Sub LoadDefaultLanguage()
    Dim s$
    s = _
    "// main form" & vbCrLf & _
    "001=Welcome to HiJackThis. This program will scan your PC and generate a log file of registry and file settings commonly manipulated by malware as well as good software." & vbCrLf & _
    "002=HiJackThis is already running." & vbCrLf & _
    "003=Warning!\n\nSince HiJackThis targets browser hijacking methods instead of actual browser hijackers, entries may appear in the scan list that are not hijackers. Be careful what you delete, some system utilities can cause problems if disabled.\nFor best results, ask spyware experts for help and show them your scan log. They will advise you what to fix and what to keep.\n\nSome adware-supported programs may cease to function if the associated adware is removed." & vbCrLf & _
    "004=Below are the results of the HiJackThis scan. Be careful what you delete with the 'Fix checked' button. Scan results do not determine whether an item is bad or not. The best thing to do is to 'AnalyzeThis' and show the log file to knowledgeable folks." & vbCrLf & _
    "// main screen" & vbCrLf & _
    "009=Main Menu" & vbCrLf & _
    "010=Scan && fix stuff" & vbCrLf & _
    "011=Scan" & vbCrLf & _
    "012=Save log" & vbCrLf & _
    "013=Fix checked" & vbCrLf & _
    "014=Info on selected item..." & vbCrLf & _
    "015=Other stuff" & vbCrLf & _
    "016=Info..." & vbCrLf & _
    "017=Back" & vbCrLf & _
    "018=Config..." & vbCrLf & _
    "019=Back" & vbCrLf & _
    "020=Add checked to ignorelist" & vbCrLf & _
    "021=Nothing selected! Continue?" & vbCrLf & _
    "022=You selected to fix everything HiJackThis has found. This could mean items important to your system will be deleted and the full functionality of your system will degrade.\n\nIf you aren't sure how to use HiJackThis, you should ask for help, not blindly fix things. Many 3rd party resources are available on the web that will gladly help you with your log.\n\nAre you sure you want to fix all items in your scan results?" & vbCrLf
    s = s & "023=Fix [] selected items? This will permanently delete and/or repair what you selected" & vbCrLf & _
    "024=unless you enable backups." & vbCrLf & _
    "025=This will set HiJackThis to ignore the checked items, unless they change. Continue?" & vbCrLf & _
    "026=Write access was denied to the location you specified. Try a different location please." & vbCrLf & _
    "027=The logfile has been saved to [].\nYou can open it in a text editor like Notepad." & vbCrLf & _
    "028=Unable to list running processes" & vbCrLf & _
    "029=Running processes" & vbCrLf & _
    "// help dialog" & vbCrLf & _
    "030=Help" & vbCrLf & _
    "// config tabs" & vbCrLf & _
    "040=Configuration" & vbCrLf & _
    "041=Main" & vbCrLf & _
    "042=Ignorelist" & vbCrLf & _
    "043=Backups" & vbCrLf & _
    "044=Misc Tools" & vbCrLf & _
    "// 'main' tab" & vbCrLf & _
    "050=Mark everything found for fixing after scan" & vbCrLf & _
    "051=Make backups before fixing items" & vbCrLf & _
    "052=Confirm fixing && ignoring of items (safe mode)" & vbCrLf & _
    "053=Ignore non-standard but safe domains in IE (e.g. msn.com, microsoft.com)" & vbCrLf
    s = s & "054=Include list of running processes in logfiles" & vbCrLf & _
    "055=Do not show intro frame at startup" & vbCrLf & _
    "056=Run HiJackThis scan at startup and show it when items are found" & vbCrLf & _
    "057=Are you sure you want to enable this option? \nHiJackThis is not a 'click & fix' program. Because it targets *general* hijacking methods, false positives are a frequent occurrence.\nIf you enable this option, you might disable programs or drivers you need. However, it is highly unlikely you will break your system beyond repair. So you should only enable this option if you know what you're doing!" & vbCrLf & _
    "060=Below URLs will be used when fixing hijacked/unwanted MSIE pages:" & vbCrLf & _
    "061=Default Start Page:" & vbCrLf & _
    "062=Default Search Page:" & vbCrLf & _
    "063=Default Search Assistant:" & vbCrLf & _
    "064=Default Search Customize:" & vbCrLf & _
    "// 'ignorelist' tab" & vbCrLf & _
    "070=The following items will be ignored when scanning." & vbCrLf & _
    "071=Delete" & vbCrLf & _
    "072=Delete all" & vbCrLf & _
    "// 'backups' tab" & vbCrLf & _
    "080=This is your list of items that were backed up. You can restore them (causing HiJackThis to re-detect them unless you place them on the ignorelist) or delete them from here. (Antivirus programs may detect HiJackThis backups!)" & vbCrLf & _
    "081=Restore" & vbCrLf & _
    "082=Delete" & vbCrLf & _
    "083=Delete all" & vbCrLf & _
    "084=Are you sure you want to delete this backup?" & vbCrLf & _
    "085=Are you sure you want to delete these [] backups?" & vbCrLf
    s = s & "086=Restore this item?" & vbCrLf & _
    "087=Restore these [] items?" & vbCrLf & _
    "088=Are you sure you want to delete ALL backups?" & vbCrLf & _
    "// 'misc tools' tab" & vbCrLf & _
    "090=StartupList" & vbCrLf & _
    "091=Generate StartupList log" & vbCrLf & _
    "092=List also minor sections (full)" & vbCrLf & _
    "093=List empty sections (complete)" & vbCrLf & _
    "100=System tools" & vbCrLf & _
    "101=Process manager" & vbCrLf & _
    "102=Opens a small process manager, working much like the Task Manager." & vbCrLf & _
    "103=Hosts file manager" & vbCrLf & _
    "104=Opens an editor for the 'hosts' file." & vbCrLf & _
    "105=Delete a file on reboot..." & vbCrLf & _
    "106=If a file cannot be removed from memory, Windows can be setup to delete it when the system is restarted." & vbCrLf & _
    "107=Delete an Windows service..." & vbCrLf & _
    "108=Delete a Windows Service (O23). USE WITH CAUTION! (WinNT4/2k/XP only)" & vbCrLf & _
    "109=ADS Spy..." & vbCrLf & _
    "110=Open the integrated ADS Spy utility to scan for hidden data streams." & vbCrLf & _
    "111=Open a utility to manage the items in the Add/Remove Software list." & vbCrLf
    s = s & "113=Enter the exact service name as it appears in the scan results, or the short name between brackets if that is listed.\nThe service needs to be stopped and disabled.\nWARNING! When the service is deleted, it cannot be restored!" & vbCrLf & _
    "114=Delete a Windows NT Service" & vbCrLf & _
    "115=Service '[]' was not found in the Registry.\nMake sure you entered the name of the service correctly." & vbCrLf & _
    "116=The service '[]' is enabled and/or running. Disable it first, using HiJackThis itself (from the scan results) or the Services.msc window." & vbCrLf & _
    "117=The following service was found:" & vbCrLf & _
    "118=Are you absolutely sure you want to delete this service?" & vbCrLf & _
    "120=Advanced settings (these will not be saved)" & vbCrLf & _
    "121=Calculate MD5 of files if possible" & vbCrLf & _
    "122=Include environment variables in logfile" & vbCrLf & _
    "130=Language files" & vbCrLf & _
    "131=Load this file" & vbCrLf & _
    "132=Reset to default" & vbCrLf & _
    "140=Update check" & vbCrLf & _
    "141=Check for update online" & vbCrLf & _
    "142=" & vbCrLf & _
    "143=Use this proxy server (host:port) :" & vbCrLf & _
    "150=Uninstall HiJackThis" & vbCrLf & _
    "151=Uninstall HiJackThis && exit" & vbCrLf & _
    "152=Removes all Registry settings and exits." & vbCrLf
    s = s & "153=This will remove HiJackThis' settings from the Registry and exit. Note that you will have to delete the HiJackThis.exe file manually.\n\nContinue with uninstall?" & vbCrLf & _
    "// n00b screen" & vbCrLf & _
    "160=Main Menu" & vbCrLf & _
    "161=What would you like to do?" & vbCrLf & _
    "162=Do a system scan and save a logfile" & vbCrLf & _
    "163=Do a system scan only" & vbCrLf & _
    "164=List of Backups" & vbCrLf & _
    "165=Misc Tools" & vbCrLf & _
    "166=Online Guide" & vbCrLf & _
    "167=" & vbCrLf & _
    "168=Start the program" & vbCrLf & _
    "169=Do not show this menu again" & vbCrLf & _
    "183=Language:" & vbCrLf & _
    "// Process Manager" & vbCrLf & _
    "170=Process Manager" & vbCrLf & _
    "171=Running processes:" & vbCrLf & _
    "172=Show DLLs" & vbCrLf & _
    "173=Kill process" & vbCrLf & _
    "174=Refresh" & vbCrLf & _
    "175=Run.." & vbCrLf & _
    "176=Back" & vbCrLf
    s = s & "177=Double-click a file to view its properties" & vbCrLf & _
    "178=Loaded DLL libraries by selected process:" & vbCrLf & _
    "179=Are you sure you want to close the selected processes?" & vbCrLf & _
    "180=Any unsaved data in it will be lost." & vbCrLf & _
    "181=Run" & vbCrLf & _
    "182=Type the name of a program, folder, document or Internet resource, and Windows will open it for you." & vbCrLf & _
    "// Hosts file manager" & vbCrLf & _
    "270=Hosts file manager" & vbCrLf & _
    "271=Hosts file is located at:" & vbCrLf & _
    "272=Delete line(s)" & vbCrLf & _
    "273=Toggle line(s)" & vbCrLf & _
    "274=Open in Notepad" & vbCrLf & _
    "275=Back" & vbCrLf & _
    "276=Note: changes to the hosts file take effect when you restart your browser." & vbCrLf & _
    "// ADS Spy" & vbCrLf & _
    "190=ADS Spy" & vbCrLf & _
    "191=Quick scan (Windows base folder only)" & vbCrLf & _
    "192=Ignore safe system info streams" & vbCrLf & _
    "193=Calculate MD5 checksum of streams" & vbCrLf & _
    "194=What's this?" & vbCrLf
    s = s & "195=Help" & vbCrLf & _
    "196=Scan" & vbCrLf & _
    "197=Save log..." & vbCrLf & _
    "198=Remove selected" & vbCrLf & _
    "199=Back" & vbCrLf & _
    "200=Ready." & vbCrLf & _
    "201=Using ADS Spy is very easy: just click 'Scan', wait until the scan completes, then select the ADS streams you want to remove and click 'Remove selected'. If you are unsure which streams to remove, ask someone for help. Don't delete streams if you don't know what they are!\n\nThe three checkboxes are:\n\nQuick Scan: only scans the Windows folder. So far all known malware that uses ADS to hide itself, hides in the Windows folder. Unchecking this will make ADS Spy scan the entire system (i.e. all drives).\n\nIgnore safe system info streams: Windows, Internet Explorer and a few antivirus programs use ADS to store metadata for certain folders and files. These streams can safely be ignored, they are harmless.\n\nCalculate MD5 checksums of streams: For antispyware program development or antivirus analysis only.\n\nNote: the default settings of above three checkboxes should be fine for most people. There's no need to change any of them unless you are a developer or anti-malware expert." & vbCrLf & _
    "202=Abort" & vbCrLf & _
    "203=Scan aborted!" & vbCrLf & _
    "204=Scan complete." & vbCrLf & _
    "205=Alternate Data Streams (ADSs) are pieces of info hidden as metadata on files. They are not visible in Explorer and the size they take up is not reported by Windows. Recent browser hijackers started hiding their files inside ADSs, and very few anti-malware scanners detect this (yet).\nUse ADS Spy to find and remove these streams.\nNote: this app also displays legitimate ADS streams. Do not delete streams if you are not completely sure they are malicious!" & vbCrLf & _
    "// Uninstall Manager" & vbCrLf & _
    "210=Add/Remove Programs Manager" & vbCrLf & _
    "211=Here you can see the list of programs in the Add/Remove Software list in the Control Panel. You can edit the uninstall command or delete an item completely. Beware, restoring a deleted item is not possible!" & vbCrLf & _
    "212=Name:" & vbCrLf & _
    "213=Uninstall command:" & vbCrLf & _
    "214=Delete this entry" & vbCrLf & _
    "215=Edit uninstall command" & vbCrLf & _
    "216=Open Add/Remove Software list" & vbCrLf & _
    "217=Refresh list" & vbCrLf
    s = s & "218=Save list..." & vbCrLf & _
    "219=Back" & vbCrLf & _
    "220=Are you sure you want to delete this item from the list?" & vbCrLf & _
    "221=Enter the new uninstall command for this program" & vbCrLf & _
    "222=New uninstall string saved!" & vbCrLf & _
    "223=Are you sure, you want to uninstall " & vbCrLf & _
    "// Status during scan" & vbCrLf & _
    "230=IE Registry values" & vbCrLf & _
    "231=INI file values" & vbCrLf & _
    "232=Netscape/Mozilla homepage" & vbCrLf & _
    "233=O1 - Hosts file redirection" & vbCrLf & _
    "234=O2 - BHO enumeration" & vbCrLf & _
    "235=O3 - Toolbar enumeration" & vbCrLf & _
    "236=O4 - Registry && Start Menu autoruns" & vbCrLf & _
    "237=O5 - Control Panel: IE Options" & vbCrLf & _
    "238=O6 - Policies: IE Options menuitem" & vbCrLf & _
    "239=O7 - Policies: Regedit" & vbCrLf & _
    "240=O8 - IE contextmenu items enumeration" & vbCrLf & _
    "241=O9 - IE 'Tools' menuitems and buttons enumeration" & vbCrLf & _
    "242=O10 - Winsock LSP hijackers" & vbCrLf & _
    "243=O11 - Extra options groups in IE Advanced Options" & vbCrLf
    s = s & "244=O12 - IE plugins enumeration" & vbCrLf & _
    "245=O13 - DefaultPrefix hijack" & vbCrLf & _
    "246=O14 - IERESET.INF hijack" & vbCrLf & _
    "247=O15 - Trusted Zone enumeration" & vbCrLf & _
    "248=O16 - DPF object enumeration" & vbCrLf & _
    "249=O17 - DNS && DNS Suffix settings" & vbCrLf & _
    "250=O18 - Protocol && Filter enumeration" & vbCrLf & _
    "251=O19 - User stylesheet hijack" & vbCrLf & _
    "252=O20 - AppInit_DLLs Registry value" & vbCrLf & _
    "253=O21 - ShellServiceObjectDelayLoad Registry key" & vbCrLf & _
    "254=O22 - SharedTaskScheduler Registry key" & vbCrLf & _
    "255=O23 - NT Services" & vbCrLf & _
    "256=Scan completed!" & vbCrLf & _
    "257=O24 - ActiveX Desktop Components" & vbCrLf & _
    "// msgboxW'es and stuff in scan methods" & vbCrLf & _
    "300=For some reason your system denied write access to the Hosts file. If any hijacked domains are in this file, HiJackThis may NOT be able to fix this.\n\nIf that happens, you need to edit the file yourself. To do this, click Start, Run and type:\n\n   notepad []\n\nand press Enter. Find the line(s) HiJackThis reports and delete them. Save the file as 'hosts.' (with quotes), and reboot.\n\nFor Vista and above: simply, exit HiJackThis, right click on the HiJackThis icon, choose 'Run as administrator'." & vbCrLf & _
    "301=Your hosts file has invalid linebreaks and HiJackThis is unable to fix this. O1 items will not be displayed.\n\nClick OK to continue the rest of the scan." & vbCrLf & _
    "302=You have an particularly large amount of hijacked domains. It's probably better to delete the file itself then to fix each item (and create a backup).\n\nIf you see the same IP address in all the reported O1 items, consider deleting your Hosts file, which is located at []." & vbCrLf & _
    "303=HiJackThis could not write the selected changes to your hosts file. The probably cause is that some program is denying access to it, or that your user account doesn't have the rights to write to it." & vbCrLf & _
    "310=HiJackThis is about to remove a BHO and the corresponding file from your system. Close all Internet Explorer windows AND all Windows Explorer windows before continuing for the best chance of success." & vbCrLf & _
    "320=Unable to delete the file:\n[]\n\nThe file may be in use." & vbCrLf
    s = s & "321=Use Task Manager to shutdown the program" & vbCrLf & _
    "322=Use a process killer like ProcView to shutdown the program" & vbCrLf & _
    "323=and run HiJackThis again to delete the file." & vbCrLf & _
    "330=HiJackThis is about to remove a plugin from your system. Close all Internet Explorer windows before continuing for the best chance of success." & vbCrLf & _
    "340=It looks like you're running HiJackThis from a read-only device like a CD or locked floppy disk.If you want to make backups of items you fix, you must copy HiJackThis.exe to your hard disk first, and run it from there.\n\nIf you continue, you might get 'Path/File Access' errors." & vbCrLf & _
    "341=HiJackThis appears to have been started from a temporary folder. Since temp folders tend to be be emptied regularly, it's wise to copy HiJackThis.exe to a folder of its own, for instance C:\Program Files\HiJackThis.\nThis way, any backups that will be made of fixed items won't be lost.\n\nPlease quit HiJackThis and copy it to a separate folder first before fixing any items." & vbCrLf & _
    "342=The file '[]' will be deleted by Windows when the system restarts." & vbCrLf & _
    "343=The service '[]'  has been marked for deletion." & vbCrLf & _
    "344=Unable to delete the service '[]'. Make sure the name is correct and the service is not running." & vbCrLf & _
    "// = misc stuff =" & vbCrLf & _
    "// HJT quickstart URL" & vbCrLf & _
    "http://sourceforge.net/projects/hjt/files/2.0.4/"
    
    sLines = Split(s, vbCrLf)
    ReloadLanguage
End Sub

Public Sub LangRU()
    Dim s$

    s = _
    "// основной вид" & vbCrLf & _
    "001=Добро пожаловать в HiJackThis. Эта программа сканирует  ПК и создает отчет по файлам и настроек реестра, которые обычно используются зловредами." & vbCrLf & _
    "002=HiJackThis уже запущен." & vbCrLf & _
    "003=Предупреждение!\n\nПоскольку целью HiJackThis является исследование методов вмешательства вредоносного ПО в работу браузеров вместо поиска конкретных видов Hijacker, результаты сканирования могут содержать легитимные записи, не Hijacker. Будьте осторожны, когда что-то удаляете. В части системных программ могут возникнуть неполадки. Лучше всего, обратиться за помощью к экспертам и показать им свой отчет. Они смогут посоветовать, что именно нужно исправлять. При удалении рекламного ПО, некоторые легитимные программы, работающие в связке с ним, могут перестать функционировать." & vbCrLf & _
    "004=Ниже приведены результаты сканирования HiJackThis. Будьте осторожны при совершении действий при помощи кнопки 'Исправить отмеченное'. В результате сканирования легитимность или зловредность элементов не определяется. Для определения нажмите 'АнализироватьЭто' и показать файл-отчет знающим людям." & vbCrLf & _
    "// main screen" & vbCrLf & _
    "010=Сканировать и исправить" & vbCrLf & _
    "011=Сканировать" & vbCrLf & _
    "012=Сохранить отчет" & vbCrLf & _
    "013=Исправить отмеченное" & vbCrLf & _
    "014=Сведения о выбранном элементе..." & vbCrLf & _
    "015=Другие данные" & vbCrLf & _
    "016=Сведения..." & vbCrLf & _
    "017=Назад" & vbCrLf & _
    "018=Конфигурация..." & vbCrLf & _
    "019=Назад" & vbCrLf & _
    "020=Добавить отмеченное в игнор-лист" & vbCrLf & _
    "021=Ничего не выбрано! Продолжить?" & vbCrLf & _
    "022=Вы выбрали для исправления все, что было найдено программой HiJackThis. В списке могут быть важные элементы системы, удаление которых, приведет к ухудшению функциональности ПК.\n\nЕсли Вы не знаете, как пользоваться программой HiJackThis, Вам следует обратиться за помощью, чем производить неосознанные действия. В Интернете достаточно сайтов, на которых Вам  смогут оказать помощь с отчетом.\n\nВы уверены, что хотите исправить все элементы, найденные в ходе сканирования?" & vbCrLf
    s = s & "023=Исправить [] выбранные элементы? Данное действие приведет к удалению или исправлению элементов" & vbCrLf & _
    "024=если Вы не включите резервное копирование." & vbCrLf & _
    "025=Настроить HijsckThis на игнорирование отмеченных элементов, в данном случае они не будут изменены. Продолжить?" & vbCrLf & _
    "026=Нет доступа на запись в указанном месте. Пожалуйста, попробуйте другое расположение." & vbCrLf & _
    "027=Файл-отчета сохранен в [].\nВы можете открыть его в текстовом редакторе, подобному Notepad." & vbCrLf & _
    "028=Невозможно отобразить запущенные процессы" & vbCrLf & _
    "029=Запущенные процессы" & vbCrLf & _
    "// справка" & vbCrLf & _
    "030=Помощь" & vbCrLf & _
    "// вкладки конфигурации" & vbCrLf & _
    "040=Конфигурация" & vbCrLf & _
    "041=Главное" & vbCrLf & _
    "042=Игнор-лист" & vbCrLf & _
    "043=Резервные копии" & vbCrLf & _
    "044=Инструменты" & vbCrLf & _
    "// 'главная' вкладка" & vbCrLf & _
    "050=Отметить все найденное для исправления после сканирования" & vbCrLf & _
    "051=Сделать резервные копии до исправления пунктов" & vbCrLf & _
    "052=Подтвердить исправление и игнорирование пунктов (безопасный режим)" & vbCrLf & _
    "053=Игнорировать не стандартные но доверенные домены в IE (e.g. msn.com, microsoft.com)" & vbCrLf
    s = s & "054=Внести список запущенных процессов в файл-отчет" & vbCrLf & _
    "055=Не показывать начальное окно при запуске" & vbCrLf & _
    "056=Начать сканирование системы при ее запуске и вывод результата после сканирования" & vbCrLf & _
    "057=Вы уверены, что хотите включить эту опцию? \nHiJackThis не является программой 'нажми-исправь'. Потому, что он нацелен на «общие» методы угона, и ложные срабатывания являются частыми явлениями.\nВключение данной опции, может отключить необходимые программы и драйвера. Тем не менее, маловероятно, что вы можете повредить систему так, чтоб ее нельзя было восстановить. Таким образом, Вам следует включить только эту опцию, если Вы знаете, что делаете!" & vbCrLf & _
    "060=Снизу интернет-адреса, использующиеся при исправлении угонщиков браузера(нежелательные страницы IE):" & vbCrLf & _
    "061=Начальная страница по умолчанию:" & vbCrLf & _
    "062=Поисковая страница по умолчанию:" & vbCrLf & _
    "063=Поисковый ассистент по умолчанию:" & vbCrLf & _
    "064=Поисковый настройщик по умолчанию:" & vbCrLf & _
    "// 'игнор-лист' вкладка" & vbCrLf & _
    "070=Следующие элементы будут игнорироваться при сканировании." & vbCrLf & _
    "071=Удалить" & vbCrLf & _
    "072=Удалить все" & vbCrLf & _
    "// 'Резервные копии' вкладка" & vbCrLf & _
    "080=Список элементов резервных копий. Вы можете восстановить их (вызвав HiJackThis обнаружить их снова, если Вы  предварительно произвели резервное копирование) или удалить отсюда. (Антивирусные программы могут обнаружить элементы в резервных копиях!)" & vbCrLf & _
    "081=Восстановление" & vbCrLf & _
    "082=Удалить" & vbCrLf & _
    "083=Удалить все" & vbCrLf & _
    "084=Вы уверены, что хотите удалить эту резервную копию?" & vbCrLf & _
    "085=Вы уверены, что хотите удалить эти [] резервные копии?" & vbCrLf
    s = s & "086=Восстановить этот элемент?" & vbCrLf & _
    "087=Восстановить эти [] элементы?" & vbCrLf & _
    "088=Вы уверены, что хотите удалить все резервные копии?" & vbCrLf & _
    "// 'инструменты' вкладка" & vbCrLf & _
    "090=Список автозагрузки (интегрированна: v1.52)" & vbCrLf & _
    "091=Создать отчет списка автозагрузки" & vbCrLf & _
    "092=Перечислить также список незначительных секций(полный)" & vbCrLf & _
    "093=Перечислить пустые секции (полностью)" & vbCrLf & _
    "100=Системные инструменты" & vbCrLf & _
    "101=Открыть менеджер процессов" & vbCrLf & _
    "102=Открыть профильный менеджер процессов, подобный диспетчеру задач." & vbCrLf & _
    "103=Открыть менеджер hosts файла" & vbCrLf & _
    "104=Открыть редактор hosts файла" & vbCrLf & _
    "105=Удалить файл после перезагрузки..." & vbCrLf & _
    "106=Если файл не может быть удален из памяти, Windows может быть настроен так, чтоб удалить его после перезагрузки." & vbCrLf & _
    "107=Удалить службу Windows.." & vbCrLf & _
    "108=Удалить службу Windows (O23). ИСПОЛЬЗОВАТЬ С ОСТОРОЖНОСТЬЮ! (WinNT4/2k/XP только)" & vbCrLf & _
    "109=Открыть ADS Spy..." & vbCrLf & _
    "110=Открыть интегрированную утилиту ADS Spy для сканирования данных скрытых потоков." & vbCrLf & _
    "111=Открыть менеджер удаления..." & vbCrLf
    s = s & "112=Открыть оснастку Установка и удаление программ ." & vbCrLf & _
    "113=Введите имя службы так, как оно отображается в результатах проверки или короткое имя в скобках, если оно указано.\nСлужба должна быть остановлена и отключена.\nВНИМАНИЕ! Если служба удалена, то ее нельзя восстановить!" & vbCrLf & _
    "114=Удалить службу Windows NT" & vbCrLf & _
    "115=Служба '[]' не найдена в Реестре.\nУбедитесь, что Вы ввели имя службы правильно." & vbCrLf & _
    "116=Служба '[]' включена и/или выполняется. Отключите ее, при помощи HiJackThis (из результата сканирования) или через консоль управления службами." & vbCrLf & _
    "117=Следующая служба была найдена:" & vbCrLf & _
    "118=Вы абсолютно уверены, что хотите удалить эту службу?" & vbCrLf & _
    "120=Дополнительные настройки (данные не будут сохранены)" & vbCrLf & _
    "121=Вычислить MD5 файлов, если это возможно" & vbCrLf & _
    "122=Включить переменные среды в файл-отчет" & vbCrLf & _
    "130=Языковые файлы" & vbCrLf & _
    "131=Загрузить этот файл" & vbCrLf & _
    "132=Восстановить значения по умолчанию" & vbCrLf & _
    "140=Проверить обновления" & vbCrLf & _
    "141=Проверить обновления онлайн" & vbCrLf & _
    "142=" & vbCrLf & _
    "143=Использовать этот прокси-сервер (хост:порт) :" & vbCrLf & _
    "150=Удалить HiJackThis" & vbCrLf & _
    "151=Выйти и удалить HiJackThis" & vbCrLf & _
    "152=Удалить все записи Реестра и выйти." & vbCrLf
    s = s & "153=Удалить настройки HiJackThis в Реестре и выйти. Обратите внимание, Вам придется удалить файл HiJackThis.exe вручную.\n\nПродолжить удаление?" & vbCrLf & _
    "// n00b экран" & vbCrLf & _
    "160=Главное Меню" & vbCrLf & _
    "161=Что бы Вы хотели сделать?" & vbCrLf & _
    "162=Проверить систему и сохранить отчет" & vbCrLf & _
    "163=Только проверить систему" & vbCrLf & _
    "164=Резервные копии" & vbCrLf & _
    "165=Инструменты" & vbCrLf & _
    "166=Онлайн руководство" & vbCrLf & _
    "167=" & vbCrLf & _
    "168=Запустить программу" & vbCrLf & _
    "169=Больше не показывать это окно при запуске HiJackThis" & vbCrLf & _
    "183=Изменить язык:" & vbCrLf & _
    "// Менеджер процессов" & vbCrLf & _
    "170=Менеджер процессов" & vbCrLf & _
    "171=Запущенные процессы:" & vbCrLf & _
    "172=Показать библиотеки DDL" & vbCrLf & _
    "173=Завершить процесс" & vbCrLf & _
    "174=Обновить" & vbCrLf & _
    "175=Запустить.." & vbCrLf & _
    "176=Назад" & vbCrLf
    s = s & "177=Дважды щелкните по файлу, чтоб просмотреть свойства" & vbCrLf & _
    "178=Загруженные DLL библиотеки выбранного процесса:" & vbCrLf & _
    "179=Вы уверены, что хотите закрыть выбранные процессы?" & vbCrLf & _
    "180=Все не сохраненные данные будут потеряны." & vbCrLf & _
    "181=Запустить" & vbCrLf & _
    "182=Введите имя программы, папки, документа или ресурса Интернета, и Windows откроет его." & vbCrLf & _
    "// Менеджер hosts файла" & vbCrLf & _
    "270=Менеджер hosts файла" & vbCrLf & _
    "271=Hosts файл расположен в:" & vbCrLf & _
    "272=Удалить строку" & vbCrLf & _
    "273=Переключить строку" & vbCrLf & _
    "274=Открыть в Notepad" & vbCrLf & _
    "275=Назад" & vbCrLf & _
    "276=Примечание: изменения в файле hosts, вступят в силу после перезагрузки браузера." & vbCrLf & _
    "// ADS Spy" & vbCrLf & _
    "190=ADS Spy" & vbCrLf & _
    "191=Быстрое сканирование (только папки Windows)" & vbCrLf & _
    "192=Игнорировать сведения системы безопасности потоков" & vbCrLf & _
    "193=Вычислить MD5 потоков" & vbCrLf & _
    "194=Что это?" & vbCrLf
    s = s & "195=Помощь" & vbCrLf & _
    "196=Сканировать" & vbCrLf & _
    "197=Сохранить файл-отчет..." & vbCrLf & _
    "198=Удалить выбранное" & vbCrLf & _
    "199=Назад" & vbCrLf & _
    "200=Готово." & vbCrLf & _
    "201=Использовать ADS Spy очень легко: просто нажмите 'Сканировать', подождите пока сканирование завершится, затем выберите ADS потоки, которые Вы хотите удалить и нажмите 'Удалить выбранное'. Если Вы не знаете, какие потоки удалять, попросите кого-то помочь Вам. Не удаляйте потоки, если Вы не знаете их назначение!\n\nТри флажка:\n\nБыстрое сканирование: сканирует только папку Windows. До сих пор, все известные зловреды использующие ADS для собственного сокрытия, прячутся в папке Windows. Если не выбрать этот пункт, то ADS Spy будет сканировать всю систему в том числе все диски.\n\nИгнорирование сведений системы безопасности потоков: Windows, Internet Explorer и немногие антивирусные программы используют ADS для хранения метаданных определенных файлов и папок. Эти потоки можно игнорировать, они безвредны.\n\nВычислить MD5 потоков: для развития антишпионских программ или антивирусного анализа.\n\n" & _
    "Примечание: настройки по умолчанию из трех флажков выше, должны подойти для большинства пользователей. Нет никакой необходимости изменять их, если Вы не разработчик или антивирусный аналитик." & vbCrLf & _
    "202=Прервать" & vbCrLf & _
    "203=Сканирование прервано!" & vbCrLf & _
    "204=Сканирование выполнено." & vbCrLf & _
    "205=Альтернативные потоки данных (ADS) - это фрагменты информации в виде скрытых метаданных. Они не видны в проводнике и занимаемый ими размер на диске, система не показывает. Последние угонщики браузеров стали прятать свои файлы в ADS, и очень мало антивирусных программ обнаруживают их.\nИспользуйте ADS Spy чтобы найти и удалить эти потоки.\nПримечание: это приложение отображает также легитимные ADS. Не удаляйте потоки, если Вы не полностью уверены в их зловредности!" & vbCrLf & _
    "// Менеджер удаления" & vbCrLf & _
    "210=Менеджер Установки/удаления программ" & vbCrLf & _
    "211=Здесь, Вы можете увидеть список программ находящиеся в оснастке Установка и удаление программ Панели управления. Вы можете редактировать команду удаления, либо полностью удалить элемент. Будьте осторожны, восстановление удаленных программ не представляется возможным!" & vbCrLf & _
    "212=Название:" & vbCrLf & _
    "213=Команда удаления:" & vbCrLf & _
    "214=Удалить запись" & vbCrLf & _
    "215=Редактировать команду удаления" & vbCrLf & _
    "216=Открыть менеджер Установки и удаления программ" & vbCrLf & _
    "217=Обновить список" & vbCrLf
    s = s & "218=Сохранить список..." & vbCrLf & _
    "219=Назад" & vbCrLf & _
    "220=Вы уверены, что хотите удалить этот элемент из списка?" & vbCrLf & _
    "221=Введите новую команду для удаления этой программы" & vbCrLf & _
    "222=Новая строка удаления сохранена!" & vbCrLf & _
    "// Статус во время сканирования" & vbCrLf & _
    "230=Значения реестра IE" & vbCrLf & _
    "231=Значения INI файлов" & vbCrLf & _
    "232=Домашняя страница Netscape/Mozilla" & vbCrLf & _
    "233=O1 - Редирект на hosts файл" & vbCrLf & _
    "234=O2 - Перечень BHO" & vbCrLf & _
    "235=O3 - Перечень панели инструментов" & vbCrLf & _
    "236=O4 - Реестр и автозапуск" & vbCrLf & _
    "237=O5 - Панель управления: свойства обозревателя" & vbCrLf & _
    "238=O6 - Политики: свойства обозревателя, " & vbCrLf & _
    "239=O7 - Политики: Редактора реестра" & vbCrLf & _
    "240=O8 - Элементы контекстного меню IE" & vbCrLf & _
    "241=O9 - Меню инструменты и перечень кнопок IE" & vbCrLf & _
    "242=O10 - Winsock LSP угонщики" & vbCrLf & _
    "243=O11 - Дополнительные настройки IE" & vbCrLf
    s = s & "244=O12 - Перечень плагинов IE" & vbCrLf & _
    "245=O13 - Префикс перехватчика по умолчанию" & vbCrLf & _
    "246=O14 - IERESET.INF перехватчик" & vbCrLf & _
    "247=O15 - Перечень Доверенных зон" & vbCrLf & _
    "248=O16 - Перечень объектов DPF" & vbCrLf & _
    "249=O17 - Настройки DNS и DNS-суффикса" & vbCrLf & _
    "250=O18 - Протокол и перечень фильтров" & vbCrLf & _
    "251=O19 - Пользовательский стиль перехватчика" & vbCrLf & _
    "252=O20 - Значение реестра AppInit_DLL" & vbCrLf & _
    "253=O21 - Ключ реестра Службы объекта отложенной загрузки(SSODL)" & vbCrLf & _
    "254=O22 - Ключ реестра Расширенного планировщика задач" & vbCrLf & _
    "255=O23 - NT Службы" & vbCrLf & _
    "256=Сканирование выполнено!" & vbCrLf & _
    "257=O24 - Компоненты рабочего стола ActiveX" & vbCrLf & _
    "// msgблоки и данные метода сканирования" & vbCrLf & _
    "300=По неизвестной причине, системе запрещен доступ на запись в файле hosts. HiJackThis может быть бессилен, против некоторых перехватчиков доменов, находящихся в файле hosts.\n\nЕсли так случится, то Вам нужно отредактировать файл самостоятельно. Чтобы сделать это, нажмите Пуск, Выполнить и введите:\n\n   notepad []\n\nи  нажмите Enter. Найдите нужную строку и удалите ее. Сохраните файл как 'hosts.' (с комментариями), и перезагрузите.\n\nДля операционных систем Vista и выше: закройте HiJackThis, затем правый клик по файлу HiJackThis.exe и Запуск от имени администратора'." & vbCrLf & _
    "301=Ваш hosts файл содержит недопустимые переносы строк, вследствие чего, HiJackThis не в состоянии исправить это . O1 элементы не будут отображаться.\n\nНажмите OK, чтобы продолжить сканирование." & vbCrLf & _
    "302=У Вас имеется большое количество угонщиков доменов. Вероятно, лучше удалить сам файл и создать новый, чем исправлять каждый элемент.\n\nЕсли Вы видите одинаковые IP адреса, в пункте O1, рассмотрите вариант удаления файла hosts, находящийся в []." & vbCrLf & _
    "303=HiJackThis не удается произвести изменения в файле hosts. Возможно, доступ блокируется работой какой-то программой, или ваша учетная запись не имеет соответствующих прав." & vbCrLf & _
    "310=HiJackThis собирается удалить BHO и соответствующий файл из системы. Закройте все браузеры и окна проводника, прежде чем продолжить работу." & vbCrLf & _
    "320=Не удается удалить файл:\n[]\n\nВозможно файл используется." & vbCrLf
    s = s & "321=Для завершения программы используйте Диспетчер задач" & vbCrLf & _
    "322=Используйте выключатель программ, как ProcView для завершения программ" & vbCrLf & _
    "323=и запустите HiJackThis снова, для удаления файла." & vbCrLf & _
    "330=HiJackThis собирается удалить плагин из системы. Закройте все браузеры и окна проводника, прежде чем продолжить работу." & vbCrLf & _
    "340=Похоже, Вы запустили HiJackThis с устройства, позволяющее только считывать информацию, на подобии CD диска,  или с заблокированной дискеты. Если Вы хотите сделать резервные копии элементов, которые собираетесь исправлять, Вам нужно скопировать HiJackThis.exe на жесткий диск и запустить его.\n\nЕсли Вы продолжите, то можете получить  'путь/файл доступ' ошибки." & vbCrLf & _
    "341=Похоже, HiJackThis запущен из временной папки. Поскольку, временные папки имеют тенденцию регулярно очищаться, желательно заранее скопировать HiJackThis.exe в свою папку, например, C:\Program Files\HiJackThis.\nТаким образом, резервные копии элементов не будут потеряны.\n\nПожалуйста, закройте HiJackThis и скопируйте его в отдельную папку, прежде чем исправлять какие-либо элементы." & vbCrLf & _
    "342=Файл '[]' будет удален из системы после перезагрузки." & vbCrLf & _
    "343=Служба '[]'  была помечена для удаления." & vbCrLf & _
    "344=Невозможно удалить службу '[]'. Убедитесь в корректности имени и в том, что служба не запущена." & vbCrLf & _
    "// = разные данные =" & vbCrLf & _
    "// HJT быстрый старт URL" & vbCrLf & _
    "http://sourceforge.net/projects/hjt/files/2.0.4/"
    
    sLines = Split(s, vbCrLf)
    ReloadLanguage
End Sub

Public Function Translate$(iRes%)
    Dim i%
    On Error GoTo Error:
    For i = 0 To UBound(sLines)
        If Len(sLines(i)) >= 3 Then
            If Val(Left(sLines(i), 3)) = iRes Then
                Translate = Replace(Mid(sLines(i), 5), "\n", vbCrLf)
                Exit For
            End If
        End If
    Next i
    Exit Function
    
Error:
    Dim szParam As String
    szParam = CStr(iRes%)
    ErrorMsg Err, "modLanguage_Translate", szParam
    If inIDE Then Stop: Resume Next
End Function

Public Sub SetCharSet(iCharSet%)
    'this is for multibyte languages like Japanese, Chinese, etc
    Dim objText As TextBox, objBtn As CommandButton
    Dim objList As ListBox, objLbl As Label
    On Error Resume Next
    For Each objText In frmMain
        Debug.Print objText.Name
        If Not Err Then
            objText.Font.Charset = iCharSet
            Err.Clear
        End If
    Next objText
    For Each objBtn In frmMain
        Debug.Print objBtn.Name
        If Not Err Then
            objBtn.Font.Charset = iCharSet
            Err.Clear
        End If
    Next objBtn
    For Each objList In frmMain
        Debug.Print objList.Name
        If Not Err Then
            objList.Font.Charset = iCharSet
            Err.Clear
        End If
    Next objList
    For Each objLbl In frmMain
        Debug.Print objLbl.Name
        If Not Err Then
            objLbl.Font.Charset = iCharSet
            Err.Clear
        End If
    Next objLbl
End Sub

Public Function GetHelpText() As String
    Dim Help$
   
    Help = vbCrLf & _
     "* Trend Micro HiJackThis v" & App.Major & "." & App.Minor & "." & App.Revision & " *" & vbCrLf & _
     vbCrLf & vbCrLf & "See bottom for version history." & vbCrLf & vbCrLf

    Help = Help & "The different sections of hijacking " & _
     "possibilities have been separated into the following groups." & vbCrLf & _
     "You can get more detailed information about an item " & _
     "by selecting it from the list of found items OR " & _
     "highlighting the relevant line below, and clicking " & _
     "'Info on selected item'." & vbCrLf & vbCrLf & _
     " R - Registry, StartPage/SearchPage changes" & vbCrLf & _
     "    R0 - Changed registry value" & vbCrLf & _
     "    R1 - Created registry value" & vbCrLf & _
     "    R2 - Created registry key" & vbCrLf & _
     "    R3 - Created extra registry value where only one should be" & vbCrLf & _
     " F - IniFiles, autoloading entries" & vbCrLf & _
     "    F0 - Changed inifile value" & vbCrLf & _
     "    F1 - Created inifile value" & vbCrLf & _
     "    F2 - Changed inifile value, mapped to Registry" & vbCrLf & _
     "    F3 - Created inifile value, mapped to Registry" & vbCrLf
    Help = Help & _
     " N - Netscape/Mozilla StartPage/SearchPage changes" & vbCrLf & _
     "    N1 - Change in prefs.js of Netscape 4.x" & vbCrLf & _
     "    N2 - Change in prefs.js of Netscape 6" & vbCrLf & _
     "    N3 - Change in prefs.js of Netscape 7" & vbCrLf & _
     "    N4 - Change in prefs.js of Mozilla" & vbCrLf & _
     " O - Other, several sections which represent:" & vbCrLf & _
     "    O1 - Hijack of auto.search.msn.com with Hosts file" & vbCrLf & _
     "    O2 - Enumeration of existing MSIE BHO's" & vbCrLf & _
     "    O3 - Enumeration of existing MSIE toolbars" & vbCrLf & _
     "    O4 - Enumeration of suspicious autoloading Registry entries" & vbCrLf & _
     "    O5 - Blocking of loading Internet Options in Control Panel" & vbCrLf & _
     "    O6 - Disabling of 'Internet Options' Main tab with Policies" & vbCrLf & _
     "    O7 - Disabling of Regedit with Policies" & vbCrLf & _
     "    O8 - Extra MSIE context menu items" & vbCrLf
    Help = Help & _
     "    O9 - Extra 'Tools' menuitems and buttons" & vbCrLf & _
     "    O10 - Breaking of Internet access by New.Net or WebHancer" & vbCrLf & _
     "    O11 - Extra options in MSIE 'Advanced' settings tab" & vbCrLf & _
     "    O12 - MSIE plugins for file extensions or MIME types" & vbCrLf & _
     "    O13 - Hijack of default URL prefixes" & vbCrLf & _
     "    O14 - Changing of IERESET.INF" & vbCrLf & _
     "    O15 - Trusted Zone Autoadd" & vbCrLf & _
     "    O16 - Download Program Files item" & vbCrLf & _
     "    O17 - Domain hijack" & vbCrLf & _
     "    O18 - Enumeration of existing protocols and filters" & vbCrLf & _
     "    O19 - User stylesheet hijack" & vbCrLf & _
     "    O20 - AppInit_DLLs autorun Registry value, Winlogon Notify Registry keys" & vbCrLf & _
     "    O21 - ShellServiceObjectDelayLoad (SSODL) autorun Registry key" & vbCrLf & _
     "    O22 - SharedTaskScheduler autorun Registry key" & vbCrLf & _
     "    O23 - Enumeration of Windows Services" & vbCrLf & _
     "    O24 - Enumeration of ActiveX Desktop Components" & vbCrLf & vbCrLf
     
    Help = Help & _
     "Command-line parameters:" & vbCrLf & _
     "* /autolog - automatically scan the system, save a logfile and open it" & vbCrLf & _
     "* /ihatewhitelists - ignore all internal whitelists" & vbCrLf & _
     "* /uninstall - remove all HiJackThis Registry entries, backups and quit" & vbCrLf & _
     "* /silentautolog - the same as /autolog, except with no required user intervention" & vbCrLf & _
     "* /startupscan - automatically scan the system (the same as button ""Do a system scan only"")" & vbCrLf & _
     "* /deleteonreboot ""c:\file.sys"" - delete the file specified after system rebooting"

    GetHelpText = Help
End Function

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
            
        Case "O23" 'Windows Services
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
    MsgBoxW sItem & vbCrLf & vbCrLf & sMsg, vbInformation
    Exit Sub
    
Error:
    ErrorMsg Err, "modInfo_GetInfo", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

