Attribute VB_Name = "modLanguage"
Option Explicit
'language module

Private sLines$(), bDontPrompt As Boolean

Public Sub LoadLanguageFile(sFile$, Optional bSilent As Boolean = False)
    Dim i%, j%
    If sFile = vbNullString Then Exit Sub
    If Not FileExists(BuildPath(App.Path, sFile)) Then Exit Sub
    Open BuildPath(App.Path, sFile) For Input As #1
        sLines = Split(Input(FileLen(BuildPath(App.Path, sFile)), #1), vbCrLf)
    Close #1
    
    If bSilent = True Then bDontPrompt = True
    
    For i = 0 To UBound(sLines)
        For j = 0 To UBound(sLines)
            If Len(sLines(i)) >= 3 And Len(sLines(j)) >= 3 Then
                If Val(Left(sLines(i), 3)) > 0 And i <> j Then
                    If Left(sLines(i), 3) = Left(sLines(j), 3) Then
                        MsgBox "The language file '" & sFile & "' " & _
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
    Dim i%
    On Error GoTo Error:
    
    With frmMain
        For i = 0 To UBound(sLines)
            If Len(sLines(i)) >= 3 Then
                sLines(i) = Replace(sLines(i), "\n", vbCrLf)
                Select Case Left(sLines(i), 3)
                    Case "// ":
                    Case "000"
                        If Not bDontPrompt Then
                            If MsgBox("Load file for language '" & Mid(sLines(i), 5) & "'?", vbYesNo + vbQuestion) = vbNo Then
                                Exit Sub
                            End If
                        End If
                        bDontPrompt = False
                    
                    Case "001": .lblInfo(0).Caption = Mid(sLines(i), 5)
                    Case "004": .lblInfo(1).Caption = Mid(sLines(i), 5)
                    Case "010": .fraScan.Caption = Mid(sLines(i), 5)
                    Case "011": .cmdScan.Caption = Mid(sLines(i), 5)
                    Case "012":
                    Case "013": .cmdFix.Caption = Mid(sLines(i), 5)
                    Case "014": .cmdInfo.Caption = Mid(sLines(i), 5)
                    Case "015": .fraOther.Caption = Mid(sLines(i), 5)
                    Case "016": .cmdHelp.Caption = Mid(sLines(i), 5)
                    Case "017":
                    Case "018": If Not .fraConfig.Visible Then .cmdConfig.Caption = Mid(sLines(i), 5)
                    Case "019": If .fraConfig.Visible Then .cmdConfig.Caption = Mid(sLines(i), 5)
                    Case "020": .cmdSaveDef.Caption = Mid(sLines(i), 5)
                    
                    Case "030": .fraHelp.Caption = Mid(sLines(i), 5)
                    
                    Case "040": .fraConfig.Caption = Mid(sLines(i), 5)
                    Case "041": .chkConfigTabs(0).Caption = Mid(sLines(i), 5)
                    Case "042": .chkConfigTabs(1).Caption = Mid(sLines(i), 5)
                    Case "043": .chkConfigTabs(2).Caption = Mid(sLines(i), 5)
                    Case "044": .chkConfigTabs(3).Caption = Mid(sLines(i), 5)
                    
                    Case "050": .chkAutoMark.Caption = Mid(sLines(i), 5)
                    Case "051": .chkBackup.Caption = Mid(sLines(i), 5)
                    Case "052": .chkConfirm.Caption = Mid(sLines(i), 5)
                    Case "053": .chkIgnoreSafe.Caption = Mid(sLines(i), 5)
                    Case "054": .chkLogProcesses.Caption = Mid(sLines(i), 5)
                    Case "055": .chkShowN00bFrame.Caption = Mid(sLines(i), 5)
                    Case "056": .chkConfigStartupScan.Caption = Mid(sLines(i), 5)
                    
                    Case "060": .lblConfigInfo(3).Caption = Mid(sLines(i), 5)
                    Case "061": .lblConfigInfo(0).Caption = Mid(sLines(i), 5)
                    Case "062": .lblConfigInfo(1).Caption = Mid(sLines(i), 5)
                    Case "063": .lblConfigInfo(2).Caption = Mid(sLines(i), 5)
                    Case "064": .lblConfigInfo(4).Caption = Mid(sLines(i), 5)
                    
                    Case "070": .lblConfigInfo(5).Caption = Mid(sLines(i), 5)
                    Case "071": .cmdConfigIgnoreDelSel.Caption = Mid(sLines(i), 5)
                    Case "072": .cmdConfigIgnoreDelAll.Caption = Mid(sLines(i), 5)
                    
                    Case "080": .lblConfigInfo(6).Caption = Mid(sLines(i), 5)
                    Case "081": .cmdConfigBackupRestore.Caption = Mid(sLines(i), 5)
                    Case "082": .cmdConfigBackupDelete.Caption = Mid(sLines(i), 5)
                    Case "083": .cmdConfigBackupDeleteAll.Caption = Mid(sLines(i), 5)
                    
                    Case "090": .lblConfigInfo(7).Caption = Mid(sLines(i), 5)
                    Case "091": .cmdStartupList.Caption = Mid(sLines(i), 5)
                    Case "092": .chkStartupListFull.Caption = Mid(sLines(i), 5)
                    Case "093": .chkStartupListComplete.Caption = Mid(sLines(i), 5)
                    
                    Case "100": .lblConfigInfo(16).Caption = Mid(sLines(i), 5)
                    Case "101": .cmdProcessManager.Caption = Mid(sLines(i), 5)
                    Case "102": .lblConfigInfo(12).Caption = Mid(sLines(i), 5)
                    Case "103": .cmdHostsManager.Caption = Mid(sLines(i), 5)
                    Case "104": .lblConfigInfo(13).Caption = Mid(sLines(i), 5)
                    Case "105": .cmdDelOnReboot.Caption = Mid(sLines(i), 5)
                    Case "106": .lblInfo(2).Caption = Mid(sLines(i), 5)
                    Case "107": .cmdDeleteService.Caption = Mid(sLines(i), 5)
                    Case "108": .lblInfo(6).Caption = Mid(sLines(i), 5)
                    Case "109": .cmdADSSpy.Caption = Mid(sLines(i), 5)
                    Case "110": .lblInfo(3).Caption = Mid(sLines(i), 5)
                    Case "111": .cmdARSMan.Caption = Mid(sLines(i), 5)
                    Case "112": .lblInfo(7).Caption = Mid(sLines(i), 5)
                    
                    Case "120": .lblConfigInfo(17).Caption = Mid(sLines(i), 5)
                    Case "121": .chkDoMD5.Caption = Mid(sLines(i), 5)
                    Case "122": .chkAdvLogEnvVars.Caption = Mid(sLines(i), 5)
                    
                    Case "130": .lblConfigInfo(22).Caption = Mid(sLines(i), 5)
                    Case "131": .cmdLangLoad.Caption = Mid(sLines(i), 5)
                    Case "132": .cmdLangReset.Caption = Mid(sLines(i), 5)
                    
                    Case "140": .lblConfigInfo(18).Caption = Mid(sLines(i), 5)
                    Case "141": .cmdCheckUpdate.Caption = Mid(sLines(i), 5)
                    Case "142": .lblConfigInfo(10).Caption = Mid(sLines(i), 5)
                    Case "143": .lblConfigInfo(11).Caption = Mid(sLines(i), 5)
                    
                    ''Case "150": .lblConfigInfo(20).Caption = Mid(sLines(i), 5)
                    ''Case "151": .cmdUninstall.Caption = Mid(sLines(i), 5)
                    ''Case "152": .lblConfigInfo(9).Caption = Mid(sLines(i), 5)
                    
                    Case "160": .fraN00b.Caption = Mid(sLines(i), 5)
                    Case "161": .lblInfo(4).Caption = Mid(sLines(i), 5)
                    Case "162": .cmdN00bLog.Caption = Mid(sLines(i), 5)
                    Case "163": .cmdN00bScan.Caption = Mid(sLines(i), 5)
                    Case "164": .cmdN00bBackups.Caption = Mid(sLines(i), 5)
                    Case "165": .cmdN00bTools.Caption = Mid(sLines(i), 5)
                    Case "166": .cmdN00bHJTQuickStart.Caption = Mid(sLines(i), 5)
                    Case "167": .lblInfo(5).Caption = Mid(sLines(i), 5)
                    Case "168": .cmdN00bClose.Caption = Mid(sLines(i), 5)
                    Case "169": .chkShowN00b.Caption = Mid(sLines(i), 5)
                    Case "183": .lblInfo(9).Caption = Mid(sLines(i), 5)
                    
                    Case "170": .fraProcessManager.Caption = Mid(sLines(i), 5)
                    Case "171": .lblConfigInfo(8).Caption = Mid(sLines(i), 5)
                    Case "172": .chkProcManShowDLLs.Caption = Mid(sLines(i), 5)
                    Case "173": .cmdProcManKill.Caption = Mid(sLines(i), 5)
                    Case "174": .cmdProcManRefresh.Caption = Mid(sLines(i), 5)
                    Case "175": .cmdProcManRun.Caption = Mid(sLines(i), 5)
                    Case "176": .cmdProcManBack.Caption = Mid(sLines(i), 5)
                    Case "177": .lblProcManDblClick.Caption = Mid(sLines(i), 5)
                    
                    
                    Case "190": .fraADSSpy.Caption = Mid(sLines(i), 5)
                    Case "191": .chkADSSpyQuick.Caption = Mid(sLines(i), 5)
                    Case "192": .chkADSSpyIgnoreSystem.Caption = Mid(sLines(i), 5)
                    Case "193": .chkADSSpyCalcMD5.Caption = Mid(sLines(i), 5)
                    Case "194": .cmdADSSpyWhatsThis.Caption = Mid(sLines(i), 5)
                    Case "195": .cmdADSSpyHelp.Caption = Mid(sLines(i), 5)
                    Case "196": .cmdADSSpyScan.Caption = Mid(sLines(i), 5)
                    Case "197": .cmdADSSpySaveLog.Caption = Mid(sLines(i), 5)
                    Case "198": .cmdADSSpyRemove.Caption = Mid(sLines(i), 5)
                    Case "199": .cmdADSSpyBack.Caption = Mid(sLines(i), 5)
                    
                    Case "210": .fraUninstMan.Caption = Mid(sLines(i), 5)
                    Case "211": .lblInfo(11).Caption = Mid(sLines(i), 5)
                    Case "212": .lblInfo(8).Caption = Mid(sLines(i), 5)
                    Case "213": .lblInfo(10).Caption = Mid(sLines(i), 5)
                    Case "214": .cmdUninstManDelete.Caption = Mid(sLines(i), 5)
                    Case "215": .cmdUninstManEdit.Caption = Mid(sLines(i), 5)
                    Case "216": .cmdUninstManOpen.Caption = Mid(sLines(i), 5)
                    Case "217": .cmdUninstManRefresh.Caption = Mid(sLines(i), 5)
                    Case "218": .cmdUninstManSave.Caption = Mid(sLines(i), 5)
                    Case "219": .cmdUninstManBack.Caption = Mid(sLines(i), 5)
                    
                    Case "270": .fraHostsMan.Caption = Mid(sLines(i), 5)
                    Case "271": .lblConfigInfo(14).Caption = Mid(sLines(i), 5)
                    Case "272": .cmdHostsManDel.Caption = Mid(sLines(i), 5)
                    Case "273": .cmdHostsManToggle.Caption = Mid(sLines(i), 5)
                    Case "274": .cmdHostsManOpen.Caption = Mid(sLines(i), 5)
                    Case "275": .cmdHostsManBack.Caption = Mid(sLines(i), 5)
                    Case "276": .lblConfigInfo(15).Caption = Mid(sLines(i), 5)
                    
                    Case "999": SetCharSet Mid(sLines(i), 5)
                End Select
            End If
        Next i
    End With
    
    Exit Sub
Error:
    If MsgBox("Invalid language file" & _
              IIf(Err.Number > 0, " (" & Err.Description & ") ", "") & _
              ". Reset to default (English)?", vbYesNo + vbExclamation) = vbYes Then
        LoadDefaultLanguage
        ReloadLanguage
    End If
End Sub

Public Sub LoadDefaultLanguage()
    Dim s$
    s = _
    "// main form" & vbCrLf & _
    "001=Welcome to HijackThis. This program will scan your PC and generate a log file of registry and file settings commonly manipulated by malware as well as good software." & vbCrLf & _
    "002=HijackThis is already running." & vbCrLf & _
    "003=Warning!\n\nSince HijackThis targets browser hijacking methods instead of actual browser hijackers, entries may appear in the scan list that are not hijackers. Be careful what you delete, some system utilities can cause problems if disabled.\nFor best results, ask spyware experts for help and show them your scan log. They will advise you what to fix and what to keep.\n\nSome adware-supported programs may cease to function if the associated adware is removed." & vbCrLf & _
    "004=Below are the results of the HijackThis scan. Be careful what you delete with the 'Fix checked' button. Scan results do not determine whether an item is bad or not. The best thing to do is to 'AnalyzeThis' and show the log file to knowledgeable folks." & vbCrLf & _
    "// main screen" & vbCrLf & _
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
    "022=You selected to fix everything HijackThis has found. This could mean items important to your system will be deleted and the full functionality of your system will degrade.\n\nIf you aren't sure how to use HijackThis, you should ask for help, not blindly fix things. Many 3rd party resources are available on the web that will gladly help you with your log.\n\nAre you sure you want to fix all items in your scan results?" & vbCrLf
    s = s & "023=Fix [] selected items? This will permanently delete and/or repair what you selected" & vbCrLf & _
    "024=unless you enable backups." & vbCrLf & _
    "025=This will set HijackThis to ignore the checked items, unless they change. Continue?" & vbCrLf & _
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
    "056=Run HijackThis scan at startup and show it when items are found" & vbCrLf & _
    "057=Are you sure you want to enable this option? \nHijackThis is not a 'click & fix' program. Because it targets *general* hijacking methods, false positives are a frequent occurrence.\nIf you enable this option, you might disable programs or drivers you need. However, it is highly unlikely you will break your system beyond repair. So you should only enable this option if you know what you're doing!" & vbCrLf & _
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
    "080=This is your list of items that were backed up. You can restore them (causing HijackThis to re-detect them unless you place them on the ignorelist) or delete them from here. (Antivirus programs may detect HijackThis backups!)" & vbCrLf & _
    "081=Restore" & vbCrLf & _
    "082=Delete" & vbCrLf & _
    "083=Delete all" & vbCrLf & _
    "084=Are you sure you want to delete this backup?" & vbCrLf & _
    "085=Are you sure you want to delete these [] backups?" & vbCrLf
    s = s & "086=Restore this item?" & vbCrLf & _
    "087=Restore these [] items?" & vbCrLf & _
    "088=Are you sure you want to delete ALL backups?" & vbCrLf & _
    "// 'misc tools' tab" & vbCrLf & _
    "090=StartupList (integrated: v1.52)" & vbCrLf & _
    "091=Generate StartupList log" & vbCrLf & _
    "092=List also minor sections (full)" & vbCrLf & _
    "093=List empty sections (complete)" & vbCrLf & _
    "100=System tools" & vbCrLf & _
    "101=Open process manager" & vbCrLf & _
    "102=Opens a small process manager, working much like the Task Manager." & vbCrLf & _
    "103=Open hosts file manager" & vbCrLf & _
    "104=Opens an editor for the 'hosts' file." & vbCrLf & _
    "105=Delete a file on reboot..." & vbCrLf & _
    "106=If a file cannot be removed from memory, Windows can be setup to delete it when the system is restarted." & vbCrLf & _
    "107=Delete an NT service..." & vbCrLf & _
    "108=Delete a Windows NT Service (O23). USE WITH CAUTION! (WinNT4/2k/XP only)" & vbCrLf & _
    "109=Open ADS Spy..." & vbCrLf & _
    "110=Open the integrated ADS Spy utility to scan for hidden data streams." & vbCrLf & _
    "111=Open Uninstall Manager..." & vbCrLf
    s = s & "112=Open a utility to manage the items in the Add/Remove Software list." & vbCrLf & _
    "113=Enter the exact service name as it appears in the scan results, or the short name between brackets if that is listed.\nThe service needs to be stopped and disabled.\nWARNING! When the service is deleted, it cannot be restored!" & vbCrLf & _
    "114=Delete a Windows NT Service" & vbCrLf & _
    "115=Service '[]' was not found in the Registry.\nMake sure you entered the name of the service correctly." & vbCrLf & _
    "116=The service '[]' is enabled and/or running. Disable it first, using HijackThis itself (from the scan results) or the Services.msc window." & vbCrLf & _
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
    "150=Uninstall HijackThis" & vbCrLf & _
    "151=Uninstall HijackThis && exit" & vbCrLf & _
    "152=Removes all Registry settings and exits." & vbCrLf
    s = s & "153=This will remove HijackThis' settings from the Registry and exit. Note that you will have to delete the HijackThis.exe file manually.\n\nContinue with uninstall?" & vbCrLf & _
    "// n00b screen" & vbCrLf & _
    "160=Main Menu" & vbCrLf & _
    "161=What would you like to do?" & vbCrLf & _
    "162=Do a system scan and save a logfile" & vbCrLf & _
    "163=Do a system scan only" & vbCrLf & _
    "164=View the list of backups" & vbCrLf & _
    "165=Open the Misc Tools section" & vbCrLf & _
    "166=Open online HijackThis QuickStart" & vbCrLf & _
    "167=" & vbCrLf & _
    "168=None of the above, just start the program" & vbCrLf & _
    "169=Do not show this window when I start HijackThis" & vbCrLf & _
    "183=Change language:" & vbCrLf & _
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
    "// msgbox'es and stuff in scan methods" & vbCrLf & _
    "300=For some reason your system denied write access to the Hosts file. If any hijacked domains are in this file, HijackThis may NOT be able to fix this.\n\nIf that happens, you need to edit the file yourself. To do this, click Start, Run and type:\n\n   notepad []\n\nand press Enter. Find the line(s) HijackThis reports and delete them. Save the file as 'hosts.' (with quotes), and reboot.\n\nFor Vista: simply, exit HijackThis, right click on the HijackThis icon, choose 'Run as administrator'." & vbCrLf & _
    "301=Your hosts file has invalid linebreaks and HijackThis is unable to fix this. O1 items will not be displayed.\n\nClick OK to continue the rest of the scan." & vbCrLf & _
    "302=You have an particularly large amount of hijacked domains. It's probably better to delete the file itself then to fix each item (and create a backup).\n\nIf you see the same IP address in all the reported O1 items, consider deleting your Hosts file, which is located at []." & vbCrLf & _
    "303=HijackThis could not write the selected changes to your hosts file. The probably cause is that some program is denying access to it, or that your user account doesn't have the rights to write to it." & vbCrLf & _
    "310=HijackThis is about to remove a BHO and the corresponding file from your system. Close all Internet Explorer windows AND all Windows Explorer windows before continuing for the best chance of success." & vbCrLf & _
    "320=Unable to delete the file:\n[]\n\nThe file may be in use." & vbCrLf
    s = s & "321=Use Task Manager to shutdown the program" & vbCrLf & _
    "322=Use a process killer like ProcView to shutdown the program" & vbCrLf & _
    "323=and run HijackThis again to delete the file." & vbCrLf & _
    "330=HijackThis is about to remove a plugin from your system. Close all Internet Explorer windows before continuing for the best chance of success." & vbCrLf & _
    "340=It looks like you're running HijackThis from a read-only device like a CD or locked floppy disk.If you want to make backups of items you fix, you must copy HijackThis.exe to your hard disk first, and run it from there.\n\nIf you continue, you might get 'Path/File Access' errors." & vbCrLf & _
    "341=HijackThis appears to have been started from a temporary folder. Since temp folders tend to be be emptied regularly, it's wise to copy HijackThis.exe to a folder of its own, for instance C:\Program Files\HijackThis.\nThis way, any backups that will be made of fixed items won't be lost.\n\nPlease quit HijackThis and copy it to a separate folder first before fixing any items." & vbCrLf & _
    "342=The file '[]' will be deleted by Windows when the system restarts." & vbCrLf & _
    "343=The service '[]'  has been marked for deletion." & vbCrLf & _
    "344=Unable to delete the service '[]'. Make sure the name is correct and the service is not running." & vbCrLf & _
    "// = misc stuff =" & vbCrLf & _
    "// HJT quickstart URL" & vbCrLf & _
    "360=http://www.trendmicro.com/go/hjt/quickstart/"
    
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
    ErrorMsg "modLanguage_Translate", Err.Number, Err.Description, szParam
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
