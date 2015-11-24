Attribute VB_Name = "modBackup"
Option Explicit

Public Sub MakeBackup(ByVal sItem$)
    Dim sPath$, lHive&, sKey$, sValue$, sSID$, sUserName$
    Dim sData$, sDummy$, sBackup$, sLine$
    Dim sDPFKey$, sCLSID$, sOSD$, sINF$, sInProcServer32$
    Dim sNum$, sFile1$, sFile2$
    On Error Resume Next
    sPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
    
    If bNoWriteAccess Then Exit Sub
    If InStr(sPath, ".zip") > 0 Then Exit Sub 'running from zip in XP
    If Not FolderExists(sPath & "backups") Then MkDir sPath & "backups"
    If Not FolderExists(sPath & "backups") Then
        MsgBox "Unable to create folder to place backups in. Backups of fixed items cannot be saved!", vbExclamation
        bNoWriteAccess = True
        Exit Sub
    End If
    
    'create backup file name
    Randomize
    sBackup = "backup-" & Format(Date, "yyyymmdd") & "-" & Format(Time, "HhNnSs") & "-" & CStr(1000 * Format(Rnd(), "0.000"))
    If Dir(sPath & "backups\" & sBackup & "*.*") <> vbNullString Or _
       InStrRev(sBackup, "-") <> Len(sBackup) - 3 Then
        Do
            sBackup = "backup-" & Format(Date, "yyyymmdd") & "-" & Format(Time, "HhNnSs") & "-"
            Randomize
            sBackup = sBackup & CStr(1000 * Format(Rnd(), "0.000"))
        Loop Until Dir(sPath & "backups\" & sBackup & "*.*") = vbNullString And _
                   InStrRev(sBackup, "-") = Len(sBackup) - 3
    End If
    
    On Error GoTo Error:
    Select Case Trim(Left(sItem, 3))
        'these lot don't need any additional stuff
        'backed up, everything is in the sItem line
        Case "R0", "R3"
        Case "F0", "F1", "F2", "F3"
        Case "N1", "N2", "N3", "N4"
        Case "O1", "O3", "O5"
        Case "O7", "O8", "O13", "O14"
        Case "O15", "O17", "O18", "O19"
        Case "O23"
        
        'below items that DO need something else
        'backed up
        
        Case "R1"
            'R1 - Created Registry value
            'R1 - HKCU\Software\..\Subkey,Value[=Data]
            
            'need to get sData if not in sItem
            If InStr(sItem, "=") = 0 Or Right(sItem, 1) = "=" Then
                sDummy = Mid(sItem, 6)
                Select Case Left(sDummy, 4)
                    Case "HKCU": lHive = HKEY_CURRENT_USER
                    Case "HKCR": lHive = HKEY_CLASSES_ROOT
                    Case "HKLM": lHive = HKEY_LOCAL_MACHINE
                End Select
                sDummy = Mid(sDummy, 6)
                sKey = Left(sDummy, InStr(sDummy, ",") - 1)
                sValue = Mid(sDummy, InStr(sDummy, ",") + 1)
                If InStr(sValue, "=") > 0 Then sValue = Left(sValue, InStr(sValue, "=") - 1)
                sData = RegGetString(lHive, sKey, sValue)
                sItem = sItem & "=" & sData
                sData = vbNullString
            End If
            
        Case "R2"
            'R2 - Created Registry key
            'R2 - HKCU\Software\..\Subkey
            
            'don't have rules with R2 yet...
            'getting one would mean enumerating
            'all values in key and save them
            'in sData
            
            MsgBox "Not implemented yet, item '" & _
              sItem & "' will not be backed up!", _
              vbExclamation, "bad coder - no donuts"
            
        Case "O2"
            'O2 - BHO
            'O2 - BHO: BhoName - CLSID - Filename
            
            'backup BHO dll
            Dim vDummy As Variant
            vDummy = Split(sItem, " - ")
            If UBound(vDummy) <> 3 Then
                If InStr(sItem, "}") > 0 And _
                   InStr(sItem, "- ") > 0 Then
                    sDummy = Mid(sItem, InStr(InStr(sItem, "}"), sItem, "- ") + 2)
                End If
            Else
                sDummy = CStr(vDummy(3))
            End If
            On Error Resume Next
            If FileExists(sDummy) Then FileCopy sDummy, sPath & "backups\" & sBackup & ".dll"
            On Error GoTo Error:
            
        Case "O4"
            'O4 - Regrun and Startup run, also for other users
            'O4 - Common Startup: Bla.lnk = c:\dummy.exe
            If InStr(sItem, "[") = 0 Then
                'need to backup link
                sData = Mid(sItem, 6)
                If InStr(sItem, " (User '") = 0 Then 'normal item
                    sData = Left(sData, InStr(sData, ":") - 1)
                    Select Case sData
                        Case "Startup":                sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
                        Case "AltStartup":             sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AltStartup")
                        Case "User Startup":           sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
                        Case "User AltStartup":        sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AltStartup")
                        Case "Global Startup":         sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common Startup")
                        Case "Global AltStartup":      sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common AltStartup")
                        Case "Global User Startup":    sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common Startup")
                        Case "Global User AltStartup": sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common AltStartup")
                    End Select
                Else 'item from other user account
                    sSID = Left(sData, InStr(sData, " ") - 1)
                    sUserName = MapSIDToUsername(sSID)
                    sData = Mid(sData, InStr(sData, " ") + 1)
                    sData = Left(sData, InStr(sData, ":") - 1)
                    Select Case sData
                        Case "Startup":                sData = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
                        Case "AltStartup":             sData = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AltStartup")
                        Case "User Startup":           sData = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
                        Case "User AltStartup":        sData = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AltStartup")
                    End Select
                    If sData <> vbNullString And FolderExists(sData) Then
                        sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
                        If InStr(sDummy, " = ") > 0 Then
                            sDummy = Left(sDummy, InStr(sDummy, " = ") - 1)
                        End If
                        sData = sData & IIf(Right(sData, 1) = "\", "", "\") & sDummy
                        On Error Resume Next
                        If FileExists(sData) Then FileCopy sData, sPath & "backups\" & sBackup & "-" & sDummy
                        On Error GoTo Error:
                        'sdata had no relevant backup data, just dummy
                        sData = vbNullString
                    End If
                End If
                'attempt backup the file, same for both
                If sData <> vbNullString And FolderExists(sData) Then
                    sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
                    If InStr(sDummy, " = ") > 0 Then
                        sDummy = Left(sDummy, InStr(sDummy, " = ") - 1)
                    End If
                    sData = sData & IIf(Right(sData, 1) = "\", "", "\") & sDummy
                    On Error Resume Next
                    If FileExists(sData) Then
                        If (GetAttr(sData) And vbDirectory) Then
                            CopyFolder sData, sPath & "backups\" & sBackup & "-" & sDummy
                        Else
                            FileCopy sData, sPath & "backups\" & sBackup & "-" & sDummy
                        End If
                    End If
                    On Error GoTo Error:
                    'sdata had no relevant backup data, just dummy
                    sData = vbNullString
                End If
            Else
                'registry autorun, nothing to backup
            End If
            
        Case "O6"
            'O6 - IE Policies block
            'O6 - HKCU\Software\Policies\Microsoft\Internet Explorer\Restrictions
            
            'need to back up everything in those keys
            If InStr(sItem, "HKCU") > 0 And InStr(sItem, "Restrictions") > 0 Then
                Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-policy.reg"" ""HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions""", vbHide
            ElseIf InStr(sItem, "HKCU") > 0 And InStr(sItem, "Control Panel") > 0 Then
                Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-policy.reg"" ""HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Control Panel""", vbHide
            ElseIf InStr(sItem, "HKLM") > 0 And InStr(sItem, "Restrictions") > 0 Then
                Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-policy.reg"" ""HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Restrictions""", vbHide
            ElseIf InStr(sItem, "HKLM") > 0 And InStr(sItem, "Control Panel") > 0 Then
                Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-policy.reg"" ""HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Control Panel""", vbHide
            End If
            
            sData = vbNullString
            If FileExists(sPath & "backups\" & sBackup & "-policy.reg") Then
                Open sPath & "backups\" & sBackup & "-policy.reg" For Input As #1
                    Do
                        Line Input #1, sLine
                        sData = sData & sLine & vbCrLf
                    Loop Until EOF(1)
                Close #1
                DeleteFile sPath & "backups\" & sBackup & "-policy.reg"
            End If
            
        Case "O9"
            'O9 - IE Tools menu item/button
            'O9 - Extra button: Offline
            'O9 - Extra 'Tools' menuitem: Add to T&rusted Zone
            
            'O9 - Extra 'Tools' menuitem: Related - {000...000} - c:\file.dll [(HKCU)]
            
            'need to backup all values in regkey
            sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
            sDummy = Mid(sDummy, InStr(sDummy, " - ") + 3)
            sDummy = Left(sDummy, InStr(sDummy, " - ") - 1)
            
            'If InStr(sItem, "Extra button:") > 0 Then
            '    sDummy = GetCLSIDOfMSIEExtension(Mid(sItem, InStr(sItem, ":") + 2), True)
            'Else
            '    sDummy = GetCLSIDOfMSIEExtension(Mid(sItem, InStr(sItem, ":") + 2), False)
            'End If
            If sDummy = vbNullString Then Exit Sub
            
            If InStr(sItem, " (HKCU)") > 0 Then
                Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-extension.reg"" ""HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Extensions\" & sDummy & """", vbHide
            Else
                Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-extension.reg"" ""HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\" & sDummy & """", vbHide
            End If
            
            If FileExists(sPath & "backups\" & sBackup & "-extension.reg") Then
                sData = vbNullString
                Open sPath & "backups\" & sBackup & "-extension.reg" For Input As #1
                    Do
                        Line Input #1, sLine
                        sData = sData & sLine & vbCrLf
                    Loop Until EOF(1)
                Close #1
                DeleteFile sPath & "backups\" & sBackup & "-extension.reg"
            End If
            
        Case "O10"
            'O10 - Winsock hijack
            'O10 - Broken Internet access because of missing LSP provider: 'file'
            'O10 - Broken Internet access because of LSP chain gap (#2 in chain of 8 missing)"
            
            'backup even possible??
            'MsgBox "Backup of LSP hijackers is not possible " & _
            '       "because of technical limitations. (IOW, " & _
            '       "I don't know how.) Since only two programs " & _
            '       "hijack the LSP (New.Net and WebHancer) and " & _
            '       "both, this should not pose a problem." & vbCrLf & _
            '       "Should you wish to restore either for testing " & _
            '       "purposes or complete insanity, you need to " & _
            '       "reinstall the program.", vbExclamation
            Exit Sub
            
        Case "O11"
            'O11 - Extra options in MSIE 'Advanced' settings tab
            'O11 - Options group: [COMMONNAME] CommonName
            
            'need to backup everything in that key
            sDummy = Left(sItem, InStr(sItem, "]") - 1)
            sDummy = Mid(sItem, InStr(sItem, "[") + 1)
            Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-advopt.reg"" ""HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\AdvancedOptions\" & sDummy & """", vbHide
            
            If Dir(sPath & "backups\" & sBackup & "-advopt.reg") <> vbNullString Then
                sData = vbNullString
                Open sPath & "backups\" & sBackup & "-advopt.reg" For Input As #1
                    Do
                        Line Input #1, sLine
                        sData = sData & sLine & vbCrLf
                    Loop Until EOF(1)
                Close #1
                DeleteFile sPath & "backups\" & sBackup & "-advopt.reg"
            End If
            
        Case "O12"
            'O12 - MSIE plugins for file extensions or MIME types
            'O12 - Plugin for .spop: NAV.DLL
            'O12 - Plugin for text/html: NAV.DLL
            
            'need to backup subkey + 'Location' value
            If InStr(sItem, "Plugin for .") > 0 Then
                'plugin for file extension
                sDummy = Left(sItem, InStr(sItem, ":") - 1)
                sDummy = Mid(sDummy, InStr(sDummy, "."))
                sDummy = "Extension\" & sDummy
            Else
                'plugin for MIME type
                sDummy = Left(sItem, InStr(sItem, ":") - 1)
                sDummy = Mid(sItem, InStr(sItem, " for ") + 5)
                sDummy = "MIME\" & sDummy
            End If
            Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-plugin.reg"" ""HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Plugins\" & sDummy & """", vbHide
            
            If FileExists(sPath & "backups\" & sBackup & "-plugin.reg") Then
                sData = vbNullString
                Open sPath & "backups\" & sBackup & "-plugin.reg" For Input As #1
                    Do
                        Line Input #1, sLine
                        sData = sData & sLine & vbCrLf
                    Loop Until EOF(1)
                Close #1
                DeleteFile sPath & "backups\" & sBackup & "-plugin.reg"
            End If
            
        Case "O16"
            'O16 - Download Program Files item
            'O16 - DPF: Plugin name - http://bla.com/bla.cab
            'O16 - DPF: {000000} (name) - http://bla.com/bla.cab
            
            'need to export key from HKLM\..\Dist Units
            'and (if applic) HKCR\CLSID\{0000}
            'need to backup files OSD, INF, InProcServer32
            
            sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
            If Left(sDummy, 1) = "{" Then
                'name is CLSID
                sDummy = Left(sDummy, InStr(sDummy, "}"))
            Else
                'name is just name
                sDummy = Left(sDummy, InStr(sDummy, " - ") - 1)
            End If
            Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-dpf1.reg"" ""HKEY_LOCAL_MACHINE\Software\Microsoft\Code Store Database\Distribution Units\" & sDummy & """", vbHide
            If Left(sDummy, 1) = "{" Then
                Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-dpf2.reg"" ""HKEY_CLASSES_ROOT\CLSID\" & sDummy & """", vbHide
            End If
            DoEvents
            
            If FileExists(sPath & "backups\" & sBackup & "-dpf1.reg") Then
                sData = vbNullString
                If FileLen(sPath & "backups\" & sBackup & "-dpf1.reg") > 0 Then
                    Open sPath & "backups\" & sBackup & "-dpf1.reg" For Input As #1
                        Do
                            Line Input #1, sLine
                            sData = sData & sLine & vbCrLf
                        Loop Until EOF(1)
                    Close #1
                End If
                DeleteFile sPath & "backups\" & sBackup & "-dpf1.reg"
                
                If FileExists(sPath & "backups\" & sBackup & "-dpf2.reg") Then
                    If FileLen(sPath & "backups\" & sBackup & "-dpf2.reg") > 0 Then
                        Open sPath & "backups\" & sBackup & "-dpf2.reg" For Input As #1
                            'ignore first line (REGEDIT4)
                            Line Input #1, sLine
                            Do
                                Line Input #1, sLine
                                sData = sData & sLine & vbCrLf
                            Loop Until EOF(1)
                        Close #1
                    End If
                    DeleteFile sPath & "backups\" & sBackup & "-dpf2.reg"
                End If
            End If
            
            sCLSID = sDummy
            sDPFKey = "Software\Microsoft\Code Store Database\Distribution Units"
            'backup INF
            sLine = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sCLSID & "\DownloadInformation", "INF")
            If sLine <> vbNullString Then
                If FileExists(sLine) Then
                    FileCopy sLine, sPath & "backups\" & sBackup & ".inf"
                End If
            End If
            
            'backup OSD
            sLine = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sCLSID & "\DownloadInformation", "OSD")
            If sLine <> vbNullString Then
                If FileExists(sLine) Then
                    On Error Resume Next
                    FileCopy sLine, sPath & "backups\" & sBackup & ".osd"
                    On Error GoTo Error:
                End If
            End If
            
            'backup InProcServer32
            If Left(sCLSID, 1) = "{" And Right(sCLSID, 1) = "}" Then
                sLine = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", "")
                If sLine <> vbNullString Then
                    If FileExists(sLine) Then
                        On Error Resume Next
                        FileCopy sLine, sPath & "backups\" & sBackup & ".dll"
                        On Error GoTo Error:
                    End If
                End If
            End If
            
        Case "O20"
            'O20 - AppInit_DLLs: file.dll (do nothing)
            'O20 - Winlogon Notify: bla - c:\file.dll
            'todo:
            'backup regkey
            If InStr(sItem, "Winlogon Notify:") > 0 Then
                sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
                sDummy = Left(sDummy, InStr(sDummy, " - ") - 1)
                
                Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-notify.reg"" ""HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify\" & sDummy & """", vbHide
                DoEvents
                
                If FileExists(sPath & "backups\" & sBackup & "-notify.reg") Then
                    Open sPath & "backups\" & sBackup & "-notify.reg" For Binary As #1
                        sData = Input(FileLen(sPath & "backups\" & sBackup & "-notify.reg"), #1)
                    Close #1
                    DeleteFile sPath & "backups\" & sBackup & "-notify.reg"
                End If
            End If
        
        Case "O21"
            'O21 - ShellServiceObjectDelayLoad
            'O21 - SSODL: webcheck - {000....000} - c:\file.dll
            'todo:
            'backup CLSID regkey
            sCLSID = Mid(sItem, 14)
            sCLSID = Mid(sCLSID, InStr(sCLSID, " - ") + 3)
            sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
            
            Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-ssodl.reg"" ""HKEY_CLASSES_ROOT\CLSID\" & sCLSID & """", vbHide
            DoEvents
            
            If FileExists(sPath & "backups\" & sBackup & "-ssodl.reg") Then
                Open sPath & "backups\" & sBackup & "-ssodl.reg" For Binary As #1
                    sData = Input(FileLen(sPath & "backups\" & sBackup & "-ssodl.reg"), #1)
                Close #1
                DeleteFile sPath & "backups\" & sBackup & "-ssodl.reg"
            End If
        
        Case "O22"
            'O22 - SharedTaskScheduler: blah - {000...000} - file.dll
            'todo:
            'backup CLSID regkey
            sCLSID = Mid(sItem, InStr(sItem, ": ") + 2)
            sCLSID = Mid(sCLSID, InStr(sCLSID, " - ") + 3)
            sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
            
            Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-sts.reg"" ""HKEY_CLASSES_ROOT\CLSID\" & sCLSID & """", vbHide
            DoEvents
        
            If FileExists(sPath & "backups\" & sBackup & "-sts.reg") Then
                Open sPath & "backups\" & sBackup & "-sts.reg" For Binary As #1
                    sData = Input(FileLen(sPath & "backups\" & sBackup & "-sts.reg"), #1)
                Close #1
                DeleteFile sPath & "backups\" & sBackup & "-sts.reg"
            End If
        
        Case "O24"
            'O24 - Desktop Component N: blah - c:\windows\index.html
            'todo:
            'backup regkey, html file (if present)
            sNum = Mid(sItem, InStr(sItem, ":") - 1, 1)
            sFile1 = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Desktop\Components\" & sNum, "Source")
            sFile2 = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Desktop\Components\" & sNum, "SubscribedURL")
            If LCase(sFile2) = LCase(sFile1) Then sFile2 = vbNullString
            
            Shell sWinDir & "\regedit.exe /e """ & sPath & "backups\" & sBackup & "-dc.reg"" ""HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Desktop\Components\" & sNum & """", vbHide
            DoEvents
            
            If FileExists(sPath & "backups\" & sBackup & "-dc.reg") Then
                Open sPath & "backups\" & sBackup & "-dc.reg" For Binary As #1
                    sData = Input(FileLen(sPath & "backups\" & sBackup & "-dc.reg"), #1)
                Close #1
                DeleteFile sPath & "backups\" & sBackup & "-dc.reg"
            End If
            If FileExists(sFile1) Then FileCopy sFile1, sPath & "backups\" & sBackup & "-source.html"
            If FileExists(sFile2) Then FileCopy sFile2, sPath & "backups\" & sBackup & "-suburl.html"
            
        Case Else
            MsgBox "I'm so stupid I forgot to implement this. Bug me about it." & vbCrLf & sItem, vbExclamation, "d'oh!"
    End Select
        
    'winNT/2000/XP reg data workaround
    If Left(sData, 2) = "ÿþ" Then sData = Mid(sData, 3)
    sData = StrConv(sData, vbFromUnicode)
    
    'write item + any data to file
    Open sPath & "backups\" & sBackup For Output As #1
        Print #1, sItem
        If sData <> vbNullString Then Print #1, vbCrLf & sData
    Close #1
    Exit Sub
    
Error:
    Close #1
    ErrorMsg "modBackup_MakeBackup", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub RestoreBackup(ByVal sItem$)
    'format of backup files:
    'line 1: original item, e.g. O1 - Hosts: auto.search.msn
    'line 2: blank if 3 != blank
    'line 3+: any Registry data
    'format of sItem:
    ' [short date], [long time]: [original item name]
    Dim sPath$, sDate$, sTime$, sFile$, sBackup$, sSID$
    Dim sName$, sDummy$, i%, sKey1, sKey2$
    Dim sRegKey$, sRegKey2$, sRegKey3$, sRegKey4$, sRegKey5$
    Dim bBackupHasRegData As Boolean, bBackupHasDLL As Boolean
    On Error GoTo Error:
    
    sPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
    If Not FolderExists(sPath & "backups") Then Exit Sub
    
    sDate = Left(sItem, InStr(sItem, ", ") - 1)
    sTime = Mid(sItem, InStr(sItem, ", ") + 2)
    sName = Mid(sTime, InStr(sTime, ": ") + 2)
    sTime = Left(sTime, InStr(sTime, ": ") - 1)
    
    sItem = Mid(sItem, InStr(sItem, ": ") + 2)
    
    If Not bIsUSADateFormat Then
        sDate = Format(sDate, "yyyymmdd")
    Else
        'use stupid workaround for USA data format
        sDate = Format(sDate, "yyyyddmm")
    End If
    sTime = Format(sTime, "HhNnSs")
    
    'sBackup = "backup-" & sDate & "-" & sTime '& "*.*"
    sBackup = "backup-*.*"
    
    'get first file for this filemask
    'multiple backups can exist, so open each file
    'and check the first line against the item
    'we are looking for
    sFile = Dir(sPath & "backups\" & sBackup)
    If sFile = vbNullString Then
        'note the small difference with the next msg
        MsgBox "The backup files for this item were not found. It could not be restored.", vbExclamation
        
        'MsgBox "Dir(" & sPath & "backups\" & sBackup & "*.*) = vbNullString"
        Exit Sub
    End If
    Do
        If InStr(sFile, ".") = 0 Then
            Open sPath & "backups\" & sFile For Input As #1
                Line Input #1, sDummy
                If sDummy = sName Then
                    sBackup = sFile
                    Close #1
                    Exit Do
                End If
            Close #1
        End If
        sFile = Dir
    Loop Until sFile = vbNullString
    If sDummy <> sName Then
        'things like this help troubleshooting stupid bugs
        MsgBox "The backup file for this item was not found. It could not be restored.", vbExclamation
        
        'MsgBox "sDummy = " & sDummy & vbCrLf & _
        '       "sName = " & sName
        Exit Sub
    End If
    
    'file types:
    'backup*. = actual backup file /w item name + reg data
    'backup*.dll = backup of BHO file
    If FileExists(sPath & "backups\" & sFile & ".dll") Then bBackupHasDLL = True
    Open sPath & "backups\" & sFile For Input As #1
        Line Input #1, sDummy
        On Error Resume Next
        sDummy = vbNullString
        Line Input #1, sDummy
        Line Input #1, sDummy
        On Error GoTo Error:
        If sDummy <> vbNullString Then bBackupHasRegData = True
    Close #1
    
    Dim lHive&, sKey$, sVal$, sData$
    Dim sIniFile$, sSection$
    Dim sMyFile$, sMyName$, sCLSID$, sLine$
    Dim sDPFKey$, sINF$, sOSD$, sInProcServer32$
    
    Select Case Trim(Left(sName, 3))
        Case "R0", "R1" 'Changed/Created Regval
            'R0 - HKCU\Software\..\Subkey,Value=Data
            'R1 - HKCU\Software\..\Subkey,Value=Data
            sDummy = Mid(sName, 6)
            Select Case Left(sDummy, 4)
                Case "HKCU": lHive = HKEY_CURRENT_USER
                Case "HKLM": lHive = HKEY_LOCAL_MACHINE
            End Select
            sKey = Mid(sDummy, 6)
            sVal = Mid(sKey, InStr(sKey, ",") + 1)
            sKey = Left(sKey, InStr(sKey, ",") - 1)
            sData = Mid(sVal, InStr(sVal, " = ") + 3)
            sVal = Left(sVal, InStr(sVal, " = ") - 1)
            If sVal = "(Default)" Then sVal = ""
            If InStr(sData, " (obfuscated)") > 0 Then
                sData = Left(sData, InStr(sData, " (obfuscated)") - 1)
            End If
            RegSetStringVal lHive, sKey, sVal, sData
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "R2" 'Created Regkey
            'don't have this yet
            
        Case "R3" 'URLSearchHoook
            'R3 - URLSearchHook: blah - {0000} - bla.dll
            sDummy = Mid(sName, InStr(sName, "- {") + 2)
            sDummy = Left(sDummy, InStr(sDummy, "}"))
            RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks", sDummy, ""
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "F0", "F1" 'Changed/Created Inifile val
            'F0 - system.ini: Shell=Explorer.exe openme.exe
            sDummy = Mid(sName, 6)
            'sMyFile = Left(sDummy, InStr(sDummy, ":") - 1)
            If InStr(sDummy, "system.ini") = 1 Then
                sSection = "boot"
                sMyFile = "system.ini"
            ElseIf InStr(sDummy, "win.ini") = 1 Then
                sSection = "windows"
                sMyFile = "win.ini"
            End If
            sVal = Mid(sDummy, InStr(sDummy, ": ") + 2)
            'sMyFile = Left(sDummy, InStr(sDummy, ":") - 1)
            sData = Mid(sVal, InStr(sVal, "=") + 1)
            sVal = Left(sVal, InStr(sVal, "=") - 1)
            'WritePrivateProfileString sSection, sVal, sData, sMyFile
            IniSetString sMyFile, sSection, sVal, sData
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "F2", "F3"
            'F2 - REG:system.ini: Shell=Explorer.exe blah
            'F2 - REG:system.ini: Userinit=c:\windows\system32\userinit.exe,blah
            'F3 - REG:win.ini: load=blah or run=blah
            sData = Mid(sName, InStr(sName, "=") + 1)
            If InStr(sName, "system.ini") > 0 Then
                If InStr(1, sName, "Shell=", vbTextCompare) > 0 Then
                    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "Shell", sData
                ElseIf InStr(1, sName, "Userinit", vbTextCompare) > 0 Then
                    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "UserInit", sData
                End If
            ElseIf InStr(sName, "win.ini") > 0 Then
                If InStr(sName, "load=") > 0 Then
                    RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "load", sData
                ElseIf InStr(sName, "run=") > 0 Then
                    RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "run", sData
                End If
            End If
        Case "N1", "N2", "N3", "N4"
            'Changed NS4.x homepage
            'N1 - Netscape 4: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
            'Changed NS6 homepage
            'N2 - Netscape 6: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
            'Changed NS7 homepage/searchpage
            'N3 - Netscape 7: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
            'Changed Moz homepage/searchpage
            'N4 - Mozilla: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
            '               user_pref("browser.search.defaultengine", "http://url"); (c:\..\prefs.js)
            
            'get user_pref line + prefs.js location
            sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
            sMyFile = Mid(sDummy, InStrRev(sDummy, "(") + 1)
            sMyFile = Left(sMyFile, Len(sMyFile) - 1)
            sDummy = Left(sDummy, InStrRev(sDummy, "(") - 2)
            
            If Not FileExists(sMyFile) Then
                MsgBox "Could not find prefs.js file for Netscape/Mozilla, homepage has not been restored.", vbExclamation
                Exit Sub
            End If
            
            'read old file, replacing relevant line
            sData = vbNullString
            Open sMyFile For Input As #1
                Do
                    Line Input #1, sLine
                    If InStr(sLine, sDummy) > 0 Then
                        sData = sData & sDummy & vbCrLf
                    Else
                        sData = sData & sLine & vbCrLf
                    End If
                Loop Until EOF(1)
            Close #1
            
            'write new file
            If FileExists(sMyFile) Then DeleteFile sMyFile
            Open sMyFile For Output As #1
                Print #1, sData
            Close #1
            
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O1" 'Hosts file hijack
            'O1 - Hosts file: 66.123.204.8 auto.search.msn.com
            If InStr(sName, "Hosts file is located at") > 0 Then
                sDummy = Mid(sName, InStr(sName, " at") + 5)
                sDummy = Left(sDummy, Len(sDummy) - 6)
                RegSetStringVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", sDummy
            Else
                sDummy = Mid(sName, InStr(sName, ": ") + 2)
                i = GetAttr(sHostsFile)
                If (i And 2048) Then i = i - 2048
                SetAttr sHostsFile, vbNormal
                Open sHostsFile For Append As #1
                    Print #1, sDummy
                Close #1
                SetAttr sHostsFile, i
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O2" 'BHO
            'O2 - BHO: [bhoname] - [clsid] - [file]
            sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
            sMyName = Left(sDummy, InStr(sDummy, " - ") - 1)
            sCLSID = Mid(sDummy, InStr(sDummy, " - ") + 3)
            sMyFile = Mid(sCLSID, InStr(sCLSID, " - ") + 3)
            If InStr(sMyFile, "(file missing)") > 0 Then
                sMyFile = Left(sMyFile, InStr(sMyFile, "(file missing)") - 1)
            End If
            sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
            
            RegCreateKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID
            If sMyName <> "(no name)" Then RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID, vbNullString, sMyName
            If Not RegKeyExists(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32") Then
                RegCreateKey HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
                RegCreateKey HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32"
                If sMyFile <> vbNullString And sMyFile <> "(no file)" Then
                    RegSetStringVal HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, sMyFile
                End If
            End If
            
            If Not bBackupHasDLL Then
                MsgBox "BHO file for '" & sName & "' was not found. The Registry data was restored, but the file was not.", vbExclamation
            Else
                'skip errors - app could have restored
                'BHO dll by itself
                On Error Resume Next
                FileCopy sPath & "backups\" & sBackup & ".dll", sMyFile
                On Error GoTo Error:
                Shell sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\regsvr32.exe /s """ & sMyFile & """", vbHide
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            If FileExists(sPath & "backups\" & sFile & ".dll") Then DeleteFile sPath & "backups\" & sFile & ".dll"
            
        Case "O3" 'IE Toolbar
            'O3 - Toolbar: Radio - {00000000-0000-0000-0000-000000000000}
            sMyName = Mid(sItem, InStr(sItem, ": ") + 2)
            sCLSID = Mid(sMyName, InStr(sMyName, " - ") + 3)
            sCLSID = Left(sCLSID, InStr(sCLSID, "}"))
            sMyName = Left(sMyName, InStr(sMyName, " - ") - 1)
            'If sMyName = "(no name)" Then sMyName = vbNullString
            
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Toolbar", sCLSID, sMyName
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O4" 'Regrun entry
            'O4 - HKLM\..\Run: [valuename] rundll shit.dll,LoadEtc
            'O4 - Startup: bla.lnk = c:\bla.exe
            
            'O4 - HKCU\SID\Run: [bla] bla.exe (User 'bla')
            'O4 - SID Startup: bla.lnk = c:\bla.exe (User 'bla')
                
            If InStr(sItem, "[") > 0 Then
                'registry autorun
                sDummy = Mid(sItem, 6)
                Select Case Left(sDummy, 4)
                    Case "HKLM": lHive = HKEY_LOCAL_MACHINE
                    Case "HKCU": lHive = HKEY_CURRENT_USER
                    Case "HKUS": lHive = HKEY_USERS
                End Select
                If Not lHive = HKEY_USERS Then
                    sDummy = Mid(sDummy, 9)
                Else
                    sDummy = Mid(sDummy, 6)
                    sSID = Left(sDummy, InStr(sDummy, "\") - 1)
                    sDummy = Mid(sDummy, Len(sSID) + 5)
                End If
                If InStr(sDummy, "RunOnce:") = 1 Then
                    sKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
                ElseIf InStr(sDummy, "RunServices:") = 1 Then
                    sKey = "Software\Microsoft\Windows\CurrentVersion\RunServices"
                ElseIf InStr(sDummy, "RunServicesOnce:") = 1 Then
                    sKey = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
                Else
                    If InStr(1, sDummy, "Policies\", vbTextCompare) = 0 Then
                        sKey = "Software\Microsoft\Windows\CurrentVersion\Run"
                    Else
                        sKey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
                    End If
                End If
                sDummy = Mid(sDummy, InStr(sDummy, "[") + 1)
                sVal = Left(sDummy, InStrRev(sDummy, "]") - 1)
                sData = Mid(sDummy, InStrRev(sDummy, "]") + 2)
                
                If lHive <> HKEY_USERS Then
                    RegSetStringVal lHive, sKey, sVal, sData
                Else
                    sData = Left(sData, InStr(sData, "(User '") - 2)
                    RegSetStringVal lHive, sSID & "\" & sKey, sVal, sData
                End If
            Else
                'O4 - Startup: bla.lnk = c:\bla.exe
                'backup file is sPath & "backups\" & sBackup & "-" & filename
                sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
                If InStr(sDummy, " = ") > 0 Then
                    sDummy = Left(sDummy, InStr(sDummy, " = ") - 1)
                End If
                sData = Mid(sItem, 6)
                sData = Left(sData, InStr(sData, ": ") - 1)
                Select Case sData
                    Case "Startup":                sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
                    Case "User Startup":           sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
                    Case "Global Startup":         sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common Startup")
                    Case "Global User Startup":    sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common Startup")
                    Case "Global User AltStartup": sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common AltStartup")
                End Select
                If sData <> vbNullString Then
                    If (GetAttr(sData) And vbDirectory) Then
                        CopyFolder sPath & "backups\" & sBackup & "-" & sDummy, sData & IIf(Right(sData, 1) = "\", "", "\") & sDummy
                    Else
                        On Error Resume Next
                        FileCopy sPath & "backups\" & sBackup & "-" & sDummy, sData & IIf(Right(sData, 1) = "\", "", "\") & sDummy
                        On Error GoTo Error:
                    End If
                End If
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            If FileExists(sData) Then
                If (GetAttr(sData) And vbDirectory) Then
                    DeleteFolder sPath & "backups\" & sFile & "-" & sDummy
                Else
                    If FileExists(sPath & "backups\" & sFile & "-" & sDummy) Then DeleteFile sPath & "backups\" & sFile & "-" & sDummy
                End If
            Else
                If FileExists(sPath & "backups\" & sFile & "-" & sDummy) Then DeleteFile sPath & "backups\" & sFile & "-" & sDummy
            End If
            
        Case "O5" 'Control.ini IE Options block
            'O5 - control.ini: inetcpl.cpl=no
            
            'WritePrivateProfileString "don't load", "inetcpl.cpl", "no", "control.ini"
            IniSetString "control.ini", "don't load", "inetcpl.cpl", "no"
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
                        
        Case "O6", "O7", "O9", "O11", "O12"
            'Policies IE Options/Control Panel block
            'O6 - HKCU\Software\Policies\Microsoft\Internet Explorer\Restrictions present
            'Policies Regedit block
            'O7 - HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System, DisableRegedit=1
            'Extra IE 'Tools' menuitem / button
            'O9 - Extra button: Offline
            'O9 - Extra 'Tools' menuitem: Add to T&rusted Zone
            'IE Advanced Options group
            'O11 - Options group: [COMMONNAME] CommonName
            'IE Plugin
            'O12 - Plugin for .spop: NAV.DLL
            'O12 - Plugin for text/html: NAV.DLL
            
            Open sPath & "backups\" & sFile For Input As #1
                Line Input #1, sDummy
                Line Input #1, sDummy
                sMyFile = vbNullString
                Do
                    Line Input #1, sDummy
                    sMyFile = sMyFile & sDummy & vbCrLf
                Loop Until EOF(1)
            Close #1
            
            'regedit in 2000/XP tends to prefix ÿþ
            'to .reg files - they won't merge then
            If Left(sMyFile, 2) = "ÿþ" Then sMyFile = Mid(sMyFile, 3)
            
            Open sPath & "backups\" & sFile & ".reg" For Output As #1
                Print #1, sMyFile
            Close #1
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFile sPath & "backups\" & sFile & ".reg"
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O8" 'IE Context menuitem
            'O8 - Extra context menu item: &Title - C:\Windows\web\dummy.htm
            sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
            sMyFile = Mid(sDummy, InStr(sDummy, " - ") + 3)
            sDummy = Left(sDummy, InStr(sDummy, " - ") - 1)
            
            RegCreateKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & sDummy
            RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & sDummy, vbNullString, sMyFile
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
                        
        Case "O10" 'Winsock hijack
            'O10 - Broken Internet access because of missing LSP provider: 'file'
            'O10 - Broken Internet access because of LSP chain gap (#2 in chain of 8 missing)"
            
            'should not trigger
                                                
        Case "O13" 'IE DefaultPrefix hijack
            'O13 - DefaultPrefix: http://www.prolivation.com/cgi?
            'O13 - WWW Prefix: http://www.prolivation.com/cgi?
            
            sMyName = Mid(sItem, InStr(sItem, ": ") + 2)
            sDummy = Mid(sItem, 7)
            sDummy = Left(sDummy, InStr(sDummy, ": ") - 1)
            Select Case sDummy
                Case "DefaultPrefix": RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\DefaultPrefix", "", sMyName
                Case "WWW Prefix":    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "www", sMyName
                Case "WWW. Prefix":   RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "www.", sMyName
                Case "Home Prefix":   RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "home", sMyName
                Case "Mosaic Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "mosaic", sMyName
                Case "FTP Prefix":    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "ftp", sMyName
                Case "Gopher Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "gopher", sMyName
            End Select
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O14" 'IERESET.INF hijack
            'O14 - IERESET.INF: START_PAGE_URL="http://www.searchalot.com"
            
            'get value + URL to revert from sItem
            sName = Mid(sItem, InStr(sItem, ": ") + 2)
            sDummy = Mid(sItem, InStr(sItem, "=") + 1)
            sName = Left(sName, InStr(sName, "=") - 1)
            If sName <> "SearchAssistant" And sName <> "CustomizeSearch" Then sName = sName & "="
            
            sMyFile = vbNullString
            Open sWinDir & "\INF\iereset.inf" For Input As #1
                Do
                    Line Input #1, sMyName
                    If InStr(sMyName, sName) > 0 Then
                        Select Case sName
                            Case "SearchAssistant": sMyFile = sMyFile & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""SearchAssistant"",0,""" & sDummy & """" & vbCrLf
                            Case "CustomizeSearch": sMyFile = sMyFile & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""CustomizeSearch"",0,""" & sDummy & """" & vbCrLf
                            Case "START_PAGE_URL=": If InStr(sMyName, "MS_START_PAGE_URL=") = 0 Then sMyFile = sMyFile & "START_PAGE_URL=""" & sDummy & """" & vbCrLf
                            Case "SEARCH_PAGE_URL=": sMyFile = sMyFile & "SEARCH_PAGE_URL=""" & sDummy & """" & vbCrLf
                            Case "MS_START_PAGE_URL=": sMyFile = sMyFile & "MS_START_PAGE_URL=""" & sDummy & """" & vbCrLf
                        End Select
                    Else
                        sMyFile = sMyFile & sMyName & vbCrLf
                    End If
                Loop Until EOF(1)
            Close #1
            If FileExists(sWinDir & "\INF\iereset.inf") Then DeleteFile sWinDir & "\INF\iereset.inf"
            Open sWinDir & "\INF\iereset.inf" For Output As #1
                Print #1, sMyFile
            Close #1
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O15" 'Trusted Zone Autoadd
            'O15 - Trusted Zone: http://free.aol.com (HKLM)
            'O15 - Trusted IP range: http://66.66.66.* (HKLM)
            'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)
            
            sRegKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\"
            sRegKey2 = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\"
            sRegKey3 = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\ProtocolDefaults"
            sRegKey4 = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\"
            sRegKey5 = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscRanges\"
            
            sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
            If InStr(sItem, "ProtocolDefaults:") > 0 Then GoTo ProtDefs:
            If InStr(sDummy, "//") > 0 Then sDummy = Mid(sDummy, InStr(sDummy, "//") + 2)
            If InStr(sDummy, "*.") > 0 Then
                sDummy = Mid(sDummy, InStr(sDummy, "*.") + 2)
                If InStr(sDummy, ".") <> InStrRev(sDummy, ".") Then sDummy = "*." & sDummy
            End If
            
            If InStr(sItem, " (HKLM)") > 0 Then
                lHive = HKEY_LOCAL_MACHINE
                sDummy = Left(sDummy, InStr(sDummy, " (HKLM)") - 1)
            Else
                lHive = HKEY_CURRENT_USER
            End If
            
            If InStr(sItem, "http://") > 0 Then
                sVal = "http"
            Else
                sVal = "*"
            End If
            
            If InStr(sItem, "Trusted Zone:") > 0 Then
                If InStr(sDummy, ".") = InStrRev(sDummy, ".") Then
                    'domain.com
                    sKey2 = sDummy
                    sKey1 = vbNullString
                Else
                    'sub.domain.com
                    i = InStrRev(sDummy, ".")
                    i = InStrRev(sDummy, ".", i - 1)
                    If DomainHasDoubleTLD(sDummy) Then
                        i = InStrRev(sDummy, ".", i - 1)
                    End If
                    sKey2 = Mid(sDummy, i + 1)
                    sKey1 = sKey2 & "\" & Left(sDummy, i - 1)
                End If
                If InStr(sItem, "ESC Trusted") = 0 Then
                    RegCreateKey lHive, sRegKey & sKey2
                    RegCreateKey lHive, sRegKey & sKey1
                    RegSetDwordVal lHive, sRegKey & sKey2, sVal, 2
                Else
                MsgBox sRegKey & sKey2
                MsgBox sRegKey & sKey1
                    RegCreateKey lHive, sRegKey4 & sKey2
                    RegCreateKey lHive, sRegKey4 & sKey1
                    RegSetDwordVal lHive, sRegKey4 & sKey2, sVal, 2
                End If
            Else
                If InStr(sItem, "ESC Trusted") = 0 Then
                    For i = 1 To 9999
                        If Not RegKeyExists(lHive, sRegKey2 & "Range" & i) Then
                            RegCreateKey lHive, sRegKey2 & "Range" & i
                            RegSetDwordVal lHive, sRegKey2 & "Range" & i, sVal, 2
                            RegSetStringVal lHive, sRegKey2 & "Range" & i, ":Range", sDummy
                            Exit For
                        End If
                    Next i
                Else
                    For i = 1 To 9999
                        If Not RegKeyExists(lHive, sRegKey5 & "Range" & i) Then
                            RegCreateKey lHive, sRegKey5 & "Range" & i
                            RegSetDwordVal lHive, sRegKey5 & "Range" & i, sVal, 2
                            RegSetStringVal lHive, sRegKey5 & "Range" & i, ":Range", sDummy
                            Exit For
                        End If
                    Next i
                End If
                If i = 10000 Then
                    'problem!
                    MsgBox "Unable to restore this backup: too many items in your Trusted Zone!", vbCritical
                    Exit Sub
                End If
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            Exit Sub
            
ProtDefs:
            'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)
            Dim sProt$, sZone$, lZone&
            sProt = Mid(sItem, InStr(sItem, ": ") + 3)
            sProt = Left(sProt, InStr(sProt, "'") - 1)
            sZone = Mid(sItem, InStr(sItem, "is in ") + 6)
            sZone = Left(sZone, InStr(sZone, ",") - 1)
            Select Case sZone
                Case "My Computer Zone": lZone = 0
                Case "Intranet Zone": lZone = 1
                Case "Trusted Zone": lZone = 2
                Case "Internet Zone": lZone = 3
                Case "Restricted Zone": lZone = 4
                Case Else
                    MsgBox "Unable to restore item: Protocol '" & sProt & "' was set to unknown zone.", vbExclamation
                    Exit Sub
            End Select
            If InStr(sItem, "(HKLM)") > 0 Then
                lHive = HKEY_LOCAL_MACHINE
            Else
                lHive = HKEY_CURRENT_USER
            End If
            RegSetDwordVal lHive, sRegKey3, sProt, lZone
            
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O16" ' - Download Program Files item
            'O16 - DPF: Plugin name - http://bla.com/bla.cab
            'O16 - DPF: {0000} (name) - http://bla.com/bla.cab
            
            'backup has extra info with reg data
            sData = vbNullString
            Open sPath & "backups\" & sFile For Input As #1
                Line Input #1, sLine
                Line Input #1, sLine
                Do
                    Line Input #1, sLine
                    sData = sData & sLine & vbCrLf
                Loop Until EOF(1)
            Close #1
            
            'regedit in 2000/XP tends to prepend ÿþ
            'to .reg files - they won't merge then
            If Left(sData, 2) = "ÿþ" Then sData = Mid(sData, 3)
            
            Open sPath & "backups\" & sFile & ".reg" For Output As #1
                Print #1, sData
            Close #1
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFile sPath & "backups\" & sFile & ".reg"
            
            'restore all the files
            sDPFKey = "Software\Microsoft\Code Store Database\Distribution Units"
            sCLSID = Mid(sName, 12)
            If Left(sCLSID, 1) = "{" Then
                sCLSID = Left(sCLSID, InStr(sCLSID, "}"))
            Else
                sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
            End If
            sINF = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sCLSID & "\DownloadInformation", "INF")
            If sINF <> vbNullString Then
                If FileExists(sPath & "backups\" & sFile & ".inf") Then
                    On Error Resume Next
                    FileCopy sPath & "backups\" & sFile & ".inf", sINF
                    On Error GoTo Error:
                End If
            End If
            sOSD = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sCLSID & "\DownloadInformation", "OSD")
            If sOSD <> vbNullString Then
                If FileExists(sPath & "backups\" & sFile & ".osd") Then
                    On Error Resume Next
                    FileCopy sPath & "backups\" & sFile & ".osd", sOSD
                    On Error GoTo Error:
                End If
            End If
            sInProcServer32 = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", "")
            If sInProcServer32 <> vbNullString Then
                If FileExists(sPath & "backups\" & sFile & ".dll") Then
                    On Error Resume Next
                    FileCopy sPath & "backups\" & sFile & ".dll", LCase(sInProcServer32)
                    On Error GoTo Error:
                    Shell sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\regsvr32.exe /s """ & sInProcServer32 & """", vbHide
                End If
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            If FileExists(sPath & "backups\" & sFile & ".dll") Then DeleteFile sPath & "backups\" & sFile & ".dll"
            If FileExists(sPath & "backups\" & sFile & ".inf") Then DeleteFile sPath & "backups\" & sFile & ".inf"
            If FileExists(sPath & "backups\" & sFile & ".osd") Then DeleteFile sPath & "backups\" & sFile & ".osd"
            
        Case "O17" 'Domain hijack
            'O17 - HKLM\Software\..\Telephony: DomainName = blah
            'O17 - HKLM\System\CS1\Services\Tcpip\..\{00000}: Domain = blah
            
            sVal = Mid(sItem, InStrRev(sItem, ": ") + 2)
            sData = Mid(sVal, InStr(sVal, " = ") + 3)
            sVal = Left(sVal, InStr(sVal, " = ") - 1)
            sKey = Mid(sItem, 12)
            sKey = Left(sKey, InStr(sKey, ": ") - 1)
            sKey = Replace(sKey, "\CCS\", "\CurrentControlSet\")
            If InStr(sKey, "System\CS") > 0 Then
                For i = 1 To 20
                    sKey = Replace(sKey, "System\CS" & CStr(i), "System\ControlSet" & String(3 - Len(CStr(i)), "0") & CStr(i))
                Next i
            End If
            If InStr(sKey, "\..\") > 0 Then
                If InStr(sKey, "Software\") > 0 Then
                    sKey = Replace(sKey, "\..\", "\Microsoft\Windows\CurrentVersion\")
                ElseIf InStr(sKey, "\Tcpip\..\") > 0 Then
                    sKey = Replace(sKey, "\Tcpip\..\", "\Tcpip\Parameters\Interfaces\")
                End If
            End If
            RegSetStringVal HKEY_LOCAL_MACHINE, sKey, sVal, sData
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O18" 'Protocol
            'O18 - Protocol: cn - {0000000000}
            'O18 - Protocol hijack: res - {000000000}
            'O18 - Filter: text/html - {000} - file.dll
            'O18 - Filter hijack: text/xml - {000} - file.dll
            sCLSID = Mid(sItem, InStr(sItem, " - {") + 3)
            sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
            sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
            sDummy = Left(sDummy, InStr(sDummy, " - {") - 1)
            If InStr(sItem, "Protocol: ") > 0 Then
                RegCreateKey HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy
                RegSetStringVal HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy, "CLSID", sCLSID
            ElseIf InStr(sItem, "Filter: ") > 0 Then
                RegCreateKey HKEY_CLASSES_ROOT, "Protocols\Filter\" & sDummy
                RegSetStringVal HKEY_CLASSES_ROOT, "Protocols\Filter\" & sDummy, "CLSID", sCLSID
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O19" 'user stylesheet
            'O19 - User stylesheet: c:\file.css (file missing) (HKLM)
            
            sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
            If InStr(sDummy, " (HKLM)") = 0 Then
                lHive = HKEY_CURRENT_USER
            Else
                lHive = HKEY_LOCAL_MACHINE
            End If
            If InStr(sDummy, "(file missing)") > 0 Then
                sDummy = Left(sDummy, InStr(sDummy, " (file missing)") - 1)
            End If
            RegSetDwordVal lHive, "Software\Microsoft\Internet Explorer\Styles", "Use My Stylesheet", 1
            RegSetStringVal lHive, "Software\Microsoft\Internet Explorer\Styles", "User Stylesheet", sDummy
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
        
        Case "O20" 'appinit_dlls
            'O20 - AppInit_DLLs: file.dll
            'O20 - Winlogon Notify: blaat - c:\file.dll
            If InStr(sItem, "AppInit_DLLs") > 0 Then
                sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
                sDummy = Replace(sDummy, "|", Chr(0))
                RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", sDummy
            Else
                'backup has extra reg data
                sData = vbNullString
                Open sPath & "backups\" & sFile For Input As #1
                    Line Input #1, sLine
                    Line Input #1, sLine
                    Do
                        Line Input #1, sLine
                        sData = sData & sLine & vbCrLf
                    Loop Until EOF(1)
                Close #1
                If Left(sData, 2) = "ÿþ" Then sData = Mid(sData, 3)
                Open sPath & "backups\" & sFile & ".reg" For Output As #1
                    Print #1, sData
                Close #1
                Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
                DoEvents
                If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFile sPath & "backups\" & sFile & ".reg"
                    
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O21" 'ssodl
            'O21 - SSODL: webcheck - {000....000} - c:\file.dll
            'todo:
            'get, print and merge .reg data for clsid regkey
            'reconstruct reg value at SSODL regkey
            
            sName = Mid(sItem, InStr(sItem, ": ") + 2)
            sCLSID = Mid(sName, InStr(sName, " - ") + 3)
            sName = Left(sName, InStr(sName, " - ") - 1)
            sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad", sName, sCLSID
            
            'backup has extra info with reg data
            sData = vbNullString
            Open sPath & "backups\" & sFile For Input As #1
                Line Input #1, sLine
                Line Input #1, sLine
                Do
                    Line Input #1, sLine
                    sData = sData & sLine & vbCrLf
                Loop Until EOF(1)
            Close #1
            
            'regedit in 2000/XP tends to prepend ÿþ
            'to .reg files - they won't merge then
            If Left(sData, 2) = "ÿþ" Then sData = Mid(sData, 3)
            
            Open sPath & "backups\" & sFile & ".reg" For Output As #1
                Print #1, sData
            Close #1
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFile sPath & "backups\" & sFile & ".reg"
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O22" 'sharedtaskscheduler
            'O22 - SharedTaskScheduler: blah - {000...000} - file.dll
            'todo:
            'restore sts regval
            'restore clsid regkey
            
            sName = Mid(sItem, InStr(sItem, ": ") + 2)
            sCLSID = Mid(sName, InStr(sName, " - ") + 3)
            sName = Left(sName, InStr(sName, " - ") - 1)
            sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler", sCLSID, sName
            
            'backup has extra info with reg data
            sData = vbNullString
            Open sPath & "backups\" & sFile For Input As #1
                Line Input #1, sLine
                Line Input #1, sLine
                Do
                    Line Input #1, sLine
                    sData = sData & sLine & vbCrLf
                Loop Until EOF(1)
            Close #1
            
            'regedit in 2000/XP tends to prepend ÿþ
            'to .reg files - they won't merge then
            If Left(sData, 2) = "ÿþ" Then sData = Mid(sData, 3)
            
            Open sPath & "backups\" & sFile & ".reg" For Output As #1
                Print #1, sData
            Close #1
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFile sPath & "backups\" & sFile & ".reg"
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
        
        Case "O23"
            'O23 - Service: bla bla - blacorp - bla.exe
            'todo:
            'enable & start service
            Dim sServices$(), sDisplayName$
            sDisplayName = Mid(sItem, InStr(sItem, ": ") + 2)
            sDisplayName = Left(sDisplayName, InStr(sDisplayName, " - ") - 1)
            sServices = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
            If UBound(sServices) <> 0 And UBound(sServices) <> -1 Then
                For i = 0 To UBound(sServices)
                    If sDisplayName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServices(i), "DisplayName") Then
                        sName = sServices(i)
                        RegSetDwordVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start", 2
                        Shell sWinSysDir & "\NET.exe START """ & sDisplayName & """", vbHide
                        Exit For
                    End If
                Next i
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            
        Case "O24"
            'O24 - Desktop Component N: blah - c:\windows\index.html
            'todo:
            'restore reg key
            Dim sSource$
            'copy file back to Source and SubscribedURL
            sSource = Mid(sItem, InStr(sItem, ":"))
            sSource = Mid(sSource, InStr(sSource, " - ") + 3)
            If sSource <> "(no file)" Then
                'only one is backed up if they are the same
                If FileExists(sPath & "backups\" & sFile & "-source.html") Then
                    FileCopy sPath & "backups\" & sFile & "-source.html", sSource
                End If
                If FileExists(sPath & "backups\" & sFile & "-suburl.html") Then
                    FileCopy sPath & "backups\" & sFile & "-suburl.html", sSource
                End If
            End If
            
            'backup has extra info with reg data
            sData = vbNullString
            Open sPath & "backups\" & sFile For Input As #1
                Line Input #1, sLine
                Line Input #1, sLine
                Do
                    Line Input #1, sLine
                    sData = sData & sLine & vbCrLf
                Loop Until EOF(1)
            Close #1
            
            'regedit in 2000/XP tends to prepend ÿþ
            'to .reg files - they won't merge then
            If Left(sData, 2) = "ÿþ" Then sData = Mid(sData, 3)
            
            Open sPath & "backups\" & sFile & ".reg" For Output As #1
                Print #1, sData
            Close #1
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFile sPath & "backups\" & sFile & ".reg"
            If FileExists(sPath & "backups\" & sFile) Then DeleteFile sPath & "backups\" & sFile
            If FileExists(sPath & "backups\" & sFile & "-source.html") Then DeleteFile sPath & "backups\" & sFile & "-source.html"
            If FileExists(sPath & "backups\" & sFile & "-suburl.html") Then DeleteFile sPath & "backups\" & sFile & "-suburl.html"
            
    End Select
    Exit Sub
    
Error:
    Close #1
    ErrorMsg "modBackup_RestoreBackup", Err.Number, Err.Description, "sItem=" & sItem
End Sub
Public Sub ListBackups()
    Dim sPath$, sFile$, vDummy As Variant
    Dim sBackup$, sDate$, sTime$
    On Error GoTo Error:
    sPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
    sFile = Dir(sPath & "backups\" & "backup*")
    If sFile = vbNullString Then Exit Sub
    frmMain.lstBackups.Clear
    
    Do
        vDummy = Split(sFile, "-")
        If InStr(sFile, ".") = 0 And UBound(vDummy) = 3 Then
            'backup-20021024-181841-901
            '0      1        2      3
            'duh    date     time   random
            
            Open sPath & "backups\" & sFile For Input As #1
                Line Input #1, sBackup
            Close #1
            
            sDate = Right(vDummy(1), 2) & "-" & Mid(vDummy(1), 5, 2) & "-" & Mid(vDummy(1), 1, 4)
            sTime = Left(vDummy(2), 2) & ":" & Mid(vDummy(2), 3, 2) & ":" & Right(vDummy(2), 2)
            
            sBackup = Format(sDate, "Short Date") & ", " & _
                      Format(sTime, "Long Time") & ": " & _
                      sBackup
            frmMain.lstBackups.AddItem sBackup
        End If
        sFile = Dir
    Loop Until sFile = vbNullString
    Exit Sub
    
Error:
    ErrorMsg "modBackup_ListBackups", Err.Number, Err.Description
End Sub

Public Sub DeleteBackup(sBackup$)
    If sBackup = vbNullString Then
        On Error Resume Next
        DeleteFile App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "backups\backup-*.*"
        Exit Sub
    End If
    
    Dim sFile$, sDate$, sTime$
    On Error GoTo Error:
    
    sDate = Left(sBackup, InStr(sBackup, ", ") - 1)
    sTime = Mid(sBackup, InStr(sBackup, ", ") + 2)
    sTime = Left(sTime, InStr(sTime, ": ") - 1)
    
    If Not bIsUSADateFormat Then
        sDate = Format(sDate, "yyyymmdd")
    Else
        'use stupid workaround for USA date format
        sDate = Format(sDate, "yyyyddmm")
    End If
    sTime = Format(sTime, "HhNnSs")
    
    sFile = "backup-" & sDate & "-" & sTime & "*.*"
    
    On Error Resume Next
    DeleteFile App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "backups\" & sFile
    Exit Sub
    
Error:
    ErrorMsg "modBackup_DeleteBackup", Err.Number, Err.Description, "sBackup=" & sBackup
End Sub

Public Function GetCLSIDOfMSIEExtension(ByVal sName$, bButtonOrMenu As Boolean)
    Dim hKey&, i%, sCLSID$
    On Error GoTo Error:
    sName = Left(sName, InStr(sName, " (HK") - 1)
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sCLSID = String(255, 0)
        If RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, 0, ByVal 0) <> 0 Then
            RegCloseKey hKey
            GetCLSIDOfMSIEExtension = vbNullString
            Exit Function
        End If
        Do
            sCLSID = Left(sCLSID, InStr(sCLSID, Chr(0)) - 1)
            If bButtonOrMenu Then
                If sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "ButtonText") Then
                    GetCLSIDOfMSIEExtension = sCLSID
                    Exit Do
                End If
            Else
                If sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "MenuText") Then
                    GetCLSIDOfMSIEExtension = sCLSID
                    Exit Do
                End If
            End If
            
            sCLSID = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sCLSID = String(255, 0)
        If RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, 0, ByVal 0) <> 0 Then
            RegCloseKey hKey
            GetCLSIDOfMSIEExtension = vbNullString
            Exit Function
        End If
        Do
            sCLSID = Left(sCLSID, InStr(sCLSID, Chr(0)) - 1)
            If bButtonOrMenu Then
                If sName = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "ButtonText") Then
                    GetCLSIDOfMSIEExtension = sCLSID
                    Exit Do
                End If
            Else
                If sName = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "MenuText") Then
                    GetCLSIDOfMSIEExtension = sCLSID
                    Exit Do
                End If
            End If
            
            sCLSID = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    Exit Function
    
Error:
    RegCloseKey hKey
    ErrorMsg "modBackup_GetCLSIDOfMSIEExtension", Err.Number, Err.Description, "sName=" & sName & ",bButtonOrMenu=" & CStr(bButtonOrMenu)
End Function
