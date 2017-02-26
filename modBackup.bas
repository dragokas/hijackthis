Attribute VB_Name = "modBackup"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long

Public Sub MakeBackup(ByVal sItem$)
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "MakeBackup - Begin", sItem

    Dim sPath$, lHive&, sKey$, sValue$, sSID$, sUsername$
    Dim sData$, sDummy$, sBackup$, sLine$
    Dim sDPFKey$, sCLSID$, sOSD$, sInf$, sInProcServer32$
    Dim sNum$, sFile1$, sFile2$, ff%, sPrefix$, Wow64Redir As Boolean
    Dim sFullPrefix$
    
    sPath = AppPath() & IIf(Right$(AppPath(), 1) = "\", vbNullString, "\")
    
    If bNoWriteAccess Then Exit Sub
    If Not FolderExists(sPath) Then Exit Sub 'running from zip in XP
    If Not FolderExists(sPath & "backups") Then MkDir sPath & "backups"
    If Not FolderExists(sPath & "backups") Then
        'MsgBoxW "Unable to create folder to place backups in. Backups of fixed items cannot be saved!", vbExclamation
        MsgBoxW Translate(530), vbExclamation
        bNoWriteAccess = True
        Exit Sub
    End If
    
    'create backup file name
    Randomize
    sBackup = "backup-" & Format(Date, "yyyymmdd") & "-" & Format(time, "HhNnSs") & "-" & CStr(1000 * Format(Rnd(), "0.000"))
    If DirW$(sPath & "backups\" & sBackup & "*.*") <> vbNullString Or _
       InStrRev(sBackup, "-") <> Len(sBackup) - 3 Then
        Do
            sBackup = "backup-" & Format(Date, "yyyymmdd") & "-" & Format(time, "HhNnSs") & "-"
            Randomize
            sBackup = sBackup & CStr(1000 * Format(Rnd(), "0.000"))
        Loop Until DirW$(sPath & "backups\" & sBackup & "*.*") = vbNullString And _
                   InStrRev(sBackup, "-") = Len(sBackup) - 3
    End If
    
    sFullPrefix = Left$(sItem, InStr(sItem, " - ") - 1)
    sPrefix = Trim(Left$(sItem, InStr(sItem, "-") - 1))
    
    If StrEndWith(sFullPrefix, "-32") Then Wow64Redir = True
    
    On Error GoTo ErrorHandler:
    Select Case sPrefix 'Trim$(Left$(sItem, 3))
        'these lot don't need any additional stuff
        'backed up, everything is in the sItem line
        Case "R0", "R3"
        Case "F0", "F1", "F2", "F3"
        'Case "N1", "N2", "N3", "N4"
        Case "O1", "O3", "O5"
        Case "O7", "O8", "O13", "O14"
        Case "O15", "O17", "O18", "O19"
        Case "O23", "O25"
        
        'below items that DO need something else
        'backed up
        
        Case "R1" ', "R1-32"
            'R1 - Created Registry value
            'R1 - HKCU\Software\..\Subkey,Value[=Data]
            
            'need to get sData if not in sItem
            If InStr(sItem, "=") = 0 Or Right$(sItem, 1) = "=" Then
                sDummy = Mid$(sItem, 6)
                Select Case Left$(sDummy, 4)
                    Case "HKCU": lHive = HKEY_CURRENT_USER
                    Case "HKCR": lHive = HKEY_CLASSES_ROOT
                    Case "HKLM": lHive = HKEY_LOCAL_MACHINE
                    Case "HKU\": lHive = HKEY_USERS
                End Select
                sDummy = Mid$(sDummy, 6)
                sKey = Left$(sDummy, InStr(sDummy, ",") - 1)
                sValue = Mid$(sDummy, InStr(sDummy, ",") + 1)
                If InStr(sValue, "=") > 0 Then sValue = Left$(sValue, InStr(sValue, "=") - 1)
                sData = RegGetString(lHive, sKey, sValue, Wow64Redir)
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
            
            'MsgBoxW "Not implemented yet, item '" & sItem & "' will not be backed up!", vbExclamation, "bad coder - no donuts"
            MsgBoxW Replace$(Translate(531), "[]", sItem), vbExclamation, Translate(532)
            
        Case "O2" ', "O2-32"
            'O2 - BHO
            'O2 - BHO: BhoName - CLSID - Filename
            
            'backup BHO dll
            Dim vDummy As Variant
            vDummy = Split(sItem, " - ")
            If UBound(vDummy) <> 3 Then
                If InStr(sItem, "}") > 0 And _
                   InStr(sItem, "- ") > 0 Then
                    sDummy = Mid$(sItem, InStr(InStr(sItem, "}"), sItem, "- ") + 2)
                End If
            Else
                sDummy = CStr(vDummy(3))
            End If
            If FileExists(sDummy) Then FileCopyW sDummy, sPath & "backups\" & sBackup & ".dll"
            
        Case "O4" ', "O4-32"
            'O4 - Regrun and Startup run, also for other users
            'O4 - Common Startup: Bla.lnk = c:\dummy.exe
            
            '// TODO: MSConfig
            
            If InStr(sItem, "[") = 0 Then
                'need to backup link
                sData = Mid$(sItem, 6)
                If InStr(sItem, " (User '") = 0 Then 'normal item
                    sData = Left$(sData, InStr(sData, ":") - 1)
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
                    sSID = Left$(sData, InStr(sData, " ") - 1)
                    sUsername = MapSIDToUsername(sSID)
                    sData = Mid$(sData, InStr(sData, " ") + 1)
                    sData = Left$(sData, InStr(sData, ":") - 1)
                    Select Case sData
                        Case "Startup":                sData = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
                        Case "AltStartup":             sData = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AltStartup")
                        Case "User Startup":           sData = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
                        Case "User AltStartup":        sData = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AltStartup")
                    End Select
                    If sData <> vbNullString And FolderExists(sData) Then
                        sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
                        If InStr(sDummy, " = ") > 0 Then
                            sDummy = Left$(sDummy, InStr(sDummy, " = ") - 1)
                        End If
                        sData = sData & IIf(Right$(sData, 1) = "\", vbNullString, "\") & sDummy
                        On Error Resume Next
                        If FileExists(sData) Then FileCopyW sData, sPath & "backups\" & sBackup & "-" & sDummy
                        On Error GoTo ErrorHandler:
                        'sdata had no relevant backup data, just dummy
                        sData = vbNullString
                    End If
                End If
                'attempt backup the file, same for both
                If sData <> vbNullString And FolderExists(sData) Then
                    sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
                    If InStr(sDummy, " = ") > 0 Then
                        sDummy = Left$(sDummy, InStr(sDummy, " = ") - 1)
                    End If
                    sData = sData & IIf(Right$(sData, 1) = "\", vbNullString, "\") & sDummy
                    On Error Resume Next
                    If FileExists(sData) Then
                        If (GetFileAttributes(StrPtr(sData)) And vbDirectory) Then
                            CopyFolder sData, sPath & "backups\" & sBackup & "-" & sDummy
                        Else
                            FileCopyW sData, sPath & "backups\" & sBackup & "-" & sDummy
                        End If
                    End If
                    On Error GoTo ErrorHandler:
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
            '"-policy.reg"
            If InStr(sItem, "HKCU") > 0 And InStr(sItem, "Restrictions") > 0 Then
                sData = RegExportKeyToVariable(0&, "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions")
            ElseIf InStr(sItem, "HKCU") > 0 And InStr(sItem, "Control Panel") > 0 Then
                sData = RegExportKeyToVariable(0&, "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Control Panel")
            ElseIf InStr(sItem, "HKLM") > 0 And InStr(sItem, "Restrictions") > 0 Then
                sData = RegExportKeyToVariable(0&, "HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Restrictions")
            ElseIf InStr(sItem, "HKLM") > 0 And InStr(sItem, "Control Panel") > 0 Then
                sData = RegExportKeyToVariable(0&, "HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Control Panel")
            End If
            
        Case "O9" ', "O9-32"
            'O9 - IE Tools menu item/button
            'O9 - Extra button: Offline
            'O9 - Extra 'Tools' menuitem: Add to T&rusted Zone
            
            'O9 - Extra 'Tools' menuitem: Related - {000...000} - c:\file.dll [(HKCU)]
            
            'need to backup all values in regkey
            sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
            sDummy = Mid$(sDummy, InStr(sDummy, " - ") + 3)
            sDummy = Left$(sDummy, InStr(sDummy, " - ") - 1)
            
            'If InStr(sItem, "Extra button:") > 0 Then
            '    sDummy = GetCLSIDOfMSIEExtension(mid$(sItem, InStr(sItem, ":") + 2), True)
            'Else
            '    sDummy = GetCLSIDOfMSIEExtension(mid$(sItem, InStr(sItem, ":") + 2), False)
            'End If
            If sDummy = vbNullString Then Exit Sub
            
            '"-extension.reg"
            If InStr(sItem, " (HKLM)") > 0 Then
                sData = RegExportKeyToVariable(0&, "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\" & sDummy, Wow64Redir)
            Else
                sData = RegExportKeyToVariable(0&, "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Extensions\" & sDummy)
            End If
            
        Case "O10"
            'O10 - Winsock hijack
            'O10 - Broken Internet access because of missing LSP provider: 'file'
            'O10 - Broken Internet access because of LSP chain gap (#2 in chain of 8 missing)"
            
            'backup even possible??
            'msgboxW "Backup of LSP hijackers is not possible " & _
            '       "because of technical limitations. (IOW, " & _
            '       "I don't know how.) Since only two programs " & _
            '       "hijack the LSP (New.Net and WebHancer) and " & _
            '       "both, this should not pose a problem." & vbCrLf & _
            '       "Should you wish to restore either for testing " & _
            '       "purposes or complete insanity, you need to " & _
            '       "reinstall the program.", vbExclamation
            Exit Sub
            
        Case "O11" ', "O11-32"
            'O11 - Extra options in MSIE 'Advanced' settings tab
            'O11 - Options group: [COMMONNAME] CommonName
            
            'need to backup everything in that key
            sDummy = Left$(sItem, InStr(sItem, "]") - 1)
            sDummy = Mid$(sItem, InStr(sItem, "[") + 1)
            '"-advopt.reg"
            sData = RegExportKeyToVariable(0&, "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\AdvancedOptions\" & sDummy, Wow64Redir)
            
        Case "O12" ', "O12-32"
            'O12 - MSIE plugins for file extensions or MIME types
            'O12 - Plugin for .spop: NAV.DLL
            'O12 - Plugin for text/html: NAV.DLL
            
            'need to backup subkey + 'Location' value
            If InStr(sItem, "Plugin for .") > 0 Then
                'plugin for file extension
                sDummy = Left$(sItem, InStr(sItem, ":") - 1)
                sDummy = Mid$(sDummy, InStr(sDummy, "."))
                sDummy = "Extension\" & sDummy
            Else
                'plugin for MIME type
                sDummy = Left$(sItem, InStr(sItem, ":") - 1)
                sDummy = Mid$(sItem, InStr(sItem, " for ") + 5)
                sDummy = "MIME\" & sDummy
            End If
            '"-plugin.reg"
            sData = RegExportKeyToVariable(0&, "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Plugins\" & sDummy, Wow64Redir)
            
        Case "O16" ', "O16-32"
            'O16 - Download Program Files item
            'O16 - DPF: Plugin name - http://bla.com/bla.cab
            'O16 - DPF: {000000} (name) - http://bla.com/bla.cab
            
            'need to export key from HKLM\..\Dist Units
            'and (if applic) HKCR\CLSID\{0000}
            'need to backup files OSD, INF, InProcServer32
            
            sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
            If Left$(sDummy, 1) = "{" Then
                'name is CLSID
                sDummy = Left$(sDummy, InStr(sDummy, "}"))
            Else
                'name is just name
                sDummy = Left$(sDummy, InStr(sDummy, " - ") - 1)
            End If
            '"-dpf1.reg"
            sData = RegExportKeyToVariable(0&, "HKEY_LOCAL_MACHINE\Software\Microsoft\Code Store Database\Distribution Units\" & sDummy, Wow64Redir)
            If Left$(sDummy, 1) = "{" Then
                '"-dpf2.reg"
                sData = sData & RegExportKeyToVariable(0&, "HKEY_CLASSES_ROOT\CLSID\" & sDummy, Wow64Redir, False) 'no header
            End If
            
            sCLSID = sDummy
            sDPFKey = "Software\Microsoft\Code Store Database\Distribution Units"
            'backup INF
            sLine = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sCLSID & "\DownloadInformation", "INF", Wow64Redir)
            If sLine <> vbNullString Then
                If FileExists(sLine) Then
                    FileCopyW sLine, sPath & "backups\" & sBackup & ".inf"
                End If
            End If
            
            'backup OSD
            sLine = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sCLSID & "\DownloadInformation", "OSD", Wow64Redir)
            If sLine <> vbNullString Then
                If FileExists(sLine) Then
                    FileCopyW sLine, sPath & "backups\" & sBackup & ".osd"
                End If
            End If
            
            'backup InProcServer32
            If Left$(sCLSID, 1) = "{" And Right$(sCLSID, 1) = "}" Then
                sLine = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", vbNullString, Wow64Redir)
                If sLine <> vbNullString Then
                    If FileExists(sLine) Then
                        FileCopyW sLine, sPath & "backups\" & sBackup & ".dll"
                    End If
                End If
            End If
            
        Case "O20" ', "O20-32"
            'O20 - AppInit_DLLs: file.dll (do nothing)
            'O20 - Winlogon Notify: bla - c:\file.dll
            'todo:
            'backup regkey
            If InStr(sItem, "Winlogon Notify:") > 0 Then
                sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
                sDummy = Left$(sDummy, InStr(sDummy, " - ") - 1)
                
                '"-notify.reg"
                sData = RegExportKeyToVariable(0&, "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify\" & sDummy, Wow64Redir)
            End If
        
        Case "O21" ', "O21-32"
            'O21 - ShellServiceObjectDelayLoad
            'O21 - SSODL: webcheck - {000....000} - c:\file.dll
            'todo:
            'backup CLSID regkey
            sCLSID = Mid$(sItem, 14)
            sCLSID = Mid$(sCLSID, InStr(sCLSID, " - ") + 3)
            sCLSID = Left$(sCLSID, InStr(sCLSID, " - ") - 1)
            
            '"-ssodl.reg"
            sData = RegExportKeyToVariable(0&, "HKEY_CLASSES_ROOT\CLSID\" & sCLSID, Wow64Redir)

            sDummy = Mid$(sItem, 14)
            sDummy = Left$(sDummy, InStr(sDummy, " - ") - 1)
            '"-ssodl_2.reg"
            sData = sData & RegExportKeyToVariable(0&, "HKLM\Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad\" & sDummy, Wow64Redir, False) 'no header
        
        Case "O22"
            'O22 - ScheduledTask: blah - {000...000} - file.dll
            'todo:
            'backup CLSID regkey
            sCLSID = Mid$(sItem, InStr(sItem, ": ") + 2)
            sCLSID = Mid$(sCLSID, InStr(sCLSID, " - ") + 3)
            sCLSID = Left$(sCLSID, InStr(sCLSID, " - ") - 1)
            
            '"-sts.reg"
            sData = RegExportKeyToVariable(0&, "HKEY_CLASSES_ROOT\CLSID\" & sCLSID)
        
        Case "O24"
            'O24 - Desktop Component N: blah - c:\windows\index.html
            
            sNum = Mid$(sItem, InStr(sItem, ":") - 1, 1)
            sFile1 = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Desktop\Components\" & sNum, "Source")
            sFile2 = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Desktop\Components\" & sNum, "SubscribedURL")
            If LCase$(sFile2) = LCase$(sFile1) Then sFile2 = vbNullString
            
            '"-dc.reg"
            sData = RegExportKeyToVariable(0&, "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Desktop\Components\" & sNum)

            If FileExists(sFile1) Then FileCopyW sFile1, sPath & "backups\" & sBackup & "-source.html"
            If FileExists(sFile2) Then FileCopyW sFile2, sPath & "backups\" & sBackup & "-suburl.html"
            
        Case Else
            'MsgBoxW "I'm so stupid I forgot to implement this. Bug me about it." & vbCrLf & sItem, vbExclamation, "d'oh!"
            MsgBoxW Translate(533) & vbCrLf & sItem, vbExclamation, Translate(534) ':)
    End Select
        
    'winNT/2000/XP reg data workaround
    If Left$(sData, 2) = "ÿþ" Or Left$(sData, 2) = ChrW(&HFF) & ChrW(&HFE) Then sData = Mid$(sData, 3)
    sData = StrConv(sData, vbFromUnicode)
    
    'write item + any data to file
    ff = FreeFile()
    Open sPath & "backups\" & sBackup For Output As #ff
        Print #ff, sItem
        If sData <> vbNullString Then Print #ff, vbCrLf & sData
    Close #ff
    
    AppendErrorLogCustom "MakeBackup - End"
    Exit Sub
ErrorHandler:
    Close #ff
    ErrorMsg Err, "modBackup_MakeBackup", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RestoreBackup(ByVal sItem$)
    'format of backup files:
    'line 1: original item, e.g. O1 - Hosts: auto.search.msn
    'line 2: blank if 3 != blank
    'line 3+: any Registry data
    'format of sItem:
    ' [short date], [long time]: [original item name]
    Dim sPath$, sDate$, stime$, sFile$, sBackup$, sSID$, ff%
    Dim sName$, sDummy$, i&, sKey1$, sKey2$
    Dim sRegKey$, sRegKey2$, sRegKey3$, sRegKey4$, sRegKey5$
    Dim Wow64Redir As Boolean, sPrefix$, sFullPrefix$
    Dim bBackupHasRegData As Boolean, bBackupHasDLL As Boolean
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "RestoreBackup - Begin", sItem
    
    sPath = AppPath() & IIf(Right$(AppPath(), 1) = "\", vbNullString, "\")
    If Not FolderExists(sPath & "backups") Then Exit Sub
    
    sDate = Left$(sItem, InStr(sItem, ", ") - 1)
    stime = Mid$(sItem, InStr(sItem, ", ") + 2)
    sName = Mid$(stime, InStr(stime, ": ") + 2)
    stime = Left$(stime, InStr(stime, ": ") - 1)
    
    sItem = Mid$(sItem, InStr(sItem, ": ") + 2)
    
    If Not bIsUSADateFormat Then
        sDate = Format(sDate, "yyyymmdd")
    Else
        'use stupid workaround for USA data format
        sDate = Format(sDate, "yyyyddmm")
    End If
    stime = Format(stime, "HhNnSs")
    
    'sBackup = "backup-" & sDate & "-" & sTime '& "*.*"
    sBackup = "backup-*.*"
    
    'get first file for this filemask
    'multiple backups can exist, so open each file
    'and check the first line against the item
    'we are looking for
    sFile = DirW$(sPath & "backups\" & sBackup, vbFile)
    If sFile = vbNullString Then
        'note the small difference with the next msg
        'MsgBoxW "The backup files for this item were not found. It could not be restored.", vbExclamation
        MsgBoxW Translate(535), vbExclamation
        
        'msgboxW "DirW$(" & sPath & "backups\" & sBackup & "*.*) = vbNullString"
        Exit Sub
    End If
    Do
        If InStr(sFile, ".") = 0 Then
            ff = FreeFile()
            Open sPath & "backups\" & sFile For Input As #ff
                Line Input #ff, sDummy
                If sDummy = sName Then
                    sBackup = sFile
                    Close #ff
                    Exit Do
                End If
            Close #ff
        End If
        sFile = DirW$()
    Loop Until Len(sFile) = 0
    If sDummy <> sName Then
        'things like this help troubleshooting stupid bugs
        'MsgBoxW "The backup file for this item was not found. It could not be restored.", vbExclamation
        MsgBoxW Translate(536), vbExclamation
        
        'msgboxW "sDummy = " & sDummy & vbCrLf & _
        '       "sName = " & sName
        Exit Sub
    End If
    
    'file types:
    'backup*. = actual backup file /w item name + reg data
    'backup*.dll = backup of BHO file
    If FileExists(sPath & "backups\" & sFile & ".dll") Then bBackupHasDLL = True
    ff = FreeFile()
    Open sPath & "backups\" & sFile For Input As #ff
        Line Input #ff, sDummy
        On Error Resume Next
        sDummy = vbNullString
        Line Input #ff, sDummy
        Line Input #ff, sDummy
        On Error GoTo ErrorHandler:
        If sDummy <> vbNullString Then bBackupHasRegData = True
    Close #ff
    
    Dim lHive&, sKey$, sVal$, sData$
    Dim sIniFile$, sSection$
    Dim sMyFile$, sMyName$, sCLSID$, sLine$
    Dim sDPFKey$, sInf$, sOSD$, sInProcServer32$
    
    'sPrefix = Trim$(Left$(sName, 3))
    sPrefix = Trim(Left$(sItem, InStr(sItem, "-") - 1))
    sFullPrefix = Left$(sItem, InStr(sItem, " - ") - 1)
    
    If StrEndWith(sFullPrefix, "-32") Then Wow64Redir = True
    
    Select Case sPrefix
        Case "R0", "R1" 'Changed/Created Regval
            'R0 - HKCU\Software\..\Subkey,Value=Data
            'R1 - HKCU\Software\..\Subkey,Value=Data
            sDummy = Mid$(sName, 6)
            Select Case Left$(sDummy, 4)
                Case "HKCU": lHive = HKEY_CURRENT_USER
                Case "HKLM": lHive = HKEY_LOCAL_MACHINE
            End Select
            sKey = Mid$(sDummy, 6)
            sVal = Mid$(sKey, InStr(sKey, ",") + 1)
            sKey = Left$(sKey, InStr(sKey, ",") - 1)
            sData = Mid$(sVal, InStr(sVal, " = ") + 3)
            sVal = Left$(sVal, InStr(sVal, " = ") - 1)
            If sVal = "(Default)" Then sVal = vbNullString
            If InStr(sData, " (obfuscated)") > 0 Then
                sData = Left$(sData, InStr(sData, " (obfuscated)") - 1)
            End If
            RegSetStringVal lHive, sKey, sVal, sData
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "R2" 'Created Regkey
            'don't have this yet
            
        Case "R3" 'URLSearchHoook
            'R3 - URLSearchHook: blah - {0000} - bla.dll
            sDummy = Mid$(sName, InStr(sName, "- {") + 2)
            sDummy = Left$(sDummy, InStr(sDummy, "}"))
            RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks", sDummy, vbNullString
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "F0", "F1" 'Changed/Created Inifile val
            'F0 - system.ini: Shell=Explorer.exe openme.exe
            sDummy = Mid$(sName, 6)
            'sMyFile = left$(sDummy, InStr(sDummy, ":") - 1)
            If InStr(sDummy, "system.ini") = 1 Then
                sSection = "boot"
                sMyFile = "system.ini"
            ElseIf InStr(sDummy, "win.ini") = 1 Then
                sSection = "windows"
                sMyFile = "win.ini"
            End If
            sVal = Mid$(sDummy, InStr(sDummy, ": ") + 2)
            'sMyFile = left$(sDummy, InStr(sDummy, ":") - 1)
            sData = Mid$(sVal, InStr(sVal, "=") + 1)
            sVal = Left$(sVal, InStr(sVal, "=") - 1)
            'WritePrivateProfileString sSection, sVal, sData, sMyFile
            IniSetString sMyFile, sSection, sVal, sData
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "F2", "F3"
            'F2 - REG:system.ini: Shell=Explorer.exe blah
            'F2 - REG:system.ini: Userinit=c:\windows\system32\userinit.exe,blah
            'F3 - REG:win.ini: load=blah or run=blah
            sData = Mid$(sName, InStr(sName, "=") + 1)
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
'        Case "N1", "N2", "N3", "N4"
'            'Changed NS4.x homepage
'            'N1 - Netscape 4: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'            'Changed NS6 homepage
'            'N2 - Netscape 6: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'            'Changed NS7 homepage/searchpage
'            'N3 - Netscape 7: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'            'Changed Moz homepage/searchpage
'            'N4 - Mozilla: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'            '               user_pref("browser.search.defaultengine", "http://url"); (c:\..\prefs.js)
'
'            'get user_pref line + prefs.js location
'            sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
'            sMyFile = Mid$(sDummy, InStrRev(sDummy, "(") + 1)
'            sMyFile = Left$(sMyFile, Len(sMyFile) - 1)
'            sDummy = Left$(sDummy, InStrRev(sDummy, "(") - 2)
'
'            If Not FileExists(sMyFile) Then
'                'MsgBoxW "Could not find prefs.js file for Netscape/Mozilla, homepage has not been restored.", vbExclamation
'                MsgBoxW Translate(537), vbExclamation
'                Exit Sub
'            End If
'
'            'read old file, replacing relevant line
'            sData = vbNullString
'            ff = FreeFile()
'            Open sMyFile For Input As #ff
'                Do
'                    Line Input #ff, sLine
'                    If InStr(sLine, sDummy) > 0 Then
'                        sData = sData & sDummy & vbCrLf
'                    Else
'                        sData = sData & sLine & vbCrLf
'                    End If
'                Loop Until EOF(ff)
'            Close #ff
'
'            'write new file
'            If FileExists(sMyFile) Then deletefileWEx (StrPtr(sMyFile))
'            ff = FreeFile()
'            Open sMyFile For Output As #ff
'                Print #ff, sData
'            Close #ff
'
'            If FileExists(sPath & "backups\" & sFile) Then deletefileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O1" 'Hosts file hijack
            'O1 - Hosts file: 66.123.204.8 auto.search.msn.com
            'If InStr(sName, "Hosts file is located at") > 0 Then
            If InStr(sName, Translate(271)) > 0 Then
                sDummy = Mid$(sName, InStr(sName, Translate(271)) + Len(Translate(271)) + 2)
                sDummy = Left$(sDummy, Len(sDummy) - 6)
                RegSetStringVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", sDummy
            Else
                sDummy = Mid$(sName, InStr(sName, ": ") + 2)
                i = GetFileAttributes(StrPtr(sHostsFile))
                If (i And 2048) Then i = i - 2048
                SetFileAttributes StrPtr(sHostsFile), vbNormal
                ff = FreeFile()
                Open sHostsFile For Append As #ff
                    Print #ff, sDummy
                Close #ff
                SetFileAttributes StrPtr(sHostsFile), i
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O2"
            'O2 - BHO: [bhoname] - [clsid] - [file]
            
            sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
            sMyName = Left$(sDummy, InStr(sDummy, " - ") - 1)
            sCLSID = Mid$(sDummy, InStr(sDummy, " - ") + 3)
            sMyFile = Mid$(sCLSID, InStr(sCLSID, " - ") + 3)
            If InStr(sMyFile, "(file missing)") > 0 Then
                sMyFile = Left$(sMyFile, InStr(sMyFile, "(file missing)") - 1)
            End If
            sCLSID = Left$(sCLSID, InStr(sCLSID, " - ") - 1)
            
            RegCreateKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID, Wow64Redir
            If sMyName <> "(no name)" Then RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID, vbNullString, sMyName, Wow64Redir
            If Not RegKeyExists(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", Wow64Redir) Then
                RegCreateKey HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, Wow64Redir
                RegCreateKey HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", Wow64Redir
                If sMyFile <> vbNullString And sMyFile <> "(no file)" Then
                    RegSetStringVal HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, sMyFile, Wow64Redir
                End If
            End If
            
            If Not bBackupHasDLL Then
                'MsgBoxW "BHO file for '" & sName & "' was not found. The Registry data was restored, but the file was not.", vbExclamation
                MsgBoxW Replace$(Translate(538), "[]", sName), vbExclamation
            Else
                'skip errors - app could have restored
                'BHO dll by itself
                On Error Resume Next
                FileCopyW sPath & "backups\" & sBackup & ".dll", sMyFile
                On Error GoTo ErrorHandler:
                If OSver.Bitness = "x64" And FolderExists(sWinDir & "\sysnative") Then
                    Shell sWinDir & "\sysnative\regsvr32.exe /s """ & sMyFile & """", vbHide
                Else
                    Shell sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\regsvr32.exe /s """ & sMyFile & """", vbHide
                End If
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            If FileExists(sPath & "backups\" & sFile & ".dll") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".dll"))
            
        Case "O3"
            'O3 - Toolbar: Radio - {00000000-0000-0000-0000-000000000000}
            
            sMyName = Mid$(sItem, InStr(sItem, ": ") + 2)
            sCLSID = Mid$(sMyName, InStr(sMyName, " - ") + 3)
            sCLSID = Left$(sCLSID, InStr(sCLSID, "}"))
            sMyName = Left$(sMyName, InStr(sMyName, " - ") - 1)
            'If sMyName = "(no name)" Then sMyName = vbNullString
            
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Toolbar", sCLSID, sMyName, Wow64Redir
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O4" 'Regrun entry
            'O4 - HKLM\..\Run: [valuename] rundll shit.dll,LoadEtc
            'O4 - Startup: bla.lnk = c:\bla.exe
            
            'O4 - HKCU\SID\Run: [bla] bla.exe (User 'bla')
            'O4 - SID Startup: bla.lnk = c:\bla.exe (User 'bla')
                
            If InStr(sItem, "[") > 0 Then
                'registry autorun
                sDummy = Mid$(sItem, 6)
                Select Case Left$(sDummy, 4)
                    Case "HKLM": lHive = HKEY_LOCAL_MACHINE
                    Case "HKCU": lHive = HKEY_CURRENT_USER
                    Case "HKU\": lHive = HKEY_USERS
                End Select
                If Not lHive = HKEY_USERS Then
                    sDummy = Mid$(sDummy, 9)
                Else
                    sDummy = Mid$(sDummy, 6)
                    sSID = Left$(sDummy, InStr(sDummy, "\") - 1)
                    sDummy = Mid$(sDummy, Len(sSID) + 5)
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
                sDummy = Mid$(sDummy, InStr(sDummy, "[") + 1)
                sVal = Left$(sDummy, InStrRev(sDummy, "]") - 1)
                sData = Mid$(sDummy, InStrRev(sDummy, "]") + 2)
                
                If lHive <> HKEY_USERS Then
                    RegSetStringVal lHive, sKey, sVal, sData
                Else
                    sData = Left$(sData, InStr(sData, "(User '") - 2)
                    RegSetStringVal lHive, sSID & "\" & sKey, sVal, sData
                End If
            Else
                'O4 - Startup: bla.lnk = c:\bla.exe
                'backup file is sPath & "backups\" & sBackup & "-" & filename
                sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
                If InStr(sDummy, " = ") > 0 Then
                    sDummy = Left$(sDummy, InStr(sDummy, " = ") - 1)
                End If
                sData = Mid$(sItem, 6)
                sData = Left$(sData, InStr(sData, ": ") - 1)
                Select Case sData
                    Case "Startup":                sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
                    Case "User Startup":           sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
                    Case "Global Startup":         sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common Startup")
                    Case "Global User Startup":    sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common Startup")
                    Case "Global User AltStartup": sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common AltStartup")
                End Select
                If sData <> vbNullString Then
                    If (GetFileAttributes(StrPtr(sData)) And vbDirectory) Then
                        CopyFolder sPath & "backups\" & sBackup & "-" & sDummy, sData & IIf(Right$(sData, 1) = "\", vbNullString, "\") & sDummy
                    Else
                        On Error Resume Next
                        FileCopyW sPath & "backups\" & sBackup & "-" & sDummy, sData & IIf(Right$(sData, 1) = "\", vbNullString, "\") & sDummy
                        On Error GoTo ErrorHandler:
                    End If
                End If
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            If FileExists(sData) Then
                If (GetFileAttributes(StrPtr(sData)) And vbDirectory) Then
                    DeleteFolder sPath & "backups\" & sFile & "-" & sDummy
                Else
                    If FileExists(sPath & "backups\" & sFile & "-" & sDummy) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & "-" & sDummy))
                End If
            Else
                If FileExists(sPath & "backups\" & sFile & "-" & sDummy) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & "-" & sDummy))
            End If
            
        Case "O5" 'Control.ini IE Options block
            'O5 - control.ini: inetcpl.cpl=no
            
            'WritePrivateProfileString "don't load", "inetcpl.cpl", "no", "control.ini"
            IniSetString "control.ini", "don't load", "inetcpl.cpl", "no"
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
                        
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
            
            ff = FreeFile()
            Open sPath & "backups\" & sFile For Input As #ff
                Line Input #ff, sDummy
                Line Input #ff, sDummy
                sMyFile = vbNullString
                Do
                    Line Input #ff, sDummy
                    sMyFile = sMyFile & sDummy & vbCrLf
                Loop Until EOF(ff)
            Close #ff
            
            'regedit in 2000/XP tends to prefix ÿþ
            'to .reg files - they won't merge then
            If Left$(sMyFile, 2) = "ÿþ" Then sMyFile = Mid$(sMyFile, 3)
            
            ff = FreeFile()
            Open sPath & "backups\" & sFile & ".reg" For Output As #ff
                Print #ff, sMyFile
            Close #ff
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".reg"))
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O8" 'IE Context menuitem
            'O8 - Extra context menu item: &Title - C:\Windows\web\dummy.htm
            sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
            sMyFile = Mid$(sDummy, InStr(sDummy, " - ") + 3)
            sDummy = Left$(sDummy, InStr(sDummy, " - ") - 1)
            
            RegCreateKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & sDummy
            RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & sDummy, vbNullString, sMyFile
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
                        
        Case "O10" 'Winsock hijack
            'O10 - Broken Internet access because of missing LSP provider: 'file'
            'O10 - Broken Internet access because of LSP chain gap (#2 in chain of 8 missing)"
            
            'should not trigger
                                                
        Case "O13" 'IE DefaultPrefix hijack
            'O13 - DefaultPrefix: http://www.prolivation.com/cgi?
            'O13 - WWW Prefix: http://www.prolivation.com/cgi?
            
            sMyName = Mid$(sItem, InStr(sItem, ": ") + 2)
            sDummy = Mid$(sItem, 7)
            sDummy = Left$(sDummy, InStr(sDummy, ": ") - 1)
            Select Case sDummy
                Case "DefaultPrefix": RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\DefaultPrefix", vbNullString, sMyName
                Case "WWW Prefix":    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "www", sMyName
                Case "WWW. Prefix":   RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "www.", sMyName
                Case "Home Prefix":   RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "home", sMyName
                Case "Mosaic Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "mosaic", sMyName
                Case "FTP Prefix":    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "ftp", sMyName
                Case "Gopher Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\URL\Prefixes", "gopher", sMyName
            End Select
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O14" 'IERESET.INF hijack
            'O14 - IERESET.INF: START_PAGE_URL="http://www.searchalot.com"
            
            'get value + URL to revert from sItem
            sName = Mid$(sItem, InStr(sItem, ": ") + 2)
            sDummy = Mid$(sItem, InStr(sItem, "=") + 1)
            sName = Left$(sName, InStr(sName, "=") - 1)
            If sName <> "SearchAssistant" And sName <> "CustomizeSearch" Then sName = sName & "="
            
            sMyFile = vbNullString
            ff = FreeFile()
            Open sWinDir & "\INF\iereset.inf" For Input As #ff
                Do
                    Line Input #ff, sMyName
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
                Loop Until EOF(ff)
            Close #ff
            If FileExists(sWinDir & "\INF\iereset.inf") Then DeleteFileWEx (StrPtr(sWinDir & "\INF\iereset.inf"))
            ff = FreeFile()
            Open sWinDir & "\INF\iereset.inf" For Output As #ff
                Print #ff, sMyFile
            Close #ff
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O15" 'Trusted Zone Autoadd
            'O15 - Trusted Zone: http://free.aol.com (HKLM)
            'O15 - Trusted IP range: http://66.66.66.* (HKLM)
            'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)
            
            sRegKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\"
            sRegKey2 = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\"
            sRegKey3 = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\ProtocolDefaults"
            sRegKey4 = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\"
            sRegKey5 = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscRanges\"
            
            sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
            If InStr(sItem, "ProtocolDefaults:") > 0 Then GoTo ProtDefs:
            If InStr(sDummy, "//") > 0 Then sDummy = Mid$(sDummy, InStr(sDummy, "//") + 2)
            If InStr(sDummy, "*.") > 0 Then
                sDummy = Mid$(sDummy, InStr(sDummy, "*.") + 2)
                If InStr(sDummy, ".") <> InStrRev(sDummy, ".") Then sDummy = "*." & sDummy
            End If
            
            If InStr(sItem, " (HKLM)") > 0 Then
                lHive = HKEY_LOCAL_MACHINE
                sDummy = Left$(sDummy, InStr(sDummy, " (HKLM)") - 1)
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
                    sKey2 = Mid$(sDummy, i + 1)
                    sKey1 = sKey2 & "\" & Left$(sDummy, i - 1)
                End If
                If InStr(sItem, "ESC Trusted") = 0 Then
                    RegCreateKey lHive, sRegKey & sKey2
                    RegCreateKey lHive, sRegKey & sKey1
                    RegSetDwordVal lHive, sRegKey & sKey2, sVal, 2
                Else
                MsgBoxW sRegKey & sKey2
                MsgBoxW sRegKey & sKey1
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
                    'MsgBoxW "Unable to restore this backup: too many items in your Trusted Zone!", vbCritical
                    MsgBoxW Translate(539), vbCritical
                    Exit Sub
                End If
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            Exit Sub
            
ProtDefs:
            'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)
            Dim sProt$, sZone$, lZone&
            sProt = Mid$(sItem, InStr(sItem, ": ") + 3)
            sProt = Left$(sProt, InStr(sProt, "'") - 1)
            sZone = Mid$(sItem, InStr(sItem, "is in ") + 6)
            sZone = Left$(sZone, InStr(sZone, ",") - 1)
            Select Case sZone
                Case "My Computer Zone": lZone = 0
                Case "Intranet Zone": lZone = 1
                Case "Trusted Zone": lZone = 2
                Case "Internet Zone": lZone = 3
                Case "Restricted Zone": lZone = 4
                Case Else
                    'MsgBoxW "Unable to restore item: Protocol '" & sProt & "' was set to unknown zone.", vbExclamation
                    MsgBoxW Replace$(Translate(540), "[]", sProt), vbExclamation
                    Exit Sub
            End Select
            If InStr(sItem, "(HKLM)") > 0 Then
                lHive = HKEY_LOCAL_MACHINE
            Else
                lHive = HKEY_CURRENT_USER
            End If
            RegSetDwordVal lHive, sRegKey3, sProt, lZone
            
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O16" ' - Download Program Files item
            'O16 - DPF: Plugin name - http://bla.com/bla.cab
            'O16 - DPF: {0000} (name) - http://bla.com/bla.cab
            
            'backup has extra info with reg data
            sData = vbNullString
            ff = FreeFile()
            Open sPath & "backups\" & sFile For Input As #ff
                Line Input #ff, sLine
                Line Input #ff, sLine
                Do
                    Line Input #ff, sLine
                    sData = sData & sLine & vbCrLf
                Loop Until EOF(ff)
            Close #ff
            
            'regedit in 2000/XP tends to prepend ÿþ
            'to .reg files - they won't merge then
            If Left$(sData, 2) = "ÿþ" Then sData = Mid$(sData, 3)
            ff = FreeFile()
            Open sPath & "backups\" & sFile & ".reg" For Output As #ff
                Print #ff, sData
            Close #ff
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".reg"))
            
            'restore all the files
            sDPFKey = "Software\Microsoft\Code Store Database\Distribution Units"
            sCLSID = Mid$(sName, 12)
            If Left$(sCLSID, 1) = "{" Then
                sCLSID = Left$(sCLSID, InStr(sCLSID, "}"))
            Else
                sCLSID = Left$(sCLSID, InStr(sCLSID, " - ") - 1)
            End If
            sInf = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sCLSID & "\DownloadInformation", "INF")
            If sInf <> vbNullString Then
                If FileExists(sPath & "backups\" & sFile & ".inf") Then
                    On Error Resume Next
                    FileCopyW sPath & "backups\" & sFile & ".inf", sInf
                    On Error GoTo ErrorHandler:
                End If
            End If
            sOSD = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sCLSID & "\DownloadInformation", "OSD")
            If sOSD <> vbNullString Then
                If FileExists(sPath & "backups\" & sFile & ".osd") Then
                    On Error Resume Next
                    FileCopyW sPath & "backups\" & sFile & ".osd", sOSD
                    On Error GoTo ErrorHandler:
                End If
            End If
            sInProcServer32 = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", vbNullString)
            If sInProcServer32 <> vbNullString Then
                If FileExists(sPath & "backups\" & sFile & ".dll") Then
                    On Error Resume Next
                    FileCopyW sPath & "backups\" & sFile & ".dll", LCase$(sInProcServer32)
                    On Error GoTo ErrorHandler:
                    Shell sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\regsvr32.exe /s """ & sInProcServer32 & """", vbHide
                End If
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            If FileExists(sPath & "backups\" & sFile & ".dll") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".dll"))
            If FileExists(sPath & "backups\" & sFile & ".inf") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".inf"))
            If FileExists(sPath & "backups\" & sFile & ".osd") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".osd"))
            
        Case "O17" 'Domain hijack
            'O17 - HKLM\Software\..\Telephony: DomainName = blah
            'O17 - HKLM\System\CS1\Services\Tcpip\..\{00000}: Domain = blah
            
            sVal = Mid$(sItem, InStrRev(sItem, ": ") + 2)
            sData = Mid$(sVal, InStr(sVal, " = ") + 3)
            sVal = Left$(sVal, InStr(sVal, " = ") - 1)
            sKey = Mid$(sItem, 12)
            sKey = Left$(sKey, InStr(sKey, ": ") - 1)
            sKey = Replace$(sKey, "\CCS\", "\CurrentControlSet\")
            If InStr(sKey, "System\CS") > 0 Then
                For i = 1 To 20
                    sKey = Replace$(sKey, "System\CS" & CStr(i), "System\ControlSet" & String$(3 - Len(CStr(i)), "0") & CStr(i))
                Next i
            End If
            If InStr(sKey, "\..\") > 0 Then
                If InStr(sKey, "Software\") > 0 Then
                    sKey = Replace$(sKey, "\..\", "\Microsoft\Windows\CurrentVersion\")
                ElseIf InStr(sKey, "\Tcpip\..\") > 0 Then
                    sKey = Replace$(sKey, "\Tcpip\..\", "\Tcpip\Parameters\Interfaces\")
                End If
            End If
            RegSetStringVal HKEY_LOCAL_MACHINE, sKey, sVal, sData
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O18" 'Protocol
            'O18 - Protocol: cn - {0000000000}
            'O18 - Protocol hijack: res - {000000000}
            'O18 - Filter: text/html - {000} - file.dll
            'O18 - Filter hijack: text/xml - {000} - file.dll
            sCLSID = Mid$(sItem, InStr(sItem, " - {") + 3)
            sCLSID = Left$(sCLSID, InStr(sCLSID, " - ") - 1)
            sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
            sDummy = Left$(sDummy, InStr(sDummy, " - {") - 1)
            If InStr(sItem, "Protocol: ") > 0 Then
                RegCreateKey HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy
                RegSetStringVal HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy, "CLSID", sCLSID
            ElseIf InStr(sItem, "Filter: ") > 0 Then
                RegCreateKey HKEY_CLASSES_ROOT, "Protocols\Filter\" & sDummy
                RegSetStringVal HKEY_CLASSES_ROOT, "Protocols\Filter\" & sDummy, "CLSID", sCLSID
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O19" 'user stylesheet
            'O19 - User stylesheet: c:\file.css (file missing) (HKLM)
            
            sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
            If InStr(sDummy, " (HKLM)") = 0 Then
                lHive = HKEY_CURRENT_USER
            Else
                lHive = HKEY_LOCAL_MACHINE
            End If
            If InStr(sDummy, "(file missing)") > 0 Then
                sDummy = Left$(sDummy, InStr(sDummy, " (file missing)") - 1)
            End If
            RegSetDwordVal lHive, "Software\Microsoft\Internet Explorer\Styles", "Use My Stylesheet", 1
            RegSetStringVal lHive, "Software\Microsoft\Internet Explorer\Styles", "User Stylesheet", sDummy
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
        
        Case "O20" 'appinit_dlls
            'O20 - AppInit_DLLs: file.dll
            'O20 - Winlogon Notify: blaat - c:\file.dll
            If InStr(sItem, "AppInit_DLLs") > 0 Then
                sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
                sDummy = Replace$(sDummy, "|", vbNullChar)
                RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", sDummy
            Else
                'backup has extra reg data
                sData = vbNullString
                ff = FreeFile()
                Open sPath & "backups\" & sFile For Input As #ff
                    Line Input #ff, sLine
                    Line Input #ff, sLine
                    Do
                        Line Input #ff, sLine
                        sData = sData & sLine & vbCrLf
                    Loop Until EOF(ff)
                Close #ff
                If Left$(sData, 2) = "ÿþ" Then sData = Mid$(sData, 3)
                ff = FreeFile()
                Open sPath & "backups\" & sFile & ".reg" For Output As #ff
                    Print #ff, sData
                Close #ff
                Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
                DoEvents
                If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".reg"))
                    
            End If
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O21" 'ssodl
            'O21 - SSODL: webcheck - {000....000} - c:\file.dll
            'todo:
            'get, print and merge .reg data for clsid regkey
            'reconstruct reg value at SSODL regkey
            
            sName = Mid$(sItem, InStr(sItem, ": ") + 2)
            sCLSID = Mid$(sName, InStr(sName, " - ") + 3)
            sName = Left$(sName, InStr(sName, " - ") - 1)
            sCLSID = Left$(sCLSID, InStr(sCLSID, " - ") - 1)
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad", sName, sCLSID
            
            'backup has extra info with reg data
            sData = vbNullString
            ff = FreeFile()
            Open sPath & "backups\" & sFile For Input As #ff
                Line Input #ff, sLine
                Line Input #ff, sLine
                Do
                    Line Input #ff, sLine
                    sData = sData & sLine & vbCrLf
                Loop Until EOF(ff)
            Close #ff
            
            'regedit in 2000/XP tends to prepend ÿþ
            'to .reg files - they won't merge then
            If Left$(sData, 2) = "ÿþ" Then sData = Mid$(sData, 3)
            
            ff = FreeFile()
            Open sPath & "backups\" & sFile & ".reg" For Output As #ff
                Print #ff, sData
            Close #ff
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".reg"))
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O22" 'ScheduledTask
            'O22 - ScheduledTask: blah - {000...000} - file.dll
            'todo:
            'restore sts regval
            'restore clsid regkey
            
            sName = Mid$(sItem, InStr(sItem, ": ") + 2)
            sCLSID = Mid$(sName, InStr(sName, " - ") + 3)
            sName = Left$(sName, InStr(sName, " - ") - 1)
            sCLSID = Left$(sCLSID, InStr(sCLSID, " - ") - 1)
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler", sCLSID, sName
            
            'backup has extra info with reg data
            sData = vbNullString
            ff = FreeFile()
            Open sPath & "backups\" & sFile For Input As #ff
                Line Input #ff, sLine
                Line Input #ff, sLine
                Do
                    Line Input #ff, sLine
                    sData = sData & sLine & vbCrLf
                Loop Until EOF(ff)
            Close #ff
            
            'regedit in 2000/XP tends to prepend ÿþ
            'to .reg files - they won't merge then
            If Left$(sData, 2) = "ÿþ" Then sData = Mid$(sData, 3)
            ff = FreeFile()
            Open sPath & "backups\" & sFile & ".reg" For Output As #ff
                Print #ff, sData
            Close #ff
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".reg"))
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
        
        Case "O23"
            'O23 - Service: bla bla - blacorp - bla.exe
            'todo:
            'enable & start service
            Dim sServices$(), sDisplayName$
            sDisplayName = Mid$(sItem, InStr(sItem, ": ") + 2)
            sDisplayName = Left$(sDisplayName, InStr(sDisplayName, " - ") - 1)
            sServices = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
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
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            
        Case "O24"
            'O24 - Desktop Component N: blah - c:\windows\index.html
            'todo:
            'restore reg key
            Dim sSource$
            'copy file back to Source and SubscribedURL
            sSource = Mid$(sItem, InStr(sItem, ":"))
            sSource = Mid$(sSource, InStr(sSource, " - ") + 3)
            If sSource <> "(no file)" Then
                'only one is backed up if they are the same
                If FileExists(sPath & "backups\" & sFile & "-source.html") Then
                    FileCopyW sPath & "backups\" & sFile & "-source.html", sSource
                End If
                If FileExists(sPath & "backups\" & sFile & "-suburl.html") Then
                    FileCopyW sPath & "backups\" & sFile & "-suburl.html", sSource
                End If
            End If
            
            'backup has extra info with reg data
            sData = vbNullString
            ff = FreeFile()
            Open sPath & "backups\" & sFile For Input As #ff
                Line Input #ff, sLine
                Line Input #ff, sLine
                Do
                    Line Input #ff, sLine
                    sData = sData & sLine & vbCrLf
                Loop Until EOF(ff)
            Close #ff
            
            'regedit in 2000/XP tends to prepend ÿþ
            'to .reg files - they won't merge then
            If Left$(sData, 2) = "ÿþ" Then sData = Mid$(sData, 3)
            ff = FreeFile()
            Open sPath & "backups\" & sFile & ".reg" For Output As #ff
                Print #ff, sData
            Close #ff
            Shell sWinDir & "\regedit.exe /s """ & sPath & "backups\" & sFile & ".reg""", vbHide
            DoEvents
            If FileExists(sPath & "backups\" & sFile & ".reg") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & ".reg"))
            If FileExists(sPath & "backups\" & sFile) Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile))
            If FileExists(sPath & "backups\" & sFile & "-source.html") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & "-source.html"))
            If FileExists(sPath & "backups\" & sFile & "-suburl.html") Then DeleteFileWEx (StrPtr(sPath & "backups\" & sFile & "-suburl.html"))
        Case Else
            'Restore for this item is not implemented:
            MsgBoxW Translate(89) & vbCrLf & vbCrLf & sItem
        
    End Select
    
    AppendErrorLogCustom "RestoreBackup - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modBackup_RestoreBackup", "sItem=", sItem
    Close #ff
    If inIDE Then Stop: Resume Next
End Sub

Public Sub ListBackups()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ListBackups - Begin"
    
    Dim sPath$, sFile$, vDummy As Variant, ff%
    Dim sBackup$, sDate$, stime$, aBackup() As String, Cnt&, i&
    
    ReDim aBackup(100)
    
    sPath = AppPath() & IIf(Right$(AppPath(), 1) = "\", vbNullString, "\")
    sFile = DirW$(sPath & "backups\" & "backup*", vbFile)
    If Len(sFile) = 0 Then Exit Sub
    frmMain.lstBackups.Clear
    
    Do
        vDummy = Split(sFile, "-")
        If InStr(sFile, ".") = 0 And UBound(vDummy) = 3 Then
            'backup-20021024-181841-901
            '0      1        2      3
            'duh    date     time   random
            ff = FreeFile()
            Open sPath & "backups\" & sFile For Input As #ff
                Line Input #ff, sBackup
            Close #ff
            
            sDate = Right$(vDummy(1), 2) & "-" & Mid$(vDummy(1), 5, 2) & "-" & Mid$(vDummy(1), 1, 4)
            stime = Left$(vDummy(2), 2) & ":" & Mid$(vDummy(2), 3, 2) & ":" & Right$(vDummy(2), 2)
            
            sBackup = Format(sDate, "Short Date") & ", " & _
                      Format(stime, "Long Time") & ": " & _
                      sBackup
                      
            Cnt = Cnt + 1
            If UBound(aBackup) < Cnt Then ReDim Preserve aBackup(UBound(aBackup) + 100)
            aBackup(Cnt) = sBackup
        End If
        sFile = DirW$()
    Loop Until sFile = vbNullString
    
    If Cnt <> 0 Then
        For i = Cnt To 1 Step -1
            frmMain.lstBackups.AddItem aBackup(i)
        Next
    End If
    
    AppendErrorLogCustom "ListBackups - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modBackup_ListBackups"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub DeleteBackup(sBackup$)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "DeleteBackup - Begin", sBackup
    
    Dim sFile$, sDate$, stime$
    
    If sBackup = vbNullString Then
        '// TODO
        'DeleteFileWEx StrPtr(BuildPath(AppPath(), "backups\backup-*.*"))
        Exit Sub
    End If
    
    sDate = Left$(sBackup, InStr(sBackup, ", ") - 1)
    stime = Mid$(sBackup, InStr(sBackup, ", ") + 2)
    stime = Left$(stime, InStr(stime, ": ") - 1)
    
    If Not bIsUSADateFormat Then
        sDate = Format(sDate, "yyyymmdd")
    Else
        'use stupid workaround for USA date format
        sDate = Format(sDate, "yyyyddmm")
    End If
    stime = Format(stime, "HhNnSs")
    
    sFile = "backup-" & sDate & "-" & stime & "*.*"
    
    DeleteFileWEx StrPtr(BuildPath(AppPath(), "backups\" & sFile))
    
    AppendErrorLogCustom "DeleteBackup - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modBackup_DeleteBackup", "sBackup=", sBackup
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetCLSIDOfMSIEExtension(ByVal sName$, bButtonOrMenu As Boolean)
    Dim hKey&, i&, sCLSID$
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetCLSIDOfMSIEExtension - Begin"
    
    sName = Left$(sName, InStr(sName, " (HK") - 1)
    
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr("Software\Microsoft\Internet Explorer\Extensions"), 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sCLSID = String$(255, 0)
        If RegEnumKeyExW(hKey, i, StrPtr(sCLSID), 255, 0, 0, ByVal 0, ByVal 0) <> 0 Then
            RegCloseKey hKey
            GetCLSIDOfMSIEExtension = vbNullString
            Exit Function
        End If
        Do
            sCLSID = Left$(sCLSID, InStr(sCLSID, vbNullChar) - 1)
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
            
            sCLSID = String$(255, 0)
            i = i + 1
        Loop Until RegEnumKeyExW(hKey, i, StrPtr(sCLSID), 255, 0, 0, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr("Software\Microsoft\Internet Explorer\Extensions"), 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sCLSID = String$(255, 0)
        If RegEnumKeyExW(hKey, i, StrPtr(sCLSID), 255, 0, 0, ByVal 0, ByVal 0) <> 0 Then
            RegCloseKey hKey
            GetCLSIDOfMSIEExtension = vbNullString
            Exit Function
        End If
        Do
            sCLSID = Left$(sCLSID, InStr(sCLSID, vbNullChar) - 1)
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
            
            sCLSID = String$(255, 0)
            i = i + 1
        Loop Until RegEnumKeyExW(hKey, i, StrPtr(sCLSID), 255, 0, 0, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    AppendErrorLogCustom "GetCLSIDOfMSIEExtension - End"
    Exit Function
ErrorHandler:
    RegCloseKey hKey
    ErrorMsg Err, "modBackup_GetCLSIDOfMSIEExtension", "sName=", sName, "bButtonOrMenu=", CStr(bButtonOrMenu)
    If inIDE Then Stop: Resume Next
End Function

Public Function HasBOM_UTF16(sText As String) As Boolean
    If Left$(sText, 2) = Chr$(&HFF) & Chr$(&HFE) Or Left$(sText, 2) = ChrW$(&HFF) & ChrW$(&HFE) Then HasBOM_UTF16 = True
End Function
