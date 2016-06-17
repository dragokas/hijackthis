Attribute VB_Name = "modLSP"
Option Explicit

Private sKeyNameSpace$
Private sKeyProtocol$

Private Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValueW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegDeleteKeyW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
Private Declare Function RegCreateKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long


Public Sub GetLSPCatalogNames()
    sKeyNameSpace = "System\CurrentControlSet\Services\WinSock2\Parameters"
    sKeyProtocol = "System\CurrentControlSet\Services\WinSock2\Parameters"
    
    sKeyNameSpace = sKeyNameSpace & "\" & RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Current_NameSpace_Catalog")
    sKeyProtocol = sKeyProtocol & "\" & RegGetString(HKEY_LOCAL_MACHINE, sKeyProtocol, "Current_Protocol_Catalog")
End Sub

Public Sub CheckLSP()
    On Error GoTo ErrorHandler:
    
    Dim lNumNameSpace&, lNumProtocol&, i&, J& ', sSafeFiles$
    Dim sFile$, uData() As Byte, hKey&, sHit$, sDummy$
    
    lNumNameSpace = RegGetDword(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries")
    lNumProtocol = RegGetDword(HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries")
        
    'check for gaps in LSP chain
    For i = 1 To lNumNameSpace
        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String$(12 - Len(CStr(i)), "0") & CStr(i)) Then
            'all fine & peachy
        Else
            'broken LSP detected!
            sHit = "O10 - Broken Internet access because of LSP chain gap (#" & CStr(i) & " in chain of " & CStr(lNumNameSpace) & " missing)"
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
            Exit Sub
        End If
    Next i
    For i = 1 To lNumProtocol
        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String$(12 - Len(CStr(i)), "0") & CStr(i)) Then
            'all fine & dandy
        Else
            'shit, not again!
            sHit = "O10 - Broken Internet access because of LSP chain gap (#" & CStr(i) & " in chain of " & CStr(lNumProtocol) & " missing)"
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
            Exit Sub
        End If
    Next i
    
    'check all LSP providers are present
    For i = 1 To lNumNameSpace
        sFile = RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String$(12 - Len(CStr(i)), "0") & CStr(i), "LibraryPath")
        sFile = LCase$(Replace$(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare))
        sFile = LCase$(Replace$(sFile, "%windir%", sWinDir, , , vbTextCompare))
        If sFile <> vbNullString Then
            If FileExists(sFile) Or _
               FileExists(sWinDir & "\" & sFile) Or _
               FileExists(sWinSysDir & "\" & sFile) Then
                'file ok
                If InStr(sFile, "webhdll.dll") > 0 Then
                    sHit = "O10 - Hijacked Internet access by WebHancer"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                ElseIf InStr(sFile, "newdot") > 0 Then
                    sHit = "O10 - Hijacked Internet access by New.Net"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                ElseIf InStr(sFile, "cnmib.dll") > 0 Then
                    sHit = "O10 - Hijacked Internet access by CommonName"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                Else
                    sDummy = Mid$(sFile, InStrRev(sFile, "\") + 1)
                    If InStr(1, sSafeLSPFiles, sDummy, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                        sHit = "O10 - Unknown file in Winsock LSP: " & sFile
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                    End If
                End If
            Else
                'damn, file is gone
                If InStr(1, sSafeLSPFiles, sFile, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                    sHit = "O10 - Broken Internet access because of LSP provider '" & sFile & "' missing"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                End If
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To lNumProtocol
        sFile = RegGetFileFromBinary(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String$(12 - Len(CStr(i)), "0") & CStr(i), "PackedCatalogItem")
        
        If sFile <> vbNullString Then
            sFile = EnvironW(sFile)
            If FileExists(sFile) Or _
               FileExists(sWinDir & "\" & sFile) Or _
               FileExists(sWinSysDir & "\" & sFile) Then
                'file ok
                If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Then
                    sHit = "O10 - Hijacked Internet access by WebHancer"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                ElseIf InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
                    sHit = "O10 - Hijacked Internet access by New.Net"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                ElseIf InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
                    sHit = "O10 - Hijacked Internet access by CommonName"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                Else
                    sDummy = LCase$(Mid$(sFile, InStrRev(sFile, "\") + 1))
                    If InStr(1, sSafeLSPFiles, sDummy, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                        sHit = "O10 - Unknown file in Winsock LSP: " & sFile
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                    End If
                End If
            Else
                'damn - crossed again!
                If InStr(1, sSafeLSPFiles, sFile, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                    sHit = "O10 - Broken Internet access because of LSP provider '" & sFile & "' missing"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O10", sHit
                End If
                Exit Sub
            End If
        End If
    Next i
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modLSP_CheckLSP"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixLSP()
    On Error GoTo ErrorHandler:

    Dim lNumNameSpace&, lNumProtocol&
    Dim i&, J&, sFile$, hKey&, uData() As Byte
    
    If Not bSeenLSPWarning Then
        'MsgBoxW "HiJackThis cannot repair O10 Winsock LSP entries. " & vbCrLf & _
        '       "You should use WinsockReset for that, which is available " & _
        '       "from https://www.foolishit.com/vb6-projects/winsockreset/" & vbCrLf & vbCrLf & _
        '       "Would you like to visit that site?"
        
        If vbYes = MsgBoxW(Translate(580), vbExclamation Or vbYesNo) Then
            ShellExecute 0&, StrPtr("open"), StrPtr("https://www.foolishit.com/vb6-projects/winsockreset/"), 0&, 0&, 1
        End If
        bSeenLSPWarning = True
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg err, "FixLSP"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixLSP_Old()
    On Error GoTo ErrorHandler:
    
    Dim lNumNameSpace&, lNumProtocol&
    Dim i&, J&, sFile$, hKey&, uData() As Byte
    
    If Not bSeenLSPWarning Then
        MsgBox "HijackThis cannot repair O10 Winsock LSP entries. " & vbCrLf & _
               "You should use LSPFix for that, which is available " & _
               "from http://www.cexx.org/lspfix.htm." & vbCrLf & vbCrLf & _
               "If the O10 item belongs to WebHancer, New.Net or CommonName, " & _
               "Spybot S&D can remove it automatically. Spybot S&D " & _
               "is available from http://www.spybot.info/.", vbCritical
               
        bSeenLSPWarning = True
    End If
    Exit Sub
    
    lNumNameSpace = RegGetDword(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries")
    lNumProtocol = RegGetDword(HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries")
    
    'check for missing files, delete keys with those
    For i = 1 To lNumNameSpace
        sFile = RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), "LibraryPath")
        sFile = LCase(Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare))
        sFile = LCase(Replace(sFile, "%windir%", sWinDir, , , vbTextCompare))
        If sFile <> vbNullString And Dir$(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            'file ok
            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Or _
               InStr(1, sFile, "newdot", vbTextCompare) > 0 Or _
               InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
                'it's New.Net/WebHancer/CN! Kill it!
                DeleteFileWEx StrPtr(sFile)  ' error 53 = file not found
                If FileExists(sFile) Then
                    If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Then
                        MsgBox "The WebHancer Agent is currently active and can't be deleted. Use Ad-Aware from www.lavasoft.nu to remove it safely.", vbExclamation
                    ElseIf InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
                        SHRestartSystemMB frmMain.hwnd, "The NewDotNet DLL is currently active. You will need to reboot and rescan, then remove New.Net and fix the WinSock stack." & vbCrLf & vbCrLf, 0
                        Exit Sub
                    ElseIf InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
                        SHRestartSystemMB frmMain.hwnd, "The CommonName DLL is currently active. You will need to reboot and rescan, then remove CommonName and fix the WinSock stack." & vbCrLf & vbCrLf, 0
                    End If
                End If
                
                RegDelKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
                lNumNameSpace = lNumNameSpace - 1
                
                'delete New.Net startup Reg entry
                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "New.Net Startup"
                'delete WebHancer startup Reg entry
                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "webHancer Agent"
                'delete CommonName startup Reg entry
                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Zenet"
            End If
        Else
            If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
                lNumNameSpace = lNumNameSpace - 1
            End If
            RegDelKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
        End If
    Next i
    
    For i = 1 To lNumProtocol
        sFile = RegGetFileFromBinary(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), "PackedCatalogItem")
        sFile = LCase(Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare))
        sFile = LCase(Replace(sFile, "%windir%", sWinDir, , , vbTextCompare))
        If sFile <> vbNullString And Dir$(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            'file ok
            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Or _
               InStr(1, sFile, "newdotnet", vbTextCompare) > 0 Or _
               InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
                'it's New.Net/WebHancer/CN! Kill it!
                DeleteFileWEx StrPtr(sFile)  ' error 53 = file not found
                If FileExists(sFile) Then
                    If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Then
                        MsgBox "The WebHancer Agent is currently active and can't be deleted. Use Ad-Aware from www.lavasoft.nu to remove it safely.", vbExclamation
                    ElseIf InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
                        SHRestartSystemMB frmMain.hwnd, "The NewDotNet DLL is currently active. You will need to reboot and rescan, then remove New.Net and fix the WinSock stack." & vbCrLf & vbCrLf, 0
                        Exit Sub
                    ElseIf InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
                        SHRestartSystemMB frmMain.hwnd, "The CommonName DLL is currently active. You will need to reboot and rescan, then remove CommonName and fix the WinSock stack." & vbCrLf & vbCrLf, 0
                    End If
                End If
                
                RegDelKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
                lNumNameSpace = lNumNameSpace - 1
                
                'delete New.Net startup Reg entry
                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "New.Net Startup"
                
                'delete WebHancer startup Reg entry
                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "webHancer Agent"
                
                'delete CommonName startup Reg entry
                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", "Zenet"
            End If
        Else
            If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
                lNumProtocol = lNumProtocol - 1
            End If
            RegDelKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
        End If
    Next i
    
    'check LSP chain, fix gaps where found
    i = 1 'current LSP #
    J = 1 'correct LSP #
    Do
        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
            If i > J Then
                RegRenameKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(J)), "0") & CStr(J)
            End If
            J = J + 1
        Else
            'nothing, j stays the same
        End If
        i = i + 1
        'check to prevent infinite loop when
        'lNumNameSpace is wrong
        If i = 100 Then
            lNumNameSpace = J - 1
            Exit Do
        End If
    Loop Until J = lNumNameSpace + 1
    
    i = 1
    J = 1
    Do
        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
            If i > J Then
                RegRenameKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(J)), "0") & CStr(J)
            End If
            J = J + 1
        Else
            'nothing, j stays the same
        End If
        i = i + 1
        If i = 100 Then
            lNumProtocol = J - 1
            Exit Do
        End If
    Loop Until J = lNumProtocol + 1
    
    RegSetDwordVal HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries", lNumNameSpace
    RegSetDwordVal HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries", lNumProtocol
    
    bRebootNeeded = True
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modLSP_FixLSP"
    RegCloseKey hKey
End Sub

Private Sub RegRenameKey(lHive&, sKeyOldName$, sKeyNewName$)
    On Error GoTo ErrorHandler:
    
    Dim hKey&, hKey2&, i&, J&, sName$, lType&, lDataLen&
    Dim sData$, lData&, uData() As Byte
    Dim lEnumBufSize As Long
    
    lEnumBufSize = 32767&
    
    If RegOpenKeyExW(lHive, StrPtr(sKeyOldName), 0, KEY_QUERY_VALUE Or KEY_WRITE, hKey) <> 0 Then Exit Sub
    If RegOpenKeyExW(lHive, StrPtr(sKeyNewName), 0, KEY_QUERY_VALUE, hKey2) = 0 Then
        RegCloseKey hKey2
        RegDeleteKeyW lHive, StrPtr(sKeyNewName)
    End If
    If RegCreateKeyExW(lHive, StrPtr(sKeyNewName), 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_WRITE, ByVal 0, hKey2, ByVal 0) <> 0 Then Exit Sub
    
    'assume key has no subkeys (which it does not have
    'where we use it for)
    
    i = 0
    sName = String$(lEnumBufSize, 0)
    lDataLen = lEnumBufSize
    sData = String$(lDataLen, 0)
    lType = 0
    If RegEnumValueW(hKey, i, StrPtr(sName), Len(sName), 0, lType, StrPtr(sData), lDataLen) <> 0 Then
        'no values to transfer
        RegCloseKey hKey
        RegCloseKey hKey2
        RegDelKey lHive, sKeyOldName
        Exit Sub
    End If
    
    Do
        sName = Left$(sName, InStr(sName, vbNullChar) - 1)
        Select Case lType
            Case REG_SZ
                'reconstruct string
                'sData = ""
                'For j = 0 To lDataLen - 1
                '    If uData(j) = 0 Then Exit For
                '    sData = sData & Chr(uData(j))
                'Next j
                'sData = StrConv(uData, vbUnicode)
                sData = TrimNull(sData)
                If Len(sData) = 0 Then
                    RegSetValueExW hKey2, StrPtr(sName), 0, REG_SZ, ByVal 0&, 0&
                Else
                    RegSetValueExW hKey2, StrPtr(sName), 0, REG_SZ, ByVal StrPtr(sData), Len(sData) * 2 + 2
                End If
                'RegDeleteValue hKey, sName
            Case REG_DWORD
                'reconstruct dword
'                lData = 0
'                lData = CLng(Val("&H" & _
'                                 String$(2 - Len(Hex(uData(3))), "0") & Hex(uData(3)) & _
'                                 String$(2 - Len(Hex(uData(2))), "0") & Hex(uData(2)) & _
'                                 String$(2 - Len(Hex(uData(1))), "0") & Hex(uData(1)) & _
'                                 String$(2 - Len(Hex(uData(0))), "0") & Hex(uData(0))))
                lData = AscW(Mid$(sData, 1, 1)) + AscW(Mid$(sData, 2, 1)) * 256&
                RegSetValueExW hKey2, StrPtr(sName), 0, REG_DWORD, lData, 4&
                'RegDeleteValue hKey, sName
            Case REG_BINARY
                'at ease, soldier
                ReDim Preserve uData(lDataLen)
                If Len(sData) = 0 Then
                    RegSetValueExW hKey2, StrPtr(sName), 0, REG_BINARY, ByVal 0&, 0&
                Else
                    RegSetValueExW hKey2, StrPtr(sName), 0, REG_BINARY, ByVal StrPtr(sData), Len(sData) * 2
                End If
                'RegDeleteValue hKey, sName
            Case Else
                'wtf?
        End Select
        
        i = i + 1
        sName = String$(lEnumBufSize, 0)
        lDataLen = lEnumBufSize
        sData = String$(lDataLen, 0)
        lType = 0
    Loop Until RegEnumValueW(hKey, i, StrPtr(sName), Len(sName), 0, lType, StrPtr(sData), lDataLen) <> 0
    RegCloseKey hKey
    RegCloseKey hKey2
    RegDeleteKeyW lHive, StrPtr(sKeyOldName)
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modLSP_RegRenameKey", "sKeyOldName=", sKeyOldName, "sKeyNewName=", sKeyNewName
    RegCloseKey hKey
    RegCloseKey hKey2
    If inIDE Then Stop: Resume Next
End Sub

