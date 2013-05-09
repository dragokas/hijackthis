Attribute VB_Name = "modLSP"
Option Explicit
Private sKeyNameSpace$
Private sKeyProtocol$

Public Sub GetLSPCatalogNames()
    sKeyNameSpace = "System\CurrentControlSet\Services\WinSock2\Parameters"
    sKeyProtocol = "System\CurrentControlSet\Services\WinSock2\Parameters"
    
    sKeyNameSpace = sKeyNameSpace & "\" & RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Current_NameSpace_Catalog")
    sKeyProtocol = sKeyProtocol & "\" & RegGetString(HKEY_LOCAL_MACHINE, sKeyProtocol, "Current_Protocol_Catalog")
End Sub

Public Sub CheckLSP()
    Dim lNumNameSpace&, lNumProtocol&, i&, j& ', sSafeFiles$
    Dim sFile$, uData() As Byte, hKey&, sHit$, sDummy$
    On Error GoTo Error:
    lNumNameSpace = RegGetDword(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries")
    lNumProtocol = RegGetDword(HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries")
        
    'check for gaps in LSP chain
    For i = 1 To lNumNameSpace
        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
            'all fine & peachy
        Else
            'broken LSP detected!
            frmMain.lstResults.AddItem "O10 - Broken Internet access because of LSP chain gap (#" & CStr(i) & " in chain of " & CStr(lNumNameSpace) & " missing)"
            Exit Sub
        End If
    Next i
    For i = 1 To lNumProtocol
        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
            'all fine & dandy
        Else
            'shit, not again!
            frmMain.lstResults.AddItem "O10 - Broken Internet access because of LSP chain gap (#" & CStr(i) & " in chain of " & CStr(lNumProtocol) & " missing)"
            Exit Sub
        End If
    Next i
    
    'check all LSP providers are present
    For i = 1 To lNumNameSpace
        sFile = RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), "LibraryPath")
        sFile = LCase(Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare))
        sFile = LCase(Replace(sFile, "%windir%", sWinDir, , , vbTextCompare))
        If sFile <> vbNullString Then
            If FileExists(sFile) Or _
               FileExists(sWinDir & "\" & sFile) Or _
               FileExists(sWinSysDir & "\" & sFile) Then
                'file ok
                If InStr(sFile, "webhdll.dll") > 0 Then
                    sHit = "O10 - Hijacked Internet access by WebHancer"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                ElseIf InStr(sFile, "newdot") > 0 Then
                    sHit = "O10 - Hijacked Internet access by New.Net"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                ElseIf InStr(sFile, "cnmib.dll") > 0 Then
                    sHit = "O10 - Hijacked Internet access by CommonName"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                Else
                    sDummy = Mid(sFile, InStrRev(sFile, "\") + 1)
                    If InStr(1, sSafeLSPFiles, sDummy, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                        sHit = "O10 - Unknown file in Winsock LSP: " & sFile
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                End If
            Else
                'damn, file is gone
                If InStr(1, sSafeLSPFiles, sFile, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                    frmMain.lstResults.AddItem "O10 - Broken Internet access because of LSP provider '" & sFile & "' missing"
                End If
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To lNumProtocol
        sFile = RegGetFileFromBinary(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), "PackedCatalogItem")
        sFile = LCase(Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare))
        sFile = LCase(Replace(sFile, "%windir%", sWinDir, , , vbTextCompare))
        If sFile <> vbNullString Then
            If FileExists(sFile) Or _
               FileExists(sWinDir & "\" & sFile) Or _
               FileExists(sWinSysDir & "\" & sFile) Then
                'file ok
                If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Then
                    sHit = "O10 - Hijacked Internet access by WebHancer"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                ElseIf InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
                    sHit = "O10 - Hijacked Internet access by New.Net"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                ElseIf InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
                    sHit = "O10 - Hijacked Internet access by CommonName"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                Else
                    sDummy = LCase(Mid(sFile, InStrRev(sFile, "\") + 1))
                    If InStr(1, sSafeLSPFiles, sDummy, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                        sHit = "O10 - Unknown file in Winsock LSP: " & sFile
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                End If
            Else
                'damn - crossed again!
                If InStr(1, sSafeLSPFiles, sFile, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                    frmMain.lstResults.AddItem "O10 - Broken Internet access because of LSP provider '" & sFile & "' missing"
                End If
                Exit Sub
            End If
        End If
    Next i
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modLSP_CheckLSP", Err.Number, Err.Description
End Sub

Public Sub FixLSP()
    Dim lNumNameSpace&, lNumProtocol&
    Dim i&, j&, sFile$, hKey&, uData() As Byte
    On Error GoTo Error:
    
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
        If sFile <> vbNullString And Dir(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            'file ok
            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Or _
               InStr(1, sFile, "newdot", vbTextCompare) > 0 Or _
               InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
                'it's New.Net/WebHancer/CN! Kill it!
                DeleteFile sFile  ' error 53 = file not found
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
        If sFile <> vbNullString And Dir(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            'file ok
            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Or _
               InStr(1, sFile, "newdotnet", vbTextCompare) > 0 Or _
               InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
                'it's New.Net/WebHancer/CN! Kill it!
                DeleteFile sFile  ' error 53 = file not found
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
    j = 1 'correct LSP #
    Do
        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
            If i > j Then
                RegRenameKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(j)), "0") & CStr(j)
            End If
            j = j + 1
        Else
            'nothing, j stays the same
        End If
        i = i + 1
        'check to prevent infinite loop when
        'lNumNameSpace is wrong
        If i = 100 Then
            lNumNameSpace = j - 1
            Exit Do
        End If
    Loop Until j = lNumNameSpace + 1
    
    i = 1
    j = 1
    Do
        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
            If i > j Then
                RegRenameKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(j)), "0") & CStr(j)
            End If
            j = j + 1
        Else
            'nothing, j stays the same
        End If
        i = i + 1
        If i = 100 Then
            lNumProtocol = j - 1
            Exit Do
        End If
    Loop Until j = lNumProtocol + 1
    
    RegSetDwordVal HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries", lNumNameSpace
    RegSetDwordVal HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries", lNumProtocol
    
    bRebootNeeded = True
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modLSP_FixLSP", Err.Number, Err.Description
End Sub

Private Sub RegRenameKey(lHive&, sKeyOldName$, sKeyNewName$)
    Dim hKey&, hKey2&, i&, j&, sName$, lType&, lDataLen&
    Dim sData$, lData&, uData() As Byte
    On Error GoTo Error:
    If RegOpenKeyEx(lHive, sKeyOldName, 0, KEY_QUERY_VALUE Or KEY_WRITE, hKey) <> 0 Then Exit Sub
    If RegOpenKeyEx(lHive, sKeyNewName, 0, KEY_QUERY_VALUE, hKey2) = 0 Then
        RegCloseKey hKey2
        RegDeleteKey lHive, sKeyNewName
    End If
    If RegCreateKeyEx(lHive, sKeyNewName, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_WRITE, ByVal 0, hKey2, ByVal 0) <> 0 Then Exit Sub
    
    'assume key has no subkeys (which it does not have
    'where we use it for)
    
    i = 0
    sName = String(lEnumBufSize, 0)
    lDataLen = lEnumBufSize
    ReDim uData(lDataLen)
    lType = 0
    If RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0 Then
        'no values to transfer
        RegCloseKey hKey
        RegCloseKey hKey2
        RegDelKey lHive, sKeyOldName
        Exit Sub
    End If
    
    Do
        sName = Left(sName, InStr(sName, Chr(0)) - 1)
        Select Case lType
            Case REG_SZ
                'reconstruct string
                'sData = ""
                'For j = 0 To lDataLen - 1
                '    If uData(j) = 0 Then Exit For
                '    sData = sData & Chr(uData(j))
                'Next j
                sData = StrConv(uData, vbUnicode)
                sData = TrimNull(sData)
                RegSetValueEx hKey2, sName, 0, REG_SZ, ByVal sData, Len(sData)
                'RegDeleteValue hKey, sName
            Case REG_DWORD
                'reconstruct dword
                lData = 0
                lData = CLng(Val("&H" & _
                                 String(2 - Len(Hex(uData(3))), "0") & Hex(uData(3)) & _
                                 String(2 - Len(Hex(uData(2))), "0") & Hex(uData(2)) & _
                                 String(2 - Len(Hex(uData(1))), "0") & Hex(uData(1)) & _
                                 String(2 - Len(Hex(uData(0))), "0") & Hex(uData(0))))
                RegSetValueEx hKey2, sName, 0, REG_DWORD, lData, 4
                'RegDeleteValue hKey, sName
            Case REG_BINARY
                'at ease, soldier
                ReDim Preserve uData(lDataLen)
                RegSetValueEx hKey2, sName, 0, REG_BINARY, uData(0), UBound(uData)
                'RegDeleteValue hKey, sName
            Case Else
                'wtf?
        End Select
        
        i = i + 1
        sName = String(lEnumBufSize, 0)
        lDataLen = lEnumBufSize
        ReDim uData(lDataLen)
        lType = 0
    Loop Until RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0
    RegCloseKey hKey
    RegCloseKey hKey2
    RegDeleteKey lHive, sKeyOldName
    Exit Sub
    
Error:
    RegCloseKey hKey
    RegCloseKey hKey2
    ErrorMsg "modLSP_RegRenameKey", Err.Number, Err.Description, "sKeyOldName=" & sKeyOldName & ",sKeyNewName=" & sKeyNewName
End Sub

Private Function RegGetFileFromBinary$(lHive&, sKey$, sValue$)
    Dim hKey&, uData() As Byte, sFile$
    Dim i&
    On Error GoTo Error:
    
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        ReDim uData(1024)
        If RegQueryValueEx(hKey, sValue, 0, 0, uData(0), 1024) = 0 Then
            sFile = ""
            For i = 0 To 1024
                If uData(i) = 0 Then Exit For
                sFile = sFile & Chr(uData(i))
            Next i
        End If
        RegCloseKey hKey
    End If
    RegGetFileFromBinary = sFile
    Exit Function
    
Error:
    RegCloseKey hKey
    ErrorMsg "modLSP_RegGetFileFromBinary", Err.Number, Err.Description, "sKey=" & sKey & ",sValue=" & sValue
End Function

