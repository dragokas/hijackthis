Attribute VB_Name = "modLSP"
Option Explicit

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(257) As Byte
    szSystemStatus(129) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type UUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type WSANAMESPACE_INFO
    NSProviderId   As UUID
    dwNameSpace    As Long
    fActive        As Long
    dwVersion      As Long
    lpszIdentifier As Long
End Type

Private Type WSAPROTOCOLCHAIN
    ChainLen As Long
    ChainEntries(6) As Long
End Type

Private Type WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As UUID
    dwCatalogEntryId As Long
    ProtocolChain As WSAPROTOCOLCHAIN
    iVersion As Long
    iAddressFamily As Long
    iMaxSockAddr As Long
    iMinSockAddr As Long
    iSocketType As Long
    iProtocol As Long
    iProtocolMaxOffset As Long
    iNetworkByteOrder As Long
    iSecurityScheme As Long
    dwMessageSize As Long
    dwProviderReserved As Long
    szProtocol As String * 256
End Type

Private Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValueW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegDeleteKeyW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
Private Declare Function RegCreateKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SHRestartSystemMB Lib "shell32.dll" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Integer, ByVal lpWSAD As Long) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function WSAEnumProtocols Lib "ws2_32.dll" Alias "WSAEnumProtocolsW" (ByVal lpiProtocols As Long, ByVal lpProtocolBuffer As Long, lpdwBufferLength As Long) As Long
Private Declare Function WSAEnumNameSpaceProviders Lib "ws2_32.dll" Alias "WSAEnumNameSpaceProvidersW" (lpdwBufferLength As Long, ByVal lpnspBuffer As Long) As Long
Private Declare Function WSCGetProviderPath Lib "ws2_32.dll" (ByVal lpProviderId As Long, ByVal lpszProviderDllPath As Long, ByVal lpProviderDllPathLen As Long, ByVal lpErrno As Long) As Long

Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As UUID, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal iMaxLength As Long) As Long
Private Declare Sub memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Const SOCKET_ERROR As Long = -1

Private sKeyNameSpace As String
Private sKeyProtocol As String


' ---------------------------------------------------------------------------------------------------
' StartupList2 routine
' ---------------------------------------------------------------------------------------------------

Public Function EnumWinsockProtocol$()
    On Error GoTo ErrorHandler:
    
    Dim i&, sEnumProt$
    Dim uWSAData As WSAData, sGUID$, sFile$
    Dim uWSAProtInfo As WSAPROTOCOL_INFO
    Dim uBuffer() As Byte, lBufferSize&
    Dim lNumProtocols&, sLSPName$, lDummy&
    
    If WSAStartup(&H202, VarPtr(uWSAData)) <> 0 Then Exit Function
    
    WSAEnumProtocols 0&, 0&, lBufferSize
    ReDim uBuffer(lBufferSize - 1)
    
    lNumProtocols = WSAEnumProtocols(0&, VarPtr(uBuffer(0)), lBufferSize)
    If lNumProtocols <> SOCKET_ERROR Then
        For i = 0 To lNumProtocols - 1
            memcpy ByVal VarPtr(uWSAProtInfo), ByVal VarPtr(uBuffer(i * LenB(uWSAProtInfo))), LenB(uWSAProtInfo)
            sGUID = GuidToString(uWSAProtInfo.ProviderId)
            sFile = GetProviderFile(uWSAProtInfo.ProviderId)
            sLSPName = TrimNull(uWSAProtInfo.szProtocol)
            If bShowCLSIDs Then
                sEnumProt = sEnumProt & "|" & sLSPName & " - " & sGUID & " - " & sFile
            Else
                sEnumProt = sEnumProt & "|" & sLSPName & " - " & sFile
            End If
        Next i
    End If
    
    WSACleanup
    
    If sEnumProt <> vbNullString Then EnumWinsockProtocol = Mid$(sEnumProt, 2)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modLSP.EnumWinsockProtocol"
    WSACleanup
    If inIDE Then Stop: Resume Next
End Function

Public Function EnumWinsockNameSpace$()
    On Error GoTo ErrorHandler:
    Dim lNumNameSpace&, sLSPName$, sEnumNamespace$
    Dim uWSANameSpaceInfo As WSANAMESPACE_INFO
    Dim uWSAData As WSAData, i&, sGUID$, sFile$
    Dim uBuffer() As Byte, lBufferSize&, strSize&
    
    If WSAStartup(&H202, VarPtr(uWSAData)) <> 0 Then Exit Function

    WSAEnumNameSpaceProviders lBufferSize, 0&
    ReDim uBuffer(lBufferSize - 1)
    
    lNumNameSpace = WSAEnumNameSpaceProviders(lBufferSize, VarPtr(uBuffer(0)))
    If lNumNameSpace <> SOCKET_ERROR Then
        For i = 0 To lNumNameSpace - 1
            memcpy ByVal VarPtr(uWSANameSpaceInfo), ByVal VarPtr(uBuffer(i * LenB(uWSANameSpaceInfo))), LenB(uWSANameSpaceInfo)
            sGUID = GuidToString(uWSANameSpaceInfo.NSProviderId)
            strSize = lstrlen(uWSANameSpaceInfo.lpszIdentifier)
            sLSPName = String$(strSize, 0)
            lstrcpyn StrPtr(sLSPName), uWSANameSpaceInfo.lpszIdentifier, strSize + 1
            sLSPName = TrimNull(sLSPName)
            sFile = GetNSProviderFile(sLSPName)
            If bShowCLSIDs Then
                sEnumNamespace = sEnumNamespace & "|" & sLSPName & " - " & sGUID & " - " & sFile
            Else
                sEnumNamespace = sEnumNamespace & "|" & sLSPName & " - " & sFile
            End If
        Next i
    End If

    WSACleanup
    
    If sEnumNamespace <> vbNullString Then EnumWinsockNameSpace = Mid$(sEnumNamespace, 2)

    Exit Function
ErrorHandler:
    ErrorMsg Err, "modLSP.EnumWinsockNameSpace"
    WSACleanup
    If inIDE Then Stop: Resume Next
End Function

Private Function GuidToString$(uGuid As UUID)
    Dim sGUID$
    sGUID = String$(39, 0)
    If StringFromGUID2(uGuid, StrPtr(sGUID), Len(sGUID)) > 0 Then
        GuidToString = TrimNull(sGUID)
    End If
End Function

Private Function GetProviderFile$(uProviderID As UUID)
    Dim sFile$, uFile() As Byte, lFileLen&, lErr&
    
    sFile = String$(MAX_PATH, 0)
    lFileLen = MAX_PATH
    
    ReDim uFile(lFileLen)
    If WSCGetProviderPath(VarPtr(uProviderID), StrPtr(sFile), VarPtr(lFileLen), VarPtr(lErr)) = 0 Then
        sFile = ExpandEnvironmentVars(TrimNull(sFile))
        GetProviderFile = sFile
    End If
End Function

Private Function GetNSProviderFile$(sName$)
    Dim sWS2Key$, sKeys$(), i&, sFile$, sDisplayName$, sBuf$
    sWS2Key = "System\CurrentControlSet\Services\Winsock2\Parameters\NameSpace_Catalog5\Catalog_Entries"
    sKeys = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sWS2Key), "|")
    For i = 0 To UBound(sKeys)
        sDisplayName = RegGetString(HKEY_LOCAL_MACHINE, sWS2Key & "\" & sKeys(i), "DisplayString")
        If Left$(sDisplayName, 1) = "@" Then
            sBuf = GetStringFromBinary(, , sDisplayName)
            If 0 <> Len(sBuf) Then sDisplayName = sBuf
        End If
        If sName = sDisplayName Then
            sFile = ExpandEnvironmentVars(RegGetString(HKEY_LOCAL_MACHINE, sWS2Key & "\" & sKeys(i), "LibraryPath"))
            GetNSProviderFile = sFile
            Exit For
        End If
    Next i
End Function

' ---------------------------------------------------------------------------------------------------
' HJT main routine
' ---------------------------------------------------------------------------------------------------

Public Sub GetLSPCatalogNames()
    sKeyNameSpace = "System\CurrentControlSet\Services\WinSock2\Parameters"
    sKeyProtocol = "System\CurrentControlSet\Services\WinSock2\Parameters"
    
    sKeyNameSpace = sKeyNameSpace & "\" & RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Current_NameSpace_Catalog")
    sKeyProtocol = sKeyProtocol & "\" & RegGetString(HKEY_LOCAL_MACHINE, sKeyProtocol, "Current_Protocol_Catalog")
End Sub

Public Sub CheckLSP()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "CheckLSP - Begin"
    
    Dim lNumNameSpace&, lNumProtocol&, i&, J& ', sSafeFiles$
    Dim sFile$, uData() As Byte, hKey&, sHit$, sDummy$, sFindFile$
    
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
            sFindFile = FindOnPath(sFile)
            If 0 <> Len(sFindFile) Then
                sFile = sFindFile
            
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
            
            sFindFile = FindOnPath(sFile)
            
            If 0 <> Len(sFindFile) Then
                sFile = sFindFile
                
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
    
    AppendErrorLogCustom "CheckLSP - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modLSP_CheckLSP"
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
    ErrorMsg Err, "FixLSP"
    If inIDE Then Stop: Resume Next
End Sub

'Public Sub FixLSP_Old()
'    On Error GoTo ErrorHandler:
'
'    Dim lNumNameSpace&, lNumProtocol&
'    Dim i&, J&, sFile$, hKey&, uData() As Byte
'
'    If Not bSeenLSPWarning Then
'        MsgBoxW "HiJackThis cannot repair O10 Winsock LSP entries. " & vbCrLf & _
'               "You should use LSPFix for that, which is available " & _
'               "from http://www.cexx.org/lspfix.htm." & vbCrLf & vbCrLf & _
'               "If the O10 item belongs to WebHancer, New.Net or CommonName, " & _
'               "Spybot S&D can remove it automatically. Spybot S&D " & _
'               "is available from http://www.spybot.info/.", vbCritical
'
'        bSeenLSPWarning = True
'    End If
'    Exit Sub
'
'    lNumNameSpace = RegGetDword(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries")
'    lNumProtocol = RegGetDword(HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries")
'
'    'check for missing files, delete keys with those
'    For i = 1 To lNumNameSpace
'        sFile = RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), "LibraryPath")
'        sFile = LCase(Replace$(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare))
'        sFile = LCase(Replace$(sFile, "%windir%", sWinDir, , , vbTextCompare))
'        If sFile <> vbNullString And DirW$(sFile, vbFile) <> vbNullString Then
'            'file ok
'            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Or _
'               InStr(1, sFile, "newdot", vbTextCompare) > 0 Or _
'               InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
'                'it's New.Net/WebHancer/CN! Kill it!
'                DeleteFileWEx StrPtr(sFile)  ' error 53 = file not found
'                If FileExists(sFile) Then
'                    If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Then
'                        MsgBoxW "The WebHancer Agent is currently active and can't be deleted. Use Ad-Aware from www.lavasoft.nu to remove it safely.", vbExclamation
'                    ElseIf InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
'                        SHRestartSystemMB frmMain.hWnd, "The NewDotNet DLL is currently active. You will need to reboot and rescan, then remove New.Net and fix the WinSock stack." & vbCrLf & vbCrLf, 0
'                        Exit Sub
'                    ElseIf InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
'                        SHRestartSystemMB frmMain.hWnd, "The CommonName DLL is currently active. You will need to reboot and rescan, then remove CommonName and fix the WinSock stack." & vbCrLf & vbCrLf, 0
'                    End If
'                End If
'
'                RegDelKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
'                lNumNameSpace = lNumNameSpace - 1
'
'                'delete New.Net startup Reg entry
'                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "New.Net Startup"
'                'delete WebHancer startup Reg entry
'                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "webHancer Agent"
'                'delete CommonName startup Reg entry
'                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Zenet"
'            End If
'        Else
'            If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
'                lNumNameSpace = lNumNameSpace - 1
'            End If
'            RegDelKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
'        End If
'    Next i
'
'    For i = 1 To lNumProtocol
'        sFile = RegGetFileFromBinary(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), "PackedCatalogItem")
'        sFile = LCase(Replace$(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare))
'        sFile = LCase(Replace$(sFile, "%windir%", sWinDir, , , vbTextCompare))
'        If sFile <> vbNullString And DirW$(sFile, vbFile) <> vbNullString Then
'            'file ok
'            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Or _
'               InStr(1, sFile, "newdotnet", vbTextCompare) > 0 Or _
'               InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
'                'it's New.Net/WebHancer/CN! Kill it!
'                DeleteFileWEx StrPtr(sFile)  ' error 53 = file not found
'                If FileExists(sFile) Then
'                    If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Then
'                        MsgBoxW "The WebHancer Agent is currently active and can't be deleted. Use Ad-Aware from www.lavasoft.nu to remove it safely.", vbExclamation
'                    ElseIf InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
'                        SHRestartSystemMB frmMain.hWnd, "The NewDotNet DLL is currently active. You will need to reboot and rescan, then remove New.Net and fix the WinSock stack." & vbCrLf & vbCrLf, 0
'                        Exit Sub
'                    ElseIf InStr(1, sFile, "cnmib.dll", vbTextCompare) > 0 Then
'                        SHRestartSystemMB frmMain.hWnd, "The CommonName DLL is currently active. You will need to reboot and rescan, then remove CommonName and fix the WinSock stack." & vbCrLf & vbCrLf, 0
'                    End If
'                End If
'
'                RegDelKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
'                lNumNameSpace = lNumNameSpace - 1
'
'                'delete New.Net startup Reg entry
'                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "New.Net Startup"
'
'                'delete WebHancer startup Reg entry
'                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "webHancer Agent"
'
'                'delete CommonName startup Reg entry
'                RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", "Zenet"
'            End If
'        Else
'            If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
'                lNumProtocol = lNumProtocol - 1
'            End If
'            RegDelKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
'        End If
'    Next i
'
'    'check LSP chain, fix gaps where found
'    i = 1 'current LSP #
'    J = 1 'correct LSP #
'    Do
'        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
'            If i > J Then
'                RegRenameKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(J)), "0") & CStr(J)
'            End If
'            J = J + 1
'        Else
'            'nothing, j stays the same
'        End If
'        i = i + 1
'        'check to prevent infinite loop when
'        'lNumNameSpace is wrong
'        If i = 100 Then
'            lNumNameSpace = J - 1
'            Exit Do
'        End If
'    Loop Until J = lNumNameSpace + 1
'
'    i = 1
'    J = 1
'    Do
'        If RegKeyExists(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
'            If i > J Then
'                RegRenameKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(J)), "0") & CStr(J)
'            End If
'            J = J + 1
'        Else
'            'nothing, j stays the same
'        End If
'        i = i + 1
'        If i = 100 Then
'            lNumProtocol = J - 1
'            Exit Do
'        End If
'    Loop Until J = lNumProtocol + 1
'
'    RegSetDwordVal HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries", lNumNameSpace
'    RegSetDwordVal HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries", lNumProtocol
'
'    bRebootNeeded = True
'    Exit Sub
'
'ErrorHandler:
'    ErrorMsg err, "modLSP_FixLSP"
'    RegCloseKey hKey
'End Sub

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
    ErrorMsg Err, "modLSP_RegRenameKey", "sKeyOldName=", sKeyOldName, "sKeyNewName=", sKeyNewName
    RegCloseKey hKey
    RegCloseKey hKey2
    If inIDE Then Stop: Resume Next
End Sub

