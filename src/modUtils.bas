Attribute VB_Name = "modUtils"
'[modUtils.bas]

'
' Various helper functions
'

Option Explicit

Private Type BROWSER_INFO
    Version As String
    'Path    As String  'future use
End Type

Public Type BROWSERS_VERSION_INFO
    Edge    As BROWSER_INFO
    IE      As BROWSER_INFO
    Firefox As BROWSER_INFO
    Opera   As BROWSER_INFO
    Chrome  As BROWSER_INFO
    Default As String
End Type

Public Enum SETTINGS_SECTION
    SETTINGS_SECTION_MAIN = 0
    SETTINGS_SECTION_ADSSPY
    SETTINGS_SECTION_SIGNCHECKER
    SETTINGS_SECTION_PROCMAN
    SETTINGS_SECTION_STARTUPLIST
    SETTINGS_SECTION_UNINSTMAN
    SETTINGS_SECTION_REGUNLOCKER
    SETTINGS_SECTION_FILEUNLOCKER
    SETTINGS_SECTION_REGKEYTYPECHECKER
    SETTINGS_SECTION_HOSTSMAN
End Enum

Public hLibPcre2        As Long
Public oRegexp          As IRegExp
Public g_bRegexpInit    As Boolean

Private lSubclassedTools    As Long
Private lSubclassedScan     As Long
Private hGetMsgHook         As Long
Private m_hColTextbox       As New clsCollectionEx

Public inIDE As Boolean

Public Sub SubClassScroll(SwitchON As Boolean)
    
    SubClassScroll_Tools SwitchON
    SubClassScroll_ScanResults SwitchON
    
    'SubClassScroll_Hotkeys SwitchON 'Replaced by Form's "KeyPreview" property
    
End Sub

Public Sub SubClassScroll_Tools(SwitchON As Boolean)
    If DisableSubclassing Then Exit Sub
    If SwitchON And Not bAutoLogSilent Then
        If lSubclassedTools = 0 Then lSubclassedTools = SetWindowSubclass(g_HwndMain, AddressOf WndProcTools, 0&)
    Else
        If lSubclassedTools Then RemoveWindowSubclass g_HwndMain, AddressOf WndProcTools, 0&: lSubclassedTools = 0
    End If
End Sub

Public Sub SubClassScroll_ScanResults(SwitchON As Boolean)
    If DisableSubclassing Then Exit Sub
    If SwitchON And Not bAutoLogSilent Then
        If Not (frmMain Is Nothing) Then
            If lSubclassedScan Then SubClassScroll_ScanResults False
            h_HwndScanResults = frmMain.lstResults.hWnd
            If lSubclassedScan = 0 Then lSubclassedScan = SetWindowSubclass(h_HwndScanResults, AddressOf WndProcScan, 0&)
        End If
    Else
        If lSubclassedScan Then RemoveWindowSubclass h_HwndScanResults, AddressOf WndProcScan, 0&: lSubclassedScan = 0
    End If
End Sub

Public Sub SubClassTextbox(hTextbox As Long, SwitchON As Boolean)
    'If DisableSubclassing Then Exit Sub
    If SwitchON Then
        If Not m_hColTextbox.ItemExists(hTextbox) Then
            SetWindowSubclass hTextbox, AddressOf Callback_WndTextbox, 0&
            m_hColTextbox.Add hTextbox
        End If
    Else
        If hTextbox = -1 Then
            Dim i As Long
            For i = 1 To m_hColTextbox.Count
                RemoveWindowSubclass m_hColTextbox(i), AddressOf Callback_WndTextbox, 0&
            Next
            m_hColTextbox.RemoveAll
        Else
            If m_hColTextbox.ItemExists(hTextbox) Then
                RemoveWindowSubclass hTextbox, AddressOf Callback_WndTextbox, 0&
                m_hColTextbox.RemoveByItem hTextbox
            End If
        End If
    End If
End Sub

Public Function IsMouseWithin(hWnd As Long) As Boolean
    Dim r As RECT
    Dim p As POINTAPI
    If GetWindowRect(hWnd, r) Then
        If GetCursorPos(p) Then
            IsMouseWithin = PtInRect(r, p.x, p.y)
        End If
    End If
End Function

Public Function GetCursorPosRel() As POINTAPI
    Dim p As POINTAPI
    If GetCursorPos(p) Then
        ScreenToClient g_HwndMain, p
    End If
    GetCursorPosRel = p
End Function

'Public Sub SubClassScroll_Hotkeys(SwitchON As Boolean)
'    If DisableSubclassing Then Exit Sub
'    If SwitchON And Not bAutoLogSilent Then
'        'hotkeys support (Thanks to ManHunter)
'        If hGetMsgHook = 0 Then hGetMsgHook = SetWindowsHookEx(WH_GETMESSAGE, AddressOf GetMsgProc, 0, App.ThreadId)
'    Else
'        If hGetMsgHook Then UnhookWindowsHookEx hGetMsgHook: hGetMsgHook = 0
'    End If
'End Sub

'Private Function GetMsgProc(ByVal nCode As Long, ByVal wParam As Long, lParam As msg) As Long
'    If lParam.message = WM_KEYDOWN Then
'        'http://www.manhunter.ru/assembler/878_obrabotka_soobscheniy_ot_klaviaturi_v_dialogbox.html
'
'        If nCode = HC_ACTION Then
'            'Debug.Print "MSG: " & Hex$(lParam.message) & ", " & _
'                "HWND: " & Hex$(lParam.hwnd) & ", " & _
'                "WPARAM: " & Hex$(lParam.wParam) & ", " & _
'                "LPARAM: " & Hex$(lParam.lParam) & ", " & _
'                "MSGREMOVED: " & wParam
'
'            If (inIDE And wParam = 0) Or Not inIDE Then
'                If lParam.wParam = Asc("A") Then                    'Ctrl + A
'                    If cMath.HIWORD(GetKeyState(VK_CONTROL)) Then
'                        MsgBoxW "Ctrl + A pressed !!!"
'                    End If
'
'                ElseIf lParam.wParam = Asc("F") Then                'Ctrl + F
'                    If cMath.HIWORD(GetKeyState(VK_CONTROL)) Then
'                        'MsgBoxW "Ctrl + F pressed !!!"
'                        If IsFormInit(frmSearch) Then
'                            frmSearch.Display
'                        Else
'                            Load frmSearch
'                        End If
'                        Exit Function
'                    End If
'                End If
'            End If
'        End If
'    End If
'    GetMsgProc = CallNextHookEx(hGetMsgHook, nCode, wParam, lParam)
'End Function

Private Function WndProcTools(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    On Error Resume Next
    
    Static MouseKeys&, Rotation&, NewValue%
    
    Select Case uMsg
    
    Case WM_NCDESTROY
        SubClassScroll_Tools False
        
    Case WM_UAHDESTROYWINDOW 'dilettante's trick
        SubClassScroll_Tools False
        
    Case WM_MOUSEWHEEL

        If Not IsMouseWithin(g_HwndMain) Then Exit Function ' mouse is outside the form
        
        If g_CurFrame = FRAME_ALIAS_MISC_TOOLS Then
            
            MouseKeys = wParam And &HFFFF&
            Rotation = wParam \ &HFFFF& 'direction
            With frmMain.vscMiscTools
                NewValue = .Value - .LargeChange * IIf(Rotation > 0, 1, -1)
                If NewValue < .Min Then NewValue = .Min
                If NewValue > .Max Then NewValue = .Max
                .Value = NewValue   'change scroll value
            End With
        End If
    
    'Case WM_KEYDOWN
    '  - is not working here because msg is intercepted by active control.
    ' WH_GETMESSAGE hook is required. To catch msg here, you can use SendMessage in GetMsg callback.

    'Case WM_HOTKEY
    '    If wParam = HOTKEY_ID_CTRL_F Then DoSearchWindow
    '    WndProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    ' Not the best option, because RegisterHotKey() intercepts hotkeys from whole system!
    ' As well as not allows to use them by another programs until UnregisterHotKey() call.
    
    End Select
    
    WndProcTools = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    
End Function

Private Function WndProcScan(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    On Error Resume Next
    
    Static Rotation&
    
    Select Case uMsg
    
    Case WM_NCDESTROY
        SubClassScroll_ScanResults False
        
    Case WM_UAHDESTROYWINDOW 'dilettante's trick
        SubClassScroll_ScanResults False
        
    Case WM_MOUSEWHEEL

        If Not IsMouseWithin(h_HwndScanResults) Then Exit Function ' mouse is outside the form
        
        If g_CurFrame = FRAME_ALIAS_SCAN Then
                
            If cMath.HIWORD(GetKeyState(VK_CONTROL)) Then
                Rotation = wParam \ &HFFFF& 'direction
                Call SetFontSizeDelta(IIf(Rotation > 0, 1, -1))
            End If
                
        End If
    
    End Select
    
    WndProcScan = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    
End Function

Private Function Callback_WndTextbox(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    On Error Resume Next
    
    Select Case uMsg
    
    Case WM_NCDESTROY
        SubClassTextbox -1, False
        
    Case WM_UAHDESTROYWINDOW 'dilettante's trick
        SubClassTextbox -1, False
        
    Case WM_KEYDOWN
        If wParam = Asc("A") And GetKeyState(VK_CONTROL) < 0 Then
            SendMessage hWnd, EM_SETSEL, ByVal 0&, ByVal -1&
            Exit Function
        End If
        
    Case WM_PASTE
        Dim txt As String
        txt = ClipboardGetText()
        If Len(txt) <> 0 Then
            SendMessage hWnd, EM_REPLACESEL, True, ByVal StrPtr(txt)
            Exit Function
        End If
    
    Case WM_COPY
        Callback_WndTextbox = DefSubclassProc(hWnd, uMsg, wParam, lParam)
        If OpenClipboardEx(hWnd) Then
            Dim hMem As Long
            Dim ptr As Long
            hMem = GlobalAlloc(GMEM_MOVEABLE, 4)
            If hMem <> 0 Then
                ptr = GlobalLock(hMem)
                If ptr <> 0 Then
                    GetMem4 OSver.LangNonUnicodeCode, ByVal ptr
                    GlobalUnlock hMem
                    If SetClipboardData(CF_LOCALE, hMem) = 0 Then
                        GlobalFree hMem
                    End If
                End If
            End If
            CloseClipboard
        End If
        Exit Function
    
    End Select
    
    Callback_WndTextbox = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    
End Function

Private Sub SetFontSizeDelta(delta As Long)
    On Error GoTo ErrorHandler:
    
    Dim lFontSize As Long
    
    g_FontName = frmMain.cmbFont.List(frmMain.cmbFont.ListIndex)
    g_FontSize = frmMain.cmbFontSize.List(frmMain.cmbFontSize.ListIndex)
    
    If g_FontSize = "Auto" Or Len(g_FontSize) = 0 Then
        lFontSize = 8
    Else
        lFontSize = CLng(g_FontSize)
    End If
    
    If (g_FontName = "MS Sans Serif") And ((lFontSize Mod 2) = 0) Then delta = delta * 2
    
    lFontSize = lFontSize + delta
    
    If lFontSize < 6 Then lFontSize = 6
    If lFontSize > 14 Then lFontSize = 14
    
    ComboSetValue frmMain.cmbFontSize, CStr(lFontSize)
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SetFontSizeDelta"
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetStringFromBinary(Optional ByVal sFile As String, Optional ByVal nid As Long, Optional ByVal FileAndIDHybrid As String) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetStringFromBinary - Begin", "File: " & sFile, "id: " & nid, "FileAndIDHybrid: " & FileAndIDHybrid
    
    Dim hModule As Long
    Dim nSize As Long
    Dim sBuf As String
    Dim pos As Long
    Dim Redirect As Boolean, bOldStatus As Boolean
    Dim sResVar As String
    Dim sInitialVar As String
    Dim bIsInf As Boolean
    
    'Get string resource from binary file
    'Source can be defined either by Filename and ID, or by hibryd (registry like) string, e.g. @%SystemRoot%\System32\my.dll,-102
    'ID with minus will be converted by module
    
    sInitialVar = FileAndIDHybrid
    
    If 0 <> Len(FileAndIDHybrid) Then
    
        If Left$(FileAndIDHybrid, 1) = "@" Then FileAndIDHybrid = mid$(FileAndIDHybrid, 2)
        If InStr(FileAndIDHybrid, "%") <> 0 Then
            If InStr(1, FileAndIDHybrid, ".inf", 1) = 0 Then
                FileAndIDHybrid = EnvironW(FileAndIDHybrid)
            End If
        End If
        pos = InStrRev(FileAndIDHybrid, ",")
        If 0 <> pos Then
            sFile = Left$(FileAndIDHybrid, pos - 1)
            sBuf = mid$(FileAndIDHybrid, pos + 1)
            
            If StrComp(GetExtensionName(sFile), ".inf", 1) = 0 Then
                bIsInf = True
                '@oem28.inf,%ImcSvcDisplayName%;System Interface Foundation Service
                '(with or without ; (semicolon) mark)
                pos = InStr(sBuf, ";")
                If pos <> 0 Then
                    sBuf = Left$(sBuf, pos - 1) 'remove name part
                End If
                sResVar = Trim$(sBuf)
                If Left$(sResVar, 1) = "%" And Right$(sResVar, 1) = "%" Then
                    sResVar = mid$(sResVar, 2, Len(sResVar) - 2)
                End If
            ElseIf IsNumeric(sBuf) And 0 <> Len(sBuf) Then
                nid = Val(sBuf)
            Else
                Exit Function
            End If
        End If
    End If
    
    If 0 = Len(sFile) Then Exit Function
    
    sFile = EnvironW(sFile)
    
    If Not FileExists(sFile) Then
        sFile = FindOnPath(sFile, , IIf(bIsInf, BuildPath(sWinDir, "inf"), vbNullString))
        If 0 = Len(sFile) Then Exit Function
    End If
    
    sBuf = String$(160, 0)
    
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    If bIsInf Then
        nSize = GetPrivateProfileString(StrPtr("Strings"), StrPtr(sResVar), StrPtr(sInitialVar), StrPtr(sBuf), Len(sBuf), StrPtr(sFile))
        If nSize <> 0 Then
            sBuf = UnQuote(Left$(sBuf, nSize))
        End If
        GetStringFromBinary = sBuf
    Else
        hModule = LoadLibraryEx(StrPtr(sFile), 0&, LOAD_LIBRARY_AS_DATAFILE)

        If hModule Then
            nSize = LoadString(hModule, Abs(nid), StrPtr(sBuf), LenB(sBuf))
            If nSize > 0 Then
                GetStringFromBinary = TrimNull(Left$(sBuf, nSize))
            End If
            FreeLibrary hModule
        End If
    End If
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "GetStringFromBinary - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetStringFromBinary", sFile
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function GetBrowsersInfo() As BROWSERS_VERSION_INFO
    AppendErrorLogCustom "GetBrowsersInfo - Begin"
    
    Dim Cmd As String
    Dim FriendlyName As String
    Dim path As String
    Dim Arguments As String
    Cmd = GetDefaultApp("http", FriendlyName)
    If Len(Cmd) = 0 Then
        Cmd = "Program is not associated"
    Else
        SplitIntoPathAndArgs Cmd, path, Arguments
    End If
    
    With GetBrowsersInfo
        .Edge.Version = GetEdgeVersion()
        .IE.Version = GetMSIEVersion()
        .Chrome.Version = GetChromeVersion()
        .Firefox.Version = GetFirefoxVersion()
        .Opera.Version = GetOperaVersion()
        .Default = Cmd & IIf(path = "(AppID)", " " & FriendlyName, IIf(Len(FriendlyName) <> 0, " (" & FriendlyName & ")", vbNullString))
    End With
    AppendErrorLogCustom "GetBrowsersInfo - End"
End Function

Public Function GetEdgeVersion() As String
    AppendErrorLogCustom "GetEdgeVersion - Begin"
    Dim EdgePath$
    '// TODO
    'maybe - HKCR\ActivatableClasses\Package\Microsoft.MicrosoftEdge_44.19041.423.0_neutral__8wekyb3d8bbwe ?
    EdgePath = sWinDir & "\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe"
    If FileExists(EdgePath) Then GetEdgeVersion = GetFilePropVersion(EdgePath)
    AppendErrorLogCustom "GetEdgeVersion - End"
End Function

Public Function GetChromeVersion() As String
    AppendErrorLogCustom "GetChromeVersion - Begin"
    Dim sVer$, sPath$
    sVer = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv")
    'not found try current user - win7(x86)
    If Len(sVer) = 0 Then
        sVer = Reg.GetString(HKEY_CURRENT_USER, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv")
    End If
    If Len(sVer) = 0 Then 'Wow6432Node
        sVer = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv", True)
    End If
    If Len(sVer) = 0 Then
        sPath = Reg.GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", vbNullString)
        If Len(sPath) <> 0 Then
            sVer = GetFilePropVersion(sPath)
        End If
    End If
    If Len(sVer) = 0 Then
        sVer = Reg.GetString(HKEY_CURRENT_USER, "SOFTWARE\Google\Chrome\BLBeacon", "version")
    End If
    GetChromeVersion = sVer
    AppendErrorLogCustom "GetChromeVersion - End"
End Function

Public Function GetFirefoxVersion() As String
    AppendErrorLogCustom "GetFirefoxVersion - Begin"
    Dim sVer$, sPath$
    
    sPath = Reg.GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe", vbNullString)
    If Len(sPath) <> 0 Then
        sVer = GetFilePropVersion(sPath)
        If Len(sVer) = 0 Then sVer = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Mozilla\Mozilla Firefox", "CurrentVersion")
    End If
    
    GetFirefoxVersion = sVer
    AppendErrorLogCustom "GetFirefoxVersion - End"
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetOperaVersion
' Purpose   : Gets the version of the installed Opera browser
' Return    : The version or an empty string if it cannot be found
'---------------------------------------------------------------------------------------
' Revision History:
' Date       Author        Purpose
' ---------  ------------  -------------------------------------------------------------
' 02Jul2013  Claire Streb  Original
' 24.03.15   Alex Dragokas Reworked/Simplified
'
Public Function GetOperaVersion() As String
    AppendErrorLogCustom "GetOperaVersion - Begin"
    Dim sOperaPath$
    sOperaPath = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\Opera.exe", vbNullString)

    If Len(sOperaPath) <> 0 Then
        sOperaPath = UnQuote(sOperaPath)
        If FileExists(sOperaPath) Then GetOperaVersion = GetFilePropVersion(sOperaPath)
    End If
    AppendErrorLogCustom "GetOperaVersion - End"
End Function

Public Function GetMSIEVersion() As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetMSIEVersion - Begin"
    
    Dim sMSIEPath$, sMSIEVersion$, sMSIEHotfixes$, HotFixVer$, i&
       
    If 0 = Len(sMSIEVersion) Then
        sMSIEPath = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE", vbNullString)
        If sMSIEPath <> vbNullString Then
            sMSIEPath = Trim$(sMSIEPath)
            If FileExists(sMSIEPath) Then
                sMSIEVersion = GetFilePropVersion(sMSIEPath)
            End If
        End If
    End If
       
    If 0 = Len(sMSIEVersion) Then
        sMSIEVersion = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "svcVersion")
        If 0 = Len(sMSIEVersion) Then
            sMSIEVersion = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "version")
        End If
    End If
    
    sMSIEHotfixes = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MinorVersion")
    
    If Len(sMSIEHotfixes) <> 0 And sMSIEHotfixes <> "0" Then
        For i = 5 To 1 Step -1
            If InStr(1, sMSIEHotfixes, "SP" & i, vbTextCompare) > 0 Then
                HotFixVer = "SP" & i
                Exit For
            End If
        Next
    End If
    
    If Len(sMSIEVersion) <> 0 Then
        GetMSIEVersion = sMSIEVersion & IIf(Len(HotFixVer) <> 0, " " & HotFixVer, vbNullString)
    Else
        GetMSIEVersion = "Unknown" '"Unable to get Internet Explorer version!"
    End If
    
    AppendErrorLogCustom "GetMSIEVersion - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetMSIEVersion"
    If inIDE Then Stop: Resume Next
End Function

Public Function IsSignPresent(FileName As String) As Boolean
    ' &H3C -> PE_Header offset
    ' PE_Header offset + &H18 = Optional_PE_Header
    ' PE_Header offset + &H78 = Data_Directories offset
    ' Data_Directories offset + &H20 = SecurityDir -> Address (dword), Size (dword) for digital signature.
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "IsSignPresent - Begin", "File: " & FileName
    
    Const IMAGE_FILE_MACHINE_I386   As Long = &H14C&
    Const IMAGE_FILE_MACHINE_IA64   As Long = &H200&
    Const IMAGE_FILE_MACHINE_AMD64  As Long = &H8664&
    
    Dim ff              As Integer
    Dim PE_offset       As Long
    Dim SignAddress     As Long
    Dim DataDir_offset  As Long
    Dim DirSecur_offset As Long
    Dim Machine(1)      As Byte
    Dim FSize           As Long
    Dim Redirect As Boolean, bOldStatus As Boolean
  
    Redirect = ToggleWow64FSRedirection(False, FileName, bOldStatus)
  
    ff = FreeFile()
    Open FileName For Binary Access Read Shared As #ff
    FSize = LOF(ff)
    
    If FSize >= &H3C& + 6& Then
        Get #ff, &H3C + 1&, PE_offset
        Get #ff, PE_offset + 4& + 1&, Machine(0)
        
        Select Case Machine(0) + CLng(Machine(1)) * 256&
            Case IMAGE_FILE_MACHINE_I386
                DataDir_offset = PE_offset + &H78&
            Case IMAGE_FILE_MACHINE_AMD64, IMAGE_FILE_MACHINE_IA64
                DataDir_offset = PE_offset + &H88&
            Case Else   'unknown architecture, not PE EXE or damaged image
                Close #ff
                Exit Function
        End Select
        
        DirSecur_offset = DataDir_offset + &H20
        
        If FSize >= DirSecur_offset + 4& Then Get #ff, DirSecur_offset + 1&, SignAddress
    End If
    
    Close #ff
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    IsSignPresent = (SignAddress <> 0)
    
    AppendErrorLogCustom "IsSignPresent - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modUtils_IsSignPresent", "File:", FileName
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function TryUnlock(ByVal FS_Object As String, Optional bRecursive As Boolean) As Boolean

    AppendErrorLogCustom "TryUnlock - Begin", "File: " & FS_Object
    
    TryUnlock = SetFileStringSD(FS_Object, GetDefaultFileSDDL(), bRecursive)
    
    AppendErrorLogCustom "TryUnlock - End"
End Function

Public Function AppPath(Optional bGetFullPath As Boolean) As String
    On Error GoTo ErrorHandler

    Static ProcPathFull  As String
    Static ProcPathShort As String
    Dim ProcPath As String
    Dim cnt      As Long
    Dim hProc    As Long
    Dim pos      As Long
    
    'Cache
    If bGetFullPath Then
        If Len(ProcPathFull) <> 0 Then
            AppPath = ProcPathFull
            Exit Function
        End If
    Else
        If Len(ProcPathShort) <> 0 Then
            AppPath = ProcPathShort
            Exit Function
        End If
    End If
    
    If inIDE Then
        If bGetFullPath Then
            AppPath = GetDOSFilename(App.path, bReverse:=True) & "\" & GetValueFromVBP(BuildPath(App.path, App.ExeName & ".vbp"), "ExeName32")
            ProcPathFull = AppPath
        Else
            AppPath = GetDOSFilename(App.path, bReverse:=True)
            ProcPathShort = AppPath
        End If
        Exit Function
    End If

    hProc = GetModuleHandle(0&)
    If hProc < 0 Then hProc = 0

    ProcPath = String$(MAX_PATH, vbNullChar)
    cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath)) 'hproc can be 0 (mean - current process)
    
    If cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
        ProcPath = Space$(MAX_PATH_W)
        cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
    End If
    
    If cnt = 0 Then                          'clear path
        ProcPath = App.path
    Else
        ProcPath = Left$(ProcPath, cnt)
        If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = sWinDir & mid$(ProcPath, 12)
        If "\??\" = Left$(ProcPath, 4) Then ProcPath = mid$(ProcPath, 5)
        
        If Not bGetFullPath Then
            ' trim to path
            pos = InStrRev(ProcPath, "\")
            If pos <> 0 Then ProcPath = Left$(ProcPath, pos - 1)
        End If
    End If
    
    ProcPath = GetDOSFilename(ProcPath, bReverse:=True)     '8.3 -> to Full
    
    AppPath = ProcPath
    
    If bGetFullPath Then
        ProcPathFull = ProcPath
    Else
        ProcPathShort = ProcPath
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.AppPath", "ProcPath:", ProcPath
    If inIDE Then Stop: Resume Next
End Function

Public Function AppExeName(Optional WithExtension As Boolean) As String
    On Error GoTo ErrorHandler

    Static ProcPathShort As String
    Static ProcPathFull  As String
    Dim ProcPath As String
    Dim cnt      As Long
    Dim hProc    As Long
    Dim pos      As Long

    'Cache
    If WithExtension Then
        If Len(ProcPathFull) <> 0 Then
            AppExeName = ProcPathFull
            Exit Function
        End If
    Else
        If Len(ProcPathShort) <> 0 Then
            AppExeName = ProcPathShort
            Exit Function
        End If
    End If

    If inIDE Then
        AppExeName = App.ExeName & IIf(WithExtension, ".exe", vbNullString)
        Exit Function
    End If

    hProc = GetModuleHandle(0&)
    If hProc < 0 Then hProc = 0

    ProcPath = String$(MAX_PATH, vbNullChar)
    cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath)) 'hproc can be 0 (mean - current process)
    
    If cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
        ProcPath = Space$(MAX_PATH_W)
        cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
    End If
    
    If cnt = 0 Then                          'clear path
        ProcPath = App.ExeName & IIf(WithExtension, ".exe", vbNullString)
    Else
        ProcPath = Left$(ProcPath, cnt)
        
        pos = InStrRev(ProcPath, "\")
        If pos <> 0 Then ProcPath = mid$(ProcPath, pos + 1)
        
        If Not WithExtension Then
            ProcPath = GetFileName(ProcPath)
        End If
    End If
    
    AppExeName = ProcPath
    
    If WithExtension Then
        ProcPathFull = ProcPath
    Else
        ProcPathShort = ProcPath
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.AppExeName", "ProcPath:", ProcPath
    If inIDE Then Stop: Resume Next
End Function

' argv(0) - this applications' exe
' argv(1 - ... argc) - tokens from the incoming "Line"
'
Public Function ParseCommandLine(Line As String, argc As Long, argv() As String, Optional bFirstArgIncomedAsAppExe As Boolean = False) As Boolean
  On Error GoTo ErrorHandler
  Dim Lex$(), nL&, nA&, Unit$, St$
  St = Line
  If Len(St) > 0 Then ParseCommandLine = True
  Lex = Split(St) '–азбиваем по пробелам на лексемы дл€ анализа знаков
  ReDim argv(0 To UBound(Lex) + 1) As String 'ќпредел€ем выходной массив до максимально возможного числа параметров
  If Not bFirstArgIncomedAsAppExe Then
    argv(0) = AppPath(True)
  End If
  If Len(St) <> 0 Then
    Do While nL <= UBound(Lex)
      Unit = Lex(nL) '«аписысаем текущую лексему как начало нового аргумента
      If Len(Unit) <> 0 Then '«ащита от двойных пробелов между аргументами
        'если в лексеме найдена кавычка или непарное их число, то начинаем процесс "квотировани€"
        If (Len(Lex(nL)) - Len(Replace$(Lex(nL), """", vbNullString))) Mod 2 = 1 Then
          Do
            nL = nL + 1
            If nL > UBound(Lex) Then Exit Do '≈сли не дождались завершающей кавычки, а больше лексем нет
            Unit = Unit & " " & Lex(nL) 'дополн€ем соседней лексемой
          ' аргумент должен завершатьс€ 1 или непарным числом кавычек лексемы со всеми прил€гающими к ней справа символами (кроме знака пробела)
          Loop Until (Len(Lex(nL)) - Len(Replace$(Lex(nL), """", vbNullString))) Mod 2 = 1
        End If
        Unit = Replace$(Unit, """", vbNullString) '”дал€ем кавычки
        nA = nA + 1 '—четчик кол-ва выходных аргументов
        If bFirstArgIncomedAsAppExe Then
          If nA > 1 Then
            argv(nA - 1) = Unit
          Else
            argv(0) = Unit
          End If
        Else
          argv(nA) = Unit
        End If
      End If
      nL = nL + 1 '—четчик текущей лексемы
    Loop
  End If
  If bFirstArgIncomedAsAppExe Then
    argc = nA - 1
  Else
    argc = nA
  End If
  If argc < 0 Then argc = 0
  ReDim Preserve argv(0 To argc) ' урезаем массив до реального числа аргументов
  Exit Function
ErrorHandler:
  ErrorMsg Err, "Parser.ParseCommandLine", "CmdLine:", Line
  If inIDE Then Stop: Resume Next
End Function

Function ExtractFilesFromCommandLine(sCmdLine As String) As String()
    Dim argc As Long
    Dim argv() As String
    Dim i As Long
    Dim N As Long
    Dim aPath() As String
    
    If ParseCommandLine(sCmdLine, argc, argv) Then
        For i = 1 To argc
            argv(i) = PathNormalize(argv(i))
            If FileExists(argv(i)) Then
                ArrayAddStr aPath, argv(i)
            End If
        Next
    End If
    ExtractFilesFromCommandLine = aPath
End Function

'By default allowing remove file on reboot
'
Public Function DeleteFileForce(sPath As String, Optional bForceMicrosoft As Boolean, Optional DisallowRemoveOnReboot As Boolean = False) As Boolean
    DeleteFileForce = DeleteFileEx(sPath, bForceMicrosoft, DisallowRemoveOnReboot)
End Function

'Delete File with unlocking DACL on failure
'
Public Function DeleteFileEx(sPath As String, Optional ForceDeleteMicrosoft As Boolean, Optional DisallowRemoveOnReboot As Boolean = True) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "DeleteFileEx - Begin"
    
    Dim iAttr As Long, lr As Long, sExt As String, sNewPath As String, lpSTR As Long
    Dim Redirect As Boolean, bOldStatus As Boolean, bMicrosoft As Boolean
    
    If Len(sPath) = 0 Then Exit Function
    
    lpSTR = StrPtr(sPath)
    
    sExt = GetExtensionName(sPath)
    
    If Not ForceDeleteMicrosoft Then
        If Not StrInParamArray(sExt, ".txt", ".log", ".tmp", ".ini") Then
            bMicrosoft = IsLolBin_ProtectedList(sPath)
            If Not bMicrosoft Then
                bMicrosoft = IsMicrosoftFile(sPath, True)
            End If
            If Not bMicrosoft Then
                bMicrosoft = IsFileSFC(sPath)
            End If
            If bMicrosoft Then
                SFC_RestoreFile sPath
                Exit Function
            End If
        End If
    End If
    
    If g_bDelModePending Then
        DeleteFileOnReboot sPath, bNoReboot:=True
        Exit Function
    End If
    
    Redirect = ToggleWow64FSRedirection(False, sPath, bOldStatus)
    
    iAttr = GetFileAttributes(lpSTR)
    If iAttr <> INVALID_FILE_ATTRIBUTES Then
        If (iAttr And (FILE_ATTRIBUTE_COMPRESSED Or FILE_ATTRIBUTE_READONLY)) Then
            iAttr = iAttr And Not (FILE_ATTRIBUTE_COMPRESSED Or FILE_ATTRIBUTE_READONLY)
            SetFileAttributes lpSTR, iAttr
        End If
    End If
    
    lr = DeleteFileW(lpSTR)
    
    If lr <> 0 Then 'success
        DeleteFileEx = True
        GoTo Finalize
    End If
    
    If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
        DeleteFileEx = True
        GoTo Finalize
    End If
    
    If Err.LastDllError = ERROR_ACCESS_DENIED Then
        TryUnlock sPath
        SetFileAttributes lpSTR, FILE_ATTRIBUTE_NORMAL
        lr = DeleteFileW(lpSTR)
    End If
    
    If lr = 0 Then 'if process still run, try rename file
        sNewPath = GetEmptyName(sPath & ".bak")
        
        'if failed
        If 0 = MoveFile(StrPtr(sPath), StrPtr(sNewPath)) Then
            'plan to delete on reboot
            If Not DisallowRemoveOnReboot Then
                DeleteFileOnReboot sPath, bNoReboot:=True
                bRebootRequired = True
            End If
        Else
            DeleteFileEx = True
        End If
    Else
        DeleteFileEx = True
    End If
    
Finalize:
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "DeleteFileEx - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DeleteFileEx", "File:", sPath
    If inIDE Then Stop: Resume Next
End Function

Public Function ConvertUnixTimeToLocalDate(Seconds As Long) As Date
    On Error GoTo ErrorHandler
    'Dim ftime               As FILETIME
    'Dim TimeZoneInfo(171)   As Byte
    Dim sTime               As SYSTEMTIME
    Dim DateTime            As Date
    
    DateTime = DateAdd("s", Seconds, #1/1/1970#)    ' time_t -> Date
    VariantTimeToSystemTime DateTime, sTime         ' Date -> SYSTEMTIME
    
    'SystemTimeToFileTime stime, ftime               ' SYSTEMTIME -> FILETIME
    'FileTimeToLocalFileTime ftime, ftime            ' учитываем смещение согласно текущим региональным настройкам в системе
    'FileTimeToSystemTime varptr(ftime), stime               ' FILETIME -> SYSTEMTIME
    
    'alternate:
    'GetTimeZoneInformation VarPtr(TimeZoneInfo(0))
    'SystemTimeToTzSpecificLocalTime VarPtr(TimeZoneInfo(0)), stime, stime
    
    SystemTimeToTzSpecificLocalTime 0&, sTime, sTime    'tz can be 0, if tz is current
    
    SystemTimeToVariantTime sTime, DateTime         ' SYSTEMTIME -> Date
    ConvertUnixTimeToLocalDate = DateTime
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modUtils.ConvertUnixTimeToLocalDate", "sec:", Seconds
    If inIDE Then Stop: Resume Next
End Function

Public Function ConvertFileTimeToLocalDate(lpFileTime As Long) As Date
    On Error GoTo ErrorHandler
    
    Dim DateTime            As Date
    Dim sTime               As SYSTEMTIME
    Dim ft                  As FILETIME
    
    memcpy ft, ByVal lpFileTime, Len(ft)
    
    FileTimeToSystemTime ft, sTime        ' FILETIME -> SYSTEMTIME
    SystemTimeToTzSpecificLocalTime 0&, sTime, sTime    ' tz can be 0, if tz is current
    SystemTimeToVariantTime sTime, DateTime             ' SYSTEMTIME -> vtDate
    
    ConvertFileTimeToLocalDate = DateTime
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modUtils.ConvertFileTimeToLocalDate", "lpFileTime:", lpFileTime
    If inIDE Then Stop: Resume Next
End Function

Private Function BuildPath$(sPath$, sFile$)
    If Right$(sPath, 1) = "\" Then
        BuildPath = sPath & sFile
    Else
        BuildPath = sPath & "\" & sFile
    End If
End Function

Public Function GetWindowsVersionTitle() As String
    Static isInit As Boolean
    Static sWinVer As String
    If isInit Then
        GetWindowsVersionTitle = sWinVer
        Exit Function
    Else
        isInit = True
        With OSver
            sWinVer = .OSName & " " & .Edition & " SP" & .SPVer & " " & _
                "(Windows " & .Platform & " " & .Major & "." & .Minor & "." & .Build & "." & .Revision & ")"
        End With
    End If
    GetWindowsVersionTitle = sWinVer
End Function

Public Sub PictureBoxRgn(pict As PictureBox, ByVal lMaskColor As Long)
    Const RGN_OR As Long = 2
    Dim hRgn As Long, hOutRgn As Long, i As Long, j As Long, lHeight As Long, lWidth As Long
    With pict
        hOutRgn = CreateRectRgn(0, 0, 0, 0)
        lWidth = .ScaleWidth: lHeight = .ScaleHeight
        For i = 0 To lWidth - 1
            For j = 0 To lHeight - 1
                If GetPixel(.hdc, i, j) <> lMaskColor Then
                    hRgn = CreateRectRgn(i, j, i + 1, j + 1)
                    Call CombineRgn(hOutRgn, hOutRgn, hRgn, RGN_OR)
                    Call DeleteObject(hRgn)
                End If
            Next j
        Next i
        Call SetWindowRgn(.hWnd, hOutRgn, True)
        Call DeleteObject(hOutRgn)
    End With
End Sub

Public Function GetDefaultApp(Protocol As String, Optional out_FriendlyName As String) As String
    On Error GoTo ErrorHandler
    
    Const ASSOCF_INIT_FIXED_PROGID  As Long = 2048
    Const ASSOCF_IS_PROTOCOL        As Long = 4096
    Const ASSOCF_INIT_FOR_FILE      As Long = 8192
    Const ASSOCF_INIT_BYEXENAME     As Long = 2
    Const ASSOCF_INIT_NOREMAPCLSID  As Long = 1
    Const ASSOCF_NOFIXUPS           As Long = &H100
    Const ASSOCF_INIT_IGNOREUNKNOWN As Long = &H400&
    Const ASSOCF_OPEN_BYEXENAME     As Long = 2&
    Const ASSOCF_INIT_DEFAULTTOSTAR As Long = 4&
    
    Const ASSOCSTR_EXECUTABLE       As Long = 2&
    Const ASSOCSTR_FRIENDLYAPPNAME  As Long = 4&
    Const ASSOCSTR_COMMAND          As Long = 1&
    
    Dim HRes    As Long
    Dim buf     As String
    Dim Size    As Long
    Dim buf2    As String
    
    buf = String$(MAX_PATH, vbNullChar)
    Size = MAX_PATH
    
    HRes = AssocQueryString(ASSOCF_INIT_DEFAULTTOSTAR Or ASSOCF_NOFIXUPS, _
        ASSOCSTR_COMMAND, StrPtr(Protocol), 0&, StrPtr(buf), Size)         'StrPtr("Open"),

    If HRes = 0 And Size <> 0 Then
        buf = Left$(buf, Size - 1)
        buf2 = String$(MAX_PATH_W, vbNullChar)
        Size = GetLongPathName(StrPtr(buf), StrPtr(buf2), Len(buf2))
        If Size <> 0 Then
            GetDefaultApp = Left$(buf2, Size)
        Else
            GetDefaultApp = buf
        End If
    Else
        If HRes = -2147023741 Then      'AL_USER
            GetDefaultApp = "(AppID)"
        Else
            GetDefaultApp = "?"
            'Err.Raise 76
        End If
    End If
    
    buf = String$(MAX_PATH, vbNullChar)
    Size = MAX_PATH
    
    HRes = AssocQueryString(ASSOCF_INIT_DEFAULTTOSTAR Or ASSOCF_NOFIXUPS, _
        ASSOCSTR_FRIENDLYAPPNAME, StrPtr(Protocol), 0&, StrPtr(buf), Size)
    
    If HRes = 0 And Size <> 0 Then
        buf = Left$(buf, Size - 1)
        out_FriendlyName = buf
    End If
    
    If 0 = Len(out_FriendlyName) And "(AppID)" = GetDefaultApp Or "?" = GetDefaultApp Then
        GetDefaultApp = ""
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetDefaultApp"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetSystemUpTime() As Date
    On Error GoTo ErrorHandler:
    
    Dim cMS As Currency
    Dim dDate As Date
    
    dDate = #12:00:00 AM#
    
    If IsProcedureAvail("GetTickCount64", "kernel32.dll") Then
        cMS = GetTickCount64()
        Dbg "GetTickCount64: " & cMS
        dDate = DateAdd("s", cMS * 10, dDate)
        Dbg "dDate: " & dDate
    Else
        dDate = DateAdd("s", GetTickCount() / 1000, dDate)
        Dbg "GetTickCount: " & GetTickCount()
        Dbg "dDate: " & dDate
    End If
    GetSystemUpTime = dDate
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetSystemUpTime"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetTimeZoneOffset(out_Offset As Long) As Boolean
    Dim tzi As TIME_ZONE_INFORMATION
    Dim dt As Date
    Dim lret As Long
    Dim dwShift As Long
    
    GetTimeZoneOffset = True
    
    'to measure shift, relative to Greenwich Mean Time: https://time.is/ru/GMT
    Select Case GetTimeZoneInformation(VarPtr(tzi))
    Case TIME_ZONE_ID_INVALID
        GetTimeZoneOffset = False
        Exit Function
    Case TIME_ZONE_ID_DAYLIGHT
        dwShift = tzi.Bias + tzi.DaylightBias
    Case TIME_ZONE_ID_STANDARD
        dwShift = tzi.Bias + tzi.StandardBias
    Case TIME_ZONE_ID_UNKNOWN
        dwShift = tzi.Bias
    End Select
    
'    to measure shift, relative to London time: https://www.timeanddate.com/worldclock/uk/london
'    without count the daylight shift (as it displayed in Windows clock timezone settings).
    
    out_Offset = dwShift
    
End Function

Public Function GetTimeZone(out_UTC As String) As Boolean
    Dim hh As Long
    Dim mm As Long
    Dim dwShift As Long
    
    GetTimeZone = GetTimeZoneOffset(dwShift)
    dwShift = dwShift * -1
    hh = dwShift \ 60
    mm = dwShift - (hh * 60)
    
    out_UTC = IIf(hh < 0 Or mm < 0, "-", "+") & Right$("0" & Abs(hh), 2) & ":" & Right$("0" & Abs(mm), 2)

End Function

Public Function ScanAfterReboot(Optional bSaveState As Boolean = True) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim dReboot     As Date
    Dim dNow        As Date
    Dim dZero       As Date
    Dim dUptime     As Date
    Dim dLastScan   As Date
    Dim sTime       As String
    
    dZero = #12:00:00 AM#
    dUptime = GetSystemUpTime()
    dNow = Now()
    dReboot = dNow - dUptime
    
    sTime = RegReadHJT("DateLastScan", vbNullString)
    
    If Len(sTime) <> 0 Then
        If StrBeginWith(sTime, "HJT:") Then
            sTime = mid$(sTime, 6)
            dLastScan = CDateEx(sTime, 1, 6, 9, 12, 15, 18)
        Else 'backward support
            If IsDate(sTime) Then
                dLastScan = CDate(sTime)
            End If
        End If
    End If
    
    If dLastScan = dZero Then
        ScanAfterReboot = True
    ElseIf dLastScan < dReboot Then
        ScanAfterReboot = True
    End If
    
    If bSaveState Then
        RegSaveHJT "DateLastScan", "HJT: " & Format$(dNow, "yyyy\/MM\/dd HH:nn:ss")
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ScanAfterReboot"
    If inIDE Then Stop: Resume Next
End Function

Public Function TrimSeconds(dDate As Date) As String
    If dDate > #11:59:59 PM# Then
        TrimSeconds = CStr(dDate)
        Dbg "TrimSeconds: " & TrimSeconds
    Else
        TrimSeconds = FormatDateTime(dDate, vbShortTime)
        Dbg "FormatDateTime: " & TrimSeconds
    End If
End Function

Public Sub SleepNoLock(ByVal nMSec As Long)
    Const nChunk As Long = 100&
    Dim nMSecCur As Long
    Do Until nMSecCur >= nMSec
        Sleep nChunk
        nMSecCur = nMSecCur + nChunk
        DoEvents
    Loop
End Sub

Public Sub GetTitleByCLSID(ByVal sCLSID As String, out_sTitle As String, Optional bRedirected As Boolean, Optional bShared As Boolean)
    On Error GoTo ErrorHandler:
    
    Dim sAppID As String
    Dim bRedirState As Boolean
    Dim out_sFile As String
    Dim sBuf As String
    Dim i As Long
    
    If Len(sCLSID) = 0 Then
        out_sTitle = STR_NO_NAME
        Exit Sub
    End If
    
    If Left$(sCLSID, 1) <> "{" Then
        sCLSID = "{" & sCLSID & "}"
    End If
    
    out_sTitle = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, bRedirected)
    If bShared And 0 = Len(out_sTitle) Then
        out_sTitle = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, Not bRedirected)
    End If
    If 0 = Len(out_sTitle) Then
        out_sTitle = STR_NO_NAME
        
        For i = 1 To 2
            If i = 1 Then
                bRedirState = bRedirected
            Else
                If Len(out_sFile) <> 0 Then Exit For
                If Not bShared Then Exit For
                bRedirState = Not bRedirected
            End If
            
            out_sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", vbNullString, bRedirState)
    
            If 0 = Len(out_sFile) Then
                sAppID = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, "AppID", bRedirState)
                If 0 <> Len(sAppID) Then
                    GetTitleByAppID sAppID, out_sTitle, bRedirState, False
                End If
            End If
        Next
    End If
        
    If Left$(out_sTitle, 1) = "@" Then
        sBuf = GetStringFromBinary(, , out_sTitle)
        If 0 <> Len(sBuf) Then out_sTitle = sBuf
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetTitleByCLSID", sCLSID, out_sTitle, bRedirected, bShared
    If inIDE Then Stop: Resume Next
End Sub

Public Sub GetFileByCLSID(ByVal sCLSID As String, out_sFile As String, Optional out_sTitle As Variant, Optional bRedirected As Boolean, Optional bShared As Boolean)
    On Error GoTo ErrorHandler:
    
    'Note: if 'bShared' = true, function will query for both WOW states,
    'but firstly it will query for redir. state defined in 'bRedirected' argument.
    
    '++ VersionIndependentProgID ?
    'e.g. http://www.checkfilename.com/view-details/Yahoo!-Widget-Engine/RespageIndex/0/sTab/2/
    
    Dim sBuf As String
    Dim sAppID As String
    Dim sServiceName As String
    Dim bRedirState As Boolean
    Dim i As Long
    
    If Len(sCLSID) = 0 Then
        out_sTitle = STR_NO_NAME
        out_sFile = STR_NO_FILE
        Exit Sub
    End If
    
    If Left$(sCLSID, 1) <> "{" Then
        sCLSID = "{" & sCLSID & "}"
    End If
    
    If Not IsMissing(out_sTitle) Then
        out_sTitle = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, bRedirected)
        If bShared And 0 = Len(out_sTitle) Then
            out_sTitle = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, Not bRedirected)
        End If
        If 0 = Len(out_sTitle) Then out_sTitle = STR_NO_NAME
        
        If Left$(out_sTitle, 1) = "@" Then
            sBuf = GetStringFromBinary(, , out_sTitle)
            If 0 <> Len(sBuf) Then out_sTitle = sBuf
        End If
    End If
    
    For i = 1 To 2
        If i = 1 Then
            bRedirState = bRedirected
        Else
            If Len(out_sFile) <> 0 Then Exit For
            If Not bShared Then Exit For
            bRedirState = Not bRedirected
        End If
    
        out_sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", vbNullString, bRedirState)
    
        If 0 = Len(out_sFile) Then
            sAppID = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, "AppID", bRedirState)
            If 0 <> Len(sAppID) Then
                If IsMissing(out_sTitle) Then
                    GetFileByAppID sAppID, out_sFile, , bRedirState, False
                Else
                    If out_sTitle <> STR_NO_NAME Then
                        GetFileByAppID sAppID, out_sFile, , bRedirState, False
                    Else
                        GetFileByAppID sAppID, out_sFile, out_sTitle, bRedirState, False
                    End If
                End If
            End If
            If Len(out_sFile) = 0 Then
                out_sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\LocalServer32", vbNullString, bRedirState)
            End If
        End If
    Next
    
    If 0 = Len(out_sFile) Then
        out_sFile = STR_NO_FILE
    Else
        out_sFile = UnQuote(EnvironW(out_sFile))
        
        If InStr(out_sFile, "\") = 0 Then
            out_sFile = FindOnPath(out_sFile, True)
        End If
        
        '8.3 -> Full
        If FileExists(out_sFile) Then
            out_sFile = GetLongPath(out_sFile)
            
    '    Else
    '        out_sFile = GetLongPath(out_sFile) & " " & STR_FILE_MISSING
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetFileByCLSID", sCLSID, out_sFile, out_sTitle, bRedirected, bShared
    If inIDE Then Stop: Resume Next
End Sub


Public Sub GetTitleByAppID(sAppID As String, out_sTitle As String, Optional bRedirected As Boolean, Optional bShared As Boolean)
    On Error GoTo ErrorHandler:
    
    Dim sBuf As String
    
    out_sTitle = Reg.GetString(HKEY_CLASSES_ROOT, "AppID\" & sAppID, vbNullString, bRedirected)
    If bShared And 0 = Len(out_sTitle) Then
        out_sTitle = Reg.GetString(HKEY_CLASSES_ROOT, "AppID\" & sAppID, vbNullString, Not bRedirected)
    End If
    If 0 = Len(out_sTitle) Then out_sTitle = STR_NO_NAME
    
    If Left$(out_sTitle, 1) = "@" Then
        sBuf = GetStringFromBinary(, , out_sTitle)
        If 0 <> Len(sBuf) Then out_sTitle = sBuf
    End If

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetTitleByAppID", sAppID, out_sTitle, bRedirected, bShared
    If inIDE Then Stop: Resume Next
End Sub

Public Sub GetFileByAppID(sAppID As String, out_sFile As String, Optional out_sTitle As Variant, Optional bRedirected As Boolean, Optional bShared As Boolean)
    On Error GoTo ErrorHandler:
    
    'https://learn.microsoft.com/en-us/windows/win32/com/appid-clsid
    'https://learn.microsoft.com/en-us/windows/win32/com/appid-key
    
    Dim sBuf As String
    Dim sServiceName As String
    Dim bRedirState As Boolean
    Dim i As Long
    
    If Not IsMissing(out_sTitle) Then
        out_sTitle = Reg.GetString(HKEY_CLASSES_ROOT, "AppID\" & sAppID, vbNullString, bRedirected)
        If bShared And 0 = Len(out_sTitle) Then
            out_sTitle = Reg.GetString(HKEY_CLASSES_ROOT, "AppID\" & sAppID, vbNullString, Not bRedirected)
        End If
        If 0 = Len(out_sTitle) Then out_sTitle = STR_NO_NAME
        
        If Left$(out_sTitle, 1) = "@" Then
            sBuf = GetStringFromBinary(, , out_sTitle)
            If 0 <> Len(sBuf) Then out_sTitle = sBuf
        End If
    End If
    
    For i = 1 To 2
        If i = 1 Then
            bRedirState = bRedirected
        Else
            If Len(out_sFile) <> 0 Then Exit For
            If Not bShared Then Exit For
            bRedirState = Not bRedirected
        End If
        
        sServiceName = Reg.GetString(HKEY_CLASSES_ROOT, "AppID\" & sAppID, "LocalService", bRedirState)
        If 0 <> Len(sServiceName) Then
            out_sFile = GetServiceDllPath(sServiceName)
            If 0 = Len(out_sFile) Then
                out_sFile = GetServiceImagePath(sServiceName)
            End If
        Else
            out_sFile = Reg.GetString(HKEY_CLASSES_ROOT, "AppID\" & sAppID, "DllSurrogate", bRedirState)
        End If
    Next

    If 0 <> Len(out_sFile) Then
        out_sFile = UnQuote(EnvironW(out_sFile))
        '8.3 -> Full
        If FileExists(out_sFile) Then
            out_sFile = GetLongPath(out_sFile)
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetFileByAppID", sAppID, out_sFile, out_sTitle, bRedirected, bShared
    If inIDE Then Stop: Resume Next
End Sub

'// Expand env. variable, unquote, normalize 8.3 path, search file on %PATH$ and append postfix "(no file)" or "(file missing)" if need.
'// Pass "sArgs" in case you want to check if argument missing for rundll32.exe process
Public Function FormatFileMissing(ByVal sFile As String, Optional sArgs As String) As String
    On Error GoTo ErrorHandler:
    
    Dim pos As Long
    
    sFile = UnQuote(EnvironW(sFile))
    
    If Len(sFile) = 0 Then
        FormatFileMissing = STR_NO_FILE
    ElseIf sFile = STR_NO_FILE Then
        FormatFileMissing = sFile
        Exit Function
    Else
        '8.3 -> Full
        sFile = GetLongPath(sFile)
        
        If Right$(sFile, 1) = "\" Then
            If FolderExists(sFile) Then
                FormatFileMissing = sFile
            Else
                FormatFileMissing = sFile & " (folder missing)"
            End If
            Exit Function
        End If
        
        If FileExists(sFile) Then
            FormatFileMissing = sFile
        Else
            If InStr(sFile, "\") <> 0 Then
                
                FormatFileMissing = sFile & " " & STR_FILE_MISSING
                
            Else 'relative path?
                Dim bFound As Boolean
                sFile = FindOnPath(sFile, True, , bFound)
                
                If bFound Then
                    FormatFileMissing = sFile
                Else
                    FormatFileMissing = sFile & " " & STR_FILE_MISSING
                End If
            End If
        End If
    End If
    
    'checking if argument file is missing, in case host is "rundll32"
    'e.g. C:\Windows\system32\Rundll32.exe C:\Windows\system32\iernonce.dll,RunOnceExProcess
    If Len(sArgs) <> 0 Then
        'If StrEndWith(sFile, "\rundll32.exe", 1) Then
        If StrComp(sFile, sWinSysDir & "\rundll32.exe", 1) = 0 Then
            Dim sDll As String
            sDll = GetRundllFile(sArgs)
            If Not FileExists(sDll) Then
                FormatFileMissing = sDll & " " & STR_FILE_MISSING
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FormatFileMissing", sFile, sArgs
    If inIDE Then Stop: Resume Next
End Function

'// concat string File + Arg, considering that "(no file)" or "(file missing)" postfixes in 'Filename' should go the last in the resulting string
Public Function ConcatFileArg(sFile As String, sArg As String) As String
    If Right$(sFile, 1) = ")" Then
        If StrEndWith(sFile, STR_FILE_MISSING) Then
            ConcatFileArg = Left$(sFile, Len(sFile) - Len(STR_FILE_MISSING)) & sArg & IIf(Len(sArg) = 0, vbNullString, " ") & STR_FILE_MISSING
            Exit Function
        ElseIf StrEndWith(sFile, STR_FOLDER_MISSING) Then
            ConcatFileArg = Left$(sFile, Len(sFile) - Len(STR_FOLDER_MISSING)) & sArg & IIf(Len(sArg) = 0, vbNullString, " ") & STR_FOLDER_MISSING
            Exit Function
        End If
    End If
    If Len(sArg) = 0 Then
        ConcatFileArg = sFile
    Else
        ConcatFileArg = sFile & " " & sArg
    End If
End Function

Function UnpackZIP(Archive As String, DestFolder As String) As Boolean
    On Error GoTo ErrorHandler
    Dim clsid   As UUID
    Dim iidSh   As UUID
    Dim shExt   As IShellExtInit
    Dim pf      As IPersistFolder2
    Dim pidl    As Long
    Dim cb      As Long
    
    CLSIDFromString StrPtr(ZipFldrCLSID), clsid
    'CLSIDFromString StrPtr(IID_IShellExtInit), iidSh
    
    iidSh = IID_IShellExtInit
    
    If CoCreateInstance(clsid, 0&, CLSCTX_INPROC_SERVER, iidSh, shExt) <> S_OK Then Exit Function
    
    Set pf = shExt
    SHParseDisplayName StrPtr(Archive), 0&, pidl, 0&, 0&
    pf.Initialize pidl
    CoTaskMemFree pidl
 
    Dim srg     As IStorage
    Dim stm     As IStream
    Dim enm     As IEnumSTATSTG
    Dim itm     As STATSTG
    Dim nam     As String
    Dim buf()   As Byte
    Dim hFile   As Long
    Dim bWrote  As Boolean
    
    Set srg = pf
    Set enm = srg.EnumElements
    
    enm.Reset
    bWrote = True
    
    Do While enm.Next(1&, itm) = S_OK
        cb = lstrlen(itm.pwcsName)
        nam = Space$(cb)
        
        lstrcpyn StrPtr(nam), itm.pwcsName, cb + 1
        CoTaskMemFree itm.pwcsName
        
        If itm.Type <> STGTY_STORAGE Then
            
            OpenW BuildPath(DestFolder, nam), FOR_OVERWRITE_CREATE, hFile
            
            Set stm = srg.OpenStream(nam, 0&, STGM_READ, 0&)
            
            ReDim buf(&HFFFF&)
            
            Do
                cb = stm.Read(buf(0), UBound(buf) + 1)
                If cb = 0 Then Exit Do
                If cb <= UBound(buf) Then ReDim Preserve buf(cb - 1)
                bWrote = bWrote And PutW(hFile, 1, VarPtr(buf(0)), cb, True)
            Loop
            CloseW hFile
        End If
    Loop
    UnpackZIP = FileExists(BuildPath(DestFolder, nam)) And bWrote
    Exit Function
ErrorHandler:
    ErrorMsg Err, "UnpackZIP", Archive, DestFolder
    If inIDE Then Stop: Resume Next
End Function

Function UnpackZIPtoArray(Archive As String, out_Buf() As Byte) As Boolean
    On Error GoTo ErrorHandler
    
    Const Chunk As Long = &HFFFF&
    
    Dim clsid   As UUID
    Dim iidSh   As UUID
    Dim shExt   As IShellExtInit
    Dim pf      As IPersistFolder2
    Dim pidl    As Long
    Dim cb      As Long
    
    CLSIDFromString StrPtr(ZipFldrCLSID), clsid
    'CLSIDFromString StrPtr(IID_IShellExtInit), iidSh
    
    iidSh = IID_IShellExtInit
    
    If CoCreateInstance(clsid, 0&, CLSCTX_INPROC_SERVER, iidSh, shExt) <> S_OK Then Exit Function
    
    Set pf = shExt
    SHParseDisplayName StrPtr(Archive), 0&, pidl, 0&, 0&
    pf.Initialize pidl
    CoTaskMemFree pidl
 
    Dim srg     As IStorage
    Dim stm     As IStream
    Dim enm     As IEnumSTATSTG
    Dim itm     As STATSTG
    Dim nam     As String
    Dim pos     As Long
    
    Set srg = pf
    Set enm = srg.EnumElements
    
    enm.Reset
    Do While enm.Next(1&, itm) = S_OK
        cb = lstrlen(itm.pwcsName)
        nam = Space$(cb)
        
        lstrcpyn StrPtr(nam), itm.pwcsName, cb + 1
        CoTaskMemFree itm.pwcsName
        
        If itm.Type <> STGTY_STORAGE Then
            
            Set stm = srg.OpenStream(nam, 0&, STGM_READ, 0&)
            pos = -1
            Do
                ReDim Preserve out_Buf(pos + Chunk)
                cb = stm.Read(out_Buf(pos + 1), Chunk)
                If cb < Chunk Then
                    If pos + cb <> -1 Then
                        ReDim Preserve out_Buf(pos + cb)
                        UnpackZIPtoArray = True
                    Else
                        Erase out_Buf
                    End If
                    Exit Do
                End If
                pos = pos + Chunk
            Loop
        End If
    Loop
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "UnpackZIP", Archive
    If inIDE Then Stop: Resume Next
End Function

Public Sub CreateUninstallKey(bCreate As Boolean, Optional EXE_Location As String = vbNullString) ' if false -> delete registry entries
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CreateUninstallKey - Begin"
    Dim Setup_Key$:   Setup_Key = "Software\Microsoft\Windows\CurrentVersion\Uninstall\HiJackThis Fork"
    
    If bCreate Then
        If Len(EXE_Location) = 0 Then EXE_Location = AppPath(True)
        
        Reg.CreateKey HKEY_LOCAL_MACHINE, Setup_Key
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "DisplayName", g_AppName & " " & AppVerString
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "UninstallString", """" & EXE_Location & """ /uninstall"
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "QuietUninstallString", """" & EXE_Location & """ /silentuninstall"
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "DisplayIcon", EXE_Location & ",0"
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "DisplayVersion", AppVerString
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "Publisher", "Alex Dragokas"
        'Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "URLInfoAbout", "http://www.spywareinfo.com/~merijn/"
        'Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "URLInfoAbout", "https://sourceforge.net/projects/hjt/"
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "URLInfoAbout", "https://github.com/dragokas/hijackthis"
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "HelpLink", "https://github.com/dragokas/hijackthis/wiki/HJT:-Tutorial"
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "InstallLocation", GetParentDir(EXE_Location)
        Reg.SetDwordVal HKEY_LOCAL_MACHINE, Setup_Key, "NoModify", 1
        Reg.SetDwordVal HKEY_LOCAL_MACHINE, Setup_Key, "NoRepair", 1
        Reg.SetDwordVal HKEY_LOCAL_MACHINE, Setup_Key, "EstimatedSize", FileLenW(EXE_Location) \ 1024 'KB
        Reg.SetDwordVal HKEY_LOCAL_MACHINE, Setup_Key, "Language", IIf(g_CurrentLangEnum = Lang_Russian Or g_CurrentLangEnum = Lang_Ukrainian, &H419&, &H409&)
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "InstallDate", Format$(Date, "yyyymmdd", vbMonday)
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "Contact", "admin@dragokas.com"
        Reg.SetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "Comments", "Creates a report of non-standard parameters of registry " & _
            "and file system for selectively removal of items related to the activities of malware and security risks"
    Else
        Reg.DelKey HKEY_LOCAL_MACHINE, Setup_Key 'HiJackThis Fork key
        Reg.DelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HiJackThis"
        Reg.DelKey HKEY_CURRENT_USER, "Software\Microsoft\Installer\Products\8A9C1670A3F861244B7A7BFAFB422AA4"
    End If
    AppendErrorLogCustom "CreateUninstallKey - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CreateUninstallKey", bCreate
    If inIDE Then Stop: Resume Next
End Sub

Public Function RegSaveHJT(sName$, sData$, Optional idSection As SETTINGS_SECTION) As Boolean
    On Error GoTo ErrorHandler:
    
    If Not OSver.IsElevated Then Exit Function
    
    Dim sSubSection As String
    sSubSection = SectionNameById(idSection)
    
    If Len(sSubSection) <> 0 Then sSubSection = "\" & sSubSection
    
    If sName Like "Ignore#*" Or sName = "ProxyPass" Then
        Dim aData() As Byte
        aData = sData
        Reg.SetBinaryVal HKEY_LOCAL_MACHINE, g_SettingsRegKey & sSubSection, sName, aData
    Else
        Reg.SetStringVal HKEY_LOCAL_MACHINE, g_SettingsRegKey & sSubSection, sName, sData
    End If
    
    RegSaveHJT = True
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegSaveHJT", sName & "," & sData & " - Section: " & sSubSection
    If inIDE Then Stop: Resume Next
End Function

Public Function RegReadHJT( _
    sName$, _
    Optional sDefault$, _
    Optional idSection As SETTINGS_SECTION) As String
    
    On Error GoTo ErrorHandler:
    
    Dim sSubSection As String
    sSubSection = SectionNameById(idSection)
    
    If Len(sSubSection) <> 0 Then sSubSection = "\" & sSubSection
    
    Dim sKeyHJT As String
    sKeyHJT = g_SettingsRegKey & sSubSection
    
    If sName Like "Ignore#*" Or sName = "ProxyPass" Then
        Dim aData() As Byte
        aData = Reg.GetBinary(HKEY_LOCAL_MACHINE, sKeyHJT, sName)
        If AryItems(aData) Then
            RegReadHJT = aData
        End If
    Else
        RegReadHJT = Reg.GetString(HKEY_LOCAL_MACHINE, sKeyHJT, sName)
    End If
    If Len(RegReadHJT) = 0 Then RegReadHJT = sDefault
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegReadHJT", sName & "," & sDefault & " - Section: " & sSubSection
    If inIDE Then Stop: Resume Next
End Function

Public Function RegDelHJT(sName$, Optional idSection As SETTINGS_SECTION) As Boolean

    If Not OSver.IsElevated Then Exit Function

    Dim sSubSection As String
    sSubSection = SectionNameById(idSection)
    
    If Len(sSubSection) <> 0 Then sSubSection = "\" & sSubSection
    
    Reg.DelVal HKEY_LOCAL_MACHINE, g_SettingsRegKey & sSubSection, sName
    
    RegDelHJT = True
    
End Function

Public Function GetFontHeight(Fnt As Font) As Long
    Dim FntPrev As Font
    Set FntPrev = frmMain.Font
    Set frmMain.Font = Fnt
    GetFontHeight = frmMain.TextHeight("A")
    Set frmMain.Font = FntPrev
End Function

Public Function isCLSID(sStr As String) As Boolean
    If Left$(sStr, 1) = "{" Then
        If sStr Like "{????????-????-????-????-????????????}*" Then isCLSID = True
    Else
        If sStr Like "????????-????-????-????-????????????*" Then isCLSID = True
    End If
End Function

Public Sub AlignCommandButtonText(Button As VBCCR17.CommandButtonW, Style As BUTTON_ALIGNMENT)
    Dim lOldStyle As Long
    Dim lret As Long
    lOldStyle = GetWindowLong(Button.hWnd, GWL_STYLE)
    lret = SetWindowLong(Button.hWnd, GWL_STYLE, Style Or lOldStyle)
    Button.Refresh
End Sub

' ≈сть ли в строке адрес URL
Public Function isURL(ByVal sText As String) As Boolean
    Static sLastURL As String
    Static bLastResult As Boolean
    
    If sText = sLastURL Then
        isURL = bLastResult
        Exit Function
    End If
    
    sLastURL = sText
    bLastResult = False

    If mid$(sText, 2, 1) = ":" Then
        If FileExists(sText) Then Exit Function
        If FolderExists(sText) Then Exit Function
    End If
    If mid$(sText, 3, 1) = ":" Then
        If Left$(sText, 1) = """" Then
            Dim sFile As String
            sFile = UnQuote(sText)
            If FileExists(sFile) Then Exit Function
            If FolderExists(sFile) Then Exit Function
        End If
    End If
    sText = UCase$(sText)
    If InStr(1&, sText, "HTTP:", vbBinaryCompare) <> 0& Then
        isURL = True
    ElseIf InStr(1&, sText, "WWW.", vbBinaryCompare) <> 0& Then
        isURL = True
    ElseIf InStr(1&, sText, "HTTPS:", vbBinaryCompare) <> 0& Then
        isURL = True
    ElseIf InStr(1&, sText, "WWW2.", vbBinaryCompare) <> 0& Then
        isURL = True
    End If
    bLastResult = isURL
End Function

Public Function GetHandleType(Handle As Long) As String
    Dim status As Long
    Dim returnedLength As Long
    Dim oti As OBJECT_TYPE_INFORMATION
    
    '//TODO: check it
    
    status = NtQueryObject(Handle, ObjectTypeInformation, oti, LenB(oti), returnedLength)
    
    'STATUS_INFO_LENGTH_MISMATCH (&HC0000004)
    
    If NT_SUCCESS(status) And returnedLength > 0 Then
        If oti.TypeName.Length > 0 Then
            GetHandleType = StringFromPtrW(oti.TypeName.Buffer)
        End If
    End If
End Function

' affected by wow64
Public Function OpenFileDialog(Optional sTitle As String, Optional InitDir As String, Optional sFilter As String, Optional hWnd As Long) As String
    On Error GoTo ErrorHandler
    Const OFN_DONTADDTORECENT As Long = &H2000000
    Const OFN_ENABLESIZING As Long = &H800000
    Const OFN_FORCESHOWHIDDEN As Long = &H10000000
    Const OFN_HIDEREADONLY As Long = 4&
    Const OFN_NODEREFERENCELINKS As Long = &H100000
    Const OFN_NOVALIDATE As Long = &H100&
    
    If Len(sFilter) = 0 Then
        'sFilter = "All Files (*.*)|*.*"
        sFilter = Translate(1003) & " (*.*)|*.*"
    End If
    
    If OSver.IsWindowsVistaOrGreater Then
        OpenFileDialog = OpenFileDialogVista_Simple(sTitle, InitDir, sFilter, hWnd)
        Exit Function
    End If
    
    If InStr(sFilter, "|") > 0 Then sFilter = Replace$(sFilter, "|", vbNullChar)
    If Right$(sFilter, 2) <> vbNullChar & vbNullChar Then sFilter = sFilter & vbNullChar & vbNullChar
    
    Dim OFN As OPENFILENAME
    Dim out As String
    
    OFN.nMaxFile = MAX_PATH_W
    out = String$(MAX_PATH_W, vbNullChar)
    
    With OFN
        .hWndOwner = IIf(hWnd = 0, g_HwndMain, hWnd)
        .lpstrTitle = StrPtr(sTitle)
        .lpstrFile = StrPtr(out)
        .lStructSize = Len(OFN)
        .lpstrInitialDir = StrPtr(InitDir)
        .lpstrFilter = StrPtr(sFilter)
        .Flags = OFN_DONTADDTORECENT Or OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN Or OFN_HIDEREADONLY Or OFN_NOVALIDATE
    End With
    If GetOpenFileName(OFN) Then OpenFileDialog = TrimNull(out)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "OpenFileDialog"
    If inIDE Then Stop: Resume Next
End Function

Public Function OpenFileDialog_Multi( _
    aPath() As String, _
    Optional sTitle As String, _
    Optional InitDir As String, _
    Optional sFilter As String, _
    Optional hWnd As Long) As Long
    
    On Error GoTo ErrorHandler
    Const OFN_DONTADDTORECENT As Long = &H2000000
    Const OFN_ENABLESIZING As Long = &H800000
    Const OFN_FORCESHOWHIDDEN As Long = &H10000000
    Const OFN_HIDEREADONLY As Long = 4&
    Const OFN_NODEREFERENCELINKS As Long = &H100000
    Const OFN_NOVALIDATE As Long = &H100&
    Const OFN_ALLOWMULTISELECT As Long = &H200&
    Const OFN_EXPLORER As Long = &H80000
    
    If Len(sFilter) = 0 Then
        'sFilter = "All Files (*.*)|*.*"
        sFilter = Translate(1003) & " (*.*)|*.*"
    End If
    
    If OSver.IsWindowsVistaOrGreater Then
        OpenFileDialog_Multi = OpenFileDialogVista_Multi(aPath, sTitle, InitDir, sFilter, hWnd)
        Exit Function
    End If
    
    If InStr(sFilter, "|") > 0 Then sFilter = Replace$(sFilter, "|", vbNullChar)
    If Right$(sFilter, 2) <> vbNullChar & vbNullChar Then sFilter = sFilter & vbNullChar & vbNullChar

    Dim OFN As OPENFILENAME
    Dim out As String
    Dim aFiles() As String
    Dim i As Long
    
    OFN.nMaxFile = MAX_PATH_W
    out = String$(MAX_PATH_W, vbNullChar)
    
    With OFN
        .hWndOwner = IIf(hWnd = 0, g_HwndMain, hWnd)
        .lpstrTitle = StrPtr(sTitle)
        .lpstrFile = StrPtr(out)
        .lStructSize = Len(OFN)
        .lpstrInitialDir = StrPtr(InitDir)
        .lpstrFilter = StrPtr(sFilter)
        .Flags = OFN_DONTADDTORECENT Or OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN Or OFN_HIDEREADONLY Or OFN_NOVALIDATE Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER
    End With
    If GetOpenFileName(OFN) Then
        aFiles = Split(RTrimNull(out), vbNullChar)
        If UBound(aFiles) = 0 Then
            ReDim aPath(1)
            aPath(1) = aFiles(0)
            OpenFileDialog_Multi = 1
        Else
            For i = 1 To UBound(aFiles)
                aFiles(i) = BuildPath(aFiles(0), aFiles(i))
            Next
            ReDim aPath(UBound(aFiles))
            For i = 1 To UBound(aFiles)
                aPath(i) = aFiles(i)
            Next
            OpenFileDialog_Multi = UBound(aPath)
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "OpenFileDialog"
    If inIDE Then Stop: Resume Next
End Function

Public Function SaveFileDialog( _
    Optional sTitle As String, _
    Optional InitDir As String, _
    Optional sDefFile As String, _
    Optional sFilter As String, _
    Optional hWnd As Long) As String
    
    Dim uOFN As OPENFILENAME, sFile$, sExt$
    On Error GoTo ErrorHandler:
    
    Const OFN_ENABLESIZING As Long = &H800000
    
    If Len(sFilter) = 0 Then
        'sFilter = "All Files (*.*)|*.*"
        sFilter = Translate(1003) & " (*.*)|*.*"
    End If
    
    If OSver.IsWindowsVistaOrGreater Then
        SaveFileDialog = SaveFileDialogVista(sTitle, InitDir, sDefFile, sFilter, hWnd)
        Exit Function
    End If
    
    If InStr(sFilter, "|") > 0 Then sFilter = Replace$(sFilter, "|", vbNullChar)
    If Right$(sFilter, 2) <> vbNullChar & vbNullChar Then sFilter = sFilter & vbNullChar & vbNullChar
    
    sFile = String$(MAX_PATH, 0)
    LSet sFile = sDefFile
    With uOFN
        .lStructSize = Len(uOFN)
        .hWndOwner = IIf(hWnd = 0, g_HwndMain, hWnd)
        .lpstrFilter = StrPtr(sFilter)
        .lpstrFile = StrPtr(sFile)
        .lpstrTitle = StrPtr(sTitle)
        .nMaxFile = Len(sFile)
        .lpstrInitialDir = StrPtr(InitDir)
        .lpstrDefExt = StrPtr(mid$(GetExtensionName(sDefFile), 2))
        .Flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_OVERWRITEPROMPT Or OFN_ENABLESIZING
    End With
    If GetSaveFileName(uOFN) = 0 Then Exit Function
    sFile = TrimNull(sFile)
    sExt = GetExtensionName(sDefFile)
    ' check if extension present
    If Len(sFile) <> 0 And Len(sExt) <> 0 Then
        If Not StrEndWith(sFile, sExt) Then
            sFile = sFile & sExt
        End If
    End If
    SaveFileDialog = sFile
    Exit Function
    
ErrorHandler:
    ErrorMsg Err, "SaveFileDialog", "sTitle=", sTitle, "sFilter=", sFilter, "sDefFile=", sDefFile
    If inIDE Then Stop: Resume Next
End Function

'Modern open/save file dialogue (c) 2015-2016 by fafalone
'
'Fork by Dragokas
' - Button_Click converted to function (only 3 variants: simple open dialogue, multiselect and save dialogue)
' - Default folder open location added as a parameter to function, otherwise "My PC" location will be open.
' - Added options FOS_FILEMUSTEXIST (FOS_PATHMUSTEXIST), FOS_FORCESHOWHIDDEN (to ensure file presence, show hidden files)
' - Added Windows\System32 to SysNative redirection by fafalone's sample (to support x64)
' - Removed example of VTable events redirection from class to module
' - Removed some comments
' - Removed several filters: left just 1 : *.* (all files)
' - Added .Unadvise to unhook events handler
' - Fixed infinite loop when cancel meltiselect dialogue

'Note: this dialogue doesn't show objects that have both "System" and "Hidden" flags set.
'This dialogue is unicode aware.

Private Function OpenFileDialogVista_Multi( _
    aPath() As String, _
    Optional sTitle As String, _
    Optional InitDir As String, _
    Optional sFilter As String, _
    Optional hWnd As Long) As Long
    
    On Error GoTo ErrorHandler:
    
    Dim isiRes      As IShellItem
    Dim isia        As IShellItemArray
    Dim iesi        As IEnumShellItems
    Dim fodSimple   As FileOpenDialog
    Dim isiDef      As IShellItem
    Dim pidlDef     As Long
    Dim nItems      As Long
    Dim nItem       As Long
    Dim lPtr        As Long
    Dim lOptions    As FILEOPENDIALOGOPTIONS
    Dim hCookie     As Long
    Dim cEvents     As clsFileDialogEvents
    Dim arrsplit()  As String
    Dim i&, idx&
    
    Set cEvents = New clsFileDialogEvents
    
    'Set up filter
    Dim FileFilter() As COMDLG_FILTERSPEC

    If InStr(sFilter, "|") = 0 Then
        'sFilter = "All files (*.*)|*.*"
        sFilter = Translate(1003) & " (*.*)|*.*"
    Else
        arrsplit = Split(sFilter, "|")
        For i = 0 To UBound(arrsplit) Step 2
            ReDim Preserve FileFilter(idx)
            FileFilter(idx).pszName = arrsplit(i)
            FileFilter(idx).pszSpec = arrsplit(i + 1)
            idx = idx + 1
        Next
    End If
    
    If Len(InitDir) <> 0 Then
        pidlDef = GetPIDLFromPathW(InitDir)
    End If
    If pidlDef = 0 Then
        Call SHGetKnownFolderIDList(FOLDERID_ComputerFolder, 0, 0, pidlDef)
    End If
    If pidlDef Then
        Call SHCreateShellItem(0, 0, pidlDef, isiDef)
    End If
    
    Set fodSimple = New FileOpenDialog
    
    With fodSimple
        .Advise cEvents, hCookie
        .SetTitle sTitle
        
        'When setting options, you should first get them
        .GetOptions lOptions
        lOptions = lOptions Or FOS_ALLOWMULTISELECT Or FOS_FILEMUSTEXIST Or FOS_FORCESHOWHIDDEN Or FOS_NOVALIDATE Or FOS_DONTADDTORECENT
        .SetOptions lOptions
        If Not (isiDef Is Nothing) Then .SetFolder isiDef
        .SetFileTypes UBound(FileFilter) + 1, VarPtr(FileFilter(0).pszName) ' number of items in filter
        
        On Error Resume Next
        .Show IIf(hWnd = 0, g_HwndMain, hWnd)
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler:
            
            .GetResults isia
            isia.EnumItems iesi
            isia.GetCount nItems
            ReDim aPath(nItems)
            
            Do While (iesi.Next(1, isiRes, 0) = 0)
                nItem = nItem + 1
                isiRes.GetDisplayName SIGDN_FILESYSPATH, lPtr
                aPath(nItem) = BStrFromLPWStr(lPtr, True)
                If PathBeginWith(aPath(nItem), sSysNativeDir) Then aPath(nItem) = sWinSysDir & mid$(aPath(nItem), Len(sSysNativeDir) + 1)
                Set isiRes = Nothing
            Loop
        End If
        On Error GoTo ErrorHandler:
        
        .Unadvise hCookie
    End With
    
    Set cEvents = Nothing
    If pidlDef Then Call CoTaskMemFree(pidlDef)
    Set isiRes = Nothing
    Set isiDef = Nothing
    Set fodSimple = Nothing
    Set iesi = Nothing
    Set isia = Nothing
    Set fodSimple = Nothing
    
    OpenFileDialogVista_Multi = nItems
    Exit Function
ErrorHandler:
    ErrorMsg Err, "OpenFileDialogVista_Multi", "InitDir: ", InitDir
    If inIDE Then Stop: Resume Next
End Function

Private Function OpenFileDialogVista_Simple( _
    Optional sTitle As String, _
    Optional InitDir As String, _
    Optional sFilter As String, _
    Optional hWnd As Long) As String
    
    On Error GoTo ErrorHandler:
    
    Dim isiRes          As IShellItem
    Dim lPtr            As Long
    Dim FileFilter()    As COMDLG_FILTERSPEC
    Dim fodSimple       As FileOpenDialog
    Dim isiDef          As IShellItem
    Dim pidlDef         As Long
    Dim lOptions        As FILEOPENDIALOGOPTIONS
    Dim hCookie         As Long
    Dim cEvents         As clsFileDialogEvents
    Dim sPath           As String
    Dim arrsplit()      As String
    Dim i&, idx&
    
    Set cEvents = New clsFileDialogEvents
    
    If Len(InitDir) <> 0 Then
        pidlDef = GetPIDLFromPathW(InitDir)
    End If
    If pidlDef = 0 Then
        Call SHGetKnownFolderIDList(FOLDERID_ComputerFolder, 0, 0, pidlDef)
    End If
    If pidlDef Then
        Call SHCreateShellItem(0, 0, pidlDef, isiDef)
    End If
    
    If InStr(sFilter, "|") = 0 Then
        'sFilter = "All files (*.*)|*.*"
        sFilter = Translate(1003) & " (*.*)|*.*"
    Else
        arrsplit = Split(sFilter, "|")
        For i = 0 To UBound(arrsplit) Step 2
            ReDim Preserve FileFilter(idx)
            FileFilter(idx).pszName = arrsplit(i)
            FileFilter(idx).pszSpec = arrsplit(i + 1)
            idx = idx + 1
        Next
    End If
    
    Set fodSimple = New FileOpenDialog
    
    With fodSimple
        .Advise cEvents, hCookie
        .SetTitle sTitle
        .SetFileTypes UBound(FileFilter) + 1, VarPtr(FileFilter(0).pszName) '1 - number of entries in the filter
        .GetOptions lOptions
        lOptions = lOptions Or FOS_FILEMUSTEXIST Or FOS_FORCESHOWHIDDEN Or FOS_NOVALIDATE Or FOS_DONTADDTORECENT
        .SetOptions lOptions
        If Not (isiDef Is Nothing) Then .SetFolder isiDef
        
        On Error Resume Next
        .Show IIf(hWnd = 0, g_HwndMain, hWnd)
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler:
            .GetResult isiRes
            isiRes.GetDisplayName SIGDN_FILESYSPATH, lPtr
            sPath = BStrFromLPWStr(lPtr, True)
            If PathBeginWith(sPath, sSysNativeDir) Then sPath = sWinSysDir & mid$(sPath, Len(sSysNativeDir) + 1)
            OpenFileDialogVista_Simple = sPath
        End If
        On Error GoTo ErrorHandler:
        .Unadvise hCookie
    End With
    
    Set cEvents = Nothing
    If pidlDef Then Call CoTaskMemFree(pidlDef)
    Set isiRes = Nothing
    Set fodSimple = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "OpenFileDialogVista_Simple", "InitDir: ", InitDir
    If inIDE Then Stop: Resume Next
End Function

Private Function SaveFileDialogVista( _
    Optional sTitle As String, _
    Optional InitDir As String, _
    Optional DefSaveFile As String, _
    Optional sFilter As String, _
    Optional hWnd As Long) As String
    
    On Error GoTo ErrorHandler:
    
    Dim isiRes          As IShellItem
    Dim lPtr            As Long
    Dim FileFilter()    As COMDLG_FILTERSPEC
    Dim fsd             As FileSaveDialog
    Dim isiDef          As IShellItem
    Dim pidlDef         As Long
    Dim lOptions        As FILEOPENDIALOGOPTIONS
    Dim hCookie         As Long
    Dim cEvents         As clsFileDialogEvents
    Dim sPath           As String
    Dim pos             As Long
    Dim i&, idx&, arrsplit() As String
    
    Set cEvents = New clsFileDialogEvents
    
    If Len(InitDir) <> 0 Then
        pidlDef = GetPIDLFromPathW(InitDir)
    End If
    If pidlDef = 0 Then
        Call SHGetKnownFolderIDList(FOLDERID_ComputerFolder, 0, 0, pidlDef)
    End If
    If pidlDef Then
        Call SHCreateShellItem(0, 0, pidlDef, isiDef)
    End If
    
    If InStr(sFilter, "|") = 0 Then
        'sFilter = "All files (*.*)|*.*"
        sFilter = Translate(1003) & " (*.*)|*.*"
    Else
        arrsplit = Split(sFilter, "|")
        For i = 0 To UBound(arrsplit) Step 2
            ReDim Preserve FileFilter(idx)
            FileFilter(idx).pszName = arrsplit(i)
            FileFilter(idx).pszSpec = arrsplit(i + 1)
            idx = idx + 1
        Next
    End If
    
    Set fsd = New FileSaveDialog
    
    With fsd
        .Advise cEvents, hCookie
        .SetTitle sTitle
        .SetFileTypes UBound(FileFilter) + 1, VarPtr(FileFilter(0).pszName) '1 - number of entries in the filter
        .GetOptions lOptions
        lOptions = lOptions Or FOS_PATHMUSTEXIST Or FOS_FORCESHOWHIDDEN Or FOS_OVERWRITEPROMPT
        .SetOptions lOptions
        If Not (isiDef Is Nothing) Then .SetFolder isiDef
        
        pos = InStr(DefSaveFile, ".")
        If pos <> 0 Then .SetDefaultExtension mid$(DefSaveFile, pos + 1)
        .SetFileName DefSaveFile
        
        On Error Resume Next
        .Show IIf(hWnd = 0, g_HwndMain, hWnd)
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler:
            .GetResult isiRes
            isiRes.GetDisplayName SIGDN_FILESYSPATH, lPtr
            sPath = BStrFromLPWStr(lPtr, True)
            If PathBeginWith(sPath, sSysNativeDir) Then sPath = sWinSysDir & mid$(sPath, Len(sSysNativeDir) + 1)
            SaveFileDialogVista = sPath
        End If
        On Error GoTo ErrorHandler:
        .Unadvise hCookie
    End With
    
    Set cEvents = Nothing
    If pidlDef Then Call CoTaskMemFree(pidlDef)
    Set isiRes = Nothing
    Set fsd = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SaveFileDialogVista", "InitDir: ", InitDir, "DefSaveFile: ", DefSaveFile
    If inIDE Then Stop: Resume Next
End Function

Public Function PathBeginWith(sPath As String, sBeginPart As String) As Boolean
    If StrComp(Left$(sPath, Len(sBeginPart)), sBeginPart, 1) = 0 Then
        If Len(sPath) = Len(sBeginPart) Or mid$(sPath, Len(sBeginPart) + 1, 1) = "\" Then PathBeginWith = True
    End If
End Function

Private Function GetPIDLFromPathW(sPath As String) As Long
   GetPIDLFromPathW = ILCreateFromPathW(StrPtr(sPath))
End Function

Public Function BStrFromLPWStr(lpWStr As Long, Optional ByVal CleanupLPWStr As Boolean = True) As String
    If lpWStr = 0 Then Exit Function
    SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
    If CleanupLPWStr Then CoTaskMemFree lpWStr
End Function

Public Function ProcessHotkey(KeyCode As Integer, frm As Form) As Boolean
    If KeyCode = Asc("F") Then                    'Ctrl + F
        If Not (cMath Is Nothing) Then
            If cMath.HIWORD(GetKeyState(VK_CONTROL)) Then LoadSearchEngine frm: ProcessHotkey = True
        End If
    ElseIf KeyCode = Asc("A") Then                    'Ctrl + A
        If Not (cMath Is Nothing) Then
            If cMath.HIWORD(GetKeyState(VK_CONTROL)) Then ControlSelectAll frm: ProcessHotkey = True
        End If
    End If
End Function

Public Function ProcessDelayedHotkey(hotkey As clsHotkey, frm As Form) As Boolean
    If hotkey.IsControlHotkey(VK_F) Then
        LoadSearchEngine frm
        ProcessDelayedHotkey = True
    ElseIf hotkey.IsControlHotkey(VK_A) Then
        ControlSelectAll frm
        ProcessDelayedHotkey = True
    End If
End Function

Public Sub LoadSearchEngine(frm As Form)
    If IsFormInit(frmSearch) Then
        frmSearch.Display frm
    Else
        Load frmSearch
        frmSearch.Display frm
    End If
End Sub

Sub ControlSelectAll(Optional frmExplicit As Form)
    Dim out_Control As Control
    Dim lst As VBCCR17.ListBoxW
    Dim txb As VBCCR17.TextBoxW
    Dim bCanSearch As Boolean
    Dim i As Long

    Select Case frmExplicit.Name
    Case "frmMain"
    
        Select Case g_CurFrame
        
'        Case FRAME_ALIAS_SCAN
'            bCanSearch = True
'            Set out_Control = frmMain.lstResults
            
        Case FRAME_ALIAS_IGNORE_LIST
            bCanSearch = True
            Set out_Control = frmMain.lstIgnore
        
        Case FRAME_ALIAS_BACKUPS
            bCanSearch = True
            Set out_Control = frmMain.lstBackups
            
        Case FRAME_ALIAS_HELP_SECTIONS, FRAME_ALIAS_HELP_KEYS, FRAME_ALIAS_HELP_PURPOSE, FRAME_ALIAS_HELP_HISTORY
            bCanSearch = True
            Set out_Control = frmMain.txtHelp
        
        End Select
        
    Case "frmADSspy"
    
        bCanSearch = True
        If frmADSspy.txtADSContent.Visible Then
            Set out_Control = frmADSspy.txtADSContent
        Else
            Set out_Control = frmADSspy.lstADSFound
        End If
    
'    Case "frmUninstMan"
'        bCanSearch = True
'        Set out_Control = frmUninstMan.lstUninstMan
    
'    Case "frmProcMan"
'        bCanSearch = True
'        If frmProcMan.ProcManDLLsHasFocus Then
'            Set out_Control = frmProcMan.lstProcManDLLs
'        Else
'            Set out_Control = frmProcMan.lstProcessManager
'        End If
        
    Case "frmCheckDigiSign"
        bCanSearch = True
        Set out_Control = frmCheckDigiSign.txtPaths
    
    Case "frmUnlockRegKey"
        bCanSearch = True
        Set out_Control = frmUnlockRegKey.txtKeys
        
    Case "frmRegTypeChecker"
        bCanSearch = True
        Set out_Control = frmRegTypeChecker.txtKeys
    
    Case "frmHostsMan"
        bCanSearch = True
        Set out_Control = frmHostsMan.lstHostsMan
    
    End Select
    
    If bCanSearch Then
        If TypeOf out_Control Is VBCCR17.ListBoxW Then
            Set lst = out_Control
            If lst.Style = LstStyleCheckbox Then
                For i = 0 To lst.ListCount - 1
                    lst.ItemChecked(i) = True
                Next
            Else
                For i = 0 To lst.ListCount - 1
                    lst.Selected(i) = True
                Next
            End If
        ElseIf TypeOf out_Control Is VBCCR17.TextBoxW Then
            Set txb = out_Control
            txb.SelStart = 0
            txb.SelLength = Len(txb.Text)
        End If
    End If
End Sub

Public Function HasCommandLineKey(ByVal sKey As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim offset As Long
    Dim bHasKey As Boolean
    If UBound(g_sCommandLineArg) > 0 Then
        
        For i = 1 To UBound(g_sCommandLineArg)
            
            bHasKey = False
            
            If StrBeginWith(g_sCommandLineArg(i), "/" & sKey) Then
                bHasKey = True
            ElseIf StrBeginWith(g_sCommandLineArg(i), "-" & sKey) Then
                bHasKey = True
            End If
            
            If bHasKey Then
                If Right$(sKey, 1) = ":" Then offset = -1
                ch = mid$(g_sCommandLineArg(i), Len(sKey) + 2 + offset, 1)
                If (Len(ch) = 0 Or ch = ":") Then
                    HasCommandLineKey = True
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Public Function SectionNameById(idSection As SETTINGS_SECTION) As String

    Dim sName As String

    Select Case idSection
        Case SETTINGS_SECTION_MAIN:                 sName = vbNullString
        Case SETTINGS_SECTION_ADSSPY:               sName = "Tools\ADSSpy"
        Case SETTINGS_SECTION_SIGNCHECKER:          sName = "Tools\SignChecker"
        Case SETTINGS_SECTION_PROCMAN:              sName = "Tools\ProcMan"
        Case SETTINGS_SECTION_STARTUPLIST:          sName = "Tools\StartupList"
        Case SETTINGS_SECTION_UNINSTMAN:            sName = "Tools\UninstMan"
        Case SETTINGS_SECTION_REGUNLOCKER:          sName = "Tools\RegUnlocker"
        Case SETTINGS_SECTION_FILEUNLOCKER:         sName = "Tools\FileUnlocker"
        Case SETTINGS_SECTION_REGKEYTYPECHECKER:    sName = "Tools\RegKeyTypeChecker"
    End Select
    
    SectionNameById = sName
    
End Function

Public Sub ArrayAdd(arr(), Value)

    If 0 = AryPtr(arr) Then
        ReDim arr(0)
    Else
        ReDim Preserve arr(UBound(arr) + 1)
    End If
    
    arr(UBound(arr)) = Value
End Sub

Public Sub ArrayAddStr(arr() As String, Value As String)

    If 0 = AryPtr(arr) Then
        ReDim arr(0)
    Else
        ReDim Preserve arr(UBound(arr) + 1)
    End If
    
    arr(UBound(arr)) = Value
End Sub

Public Sub ArrayAddLong(arr() As Long, Value As Long)

    If 0 = AryPtr(arr) Then
        ReDim arr(0)
    Else
        ReDim Preserve arr(UBound(arr) + 1)
    End If
    
    arr(UBound(arr)) = Value
End Sub

Public Function RoundUp(Num As Double) As Double
    RoundUp = -Int(-Num)
End Function

' Explain:
'
' gSIDs - S-x-x-x list. Includes all active SIDs, excluding current user's SID (read from HKU)
'
' gSID_All - S-x-x-x list. Includes all active and non-active SIDs with current user as well (read from SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList)
'
' gHives - HKU\S-x-x-x list + HKLM + HKCU
'
' aUserOfHive - user name corresponding to gHives (by index)
'
' gHivesUser - HKU\S-x-x-x of other users only (no Service) + HKCU
'
' g_LocalUserNames - all user names (read with NetUserEnum API)
'
' g_LocalGroupNames - all group names (read with NetLocalGroupEnum API)

Public Sub FillUsers()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "FillUsers - Begin"
    
    Dim i&
    
    Erase gSID_All
    Erase gSIDs
    Erase gUserOfHive
    Erase gHives
    Erase gHivesUser
    
    GetHives gHivesUser, addService:=False, addHKLM:=False, addHKCU:=True
    
    GetUserNamesAndSids gSIDs(), gUserOfHive()
    
    Call Reg.EnumSubKeysToArray(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", gSID_All)
    
    ReDim gHives(UBound(gSIDs) + 2)  '+ HKLM, HKCU
    ReDim Preserve gUserOfHive(UBound(gHives))
    
    'Convert SID -> to hive
    For i = 0 To UBound(gSIDs)
        gHives(i) = "HKU\" & gSIDs(i)
    Next
    'Add HKLM, HKCU
    gHives(UBound(gHives) - 1) = "HKLM"
    gUserOfHive(UBound(gHives) - 1) = "All users"
    
    gHives(UBound(gHives)) = "HKCU"
    gUserOfHive(UBound(gHives)) = OSver.UserName
    
    AppendErrorLogCustom "FillUsers - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FillUsers"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub GetHives(aHives() As String, _
    Optional addService As Boolean = False, _
    Optional addHKLM As Boolean = True, _
    Optional addHKCU As Boolean = True)

    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetHives - Begin"
    
    Dim i&, j&, aSID() As String, aUser() As String
    
    GetUserNamesAndSids aSID(), aUser()
    
    ReDim aHives(UBound(aSID))
    
    'Convert SID -> to hive
    For i = 0 To UBound(aSID)
        If Not addService Then
            If aSID(i) = "S-1-5-18" Then GoTo Continue
            If aSID(i) = "S-1-5-19" Then GoTo Continue
            If aSID(i) = "S-1-5-20" Then GoTo Continue
        End If
        If aSID(i) = ".DEFAULT" Then GoTo Continue
        aHives(j) = "HKU\" & aSID(i)
        j = j + 1
Continue:
    Next
    
    j = j - 1
    If j >= 0 Then ReDim Preserve aHives(j)
    
    If addHKCU Then
        j = j + 1
        ReDim Preserve aHives(j)
        aHives(j) = "HKCU"
    End If
    If addHKLM Then
        j = j + 1
        ReDim Preserve aHives(j)
        aHives(j) = "HKLM"
    End If
    
    AppendErrorLogCustom "GetHives - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetHives"
    If inIDE Then Stop: Resume Next
End Sub

Public Function IsValidBuildInUserName(sUsername As String) As Boolean
    
    Static Names() As String
    Static bInit As Boolean
    
    If Not bInit Then
        bInit = True
        ArrayAddStr Names, "System" ' OS uses localized name, however, Tasks xml is not
        ArrayAddStr Names, "LocalSystem"
        ArrayAddStr Names, MapSIDToUsername("S-1-5-18")
        ArrayAddStr Names, MapSIDToUsername("S-1-5-19")
        ArrayAddStr Names, MapSIDToUsername("S-1-5-20")
    End If
    
    ' TODO: append with everything in Well-Known SIDs:
    ' https://docs.microsoft.com/en-US/troubleshoot/windows-server/identity/security-identifiers-in-windows
    '
    ' There are also group-prefixes:
    '
    ' NT AUTHORITY\SYSTEM
    ' BUILTIN\Administrators
    ' etc.
    
    IsValidBuildInUserName = InArray(sUsername, Names, , , vbTextCompare)
    
End Function

' Cases:
'
' S-x-x-x... (SID)
' UserName
' ComputerName\UserName
'
Public Function IsValidTaskUserId(ByVal sUserId As String) As Boolean
    Dim pos As Long
    If Len(sUserId) = 0 Then Exit Function
    If StrBeginWith(sUserId, "S-") Then 'as SID ?
        If IsValidSidEx(sUserId) Then
            IsValidTaskUserId = True
            Exit Function
        End If
    End If
    pos = InStr(sUserId, "\")
    If pos <> 0 Then 'prepended by Computername?
        Dim sComputer As String
        sComputer = Left$(sUserId, pos - 1)
        If StrComp(sComputer, OSver.ComputerName, 1) <> 0 Then
            'we do not check network PCs for validity
            IsValidTaskUserId = True
            Exit Function
        End If
        sUserId = mid$(sUserId, pos + 1)
    End If
    If Len(sUserId) = 0 Then Exit Function
    If InArray(sUserId, g_LocalUserNames, , , vbTextCompare) Then
        IsValidTaskUserId = True
        Exit Function
    End If
    If IsValidBuildInUserName(sUserId) Then IsValidTaskUserId = True
End Function

' Cases:
'
' S-x-x-x... (SID)
' GroupName
' ComputerName\GroupName
'
Public Function IsValidTaskGroupId(ByVal sGroupId As String) As Boolean
    Dim pos As Long
    If Len(sGroupId) = 0 Then Exit Function
    If StrBeginWith(sGroupId, "S-") Then 'as SID ?
        If IsValidSidEx(sGroupId) Then
            IsValidTaskGroupId = True
            Exit Function
        End If
    End If
    pos = InStr(sGroupId, "\")
    If pos <> 0 Then 'prepended by Computername?
        Dim sComputer As String
        sComputer = Left$(sGroupId, pos - 1)
        If StrComp(sComputer, OSver.ComputerName, 1) <> 0 Then
            'we do not check network PCs for validity
            IsValidTaskGroupId = True
            Exit Function
        End If
        
        sGroupId = mid$(sComputer, pos + 1)
    End If
    If Len(sGroupId) = 0 Then Exit Function
    IsValidTaskGroupId = InArray(sGroupId, g_LocalGroupNames, , , vbTextCompare)
End Function

Public Function Deref(ptr As Long) As Long
    If ptr <> 0 Then GetMem4 ByVal ptr, Deref
End Function

Public Function DerefWord(ptr As Long) As Integer
    If ptr <> 0 Then GetMem2 ByVal ptr, DerefWord
End Function

Public Sub ComboSetValue(cmb As VBCCR17.ComboBoxW, sValue As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = sValue Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next
    ErrorMsg Err, "Cannot set value: " & sValue & " - for ComboBox: " & cmb.Name
End Sub

Public Function OS_SupportSHA2() As Boolean

    OS_SupportSHA2 = OSver.IsWindowsXP_SP3OrGreater

End Function

Public Function HexStringToNumber(str As String) As Long
    If IsNumeric(str) Then
        HexStringToNumber = CLng(str)
    Else
        If StrBeginWith(str, "0x") Then
            If IsNumeric(mid$(str, 3)) Then HexStringToNumber = CLng("&H" & mid$(str, 3))
        End If
    End If
End Function

Public Function PathRemoveLastSlash(path As String) As String
    Dim ch As String
    ch = Right$(path, 1)
    If ch = "\" Or ch = "/" Then
        PathRemoveLastSlash = Left$(path, Len(path) - 1)
    Else
        PathRemoveLastSlash = path
    End If
End Function

Public Sub PathRemoveLastSlashInArray(arr() As String)
    Dim i As Long
    For i = 0 To UBound(arr)
        arr(i) = PathRemoveLastSlash(arr(i))
    Next
End Sub

Public Function GetDefaultTextEditorPath() As String
    Dim Cmd As String, path As String
    Cmd = GetDefaultApp(".txt")
    SplitIntoPathAndArgs Cmd, path
    path = EnvironW(path)
    If Not FileExists(path) Then
        path = "rundll32.exe shell32,ShellExec_RunDLL"
    End If
    GetDefaultTextEditorPath = path
End Function

Public Sub OpenInTextEditor(sTextFile As String)
    On Error GoTo ErrorHandler:
    Dim editorPath As String
    editorPath = GetDefaultTextEditorPath()
    If Not FileExists(editorPath) Then
        GoTo tryDefault
    End If
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .lpVerb = StrPtr("open")
        .lpFile = StrPtr(editorPath)
        .lpParameters = StrPtr(sTextFile)
        .lpDirectory = StrPtr(GetParentDir(sTextFile))
        .hWnd = g_HwndMain
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .nShow = SW_SHOWNORMAL
    End With
    If ShellExecuteEx(uSEI) = 0 Then
        GoTo tryDefault
    End If
    If uSEI.hProcess <> 0 Then
        MakeWindowForegroundByProcessHandle uSEI.hProcess
    End If
    Exit Sub
tryDefault:
    Shell "rundll32.exe shell32,ShellExec_RunDLL" & " " & """" & sTextFile & """", vbNormalFocus
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "OpenInTextEditor"
    If inIDE Then Stop: Resume Next
End Sub

Public Function ConvertCollectionToArray(col As Collection) As String()
    Dim i As Long
    Dim a() As String
    ReDim a(col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = col.Item(i)
    Next
    ConvertCollectionToArray = a
End Function

Public Function ScreenHitLine(ByVal sLine As String) As String
    Dim i As Long
    Dim Code As Long
    Dim nStart As Long
    nStart = 1
begin:
    For i = 1 To Len(sLine)
        Code = Asc(mid$(sLine, i, 1))
        If Code >= 0 And Code < 32 Then
            sLine = Replace$(sLine, Chr$(Code), "\x" & Right$("0" & Hex(Code), 2))
            nStart = i + 3
            GoTo begin
        End If
    Next
    ScreenHitLine = doSafeURLPrefix(sLine)
End Function

Private Function GetLastCharPosWithMaxDistReverse(str As String, ch As String, iMaxDist As Long) As Long
    Dim pos As Long
    Dim prevPos As Long
    prevPos = 0
    Do
        pos = InStrRev(str, ch, prevPos - 1)
        If pos = 0 Or (Len(str) - pos) > iMaxDist Then Exit Do
        prevPos = pos
        If pos = 1 Then Exit Do
    Loop
    If prevPos > 0 Then GetLastCharPosWithMaxDistReverse = prevPos
End Function

Public Function LimitHitLineLength(sLine As String, ByVal iLimit As Long) As String
    If Len(sLine) > iLimit Then
        Dim posMark As Long
        'Preserve special marks at the end of the line
        posMark = GetLastCharPosWithMaxDistReverse(sLine, "(", 150)
        If posMark <> 0 Then
            iLimit = iLimit - (Len(sLine) - posMark)
            If iLimit < 1 Then iLimit = 1
        End If
        LimitHitLineLength = Left$(sLine, iLimit) & "... (" & (Len(sLine) - iLimit) & " more chars" & ")" & _
            IIf(posMark = 0, vbNullString, " " & mid$(sLine, posMark))
    Else
        LimitHitLineLength = sLine
    End If
End Function

Public Function IsScriptExtension(sPath As String) As Boolean
    Dim sExt As String
    sExt = GetExtensionName(sPath)
    IsScriptExtension = StrInParamArray(sExt, ".BAT", ".CMD", ".VBS", ".JS", ".PY")
End Function
