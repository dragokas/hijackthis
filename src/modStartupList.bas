Attribute VB_Name = "modStartupList"
'[modStartupList.bas]

'
' StartupList by Merijn Bellekom
'

' fork 2.11 (based on 2.10) by Dragokas
'
' WinTrustVerifyChildNodes. Fixed error with empty node
' istrusted.dll replaced by internal digital signature checking
' list of process replaced by function NtQuerySystemInformation

Option Explicit

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    '  Optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

'Private Type GENERALINPUT
'    dwType As Long
'    xi(0 To 23) As Byte
'End Type

Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32.dll" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
'Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal length As Long)
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Declare Function LoadString Lib "user32.dll" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

'Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As GUID) As Long

Private Declare Function RegOpenKeyEx Lib "Advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyEx Lib "Advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "Advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "Advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "Advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long

'Private Const HKEY_CLASSES_ROOT = &H80000000
'Private Const HKEY_CURRENT_USER = &H80000001
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const HKEY_USERS = &H80000003

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Const WM_KEYDOWN = &H100
Private Const WM_CHAR = &H102
'Private Const WM_SETTEXT = &HC

'Private Const VK_SHIFT = &H10
Private Const VK_HOME = &H24
Private Const VK_RIGHT = &H27
Private Const VK_LEFT = &H25
'Private Const VK_OEM_MINUS = &HBD
'Private Const VK_OEM_5 = &HDC
'Private Const KEYEVENTF_KEYUP = &H2
'Private Const INPUT_MOUSE = 0
'Private Const INPUT_KEYBOARD = 1
'Private Const INPUT_HARDWARE = 2

Private Const SW_SHOWNORMAL = 1

Private Const SEE_MASK_DOENVSUBST = &H200
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_INVOKEIDLIST = &HC

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_NT_WORKSTATION = 1

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3

Public bShowEmpty As Boolean
Public bShowCLSIDs As Boolean
Public bShowCmts As Boolean
Public bShowPrivacy As Boolean
Public bAutoSave As Boolean
Public sAutoSavePath$

Public bShowUsers As Boolean
Public bShowHardware As Boolean

Public lEnumBufLen&
Public sUsernames$(), sHardwareCfgs$()

Private lTicks&
'Public bDebug As Boolean
Public bSL_Abort As Boolean
Public bSL_Terminate As Boolean

Public SEC_RUNNINGPROCESSES As String
Public SEC_AUTOSTARTFOLDERS As String
Public SEC_TASKSCHEDULER As String
Public SEC_INIFILE As String
Public SEC_AUTORUNINF As String
Public SEC_BATFILES As String
Public SEC_EXPLORERCLONES As String

Public SEC_BHOS As String
Public SEC_IETOOLBARS As String
Public SEC_IEEXTENSIONS As String
Public SEC_IEBARS As String
Public SEC_IEMENUEXT As String
Public SEC_IEBANDS As String
Public SEC_DPFS As String
Public SEC_ACTIVEX As String
Public SEC_DESKTOPCOMPONENTS As String
Public SEC_URLSEARCHHOOKS As String

Public SEC_APPPATHS As String
Public SEC_SHELLEXT As String
Public SEC_COLUMNHANDLERS As String
Public SEC_CMDPROC As String
Public SEC_CONTEXTMENUHANDLERS As String
Public SEC_DRIVERFILTERS As String
Public SEC_DRIVERS32 As String
Public SEC_IMAGEFILEEXECUTION As String
Public SEC_LSAPACKAGES As String
Public SEC_MOUNTPOINTS As String
Public SEC_MPRSERVICES As String
Public SEC_ONREBOOT As String
Public SEC_POLICIES As String
Public SEC_PRINTMONITORS As String
Public SEC_PROTOCOLS As String
Public SEC_INIMAPPING As String
Public SEC_REGRUNKEYS As String
Public SEC_REGRUNEXKEYS As String
Public SEC_SECURITYPROVIDERS As String
Public SEC_SERVICES As String
Public SEC_SHAREDTASKSCHEDULER As String
Public SEC_SHELLCOMMANDS As String
Public SEC_SHELLEXECUTEHOOKS As String
Public SEC_SSODL As String
Public SEC_UTILMANAGER As String
Public SEC_WINLOGON As String
Public SEC_SCRIPTPOLICIES As String
Public SEC_WINSOCKLSP As String
Public SEC_WOW As String
Public SEC_3RDPARTY As String

Public SEC_RESETWEBSETTINGS As String
Public SEC_IEURLS As String
Public SEC_URLPREFIX As String
Public SEC_HOSTSFILEPATH As String

Public SEC_HOSTSFILE As String
Public SEC_KILLBITS As String
Public SEC_ZONES As String
Public SEC_MSCONFIG9X As String
Public SEC_MSCONFIGXP As String
Public SEC_STOPPEDSERVICES As String
Public SEC_XPSECURITY As String

Public bShowLargeHosts As Boolean, bShowLargeZones As Boolean

Private Const NUM_OF_SECTIONS As Long = 58

Public Function StartupList_UpdateCaption(frm As Form) As Long

    frm.Caption = "StartupList v." & StartupListVer & " fork" & _
        Replace$(" - " & Translate(906), "[]", NUM_OF_SECTIONS)
    
    StartupList_UpdateCaption = NUM_OF_SECTIONS
End Function
    

Public Sub Status(s$)
    If Not bSL_Terminate Then
        frmStartupList2.stbStatus.SimpleText = s
        DoEvents
    End If
End Sub

Public Function InputFile$(sFile$)
    On Error GoTo ErrorHandler:
    
    'this uses APIs instead of Input(), which is ~3x slower and doesn't cache :P
    Dim hFile&, uBuffer() As Byte, lFileSize&, lBytesRead&
    hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0, OPEN_EXISTING, 0, 0)
    If hFile = -1 Then Exit Function
    
    'second parameter is dwSizeHigh, we ignore that
    'since it's only used if the file is >2GB
    lFileSize = GetFileSize(hFile, 0)
    If lFileSize = -1 Or lFileSize = 0 Then
        CloseHandle hFile
        Exit Function
    End If
    
    ReDim uBuffer(lFileSize - 1)
    If ReadFile(hFile, uBuffer(0), lFileSize, lBytesRead, ByVal 0) > 0 Then
        If lBytesRead <> lFileSize Then
            ReDim Preserve uBuffer(lBytesRead)
        End If
        InputFile = StrConv(uBuffer, vbUnicode)
    End If
    CloseHandle hFile
    Exit Function
ErrorHandler:
    ErrorMsg Err, "InputFile"
    If inIDE Then Stop: Resume Next
End Function

Public Sub ShowFile(sFile$)
    OpenAndSelectFile PathX64(sFile)
End Sub

Public Sub SendToNotepad(sFile$)
    On Error GoTo ErrorHandler:
    If Not FileExists(sFile) Then Exit Sub
    Dim sNotepad$
    sNotepad = Reg.GetString(HKEY_CLASSES_ROOT, ".txt", vbNullString)
    sNotepad = EnvironW(Reg.GetString(HKEY_CLASSES_ROOT, sNotepad & "\shell\open\command", vbNullString))
    If sNotepad <> vbNullString Then
        sNotepad = Left$(sNotepad, InStr(1, sNotepad, ".exe", vbTextCompare) + 3)
        If Not FileExists(sNotepad) Then sNotepad = sWinDir & "\notepad.exe"
    End If
    
    Dim sSEI As SHELLEXECUTEINFO
    With sSEI
        .cbSize = Len(sSEI)
        .hwnd = frmStartupList2.hwnd
        '.lpFile = sWinDir & "\notepad.exe"
        .lpFile = sNotepad
        .lpVerb = "open"
        .lpParameters = PathX64(sFile)
        .fMask = SEE_MASK_DOENVSUBST Or SEE_MASK_FLAG_NO_UI Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_NOCLOSEPROCESS
        .nShow = 1
    End With
    ShellExecuteEx sSEI
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SendToNotepad"
    If inIDE Then Stop: Resume Next
End Sub

Public Function GuessFullpathFromAutorun$(sAutorunFile$)
    On Error GoTo ErrorHandler:
'    Dim sFile$
'    If Trim$(sAutorunFile) = vbNullString Then Exit Function
'    sFile = sAutorunFile
'
'    'already full path? return
'    If InStr(sFile, "\") > 0 And FileExists(sFile) Then
'        GuessFullpathFromAutorun = sFile
'        Exit Function
'    End If
'    'if enclosed in quotes, assume that's the full path and return
'    If InStr(sFile, """") > 0 Then
'        sFile = Mid$(sFile, 2)
'        sFile = Left$(sFile, InStr(sFile, """") - 1)
'    ElseIf InStr(sFile, "\") > 0 And InStr(sFile, " ") > 0 And InStr(1, sFile, ".exe", vbTextCompare) < Len(sFile) - 3 Then
'        'cut off everything after .exe if it's a full path
'        sFile = Left$(sFile, InStr(1, sFile, ".exe", vbTextCompare) + 3)
'    Else
'        'strip everything after the first space (parameters)
'        If InStr(sFile, " ") > 0 Then sFile = Mid$(sFile, 1, InStr(sFile, " ") - 1)
'        'add extension if not there
'        If InStr(sFile, ".") = 0 Then sFile = sFile & ".exe"
'        'try a few common paths to find the file
'        If Not FileExists(sFile) Then
'            'windir
'            If FileExists(BuildPath(sWinDir, sFile)) Then
'                sFile = BuildPath(sWinDir, sFile)
'            Else
'                'sysdir
'                If FileExists(BuildPath(sSysDir, sFile)) Then
'                    sFile = BuildPath(sSysDir, sFile)
'                Else
'                    'root
'                    If FileExists(BuildPath(Left$(sWinDir, 3), sFile)) Then
'                        sFile = BuildPath(Left$(sWinDir, 3), sFile)
'                    End If
'                End If
'            End If
'        End If
'    End If
'    If FileExists(sFile) Then
'        GuessFullpathFromAutorun = sFile
'    Else
'        GuessFullpathFromAutorun = sAutorunFile
'    End If

    Dim sFile$, sArgs$

    SplitIntoPathAndArgs sAutorunFile, sFile, sArgs, bIsRegistryData:=True
                
    GuessFullpathFromAutorun = FormatFileMissing(sFile)

    Exit Function
ErrorHandler:
    ErrorMsg Err, "GuessFullpathFromAutorun"
    If inIDE Then Stop: Resume Next
End Function

Public Sub GetUserNames()
    On Error GoTo ErrorHandler:
    ReDim sUsernames(0)
    Dim sKeys$(), i%
    sKeys = Split(Reg.EnumSubKeys(HKEY_USERS, vbNullString), "|")
    For i = 0 To UBound(sKeys)
        If InStr(1, sKeys(i), "_Classes", vbTextCompare) = 0 Then
            ReDim Preserve sUsernames(UBound(sUsernames) + 1)
            sUsernames(UBound(sUsernames) - 1) = sKeys(i)
        End If
    Next i
    If UBound(sUsernames) > 0 Then
    ReDim Preserve sUsernames(UBound(sUsernames) - 1)
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetUsernames"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub GetHardwareCfgs()
    On Error GoTo ErrorHandler:
    Dim lDefault&, lCurrent&, lLastKnownGood&, lFailed&
    lDefault = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\Select", "Default")
    lCurrent = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\Select", "Current")
    lLastKnownGood = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\Select", "LastKnownGood")
    lFailed = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\Select", "Failed")
    
    ReDim sHardwareCfgs(0)
    sHardwareCfgs(0) = "ControlSet" & Format$(lCurrent, "000")
    If lDefault <> lCurrent And lDefault > 0 Then
        sHardwareCfgs(UBound(sHardwareCfgs)) = "ControlSet" & Format$(lDefault, "000")
    End If
    If lLastKnownGood <> lCurrent And lLastKnownGood > 0 Then
        ReDim Preserve sHardwareCfgs(UBound(sHardwareCfgs) + 1)
        sHardwareCfgs(UBound(sHardwareCfgs)) = "ControlSet" & Format$(lLastKnownGood, "000")
    End If
    If lFailed <> lCurrent And lFailed > 0 Then
        ReDim Preserve sHardwareCfgs(UBound(sHardwareCfgs) + 1)
        sHardwareCfgs(UBound(sHardwareCfgs)) = "ControlSet" & Format$(lFailed, "000")
    End If
    'msgboxw Join(sHardwareCfgs, vbCrLf)
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetHardwareCfgs"
    If inIDE Then Stop: Resume Next
End Sub

Public Function MapControlSetToHardwareCfg$(sControlSet$)
    On Error GoTo ErrorHandler:
    Dim lThisCS&, lDefault&, lCurrent&, lFailed&, lLKG&
    lThisCS = Val(Right$(sControlSet, 3))
    
    lDefault = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\Select", "Default")
    lCurrent = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\Select", "Current")
    lFailed = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\Select", "Failed")
    lLKG = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\Select", "LastKnownGood")
    
    Select Case lThisCS
        Case lDefault: MapControlSetToHardwareCfg = "Default"
        Case lCurrent: MapControlSetToHardwareCfg = "Current"
        Case lFailed:  MapControlSetToHardwareCfg = "Failed"
        Case lLKG:     MapControlSetToHardwareCfg = "Last known good"
    End Select
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MapControlSetToHardwareCfg"
    If inIDE Then Stop: Resume Next
End Function

'Public Sub Jump_(sRegKey$)
'    'this sub has a bug! if regkeys exist similar to the target that
'    'contain spaces, things are screwed up. e.g. a jump to any
'    '"Internet Explorer" key will fail if there is also a key
'    '"Internet Account Manager" present. the space is somehow to blame.
'    Dim lHive&, sKey$, sKeyStrokes$
'    'verify the key actually exists
'    Select Case UCase$(Left$(sRegKey, InStr(sRegKey, "\") - 1))
'        Case "HKEY_CLASSES_ROOT": lHive = HKEY_CLASSES_ROOT
'        Case "HKEY_CURRENT_USER": lHive = HKEY_CURRENT_USER
'        Case "HKEY_LOCAL_MACHINE": lHive = HKEY_LOCAL_MACHINE
'        Case "HKEY_USERS": lHive = HKEY_USERS
'        Case Else: Exit Sub
'    End Select
'    sKey = Mid$(sRegKey, InStr(sRegKey, "\") + 1)
'    If Not Reg.KeyExists(lHive, sKey) Then Exit Sub
'
'    Shell BuildPath(sWinDir, "regedit.exe"), vbNormalFocus
'    'Shell "notepad.exe", vbNormalFocus
'
'    sKeyStrokes = sRegKey
'    sKeyStrokes = Replace$(sKeyStrokes, "{", "{{}")
'    sKeyStrokes = Replace$(sKeyStrokes, "}", "{}}")
'    sKeyStrokes = Replace$(sKeyStrokes, "{{{}}", "{{}")
'    sKeyStrokes = Replace$(sKeyStrokes, "~", "{~}")
'    sKeyStrokes = Replace$(sKeyStrokes, "%", "{%}")
'    sKeyStrokes = Replace$(sKeyStrokes, "^", "{^}")
'    sKeyStrokes = Replace$(sKeyStrokes, "(", "{(}")
'    sKeyStrokes = Replace$(sKeyStrokes, ")", "{)}")
'    sKeyStrokes = Replace$(sKeyStrokes, "+", "{+}")
'    sKeyStrokes = Replace$(sKeyStrokes, "[", "{[}")
'    sKeyStrokes = Replace$(sKeyStrokes, "]", "{]}")
'    sKeyStrokes = Replace$(sKeyStrokes, "\", "{RIGHT}")
'
'    sKeyStrokes = Replace$(sKeyStrokes, " ", vbNullString)
'
'    SendKeys "{HOME}", True
'    SendKeys sKeyStrokes, True
'    SendKeys "{RIGHT}", True
'
''    For i = 1 To Len(sRegKey)
''        Select Case Mid$(sRegKey, i, 1)
''            Case "\" 'send right arrow to expand branch
''                SendKeys "{RIGHT}"
''            'these are special characters and need curly braces
''            Case "~": SendKeys "{~}"
''            Case "%": SendKeys "{%}"
''            Case "^": SendKeys "{^}"
''            Case "(": SendKeys "{(}"
''            Case ")": SendKeys "{)}"
''            Case "+": SendKeys "{+}"
''            Case "{": SendKeys "{{}"
''            Case "}": SendKeys "{}}"
''            'the ONLY character not allowed in a regkey (apart from
''            'high-ascii crap, I suppose) is the BACKSLASH :)
''            Case Else: SendKeys Mid$(sRegKey, i, 1)
''        End Select
''        DoEvents
''    Next i
''    SendKeys "{RIGHT}"
'End Sub

Public Sub DoTicks(tvwMain As TreeView, Optional sNode$)
    If Not bDebug Then Exit Sub
    If bSL_Abort Then Exit Sub
    If sNode = vbNullString Then
        'start
        lTicks = GetTickCount
    Else
        'stop + display
        lTicks = GetTickCount - lTicks
        On Error Resume Next
        tvwMain.Nodes.Add sNode, tvwChild, sNode & "Ticks", " Time: " & lTicks & " ms", "clock"
    End If
End Sub

Private Function isCLSID(sCLSID$) As Boolean
    If sCLSID Like "{????????-????-????-????-????????????}" Then isCLSID = True
End Function

Public Function GetStringResFromDLL$(sFile$, iResID%)
    On Error GoTo ErrorHandler:
    Dim hMod&, lLen&, sBuf$
    If FileExists(sFile) Then
        hMod = LoadLibrary(sFile)
        If hMod > 0 Then
            sBuf = String$(MAX_PATH, 0)
            lLen = LoadString(hMod, Abs(iResID), sBuf, Len(sBuf))
            If lLen > 0 Then GetStringResFromDLL = TrimNull(sBuf)
            FreeLibrary hMod
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetStringResFromDLL"
    If inIDE Then Stop: Resume Next
End Function

Public Sub ShellRun(sFile$, Optional bHidden As Boolean = False)
    On Error GoTo ErrorHandler:
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .lpFile = PathX64(sFile)
        .lpVerb = "open"
        .nShow = Not Abs(CLng(bHidden))
    End With
    ShellExecuteEx uSEI
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ShellRun"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RunScannerGetMD5(sFile$, sKey$)
    On Error GoTo ErrorHandler:
    Dim sMD5$, sAppVer$, sSection$
    sMD5 = GetFileCheckSum(sFile, , True)
    sAppVer = "StartupList" & App.Major & "." & Format$(App.Minor, "00") & "." & App.Revision
    sSection = GetRunScannerItem(GetSectionFromKey(sKey), sKey)
    
    'ShellRun
    OpenURL "https://www.runscanner.net/getMD5.aspx?" & _
      "MD5=" & sMD5 & _
      "&source=" & sAppVer & _
      "&item=" & sSection
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RunScannerGetMD5"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RunScannerGetCLSID(sCLSID$, sKey$)
    On Error GoTo ErrorHandler:
    Dim sAppVer$, sSection$
    sAppVer = "StartupList" & App.Major & "." & Format$(App.Minor, "00") & "." & App.Revision
    sSection = GetRunScannerItem(GetSectionFromKey(sKey), sKey)
    
    'ShellRun
    OpenURL "https://www.runscanner.net/getGUID.aspx?GUID=" & sCLSID & _
          "&source=StartupList" & App.Major & "." & Format$(App.Minor, "00") & "." & App.Revision
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RunScannerGetCLSID"
    If inIDE Then Stop: Resume Next
End Sub

Private Function GetRunScannerItem$(sSection$, sKey$)
    On Error GoTo ErrorHandler:
    Select Case sSection
        Case "RunningProcesses"
            GetRunScannerItem = "001"
        Case "RunRegkeys"
            If InStr(sKey, "System") > 0 Then
                If InStr(sKey, "Once") > 0 Then
                    GetRunScannerItem = "136"
                Else
                    GetRunScannerItem = "002" 'system registry autorun
                End If
            End If
            If InStr(sKey, "User") > 0 Then
                If InStr(sKey, "Once") > 0 Then
                    GetRunScannerItem = "135"
                Else
                    GetRunScannerItem = "003" 'user registry autorun
                End If
            End If
        Case "AutoStartFoldersCommon Startup", "AutoStartFoldersUser Common Startup", "Windows Vista common Startup"
            GetRunScannerItem = "004" 'all users startup
        Case "AutoStartFoldersStartup", "AutoStartFoldersUser Startup"
            GetRunScannerItem = "005" 'user startup
        Case "Windows Vista roaming profile Startup", "Windows Vista roaming profile Startup 2"
            GetRunScannerItem = "007" 'roaming user startup
        Case "NTServices", "VxDServices"
            GetRunScannerItem = "010" 'installed services
        Case "ProtocolsFilter"
            GetRunScannerItem = "030" 'installed protocol filters
        Case "ProtocolsHandler"
            GetRunScannerItem = "031" 'installed protocol handlers
        Case "WinLogonL"
            If InStr(sKey, "WinLogonL0") > 0 Then
                GetRunScannerItem = "033" 'winlogon userinit
            End If
        Case "IniMapping"
            If sKey = "IniMapping0" Then
                GetRunScannerItem = "034"
            Else
                If CInt(Right$(sKey, 1)) Mod 2 = 0 Then
                    GetRunScannerItem = "140"
                Else
                    GetRunScannerItem = "139"
                End If
            End If
        Case "ActiveX"
            GetRunScannerItem = "035"
        Case "WinLogonL1"
            GetRunScannerItem = "037"
        Case "WinLogonL3"
            GetRunScannerItem = "038"
        Case "URLSearchHooks"
            GetRunScannerItem = "040"
        Case "IEToolbars"
            If InStr(sKey, "IEToolbarsUserShell") > 0 Then
                GetRunScannerItem = "045"
            ElseIf InStr(sKey, "IEToolbarsUserWeb") > 0 Then
                GetRunScannerItem = "046"
            Else
                GetRunScannerItem = "041"
            End If
        Case "IEExtensions"
            GetRunScannerItem = "042"
        Case "ShellExecuteHooks"
            GetRunScannerItem = "050"
        Case "SharedTaskScheduler"
            GetRunScannerItem = "051"
        Case "BHO"
            GetRunScannerItem = "052"
        Case "SSODL"
            GetRunScannerItem = "060"
        Case "ShellExts"
            GetRunScannerItem = "061"
        Case "ColumnHandlers"
            GetRunScannerItem = "062"
        Case "OnRebootActionsBootExecute"
            GetRunScannerItem = "063"
        Case "WOWKnownDlls", "WOWKnownDlls32b"
            GetRunScannerItem = "064"
        Case "ImageFileExecution"
            GetRunScannerItem = "065"
        Case "WinLogonL4"
            GetRunScannerItem = "066"
        Case "WinLogonNotify"
            GetRunScannerItem = "067"
        Case "WinsockLSPProtocols"
            GetRunScannerItem = "068"
        Case "PrintMonitors"
            GetRunScannerItem = "069"
        Case "TaskSchedulerJobs"
            If InStr(sKey, "System") = 0 Then
                GetRunScannerItem = "073"
            Else
                GetRunScannerItem = "074"
            End If
        Case "IEURLs"
            GetRunScannerItem = "100"
        Case "IEExplBars"
            GetRunScannerItem = "102"
        Case "DPF"
            GetRunScannerItem = "104"
        Case "WinsockLSPNamespaces"
            GetRunScannerItem = "107"
        Case "WinLogonW"
            If InStr(sKey, "WinLogonW0") > 0 Then
                GetRunScannerItem = "121"
            End If
        Case "WinLogonGinaDLL"
            GetRunScannerItem = "122"
        Case "RunExRegkeys"
            If InStr(sKey, "System") > 0 Then
                If InStr(sKey, "Ex") > 0 Then
                    GetRunScannerItem = "138"
                Else
                    GetRunScannerItem = "136"
                End If
            ElseIf InStr(sKey, "User") > 0 Then
                If InStr(sKey, "Ex") > 0 Then
                    GetRunScannerItem = "137"
                Else
                    GetRunScannerItem = "135"
                End If
            End If
        Case "DriverFiltersClass", "DriverFiltersDevice"
            If InStr(sKey, "Upper") > 0 Then
                GetRunScannerItem = "145"
            End If
        Case "SafeBootAltShell"
            GetRunScannerItem = "146"
        Case "SecurityProviders"
            GetRunScannerItem = "147"
        Case "WOW"
            If sKey = "WOW1" Then
                GetRunScannerItem = "148"
            ElseIf sKey = "WOW2" Then
                GetRunScannerItem = "149"
            End If
        Case "XPSecurityRestore"
            GetRunScannerItem = "150"
        Case "Policies"
            If InStr(sKey, "System") > 0 Then
                GetRunScannerItem = "161"
            ElseIf InStr(sKey, "User") > 0 Then
                GetRunScannerItem = "160"
            End If
        Case "MountPoints", "MountPoints2"
            GetRunScannerItem = "170"
        Case "IniFiles"
            If InStr(sKey, "IniFilessystem.ini3") > 0 Then
                GetRunScannerItem = "171"
            End If
        Case "ContextMenuHandlers"
            GetRunScannerItem = "173"
        Case "ShellCommandsbat", "ShellCommandscmd", "ShellCommandscom", "ShellCommandsexe", "ShellCommandshta", "ShellCommandsjs", "ShellCommandsjse", "ShellCommandspif", "ShellCommandsscr", "ShellCommandstxt", "ShellCommandsvbe", "ShellCommandsvbs", "ShellCommandswsf", "ShellCommandswsh"
            GetRunScannerItem = "180"
        
    End Select
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetRunScannerItem"
    If inIDE Then Stop: Resume Next
End Function

Public Function NodeIsValidFile(objNode As Node) As Boolean
    On Error GoTo ErrorHandler:
    NodeIsValidFile = False
    If objNode.Tag <> vbNullString Then
        If FileExists(objNode.Tag) And Not IsFolder(objNode.Tag) Then
            NodeIsValidFile = True
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "NodeIsValidFile"
    If inIDE Then Stop: Resume Next
End Function

Public Function NodeIsValidRegkey(objNode As Node) As Boolean
    On Error GoTo ErrorHandler:
    NodeIsValidRegkey = False
    If InStr(1, objNode.Tag, "HKEY_") <> 1 Then
        'selected item is not a regkey but a file - climb up in the
        'tree until we find a regkey
        Dim MyNode As Node
        Set MyNode = objNode
        With frmStartupList2.tvwMain
            Do Until MyNode = .Nodes("System") Or _
                     MyNode = .Nodes("Users") Or _
                     MyNode = .Nodes("Hardware")
                Set MyNode = MyNode.Parent
                If InStr(1, MyNode.Tag, "HKEY_") = 1 Then
                    NodeIsValidRegkey = True
                    Exit Function
                End If
            Loop
        End With
    Else
        NodeIsValidRegkey = True
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "NodeIsValidRegkey"
    If inIDE Then Stop: Resume Next
End Function

Public Function NodeExists(sKey$) As Boolean
    Dim s$
    On Error Resume Next
    s = frmStartupList2.tvwMain.Nodes(sKey).Text
    If Err.Number Then
    'If s <> vbNullString Then
        NodeExists = False
    Else
        NodeExists = True
    End If
End Function

Private Function IsFolder(sFile$) As Boolean
    On Error GoTo ErrorHandler:
    If GetFileAttributes(sFile) And FILE_ATTRIBUTE_DIRECTORY Then
        IsFolder = True
    Else
        IsFolder = False
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IsFolder"
    If inIDE Then Stop: Resume Next
End Function

Public Sub RegEnumIEBands(tvwMain As TreeView)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "RegEnumIEBands - Begin"
    
    If bSL_Abort Then Exit Sub
    'Loading... Internet Explorer Bands
    Status Translate(921)
    'HKCR\CLSID\*\Implemented Categories\{00021493-0000-0000-C000-000000000046}
    'HKCR\CLSID\*\Implemented Categories\{00021494-0000-0000-C000-000000000046}
    tvwMain.Nodes.Add "System", tvwChild, "IEBands", SEC_IEBANDS, "msie"
    tvwMain.Nodes("IEBands").Tag = "HKEY_CLASSES_ROOT\CLSID"
    
    Dim hKey&, i&, lNumItems&, sCLSID$, sName$, sFile$
    If RegOpenKeyEx(HKEY_CLASSES_ROOT, "CLSID", 0, KEY_READ, hKey) = 0 Then
        RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        
        sCLSID = String$(MAX_PATH, 0)
        Do Until RegEnumKeyEx(hKey, i, sCLSID, Len(sCLSID), 0, vbNullString, 0, ByVal 0) <> 0
            sCLSID = TrimNull(sCLSID)
    
            If Reg.KeyExists(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\Implemented Categories\{00021493-0000-0000-C000-000000000046}") Or _
               Reg.KeyExists(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\Implemented Categories\{00021494-0000-0000-C000-000000000046}") Then
                sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
                sFile = EnvironW(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
                sFile = GetLongFilename(sFile)
                If bShowCLSIDs Then
                    tvwMain.Nodes.Add "IEBands", tvwChild, "IEBands" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
                Else
                    tvwMain.Nodes.Add "IEBands", tvwChild, "IEBands" & i, sName & " - " & sFile, "dll"
                End If
                tvwMain.Nodes("IEBands" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
    
            sCLSID = String$(MAX_PATH, 0)
            i = i + 1
            If i Mod 100 = 0 And lNumItems > 0 Then
                'Loading... Internet Explorer Bands
                Status Translate(921) & " (" & CInt(i * 100 / lNumItems) & "%, " & i & " CLSIDs)"
            End If
        
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    
    If tvwMain.Nodes("IEBands").Children > 0 Then
        tvwMain.Nodes("IEBands").Text = tvwMain.Nodes("IEBands").Text & " (" & tvwMain.Nodes("IEBands").Children & ")"
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "IEBands"
        End If
    End If
    
    AppendErrorLogCustom "RegEnumIEBands - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RegEnumIEBands"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RegEnumKillBits(tvwMain As TreeView)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "RegEnumKillBits - Begin"
    
    If bSL_Abort Then Exit Sub
    'Loading... ActiveX killbits
    Status Translate(922)
    'HKLM\Software\Microsoft\Internet Explorer\ActiveXCompatibility
    'note: this sub will not show all set Killbits - only those that
    'are actually blocking a CLSID+File that exists on the system.
    Dim sKey$
    sKey = "Software\Microsoft\Internet Explorer\ActiveX Compatibility"
    tvwMain.Nodes.Add "DisabledEnums", tvwChild, "Killbits", SEC_KILLBITS, "msie"
    tvwMain.Nodes("Killbits").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    tvwMain.Nodes("Killbits").Sorted = True
    Dim hKey&, sCLSID$, i&, sName$, sFile$, lKill&
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0, KEY_READ, hKey) = 0 Then
        'RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        
        sCLSID = String$(MAX_PATH, 0)
        Do Until RegEnumKeyEx(hKey, i, sCLSID, Len(sCLSID), 0, vbNullString, 0, ByVal 0) <> 0
            sCLSID = TrimNull(sCLSID)
        
            lKill = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\" & sCLSID, "Compatibility Flags")
            If lKill = 1024 Then
                sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
                sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
                sFile = GetLongFilename(sFile)
                If sFile <> vbNullString Then
                    If sName = vbNullString Then sName = "(no name)"
                    If Not bShowCLSIDs Then
                        tvwMain.Nodes.Add "Killbits", tvwChild, "Killbits" & i, sName & " - " & sFile, "dll"
                    Else
                        tvwMain.Nodes.Add "Killbits", tvwChild, "Killbits" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
                    End If
                    tvwMain.Nodes("Killbits" & i).Tag = GuessFullpathFromAutorun(sFile)
                End If
            End If
            
            sCLSID = String$(MAX_PATH, 0)
            i = i + 1
            'If i Mod 100 = 0 And lNumItems<> 0Then
            '    Status "Loading... ActiveX killbits (" & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " CLSIDs)"
            'End If
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
        
        tvwMain.Nodes("Killbits").Text = tvwMain.Nodes("Killbits").Text & " (" & tvwMain.Nodes("Killbits").Children & ")"
        If tvwMain.Nodes("Killbits").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove "Killbits"
        End If
    End If

    '----------------------------------------------------------------
    'nothing - this is system-wide
    
    AppendErrorLogCustom "RegEnumKillBits - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RegEnumKillBits"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RegEnumZones(tvwMain As TreeView)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "RegEnumZones - Begin"
    
    Dim sKey$, sZoneNames$(), i&, lNumItems&
    Dim hKey&, sDomain$, lZone&, sIcon$, sSubkeys$(), j&, sRange$
    If bSL_Abort Then Exit Sub
    'Loading... Trusted sites & Restricted sites
    Status Translate(923)
    tvwMain.Nodes.Add "DisabledEnums", tvwChild, "Zones", SEC_ZONES, "internet"
    sKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    
    'Loading... Trusted sites & Restricted sites (this user)
    Status Translate(924)
    sZoneNames = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sKey & "\Zones"), "|")
    For i = 0 To UBound(sZoneNames)
        sZoneNames(i) = Reg.GetString(HKEY_CURRENT_USER, sKey & "\Zones\" & sZoneNames(i), "DisplayName")
    Next i
    tvwMain.Nodes.Add "Zones", tvwChild, "ZonesUser", "This user", "user"
    'add root keys for zones
    For i = 0 To UBound(sZoneNames)
        tvwMain.Nodes.Add "ZonesUser", tvwChild, "ZonesUser" & i, sZoneNames(i), "internet"
        tvwMain.Nodes("ZonesUser" & i).Tag = "HKEY_CURRENT_USER\" & sKey & "\ZoneMap\Domains"
    Next i
    If RegOpenKeyEx(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains", 0, KEY_READ, hKey) = 0 Then
        RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        If lNumItems > 1000 And Not bShowLargeZones Then
            'Skipping Zones for this user, since there are over 1000 domains in them. (" & lNumItems & " to be exact)
            frmStartupList2.ShowError Replace$(Translate(2101), "[]", lNumItems)
            RegCloseKey hKey
            GoTo CheckHKCURanges:
        End If
        sDomain = String$(MAX_PATH, 0)
        i = 0
        'loop through subkeys and add them to proper zone
        Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
            If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "http") Then
                lZone = Reg.GetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "http")
            Else
                If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                    lZone = Reg.GetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "*")
                End If
            End If
            
            If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "http") Or _
               RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                Select Case lZone
                    Case 0, 1: sIcon = "system"
                    Case 2: sIcon = "good"
                    Case 3: sIcon = "internet"
                    Case 4: sIcon = "bad"
                    Case Else: sIcon = "internet"
                End Select
                tvwMain.Nodes.Add "ZonesUser" & CStr(lZone), tvwChild, "ZonesUser" & CStr(lZone) & i, sDomain, sIcon
            End If
            'check for subdomains
            sSubkeys = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain), "|")
            If UBound(sSubkeys) > -1 Then
                For j = 0 To UBound(sSubkeys)
                    If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Then
                        lZone = Reg.GetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http")
                    Else
                        If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                            lZone = Reg.GetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*")
                        End If
                    End If
                    
                    If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Or _
                       RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                        Select Case lZone
                            Case 0, 1: sIcon = "system"
                            Case 2: sIcon = "good"
                            Case 3: sIcon = "internet"
                            Case 4: sIcon = "bad"
                            Case Else: sIcon = "internet"
                        End Select
                        tvwMain.Nodes.Add "ZonesUser" & CStr(lZone), tvwChild, "ZonesUser" & CStr(lZone) & i & "s" & j, sSubkeys(j) & "." & sDomain, sIcon
                    End If
                Next j
            End If
            sDomain = String$(MAX_PATH, 0)
            i = i + 1
            If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                'Loading... Trusted sites & Restricted sites (this user,
                Status Replace$(Translate(924), ")", ", " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " domains)")
            End If
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    
CheckHKCURanges:
    If RegOpenKeyEx(HKEY_CURRENT_USER, sKey & "\ZoneMap\Ranges", 0, KEY_READ, hKey) = 0 Then
        'RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        sDomain = String$(MAX_PATH, 0)
        i = 0
        Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
            sDomain = TrimNull(sDomain)
            sRange = Reg.GetString(HKEY_CURRENT_USER, sKey & "\ZoneMap\Ranges\" & sDomain, ":Range")
            lZone = Reg.GetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Ranges\" & sDomain, "*")
            
            If Trim$(sRange) <> vbNullString Then
                Select Case lZone
                    Case 0, 1: sIcon = "system"
                    Case 2: sIcon = "good"
                    Case 3: sIcon = "internet"
                    Case 4: sIcon = "bad"
                    Case Else: sIcon = "internet"
                End Select
                tvwMain.Nodes.Add "ZonesUser" & CStr(lZone), tvwChild, "ZonesUser" & CStr(lZone) & "Range" & i, sRange, sIcon
            End If
            
            sDomain = String$(MAX_PATH, 0)
            i = i + 1
            If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                'Loading... Trusted sites & Restricted sites (this user,
                Status Replace$(Translate(924), ")", ", " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " IP)")
            End If
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    
    For i = 0 To UBound(sZoneNames)
        If tvwMain.Nodes("ZonesUser" & i).Children > 0 Then
            tvwMain.Nodes("ZonesUser" & i).Text = tvwMain.Nodes("ZonesUser" & i).Text & " (" & tvwMain.Nodes("ZonesUser" & i).Children & ")"
            tvwMain.Nodes("ZonesUser" & i).Sorted = True
        Else
            If Not bShowEmpty Then
                tvwMain.Nodes.Remove "ZonesUser" & i
            End If
        End If
    Next i
    If tvwMain.Nodes("ZonesUser").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ZonesUser"
    End If
    
    '---------------------------------
    
    'Loading... Trusted sites & Restricted sites (all users)
    Status Translate(925)
    sZoneNames = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey & "\Zones"), "|")
    For i = 0 To UBound(sZoneNames)
        sZoneNames(i) = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\Zones\" & sZoneNames(i), "DisplayName")
    Next i
    tvwMain.Nodes.Add "Zones", tvwChild, "ZonesSystem", "All users", "users"
    For i = 0 To UBound(sZoneNames)
        tvwMain.Nodes.Add "ZonesSystem", tvwChild, "ZonesSystem" & i, sZoneNames(i), "internet"
        tvwMain.Nodes("ZonesSystem" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKey & "\ZoneMap\Domains"
    Next i
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains", 0, KEY_READ, hKey) = 0 Then
        RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        If lNumItems > 1000 And Not bShowLargeZones Then
            'Skipping Zones for all users, since there are over 1000 domains in them. (" & lNumItems & " to be exact)
            frmStartupList2.ShowError Replace$(Translate(2102), "[]", lNumItems)
            RegCloseKey hKey
            GoTo CheckHKLMRanges:
        End If
        
        sDomain = String$(MAX_PATH, 0)
        i = 0
        Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
            If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "http") Then
                lZone = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "http")
            Else
                If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                    lZone = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "*")
                End If
            End If
            
            If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "http") Or _
               RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                Select Case lZone
                    Case 0, 1: sIcon = "system"
                    Case 2: sIcon = "good"
                    Case 3: sIcon = "internet"
                    Case 4: sIcon = "bad"
                    Case Else: sIcon = "internet"
                End Select
                tvwMain.Nodes.Add "ZonesSystem" & CStr(lZone), tvwChild, "ZonesSystem" & CStr(lZone) & i, sDomain, sIcon
            End If
            'check for subdomains
            sSubkeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain), "|")
            If UBound(sSubkeys) > -1 Then
                For j = 0 To UBound(sSubkeys)
                    If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Then
                        lZone = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http")
                    Else
                        If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                            lZone = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*")
                        End If
                    End If
                    
                    If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Or _
                       RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                        Select Case lZone
                            Case 0, 1: sIcon = "system"
                            Case 2: sIcon = "good"
                            Case 3: sIcon = "internet"
                            Case 4: sIcon = "bad"
                            Case Else: sIcon = "internet"
                        End Select
                        tvwMain.Nodes.Add "ZonesSystem" & CStr(lZone), tvwChild, "ZonesUser" & CStr(lZone) & i & "s" & j, sSubkeys(j) & "." & sDomain, sIcon
                    End If
                Next j
            End If
            
            sDomain = String$(MAX_PATH, 0)
            i = i + 1
            If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                'Loading... Trusted sites & Restricted sites (all users,
                Status Replace$(Translate(925), ")", ", " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " domains)")
            End If
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    
CheckHKLMRanges:
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Ranges", 0, KEY_READ, hKey) = 0 Then
        'RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        sDomain = String$(MAX_PATH, 0)
        i = 0
        Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
            sDomain = TrimNull(sDomain)
            sRange = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Ranges\" & sDomain, ":Range")
            lZone = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Ranges\" & sDomain, "*")
            
            If Trim$(sRange) <> vbNullString Then
                Select Case lZone
                    Case 0, 1: sIcon = "system"
                    Case 2: sIcon = "good"
                    Case 3: sIcon = "internet"
                    Case 4: sIcon = "bad"
                    Case Else: sIcon = "internet"
                End Select
                tvwMain.Nodes.Add "ZonesSystem" & CStr(lZone), tvwChild, "ZonesSystem" & CStr(lZone) & "Range" & i, sRange, sIcon
            End If
            
            sDomain = String$(MAX_PATH, 0)
            i = i + 1
            If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                'Loading... Trusted sites & Restricted sites (all users,
                Status Replace$(Translate(925), ")", ", " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " IPs)")
            End If
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    For i = 0 To UBound(sZoneNames)
        If tvwMain.Nodes("ZonesSystem" & i).Children > 0 Then
            tvwMain.Nodes("ZonesSystem" & i).Text = tvwMain.Nodes("ZonesSystem" & i).Text & " (" & tvwMain.Nodes("ZonesSystem" & i).Children & ")"
            tvwMain.Nodes("ZonesSystem" & i).Sorted = True
        Else
            If Not bShowEmpty Then
                tvwMain.Nodes.Remove "ZonesSystem" & i
            End If
        End If
    Next i
    If tvwMain.Nodes("ZonesSystem").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ZonesSystem"
    End If
        
    If tvwMain.Nodes("Zones").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "Zones"
    End If

    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            'Loading... Trusted sites & Restricted sites
            Status Translate(923) & " (" & sUsername & ")"
            tvwMain.Nodes.Add sUsernames(L) & "DisabledEnums", tvwChild, sUsernames(L) & "Zones", SEC_ZONES, "internet"
            
            For i = 0 To UBound(sZoneNames)
                tvwMain.Nodes.Add sUsernames(L) & "Zones", tvwChild, sUsernames(L) & "ZonesUser" & i, sZoneNames(i), "internet"
                tvwMain.Nodes(sUsernames(L) & "ZonesUser" & i).Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKey & "\ZoneMap\Domains"
            Next i
            If RegOpenKeyEx(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains", 0, KEY_READ, hKey) = 0 Then
                RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
                If lNumItems > 1000 And Not bShowLargeZones Then
                    'Skipping Zones for user " & sUsername & ", since there are over 1000 domains in them. (" & lNumItems & " to be exact)
                    frmStartupList2.ShowError Replace$(Replace$(Translate(2103), "[*]", sUsername), "[**]", lNumItems)
                    RegCloseKey hKey
                    GoTo CheckUserRanges:
                End If
                
                'loop through subkeys and add them to proper zone
                sDomain = String$(MAX_PATH, 0)
                i = 0
                Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
                    If RegValExists(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "http") Then
                        lZone = Reg.GetDword(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "http")
                    Else
                        If RegValExists(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                            lZone = Reg.GetDword(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "*")
                        End If
                    End If
                    
                    If RegValExists(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "http") Or _
                       RegValExists(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                        Select Case lZone
                            Case 0, 1: sIcon = "system"
                            Case 2: sIcon = "good"
                            Case 3: sIcon = "internet"
                            Case 4: sIcon = "bad"
                            Case Else: sIcon = "internet"
                        End Select
                        tvwMain.Nodes.Add sUsernames(L) & "ZonesUser" & CStr(lZone), tvwChild, sUsernames(L) & "ZonesUser" & CStr(lZone) & i, sDomain, sIcon
                    End If
                    'check for subdomains
                    sSubkeys = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain), "|")
                    If UBound(sSubkeys) > -1 Then
                        For j = 0 To UBound(sSubkeys)
                            
                            If RegValExists(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Then
                                lZone = Reg.GetDword(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http")
                            Else
                                If RegValExists(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                                    lZone = Reg.GetDword(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*")
                                End If
                            End If
                            
                            If RegValExists(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Or _
                               RegValExists(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                                Select Case lZone
                                    Case 0, 1: sIcon = "system"
                                    Case 2: sIcon = "good"
                                    Case 3: sIcon = "internet"
                                    Case 4: sIcon = "bad"
                                    Case Else: sIcon = "internet"
                                End Select
                                tvwMain.Nodes.Add sUsernames(L) & "ZonesUser" & CStr(lZone), tvwChild, sUsernames(L) & "ZonesUser" & CStr(lZone) & i & "s" & j, sSubkeys(j) & "." & sDomain, sIcon
                            End If
                        Next j
                    End If
                    
                    i = i + 1
                    sDomain = String$(MAX_PATH, 0)
                    If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                        'Loading... Trusted sites & Restricted sites
                        Status Translate(925) & " (" & sUsername & ", " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " domains)"
                    End If
                    If bSL_Abort Then
                        RegCloseKey hKey
                        Exit Sub
                    End If
                Loop
                RegCloseKey hKey
            End If
            
CheckUserRanges:
            If RegOpenKeyEx(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Ranges", 0, KEY_READ, hKey) = 0 Then
                'RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
                sDomain = String$(MAX_PATH, 0)
                i = 0
                Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
                    sDomain = TrimNull(sDomain)
                    sRange = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Ranges\" & sDomain, ":Range")
                    lZone = Reg.GetDword(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ZoneMap\Ranges\" & sDomain, "*")
                    
                    If lZone > 0 And Trim$(sRange) <> vbNullString Then
                        Select Case lZone
                            Case 0, 1: sIcon = "system"
                            Case 2: sIcon = "good"
                            Case 3: sIcon = "internet"
                            Case 4: sIcon = "bad"
                            Case Else: sIcon = "internet"
                        End Select
                        tvwMain.Nodes.Add sUsernames(L) & "ZonesUser" & CStr(lZone), tvwChild, sUsernames(L) & "ZonesUser" & CStr(lZone) & "Range" & i, sRange, sIcon
                    End If
                    
                    sDomain = String$(MAX_PATH, 0)
                    i = i + 1
                    If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                        'Loading... Trusted sites & Restricted sites
                        Status Translate(923) & " (" & sUsername & ", " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " IPs)"
                    End If
                    If bSL_Abort Then
                        RegCloseKey hKey
                        Exit Sub
                    End If
                Loop
                RegCloseKey hKey
            End If
            
            For i = 0 To UBound(sZoneNames)
                If tvwMain.Nodes(sUsernames(L) & "ZonesUser" & i).Children > 0 Then
                    tvwMain.Nodes(sUsernames(L) & "ZonesUser" & i).Text = tvwMain.Nodes(sUsernames(L) & "ZonesUser" & i).Text & " (" & tvwMain.Nodes(sUsernames(L) & "ZonesUser" & i).Children & ")"
                    tvwMain.Nodes(sUsernames(L) & "ZonesUser" & i).Sorted = True
                Else
                    If Not bShowEmpty Then
                        tvwMain.Nodes.Remove sUsernames(L) & "ZonesUser" & i
                    End If
                End If
            Next i
            
            If tvwMain.Nodes(sUsernames(L) & "Zones").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "Zones"
            End If
        End If
    Next L
    
    AppendErrorLogCustom "RegEnumZones - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RegEnumZones"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RegEnumDriverFilters(tvwMain As TreeView)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "RegEnumDriverFilters - Begin"
    
    'enumerate UpperFilters, LowerFilters on:
    'HKLM\System\CCS\Control\Class\* (Class Lower/Upper Filters)
    'HKLM\System\CCS\Enum\*\*\*      (Device Lower/Upper Filters)
    'HKLM\System\CS?\..etc..
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "DriverFilters", SEC_DRIVERFILTERS, "dll"
    
    Dim hKey&, i&, j&, sKey$, sName$, sLFilters$(), sUFilters$()
    Dim sClassKey$, sDeviceKey$
    sClassKey = "System\CurrentControlSet\Control\Class"
    sDeviceKey = "System\CurrentControlSet\Enum"
    
    tvwMain.Nodes.Add "DriverFilters", tvwChild, "DriverFiltersClass", "Class filters", "dll"
    tvwMain.Nodes("DriverFiltersClass").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey
    tvwMain.Nodes("DriverFiltersClass").Sorted = True
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sClassKey, 0, KEY_READ, hKey) = 0 Then
        sKey = String$(MAX_PATH, 0)
        Do Until RegEnumKeyEx(hKey, i, sKey, Len(sKey), 0, vbNullString, 0, ByVal 0) <> 0
            sKey = TrimNull(sKey)
            sName = Reg.GetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, vbNullString)
            If sName = vbNullString Then sName = "(no name)"
            sLFilters = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, "LowerFilters", False), Chr$(0))
            sUFilters = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, "UpperFilters", False), Chr$(0))
            'root key for device
            If UBound(sLFilters) > 0 Or UBound(sUFilters) > 0 Then
                tvwMain.Nodes.Add "DriverFiltersClass", tvwChild, "DriverFiltersClass" & i, sName, "hardware"
                tvwMain.Nodes("DriverFiltersClass" & i).Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
            End If
            'upper filters
            If UBound(sUFilters) > 0 Then
                tvwMain.Nodes.Add "DriverFiltersClass" & i, tvwChild, "DriverFiltersClass" & i & "Upper", "Upper filters", "dll"
                tvwMain.Nodes("DriverFiltersClass" & i & "Upper").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                For j = 0 To UBound(sUFilters)
                    If Trim$(sUFilters(j)) <> vbNullString Then
                        sName = sUFilters(j) & ".sys"
                        If FileExists(sSysDir & "\drivers\" & sName) Then
                            sName = BuildPath(sSysDir & "\drivers\", sName)
                        Else
                            sName = GuessFullpathFromAutorun(sName)
                        End If
                        tvwMain.Nodes.Add "DriverFiltersClass" & i & "Upper", tvwChild, "DriverFiltersClass" & i & "Upper" & j, sUFilters(j) & ".sys", "dll"
                        tvwMain.Nodes("DriverFiltersClass" & i & "Upper" & j).Tag = sName
                    End If
                Next j
            End If
            'lower filters
            If UBound(sLFilters) > 0 Then
                tvwMain.Nodes.Add "DriverFiltersClass" & i, tvwChild, "DriverFiltersClass" & i & "Lower", "Lower filters", "dll"
                tvwMain.Nodes("DriverFiltersClass" & i & "Lower").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                For j = 0 To UBound(sLFilters)
                    If Trim$(sLFilters(j)) <> vbNullString Then
                        sName = sLFilters(j) & ".sys"
                        If FileExists(sSysDir & "\drivers\" & sName) Then
                            sName = BuildPath(sSysDir & "\drivers\", sName)
                        Else
                            sName = GuessFullpathFromAutorun(sName)
                        End If
                        tvwMain.Nodes.Add "DriverFiltersClass" & i & "Lower", tvwChild, "DriverFiltersClass" & i & "Lower" & j, sLFilters(j) & ".sys", "dll"
                        tvwMain.Nodes("DriverFiltersClass" & i & "Lower" & j).Tag = sName
                    End If
                Next j
            End If
            
            
            sKey = String$(MAX_PATH, 0)
            i = i + 1
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    If tvwMain.Nodes("DriverFiltersClass").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "DriverFiltersClass"
    End If
    '---------------------
    
    tvwMain.Nodes.Add "DriverFilters", tvwChild, "DriverFiltersDevice", "Device filters", "dll"
    tvwMain.Nodes("DriverFiltersDevice").Tag = "HKEY_LOCAL_MACHINE\" & sDeviceKey
    tvwMain.Nodes("DriverFiltersDevice").Sorted = True
    Dim sSections$(), sDevices$(), sSubkeys$(), k&, m&
    'this fucking sucks
    sSections = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey), "|")
    For i = 0 To UBound(sSections)
        sDevices = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i)), "|")
        For j = 0 To UBound(sDevices)
            sSubkeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j)), "|")
            For k = 0 To UBound(sSubkeys)
                sName = Reg.GetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "DeviceDesc")
                If sName = vbNullString Then sName = "(no name)"
                sUFilters = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "UpperFilters", False), Chr$(0))
                sLFilters = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "LowerFilters", False), Chr$(0))
                If UBound(sUFilters) > 0 Or UBound(sLFilters) > 0 Then
                    tvwMain.Nodes.Add "DriverFiltersDevice", tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k, sName, "hardware"
                End If
                If UBound(sUFilters) > 0 Then
                    tvwMain.Nodes.Add "DriverFiltersDevice" & i & "." & j & "." & k, tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k & "Upper", "Upper filters", "dll"
                    For m = 0 To UBound(sUFilters)
                        If Trim$(sUFilters(m)) <> vbNullString Then
                            sName = sUFilters(m) & ".sys"
                            If FileExists(sSysDir & "\drivers\" & sName) Then
                                sName = BuildPath(sSysDir & "\drivers\", sName)
                            Else
                                sName = GuessFullpathFromAutorun(sName)
                            End If
                            tvwMain.Nodes.Add "DriverFiltersDevice" & i & "." & j & "." & k & "Upper", tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k & "Upper" & m, sUFilters(m) & ".sys", "dll"
                            tvwMain.Nodes("DriverFiltersDevice" & i & "." & j & "." & k & "Upper" & m).Tag = sName
                        End If
                    Next m
                End If
                If UBound(sLFilters) > 0 Then
                    tvwMain.Nodes.Add "DriverFiltersDevice" & i & "." & j & "." & k, tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k & "Lower", "Lower filters", "dll"
                    For m = 0 To UBound(sLFilters)
                        If Trim$(sLFilters(m)) <> vbNullString Then
                            sName = sLFilters(m) & ".sys"
                            If FileExists(sSysDir & "\drivers\" & sName) Then
                                sName = BuildPath(sSysDir & "\drivers\", sName)
                            Else
                                sName = GuessFullpathFromAutorun(sName)
                            End If
                            tvwMain.Nodes.Add "DriverFiltersDevice" & i & "." & j & "." & k & "Lower", tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k & "Lower" & m, sLFilters(m) & ".sys", "dll"
                            tvwMain.Nodes("DriverFiltersDevice" & i & "." & j & "." & k & "Lower" & m).Tag = sName
                        End If
                    Next m
                End If
                If bSL_Abort Then Exit Sub
            Next k
        Next j
    Next i
    If tvwMain.Nodes("DriverFiltersDevice").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "DriverFiltersDevice"
    End If
    
    If tvwMain.Nodes("DriverFilters").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "DriverFilters"
    End If
    
    If Not bShowHardware Then Exit Sub
    '-------------------------------------------------------------------------
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sClassKey = "System\" & sHardwareCfgs(L) & "\Control\Class"
        sDeviceKey = "System\" & sHardwareCfgs(L) & "\Enum"
        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "DriverFilters", SEC_DRIVERFILTERS, "dll"
        
        tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFilters", tvwChild, sHardwareCfgs(L) & "DriverFiltersClass", "Class filters", "dll"
        tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersClass").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey
        tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersClass").Sorted = True
        If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sClassKey, 0, KEY_READ, hKey) = 0 Then
            sKey = String$(MAX_PATH, 0)
            Do Until RegEnumKeyEx(hKey, i, sKey, Len(sKey), 0, vbNullString, 0, ByVal 0) <> 0
                sKey = TrimNull(sKey)
                sName = Reg.GetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, vbNullString)
                If sName = vbNullString Then sName = "(no name)"
                sLFilters = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, "LowerFilters", False), Chr$(0))
                sUFilters = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, "UpperFilters", False), Chr$(0))
                'root key for device
                If UBound(sLFilters) > 0 Or UBound(sUFilters) > 0 Then
                    tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersClass", tvwChild, sHardwareCfgs(L) & "DriverFiltersClass" & i, sName, "hardware"
                    tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersClass" & i).Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                End If
                'upper filters
                If UBound(sUFilters) > 0 Then
                    tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersClass" & i, tvwChild, sHardwareCfgs(L) & "DriverFiltersClass" & i & "Upper", "Upper filters", "dll"
                    tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersClass" & i & "Upper").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                    For j = 0 To UBound(sUFilters)
                        If Trim$(sUFilters(j)) <> vbNullString Then
                            sName = sUFilters(j) & ".sys"
                            If FileExists(sSysDir & "\drivers\" & sName) Then
                                sName = BuildPath(sSysDir & "\drivers\", sName)
                            Else
                                sName = GuessFullpathFromAutorun(sName)
                            End If
                            tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersClass" & i & "Upper", tvwChild, sHardwareCfgs(L) & "DriverFiltersClass" & i & "Upper" & j, sUFilters(j) & ".sys", "dll"
                            tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersClass" & i & "Upper" & j).Tag = sName
                        End If
                    Next j
                End If
                'lower filters
                If UBound(sLFilters) > 0 Then
                    tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersClass" & i, tvwChild, sHardwareCfgs(L) & "DriverFiltersClass" & i & "Lower", "Lower filters", "dll"
                    tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersClass" & i & "Lower").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                    For j = 0 To UBound(sLFilters)
                        If Trim$(sLFilters(j)) <> vbNullString Then
                            sName = sLFilters(j) & ".sys"
                            If FileExists(sSysDir & "\drivers\" & sName) Then
                                sName = BuildPath(sSysDir & "\drivers\", sName)
                            Else
                                sName = GuessFullpathFromAutorun(sName)
                            End If
                            tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersClass" & i & "Lower", tvwChild, sHardwareCfgs(L) & "DriverFiltersClass" & i & "Lower" & j, sLFilters(j) & ".sys", "dll"
                            tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersClass" & i & "Lower" & j).Tag = sName
                        End If
                    Next j
                End If
                
                
                sKey = String$(MAX_PATH, 0)
                i = i + 1
                If bSL_Abort Then
                    RegCloseKey hKey
                    Exit Sub
                End If
            Loop
            RegCloseKey hKey
        End If
        If tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersClass").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove sHardwareCfgs(L) & "DriverFiltersClass"
        End If
    
        tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFilters", tvwChild, sHardwareCfgs(L) & "DriverFiltersDevice", "Device filters", "dll"
        tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersDevice").Tag = "HKEY_LOCAL_MACHINE\" & sDeviceKey
        tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersDevice").Sorted = True
        'this fucking sucks - again
        sSections = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey), "|")
        For i = 0 To UBound(sSections)
            sDevices = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i)), "|")
            For j = 0 To UBound(sDevices)
                sSubkeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j)), "|")
                For k = 0 To UBound(sSubkeys)
                    sName = Reg.GetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "DeviceDesc")
                    If sName = vbNullString Then sName = "(no name)"
                    sUFilters = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "UpperFilters", False), Chr$(0))
                    sLFilters = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "LowerFilters", False), Chr$(0))
                    If UBound(sUFilters) > 0 Or UBound(sLFilters) > 0 Then
                        tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersDevice", tvwChild, sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k, sName, "hardware"
                    End If
                    If UBound(sUFilters) > 0 Then
                        tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k, tvwChild, sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k & "Upper", "Upper filters", "dll"
                        For m = 0 To UBound(sUFilters)
                            If Trim$(sUFilters(m)) <> vbNullString Then
                                sName = sUFilters(m) & ".sys"
                                If FileExists(sSysDir & "\drivers\" & sName) Then
                                    sName = BuildPath(sSysDir & "\drivers\", sName)
                                Else
                                    sName = GuessFullpathFromAutorun(sName)
                                End If
                                tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k & "Upper", tvwChild, sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k & "Upper" & m, sUFilters(m) & ".sys", "dll"
                                tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k & "Upper" & m).Tag = sName
                            End If
                        Next m
                    End If
                    If UBound(sLFilters) > 0 Then
                        tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k, tvwChild, sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k & "Lower", "Lower filters", "dll"
                        For m = 0 To UBound(sLFilters)
                            If Trim$(sLFilters(m)) <> vbNullString Then
                                sName = sLFilters(m) & ".sys"
                                If FileExists(sSysDir & "\drivers\" & sName) Then
                                    sName = BuildPath(sSysDir & "\drivers\", sName)
                                Else
                                    sName = GuessFullpathFromAutorun(sName)
                                End If
                                tvwMain.Nodes.Add sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k & "Lower", tvwChild, sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k & "Lower" & m, sLFilters(m) & ".sys", "dll"
                                tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersDevice" & i & "." & j & "." & k & "Lower" & m).Tag = sName
                            End If
                        Next m
                    End If
                    If bSL_Abort Then Exit Sub
                Next k
            Next j
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "DriverFiltersDevice").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove sHardwareCfgs(L) & "DriverFiltersDevice"
        End If
    Next L
    
    AppendErrorLogCustom "RegEnumDriverFilters - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RegEnumDriverFilters"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RegEnumPolicies(tvwMain As TreeView)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "RegEnumPoliciesError - Begin"
    
    Dim Stady As Long
    
    If bSL_Abort Then Exit Sub
    'policies - EVERYTHING
    
    Stady = 1
    tvwMain.Nodes.Add "System", tvwChild, "Policies", SEC_POLICIES, "policy"
    tvwMain.Nodes.Add "Policies", tvwChild, "PoliciesUser", "This user", "user"
    tvwMain.Nodes.Add "Policies", tvwChild, "PoliciesSystem", "All users", "users"
    'enum the tree structures below:
    ' Software\Policies
    ' Software\Microsoft\Windows\CurrentVersion\policies
    ' SOFTWARE\Microsoft\Security Center
    'and then enum all values (REG_SZ, REG_DWORD) in there

    Stady = 2
    Dim sPolicyKeys$(), sPolicyNames$(), k&
    ReDim sPolicyNames(1)
    sPolicyNames(0) = "Primary policies"
    sPolicyNames(1) = "Alternate policies"
    'sPolicyNames(2) = "Security Center policies" - moved to XPSecurityCenter
    ReDim sPolicyKeys(1)
    sPolicyKeys(0) = "Software\Policies"
    sPolicyKeys(1) = "Software\Microsoft\Windows\CurrentVersion\policies"
    'sPolicyKeys(2) = "Software\Microsoft\Security Center" - moved to XPSecurityCenter

    Dim sRegKeysUser$(), sRegKeysSystem$(), sValues$(), i&, j&


    Stady = 3
    For k = 0 To UBound(sPolicyKeys)
        Stady = 4
        tvwMain.Nodes.Add "PoliciesUser", tvwChild, "Policies" & k & "User", sPolicyNames(k), "winlogon"
        tvwMain.Nodes.Add "PoliciesSystem", tvwChild, "Policies" & k & "System", sPolicyNames(k), "winlogon"
        
        Stady = 5
        tvwMain.Nodes("Policies" & k & "User").Tag = "HKEY_CURRENT_USER\" & sPolicyKeys(k)
        tvwMain.Nodes("Policies" & k & "System").Tag = "HKEY_LOCAL_MACHINE\" & sPolicyKeys(k)
    
        Stady = 6
        sValues = Split(RegEnumValues(HKEY_CURRENT_USER, sPolicyKeys(k), , , False), "|")
        For j = 0 To UBound(sValues)
            Stady = 7
            tvwMain.Nodes.Add "Policies" & k & "User", tvwChild, "Policies" & k & "User" & j & "Root", sValues(j), "reg"
        Next j

        Stady = 10
        sRegKeysUser = Split(EnumSubKeysTree(HKEY_CURRENT_USER, sPolicyKeys(k)), "|")
        
        Stady = 11
        For i = 0 To UBound(sRegKeysUser)
            Stady = 12
            sValues = Split(RegEnumValues(HKEY_CURRENT_USER, sRegKeysUser(i), , , False), "|")
            If UBound(sValues) > -1 Then
                Stady = 13
                tvwMain.Nodes.Add "Policies" & k & "User", tvwChild, "Policies" & k & "User" & i, sRegKeysUser(i), "registry"
                Stady = 14
                tvwMain.Nodes("Policies" & k & "User" & i).Tag = "HKEY_CURRENT_USER\" & sRegKeysUser(i)
                Stady = 15
                For j = 0 To UBound(sValues)
                    Stady = 16
                    tvwMain.Nodes.Add "Policies" & k & "User" & i, tvwChild, "Policies" & k & "User" & i & "." & j, sValues(j), "reg"
                Next j
                Stady = 17
                tvwMain.Nodes("Policies" & k & "User" & i).Text = tvwMain.Nodes("Policies" & k & "User" & i).Text & " (" & tvwMain.Nodes("Policies" & k & "User" & i).Children & ")"
            End If
            If bSL_Abort Then Exit Sub
        Next i
        
        sValues = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sPolicyKeys(k), , , False), "|")
        For j = 0 To UBound(sValues)
            Stady = 9
            tvwMain.Nodes.Add "Policies" & k & "System", tvwChild, "Policies" & k & "System" & j & "Root", sValues(j), "reg"
        Next j
        
        sRegKeysSystem = Split(EnumSubKeysTree(HKEY_LOCAL_MACHINE, sPolicyKeys(k)), "|")
        
        For i = 0 To UBound(sRegKeysSystem)
            Stady = 18
            sValues = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sRegKeysSystem(i), , , False), "|")
            If UBound(sValues) > -1 Then
                Stady = 19
                tvwMain.Nodes.Add "Policies" & k & "System", tvwChild, "Policies" & k & "System" & i, sRegKeysSystem(i), "registry"
                Stady = 20
                tvwMain.Nodes("Policies" & k & "System" & i).Tag = "HKEY_LOCAL_MACHINE\" & sRegKeysSystem(i)
                For j = 0 To UBound(sValues)
                    Stady = 21
                    tvwMain.Nodes.Add "Policies" & k & "System" & i, tvwChild, "Policies" & k & "System" & i & "." & j, sValues(j), "reg"
                Next j
                Stady = 22
                tvwMain.Nodes("Policies" & k & "System" & i).Text = tvwMain.Nodes("Policies" & k & "System" & i).Text & " (" & tvwMain.Nodes("Policies" & k & "System" & i).Children & ")"
            End If
            If bSL_Abort Then Exit Sub
        Next i

        Stady = 23
        If tvwMain.Nodes("Policies" & k & "User").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove "Policies" & k & "User"
        End If
        Stady = 24
        If tvwMain.Nodes("Policies" & k & "System").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove "Policies" & k & "System"
        End If
        If bSL_Abort Then Exit Sub
    Next k

    Stady = 25
    If tvwMain.Nodes("PoliciesUser").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "PoliciesUser"
    End If
    Stady = 26
    If tvwMain.Nodes("PoliciesSystem").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "PoliciesSystem"
    End If
    Stady = 27
    If tvwMain.Nodes("Policies").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "Policies"
    End If

    If Not bShowUsers Then Exit Sub
    '-----------------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        Stady = 28
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            Stady = 29
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "PoliciesUser", SEC_POLICIES, "policy"

            Stady = 30
            For k = 0 To UBound(sPolicyKeys)
                Stady = 31
                tvwMain.Nodes.Add sUsernames(L) & "PoliciesUser", tvwChild, sUsernames(L) & "Policies" & k & "User", sPolicyNames(k), "winlogon"
                Stady = 32
                tvwMain.Nodes(sUsernames(L) & "Policies" & k & "User").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sPolicyKeys(k)

                Stady = 33
                sValues = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\" & sPolicyKeys(k), , , False), "|")
                For j = 0 To UBound(sValues)
                    Stady = 34
                    tvwMain.Nodes.Add sUsernames(L) & "Policies" & k & "User", tvwChild, sUsernames(L) & "Policies" & k & "User" & j & "Root", sValues(j), "reg"
                Next j
    
                Stady = 35
                sRegKeysUser = Split(EnumSubKeysTree(HKEY_USERS, sUsernames(L) & "\" & sPolicyKeys(k)), "|")

                For i = 0 To UBound(sRegKeysUser)
                    Stady = 36
                    sValues = Split(RegEnumValues(HKEY_USERS, sRegKeysUser(i), , , False), "|")
                    If UBound(sValues) > -1 Then
                        Stady = 37
                        tvwMain.Nodes.Add sUsernames(L) & "Policies" & k & "User", tvwChild, sUsernames(L) & "Policies" & k & "User" & i, Mid$(sRegKeysUser(i), Len(sUsernames(L)) + 2), "registry"
                        Stady = 38
                        tvwMain.Nodes(sUsernames(L) & "Policies" & k & "User" & i).Tag = "HKEY_USERS\" & sRegKeysUser(i)
                        For j = 0 To UBound(sValues)
                            Stady = 39
                            tvwMain.Nodes.Add sUsernames(L) & "Policies" & k & "User" & i, tvwChild, sUsernames(L) & "Policies" & k & "User" & i & "." & j, sValues(j), "reg"
                        Next j
                        Stady = 40
                        tvwMain.Nodes(sUsernames(L) & "Policies" & k & "User" & i).Text = tvwMain.Nodes(sUsernames(L) & "Policies" & k & "User" & i).Text & " (" & tvwMain.Nodes(sUsernames(L) & "Policies" & k & "User" & i).Children & ")"
                    End If
                    If bSL_Abort Then Exit Sub
                Next i

                Stady = 41
                If tvwMain.Nodes(sUsernames(L) & "Policies" & k & "User").Children = 0 And Not bShowEmpty Then
                    Stady = 42
                    tvwMain.Nodes.Remove sUsernames(L) & "Policies" & k & "User"
                End If
            Next k
            
            Stady = 43
            If tvwMain.Nodes(sUsernames(L) & "PoliciesUser").Children = 0 And Not bShowEmpty Then
                Stady = 44
                tvwMain.Nodes.Remove sUsernames(L) & "PoliciesUser"
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next L
    
    AppendErrorLogCustom "RegEnumPoliciesError - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RegEnumPolicies", "Stady: " & Stady & ", Iteration: [K] = " & k, " [J] = " & j & " [i] = " & i & " [L] = " & L
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RegEnumDrivers32(tvwMain As TreeView)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "RegEnumDrivers32 - Begin"
    
    If bSL_Abort Then Exit Sub
    Const sDrivers$ = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Drivers32"
    
    tvwMain.Nodes.Add "System", tvwChild, "Drivers32", SEC_DRIVERS32, "dll"
    tvwMain.Nodes("Drivers32").Tag = "HKEY_LOCAL_MACHINE\" & sDrivers
    tvwMain.Nodes("Drivers32").Sorted = True
    Dim i&, sDriverKeys$()
    sDriverKeys = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sDrivers), "|")
    For i = 0 To UBound(sDriverKeys)
        tvwMain.Nodes.Add "Drivers32", tvwChild, "Drivers32" & i, sDriverKeys(i), "dll", "dll"
        tvwMain.Nodes("Drivers32" & i).Tag = GuessFullpathFromAutorun(Mid$(sDriverKeys(i), InStrRev(sDriverKeys(i), " = ") + 3))
        If bSL_Abort Then Exit Sub
    Next i
    
    tvwMain.Nodes.Add "Drivers32", tvwChild, "Drivers32RDP", " Terminal Services", "internet", "internet"
    tvwMain.Nodes("Drivers32RDP").Tag = "HKEY_LOCAL_MACHINE\" & sDrivers & "\Terminal Server\RDP"
    tvwMain.Nodes("Drivers32RDP").Sorted = True
    sDriverKeys = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sDrivers & "\Terminal Server\RDP"), "|")
    For i = 0 To UBound(sDriverKeys)
        tvwMain.Nodes.Add "Drivers32RDP", tvwChild, "Drivers32RDP" & i, sDriverKeys(i), "dll", "dll"
        tvwMain.Nodes("Drivers32RDP" & i).Tag = GuessFullpathFromAutorun(Mid$(sDriverKeys(i), InStrRev(sDriverKeys(i), " = ") + 3))
    Next i
    
    If tvwMain.Nodes("Drivers32RDP").Children > 0 Then
        tvwMain.Nodes("Drivers32RDP").Text = tvwMain.Nodes("Drivers32RDP").Text & " (" & tvwMain.Nodes("Drivers32RDP").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "Drivers32RDP"
    End If
    If tvwMain.Nodes("Drivers32").Children > 0 Then
        tvwMain.Nodes("Drivers32").Text = tvwMain.Nodes("Drivers32").Text & " (" & tvwMain.Nodes("Drivers32").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "Drivers32"
    End If
    
    AppendErrorLogCustom "RegEnumDrivers32 - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RegEnumDrivers32"
    If inIDE Then Stop: Resume Next
End Sub

Private Function EnumSubKeysTree$(lHive&, sRootKey$)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumSubKeysTree - Begin"
    
    Dim hKey&, i&, sName$, sList$
    If bSL_Abort Then Exit Function
    If RegOpenKeyEx(lHive, sRootKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(MAX_PATH, 0)
        Do Until RegEnumKeyEx(hKey, i, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0
            sName = TrimNull(sName)
            
            sList = sList & "|" & sRootKey & "\" & sName
            sList = sList & "|" & EnumSubKeysTree(lHive, sRootKey & "\" & sName)
            
            i = i + 1
            sName = String$(MAX_PATH, 0)
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then EnumSubKeysTree = Mid$(Replace$(sList, "||", "|"), 2)
    
    AppendErrorLogCustom "EnumSubKeysTree - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumSubKeysTree"
    If inIDE Then Stop: Resume Next
End Function

Public Function ExpandEnvironmentVars(sEnvStr As String) As String
    ExpandEnvironmentVars = EnvironW(sEnvStr)
End Function

Private Function GetString(lHive&, sKey$, sVal$, Optional bTrimNull As Boolean = True) As String
    On Error GoTo ErrorHandler:
    Dim hKey&, uData() As Byte, lDataLen&, sData$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        RegQueryValueEx hKey, sVal, 0, 0, ByVal 0, lDataLen
        ReDim uData(lDataLen)
        If RegQueryValueEx(hKey, sVal, 0, 0, uData(0), lDataLen) = 0 Then
            If bTrimNull Then
                sData = StrConv(uData, vbUnicode)
                sData = TrimNull(sData)
            Else
                If lDataLen > 2 Then
                    ReDim Preserve uData(lDataLen - 2)
                    sData = StrConv(uData, vbUnicode)
                End If
            End If
            GetString = sData
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetString"
    If inIDE Then Stop: Resume Next
End Function

Private Function GetDword&(lHive$, sKey$, sVal$)
    On Error GoTo ErrorHandler:
    Dim hKey&, lData&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        If RegQueryValueEx(hKey, sVal, 0, 0, lData, 4) = 0 Then
            GetDword = lData
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetDword"
    If inIDE Then Stop: Resume Next
End Function

Private Function KeyExists(lHive&, sKey$) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        KeyExists = True
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "KeyExists"
    If inIDE Then Stop: Resume Next
End Function

Private Function RegValExists(lHive&, sKey$, sVal$) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lDataLen&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        If RegQueryValueEx(hKey, sVal, 0, 0, ByVal 0, lDataLen) = 0 Then
            RegValExists = True
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegValExists"
    If inIDE Then Stop: Resume Next
End Function

Private Function EnumSubKeys$(lHive&, sKey$)
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, sName$, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(MAX_PATH, 0)
        Do Until RegEnumKeyEx(hKey, i, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0
            sName = TrimNull(sName)
            sList = sList & "|" & sName
            i = i + 1
            sName = String$(MAX_PATH, 0)
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then EnumSubKeys = Mid$(sList, 2)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumSubKeys"
    If inIDE Then Stop: Resume Next
End Function

Private Function RegEnumValues$(lHive&, sKey$, Optional bNullSep As Boolean = False, Optional bIgnoreBinaries As Boolean = True, Optional bIgnoreDwords As Boolean = True)
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, sName$, uData() As Byte, lDataLen&
    Dim lType&, sData$, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(lEnumBufLen, 0)
        ReDim uData(32768)
        lDataLen = UBound(uData)
        Do Until RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0
            
            sName = TrimNull(sName)
            If sName = vbNullString Then sName = "@"
            
            If lType = REG_SZ Then
                ReDim Preserve uData(lDataLen)
                sData = TrimNull(StrConv(uData, vbUnicode))
                If bNullSep Then
                    sList = sList & Chr$(0) & sName & " = " & sData
                Else
                    sList = sList & "|" & sName & " = " & sData
                End If
            End If
            
            If lType = REG_BINARY And Not bIgnoreBinaries Then
                sList = sList & "|" & sName & " (binary)"
            End If
            
            If lType = REG_DWORD And Not bIgnoreDwords Then
                'look at me! I'm haxxoring word values from binary!
                'sData = "dword: " & Hex$(uData(0)) & "." & Hex$(uData(1)) & "." & Hex$(uData(2)) & "." & Hex$(uData(3))
                'sData = "dword: " & Val("&H" & Hex$(uData(3)) & Hex$(uData(2)) & Hex$(uData(1)) & Hex$(uData(0)))
                sData = "dword: " & CStr(16 ^ 6 * uData(3) + 16 ^ 4 * uData(2) + 16 ^ 2 * uData(1) + uData(0))
                sList = sList & "|" & sName & " = " & sData
            End If
            sName = String$(lEnumBufLen, 0)
            ReDim uData(32768)
            lDataLen = UBound(uData)
            i = i + 1
            
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then RegEnumValues = Mid$(sList, 2)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegEnumValues"
    If inIDE Then Stop: Resume Next
End Function

Private Function RegEnumDwordValues$(lHive&, sKey$)
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, sName$, uData() As Byte, lDataLen&
    Dim lType&, lData&, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(lEnumBufLen, 0)
        ReDim uData(32768)
        lDataLen = UBound(uData)
        Do Until RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0
            If lType = REG_DWORD And lDataLen = 4 Then
                sName = TrimNull(sName)
                If sName = vbNullString Then sName = "@"
                ReDim Preserve uData(4)
                CopyMemory lData, uData(0), 4
                sList = sList & "|" & sName & " = " & CStr(lData)
            End If
            sName = String$(lEnumBufLen, 0)
            ReDim uData(32768)
            lDataLen = UBound(uData)
            i = i + 1
        
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then RegEnumDwordValues = Mid$(sList, 2)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegEnumDwordValues"
    If inIDE Then Stop: Resume Next
End Function

Public Sub StartupList_Abort()
    bSL_Abort = True
    bSL_Terminate = True
End Sub
