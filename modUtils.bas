Attribute VB_Name = "modUtils"
Option Explicit

Private Type BROWSER_INFO
    Version As String
    'Path    As String  'future use
End Type

Public Type MY_BROWSERS
    Edge    As BROWSER_INFO
    IE      As BROWSER_INFO
    Firefox As BROWSER_INFO
    Opera   As BROWSER_INFO
    Chrome  As BROWSER_INFO
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(255) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type


Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExW" (ByVal lpFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadString Lib "user32.dll" Alias "LoadStringW" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As Long, ByVal nBufferMax As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStrDest As Long, ByVal lpStrSrc As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" (lpSystemTime As SYSTEMTIME, vtime As Date) As Long
Private Declare Function VariantTimeToSystemTime Lib "oleaut32.dll" (ByVal vtime As Date, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32.dll" (ByVal lpTimeZone As Any, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByVal lpFileTime As Long, lpSystemTime As SYSTEMTIME) As Long
'Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
'Private Declare Function GetTimeZoneInformation Lib "kernel32.dll" (ByVal lpTimeZoneInformation As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32.dll" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As OSVERSIONINFOEX) As Long

Private Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2   'Read Only      ( do not execute DllMain )

Public Const GWL_WNDPROC    As Long = &HFFFFFFFC
Public Const WM_MOUSEWHEEL  As Long = &H20A&

Public Const SHCNE_DELETE       As Long = 4&
Public Const SHCNF_PATH         As Long = 1&
Public Const SHCNF_FLUSHNOWAIT  As Long = &H2000&
Public Const SHCNE_CREATE       As Long = 2&
Public Const SHCNE_RENAMEITEM   As Long = 1&
Public Const SHCNE_ATTRIBUTES   As Long = &H800&

Private Const FILE_ATTRIBUTE_READONLY   As Long = 1&
Private Const ERROR_FILE_NOT_FOUND      As Long = 2&
Private Const ERROR_ACCESS_DENIED       As Long = 5&

Public BROWSERS As MY_BROWSERS

Dim OFName As OPENFILENAME

Private lpPrevWndProc As Long

Public Sub SubClassScroll(SwitchON As Boolean)
    If DisableSubclassing Then Exit Sub
    If SwitchON Then
        If lpPrevWndProc = 0 Then lpPrevWndProc = SetWindowLong(frmMain.hWnd, GWL_WNDPROC, AddressOf WndProc)
    Else
        If lpPrevWndProc Then SetWindowLong frmMain.hWnd, GWL_WNDPROC, lpPrevWndProc: lpPrevWndProc = 0
    End If
End Sub

Function IsMouseWithin(hWnd As Long) As Boolean
    Dim r As RECT
    Dim p As POINTAPI
    If GetWindowRect(hWnd, r) Then
        If GetCursorPos(p) Then
            IsMouseWithin = PtInRect(r, p.X, p.Y)
        End If
    End If
End Function

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
    Case WM_MOUSEWHEEL
        If Not IsMouseWithin(frmMain.hWnd) Then Exit Function ' mouse is outside the form
        Dim MouseKeys&, Rotation&, NewValue%
        MouseKeys = wParam And &HFFFF&
        Rotation = wParam / &HFFFF& 'direction
        With frmMain.vscMiscTools
            NewValue = .value - .LargeChange * IIf(Rotation > 0, 1, -1)
            If NewValue < .Min Then NewValue = .Min
            If NewValue > .Max Then NewValue = .Max
            .value = NewValue   'change scroll value
        End With
    Case Else
        WndProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function

' affected by wow64
Public Function OpenFileDialog(Optional sTitle As String, Optional InitDir As String, Optional hWnd As Long) As String
    On Error GoTo ErrorHandler
    Const OFN_DONTADDTORECENT As Long = &H2000000
    Const OFN_ENABLESIZING As Long = &H800000
    Const OFN_FORCESHOWHIDDEN As Long = &H10000000
    Const OFN_HIDEREADONLY As Long = 4&
    Const OFN_NODEREFERENCELINKS As Long = &H100000
    Const OFN_NOVALIDATE As Long = &H100&
    
    Dim ofn As OPENFILENAME
    Dim out As String
    Dim Filter As String, i As Long
    
    ofn.nMaxFile = MAX_PATH_W
    out = String$(MAX_PATH_W, vbNullChar)
    'Filter = IIf(isTextFile, "Текстовые файлы" & vbNullChar & "*.txt" & vbNullChar, _
    '    "Исполняемые файлы" & vbNullChar & "*.exe;*.dll" & vbNullChar)
    Filter = "All Files" & vbNullChar & "*.*" & vbNullChar
    With ofn
        .hWndOwner = IIf(hWnd = 0, frmMain.hWnd, hWnd)
        .lpstrTitle = StrPtr(sTitle)
        .lpstrFile = StrPtr(out)
        .lStructSize = Len(ofn)
        .lpstrFilter = StrPtr(Filter)
        .lpstrInitialDir = StrPtr(InitDir)
        .flags = OFN_DONTADDTORECENT Or OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN Or OFN_HIDEREADONLY Or OFN_NODEREFERENCELINKS Or OFN_NOVALIDATE
    End With
    If GetOpenFileName(ofn) Then OpenFileDialog = TrimNull(out)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "OpenFileDialog"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetStringFromBinary(Optional ByVal sFile As String, Optional ByVal nid As Long, Optional ByVal FileAndIDHybrid As String) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetStringFromBinary - Begin", "File: " & sFile, "id: " & nid, "FileAndIDHybrid: " & FileAndIDHybrid

    Dim hModule As Long
    Dim nSize As Long
    Dim sBuf As String
    Dim pos As Long
    Dim Redirect As Boolean, bOldStatus As Boolean
    
    'Get string resource from binary file
    'Source can be defined either by Filename and ID, or by hibryd (registry like) string, e.g. @%SystemRoot%\System32\my.dll,-102
    'ID with minus will be converted by module
    
    If 0 <> Len(FileAndIDHybrid) Then
    
        If Left$(FileAndIDHybrid, 1) = "@" Then FileAndIDHybrid = Mid$(FileAndIDHybrid, 2)
        If InStr(FileAndIDHybrid, "%") <> 0 Then FileAndIDHybrid = EnvironW(FileAndIDHybrid)
        pos = InStrRev(FileAndIDHybrid, ",")
        If 0 <> pos Then
            sFile = Left$(FileAndIDHybrid, pos - 1)
            sBuf = Mid$(FileAndIDHybrid, pos + 1)
            If IsNumeric(sBuf) And 0 <> Len(sBuf) Then nid = Val(sBuf) Else Exit Function
        End If
    End If
    
    If 0 = Len(sFile) Then Exit Function
    
    If Not FileExists(sFile) Then
        sFile = FindOnPath(sFile)
        If 0 = Len(sFile) Then Exit Function
    End If
    
    sBuf = String$(160, 0)
    
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    hModule = LoadLibraryEx(StrPtr(sFile), 0&, LOAD_LIBRARY_AS_DATAFILE)
    
    If hModule Then
        nSize = LoadString(hModule, Abs(nid), StrPtr(sBuf), LenB(sBuf))
        If nSize > 0 Then
            GetStringFromBinary = TrimNull(Left$(sBuf, nSize))
        End If
        FreeLibrary hModule
    End If
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "GetStringFromBinary - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetStringFromBinary", sFile
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Sub GetBrowsersInfo()
    AppendErrorLogCustom "GetBrowsersInfo - Begin"
    With BROWSERS
        .Edge.Version = GetEdgeVersion()
        .IE.Version = GetMSIEVersion()
        .Chrome.Version = GetChromeVersion()
        .Firefox.Version = GetFirefoxVersion()
        .Opera.Version = GetOperaVersion()
    End With
    AppendErrorLogCustom "GetBrowsersInfo - End"
End Sub

Public Function GetEdgeVersion() As String
    AppendErrorLogCustom "GetEdgeVersion - Begin"
    Dim EdgePath$
    EdgePath = sWinDir & "\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe"
    If FileExists(EdgePath) Then GetEdgeVersion = GetFilePropVersion(EdgePath)
    AppendErrorLogCustom "GetEdgeVersion - End"
End Function

Public Function GetChromeVersion() As String
    AppendErrorLogCustom "GetChromeVersion - Begin"
    Dim sVer$, sPath$
    sVer = RegGetString(HKEY_LOCAL_MACHINE, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv")
    'not found try current user - win7(x86)
    If Len(sVer) = 0 Then
        sVer = RegGetString(HKEY_CURRENT_USER, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv")
    End If
    If Len(sVer) = 0 Then 'Wow6432Node
        sVer = RegGetString(HKEY_LOCAL_MACHINE, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv", True)
    End If
    If Len(sVer) = 0 Then
        sPath = RegGetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", vbNullString)
        If sPath <> "" Then
            sVer = GetFilePropVersion(sPath)
        End If
    End If
    If Len(sVer) = 0 Then
        sVer = RegGetString(HKEY_CURRENT_USER, "SOFTWARE\Google\Chrome\BLBeacon", "version")
    End If
    GetChromeVersion = sVer
    AppendErrorLogCustom "GetChromeVersion - End"
End Function

Public Function GetFirefoxVersion() As String
    AppendErrorLogCustom "GetFirefoxVersion - Begin"
    Dim sVer$, sPath$
    
    sPath = RegGetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe", vbNullString)
    If Len(sPath) <> 0 Then
        sVer = GetFilePropVersion(sPath)
        If Len(sVer) = 0 Then sVer = RegGetString(HKEY_LOCAL_MACHINE, "Software\Mozilla\Mozilla Firefox", "CurrentVersion")
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
    sOperaPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\Opera.exe", vbNullString)

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
        sMSIEPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE", vbNullString)
        If sMSIEPath <> vbNullString Then
            sMSIEPath = Trim$(sMSIEPath)
            If FileExists(sMSIEPath) Then
                sMSIEVersion = GetFilePropVersion(sMSIEPath)
            End If
        End If
    End If
       
    If 0 = Len(sMSIEVersion) Then
        sMSIEVersion = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "svcVersion")
        If 0 = Len(sMSIEVersion) Then
            sMSIEVersion = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "version")
        End If
    End If
    
    sMSIEHotfixes = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MinorVersion")
    
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

' Особая процедура удаления файла (с разблокировкой NTFS привилегий)
Public Function DeleteFileForce(File As String) As Long
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "DeleteFileForce - Begin", "File: " & File
    
    Const FILE_ATTRIBUTE_NORMAL     As Long = &H80&
    Const FILE_ATTRIBUTE_READONLY   As Long = 1&
    
    Dim lr          As Long
    Dim Attrib      As Long
    Dim isDeleted   As Boolean
    Dim Redirect As Boolean, bOldStatus As Boolean

    If IsMicrosoftFile(File) Then Exit Function

    Redirect = ToggleWow64FSRedirection(False, File, bOldStatus)

    Attrib = GetFileAttributes(StrPtr(File))

    If Attrib And FILE_ATTRIBUTE_READONLY Then
        SetFileAttributes StrPtr(File), Attrib And Not FILE_ATTRIBUTE_READONLY
    End If
    
    lr = DeleteFileW(StrPtr(File))  'not 0 - success
    If lr = 0 Then
        Debug.Print "Error " & Err.LastDllError & " when deleting file: " & File
    End If

    ' -> в случае неудачи, попытка получения прав NTFS + смена владельца на локальную группу "Администраторы"
    ' цель и аргументы передаются только для включение в отчет

    isDeleted = Not FileExists(File)
    
    If Not isDeleted Then
        TryUnlock File
        SetFileAttributes StrPtr(File), FILE_ATTRIBUTE_NORMAL
        
        Call DeleteFileW(StrPtr(File))
        'lr = Err.LastDllError
        
        isDeleted = Not FileExists(File)
    End If
    
    If isDeleted Then SHChangeNotify SHCNE_DELETE, SHCNF_PATH Or SHCNF_FLUSHNOWAIT, StrPtr(File), ByVal 0&

    DeleteFileForce = isDeleted 'lr

    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "DeleteFileForce - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DeleteFileEx", "File:", File
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Sub TryUnlock(ByVal File As String)  'получения прав NTFS + смена владельца на локальную группу "Администраторы"
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "TryUnlock - Begin", "File: " & File
    
    Dim TakeOwn As String
    Dim Icacls As String
    Dim DosName As String
    
    DosName = GetDOSFilename(File)
    If Len(DosName) <> 0 Then File = DosName
    
    If Not OSver.bIsVistaOrLater Then Exit Sub
    
    If OSver.Bitness = "x64" And FolderExists(sWinDir & "\sysnative") Then
        TakeOwn = EnvironW("%SystemRoot%") & "\Sysnative\takeown.exe"
        Icacls = EnvironW("%SystemRoot%") & "\Sysnative\icacls.exe"
    Else
        TakeOwn = EnvironW("%SystemRoot%") & "\System32\takeown.exe"
        Icacls = EnvironW("%SystemRoot%") & "\System32\icacls.exe"
    End If
    
    If FileExists(TakeOwn) Then
        Proc.ProcessRun TakeOwn, "/F " & """" & File & """" & " /A", , 0
        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
            Proc.ProcessClose , , True
        End If
    End If
    
    If FileExists(Icacls) Then
        Proc.ProcessRun Icacls, """" & File & """" & " /grant:r *S-1-1-0:F", , 0
        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
            Proc.ProcessClose , , True
        End If
        
        Proc.ProcessRun Icacls, """" & File & """" & " /grant:r *S-1-5-32-544:F", , 0
        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
            Proc.ProcessClose , , True
        End If
    
        If 0 <> Len(OSInfo.SID_CurrentProcess) Then
            Proc.ProcessRun Icacls, """" & File & """" & " /grant:r *" & OSInfo.SID_CurrentProcess & ":F", , 0
        Else
            Proc.ProcessRun Icacls, """" & File & """" & " /grant:r """ & EnvironW("%UserName%") & """:F", , 0
        End If
        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
            Proc.ProcessClose , , True
        End If
    End If
    
    AppendErrorLogCustom "TryUnlock - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "TryUnlock", "File:", File
    If inIDE Then Stop: Resume Next
End Sub

Public Function AppPath(Optional bGetFullPath As Boolean) As String
    On Error GoTo ErrorHandler

    Static ProcPathFull  As String
    Static ProcPathShort As String
    Dim ProcPath As String
    Dim Cnt      As Long
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
        AppPath = GetDOSFilename(App.Path, bReverse:=True)
        'bGetFullPath does not supported in IDE
        Exit Function
    End If

    hProc = GetModuleHandle(0&)
    If hProc < 0 Then hProc = 0

    ProcPath = String$(MAX_PATH, vbNullChar)
    Cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath)) 'hproc can be 0 (mean - current process)
    
    If Cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
        ProcPath = Space$(MAX_PATH_W)
        Cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
    End If
    
    If Cnt = 0 Then                          'clear path
        ProcPath = App.Path
    Else
        ProcPath = Left$(ProcPath, Cnt)
        If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = sWinDir & Mid$(ProcPath, 12)
        If "\??\" = Left$(ProcPath, 4) Then ProcPath = Mid$(ProcPath, 5)
        
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
    Dim Cnt      As Long
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
        AppExeName = App.EXEName & IIf(WithExtension, ".exe", "")
        Exit Function
    End If

    hProc = GetModuleHandle(0&)
    If hProc < 0 Then hProc = 0

    ProcPath = String$(MAX_PATH, vbNullChar)
    Cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath)) 'hproc can be 0 (mean - current process)
    
    If Cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
        ProcPath = Space$(MAX_PATH_W)
        Cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
    End If
    
    If Cnt = 0 Then                          'clear path
        ProcPath = App.EXEName & IIf(WithExtension, ".exe", "")
    Else
        ProcPath = Left$(ProcPath, Cnt)
        
        pos = InStrRev(ProcPath, "\")
        If pos <> 0 Then ProcPath = Mid$(ProcPath, pos + 1)
        
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


Function ParseCommandLine(Line As String, argc As Long, argv() As String) As Boolean
  On Error GoTo ErrorHandler
  Dim Lex$(), nL&, nA&, Unit$, St$
  St = Line
  If Len(St) > 0 Then ParseCommandLine = True
  Lex = Split(St) 'Разбиваем по пробелам на лексемы для анализа знаков
  ReDim argv(0 To UBound(Lex) + 1) As String 'Определяем выходной массив до максимально возможного числа параметров
  argv(0) = AppPath(True)
  If Len(St) <> 0 Then
    Do While nL <= UBound(Lex)
      Unit = Lex(nL) 'Записысаем текущую лексему как начало нового аргумента
      If Len(Unit) <> 0 Then 'Защита от двойных пробелов между аргументами
        'если в лексеме найдена кавычка или непарное их число, то начинаем процесс "квотирования"
        If (Len(Lex(nL)) - Len(Replace$(Lex(nL), """", ""))) Mod 2 = 1 Then
          Do
            nL = nL + 1
            If nL > UBound(Lex) Then Exit Do 'Если не дождались завершающей кавычки, а больше лексем нет
            Unit = Unit & " " & Lex(nL) 'дополняем соседней лексемой
          ' аргумент должен завершаться 1 или непарным числом кавычек лексемы со всеми прилягающими к ней справа символами (кроме знака пробела)
          Loop Until (Len(Lex(nL)) - Len(Replace$(Lex(nL), """", ""))) Mod 2 = 1
        End If
        Unit = Replace$(Unit, """", "") 'Удаляем кавычки
        nA = nA + 1 'Счетчик кол-ва выходных аргументов
        argv(nA) = Unit
      End If
      nL = nL + 1 'Счетчик текущей лексемы
    Loop
  End If
  ReDim Preserve argv(0 To nA) ' урезаем массив до реального числа аргументов
  argc = nA
  Exit Function
ErrorHandler:
  ErrorMsg Err, "Parser.ParseCommandLine", "CmdLine:", Line
  If inIDE Then Stop: Resume Next
End Function

'Delete File with unlock access rights on failure. Return non 0 on success.
Public Function DeleteFileWEx(lpSTR As Long, Optional ForceDeleteMicrosoft As Boolean) As Long
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "DeleteFileWEx - Begin"

    Dim iAttr As Long, lr As Long, sExt As String
    Dim Redirect As Boolean, bOldStatus As Boolean

    Dim FileName$
    FileName = String$(lstrlen(lpSTR), vbNullChar)
    If Len(FileName) <> 0 Then
        lstrcpy StrPtr(FileName), lpSTR
    Else
        Exit Function
    End If
    
    sExt = GetExtensionName(FileName)
    
    If Not ForceDeleteMicrosoft Then
        If Not StrInParamArray(sExt, ".txt", ".log", ".tmp") Then
            If IsMicrosoftFile(FileName) Then Exit Function
        End If
    End If
    
    Redirect = ToggleWow64FSRedirection(False, FileName, bOldStatus)
    
    iAttr = GetFileAttributes(lpSTR)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    
    If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes lpSTR, iAttr And Not FILE_ATTRIBUTE_READONLY
    lr = DeleteFileW(lpSTR)
    
    If lr <> 0 Then
        DeleteFileWEx = lr
        GoTo Finalize
    End If
    
    If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
        DeleteFileWEx = 1
        GoTo Finalize
    End If
    
    If Err.LastDllError = ERROR_ACCESS_DENIED Then
        TryUnlock FileName
        If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes lpSTR, iAttr And Not FILE_ATTRIBUTE_READONLY
        lr = DeleteFileW(lpSTR)
    End If
    
Finalize:
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "DeleteFileWEx - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.DeleteFileWEx", "File:", FileName
    If inIDE Then Stop: Resume Next
End Function

Public Function ConvertUnixTimeToLocalDate(Seconds As Long) As Date
    On Error GoTo ErrorHandler
    'Dim ftime               As FILETIME
    'Dim TimeZoneInfo(171)   As Byte
    Dim stime               As SYSTEMTIME
    Dim DateTime            As Date
    
    DateTime = DateAdd("s", Seconds, #1/1/1970#)    ' time_t -> Date
    VariantTimeToSystemTime DateTime, stime         ' Date -> SYSTEMTIME
    
    'SystemTimeToFileTime stime, ftime               ' SYSTEMTIME -> FILETIME
    'FileTimeToLocalFileTime ftime, ftime            ' учитываем смещение согласно текущим региональным настройкам в системе
    'FileTimeToSystemTime varptr(ftime), stime               ' FILETIME -> SYSTEMTIME
    
    'alternate:
    'GetTimeZoneInformation VarPtr(TimeZoneInfo(0))
    'SystemTimeToTzSpecificLocalTime VarPtr(TimeZoneInfo(0)), stime, stime
    
    SystemTimeToTzSpecificLocalTime 0&, stime, stime    'tz can be 0, if tz is current
    
    SystemTimeToVariantTime stime, DateTime         ' SYSTEMTIME -> Date
    ConvertUnixTimeToLocalDate = DateTime
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modUtils.ConvertUnixTimeToLocalDate", "sec:", Seconds
    If inIDE Then Stop: Resume Next
End Function

Public Function ConvertFileTimeToLocalDate(lpFileTime As Long) As Date
    On Error GoTo ErrorHandler
    
    Dim DateTime            As Date
    Dim stime               As SYSTEMTIME
    
    FileTimeToSystemTime ByVal lpFileTime, stime        ' FILETIME -> SYSTEMTIME
    SystemTimeToTzSpecificLocalTime 0&, stime, stime    ' tz can be 0, if tz is current
    SystemTimeToVariantTime stime, DateTime             ' SYSTEMTIME -> vtDate
    
    ConvertFileTimeToLocalDate = DateTime
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modUtils.ConvertFileTimeToLocalDate", "lpFileTime:", lpFileTime
    If inIDE Then Stop: Resume Next
End Function

Public Function BuildPath$(sPath$, sFile$)
    BuildPath = sPath & IIf(Right$(sPath, 1) = "\", vbNullString, "\") & sFile
End Function

Public Function GetWindowsVersion() As String    'Init by Form_load.
                                                 
    AppendErrorLogCustom "modUtils.GetWindowsVersion - Begin"
    
    ' Result -> sWinVersion (global)
    ' bIsWin64 (global)
    ' bIsWin32 (global)
    ' bIsWinVistaOrLater (global)
    ' OSver struct (global)
    ' OSInfo class (global)
    ' bIsWinVistaOrLater (global)
    ' bIsWin9x (global)
    ' bIsWinME (global)
    ' bIsWinNT (global)
    
    On Error GoTo ErrorHandler:
    
    Dim OVI As OSVERSIONINFOEX, sWinVer$
    
    Set OSInfo = New clsOSInfo
    
    bIsWin64 = (OSInfo.Bitness = "x64")
    bIsWOW64 = bIsWin64 ' mean VB6 app-s are always x32 bit.
    bIsWin32 = Not bIsWin64
    
    sWinSysDir = Environ("SystemRoot") & "\System32"

    'enable redirector
    If bIsWin64 Then ToggleWow64FSRedirection True
    
    OSver.bIsSafeBoot = OSInfo.IsSafeBoot
    OSver.BootMode = OSInfo.SafeBootMode

    OSver.bIsAdmin = OSInfo.IsElevated
    
    OVI.dwOSVersionInfoSize = LenB(OVI)
    GetVersionEx OVI
    
    OSver.Major = OVI.dwMajorVersion
    OSver.Minor = OVI.dwMinorVersion
    
    OSver.MajorMinor = OSInfo.MajorMinor
    OSver.SPVer = OSInfo.SPVer
    OSver.Build = OVI.dwBuildNumber
    OSver.Platform = OSInfo.Platform
    OSver.OSName = OSInfo.OSName
    OSver.Edition = OSInfo.Edition
    OSver.bIsWin64 = bIsWin64
    OSver.Bitness = OSInfo.Bitness

    With OVI
        bIsWinVistaOrLater = .dwMajorVersion >= 6
        OSver.bIsVistaOrLater = bIsWinVistaOrLater
        
        Select Case .dwPlatformId
            Case 0: GetWindowsVersion = "Detected: Windows 3.x running Win32s": Exit Function
            Case 1: bIsWin9x = True: bIsWinNT = False
            Case 2: bIsWinNT = True: bIsWin9x = False
        End Select
        
        If bIsWin9x Then
            If .dwMajorVersion = 4 Then
                If .dwMinorVersion = 90 Then 'Windows Millennium Edition
                    bIsWinME = True
                End If
            End If
        End If

        sWinVer = OSver.OSName & " " & OSver.Edition & " SP" & OSver.SPVer & " " & _
            "(Windows " & OSver.Platform & " " & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & ")"

    End With

    GetWindowsVersion = sWinVer

    AppendErrorLogCustom "modUtils.GetWindowsVersion - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetWindowsVersion"
    If inIDE Then Stop: Resume Next
End Function


'Ищет полный путь к файлу.
'Функция не гарантирует, что файл существует.
'Результатом могут быть даже служебные символы, если таковые были поданы на вход.
'Это конечно странно, но пока что не трогаю это поведение функции, чтобы ничего не сломать вверх по стеку вызовов.

'Public Function GetLongPath$(ByVal sFile$)
'    'sub applies to NT only, checked in ListRunningProcesses()
'    'attempt to find location of given file
'    On Error GoTo ErrorHandler:
'    AppendErrorLogCustom "GetLongPath - Begin", "File: " & sFile
'
'    Dim pos&
'
'    'evading parasites that put html or garbled data in
'    'O4 autorun entries :P
'    If InStr(sFile, "<") > 0 Or InStr(sFile, ">") > 0 Or _
'       InStr(sFile, "|") > 0 Or InStr(sFile, "*") > 0 Or _
'       InStr(sFile, "?") > 0 Then  'Or InStr(sFile, "/") > 0 Or InStr(sFile, ":") > 0 Then
'        GetLongPath = sFile ' ???? //TODO
'        Exit Function
'    End If
'
'    If InStr(sFile, "/") <> 0 Then sFile = Replace$(sFile, "/", "\")
'
'    If Left$(sFile, 1) = """" Then
'        pos = InStr(2, sFile, """")
'        If pos <> 0 Then
'            sFile = Mid$(sFile, 2, pos - 2)
'        Else
'            sFile = Mid$(sFile, 2)
'        End If
'    End If
'
'    GetLongPath = FindOnPath(sFile)
'    If 0 <> Len(GetLongPath) Then Exit Function
'
'    pos = InStrRev(sFile, ".exe", -1, 1)
'    If 0 <> pos And pos <> Len(sFile) - 3 Then
'        sFile = Left$(sFile, pos + 3)
'        GetLongPath = FindOnPath(sFile)
'        If 0 <> Len(GetLongPath) Then Exit Function
'    End If
'
'    'If sFile = "[System Process]" Or sFile = "System" Then
'    '    GetLongPath = sFile
'    '    Exit Function
'    'End If
'
'    If InStr(sFile, "\") > 0 Then
'        'filename is already full path
'        GetLongPath = sFile
'        Exit Function
'    End If
'
''    'check if file is self
''    If LCase$(sFile) = LCase$(AppExeName(True)) Then
''        GetLongPath = AppPath() & IIf(Right$(AppPath(), 1) = "\", vbNullString, "\") & sFile
''        Exit Function
''    End If
'
'    Dim hKey, sData$, i&, sDummy$, sProgramFiles$
'    'check App Paths regkey
'    sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" & sFile, vbNullString)
'    If sData <> vbNullString Then
'        GetLongPath = sData
'        Exit Function
'    End If
'
'    'check own folder
'    If FileExists(BuildPath(AppPath(), sFile)) Then
'        GetLongPath = BuildPath(AppPath(), sFile)
'        Exit Function
'    End If
'
'    'check windir\system32
'    If FileExists(sWinDir & "\system32\" & sFile) Then
'        GetLongPath = sWinDir & "\system32\" & sFile
'        Exit Function
'    End If
'
'    If OSver.Bitness = "x64" Then
'        If FileExists(sWinDir & "\syswow64\" & sFile) Then
'            GetLongPath = sWinDir & "\syswow64\" & sFile
'            Exit Function
'        End If
'    End If
'
'    'check windir
'    If FileExists(sWinDir & "\" & sFile) Then
'        GetLongPath = sWinDir & "\" & sFile
'        Exit Function
'    End If
'
'    'check windir\system
'    If FileExists(sWinDir & "\system\" & sFile) Then
'        GetLongPath = sWinDir & "\system\" & sFile
'        Exit Function
'    End If
'
'    If InStr(sFile, ".") > 0 Then
'        'prog.exe -> prog
'        sDummy = Left$(sFile, InStr(sFile, ".") - 1)
'        'sProgramFiles = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
'        sProgramFiles = PF_64
'
'        'check x:\program files\prog\prog.exe
'        If FileExists(sProgramFiles & "\" & sDummy & "\" & sFile) Then
'            GetLongPath = sProgramFiles & "\" & sDummy & "\" & sFile
'            Exit Function
'        End If
'
'        'check c:\prog\prog.exe
'        If FileExists(SysDisk & "\" & sDummy & "\" & sFile) Then
'            GetLongPath = SysDisk & "\" & sDummy & "\" & sFile
'            Exit Function
'        End If
'
'        'check x:\program files\prog32\prog.exe
'        If FileExists(sProgramFiles & "\" & sDummy & "32\" & sFile) Then
'            GetLongPath = sProgramFiles & "\" & sDummy & "32\" & sFile
'            Exit Function
'        End If
'        If FileExists(sProgramFiles & "\" & sDummy & "16\" & sFile) Then
'            GetLongPath = sProgramFiles & "\" & sDummy & "16\" & sFile
'            Exit Function
'        End If
'
'        'check c:\prog32\prog.exe
'        If FileExists(SysDisk & "\" & sDummy & "32\" & sFile) Then
'            GetLongPath = SysDisk & "\" & sDummy & "32\" & sFile
'            Exit Function
'        End If
'        If FileExists(SysDisk & "\" & sDummy & "16\" & sFile) Then
'            GetLongPath = SysDisk & "\" & sDummy & "16\" & sFile
'            Exit Function
'        End If
'
'        If Right$(sDummy, 2) = "32" Or Right$(sDummy, 2) = "16" Then
'            'asssuming sFile is prog32.exe,
'            'check x:\program files\prog\prog32.exe
'            sDummy = Left$(sDummy, Len(sDummy) - 2)
'            If FileExists(sProgramFiles & "\" & sDummy & "\" & sFile) Then
'                GetLongPath = sProgramFiles & "\" & sDummy & "\" & sFile
'                Exit Function
'            End If
'
'            'check c:\prog\prog32.exe
'            If DirW$(SysDisk & "\" & sDummy & "\" & sFile) Then
'                GetLongPath = SysDisk & "\" & sDummy & "\" & sFile
'                Exit Function
'            End If
'        End If
'    End If
'
'    'can't find it!
'    GetLongPath = "?:\?\" & sFile
'
'    AppendErrorLogCustom "GetLongPath - End"
'    Exit Function
'ErrorHandler:
'    ErrorMsg Err, "GetLongPath", sFile
'    If inIDE Then Stop: Resume Next
'End Function


'Public Function GetFileFromAutostart(sAutostart$, Optional bGetMD5 As Boolean = True) As String
'    Dim sDummy$
'    On Error GoTo ErrorHandler:
'    AppendErrorLogCustom "GetFileFromAutostart - Begin", "File: " & sAutostart
'
'    If InStr(sAutostart, "(file missing)") > 0 Then Exit Function
'
'    sDummy = sAutostart
'
'    'forms we can find the file in:
'    'c:\bla\bla.exe
'    'c:\bla.exe
'    'bla.exe
'    'bla
'    '
'    'also possible:
'    '* surrounding quotes
'    '* arguments (possibly files)
'
'    If Not FileExists(sDummy) Then
'      If Left$(sDummy, 1) = """" Then
'        'has quotes
'        'stripping like this also removes any
'        'arguments, so a path means it's finished
'        sDummy = Mid$(sDummy, 2)
'        sDummy = Left$(sDummy, InStr(sDummy, """") - 1)
'
'        If InStr(sDummy, "\") = 0 Then
'            'GoTo FindFullPath:
'            If InStr(sDummy, "\") = 0 Then
'                'no path - so search for file
'                sDummy = GetLongPath(sDummy)
'            End If
'        End If
'      End If
'    End If
'
'    If Not FileExists(sDummy) Then
'      If LCase$(Right$(sDummy, 4)) <> ".exe" And _
'       LCase$(Right$(sDummy, 4)) <> ".com" Then
'        'has arguments, or no extension
'        If InStr(sDummy, " ") = 0 Then
'            'only one word, so no extension
'            sDummy = GetLongPath(sDummy & ".exe")
'            If InStr(sDummy, "\") = 0 Then
'                sDummy = GetLongPath(sDummy & ".com")
'            End If
'        Else
'            'multiple words, the first is the program
'            If FileExists(Left$(sDummy, InStr(sDummy, " ") - 1)) Then
'                sDummy = Left$(sDummy, InStr(sDummy, " ") - 1)
'                sDummy = GetLongPath(sDummy)
'            Else
'                sDummy = Left$(sDummy, InStrRev(sDummy, " ") - 1)
'                sDummy = GetLongPath(sDummy)
'            End If
'        End If
'      End If
'    End If
'
'    If FileExists(sDummy) Then
'        If bGetMD5 Then
'            GetFileFromAutostart = GetFileMD5(sDummy)
'        Else
'            GetFileFromAutostart = sDummy
'        End If
'    End If
'
'    AppendErrorLogCustom "GetFileFromAutostart - End"
'    Exit Function
'ErrorHandler:
'    ErrorMsg Err, "modMD5_GetFileFromAutostart", sAutostart
'    If inIDE Then Stop: Resume Next
'End Function


