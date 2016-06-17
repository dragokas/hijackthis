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
    x As Long
    y As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExW" (ByVal lpFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadString Lib "user32.dll" Alias "LoadStringW" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As Long, ByVal nBufferMax As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStrDest As Long, ByVal lpStrSrc As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long

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
        If lpPrevWndProc = 0 Then lpPrevWndProc = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf WndProc)
    Else
        If lpPrevWndProc Then SetWindowLong frmMain.hwnd, GWL_WNDPROC, lpPrevWndProc: lpPrevWndProc = 0
    End If
End Sub

Function IsMouseWithin(hwnd As Long) As Boolean
    Dim R As RECT
    Dim p As POINTAPI
    If GetWindowRect(hwnd, R) Then
        If GetCursorPos(p) Then
            IsMouseWithin = PtInRect(R, p.x, p.y)
        End If
    End If
End Function

Private Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
    Case WM_MOUSEWHEEL
        If Not IsMouseWithin(frmMain.hwnd) Then Exit Function ' mouse is outside the form
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
        WndProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function

' affected by wow64
Public Function OpenFileDialog(Optional sTitle As String, Optional InitDir As String, Optional hwnd As Long) As String
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
        .hwndOwner = IIf(hwnd = 0, frmMain.hwnd, hwnd)
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
    ErrorMsg err, "OpenFileDialog"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetStringFromBinary(Optional ByVal sFile As String, Optional ByVal nID As Long, Optional ByVal FileAndIDHybrid As String) As String
    On Error GoTo ErrorHandler:

    Dim hModule As Long
    Dim nSize As Long
    Dim sBuf As String
    Dim pos As Long
    
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
            If IsNumeric(sBuf) And 0 <> Len(sBuf) Then nID = Val(sBuf) Else Exit Function
        End If
    End If
    
    If 0 = Len(sFile) Then Exit Function
    
    If Not FileExists(sFile) Then
        sFile = GetLongPath(sFile)
        If Not FileExists(sFile) Then Exit Function
    End If
    
    sBuf = String$(160, vbNullChar)
    
    ToggleWow64FSRedirection False, sFile
    hModule = LoadLibraryEx(StrPtr(sFile), 0&, LOAD_LIBRARY_AS_DATAFILE)
    
    If hModule Then
        nSize = LoadString(hModule, Abs(nID), StrPtr(sBuf), LenB(sBuf))
        If nSize > 0 Then
            GetStringFromBinary = TrimNull(Left$(sBuf, nSize))
        End If
        FreeLibrary hModule
    End If
    ToggleWow64FSRedirection True
    Exit Function
ErrorHandler:
    ErrorMsg err, "GetStringFromBinary", sFile
    If inIDE Then Stop: Resume Next
End Function

'Public Function NormalizePath$(sFile$)
'
'    Dim sBegin$, sValue$, sNext$
'    Dim EnvVar As String
'    Dim RealEnvVar As String
'
'    If False Then
''        Dim EnvRegExp As RegExp
''        Dim ObjMatch As Match
''        Dim ObjMatches As MatchCollection
''        'Dim EnvVar As String
''
''        Set EnvRegExp = New RegExp
''        EnvRegExp.Pattern = "%[\w_-]+%"
''        EnvRegExp.IgnoreCase = True
''        EnvRegExp.Global = True
''
''        If EnvRegExp.Test(sFile) = True Then
''            Set ObjMatches = EnvRegExp.Execute(sFile)
''            For Each ObjMatch In ObjMatches
''                EnvVar = Replace$(ObjMatch.value, "%", "", , , vbTextCompare)
''                If Len(Environ$(EnvVar)) > 0 Then
''                    sFile = Replace$(sFile, ObjMatch.value, Environ$(EnvVar), , , vbTextCompare)
''                End If
''            Next
''        End If
'    End If
    
'    'If False Then
'    sBegin = 1
'    Do
'        sValue = InStr(sBegin, sFile, "%", vbTextCompare)
'        If sValue = 0 Or sValue = Len(sFile) Or sBegin > Len(sFile) Then
'            Exit Do
'        End If
'
'        sBegin = sValue + 1
'        sNext = InStr(sBegin + 1, sFile, "%", vbTextCompare)
'        If sNext = 0 Or sNext > Len(sFile) Or sBegin > Len(sFile) Then
'            Exit Do
'        End If
'
'        EnvVar = mid$(sFile, sValue, sNext - sValue + 1)
'        RealEnvVar = mid$(sFile, sValue + 1, sNext - sValue - 1)
'
'        If Len(Environ$(RealEnvVar)) > 0 Then
'            sFile = Replace$(sFile, EnvVar, Environ$(RealEnvVar), sValue, sNext - sValue + 1, vbTextCompare)
'            sBegin = sNext + 1 + Len(Environ$(RealEnvVar)) - Len(EnvVar)
'        Else
'            sBegin = sNext + 1
'        End If
'
'    Loop While True
'    'End If
'
'    NormalizePath = sFile
'    NormalizePath = EnvironW(sFile)
'End Function

Public Sub GetBrowsersInfo()
    With BROWSERS
        .Edge.Version = GetEdgeVersion()
        .IE.Version = GetMSIEVersion()
        .Chrome.Version = GetChromeVersion()
        .Firefox.Version = GetFirefoxVersion()
        .Opera.Version = GetOperaVersion()
    End With
End Sub

Public Function GetEdgeVersion() As String
    Dim EdgePath$
    EdgePath = sWinDir & "\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe"
    If FileExists(EdgePath) Then GetEdgeVersion = GetFilePropVersion(EdgePath)
End Function

Public Function GetChromeVersion() As String
    Dim sVer$
    sVer = RegGetString(HKEY_LOCAL_MACHINE, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv")
    'not found try current user - win7(x86)
    If Len(sVer) = 0 Then
        sVer = RegGetString(HKEY_CURRENT_USER, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv")
    End If
    If Len(sVer) = 0 Then 'Wow6432Node
        sVer = RegGetString(HKEY_LOCAL_MACHINE, "Software\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv", True)
    End If

    GetChromeVersion = sVer
End Function

Public Function GetFirefoxVersion() As String
    Dim sVer$, sPath$
    
    sPath = RegGetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe", vbNullString)
    If Len(sPath) <> 0 Then
        sVer = GetFilePropVersion(sPath)
        If Len(sVer) = 0 Then sVer = RegGetString(HKEY_LOCAL_MACHINE, "Software\Mozilla\Mozilla Firefox", "CurrentVersion")
    End If
    
    GetFirefoxVersion = sVer
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
    Dim sOperaPath$
    sOperaPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\Opera.exe", vbNullString)

    If Len(sOperaPath) <> 0 Then
        sOperaPath = UnQuote(sOperaPath)
        If FileExists(sOperaPath) Then GetOperaVersion = GetFilePropVersion(sOperaPath)
    End If
End Function

Public Function GetMSIEVersion() As String
    On Error GoTo ErrorHandler:
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
            sMSIEVersion = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "Version")
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
    Exit Function
    
ErrorHandler:
    ErrorMsg err, "GetMSIEVersion"
    If inIDE Then Stop: Resume Next
End Function

Public Function IsSignPresent(FileName As String) As Boolean
    ' &H3C -> PE_Header offset
    ' PE_Header offset + &H18 = Optional_PE_Header
    ' PE_Header offset + &H78 = Data_Directories offset
    ' Data_Directories offset + &H20 = SecurityDir -> Address (dword), Size (dword) for digital signature.
    
    On Error GoTo ErrorHandler:
    
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
    
    IsSignPresent = (SignAddress <> 0)
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "modUtils_IsSignPresent", "File:", FileName
    If inIDE Then Stop: Resume Next
End Function

' Особая процедура удаления файла (с разблокировкой NTFS привилегий)
Public Function DeleteFileForce(File As String) As Long
    On Error GoTo ErrorHandler:
    
    Const FILE_ATTRIBUTE_NORMAL     As Long = &H80&
    Const FILE_ATTRIBUTE_READONLY   As Long = 1&
    
    Dim lr          As Long
    Dim Attrib      As Long
    Dim isDeleted   As Boolean

    If IsMicrosoftFile(File) Then Exit Function

    ToggleWow64FSRedirection False, File

    Attrib = GetFileAttributes(StrPtr(File))

    If Attrib And FILE_ATTRIBUTE_READONLY Then
        SetFileAttributes StrPtr(File), Attrib And Not FILE_ATTRIBUTE_READONLY
    End If
    
    lr = DeleteFileW(StrPtr(File))  'not 0 - success
    If lr = 0 Then
        Debug.Print "Error " & err.LastDllError & " when deleting file: " & File
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

    ToggleWow64FSRedirection True
    Exit Function
ErrorHandler:
    ErrorMsg err, "DeleteFileEx", "File:", File
    If inIDE Then Stop: Resume Next
End Function

Sub TryUnlock(ByVal File As String)  'получения прав NTFS + смена владельца на локальную группу "Администраторы"
    On Error GoTo ErrorHandler:
    Dim TakeOwn As String
    Dim Icacls As String
    Dim DosName As String
    
    DosName = GetDOSFilename(File)
    If Len(DosName) <> 0 Then File = DosName
    
    If Not OSVer.bIsVistaOrLater Then Exit Sub
    
    If OSVer.Bitness = "x64" And FolderExists(sWinDir & "\sysnative") Then
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
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "TryUnlock", "File:", File
    If inIDE Then Stop: Resume Next
End Sub

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
        AppPath = GetDOSFilename(App.Path, bReverse:=True)
        'bGetFullPath does not supported in IDE
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
        ProcPath = App.Path
    Else
        ProcPath = Left$(ProcPath, cnt)
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
    ErrorMsg err, "Parser.AppPath", "ProcPath:", ProcPath
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
        AppExeName = App.EXEName & IIf(WithExtension, ".exe", "")
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
        ProcPath = App.EXEName & IIf(WithExtension, ".exe", "")
    Else
        ProcPath = Left$(ProcPath, cnt)
        
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
    ErrorMsg err, "Parser.AppExeName", "ProcPath:", ProcPath
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
  ErrorMsg err, "Parser.ParseCommandLine", "CmdLine:", Line
  If inIDE Then Stop: Resume Next
End Function

'Delete File with unlock access rights on failure. Return non 0 on success.
Public Function DeleteFileWEx(lpStr As Long) As Long
    On Error GoTo ErrorHandler:

    Dim iAttr As Long, lr As Long

    Dim FileName$
    FileName = String$(lstrlen(lpStr), vbNullChar)
    If Len(FileName) <> 0 Then
        lstrcpy StrPtr(FileName), lpStr
    Else
        Exit Function
    End If
    
    If IsMicrosoftFile(FileName) Then Exit Function
    
    ToggleWow64FSRedirection False, FileName
    
    iAttr = GetFileAttributes(lpStr)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    
    If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes lpStr, iAttr And Not FILE_ATTRIBUTE_READONLY
    lr = DeleteFileW(lpStr)
    
    If lr <> 0 Then
        DeleteFileWEx = lr
        ToggleWow64FSRedirection True, FileName
        Exit Function
    End If
    
    If err.LastDllError = ERROR_FILE_NOT_FOUND Then
        DeleteFileWEx = 1
        ToggleWow64FSRedirection True, FileName
        Exit Function
    End If
    
    If err.LastDllError = ERROR_ACCESS_DENIED Then
        TryUnlock FileName
        If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes lpStr, iAttr And Not FILE_ATTRIBUTE_READONLY
        lr = DeleteFileW(lpStr)
    End If
    ToggleWow64FSRedirection True, FileName
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "Parser.DeleteFileWEx", "File:", FileName
    If inIDE Then Stop: Resume Next
End Function


'    Dim hLib     As Long
'    Dim hModule  As Long
'    Dim hHandle As Long
'    hLib = LoadLibrary(StrPtr("c:\Windows\SysWOW64\ntdll.dll"))
'    If hLib <> 0 Then
'        hModule = GetProcAddress(hLib, "A_SHAFinal")
'        hHandle = GetModuleHandle(StrPtr("c:\Windows\SysWOW64\ntdll.dll"))
'        Debug.Print "Addr func: " & hModule
'        Debug.Print "Addr Base: " & hHandle
'        Debug.Print "RVA: 0x" & Hex(hModule - hHandle)
'    End If
'    Stop
