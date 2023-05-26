Attribute VB_Name = "Common"
Option Explicit
Private Type MSGBOXPARAMS
cbSize As Long
hWndOwner As Long
hInstance As Long
lpszText As Long
lpszCaption As Long
dwStyle As Long
lpszIcon As Long
dwContextHelpID As Long
lpfnMsgBoxCallback As Long
dwLanguageId As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type BITMAP
BMType As Long
BMWidth As Long
BMHeight As Long
BMWidthBytes As Long
BMPlanes As Integer
BMBitsPixel As Integer
BMBits As Long
End Type
Private Type SAFEARRAYBOUND
cElements As Long
lLbound As Long
End Type
Private Type SAFEARRAY1D
cDims As Integer
fFeatures As Integer
cbElements As Long
cLocks As Long
pvData As Long
Bounds As SAFEARRAYBOUND
End Type
Private Type PICTDESC
cbSizeOfStruct As Long
PicType As Long
hImage As Long
XExt As Long
YExt As Long
End Type
Private Type CLSID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type
Private Const MAX_PATH As Long = 260
Private Type WIN32_FIND_DATA
dwFileAttributes As Long
FTCreationTime As FILETIME
FTLastAccessTime As FILETIME
FTLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
lpszFileName(0 To ((MAX_PATH * 2) - 1)) As Byte
lpszAlternateFileName(0 To ((14 * 2) - 1)) As Byte
End Type
Private Type WIN32_FILE_ATTRIBUTE_DATA
dwFileAttributes As Long
FTCreationTime As FILETIME
FTLastAccessTime As FILETIME
FTLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
End Type
Private Type VS_FIXEDFILEINFO
dwSignature As Long
dwStrucVersionLo As Integer
dwStrucVersionHi As Integer
dwFileVersionMSLo As Integer
dwFileVersionMSHi As Integer
dwFileVersionLSLo As Integer
dwFileVersionLSHi As Integer
dwProductVersionMSLo As Integer
dwProductVersionMSHi As Integer
dwProductVersionLSLo As Integer
dwProductVersionLSHi As Integer
dwFileFlagsMask As Long
dwFileFlags As Long
dwFileOS As Long
dwFileType As Long
dwFileSubtype As Long
dwFileDateMS As Long
dwFileDateLS As Long
End Type
Private Type MONITORINFO
cbSize As Long
RCMonitor As RECT
RCWork As RECT
dwFlags As Long
End Type
Private Type FLASHWINFO
cbSize As Long
hWnd As Long
dwFlags As Long
uCount As Long
dwTimeout As Long
End Type
Private Const LF_FACESIZE As Long = 32
Private Const DEFAULT_QUALITY As Long = 0
Private Type LOGFONT
LFHeight As Long
LFWidth As Long
LFEscapement As Long
LFOrientation As Long
LFWeight As Long
LFItalic As Byte
LFUnderline As Byte
LFStrikeOut As Byte
LFCharset As Byte
LFOutPrecision As Byte
LFClipPrecision As Byte
LFQuality As Byte
LFPitchAndFamily As Byte
LFFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Var() As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectW" (ByRef lpMsgBoxParams As MSGBOXPARAMS) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributesEx Lib "kernel32" Alias "GetFileAttributesExW" (ByVal lpFileName As Long, ByVal fInfoLevelId As Long, ByVal lpFileInformation As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByVal lpFileTime As Long, ByVal lpLocalFileTime As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByVal lpFileTime As Long, ByVal lpSystemTime As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As Long, ByRef lpMI As MONITORINFO) As Long
Private Declare Function GetVolumePathName Lib "kernel32" Alias "GetVolumePathNameW" (ByVal lpFileName As Long, ByVal lpVolumePathName As Long, ByVal cch As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationW" (ByVal lpRootPathName As Long, ByVal lpVolumeNameBuffer As Long, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As Long, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryW" (ByVal lpPathName As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version" Alias "GetFileVersionInfoW" (ByVal lpFileName As Long, ByVal dwHandle As Long, ByVal dwLen As Long, ByVal lpData As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version" Alias "GetFileVersionInfoSizeW" (ByVal lpFileName As Long, ByVal lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version" Alias "VerQueryValueW" (ByVal lpBlock As Long, ByVal lpSubBlock As Long, ByRef lplpBuffer As Long, ByRef puLen As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Private Declare Function PathGetArgs Lib "shlwapi" Alias "PathGetArgsW" (ByVal lpszPath As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pbString As Long, ByVal pszStrPtr As Long) As Long
Private Declare Function VarDecFromI8 Lib "oleaut32" (ByVal LoDWord As Long, ByVal HiDWord As Long, ByRef pDecOut As Variant) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetSystemWindowsDirectory Lib "kernel32" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryW" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByVal lpdwProcessId As Long) As Long
Private Declare Function FlashWindowEx Lib "user32" (ByRef pFWI As FLASHWINFO) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GdiAlphaBlend Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BlendFunc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hPal As Long, ByRef RGBResult As Long) As Long
Private Declare Function OleLoadPicture Lib "oleaut32" (ByVal pStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As Any, ByRef pIPicture As IPicture) As Long
Private Declare Function OleLoadPicturePath Lib "oleaut32" (ByVal lpszPath As Long, ByVal pUnkCaller As Long, ByVal dwReserved As Long, ByVal ClrReserved As OLE_COLOR, ByRef riid As CLSID, ByRef pIPicture As IPicture) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (ByRef pPictDesc As PICTDESC, ByRef riid As Any, ByVal fPictureOwnsHandle As Long, ByRef pIPicture As IPicture) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef pStream As IUnknown) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

' (VB-Overwrite)
Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String) As VbMsgBoxResult
Dim MSGBOXP As MSGBOXPARAMS
With MSGBOXP
.cbSize = LenB(MSGBOXP)
If (Buttons And vbSystemModal) = 0 Then
    If Not Screen.ActiveForm Is Nothing Then
        .hWndOwner = Screen.ActiveForm.hWnd
    Else
        .hWndOwner = GetActiveWindow()
    End If
Else
    .hWndOwner = GetForegroundWindow()
End If
.hInstance = App.hInstance
.lpszText = StrPtr(Prompt)
If Title = vbNullString Then Title = App.Title
.lpszCaption = StrPtr(Title)
.dwStyle = Buttons
End With
MsgBox = MessageBoxIndirect(MSGBOXP)
End Function

' (VB-Overwrite)
Public Sub SendKeys(ByRef Text As String, Optional ByRef Wait As Boolean)
CreateObject("WScript.Shell").SendKeys Text, Wait
End Sub

' (VB-Overwrite)
Public Function GetAttr(ByVal PathName As String) As VbFileAttribute
Const INVALID_FILE_ATTRIBUTES As Long = (-1)
Const FILE_ATTRIBUTE_NORMAL As Long = &H80
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
Dim dwAttributes As Long
dwAttributes = GetFileAttributes(StrPtr("\\?\" & PathName))
If dwAttributes = INVALID_FILE_ATTRIBUTES Then
    Err.Raise 53
ElseIf dwAttributes = FILE_ATTRIBUTE_NORMAL Then
    GetAttr = vbNormal
Else
    GetAttr = dwAttributes
End If
End Function

' (VB-Overwrite)
Public Sub SetAttr(ByVal PathName As String, ByVal Attributes As VbFileAttribute)
Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Dim dwAttributes As Long
If Attributes = vbNormal Then
    dwAttributes = FILE_ATTRIBUTE_NORMAL
Else
    If (Attributes And (vbVolume Or vbDirectory Or vbAlias)) <> 0 Then Err.Raise 5
    dwAttributes = Attributes
End If
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
If SetFileAttributes(StrPtr("\\?\" & PathName), dwAttributes) = 0 Then Err.Raise 53
End Sub

' (VB-Overwrite)
Public Function Dir(Optional ByVal PathMask As String, Optional ByVal Attributes As VbFileAttribute = vbNormal) As String
Const INVALID_HANDLE_VALUE As Long = (-1)
Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Static hFindFile As Long, AttributesCache As VbFileAttribute
If Attributes = vbVolume Then ' Exact match
    ' If any other attribute is specified, vbVolume is ignored.
    If hFindFile <> 0 Then
        FindClose hFindFile
        hFindFile = 0
    End If
    Dim VolumePathBuffer As String, VolumeNameBuffer As String
    If Len(PathMask) = 0 Then
        VolumeNameBuffer = String$(MAX_PATH, vbNullChar)
        If GetVolumeInformation(0, StrPtr(VolumeNameBuffer), Len(VolumeNameBuffer), ByVal 0, ByVal 0, ByVal 0, 0, 0) <> 0 Then Dir = Left$(VolumeNameBuffer, InStr(VolumeNameBuffer, vbNullChar) - 1)
    Else
        VolumePathBuffer = String$(MAX_PATH, vbNullChar)
        If Left$(PathMask, 2) = "\\" Then PathMask = "UNC\" & Mid$(PathMask, 3)
        If GetVolumePathName(StrPtr("\\?\" & PathMask), StrPtr(VolumePathBuffer), Len(VolumePathBuffer)) <> 0 Then
            VolumePathBuffer = Left$(VolumePathBuffer, InStr(VolumePathBuffer, vbNullChar) - 1)
            VolumeNameBuffer = String$(MAX_PATH, vbNullChar)
            If GetVolumeInformation(StrPtr(VolumePathBuffer), StrPtr(VolumeNameBuffer), Len(VolumeNameBuffer), ByVal 0, ByVal 0, ByVal 0, 0, 0) <> 0 Then Dir = Left$(VolumeNameBuffer, InStr(VolumeNameBuffer, vbNullChar) - 1)
        End If
    End If
Else
    Dim FD As WIN32_FIND_DATA, dwMask As Long
    If Len(PathMask) = 0 Then
        If hFindFile <> 0 Then
            If FindNextFile(hFindFile, FD) = 0 Then
                FindClose hFindFile
                hFindFile = 0
                Exit Function
            End If
        Else
            Err.Raise 5
            Exit Function
        End If
    Else
        If hFindFile <> 0 Then
            FindClose hFindFile
            hFindFile = 0
        End If
        Select Case Right$(PathMask, 1)
            Case "\", ":", "/"
                PathMask = PathMask & "*.*"
        End Select
        AttributesCache = Attributes
        If Left$(PathMask, 2) = "\\" Then PathMask = "UNC\" & Mid$(PathMask, 3)
        hFindFile = FindFirstFile(StrPtr("\\?\" & PathMask), FD)
        If hFindFile = INVALID_HANDLE_VALUE Then
            hFindFile = 0
            If Err.LastDllError > 12 Then Err.Raise 52
            Exit Function
        End If
    End If
    Do
        If FD.dwFileAttributes = FILE_ATTRIBUTE_NORMAL Then
            dwMask = 0 ' Found
        Else
            dwMask = FD.dwFileAttributes And (Not AttributesCache) And &H16
        End If
        If dwMask = 0 Then
            Dir = Left$(FD.lpszFileName(), InStr(FD.lpszFileName(), vbNullChar) - 1)
            If FD.dwFileAttributes And vbDirectory Then
                If Dir <> "." And Dir <> ".." Then Exit Do ' Exclude self and relative path aliases
            Else
                Exit Do
            End If
        End If
        If FindNextFile(hFindFile, FD) = 0 Then
            FindClose hFindFile
            hFindFile = 0
            Exit Do
        End If
    Loop
End If
End Function

' (VB-Overwrite)
Public Sub MkDir(ByVal PathName As String)
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
If CreateDirectory(StrPtr("\\?\" & PathName), 0) = 0 Then
    Const ERROR_PATH_NOT_FOUND As Long = 3
    If Err.LastDllError = ERROR_PATH_NOT_FOUND Then
        Err.Raise 76
    Else
        Err.Raise 75
    End If
End If
End Sub

' (VB-Overwrite)
Public Sub RmDir(ByVal PathName As String)
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
If RemoveDirectory(StrPtr("\\?\" & PathName)) = 0 Then
    Const ERROR_FILE_NOT_FOUND As Long = 2
    If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
        Err.Raise 76
    Else
        Err.Raise 75
    End If
End If
End Sub

' (VB-Overwrite)
Public Function FileLen(ByVal PathName As String) As Variant
Dim FAD As WIN32_FILE_ATTRIBUTE_DATA
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
If GetFileAttributesEx(StrPtr("\\?\" & PathName), 0, VarPtr(FAD)) <> 0 Then
    FileLen = CDec(0)
    VarDecFromI8 FAD.nFileSizeLow, FAD.nFileSizeHigh, FileLen
Else
    Err.Raise Number:=53, Description:="File not found: '" & PathName & "'"
End If
End Function

' (VB-Overwrite)
Public Function FileDateTime(ByVal PathName As String) As Date
Dim FAD As WIN32_FILE_ATTRIBUTE_DATA
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
If GetFileAttributesEx(StrPtr("\\?\" & PathName), 0, VarPtr(FAD)) <> 0 Then
    Dim FT As FILETIME, ST As SYSTEMTIME
    FileTimeToLocalFileTime VarPtr(FAD.FTLastWriteTime), VarPtr(FT)
    FileTimeToSystemTime VarPtr(FT), VarPtr(ST)
    FileDateTime = DateSerial(ST.wYear, ST.wMonth, ST.wDay) + TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
Else
    Err.Raise Number:=53, Description:="File not found: '" & PathName & "'"
End If
End Function

' (VB-Overwrite)
Public Function Command$()
If InIDE() = False Then
    SysReAllocString VarPtr(Command$), PathGetArgs(GetCommandLine())
    Command$ = LTrim$(Command$)
Else
    Command$ = VBA.Command$()
End If
End Function

Public Function FileExists(ByVal PathName As String) As Boolean
On Error Resume Next
Dim Attributes As VbFileAttribute, ErrVal As Long
Attributes = GetAttr(PathName)
ErrVal = Err.Number
On Error GoTo 0
If (Attributes And (vbDirectory Or vbVolume)) = 0 And ErrVal = 0 Then FileExists = True
End Function

Public Function AppPath() As String
If InIDE() = False Then
    Const MAX_PATH_W As Long = 32767
    Dim Buffer As String, RetVal As Long
    Buffer = String(MAX_PATH, vbNullChar)
    RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH)
    If RetVal = MAX_PATH Then ' Path > MAX_PATH
        Buffer = String(MAX_PATH_W, vbNullChar)
        RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH_W)
    End If
    If RetVal > 0 Then
        Buffer = Left$(Buffer, RetVal)
        AppPath = Left$(Buffer, InStrRev(Buffer, "\"))
    Else
        AppPath = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\")
    End If
Else
    AppPath = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\")
End If
End Function

Public Function AppEXEName() As String
If InIDE() = False Then
    Const MAX_PATH_W As Long = 32767
    Dim Buffer As String, RetVal As Long
    Buffer = String(MAX_PATH, vbNullChar)
    RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH)
    If RetVal = MAX_PATH Then ' Path > MAX_PATH
        Buffer = String(MAX_PATH_W, vbNullChar)
        RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH_W)
    End If
    If RetVal > 0 Then
        Buffer = Left$(Buffer, RetVal)
        Buffer = Right$(Buffer, Len(Buffer) - InStrRev(Buffer, "\"))
        AppEXEName = Left$(Buffer, InStrRev(Buffer, ".") - 1)
    Else
        AppEXEName = App.EXEName
    End If
Else
    AppEXEName = App.EXEName
End If
End Function

Public Function AppMajor() As Integer
If InIDE() = False Then
    With GetAppVersionInfo()
    AppMajor = .dwFileVersionMSHi
    End With
Else
    AppMajor = App.Major
End If
End Function

Public Function AppMinor() As Integer
If InIDE() = False Then
    With GetAppVersionInfo()
    AppMinor = .dwFileVersionMSLo
    End With
Else
    AppMinor = App.Minor
End If
End Function

Public Function AppRevision() As Integer
If InIDE() = False Then
    With GetAppVersionInfo()
    AppRevision = .dwFileVersionLSLo
    End With
Else
    AppRevision = App.Revision
End If
End Function

Private Function GetAppVersionInfo() As VS_FIXEDFILEINFO
Static Done As Boolean, Value As VS_FIXEDFILEINFO
If Done = False Then
    Const MAX_PATH_W As Long = 32767
    Dim Buffer As String, RetVal As Long
    Buffer = String(MAX_PATH, vbNullChar)
    RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH)
    If RetVal = MAX_PATH Then ' Path > MAX_PATH
        Buffer = String(MAX_PATH_W, vbNullChar)
        RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH_W)
    End If
    If RetVal > 0 Then
        Dim ImagePath As String, Length As Long
        ImagePath = Left$(Buffer, RetVal)
        Length = GetFileVersionInfoSize(StrPtr(ImagePath), 0)
        If Length > 0 Then
            Dim DataBuffer() As Byte
            ReDim DataBuffer(0 To (Length - 1)) As Byte
            If GetFileVersionInfo(StrPtr(ImagePath), 0, Length, VarPtr(DataBuffer(0))) <> 0 Then
                Dim hData As Long
                If VerQueryValue(VarPtr(DataBuffer(0)), StrPtr("\"), hData, Length) <> 0 Then
                    If hData <> 0 Then CopyMemory Value, ByVal hData, LenB(Value)
                End If
            End If
        End If
    End If
    Done = True
End If
LSet GetAppVersionInfo = Value
End Function

Public Function GetClipboardText() As String
Const CF_UNICODETEXT As Long = 13
Dim lpText As Long, lpMem As Long, Length As Long
If OpenClipboard(0) <> 0 Then
    If IsClipboardFormatAvailable(CF_UNICODETEXT) <> 0 Then
        lpText = GetClipboardData(CF_UNICODETEXT)
        If lpText <> 0 Then
            lpMem = GlobalLock(lpText)
            If lpMem <> 0 Then
                Length = lstrlen(lpMem)
                If Length > 0 Then
                    GetClipboardText = String(Length, vbNullChar)
                    lstrcpy StrPtr(GetClipboardText), lpMem
                End If
                GlobalUnlock lpMem
            End If
        End If
    End If
    CloseClipboard
End If
End Function

Public Sub SetClipboardText(ByRef Text As String)
Const CF_UNICODETEXT As Long = 13
Const GMEM_MOVEABLE As Long = &H2
Dim Buffer As String, Length As Long
Dim hMem As Long, lpMem As Long
If OpenClipboard(0) <> 0 Then
    EmptyClipboard
    Buffer = Text & vbNullChar
    Length = LenB(Buffer)
    hMem = GlobalAlloc(GMEM_MOVEABLE, Length)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            CopyMemory ByVal lpMem, ByVal StrPtr(Buffer), Length
            GlobalUnlock hMem
            SetClipboardData CF_UNICODETEXT, hMem
        End If
    End If
    CloseClipboard
End If
End Sub

Public Function AccelCharCode(ByVal Caption As String) As Integer
If Caption = vbNullString Then Exit Function
Dim Pos As Long, Length As Long
Length = Len(Caption)
Pos = Length
Do
    If Mid$(Caption, Pos, 1) = "&" And Pos < Length Then
        AccelCharCode = Asc(UCase$(Mid$(Caption, Pos + 1, 1)))
        If Pos > 1 Then
            If Mid$(Caption, Pos - 1, 1) = "&" Then AccelCharCode = 0
        Else
            If AccelCharCode = vbKeyUp Then AccelCharCode = 0
        End If
        If AccelCharCode <> 0 Then Exit Do
    End If
    Pos = Pos - 1
Loop Until Pos = 0
End Function

Public Function ProperControlName(ByVal Control As VB.Control) As String
Dim Index As Long
On Error Resume Next
Index = Control.Index
If Err.Number <> 0 Or Index < 0 Then ProperControlName = Control.Name Else ProperControlName = Control.Name & "(" & Index & ")"
On Error GoTo 0
End Function

Public Function GetTopUserControl(ByVal UserControl As Object) As VB.UserControl
If UserControl Is Nothing Then Exit Function
Dim TopUserControl As VB.UserControl, TempUserControl As VB.UserControl
CopyMemory TempUserControl, ObjPtr(UserControl), 4
Set TopUserControl = TempUserControl
CopyMemory TempUserControl, 0&, 4
With TopUserControl
If .ParentControls.Count > 0 Then
    Dim OldParentControlsType As VBRUN.ParentControlsType
    OldParentControlsType = .ParentControls.ParentControlsType
    .ParentControls.ParentControlsType = vbExtender
    If TypeOf .ParentControls(0) Is VB.VBControlExtender Then
        .ParentControls.ParentControlsType = vbNoExtender
        CopyMemory TempUserControl, ObjPtr(.ParentControls(0)), 4
        Set TopUserControl = TempUserControl
        CopyMemory TempUserControl, 0&, 4
        Dim TempParentControlsType As VBRUN.ParentControlsType
        Do
            With TopUserControl
            If .ParentControls.Count = 0 Then Exit Do
            TempParentControlsType = .ParentControls.ParentControlsType
            .ParentControls.ParentControlsType = vbExtender
            If TypeOf .ParentControls(0) Is VB.VBControlExtender Then
                .ParentControls.ParentControlsType = vbNoExtender
                CopyMemory TempUserControl, ObjPtr(.ParentControls(0)), 4
                Set TopUserControl = TempUserControl
                CopyMemory TempUserControl, 0&, 4
                .ParentControls.ParentControlsType = TempParentControlsType
            Else
                .ParentControls.ParentControlsType = TempParentControlsType
                Exit Do
            End If
            End With
        Loop
    End If
    .ParentControls.ParentControlsType = OldParentControlsType
End If
End With
Set GetTopUserControl = TopUserControl
End Function

Public Function MousePointerID(ByVal MousePointer As Integer) As Long
Select Case MousePointer
    Case vbArrow
        Const IDC_ARROW As Long = 32512
        MousePointerID = IDC_ARROW
    Case vbCrosshair
        Const IDC_CROSS As Long = 32515
        MousePointerID = IDC_CROSS
    Case vbIbeam
        Const IDC_IBEAM As Long = 32513
        MousePointerID = IDC_IBEAM
    Case vbIconPointer ' Obselete, replaced Icon with Hand
        Const IDC_HAND As Long = 32649
        MousePointerID = IDC_HAND
    Case vbSizePointer, vbSizeAll
        Const IDC_SIZEALL As Long = 32646
        MousePointerID = IDC_SIZEALL
    Case vbSizeNESW
        Const IDC_SIZENESW As Long = 32643
        MousePointerID = IDC_SIZENESW
    Case vbSizeNS
        Const IDC_SIZENS As Long = 32645
        MousePointerID = IDC_SIZENS
    Case vbSizeNWSE
        Const IDC_SIZENWSE As Long = 32642
        MousePointerID = IDC_SIZENWSE
    Case vbSizeWE
        Const IDC_SIZEWE As Long = 32644
        MousePointerID = IDC_SIZEWE
    Case vbUpArrow
        Const IDC_UPARROW As Long = 32516
        MousePointerID = IDC_UPARROW
    Case vbHourglass
        Const IDC_WAIT As Long = 32514
        MousePointerID = IDC_WAIT
    Case vbNoDrop
        Const IDC_NO As Long = 32648
        MousePointerID = IDC_NO
    Case vbArrowHourglass
        Const IDC_APPSTARTING As Long = 32650
        MousePointerID = IDC_APPSTARTING
    Case vbArrowQuestion
        Const IDC_HELP As Long = 32651
        MousePointerID = IDC_HELP
    Case 16
        Const IDC_WAITCD As Long = 32663 ' Undocumented
        MousePointerID = IDC_WAITCD
End Select
End Function

Public Sub RefreshMousePointer(Optional ByVal hWndFallback As Long)
Const WM_SETCURSOR As Long = &H20, WM_NCHITTEST As Long = &H84, WM_MOUSEMOVE As Long = &H200
Dim P As POINTAPI, hWndCursor As Long
GetCursorPos P
hWndCursor = GetCapture()
If hWndCursor = 0 Then hWndCursor = WindowFromPoint(P.X, P.Y)
If hWndCursor <> 0 Then
    If GetWindowThreadProcessId(hWndCursor, 0) <> App.ThreadID Then hWndCursor = hWndFallback
Else
    hWndCursor = hWndFallback
End If
If hWndCursor <> 0 Then SendMessage hWndCursor, WM_SETCURSOR, hWndCursor, ByVal MakeDWord(SendMessage(hWndCursor, WM_NCHITTEST, 0, ByVal Make_XY_lParam(P.X, P.Y)), WM_MOUSEMOVE)
End Sub

Public Function OLEFontIsEqual(ByVal Font As StdFont, ByVal FontOther As StdFont) As Boolean
If Font Is Nothing Then
    If FontOther Is Nothing Then OLEFontIsEqual = True
ElseIf FontOther Is Nothing Then
    If Font Is Nothing Then OLEFontIsEqual = True
Else
    If Font.Name = FontOther.Name And Font.Size = FontOther.Size And Font.Charset = FontOther.Charset And Font.Weight = FontOther.Weight And _
    Font.Underline = FontOther.Underline And Font.Italic = FontOther.Italic And Font.Strikethrough = FontOther.Strikethrough Then
        OLEFontIsEqual = True
    End If
End If
End Function

Public Function CreateGDIFontFromOLEFont(ByVal Font As IFont) As Long
If Font Is Nothing Then Exit Function
Dim LF As LOGFONT
' hFont will be cleared when the IFont reference goes out of scope or is set to nothing.
GetObjectAPI Font.hFont, LenB(LF), LF
CreateGDIFontFromOLEFont = CreateFontIndirect(LF)
End Function

Public Function CloneOLEFont(ByVal Font As IFont) As StdFont
If Not Font Is Nothing Then Font.Clone CloneOLEFont
End Function

Public Function GetNumberGroupDigit() As String
GetNumberGroupDigit = Mid$(FormatNumber(1000, 0, , , vbTrue), 2, 1)
If GetNumberGroupDigit = "0" Then GetNumberGroupDigit = vbNullString
End Function

Public Function GetDecimalChar() As String
GetDecimalChar = Mid$(CStr(1.1), 2, 1)
End Function

Public Function IsFormLoaded(ByVal FormName As String) As Boolean
Dim i As Integer
For i = 0 To Forms.Count - 1
    If StrComp(Forms(i).Name, FormName, vbTextCompare) = 0 Then
        IsFormLoaded = True
        Exit For
    End If
Next i
End Function

Public Function GetWindowTitle(ByVal hWnd As Long) As String
Dim Buffer As String
Buffer = String(GetWindowTextLength(hWnd) + 1, vbNullChar)
GetWindowText hWnd, StrPtr(Buffer), Len(Buffer)
GetWindowTitle = Left$(Buffer, Len(Buffer) - 1)
End Function

Public Function GetWindowClassName(ByVal hWnd As Long) As String
Dim Buffer As String, RetVal As Long
Buffer = String(256, vbNullChar)
RetVal = GetClassName(hWnd, StrPtr(Buffer), Len(Buffer))
If RetVal <> 0 Then GetWindowClassName = Left$(Buffer, RetVal)
End Function

Public Sub CenterFormToScreen(ByVal Form As VB.Form, Optional ByVal RefForm As VB.Form)
Const MONITOR_DEFAULTTOPRIMARY As Long = &H1
If RefForm Is Nothing Then Set RefForm = Form
Dim hMonitor As Long, MI As MONITORINFO, WndRect As RECT
hMonitor = MonitorFromWindow(RefForm.hWnd, MONITOR_DEFAULTTOPRIMARY)
MI.cbSize = LenB(MI)
GetMonitorInfo hMonitor, MI
GetWindowRect Form.hWnd, WndRect
If TypeOf Form Is VB.MDIForm Then
    Dim MDIForm As VB.MDIForm
    Set MDIForm = Form
    MDIForm.Move (MI.RCMonitor.Left + (((MI.RCMonitor.Right - MI.RCMonitor.Left) - (WndRect.Right - WndRect.Left)) \ 2)) * (1440 / DPI_X()), (MI.RCMonitor.Top + (((MI.RCMonitor.Bottom - MI.RCMonitor.Top) - (WndRect.Bottom - WndRect.Top)) \ 2)) * (1440 / DPI_Y())
Else
    Form.Move (MI.RCMonitor.Left + (((MI.RCMonitor.Right - MI.RCMonitor.Left) - (WndRect.Right - WndRect.Left)) \ 2)) * (1440 / DPI_X()), (MI.RCMonitor.Top + (((MI.RCMonitor.Bottom - MI.RCMonitor.Top) - (WndRect.Bottom - WndRect.Top)) \ 2)) * (1440 / DPI_Y())
End If
End Sub

Public Sub FlashForm(ByVal Form As VB.Form)
Const FLASHW_CAPTION As Long = &H1, FLASHW_TRAY As Long = &H2, FLASHW_TIMERNOFG As Long = &HC
Dim FWI As FLASHWINFO
With FWI
.cbSize = LenB(FWI)
.dwFlags = FLASHW_CAPTION Or FLASHW_TRAY Or FLASHW_TIMERNOFG
.hWnd = Form.hWnd
.dwTimeout = 0 ' Default cursor blink rate
.uCount = 0
End With
FlashWindowEx FWI
End Sub

Public Function GetFormTitleBarHeight(ByVal Form As VB.Form) As Single
Const SM_CYCAPTION As Long = 4, SM_CYMENU As Long = 15
Const SM_CYSIZEFRAME As Long = 33, SM_CYFIXEDFRAME As Long = 8
Dim CY As Long
CY = GetSystemMetrics(SM_CYCAPTION)
If GetMenu(Form.hWnd) <> 0 Then CY = CY + GetSystemMetrics(SM_CYMENU)
Select Case Form.BorderStyle
    Case vbSizable, vbSizableToolWindow
        CY = CY + GetSystemMetrics(SM_CYSIZEFRAME)
    Case vbFixedSingle, vbFixedDialog, vbFixedToolWindow
        CY = CY + GetSystemMetrics(SM_CYFIXEDFRAME)
End Select
If CY > 0 Then GetFormTitleBarHeight = Form.ScaleY(CY, vbPixels, Form.ScaleMode)
End Function

Public Function GetFormNonScaleHeight(ByVal Form As VB.Form) As Single
Const SM_CYCAPTION As Long = 4, SM_CYMENU As Long = 15
Const SM_CYSIZEFRAME As Long = 33, SM_CYFIXEDFRAME As Long = 8
Dim CY As Long
CY = GetSystemMetrics(SM_CYCAPTION)
If GetMenu(Form.hWnd) <> 0 Then CY = CY + GetSystemMetrics(SM_CYMENU)
Select Case Form.BorderStyle
    Case vbSizable, vbSizableToolWindow
        CY = CY + (GetSystemMetrics(SM_CYSIZEFRAME) * 2)
    Case vbFixedSingle, vbFixedDialog, vbFixedToolWindow
        CY = CY + (GetSystemMetrics(SM_CYFIXEDFRAME) * 2)
End Select
If CY > 0 Then GetFormNonScaleHeight = Form.ScaleY(CY, vbPixels, Form.ScaleMode)
End Function

Public Sub SetWindowRedraw(ByVal hWnd As Long, ByVal Enabled As Boolean)
Const WM_SETREDRAW As Long = &HB
SendMessage hWnd, WM_SETREDRAW, IIf(Enabled = True, 1, 0), ByVal 0&
If Enabled = True Then
    Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
    RedrawWindow hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End If
End Sub

Public Function GetWindowsDir() As String
Static Done As Boolean, Value As String
If Done = False Then
    Dim Buffer As String
    Buffer = String(MAX_PATH, vbNullChar)
    If GetSystemWindowsDirectory(StrPtr(Buffer), MAX_PATH) <> 0 Then
        Value = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
        Value = Value & IIf(Right$(Value, 1) = "\", "", "\")
    End If
    Done = True
End If
GetWindowsDir = Value
End Function

Public Function GetSystemDir() As String
Static Done As Boolean, Value As String
If Done = False Then
    Dim Buffer As String
    Buffer = String(MAX_PATH, vbNullChar)
    If GetSystemDirectory(StrPtr(Buffer), MAX_PATH) <> 0 Then
        Value = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
        Value = Value & IIf(Right$(Value, 1) = "\", "", "\")
    End If
    Done = True
End If
GetSystemDir = Value
End Function

Public Function GetShiftStateFromParam(ByVal wParam As Long) As ShiftConstants
Const MK_SHIFT As Long = &H4, MK_CONTROL As Long = &H8
If (wParam And MK_SHIFT) = MK_SHIFT Then GetShiftStateFromParam = vbShiftMask
If (wParam And MK_CONTROL) = MK_CONTROL Then GetShiftStateFromParam = GetShiftStateFromParam Or vbCtrlMask
If GetKeyState(vbKeyMenu) < 0 Then GetShiftStateFromParam = GetShiftStateFromParam Or vbAltMask
End Function

Public Function GetMouseStateFromParam(ByVal wParam As Long) As MouseButtonConstants
Const MK_LBUTTON As Long = &H1, MK_RBUTTON As Long = &H2, MK_MBUTTON As Long = &H10
If (wParam And MK_LBUTTON) = MK_LBUTTON Then GetMouseStateFromParam = vbLeftButton
If (wParam And MK_RBUTTON) = MK_RBUTTON Then GetMouseStateFromParam = GetMouseStateFromParam Or vbRightButton
If (wParam And MK_MBUTTON) = MK_MBUTTON Then GetMouseStateFromParam = GetMouseStateFromParam Or vbMiddleButton
End Function

Public Function GetShiftStateFromMsg() As ShiftConstants
If GetKeyState(vbKeyShift) < 0 Then GetShiftStateFromMsg = vbShiftMask
If GetKeyState(vbKeyControl) < 0 Then GetShiftStateFromMsg = GetShiftStateFromMsg Or vbCtrlMask
If GetKeyState(vbKeyMenu) < 0 Then GetShiftStateFromMsg = GetShiftStateFromMsg Or vbAltMask
End Function

Public Function GetMouseStateFromMsg() As MouseButtonConstants
If GetKeyState(vbLeftButton) < 0 Then GetMouseStateFromMsg = vbLeftButton
If GetKeyState(vbRightButton) < 0 Then GetMouseStateFromMsg = GetMouseStateFromMsg Or vbRightButton
If GetKeyState(vbMiddleButton) < 0 Then GetMouseStateFromMsg = GetMouseStateFromMsg Or vbMiddleButton
End Function

Public Function GetShiftState() As ShiftConstants
GetShiftState = (-vbShiftMask * KeyPressed(vbKeyShift))
GetShiftState = GetShiftState Or (-vbCtrlMask * KeyPressed(vbKeyControl))
GetShiftState = GetShiftState Or (-vbAltMask * KeyPressed(vbKeyMenu))
End Function

Public Function GetMouseState() As MouseButtonConstants
Const SM_SWAPBUTTON As Long = 23
' GetAsyncKeyState requires a mapping of physical mouse buttons to logical mouse buttons.
GetMouseState = (-vbLeftButton * KeyPressed(IIf(GetSystemMetrics(SM_SWAPBUTTON) = 0, vbLeftButton, vbRightButton)))
GetMouseState = GetMouseState Or (-vbRightButton * KeyPressed(IIf(GetSystemMetrics(SM_SWAPBUTTON) = 0, vbRightButton, vbLeftButton)))
GetMouseState = GetMouseState Or (-vbMiddleButton * KeyPressed(vbMiddleButton))
End Function

Public Function KeyToggled(ByVal KeyCode As KeyCodeConstants) As Boolean
KeyToggled = CBool(LoByte(GetKeyState(KeyCode)) = 1)
End Function
 
Public Function KeyPressed(ByVal KeyCode As KeyCodeConstants) As Boolean
KeyPressed = CBool((GetAsyncKeyState(KeyCode) And &H8000&) = &H8000&)
End Function

Public Function InIDE(Optional ByRef B As Boolean = True) As Boolean
If B = True Then Debug.Assert Not InIDE(InIDE) Else B = True
End Function

Public Function PtrToObj(ByVal ObjectPointer As Long) As Object
Dim TempObj As Object
CopyMemory TempObj, ObjectPointer, 4
Set PtrToObj = TempObj
CopyMemory TempObj, 0&, 4
End Function

Public Function ProcPtr(ByVal Address As Long) As Long
ProcPtr = Address
End Function

Public Function LoByte(ByVal Word As Integer) As Byte
LoByte = Word And &HFF
End Function

Public Function HiByte(ByVal Word As Integer) As Byte
HiByte = (Word And &HFF00&) \ &H100
End Function

Public Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer
If (HiByte And &H80) <> 0 Then
    MakeWord = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
Else
    MakeWord = (HiByte * &H100) Or LoByte
End If
End Function

Public Function LoWord(ByVal DWord As Long) As Integer
If DWord And &H8000& Then
    LoWord = DWord Or &HFFFF0000
Else
    LoWord = DWord And &HFFFF&
End If
End Function

Public Function HiWord(ByVal DWord As Long) As Integer
HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

Public Function Get_X_lParam(ByVal lParam As Long) As Long
Get_X_lParam = lParam And &H7FFF&
If lParam And &H8000& Then Get_X_lParam = Get_X_lParam Or &HFFFF8000
End Function

Public Function Get_Y_lParam(ByVal lParam As Long) As Long
Get_Y_lParam = (lParam And &H7FFF0000) \ &H10000
If lParam And &H80000000 Then Get_Y_lParam = Get_Y_lParam Or &HFFFF8000
End Function

Public Function Make_XY_lParam(ByVal X As Long, ByVal Y As Long) As Long
Make_XY_lParam = LoWord(X) Or (&H10000 * LoWord(Y))
End Function

Public Function UTF32CodePoint_To_UTF16(ByVal CodePoint As Long) As String
If CodePoint >= &HFFFF8000 And CodePoint <= &H10FFFF Then
    Dim HW As Integer, LW As Integer
    If CodePoint < &H10000 Then
        HW = 0
        LW = CUIntToInt(CodePoint And &HFFFF&)
    Else
        CodePoint = CodePoint - &H10000
        HW = (CodePoint \ &H400) + &HD800
        LW = (CodePoint Mod &H400) + &HDC00
    End If
    If HW = 0 Then UTF32CodePoint_To_UTF16 = ChrW(LW) Else UTF32CodePoint_To_UTF16 = ChrW(HW) & ChrW(LW)
End If
End Function

Public Function UTF16_To_UTF8(ByRef Source As String) As Byte()
Const CP_UTF8 As Long = 65001
Dim Length As Long, Pointer As Long, Size As Long
Length = Len(Source)
Pointer = StrPtr(Source)
Size = WideCharToMultiByte(CP_UTF8, 0, Pointer, Length, 0, 0, 0, 0)
If Size > 0 Then
    Dim Buffer() As Byte
    ReDim Buffer(0 To Size - 1) As Byte
    WideCharToMultiByte CP_UTF8, 0, Pointer, Length, VarPtr(Buffer(0)), Size, 0, 0
    UTF16_To_UTF8 = Buffer()
End If
End Function

Public Function UTF8_To_UTF16(ByRef Source() As Byte) As String
If (0 / 1) + (Not Not Source()) = 0 Then Exit Function
Const CP_UTF8 As Long = 65001
Dim Size As Long, Pointer As Long, Length As Long
Size = UBound(Source) - LBound(Source) + 1
Pointer = VarPtr(Source(LBound(Source)))
Length = MultiByteToWideChar(CP_UTF8, 0, Pointer, Size, 0, 0)
If Length > 0 Then
    UTF8_To_UTF16 = Space$(Length)
    MultiByteToWideChar CP_UTF8, 0, Pointer, Size, StrPtr(UTF8_To_UTF16), Length
End If
End Function

Public Function StrToVar(ByVal Text As String) As Variant
If Text = vbNullString Then
    StrToVar = Empty
Else
    Dim B() As Byte
    B() = Text
    StrToVar = B()
End If
End Function

Public Function VarToStr(ByVal Bytes As Variant) As String
If IsEmpty(Bytes) Then
    VarToStr = vbNullString
Else
    Dim B() As Byte
    B() = Bytes
    VarToStr = B()
End If
End Function

Public Function UnsignedAdd(ByVal Start As Long, ByVal Incr As Long) As Long
UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
End Function

Public Function UnsignedSub(ByVal Start As Long, ByVal Decr As Long) As Long
UnsignedSub = ((Start And &H7FFFFFFF) - (Decr And &H7FFFFFFF)) Xor ((Start Xor Decr) And &H80000000)
End Function

Public Function CUIntToInt(ByVal Value As Long) As Integer
Const OFFSET_2 As Long = 65536
Const MAXINT_2 As Integer = 32767
If Value < 0 Or Value >= OFFSET_2 Then Err.Raise 6
If Value <= MAXINT_2 Then
    CUIntToInt = Value
Else
    CUIntToInt = Value - OFFSET_2
End If
End Function

Public Function CIntToUInt(ByVal Value As Integer) As Long
Const OFFSET_2 As Long = 65536
If Value < 0 Then
    CIntToUInt = Value + OFFSET_2
Else
    CIntToUInt = Value
End If
End Function

Public Function CULngToLng(ByVal Value As Double) As Long
Const OFFSET_4 As Double = 4294967296#
Const MAXINT_4 As Long = 2147483647
If Value < 0 Or Value >= OFFSET_4 Then Err.Raise 6
If Value <= MAXINT_4 Then
    CULngToLng = Value
Else
    CULngToLng = Value - OFFSET_4
End If
End Function

Public Function CLngToULng(ByVal Value As Long) As Double
Const OFFSET_4 As Double = 4294967296#
If Value < 0 Then
    CLngToULng = Value + OFFSET_4
Else
    CLngToULng = Value
End If
End Function

Public Function DPI_X() As Long
Const LOGPIXELSX As Long = 88
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    DPI_X = GetDeviceCaps(hDCScreen, LOGPIXELSX)
    ReleaseDC 0, hDCScreen
End If
End Function

Public Function DPI_Y() As Long
Const LOGPIXELSY As Long = 90
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    DPI_Y = GetDeviceCaps(hDCScreen, LOGPIXELSY)
    ReleaseDC 0, hDCScreen
End If
End Function

Public Function DPICorrectionFactor() As Single
Static Done As Boolean, Value As Single
If Done = False Then
    Value = ((96 / DPI_X()) * 15) / Screen.TwipsPerPixelX
    Done = True
End If
' Returns exactly 1 when no corrections are required.
DPICorrectionFactor = Value
End Function

Public Function CHimetricToPixel_X(ByVal Width As Long) As Long
Const HIMETRIC_PER_INCH As Long = 2540
CHimetricToPixel_X = (Width * DPI_X()) / HIMETRIC_PER_INCH
End Function

Public Function CHimetricToPixel_Y(ByVal Height As Long) As Long
Const HIMETRIC_PER_INCH As Long = 2540
CHimetricToPixel_Y = (Height * DPI_Y()) / HIMETRIC_PER_INCH
End Function

Public Function PixelsPerDIP_X() As Single
Static Done As Boolean, Value As Single
If Done = False Then
    Value = (DPI_X() / 96)
    Done = True
End If
PixelsPerDIP_X = Value
End Function

Public Function PixelsPerDIP_Y() As Single
Static Done As Boolean, Value As Single
If Done = False Then
    Value = (DPI_Y() / 96)
    Done = True
End If
PixelsPerDIP_Y = Value
End Function

Public Function WinColor(ByVal Color As Long, Optional ByVal hPal As Long) As Long
If OleTranslateColor(Color, hPal, WinColor) <> 0 Then WinColor = -1
End Function

Public Function PictureFromByteStream(ByRef ByteStream As Variant) As IPictureDisp
Const GMEM_MOVEABLE As Long = &H2
Dim IID As CLSID, Stream As IUnknown, NewPicture As IPicture
Dim B() As Byte, ByteCount As Long
Dim hMem As Long, lpMem As Long
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(3) = &HAA
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With
If VarType(ByteStream) = (vbArray + vbByte) Then
    B() = ByteStream
    ByteCount = (UBound(B()) - LBound(B())) + 1
    hMem = GlobalAlloc(GMEM_MOVEABLE, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            CopyMemory ByVal lpMem, B(LBound(B())), ByteCount
            GlobalUnlock hMem
            If CreateStreamOnHGlobal(hMem, 1, Stream) = 0 Then
                If OleLoadPicture(Stream, ByteCount, 0, IID, NewPicture) = 0 Then Set PictureFromByteStream = NewPicture
            End If
        End If
    End If
End If
End Function

Public Function PictureFromPath(ByVal PathName As String) As IPictureDisp
Dim IID As CLSID, NewPicture As IPicture
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(3) = &HAA
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With
If OleLoadPicturePath(StrPtr(PathName), 0, 0, 0, IID, NewPicture) = 0 Then Set PictureFromPath = NewPicture
End Function

Public Function PictureFromHandle(ByVal Handle As Long, ByVal PicType As VBRUN.PictureTypeConstants) As IPictureDisp
If Handle = 0 Then Exit Function
Dim PICD As PICTDESC, IID As CLSID, NewPicture As IPicture
With PICD
.cbSizeOfStruct = LenB(PICD)
.PicType = PicType
.hImage = Handle
End With
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(3) = &HAA
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With
If OleCreatePictureIndirect(PICD, IID, 1, NewPicture) = 0 Then Set PictureFromHandle = NewPicture
End Function

Public Function BitmapHandleFromPicture(ByVal Picture As IPictureDisp, Optional ByVal BackColor As OLE_COLOR) As Long
If Picture Is Nothing Then Exit Function
With Picture
If .Handle <> 0 Then
    Dim hDCScreen As Long, hDC As Long, hBmp As Long, hBmpOld As Long
    Dim CX As Long, CY As Long, Brush As Long
    CX = CHimetricToPixel_X(.Width)
    CY = CHimetricToPixel_Y(.Height)
    Brush = CreateSolidBrush(WinColor(BackColor))
    hDCScreen = GetDC(0)
    If hDCScreen <> 0 Then
        hDC = CreateCompatibleDC(hDCScreen)
        If hDC <> 0 Then
            hBmp = CreateCompatibleBitmap(hDCScreen, CX, CY)
            If hBmp <> 0 Then
                hBmpOld = SelectObject(hDC, hBmp)
                If .Type = vbPicTypeIcon Then
                    Const DI_NORMAL As Long = &H3
                    DrawIconEx hDC, 0, 0, .Handle, CX, CY, 0, Brush, DI_NORMAL
                Else
                    Dim RC As RECT
                    RC.Right = CX
                    RC.Bottom = CY
                    FillRect hDC, RC, Brush
                    .Render hDC Or 0&, 0&, 0&, CX Or 0&, CY Or 0&, 0&, .Height, .Width, -.Height, ByVal 0&
                End If
                SelectObject hDC, hBmpOld
                BitmapHandleFromPicture = hBmp
            End If
            DeleteDC hDC
        End If
        ReleaseDC 0, hDCScreen
    End If
    DeleteObject Brush
End If
End With
End Function

Public Sub RenderPicture(ByVal Picture As IPicture, ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal CX As Long, Optional ByVal CY As Long, Optional ByRef RenderFlag As Integer)
' RenderFlag is passed as a optional parameter ByRef.
' It is ignored for icons and metafiles.
' 0 = render method unknown, determine it and update parameter
' 1 = StdPicture.Render
' 2 = GdiAlphaBlend
If Picture Is Nothing Then Exit Sub
With Picture
If .Handle <> 0 Then
    If CX = 0 Then CX = CHimetricToPixel_X(.Width)
    If CY = 0 Then CY = CHimetricToPixel_Y(.Height)
    If .Type = vbPicTypeIcon Then
        Const DI_NORMAL As Long = &H3
        DrawIconEx hDC, X, Y, .Handle, CX, CY, 0, 0, DI_NORMAL
    Else
        Dim HasAlpha As Boolean
        If .Type = vbPicTypeBitmap Then
            If RenderFlag = 0 Then
                Const PICTURE_TRANSPARENT As Long = &H2
                If (.Attributes And PICTURE_TRANSPARENT) = 0 Then ' Exclude GIF
                    Dim Bmp As BITMAP
                    GetObjectAPI .Handle, LenB(Bmp), Bmp
                    If Bmp.BMBitsPixel = 32 And Bmp.BMBits <> 0 Then
                        Dim SA1D As SAFEARRAY1D, B() As Byte
                        With SA1D
                        .cDims = 1
                        .fFeatures = 0
                        .cbElements = 1
                        .cLocks = 0
                        .pvData = Bmp.BMBits
                        .Bounds.lLbound = 0
                        .Bounds.cElements = Bmp.BMWidthBytes * Bmp.BMHeight
                        End With
                        CopyMemory ByVal ArrPtr(B()), VarPtr(SA1D), 4
                        Dim i As Long, j As Long, Pos As Long
                        For i = 0 To (Abs(Bmp.BMHeight) - 1)
                            Pos = i * Bmp.BMWidthBytes
                            For j = (Pos + 3) To (Pos + Bmp.BMWidthBytes - 1) Step 4
                                If HasAlpha = False Then HasAlpha = (B(j) > 0)
                                If HasAlpha = True Then
                                    If B(j - 1) > B(j) Then
                                        HasAlpha = False
                                        i = Abs(Bmp.BMHeight) - 1
                                        Exit For
                                    ElseIf B(j - 2) > B(j) Then
                                        HasAlpha = False
                                        i = Abs(Bmp.BMHeight) - 1
                                        Exit For
                                    ElseIf B(j - 3) > B(j) Then
                                        HasAlpha = False
                                        i = Abs(Bmp.BMHeight) - 1
                                        Exit For
                                    End If
                                End If
                            Next j
                        Next i
                        CopyMemory ByVal ArrPtr(B()), 0&, 4
                    End If
                End If
                If HasAlpha = False Then RenderFlag = 1 Else RenderFlag = 2
            ElseIf RenderFlag = 2 Then
                HasAlpha = True
            End If
        End If
        If HasAlpha = False Then
            .Render hDC Or 0&, X Or 0&, Y Or 0&, CX Or 0&, CY Or 0&, 0&, .Height, .Width, -.Height, ByVal 0&
        Else
            Dim hDCBmp As Long, hBmpOld As Long
            hDCBmp = CreateCompatibleDC(0)
            If hDCBmp <> 0 Then
                hBmpOld = SelectObject(hDCBmp, .Handle)
                GdiAlphaBlend hDC, X, Y, CX, CY, hDCBmp, 0, 0, CHimetricToPixel_X(.Width), CHimetricToPixel_Y(.Height), &H1FF0000
                SelectObject hDCBmp, hBmpOld
                DeleteDC hDCBmp
            End If
        End If
    End If
End If
End With
End Sub
