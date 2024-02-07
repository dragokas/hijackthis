Attribute VB_Name = "modShortcut"
'[modShortcut.bas]

'
' Shortcut parser (light version) module by Alex Dragokas
'

Option Explicit

Private Type EXP_SZ_LINK
    cbSize                      As Long
    dwSignature                 As Long
    szTarget(0 To MAX_PATH - 1) As Byte
    swzTarget                   As String * MAX_PATH
End Type

Private Type UUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszGuid As Long, pGuid As UUID) As Long
Private Declare Function CoCreateInstance Lib "ole32.dll" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, pvarResult As Object) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, ByVal lpFilePart As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
'Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
'Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, ByVal pszStrPtr As Long) As Long

Private Declare Function CallWindowProcA Lib "user32.dll" (ByVal pFunc As Long, ByVal pESL As Long, ByVal pStrOut As Long, Optional ByVal Reserved1 As Long, Optional ByVal Reserved2 As Long) As Long
'Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
'Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
'Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
'Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
'Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
'Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long

Private Const MAX_PATH_W            As Long = 32767&
Private Const ERROR_MORE_DATA       As Long = 234&

Dim oPFile          As IPersistFile
Dim oSLink          As IShellLinkW
Dim oSLDL           As IShellLinkDataList

Dim CLSID_InternetShortcut  As UUID

Private LnkHeader(19)        As Byte


Public Function GetFileFromShortcut(path As String, Optional out_Args As String, Optional IsTypeLnk As Boolean) As String
    On Error GoTo ErrorHandler

    Dim Target  As String
    Dim ObjPath As String
    Dim sExt    As String
    
    If IsTypeLnk Then
        sExt = ".LNK"
    Else
        sExt = UCase$(GetExtensionName(path))
    End If
    
    Select Case sExt
    
        Case ".LNK"
        
            GetTargetShellLinkW path, Target, out_Args
    
            ' IDL on target ?  -> expand
            If Left$(Target, 3&) = "::{" Or Left$(Target, 4&) = "::\{" Then
                ObjPath = GetPathFromIDL(Target)
                If Len(ObjPath) <> 0 Then Target = ObjPath
            End If
    
        Case ".URL", ".WEBSITE"
            
            Target = GetUrlTargetW(path)
        
        Case ".PIF"
    
            If GetPIF_target(path, ObjPath, out_Args) Then Target = ObjPath
    
    End Select
    
    GetFileFromShortcut = Target
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFileFromShortcut", "Path: " & path
    If inIDE Then Stop: Resume Next
End Function


Public Function GetPathFromIDL(sIDL As String) As String
    On Error Resume Next
    
    Dim Shl     As Object
    Dim fld     As Object
    Dim path    As String
    Dim itm     As Variant
    
    AppendErrorLogCustom "GetPathFromIDL - Begin", "IDL: " & sIDL
    
    Set Shl = CreateObject("shell.application")
    If Err.Number <> 0 Then
        'Library or registry entries are damaged
        ErrorMsg Err, "Parser.GetPathFromIDL", Translate(512) & ": Shell32.dll"
        Exit Function
    End If
    If Left$(sIDL, 4) = "::\{" Then sIDL = "::" & mid$(sIDL, 4) 'trim \
    
    Set fld = Shl.NameSpace$(CVar(sIDL))
    
    If (Err.Number <> 0) Or (fld Is Nothing) Then Exit Function
    
    On Error GoTo ErrorHandler
    path = fld.Self.path
    
    If Len(path) <> 0 And StrComp(path, sIDL, 1) <> 0 Then
        GetPathFromIDL = path
    Else
        For Each itm In fld.Items
            path = itm.path
            GetPathFromIDL = GetPathName(path)
            Exit For
        Next
    End If
    Set fld = Nothing
    Set Shl = Nothing
    
    AppendErrorLogCustom "GetPathFromIDL - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetPathFromIDL", "IDL: " & sIDL
    If inIDE Then Stop: Resume Next
End Function


' получает цель и аргументы LNK
Public Sub GetTargetShellLinkW(LNK_file As String, Optional Target As String, Optional Argument As String)
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "GetTargetShellLinkW - Begin", "File: " & LNK_file
    
    Dim fd              As WIN32_FIND_DATAW
    Dim ptr             As Long
    Dim lr              As Long
    Dim Flags           As Long
    
    Static bTerminalServerEmulation As Boolean
    Static SysRoot2                 As String
    Static isInit                   As Boolean
    
    If Not isInit Then
        isInit = True
        If OSver.IsServer Then
            SysRoot2 = String$(MAX_PATH, 0&)
            lr = GetWindowsDirectory(StrPtr(SysRoot2), MAX_PATH)
            If lr Then SysRoot2 = Left$(SysRoot2, lr)
            If StrComp(sWinDir, SysRoot2, 1) <> 0 Then bTerminalServerEmulation = True
        End If
    End If

    If Not FileExists(LNK_file) Then Exit Sub

    ' Проверяем целостность заголовка LNK
    
    Dim FileHeader() As Byte
    
    FileHeader = GetHeaderFromFile(LNK_file, 20&)
    
    If StrComp(LnkHeader, FileHeader) <> 0 Then

        Argument = vbNullString
        Target = "(lnk is corrupted)"

        Exit Sub
    End If
    
    oPFile.Load LNK_file, STGM_READ
    
    If (oSLDL Is Nothing) Then
        If inIDE Then Debug.Print "oPFile.Load is failed. Error: " & Err.LastDllError & ". File: " & LNK_file
    Else
        If S_OK = oSLDL.GetFlags(Flags) Then
        
            If Flags And SLDF_HAS_EXP_SZ Then
                
                If S_OK = oSLDL.CopyDataBlock(EXP_SZ_LINK_SIG, ptr) Then
        
                    If ptr Then
                        CallWindowProcA AddressOf DerefDataBlock, ptr, VarPtr(Target)
            
                        ptr = LocalFree(ptr)
            
                        Target = EnvironW(Target)
                    End If
                End If
            End If
        End If
    End If
    
    If 0 = Len(Target) Then
    
        Target = String$(MAX_PATH, vbNullChar)
        oSLink.getPath Target, MAX_PATH, fd, SLGP_UNCPRIORITY
        Target = Left$(Target, lstrlen(StrPtr(Target)))
        
        If Len(Target) > 255 Then
            Target = String$(MAX_PATH_W, vbNullChar)
            oSLink.getPath Target, MAX_PATH_W, fd, SLGP_UNCPRIORITY
            Target = Left$(Target, lstrlen(StrPtr(Target)))
        End If
        
        If bTerminalServerEmulation Then
        
            If StrBeginWith(Target, SysRoot2) Then
                Target = Replace$(Target, SysRoot2, sWinDir, 1, 1, vbTextCompare)
            End If
        End If
        
        'If OSver.IsLocalSystemContext Then
        '    Target = PathSubstituteProfile(Target, LNK_file)
        'End If
        'temporarily hack - substitute profile in any case
        'to do it normally, I need make manual parsing of LNK (already done) and return 'Special Folder' ID,
        'or just check if first token in IDList represent link to 'Special folder ID'. In such case call PathSubstituteProfile
        
        Target = PathSubstituteProfile(Target, ExtractProfilePath(LNK_file))
        
        Target = GetFullPath(Target)
    Else
        Target = GetFullPath(Left$(Target, lstrlen(StrPtr(Target))))
    End If
    
    Argument = String$(MAX_PATH, 0)
    oSLink.GetArguments Argument, MAX_PATH
    Argument = Left$(Argument, lstrlen(StrPtr(Argument)))
    
    If Len(Argument) > 255 Then
        Argument = String$(MAX_PATH_W, 0)
        oSLink.GetArguments Argument, MAX_PATH_W
        Argument = Left$(Argument, lstrlen(StrPtr(Argument)))
    End If
    
    'добавил trim пробелов (приём игры в прятки вирмейкеров :)
    If 0 <> Len(Argument) Then Argument = Trim$(Argument)

    AppendErrorLogCustom "GetTargetShellLinkW - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Parser.GetTargetShellLinkW", "File: ", LNK_file
    If inIDE Then Stop: Resume Next
End Sub

Private Function DerefDataBlock(ByRef ESL As EXP_SZ_LINK, ByRef StrOut As String, Optional ByVal Reserved1 As Long, Optional ByVal Reserved2 As Long) As Long
    SysReAllocString VarPtr(StrOut), StrPtr(ESL.swzTarget)
End Function

Private Function ExtractProfilePath(sPath As String) As String
    Dim pos As Long
    pos = InStr(Len(ProfilesDir) + 2, sPath, "\")
    If pos <> 0 Then
        ExtractProfilePath = Left$(sPath, pos - 1)
    Else
        ExtractProfilePath = sPath
    End If
End Function


' Возвращает заголовок файла
Public Function GetHeaderFromFile(FileName As String, BytesCnt As Long) As Byte()
    On Error GoTo ErrorHandler
    
    Dim ff As Long
    Dim Size As Currency
    Dim Data() As Byte
    
    OpenW FileName, FOR_READ, ff, g_FileBackupFlag
    If ff < 1 Then Exit Function
    
    Size = LOFW(ff)
    If Size = 0@ Then CloseW ff: ff = 0: Exit Function
    If BytesCnt > Size Then BytesCnt = Size
    
    ReDim Data(BytesCnt - 1)
    GetW ff, 1&, , VarPtr(Data(0)), BytesCnt
    CloseW ff: ff = 0
    
    GetHeaderFromFile = Data
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetHeaderFromFile", "File:", FileName
    If ff <> 0 Then CloseW ff: ff = 0
End Function

Private Function isFileFilledByNUL(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim Data As String
    Dim i As Long
    
    Data = ReadFileContents(FileName, False)
    
    isFileFilledByNUL = True
    
    For i = 1 To Len(Data)
        If AscW(mid$(Data, i, 1)) <> 0& Then isFileFilledByNUL = False: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.isFileFilledByNUL", "File:", FileName
End Function

' Инициализация интерфейса IShellLink
Public Sub ISL_Init()
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "ISL_Init - Begin"
    
    Const CLSIDSTR_ShellLink As String = "{00021401-0000-0000-C000-000000000046}"
    Const IIDSTR_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
    Const CLSCTX_INPROC_SERVER As Long = 1
    
    Dim CLSID_ShellLink As UUID
    Dim IID_IUnknown    As UUID
    Dim oUnknown        As IUnknown

    CLSIDFromString StrPtr(CLSIDSTR_ShellLink), CLSID_ShellLink
    CLSIDFromString StrPtr(IIDSTR_IUnknown), IID_IUnknown
    CoCreateInstance CLSID_ShellLink, 0&, CLSCTX_INPROC_SERVER, IID_IUnknown, oUnknown
 
    LnkHeader(0) = &H4C
    LnkHeader(4) = 1
    LnkHeader(5) = &H14
    LnkHeader(6) = 2
    LnkHeader(12) = &HC0
    LnkHeader(19) = &H46
 
    Set oPFile = oUnknown
    Set oSLink = oUnknown
    Set oSLDL = oUnknown
    
    AppendErrorLogCustom "ISL_Init - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Parser.ISL_Init"
    If inIDE Then Stop: Resume Next
End Sub

' Освобождение памяти, занятой объектом IUnknown для интерфейса IShellLink
Public Sub ISL_Dispatch()
    Set oPFile = Nothing
    Set oSLink = Nothing
    Set oSLDL = Nothing
End Sub

Public Function GetUrlTargetW(URLpathW As String) As String
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "GetUrlTargetW - Begin", "Path: " & URLpathW

    Dim lr          As Long
    Dim buf         As String
    Dim sTemp       As String
    Dim CodePage    As Long
    Dim aBuf()      As Byte
    Dim cpPercent   As Long
    Dim sAppend     As String
    
    buf = String$(256&, 0)
    lr = GetPrivateProfileString(StrPtr("InternetShortcut"), StrPtr("URL"), StrPtr(""), StrPtr(buf), Len(buf), StrPtr(URLpathW))
    If Err.LastDllError = ERROR_MORE_DATA Then
        buf = String$(1001&, 0)
        lr = GetPrivateProfileString(StrPtr("InternetShortcut"), StrPtr("URL"), StrPtr(""), StrPtr(buf), Len(buf), StrPtr(URLpathW))
        If Err.LastDllError = ERROR_MORE_DATA Then
            buf = String$(10001&, 0)
            lr = GetPrivateProfileString(StrPtr("InternetShortcut"), StrPtr("URL"), StrPtr(""), StrPtr(buf), Len(buf), StrPtr(URLpathW))
            If Err.LastDllError = ERROR_MORE_DATA Then
                sAppend = "(" & Translate(1080) & " > 10000 " & Translate(1081) & ")" 'length of address > X characters
            Else
                sAppend = "(" & Translate(1080) & " = " & lr & " " & Translate(1081) & ")"
            End If
        End If
    End If
    
    If lr <> 0 Then
        If lr > 1000 Then
            buf = Left$(buf, 1000&)
        Else
            buf = Left$(buf, lr)
        End If
        
        sTemp = UnEscape(buf)
        If Len(sTemp) <> 0 Then buf = sTemp
        
        aBuf() = StrConv(buf, vbFromUnicode, OSver.LangNonUnicodeCode)
        
        CodePage = GetEncoding(aBuf, cpPercent, URLpathW)
        
        If (CP_UTF8 = CodePage) And (cpPercent = -1 Or cpPercent > 10) Then
            sTemp = ConvertCodePage(VarPtr(aBuf(0)), CP_UTF8)
            If Len(sTemp) <> 0 Then buf = sTemp
        End If
        
        GetUrlTargetW = buf & IIf(Len(sAppend) <> 0, " ... " & sAppend, vbNullString)
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetUrlTargetW", "File: ", URLpathW
    If inIDE Then Stop: Resume Next
End Function

' Нормализация пути
Public Function GetFullPath(sFilename As String) As String
    On Error GoTo ErrorHandler
    Dim cnt        As Long
    Dim sFullName  As String
    
    sFullName = String$(MAX_PATH_W, 0)
    cnt = GetFullPathName(StrPtr(sFilename), MAX_PATH_W, StrPtr(sFullName), 0&)
    If cnt Then
        GetFullPath = Left$(sFullName, cnt)
    Else
        GetFullPath = sFilename
    End If
    If Right$(GetFullPath, 1) = "\" Then GetFullPath = Left$(GetFullPath, Len(GetFullPath) - 1)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFullPath"
    If inIDE Then Stop: Resume Next
End Function

' Раскрытие цели и аргумента ярлыков PIF
Public Function GetPIF_target(FileName As String, Target As String, Argument As String) As Boolean
    'thanks to Sergey Merzlikin  ( http://www.smsoft.ru/ru/pifdoc.htm )
    
    ' offset 0x24 (длина: 63 байта) - цель
    ' offset 0xA5 (длина: 64 байта) - аргумент
    
    Dim pif_Target  As String
    Dim pif_Arg     As String
        
    On Error GoTo ErrorHandler
    
    Dim sBuffer As String
    Dim Header  As String
    Dim FLen    As Currency
    Dim cnt     As Long
    Dim ff      As Long
    
    pif_Target = String$(63&, vbNullChar)
    pif_Arg = String$(64&, vbNullChar)
  
    If Not OpenW(FileName, FOR_READ, ff, g_FileBackupFlag) Then Exit Function
    FLen = LOFW(ff)
    
    If FLen >= &H187& Then    '  NT / 2000
        ' Check header
        Header = String$(15&, vbNullChar) ' 16-th is NULL char
        GetW ff, &H171& + 1&, Header
        If Header <> "MICROSOFT PIFEX" Then CloseHandle ff: ff = 0: Exit Function 'incorrect header
        
    ElseIf FLen = &H171& Then ' Windows 1.X
        ' It's Okay (no header)
        
    ElseIf FLen < &H171& Then
        ' incorrect PIF
        CloseHandle ff: ff = 0
        Exit Function
    End If
    
    GetW ff, &H24& + 1&, pif_Target
    GetW ff, &HA5& + 1&, pif_Arg
    CloseHandle ff: ff = 0
    
    pif_Arg = Left$(pif_Arg, lstrlen(StrPtr(pif_Arg)))
    pif_Target = Left$(pif_Target, lstrlen(StrPtr(pif_Target)))
    
    If FileExists(pif_Target) Then    'DOS -> to Full name
        sBuffer = String$(MAX_PATH_W, vbNullChar)
        cnt = GetLongPathName(StrPtr(pif_Target), StrPtr(sBuffer), Len(sBuffer))
        If cnt Then
            pif_Target = Left$(sBuffer, cnt)
        End If
    End If
    
    GetPIF_target = True
    Target = pif_Target
    Argument = pif_Arg
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetPIF_target", "File:", FileName
    If inIDE Then Stop: Resume Next
End Function

Public Function GetEncoding(aBytes() As Byte, Optional Percent As Long, Optional sSource As String) As Long
    On Error GoTo ErrorHandler

    AppendErrorLogCustom "Parser.GetEncoding - Begin"

    Dim MLang       As CMultiLanguage
    Dim IMLang2     As IMultiLanguage2
    Dim Encoding()  As tagDetectEncodingInfo
    Dim encCount    As Long
    Dim Index       As Long
    
    Set MLang = New CMultiLanguage
    Set IMLang2 = MLang
    
    encCount = 16
    ReDim Encoding(encCount - 1)
    IMLang2.DetectInputCodepage 0, 0, aBytes(0), UBound(aBytes) + 1, Encoding(0), encCount
    
    For Index = 0 To encCount - 1
        
        Percent = Encoding(Index).nDocPercent
        'BytesCnt = Encoding(index).nConfidence
        GetEncoding = Encoding(Index).nCodePage
        
        Exit Function
    Next
    
    AppendErrorLogCustom "Parser.GetEncoding - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetEncoding", "Source:", sSource, "Data: 0x" & ByteArrayToHex(aBytes)
    'try UTF-8 recognize alternative
    GetEncoding = GetEncoding_UTF8(aBytes(), Percent)
End Function

Function GetEncoding_UTF8(aBytes() As Byte, Optional Percent As Long) As idCodePage
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "Parser.GetEncoding_UTF8 - Begin"
    
    Dim c As Long, N As Long, NTotal As Long, i As Long, bSuccess As Boolean, BTC As Long 'bytes to check
    
    Do
        '2-bytes seq.: 110x xxxx, 10xx xxxx (0xC0, 0x80)
        '              110...= 0xC0, 10...= 0x80
        '              111...= 0xE0, 11...= 0xC0
        
        '3-bytes seq.: 1110 xxxx, 10xx xxxx, 10xx xxxx (0xE0, 0x80, 0x80)
        '              111... = 0xE0
        '              1111...= 0xF0
        
        '4-bytes seq.: 1111 0xxx, 10xx xxxx, 10xx xxxx, 10xx xxxx (0xF0, 0x80, 0x80, 0x80)
        '              1111...  = 0xF0
        '              1111 1...= 0xF8
        
        BTC = 0
        If ((aBytes(c) Xor &HC0) And &HE0) = 0 Then
            BTC = 1
        ElseIf ((aBytes(c) Xor &HE0) And &HF0) = 0 Then
            BTC = 2
        ElseIf ((aBytes(c) Xor &HF0) And &HF8) = 0 Then
            BTC = 3
        End If
        
        If (BTC > 0) And ((c + BTC) <= UBound(aBytes)) Then
            bSuccess = True
            For i = c + 1 To c + BTC
                If ((aBytes(i) Xor &H80) And &HC0) <> 0 Then bSuccess = False: Exit For
            Next
            If bSuccess Then N = N + 1
        End If
        
        If aBytes(c) <> 13 And aBytes(c) <> 10 Then
            NTotal = NTotal + 1
        End If
        
        c = c + 1 + BTC
        
    Loop While c <= UBound(aBytes)
    
    If NTotal <> 0 Then
        Percent = N * 100& / NTotal
    End If
    
    If Percent >= 5 Then Percent = -1: GetEncoding_UTF8 = CP_UTF8
    
    AppendErrorLogCustom "Parser.GetEncoding_UTF8 - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetEncoding_UTF8"
End Function

Public Function ByteArrayToHex(arr() As Byte) As String
    Dim i&, s As String
    For i = 0 To UBound(arr)
        s = s & Right$("0" & Hex$(arr(i)), 2)
    Next
    ByteArrayToHex = s
End Function


Public Function CreateHJTShortcuts(HJT_Location As String) As Boolean
    Dim bSuccess As Boolean
    Dim hFile As Long
    bSuccess = True
    bSuccess = bSuccess And MkDirW(BuildPath(StartMenuPrograms, "HiJackThis+"), , True)
    bSuccess = bSuccess And MkDirW(BuildPath(StartMenuPrograms, "HiJackThis+\Tools"), , True)
    bSuccess = bSuccess And MkDirW(BuildPath(StartMenuPrograms, "HiJackThis+\Plugins"), , True)
    
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\HiJackThis.lnk"), HJT_Location)
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Uninstall HJT.lnk"), HJT_Location, "/uninstall")
    
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Tools\StartupList.lnk"), HJT_Location, "/tool+StartupList", , , "Lists startup items in a convenient way")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Tools\Uninstall Manager.lnk"), HJT_Location, "/tool+UninstMan", , , "Manage the list of installed software")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Tools\Digital Signature Checker.lnk"), HJT_Location, "/tool+DigiSign", , , "Check your PE EXE files digital signature")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Tools\Registry Key Unlocker.lnk"), HJT_Location, "/tool+RegUnlocker", , , "Unlock and reset permissions on registry key")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Tools\Registry Key Type Analyzer.lnk"), HJT_Location, "/tool+RegTypeChecker", , , "Retrieves various information about registry key")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Tools\ADS Spy.lnk"), HJT_Location, "/tool+ADSSpy", , , "Alternative Data Streams Scanner & Remover")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Tools\Hosts File Manager.lnk"), HJT_Location, "/tool+Hosts", , , "Manage entries in hosts file")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Tools\Process Manager.lnk"), HJT_Location, "/tool+ProcMan", , , "Little tool like Task Manager")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Plugins\Check Browsers' LNK.lnk"), HJT_Location, "/tool+CheckLNK", , , "Checks your PC for shortcuts infection")
    bSuccess = bSuccess And CreateShortcut(BuildPath(StartMenuPrograms, "HiJackThis+\Plugins\ClearLNK.lnk"), HJT_Location, "/tool+ClearLNK", , , "LNK / URL Shortcuts cleaner & restorer")
    
    'Users manual url shortcut
    If IsRussianAreaLangCode(OSver.LangSystemCode) Or IsRussianAreaLangCode(OSver.LangDisplayCode) Then
    
        If OpenW(BuildPath(StartMenuPrograms, "HiJackThis+\" & LoadResString(607) & ".url"), FOR_OVERWRITE_CREATE, hFile) Then
            PrintLineW hFile, "[InternetShortcut]", False
            PrintLineW hFile, "URL=https://regist.safezone.cc/hijackthis_help/hijackthis.html", False
            CloseW hFile
        End If
        If OpenW(BuildPath(StartMenuPrograms, "HiJackThis+\" & LoadResString(608) & ".url"), FOR_OVERWRITE_CREATE, hFile) Then
            PrintLineW hFile, "[InternetShortcut]", False
            PrintLineW hFile, "URL=https://www.safezone.cc/threads/dopolnenie-v-rukovodstvo-po-hijackthis.27470/", False
            CloseW hFile
        End If
    Else
        If OpenW(BuildPath(StartMenuPrograms, "HiJackThis+\Users manual.url"), FOR_OVERWRITE_CREATE, hFile) Then
            PrintLineW hFile, "[InternetShortcut]", False
            PrintLineW hFile, "URL=" & GetTutorialURL(), False
            CloseW hFile
        End If
    End If
    CreateHJTShortcuts = bSuccess
End Function

Public Function CreateHJTShortcutDesktop(HJT_Location As String) As Boolean
    CreateHJTShortcutDesktop = CreateShortcut(BuildPath(Desktop, "HiJackThis+.lnk"), HJT_Location)
End Function

Public Sub RemoveHJTShortcuts()
    Call DeleteFolderForce(BuildPath(StartMenuPrograms, "HiJackThis Fork"))
    Call DeleteFolderForce(BuildPath(StartMenuPrograms, "HiJackThis+"))
    Call DeleteFileW(StrPtr(BuildPath(Desktop, "HiJackThis Fork.lnk")))
    Call DeleteFileW(StrPtr(BuildPath(Desktop, "HiJackThis+.lnk")))
End Sub

Public Function CreateShortcut( _
    sPathLnk As String, _
    sTarget As String, _
    Optional sArg As String = vbNullString, _
    Optional sIcon As String = vbNullString, _
    Optional iShowCmd As Long = 1, _
    Optional sDescription As String = vbNullString) As Boolean
    
    If Len(sIcon) = 0 Then sIcon = sTarget
    
    DeleteFileW StrPtr(sPathLnk)
    
    With oSLink
        .SetPath sTarget
        .SetArguments sArg
        .SetWorkingDirectory GetParentDir(sTarget)
        .SetIconLocation sIcon, 0
        .SetShowCmd iShowCmd
        .SetDescription sDescription
    End With
    oPFile.Save sPathLnk, 1&
    oPFile.SaveCompleted sPathLnk
    
    CreateShortcut = FileExists(sPathLnk)
End Function
