Attribute VB_Name = "modShortcut"
'
' Shortcut parser (light version) module by Alex Dragokas
'

Option Explicit

Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszGuid As Long, pGuid As UUID) As Long
Private Declare Function CoCreateInstance Lib "ole32.dll" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, pvarResult As Object) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, ByVal lpFilePart As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long

Const MAX_PATH_W    As Long = 32767&

Dim oPFile          As IPersistFile
Dim oSLink          As IShellLinkW

Dim CLSID_InternetShortcut  As UUID


Public Function GetFileFromShortcut(Path As String, Optional out_Args As String, Optional ForceLNK As Boolean) As String
    On Error GoTo ErrorHandler

    Dim Target  As String
    Dim ObjPath As String
    Dim sExt    As String

    If ForceLNK Then
        sExt = ".LNK"
    Else
        sExt = UCase$(GetExtensionName(Path))
    End If

    Select Case sExt
    
        Case ".LNK"
        
            GetTargetShellLinkW Path, Target, out_Args
    
            ' IDL on target ?  -> expand
            If Left$(Target, 3&) = "::{" Or Left$(Target, 4&) = "::\{" Then
                ObjPath = GetPathFromIDL(Target)
                If Len(ObjPath) <> 0 Then Target = ObjPath
            End If
    
        Case ".URL", ".WEBSITE"
            
            Target = GetUrlTargetW(Path)
        
        Case ".PIF"
    
            If GetPIF_target(Path, ObjPath, out_Args) Then Target = ObjPath
    
    End Select
    
    GetFileFromShortcut = Target
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFileFromShortcut", "Path: " & Path
    If inIDE Then Stop: Resume Next
End Function


Public Function GetPathFromIDL(sIDL As String) As String
    On Error Resume Next
    
    Dim shl     As Object
    Dim fld     As Object
    Dim Path    As String
    Dim itm     As Variant
    
    AppendErrorLogCustom "GetPathFromIDL - Begin", "IDL: " & sIDL
    
    Set shl = CreateObject("shell.application")
    If Err.Number <> 0 Then
        'Library or registry entries are damaged
        ErrorMsg Err, "Parser.GetPathFromIDL", Translate(512) & ": Shell32.dll"
        Exit Function
    End If
    If Left$(sIDL, 4) = "::\{" Then sIDL = "::" & Mid$(sIDL, 4) 'trim \
    
    Set fld = shl.Namespace(CVar(sIDL))
    
    If (Err.Number <> 0) Or (fld Is Nothing) Then Exit Function
    
    On Error GoTo ErrorHandler
    Path = fld.self.Path
    
    If Len(Path) <> 0 And StrComp(Path, sIDL, 1) <> 0 Then
        GetPathFromIDL = Path
    Else
        For Each itm In fld.Items
            Path = itm.Path
            GetPathFromIDL = GetPathName(Path)
            Exit For
        Next
    End If
    Set fld = Nothing
    Set shl = Nothing
    
    AppendErrorLogCustom "GetPathFromIDL - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetPathFromIDL", "IDL: " & sIDL
    If inIDE Then Stop: Resume Next
End Function


' получает цель и аргументы, как есть (IShellLink интерфейс)
Public Sub GetTargetShellLinkW(LNK_file As String, Optional Target As String, Optional Argument As String)
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "GetTargetShellLinkW - Begin", "File: " & LNK_file
    
    Dim fd              As WIN32_FIND_DATAW
    
    If Not FileExists(LNK_file) Then Exit Sub
    
    oPFile.Load LNK_file, STGM_READ
    
    Target = String$(MAX_PATH_W, vbNullChar)
    oSLink.GetPath Target, MAX_PATH_W, fd, SLGP_UNCPRIORITY
    Target = GetFullPath(Left$(Target, lstrlen(StrPtr(Target))))  ' НЕ ЗАМЕНЯТЬ http -> hxxp !!!
    
    Argument = String$(MAX_PATH_W, vbNullChar)
    oSLink.GetArguments Argument, MAX_PATH_W
    Argument = Left$(Argument, lstrlen(StrPtr(Argument)))
    
    AppendErrorLogCustom "GetTargetShellLinkW - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Parser.GetTargetShellLinkW", "File: ", LNK_file
    If inIDE Then Stop: Resume Next
End Sub

' Инициализация интерфейса IShellLink
Public Sub ISL_Init()
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "ISL_Init - Begin"
    
    Dim CLSID_ShellLink As UUID
    Dim IID_IUnknown    As UUID
    Dim oUnknown        As IUnknown

    CLSIDFromString StrPtr(CLSIDSTR_ShellLink), CLSID_ShellLink
    CLSIDFromString StrPtr(IIDSTR_IUnknown), IID_IUnknown
    CoCreateInstance CLSID_ShellLink, 0&, CLSCTX_INPROC_SERVER, IID_IUnknown, oUnknown
 
    Set oPFile = oUnknown
    Set oSLink = oUnknown
    
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
    Dim Stady       As Long
    
    buf = String$(256&, 0)
    lr = GetPrivateProfileString(StrPtr("InternetShortcut"), StrPtr("URL"), StrPtr(""), StrPtr(buf), Len(buf), StrPtr(URLpathW))
    If Err.LastDllError = ERROR_MORE_DATA Then
        buf = String$(1001&, 0)
        lr = GetPrivateProfileString(StrPtr("InternetShortcut"), StrPtr("URL"), StrPtr(""), StrPtr(buf), Len(buf), StrPtr(URLpathW))
        If Err.LastDllError = ERROR_MORE_DATA Then
            buf = String$(10001&, 0)
            lr = GetPrivateProfileString(StrPtr("InternetShortcut"), StrPtr("URL"), StrPtr(""), StrPtr(buf), Len(buf), StrPtr(URLpathW))
            If Err.LastDllError = ERROR_MORE_DATA Then
                'sAppend = "(" & "длина адреса" & " > 10000 " & "символов" & ")"
                sAppend = "(" & Translate(51) & " > 10000 " & Translate(52) & ")"
            Else
                'sAppend = "(" & "длина адреса" & " = " & lr & " " & "символов" & ")"
                sAppend = "(" & Translate(51) & " = " & lr & " " & Translate(52) & ")"
            End If
        End If
    End If
    
    Stady = 1
    
    If lr <> 0 Then
        If lr > 1000 Then
            buf = Left$(buf, 1000&)
        Else
            buf = Left$(buf, lr)
        End If
        
        Stady = 2
        
        sTemp = UnEscape(buf)
        If Len(sTemp) <> 0 Then buf = sTemp
        
        Stady = 3
        
        'identify codepage
        aBuf() = StrConv(buf, vbFromUnicode, &H419&)

        Stady = 4

        CodePage = GetEncoding(aBuf, cpPercent, URLpathW)
        
        Stady = 5
        
        If (UTF8 = CodePage) And (cpPercent = -1 Or cpPercent > 10) Then
            sTemp = ConvertCodePageW(buf, UTF8)
            If Len(sTemp) <> 0 Then buf = sTemp
        End If
        
        GetUrlTargetW = buf & IIf(Len(sAppend) <> 0, " ... " & sAppend, "")
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetUrlTargetW", "File: ", URLpathW, "Stady:", Stady
    If inIDE Then Stop: Resume Next
End Function

' Нормализация пути
Public Function GetFullPath(sFileName As String) As String
    On Error GoTo ErrorHandler
    Dim Cnt        As Long
    Dim sFullName  As String
    
    sFullName = String$(MAX_PATH_W, 0)
    Cnt = GetFullPathName(StrPtr(sFileName), MAX_PATH_W, StrPtr(sFullName), 0&)
    If Cnt Then
        GetFullPath = Left$(sFullName, Cnt)
    Else
        GetFullPath = sFileName
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFullPath"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetPathName(Path As String) As String   ' получить родительский каталог
    Dim pos As Long
    pos = InStrRev(Path, "\")
    If pos <> 0 Then GetPathName = Left$(Path, pos - 1)
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
    Dim Cnt     As Long
    Dim ff      As Long
    
    pif_Target = String$(63&, vbNullChar)
    pif_Arg = String$(64&, vbNullChar)
  
    If Not OpenW(FileName, FOR_READ, ff) Then Exit Function
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
        Cnt = GetLongPathName(StrPtr(pif_Target), StrPtr(sBuffer), Len(sBuffer))
        If Cnt Then
            pif_Target = Left$(sBuffer, Cnt)
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
    Dim inp()       As Byte
    Dim index       As Long
    Dim J           As Long
    
'    Open file For Binary As #1
'    ReDim inp(LOF(1) - 1)
'    Get #1, , inp()
'    Close #1
    
    Set MLang = New CMultiLanguage
    Set IMLang2 = MLang
    
    encCount = 16
    ReDim Encoding(encCount - 1)
    'IMLang2.DetectInputCodepage 0, 0, inp(0), UBound(inp) + 1, Encoding(0), encCount
    IMLang2.DetectInputCodepage 0, 0, aBytes(0), UBound(aBytes) + 1, Encoding(0), encCount
    
    For index = 0 To encCount - 1
        
'        Debug.Print file
'        Debug.Print "Задетектирован " & Encoding(index).nCodePage & _
'            ", кол-во: " & Encoding(index).nDocPercent & "%" & _
'            ", вероятность " & Encoding(index).nConfidence & "%"
        
        Percent = Encoding(index).nDocPercent
        'BytesCnt = Encoding(index).nConfidence
        GetEncoding = Encoding(index).nCodePage
        
'        If inIDE Then
'            If encCount > 1 Then
'                Debug.Print "Encoding Cnt: " & encCount
'                For j = 0 To encCount - 1
'                    Debug.Print "Задетектирован " & Encoding(j).nCodePage & _
'                        ", кол-во: " & Encoding(j).nDocPercent & "%" & _
'                        ", вероятность " & Encoding(j).nConfidence & "%"
'                Next
'            End If
'        End If
        
        Exit Function
    Next
    
    AppendErrorLogCustom "Parser.GetEncoding - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetEncoding", "Source:", sSource, "Data: 0x" & ByteArrayToHex(aBytes)
    'try UTF-8 recognize alternative
    GetEncoding = GetEncoding_UTF8(aBytes(), Percent)
End Function

Function GetEncoding_UTF8(aBytes() As Byte, Optional Percent As Long) As Long
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "Parser.GetEncoding_UTF8 - Begin"
    
    Dim c As Long, n As Long, i As Long, bSuccess As Boolean, btc As Long 'bytes to check
    
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
        
        btc = 0
        If ((aBytes(c) Xor &HC0) And &HE0) = 0 Then
            btc = 1
        ElseIf ((aBytes(c) Xor &HE0) And &HF0) = 0 Then
            btc = 2
        ElseIf ((aBytes(c) Xor &HF0) And &HF8) = 0 Then
            btc = 3
        End If
        
        If (btc > 0) And ((c + btc) <= UBound(aBytes)) Then
            bSuccess = True
            For i = c + 1 To c + btc
                If ((aBytes(c + 1) Xor &H80) And &HC0) <> 0 Then bSuccess = False: Exit For
            Next
            If bSuccess Then n = n + 1
        End If
        
        c = c + 1 + btc
        
    Loop Until c >= UBound(aBytes)
    
    Percent = n / UBound(aBytes) * 100&
    
    If Percent > 10 Then Percent = -1: GetEncoding_UTF8 = 65001
    
    AppendErrorLogCustom "Parser.GetEncoding_UTF8 - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetEncoding_UTF8"
End Function

Public Function ByteArrayToHex(arr() As Byte) As String
    Dim i&, s&
    For i = 0 To UBound(arr)
        s = s & Right$("0" & Hex(arr(i)), 2)
    Next
    ByteArrayToHex = s
End Function

