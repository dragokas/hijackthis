Attribute VB_Name = "modShortcut"
'
' Shortcut parser (light version) module by Alex Dragokas
'

Option Explicit

Private Type STRUCT_IDList
    size As Integer
    Type As Byte
    Unknown As Byte
    Data() As Byte
End Type

Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function MsiGetShortcutTarget Lib "Msi.dll" Alias "MsiGetShortcutTargetW" (ByVal szShortcutTarget As Long, ByVal szProductCode As Long, ByVal szFeatureId As Long, ByVal szComponentCode As Long) As Long
Private Declare Function MsiGetComponentPath Lib "Msi.dll" Alias "MsiGetComponentPathW" (ByVal szProduct As Long, ByVal szComponent As Long, ByVal lpPathBuf As Long, pcchBuf As Long) As Long

Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszGuid As Long, pGuid As UUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As UUID, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function CoCreateInstance Lib "ole32.dll" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, pvarResult As Object) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, ByVal lpFilePart As Long) As Long
Private Declare Function PathRemoveFileSpec Lib "Shlwapi.dll" Alias "PathRemoveFileSpecW" (ByVal pszPath As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As Long, ByVal IBindCtx As Long, ppidl As Long, ByVal sfgaoIn As Long, psfgaoOut As Long) As Long
Private Declare Function SHGetDesktopFolder Lib "shell32.dll" (ISF As IShellFolder) As Long
Private Declare Function StrRetToStr Lib "Shlwapi.dll" Alias "StrRetToStrA" (pstr As STRRET, ByVal pIDL As Long, ppsz As String) As Long
Private Declare Function ILFree Lib "shell32.dll" (ByVal pidlFree As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long

Const MaxFeatureLength          As Long = 38&
Const MaxGuidLength             As Long = 38&
Const MaxPathLength             As Long = 1024&
Const INSTALLSTATE_LOCAL        As Long = 3&

Const MAX_PATH_W    As Long = 32767&

Dim oPFile          As IPersistFile
Dim oSLink          As IShellLinkW
Dim IURL            As IUniformResourceLocatorW
Dim IPF_URL         As IPersistFile

Dim IID_IURLW               As UUID
Dim CLSID_InternetShortcut  As UUID




Public Function GetFileFromShortcut(Path As String) As String
    Dim Target  As String
    Dim ObjPath As String
    Dim LNK_Arg As String

    Select Case UCase$(GetExtensionName(Path))
    
        Case ".LNK"
        
            ' if Windows Installer LNK
            Target = GetMSILinkTarget(Path)
        
            If Len(Target) = 0& Then
                GetTargetShellLinkW Path, Target, LNK_Arg
            Else
                ' для MSI LNK также может быть аргумент
            End If
    
            ' IDL on target ?  -> expand
            If Left$(Target, 3&) = "::{" Or Left$(Target, 4&) = "::\{" Then
                ObjPath = GetPathFromIDL(Target)
                If Len(ObjPath) <> 0 Then Target = ObjPath
            End If
    
        Case ".URL", ".WEBSITE"
            
            Target = GetUrlTargetW(Path)
        
        Case ".PIF"
    
            If GetPIF_target(Path, ObjPath, LNK_Arg) Then Target = ObjPath
    
    End Select
    
    GetFileFromShortcut = Target
End Function


Public Function GetIDLDisplayName(sIDL As String) As String
    On Error GoTo ErrorHandler
    
    Const STRRET_WSTR   As Long = 0&
    Const STRRET_OFFSET As Long = 1&
    Const STRRET_CSTR   As Long = 2&
    
    Dim ISF        As IShellFolder
    Dim oStrRet    As STRRET
    Dim pIDL       As Long
    Dim pstr       As Long
    Dim lr         As Long
    Dim ObjName    As String
    Dim ptr        As Long
    
    If Left$(sIDL, 4) = "::\{" Then sIDL = "::" & Mid$(sIDL, 4) 'trim \
    
    ' Создать Item ID List из строки и вернуть указатель на него
    If SHParseDisplayName(StrPtr(sIDL), ByVal 0&, pIDL, ByVal 0&, ByVal 0&) <> S_OK Then Exit Function
    
    If 0 = pIDL Then Exit Function
    
    ' Получаем указатель на интерфейс IShellFolder корневого объекта - Desktop
    lr = SHGetDesktopFolder(ISF)
    
    If S_OK <> lr Then Exit Function
    
    ' Преобразуем IDL в имя, отображаемое по-умолчанию в Windows Explorer
    ISF.GetDisplayNameOf pIDL, ByVal 0&, oStrRet
    
    ' Проверка валидности данных в структуре STRRET
    
    ' если тип = указатель на строку
    If STRRET_WSTR = oStrRet.uType Then
        ' проверяем, является ли это указателем
        'GetMem4 ByVal VarPtr(oStrRet) + 4, ptr
        'If 0 = ptr Then
        With oStrRet
          If .CStr(0) + .CStr(1) + .CStr(2) + .CStr(3) = 0 Then
            Set ISF = Nothing
            ILFree pIDL
            Exit Function
          End If
        End With
    End If
    
    ' Структуру STRRET -> в строку
    ObjName = String$(MAX_PATH, vbNullChar)
    
    If StrRetToStr(oStrRet, pIDL, ObjName) = S_OK Then
        GetIDLDisplayName = RTrimNull(ObjName)
        CoTaskMemFree StrPtr(ObjName)
    End If
    
    Set ISF = Nothing
    ILFree pIDL
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "Parser.GetIDLDisplayName", "IDL: ", sIDL
    If inIDE Then Stop: Resume Next
End Function

Public Function GetPathFromIDL(sIDL As String) As String
    On Error GoTo ErrorHandler
    Dim shl     As Object
    Dim fld     As Object
    Dim Path    As String
    Dim itm     As Variant
    
    Set shl = CreateObject("shell.application")
    If err.Number <> 0 Then
        ' "Подробности: повреждена или отсутствует регистрация библиотеки Shell32.dll"
        ErrorMsg err, "Parser.GetPathFromIDL", "Library is damaged or not exists: " & " Shell32.dll"
        Exit Function
    End If
    If Left$(sIDL, 4) = "::\{" Then sIDL = "::" & Mid$(sIDL, 4) 'trim \
    
    Set fld = shl.Namespace(CVar(sIDL))
    
    If (err.Number <> 0) Or (fld Is Nothing) Then Exit Function
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
    Exit Function
ErrorHandler:
    ErrorMsg err, "Parser.GetPathFromIDL", "IDL: ", sIDL
    If inIDE Then Stop: Resume Next
End Function

Public Function GetMSILinkTarget(Link As String) As String
    On Error GoTo ErrorHandler

    Dim lr              As Long
    Dim TargetSize      As Long
    Dim ProductCode     As String
    Dim FeatureID       As String
    Dim ComponentCode   As String
    Dim Target          As String
    Static MSI_init     As Boolean
    Static MSI_ok       As Boolean
    
    If MSI_init Then
        If Not MSI_ok Then Exit Function
    Else
        Dim hLib     As Long
        Dim hModule  As Long
        MSI_init = True
        hLib = LoadLibrary(StrPtr("msi.dll"))
        If hLib <> 0 Then
            hModule = GetProcAddress(hLib, "MsiGetShortcutTargetW")
            If hModule <> 0 Then
                hModule = 0
                hModule = GetProcAddress(hLib, "MsiGetComponentPathW")
                MSI_ok = (hModule <> 0)
            End If
            FreeLibrary hLib
        End If
        If Not MSI_ok Then
            
            If Not FileExists(sWinDir & "\System32\msi.dll") Then
                err.Clear
                ErrorMsg err, "Library is not exists" & ": msi.dll"   '"Отсутствует библиотека"
            Else
                err.Clear
                ErrorMsg err, "Library is damaged" & ": msi.dll"    '"Повреждена библиотека"
            End If
            Exit Function
        End If
    End If
    
    ProductCode = Space$(MaxGuidLength + 1)
    FeatureID = Space$(MaxFeatureLength + 1)
    ComponentCode = Space$(MaxGuidLength + 1)
    
    lr = MsiGetShortcutTarget(StrPtr(Link), StrPtr(ProductCode), StrPtr(FeatureID), StrPtr(ComponentCode))
    
    If lr = 0 Then
        TargetSize = MaxPathLength
        Target = Space(TargetSize)
    
        lr = MsiGetComponentPath(StrPtr(ProductCode), StrPtr(ComponentCode), StrPtr(Target), TargetSize)
    
        If lr = INSTALLSTATE_LOCAL Then
            GetMSILinkTarget = Trim$(Left$(Target, TargetSize))
        Else
            GetMSILinkTarget = "INSTALLSTATE_UNKNOWN"
        End If
    ElseIf Len(Trim$(ProductCode)) <> 0 Then
        GetMSILinkTarget = "INSTALLSTATE_UNKNOWN"
    End If
    
    'Alternative:
    'Get Pidl from GetIDList
    '
    'SHGetPathFromIDList(pidl, target)
    'hres = pShLink->GetIDList(&pidl);
    'if (SUCCEEDED(hres))
    '{
    '    SHGetPathFromIDList(pidl, target);
    '}
    Exit Function
ErrorHandler:
    ErrorMsg err, "Parser.GetMSILinkTarget", "File: ", Link, "GetMSILinkTarget: ", GetMSILinkTarget
    If inIDE Then Stop: Resume Next
End Function


' получает цель и аргументы, как есть (IShellLink интерфейс)
Public Sub GetTargetShellLinkW(LNK_file As String, Optional Target As String, Optional Argument As String)
    On Error GoTo ErrorHandler
    Dim fd              As WIN32_FIND_DATAW
    
    If Not FileExists(LNK_file) Then Exit Sub
    
    oPFile.Load LNK_file, STGM_READ
    
    Target = String$(MAX_PATH_W, vbNullChar)
    oSLink.GetPath Target, MAX_PATH_W, fd, SLGP_UNCPRIORITY
    Target = GetFullPath(Left$(Target, lstrlen(StrPtr(Target))))  ' НЕ ЗАМЕНЯТЬ http -> hxxp !!!
    
    Argument = String$(MAX_PATH_W, vbNullChar)
    oSLink.GetArguments Argument, MAX_PATH_W
    Argument = Left$(Argument, lstrlen(StrPtr(Argument)))
    
    'If Len(Target) = 0 Then
    '    If FileLenW(LNK_file) = 0 Then
    '        Raise_LNK_Struct_Error LNK_file, FileSizeIs0byte:=True
    '        Target = "?( 0 " & Translate(25) & " )"     '"?( 0 байт )"
    '    End If
    'End If
    Exit Sub
ErrorHandler:
    'If Err.Number = -2147467259 Then ' (Неопознанная ошибка) Automation error - Unspecified error . LastDllError = 0
    '    Raise_LNK_Struct_Error LNK_file
    'Else
        ErrorMsg err, "Parser.GetTargetShellLinkW", "File: ", LNK_file
        If inIDE Then Stop: Resume Next
    'End If
End Sub

' Инициализация интерфейса IShellLink
Public Sub ISL_Init()
    On Error GoTo ErrorHandler
    Dim CLSID_ShellLink As UUID
    Dim IID_IUnknown    As UUID
    Dim oUnknown        As IUnknown

    CLSIDFromString StrPtr(CLSIDSTR_ShellLink), CLSID_ShellLink
    CLSIDFromString StrPtr(IIDSTR_IUnknown), IID_IUnknown
    CoCreateInstance CLSID_ShellLink, 0&, CLSCTX_INPROC_SERVER, IID_IUnknown, oUnknown
 
    Set oPFile = oUnknown
    Set oSLink = oUnknown
    Exit Sub
ErrorHandler:
    ErrorMsg err, "Parser.ISL_Init"
    If inIDE Then Stop: Resume Next
End Sub

' Освобождение памяти, занятой объектом IUnknown для интерфейса IShellLink
Public Sub ISL_Dispatch()
    Set oPFile = Nothing
    Set oSLink = Nothing
End Sub

Sub IURL_Init()
    On Error GoTo ErrorHandler
    Const CLSIDSTR_InternetShortcut As String = "{FBF23B40-E3F0-101B-8488-00AA003E56F8}"
    Const IIDSTR_IURLW              As String = "{CABB0DA0-DA57-11CF-9974-0020AFD79762}"
      
    CLSIDFromString StrPtr(IIDSTR_IURLW), IID_IURLW
    CLSIDFromString StrPtr(CLSIDSTR_InternetShortcut), CLSID_InternetShortcut
    Exit Sub
ErrorHandler:
    ErrorMsg err, "Parser.IURL_Init"
    If inIDE Then Stop: Resume Next
End Sub


'' Получить цель из ярлыка URL
'Public Function GetUrlTargetW_Old(URLpathW As String) As String
'    On Error GoTo ErrorHandler
'    Dim strLen      As Long
'    Dim ptr         As Long
'    Dim URLtarget   As String
'
'    CoCreateInstance CLSID_InternetShortcut, 0&, CLSCTX_INPROC_SERVER, IID_IURLW, IURL
'    Set IPF_URL = IURL
'
'    ' Загружаем ярлык URL
'    IPF_URL.Load URLpathW, STGM_READ
'    ' Получаем указатель на строку с URL
'    ptr = IURL.GetUrl
'    strLen = lstrlen(ptr)
'    URLtarget = Space(strLen)
'    lstrcpyn StrPtr(URLtarget), ptr, strLen + 1
'    ' Освобождаем ресурсы
'    CoTaskMemFree ptr
'    Set IPF_URL = Nothing
'    Set IURL = Nothing
'
''    If Len(URLtarget) = 0 Then  'если цель не получена
''        Dim ff      As Long
''        Dim FSize   As Currency
''        Dim iData   As Integer
''
''        If OpenW(URLpathW, FOR_READ, ff) Then
''            FSize = LOFW(ff)
''            If FSize <> 0 Then  ' попробуем прочитать пару байт
''                GetW ff, 1&, iData
''            End If
''            CloseHandle ff
''        End If
''    End If
'
'    GetUrlTargetW_Old = URLtarget
'
'    Exit Function
'ErrorHandler:
'    ErrorMsg err, "Parser.GetUrlTargetW_Old", "File: ", URLpathW
'    If inIDE Then Stop: Resume Next
'End Function

Public Function GetUrlTargetW(URLpathW As String) As String
    On Error GoTo ErrorHandler
    Dim lr As Long
    Dim buf As String
    buf = String$(255&, vbNullChar)
    lr = GetPrivateProfileString(StrPtr("InternetShortcut"), StrPtr("URL"), StrPtr(""), StrPtr(buf), Len(buf), StrPtr(URLpathW))
    If lr <> 0 Then
        GetUrlTargetW = Left$(buf, lr)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "Parser.GetUrlTargetW", "File: ", URLpathW
    If inIDE Then Stop: Resume Next
End Function

' Нормализация пути
Public Function GetFullPath(sFileName As String) As String
    On Error GoTo ErrorHandler
    Dim cnt        As Long
    Dim sFullName  As String
    
    sFullName = Space(MAX_PATH_W)
    cnt = GetFullPathName(StrPtr(sFileName), MAX_PATH_W, StrPtr(sFullName), 0&)
    If cnt Then
        GetFullPath = Left$(sFullName, cnt)
    Else
        GetFullPath = sFileName
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "Parser.GetFullPath"
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
    Dim cnt     As Long
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
    ErrorMsg err, "Parser.GetPIF_target", "File:", FileName
    If inIDE Then Stop: Resume Next
End Function
