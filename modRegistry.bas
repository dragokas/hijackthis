Attribute VB_Name = "modRegistry"
Option Explicit
'
' Registry functions by Alex Dragokas
'
' This module is a part of HiJackThis project

' 2.0.7 - registry functions reworked completely

'Revision 2.2 (28.05.2016)

Public Const MAX_PATH       As Long = 260&
Public Const MAX_PATH_W     As Long = 32767&
Public Const MAX_KEYNAME    As Long = 255& 'https://msdn.microsoft.com/en-us/library/windows/desktop/ms724872(v=vs.85).aspx
Public Const MAX_VALUENAME  As Long = 32767&

Public Enum REG_VALUE_TYPE
    REG_NONE = 0&
    REG_SZ = 1&
    REG_EXPAND_SZ = 2&
    REG_BINARY = 3&
    REG_DWORD = 4&
    REG_DWORDLittleEndian = 4&
    REG_DWORDBigEndian = 5&
    REG_LINK = 6&
    REG_MULTI_SZ = 7&
    REG_ResourceList = 8&
    REG_FullResourceDescriptor = 9&
    REG_ResourceRequirementsList = 10&
    REG_QWORD = 11&
    REG_QWORD_LITTLE_ENDIAN = 11&
End Enum

Public Enum FLAG_REG_TYPE   'flags to be able to map bit mask and default registry type constants
    FLAG_REG_ALL = -1&
    FLAG_REG_NONE = 1&
    FLAG_REG_SZ = 2&
    FLAG_REG_EXPAND_SZ = 4&
    FLAG_REG_BINARY = 8&
    FLAG_REG_DWORD = &H10&
    FLAG_REG_DWORDLittleEndian = &H10&
    FLAG_REG_DWORDBigEndian = &H20&
    FLAG_REG_LINK = &H40&
    FLAG_REG_MULTI_SZ = &H80&
    FLAG_REG_ResourceList = &H100&
    FLAG_REG_FullResourceDescriptor = &H200&
    FLAG_REG_ResourceRequirementsList = &H400&
    FLAG_REG_QWORD = &H800&
    FLAG_REG_QWORD_LITTLE_ENDIAN = &H1000&
End Enum

Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes        As Long
    ftCreationTime          As FILETIME
    ftLastAccessTime        As FILETIME
    ftLastWriteTime         As FILETIME
    nFileSizeHigh           As Long
    nFileSizeLow            As Long
    dwReserved0             As Long
    dwReserved1             As Long
    lpszFileName(MAX_PATH)  As Integer
    lpszAlternate(14)       As Integer
End Type

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyW" (ByVal hKey As Long, ByVal lpClass As Long, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
'Private Declare Function RegGetValue Lib "advapi32.dll" Alias "RegGetValueW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal lpValue As Long, ByVal dwFlags As Long, pdwType As Long, ByVal pvData As Long, pcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueW" (ByVal hKey As Long, ByVal lpValueName As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
Private Declare Function RegDeleteKeyEx Lib "advapi32.dll" Alias "RegDeleteKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal samDesired As Long, ByVal Reserved As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long
'Private Declare Function SHFileExists Lib "shell32.dll" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function SHDeleteKey Lib "Shlwapi.dll" Alias "SHDeleteKeyW" (ByVal lRootKey As Long, ByVal szKeyToDelete As Long) As Long
Private Declare Function RegSaveKeyEx Lib "advapi32.dll" Alias "RegSaveKeyExW" (ByVal hKey As Long, ByVal lpFile As Long, ByVal lpSecurityAttributes As Long, ByVal flags As Long) As Long

Public Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Public Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" (lpSystemTime As SYSTEMTIME, vtime As Date) As Long

Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Sub memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetMem4 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long
Private Declare Function GetMem8 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long

Private Const INVALID_FILE_ATTRIBUTES As Long = -1&

Public Const HKEY_CLASSES_ROOT       As Long = &H80000000
Public Const HKEY_CURRENT_USER       As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE      As Long = &H80000002
Public Const HKEY_USERS              As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA   As Long = &H80000004
Public Const HKEY_CURRENT_CONFIG     As Long = &H80000005
Public Const HKEY_DYN_DATA           As Long = &H80000006

Public Const KEY_CREATE_SUB_KEY     As Long = &H4
Public Const KEY_QUERY_VALUE        As Long = &H1
Public Const KEY_SET_VALUE          As Long = &H2
Public Const READ_CONTROL           As Long = &H20000
Public Const WRITE_OWNER            As Long = &H80000
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const SYNCHRONIZE            As Long = &H100000
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const REG_OPTION_NON_VOLATILE As Long = 0
Public Const KEY_WOW64_64KEY        As Long = &H100&

Public Const ERROR_MORE_DATA        As Long = 234&
Public Const ERROR_SUCCESS          As Long = 0&
Public Const ERROR_ACCESS_DENIED    As Long = 5&

Private Const REG_STANDARD_FORMAT   As Long = 1&
Private Const REG_LATEST_FORMAT     As Long = 2&
Private Const RRF_RT_ANY            As Long = &HFFFF&
Private Const RRF_NOEXPAND          As Long = &H10000000


Private Function GetHKey(ByVal HKeyName As String) As Long 'Get handle of main hive
    On Error GoTo ErrorHandler:
    Dim pos As Long
    pos = InStr(HKeyName, "\")
    If pos <> 0 Then HKeyName = Left$(HKeyName, pos - 1)
    Select Case UCase$(HKeyName)
        Case "HKEY_CLASSES_ROOT", "HKCR"
            GetHKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER", "HKCU"
            GetHKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE", "HKLM"
            GetHKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS", "HKU"
            GetHKey = HKEY_USERS
        Case "HKEY_PERFORMANCE_DATA"
            GetHKey = HKEY_PERFORMANCE_DATA
        Case "HKEY_CURRENT_CONFIG", "HKCC"
            GetHKey = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            GetHKey = HKEY_DYN_DATA
    End Select
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetHKey", HKeyName
    If inIDE Then Stop: Resume Next
End Function


Public Function GetHiveNameByHandle(Handle As Long) As String
    On Error GoTo ErrorHandler:
    Select Case Handle
        Case HKEY_CLASSES_ROOT
            GetHiveNameByHandle = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER
            GetHiveNameByHandle = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            GetHiveNameByHandle = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            GetHiveNameByHandle = "HKEY_USERS"
        Case HKEY_PERFORMANCE_DATA
            GetHiveNameByHandle = "HKEY_PERFORMANCE_DATA"
        Case HKEY_CURRENT_CONFIG
            GetHiveNameByHandle = "HKEY_CURRENT_CONFIG"
        Case HKEY_DYN_DATA
            GetHiveNameByHandle = "HKEY_DYN_DATA"
    End Select
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetHiveNameByHandle", Handle
    If inIDE Then Stop: Resume Next
End Function


Public Function GetShortHiveName(ByVal FullHiveName As String) As String
    On Error GoTo ErrorHandler:
    Dim pos As Long
    pos = InStr(FullHiveName, "\")
    If pos <> 0 Then FullHiveName = Left$(FullHiveName, pos - 1)
    Select Case UCase$(FullHiveName)
        Case "HKEY_CLASSES_ROOT", "HKCR"
            GetShortHiveName = "HKCR"
        Case "HKEY_CURRENT_USER", "HKCU"
            GetShortHiveName = "HKCU"
        Case "HKEY_LOCAL_MACHINE", "HKLM"
            GetShortHiveName = "HKLM"
        Case "HKEY_USERS", "HKU"
            GetShortHiveName = "HKU"
        Case "HKEY_PERFORMANCE_DATA"
            GetShortHiveName = "HKPD"
        Case "HKEY_CURRENT_CONFIG", "HKCC"
            GetShortHiveName = "HKCC"
        Case "HKEY_DYN_DATA"
            GetShortHiveName = "HKDD"
    End Select
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetShortHiveName", FullHiveName
    If inIDE Then Stop: Resume Next
End Function


Private Function SwapEndian(ByVal dw As Long) As Long
    memcpy ByVal VarPtr(SwapEndian) + 3&, dw, 1&
    memcpy ByVal VarPtr(SwapEndian) + 2&, ByVal VarPtr(dw) + 1&, 1&
    memcpy ByVal VarPtr(SwapEndian) + 1&, ByVal VarPtr(dw) + 2&, 1&
    memcpy SwapEndian, ByVal VarPtr(dw) + 3&, 1&
End Function


Private Function ExpandEnvStr(sData As String) As String
'    Dim lRet     As Long
'    Dim sTemp    As String
'    lRet = ExpandEnvironmentStrings(StrPtr(sData), StrPtr(sTemp), lRet) 'get buffer size needed
'    sTemp = Space(lRet - 1&)
'    lRet = ExpandEnvironmentStrings(StrPtr(sData), StrPtr(sTemp), lRet)
'    If lRet Then
'        ExpandEnvStr = Left$(sTemp, lRet - 1&)
'    Else
'        ExpandEnvStr = sData
'    End If
    ExpandEnvStr = EnvironW(sData)
End Function


' if main hive handle wasn't defined, assigns handle according to hive's name defined by Full key name directed
Sub NormalizeKeyNameAndHiveHandle(ByRef lHive As Long, ByRef KeyName As String)
    Dim iPos        As Long
    If lHive = 0 Then
        lHive = GetHKey(KeyName)
        iPos = InStr(KeyName, "\")
        If (iPos <> 0&) Then KeyName = Mid$(KeyName, iPos + 1&) Else KeyName = vbNullString
    End If
End Sub


Public Function RegCreateKey(lHive&, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    lret = RegCreateKeyEx(lHive, StrPtr(sKey), 0&, ByVal 0&, 0&, KEY_CREATE_SUB_KEY Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), ByVal 0&, hKey, ByVal 0&)
    
    If lret = ERROR_ACCESS_DENIED Then
    
        If modPermissions.RegKeyResetDACL(lHive, sKey, bUseWow64, False) Then
            lret = RegCreateKeyEx(lHive, StrPtr(sKey), 0&, ByVal 0&, 0&, KEY_CREATE_SUB_KEY Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), ByVal 0&, hKey, ByVal 0&)
        End If
    End If
    RegCreateKey = (ERROR_SUCCESS = lret)
    If hKey <> 0 Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegCreateKey", lHive & "," & sKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegGetString(lHive&, sKey$, sValue$, Optional bUseWow64 As Boolean) As String
    RegGetString = GetRegData(lHive, sKey, sValue, bUseWow64)  '-> redirection to common function, just in case REG type is wrong
End Function

Public Function RegGetDword(lHive&, sKey$, sValue$, Optional bUseWow64 As Boolean) As Long
    Dim tmp As String
    tmp = GetRegData(lHive, sKey, sValue, bUseWow64)  '-> redirection to common function, just in case REG type is wrong
    If IsNumeric(tmp) Then RegGetDword = CLng(tmp)
End Function

Public Function RegGetBinaryToString(lHive&, sKey$, sValue$, Optional bUseWow64 As Boolean) As String
    RegGetBinaryToString = GetRegData(lHive, sKey, sValue, bUseWow64)   '-> redirection to common function, just in case REG type is wrong
End Function

Public Function RegGetBinary(hHive As Long, ByVal KeyName As String, ByVal ValueName As String, Optional bUseWow64 As Boolean) As Byte()
    On Error GoTo ErrorHandler:

    Dim abData()     As Byte
    Dim cData        As Long
    Dim hKey         As Long
    Dim lret         As Long
    Dim ordType      As Long

    Call NormalizeKeyNameAndHiveHandle(hHive, KeyName)
    
    If ERROR_SUCCESS <> RegOpenKeyEx(hHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then Exit Function
    
    'get size of buffer needed
    lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, ByVal 0&, cData)
    
    If ERROR_SUCCESS <> lret And ERROR_MORE_DATA <> lret Then
        RegCloseKey hKey
        Exit Function
    End If
    
    If ordType = REG_BINARY Then
        If cData > 0 Then
            ReDim abData(cData - 1) As Byte
            lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, VarPtr(abData(0&)), cData)
            If lret = ERROR_SUCCESS Then
                RegGetBinary = abData()
            End If
        End If
    End If
    
    If hKey <> 0& Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modRegistry.RegGetBinary", "hHive:", hHive, KeyName & "\" & ValueName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function RegGetMultiSZ(hHive As Long, ByVal KeyName As String, ByVal ValueName As String, Optional bUseWow64 As Boolean) As String()
    On Error GoTo ErrorHandler:

    Dim cData        As Long
    Dim hKey         As Long
    Dim lret         As Long
    Dim sData        As String
    Dim ordType      As Long
    Dim aData()      As String
    
    Call NormalizeKeyNameAndHiveHandle(hHive, KeyName)
    
    If ERROR_SUCCESS <> RegOpenKeyEx(hHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then Exit Function
    
    'get size of buffer needed
    lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, ByVal 0&, cData)
    
    If ERROR_SUCCESS <> lret And ERROR_MORE_DATA <> lret Then
        RegCloseKey hKey
        Exit Function
    End If
    
    '1 nul (2-byte) -> param is empty
    '1 character + 1 nul (2 byte + 2 byte)
    '2 characters + 2 nul (total: 8 bytes)
    'multiline: 1 char. + 1 nul + 1 char. + 2 nul. (10 bytes)
    
    If ordType = REG_MULTI_SZ Then
        If cData > 2 Then
            sData = String$(cData \ 2, vbNullChar)
            
            lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, StrPtr(sData), cData)
            If lret = ERROR_SUCCESS Then
                Do While AscW(Right$(sData, 1)) = 0
                    sData = Left$(sData, Len(sData) - 1)
                Loop
                RegGetMultiSZ = Split(sData, vbNullChar) 'struct is: item1 nul item2 nul ... itemN nul nul
            End If
        End If
    End If
    
    If hKey <> 0& Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modRegistry.RegGetMultiSZ", "hHive:", hHive, KeyName & "\" & ValueName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function GetRegData(hHive As Long, ByVal KeyName As String, ByVal ValueName As String, Optional bUseWow64 As Boolean) As Variant
    On Error GoTo ErrorHandler:
    
    Dim abData()     As Byte
    Dim cData        As Long
    Dim hKey         As Long
    Dim lData        As Long
    Dim qData        As Currency
    Dim lret         As Long
    Dim ordType      As Long
    Dim iPos         As Long
    Dim sData        As String
    Dim vValue       As Variant
    
    Call NormalizeKeyNameAndHiveHandle(hHive, KeyName)
    
    If ERROR_SUCCESS <> RegOpenKeyEx(hHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then Exit Function
    
    'get size of buffer needed
    lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, ByVal 0&, cData)
    
    If ERROR_SUCCESS <> lret And ERROR_MORE_DATA <> lret Then
        RegCloseKey hKey
        Exit Function
    End If
    
    Select Case ordType
        
        Case REG_SZ
            If cData > 1 Then
                'RegGetValue - Win 2003 SP1 +
'                sData = String$(cData - 1&, vbNullChar)
'                lret = RegGetValue(hKey, ByVal 0&, StrPtr(ValueName), RRF_RT_ANY, ByVal 0&, StrPtr(sData), cData)
'
'                If lret = ERROR_MORE_DATA Then
'                    sData = String$(cData - 1&, vbNullChar)
'                    lret = RegGetValue(hKey, ByVal 0&, StrPtr(ValueName), RRF_RT_ANY, ByVal 0&, StrPtr(sData), cData)
'                End If

                sData = String$(cData \ 2 + 1, vbNullChar)  ' (this API doesn't ensure that result buffer will contain null char, so I'll add extra 2 bytes)
                
                lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, StrPtr(sData), cData)
                If lret = ERROR_SUCCESS Then
                    vValue = Left$(sData, lstrlen(StrPtr(sData)))
                End If
            End If
            
        Case REG_EXPAND_SZ
            If cData > 1 Then
            
                'RegGetValue - Win 2003 SP1 +
'                sData = String$(cData - 1&, vbNullChar)
'                lret = RegGetValue(hKey, ByVal 0&, StrPtr(ValueName), RRF_RT_ANY Or RRF_NOEXPAND, ByVal 0&, StrPtr(sData), cData)
'
'                'Note: if you don't set RRF_NOEXPAND flag, you should prepare a bit bigger buffer in case of REG_EXPAND_SZ type,
'                'because it must be large enought to get string with expanded environment variables. And also anyway, be ready to get ERROR_MORE_DATA error.
'
'                If lret = ERROR_MORE_DATA Then
'                    sData = String$(cData - 1&, vbNullChar)
'                    lret = RegGetValue(hKey, ByVal 0&, StrPtr(ValueName), RRF_RT_ANY Or RRF_NOEXPAND, ByVal 0&, StrPtr(sData), cData)
'                End If
                
                sData = String$(cData \ 2 + 1, vbNullChar)  ' (this API doesn't ensure that result buffer will contain null char, so I'll add extra 2 bytes)
                
                lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, StrPtr(sData), cData)
                If lret = ERROR_SUCCESS Then
                    vValue = ExpandEnvStr(Left$(sData, lstrlen(StrPtr(sData))))
                End If
            End If
        
        Case REG_MULTI_SZ
            '//TODO: Check it: https://msdn.microsoft.com/en-us/library/windows/desktop/aa365240(v=vs.85).aspx
            'MSDN Note: Although \0\0 is technically not allowed in a REG_MULTI_SZ node, it can because the file is considered to be renamed to a null name.
        
            If cData > 2 Then
                'RegGetValue - Win 2003 SP1 +
'                sData = String$(cData - 1&, vbNullChar)
'                lret = RegGetValue(hKey, ByVal 0&, StrPtr(ValueName), RRF_RT_ANY, ByVal 0&, StrPtr(sData), cData)
                
                sData = String$(cData \ 2, vbNullChar)
                
                lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, StrPtr(sData), cData)
                If lret = ERROR_SUCCESS Then
                    If Right$(sData, 2) = vbNullChar & vbNullChar Then  'struct is: item1 nul item2 nul ... itemN nul nul
                        vValue = Left$(sData, Len(sData) - 2)
                    ElseIf Right$(sData, 1) = vbNullChar Then           'struct is: item1 nul
                        vValue = Left$(sData, Len(sData) - 1)
                    End If
                End If
            End If
        
        Case REG_DWORD, REG_DWORDLittleEndian
            cData = 4 'strict size
            lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, VarPtr(lData), cData)
            If lret = ERROR_SUCCESS Then
                vValue = lData
            End If
        
        Case REG_DWORDBigEndian
            cData = 4 'strict size
            lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, VarPtr(lData), cData)
            If lret = ERROR_SUCCESS Then
                vValue = SwapEndian(lData)
            End If
        
        Case REG_QWORD, REG_QWORD_LITTLE_ENDIAN
            cData = 8 'strict size
            lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, VarPtr(qData), cData)
            If lret = ERROR_SUCCESS Then
                vValue = qData
            End If
        
        Case Else ' other types -> byte data -> string
            If cData > 0 Then
                ReDim abData(cData - 1) As Byte
                lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, VarPtr(abData(0&)), cData)
                If lret = ERROR_SUCCESS Then
                    vValue = StrConv(abData, vbUnicode)
                End If
            End If
    
    End Select
    GetRegData = vValue
    If hKey <> 0& Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetRegData", "hHive:", hHive, KeyName & "\" & ValueName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function GetRegValuesAndData(lHive As Long, ByVal KeyName As String, _
    uRegTypesToQuery As FLAG_REG_TYPE, _
    aValueNames() As String, _
    aDataValues() As Variant, _
    aTypes() As Long, _
    Optional bExpandValues As Boolean = True, _
    Optional bUseWow64 As Boolean) As Long
    
    On Error GoTo ErrorHandler
    
    Dim cNameMax     As Long
    Dim cDataMax     As Long
    Dim CurValueN    As Long
    Dim hKey         As Long
    Dim lIndex       As Long
    Dim lNameSize    As Long
    Dim lDataSize    As Long
    Dim lret         As Long
    Dim hHive        As Long
    Dim iPos         As Long
    Dim sName        As String
    Dim ValuesCnt    As Long
    Dim lType        As Long
    Dim bData()      As Byte
    Dim lData        As Long
    Dim sData        As String
    Dim qData        As Currency
    
    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    lret = RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    
    If lret <> ERROR_SUCCESS Then
    
        ReDim aValueNames(0)
        ReDim aDataValues(0)
        ReDim aTypes(0)
        Exit Function
    
    Else
        
        lret = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, ByVal 0&, ByVal 0&, ByVal 0&, ValuesCnt, cNameMax, cDataMax, ByVal 0&, ByVal 0&)
        
        If lret <> ERROR_SUCCESS Then ValuesCnt = 1
        
        If cNameMax = 0 Then cNameMax = MAX_VALUENAME
        If cDataMax = 0 Then cDataMax = MAX_VALUENAME
        
        If ValuesCnt > 0 Then
        
          cNameMax = cNameMax + 1   'Nul
          cDataMax = cDataMax + 2   '2x Nul (REG_MULTI_SZ)
        
          ReDim aValueNames(1& To ValuesCnt)
          ReDim aDataValues(1& To ValuesCnt)
          ReDim aTypes(1& To ValuesCnt)
        
          Do
            
            lNameSize = cNameMax
            lDataSize = cDataMax
            
            sName = String$(lNameSize, vbNullChar)
            ReDim bData(lDataSize - 1)
            
            lret = RegEnumValue(hKey, lIndex, StrPtr(sName), lNameSize, 0&, lType, VarPtr(bData(0)), lDataSize)
            
            If lret = ERROR_MORE_DATA Then
                lNameSize = MAX_VALUENAME
                
                sName = String$(lNameSize, vbNullChar)
                ReDim bData(lDataSize - 1)
                
                lret = RegEnumValue(hKey, lIndex, StrPtr(sName), lNameSize, 0&, lType, VarPtr(bData(0)), lDataSize)
            End If
            
            If (lret = ERROR_SUCCESS) And ((2 ^ lType) And uRegTypesToQuery) Then
            
                If lDataSize <> 0 Then ReDim Preserve bData(lDataSize - 1)
            
                sName = Left$(sName, lstrlen(StrPtr(sName)))
                
                CurValueN = CurValueN + 1&
                
                If CurValueN > ValuesCnt Then
                    ReDim Preserve aValueNames(1& To CurValueN)
                    ReDim Preserve aDataValues(1& To CurValueN)
                    ReDim Preserve aTypes(1& To CurValueN)
                End If
                
                aValueNames(CurValueN) = sName
                aTypes(CurValueN) = lType
                
                If lDataSize > 0 Then
                
                    Select Case lType
                
                    Case REG_SZ
                        sData = String$(lDataSize \ 2, vbNullChar)
                        memcpy ByVal StrPtr(sData), ByVal VarPtr(bData(0)), lDataSize
                        aDataValues(CurValueN) = Left$(sData, lstrlen(StrPtr(sData)))
                    
                    Case REG_EXPAND_SZ
                        sData = String$(lDataSize \ 2, vbNullChar)
                        memcpy ByVal StrPtr(sData), ByVal VarPtr(bData(0)), lDataSize
                        If bExpandValues Then
                            aDataValues(CurValueN) = ExpandEnvStr(Left$(sData, lstrlen(StrPtr(sData))))
                        End If
                    
                    Case REG_MULTI_SZ
                        sData = String$(lDataSize \ 2, vbNullChar)
                        memcpy ByVal StrPtr(sData), ByVal VarPtr(bData(0)), lDataSize
                        If Right$(sData, 2) = vbNullChar & vbNullChar Then  'struct is: item1 nul item2 nul ... itemN nul nul
                            aDataValues(CurValueN) = Left$(sData, Len(sData) - 2)
                        ElseIf Right$(sData, 1) = vbNullChar Then           'struct is: item1 nul
                            aDataValues(CurValueN) = Left$(sData, Len(sData) - 1)
                        End If
                    
                    Case REG_DWORD, REG_DWORDLittleEndian
'                        GetMem4 ByVal VarPtr(bData(0)), ByVal VarPtr(lData)
'                        aDataValues(CurValueN) = lData
                        aDataValues(CurValueN) = bData(0) + bData(1) * 256&
                        
                    Case REG_DWORDBigEndian
                        'GetMem4 ByVal VarPtr(bData(0)), ByVal VarPtr(lData)
                        lData = bData(0) + bData(1) * 256&
                        aDataValues(CurValueN) = SwapEndian(lData)
                        
                    Case REG_QWORD, REG_QWORD_LITTLE_ENDIAN
                        GetMem8 ByVal VarPtr(bData(0)), ByVal VarPtr(qData)
                        aDataValues(CurValueN) = qData                          '4 zeros after dot
                        
                    Case Else   'binary -> string
                        aDataValues(CurValueN) = StrConv(bData, vbUnicode)
                
                    End Select
                End If
            End If
            
            lIndex = lIndex + 1
        
          Loop While lret = ERROR_SUCCESS
        End If
    
    End If
    
    If CurValueN < ValuesCnt Then
        ReDim Preserve aValueNames(1& To CurValueN)
        ReDim Preserve aDataValues(1& To CurValueN)
        ReDim Preserve aTypes(1& To CurValueN)
    End If
    
    If (hKey <> 0&) Then RegCloseKey hKey
    GetRegValuesAndData = CurValueN

    Exit Function
ErrorHandler:
    ErrorMsg Err, "modRegistry.GetRegValues", "lHive:", lHive, "Key:", KeyName, "uTypesToQuery:", uRegTypesToQuery, "bUseWow64:", bUseWow64
    If (hKey <> 0&) Then RegCloseKey hKey
End Function


Public Function RegSetStringVal(lHive&, ByVal sKey$, sValue$, sData$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, ret&, lret&
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not RegKeyExists(lHive, sKey, bUseWow64) Then
        RegCreateKey lHive, sKey, bUseWow64
    End If
    lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    
    If lret <> ERROR_SUCCESS Then
        If lret = ERROR_ACCESS_DENIED Then
            modPermissions.RegKeyResetDACL lHive, sKey, bUseWow64, False
            lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
            If lret <> ERROR_SUCCESS Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    If hKey <> 0 Then
        If Len(sData) = 0 Then
            RegSetStringVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_SZ, ByVal 0&, 0&))
        Else
            RegSetStringVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_SZ, ByVal StrPtr(sData), Len(sData) * 2 + 2))
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegSetStringVal", lHive & "," & sKey & "\" & sValue & "," & sData
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegSetExpandStringVal(lHive&, ByVal sKey$, sValue$, sData$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not RegKeyExists(lHive, sKey, bUseWow64) Then
        RegCreateKey lHive, sKey, bUseWow64
    End If
    lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    If lret <> ERROR_SUCCESS Then
        If lret = ERROR_ACCESS_DENIED Then
            modPermissions.RegKeyResetDACL lHive, sKey, bUseWow64, False
            lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
            If lret <> ERROR_SUCCESS Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    RegSetExpandStringVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_EXPAND_SZ, ByVal StrPtr(sData), Len(sData) * 2 + 2))
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegSetExpandStringVal", lHive & "," & sKey & "\" & sValue & "," & sData
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegSetDwordVal(lHive&, ByVal sKey$, sValue$, lData&, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not RegKeyExists(lHive, sKey, bUseWow64) Then
        RegCreateKey lHive, sKey, bUseWow64
    End If
    lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    If lret <> ERROR_SUCCESS Then
        If lret = ERROR_ACCESS_DENIED Then
            modPermissions.RegKeyResetDACL lHive, sKey, bUseWow64, False
            lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
            If lret <> ERROR_SUCCESS Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    RegSetDwordVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_DWORD, lData, 4&))
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegSetDwordVal", lHive & "," & sKey & "\" & sValue & "," & lData
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegSetBinaryVal(lHive&, ByVal sKey$, sValue$, aData() As Byte, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not RegKeyExists(lHive, sKey, bUseWow64) Then
        RegCreateKey lHive, sKey, bUseWow64
    End If
    lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    If lret <> ERROR_SUCCESS Then
        If lret = ERROR_ACCESS_DENIED Then
            modPermissions.RegKeyResetDACL lHive, sKey, bUseWow64, False
            lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
            If lret <> ERROR_SUCCESS Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    RegSetBinaryVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_BINARY, ByVal VarPtr(aData(0)), UBound(aData) + 1))
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegSetBinaryVal", lHive & "," & sKey & "\" & sValue & ",UBound=" & UBound(aData)
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegDelVal(lHive&, ByVal sKey$, sValue$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    If lret <> ERROR_SUCCESS Then
        If lret = ERROR_ACCESS_DENIED Then
            modPermissions.RegKeyResetDACL lHive, sKey, bUseWow64, False
            lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_SET_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
            If lret <> ERROR_SUCCESS Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    RegDelVal = (ERROR_SUCCESS = RegDeleteValue(hKey, StrPtr(sValue)))
    If hKey <> 0 Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegDelVal", lHive & "," & sKey & "\" & sValue
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegDelKey(lHive&, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''   Function   ''  Can recursively    ''  Support flag    '' Must close  ''   Minimum    ''
    '''     name     '' delete all subkeys  '' KEY_WOW64_64KEY  '' all handles ''  OS support  ''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SHDeleteKey              Yes             Partially (Win7+)        ?         Win 2000
    ' RegDeleteKey             No                   No                Yes !       Win 2000
    ' RegDeleteKeyEx           No                   Yes               Yes !       Win XP x64+
    ' RegDeleteTree            Yes                  Yes                 ?         Win Vista !
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' So, thank you, M$ ^_^ ):. We should write recursive function on one's own.
    
    Dim lret&
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
       
    If OSver.Bitness = "x32" And Not OSver.bIsVistaOrLater Then
        'XP x32
        SHDeleteKey lHive, StrPtr(sKey)
        If RegKeyExists(lHive, sKey, True) Then
        
            If modPermissions.RegKeyResetDACL(lHive, sKey, True, True) Then
                SHDeleteKey lHive, StrPtr(sKey)
            End If
        End If
    Else
        RegDeleteAllSubKeys lHive, sKey, bUseWow64
        
        lret = RegDeleteKeyEx(lHive, StrPtr(sKey), bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64, 0&)
        
        If lret = ERROR_ACCESS_DENIED Then   'root key self
        
            If modPermissions.RegKeyResetDACL(lHive, sKey, bUseWow64, False) Then
                RegDeleteKeyEx lHive, StrPtr(sKey), bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64, 0&
            End If
        End If
        
    End If
    RegDelKey = Not RegKeyExists(lHive, sKey, bUseWow64)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegDelKey", lHive & "," & sKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegDeleteAllSubKeys(lHive&, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    'del subkeys without root
    On Error GoTo ErrorHandler:
    Dim i&, sSubKey$, aSubKeys() As String, Flag As Boolean, lret&
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    Flag = True
    
    For i = 1 To RegEnumSubkeysToArray(lHive, sKey, aSubKeys(), bUseWow64, True)
        sSubKey = sKey & "\" & aSubKeys(i)
        Flag = Flag And RegDeleteAllSubKeys(lHive, sSubKey, bUseWow64)   '< --- recursively
        lret = RegDeleteKeyEx(lHive, StrPtr(sSubKey), bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64, 0&)
        
        If lret = ERROR_ACCESS_DENIED Then
            If modPermissions.RegKeyResetDACL(lHive, sSubKey, bUseWow64, False) Then
                lret = RegDeleteKeyEx(lHive, StrPtr(sSubKey), bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64, 0&)
            End If
        End If
        If lret <> ERROR_SUCCESS Then Flag = False
    Next i
    RegDeleteAllSubKeys = Flag
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegDeleteAllSubKeys", lHive & "," & sKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegKeyExists(lHive&, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If ERROR_SUCCESS = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, WRITE_OWNER Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        RegKeyExists = True
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegKeyExists", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegValueExists(lHive&, ByVal sKey$, sValue$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, WRITE_OWNER Or KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) <> 0 Then Exit Function
    RegValueExists = (ERROR_SUCCESS = RegQueryValueEx(hKey, StrPtr(sValue), 0&, ByVal 0&, ByVal 0&, ByVal 0&))
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegValueExists", lHive & "," & sKey & "\" & sValue
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function RegKeyHasSubKeys(lHive&, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, sBuf$, cbBuf&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    RegKeyHasSubKeys = False
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) = 0& Then
        cbBuf = MAX_KEYNAME
        sBuf = String(cbBuf, vbNullChar)
        If RegEnumKeyEx(hKey, 0&, StrPtr(sBuf), cbBuf, 0&, ByVal 0&, ByVal 0&, ByVal 0&) = 0& Then
            RegKeyHasSubKeys = True
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegKeyHasSubKeys", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegGetFirstSubKey(lHive&, ByVal sKey$, Optional bUseWow64 As Boolean) As String
    On Error GoTo ErrorHandler:
    Dim hKey&, sName$, cbBuf&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) = 0& Then
        cbBuf = MAX_KEYNAME
        sName = String(cbBuf, vbNullChar)
        If ERROR_SUCCESS = RegEnumKeyEx(hKey, 0&, StrPtr(sName), cbBuf, 0&, ByVal 0&, ByVal 0&, ByVal 0&) Then
            RegGetFirstSubKey = Left$(sName, lstrlen(StrPtr(sName)))
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegGetFirstSubKey", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegKeyHasValues(lHive&, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, sName$, cbBuf&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) = 0& Then
        cbBuf = MAX_VALUENAME
        sName = String$(cbBuf, vbNullChar)
        If RegEnumValue(hKey, 0&, StrPtr(sName), cbBuf, 0&, ByVal 0&, ByVal 0&, ByVal 0&) = 0& Then
            RegKeyHasValues = True
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegKeyHasValues", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegEnumSubKeys(lHive&, ByVal sKey$, Optional bUseWow64 As Boolean, Optional ForceUnlock As Boolean) As String
    On Error GoTo ErrorHandler:
    Dim hKey        As Long
    Dim sName       As String
    Dim sDummy      As String
    Dim lret        As Long
    Dim cMaxSubKeyLen As Long
    Dim lNameSize   As Long
    Dim idx         As Long
    Dim SubkeysCnt  As Long
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    
    If lret = ERROR_ACCESS_DENIED Then
        lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
        
        If lret = ERROR_ACCESS_DENIED And ForceUnlock Then
        
            If modPermissions.RegKeyResetDACL(lHive, sKey, bUseWow64, False) Then
            
                lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
            End If
        End If
    End If
    
    If lret = ERROR_SUCCESS Then
    
        lret = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, SubkeysCnt, cMaxSubKeyLen, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
        
        If Not (lret = ERROR_SUCCESS And SubkeysCnt = 0) Then
        
          If cMaxSubKeyLen = 0 Then
            cMaxSubKeyLen = MAX_KEYNAME
          Else
            cMaxSubKeyLen = cMaxSubKeyLen + 1
          End If
        
          Do
            lNameSize = cMaxSubKeyLen
            sName = String$(lNameSize, vbNullChar)
            
            lret = RegEnumKeyEx(hKey, idx, StrPtr(sName), lNameSize, 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            
            If lret = ERROR_MORE_DATA Then
                cMaxSubKeyLen = cMaxSubKeyLen * 2&
                
                lNameSize = cMaxSubKeyLen
                
                sName = String$(lNameSize, vbNullChar)
                
                lret = RegEnumKeyEx(hKey, idx, StrPtr(sName), lNameSize, 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            End If
            
            If (lret = ERROR_SUCCESS) Then
                sName = Left$(sName, lstrlen(StrPtr(sName)))
                sDummy = sDummy & sName & "|"
            End If
            
            idx = idx + 1
          Loop While lret = ERROR_SUCCESS
        End If
    End If
    
    If hKey <> 0 Then RegCloseKey hKey
    If Len(sDummy) <> 0 Then RegEnumSubKeys = Left$(sDummy, Len(sDummy) - 1)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegEnumSubkeys", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function RegEnumSubkeysToArray(lHive&, ByVal sKey$, aSubKeys() As String, Optional bUseWow64 As Boolean, Optional ForceUnlock As Boolean) As Long
    On Error GoTo ErrorHandler:
    Dim hKey        As Long
    Dim sName       As String
    Dim lret        As Long
    Dim cMaxSubKeyLen As Long
    Dim lNameSize   As Long
    Dim idx         As Long
    Dim SubkeysCnt  As Long
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    
    If lret = ERROR_ACCESS_DENIED Then
        lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
        
        If lret = ERROR_ACCESS_DENIED And ForceUnlock Then
        
            If modPermissions.RegKeyResetDACL(lHive, sKey, bUseWow64, False) Then
            
                lret = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
            End If
        End If
    End If
    
    If lret = ERROR_SUCCESS Then
    
        lret = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, SubkeysCnt, cMaxSubKeyLen, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
        
        If Not (lret = ERROR_SUCCESS And SubkeysCnt = 0) Then
        
          If lret = ERROR_SUCCESS Then
            ReDim aSubKeys(1 To SubkeysCnt)
          Else
            ReDim aSubKeys(1 To 100)
          End If
        
          If cMaxSubKeyLen = 0 Then
            cMaxSubKeyLen = MAX_KEYNAME
          Else
            cMaxSubKeyLen = cMaxSubKeyLen + 1
          End If
        
          Do
            lNameSize = cMaxSubKeyLen
            sName = String$(lNameSize, vbNullChar)
            
            lret = RegEnumKeyEx(hKey, idx, StrPtr(sName), lNameSize, 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            
            If lret = ERROR_MORE_DATA Then
                cMaxSubKeyLen = cMaxSubKeyLen * 2&
                
                lNameSize = cMaxSubKeyLen
                
                sName = String$(lNameSize, vbNullChar)
                
                lret = RegEnumKeyEx(hKey, idx, StrPtr(sName), lNameSize, 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            End If
            
            idx = idx + 1
            
            If (lret = ERROR_SUCCESS) Then
                sName = Left$(sName, lstrlen(StrPtr(sName)))
                
                If UBound(aSubKeys) < idx Then ReDim Preserve aSubKeys(UBound(aSubKeys) + 100)
                aSubKeys(idx) = sName
            Else
                If idx > 1 Then
                    ReDim Preserve aSubKeys(1 To idx - 1)
                    RegEnumSubkeysToArray = idx - 1
                End If
            End If
            
          Loop While lret = ERROR_SUCCESS
        End If
    End If
    
    If hKey <> 0 Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegEnumSubkeys", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function GetEnumValues(lHive&, ByVal KeyName$, Optional bUseWow64 As Boolean) As String
    On Error GoTo ErrorHandler:
    Dim iNameMax     As Long
    Dim hKey         As Long
    Dim idx          As Long
    Dim lNameSize    As Long
    Dim lret         As Long
    Dim sName        As String
    Dim sDummy       As String
    Dim iValueMax    As Long
    Dim ValuesCnt    As Long
    
    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    If ERROR_SUCCESS = RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        
        lret = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, ByVal 0&, ByVal 0&, ByVal 0&, ValuesCnt, iNameMax, ByVal 0&, ByVal 0&, ByVal 0&)
        
        If Not (lret = ERROR_SUCCESS And ValuesCnt = 0) Then
        
          If iNameMax = 0 Then
            iNameMax = MAX_VALUENAME
          Else
            iNameMax = iNameMax + 1
          End If
        
          Do
            lNameSize = iNameMax
            sName = String$(lNameSize, vbNullChar)
            
            lret = RegEnumValue(hKey, idx, StrPtr(sName), lNameSize, 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            
            If lret = ERROR_MORE_DATA Then
                iNameMax = MAX_VALUENAME
                
                lNameSize = iNameMax
                
                sName = String$(lNameSize, vbNullChar)
                
                lret = RegEnumValue(hKey, idx, StrPtr(sName), lNameSize, 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            End If
            
            If (lret = ERROR_SUCCESS Or lret = ERROR_MORE_DATA) Then
                sName = Left$(sName, lstrlen(StrPtr(sName)))
                sDummy = sDummy & sName & "|"
            End If
            
            idx = idx + 1
          Loop While lret = ERROR_SUCCESS
        End If
    End If
    If (hKey <> 0&) Then RegCloseKey hKey
    If Len(sDummy) <> 0 Then GetEnumValues = Left$(sDummy, Len(sDummy) - 1)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetEnumValues", lHive & "," & KeyName
    If (hKey <> 0&) Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function GetEnumValuesToArray(lHive&, ByVal KeyName$, aValues() As String, Optional bUseWow64 As Boolean) As Long
    On Error GoTo ErrorHandler:
    Dim iNameMax     As Long
    Dim hKey         As Long
    Dim idx          As Long
    Dim lNameSize    As Long
    Dim lret         As Long
    Dim sName        As String
    Dim sDummy       As String
    Dim iValueMax    As Long
    Dim ValuesCnt    As Long
    
    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    If ERROR_SUCCESS = RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        
        lret = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, ByVal 0&, ByVal 0&, ByVal 0&, ValuesCnt, iNameMax, ByVal 0&, ByVal 0&, ByVal 0&)
        
        If Not (lret = ERROR_SUCCESS And ValuesCnt = 0) Then
        
          If lret = ERROR_SUCCESS Then
            ReDim aValues(1 To ValuesCnt)
          Else
            ReDim aValues(1 To 100)
          End If
        
          If iNameMax = 0 Then
            iNameMax = MAX_VALUENAME
          Else
            iNameMax = iNameMax + 1
          End If
        
          Do
            lNameSize = iNameMax
            sName = String$(lNameSize, vbNullChar)
            
            lret = RegEnumValue(hKey, idx, StrPtr(sName), lNameSize, 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            
            If lret = ERROR_MORE_DATA Then
                iNameMax = MAX_VALUENAME
                
                lNameSize = iNameMax
                
                sName = String$(lNameSize, vbNullChar)
                
                lret = RegEnumValue(hKey, idx, StrPtr(sName), lNameSize, 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            End If
            
            idx = idx + 1
            
            If (lret = ERROR_SUCCESS Or lret = ERROR_MORE_DATA) Then
                sName = Left$(sName, lstrlen(StrPtr(sName)))
                If UBound(aValues) < idx Then ReDim Preserve aValues(UBound(aValues) + 100)
                aValues(idx) = sName
            Else
                If idx > 1 Then
                    ReDim Preserve aValues(1 To idx - 1)
                    GetEnumValuesToArray = idx - 1
                End If
            End If
            
          Loop While lret = ERROR_SUCCESS
        End If
    End If
    If (hKey <> 0&) Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetEnumValues", lHive & "," & KeyName
    If (hKey <> 0&) Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function GetRegKeyTime(lHive&, ByVal KeyName$, Optional bUseWow64 As Boolean) As Date
    On Error GoTo ErrorHandler:
    Dim hKey         As Long
    Dim lret         As Long
    Dim ftime        As FILETIME
    Dim stime        As SYSTEMTIME
    Dim DateTime     As Date

    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    If ERROR_SUCCESS = RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        
        If ERROR_SUCCESS = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ftime) Then
        
            lret = FileTimeToLocalFileTime(ftime, ftime)    ' учитываем смещение согласно текущим региональным настройкам в системе
            lret = FileTimeToSystemTime(ftime, stime)       ' FILETIME -> SYSTEMTIME
        
            'GetRegKeyTime = DateSerial(fTime.wYear, fTime.wMonth, fTime.wDay) + TimeSerial(fTime.wHour, fTime.wMinute, fTime.wSecond)
            
            SystemTimeToVariantTime stime, DateTime         ' SYSTEMTIME -> Date
            GetRegKeyTime = DateTime
        
        End If
        
    End If
    
    If (hKey <> 0&) Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetRegKeyTime", lHive & "," & KeyName
    If (hKey <> 0&) Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

'it makes registry export (temporary .reg-file), then read its contents to variable
Public Function RegExportKeyToVariable( _
    lHive&, _
    ByVal sKey$, _
    Optional bUseWow64 As Boolean, _
    Optional SaveHeader As Boolean = True, _
    Optional ConvertToANSI As Boolean = False) As String
    
    On Error GoTo ErrorHandler:

    Dim sTempFile$, sData$, hFile&, lret&, FSize@, isUnicode As Boolean, pos&

    Randomize
    sTempFile$ = BuildPath(AppPath(), "backups\backup_" & Int(Rnd * 10000) & ".reg")
    
    If lHive <> 0 Then
        sKey = GetHiveNameByHandle(lHive) & "\" & sKey
    End If
    
    If FileExists(sWinSysDir & "\reg.exe", True) Then
        
        If Proc.ProcessRun(sWinSysDir & "\reg.exe", "export " & """" & sKey & """" & " " & """" & sTempFile & """" & " /y" & IIf(bUseWow64, "", " /reg:64"), , vbHide) Then

            If ERROR_SUCCESS <> Proc.WaitForTerminate(, , False, 10000) Then    '10 sec. timeout
                Proc.ProcessClose , , True
            End If
            
            If FileExists(sTempFile) Then
                    
                If OpenW(sTempFile, FOR_READ, hFile) Then
                    
                    FSize = LOFW(hFile)
                    
                    If FSize > 0 Then
                        
                        sData = String$(FSize, 0)
                        
                        If GetW(hFile, 1, sData) Then
                            
                            isUnicode = (Left$(sData, 2) = "яю")
                            
                            'adding 2 extra CrLf for future purposes - if we would like to concat. several export. variables
                            If isUnicode Then
                                sData = sData & vbCr & vbNullChar & vbLf & vbNullChar & vbCr & vbNullChar & vbLf & vbNullChar
                            Else
                                sData = sData & vbCrLf & vbCrLf
                            End If
                            
                            If isUnicode Then
                                If ConvertToANSI Then
                                    sData = StrConv(Mid$(sData, 3), vbFromUnicode)
                                    isUnicode = False
                                End If
                            End If
                            
                            If Not SaveHeader Then
                                If isUnicode Then
                                    pos = InStr(sData, vbCr & vbNullChar & vbLf & vbNullChar)
                                    If pos <> 0 Then
                                        sData = Mid$(sData, pos + 4)
                                    End If
                                Else
                                    pos = InStr(sData, vbCrLf)
                                    If pos <> 0 Then
                                        sData = Mid$(sData, pos + 2)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    CloseW hFile
                End If
                DeleteFileWEx (StrPtr(sTempFile))
            End If
        End If
    End If
    RegExportKeyToVariable = sData
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegExportKeyToVariable", lHive & "," & sKey, "bUseWow64:", bUseWow64, "SaveHeader:", SaveHeader
    If inIDE Then Stop: Resume Next
End Function

Public Function RegGetFileFromBinary(lHive&, ByVal sKey$, sValue$, Optional bUseWow64 As Boolean) As String
    On Error GoTo ErrorHandler:
    
    Dim hKey&, sData$, sFile$, cbData&
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    cbData = MAX_VALUENAME
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) = 0 Then
        ReDim uData(MAX_VALUENAME - 1) As Byte
        If RegQueryValueEx(hKey, StrPtr(sValue), 0, ByVal 0&, ByVal VarPtr(uData(0)), cbData) = 0 Then
            sFile = RTrimNull(StrConv(uData, vbUnicode))
        End If
        RegCloseKey hKey
    End If
    RegGetFileFromBinary = sFile
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegGetFileFromBinary", lHive & "," & sKey, "sParam:", sValue
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function KeyExportToBinary(lHive&, ByVal KeyName$, destFile As String, Optional bUseWow64 As Boolean) As Boolean     'Save key to binary .hiv file
    On Error GoTo ErrorHandler:
    Dim hKey         As Long
    Dim lret         As Long

    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    If ERROR_SUCCESS = RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        
        If FileExists(destFile) Then DeleteFileWEx (StrPtr(destFile))
        
        'If SetCurrentProcessPrivileges("SeBackupPrivilege") Then
        ' (already defined in main form module)
            KeyExportToBinary = ERROR_SUCCESS = RegSaveKeyEx(hKey, StrPtr(destFile), ByVal 0&, IIf(OSver.MajorMinor < 5.1, REG_STANDARD_FORMAT, REG_LATEST_FORMAT))
        'End If
    End If
    
    If (hKey <> 0&) Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "KeyExportToBinary", lHive & "," & KeyName & "," & destFile
    If (hKey <> 0&) Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function IniGetString(sFile$, sSection$, sValue$, Optional bMultiple As Boolean = False) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "IniGetString - Begin", "File: " & sFile
    
    Dim sIniFile$, i&, iSect&, vContents As Variant, sData$, iAttr&, ff%
    Dim Redirect As Boolean, bOldStatus As Boolean
    
    If Not FileExists(sFile) Then
        If FileExists(sWinSysDir & "\" & sFile) Then
            sIniFile = sWinSysDir & "\" & sFile
        End If
        If FileExists(sWinDir & "\" & sFile) Then
            sIniFile = sWinDir & "\" & sFile
        End If
        If 0 = Len(sIniFile) Then Exit Function
        'If Not FileExists(sIniFile) Then Exit Function
    Else
        sIniFile = sFile
    End If
    If FileLenW(sIniFile) = 0 Then Exit Function
    
'    iAttr = GetFileAttributes(StrPtr(sIniFile))
'    If (iAttr And 2048) Then
'        iAttr = iAttr And Not 2048 'compression flag
'        SetFileAttributes StrPtr(sIniFile), iAttr
'    End If
     
    Redirect = ToggleWow64FSRedirection(False, sIniFile, bOldStatus)
    
    ff = FreeFile()
    Open sIniFile For Binary Access Read As #ff
        vContents = Split(Input(LOF(ff), #ff), vbCrLf)
    Close #ff
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    
    If UBound(vContents) = -1 Then Exit Function 'file is empty
    
'    For i = 0 To UBound(vContents)
'        If Trim$(vContents(i)) <> vbNullString Then
'            If InStr(1, LTrim$(vContents(i)), "[" & sSection & "]", vbTextCompare) = 1 Then
'                'found the correct section
'                iSect = i
'                Exit For
'            End If
'        End If
'    Next i
'    If i = UBound(vContents) + 1 Then Exit Function 'section not found
'
'    For i = iSect + 1 To UBound(vContents)
'        'check if new section was started
'        If InStr(LTrim$(vContents(i)), "[") = 1 Then Exit For 'value not found
'
'        If InStr(1, LTrim$(vContents(i)), sValue & "=", vbTextCompare) = 1 Then
'            'found the value!
'            IniGetString = Mid$(vContents(i), InStr(vContents(i), "=") + 1)
'            Exit Function
'        End If
'    Next i

    Do Until InStr(1, vContents(i), "[" & sSection & "]", vbTextCompare) = 1
        i = i + 1
        If i > UBound(vContents) Then Exit Function
    Loop
    i = i + 1
    Do Until Left$(vContents(i), 1) = "["
        If InStr(1, vContents(i), sValue, vbTextCompare) = 1 Then
            sData = sData & "|" & vContents(i)
            If Not bMultiple Then Exit Do
        End If
        i = i + 1
        If i > UBound(vContents) Then Exit Do
    Loop
    'IniGetString = Mid$(vContents(i), InStr(vContents(i), "=") + 1)
    If sData <> vbNullString Then
        If Not bMultiple Then
            IniGetString = Mid$(sData, InStr(sData, "=") + 1)
        Else
            IniGetString = Replace$(Mid$(sData, 2), "=", " = ")
        End If
    End If
    
    AppendErrorLogCustom "IniGetString - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modRegistry_IniGetString", "sFile=", sFile, "sSection=", sSection, "sValue=", sValue
    Close #ff
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function


Public Sub IniSetString(sFile$, sSection$, sValue$, sData$)
    On Error GoTo ErrorHandler:
    Dim sIniFile$, i&, iSect&, vContents As Variant, sNewData$, iAttr&, ff%
    Dim Redirect As Boolean, bOldStatus As Boolean
    If Not FileExists(sFile) Then
        If FileExists(sWinSysDir & "\" & sFile) Then
            sIniFile = sWinSysDir & "\" & sFile
        End If
        If FileExists(sWinDir & "\" & sFile) Then
            sIniFile = sWinDir & "\" & sFile
        End If
        If Not FileExists(sIniFile) Then Exit Sub
    Else
        sIniFile = sFile
    End If
    
    On Error Resume Next
    Redirect = ToggleWow64FSRedirection(False, sIniFile, bOldStatus)
    ff = FreeFile()
    Open sIniFile For Binary Access Read As #ff
    If Err.Number <> 0 Then
        TryUnlock sIniFile
        Err.Clear
        Open sIniFile For Binary Access Read As #ff
        If Err.Number <> 0 Then
            If Not bAutoLogSilent Then
                MsgBoxW "Failed to open the settings file '" & sIniFile & "'. Please verify " & _
                    "that read access is allowed to that file.", vbCritical
            End If
            If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
            Exit Sub
        End If
    End If
    On Error GoTo ErrorHandler:

    vContents = Split(Input(LOF(ff), #ff), vbCrLf)
    Close #ff
    iAttr = GetFileAttributes(StrPtr(sIniFile))
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)

    If UBound(vContents) = -1 Then 'file is empty
        sNewData = "[" & sSection & "]" & vbCrLf & sValue & "=" & sData
        WriteDataToFile sIniFile, sNewData, iAttr, True
        Exit Sub
    End If
    
    For i = 0 To UBound(vContents)
        If Trim$(vContents(i)) <> vbNullString Then
            If InStr(1, vContents(i), "[" & sSection & "]", vbTextCompare) = 1 Then
                'found the correct section
                iSect = i
                Exit For
            End If
        End If
    Next i
    If i = UBound(vContents) + 1 Then Exit Sub   'section not found
    
    For i = iSect + 1 To UBound(vContents)
        
        If InStr(vContents(i), "[") = 1 Or _
          InStr(1, vContents(i), sValue & "=", vbTextCompare) = 1 Then
        
            'value not found ("[" - mean next section)
            If InStr(vContents(i), "[") = 1 Then
                sNewData = sValue & "=" & sData
                vContents(i) = sNewData & vbCrLf & vContents(i)
            Else
                'found the value!
                sNewData = Left$(vContents(i), InStr(vContents(i), "=")) & sData
                vContents(i) = sNewData
            End If
            
            'input new data, replace file
            WriteDataToFile sIniFile, Join(vContents, vbCrLf), iAttr, True
            Exit Sub
        End If
    Next i
    
    'Last section, but no value
    sNewData = sValue & "=" & sData
    WriteDataToFile sIniFile, Join(vContents, vbCrLf) & vbCrLf & sNewData, iAttr, True
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modRegistry_IniSetString", "sFile=", sFile, "sSection=", sSection, "sValue=", sValue, "sData=", sData
    Close #ff
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Sub

Private Sub WriteDataToFile(sFile$, sContents$, iAttr&, Optional bShowWarning As Boolean)
    On Error GoTo ErrorHandler:
    Dim ff As Integer
    Dim Redirect As Boolean, bOldStatus As Boolean

    If 0 = DeleteFileWEx(StrPtr(sFile)) Then
        If Not bAutoLogSilent And bShowWarning Then
            'The value '[*]' could not be written to the settings file '[**]'. Please verify that write access is allowed to that file.
            MsgBoxW Replace$(Replace$(Translate(1008), "[*]", sContents), "[**]", sFile), vbCritical
        End If
        Exit Sub
    End If
    
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    ff = FreeFile()
    Open sFile For Output As #ff
        Print #ff, sContents
    Close #ff
    SetFileAttributes StrPtr(sFile), iAttr
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modRegistry_WriteDataToFile", "sFile=", sFile, "sContents=", sContents, "iAttr=", iAttr, "bShowWarning=", bShowWarning
    Close #ff
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CreateUninstallKey(bCreate As Boolean) ' if false -> delete registry entries
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CreateUninstallKey - Begin"
    Dim Setup_Key$:   Setup_Key = "Software\Microsoft\Windows\CurrentVersion\Uninstall\HiJackThis"
    If bNoWriteAccess Then bCreate = False
    If bCreate Then
        If RegGetString(HKEY_LOCAL_MACHINE, Setup_Key, "DisplayName") <> _
                   "HiJackThis " & App.Major & "." & App.Minor & "." & App.Revision Then
            RegCreateKey HKEY_LOCAL_MACHINE, Setup_Key
            RegSetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "DisplayName", "HiJackThis Fork " & App.Major & "." & App.Minor & "." & App.Revision
            RegSetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "UninstallString", """" & AppPath(True) & """ /uninstall"
            RegSetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "DisplayIcon", AppPath(True)
            RegSetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "DisplayVersion", App.Major & "." & App.Minor & "." & App.Revision
            RegSetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "Publisher", "Alex Dragokas"
            'RegSetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "URLInfoAbout", "http://www.spywareinfo.com/~merijn/"
            'RegSetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "URLInfoAbout", "https://sourceforge.net/projects/hjt/"
            RegSetStringVal HKEY_LOCAL_MACHINE, Setup_Key, "URLInfoAbout", "https://sourceforge.net/projects/hjt/"
        End If
    Else
        RegDelKey HKEY_LOCAL_MACHINE, Setup_Key
        RegDelKey HKEY_CURRENT_USER, "Software\Microsoft\Installer\Products\8A9C1670A3F861244B7A7BFAFB422AA4"
    End If
    AppendErrorLogCustom "CreateUninstallKey - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CreateUninstallKey", bCreate
    If inIDE Then Stop: Resume Next
End Sub


Public Function GetFirstSubFolder(sFolder$) As String
    On Error GoTo ErrorHandler:
    Dim sBla$
    Dim Redirect As Boolean, bOldStatus As Boolean
    
    Redirect = ToggleWow64FSRedirection(False, sFolder, bOldStatus)
    sBla = DirW$(sFolder & "\", vbAll, True)
    If Len(sBla) <> 0 Then
        GetFirstSubFolder = sBla
    End If
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modRegistry_GetFirstSubFolder", "sFolder=", sFolder
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function


Public Function ExtractFilename(sLine$) As String
    On Error GoTo ErrorHandler:
    'Parse rule:
    '
    '1) "1 1.exe" arg -> '1 1.exe'
    ' if path or name contains spaces they must be quoted, otherwise name can be truncated to first space (see example below).
    '
    '2) 1.exe arg -> '1.exe'
    '   1 1.exe arg -> '1 1.exe'
    '   1 1.cmd arg -> '1'
    '
    ' Note: '.exe ' - is a marker of the end of filename
    ' Note: This function does not remove path.
    '
    Dim s$, pos&, pos2&
    s = Trim$(sLine)
    If Left$(s, 1) = """" Then
        pos = InStr(2, s, """")
        If pos > 0 Then
            ExtractFilename = Mid$(s, 2, pos - 2) ' remove first and last quote
        Else
            ExtractFilename = Mid$(s, 2) 'no close quote... lol
        End If
    ' if there are no quote
    Else
        pos = InStr(1, s, ".exe ", vbTextCompare) ' mark -> '.exe' + space
        If pos Then
            ExtractFilename = Left$(s, pos + 3)
        Else
            pos = InStr(pos, s, " ")
            If pos > 0 Then
                ExtractFilename = Left$(s, pos - 1)
            Else
                ExtractFilename = s
            End If
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ExtractFilename", sLine
    If inIDE Then Stop: Resume Next
End Function


Public Function ExtractArguments(sLine$) As String
    On Error GoTo ErrorHandler:
    Dim s$, pos&
    s = Trim$(sLine)
    If Left$(s, 1) = """" Then
        pos = InStr(2, s, """")
        If pos > 0 Then
            ExtractArguments = Trim$(Mid$(s, pos + 1))
        Else
            ExtractArguments = vbNullString 'no close quote... lol
        End If
    Else
        pos = InStr(s, " ")
        If pos > 0 Then
            ExtractArguments = Trim$(Mid$(s, pos + 1))
        Else
            ExtractArguments = vbNullString
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ExtractArguments", sLine
    If inIDE Then Stop: Resume Next
End Function


Public Sub RegSaveHJT(sName$, sData$)
    On Error GoTo ErrorHandler:
    If sName Like "Ignore#*" Then
        Dim aData() As Byte
        aData = StrConv(sData, vbFromUnicode)
        RegSetBinaryVal HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", sName, aData
    Else
        RegSetStringVal HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", sName, sData
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RegSaveHJT", sName & "," & sData
    If inIDE Then Stop: Resume Next
End Sub


Public Function RegReadHJT(sName$, Optional sDefault$) As String
    On Error GoTo ErrorHandler:
    If sName Like "Ignore#*" Then
        RegReadHJT = RegGetBinaryToString(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", sName)
    Else
        RegReadHJT = RegGetString(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", sName)
    End If
    If Len(RegReadHJT) = 0 Then RegReadHJT = sDefault
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegReadHJT", sName & "," & sDefault
    If inIDE Then Stop: Resume Next
End Function


Public Sub RegDelHJT(sName$)
    RegDelVal HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", sName
End Sub

