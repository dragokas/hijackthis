VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsRegistry.cls]

Option Explicit
'
' Registry functions class by Alex Dragokas
'

'Revision 3.3 (05.08.2020)

'Private Const MAX_PATH       As Long = 260&
'Private Const MAX_PATH_W     As Long = 32767&
'Private Const MAX_KEYNAME    As Long = 255& 'https://msdn.microsoft.com/en-us/library/windows/desktop/ms724872(v=vs.85).aspx
'Private Const MAX_VALUENAME  As Long = 32767&

'Public Enum ENUM_REG_HIVE
'    HKEY_USER_SPECIFIED = 0
'    HKEY_CLASSES_ROOT = &H80000000
'    HKEY_CURRENT_USER = &H80000001
'    HKEY_LOCAL_MACHINE = &H80000002
'    HKEY_USERS = &H80000003
'    HKEY_PERFORMANCE_DATA = &H80000004
'    HKEY_CURRENT_CONFIG = &H80000005
'    HKEY_DYN_DATA = &H80000006
'    HKCR = &H80000000
'    HKCU = &H80000001
'    HKLM = &H80000002
'    HKU = &H80000003
'    HKPD = &H80000004
'    HKCC = &H80000005
'    HKDD = &H80000006
'End Enum
'#If False Then
'    Dim HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS, HKEY_USER_SPECIFIED
'    Dim HKCR, HKCU, HKLM, HKU
'#End If
'
'Public Enum REG_VALUE_TYPE
'    REG_NONE = 0&
'    REG_SZ = 1&
'    REG_EXPAND_SZ = 2&
'    REG_BINARY = 3&
'    REG_DWORD = 4&
'    REG_DWORDLittleEndian = 4&
'    REG_DWORDBigEndian = 5&
'    REG_LINK = 6&
'    REG_MULTI_SZ = 7&
'    REG_ResourceList = 8&
'    REG_FullResourceDescriptor = 9&
'    REG_ResourceRequirementsList = 10&
'    REG_QWORD = 11&
'    REG_QWORD_LITTLE_ENDIAN = 11&
'End Enum
'#If False Then
'    Dim REG_NONE, REG_SZ, REG_EXPAND_SZ, REG_BINARY, REG_DWORD, REG_DWORDLittleEndian, REG_DWORDBigEndian, REG_LINK, REG_MULTI_SZ, REG_ResourceList
'    Dim REG_FullResourceDescriptor, REG_ResourceRequirementsList, REG_QWORD, REG_QWORD_LITTLE_ENDIAN
'#End If
'
'Public Enum FLAG_REG_TYPE   'flags to be able to map bit mask and default registry type constants
'    FLAG_REG_ALL = -1&
'    FLAG_REG_NONE = 1&
'    FLAG_REG_SZ = 2&
'    FLAG_REG_EXPAND_SZ = 4&
'    FLAG_REG_BINARY = 8&
'    FLAG_REG_DWORD = &H10&
'    FLAG_REG_DWORDLittleEndian = &H10&
'    FLAG_REG_DWORDBigEndian = &H20&
'    FLAG_REG_LINK = &H40&
'    FLAG_REG_MULTI_SZ = &H80&
'    FLAG_REG_ResourceList = &H100&
'    FLAG_REG_FullResourceDescriptor = &H200&
'    FLAG_REG_ResourceRequirementsList = &H400&
'    FLAG_REG_QWORD = &H800&
'    FLAG_REG_QWORD_LITTLE_ENDIAN = &H1000&
'End Enum
'#If False Then
'    Dim FLAG_REG_ALL, FLAG_REG_NONE, FLAG_REG_SZ, FLAG_REG_EXPAND_SZ, FLAG_REG_BINARY, FLAG_REG_DWORD, FLAG_REG_DWORDLittleEndian, FLAG_REG_DWORDBigEndian
'    Dim FLAG_REG_LINK, FLAG_REG_MULTI_SZ, FLAG_REG_ResourceList, FLAG_REG_FullResourceDescriptor, FLAG_REG_ResourceRequirementsList, FLAG_REG_QWORD, FLAG_REG_QWORD_LITTLE_ENDIAN
'#End If

'Private Type SYSTEMTIME
'    wYear           As Integer
'    wMonth          As Integer
'    wDayOfWeek      As Integer
'    wDay            As Integer
'    wHour           As Integer
'    wMinute         As Integer
'    wSecond         As Integer
'    wMilliseconds   As Integer
'End Type
'
'Private Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type
'
'Private Type WIN32_FIND_DATA
'    dwFileAttributes        As Long
'    ftCreationTime          As FILETIME
'    ftLastAccessTime        As FILETIME
'    ftLastWriteTime         As FILETIME
'    nFileSizeHigh           As Long
'    nFileSizeLow            As Long
'    dwReserved0             As Long
'    dwReserved1             As Long
'    lpszFileName(MAX_PATH - 1) As Integer
'    lpszAlternate(13)       As Integer
'End Type
'
'Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyW" (ByVal hKey As Long, ByVal lpClass As Long, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
'Private Declare Function RegGetValue Lib "advapi32.dll" Alias "RegGetValueW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal lpValue As Long, ByVal dwFlags As Long, pdwType As Long, ByVal pvData As Long, pcbData As Long) As Long
'Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
'Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueW" (ByVal hKey As Long, ByVal lpValueName As Long) As Long
'Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
'Private Declare Function RegDeleteKeyEx Lib "advapi32.dll" Alias "RegDeleteKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal samDesired As Long, ByVal Reserved As Long) As Long
'Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
'Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long
'
'Private Declare Function SHDeleteKey Lib "Shlwapi.dll" Alias "SHDeleteKeyW" (ByVal lRootKey As Long, ByVal szKeyToDelete As Long) As Long
'Private Declare Function RegSaveKeyEx Lib "advapi32.dll" Alias "RegSaveKeyExW" (ByVal hKey As Long, ByVal lpFile As Long, ByVal lpSecurityAttributes As Long, ByVal flags As Long) As Long
'
'Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'Private Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" (lpSystemTime As SYSTEMTIME, vtime As Date) As Long
'
'Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long
'Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
'Private Declare Sub memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'Private Declare Function GetMem4 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long
'Private Declare Function GetMem8 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long
'
'Private Declare Function DispCallFunc Lib "oleaut32" (ByVal PPV As Long, ByVal oVft As Long, ByVal cc As Long, ByVal rtTYP As VbVarType, ByVal paCNT As Long, paTypes As Any, paValues As Any, ByRef fuReturn As Variant) As Long
'Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'
'Private Const CC_STDCALL As Long = 4
'
'Private Const INVALID_FILE_ATTRIBUTES As Long = -1&
'
'Private Const KEY_CREATE_SUB_KEY     As Long = &H4
'Private Const KEY_QUERY_VALUE        As Long = &H1
'Private Const KEY_SET_VALUE          As Long = &H2
'Private Const READ_CONTROL           As Long = &H20000
'Private Const WRITE_OWNER            As Long = &H80000
'Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
'Private Const SYNCHRONIZE            As Long = &H100000
'Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
'Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
'Private Const REG_OPTION_NON_VOLATILE As Long = 0
'Private Const KEY_WOW64_64KEY        As Long = &H100&
'
'Private Const ERROR_MORE_DATA        As Long = 234&
'Private Const ERROR_SUCCESS          As Long = 0&
'Private Const ERROR_ACCESS_DENIED    As Long = 5&
'
'Private Const REG_STANDARD_FORMAT   As Long = 1&
'Private Const REG_LATEST_FORMAT     As Long = 2&
'Private Const RRF_RT_ANY            As Long = &HFFFF&
'Private Const RRF_NOEXPAND          As Long = &H10000000

'Public Enum KEY_VIRTUAL_TYPE
'    KEY_VIRTUAL_NOT_EXIST
'    KEY_VIRTUAL_USUAL
'    KEY_VIRTUAL_SHARED
'    KEY_VIRTUAL_REDIRECTED
'End Enum
'#If False Then
'    Dim KEY_VIRTUAL_NOT_EXIST, KEY_VIRTUAL_USUAL, KEY_VIRTUAL_SHARED, KEY_VIRTUAL_REDIRECTED, KEY_VIRTUAL_SYMLINK
'#End If

Private Enum KEY_INFORMATION_CLASS
    KeyBasicInformation = 0
    KeyNodeInformation
    KeyFullInformation
    KeyNameInformation
    KeyCachedInformation
    KeyFlagsInformation
    KeyVirtualizationInformation
    KeyHandleTagsInformation
    MaxKeyInfoClass
End Enum

Private Enum OBJECT_INFORMATION_CLASS
    ObjectBasicInformation = 0
    ObjectNameInformation
    ObjectTypeInformation
    ObjectAllTypesInformation
    ObjectHandleInformation
    ObjectSessionInformation
End Enum

'Private Enum OBJ_ATTRIBUTES
'    OBJ_INHERIT = 2&
'    OBJ_PERMANENT = &H10&
'    OBJ_EXCLUSIVE = &H20&
'    OBJ_CASE_INSENSITIVE = &H40&
'    OBJ_OPENIF = &H80&
'    OBJ_OPENLINK = &H100&
'    OBJ_KERNEL_HANDLE = &H200&
'    OBJ_FORCE_ACCESS_CHECK = &H400&
'    OBJ_VALID_ATTRIBUTES = &H7F2&
'End Enum

Private Enum KEY_SET_INFORMATION_CLASS
    KeyWriteTimeInformation
    KeyWow64FlagsInformation
    KeyControlFlagsInformation
    KeySetVirtualizationInformation
    KeySetDebugInformation
    KeySetHandleTagsInformation
    KeySetLayerInformation
    MaxKeySetInfoClass
End Enum

Private Type PUBLIC_OBJECT_BASIC_INFORMATION
    Attributes As Long
    GrantedAccess As Long
    HandleCount As Long
    PointerCount As Long
    Reserved(9) As Long
End Type

Private Type OBJECT_ATTRIBUTES
    Length                   As Long
    RootDirectory            As Long
    ObjectName               As Long
    Attributes               As Long
    SecurityDescriptor       As Long
    SecurityQualityOfService As Long
End Type

Private Type KEY_VIRTUALIZATION_INFORMATION
    VirtualizationCandidate As Long
    VirtualizationEnabled   As Long
    VirtualTarget           As Long
    VirtualStore            As Long
    VirtualSource           As Long
    Reserved                As Long
End Type

Private Type KEY_FULL_INFORMATION
    LastWriteTime           As Currency
    TitleIndex              As Long
    ClassOffset             As Long
    ClassLength             As Long
    SubKeys                 As Long
    MaxNameLen              As Long
    MaxClassLen             As Long
    Values                  As Long
    MaxValueNameLen         As Long
    MaxValueDataLen         As Long
    Class(0)                As Integer 'WCHAR Class[1]
End Type

Private Const STANDARD_RIGHTS_READ   As Long = &H20000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const REG_OPTION_OPEN_LINK   As Long = 8&
Private Const REG_OPTION_BACKUP_RESTORE As Long = 4&
Private Const STATUS_SUCCESS         As Long = 0&
Private Const STATUS_INVALID_PARAMETER As Long = &HC000000D

'Private Declare Function CompareObjectHandles Lib "Kernelbase.dll" (ByVal hFirstObjectHandle As Long, ByVal hSecondObjectHandle As Long) As Long
Private Declare Function NtQueryKey Lib "ntdll.dll" (ByVal KeyHandle As Long, ByVal KeyInformationClass As KEY_INFORMATION_CLASS, KeyInformation As Any, ByVal Length As Long, ResultLength As Long) As Long
Private Declare Function NtQueryObject Lib "ntdll.dll" (ByVal Handle As Long, ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, ObjectInformation As Any, ByVal ObjectInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function NtOpenKeyEx Lib "ntdll.dll" (KeyHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As OBJECT_ATTRIBUTES, ByVal OpenOptions As Long) As Long
Private Declare Function NtOpenKey Lib "ntdll.dll" (KeyHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As OBJECT_ATTRIBUTES) As Long
Private Declare Function RegFlushKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function NtSetInformationKey Lib "ntdll.dll" (ByVal KeyHandle As Long, ByVal KeySetInformationClass As KEY_SET_INFORMATION_CLASS, ByVal KeySetInformation As Long, ByVal KeySetInformationLength As Long) As Long
Private Declare Function NtClose Lib "ntdll.dll" (ByVal Handle As Long) As Long
Private Declare Function NtEnumerateKey Lib "ntdll.dll" (ByVal KeyHandle As Long, ByVal Index As Long, ByVal KeyInformationClass As KEY_INFORMATION_CLASS, ByVal KeyInformation As Any, ByVal Length As Long, ResultLength As Long) As Long
Private Declare Function NtRenameKey Lib "ntdll.dll" (ByVal KeyHandle As Long, ByVal NewName As Long) As Long


Public Function GetHKey(ByVal HKeyName As Variant) As ENUM_REG_HIVE 'Get handle of main hive
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

Public Function GetHiveNameByHandle(Handle As ENUM_REG_HIVE) As String
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

'to extract root hive name from full key
'returns short hive name
'specify bIncludeSID to get e.g., HKU\S-1-5-20
Public Function GetShortHiveName(ByVal FullHiveName As String, Optional bIncludeSID As Boolean) As String
    On Error GoTo ErrorHandler:
    Dim pos As Long
    Dim sRoot As String
    Dim sSidPart As String
    pos = InStr(FullHiveName, "\")
    If pos <> 0 Then
        sRoot = Left$(FullHiveName, pos - 1)
    Else
        sRoot = FullHiveName
    End If
    Select Case UCase$(sRoot)
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
    If bIncludeSID Then
        If GetShortHiveName = "HKU" Then
            If pos <> 0 Then
                sSidPart = Mid$(FullHiveName, pos + 1)
                If Len(sSidPart) <> 0 Then
                    pos = InStr(sSidPart, "\")
                    If pos <> 0 Then
                        sSidPart = Left$(sSidPart, pos - 1)
                    End If
                    GetShortHiveName = GetShortHiveName & "\" & sSidPart
                End If
            End If
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetShortHiveName", FullHiveName
    If inIDE Then Stop: Resume Next
End Function

Public Function MapRegTypeToString(RegType As REG_VALUE_TYPE) As String
    On Error GoTo ErrorHandler:
    Dim sRegType As String
    Select Case RegType
        Case REG_NONE
            sRegType = "REG_NONE"
        Case REG_SZ
            sRegType = "REG_SZ"
        Case REG_EXPAND_SZ
            sRegType = "REG_EXPAND_SZ"
        Case REG_BINARY
            sRegType = "REG_BINARY"
        Case REG_DWORD
            sRegType = "REG_DWORD"
        Case REG_DWORDLittleEndian
            sRegType = "REG_DWORDLittleEndian"
        Case REG_DWORDBigEndian
            sRegType = "REG_DWORDBigEndian"
        Case REG_LINK
            sRegType = "REG_LINK"
        Case REG_MULTI_SZ
            sRegType = "REG_MULTI_SZ"
        Case REG_ResourceList
            sRegType = "REG_ResourceList"
        Case REG_FullResourceDescriptor
            sRegType = "REG_FullResourceDescriptor"
        Case REG_ResourceRequirementsList
            sRegType = "REG_ResourceRequirementsList"
        Case REG_QWORD
            sRegType = "REG_QWORD"
        Case REG_QWORD_LITTLE_ENDIAN
            sRegType = "REG_QWORD_LITTLE_ENDIAN"
    End Select
    MapRegTypeToString = sRegType
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MapRegTypeToString", RegType
    If inIDE Then Stop: Resume Next
End Function

Public Function MapStringToRegType(sRegType As String) As REG_VALUE_TYPE
    On Error GoTo ErrorHandler:
    Dim RegType As REG_VALUE_TYPE
    
    Select Case UCase$(sRegType)
        Case "REG_NONE"
            RegType = REG_NONE
        Case "REG_SZ"
            RegType = REG_SZ
        Case "REG_EXPAND_SZ"
            RegType = REG_EXPAND_SZ
        Case "REG_BINARY"
            RegType = REG_BINARY
        Case "REG_DWORD"
            RegType = REG_DWORD
        Case UCase$("REG_DWORDLittleEndian")
            RegType = REG_DWORDLittleEndian
        Case UCase$("REG_DWORDBigEndian")
            RegType = REG_DWORDBigEndian
        Case "REG_LINK"
            RegType = REG_LINK
        Case "REG_MULTI_SZ"
            RegType = REG_MULTI_SZ
        Case UCase$("REG_ResourceList")
            RegType = REG_ResourceList
        Case UCase$("REG_FullResourceDescriptor")
            RegType = REG_FullResourceDescriptor
        Case UCase$("REG_ResourceRequirementsList")
            RegType = REG_ResourceRequirementsList
        Case "REG_QWORD"
            RegType = REG_QWORD
        Case "REG_QWORD_LITTLE_ENDIAN"
            RegType = REG_QWORD_LITTLE_ENDIAN
    End Select
    MapStringToRegType = RegType
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MapStringToRegType", sRegType
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
'    sTemp = space$(lRet - 1&)
'    lRet = ExpandEnvironmentStrings(StrPtr(sData), StrPtr(sTemp), lRet)
'    If lRet Then
'        ExpandEnvStr = Left$(sTemp, lRet - 1&)
'    Else
'        ExpandEnvStr = sData
'    End If
    ExpandEnvStr = EnvironW(sData)
End Function


' if main hive handle wasn't defined, assigns handle according to hive's name defined by Full key name directed
Public Sub NormalizeKeyNameAndHiveHandle(ByRef lHive As ENUM_REG_HIVE, ByRef in_out_KeyName As String)
    Dim iPos        As Long
    If lHive = 0 Then
        lHive = GetHKey(in_out_KeyName)
        iPos = InStr(in_out_KeyName, "\")
        If (iPos <> 0&) Then in_out_KeyName = Mid$(in_out_KeyName, iPos + 1&) Else in_out_KeyName = vbNullString
    End If
End Sub


Public Function CreateKey(lHive As ENUM_REG_HIVE, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    lret = RegCreateKeyEx(lHive, StrPtr(sKey), 0&, ByVal 0&, 0&, KEY_CREATE_SUB_KEY Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), ByVal 0&, hKey, ByVal 0&)
    
    If lret = ERROR_ACCESS_DENIED Then
    
        If modPermissions.RegKeyResetDACL(lHive, GetParentDir(sKey), bUseWow64, False) Then
            lret = RegCreateKeyEx(lHive, StrPtr(sKey), 0&, ByVal 0&, 0&, KEY_CREATE_SUB_KEY Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), ByVal 0&, hKey, ByVal 0&)
            modPermissions.RegKeyResetDACL lHive, sKey, bUseWow64, True
        End If
    End If
    CreateKey = (ERROR_SUCCESS = lret)
    If hKey <> 0 Then
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CreateKey", lHive & "," & sKey
    If inIDE Then Stop: Resume Next
End Function


Public Function GetString(hHive As ENUM_REG_HIVE, sKey$, sValue$, Optional bUseWow64 As Boolean, Optional bDoNotExpand As Boolean) As String
    GetString = GetData(hHive, sKey, sValue, bUseWow64, bDoNotExpand) '-> redirection to common function, just in case REG type is wrong
End Function

Public Function GetDword(hHive As ENUM_REG_HIVE, sKey$, sValue$, Optional bUseWow64 As Boolean) As Long
    On Error GoTo ErrorHandler:

    Dim tmp As Variant
    tmp = GetData(hHive, sKey, sValue, bUseWow64)  '-> redirection to common function, just in case REG type is wrong
    If IsNumeric(tmp) Then GetDword = CLng(tmp)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetDword", hHive & "," & sKey & "," & sValue
    If inIDE Then Stop: Resume Next
End Function

Public Function GetQword(hHive As ENUM_REG_HIVE, sKey$, sValue$, Optional bUseWow64 As Boolean) As Currency
    On Error GoTo ErrorHandler:
    
    Dim tmp As Variant
    tmp = GetData(hHive, sKey, sValue, bUseWow64)  '-> redirection to common function, just in case REG type is wrong
    If IsNumeric(tmp) Then GetQword = CCur(tmp)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetQword", hHive & "," & sKey & "," & sValue
    If inIDE Then Stop: Resume Next
End Function

Public Function GetBinaryToString(hHive As ENUM_REG_HIVE, sKey$, sValue$, Optional bUseWow64 As Boolean) As String
    GetBinaryToString = GetData(hHive, sKey, sValue, bUseWow64)   '-> redirection to common function, just in case REG type is wrong
End Function

Public Function GetBinary(hHive As ENUM_REG_HIVE, ByVal KeyName As String, ValueName As String, Optional bUseWow64 As Boolean) As Byte()
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
                GetBinary = abData()
            End If
        End If
    End If
    
    If hKey <> 0& Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetBinary", "hHive:", hHive, KeyName & "\" & ValueName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function GetMultiSZ(hHive As ENUM_REG_HIVE, ByVal KeyName As String, ValueName As String, Optional bUseWow64 As Boolean) As String()
    On Error GoTo ErrorHandler:

    Dim cData        As Long
    Dim hKey         As Long
    Dim lret         As Long
    Dim sData        As String
    Dim ordType      As Long
    
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
                Do While AscW(Right$(sData, 1)) = 0 'safe rtim 'nul'
                    sData = Left$(sData, Len(sData) - 1)
                Loop
                GetMultiSZ = Split(sData, vbNullChar) 'struct is: item1 nul item2 nul ... itemN nul nul
            End If
        End If
    End If
    
    If hKey <> 0& Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetMultiSZ", "hHive:", hHive, KeyName & "\" & ValueName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function GetValueType(hHive As ENUM_REG_HIVE, ByVal KeyName As String, ValueName As String, Optional bUseWow64 As Boolean) As REG_VALUE_TYPE
    On Error GoTo ErrorHandler:
    
    Dim hKey         As Long
    Dim lret         As Long
    Dim ordType      As Long
    
    Call NormalizeKeyNameAndHiveHandle(hHive, KeyName)
    
    If ERROR_SUCCESS <> RegOpenKeyEx(hHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then Exit Function
    
    lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, ByVal 0&, ByVal 0&)
    
    GetValueType = ordType
    
    If hKey <> 0& Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetValueType", "hHive:", hHive, KeyName & "\" & ValueName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function
    

Public Function GetData(hHive As ENUM_REG_HIVE, ByVal KeyName As String, ValueName As String, _
    Optional bUseWow64 As Boolean, _
    Optional bDoNotExpand As Boolean, _
    Optional bUseHexFormatForBinary As Boolean, _
    Optional out_KeyType As Long) As Variant
    
    On Error GoTo ErrorHandler:
    
    'bDoNotExpand - Win 2003 SP1+ supported only
    
    Dim abData()     As Byte
    Dim cData        As Long
    Dim hKey         As Long
    Dim lData        As Long
    Dim qData        As Currency
    Dim lret         As Long
    Dim ordType      As Long
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
    
    out_KeyType = ordType
    
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
                Dim bRegGetValueUsed As Boolean
            
                'RegGetValue - Win 2003 SP1 +
                If bDoNotExpand Then
                    If OSver.MajorMinor >= 5.2 Then
                        If IsProcedureAvail("RegGetValueW", "Advapi32.dll") Then
                            sData = String$(cData - 1&, vbNullChar)
                            lret = RegGetValue(hKey, ByVal 0&, StrPtr(ValueName), RRF_RT_ANY Or RRF_NOEXPAND, ByVal 0&, StrPtr(sData), cData)

                            'Note: if you don't set RRF_NOEXPAND flag, you should prepare a bit bigger buffer in case of REG_EXPAND_SZ type,
                            'because it must be large enought to get string with expanded environment variables. And also anyway, be ready to get ERROR_MORE_DATA error.

                            If lret = ERROR_MORE_DATA Then
                                sData = String$(cData - 1&, vbNullChar)
                                lret = RegGetValue(hKey, ByVal 0&, StrPtr(ValueName), RRF_RT_ANY Or RRF_NOEXPAND, ByVal 0&, StrPtr(sData), cData)
                            End If
                            
                            vValue = Left$(sData, lstrlen(StrPtr(sData)))
                            bRegGetValueUsed = True
                        End If
                    End If
                End If
                
                If Not bRegGetValueUsed Then
                    sData = String$(cData \ 2 + 1, vbNullChar)  ' (this API doesn't ensure that result buffer will contain null char, so I'll add extra 2 bytes)
                
                    lret = RegQueryValueEx(hKey, StrPtr(ValueName), 0&, ordType, StrPtr(sData), cData)
                    If lret = ERROR_SUCCESS Then
                        vValue = ExpandEnvStr(Left$(sData, lstrlen(StrPtr(sData))))
                    End If
                    
                    If bDoNotExpand Then
                        vValue = EnvironUnexpand(vValue, False)
                    End If
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
                    If bUseHexFormatForBinary Then
                        vValue = SerializeByteArray(abData)
                    Else
                        vValue = StrConv(abData, vbUnicode)
                    End If
                End If
            End If
    
    End Select
    GetData = vValue
    If hKey <> 0& Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetData", "hHive:", hHive, KeyName & "\" & ValueName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function GetKeysCount(hHive As ENUM_REG_HIVE, ByVal KeyName As String, Optional bUseWow64 As Boolean) As Long
    On Error GoTo ErrorHandler:
    Dim KeysCnt As Long
    Dim hKey As Long
    Dim lret As Long
    
    Call NormalizeKeyNameAndHiveHandle(hHive, KeyName)
    
    lret = RegOpenKeyEx(hHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    
    If lret = ERROR_SUCCESS Then
        If ERROR_SUCCESS = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, KeysCnt, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&) Then
            GetKeysCount = KeysCnt
        End If
        CloseHandle hKey
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetValuesCount", "hHive:", hHive, KeyName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function GetValuesCount(hHive As ENUM_REG_HIVE, ByVal KeyName As String, Optional bUseWow64 As Boolean) As Long
    On Error GoTo ErrorHandler:
    Dim ValuesCnt As Long
    Dim hKey As Long
    Dim lret As Long
    
    Call NormalizeKeyNameAndHiveHandle(hHive, KeyName)
    
    lret = RegOpenKeyEx(hHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    
    If lret = ERROR_SUCCESS Then
        If ERROR_SUCCESS = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, ByVal 0&, ByVal 0&, ByVal 0&, ValuesCnt, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&) Then
            GetValuesCount = ValuesCnt
        End If
        CloseHandle hKey
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetValuesCount", "hHive:", hHive, KeyName, "bUseWow64:", bUseWow64
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function EnumValuesAndData(lHive As ENUM_REG_HIVE, ByVal KeyName As String, _
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
    Dim sName        As String
    Dim ValuesCnt    As Long
    Dim lType        As Long
    Dim bData()      As Byte
    Dim lData        As Long
    Dim sData        As String
    Dim qData        As Currency
    
    Erase aValueNames
    Erase aDataValues
    Erase aTypes
    
    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    lret = RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    
    If lret <> ERROR_SUCCESS Then
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
    EnumValuesAndData = CurValueN

    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumValuesAndData", "lHive:", lHive, "Key:", KeyName, "uTypesToQuery:", uRegTypesToQuery, "bUseWow64:", bUseWow64
    If (hKey <> 0&) Then RegCloseKey hKey
End Function


Public Function SetStringVal(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, sData$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not KeyExists(lHive, sKey, bUseWow64) Then
        CreateKey lHive, sKey, bUseWow64
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
            SetStringVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_SZ, ByVal 0&, 0&))
        Else
            SetStringVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_SZ, ByVal StrPtr(sData), Len(sData) * 2 + 2))
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetStringVal", lHive & "," & sKey & "\" & sValue & "," & sData
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function SetExpandStringVal(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, sData$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not KeyExists(lHive, sKey, bUseWow64) Then
        CreateKey lHive, sKey, bUseWow64
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
    SetExpandStringVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_EXPAND_SZ, ByVal StrPtr(sData), Len(sData) * 2 + 2))
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetExpandStringVal", lHive & "," & sKey & "\" & sValue & "," & sData
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function SetDwordVal(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, lData&, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not KeyExists(lHive, sKey, bUseWow64) Then
        CreateKey lHive, sKey, bUseWow64
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
    SetDwordVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_DWORD, lData, 4&))
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetDwordVal", lHive & "," & sKey & "\" & sValue & "," & lData
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function SetQwordVal(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, vData As Variant, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&, cData As Currency
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not KeyExists(lHive, sKey, bUseWow64) Then
        CreateKey lHive, sKey, bUseWow64
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
    
    If Len(vData) Then
        cData = CCur(vData)
        SetQwordVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_QWORD, cData, 8&))
    Else
        'for empty values
        SetQwordVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_QWORD, cData, 0&))
    End If
    
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetQwordVal", lHive & "," & sKey & "\" & sValue & "," & cData
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function SetBinaryVal(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, aData() As Byte, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not KeyExists(lHive, sKey, bUseWow64) Then
        CreateKey lHive, sKey, bUseWow64
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
    If AryItems(aData) Then
        SetBinaryVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_BINARY, ByVal VarPtr(aData(0)), UBound(aData) + 1&))
    Else
        SetBinaryVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_BINARY, ByVal 0&, 0&))
    End If
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetBinaryVal", lHive & "," & sKey & "\" & sValue & ",UBound=" & UBound(aData)
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function SetMultiSZVal(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, aData() As String, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lret&, sData$, i&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If Not KeyExists(lHive, sKey, bUseWow64) Then
        CreateKey lHive, sKey, bUseWow64
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
    
    If AryItems(aData) Then

        For i = 0 To UBound(aData)
            sData = sData & aData(i) & vbNullChar
        Next
        If Len(sData) = 1 Then sData = ""
        
        SetMultiSZVal = (ERROR_SUCCESS = RegSetValueEx(hKey, StrPtr(sValue), 0&, REG_MULTI_SZ, ByVal StrPtr(sData), Len(sData) * 2 + 2)) 'data + nul + nul
    End If
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetBinaryVal", lHive & "," & sKey & "\" & sValue & ",UBound=" & UBound(aData)
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function DelVal(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, Optional bUseWow64 As Boolean) As Boolean
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
        ElseIf lret = ERROR_FILE_NOT_FOUND Then
            DelVal = True
            Exit Function
        Else
            Exit Function
        End If
    End If
    lret = RegDeleteValue(hKey, StrPtr(sValue))
    
    If lret = ERROR_SUCCESS Or lret = ERROR_FILE_NOT_FOUND Then
        DelVal = True
    End If
    
    If hKey <> 0 Then
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DelVal", lHive & "," & sKey & "\" & sValue
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function DelKey(lHive As ENUM_REG_HIVE, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
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
       
    If OSver.Bitness = "x32" And Not OSver.IsWindowsVistaOrGreater Then
        'XP x32
        SHDeleteKey lHive, StrPtr(sKey)
        If KeyExists(lHive, sKey, True) Then
        
            If modPermissions.RegKeyResetDACL(lHive, sKey, True, True) Then
                SHDeleteKey lHive, StrPtr(sKey)
            End If
        End If
    Else
        DeleteAllSubKeys lHive, sKey, bUseWow64
        
        lret = RegDeleteKeyEx(lHive, StrPtr(sKey), bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64, 0&)
        
        If lret = ERROR_ACCESS_DENIED Then   'root key self
        
            If modPermissions.RegKeyResetDACL(lHive, sKey, bUseWow64, False) Then
                RegDeleteKeyEx lHive, StrPtr(sKey), bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64, 0&
            End If
        End If
        
    End If
    DelKey = Not KeyExists(lHive, sKey, bUseWow64)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DelKey", lHive & "," & sKey
    If inIDE Then Stop: Resume Next
End Function


Public Function DeleteAllSubKeys(lHive As ENUM_REG_HIVE, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    'del subkeys without root
    On Error GoTo ErrorHandler:
    Dim i&, sSubKey$, aSubKeys() As String, Flag As Boolean, lret&
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    Flag = True
    
    For i = 1 To EnumSubKeysToArray(lHive, sKey, aSubKeys(), bUseWow64, True)
        sSubKey = sKey & "\" & aSubKeys(i)
        Flag = Flag And DeleteAllSubKeys(lHive, sSubKey, bUseWow64)   '< --- recursively
        lret = RegDeleteKeyEx(lHive, StrPtr(sSubKey), bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64, 0&)
        
        If lret = ERROR_ACCESS_DENIED Then
            If modPermissions.RegKeyResetDACL(lHive, sSubKey, bUseWow64, False) Then
                lret = RegDeleteKeyEx(lHive, StrPtr(sSubKey), bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64, 0&)
            End If
        End If
        If lret <> ERROR_SUCCESS Then Flag = False
    Next i
    DeleteAllSubKeys = Flag
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DeleteAllSubKeys", lHive & "," & sKey
    If inIDE Then Stop: Resume Next
End Function


Public Function KeyExists(lHive As ENUM_REG_HIVE, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lStatus&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    lStatus = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, WRITE_OWNER Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey)
    
    If lStatus = ERROR_SUCCESS Then
        KeyExists = True
        RegCloseKey hKey
        
    ElseIf lStatus = ERROR_ACCESS_DENIED Then
        
        'for 'Limited User'
        If ERROR_SUCCESS = RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        
            KeyExists = True
            RegCloseKey hKey
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "KeyExists", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function ValueExists(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, WRITE_OWNER Or KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) <> 0 Then
        If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) <> 0 Then Exit Function
    End If
    ValueExists = (ERROR_SUCCESS = RegQueryValueEx(hKey, StrPtr(sValue), 0&, ByVal 0&, ByVal 0&, ByVal 0&))
    RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ValueExists", lHive & "," & sKey & "\" & sValue
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function KeyHasSubKeys(lHive As ENUM_REG_HIVE, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, sBuf$, cbBuf&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    KeyHasSubKeys = False
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) = 0& Then
        cbBuf = MAX_KEYNAME
        sBuf = String$(cbBuf, vbNullChar)
        If RegEnumKeyEx(hKey, 0&, StrPtr(sBuf), cbBuf, 0&, ByVal 0&, ByVal 0&, ByVal 0&) = 0& Then
            KeyHasSubKeys = True
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "KeyHasSubKeys", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function GetFirstSubKey(lHive As ENUM_REG_HIVE, ByVal sKey$, Optional bUseWow64 As Boolean) As String
    On Error GoTo ErrorHandler:
    Dim hKey&, sName$, cbBuf&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) = 0& Then
        cbBuf = MAX_KEYNAME
        sName = String$(cbBuf, vbNullChar)
        If ERROR_SUCCESS = RegEnumKeyEx(hKey, 0&, StrPtr(sName), cbBuf, 0&, ByVal 0&, ByVal 0&, ByVal 0&) Then
            GetFirstSubKey = Left$(sName, lstrlen(StrPtr(sName)))
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFirstSubKey", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function KeyHasValues(lHive As ENUM_REG_HIVE, ByVal sKey$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, sName$, cbBuf&
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) = 0& Then
        cbBuf = MAX_VALUENAME
        sName = String$(cbBuf, vbNullChar)
        If RegEnumValue(hKey, 0&, StrPtr(sName), cbBuf, 0&, ByVal 0&, ByVal 0&, ByVal 0&) = 0& Then
            KeyHasValues = True
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "KeyHasValues", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function EnumSubKeys(lHive As ENUM_REG_HIVE, ByVal sKey$, Optional bUseWow64 As Boolean, Optional ForceUnlock As Boolean) As String
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
    If Len(sDummy) <> 0 Then EnumSubKeys = Left$(sDummy, Len(sDummy) - 1)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumSubKeys", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

'// return array of key names (or full key without hive name, if bRecursively == true)
'in bounds [1 to count], when [count] is a number returned by function

'Warning: the caller is responsible for erasing aSubKeys() array
Public Function EnumSubKeysToArray( _
    lHive As ENUM_REG_HIVE, _
    ByVal sKey As String, _
    aSubKeys() As String, _
    Optional bUseWow64 As Boolean, _
    Optional ForceUnlock As Boolean, _
    Optional bRecursively As Boolean, _
    Optional bEraseArray As Boolean = True) As Long
    
    On Error GoTo ErrorHandler:
    Dim hKey        As Long
    Dim sName       As String
    Dim lret        As Long
    Dim cMaxSubKeyLen As Long
    Dim lNameSize   As Long
    Dim idx         As Long
    Dim SubkeysCnt  As Long
    Dim BeginIdx    As Long
    Dim i           As Long
    
    If bEraseArray Then Erase aSubKeys
    
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

          If 0 = AryItems(aSubKeys) Or (Not bRecursively) Then
            If lret = ERROR_SUCCESS Then
                ReDim aSubKeys(1 To SubkeysCnt)
            Else
                ReDim aSubKeys(1 To 100)
            End If
          Else
            BeginIdx = UBound(aSubKeys)
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
            
            If (lret = ERROR_SUCCESS) Then
                idx = idx + 1
                sName = Left$(sName, lstrlen(StrPtr(sName)))
                
                If UBound(aSubKeys) < (BeginIdx + idx) Then
                    ReDim Preserve aSubKeys(1 To UBound(aSubKeys) + 100)
                End If
                If Not bRecursively Then
                    aSubKeys(BeginIdx + idx) = sName
                Else
                    aSubKeys(BeginIdx + idx) = IIf(sKey <> "", sKey & "\" & sName, sName)
                End If
            Else
                ReDim Preserve aSubKeys(1 To BeginIdx + idx)
            End If
            
          Loop While lret = ERROR_SUCCESS
        
          If bRecursively Then
            If hKey <> 0 Then RegCloseKey hKey: hKey = 0
            
            For i = BeginIdx + 1 To UBound(aSubKeys)
                Call EnumSubKeysToArray(lHive, aSubKeys(i), aSubKeys, bUseWow64, ForceUnlock, True, False)
            Next
          End If
        End If
    End If
    
    If hKey <> 0 Then RegCloseKey hKey
    EnumSubKeysToArray = UBoundSafe(aSubKeys)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumSubKeysToArray", lHive & "," & sKey
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

Public Function NtEnumSubKeysToArray( _
    lHive As ENUM_REG_HIVE, _
    ByVal sKey As String, _
    aSubKeys() As String, _
    Optional bUseWow64 As Boolean, _
    Optional ForceUnlock As Boolean, _
    Optional bRecursively As Boolean, _
    Optional bEraseArray As Boolean = True) As Long
    
    On Error GoTo ErrorHandler:
    Dim hKey        As Long
    Dim sName       As String
    Dim lret        As Long
    Dim BeginIdx    As Long
    Dim idx         As Long
    Dim i           As Long
    Dim buf()       As Byte
    Dim cchBuf      As Long
    Dim reqSize     As Long
    Dim kfi         As KEY_FULL_INFORMATION
    Dim strLength   As Long
    
    If bEraseArray Then Erase aSubKeys
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    lret = WrapNtOpenKeyEx(lHive, sKey, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE, hKey, , bUseWow64)
    
    If lret = STATUS_ACCESS_DENIED Then
        lret = WrapNtOpenKeyEx(lHive, sKey, KEY_ENUMERATE_SUB_KEYS, hKey, , bUseWow64)
        
        If lret = STATUS_ACCESS_DENIED And ForceUnlock Then
        
            If modPermissions.RegKeyResetDACL(lHive, sKey, bUseWow64, False) Then
            
                lret = WrapNtOpenKeyEx(lHive, sKey, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE, hKey, , bUseWow64)
            End If
        End If
    End If
    
    If lret <> STATUS_SUCCESS Then Exit Function
    
    If 0 = AryItems(aSubKeys) Or (Not bRecursively) Then
        ReDim aSubKeys(1 To 100)
    Else
        BeginIdx = UBound(aSubKeys)
    End If
    
    Do
        ReDim buf(0): buf(0) = 0: cchBuf = 0
        
        'Note: KeyNameInformation is not supported here
        lret = NtEnumerateKey(hKey, idx, KeyBasicInformation, VarPtr(buf(0)), cchBuf, reqSize)
        
        If lret = STATUS_BUFFER_TOO_SMALL Then
            ReDim buf(reqSize - 1)
            cchBuf = reqSize
            
            lret = NtEnumerateKey(hKey, idx, KeyBasicInformation, VarPtr(buf(0)), cchBuf, reqSize)
            
            If lret = STATUS_SUCCESS Then
                GetMem4 buf(12), strLength
                
                If strLength <> 0 Then
                    sName = String$(strLength \ 2, 0)
                    memcpy ByVal StrPtr(sName), buf(16), strLength
                End If
            End If
        End If
        
        If (lret = STATUS_SUCCESS) Then
            
            idx = idx + 1
            If UBound(aSubKeys) < (BeginIdx + idx) Then
                ReDim Preserve aSubKeys(1 To UBound(aSubKeys) + 100)
            End If
            
            If Not bRecursively Then
                aSubKeys(BeginIdx + idx) = sName
            Else
                aSubKeys(BeginIdx + idx) = IIf(sKey <> "", sKey & "\" & sName, sName)
            End If
        End If
    
    Loop While lret = STATUS_SUCCESS
    
    If (BeginIdx + idx) > 0 Then
        ReDim Preserve aSubKeys(1 To BeginIdx + idx)
    Else
        Erase aSubKeys
    End If
    
    If hKey <> 0 Then NtClose hKey: hKey = 0
    
    If bRecursively Then
        For i = BeginIdx + 1 To UBound(aSubKeys)
            Call NtEnumSubKeysToArray(lHive, aSubKeys(i), aSubKeys, bUseWow64, ForceUnlock, True, False)
        Next
    End If
    
    NtEnumSubKeysToArray = UBoundSafe(aSubKeys)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "NtEnumSubKeysToArray", lHive & "," & sKey
    If hKey <> 0 Then NtClose hKey: hKey = 0
    If inIDE Then Stop: Resume Next
End Function

Public Function EnumValues(lHive As ENUM_REG_HIVE, ByVal KeyName$, Optional bUseWow64 As Boolean) As String
    On Error GoTo ErrorHandler:
    Dim iNameMax     As Long
    Dim hKey         As Long
    Dim idx          As Long
    Dim lNameSize    As Long
    Dim lret         As Long
    Dim sName        As String
    Dim sDummy       As String
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
    If Len(sDummy) <> 0 Then EnumValues = Left$(sDummy, Len(sDummy) - 1)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumValues", lHive & "," & KeyName
    If (hKey <> 0&) Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

'// returns array [1 to count], where [count] is a number returned by this function

'Warning: the caller is responsible for erasing aValues() array
Public Function EnumValuesToArray(lHive As ENUM_REG_HIVE, ByVal KeyName$, aValues() As String, Optional bUseWow64 As Boolean) As Long
    On Error GoTo ErrorHandler:
    Dim iNameMax     As Long
    Dim hKey         As Long
    Dim idx          As Long
    Dim lNameSize    As Long
    Dim lret         As Long
    Dim sName        As String
    Dim ValuesCnt    As Long
    
    Erase aValues()
    
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
                If UBound(aValues) < idx Then ReDim Preserve aValues(1 To UBound(aValues) + 100)
                aValues(idx) = sName
            Else
                If idx > 1 Then
                    ReDim Preserve aValues(1 To idx - 1)
                    EnumValuesToArray = idx - 1
                End If
            End If
            
          Loop While lret = ERROR_SUCCESS
        End If
    End If
    If (hKey <> 0&) Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumValues", lHive & "," & KeyName
    If (hKey <> 0&) Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function GetKeyTime(lHive As ENUM_REG_HIVE, ByVal KeyName$, Optional bUseWow64 As Boolean) As Date
    On Error GoTo ErrorHandler:
    Dim hKey         As Long
    Dim lret         As Long
    Dim fTime        As FILETIME
    Dim sTime        As SYSTEMTIME
    Dim DateTime     As Date

    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    If ERROR_SUCCESS = RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        
        If ERROR_SUCCESS = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, fTime) Then
        
            lret = FileTimeToLocalFileTime(fTime, fTime)    ' учитываем смещение согласно текущим региональным настройкам в системе
            lret = FileTimeToSystemTime(fTime, sTime)       ' FILETIME -> SYSTEMTIME
        
            'GetKeyTime = DateSerial(fTime.wYear, fTime.wMonth, fTime.wDay) + TimeSerial(fTime.wHour, fTime.wMinute, fTime.wSecond)
            
            SystemTimeToVariantTime sTime, DateTime         ' SYSTEMTIME -> Date
            GetKeyTime = DateTime
        End If
    End If
    
    If (hKey <> 0&) Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetKeyTime", lHive & "," & KeyName
    If (hKey <> 0&) Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function SetKeyTime(lHive As ENUM_REG_HIVE, ByVal KeyName$, dDate As Date, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey         As Long
    Dim lret         As Long
    Dim fTime        As FILETIME
    Dim sTime        As SYSTEMTIME
    
    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    If ERROR_SUCCESS = RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_WRITE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        
        VariantTimeToSystemTime dDate, sTime
        SystemTimeToFileTime sTime, fTime
        LocalFileTimeToFileTime fTime, fTime
        
        If STATUS_SUCCESS = NtSetInformationKey(hKey, KeyWriteTimeInformation, VarPtr(fTime), LenB(fTime)) Then
            SetKeyTime = True
        End If
    End If
    
    If (hKey <> 0&) Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetKeyTime", lHive & "," & KeyName
    If (hKey <> 0&) Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


'it makes registry export (temporary .reg-file), then read its contents to variable
Public Function ExportKeyToVariable( _
    lHive As ENUM_REG_HIVE, _
    ByVal sKey$, _
    Optional bUseWow64 As Boolean, _
    Optional SaveHeader As Boolean = True, _
    Optional ConvertToANSI As Boolean = False) As String
    
    On Error GoTo ErrorHandler:

    Dim sTempFile$, sData$, hFile&, FSize@, isUnicode As Boolean, pos&

    Randomize
    sTempFile$ = BuildPath(TempCU, "backup_" & Int(Rnd * 10000) & ".reg")
    
    If lHive <> 0 Then
        sKey = GetHiveNameByHandle(lHive) & "\" & sKey
    End If
    
    '// TODO: Wow6432Node is not substitutes to the name of the key in export file when you enum under Wow64
    
    If FileExists(sWinSysDir & "\reg.exe", True) Then
        
        If Proc.ProcessRun(sWinSysDir & "\reg.exe", "export " & """" & sKey & """" & " " & """" & sTempFile & """" & " /y" & IIf(bUseWow64 Or OSver.IsWin32, "", " /reg:64"), , vbHide) Then

            If ERROR_SUCCESS <> Proc.WaitForTerminate(, , False, 10000) Then    '10 sec. timeout
                Proc.ProcessClose , , True
            End If
            
            If FileExists(sTempFile) Then
                    
                If OpenW(sTempFile, FOR_READ, hFile) Then
                    
                    FSize = LOFW(hFile)
                    
                    If FSize > 0 Then
                        
                        sData = String$(FSize, 0)
                        
                        If GetW(hFile, 1, sData) Then
                            
                            isUnicode = (AscW(Left$(sData, 1)) = 1103 And (AscW(Mid$(sData, 2, 1)) = 1102)) '"яю"
                            
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
    ExportKeyToVariable = sData
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ExportKeyToVariable", lHive & "," & sKey, "bUseWow64:", bUseWow64, "SaveHeader:", SaveHeader
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFilenameFromBinaryKey(lHive As ENUM_REG_HIVE, ByVal sKey$, sValue$, Optional bUseWow64 As Boolean) As String
    On Error GoTo ErrorHandler:
    
    Dim hKey&, sFile$, cbData&
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    cbData = MAX_VALUENAME
    If RegOpenKeyEx(lHive, StrPtr(sKey), 0, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not bUseWow64), hKey) = 0 Then
        ReDim uData(MAX_VALUENAME - 1) As Byte
        If RegQueryValueEx(hKey, StrPtr(sValue), 0, ByVal 0&, ByVal VarPtr(uData(0)), cbData) = 0 Then
            sFile = TrimNull(StrConv(uData, vbUnicode))
        End If
        RegCloseKey hKey
    End If
    GetFilenameFromBinaryKey = sFile
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilenameFromBinaryKey", lHive & "," & sKey, "sParam:", sValue
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function


Public Function KeyExportToBinary(lHive As ENUM_REG_HIVE, ByVal KeyName$, destFile As String, Optional bUseWow64 As Boolean) As Boolean     'Save key to binary .hiv file
    On Error GoTo ErrorHandler:
    Dim hKey         As Long

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


Public Function GetDefaultProgram(sExtension As String) As String 'e.g., ".txt"
    On Error GoTo ErrorHandler:

    Dim TXTProg     As String
    Dim TXTClassID  As String
    Dim sFile       As String
    Dim sVerb       As String
    
    TXTClassID = GetData(HKEY_CLASSES_ROOT, sExtension, "")
    
    If TXTClassID <> "" Then
        sVerb = GetData(HKEY_CLASSES_ROOT, TXTClassID & "\shell", "")
        
        If sVerb = "" Then sVerb = "open"
    
        TXTProg = EnvironW(GetData(HKEY_CLASSES_ROOT, TXTClassID & "\shell\" & sVerb & "\command", ""))
        
        SplitIntoPathAndArgs TXTProg, sFile, , True
        TXTProg = sFile
    End If
    
    If Not FileExists(TXTProg) Then
        TXTProg = "rundll32.exe shell32,ShellExec_RunDLL"
    End If
    GetDefaultProgram = TXTProg
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetDefaultProgram", "Extension=", sExtension
    If inIDE Then Stop: Resume Next
End Function


Public Function RenameKey(hHive As ENUM_REG_HIVE, sKey As String, sKeyNewName As String, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:

    Dim hKey As Long, lret As Long
    Call NormalizeKeyNameAndHiveHandle(hHive, sKey)
    
    lret = WrapNtOpenKeyEx(hHive, sKey, KEY_WRITE, hKey, , bUseWow64)
    
    If STATUS_SUCCESS = lret Then
    
        Dim UniStr As UNICODE_STRING
    
        With UniStr
            .Length = LenB(sKeyNewName)
            .MaximumLength = .Length
            .Buffer = StrPtr(sKeyNewName)
        End With

        'see also: https://www.unsafehex.com/index.php/2018/04/14/undocumented-function-regrenamekey/
        
        RenameKey = (STATUS_SUCCESS = NtRenameKey(hKey, VarPtr(UniStr)))
        
        NtClose hKey
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RenameKey", "sKeyOldPath=", sKey, "sKeyNewName=", sKeyNewName
    If inIDE Then Stop: Resume Next
End Function


Public Function RegisterDll(Path As String, Optional Unload As Boolean = True) As Boolean
    Dim hLib    As Long
    Dim pAddr   As Long
    Dim lret    As Variant
    Dim bOldRedir As Boolean
    
    ToggleWow64FSRedirection False, Path, bOldRedir
    
    hLib = LoadLibrary(StrPtr(Path))
    
    ToggleWow64FSRedirection bOldRedir
    
    If hLib = 0 Then Exit Function
    
    pAddr = GetProcAddress(hLib, StrPtr(StrConv("DllRegisterServer", vbFromUnicode)))
    
    If pAddr Then
    
        If DispCallFunc(0, pAddr, CC_STDCALL, vbLong, 0, 0, 0, lret) = 0 Then
        
            If lret = 0 Then RegisterDll = True
        
        End If
        
    End If
    
    If Unload Then FreeLibrary hLib
    
End Function

Public Function UnRegisterDll(Path As String) As Boolean
    Dim hLib    As Long
    Dim pAddr   As Long
    Dim lret    As Variant
    Dim bOldRedir As Boolean
    
    ToggleWow64FSRedirection False, Path, bOldRedir
    
    hLib = LoadLibrary(StrPtr(Path))
   
    ToggleWow64FSRedirection bOldRedir
   
    If hLib = 0 Then Exit Function
    
    pAddr = GetProcAddress(hLib, StrPtr(StrConv("DllUnregisterServer", vbFromUnicode)))
    
    If pAddr Then
    
        If DispCallFunc(0, pAddr, CC_STDCALL, vbLong, 0, 0, 0, lret) = 0 Then
        
            If lret = 0 Then UnRegisterDll = True
        
        End If
        
    End If
    
    FreeLibrary hLib
    
End Function

Public Function SetData(hHive As ENUM_REG_HIVE, ByVal KeyName As String, ValueName As String, ValueType As REG_VALUE_TYPE, vData As Variant, Optional bUseWow64 As Boolean) As Boolean
    
    On Error GoTo ErrorHandler:
    
    Call NormalizeKeyNameAndHiveHandle(hHive, KeyName)
    
    Select Case ValueType
        Case REG_BINARY
            Dim bData() As Byte
            bData = DeSerializeToByteArray(CStr(vData)) 'hex string to array
            SetData = SetBinaryVal(hHive, KeyName, ValueName, bData(), bUseWow64)
            
        Case REG_DWORD, REG_DWORDBigEndian, REG_DWORDLittleEndian
            SetData = SetDwordVal(hHive, KeyName, ValueName, CLng(vData), bUseWow64)
            
        Case REG_EXPAND_SZ
            SetData = SetExpandStringVal(hHive, KeyName, ValueName, CStr(vData), bUseWow64)
            
        Case REG_MULTI_SZ
            Dim sData() As String
            sData = Split(vData, vbCrLf)
            SetData = SetMultiSZVal(hHive, KeyName, ValueName, sData(), bUseWow64)
            
        Case REG_SZ
            SetData = SetStringVal(hHive, KeyName, ValueName, CStr(vData), bUseWow64)
        
        Case REG_QWORD
            SetData = SetQwordVal(hHive, KeyName, ValueName, vData, bUseWow64)
        
        Case Else
            MsgBoxW "This registry value type is not supported: " & ValueType, vbExclamation
            SetData = False
    End Select
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetData"
    If inIDE Then Stop: Resume Next
End Function

Public Sub Jump(lHive As ENUM_REG_HIVE, ByVal sKey As String, Optional sValue As String, Optional bUseWow64 As Boolean)
    On Error GoTo ErrorHandler:
    
    Dim i&, iProcId&, uSEI As SHELLEXECUTEINFO
    Dim hwndRegedit&, hwndTreeView&, hwndListView&
    Dim bVista As Boolean, sPrefix As String, sFullKey As String
    Dim bNoValue As Boolean, nMSec As Single
    
    Call NormalizeKeyNameAndHiveHandle(lHive, sKey)
    
    'verify the key actually exists
    Do Until KeyExists(lHive, sKey, bUseWow64) Or Len(sKey) = 0
        'move to parent key if target doesn't exist
        sKey = GetParentDir(sKey)
    Loop
    
    If sValue <> "" Then
        'value exists ?
        bNoValue = Not ValueExists(lHive, sKey, sValue, bUseWow64)
    End If
    
    'take account of the reg redirection
    If bUseWow64 Then
        If Not ((lHive = HKCR) And StrBeginWith(sKey, "Wow6432Node\")) And _
          Not StrBeginWith(sKey, "SOFTWARE\Classes\Wow6432Node") And _
          Not StrBeginWith(sKey, "SOFTWARE\Wow6432Node") Then
            If StrBeginWith(sKey, "SOFTWARE\Classes\") Or StrComp(sKey, "SOFTWARE\Classes", 1) = 0 Then
                sKey = "SOFTWARE\Classes\Wow6432Node" & Mid$(sKey, Len("SOFTWARE\Classes\"))
            ElseIf StrBeginWith(sKey, "SOFTWARE\") Or StrComp(sKey, "SOFTWARE", 1) = 0 Then
                sKey = "SOFTWARE\Wow6432Node" & Mid$(sKey, Len("SOFTWARE\"))
            ElseIf lHive = HKCR Then
                sKey = "Wow6432Node\" & sKey
            End If
        End If
    End If
    
    sFullKey = GetHiveNameByHandle(lHive) & "\" & sKey
    
    bVista = (OSver.MajorMinor = 6)
    If Not bVista Then
        sPrefix = GetString(HKCU, "Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", "Lastkey")
        If sPrefix = "" Then
            If OSver.LangDisplayCode = &H419& Then
                'sPrefix = "Компьютер\"
                sPrefix = STR_CONST.RU_PC
            Else
                sPrefix = "Computer\"
            End If
        Else
            i = InStr(sPrefix, "\")
            If i <> 0 Then
                sPrefix = Left$(sPrefix, i)
            Else
                sPrefix = sPrefix & "\"
            End If
        End If
        SetStringVal HKCU, "Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", "Lastkey", sPrefix & sFullKey
    End If
    
    'If ProcessExist("regedit.exe", True) Then
        Proc.ProcessClose , "regedit.exe", False, 5000, , False
    'End If
    
    'start regedit and wait until it's done loading
    With uSEI
        .cbSize = Len(uSEI)
        .lpVerb = StrPtr("open")
        If OSver.Bitness = "x64" Then
            If FileExists(BuildPath(sWinDir & "\sysnative", "regedt32.exe")) Then
                .lpFile = StrPtr(BuildPath(sWinDir & "\sysnative", "regedt32.exe"))
            Else
                .lpFile = StrPtr(BuildPath(sWinDir, "regedit.exe"))
            End If
        Else
            .lpFile = StrPtr(BuildPath(sWinDir, "regedit.exe"))
        End If
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .nShow = SW_SHOWNORMAL
    End With
    If ShellExecuteEx(uSEI) = 0 Then
        'Unable to start Regedit.
        MsgBoxW Translate(2100)
        Exit Sub
    End If
    
    If uSEI.hProcess <> 0 Then
        WaitForInputIdle uSEI.hProcess, 10000
    Else
        'Unable to start Regedit.
        MsgBoxW Translate(2100) & ", hProcess = 0"
        Exit Sub
    End If
    
    'find the regedit window and it's components
    Do
        Sleep 200
        nMSec = nMSec + 200
        hwndRegedit = FindWindow(StrPtr("RegEdit_RegEdit"), 0)
    Loop Until hwndRegedit <> 0 Or nMSec > 5000 '5 Sec. timeout
    
    hwndTreeView = FindWindowEx(hwndRegedit, 0, StrPtr("SysTreeView32"), 0)
    hwndListView = FindWindowEx(hwndRegedit, 0, StrPtr("SysListView32"), 0)
    If hwndTreeView = 0 Or hwndListView = 0 Then
        'Unable to start Regedit.
        'MsgBoxW Translate(2100)
        MsgBoxW "Cannot select regedit value. hRegedit = " & hwndRegedit & ", hTreeView = " & hwndTreeView & ", hListView = " & hwndListView
        Exit Sub
    End If
    
    GetWindowThreadProcessId hwndRegedit, iProcId
    
    ShowWindow hwndRegedit, SW_RESTORE
    If iProcId <> 0 Then AllowSetForegroundWindow iProcId
    SetForegroundWindow hwndRegedit
    
    If bVista Then
        'if regedit was already open, collapse any open keys
        For i = 0 To 20
            SendMessage hwndTreeView, WM_KEYDOWN, VK_LEFT, ByVal 0
        Next i
        SendMessage hwndTreeView, WM_KEYDOWN, VK_HOME, ByVal 0
        SendMessage hwndTreeView, WM_KEYDOWN, VK_RIGHT, ByVal 0
    End If
    
    If sValue <> "" And (Not bNoValue) Then
    
        'setting the focus on SysListView32 control
        Dim thrCurrent As Long
        Dim thrRegedit As Long
        
        thrCurrent = GetWindowThreadProcessId(g_HwndMain, ByVal 0)
        thrRegedit = GetWindowThreadProcessId(hwndListView, ByVal 0)
        
        If thrCurrent <> 0 And thrRegedit <> 0 Then
            If AttachThreadInput(thrRegedit, thrCurrent, True) Then
                SetFocus2 hwndListView
                'detach
                AttachThreadInput thrRegedit, thrCurrent, False
            End If
        End If
    
        ' Using of LVM_GETITEMTEXTW and LVM_SETITEMSTATE requires allocation LVITEM within address space of a target process
        ' together with 'string', pszText pointed to
        
        'PS. This routine doesn't want to work for x64, however it should work, if you'll prepare correctly x64-bit structures,
        'like LVITEMW_64. Currently, I unable to do that, so I'm using SendMessage + WM_CHAR method instead.
    
        If OSver.MajorMinor >= 6 And OSver.Bitness = "x32" Then
        
            Dim ItemCount As Long, iItem As Long
            Dim lvi As LVITEMW
            Dim pHandle As Long, ProcessID As Long, Index As Long
            Dim pStrBufferMemory As Long, pMyItemMemory As Long, strLength As Long
            Dim strBuffer As String, result As Long
        
            ItemCount = SendMessage(hwndListView, LVM_GETITEMCOUNT, 0, 0)
            
            If ItemCount > 0 Then
        
                GetWindowThreadProcessId hwndListView, ProcessID
 
                pHandle = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, ProcessID)
    
                If pHandle <> 0 Then
    
                    pStrBufferMemory = VirtualAllocEx(pHandle, 0, 255, MEM_COMMIT, PAGE_READWRITE)

                    If pStrBufferMemory <> 0 Then

                        lvi.iSubItem = 0
                        lvi.cchTextMax = 255
                        lvi.pszText = pStrBufferMemory
                        lvi.Mask = LVIF_TEXT

                        pMyItemMemory = VirtualAllocEx(pHandle, 0, LenB(lvi), MEM_COMMIT, PAGE_READWRITE)

                        If pMyItemMemory <> 0 Then
                        
                            result = WriteProcessMemory(pHandle, pMyItemMemory, lvi, LenB(lvi), 0)
                            
                            If result Then

                                For iItem = 0 To ItemCount - 1

                                    strLength = SendMessage(hwndListView, LVM_GETITEMTEXTW, iItem, ByVal pMyItemMemory)
                                    
                                    If strLength Then

                                        strBuffer = String$(strLength, 0)
                                        
                                        result = ReadProcessMemory(pHandle, pStrBufferMemory, ByVal StrPtr(strBuffer), strLength * 2, 0)
                                        
                                        If StrComp(strBuffer, sValue, 1) = 0 Then
                                            
                                            lvi.Mask = LVIF_STATE
                                            lvi.State = LVIS_SELECTED Or LVIS_FOCUSED
                                            lvi.stateMask = LVIS_SELECTED Or LVIS_FOCUSED
                                            result = WriteProcessMemory(pHandle, pMyItemMemory, lvi, LenB(lvi), 0)
                                            
                                            If result Then
                                                result = SendMessage(hwndListView, LVM_SETITEMSTATE, iItem, ByVal pMyItemMemory)
                                            End If
                                            
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            result = VirtualFreeEx(pHandle, pMyItemMemory, 0&, MEM_RELEASE)
                        End If
                        result = VirtualFreeEx(pHandle, pStrBufferMemory, 0&, MEM_RELEASE)
                    End If
                    result = CloseHandle(pHandle)
                End If
            End If
        Else
            If bVista Then
                'type out the key we want to jump to
                For i = 1 To Len(sFullKey)
                    Select Case Mid$(sFullKey, i, 1)
                        Case "\": SendMessage hwndTreeView, WM_KEYDOWN, VK_RIGHT, ByVal 0
                        Case Else: SendMessage hwndTreeView, WM_CHAR, Asc(Mid$(sFullKey, i, 1)), ByVal 0
                    End Select
                    DoEvents
                    Sleep 50
                Next i
                SendMessage hwndTreeView, WM_KEYDOWN, VK_RIGHT, ByVal 0
            End If
            'type out the value we want to jump to
            For i = 1 To Len(sValue)
                SendMessage hwndListView, WM_CHAR, Asc(Mid$(sValue, i, 1)), ByVal 0
                DoEvents
                Sleep 50
            Next i
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Jump", lHive, sKey, sValue, bUseWow64
    If inIDE Then Stop: Resume Next
End Sub

Public Function WrapNtOpenKeyEx( _
    hHive As ENUM_REG_HIVE, _
    sKey As String, _
    dwAccess As Long, _
    retHandle As Long, _
    Optional dwAttributes As OBJ_ATTRIBUTES = OBJ_CASE_INSENSITIVE, _
    Optional bUseWow64 As Boolean) As Long
    
    Dim bOS_32          As Boolean
    Dim OA              As OBJECT_ATTRIBUTES
    Dim UniStr          As UNICODE_STRING
    Dim lret            As Long
    Dim sNtKey          As String
    
    Call NormalizeKeyNameAndHiveHandle(hHive, sKey)
    
    bOS_32 = Not bIsWOW64
    
    sNtKey = MapHiveToNtName(hHive) & IIf(Len(sKey) = 0, "", "\" & sKey)
    
    If Not bOS_32 And Not bUseWow64 Then
        'walkaround for WOW since Native API doesn't support KEY_WOW64_64KEY access flag (thanks to David Heffernan (StackOverflow))
        '(opening root with Win32 API first + KEY_WOW64_64KEY, then pass the handle to the Native API)
        lret = RegOpenKeyEx(hHive, 0&, 0&, dwAccess Or (KEY_WOW64_64KEY), OA.RootDirectory)
        
        If OA.RootDirectory <> 0 Then
            sNtKey = sKey
        Else
            Debug.Print "Error happened when opening reg. root node. Code = " & lret
        End If
    End If
    
    With UniStr
        .Length = LenB(sNtKey)
        .MaximumLength = .Length
        .Buffer = StrPtr(sNtKey)
    End With
    
    With OA
        .Length = LenB(OA)
        .ObjectName = VarPtr(UniStr)
        If bOS_32 Then
            .RootDirectory = 0
        End If
        .Attributes = dwAttributes
        .SecurityDescriptor = 0
        .SecurityQualityOfService = 0
    End With
    
    '// TODO: ensure SeBackupPrivilege is granted
    
    WrapNtOpenKeyEx = NtOpenKeyEx(retHandle, dwAccess, OA, REG_OPTION_BACKUP_RESTORE) '// TODO: check it
    'WrapNtOpenKeyEx = NtOpenKey(retHandle, dwAccess, OA)
    
    If STATUS_SUCCESS <> lret Then
        Debug.Print "NtOpenKey(Ex) returns: 0x" & Hex(lret)
    End If
    
    If OA.RootDirectory <> 0 Then RegCloseKey OA.RootDirectory
End Function

Public Function IsKeySymLink(hHive As ENUM_REG_HIVE, sKey As String, Optional out_KeyTarget As String) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim hKey1           As Long
    Dim hKeyLink        As Long
    Dim sNtName1        As String
    Dim sNtNameLink     As String
    Dim sNtKey          As String
    Dim lret            As Long
    Dim bOS_32          As Boolean
    Dim OA              As OBJECT_ATTRIBUTES
    Dim UniStr          As UNICODE_STRING
    Dim cData           As Long
    
    Call NormalizeKeyNameAndHiveHandle(hHive, sKey)
    
    'you should obtain SeBackupPrivilege for REG_OPTION_BACKUP_RESTORE
    
    bOS_32 = Not bIsWOW64
    
    'open key as native (x64)
    lret = RegOpenKeyEx(hHive, StrPtr(sKey), 0&, WRITE_OWNER Or (KEY_WOW64_64KEY And Not bOS_32), hKey1)
    If ERROR_SUCCESS <> lret Then
        'key is not exist
        Exit Function
    End If
    
    'get native phisical key name
    sNtName1 = RegGetKeyInfoNameEx(hKey1)
    RegCloseKey hKey1: hKey1 = 0
    
    If Len(sNtName1) = 0 Then
        'unknown error
        Debug.Print "Error in IsKeySymLink when open hKey1"
        Exit Function
    End If
    
    sNtKey = MapHiveToNtName(hHive) & "\" & sKey
    
    If Not bOS_32 Then
        'walkaround for WOW since Native API doesn't support KEY_WOW64_64KEY access flag (thanks to David Heffernan (StackOverflow))
        '(opening root with Win32 API first + KEY_WOW64_64KEY, then pass the handle to the Native API)
        lret = RegOpenKeyEx(hHive, 0&, 0&, WRITE_OWNER Or (KEY_WOW64_64KEY), OA.RootDirectory)
        
        If OA.RootDirectory <> 0 Then
            sNtKey = sKey
        Else
            Debug.Print "Error happened when opening root node 'SOFTWARE'. Code = " & lret
        End If
    End If
    
    With UniStr
        .Length = LenB(sNtKey)
        .MaximumLength = .Length
        .Buffer = StrPtr(sNtKey)
    End With
    
    With OA
        .Length = LenB(OA)
        .ObjectName = VarPtr(UniStr)
        If bOS_32 Then
            .RootDirectory = 0
        End If
        .Attributes = OBJ_CASE_INSENSITIVE Or OBJ_OPENLINK
        .SecurityDescriptor = 0
        .SecurityQualityOfService = 0
    End With
    
    'open key as a source of a symlink
    'lRet = NtOpenKeyEx(hKeyLink, WRITE_OWNER, OA, REG_OPTION_OPEN_LINK Or REG_OPTION_BACKUP_RESTORE)
    lret = NtOpenKey(hKeyLink, KEY_READ, OA)
    
    If STATUS_SUCCESS = lret Then
        sNtNameLink = RegGetKeyInfoNameEx(hKeyLink)
        
        'phisical names are different ?
        If StrComp(sNtName1, sNtNameLink, 1) <> 0 Then
        
            'key is a symlink
            IsKeySymLink = True
            
            '... here we'll get a target of a symlink by quering "SymbolicLinkValue" value
            
            'get size of buffer required
            lret = RegQueryValueEx(hKeyLink, StrPtr("SymbolicLinkValue"), 0&, ByVal 0&, ByVal 0&, cData)
            
            If ERROR_SUCCESS = lret Or ERROR_MORE_DATA = lret Then
                
                out_KeyTarget = String$(cData \ 2 + 1, 0&)
                lret = RegQueryValueEx(hKeyLink, StrPtr("SymbolicLinkValue"), 0&, ByVal 0&, StrPtr(out_KeyTarget), cData)
                
                If ERROR_SUCCESS = lret Then
                    out_KeyTarget = Left$(out_KeyTarget, lstrlen(StrPtr(out_KeyTarget)))
                    '// TODO.
                    'add RegMapNtNameToHive
                Else
                    out_KeyTarget = ""
                End If
            Else
                Debug.Print "Error in query SymbolicLinkValue for: " & hHive & "\" & sKey & ". Code = " & lret
            End If
        End If
    End If
    
    If hKeyLink <> 0 Then NtClose hKeyLink: hKeyLink = 0
    If OA.RootDirectory <> 0 Then RegCloseKey OA.RootDirectory: OA.RootDirectory = 0
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IsKeySymLink", hHive, sKey
    If inIDE Then Stop: Resume Next
    If hKey1 <> 0 Then RegCloseKey hKey1
    If hKeyLink <> 0 Then RegCloseKey hKeyLink
    If OA.RootDirectory <> 0 Then RegCloseKey OA.RootDirectory
End Function


Public Function GetKeyVirtualType( _
    hHive As ENUM_REG_HIVE, _
    sKey As String, _
    Optional bFlagCheckForSymlink As Boolean = False, _
    Optional out_SymlinkTarget As String) As KEY_VIRTUAL_TYPE
    
    On Error GoTo ErrorHandler:
    
    Dim hKey1           As Long
    Dim hKey2           As Long
    Dim sNtName1        As String
    Dim sNtName2        As String
    Dim sKeyWOW         As String
    Dim lret            As Long
    Dim bOS_32          As Boolean
    Dim pos             As Long
    Dim bHKU_Class      As Boolean
    
    If bFlagCheckForSymlink Then
        If IsKeySymLink(hHive, sKey, out_SymlinkTarget) Then
        GetKeyVirtualType = KEY_VIRTUAL_SYMLINK
        End If
    End If
    
    'check OS bitness
    bOS_32 = Not bIsWOW64
    
    'For reliability: SeTakeOwnershipPrivilege should be taken first.
    '
    'Note: this routine doesn't do any write operation. WRITE_OWNER is just a protection against restricted DACL.
    
    'open key as native (x64)
    lret = RegOpenKeyEx(hHive, StrPtr(sKey), 0&, WRITE_OWNER Or (KEY_WOW64_64KEY And Not bOS_32), hKey1)
    If ERROR_SUCCESS <> lret Then
        'native key is not exist
        If bOS_32 Then
            GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_NOT_EXIST
            Exit Function
        End If
        'check if redirected key exist
        lret = RegOpenKeyEx(hHive, StrPtr(sKey), 0&, WRITE_OWNER, hKey2)
        If ERROR_SUCCESS = lret Then
            'if alternate view exists -> key is redirected
            GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_REDIRECTED
            RegCloseKey hKey2
        Else
            GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_NOT_EXIST
        End If
        Exit Function
    End If
    
    'if we are trying to check a root key
    If 0 = Len(sKey) Then
        GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_USUAL
        Exit Function
    End If
    
    'get native phisical key name
    sNtName1 = RegGetKeyInfoNameEx(hKey1)
    RegCloseKey hKey1: hKey1 = 0
    
    If Len(sNtName1) = 0 Then
        'unknown error
        GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_USUAL
        Debug.Print "Error in GetKeyVirtualType when open hKey1"
        Exit Function
    End If
    
    If bOS_32 Then
        'OS bit. == x86
        GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_USUAL
        Exit Function
    End If
    
    'open key as WOW
    lret = RegOpenKeyEx(hHive, StrPtr(sKey), 0&, WRITE_OWNER, hKey2)
    If ERROR_SUCCESS <> lret Then
        'x64 opened, but WOW is not -> it's a redirected key with no corresponding WOW para
        GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_REDIRECTED
        Exit Function
    End If
    
    'get native phisical key name
    sNtName2 = RegGetKeyInfoNameEx(hKey2)
    RegCloseKey hKey2: hKey2 = 0
    
    'success query names?
    If Len(sNtName1) <> 0 And Len(sNtName2) <> 0 Then
        'redirected ? (phisical keys should become different)
        ' !=
        If StrComp(sNtName1, sNtName2, 1) <> 0 Then
            'it's a redirected key
            GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_REDIRECTED
        Else
            'check whether it is shared or usual key
            'convert name to WOW
            'only "Software" node can be redirected according to MSDN
            If StrBeginWith(sKey, "SOFTWARE\") Or StrComp(sKey, "SOFTWARE", 1) = 0 Then
                'append Wow6432Node manually
                'no matter where Wow6432Node to write, after "Software" or after "Classes" (there is always a symlink and NtQueryKey will normalize it)
                'we'll insert "Wow6432Node" inside phisical name (NtName), because original name can present a symlink,
                'so SOFTWARE\Wow6432Node\classes\wow6432node\appid will not be a correct registry path here
                
                'Note: For HKU hives, "Software\Classes" is stored in NT namespace as "\REGISTRY\USER\<SID>_CLASSES\"
                bHKU_Class = False
                If StrBeginWith(sNtName1, "\REGISTRY\USER\") Then
                    If StrBeginWith(sKey, "SOFTWARE\CLASSES\") Or StrComp(sKey, "SOFTWARE\CLASSES", 1) = 0 Then
                        bHKU_Class = True
                        pos = InStr(1, sNtName1, "_CLASSES\", 1)
                        If pos <> 0 Then
                            sKeyWOW = Mid$(sNtName1, Len("\REGISTRY\USER\") + 1)
                            sKeyWOW = Replace$(sKeyWOW, "_CLASSES\", "_CLASSES\Wow6432Node\", 1, 1, 1)
                        Else
                            'unknown error. <SID>_CLASSES should always be there. Possible, an unknown symlink (?)
                            Debug.Print "No <SID>_CLASSES node in phisical registry path after quering by RegGetKeyInfoNameEx. Key: " & sKey
                        End If
                    End If
                End If
                
                If Not bHKU_Class Then
                    pos = InStr(1, sNtName1, "SOFTWARE", 1)
                    If pos <> 0 Then
                        sKeyWOW = "SOFTWARE\Wow6432Node\" & Mid$(sNtName1, pos + Len("SOFTWARE\"))
                    Else
                        'unknown error. SOFTWARE should always be there. Possible, an unknown symlink (?)
                        Debug.Print "No SOFTWARE node in phisical registry path after quering by RegGetKeyInfoNameEx. Key: " & sKey
                    End If
                End If
                
                'open Wow6432Node key as native
                If ERROR_SUCCESS <> RegOpenKeyEx(hHive, StrPtr(sKeyWOW), 0&, WRITE_OWNER Or (KEY_WOW64_64KEY), hKey2) Then
                    'second key doesn't exist (however, for 'shared' keys it always should be there)
                    'so, it's usual key
                    GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_USUAL
                    Exit Function
                End If
                
                sNtName2 = RegGetKeyInfoNameEx(hKey2)
                RegCloseKey hKey2: hKey2 = 0
                
                'comparing where Wow6432Node and native keys point to
                ' ==
                If StrComp(sNtName1, sNtName2, 1) = 0 Then
                    'handles point to the same phisical name -> it's a shared key
                    GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_SHARED
                Else
                    'usual key
                    GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_USUAL
                End If
            Else
                'usual key
                GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_USUAL
            End If
        End If
    Else
        'error happened during query
        GetKeyVirtualType = GetKeyVirtualType Or KEY_VIRTUAL_USUAL
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetKeyVirtualType", hHive, sKey
    If inIDE Then Stop: Resume Next
    If hKey1 <> 0 Then RegCloseKey hKey1
    If hKey2 <> 0 Then RegCloseKey hKey2
End Function


Public Function NtGetKeyVirtualType( _
    hHive As ENUM_REG_HIVE, _
    sKey As String, _
    Optional bFlagCheckForSymlink As Boolean = False, _
    Optional out_SymlinkTarget As String) As KEY_VIRTUAL_TYPE
    
    On Error GoTo ErrorHandler:
    
    Dim hKey1           As Long
    Dim hKey2           As Long
    Dim sNtName1        As String
    Dim sNtName2        As String
    Dim sKeyWOW         As String
    Dim lret            As Long
    Dim bOS_32          As Boolean
    Dim pos             As Long
    Dim bHKU_Class      As Boolean
    
    If bFlagCheckForSymlink Then
        If IsKeySymLink(hHive, sKey, out_SymlinkTarget) Then
        NtGetKeyVirtualType = KEY_VIRTUAL_SYMLINK
        End If
    End If
    
    'check OS bitness
    bOS_32 = Not bIsWOW64
    
    'For reliability: SeTakeOwnershipPrivilege should be taken first.
    '
    'Note: this routine doesn't do any write operation. WRITE_OWNER is just a protection against restricted DACL.
    
    'open key as native (x64)
    lret = WrapNtOpenKeyEx(hHive, sKey, WRITE_OWNER, hKey1)
    If STATUS_SUCCESS <> lret Then
        'native key is not exist
        If bOS_32 Then
            NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_NOT_EXIST
            Exit Function
        End If
        'check if redirected key exist
        lret = WrapNtOpenKeyEx(hHive, sKey, WRITE_OWNER, hKey2, , True)
        If STATUS_SUCCESS = lret Then
            'if alternate view exists -> key is redirected
            NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_REDIRECTED
            NtClose hKey2
        Else
            NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_NOT_EXIST
        End If
        Exit Function
    End If
    
    'if we are trying to check a root key
    If 0 = Len(sKey) Then
        NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_USUAL
        Exit Function
    End If
    
    'get native phisical key name
    sNtName1 = RegGetKeyInfoNameEx(hKey1)
    NtClose hKey1: hKey1 = 0
    
    If Len(sNtName1) = 0 Then
        'unknown error
        NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_USUAL
        Debug.Print "Error in NtGetKeyVirtualType when open hKey1"
        Exit Function
    End If
    
    If bOS_32 Then
        'OS bit. == x86
        NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_USUAL
        Exit Function
    End If
    
    'open key as WOW
    lret = WrapNtOpenKeyEx(hHive, sKey, WRITE_OWNER, hKey2, , True)
    If STATUS_SUCCESS <> lret Then
        'x64 opened, but WOW is not -> it's a redirected key with no corresponding WOW para
        NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_REDIRECTED
        Exit Function
    End If
    
    'get native phisical key name
    sNtName2 = RegGetKeyInfoNameEx(hKey2)
    NtClose hKey2: hKey2 = 0
    
    'success query names?
    If Len(sNtName1) <> 0 And Len(sNtName2) <> 0 Then
        'redirected ? (phisical keys should become different)
        ' !=
        If StrComp(sNtName1, sNtName2, 1) <> 0 Then
            'it's a redirected key
            NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_REDIRECTED
        Else
            'check whether it is shared or usual key
            'convert name to WOW
            'only "Software" node can be redirected according to MSDN
            If StrBeginWith(sKey, "SOFTWARE\") Or StrComp(sKey, "SOFTWARE", 1) = 0 Then
                'append Wow6432Node manually
                'no matter where Wow6432Node to write, after "Software" or after "Classes" (there is always a symlink and NtQueryKey will normalize it)
                'we'll insert "Wow6432Node" inside phisical name (NtName), because original name can present a symlink,
                'so SOFTWARE\Wow6432Node\classes\wow6432node\appid will not be a correct registry path here
                
                'Note: For HKU hives, "Software\Classes" is stored in NT namespace as "\REGISTRY\USER\<SID>_CLASSES\"
                bHKU_Class = False
                If StrBeginWith(sNtName1, "\REGISTRY\USER\") Then
                    If StrBeginWith(sKey, "SOFTWARE\CLASSES\") Or StrComp(sKey, "SOFTWARE\CLASSES", 1) = 0 Then
                        bHKU_Class = True
                        pos = InStr(1, sNtName1, "_CLASSES\", 1)
                        If pos <> 0 Then
                            sKeyWOW = Mid$(sNtName1, Len("\REGISTRY\USER\") + 1)
                            sKeyWOW = Replace$(sKeyWOW, "_CLASSES\", "_CLASSES\Wow6432Node\", 1, 1, 1)
                        Else
                            'unknown error. <SID>_CLASSES should always be there. Possible, an unknown symlink (?)
                            Debug.Print "No <SID>_CLASSES node in phisical registry path after quering by RegGetKeyInfoNameEx. Key: " & sKey
                        End If
                    End If
                End If
                
                If Not bHKU_Class Then
                    pos = InStr(1, sNtName1, "SOFTWARE", 1)
                    If pos <> 0 Then
                        sKeyWOW = "SOFTWARE\Wow6432Node\" & Mid$(sNtName1, pos + Len("SOFTWARE\"))
                    Else
                        'unknown error. SOFTWARE should always be there. Possible, an unknown symlink (?)
                        Debug.Print "No SOFTWARE node in phisical registry path after quering by RegGetKeyInfoNameEx. Key: " & sKey
                    End If
                End If
                
                'open Wow6432Node key as native
                If STATUS_SUCCESS <> WrapNtOpenKeyEx(hHive, sKeyWOW, WRITE_OWNER, hKey2) Then
                    'second key doesn't exist (however, for 'shared' keys it always should be there)
                    'so, it's usual key
                    NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_USUAL
                    Exit Function
                End If
                
                sNtName2 = RegGetKeyInfoNameEx(hKey2)
                NtClose hKey2: hKey2 = 0
                
                'comparing where Wow6432Node and native keys point to
                ' ==
                If StrComp(sNtName1, sNtName2, 1) = 0 Then
                    'handles point to the same phisical name -> it's a shared key
                    NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_SHARED
                Else
                    'usual key
                    NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_USUAL
                End If
            Else
                'usual key
                NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_USUAL
            End If
        End If
    Else
        'error happened during query
        NtGetKeyVirtualType = NtGetKeyVirtualType Or KEY_VIRTUAL_USUAL
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "NtGetKeyVirtualType", hHive, sKey
    If inIDE Then Stop: Resume Next
    If hKey1 <> 0 Then NtClose hKey1
    If hKey2 <> 0 Then NtClose hKey2
End Function

'// Getting the phisical location of the key by handle (thanks to greg zaharov, jupman, jr_ (CyberForum))
Private Function RegGetKeyInfoNameEx(hKey As Long) As String
    On Error GoTo ErrorHandler:

    Const STATUS_SUCCESS  As Long = 0&

    Dim lret As Long
    Dim buf() As Long
    Dim cchBuf As Long
    Dim reqSize As Long
    Dim sKeyName As String
    Dim strLength As Long
    
    If hKey = 0 Then Exit Function
    
    ReDim buf(0)
    buf(0) = 0
    cchBuf = 0
    
    lret = NtQueryKey(hKey, KeyNameInformation, ByVal VarPtr(buf(0)), cchBuf, reqSize)
    
    If lret = STATUS_BUFFER_TOO_SMALL Then
        ReDim buf(reqSize \ 4)
        cchBuf = (UBound(buf) + 1) * 4
        lret = NtQueryKey(hKey, KeyNameInformation, ByVal VarPtr(buf(0)), cchBuf, reqSize)
        
        If lret = STATUS_SUCCESS Then
            strLength = buf(0)
            If strLength <> 0 Then
                sKeyName = String$(strLength \ 2, 0)
                memcpy ByVal StrPtr(sKeyName), buf(1), strLength
                RegGetKeyInfoNameEx = sKeyName
            Else
                Debug.Print "Error in RegGetKeyInfoNameEx (strLength = 0), code = 0x" & Hex(lret)
            End If
        Else
            Debug.Print "Error in RegGetKeyInfoNameEx (NtQueryKey, second call), code = 0x" & Hex(lret)
        End If
    Else
        Debug.Print "Error in RegGetKeyInfoNameEx (NtQueryKey, first call), code = 0x" & Hex(lret)
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegGetKeyInfoNameEx", hKey
    If inIDE Then Stop: Resume Next
End Function


'// e.g. HKLM -> \Registry\Machine
Public Function MapHiveToNtName(hHive As ENUM_REG_HIVE) As String
    On Error GoTo ErrorHandler:
    
    Dim hKey As Long
    Dim lret As Long
    Dim bOS_32 As Boolean
    
    bOS_32 = Not bIsWOW64
    
    lret = RegOpenKeyEx(hHive, 0&, 0&, KEY_QUERY_VALUE Or (KEY_WOW64_64KEY And Not bOS_32), hKey)
    
    If ERROR_SUCCESS <> lret Then
        Debug.Print "Error in RegOpenKeyEx when open hKey1, code = " & lret
        Exit Function
    End If
    
    MapHiveToNtName = RegGetKeyInfoNameEx(hKey)
    RegCloseKey hKey
    
    '// TODO: maybe replace by manual parsing?
    
    'HKLM = \REGISTRY\MACHINE
    'HKCU = \REGISTRY\USER\<SID>
    'HKU = \REGISTRY\USER
    'HKCC = \REGISTRY\MACHINE\SYSTEM\ControlSet<Current>\Hardware Profiles\<CurrentID>
    '     or by symlink: "\REGISTRY\MACHINE\SYSTEM\CurrentControlSet\Hardware Profiles\Current"
    'HKCR = {depending on token}:
    ' 1. User:   \REGISTRY\USER\<SID>_CLASSES
    ' 2. System: \Registry\Machine\SOFTWARE\Classes
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MapHiveToNtName", hHive
    If inIDE Then Stop: Resume Next
End Function


'// e.g. \Registry\Machine\Software -> HKLM, Software
Public Function MapNtNameToHive(sNtKey As String, out_Hive As ENUM_REG_HIVE, out_Key As String) As Boolean
    On Error GoTo ErrorHandler:
    
    '// TODO

    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegMapNtNameToHive", sNtKey
    If inIDE Then Stop: Resume Next
End Function

Public Function Flush(hHive As ENUM_REG_HIVE) As Boolean
    Flush = (0 = RegFlushKey(hHive))
End Function

Public Function FlushAll() As Boolean
    FlushAll = True
    FlushAll = FlushAll And (0 = RegFlushKey(HKEY_CURRENT_USER))
    FlushAll = FlushAll And (0 = RegFlushKey(HKEY_LOCAL_MACHINE))
    FlushAll = FlushAll And (0 = RegFlushKey(HKEY_USERS))
End Function

Private Sub KeyShowInfo(hHive As Long, sSubKey As String, hFile As Long)
    Dim sSymTarget As String
    Dim sLog As String
    Dim bExist As Boolean
    Dim bExist2 As Boolean
    Dim hKey As Long
    Dim hKey2 As Long

    bExist = False
    If STATUS_SUCCESS = WrapNtOpenKeyEx(hHive, sSubKey, WRITE_OWNER, hKey, , False) Then
        bExist = True
        NtClose hKey
    End If

    bExist2 = False
    If STATUS_SUCCESS = WrapNtOpenKeyEx(hHive, sSubKey, WRITE_OWNER, hKey, , True) Then
        bExist2 = True
        NtClose hKey2
    End If

    If bExist <> bExist2 Then
        sLog = "Warning: Key exist (with flag KEY_WOW64_64KEY) returns " & bExist & ", when without flag it returns " & bExist2
        PrintW hFile, sLog, True
    End If

    Call NtGetKeyVirtualType(hHive, sSubKey, True, sSymTarget)

    sLog = "HKLM\" & sSubKey & " => " & IIf(bExist, " (exist, handle = " & hKey & ")", "(not exist, handle = " & hKey & ")") & _
        IIf(Len(sSymTarget) <> 0, " => (symlink) " & sSymTarget, "")

    Debug.Print sLog
End Sub

