VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOSInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'[clsOSInfo.cls]

Option Explicit

' OSInfo class by Alex Dragokas
' ver 1.10
'

'Tips:

' Versions:

' 10.0* | Windows Server 2016
' 10.0* | Windows 10 (builds 9885+, Insider Preview)
' 6.4*  | Windows 10 Technical Preview (builds 9879-)
' 6.3*  | Windows Server 2012 R2
' 6.3*  | Windows 8.1
' 6.2   | Windows Server 2012
' 6.2   | Windows 8
' 6.1   | Windows Server 2008 R2
' 6.1   | Windows 7
' 6.0   | Windows Server 2008
' 6.0   | Windows Vista
' 5.2   | Windows Server 2003 R2
' 5.2   | Windows Server 2003
' 5.2   | Windows XP 64-Bit Edition
' 5.1   | Windows XP
' 5.0   | Windows 2000
' 4.90  | Windows Me
' 4.10  | Windows 98
' 4.0   | Windows NT Workstation 4.0
' 4.0/4.03 | Windows 95

' Last Service Packs:

' 10     | SP0, ReleaseId=1803+, Build=17134+ (Creators Update)
' 8.1    | SP0, Update 1+, Build=9600
' 8      | SP0, Build=9200
' 7      | SP1, Build=7601
' Vista  | SP2, Build=6002
' XP x64 | SP2, Build=3790
' XP x32 | SP3, Build=2600
' 2000   | SP4, Build=2195
' ME     | Build=3000
' NT 4.0 | SP6a, Build=1381
' 98     | Build=2222 A
' 95     | Build=950

Private Enum ePlatformID
    Win32S = 0
    Win32Windows = 1
    Win32NT = 2
    WinCE = 3
    UNIX = 4
    Xbox = 5
    MacOSX = 6
End Enum

Private Type RTL_OSVERSIONINFOEXW
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(127) As Integer
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type OSVERSIONINFOEXW
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(127) As Integer
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

Private Type SID_AND_ATTRIBUTES
    SID As Long
    Attributes As Long
End Type

Private Type TOKEN_GROUPS
    GroupCount As Long
    Groups(20) As SID_AND_ATTRIBUTES
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount  As Long
    LuidLowPart     As Long
    LuidHighPart    As Long
    Attributes      As Long
End Type

Private Type OBJECT_ATTRIBUTES
    Length                   As Long
    RootDirectory            As Long
    ObjectName               As Long
    Attributes               As Long
    SecurityDescriptor       As Long
    SecurityQualityOfService As Long
End Type

Private Type UNICODE_STRING
    Length                          As Integer
    MaximumLength                   As Integer
    Buffer                          As Long
End Type

Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoW" (ByVal lcid As Long, ByVal LCTYPE As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDefaultUILanguage Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32.dll" () As Long
'Private Declare Function GetUserDefaultLCID Lib "kernel32.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32.dll" (ByVal hProc As Long, bWow64Process As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As RTL_OSVERSIONINFOEXW) As Long 'OSVERSIONINFOEXW
Private Declare Function RtlGetVersion Lib "ntdll.dll" (lpVersionInformation As RTL_OSVERSIONINFOEXW) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetProductInfo Lib "kernel32.dll" (ByVal dwOSMajorVersion As Long, ByVal dwOSMinorVersion As Long, ByVal dwSpMajorVersion As Long, ByVal dwSpMinorVersion As Long, pdwReturnedProductType As Long) As Long
'Private Declare Function CheckTokenMembership Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal SidToCheck As Long, IsMember As Long) As Long
Private Declare Function OpenProcessToken Lib "Advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "Advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
'Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
'Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal lSize As Long)
Private Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Sub FreeSid Lib "Advapi32.dll" (ByVal pSid As Long)
Private Declare Function AllocateAndInitializeSid Lib "Advapi32.dll" (pIdentifierAuthority As Any, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function IsValidSid Lib "Advapi32.dll" (ByVal pSid As Long) As Long
Private Declare Function GetSidSubAuthority Lib "Advapi32.dll" (ByVal pSid As Long, ByVal nSubAuthority As Long) As Long
Private Declare Function GetSidSubAuthorityCount Lib "Advapi32.dll" (ByVal pSid As Long) As Long
Private Declare Function EqualSid Lib "Advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
Private Declare Function RegOpenKeyEx Lib "Advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExStr Lib "Advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal szData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "Advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, szData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "Advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32.dll" (ByVal lpStrDest As Long, ByVal lpStrSrc As Long) As Long
Private Declare Function ConvertSidToStringSid Lib "Advapi32.dll" Alias "ConvertSidToStringSidW" (ByVal pSid As Long, StringSid As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Sub RtlGetNtVersionNumbers Lib "ntdll.dll" (Major As Long, Minor As Long, Build As Long)
Private Declare Function LookupPrivilegeValue Lib "Advapi32.dll" Alias "LookupPrivilegeValueW" (ByVal lpSystemName As Long, ByVal lpName As Long, lpLuid As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "Advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long
Private Declare Function OpenThreadToken Lib "Advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
Private Declare Function NtOpenSymbolicLinkObject Lib "ntdll.dll" (LinkHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As OBJECT_ATTRIBUTES) As Long
Private Declare Function NtOpenDirectoryObject Lib "ntdll.dll" (DirectoryHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As OBJECT_ATTRIBUTES) As Long
Private Declare Function NtClose Lib "ntdll.dll" (ByVal Handle As Long) As Long
Private Declare Function IsOS Lib "Shlwapi.dll" Alias "#437" (ByVal dwOS As Long) As Long

Private Const SM_SERVERR2               As Long = 89&
Private Const SM_MEDIACENTER            As Long = 87&
Private Const SM_STARTER                As Long = 88&
Private Const SM_TABLETPC               As Long = 86&
Private Const SM_CLEANBOOT              As Long = 67&

Private Const VER_SUITE_STORAGE_SERVER  As Long = &H2000&
Private Const VER_SUITE_DATACENTER      As Long = &H80&
Private Const VER_SUITE_PERSONAL        As Long = &H200&
Private Const VER_SUITE_ENTERPRISE      As Long = 2&
Private Const VER_SUITE_BACKOFFICE      As Long = 4
Private Const VER_SUITE_BLADE           As Long = &H400&
Private Const VER_SUITE_COMPUTE_SERVER  As Long = &H4000&
Private Const VER_SUITE_EMBEDDEDNT      As Long = &H40&
Private Const VER_SUITE_SINGLEUSERTS    As Long = &H100&
Private Const VER_SUITE_SMALLBUSINESS   As Long = 1
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20&
Private Const VER_SUITE_TERMINAL        As Long = &H10&
Private Const VER_SUITE_WH_SERVER       As Long = &H8000&
Private Const VER_NT_WORKSTATION        As Long = 1&
Private Const VER_NT_DOMAIN_CONTROLLER  As Long = 2&
Private Const VER_NT_SERVER             As Long = 3&
Private Const LOCALE_SYSTEM_DEFAULT     As Long = &H800&
Private Const LOCALE_USER_DEFAULT       As Long = &H400&
Private Const LOCALE_SENGLANGUAGE       As Long = &H1001&
Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Private Const KEY_WOW64_64KEY           As Long = &H100&
Private Const DIRECTORY_QUERY           As Long = 1&

Private Const STATUS_SUCCESS            As Long = 0&
Private Const ERROR_SUCCESS             As Long = 0&
Private Const TOKEN_QUERY               As Long = &H8&
Private Const ERROR_NO_TOKEN            As Long = 1008&
Private Const TokenElevation            As Long = 20&
Private Const TOKEN_ADJUST_PRIVILEGES   As Long = &H20
Private Const SE_PRIVILEGE_ENABLED      As Long = 2&
Private Const STATUS_ACCESS_DENIED      As Long = &HC0000022
Private Const STATUS_OBJECT_NAME_NOT_FOUND As Long = &HC0000034


Dim osi As RTL_OSVERSIONINFOEXW

Dim OSName_                 As String
Dim Bitness_                As String
Dim Edition_                As String
Dim MajorMinor_             As Single
Dim MajorMinorNTDLL_        As Single
Dim NtDllVersion_           As String
Dim Revision_               As Long
Dim SPver_                  As Single
Dim IsSafeBoot_             As Boolean
Dim IsElevated_             As Boolean
Dim IntegrityLevel_         As String
Dim UserType_               As String
Dim UserName_               As String
Dim ComputerName_               As String
Dim IsWindowsXPOrGreater_   As Boolean
Dim IsWindowsVistaOrGreater_ As Boolean
Dim IsWindows7OrGreater_    As Boolean
Dim IsWindows8OrGreater_    As Boolean
Dim IsWindows8Point1OrGreater_ As Boolean
Dim IsWindows10OrGreater_   As Boolean
Dim LangSystemName_         As String
Dim LangSystemNameFull_     As String
Dim LangSystemCode_         As Long
Dim LangDisplayName_        As String
Dim LangDisplayNameFull_    As String
Dim LangDisplayCode_        As Long
Dim LangNonUnicodeName_     As String
Dim LangNonUnicodeNameFull_ As String
Dim LangNonUnicodeCode_     As Long
Dim CodepageOEM_            As String
Dim CodepageANSI_           As String
Dim CodepageOEM_File_       As String
Dim CodepageANSI_File_      As String
Dim SuiteMask_              As String 'As Integer
Dim ProductType_            As String 'As Byte
Dim SID_CurrentProcess_     As String
Dim SafeBootMode_           As String
Dim Platform_               As String 'Win9x / WinNT
Dim ReleaseId_              As Long
Dim IsServer_               As Boolean
Dim IsDomainController_     As Boolean
Dim inIDE                   As Boolean
Dim IsWin64_                As Boolean
Dim SecureBoot_             As Boolean
Dim PlatformID_             As Long
Dim IsLocalSystemContext_   As Boolean
Dim IsAdmin_                As Boolean
Dim IsSystemCaseSensitive_  As Boolean
Dim IsEmbedded_             As Boolean


Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "clsOSver.Class_Initialize - Begin"
    
    Dim dec             As Single
    Dim ProductType     As Long
    Dim lret            As Long
    Dim pos             As Long
    Dim tmp             As String
    Dim hLib            As Long
    Dim dwMajor         As Long
    Dim dwMinor         As Long
    Dim dwBuild         As Long
    
    inIDE = (App.LogMode = 0)
    
    LangDisplayCode_ = GetUserDefaultUILanguage Mod &H10000
    LangDisplayName_ = GetLangNameByCultureCode(LangDisplayCode_)
    LangDisplayNameFull_ = GetLangNameFullByCultureCode(LangDisplayCode_)
    
    LangSystemCode_ = GetSystemDefaultUILanguage Mod &H10000
    LangSystemName_ = GetLangNameByCultureCode(LangSystemCode_)
    LangSystemNameFull_ = GetLangNameFullByCultureCode(LangSystemCode_)
    
    LangNonUnicodeCode_ = GetSystemDefaultLCID Mod &H10000
    LangNonUnicodeName_ = GetLangNameByCultureCode(LangNonUnicodeCode_)
    LangNonUnicodeNameFull_ = GetLangNameFullByCultureCode(LangNonUnicodeCode_)
    
    osi.dwOSVersionInfoSize = Len(osi)
    
    If (STATUS_SUCCESS <> RtlGetVersion(osi)) Then
        GetVersionEx osi
    End If
    
    If inIDE Then
        'query correct OS version without manifest
        Dim oWMI As Object
        Dim colOS As Object
        Dim oOS As Object
        Dim Ver() As String
    
        Set oWMI = CreateObject("winmgmts:{impersonationLevel=Impersonate}!\\.\root\cimv2")
        Set colOS = oWMI.ExecQuery("Select * from Win32_OperatingSystem")
        
        For Each oOS In colOS
            pos = InStr(oOS.Version, ".")
            If pos = 0 Then
                osi.dwMajorVersion = CLng(oOS.Version)
                osi.dwMinorVersion = 0
                osi.dwBuildNumber = 0
            Else
                Ver = Split(oOS.Version, ".")
                If UBound(Ver) = 1 Then
                    osi.dwMajorVersion = CLng(Ver(0))
                    osi.dwMinorVersion = CLng(Ver(1))
                    osi.dwBuildNumber = 0
                Else
                    osi.dwMajorVersion = CLng(Ver(0))
                    osi.dwMinorVersion = CLng(Ver(1))
                    osi.dwBuildNumber = CLng(Ver(2))
                End If
            End If
        Next
        Set oOS = Nothing
        Set colOS = Nothing
        Set oWMI = Nothing
    End If
    
    Bitness_ = IIf(IsWow64(), "x64", "x32") 'VB6 is always x86
    
    IsWin64_ = (Bitness_ = "x64")
    
    lret = GetSystemMetrics(SM_CLEANBOOT)
    IsSafeBoot_ = (lret > 0)
    Select Case lret
        Case 1
            SafeBootMode_ = "Safe mode"
        Case 2
            SafeBootMode_ = "Safe mode with network support"
        Case Else
            SafeBootMode_ = "Normal"
    End Select
    
    ' OS Major + Minor
    dec = osi.dwMinorVersion
    If dec <> 0 Then Do: dec = dec / 10: Loop Until dec < 1
    MajorMinor_ = osi.dwMajorVersion + dec
    
    ' Service Pack Major + Minor
    dec = osi.wServicePackMinor
    If dec <> 0 Then Do: dec = dec / 10: Loop Until dec < 1
    SPver_ = osi.wServicePackMajor + dec
    
    'getting ntdll.dll version
    hLib = GetProcAddress(GetModuleHandle(StrPtr("ntdll.dll")), "RtlGetNtVersionNumbers")
    
    If hLib <> 0 Then
        RtlGetNtVersionNumbers dwMajor, dwMinor, dwBuild
        
        dwBuild = dwBuild And &H3FFF&
        
        NtDllVersion_ = dwMajor & "." & dwMinor & "." & dwBuild
        dec = dwMinor
        If dec <> 0 Then Do: dec = dec / 10: Loop Until dec < 1
        MajorMinorNTDLL_ = dwMajor + dec
    End If
    
    IsWindowsXPOrGreater_ = (MajorMinor_ >= 5.1)
    IsWindowsVistaOrGreater_ = (MajorMinor_ >= 6)
    IsWindows7OrGreater_ = (MajorMinor_ >= 6.1)
    IsWindows8OrGreater_ = (MajorMinor_ >= 6.2)
    IsWindows8Point1OrGreater_ = (MajorMinor_ >= 6.3)
    IsWindows10OrGreater_ = (MajorMinor_ >= 6.4)
    
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
    
    PlatformID_ = osi.dwPlatformId
    
    If osi.dwPlatformId = Win32S Then
        OSName_ = "Windows 3.1"
        Platform_ = "Win3x"
        
    ElseIf osi.dwPlatformId = Win32Windows Then
        'Win 9x
        Platform_ = "Win9x"
        
        If osi.dwMajorVersion = 4 Then
            Select Case osi.dwMinorVersion
                '4.0
                Case 0 'Windows 95 [A/B/C]
                    OSName_ = "Windows 95"
                    If ChrW$(osi.szCSDVersion(0)) = "B" Or ChrW$(osi.szCSDVersion(0)) = "C" Then
                        Edition_ = "OSR2"
                    End If
                '4.10
                Case 10 'Windows 98 [Gold/SE]
                    OSName_ = "Windows 98"
                    If ChrW$(osi.szCSDVersion(0)) = "B" Or ChrW$(osi.szCSDVersion(0)) = "C" Then
                        Edition_ = "Second Edition"
                    Else
                        Edition_ = "Gold"
                    End If
                '4.90
                Case 90 'Windows Millennium Edition
                    OSName_ = "Windows ME"
            End Select
        End If
    ElseIf osi.dwPlatformId = Win32NT Then
        Platform_ = "WinNT"
        
        Select Case MajorMinor_
        Case 10
            If osi.wProductType = VER_NT_WORKSTATION Then
                OSName_ = "Windows 10"
            Else
                OSName_ = "Windows Server 2016"
            End If
        Case 6.4
            OSName_ = "Windows 10"
            Edition_ = "Technical Preview"
        Case 6.3
            If osi.wProductType = VER_NT_WORKSTATION Then
                OSName_ = "Windows 8.1"
            Else
                OSName_ = "Windows Server 2012 R2"
            End If
        Case 6.2
            If osi.wProductType = VER_NT_WORKSTATION Then
                OSName_ = "Windows 8"
            Else
                OSName_ = "Windows Server 2012"
            End If
        Case 6.1
            If osi.wProductType = VER_NT_WORKSTATION Then
                OSName_ = "Windows 7"
            Else
                OSName_ = "Windows Server 2008 R2"
            End If
        Case 6
            If osi.wProductType = VER_NT_WORKSTATION Then
                OSName_ = "Windows Vista"
            Else
                OSName_ = "Windows Server 2008"
            End If
        Case 5.2
            If osi.wProductType = VER_NT_SERVER Or osi.wProductType = VER_NT_DOMAIN_CONTROLLER Then
                OSName_ = "Windows Server 2003"
                If GetSystemMetrics(SM_SERVERR2) Then
                    Edition_ = "R2"
                ElseIf osi.wSuiteMask And VER_SUITE_DATACENTER Then
                    Edition_ = "Datacenter"
                ElseIf osi.wSuiteMask And VER_SUITE_ENTERPRISE Then
                    Edition_ = "Enterprise"
                ElseIf osi.wSuiteMask And VER_SUITE_BLADE Then
                    Edition_ = "Web Edition"
                ElseIf osi.wSuiteMask And VER_SUITE_STORAGE_SERVER Then
                    Edition_ = "Storage"
                ElseIf osi.wSuiteMask And VER_SUITE_WH_SERVER Then
                    Edition_ = "Home"
                Else
                    Edition_ = "Standard"
                End If
            ElseIf osi.wProductType = VER_NT_WORKSTATION And Bitness_ = "x64" Then
                OSName_ = "Windows XP"
                Edition_ = "Professional"
            End If
        Case 5.1
            OSName_ = "Windows XP"
            If GetSystemMetrics(SM_MEDIACENTER) Then
                Edition_ = "Media Center Edition"
            ElseIf GetSystemMetrics(SM_STARTER) Then
                Edition_ = "Starter Edition"
            ElseIf GetSystemMetrics(SM_TABLETPC) Then
                Edition_ = "Tablet PC Edition"
            ElseIf osi.wSuiteMask = VER_SUITE_PERSONAL Then
                Edition_ = "Home Edition"
            Else
                Edition_ = "Professional"
            End If
        Case 5
            OSName_ = "Windows 2000"
            If osi.wProductType = VER_NT_WORKSTATION Then
                If osi.wSuiteMask And VER_SUITE_PERSONAL Then
                    Edition_ = "Home"
                ElseIf GetSystemMetrics(SM_TABLETPC) = 0 Then
                    Edition_ = "Professional"
                Else
                    Edition_ = "Tablet Edition"
                End If
            Else
                If osi.wSuiteMask And VER_SUITE_DATACENTER Then
                    Edition_ = "Datacenter Server"
                ElseIf osi.wSuiteMask And VER_SUITE_ENTERPRISE Then
                    Edition_ = "Advanced Server"
                Else
                    Edition_ = "Server"
                End If
            End If
        Case 4
            OSName_ = "Windows NT 4.0"
            If osi.wProductType = VER_NT_WORKSTATION Then
                Edition_ = "Workstation"
            Else
                If osi.wSuiteMask And VER_SUITE_ENTERPRISE Then
                    Edition_ = "Enterprise Server"
                Else
                    Edition_ = "Standard Server"
                End If
            End If
        Case 3
            OSName_ = "Windows NT 3.51"
        End Select
    ElseIf osi.dwPlatformId = WinCE Then
        OSName_ = "Windows CE"
        Platform_ = "WinCE"
    ElseIf osi.dwPlatformId = UNIX Then
        OSName_ = "Unix"
        Platform_ = "Unix"
    End If
    
    If OSName_ = "" Then
        OSName_ = "Windows Unknown" & "(ver: " & MajorMinor_ & ". PT: " & osi.wProductType & ". Build: " & osi.dwBuildNumber & ")" & " Registry's data: " & _
            RegRead_REG_SZ("SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    End If
    
    If MajorMinor_ >= 10 Then
        'Get ReleaseId for Windows 10
        tmp = RegRead_REG_SZ("SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId")
        If IsNumeric(tmp) Then ReleaseId_ = CLng(tmp)
    End If
    
    'Редакция
    If Edition_ = "" Then
        If MajorMinor_ >= 6 Then
            If GetProductInfo(osi.dwMajorVersion, osi.dwMinorVersion, osi.wServicePackMajor, osi.wServicePackMinor, ProductType) Then
                Edition_ = GetProductName(ProductType)
            End If
        End If
    End If
    
    Revision_ = RegRead_REG_DWORD("SOFTWARE\Microsoft\Windows NT\CurrentVersion", "UBR")
    
    IsServer_ = (osi.wProductType <> VER_NT_WORKSTATION)
    
    If osi.wProductType = VER_NT_DOMAIN_CONTROLLER Then
        IsDomainController_ = True
        Edition_ = LTrim$(Edition_ & " (Domain Controller)")
    End If
    
    IsElevated_ = IsProcessElevated()
    
    IntegrityLevel_ = GetIntegrityLevel()
    
    UserType_ = GetUserType()
    UserName_ = GetUserName_()
    ComputerName_ = GetComputerName_()
    IsAdmin_ = (UserType_ = "Administrator")
    
    GetCodePageInfo
    
    ' wSuiteMask
    If osi.wSuiteMask And VER_SUITE_BACKOFFICE Then SuiteMask_ = SuiteMask_ & " + " & "BackOffice"
    If osi.wSuiteMask And VER_SUITE_BLADE Then SuiteMask_ = SuiteMask_ & " + " & "Blade"
    If osi.wSuiteMask And VER_SUITE_COMPUTE_SERVER Then SuiteMask_ = SuiteMask_ & " + " & "Compute Server"
    If osi.wSuiteMask And VER_SUITE_DATACENTER Then SuiteMask_ = SuiteMask_ & " + " & "DataCenter"
    If osi.wSuiteMask And VER_SUITE_ENTERPRISE Then SuiteMask_ = SuiteMask_ & " + " & "Enterprise"
    If osi.wSuiteMask And VER_SUITE_EMBEDDEDNT Then SuiteMask_ = SuiteMask_ & " + " & "EmbeddedNT"
    If osi.wSuiteMask And VER_SUITE_PERSONAL Then SuiteMask_ = SuiteMask_ & " + " & "Personal"
    If osi.wSuiteMask And VER_SUITE_SINGLEUSERTS Then SuiteMask_ = SuiteMask_ & " + " & "SingleUserTS"
    If osi.wSuiteMask And VER_SUITE_SMALLBUSINESS Then SuiteMask_ = SuiteMask_ & " + " & "SmallBusiness"
    If osi.wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED Then SuiteMask_ = SuiteMask_ & " + " & "SmallBusiness Restricted"
    If osi.wSuiteMask And VER_SUITE_STORAGE_SERVER Then SuiteMask_ = SuiteMask_ & " + " & "Storage Server"
    If osi.wSuiteMask And VER_SUITE_TERMINAL Then SuiteMask_ = SuiteMask_ & " + " & "Terminal"
    If osi.wSuiteMask And VER_SUITE_WH_SERVER Then SuiteMask_ = SuiteMask_ & " + " & "WH Server"
    
    If 0 <> Len(SuiteMask_) Then SuiteMask_ = Mid$(SuiteMask_, 4)
    
    ' wProductType
    Select Case osi.wProductType
        Case VER_NT_DOMAIN_CONTROLLER
            ProductType_ = "Domain Controller"
        Case VER_NT_SERVER
            ProductType_ = "Server"
        Case VER_NT_WORKSTATION
            ProductType_ = "Workstation"
    End Select
    
    SID_CurrentProcess_ = GetCurrentProcessSID()
    
    IsLocalSystemContext_ = (SID_CurrentProcess_ = "S-1-5-18")
    
    SecureBoot_ = SecureBootState()
    
    'http://www.nicklowe.org/2012/02/understanding-case-sensitivity-in-windows-obcaseinsensitive-file_case_sensitive_search/
    Dim hDir    As Long
    Dim OA      As OBJECT_ATTRIBUTES
    Dim UniStr  As UNICODE_STRING
    
    If Me.MajorMinor <= 5 Then  'Win2000-
        IsSystemCaseSensitive_ = True
    Else
        With UniStr
            .Length = LenB("\SYSTEMROOT")
            .MaximumLength = .Length
            .Buffer = StrPtr("\SYSTEMROOT")
        End With
    
        With OA
            .Length = LenB(OA)
            .ObjectName = VarPtr(UniStr)
            .Attributes = 0
        End With
        
        lret = NtOpenSymbolicLinkObject(hDir, DIRECTORY_QUERY, OA)
        
        If hDir <> 0 Then NtClose hDir
        
        If lret = STATUS_OBJECT_NAME_NOT_FOUND Then
            IsSystemCaseSensitive_ = True
        End If
    End If
    
    'For some reason, doesn't return correct result (at least, for Windows 8.1 Embedded Industry Pro)
    '
    'Const OS_EMBEDDED As Long = 13&
    'Dim iHasFeature As Long
    'iHasFeature = IsOS(OS_EMBEDDED)
    
    Dim sEditionID$, sProductName$
    sEditionID = RegRead_REG_SZ("SOFTWARE\Microsoft\Windows NT\CurrentVersion", "EditionID")
    sProductName = RegRead_REG_SZ("SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    
    IsEmbedded_ = (osi.wSuiteMask And VER_SUITE_EMBEDDEDNT) Or _
        (InStr(1, sEditionID, "Embedded", 1) <> 0) Or _
        (InStr(1, sProductName, "Embedded", 1) <> 0)
    
    If InStr(1, Edition_, "Unknown", 1) <> 0 Then Edition_ = Edition_ & " (" & sEditionID & ")"
    If IsEmbedded_ Then Edition_ = Edition_ & " (Embedded)"
    
    AppendErrorLogCustom "clsOSver.Class_Initialize - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsOSver.Class_Initialize"
    Resume Next
End Sub

Private Sub GetCodePageInfo()
    On Error GoTo ErrorHandler

    AppendErrorLogCustom "GetCodePageInfo - Begin"

    CodepageOEM_ = RegRead_REG_SZ("SYSTEM\CurrentControlSet\Control\Nls\CodePage", "OEMCP")
 
    If 0 <> Len(CodepageOEM_) Then
        CodepageOEM_File_ = RegRead_REG_SZ("SYSTEM\CurrentControlSet\Control\Nls\CodePage", CodepageOEM_)
    End If
    
    CodepageANSI_ = RegRead_REG_SZ("SYSTEM\CurrentControlSet\Control\Nls\CodePage", "ACP")
 
    If 0 <> Len(CodepageANSI_) Then
        CodepageANSI_File_ = RegRead_REG_SZ("SYSTEM\CurrentControlSet\Control\Nls\CodePage", CodepageANSI_)
    End If
    
    AppendErrorLogCustom "GetCodePageInfo - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsOSver.GetCodePageInfo"
End Sub

Private Function GetLangNameFullByCultureCode(lcid As Long) As String
    Dim buf As String
    Dim lr  As Long
    buf = String$(1000, vbNullChar)
    lr = GetLocaleInfo(lcid, LOCALE_SENGLANGUAGE, StrPtr(buf), ByVal 1000&)
    If lr Then
        GetLangNameFullByCultureCode = Left$(buf, lr - 1)
    End If
End Function

Private Function RegRead_REG_SZ(strKey As String, strParam As String) As String
    On Error GoTo ErrorHandler
    
    Const KEY_QUERY_VALUE       As Long = &H1&
    
    Dim buf As String
    Dim hKey As Long
    Dim ordType As Long
    Dim cData As Long

    RegOpenKeyEx HKEY_LOCAL_MACHINE, StrPtr(strKey), 0&, KEY_QUERY_VALUE Or (IsWin64_ And KEY_WOW64_64KEY), hKey
    If hKey = 0 Then Exit Function
    
    RegQueryValueExLong hKey, StrPtr(strParam), 0&, ordType, 0&, cData
    If cData > 1 Then
        buf = String$(cData - 1&, vbNullChar)
        RegQueryValueExStr hKey, StrPtr(strParam), 0&, ordType, StrPtr(buf), cData
    End If
    buf = Left$(buf, lstrlenW(StrPtr(buf)))
    
    RegRead_REG_SZ = buf
    If hKey <> 0 Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsOSver.RegRead_REG_SZ"
    If hKey <> 0 Then RegCloseKey hKey
End Function

Private Function RegRead_REG_DWORD(strKey As String, strParam As String) As Long
    On Error GoTo ErrorHandler
    
    Const KEY_QUERY_VALUE       As Long = &H1&
    
    Dim lData As Long
    Dim hKey As Long
    Dim ordType As Long
    Dim cData As Long

    RegOpenKeyEx HKEY_LOCAL_MACHINE, StrPtr(strKey), 0&, KEY_QUERY_VALUE Or (IsWin64_ And KEY_WOW64_64KEY), hKey
    If hKey = 0 Then Exit Function
    
    RegQueryValueExLong hKey, StrPtr(strParam), 0&, ordType, 0&, cData
    If cData = 4 Then
        If ERROR_SUCCESS = RegQueryValueEx(hKey, StrPtr(strParam), 0&, ordType, VarPtr(lData), cData) Then
            RegRead_REG_DWORD = lData
        End If
    End If
    
    If hKey <> 0 Then RegCloseKey hKey
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsOSver.RegRead_REG_SZ"
    If hKey <> 0 Then RegCloseKey hKey
End Function

Public Function IsProcessElevated(Optional hProcess As Long) As Boolean
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "IsProcessElevated - Begin"
    
    Dim hToken           As Long
    Dim dwLengthNeeded   As Long
    Dim dwIsElevated     As Long
    
    ' < Win Vista. Устанавливаем true, если пользователь состоит в группе "Администраторы"
    If osi.dwMajorVersion < 6 Then IsProcessElevated = (GetUserType() = "Administrator"): Exit Function

    If hProcess = 0 Then hProcess = GetCurrentProcess()
    
    If OpenProcessToken(hProcess, TOKEN_QUERY, hToken) Then

        If 0 <> GetTokenInformation(hToken, TokenElevation, dwIsElevated, 4&, dwLengthNeeded) Then
            IsProcessElevated = (dwIsElevated <> 0)
        End If
        
        CloseHandle hToken: hToken = 0
    End If
    
    AppendErrorLogCustom "IsProcessElevated - End"
    Exit Function
ErrorHandler:
    ErrorMsg Now, "clsOSver.IsProcessElevated"
    If hToken Then CloseHandle hToken: hToken = 0
End Function

Private Function GetUserType() As String
    On Error GoTo ErrorHandler

    AppendErrorLogCustom "GetUserType - Begin"

    Const TOKEN_QUERY                   As Long = &H8&
    Const SECURITY_NT_AUTHORITY         As Long = 5&
    Const TokenGroups                   As Long = 2&
    Const SECURITY_BUILTIN_DOMAIN_RID   As Long = &H20&
    Const DOMAIN_ALIAS_RID_ADMINS       As Long = &H220&
    Const DOMAIN_ALIAS_RID_USERS        As Long = &H221&
    Const DOMAIN_ALIAS_RID_GUESTS       As Long = &H222&
    Const DOMAIN_ALIAS_RID_POWER_USERS  As Long = &H223&

    Dim hProcessToken   As Long
    Dim BufferSize      As Long
    Dim psidAdmin       As Long
    Dim psidPower       As Long
    Dim psidUser        As Long
    Dim psidGuest       As Long
    Dim i               As Long
    Dim lResult         As Long
    Dim tpTokens        As TOKEN_GROUPS
    Dim tpSidAuth       As SID_IDENTIFIER_AUTHORITY
    
    GetUserType = "Unknown"
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
    
    ' Obtain current process token
    If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) Then
        Call OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken)
    End If
    
    If hProcessToken Then
        
        ' Определяем требуемый размер буфера
        GetTokenInformation hProcessToken, ByVal TokenGroups, 0&, 0&, BufferSize
        
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long  ' Переводим размер byte -> Long
            
            ' Получаем информацию о SID-ах групп, ассоциированных с этим токеном
            If 0 <> GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize) Then
            
                ' Заполняем структуру из буфера
                Call CopyMemory(tpTokens, InfoBuffer(0), Len(tpTokens))
            
                ' Получаем SID-ы каждого типа пользователей
                lResult = AllocateAndInitializeSid(tpSidAuth, 2&, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0&, 0&, 0&, 0&, 0&, 0&, psidAdmin)
                lResult = AllocateAndInitializeSid(tpSidAuth, 2&, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_POWER_USERS, 0&, 0&, 0&, 0&, 0&, 0&, psidPower)
                lResult = AllocateAndInitializeSid(tpSidAuth, 2&, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_USERS, 0&, 0&, 0&, 0&, 0&, 0&, psidUser)
                lResult = AllocateAndInitializeSid(tpSidAuth, 2&, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_GUESTS, 0&, 0&, 0&, 0&, 0&, 0&, psidGuest)
            
                If IsValidSid(psidAdmin) And IsValidSid(psidPower) And IsValidSid(psidUser) And IsValidSid(psidGuest) Then
                  
                    For i = 0 To tpTokens.GroupCount
                        ' Берем SID каждой из ассоциированных групп
                        If IsValidSid(tpTokens.Groups(i).SID) Then
                            ' Проверяем на соответствие
                            If EqualSid(ByVal tpTokens.Groups(i).SID, ByVal psidAdmin) Then
                                GetUserType = "Administrator":  Exit For
                            ElseIf EqualSid(ByVal tpTokens.Groups(i).SID, ByVal psidPower) Then
                                GetUserType = "Power User":     Exit For
                            ElseIf EqualSid(ByVal tpTokens.Groups(i).SID, ByVal psidUser) Then
                                GetUserType = "Limited User":   Exit For
                            ElseIf EqualSid(ByVal tpTokens.Groups(i).SID, ByVal psidGuest) Then
                                GetUserType = "Guest":          Exit For
                            End If
                        End If
                    Next
                End If
                If psidAdmin Then FreeSid psidAdmin: psidAdmin = 0
                If psidPower Then FreeSid psidPower: psidPower = 0
                If psidUser Then FreeSid psidUser: psidUser = 0
                If psidGuest Then FreeSid psidGuest: psidGuest = 0
            End If
        End If
        CloseHandle hProcessToken: hProcessToken = 0
    End If
    
    AppendErrorLogCustom "GetUserType - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsOSver.GetUserType"
    If hProcessToken Then CloseHandle hProcessToken: hProcessToken = 0
    If psidAdmin Then FreeSid psidAdmin
    If psidPower Then FreeSid psidPower
    If psidUser Then FreeSid psidUser
    If psidGuest Then FreeSid psidGuest
End Function

Private Function GetCurrentProcessSID() As String
    On Error GoTo ErrorHandler

    AppendErrorLogCustom "GetCurrentProcessSID - Begin"

    Const TokenUser As Long = 1&
        
    Dim hProcess As Long, hProcessToken As Long, tkUser As SID_AND_ATTRIBUTES, dwSize As Long, sSID As String, psSID As Long

    If hProcess = 0 Then hProcess = GetCurrentProcess()
    If 0 = OpenProcessToken(hProcess, TOKEN_QUERY, hProcessToken) Then Exit Function
    
    If hProcessToken Then
        
        ' Определяем требуемый размер буфера
        GetTokenInformation hProcessToken, ByVal TokenUser, 0&, 0&, dwSize
        
        If dwSize Then
            
            ReDim InfoBuffer((dwSize \ 4) - 1) As Long
            
            ' Получаем информацию о SID-ах групп, ассоциированных с этим токеном
            If 0 <> GetTokenInformation(hProcessToken, ByVal TokenUser, InfoBuffer(0), dwSize, dwSize) Then
            
                Call CopyMemory(tkUser, InfoBuffer(0), Len(tkUser))
            
                ' Заполняем структуру из буфера
                'Call CopyMemory(tpTokens, InfoBuffer(0), Len(tpTokens))
                
                If IsValidSid(tkUser.SID) Then
                    If 0 <> ConvertSidToStringSid(tkUser.SID, psSID) Then
                        dwSize = lstrlenW(psSID)
                        sSID = String$(dwSize + 1, 0)   '4 Nul
                        If dwSize <> 0 Then
                            lstrcpyW StrPtr(sSID), psSID
                        End If
                        GetCurrentProcessSID = Left$(sSID, lstrlenW(StrPtr(sSID)))
                        LocalFree psSID
                    End If
                End If
            End If
        End If
        CloseHandle hProcessToken
    End If
    
    AppendErrorLogCustom "GetCurrentProcessSID - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsOSver.GetCurrentProcessSID"
End Function

Private Function GetIntegrityLevel(Optional hProcess As Long) As String
    'https://msdn.microsoft.com/en-us/library/bb625963.aspx
    'https://msdn.microsoft.com/en-us/library/bb625966.aspx
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "GetIntegrityLevel - Begin"
    
    Const SECURITY_MANDATORY_UNTRUSTED_RID          As Long = 0&
    Const SECURITY_MANDATORY_LOW_RID                As Long = &H1000&
    Const SECURITY_MANDATORY_MEDIUM_RID             As Long = &H2000&
    Const SECURITY_MANDATORY_UIACCESS_RID           As Long = SECURITY_MANDATORY_MEDIUM_RID + &H10&
    Const SECURITY_MANDATORY_HIGH_RID               As Long = &H3000&
    Const SECURITY_MANDATORY_SYSTEM_RID             As Long = &H4000&
    Const SECURITY_MANDATORY_PROTECTED_PROCESS_RID  As Long = &H5000&
    
    Const TokenIntegrityLevel       As Long = 25&
    Const TOKEN_QUERY               As Long = &H8&
    Const ERROR_INSUFFICIENT_BUFFER As Long = &H7A&
    
    Dim hToken           As Long
    Dim dwLengthNeeded   As Long
    Dim bTIL()           As Byte
    Dim pSidSub          As Long
    Dim dwIntegrityLevel As Long
    Dim pSidAuthCnt      As Long
    Dim SidAuthCnt       As Long
    Dim pILSid           As Long
    Dim iLevel           As String
    
    If osi.dwMajorVersion < 6 Then GetIntegrityLevel = "Not supported": Exit Function ' < Win Vista
    
    iLevel = "Unknown"
    
    If hProcess = 0 Then hProcess = GetCurrentProcess()
    
    If OpenProcessToken(hProcess, TOKEN_QUERY, hToken) Then
    
        GetTokenInformation hToken, TokenIntegrityLevel, 0&, 0&, dwLengthNeeded
        
        If ERROR_INSUFFICIENT_BUFFER = Err.LastDllError Then
        
            ReDim bTIL(dwLengthNeeded - 1)
        
            If 0 <> GetTokenInformation(hToken, TokenIntegrityLevel, bTIL(0), dwLengthNeeded, dwLengthNeeded) Then
        
                GetMem4 bTIL(0), pILSid
                
                If IsValidSid(pILSid) Then

                    pSidAuthCnt = GetSidSubAuthorityCount(pILSid)
                    
                    If pSidAuthCnt Then
                    
                        GetMem4 ByVal pSidAuthCnt, SidAuthCnt
                        
                        If SidAuthCnt Then
                        
                            pSidSub = GetSidSubAuthority(pILSid, SidAuthCnt - 1)
                    
                            If pSidSub Then GetMem4 ByVal pSidSub, dwIntegrityLevel
                    
                            Select Case dwIntegrityLevel
                            
                                Case Is < SECURITY_MANDATORY_UNTRUSTED_RID
                                    iLevel = "Unknown"
                                Case Is < SECURITY_MANDATORY_LOW_RID
                                    iLevel = "Untrusted"
                                Case Is < SECURITY_MANDATORY_MEDIUM_RID
                                    iLevel = "Low"
                                Case SECURITY_MANDATORY_UIACCESS_RID
                                    iLevel = "Medium (UIAccess)"
                                Case Is < SECURITY_MANDATORY_HIGH_RID
                                    iLevel = "Medium"
                                Case Is < SECURITY_MANDATORY_SYSTEM_RID
                                    iLevel = "High"
                                Case Is < SECURITY_MANDATORY_PROTECTED_PROCESS_RID
                                    iLevel = "System"
                                Case Else
                                    iLevel = "ProtectedProcess"
                            End Select
                        End If
                    End If
                    'FreeSid pILSid
                End If
            End If
        End If
        CloseHandle hToken: hToken = 0
    End If
    GetIntegrityLevel = iLevel
    
    AppendErrorLogCustom "GetIntegrityLevel - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsOSver.GetIntegrityLevel"
    'If pILSid Then FreeSid pILSid
    If hToken Then CloseHandle hToken: hToken = 0
End Function

'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-mde2/d92ead8f-faf3-47a8-a341-1921dc2c463b
'https://docs.microsoft.com/en-us/windows/win32/api/sysinfoapi/nf-sysinfoapi-getproductinfo
Private Function GetProductName(ProductType As Long) As String
    On Error GoTo ErrorHandler
    
    Dim ProductName As String
    
    Select Case ProductType
    Case &H1&:        ProductName = "Ultimate"
    Case &H2&:        ProductName = "Home Basic"
    Case &H3&:        ProductName = "Home Premium"
    Case &H4&:        ProductName = "Enterprise"
    Case &H5&:        ProductName = "Home Basic N"
    Case &H6&:        ProductName = "Business"
    Case &H7&:        ProductName = "Server Standard"
    Case &H8&:        ProductName = "Server Datacenter"
    Case &H9&:        ProductName = "Small Business Server"
    Case &HA&:        ProductName = "Server Enterprise"
    Case &HB&:        ProductName = "Starter"
    Case &HC&:        ProductName = "Server Datacenter Core"
    Case &HD&:        ProductName = "Server Standard Core"
    Case &HE&:        ProductName = "Server Enterprise Core"
    Case &HF&:        ProductName = "Server Enterprise for Itanium-based Systems"
    Case &H10&:        ProductName = "Business N"
    Case &H11&:        ProductName = "Web Server"
    Case &H12&:        ProductName = "Cluster Server" 'HPC Edition?
    Case &H13&:        ProductName = "Home Server" 'Storage Server 2008 R2 Essentials?
    Case &H14&:        ProductName = "Storage Server Express"
    Case &H15&:        ProductName = "Storage Server Standard"
    Case &H16&:        ProductName = "Storage Server Workgroup"
    Case &H17&:        ProductName = "Storage Server Enterprise"
    Case &H18&:        ProductName = "Server for Small Business" 'Server 2008 for Windows Essential Server Solutions?
    Case &H19&:        ProductName = "Small Business Server Premium"
    Case &H1A&:        ProductName = "Home Premium N"
    Case &H1B&:        ProductName = "Enterprise N"
    Case &H1C&:        ProductName = "Ultimate N"
    Case &H1D&:        ProductName = "Web Server Core"
    Case &H1E&:        ProductName = "Medium Business Server Management 'Essential Business Server Management Server?"
    Case &H1F&:        ProductName = "Medium Business Server Security" 'Essential Business Server Security Server
    Case &H20&:        ProductName = "Medium Business Server Messaging" 'Essential Business Server Messaging Server
    Case &H21&:        ProductName = "Server Foundation"
    Case &H22&:        ProductName = "Home Premium Server" 'Home Server 2011?
    Case &H23&:        ProductName = "Server for Small Business without Hyper-V" 'Server 2008 without Hyper-V for Windows Essential Server Solutions?
    Case &H24&:        ProductName = "Server Standard without Hyper-V"
    Case &H25&:        ProductName = "Server Datacenter without Hyper-V"
    Case &H26&:        ProductName = "Server Enterprise without Hyper-V"
    Case &H27&:        ProductName = "Server Datacenter without Hyper-V Core"
    Case &H28&:        ProductName = "Server Standard without Hyper-V Core"
    Case &H29&:        ProductName = "Server Enterprise without Hyper-V Core"
    Case &H2A&:        ProductName = "Microsoft Hyper-V Server"
    Case &H2B&:        ProductName = "Storage Server Express Core"
    Case &H2C&:        ProductName = "Storage Server Standard Core"
    Case &H2D&:        ProductName = "Storage Server Workgroup Core"
    Case &H2E&:        ProductName = "Storage Server Enterprise Core"
    Case &H2F&:        ProductName = "Starter N"
    Case &H30&:        ProductName = "Pro"
    Case &H31&:        ProductName = "Pro N"
    Case &H32&:        ProductName = "Small Business Solution Server" 'Small Business Server 2011 Essentials?
    Case &H33&:        ProductName = "Small Business Solution Server"
    Case &H34&:        ProductName = "Server Solutions Standard" 'Server Solutions Premium
    Case &H35&:        ProductName = "Server Solutions Standard Core" 'Server Solutions Premium Core
    Case &H36&:        ProductName = "Server For Small Business Solutions EM"
    Case &H37&:        ProductName = "Server For Small Business Solutions EM"
    Case &H38&:        ProductName = "Solution Embedded Server" ' MultiPoint Server?
    Case &H39&:        ProductName = "Solution Embedded Server Core"
    Case &H3B&:        ProductName = "Essential Server Solution Management"
    Case &H3C&:        ProductName = "Essential Server Solution Additional"
    Case &H3D&:        ProductName = "Essential Server Solution Management SVC"
    Case &H3E&:        ProductName = "Essential Server Solution Additional SVC"
    Case &H3F&:        ProductName = "Small Business Server Premium Core"
    Case &H40&:        ProductName = "Cluster Server V" 'Server Hyper Core V?
    Case &H41&:        ProductName = "Embedded"
    Case &H42&:        ProductName = "Starter Enterprise"
    Case &H43&:        ProductName = "Home Basic Enterprise"
    Case &H44&:        ProductName = "Home Premium Enterprise"
    Case &H45&:        ProductName = "Pro Enterprise"
    Case &H46&:        ProductName = "Enterprise E"
    Case &H47&:        ProductName = "Ultimate Enterprise"
    Case &H48&:        ProductName = "Enterprise Evaluation" '"Server Enterprise (EI)"
    Case &H4A&:        ProductName = "Consumer Preview"
    Case &H4C&:        ProductName = "MultiPoint Server Standard"
    Case &H4D&:        ProductName = "MultiPoint Server Premium"
    Case &H4F&:        ProductName = "Server Standard Evaluation"
    Case &H50&:        ProductName = "Server Datacenter Evaluation"
    Case &H54&:        ProductName = "Enterprise N Evaluation"
    Case &H55&:        ProductName = "Embedded Automotive"
    Case &H56&:        ProductName = "Embedded Industry A"
    Case &H57&:        ProductName = "Thin PC"
    Case &H58&:        ProductName = "Embedded A"
    Case &H59&:        ProductName = "Embedded Industry" 'Industry Pro?
    Case &H5A&:        ProductName = "Embedded E"
    Case &H5B&:        ProductName = "Embedded Industry E"
    Case &H5C&:        ProductName = "Embedded Industry A E"
    Case &H5F&:        ProductName = "Storage Server Workgroup Evaluation"
    Case &H60&:        ProductName = "Storage Server Standard Evaluation"
    Case &H61&:        ProductName = "Core ARM"
    Case &H62&:        ProductName = "Home N" ' "Windows 8 N" ' Core N
    Case &H63&:        ProductName = "Home China" ' "Windows 8 China" 'Core Country Specific
    Case &H64&:        ProductName = "Home Single Language" ' "Windows 8 Single Language"
    Case &H65&:        ProductName = "Home" '"Windows 8 Core"
    Case &H67&:        ProductName = "Professional with Media Center"
    Case &H68&:        ProductName = "Mobile Core"
    Case &H69&:        ProductName = "Embedded Industry Eval"
    Case &H6A&:        ProductName = "Embedded Industry Eval E"
    Case &H6B&:        ProductName = "Embedded Eval"
    Case &H6C&:        ProductName = "Embedded Eval E"
    Case &H6D&:        ProductName = "Nano Server"
    Case &H6E&:        ProductName = "Cloud Storage Server"
    Case &H6F&:        ProductName = "Core Connected"
    Case &H70&:        ProductName = "Professional Student"
    Case &H71&:        ProductName = "Core Connected N"
    Case &H72&:        ProductName = "Professional Student N"
    Case &H73&:        ProductName = "Core Connected Single Language"
    Case &H74&:        ProductName = "Core Connected China"
    Case &H75&:        ProductName = "Connected Car"
    Case &H76&:        ProductName = "Industry Handheld"
    Case &H77&:        ProductName = "PPI Pro"
    Case &H78&:        ProductName = "Server ARM64"
    Case &H79&:        ProductName = "Education"
    Case &H7A&:        ProductName = "Education N"
    Case &H7B&:        ProductName = "IoT Core"
    Case &H7C&:        ProductName = "Cloud Host Infrastructure Server"
    Case &H7D&:        ProductName = "Enterprise LTSB" 'Enterprise 2015 LTSB?
    Case &H7E&:        ProductName = "Enterprise LTSB N" 'Enterprise 2015 LTSB N?
    Case &H7F&:        ProductName = "Pro LTSB"
    Case &H80&:        ProductName = "Pro LTSB N"
    Case &H81&:        ProductName = "Enterprise LTSB Evaluation" 'Enterprise 2015 LTSB Evaluation?
    Case &H82&:        ProductName = "Enterprise LTSB N Evaluation" 'Enterprise 2015 LTSB N Evaluation?
    Case &H83&:        ProductName = "IoT Core Commercial"
    Case &H85&:        ProductName = "Mobile Enterprise"
    Case &H87&:        ProductName = "Holographic"
    Case &H88&:        ProductName = "Holographic Business"
    Case &H91&:        ProductName = "Server Datacenter (Semi-Annual Channel Core)"
    Case &H92&:        ProductName = "Server Standard (Semi-Annual Channel Core)"
    Case &HA1&:        ProductName = "Pro for Workstations"
    Case &HA2&:        ProductName = "Pro for Workstations N"
    Case 0:
        ProductName = "An unknown product"
    Case Else
        ProductName = "Unknown Edition (0x" & Hex$(ProductType) & ")"
    End Select
    
    GetProductName = ProductName
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsOSver.GetProductName"
End Function


Private Function GetLangNameByCultureCode(CultureCode As Long) As String
    Dim Lang$
    Select Case CultureCode
        Case &H419&
            'Lang = "ru-RU"
            Lang = "RU"
        Case &H409&
            'Lang = "en-US"
            Lang = "EN"
        Case &H422&
            'Lang = "uk-UA"
            Lang = "UA"
        Case &H423&
            Lang = "be-BY"
        Case &H402&
            'Lang = "bg-BG"
            Lang = "BG"
        Case &H436&
            Lang = "af-ZA"
        Case &H41C&
            Lang = "sq-AL"
        Case &H1401&
            Lang = "ar-DZ"
        Case &H3C01&
            Lang = "ar-BH"
        Case &HC01&
            Lang = "ar-EG"
        Case &H801&
            Lang = "ar-IQ"
        Case &H2C01&
            Lang = "ar-JO"
        Case &H3401&
            Lang = "ar-KW"
        Case &H3001&
            Lang = "ar-LB"
        Case &H1001&
            Lang = "ar-LY"
        Case &H1801&
            Lang = "ar-MA"
        Case &H2001&
            Lang = "ar-OM"
        Case &H4001&
            Lang = "ar-QA"
        Case &H401&
            Lang = "ar-SA"
        Case &H2801&
            Lang = "ar-SY"
        Case &H1C01&
            Lang = "ar-TN"
        Case &H3801&
            Lang = "ar-AE"
        Case &H2401&
            Lang = "ar-YE"
        Case &H42B&
            Lang = "hy-AM"
        Case &H82C&
            Lang = "Cy-az-AZ"
        Case &H42C&
            Lang = "Lt-az-AZ"
        Case &H42D&
            Lang = "eu-ES"
        Case &H403&
            Lang = "ca-ES"
        Case &H804&
            Lang = "zh-CN"
        Case &HC04&
            Lang = "zh-HK"
        Case &H1404&
            Lang = "zh-MO"
        Case &H1004&
            Lang = "zh-SG"
        Case &H404&
            Lang = "zh-TW"
        Case &H4&
            Lang = "zh-CHS"
        Case &H7C04&
            Lang = "zh-CHT"
        Case &H41A&
            Lang = "hr-HR"
        Case &H405&
            Lang = "cs-CZ"
        Case &H406&
            Lang = "da-DK"
        Case &H465&
            Lang = "div-MV"
        Case &H813&
            Lang = "nl-BE"
        Case &H413&
            Lang = "nl-NL"
        Case &HC09&
            Lang = "en-AU"
        Case &H2809&
            Lang = "en-BZ"
        Case &H1009&
            Lang = "en-CA"
        Case &H2409&
            Lang = "en-CB"
        Case &H1809&
            Lang = "en-IE"
        Case &H2009&
            Lang = "en-JM"
        Case &H1409&
            Lang = "en-NZ"
        Case &H3409&
            Lang = "en-PH"
        Case &H1C09&
            Lang = "en-ZA"
        Case &H2C09&
            Lang = "en-TT"
        Case &H809&
            Lang = "en-GB"
        Case &H3009&
            Lang = "en-ZW"
        Case &H425&
            Lang = "et-EE"
        Case &H438&
            Lang = "fo-FO"
        Case &H429&
            Lang = "fa-IR"
        Case &H40B&
            Lang = "fi-FI"
        Case &H80C&
            Lang = "fr-BE"
        Case &HC0C&
            Lang = "fr-CA"
        Case &H40C&
            Lang = "fr-FR"
        Case &H140C&
            Lang = "fr-LU"
        Case &H180C&
            Lang = "fr-MC"
        Case &H100C&
            Lang = "fr-CH"
        Case &H456&
            Lang = "gl-ES"
        Case &H437&
            Lang = "ka-GE"
        Case &HC07&
            Lang = "de-AT"
        Case &H407&
            Lang = "de-DE"
        Case &H1407&
            Lang = "de-LI"
        Case &H1007&
            Lang = "de-LU"
        Case &H807&
            Lang = "de-CH"
        Case &H408&
            Lang = "el-GR"
        Case &H447&
            Lang = "gu-IN"
        Case &H40D&
            Lang = "he-IL"
        Case &H439&
            Lang = "hi-IN"
        Case &H40E&
            Lang = "hu-HU"
        Case &H40F&
            Lang = "is-IS"
        Case &H421&
            Lang = "id-ID"
        Case &H410&
            Lang = "it-IT"
        Case &H810&
            Lang = "it-CH"
        Case &H411&
            Lang = "ja-JP"
        Case &H44B&
            Lang = "kn-IN"
        Case &H43F&
            Lang = "kk-KZ"
        Case &H457&
            Lang = "kok-IN"
        Case &H412&
            Lang = "ko-KR"
        Case &H440&
            Lang = "ky-KZ"
        Case &H426&
            Lang = "lv-LV"
        Case &H427&
            Lang = "lt-LT"
        Case &H42F&
            Lang = "mk-MK"
        Case &H83E&
            Lang = "ms-BN"
        Case &H43E&
            Lang = "ms-MY"
        Case &H44E&
            Lang = "mr-IN"
        Case &H450&
            Lang = "mn-MN"
        Case &H414&
            Lang = "nb-NO"
        Case &H814&
            Lang = "nn-NO"
        Case &H415&
            Lang = "pl-PL"
        Case &H416&
            Lang = "pt-BR"
        Case &H816&
            Lang = "pt-PT"
        Case &H446&
            Lang = "pa-IN"
        Case &H418&
            Lang = "ro-RO"
        Case &H44F&
            Lang = "sa-IN"
        Case &HC1A&
            Lang = "Cy-sr-SP"
        Case &H81A&
            Lang = "Lt-sr-SP"
        Case &H41B&
            Lang = "sk-SK"
        Case &H424&
            Lang = "sl-SI"
        Case &H2C0A&
            Lang = "es-AR"
        Case &H400A&
            Lang = "es-BO"
        Case &H340A&
            Lang = "es-CL"
        Case &H240A&
            Lang = "es-CO"
        Case &H140A&
            Lang = "es-CR"
        Case &H1C0A&
            Lang = "es-DO"
        Case &H300A&
            Lang = "es-EC"
        Case &H440A&
            Lang = "es-SV"
        Case &H100A&
            Lang = "es-GT"
        Case &H480A&
            Lang = "es-HN"
        Case &H80A&
            Lang = "es-MX"
        Case &H4C0A&
            Lang = "es-NI"
        Case &H180A&
            Lang = "es-PA"
        Case &H3C0A&
            Lang = "es-PY"
        Case &H280A&
            Lang = "es-PE"
        Case &H500A&
            Lang = "es-PR"
        Case &HC0A&
            Lang = "es-ES"
        Case &H380A&
            Lang = "es-UY"
        Case &H200A&
            Lang = "es-VE"
        Case &H441&
            Lang = "sw-KE"
        Case &H81D&
            Lang = "sv-FI"
        Case &H41D&
            Lang = "sv-SE"
        Case &H45A&
            Lang = "syr-SY"
        Case &H449&
            Lang = "ta-IN"
        Case &H444&
            Lang = "tt-RU"
        Case &H44A&
            Lang = "te-IN"
        Case &H41E&
            Lang = "th-TH"
        Case &H41F&
            Lang = "tr-TR"
        Case &H420&
            Lang = "ur-PK"
        Case &H843&
            Lang = "Cy-uz-UZ"
        Case &H443&
            Lang = "Lt-uz-UZ"
        Case &H42A&
            Lang = "vi-VN"
        Case Else
            Lang = "unknown"
    End Select
    GetLangNameByCultureCode = Lang
End Function

Public Function IsWow64() As Boolean   ' Разрядность ОС
    Dim hModule As Long, procAddr As Long, lIsWin64 As Long
    Static isInit As Boolean, result As Boolean
    
    If isInit Then
        IsWow64 = result
    Else
        isInit = True
        hModule = LoadLibrary(StrPtr("kernel32.dll"))
        If hModule Then
            procAddr = GetProcAddress(hModule, "IsWow64Process")
            If procAddr <> 0 Then
                IsWow64Process GetCurrentProcess(), lIsWin64
                result = CBool(lIsWin64)
                IsWow64 = result
            End If
            FreeLibrary hModule
        End If
    End If
End Function

Private Function SecureBootState() As Boolean
    Dim lState As Long
    '// TODO:
    'improvement - https://www.cyberforum.ru/win-api/thread2046367.html#post10789948
    '++
    If MajorMinor_ >= 6.2 Then ' Win8+
        lState = RegRead_REG_DWORD("SYSTEM\CurrentControlSet\Control\SecureBoot\State", "UEFISecureBootEnabled")
        If lState = 1 Then SecureBootState = True
    End If
End Function

Private Function GetUserName_() As String
    Dim sUsername$
    sUsername = String$(MAX_PATH, vbNullChar)
    If 0 <> GetUserName(StrPtr(sUsername), MAX_PATH) Then
        sUsername = Left$(sUsername, lstrlen(StrPtr(sUsername)))
    End If
    GetUserName_ = sUsername
End Function

Private Function GetComputerName_() As String
    Dim sComputerName$
    sComputerName = String$(MAX_PATH, vbNullChar)
    If 0 <> GetComputerName(StrPtr(sComputerName), MAX_PATH) Then
        sComputerName = Left$(sComputerName, lstrlen(StrPtr(sComputerName)))
    End If
    GetComputerName_ = sComputerName
End Function

Public Property Get Bitness() As String
    Bitness = Bitness_
End Property

Public Property Get Major() As Long
    Major = osi.dwMajorVersion
End Property

Public Property Get Minor() As Long
    Minor = osi.dwMinorVersion
End Property

Public Property Get MajorMinor() As Single
    MajorMinor = MajorMinor_
End Property

Public Property Get MajorMinorNTDLL() As Single
    MajorMinorNTDLL = MajorMinorNTDLL_
End Property

Public Property Get NtDllVersion() As String
    NtDllVersion = NtDllVersion_
End Property

Public Property Get Build() As Long
    Build = osi.dwBuildNumber
End Property

Public Property Get Revision() As Long
    Revision = Revision_
End Property

Public Property Get SPVer() As Single
    SPVer = SPver_
End Property

Public Property Get OSName() As String
    OSName = OSName_
End Property

Public Property Get Edition() As String
    Edition = Edition_
End Property

Public Property Get IsElevated() As Boolean
    IsElevated = IsElevated_
End Property

Public Property Get IntegrityLevel() As String
    IntegrityLevel = IntegrityLevel_
End Property

Public Property Get UserType() As String
    UserType = UserType_
End Property

Public Property Get UserName() As String
    UserName = UserName_
End Property

Public Property Get ComputerName() As String
    ComputerName = ComputerName_
End Property

Public Property Get IsSafeBoot() As Boolean
    IsSafeBoot = IsSafeBoot_
End Property

Public Property Get LangSystemCode() As Long
    LangSystemCode = LangSystemCode_
End Property

Public Property Get LangSystemName() As String
    LangSystemName = LangSystemName_
End Property

Public Property Get LangSystemNameFull() As String
    LangSystemNameFull = LangSystemNameFull_
End Property

Public Property Get LangNonUnicodeCode() As Long
    LangNonUnicodeCode = LangNonUnicodeCode_
End Property

Public Property Get LangNonUnicodeName() As String
    LangNonUnicodeName = LangNonUnicodeName_
End Property

Public Property Get LangNonUnicodeNameFull() As String
    LangNonUnicodeNameFull = LangNonUnicodeNameFull_
End Property

Public Property Get LangDisplayCode() As Long
    LangDisplayCode = LangDisplayCode_
End Property

Public Property Get LangDisplayName() As String
    LangDisplayName = LangDisplayName_
End Property

Public Property Get LangDisplayNameFull() As String
    LangDisplayNameFull = LangDisplayNameFull_
End Property

Public Property Get IsWindowsXPOrGreater() As Boolean
    IsWindowsXPOrGreater = IsWindowsXPOrGreater_
End Property

Public Property Get IsWindowsVistaOrGreater() As Boolean
    IsWindowsVistaOrGreater = IsWindowsVistaOrGreater_
End Property

Public Property Get IsWindows7OrGreater() As Boolean
    IsWindows7OrGreater = IsWindows7OrGreater_
End Property

Public Property Get IsWindows8OrGreater() As Boolean
    IsWindows8OrGreater = IsWindows8OrGreater_
End Property

Public Property Get IsWindows8Point1OrGreater() As Boolean
    IsWindows8Point1OrGreater = IsWindows8Point1OrGreater_
End Property

Public Property Get IsWindows10OrGreater() As Boolean
    IsWindows10OrGreater = IsWindows10OrGreater_
End Property

Public Property Get CodepageOEM() As String
    CodepageOEM = CodepageOEM_
End Property

Public Property Get CodepageOEM_File() As String
    CodepageOEM_File = CodepageOEM_File_
End Property

Public Property Get CodepageANSI() As String
    CodepageANSI = CodepageANSI_
End Property

Public Property Get CodepageANSI_File() As String
    CodepageANSI_File = CodepageANSI_File_
End Property

Public Property Get SuiteMask() As String
    SuiteMask = SuiteMask_
End Property

Public Property Get ProductType() As String
    ProductType = ProductType_
End Property

Public Property Get SID_CurrentProcess() As String
    SID_CurrentProcess = SID_CurrentProcess_
End Property

Public Property Get SafeBootMode() As String
    SafeBootMode = SafeBootMode_
End Property

Public Property Get Platform() As String
    Platform = Platform_
End Property

Public Property Get ReleaseId() As Long
    ReleaseId = ReleaseId_
End Property

Public Property Get IsServer() As Boolean
    IsServer = IsServer_
End Property

Public Property Get IsDomainController() As Boolean
    IsDomainController = IsDomainController_
End Property

Public Property Get IsWin64() As Boolean
    IsWin64 = IsWin64_
End Property

Public Property Get IsWin32() As Boolean
    IsWin32 = Not IsWin64_
End Property

Public Property Get PlatformID() As Long
    PlatformID = PlatformID_
End Property

Public Property Get SecureBoot() As Boolean
    SecureBoot = SecureBoot_
End Property

Public Property Get IsLocalSystemContext() As Boolean
    IsLocalSystemContext = IsLocalSystemContext_
End Property

Public Property Get IsAdmin() As Boolean
    IsAdmin = IsAdmin_
End Property

Public Property Get IsSystemCaseSensitive() As Boolean
    IsSystemCaseSensitive = IsSystemCaseSensitive_
End Property

Public Property Get IsEmbedded() As Boolean
    IsEmbedded = IsEmbedded_
End Property
