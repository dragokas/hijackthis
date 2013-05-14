Attribute VB_Name = "modMain"
'R0 - Changed Registry value (MSIE)
'R1 - Created Registry value
'R2 - Created Registry key
'R3 - Created extra value in regkey where only one should be
'F0 - Changed inifile value (system.ini)
'F1 - Created inifile value (win.ini)
'N1 - Changed NS4.x homepage
'N2 - Changed NS6 homepage
'N3 - Changed NS7 homepage/searchpage
'N4 - Changed Moz homepage/searchpage
'O1 - Hosts file hijack
'O2 - BHO
'O3 - IE Toolbar
'O4 - Regrun entry
'O5 - Control.ini IE Options block
'O6 - Policies IE Options/Control Panel block
'O7 - Policies Regedit block
'O8 - IE Context menuitem
'O9 - IE Tools menuitem/button
'O10 - Winsock hijack
'O11 - IE Advanced Options group
'O12 - IE Plugin
'O13 - IE DefaultPrefix hijack
'O14 - IERESET.INF hijack
'O15 - Trusted Zone autoadd, e.g. free.aol.com
'O16 - Downloaded Program Files
'O17 - Domain hijacks in CurrentControlSet
'O18 - Protocol & Filter enum
'O19 - User style sheet hijack
'O20 - AppInit_DLLs registry value + Winlogon Notify subkeys
'O21 - ShellServiceObjectDelayLoad enumeration
'O22 - SharedTaskScheduler enumeration

'Next possible methods:
'* SearchAccurates 'URL' method in a InitPropertyBag (??)
'* HKLM\..\CurrentVersion\ModuleUsage
'* HKLM\..\CurrentVersion\Explorer\ShellExecuteHooks (eudora)
'* HKLM\..\Internet Explorer\SafeSites (searchaccurate)

Option Explicit
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long

Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal lSize As Long)
Private Declare Function OpenProcessToken Lib "Advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function OpenThreadToken Lib "Advapi32" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetTokenInformation Lib "Advapi32" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "Advapi32" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function IsValidSid Lib "Advapi32" (ByVal pSid As Long) As Long
Private Declare Function EqualSid Lib "Advapi32" (pSid1 As Any, pSid2 As Any) As Long
Private Declare Sub FreeSid Lib "Advapi32" (pSid As Any)

Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (ByRef OldValue As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByRef OldValue As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'For O24
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_UPDATEINIFILE = &H1


Private Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

Private Type SID_AND_ATTRIBUTES
    Sid As Long
    Attributes As Long
End Type

Private Type TOKEN_GROUPS
    GroupCount As Long
    Groups(20) As SID_AND_ATTRIBUTES
End Type

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

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
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
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

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const SECURITY_NT_AUTHORITY = &H5
Private Const TOKEN_QUERY = &H8
Private Const TokenGroups = 2
Private Const SECURITY_BUILTIN_DOMAIN_RID = &H20
Private Const DOMAIN_ALIAS_RID_ADMINS = &H220
Private Const DOMAIN_ALIAS_RID_USERS = &H221
Private Const DOMAIN_ALIAS_RID_GUESTS = &H222
Private Const DOMAIN_ALIAS_RID_POWER_USERS = &H223

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_SILENT = &H4

Private Const SM_CLEANBOOT = &H43 '67

Private Const SC_MANAGER_CREATE_SERVICE = &H2
Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
'Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
Private Const CHINESESIMPLIFIED_CHARSET = 134

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_OVERWRITEPROMPT = &H2

Private lWow64Old&
Public iItems%
Public sRegVals$()
Public sFileVals$()

Public bAutoSelect As Boolean
Public bConfirm As Boolean
Public bMakeBackup As Boolean
Public bIgnoreSafe As Boolean
Public bLogProcesses As Boolean
Public sHostsFile$
Public bIsWinNT As Boolean, bIsWinME As Boolean
'Public sWinDir$, sWinSysDir$
Private sIgnoreList() As String
Public bDebugMode As Boolean
Public sWinVersion$, sMSIEVersion$
Public bRebootNeeded As Boolean

Public bIsUSADateFormat As Boolean
Public bNoWriteAccess As Boolean
Public bSeenLSPWarning As Boolean

Public sSafeLSPFiles$
Public sSafeProtocols$()
Public sSafeRegDomains$()
Public sSafeSSODL$()
Public sSafeFilters$()
Public sSafeAppInit$
Public sSafeWinlogonNotify$

Public sProgramVersion$ 'encryption phrase

Public bShownBHOWarning As Boolean
Public bShownToolbarWarning As Boolean

Public bMD5 As Boolean
Public bIgnoreAllWhitelists As Boolean
Public bAutoLog As Boolean, bAutoLogSilent As Boolean
Public bLogEnvVars As Boolean
Private bTriedFixUnixHostsFile As Boolean
Public bSeenHostsFileAccessDeniedWarning

Public Sub LoadStuff()
    On Error GoTo Error:
    '=== LOAD FILEVALS ===
    'syntax:
    ' inifile,section,value,resetdata,baddata
    ' |       |       |     |         |
    ' |       |       |     |         1) data that shouldn't be (never used)
    ' |       |       |     2) data to reset to
    ' |       |       |        (delete all if empty)
    ' |       |       3) value to check
    ' |       4) section to check
    ' 5) file to check
    
    ReDim sFileVals(6)
    sFileVals(0) = "system.ini,boot,Shell,explorer.exe,"
    'sFileVals(0) = "icdk'bvL\_LR`e6!IOHZbtVo2aYUShNUi[L"
    sFileVals(1) = "win.ini,windows,load,,"
    'sFileVals(1) = "mS_%+cSme_0T`m5!bVDR""t"
    sFileVals(2) = "win.ini,windows,run,,"
    'sFileVals(2) = "mS_%+cSme_0T`m5!h\Qx"""
    sFileVals(3) = "REG:system.ini,boot,Shell,explorer.exe,"
    'sFileVals(3) = "H/815n]WScNY__LWeVWxIRVc.!O[^b1bVhNZnLm"
    sFileVals(4) = "REG:win.ini,windows,load,,"
    'sFileVals(4) = "H/819^XoWd+zh_0Ye^VxbYR[L!"
    sFileVals(5) = "REG:win.ini,windows,run,,"
    'sFileVals(5) = "H/819^XoWd+zh_0Ye^Vxh__#L"
    sFileVals(6) = "REG:system.ini,boot,UserInit,$WINDIR\System32\UserInit.exe,"
    'sFileVals(6) = "H/815n]WScNY__LWeVWxK]ViicSWxxw9?:iGR:\ajO^*RQ?VShi^ZjNZnLm"

    '=== LOAD REGVALS ===
    'syntax:
    '  regkey,regvalue,resetdata,baddata
    '  |      |        |          |
    '  |      |        |          1) data that shouldn't be (never used)
    '  |      |        2) data to reset to
    '  |      3) value to check
    '  4) regkey to check
    '
    'when empty:
    '1) everything is considered bad (always used), change to resetdata
    '2) value being present is considered bad, delete value
    '3) key being present is considered bad, delete key (not used)
    '4) [invalid]
    
    ReDim sRegVals(103)
    '             HKCU\Software\Microsoft\Internet Explorer,Default_Page_URL,,
    sRegVals(0) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUSht5\(V_ObUpQX[!JH3mx"
    '             HKCU\Software\Microsoft\Internet Explorer,Default_Search_URL,,
    sRegVals(1) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUSht5\(V_ObUsURh%]U<5:""t"
    '             HKCU\Software\Microsoft\Internet Explorer,SearchAssistant,,
    sRegVals(2) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMK/i5Ydj#cjqm"
    '             HKCU\Software\Microsoft\Internet Explorer,CustomizeSearch,,
    sRegVals(3) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUSht4l5iYPWp'CVW4X^qm"
    '             HKCU\Software\Microsoft\Internet Explorer,Search,,
    sRegVals(4) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMKx"""
    '             HKCU\Software\Microsoft\Internet Explorer,Search Bar,,
    sRegVals(5) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMKl8#b{"""
    '             HKCU\Software\Microsoft\Internet Explorer,Search Page,,
    sRegVals(6) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMKlF#WV""L"
    '             HKCU\Software\Microsoft\Internet Explorer,Start Page,,
    sRegVals(7) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtDk#g^a>W)U{"""
    '             HKCU\Software\Microsoft\Internet Explorer,SearchURL,,
    sRegVals(8) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMKCHlz{"
    '             HKCU\Software\Microsoft\Internet Explorer,(Default),,
    sRegVals(9) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtw;'[KXZjIz{"
    '             HKCU\Software\Microsoft\Internet Explorer,www,,
    sRegVals(10) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShthn9!t"
   
    '             HKLM\Software\Microsoft\Internet Explorer,Default_Page_URL,,
    sRegVals(11) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUSht5\(V_ObUpQX[!JH3mx"
    '             HKLM\Software\Microsoft\Internet Explorer,Default_Search_URL,,
    sRegVals(12) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUSht5\(V_ObUsURh%]U<5:""t"
    '             HKLM\Software\Microsoft\Internet Explorer,SearchAssistant,,
    sRegVals(13) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMK/i5Ydj#cjqm"
    '             HKLM\Software\Microsoft\Internet Explorer,CustomizeSearch,,
    sRegVals(14) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUSht4l5iYPWp'CVW4X^qm"
    '             HKLM\Software\Microsoft\Internet Explorer,Search,,
    sRegVals(15) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMKx"""
    '             HKLM\Software\Microsoft\Internet Explorer,Search Bar,,
    sRegVals(16) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMKl8#b{"""
    '             HKLM\Software\Microsoft\Internet Explorer,Search Page,,
    sRegVals(17) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMKlF#WV""L"
    '             HKLM\Software\Microsoft\Internet Explorer,Start Page,,
    sRegVals(18) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtDk#g^a>W)U{"""
    '             HKLM\Software\Microsoft\Internet Explorer,SearchURL,,
    sRegVals(19) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtD\#gMKCHlz{"
    '             HKLM\Software\Microsoft\Internet Explorer,(Default),,
    sRegVals(20) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShtw;'[KXZjIz{"
    '             HKLM\Software\Microsoft\Internet Explorer,www,,
    sRegVals(21) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShthn9!t"

    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,Default_Page_URL,,
    sRegVals(22) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShL4V\#jb[B>WQVVuG6mx"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,Default_Search_URL,,
    sRegVals(23) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShL4V\#jb[BA[KcZ*T?5:""L"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,SearchAssistant,,
    sRegVals(24) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShLCVW4X^(Va_]eX0itm"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,CustomizeSearch,,
    sRegVals(25) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShL3fi6dcP]SIORi%]tm"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,Search,,
    sRegVals(26) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShLCVW4X^qm"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,Search Bar,,
    sRegVals(27) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShLCVW4X^e%Oht{"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,Search Page,,
    sRegVals(28) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShLCVW4X^e3O]O{#"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,Start Page,,
    sRegVals(29) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShLCeW4it7DU[t{"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,SearchURL,,
    sRegVals(30) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShLCVW4X^<5:""t"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,(Default),,
    sRegVals(31) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShLv5[(VkSWu""t"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer,www,,
    sRegVals(32) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUShLghmL!"
    
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Default_Page_URL,,
    sRegVals(33) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct'S\#e]j!EWNHMK<=#L"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Default_Search_URL,,
    sRegVals(34) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct'S\#e]j!H[HUQ^IFIl!t"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,SearchAssistant,,
    sRegVals(35) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6SW4SY75h_ZWOd^{#"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,CustomizeSearch,,
    sRegVals(36) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct&ci6_^_<ZILD`YR{#"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Search Bar,,
    sRegVals(37) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6SW4SYtbVhqm"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Search Page,,
    sRegVals(38) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6SW4SYtpV]Lmx"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Start Page,$DEFSTARTPAGE,
    sRegVals(39) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6bW4doF#\[qe2;0DKaG>3/=ez"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,SearchURL,,
    sRegVals(40) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6SW4SYKrA""q"

    '             HKLM\Software\Microsoft\Internet Explorer\Main,Default_Page_URL,,
    sRegVals(41) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct'S\#e]j!EWNHMK<=#L"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,Default_Search_URL,,
    sRegVals(42) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct'S\#e]j!H[HUQ^IFIl!t"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,SearchAssistant,,
    sRegVals(43) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6SW4SY75h_ZWOd^{#"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,CustomizeSearch,,
    sRegVals(44) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct&ci6_^_<ZILD`YR{#"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,Search Bar,,
    sRegVals(45) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6SW4SYtbVhqm"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,Search Page,,
    sRegVals(46) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6SW4SYtpV]Lmx"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,Start Page,$DEFSTARTPAGE,
    sRegVals(47) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6bW4doF#\[qe2;0DKaG>3/=ez"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,SearchURL,,
    sRegVals(48) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6SW4SYKrA""q"
    
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Default_Page_URL,,
    sRegVals(49) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!:LIOkVeVpVQHMKr<{"""
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Default_Search_URL,,
    sRegVals(50) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!:LIOkVeVsZKUQ^!ECBL!"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,SearchAssistant,,
    sRegVals(51) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!ILD`YR2j5^]WOd6z{"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,CustomizeSearch,,
    sRegVals(52) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!9\VbeWZq'HOD`Y*z{"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Search Bar,,
    sRegVals(53) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!ILD`YRo9#gtm"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Search Page,,
    sRegVals(54) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!ILD`YRoG#\Omx"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Start Page,$DEFSTARTPAGE,
    sRegVals(55) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!I[D`jhAX)Zte2;fCE7rIF(*3"""
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,SearchURL,,
    sRegVals(56) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!ILD`YRFIl!t"

    '             HKCU\Software\Microsoft\Internet Explorer\Search,Default_Search_URL,,
    sRegVals(57) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKx:'VRk.iU:HOhMYVuG6mx"
    '             HKCU\Software\Microsoft\Internet Explorer\Search,SearchAssistant,,
    sRegVals(58) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKxI'QcY*6iZLajK_kL!"
    '             HKCU\Software\Microsoft\Internet Explorer\Search,CustomizeSearch,,
    sRegVals(59) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKx97cee/^pL6SW\T_L!"
    '             HKCU\Software\Microsoft\Internet Explorer\Search,(Default),,
    sRegVals(60) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKx|dUWW7ajnmx"
    '             HKCU\Software\Microsoft\Internet Explorer\Search,Default_Page_URL,,
    sRegVals(61) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKx:'VRk.iU7DU[IFIl!t"

    '             HKLM\Software\Microsoft\Internet Explorer\Search,Default_Search_URL,,
    sRegVals(62) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKx:'VRk.iU:HOhMYVuG6mx"
    '             HKLM\Software\Microsoft\Internet Explorer\Search,SearchAssistant,$DEFSEARCHASS,
    sRegVals(63) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKxI'QcY*6iZLajK_kLw.(4Ie1C9h6I:m"
    '             HKLM\Software\Microsoft\Internet Explorer\Search,CustomizeSearch,$DEFSEARCHCUST,
    sRegVals(64) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKx97cee/^pL6SW\T_Lw.(4Ie1C9h8K:7x"
    '             HKLM\Software\Microsoft\Internet Explorer\Search,(Default),,
    sRegVals(65) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKx|dUWW7ajnmx"
    '             HKLM\Software\Microsoft\Internet Explorer\Search,Default_Page_URL,,
    sRegVals(66) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKx:'VRk.iU7DU[IFIl!t"
    
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Search,Default_Search_URL,,
    sRegVals(67) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|CVW4X^q'S\Kfc6T=HOh%XPKrA""q"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Search,SearchAssistant,$DEFSEARCHASS,
    sRegVals(68) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|CVW4X^q6SW\T_ah]Laj#^e""D9;-637<4?aH=m"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Search,CustomizeSearch,$DEFSEARCHCUST,
    sRegVals(69) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|CVW4X^q&ci^`d+oO6SW4SY""D9;-637<4?cJ=7x"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Search,(Default),,
    sRegVals(70) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|CVW4X^qi2[PRl.iqmx"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Search,Default_Page_URL,,
    sRegVals(71) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|CVW4X^q'S\Kfc6T:DU[!ECBL!"

    '             HKCU\Software\Microsoft\Internet Explorer\SearchURL,(Default),,
    sRegVals(72) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKCHlzw:'[W\Ob}t{"
    '             HKCU\Software\Microsoft\Internet Explorer\SearchURL,SearchURL,,
    sRegVals(73) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKCHlzD[#gYO8@Bt{"

    '             HKLM\Software\Microsoft\Internet Explorer\SearchURL,(Default),,
    sRegVals(74) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKCHlzw:'[W\Ob}t{"
    '             HKLM\Software\Microsoft\Internet Explorer\SearchURL,SearchURL,,
    sRegVals(75) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFD\#gMKCHlzD[#gYO8@Bt{"

    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\SearchURL,(Default),,
    sRegVals(76) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|CVW4X^<5:""p5\(V_Ob}Lz"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\SearchURL,SearchURL,,
    sRegVals(77) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|CVW4X^<5:""=VX4XR8@BLz"
    
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Startpagina,,
    sRegVals(78) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6bW4daW)^dHmx"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,First Home Page,,
    sRegVals(79) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct)Wh5do>1b[e3O]O{#"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Local Page,$WINSYSDIR\blank.htm,
    sRegVals(80) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct/]Y#\oF#\[qeE?8DPs935JX.Q_aN]jTm"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Start Page_bak,,
    sRegVals(81) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6bW4doF#\[FEOat{"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,HomeOldSP,,
    sRegVals(82) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct+]c'?]ZsE""q"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,Use Custom Search URL,,
    'sRegVals() = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct8a[@3fi6dce6SW\T_@J</x"""
    '                HKCU\Software\Microsoft\Internet Explorer\Main,Úvodní stránka,,
    'sRegVals() = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ctÚde&^ít5iháQYWt{"

    '             HKLM\Software\Microsoft\Internet Explorer\Main,Startpagina,,
    sRegVals(83) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6bW4daW)^dHmx"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,First Home Page,,
    sRegVals(84) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct)Wh5do>1b[e3O]O{#"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,Local Page,$WINSYSDIR\blank.htm,
    sRegVals(85) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct/]Y#\oF#\[qeE?8DPs935JX.Q_aN]jTm"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,Start Page_bak,,
    sRegVals(86) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct6bW4doF#\[FEOat{"
    '             HKLM\Software\Microsoft\Internet Explorer\Main,HomeOldSP,,
    sRegVals(87) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct+]c'?]ZsE""q"
    '                HKLM\Software\Microsoft\Internet Explorer\Main,Úvodní stránka,,
    'sRegVals() = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ctÚde&^ít5iháQYWt{"
    
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Startpagina,,
    sRegVals(88) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!I[D`jZR^+cKmx"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,First Home Page,,
    sRegVals(89) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!<PUajh9f/Zh3O]'z{"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Local Page,$WINSYSDIR\blank.htm,
    sRegVals(90) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!BVFObhAX)ZteE?nCJId>HCEZWX\%*iWm"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Start Page_bak,,
    sRegVals(91) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!I[D`jhAX)ZIEOaLz"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,HomeOldSP,,
    sRegVals(92) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!>VPSEVUJp!t"
    '                HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,Úvodní stránka,,
    'sRegVals() = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!Ú]RRdíoj6gáQYWLz"
    
    '             HKLM\Software\Microsoft\Internet Explorer\Main,YAHOOSubst,,
    sRegVals(93) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct</>o?Dk$hjqm"
    '             HKCU\Software\Microsoft\Internet Explorer\Main,YAHOOSubst,,
    sRegVals(94) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct</>o?Dk$hjqm"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Explorer\Main,YAHOOSubst,,
    sRegVals(95) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh6o2aYUSh|=R_0!O(+=E=fY5itm"
    
    '             HKCU\Software\Microsoft\Internet Connection Wizard,ShellNext,,
    sRegVals(96) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do91cdLFb_Y_uw^dD`ZLCY[.aDL[b""t"
    '             HKLM\Software\Microsoft\Internet Connection Wizard,ShellNext,,
    sRegVals(97) = ">5=D|HYIbm#bVRm^YYRaePeSic^H`d'do91cdLFb_Y_uw^dD`ZLCY[.aDL[b""t"
    '             HKUS\.DEFAULT\Software\Microsoft\Internet Connection Wizard,ShellNext,,
    sRegVals(98) = ">5FJ|#.(47u<ERsd\[ZOhOMD+X\Rae(dM?0i[YQSjh4f0cOFb_1^oM+oWYGxIRVc.CO[b""L"

    '             HKCU\Software\Microsoft\Internet Explorer\Main,Window Title,,
    sRegVals(99) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShF>X+ct:Wd&_htt^jSHx"""

    '             HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings,AutoConfigURL,,
    sRegVals(100) = ">54L|HYIbm#bVRm^YYRaePeSw^XG]m5L4k4g[UWD[\d`1cF,\j'b_[6sILWb_XXjL6_W]91^W_)JH3mx"
    '             HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ProxyServer,,
    sRegVals(101) = ">54L|HYIbm#bVRm^YYRaePeSw^XG]m5L4k4g[UWD[\d`1cF,\j'b_[6sILWb_XXjLE\RfosUcl'g""q"
    '             HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ProxyOverride,,
    sRegVals(102) = ">54L|HYIbm#bVRm^YYRaePeSw^XG]m5L4k4g[UWD[\d`1cF,\j'b_[6sILWb_XXjLE\RfoofVh4^ZLmx"
    '             HKCU\Software\Microsoft\Internet Explorer\Toolbar,LinksFolderName,Links,
    sRegVals(103) = ">54L|HYIbm#bVRm^YYRaePeSic^H`d'do;:ebVUShFEf1aLD`""lY_a5;eSGSh8Rd'!6L\a5z"

    
    ' === LOAD NONSTANDARD-BUT-SAFE-DOMAINS LIST ===
    ReDim sSafeRegDomains(11)
    'sSafeRegDomains(0) = "http://www.microsoft.com"
    sSafeRegDomains(0) = "^^egZ$wZemN]ZY4diVIb$M`d"
    'sSafeRegDomains(1) = "http://home.microsoft.com"
    sSafeRegDomains(1) = "^^egZ$wK]c'|^_%geZRTjvTf/"
    'sSafeRegDomains(2) = "http://www.msn.com"
    sSafeRegDomains(2) = "^^egZ$wZemN]ddNXeT"
    'sSafeRegDomains(3) = "http://search.msn.com"
    sSafeRegDomains(3) = "^^egZ$wVSW4SY$/hdsF]c"
    'sSafeRegDomains(4) = "http://ie.search.msn.com"
    sSafeRegDomains(4) = "^^egZ$wLS$5URh%]$TV\$M`d"
    'sSafeRegDomains(5) = "ie.search.msn.com"
    sSafeRegDomains(5) = "_O}j'V\FV$/c_$%dc"
    'sSafeRegDomains(6) = "http://my.yahoo.com"
    sSafeRegDomains(6) = "^^egZ$wPg$;QYe1#YVP"
    'sSafeRegDomains(7) = "http://www.aol.com"
    sSafeRegDomains(7) = "^^egZ$wZemNQ`bNXeT"
    'sSafeRegDomains(8) = "<local>" 'for proxy check
    sSafeRegDomains(8) = "2V`Z#a("
    'sSafeRegDomains(9) = "http://www.google.com"
    sSafeRegDomains(9) = "^^egZ$wZemNW`e)a[sF]c"
    'sSafeRegDomains(10) = "127.0.0.1;localhost"
    sSafeRegDomains(10) = "'z(%P#xo}1._TW.]eZW"
    'sSafeRegDomains(11) = "iexplore"
    sSafeRegDomains(11) = "_Oig.d\H"

    ' === LOAD LSP PROVIDERS SAFELIST ===
    'asterisk is used for filename separation, because.
    'did you ever see a filename with an asterisk?
    sSafeLSPFiles = "*A2antispamlsp.dll*Adlsp.dll*Agbfilt.dll*Antiyfilter.dll*Ao2lsp.dll*Aphish.dll*Asdns.dll*Aslsp.dll*Asnsp.dll*Avgfwafu.dll*Avsda.dll*Betsp.dll*Biolsp.dll*Bmi_lsp.dll*Caslsp.dll*Cavemlsp.dll*Cdnns.dll*Connwsp.dll*Cplsp.dll*Csesck32.dll*Cslsp.dll*Cssp.al*Ctxlsp.dll*Ctxnsp.dll*Cwhook.dll*Cwlsp.dll*Dcsws2.dll*Disksearchservicestub.dll*Drwebsp.dll*Drwhook.dll*Espsock2.dll*Farlsp.dll*Fbm.dll*Fbm_lsp.dll*Fortilsp.dll*Fslsp.dll*Fwcwsp.dll*Fwtunnellsp.dll*Gapsp.dll*Googledesktopnetwork1.dll*Hclsock5.dll*Iapplsp.dll*Iapp_lsp.dll*Ickgw32i.dll*Ictload.dll*Idmmbc.dll*Iga.dll*Imon.dll*Imslsp.dll*Inetcntrl.dll*Ippsp.dll*Ipsp.dll*Iss_clsp.dll*Iss_slsp.dll*Kvwsp.dll*Kvwspxp.dll*Lslsimon.dll*Lsp32.dll*" & _
        "Lspcs.dll*Mclsp.dll*Mdnsnsp.dll*Msafd.dll*Msniffer.dll*Mswsock.dll*Mswsosp.dll*Mwtsp.dll*Mxavlsp.dll*Napinsp.dll*Nblsp.dll*Ndpwsspr.dll*Netd.dll*Nihlsp.dll*Nlaapi.dll*Nl_lsp.dll*Nnsp.dll*Normanpf.dll*Nutafun4.dll*Nvappfilter.dll*Nwws2nds.dll*Nwws2sap.dll*Nwws2slp.dll*Odsp.dll*Pavlsp.dll*Pclsp.dll*Pctlsp.dll*Pfftsp.dll*Pgplsp.dll*Pidlsp.dll*Pnrpnsp.dll*Prifw.dll*Proxy.dll*Prplsf.dll*Pxlsp.dll*Rnr20.dll*Rsvpsp.dll*S5spi.dll*Samnsp.dll*Sarah.dll*Scopinet.dll*Skysocks.dll*Sliplsp.dll*Smnsp.dll*Spacklsp.dll*Spampallsp.dll*Spi.dll*Spidll.dll*Spishare.dll*Spsublsp.dll*Sselsp.dll*Stplayer.dll*Syspy.dll*Tasi.dll*Tasp.dll*Tcpspylsp.dll*Ua_lsp.dll*Ufilter.dll*Vblsp.dll*Vetredir.dll*Vlsp.dll*Vnsp.dll*" & _
        "Wglsp.dll*Whllsp.dll*Whlnsp.dll*Winrnr.dll*Wins4f.dll*Winsflt.dll*WinSysAM.dll*Wps.dll*Wshbth.dll*Wspirda.dll*Wspwsp.dll*Xfilter.dll*xfire_lsp.dll*Xnetlsp.dll*Ypclsp.dll*Zklspr.dll*_Easywall.dll*_Handywall.dll*"
    
    ' === LOAD PROTOCOL SAFELIST ===
    ReDim sSafeProtocols(104) '(O18)
    'sSafeProtocols(0) = "about|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
    sSafeProtocols(0) = "WL`l6qet|+P6%&V""/}%##y"":f"",%&(M~!7a%&)'1;x3t"
    'sSafeProtocols(1) = "belarc|{6318E0AB-2E93-11D1-B8ED-00608CC9A71F}"
    sSafeProtocols(1) = "XO]X4Xf^$)Q(6&a7#w(')u""(d&u%&;d{!&V%.*&'7!""=?"
    'sSafeProtocols(2) = "BPC|{3A1096B3-9BFA-11D1-AE77-00C04FBBDEBC}"
    sSafeProtocols(2) = "8:4s=(+r|/V2$#Y7<(n}'.""$a:!xy&P3!*f78+(09g"
    'sSafeProtocols(3) = "CDL|{3DD53D40-7B8B-11D0-B013-00AA0059CE02}"
    sSafeProtocols(3) = "9.=s=(.'#)d$!#W7.)n}'.!$b%yty&P12&P*/*(|(g"
    'sSafeProtocols(4) = "cdo|{CD00020A-8B95-11D1-82DB-00C04FB1625D}"
    sSafeProtocols(4) = "YN`s=8.q|&R~2#X7/zn}'.""$X'.%y&P3!*f7'{s#:g"
    'sSafeProtocols(5) = "copernicagentcache|{AAC34CFD-274D-4A9D-B0DC-C74C05A67E1D}"
    sSafeProtocols(5) = "YYa\4cSFO]'^eY#X^L_i7+4*T80'y(W$5#T6/+n0&.4$c,|&|+a&(;Q9s"
    'sSafeProtocols(6) = "copernicagent|{A979B6BD-E40B-4A07-ABDD-A62C64A4EBF6}"
    sSafeProtocols(6) = "YYa\4cSFO]'^er=6/|z0,,5$e)x%y*a~(#a7:+n/,z4-T6|(0<Vm"
    'sSafeProtocols(7) = "dodots|{9446C008-3810-11D4-901D-00B0D04158D2}"
    sSafeProtocols(7) = "ZYUf6hf^'*T&4&P-#xy}&u""(d)uz|'d{!&b%:uu}+""5)?"
    'sSafeProtocols(8) = "DVD|{12D51199-0DB5-46FE-A120-47A3D7D937CC}"
    sSafeProtocols(8) = ":@5s=&z'#'Q)*#P98zn"",06$a&zqy*W1$:W9/xx19g"
    'sSafeProtocols(9) = "file|{79EAC9E7-BAF9-11CE-8C82-00AA004BA90B}"
    sSafeProtocols(9) = "\S]\>p!z37c)6-M77-zy'y4<M--y~#P~27P%*)$'&,n"
    'sSafeProtocols(10) = "ftp|{79EAC9E3-BAF9-11CE-8C82-00AA004BA90B}"
    sSafeProtocols(10) = "\^as=,#(/9Y5$#b6<~n}'-6$X8""sy&P12&P)8(z|8g"
    'sSafeProtocols(11) = "gopher|{79EAC9E4-BAF9-11CE-8C82-00AA004BA90B}"
    sSafeProtocols(11) = "]Ya_'gf^%/e14/e)#)$4/u""(c:uy1.R{!&a6&uu07#!9?"
    'sSafeProtocols(12) = "https|{79EAC9E5-BAF9-11CE-8C82-00AA004BA90B}"
    sSafeProtocols(12) = "^^eg5qex';a3*;U""8()'#y"":e""""&&(M~!7a%&y%//x3t"
    'sSafeProtocols(13) = "http|{79EAC9E2-BAF9-11CE-8C82-00AA004BA90B}"
    sSafeProtocols(13) = "^^eg>p!z37c)6(M77-zy'y4<M--y~#P~27P%*)$'&,n"
    'sSafeProtocols(14) = "ic32pp|{BBCA9F81-8F4F-11D2-90FF-0080C83D3571}"
    sSafeProtocols(14) = "_M$)2ef^08c1*<X&#})""<u""(d'uz|<f{!&X%9}t2)}((?"
    'sSafeProtocols(15) = "ipp|"
    sSafeProtocols(15) = "_Zas"
    'sSafeProtocols(16) = "its|{9D148291-B9C8-11D0-A4CC-0000F80149F6}"
    sSafeProtocols(16) = "_^ds=..r"".R)""#b.9}n}'.!$a)-&y&P~!<X%'yz4,g"
    'sSafeProtocols(17) = "javascript|{3050F3B2-98B5-11CF-BB82-00AA00BDCE0B}"
    sSafeProtocols(17) = "`KgX5X\L^j>k$&U%<x%~##)9U""yr1<M23.R""&u$/&x3;c:x%k"
    'sSafeProtocols(18) = "junomsg|{C4D10830-379D-11D4-9B2D-00C04F1579A5}"
    sSafeProtocols(18) = "`__f/hQ_i9T4""&X(&rt%/.|(Q9|n'8R4|&P8&y)}+!*8Ur"
    'sSafeProtocols(19) = "lid|{5C135180-9973-46D9-ABF4-148267CBB8BF}"
    sSafeProtocols(19) = "bSUs=*-r!+Q(!#Y.-xn"",.*$a70uy'T(#,W88)y0<g"
    'sSafeProtocols(20) = "local|{79EAC9E7-BAF9-11CE-8C82-00AA004BA90B}"
    sSafeProtocols(20) = "bYTX.qex';a3*;W""8()'#y"":e""""&&(M~!7a%&y%//x3t"
    'sSafeProtocols(21) = "mailto|{3050F3DA-98B5-11CF-BB82-00AA00BDCE0B}"
    sSafeProtocols(21) = "cKZc6df^!&U~7)d6#~y0+u""(c;u%0.R{!&a6&u%29/!9?"
    'sSafeProtocols(22) = "mctp|{D7B95390-B1C5-11D0-B111-0080C712FE82}"
    sSafeProtocols(22) = "cMeg>p.x0/U#*&M7'*vy'y5'M7yr}#P~)&c,'w)3.zn"
    'sSafeProtocols(23) = "mhtml|{05300401-BCBC-11D0-85E3-00C04FD85AB4}"
    sSafeProtocols(23) = "cRed.qeq#)P~%&Q""8*%1#y"";P""""v3)M~!9P)<+y#7,%t"
    'sSafeProtocols(24) = "mk|{79EAC9E6-BAF9-11CE-8C82-00AA004BA90B}"
    sSafeProtocols(24) = "cUmrW./$1/e&|8a;/rr}9/|/c-zn|&a1!&T77~q0s"
    'sSafeProtocols(25) = "ms-its50|{F8606A00-F5CF-11D1-B6BB-0000F80149F6}"
    sSafeProtocols(25) = "c]|`6h}qjqf('&V6&un4+-7$Q&.ry8V23#P%&u)&&y%0f+g"
    'sSafeProtocols(26) = "ms-its51|{F6F1E82D-DE4D-11D2-875C-0000F8105754}"
    sSafeProtocols(26) = "c]|`6h}rjqf&7'e-(+n2;|5$Q&.sy.W%4#P%&u)&'x&.U)g"
    'sSafeProtocols(27) = "ms-itss|{0A9007C0-4076-11D3-8789-0000F8105754}"
    sSafeProtocols(27) = "c]|`6h]_i&a)!&W8&ru|-~|(Q9{n&-X)|&P%&-y}&}(,Tr"
    'sSafeProtocols(28) = "ms-its|{9D148291-B9C8-11D0-A4CC-0000F80149F6}"
    sSafeProtocols(28) = "c]|`6hf^':Q$)(Y&#)z1.u""(d%u$""9c{!&P%<}q}*#7-?"
    'sSafeProtocols(29) = "msdaipp|"
    sSafeProtocols(29) = "c]UX+eZ_"
    'sSafeProtocols(30) = "mso-offdap|{3D9F03FA-7A94-11D3-BE81-0050048385D1}"
    sSafeProtocols(30) = "c]`$1[PGOf>k$:Y;&x)/#!20T""yr2)M26.Q""&uv|&|)*X*.rk"
    'sSafeProtocols(31) = "ndwiat|{13F3EA8B-91D7-4F0A-AD76-D2853AC8BECE}"
    sSafeProtocols(31) = "dNh`#if^})f#67X7#~r2-u%=P6u$2-V{5(X*)(&&8/4<?"
    'sSafeProtocols(32) = "res|{3050F3BC-98B5-11CF-BB82-00AA00BDCE0B}"
    sSafeProtocols(32) = "hOds=(xv|<S24#Y-8zn}'-7$b7""sy&P12&P7:*(|8g"
    'sSafeProtocols(33) = "sysimage|{76E67A63-06E9-11D2-A840-006008059382}"
    sSafeProtocols(33) = "icd`/VQHjqW&6,W6,xn|,/*$Q&.sy7X$!#P%,uq&&}**X'g"
    'sSafeProtocols(34) = "tve-trigger|{CBD30859-AF45-11D2-B6D6-00C04FBBDE6E}"
    sSafeProtocols(34) = "j`V$6gSJU[4ll9b9)uy#/u2=T*ur}:R{3,d+#uq1&|79b9/w3s"
    'sSafeProtocols(35) = "tv|{CBD30858-AF45-11D2-B6D6-00C04FBBDE6E}"
    sSafeProtocols(35) = "j`mrc7.t|.U(|7f)+rr}:z|9V9~n|&c~%<b7:,w3s"
    'sSafeProtocols(36) = "vbscript|{3050F3B2-98B5-11CF-BB82-00AA00BDCE0B}"
    sSafeProtocols(36) = "lLdZ4^ZWjqS~&&f(8wn'.,&$Q&-)y8b(##P%7(q|8.4<P7g"
    'sSafeProtocols(37) = "vnd.ms.radio|{3DA2AA3B-3D96-11D2-9BD2-204C4F4F5020}"
    sSafeProtocols(37) = "lXU%/hvUOZ+_mqS97w$/),|*d.~n}'d""|/b9(rs|*-%=T;}q~&?"
    'sSafeProtocols(38) = "wia|{13F3EA8B-91D7-4F0A-AD76-D2853AC8BECE}"
    sSafeProtocols(38) = "mSRs=&{)!;a(3#Y&:|n""<x2$a9!wy:R(&)a8.)(1;g"
    'sSafeProtocols(39) = "mso-offdap11|{32505114-5902-49B2-880A-1F7738E5A384}"
    sSafeProtocols(39) = "c]`$1[PGOfQ!mqS'+uv}'||,Y%zn""/b""|.X%7rr4-!$/e*+t&*?"
    'sSafeProtocols(40) = "DirectDVD|{85A81A02-336B-43FF-998B-FE8E194FBA4D}"
    sSafeProtocols(40) = ":Sc\%i.92r=(&7X&7usy){'9M){)4#Y))8M;;}(}/|79a).`"
    'sSafeProtocols(41) = "pcn|{D540F040-F3D9-11D0-95BE-00C04FD93CA5}"
    sSafeProtocols(41) = "fM_s=9}u|<P$!#f(:~n}'.!$Y*,(y&P3!*f9/x&/+g"
    'sSafeProtocols(42) = "msencarta|{74D92DF3-6D9D-11D1-8B38-006097DBED7A}"
    sSafeProtocols(42) = "c]Ve%V\WOr='%:Y':-ty,.*;M&y'}#X2$.M%&{q'-.3<d,+`"
    'sSafeProtocols(43) = "msero|{B0D92A71-886B-453B-A649-1B91F93801E7}"
    sSafeProtocols(43) = "c]Vi1qe%|:Y""2-Q"".}w0#|&*b""+w""/M!3/Q;/xy|'/(t"
    'sSafeProtocols(44) = "msref|{74D92DF3-6D9D-11D1-8B38-006097DBED7A}"
    sSafeProtocols(44) = "c]c\(qex"":Y""5<S"",+z2#y"";Q""""%!.M~!,P.-+%3:!2t"
    'sSafeProtocols(45) = "df2|{219A97F3-D661-4766-B658-646A771AE49E}"
    sSafeProtocols(45) = "ZP#s='yz//W6$#d+,vn""-~'$b+}yy,T&2-W&7,u';g"
    'sSafeProtocols(46) = "df3|{219A97F3-D661-4766-B658-646A771AE49E}"
    sSafeProtocols(46) = "ZP$s='yz//W6$#d+,vn""-~'$b+}yy,T&2-W&7,u';g"
    'sSafeProtocols(47) = "df4|{219A97F3-D661-4766-B658-646A771AE49E}"
    sSafeProtocols(47) = "ZP%s='yz//W6$#d+,vn""-~'$b+}yy,T&2-W&7,u';g"
    'sSafeProtocols(48) = "df5|{219A97F3-D661-4766-B658-646A771AE49E}"
    sSafeProtocols(48) = "ZP&s='yz//W6$#d+,vn""-~'$b+}yy,T&2-W&7,u';g"
    'sSafeProtocols(49) = "df23chat|{219A97F3-D661-4766-B658-646A771AE49E}"
    sSafeProtocols(49) = "ZP#*%]KWjqR!*7Y,<xn2,~""$T,~wy8V%)#V),(x%'+6+Y:g"
    'sSafeProtocols(50) = "df5demo|{219A97F3-D661-4766-B658-646A771AE49E}"
    sSafeProtocols(50) = "ZP&['bY_i(Q)2/W;)r'$,y|+W+~n0,U(|,T+7|x}7/%0er"
    'sSafeProtocols(51) = "ofpjoin|{219A97F3-D661-4766-B658-646A771AE49E}"
    sSafeProtocols(51) = "ePaa1^X_i(Q)2/W;)r'$,y|+W+~n0,U(|,T+7|x}7/%0er"
    'sSafeProtocols(52) = "saphtmlp|{D1F8BD1E-7967-11D2-B43A-006094B9EADB}"
    sSafeProtocols(52) = "iKa_6bVSjqd!7.b9',n%/~($Q&.sy8T#2#P%,uz""8#68d7g"
    'sSafeProtocols(53) = "sapr3|{D1F8BD1E-7967-11D2-B43A-006094B9EADB}"
    sSafeProtocols(53) = "iKaiSqe'}<X25'e""-~w%#y"";R"",u!7M~!,P.*)z37.3t"
    'sSafeProtocols(54) = "lbxfile|{56831180-F115-11D2-B6AA-00104B2B9943}"
    sSafeProtocols(54) = "bLi]+aO_i+V($'Q-&r)}'}|(Q9zn0,a1|&P&&y%~8#*+Sr"
    'sSafeProtocols(55) = "lbxres|{24508F1B-9E94-40EE-9759-9AF5795ADF52}"
    sSafeProtocols(55) = "bLii'hf^~*U~)<Q7#~('*u%'e:uz%+Y{*7f*-~v/:0&)?"
    'sSafeProtocols(56) = "cetihpz|{CF184AD3-CDCB-4168-A3F7-8E447D129300}"
    sSafeProtocols(56) = "YOe`*ed_i9f!)*a9)r&29,|+Q+""n/)f'|.e)*|'}(#$'Pr"
    '= added in HJT 1.99.2 final=
    'sSafeProtocols(57)  = "aim|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}
    sSafeProtocols(57) = "WS^s=(xv|<T~'#Y-8zn}'-7$b7""sy&P12&P7:*(|8g"
    'sSafeProtocols(58)  = "shell|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}
    sSafeProtocols(58) = "iRVc.qet|+P6%&V""/}%##y"":f"",%&(M~!7a%&)'1;x3t"
    'sSafeProtocols(59)  = "asp|{8D32BA61-D15B-11D4-894B-000000000000}
    sSafeProtocols(59) = "W]as=-.t~8a&""#d&+)n}'.%$X.|%y&P~!&P%&uq|&g"
    'sSafeProtocols(60)  = "hsp|{8D32BA61-D15B-11D4-894B-000000000000}
    sSafeProtocols(60) = "^]as=-.t~8a&""#d&+)n}'.%$X.|%y&P~!&P%&uq|&g"
    'sSafeProtocols(61)  = "x-asp|{8D32BA61-D15B-11D4-894B-000000000000}
    sSafeProtocols(61) = "nuRj2qey2)R22,Q"":vv0#y"";T""""z""8M~!&P%&uq|&x!t"
    'sSafeProtocols(62)  = "x-hsp|{8D32BA61-D15B-11D4-894B-000000000000}
    sSafeProtocols(62) = "nuYj2qey2)R22,Q"":vv0#y"";T""""z""8M~!&P%&uq|&x!t"
    'sSafeProtocols(63)  = "x-zip|{8D32BA61-D15B-11D4-894B-000000000000}
    sSafeProtocols(63) = "nuk`2qey2)R22,Q"":vv0#y"";T""""z""8M~!&P%&uq|&x!t"
    'sSafeProtocols(64)  = "zip|{8D32BA61-D15B-11D4-894B-000000000000}
    sSafeProtocols(64) = "pSas=-.t~8a&""#d&+)n}'.%$X.|%y&P~!&P%&uq|&g"
    'sSafeProtocols(65)  = "bega|{A57721C9-B905-49B3-8BCA-B99FBB8C627E}
    sSafeProtocols(65) = "XOXX>p+v%-R!4/M7/uvy*#3*M-,&/#b)*<b7.*w~-/n"
    'sSafeProtocols(66)  = "bt2|{1730B77B-F429-498F-9B15-4514D83C8294}
    sSafeProtocols(66) = "X^#s=&!t|8W'3#f)(~n""/""7$Y7yvy*U!%:X(9}s'*g"
    'sSafeProtocols(67)  = "cetihpz|{CF184AD3-CDCB-4168-A3F7-8E447D129300}
    sSafeProtocols(67) = "YOe`*ed_i9f!)*a9)r&29,|+Q+""n/)f'|.e)*|'}(#$'Pr"
    'sSafeProtocols(68)  = "copernicdesktopsearch|{D9656C75-5090-45C3-B27E-436FBC7ACFA7}
    sSafeProtocols(68) = "YYa\4cSFR[5[ee2h[HUQ^fl;Y+}w1-U{&&Y%#yv1)u3)W:uu!,f24-a8<(xk"
    'sSafeProtocols(69)  = "crick|{B861500A-A326-11D3-A248-0080C8F7DE1E}
    sSafeProtocols(69) = "Y\ZZ-qe%&,Q%!&a""7xs$#y"";S""+s"".M~!.P8.-x2;y6t"
    'sSafeProtocols(70)  = "dadb|{82D6F09F-4AC2-11D3-8BD9-0080ADB8683C}
    sSafeProtocols(70) = "ZKUY>p""s2,f~*<M)7*sy'y5*M-,''#P~)&a98}w&)-n"
    'sSafeProtocols(71)  = "dialux|{8352FA4C-39C6-11D3-ADBA-00A0244FB1A2}
    sSafeProtocols(71) = "ZSRc7mf^&)U""77T8#xz1,u""(d(u$28a{!&a%(yu48y2)?"
    'sSafeProtocols(72)  = "emistp|{0EFAEA2E-11C9-11D3-88E3-0000E867A001}
    sSafeProtocols(72) = "[WZj6ef^|;f167R:#vr1/u""(d(uy&;S{!&P%;}w%7x!(?"
    'sSafeProtocols(73)  = "ezstor|{6344A3A0-96A7-11D4-88CC-000000000000}
    sSafeProtocols(73) = "[ddk1gf^$)T$2)a%#~w/-u""(d)uy&9c{!&P%&uq|&x!'?"
    'sSafeProtocols(74)  = "flowto|{C7101FB0-28FB-11D5-883A-204C4F4F5020}
    sSafeProtocols(74) = "\V`n6df^1-Q~""<b%#wy48u""(d*uy&)a{#&T8*-u4+x#'?"
    'sSafeProtocols(75)  = "g7ps|{9EACF0FB-4FC7-436E-989B-3197142AD979}
    sSafeProtocols(75) = "]!aj>p#(/9f~78M)<*xy*{'<M.""z0#S!*-Q)((''-#n"
    'sSafeProtocols(76)  = "intu-res|{9CE7D474-16F9-4889-9BB9-53E2008EAE8A}
    sSafeProtocols(76) = "_XelMgOVjqY36-d)-yn},0*$T-""zy/b2*#U(;wq|./2<X6g"
    'sSafeProtocols(77)  = "iwd|{EA5F5649-A6C7-11D4-9E3C-0020AF0FFB56}
    sSafeProtocols(77) = "_aUs=:+v4+V$*#a+9|n}'.%$Y:{&y&P""!7f%<-%#,g"
    'sSafeProtocols(78)  = "mavencache|{DB47FDC2-8C38-4413-9C78-D1A68BF24EED}
    sSafeProtocols(78) = "cKg\0XKFV[>k58T,<+&~#""4*X""|u})M)4-X"":v$$.,7)T:/'k"
    'sSafeProtocols(79)  = "ms-help|{314111C7-A502-11D2-BBCA-00C04F8EC294}
    sSafeProtocols(79) = "c]|_'aZ_i)Q$""'Q8-r$#&z|(Q9zn08c1|&P8&y)&;-#0Tr"
    'sSafeProtocols(80)  = "msnim|{828030A1-22C1-4009-854F-8E305202313F}
    sSafeProtocols(80) = "c]_`/qey~.P#!7Q""(w&}#|!'Y""""v""<M(6)P*(us!'{7t"
    'sSafeProtocols(81)  = "myrm|{4D034FC3-013F-4B95-B544-44D49ABE3E76}
    sSafeProtocols(81) = "cccd>p|'|)T64)M%'x)y*,*,M7}u""#T$5*Y68,t3-~n"
    'sSafeProtocols(82)  = "nbso|{DF700763-3EAD-4B64-9626-22BEEFF3EA47}
    sSafeProtocols(82) = "dLdf>p.)%&P'')M(;('y*,'+M.~s$#R""3;e;<x(/*!n"
    'sSafeProtocols(83)  = "nim|{3D206AE2-3039-413B-B748-3ACC562EC22A}
    sSafeProtocols(83) = "dS^s=(.s|,a5##S%)~n""'{3$b,|yy)a34+V';*s~7g"
    'sSafeProtocols(84)  = "OWC11.mso-offdap|{32505114-5902-49B2-880A-1F7738E5A384}
    sSafeProtocols(84) = "EA4(Q#WV]#1VWZ#erbt~+x&(Q)uv'&R{%/b'#}y|7u""=W,{y3+a#)*?"
    'sSafeProtocols(85)  = "pcl|{182D0C85-206F-4103-B4FA-DCC1FB0A0A44}
    sSafeProtocols(85) = "fM]s=&""s2&c(&#R%,-n""'x$$b)0$y:c3""<b%7u$""*g"
    'sSafeProtocols(86)  = "pure-go|{4746C79A-2042-4332-8650-48966E44ABA8}
    sSafeProtocols(86) = "f_c\M\Y_i*W$'9W.7rs|*z|+S(zn&,U~|*X.,{(""*+38Xr"
    'sSafeProtocols(87)  = "qrev|{9DE24BAC-FC3C-42C4-9FC4-76B3FAFDBD90}
    sSafeProtocols(87) = "g\Vm>p#'3(T229M;9x&y*z4+M.0&""#W&3)f6<+%2/xn"
    'sSafeProtocols(88)  = "rmh|{23C585BB-48FF-4865-8934-185F0A7EB84C}
    sSafeProtocols(88) = "hWYs='{&#.U23#T-<-n"".~&$X.{uy'X%7&a,;)y""9g"
    'sSafeProtocols(89)  = "SafeAuthenticate|{8125919B-9BE9-4213-A1D6-75188A22D21E}
    sSafeProtocols(89) = "IKW\aj^KSd6YTW6Zrby}(}*(Y7uz0;Y{%(Q(#(r2,u(,Q-""$~(d"""";?"
    'sSafeProtocols(90)  = "sds|{79E0F14C-9C52-4218-89A7-7C4B0563D121}
    sSafeProtocols(90) = "iNds=,#(|<Q$4#Y8+wn""(y)$X.+xy-c$3&U+)+r~'g"
    'sSafeProtocols(91)  = "siteadvisor|{3A5DC592-7723-4EAA-9EE6-AF4222BCF879}
    sSafeProtocols(91) = "iSe\#Y`Lae4ll)a*:*v'(u(.R(uu37a{*;e+#()""(z#9c;""x's"
    'sSafeProtocols(92)  = "smscrd|{FA3F5003-93D4-11D2-8E48-00A0C98BD8C3}
    sSafeProtocols(92) = "iWdZ4Yf^47S6&&P(#~t2*u""(d'uy3*X{!&a%9~y0:""4*?"
    'sSafeProtocols(93)  = "stibo|{FFAD3420-6D61-44F6-BA25-293F17152D79}
    sSafeProtocols(93) = "i^ZY1qe)47d#%(P"",+w}#|%=V"",$~+M""*)f&-vv~:!*t"
    'sSafeProtocols(94)  = "textwareilluminatorbase|{CE5CD329-1650-414A-8DB0-4CBF72FAED87}
    sSafeProtocols(94) = "jOik9V\HWb.e^_0VjVUPW]Vs=8/v1:S""*#Q++un""'|2$X9,qy*c27-R;7,'&-g"
    'sSafeProtocols(95)  = "widimg|{EE7C2AFF-5742-44FF-BD0E-E521B0D3C3BA}
    sSafeProtocols(95) = "mSU`/\f^3;W3#7f;#zx""(u%+f;u%2&e{6+R&8u'!9{38?"
    'sSafeProtocols(96)  = "wlmailhtml|{03C514A3-1EFB-4856-9F99-10D7BE1653C0}
    sSafeProtocols(96) = "mV^X+aRW[b>k!)c*'y$!#y6=b""|y#,M)7/Y""'u'%8/""-U(-qk"
    'sSafeProtocols(97)  = "x-atng|{7E8717B0-D862-11D5-8C9E-00010304F989}v
    sSafeProtocols(97) = "nuRk0\f^%;X'""-b%#+y$(u""(d*uy1/e{!&P&&xq""<#)0?"
    'sSafeProtocols(98)  = "x-excid|{9D6CC632-1337-4A33-9214-2DA092E776F4}
    sSafeProtocols(98) = "nuVo%^N_i/d&49V((rr!)!|+a({n'(Q$|(d6&~s3-!'=Tr"
    'sSafeProtocols(99)  = "x-mem1|{C3719F83-7EF8-4BA0-89B0-3360C7AFB7CC}
    sSafeProtocols(99) = "nu^\/&f^1)W!*<X(#|(4.u%9a%uy'8P{$)V%9|$48!4:?"
    'sSafeProtocols(100)  = "x-mem3|{4F6D06DD-44AB-4F89-BF13-9027B505B15A}
    sSafeProtocols(100) = "nu^\/(f^""<V4!,d9#yu/8u%=X.u%4'S{*&R,8zq#8y&8?"
    'sSafeProtocols(101)  = "ct|{774E529C-2458-48A2-8F57-3ED3105D8612}
    sSafeProtocols(101) = "Y^mrW,|(#(Y3|(T*.ru&7z|/f*!n!;d#""&U9.{r~s"
    'sSafeProtocols(102)  = "cw|{774E529C-2458-48A2-8F57-3ED3105D8612}
    sSafeProtocols(102) = "YamrW,|(#(Y3|(T*.ru&7z|/f*!n!;d#""&U9.{r~s"
    'sSafeProtocols(103)  = "eti|{3AAE7392-E7AA-11D2-969E-00105A088846}
    sSafeProtocols(103) = "[^Zs=(+$3-S)##e,7(n}'.#$Y+#(y&P!!+a%.}y"",g"
    'sSafeProtocols(104) = "livecall|{828030A1-22C1-4009-854F-8E305202313F}"
    sSafeProtocols(104) = "bSg\%VVOjqX"")&S%7vn~(-""$T%xzy.U$7#X:)uv~&z$(S;g"
    
    ReDim sSafeFilters(24) '(O18)
    sSafeFilters(0) = "application/octet-stream|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
    sSafeFilters(1) = "application/x-complus|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
    sSafeFilters(2) = "application/x-msdownload|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
    sSafeFilters(3) = "Class Install Handler|{32B533BB-EDAE-11d0-BD5A-00AA00B92AF1}"
    sSafeFilters(4) = "deflate|{8f6b0360-b80d-11d0-a9b3-006097942311}"
    sSafeFilters(5) = "gzip|{8f6b0360-b80d-11d0-a9b3-006097942311}"
    sSafeFilters(6) = "lzdhtml|{8f6b0360-b80d-11d0-a9b3-006097942311}"
    sSafeFilters(7) = "text/webviewhtml|{733AC4CB-F1A4-11d0-B951-00A0C90312E1}"
    sSafeFilters(8) = "text/xml|{807553E5-5146-11D5-A672-00B0D022E945}"
    sSafeFilters(9) = "application/x-icq|{db40c160-09a1-11d3-baf2-000000000000}"
    'added in HJT 1.99.2 final
    sSafeFilters(10) = "application/msword|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
    sSafeFilters(11) = "application/vnd.ms-excel|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
    sSafeFilters(12) = "application/vnd.ms-powerpoint|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
    sSafeFilters(13) = "application/x-microsoft-rpmsg-message|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
    sSafeFilters(14) = "application/vnd-backup-octet-stream|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
    sSafeFilters(15) = "application/vnd-viewer|{CD4527E8-4FC7-48DB-9806-10537B501237}"
    sSafeFilters(16) = "application/x-bt2|{6E1DDCE8-76BC-4390-9488-806E8FB1AD77}"
    sSafeFilters(17) = "application/x-internet-signup|{A173B69A-1F9B-4823-9FDA-412F641E65D6}"
    sSafeFilters(18) = "text/html|{8D42AD12-D7A1-4797-BCB7-AD89E5FCE4F7}"
    sSafeFilters(19) = "text/html|{F79B2338-A6E7-46D4-9201-422AA6E74F43}"
    sSafeFilters(20) = "text/x-mrml|{C51721BE-858B-4A66-A8BF-D2882FF49820}"
    sSafeFilters(21) = "text/xml|{807563E5-5146-11D5-A672-00B0D022E945}"
    sSafeFilters(22) = "application/octet-stream|{F969FE8E-1937-45AD-AF42-8A4D11CBDC2A}"
    sSafeFilters(23) = "application/xhtml+xml|{32F66A26-7614-11D4-BD11-00104BD3F987}"
    sSafeFilters(24) = "text/xml|{32F66A26-7614-11D4-BD11-00104BD3F987}"

    'LOAD APPINIT_DLLS SAFELIST (O20)
    sSafeAppInit = "*aakah.dll*akdllnt.dll*ROUSRNT.DLL*ssohook*KATRACK.DLL*APITRAP.DLL*UmxSbxExw.dll*sockspy.dll*scorillont.dll*wbsys.dll*NVDESK32.DLL*hplun.dll*mfaphook.dll*PAVWAIT.DLL*OCMAPIHK.DLL*MsgPlusLoader.dll*IconCodecService.dll*wl_hook.dll*Google\GOOGLE~1\GOEC62~1.DLL*adialhk.dll*wmfhotfix.dll*interceptor.dll*qaphooks.dll*RMProcessLink.dll*msgrmate.dll*wxvault.dll*ctu33.dll*ati2evxx.dll*vsmvhk.dll*"
    
    'LOAD SSODL SAFELIST (O21)
    ReDim sSafeSSODL(11)
    sSafeSSODL(0) = "{E6FB5E20-DE35-11CF-9C87-00AA005127ED}"  'WebCheck: E:\WINDOWS\System32\webcheck.dll (WinAll)
    sSafeSSODL(1) = "{35CEC8A3-2BE6-11D2-8773-92E220524153}"  'SysTray: E:\WINDOWS\System32\stobject.dll (Win2k/XP)
    sSafeSSODL(2) = "{7849596a-48ea-486e-8937-a2a3009f31a9}"  'PostBootReminder: E:\WINDOWS\system32\SHELL32.dll (WinXP)
    sSafeSSODL(3) = "{fbeb8a05-beee-4442-804e-409d6c4515e9}"  'CDBurn: E:\WINDOWS\system32\SHELL32.dll (WinXP)
    sSafeSSODL(4) = "{11566B38-955B-4549-930F-7B7482668782}"  'AUHook: C:\WINDOWS\SYSTEM\AUHOOK.DLL (WinME)
    sSafeSSODL(5) = "{7007ACCF-3202-11D1-AAD2-00805FC1270E}"  'Network.ConnectionTray: C:\WINNT\system32\NETSHELL.dll (Win2k)
    sSafeSSODL(6) = "{e57ce738-33e8-4c51-8354-bb4de9d215d1}"  'UPnPMonitor: C:\WINDOWS\SYSTEM\UPNPUI.DLL (WinME/XP)
    sSafeSSODL(7) = "{BCBCD383-3E06-11D3-91A9-00C04F68105C}"  'AUHook: C:\WINDOWS\SYSTEM\AUHOOK.DLL (WinME)
    sSafeSSODL(8) = "{F5DF91F9-15E9-416B-A7C3-7519B11ECBFC}"  '0aMCPClient: C:\Program Files\StarDock\MCPCore.dll
    sSafeSSODL(9) = "{AAA288BA-9A4C-45B0-95D7-94D524869DB5}"  'WPDShServiceObj   WPDShServiceObj.dll Windows Portable Device Shell Service Object
    sSafeSSODL(10) = "{1799460C-0BC8-4865-B9DF-4A36CD703FF0}" 'IconPackager Repair  iprepair.dll    Stardock\Object Desktop\ ThemeManager
    sSafeSSODL(11) = "{6D972050-A934-44D7-AC67-7C9E0B264220}" 'EnhancedDialog   enhdlginit.dll  EnhancedDialog by Stardock
    
    'LOAD WINLOGON NOTIFY SAFELIST (O20)
    'second line added in HJT 1.99.2 final
    sSafeWinlogonNotify = "crypt32chain*cryptnet*cscdll*ScCertProp*Schedule*SensLogn*termsrv*wlballoon*igfxcui*AtiExtEvent*wzcnotif*" & _
                          "ActiveSync*atmgrtok*avldr*Caveo*ckpNotify*Command AntiVirus Download*ComPlusSetup*CwWLEvent*dimsntfy*DPWLN*EFS*FolderGuard*GoToMyPC*IfxWlxEN*igfxcui*IntelWireless*klogon*LBTServ*LBTWlgn*LMIinit*loginkey*MCPClient*MetaFrame*NavLogon*NetIdentity Notification*nwprovau*OdysseyClient*OPXPGina*PCANotify*pcsinst*PFW*PixVue*ppeclt*PRISMAPI.DLL*PRISMGNA.DLL*psfus*QConGina*RAinit*RegCompact*SABWinLogon*SDNotify*Sebring*STOPzilla*sunotify*SymcEventMonitors*T3Notify*TabBtnWL*Timbuktu Pro*tpfnf2*tpgwlnotify*tphotkey*VESWinlogon*WB*WBSrv*WgaLogon*wintask*WLogon*WRNotifier*Zboard*zsnotify*sclgntfy"
    
    Exit Sub
    
Error:
    ErrorMsg "modMain_LoadStuff", Err.Number, Err.Description
End Sub

Public Sub StartScan()
    On Error GoTo Error:
    frmMain.shpBackground.Tag = CStr(iItems)
    frmMain.shpProgress.Tag = "0"
    frmMain.txtNothing.Visible = False
    
    Dim sRule$, i%
    'load ignore list
    IsOnIgnoreList ""
    
    frmMain.lstResults.Clear
    
    'Registry
    
    'decrypt nonstandard safelist domains
    For i = 0 To UBound(sSafeRegDomains)
        If sSafeRegDomains(i) = vbNullString Then Exit For
        sSafeRegDomains(i) = Crypt(sSafeRegDomains(i), sProgramVersion)
    Next i
    
    'i = 0
    For i = 0 To UBound(sRegVals)
        'If sRegVals(i) = "" Then Exit For
        ProcessRuleReg sRegVals(i)
        UpdateProgressBar
        'i = i + 1
    Next i
    
    're-encrypt nonstandard domains safelist
    For i = 0 To UBound(sSafeRegDomains)
        If sSafeRegDomains(i) = vbNullString Then Exit For
        sSafeRegDomains(i) = Crypt(sSafeRegDomains(i), sProgramVersion, True)
    Next i
    
    CheckRegistry3Item
    UpdateProgressBar
        
    'File
    'i = 0
    For i = 0 To UBound(sFileVals)
        If sFileVals(i) = "" Then Exit For
        ProcessRuleIniFile sFileVals(i)
        UpdateProgressBar
        'i = i + 1
    Next i
        
    'Netscape/Mozilla stuff
    CheckNetscapeMozilla        'N1-4
    UpdateProgressBar
    
    'Other options
    CheckOther1Item
    UpdateProgressBar
    CheckOther2Item
    UpdateProgressBar
    CheckOther3Item
    UpdateProgressBar
    CheckOther4Item
    UpdateProgressBar
    CheckOther5Item
    UpdateProgressBar
    CheckOther6Item
    UpdateProgressBar
    CheckOther7Item
    UpdateProgressBar
    CheckOther8Item
    UpdateProgressBar
    CheckOther9Item
    UpdateProgressBar
    CheckOther10Item
    UpdateProgressBar
    CheckOther11Item
    UpdateProgressBar
    CheckOther12Item
    UpdateProgressBar
    CheckOther13Item
    UpdateProgressBar
    CheckOther14Item
    UpdateProgressBar
    CheckOther15Item
    UpdateProgressBar
    CheckOther16Item
    UpdateProgressBar
    CheckOther17Item
    UpdateProgressBar
    CheckOther18Item
    UpdateProgressBar
    CheckOther19Item
    UpdateProgressBar
    CheckOther20Item
    UpdateProgressBar
    CheckOther21Item
    UpdateProgressBar
    CheckOther22Item
    UpdateProgressBar
    CheckOther23Item
    UpdateProgressBar
    'added in HJT 1.99.2: Desktop Components
    CheckOther24Item
    UpdateProgressBar
    
   
    With frmMain
        .shpBackground.Visible = False
        .shpProgress.Visible = False
        .lblMD5.Visible = False
        .lblInfo(1).Visible = True
        .picPaypal.Visible = True
        If .lstResults.ListCount > 0 Then
            If bAutoSelect Then
                For i = 0 To .lstResults.ListCount - 1
                    .lstResults.Selected(i) = True
                Next i
            End If
            .txtNothing.Visible = False
            .cmdFix.Enabled = True
            .cmdSaveDef.Enabled = True
        Else
            .txtNothing.Visible = True
            .cmdFix.Enabled = False
            .cmdSaveDef.Enabled = False
        End If
    End With
    Exit Sub
    
Error:
    ErrorMsg "modMain_StartScan", Err.Number, Err.Description
End Sub

Public Sub UpdateProgressBar()
    Dim nFract As Double
    On Error GoTo Error:
    With frmMain
        If Not IsNumeric(.shpProgress.Tag) Then .shpProgress.Tag = "0"
        If Not IsNumeric(.shpBackground.Tag) Then .shpBackground.Tag = "1"
        'nFract = 0.05
        nFract = CDbl(.shpProgress.Tag) / CDbl(.shpBackground.Tag)
        If nFract > 1 Then nFract = 1
        .shpProgress.Width = .shpBackground.Width * nFract
        .shpProgress.Tag = CStr(CInt(.shpProgress.Tag) + 1)
        
        
        'frmMain.lblStatus.Caption = .shpProgress.Tag
        Select Case Val(.shpProgress.Tag)
            Case 1 To 103: frmMain.lblStatus.Caption = Translate(230) & "... (" & Val(.shpProgress.Tag) - 0 & "/" & UBound(sRegVals) + 1 & ")"
            Case 104 To 111: frmMain.lblStatus.Caption = Translate(231) & "... (" & Val(.shpProgress.Tag) - 103 & "/" & UBound(sFileVals) + 1 & ")"
            Case 112: frmMain.lblStatus.Caption = Translate(232) & "..."
            Case 113: frmMain.lblStatus.Caption = Translate(233) & "..."
            Case 114: frmMain.lblStatus.Caption = Translate(234) & "..."
            Case 115: frmMain.lblStatus.Caption = Translate(235) & "..."
            Case 116: frmMain.lblStatus.Caption = Translate(236) & "..."
            Case 117: frmMain.lblStatus.Caption = Translate(237) & "..."
            Case 118: frmMain.lblStatus.Caption = Translate(238) & "..."
            Case 119: frmMain.lblStatus.Caption = Translate(239) & "..."
            Case 120: frmMain.lblStatus.Caption = Translate(240) & "..."
            Case 121: frmMain.lblStatus.Caption = Translate(241) & "..."
            Case 122: frmMain.lblStatus.Caption = Translate(242) & "..."
            Case 123: frmMain.lblStatus.Caption = Translate(243) & "..."
            Case 124: frmMain.lblStatus.Caption = Translate(244) & "..."
            Case 125: frmMain.lblStatus.Caption = Translate(245) & "..."
            Case 126: frmMain.lblStatus.Caption = Translate(246) & "..."
            Case 127: frmMain.lblStatus.Caption = Translate(247) & "..."
            Case 128: frmMain.lblStatus.Caption = Translate(248) & "..."
            Case 129: frmMain.lblStatus.Caption = Translate(249) & "..."
            Case 130: frmMain.lblStatus.Caption = Translate(250) & "..."
            Case 131: frmMain.lblStatus.Caption = Translate(251) & "..."
            Case 132: frmMain.lblStatus.Caption = Translate(252) & "..."
            Case 133: frmMain.lblStatus.Caption = Translate(253) & "..."
            Case 134: frmMain.lblStatus.Caption = Translate(254) & "..."
            Case 135: frmMain.lblStatus.Caption = Translate(255) & "..."
            Case 136: frmMain.lblStatus.Caption = Translate(257) & "..."
            
            Case Else: frmMain.lblStatus.Caption = Translate(256)
'            Case 1 To 103: frmMain.lblStatus.Caption = "IE Registry values... (" & Val(.shpProgress.Tag) - 0 & "/" & UBound(sRegVals) + 1 & ")"
'            Case 104 To 111: frmMain.lblStatus.Caption = "INI file values... (" & Val(.shpProgress.Tag) - 103 & "/" & UBound(sFileVals) + 1 & ")"
'            Case 112: frmMain.lblStatus.Caption = "Netscape/Mozilla homepage..."
'            Case 113: frmMain.lblStatus.Caption = "O1 - Hosts file redirection..."
'            Case 114: frmMain.lblStatus.Caption = "O2 - BHO enumeration..."
'            Case 115: frmMain.lblStatus.Caption = "O3 - Toolbar enumeration..."
'            Case 116: frmMain.lblStatus.Caption = "O4 - Registry && Start Menu autoruns..."
'            Case 117: frmMain.lblStatus.Caption = "O5 - Control Panel: IE Options..."
'            Case 118: frmMain.lblStatus.Caption = "O6 - Policies: IE Options menuitem..."
'            Case 119: frmMain.lblStatus.Caption = "O7 - Policies: Regedit..."
'            Case 120: frmMain.lblStatus.Caption = "O8 - IE contextmenu items enumeration..."
'            Case 121: frmMain.lblStatus.Caption = "O9 - IE 'Tools' menuitems and buttons enumeration..."
'            Case 122: frmMain.lblStatus.Caption = "O10 - Winsock LSP hijackers..."
'            Case 123: frmMain.lblStatus.Caption = "O11 - Extra options groups in IE Advanced Options..."
'            Case 124: frmMain.lblStatus.Caption = "O12 - IE plugins enumeration..."
'            Case 125: frmMain.lblStatus.Caption = "O13 - DefaultPrefix hijack..."
'            Case 126: frmMain.lblStatus.Caption = "O14 - IERESET.INF hijack..."
'            Case 127: frmMain.lblStatus.Caption = "O15 - Trusted Zone enumeration..."
'            Case 128: frmMain.lblStatus.Caption = "O16 - DPF object enumeration..."
'            Case 129: frmMain.lblStatus.Caption = "O17 - DNS && DNS Suffix settings..."
'            Case 130: frmMain.lblStatus.Caption = "O18 - Protocol && Filter enumeration..."
'            Case 131: frmMain.lblStatus.Caption = "O19 - User stylesheet hijack..."
'            Case 132: frmMain.lblStatus.Caption = "O20 - AppInit_DLLs Registry value..."
'            Case 133: frmMain.lblStatus.Caption = "O21 - ShellServiceObjectDelayLoad Registry key..."
'            Case 134: frmMain.lblStatus.Caption = "O22 - SharedTaskScheduler Registry key..."
'            Case 135: frmMain.lblStatus.Caption = "O23 - NT Services..."
'            Case 136: frmMain.lblStatus.Caption = "O24 - ActiveX Desktop Components..."
'
'            Case Else: frmMain.lblStatus.Caption = "Scan completed!"
        End Select
        'Sleep 100
    End With
    DoEvents
    Exit Sub
    
Error:
    ErrorMsg "modMain_UpdateProgressBar", Err.Number, Err.Description, "shpProgress.Tag=" & frmMain.shpProgress.Tag & ",shpBackground.Tag=" & frmMain.shpBackground.Tag
End Sub

Private Sub ProcessRuleReg(ByVal sRule$)
    Dim vRule As Variant, iMode%, i%, bIsNSBSD As Boolean
    Dim sValue$, lHive&, sHit$
    On Error GoTo Error:
    If sRule = vbNullString Then Exit Sub
    
    'decrypt rule
    sRule = Crypt(sRule, sProgramVersion)
    
    If Right(sRule, 1) = Chr(0) Then sRule = Left(sRule, Len(sRule) - 1)
    'Registry rule syntax:
    '[regkey],[regvalue],[infected data],[default data]
    '* [regkey]           = "" -> abort - no way man!
    ' * [regvalue]        = "" -> delete entire key
    '  * [default data]   = "" -> delete value
    '   * [infected data] = "" -> any value (other than default) is considered infected
    vRule = Split(sRule, ",")
    If UBound(vRule) <> 3 Or _
       Left(CStr(vRule(0)), 2) <> "HK" Then
        'decryption failed or spelling error
        Exit Sub
    End If
    
    ' iMode = 0 -> check if value is infected
    ' iMode = 1 -> check if value is present
    ' iMode = 2 -> check if regkey is present
    If CStr(vRule(0)) = "" Then Exit Sub
    If CStr(vRule(3)) = "" Then iMode = 0
    If CStr(vRule(2)) = "" Then iMode = 1
    If CStr(vRule(1)) = "" Then iMode = 2
    
    Select Case Left(CStr(vRule(0)), 4)
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case Else: Exit Sub
    End Select
    vRule(0) = Mid(CStr(vRule(0)), 6)
    If CStr(vRule(1)) = "(Default)" Then vRule(1) = ""
    
    Select Case iMode
        Case 0 'check for incorrect value
            sValue = RegGetString(lHive, CStr(vRule(0)), CStr(vRule(1)))
            If InStr(1, sValue, "%SYSTEMROOT%", vbTextCompare) Then
                sValue = Replace(sValue, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
                sValue = LCase(sValue)
                vRule(2) = LCase(CStr(vRule(2)))
            End If
            
            'use instr instead of = to prevent stupid VB errs
            If InStr(1, sValue, CStr(vRule(2)), vbTextCompare) <> 1 Then
                bIsNSBSD = False
                For i = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sValue, sSafeRegDomains(i), vbTextCompare) = 1 _
                       And sSafeRegDomains(i) <> vbNullString Then
                        bIsNSBSD = True
                        Exit For
                    End If
                Next i
                If bIgnoreSafe = False Then bIsNSBSD = False
                If Not bIsNSBSD Then
                    If InStr(1, sValue, "%2e", vbTextCompare) > 0 Then sValue = Unescape(sValue)
                    sHit = "R0 - " & Left(sRule, InStr(sRule, ",") - 1) & "," & CStr(vRule(1)) & " = " & sValue
                    If IsOnIgnoreList(sHit) Then Exit Sub
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        Case 1  'check for present value
            sValue = RegGetString(lHive, CStr(vRule(0)), CStr(vRule(1)))
            If sValue <> vbNullString Then
                'check if domain is on safe list
                bIsNSBSD = False
                For i = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sValue, sSafeRegDomains(i), vbTextCompare) = 1 _
                       And sSafeRegDomains(i) <> vbNullString Then
                        bIsNSBSD = True
                        Exit For
                    End If
                Next i
                If bIgnoreSafe = False Then bIsNSBSD = False
                'make hit
                If Not bIsNSBSD Then
                    If InStr(1, sValue, "%2e", vbTextCompare) > 0 Then sValue = Unescape(sValue)
                    sHit = "R1 - " & Left(sRule, InStr(sRule, ",") - 1) & "," & IIf(CStr(vRule(1)) = "", "(Default)", CStr(vRule(1))) & IIf(sValue <> vbNullString, " = " & sValue, "")
                    If IsOnIgnoreList(sHit) Then Exit Sub
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        Case 2
            If RegKeyExists(lHive, CStr(vRule(0))) Then
                sHit = "R2 - " & Left(sRule, InStr(sRule, ",") - 1)
                If IsOnIgnoreList(sHit) Then Exit Sub
                frmMain.lstResults.AddItem sHit
            End If
        Case Else: Exit Sub
    End Select
    Exit Sub
    
Error:
    ErrorMsg "modMain_ProcessRuleReg", Err.Number, Err.Description, "sRule=" & sRule
End Sub

Private Sub ProcessRuleIniFile(ByVal sRule$)
    Dim vRule As Variant, iMode%, sValue$, sHit$
    On Error GoTo Error:
    'IniFile rule syntax:
    '[inifile],[section],[value],[default data],[infected data]
    '* [inifile]          = "" -> abort
    ' * [section]         = "" -> abort
    '  * [value]          = "" -> abort
    '   * [default data]  = "" -> delete if found
    '    * [infected data]= "" -> fix if infected
    
    'decrypt rule
    'sRule = Crypt(sRule, sProgramVersion)
    
    If Right(sRule, 1) = Chr(0) Then sRule = Left(sRule, Len(sRule) - 1)
    vRule = Split(sRule, ",")
    If UBound(vRule) <> 4 Or _
       InStr(CStr(vRule(0)), ".ini") = 0 Then
        'spelling error or decrypting error
        Exit Sub
    End If
    If CStr(vRule(0)) = "" Then Exit Sub
    If CStr(vRule(1)) = "" Then Exit Sub
    If CStr(vRule(2)) = "" Then Exit Sub
    If CStr(vRule(4)) = "" Then iMode = 0
    If CStr(vRule(3)) = "" Then iMode = 1
    
    If InStr(CStr(vRule(3)), "UserInit") > 0 Then vRule(3) = CStr(vRule(3)) & ","
    
    If Left(CStr(vRule(0)), 3) = "REG" Then
        If Not bIsWinNT Then Exit Sub
        
        If CStr(vRule(4)) = "" Then iMode = 2
        If CStr(vRule(3)) = "" Then iMode = 3
    End If
    
    'iMode:
    ' 0 = check if value is infected
    ' 1 = check if value is present
    ' 2 = check if value is infected, in the Registry
    ' 3 = check if value is present, in the Registry
    
    Select Case iMode
        Case 0
            'sValue = String(255, " ")
            'GetPrivateProfileString CStr(vRule(1)), CStr(vRule(2)), "", sValue, 255, CStr(vRule(0))
            'sValue = RTrim(sValue)
            sValue = IniGetString(CStr(vRule(0)), CStr(vRule(1)), CStr(vRule(2)))
            If Right(sValue, 1) = Chr(0) Then sValue = Left(sValue, Len(sValue) - 1)
            'If RightB(sValue, 2) = Chr(0) Then sValue = LeftB(sValue, LenB(sValue) - 2)
            If Trim(LCase(sValue)) <> LCase(CStr(vRule(3))) Then
                If bIsWinNT And Trim(LCase(sValue)) <> vbNullString Then
                    sHit = "F0 - " & CStr(vRule(0)) & ": " & CStr(vRule(2)) & "=" & sValue
                    If IsOnIgnoreList(sHit) Then Exit Sub
                    If bMD5 Then sHit = sHit & GetFileFromAutostart(sValue)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        Case 1
            'sValue = String(255, " ")
            'GetPrivateProfileString CStr(vRule(1)), CStr(vRule(2)), "", sValue, 255, CStr(vRule(0))
            'sValue = RTrim(sValue)
            sValue = IniGetString(CStr(vRule(0)), CStr(vRule(1)), CStr(vRule(2)))
            If Right(sValue, 1) = Chr(0) Then sValue = Left(sValue, Len(sValue) - 1)
            'If RightB(sValue, 2) = Chr(0) Then sValue = LeftB(sValue, LenB(sValue) - 2)
            If Trim(sValue) <> vbNullString Then
                sHit = "F1 - " & CStr(vRule(0)) & ": " & CStr(vRule(2)) & "=" & sValue
                If IsOnIgnoreList(sHit) Then Exit Sub
                If bMD5 Then sHit = sHit & GetFileFromAutostart(sValue)
                frmMain.lstResults.AddItem sHit
            End If
        Case 2
            'so far F2 is only reg:Shell and reg:UserInit
            sValue = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", CStr(vRule(2)))
            If LCase(sValue) <> LCase(CStr(vRule(3))) Then
                sHit = "F2 - " & CStr(vRule(0)) & ": " & CStr(vRule(2)) & "=" & sValue
                If IsOnIgnoreList(sHit) Then Exit Sub
                If bMD5 Then sHit = sHit & GetFileFromAutostart(sValue)
                frmMain.lstResults.AddItem sHit
            End If
        Case 3
            'this is not really smart when more INIFile items get
            'added, but so far F3 is only reg:load and reg:run
            sValue = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", CStr(vRule(2)))
            If sValue <> vbNullString Then
                sHit = "F3 - " & CStr(vRule(0)) & ": " & CStr(vRule(2)) & "=" & sValue
                If IsOnIgnoreList(sHit) Then Exit Sub
                If bMD5 Then sHit = sHit & GetFileFromAutostart(sValue)
                frmMain.lstResults.AddItem sHit
            End If
    End Select
    Exit Sub
    
Error:
    ErrorMsg "modMain_ProcessRuleIniFile", Err.Number, Err.Description, "sRule=" & sRule
End Sub

Public Sub GetHostsAndWinDir()
    Dim uOVI As OSVERSIONINFO, sDatabasePath$
    On Error GoTo Error:
    sWinDir = String(255, 0)
    GetWindowsDirectory sWinDir, 255
    sWinDir = Left(sWinDir, InStr(sWinDir, Chr(0)) - 1)
    If Right(sWinDir, 1) = "\" Then sWinDir = Left(sWinDir, Len(sWinDir) - 1)
    sWinSysDir = sWinDir & "\" & IIf(bIsWinNT, "system32", "system")
    
    uOVI.dwOSVersionInfoSize = Len(uOVI)
    GetVersionEx uOVI
    Select Case uOVI.dwPlatformId
        Case VER_PLATFORM_WIN32s: End
        Case VER_PLATFORM_WIN32_WINDOWS
            sHostsFile = sWinDir & "\hosts"
            sWinVersion = "Windows 9x "
            'frmMain.chkProcManShowDLLs.Visible = False
            lEnumBufSize = 260
        Case VER_PLATFORM_WIN32_NT
            sHostsFile = sWinDir & "\system32\drivers\etc\hosts"
            sWinVersion = "Windows NT "
            bIsWinNT = True
            lEnumBufSize = 16400
    End Select
    
    With uOVI
        sWinVersion = sWinVersion & _
            CStr(.dwMajorVersion) & "." & _
            String(2 - Len(CStr(.dwMinorVersion)), "0") & _
            CStr(.dwMinorVersion) & "." & _
            CStr(.dwBuildNumber And &HFFF)
        If Not bIsWinNT And _
           .dwMajorVersion = 4 And _
           .dwMinorVersion = 90 Then bIsWinME = True
    End With
    sMSIEVersion = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "Version")
    If sMSIEVersion = vbNullString Then sMSIEVersion = "Unknown"
    sWinSysDir = sWinDir & IIf(bIsWinNT, "\SYSTEM32", "\SYSTEM")
    
    If bIsWinNT Then
        sDatabasePath = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DataBasePath")
        '%systemroot% may be in path - replace it
        sDatabasePath = Replace(sDatabasePath, "%SystemRoot%", sWinDir, , , vbTextCompare)
        sHostsFile = sDatabasePath & "\hosts"
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_GetHostsAndWinDir", Err.Number, Err.Description
End Sub

Private Sub CheckOther1Item()
    Dim sLine$, sFile$, sHit$, sDomains$(), i%
    Dim iAttr%
    On Error GoTo Error:
    
    If Not FileExists(sHostsFile) Then Exit Sub
    If FileLen(sHostsFile) = 0 Then Exit Sub
    
    On Error Resume Next
    iAttr = GetAttr(sHostsFile)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetAttr sHostsFile, vbNormal
    SetAttr sHostsFile, vbArchive
    If Err Then
        MsgBox Replace(Translate(300), "[]", sHostsFile), vbExclamation
'        MsgBox "For some reason your system denied write " & _
'        "access to the Hosts file." & vbCrLf & "If any hijacked domains " & _
'        "are in this file, HijackThis may NOT be able to " & _
'        "fix this." & vbCrLf & vbCrLf & "If that happens, you need " & _
'        "to edit the file yourself. To do this, click " & _
'        "Start, Run and type:" & vbCrLf & vbCrLf & _
'        "   notepad """ & sHostsFile & """" & vbCrLf & vbCrLf & _
'        "and press Enter. Find the line(s) HijackThis " & _
'        "reports and delete them. Save the file as " & _
'        """hosts."" (with quotes), and reboot.", vbExclamation
    End If
    SetAttr sHostsFile, iAttr
    On Error GoTo Error:
    
    If LCase(sHostsFile) <> LCase(sWinDir & "\hosts") And _
       LCase(sHostsFile) <> LCase(sWinSysDir & "\drivers\etc\hosts") Then
        sHit = "O1 - Hosts file is located at: " & sHostsFile
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    Open sHostsFile For Input As #1
        Do
            Line Input #1, sLine
            If InStr(sLine, Chr(10)) > 0 Then
                'hosts file has line delimiters
                'that confuse Line Input - so
                'convert them to vbCrLf :)
                Close #1
                If Not bTriedFixUnixHostsFile Then
                    FixUNIXHostsFile
                    bTriedFixUnixHostsFile = True
                    CheckOther1Item
                Else
                    MsgBox Translate(301), vbExclamation
'                    MsgBox "Your hosts file has invalid linebreaks and " & _
'                           "HijackThis is unable to fix this. O1 items will " & _
'                           "not be displayed." & vbCrLf & vbCrLf & _
'                           "Click OK to continue the rest of the scan.", vbExclamation
                End If
                Exit Sub
            End If
            
            'ignore all lines that start with loopback
            '(127.0.0.1), null (0.0.0.0) and private IPs
            '(192.168. / 10.)
            sLine = Replace(sLine, vbTab, " ")
            sLine = Trim(sLine)
            If sLine <> vbNullString Then
                If InStr(sLine, "127.0.0.1") <> 1 And _
                   InStr(sLine, "0.0.0.0") <> 1 And _
                   InStr(sLine, "192.168.") <> 1 And _
                   InStr(sLine, "10.") <> 1 And _
                   InStr(sLine, "#") <> 1 And _
                   Not (bIgnoreSafe And InStr(sLine, "216.239.37.101") > 0) Or _
                   bIgnoreAllWhitelists Then
                    '216.239.37.101 = google.com
                    Do
                        sLine = Replace(sLine, "  ", " ")
                    Loop Until InStr(sLine, "  ") = 0
                    
                    sHit = "O1 - Hosts: " & sLine
                    If Not IsOnIgnoreList(sHit) Then
                        frmMain.lstResults.AddItem sHit
                        i = i + 1
                    End If
                    
                    If i > 100 Then
                        MsgBox Replace(Translate(302), "[]", sHostsFile), vbExclamation
'                        MsgBox "You have an particularly large " & _
'                        "amount of hijacked domains. It's probably " & _
'                        "better to delete the file itself then to " & _
'                        "fix each item (and create a backup)." & vbCrLf & _
'                        vbCrLf & "If you see the same IP address in all " & _
'                        "the reported O1 items, consider deleting your " & _
'                        "Hosts file, which is located at " & sHostsFile & _
'                        ".", vbExclamation
                        Close #1
                        Exit Sub
                    End If
                End If
            End If
        Loop Until EOF(1)
    Close #1
    On Error Resume Next
    SetAttr sHostsFile, iAttr
    Exit Sub
    
Error:
    Close #1
    ErrorMsg "modMain_CheckOther1Item", Err.Number, Err.Description
End Sub

Private Sub CheckOther5Item()
    Dim sControlIni$, sDummy$, sHit$
    On Error GoTo Error:
    
    sControlIni = String(255, 0)
    GetWindowsDirectory sControlIni, 255
    sControlIni = Left(sControlIni, InStr(sControlIni, Chr(0)) - 1) & "\control.ini"
    If sControlIni = "\control.ini" Then Exit Sub
    If Dir(sControlIni) = vbNullString Then Exit Sub
    
    sDummy = String(5, " ")
    'GetPrivateProfileString "don't load", "inetcpl.cpl", "", sDummy, 5, sControlIni
    IniGetString sControlIni, "don't load", "inetcpl.cpl"
    sDummy = RTrim(sDummy)
    If Right(sDummy, 1) = Chr(0) Then sDummy = Left(sDummy, Len(sDummy) - 1)
    If sDummy <> vbNullString Then
        sHit = "O5 - control.ini: inetcpl.cpl=" & sDummy
        If IsOnIgnoreList(sHit) Then Exit Sub
        frmMain.lstResults.AddItem sHit
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_CheckOther5Item", Err.Number, Err.Description
End Sub

Private Sub CheckOther6Item()
    'HKEY_CURRENT_USER\ software\ policies\ microsoft\
    'internet explorer. If there are sub folders called
    '"restrictions" and/or "control panel", delete them
    
    Dim sHit$
    On Error GoTo Error:
    If RegKeyExists(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Restrictions") And _
       RegKeyHasValues(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Restrictions") Then
        sHit = "O6 - HKCU\Software\Policies\Microsoft\Internet Explorer\Restrictions present"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    If RegKeyExists(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel") And _
       RegKeyHasValues(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel") Then
        sHit = "O6 - HKCU\Software\Policies\Microsoft\Internet Explorer\Control Panel present"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    If RegKeyExists(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions") And _
       RegKeyHasValues(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions") Then
        sHit = "O6 - HKCU\Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions present"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    If RegKeyExists(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions") And _
       RegKeyHasValues(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions") Then
        sHit = "O6 - HKLM\Software\Policies\Microsoft\Internet Explorer\Restrictions present"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    If RegKeyExists(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Control Panel") And _
       RegKeyHasValues(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Control Panel") Then
        sHit = "O6 - HKLM\Software\Policies\Microsoft\Internet Explorer\Control Panel present"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    If RegKeyExists(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions") And _
       RegKeyHasValues(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions") Then
        sHit = "O6 - HKLM\Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions present"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_CheckOther6Item", Err.Number, Err.Description
End Sub

Private Sub CheckOther7Item()
    Dim lValue&, sHit$
    On Error GoTo Error:
    
    lValue = RegGetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools")
    If lValue = 1 Then
        sHit = "O7 - HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System, DisableRegedit=1"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    lValue = RegGetDword(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools")
    If lValue = 1 Then
        sHit = "O7 - HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System, DisableRegedit=1"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_CheckOther7Item", Err.Number, Err.Description
End Sub

Public Function CmnDlgSaveFile(sTitle$, sFilter$, Optional sDefFile$)
    Dim uOFN As OPENFILENAME, sFile$
    On Error GoTo Error:
    
    sFile = sDefFile & String(256 - Len(sDefFile), 0)
    With uOFN
        .lStructSize = Len(uOFN)
        If InStr(sFilter, "|") > 0 Then sFilter = Replace(sFilter, "|", Chr(0))
        If Right(sFilter, 2) <> Chr(0) & Chr(0) Then sFilter = sFilter & Chr(0) & Chr(0)
        .lpstrFilter = sFilter
        .lpstrFile = sFile
        .lpstrTitle = sTitle
        .nMaxFile = 256
        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_OVERWRITEPROMPT
    End With
    If GetSaveFileName(uOFN) = 0 Then Exit Function
    sFile = Left(uOFN.lpstrFile, InStr(uOFN.lpstrFile, Chr(0)) - 1)
    CmnDlgSaveFile = sFile
    Exit Function
    
Error:
    ErrorMsg "modMain_CmnDlgSaveFile", Err.Number, Err.Description, "sTitle=" & sTitle & ",sFilter=" & sFilter & ",sDefFile=" & sDefFile
End Function

Public Function CmnDlgOpenFile(sTitle$, sFilter$, Optional sDefFile$)
    Dim uOFN As OPENFILENAME, sFile$
    On Error GoTo Error:
    
    sFile = sDefFile & String(256 - Len(sDefFile), 0)
    With uOFN
        .lStructSize = Len(uOFN)
        If InStr(sFilter, "|") > 0 Then sFilter = Replace(sFilter, "|", Chr(0))
        If Right(sFilter, 2) <> Chr(0) & Chr(0) Then sFilter = sFilter & Chr(0) & Chr(0)
        .lpstrFilter = sFilter
        .lpstrFile = sFile
        .lpstrTitle = sTitle
        .nMaxFile = 256
        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_PATHMUSTEXIST
    End With
    If GetOpenFileName(uOFN) = 0 Then Exit Function
    sFile = TrimNull(uOFN.lpstrFile)
    CmnDlgOpenFile = sFile
    Exit Function
    
Error:
    ErrorMsg "modMain_CmnDlgOpenFile", Err.Number, Err.Description, "sTitle=" & sTitle & ", sFilter=" & sFilter & ", sDefFile=" & sDefFile
End Function

Public Sub FixRegItem(sItem$)
    'R0 - HKCU\Software\..\Main,Window Title
    'R1 - HKCU\Software\..\Main,Window Title=MSIE 5.01
    'R2 - HKCU\Software\..\Main
    Dim lHive&, sKey$, sValue$, i%, sFixed$, sDummy$
    On Error GoTo Error:
    
    For i = 0 To UBound(sRegVals)
        If sRegVals(i) = vbNullString Then Exit For
        sRegVals(i) = Crypt(sRegVals(i), sProgramVersion)
    Next i
    
    If InStr(sItem, " (obfuscated)") > 0 Then
        'remove unescape tag, just in case
        sItem = Replace(sItem, " (obfuscated)", vbNullString)
    End If
    Select Case Mid(sItem, 6, 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
    End Select
    
    sKey = Mid(sItem, 11)
    sValue = Mid(sKey, InStr(sKey, ",") + 1)
    sKey = Left(sKey, InStr(sKey, ",") - 1)
    If InStr(sValue, " = ") > 0 Then sValue = Left(sValue, InStr(sValue, " = ") - 1)
    If sValue = "(Default)" Then sValue = ""
    
    If Left(sItem, 2) = "R0" Then
        'restore value
        'find item in reflist
        sDummy = Mid(sItem, 6)
        sDummy = Left(sDummy, InStr(sDummy, " = ") - 1)
        For i = 0 To UBound(sRegVals)
            If InStr(1, sRegVals(i), sDummy, vbTextCompare) Then Exit For
        Next i
        If i = UBound(sRegVals) + 1 Then GoTo CleanUp
        'get fixed data for value
        sFixed = Left(sRegVals(i), Len(sRegVals(i)) - 1)
        sFixed = Mid(sFixed, InStrRev(sFixed, ",") + 1)
        RegSetStringVal lHive, sKey, sValue, sFixed
    ElseIf Left(sItem, 2) = "R1" Then
        'delete value
        RegDelVal lHive, sKey, sValue
    ElseIf Left(sItem, 2) = "R2" Then
        'delete key
        RegDelKey lHive, sKey
    End If
    
CleanUp:
    For i = 0 To UBound(sRegVals)
        If sRegVals(i) = vbNullString Then Exit For
        sRegVals(i) = Crypt(sRegVals(i), sProgramVersion, True)
    Next i
    
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixRegItem", Err.Number, Err.Description, "sItem=" & sItem
    GoTo CleanUp
End Sub

Public Sub FixFileItem(sItem$)
    'F0 - system.ini: Shell=c:\win98\explorer.exe openme.exe
    'F1 - win.ini: load=hpfsch
    On Error GoTo Error:
    'coding is easy if you cheat :)
    
    If Left(sItem, 2) = "F0" Then
        'restore value
        If InStr(sItem, "system.ini: Shell=") > 0 Then
            'WritePrivateProfileString "boot", "shell", "explorer.exe", "system.ini"
            IniSetString "system.ini", "boot", "shell", "explorer.exe"
        End If
    ElseIf Left(sItem, 2) = "F1" Then
        'delete value
        If InStr(sItem, "win.ini: load=") > 0 Then
            'WritePrivateProfileString "windows", "load", "", "win.ini"
            IniSetString "win.ini", "windows", "load", ""
        ElseIf InStr(sItem, "win.ini: run=") > 0 Then
            'WritePrivateProfileString "windows", "run", "", "win.ini"
            IniSetString "win.ini", "windows", "run", ""
        End If
    ElseIf Left(sItem, 2) = "F2" Then
        'restore registry value
        If InStr(sItem, "system.ini: Shell=") > 0 Then
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "Shell", "explorer.exe"
        ElseIf InStr(sItem, "system.ini: UserInit=") > 0 Then
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "UserInit", sWinSysDir & "\Userinit.exe,"
        End If
    ElseIf Left(sItem, 2) = "F3" Then
        'delete registry value
        If InStr(sItem, "win.ini: load=") > 0 Then
            RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "load"
        ElseIf InStr(sItem, "win.ini: run=") > 0 Then
            RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "run"
        End If
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixFileItem", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub FixOther1Item(sItem$)
    'O1 - Hijack of auto.search.msn.com etc with Hosts file
    
    Dim sLine$, sHijacker$, iAttr%
    On Error GoTo Error:
    If Not FileExists(sHostsFile) Then Exit Sub
    If FileLen(sHostsFile) = 0 Then Exit Sub
    iAttr = GetAttr(sHostsFile)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    On Error Resume Next
    SetAttr sHostsFile, vbNormal
    On Error GoTo Error:
    
    'fix stupid 'permission denied' errors on Win2000
    If GetAttr(sHostsFile) <> vbNormal Then
        On Error Resume Next
    End If
    
    If InStr(sItem, "Hosts file is located at") > 0 Then
        'hosts file relocation - always bad
        RegSetExpandStringVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", "%SystemRoot%\System32\drivers\etc"
        Exit Sub
    End If
    
    sHijacker = Mid(sItem, 6)
    sHijacker = Replace(sHijacker, vbTab, " ")
    If InStr(sHijacker, " ") = 0 Then Exit Sub
    sHijacker = Mid(LTrim(sHijacker), InStr(LTrim(sHijacker), " ") + 1)
    sHijacker = RTrim(sHijacker)
    If InStr(sHijacker, " ") > 0 Then sHijacker = Left(sHijacker, InStr(sHijacker, " ") - 1)
    If InStr(sHijacker, vbTab) > 0 Then sHijacker = Left(sHijacker, InStr(sHijacker, vbTab) - 1)
    
    If FileExists(sHostsFile & ".new") Then
        SetAttr sHostsFile & ".new", vbNormal
        DeleteFile sHostsFile & ".new"
    End If
    
    Open sHostsFile For Input As #1
    Open sHostsFile & ".new" For Output As #2
        Do
            Line Input #1, sLine
            If InStr(sLine, sHijacker) > 0 Then
                'don't write line to hosts file
            Else
                Print #2, sLine
            End If
        Loop Until EOF(1)
    Close #1
    Close #2
    DeleteFile sHostsFile
    Name sHostsFile & ".new" As sHostsFile
    SetAttr sHostsFile, iAttr
    Exit Sub
    
Error:
    Close
    If Err.Number = 70 And Not bSeenHostsFileAccessDeniedWarning Then
        'permission denied
        MsgBox Translate(303), vbExclamation
'        MsgBox "HijackThis could not write the selected changes to your " & _
'               "hosts file. The probably cause is that some program is " & _
'               "denying access to it, or that your user account doesn't have " & _
'               "the rights to write to it.", vbExclamation
        bSeenHostsFileAccessDeniedWarning = True
    Else
        ErrorMsg "modMain_FixOther1Item", Err.Number, Err.Description, "sItem=" & sItem
    End If
End Sub

Public Sub FixOther5Item(sItem$)
    'O5 - Blocking of loading Internet Options in Control Panel
    'WritePrivateProfileString "don't load", "inetcpl.cpl", vbNullString, "control.ini"
    On Error GoTo Error:
    IniSetString "control.ini", "don't load", "inetcpl.cpl", ""
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther5Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub FixOther6Item(sItem$)
    'O6 - Disabling of Internet Options' Main tab with Policies
    Dim lHive&
    On Error GoTo Error:
    If Mid(sItem, 6, 4) = "HKLM" Then
        lHive = HKEY_LOCAL_MACHINE
    ElseIf Mid(sItem, 6, 4) = "HKCU" Then
        lHive = HKEY_CURRENT_USER
    End If
    If InStr(sItem, "Restrictions") > 0 Then
        RegDelKey lHive, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    ElseIf InStr(sItem, "Control Panel") > 0 Then
        RegDelKey lHive, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther6Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub FixOther7Item(sItem$)
    'O7 - Disabling of Regedit with Policies
    On Error GoTo Error:
    If Mid(sItem, 6, 4) = "HKLM" Then
        RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
    ElseIf Mid(sItem, 6, 4) = "HKCU" Then
        RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther7Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub FixOther8Item(sItem$)
    'O8 - Extra context menu items
    'O8 - Extra context menu item: [name] - html file
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt
    
    Dim sName$
    On Error GoTo Error:
    sName = Mid(sItem, InStr(sItem, ": ") + 2)
    sName = Left(sName, InStrRev(sName, " - ") - 1)
    RegDelKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & sName
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther8Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther2Item()
    Dim hKey&, i&, j&, sName$, sCLSID$, sFile$, sHit$
    On Error GoTo Error:
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then Exit Sub
    Do
        sCLSID = String(255, 0)
        If RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, 0, ByVal 0) <> 0 Then Exit Do
        sCLSID = Left(sCLSID, InStr(sCLSID, Chr(0)) - 1)
        If sCLSID <> vbNullString And _
           InStr(1, sCLSID, "MSHist", vbTextCompare) <> 1 Then
            'get filename from HKCR\CLSID\sName
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", "")
            
            If InStr(sFile, "__BHODemonDisabled") > 0 Then
                sFile = Left(sFile, InStr(sFile, "__BHODemonDisabled") - 1) & _
                " (disabled by BHODemon)"
            Else
                If sFile <> vbNullString And Not FileExists(sFile) Then sFile = sFile & " (file missing)"
            End If
            If sFile = vbNullString Then sFile = "(no file)"
            
            'get bho name from BHO regkey
            sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID, "")
            If sName = vbNullString Then
                'get BHO name from CLSID regkey
                sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, "")
                If sName = vbNullString Then sName = "(no name)"
            End If
            
            sHit = "O2 - BHO: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                frmMain.lstResults.AddItem sHit
            End If
            
            If InStr(sCLSID, "}}") > 0 Then
                'the new searchwww.com trick - use a double
                '}} in the IE toolbar registration, reg the toolbar
                'with only one } - IE ignores the double }}, but
                'HT didn't. It does now!
                
                sCLSID = Left(sCLSID, Len(sCLSID) - 1)
            
                sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", "")
                If InStr(sFile, "__BHODemonDisabled") > 0 Then
                    sFile = Left(sFile, InStr(sFile, "__BHODemonDisabled") - 1) & _
                    " (disabled by BHODemon)"
                Else
                    If sFile <> vbNullString And Not FileExists(sFile) Then sFile = sFile & " (file missing)"
                End If
                
                If sFile = vbNullString Then sFile = "(no file)"
                sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID, "")
                If sName = vbNullString Then sName = "(no name)"
                
                sHit = "O2 - BHO: " & sName & " - " & sCLSID & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        End If
        i = i + 1
    Loop
    RegCloseKey hKey
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modMain_CheckOther2Item", Err.Number, Err.Description
End Sub

Public Sub CheckOther8Item()
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt
    
    On Error GoTo Error:
    Dim hKey&, hKey2&, i&, sName$, sData$, sHit$
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        i = 0
        sName = String(255, 0)
        If RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then RegCloseKey hKey: Exit Sub
        Do
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
            sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & sName, vbNullString)
            If sData <> vbNullString Then
                sHit = "O8 - Extra context menu item: " & sName & " - " & sData
                If Not IsOnIgnoreList(sHit) Then
                    'md5 doesn't seem useful here
                    If bMD5 Then sHit = sHit & GetFileMD5(sData)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
            sName = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modMain_CheckOther8Item", Err.Number, Err.Description
End Sub

Public Sub CheckOther9Item()
    'HKLM\Software\Microsoft\Internet Explorer\Extensions
    'HKCU\..\etc
    
    On Error GoTo Error:
    Dim hKey&, hKey2&, i&, sData$, sCLSID$, sCLSID2$, sFile$, sHit$
    'open root key
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        i = 0
        sCLSID = String(255, 0)
        'start enum of root key subkeys (i.e., extensions)
        If RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then RegCloseKey hKey: Exit Sub
        Do
            sCLSID = TrimNull(sCLSID)
            If sCLSID = "CmdMapping" Then GoTo NextExtHKLM:
            
            'check for 'MenuText' or 'ButtonText'
            sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "ButtonText")
            
            'this clsid is mostly useless, always pointing to SHDOCVW.DLL
            'places to look for correct dll:
            '* Exec
            '* Script
            '* BandCLSID
            '* CLSIDExtension
            '* CLSIDExtension -> TreatAs CLSID
            '* CLSID
            '* ???
            '* actual CLSID of regkey (not used)
            sFile = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "Exec")
            If sFile = vbNullString Then
                sFile = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "Script")
                If sFile = vbNullString Then
                    sCLSID2 = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "BandCLSID")
                    sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
                    If sFile = vbNullString Then
                        sCLSID2 = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "CLSIDExtension")
                        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
                        If sFile = vbNullString Then
                            sCLSID2 = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\TreatAs", "")
                            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
                            If sFile = vbNullString Then
                                sCLSID2 = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "CLSID")
                                sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
                            End If
                        End If
                    End If
                End If
            End If
            
            If sFile <> vbNullString Then
                'expand %systemroot% var
                'sFile = Replace(sFile, "%systemroot%", sWinDir, , , vbTextCompare)
                sFile = NormalizePath(sFile)
                
                'strip stuff from res://[dll]/page.htm to just [dll]
                If InStr(1, sFile, "res://", vbTextCompare) = 1 And _
                   (LCase(Right(sFile, 4)) = ".htm" Or LCase(Right(sFile, 4)) = "html") Then
                    sFile = Mid(sFile, 7)
                    sFile = Left(sFile, InStrRev(sFile, "/") - 1)
                End If
                
                'remove other stupid prefixes
                If InStr(sFile, "file://") = 1 And _
                   InStr(sFile, "http://") <> 1 Then
                    If Not FileExists(Mid(sFile, 8)) Then sFile = sFile & " (file missing)"
                Else
                    If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
                End If
            Else
                sFile = "(no file)"
            End If
            
            If sData = vbNullString Then sData = "(no name)"
            If InStr(sData, "@shdoclc.dll,-866") > 0 Then sData = "Related"
            
            sHit = "O9 - Extra button: " & sData & " - " & sCLSID & " - " & sFile '& " (HKLM)"
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                frmMain.lstResults.AddItem sHit
            End If
                
            sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "MenuText")
            'don't show it again in case sdata=null
            If sData <> vbNullString Then
                If InStr(sData, "@shdoclc.dll,-864") > 0 Then sData = "Show &Related Links"
                sHit = "O9 - Extra 'Tools' menuitem: " & sData & " - " & sCLSID & " - " & sFile '& " (HKLM)"
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
NextExtHKLM:
            sCLSID = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    '-----------------------------
    'repeat for HKCU
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        i = 0
        sCLSID = String(255, 0)
        'start enum of root key subkeys (i.e., extensions)
        If RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then RegCloseKey hKey: Exit Sub
        Do
            sCLSID = TrimNull(sCLSID)
            If sCLSID = "CmdMapping" Then GoTo NextExtHKCU:
            
            'check for 'MenuText' or 'ButtonText'
            sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "ButtonText")
            
            sFile = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "Exec")
            If sFile = vbNullString Then
                sFile = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "Script")
                If sFile = vbNullString Then
                    sCLSID2 = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "BandCLSID")
                    sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
                    If sFile = vbNullString Then
                        sCLSID2 = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "CLSIDExtension")
                        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
                        If sFile = vbNullString Then
                            sCLSID2 = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\TreatAs", "")
                            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
                            If sFile = vbNullString Then
                                sCLSID2 = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "CLSID")
                                sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
                            End If
                        End If
                    End If
                End If
            End If
            
'            sFile = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "Exec")
'            If sFile = vbNullString Then
'                sFile = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "Script")
'                If sFile = vbNullString Then
'                    sCLSID2 = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "BandCLSID")
'                    sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", "")
'                End If
'            End If
            
            If sFile <> vbNullString Then
                'sFile = Replace(sFile, "%systemroot%", sWinDir, , , vbTextCompare)
                sFile = NormalizePath(sFile)
                
                If InStr(sFile, "file://") = 1 And InStr(sFile, "http://") <> 1 Then
                    If Not FileExists(Mid(sFile, 8)) Then sFile = sFile & " (file missing)"
                Else
                    If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
                End If
            Else
                sFile = "(no file)"
            End If
            
            If sData = vbNullString Then sData = "(no name)"
            If InStr(sData, "@shdoclc.dll,-866") > 0 Then sData = "Related"
            
            sHit = "O9 - Extra button: " & sData & " - " & sCLSID & " - " & sFile & " (HKCU)"
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                frmMain.lstResults.AddItem sHit
            End If
                
            sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "MenuText")
            If sData <> vbNullString Then
                If InStr(sData, "@shdoclc.dll,-864") > 0 Then sData = "Show &Related Links"
                sHit = "O9 - Extra 'Tools' menuitem: " & sData & " - " & sCLSID & " - " & sFile & " (HKCU)"
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
NextExtHKCU:
            sCLSID = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sCLSID, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    Exit Sub
    
Error:
    ErrorMsg "modMain_CheckOther9Item", Err.Number, Err.Description
End Sub

Public Sub FixOther2Item(sItem$)
    'O2 - Enumeration of existing MSIE BHO's
    'O2 - BHO: AcroIEHlprObj Class - {00000...000} - C:\PROGRAM FILES\ADOBE\ACROBAT 5.0\ACROBAT\ACTIVEX\ACROIEHELPER.OCX
    'O2 - BHO: ... (no file)
    'O2 - BHO: ... c:\bla.dll (file missing)
    'O2 - BHO: ... c:\bla.dll (disabled by BHODemon)
    
    Dim hKey&, i&, sData$
    Dim sName$, sCLSID$, sFile$, vBlah As Variant
    On Error GoTo Error:
    sName = Mid(sItem, 11)
    vBlah = Split(sName, " - ")
    If UBound(vBlah) = 2 Then
        'new method, should take care of those stupid
        'tricks with extra/missing brackets forever
        sName = CStr(vBlah(0))
        sCLSID = CStr(vBlah(1))
        sFile = CStr(vBlah(2))
    Else
        'something really odd is going on, so use
        'old method.
        If InStr(sName, "- -{") > 0 Then
            'stupid stupid trick
            sCLSID = Mid(sName, InStr(sName, " - -{") + 3)
        Else
            sCLSID = Mid(sName, InStr(sName, " - {") + 3)
        End If
        sFile = Mid(sCLSID, InStr(sCLSID, " - ") + 3)
        sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
        sName = Left(sName, InStr(sName, " - ") - 1)
    End If
    
    'extra strings appended to sFile
    If InStr(sFile, " (disabled by BHODemon)") > 0 Then
        sFile = Left(sFile, InStr(sFile, " (disabled by BHODemon)") - 1)
    End If
    If InStr(sFile, " (file missing)") > 0 Then sFile = vbNullString
    If sFile = "(no file)" Then sFile = vbNullString
    
    If Not bShownBHOWarning Then
        MsgBox Translate(310), vbExclamation
'        MsgBox "HijackThis is about to remove a " & _
'               "BHO and the corresponding file from " & _
'               "your system. Close all Internet " & _
'               "Explorer windows AND all Windows " & _
'               "Explorer windows before continuing for " & _
'               "the best chance of success.", vbExclamation
        bShownBHOWarning = True
    End If
    
    'On Error Resume Next
    'If sFile <> vbNullString Then
    '    If InStr(1, sFile, "dreplace.dll", vbTextCompare) = 0 And _
    '       InStr(1, sFile, "dnse.dll", vbTextCompare) = 0 Then
    '        Shell sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\regsvr32.exe /u /s """ & sFile & """", vbHide
    '        DoEvents
    '    End If
    'End If
    'On Error GoTo Error:
    
    'RegDelSubKeys HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID
    RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID
    'RegDelSubKeys HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
    RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
    On Error Resume Next
    If sFile <> vbNullString Then DeleteFile sFile
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther2Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub FixOther3Item(sItem$)
    'O3 - Enumeration of existing MSIE toolbars
    
    On Error GoTo Error:
    Dim sCLSID$
    If InStr(sItem, "{") > 0 And InStr(sItem, "}") > 0 Then
        sCLSID = Mid(sItem, InStr(sItem, "{"))
        sCLSID = Left(sCLSID, InStrRev(sCLSID, "}"))
        RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Toolbar", sCLSID
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther3Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub FixOther4Item(sItem$)
    'O4 - Enumeration of autoloading Regedit entries
    'O4 - HKLM\..\Run: [blah] program.exe
    'O4 - Startup: bla.lnk = c:\bla.exe
    'O4 - HKUS\S-1-5-19\..\Run: [blah] program.exe (Username 'Joe')
    
    'O4 - Startup: bla.exe
    
    On Error GoTo Error:
    Dim lHive&, sKey$, sVal$, sData$
    If InStr(sItem, "[") = 0 Then GoTo FixShortCut
    sItem = Mid(sItem, 6)
    Select Case Left(sItem, 4)
        Case "HKCU"
            lHive = HKEY_CURRENT_USER
        Case "HKLM"
            lHive = HKEY_LOCAL_MACHINE
        Case "HKUS"
            FixOther4ItemUsers "O4 - " & sItem
            Exit Sub
    End Select
    
    If InStr(sItem, "\RunServices:") > 0 Then
        sKey = "Software\Microsoft\Windows\CurrentVersion\RunServices"
    ElseIf InStr(sItem, "\RunOnce:") > 0 Then
        sKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
    ElseIf InStr(sItem, "\RunServicesOnce:") > 0 Then
        sKey = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    Else
        If InStr(1, sItem, "\Policies\", vbTextCompare) = 0 Then
            sKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        Else
            sKey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
        End If
    End If

    sVal = Mid(sItem, InStr(sItem, "[") + 1)
    sData = Mid(sVal, InStrRev(sVal, "]") + 2)
    KillProcessByFile GetFileFromAutostart(sData, False)
    'some wankers used a garbled value name with a ']' in it.
    'assuming no one ever uses a filename with a ']' in it in the
    'future, this workaround should work (InStrRev instead of InStr)
    'update: autorun with sol[1].exe - doh!
    sVal = Left(sVal, InStrRev(sVal, "] ") - 1)
    
    RegDelVal lHive, sKey, sVal
    Exit Sub
    
FixShortCut:
    'O4 - Startup: bla.lnk = c:\bla.exe
    Dim sPath$, sFile$
    If InStr(sItem, " (User '") > 0 Then
        FixOther4ItemUsers sItem
        Exit Sub
    End If
    sPath = Mid(sItem, 6)
    If InStr(sPath, ": ") = 0 Then Exit Sub
    sPath = Left(sPath, InStr(sPath, ": ") - 1)
    Select Case sPath
        Case "Startup":                sPath = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
        Case "User Startup":           sPath = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
        Case "Global Startup":         sPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common Startup")
        Case "Global User Startup":    sPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common Startup")
        Case "Global User AltStartup": sPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common AltStartup")
    End Select
    If sPath = vbNullString Then Exit Sub
    sFile = Mid(sItem, InStr(sItem, ": ") + 2)
    If InStr(sFile, " = ") > 0 Then
        sData = Mid(sFile, InStr(sFile, " = ") + 3)
        sFile = Left(sFile, InStr(sFile, " = ") - 1)
    Else
        sData = sPath & "\" & sFile
    End If
    sFile = sPath & IIf(Right(sPath, 1) = "\", "", "\") & sFile
    If FileExists(sFile) Then
        On Error Resume Next
        If (GetAttr(sFile) And vbDirectory) Then
            DeleteFolder sFile
        Else
            KillProcessByFile GetFileFromAutostart(sData)
            DeleteFile sFile
        End If
        If Err Then
            MsgBox Err.Description
            MsgBox Replace(Translate(320), "[]", sItem) & " " & _
                   IIf(bIsWinNT, Translate(321), Translate(322)) & _
                   " " & Translate(323), vbExclamation
'            MsgBox "Unable to delete the file:" & vbCrLf & _
'                   sItem & vbCrLf & vbCrLf & "The file " & _
'                   "may be in use. Use " & IIf(bIsWinNT, _
'                   "Task Manager", "a process killer like " & _
'                   "ProcView") & " to shutdown the program " & _
'                   "and run HijackThis again to delete the file.", vbExclamation
        End If
        On Error GoTo Error:
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther4Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub FixOther9Item(sItem$)
    'O9 - Extra buttons/Tools menu items
    'O9 - Extra button: [name] - [CLSID] - [file] [(HKCU)]
    
    On Error GoTo Error:
    Dim hKey&, i&, sName$, sData$, sCaption$, sCLSID$, lHive&
    sCLSID = Mid(sItem, InStr(sItem, ": ") + 2)
    sCLSID = Mid(sCLSID, InStr(sCLSID, " - ") + 3)
    sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
    
    If InStr(sItem, " (HKCU)") > 0 Then
        lHive = HKEY_CURRENT_USER
    Else
        lHive = HKEY_LOCAL_MACHINE
    End If
    
    RegDelKey lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID
    Exit Sub
    
    'outdated stuff
    If RegOpenKeyEx(lHive, "Software\Microsoft\Internet Explorer\Extensions", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sName = String(255, 0)
        i = 0
        If RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then RegCloseKey hKey: Exit Sub
        Do
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
            sData = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sName, "ButtonText")
            If sData = sCaption Then
                RegCloseKey hKey
                RegDelKey lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sName
                Exit Sub
            End If
                
            sData = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sName, "MenuText")
            If sData = sCaption Then
                RegCloseKey hKey
                RegDelKey lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sName
                Exit Sub
            End If
            sName = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther9Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther4Item()
    Dim i%, j%, k%, hKey&, sName$, uData() As Byte, sMD5$
    Dim lHive&, sKey$, sRegRuns$(1 To 10), sData$, sHit$
    On Error GoTo Error:
    
    sRegRuns(1) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run"
    sRegRuns(2) = "HKLM\Software\Microsoft\Windows\CurrentVersion\RunServices"
    sRegRuns(3) = "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    sRegRuns(4) = "HKLM\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    sRegRuns(5) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
    sRegRuns(6) = "HKCU\Software\Microsoft\Windows\CurrentVersion\RunServices"
    sRegRuns(7) = "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    sRegRuns(8) = "HKCU\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    'added in 1.99.2
    sRegRuns(9) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
    sRegRuns(10) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
    'also see CheckOther4ItemUsers()
    
    For k = 1 To UBound(sRegRuns)
        If Left(sRegRuns(k), 4) = "HKLM" Then
            lHive = HKEY_LOCAL_MACHINE
        ElseIf Left(sRegRuns(k), 4) = "HKCU" Then
            lHive = HKEY_CURRENT_USER
        End If
        sKey = Mid(sRegRuns(k), 6)
    
        RegOpenKeyEx lHive, sKey, 0, KEY_QUERY_VALUE, hKey
        If hKey <> 0 Then
            Do
                sName = String(lEnumBufSize, 0)
                ReDim uData(lEnumBufSize)
                If RegEnumValue(hKey, i, sName, Len(sName), 0, ByVal 0, uData(0), UBound(uData)) = 0 Then
                    sName = TrimNull(sName)
                    'sData = ""
                    'For j = 0 To 510
                    '    If uData(j) = 0 Then Exit For
                    '    sData = sData & Chr(uData(j))
                    'Next j
                    sData = StrConv(uData, vbUnicode)
                    sData = TrimNull(sData)
                    
                    If sData <> vbNullString Then
                        Select Case k
                            Case 1: sHit = "O4 - HKLM\..\Run: "
                            Case 2: sHit = "O4 - HKLM\..\RunServices: "
                            Case 3: sHit = "O4 - HKLM\..\RunOnce: "
                            Case 4: sHit = "O4 - HKLM\..\RunServicesOnce: "
                            Case 5: sHit = "O4 - HKCU\..\Run: "
                            Case 6: sHit = "O4 - HKCU\..\RunServices: "
                            Case 7: sHit = "O4 - HKCU\..\RunOnce: "
                            Case 8: sHit = "O4 - HKCU\..\RunServicesOnce: "
                            Case 9: sHit = "O4 - HKLM\..\Policies\Explorer\Run: "
                            Case 10: sHit = "O4 - HKCU\..\Policies\Explorer\Run: "
                        End Select
                        sHit = sHit & "[" & sName & "] " & sData
                        If Not IsOnIgnoreList(sHit) Then
                            If bMD5 Then sMD5 = GetFileFromAutostart(sData)
                            sHit = sHit & sMD5
                            frmMain.lstResults.AddItem sHit
                        End If
                    End If
                Else
                    Exit Do
                End If
                i = i + 1
            Loop
            'this last one makes some registry problems
            'and I can't figure out why
            RegCloseKey hKey
        End If
        i = 0
    Next k
    'added in HJT 1.99.2
    CheckOther4ItemUsers
    
    Dim sAutostartFolder$(1 To 8), sFile$, sShortCut$
    sAutostartFolder(1) = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
    sAutostartFolder(2) = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AltStartup")
    sAutostartFolder(3) = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
    sAutostartFolder(4) = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AltStartup")
    sAutostartFolder(5) = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common Startup")
    sAutostartFolder(6) = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common AltStartup")
    sAutostartFolder(7) = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common Startup")
    sAutostartFolder(8) = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common AltStartup")
    
    For k = 1 To UBound(sAutostartFolder)
        If sAutostartFolder(k) <> vbNullString And _
           FolderExists(sAutostartFolder(k)) Then
            sShortCut = Dir(sAutostartFolder(k) & "\*.*", vbArchive + vbHidden + vbReadOnly + vbSystem + vbDirectory)
            If sShortCut <> vbNullString Then
                Do
                    Select Case k
                        Case 1: sHit = "O4 - Startup: "
                        Case 2: sHit = "O4 - AltStartup: "
                        Case 3: sHit = "O4 - User Startup: "
                        Case 4: sHit = "O4 - User AltStartup: "
                        Case 5: sHit = "O4 - Global Startup: "
                        Case 6: sHit = "O4 - Global AltStartup: "
                        Case 7: sHit = "O4 - Global User Startup: "
                        Case 8: sHit = "O4 - Global User AltStartup: "
                    End Select
                    sFile = GetFileFromShortCut(sAutostartFolder(k) & "\" & sShortCut)
                    sHit = sHit & sShortCut & sFile
                    If LCase(sShortCut) <> "desktop.ini" And _
                       sShortCut <> "." And sShortCut <> ".." And _
                       Not IsOnIgnoreList(sHit) Then
                        If bMD5 And sFile <> vbNullString And sFile <> " = ?" Then
                            sHit = sHit & GetFileMD5(Mid(sFile, 4))
                        End If
                        frmMain.lstResults.AddItem sHit
                    End If
                    
                    sShortCut = Dir
                Loop Until sShortCut = vbNullString
            End If
        End If
    Next k
    Exit Sub
    
Error:
    ErrorMsg "modMain_CheckOther4Item", Err.Number, Err.Description
End Sub

Public Sub CheckOther3Item()
    'HKLM\Software\Microsoft\Internet Explorer\Toolbar
    On Error GoTo Error:
    
    Dim hKey&, hKey2&, i%, j%, sCLSID$, sName$
    Dim uData() As Byte, sFile$, sHit$
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Toolbar", 0, KEY_QUERY_VALUE, hKey) <> 0 Then Exit Sub
    Do
        sCLSID = String(lEnumBufSize, 0)
        ReDim uData(lEnumBufSize)
        
        'enumerate MSIE toolbars
        If RegEnumValue(hKey, i, sCLSID, Len(sCLSID), 0, ByVal 0, uData(0), UBound(uData)) <> 0 Then Exit Do
        sCLSID = Left(sCLSID, InStr(sCLSID, Chr(0)) - 1)
        
        'found one? then check corresponding HKCR key
        sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, "")
        If sName = vbNullString Then sName = "(no name)"
        'If HasSpecialCharacters(sName) Then
            'when japanese characters are in toolbar name,
            'it tends to screw up things
        '    sName = "?????"
        'End If
        
        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString)
        If sFile = vbNullString Then
            sFile = "(no file)"
        Else
            If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
        End If
        
        '   sCLSID <> "BrandBitmap" And _
        '   sCLSID <> "SmBrandBitmap" And _
        '   sCLSID <> "BackBitmap" And _
        '   sCLSID <> "BackBitmapIE5" And _
        '   sCLSID <> "OLE (Part 1 of 5)" And _
        '   sCLSID <> "OLE (Part 2 of 5)" And _
        '   sCLSID <> "OLE (Part 3 of 5)" And _
        '   sCLSID <> "OLE (Part 4 of 5)" And _
        '   sCLSID <> "OLE (Part 5 of 5)" Then
        
        If sName <> vbNullString And _
           InStr(sCLSID, "{") > 0 Then
            sHit = "O3 - Toolbar: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                frmMain.lstResults.AddItem sHit
            End If
        End If
        
        If InStr(sCLSID, "}}") > 0 Then
            'the new searchwww.com trick - use a double
            '}} in the IE toolbar registration, reg the toolbar
            'with only one } - IE ignores the double }}, but
            'HT didn't. It does now!
            
            sCLSID = Left(sCLSID, Len(sCLSID) - 1)
        
            sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, "")
            If sName = vbNullString Then sName = "(no name)"
            'If HasSpecialCharacters(sName) Then sName = "?????"
            
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString)
            If sFile = vbNullString Then
                sFile = "(no file)"
            Else
                If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
            End If
            If sName <> vbNullString And _
               sCLSID <> "BrandBitmap" And _
               sCLSID <> "SmBrandBitmap" Then
                sHit = "O3 - Toolbar: " & sName & " - " & sCLSID & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        End If
        
        i = i + 1
    Loop
    RegCloseKey hKey
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modMain_CheckOther3Item", Err.Number, Err.Description
End Sub

Public Sub CheckOther11Item()
    'HKLM\Software\Microsoft\Internet Explorer\AdvancedOptions
    Dim hKey&, i&, sKey$, sName$, sHit$
    On Error GoTo Error:
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\AdvancedOptions", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sKey = String(255, 0)
        If RegEnumKeyEx(hKey, i, sKey, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then RegCloseKey hKey: Exit Sub
        Do
            sKey = Left(sKey, InStr(sKey, Chr(0)) - 1)
            sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\AdvancedOptions\" & sKey, "Text")
            If InStr("JAVA_VM.JAVA_SUN.BROWSE.ACCESSIBILITY.SEARCHING." & _
                     "HTTP1.1.MULTIMEDIA.Multimedia.CRYPTO.PRINT." & _
                     "TOEGANKELIJKHEID.TABS.INTERNATIONAL*", sKey) = 0 And _
               sName <> vbNullString Then
                sHit = "O11 - Options group: [" & sKey & "] " & sName
                If bIgnoreAllWhitelists = True Then
                    frmMain.lstResults.AddItem sHit
                ElseIf Not IsOnIgnoreList(sHit) Then
                    frmMain.lstResults.AddItem sHit
                End If
            End If
            sKey = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sKey, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_CheckOther11Item", Err.Number, Err.Description
End Sub

Public Sub FixOther11Item(sItem$)
    'O11 - Options group: [BLA] Blah"
    Dim sKey$, hKey&, i&, sName$, sSubKeys$()
    On Error GoTo Error:
    sKey = Mid(sItem, InStr(sItem, "[") + 1)
    sKey = Left(sKey, InStr(sKey, "]") - 1)
    
    'RegDelSubKeys HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\AdvancedOptions\" & sKey
    RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\AdvancedOptions\" & sKey
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther11Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther12Item()
    'HKLM\Software\Microsoft\Internet Explorer\Plugins\Extensions
    'HKLM\Software\Microsoft\Internet Explorer\Plugins\MIME
    
    Dim hKey&, i&, sName$, sFile$, sHit$
    On Error GoTo Error:
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Plugins\Extension", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sName = String(255, 0)
        If RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then RegCloseKey hKey: Exit Sub
        Do
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
            sFile = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Plugins\Extension\" & sName, "Location")
            If sFile <> vbNullString Then
                sHit = "O12 - Plugin for " & sName & ": " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
            
            sName = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    hKey = 0
    i = 0
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Plugins\MIME", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sName = String(255, 0)
        If RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then RegCloseKey hKey: Exit Sub
        Do
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
            sFile = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Plugins\MIME\" & sName, "Location")
            If sFile <> vbNullString Then
                sHit = "O12 - Plugin for " & sName & ": " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
            
            sName = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modMain_CheckOther12Item", Err.Number, Err.Description
End Sub

Public Sub FixOther12Item(sItem$)
    'O12 - Plugin for .ofb: C:\Win98\blah.dll
    'O12 - Plugin for text/blah: C:\Win98\blah.dll
    
    Dim sKey$, sFile$, sType$
    On Error GoTo Error:
    sFile = Mid(sItem, InStr(sItem, ": ") + 2)
    sKey = Mid(sItem, InStr(sItem, "for ") + 4)
    sKey = Left(sKey, InStr(sKey, ": ") - 1)
    If InStr(sKey, ".") > 0 Then
        sType = "Extension\"
    Else
        sType = "MIME\"
    End If
    
    If Not bShownToolbarWarning Then
        MsgBox Translate(330), vbExclamation
'        MsgBox "HijackThis is about to remove a " & _
'               "plugin from " & _
'               "your system. Close all Internet " & _
'               "Explorer windows before continuing for " & _
'               "the best chance of success.", vbExclamation
        bShownToolbarWarning = True
    End If
    RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Plugins\" & sType & sKey
    On Error Resume Next
    DeleteFile sFile
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther12Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther13Item()
    'O13
    Dim sDummy$, sKeyURL$, sHit$
    On Error GoTo Error:
    sKeyURL = "Software\Microsoft\Windows\CurrentVersion\URL"
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeyURL & "\DefaultPrefix", "")
    If sDummy <> "http://" Then
        'infected!
        sHit = "O13 - DefaultPrefix: " & sDummy
        If Not IsOnIgnoreList(sHit) Then
            frmMain.lstResults.AddItem sHit
        End If
    End If
    
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "www")
    If sDummy <> "http://" Then
        'infected!
        sHit = "O13 - WWW Prefix: " & sDummy
        If Not IsOnIgnoreList(sHit) Then
            frmMain.lstResults.AddItem sHit
        End If
    End If
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "www.")
    If sDummy <> vbNullString Then
        'infected!
        sHit = "O13 - WWW. Prefix: " & sDummy
        If Not IsOnIgnoreList(sHit) Then
            frmMain.lstResults.AddItem sHit
        End If
    End If
    
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "home")
    If sDummy <> "http://" Then
        'infected!
        sHit = "O13 - Home Prefix: " & sDummy
        If Not IsOnIgnoreList(sHit) Then
            frmMain.lstResults.AddItem sHit
        End If
    End If
    
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "mosaic")
    If sDummy <> "http://" Then
        'infected!
        sHit = "O13 - Mosaic Prefix: " & sDummy
        If Not IsOnIgnoreList(sHit) Then
            frmMain.lstResults.AddItem sHit
        End If
    End If
    
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "ftp")
    If sDummy <> "ftp://" Then
        sHit = "O13 - FTP Prefix: " & sDummy
        If Not IsOnIgnoreList(sHit) Then
            frmMain.lstResults.AddItem sHit
        End If
    End If
    
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "gopher")
    If sDummy <> "gopher://" And sDummy <> vbNullString Then
        sHit = "O13 - Gopher Prefix: " & sDummy
        If Not IsOnIgnoreList(sHit) Then
            frmMain.lstResults.AddItem sHit
        End If
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_CheckOther13Item", Err.Number, Err.Description
End Sub

Public Sub FixOther13Item(sItem$)
    'defaultprefix fix
    'O13 - DefaultPrefix: http://www.hijacker.com/redir.cgi?
    'O13 - [WWW/Home/Mosaic/FTP/Gopher] Prefix: ..
    
    Dim sDummy$, sKeyURL$
    On Error GoTo Error:
    sKeyURL = "Software\Microsoft\Windows\CurrentVersion\URL"
    sDummy = Left(sItem, InStr(sItem, ":") - 1)
    sDummy = Mid(sDummy, 7)
    Select Case sDummy
        Case "DefaultPrefix": RegSetStringVal HKEY_LOCAL_MACHINE, sKeyURL & "\DefaultPrefix", "", "http://"
        Case "WWW Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "www", "http://"
        Case "WWW. Prefix": RegDelVal HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "www."
        Case "Home Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "home", "http://"
        Case "Mosaic Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "mosaic", "http://"
        Case "FTP Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "ftp", "ftp://"
        Case "Gopher Prefix": RegSetStringVal HKEY_LOCAL_MACHINE, sKeyURL & "\Prefixes", "gopher", "gopher://"
    End Select
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther13Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther14Item()
    'O14 - Reset Websettings check
    Dim sLine$, sHit$
    Dim sStartPage$, sSearchPage$, sMsStartPage$
    Dim sSearchAssis$, sCustSearch$
    On Error GoTo Error:
    
    If Not FileExists(sWinDir & "\inf\iereset.inf") Then Exit Sub
    If FileLen(sWinDir & "\INF\iereset.inf") = 0 Then Exit Sub
    Open sWinDir & "\INF\iereset.inf" For Input As #1
        Do
            Line Input #1, sLine
            If InStr(sLine, "SearchAssistant") > 0 Then
                sSearchAssis = Mid(sLine, InStr(sLine, "http://"))
                sSearchAssis = Left(sSearchAssis, Len(sSearchAssis) - 1)
            End If
            If InStr(sLine, "CustomizeSearch") > 0 Then
                sCustSearch = Mid(sLine, InStr(sLine, "http://"))
                sCustSearch = Left(sCustSearch, Len(sCustSearch) - 1)
            End If
            If InStr(sLine, "START_PAGE_URL=") = 1 And _
               InStr(sLine, "MS_START_PAGE_URL") = 0 Then
                sStartPage = Mid(sLine, InStr(sLine, "=") + 1)
                If Left(sStartPage, 1) = """" Then sStartPage = Mid(sStartPage, 2)
                If Right(sStartPage, 1) = """" Then sStartPage = Left(sStartPage, Len(sStartPage) - 1)
            End If
            If InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sSearchPage = Mid(sLine, InStr(sLine, "=") + 1)
                If Left(sSearchPage, 1) = """" Then sSearchPage = Mid(sSearchPage, 2)
                If Right(sSearchPage, 1) = """" Then sSearchPage = Left(sSearchPage, Len(sSearchPage) - 1)
            End If
            If InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sMsStartPage = Mid(sLine, InStr(sLine, "=") + 1)
                If Left(sMsStartPage, 1) = """" Then sMsStartPage = Mid(sMsStartPage, 2)
                If Right(sMsStartPage, 1) = """" Then sMsStartPage = Left(sMsStartPage, Len(sMsStartPage) - 1)
            End If
        Loop Until EOF(1)
    Close #1
    
    'SearchAssistant = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm
    If sSearchAssis <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm" Then
        sHit = "O14 - IERESET.INF: SearchAssistant=" & sSearchAssis
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'CustomizeSearch = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm
    If sCustSearch <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm" Then
        sHit = "O14 - IERESET.INF: CustomizeSearch=" & sCustSearch
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'SEARCH_PAGE_URL = http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch
    If sSearchPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch" Then
        sHit = "O14 - IERESET.INF: SEARCH_PAGE_URL=" & sSearchPage
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'START_PAGE_URL  = http://www.msn.com
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If sStartPage <> "http://www.msn.com" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome" Then
        sHit = "O14 - IERESET.INF: START_PAGE_URL=" & sStartPage
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'MS_START_PAGE_URL=http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '(=START_PAGE_URL) http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If sMsStartPage <> vbNullString Then
        If sMsStartPage <> "http://www.msn.com" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome" Then
            sHit = "O14 - IERESET.INF: MS_START_PAGE_URL=" & sMsStartPage
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_CheckOther14Item", Err.Number, Err.Description
End Sub

Public Sub FixOther14Item(sItem$)
    'resetwebsettings fix
    'O14 - IERESET.INF: [item]=[URL]
    
    Dim sLine$, sFixedIeResetInf$
    On Error GoTo Error:
    If Not FileExists(sWinDir & "\INF\iereset.inf") Then Exit Sub
    Open sWinDir & "\INF\iereset.inf" For Input As #1
        Do
            Line Input #1, sLine
            If InStr(sLine, "SearchAssistant") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""SearchAssistant"",0,""http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm""" & vbCrLf
            ElseIf InStr(sLine, "CustomizeSearch") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""CustomizeSearch"",0,""http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm""" & vbCrLf
            ElseIf InStr(sLine, "START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "START_PAGE_URL=""http://www.msn.com""" & vbCrLf
            ElseIf InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "SEARCH_PAGE_URL=""http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch""" & vbCrLf
            ElseIf InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "MS_START_PAGE_URL=""http://www.msn.com""" & vbCrLf
            Else
                sFixedIeResetInf = sFixedIeResetInf & sLine & vbCrLf
            End If
        Loop Until EOF(1)
    Close #1
    
    SetAttr sWinDir & "\INF\iereset.inf", vbArchive
    DeleteFile sWinDir & "\INF\iereset.inf"
    Open sWinDir & "\INF\iereset.inf" For Output As #1
        Print #1, Left(sFixedIeResetInf, Len(sFixedIeResetInf) - 2)
    Close #1
    Exit Sub
    
Error:
    Close #1
    ErrorMsg "modMain_FixOther14Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther10Item()
    CheckLSP
End Sub

Public Sub CheckOther15Item()
    'the value * or http denotes the protocol for which
    'the rule is valid. it's 2 for Trusted Zone and
    '4 for Restricted Zone.
    
    'Checks:
    '* ZoneMap\Domains          - trusted domains
    '* ZoneMap\Ranges           - trusted IPs and IP ranges
    '* ZoneMap\ProtocolDefaults - what zone rules does a protocol obey
    'added in 1.99.2
    '* ZoneMap\EscDomains       - trusted domains for Enhanced Security Configuration
    '* ZoneMap\EscRanges        - trusted IPs and IP ranges for ESC
    
    Dim sZoneMapDomains$, sZoneMapRanges$, sZoneMapProtDefs$
    Dim sZoneMapEscDomains$, sZoneMapEscRanges$
    Dim sDomains$(), sSubDomains$()
    Dim i&, j%, sHit$, sIPRange$
    sZoneMapDomains = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"
    sZoneMapRanges = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    sZoneMapProtDefs = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\ProtocolDefaults"
    sZoneMapEscDomains = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains"
    sZoneMapEscRanges = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscRanges"
    
    'decrypt NSBSD list
    For i = 0 To UBound(sSafeRegDomains)
        sSafeRegDomains(i) = Crypt(sSafeRegDomains(i), sProgramVersion, False)
    Next i
    
    'enum all subkeys (i.e. all domains)
    sDomains = Split(RegEnumSubkeys(HKEY_CURRENT_USER, sZoneMapDomains), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            If bIgnoreSafe Then
                For j = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sSafeRegDomains(j), CStr(sDomains(i)), vbTextCompare) > 0 Then
                        If InStr(sDomains(i), "msn.com") = 0 Then
                            'it's a safe domain!
                            GoTo NextDomain
                        Else
                            '*.msn.com is added by CWS - coupled
                            'with the Hosts file hijack this
                            'could reinstall it. so this is
                            'an exception to the whitelist :)
                            Exit For
                        End If
                    End If
                Next j
            End If
            sSubDomains = Split(RegEnumSubkeys(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i)), "|")
            If UBound(sSubDomains) <> -1 Then
                'list any trusted subdomains for main domain
                For j = 0 To UBound(sSubDomains)
                    If RegGetDword(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i) & "\" & sSubDomains(j), "*") = 2 Then
                        'entire subdomain is trusted
                        sHit = "O15 - Trusted Zone: " & sSubDomains(j) & "." & sDomains(i)
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                    If RegGetDword(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i) & "\" & sSubDomains(j), "http") = 2 Then
                        'only http on subdomain is trusted
                        sHit = "O15 - Trusted Zone: http://" & sSubDomains(j) & "." & sDomains(i)
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                Next j
            End If
            'list main domain as well if that's trusted too (*grumble*)
            If RegGetDword(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i), "*") = 2 Then
                'entire domain is trusted
                sHit = "O15 - Trusted Zone: *." & sDomains(i)
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            If RegGetDword(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i), "http") = 2 Then
                'only http on domain is trusted
                sHit = "O15 - Trusted Zone: http://*." & sDomains(i)
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
NextDomain:
        Next i
    End If
    
    'repeat for HKLM (domains)
    sDomains = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sZoneMapDomains), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            If bIgnoreSafe Then
                For j = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sSafeRegDomains(j), CStr(sDomains(i)), vbTextCompare) > 0 Then
                        If InStr(sDomains(i), "msn.com") = 0 Then
                            'it's a safe domain!
                            GoTo NextDomain2
                        Else
                            Exit For
                        End If
                    End If
                Next j
            End If
            sSubDomains = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i)), "|")
            If UBound(sSubDomains) <> -1 Then
                'list any trusted subdomains for main domain
                For j = 0 To UBound(sSubDomains)
                    If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i) & "\" & sSubDomains(j), "*") = 2 Then
                        'entire subdomain is trusted
                        sHit = "O15 - Trusted Zone: " & sSubDomains(j) & "." & sDomains(i) & " (HKLM)"
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                    If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i) & "\" & sSubDomains(j), "http") = 2 Then
                        'only http on subdomain is trusted
                        sHit = "O15 - Trusted Zone: http://" & sSubDomains(j) & "." & sDomains(i) & " (HKLM)"
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                Next j
            End If
            'list main domain as well, if applicable
            If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i), "*") = 2 Then
                'entire domain is trusted
                sHit = "O15 - Trusted Zone: *." & sDomains(i) & " (HKLM)"
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i), "http") = 2 Then
                'only http on domain is trusted
                sHit = "O15 - Trusted Zone: http://*." & sDomains(i) & " (HKLM)"
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
NextDomain2:
        Next i
    End If
    
    'enum all IP ranges
    sDomains = Split(RegEnumSubkeys(HKEY_CURRENT_USER, sZoneMapRanges), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            sIPRange = RegGetString(HKEY_CURRENT_USER, sZoneMapRanges & "\" & sDomains(i), ":Range")
            If Left(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                If RegGetDword(HKEY_CURRENT_USER, sZoneMapRanges & "\" & sDomains(i), "*") = 2 Then
                    'all protocols for this ip range is trusted
                    sHit = "O15 - Trusted IP range: " & sIPRange
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
                If RegGetDword(HKEY_CURRENT_USER, sZoneMapRanges & "\" & sDomains(i), "http") = 2 Then
                    'only http protocol for this ip range is trusted
                    sHit = "O15 - Trusted IP range: http://" & sIPRange
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
            End If
        Next i
    End If
    
    'repeat for HKLM (ip ranges)
    sDomains = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sZoneMapRanges), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            sIPRange = RegGetString(HKEY_LOCAL_MACHINE, sZoneMapRanges & "\" & sDomains(i), ":Range")
            If Left(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapRanges & "\" & sDomains(i), "*") = 2 Then
                    'all protocols for this ip range is trusted
                    sHit = "O15 - Trusted IP range: " & sIPRange & " (HKLM)"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
                If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapRanges & "\" & sDomains(i), "http") = 2 Then
                    'only http protocol for this ip range is trusted
                    sHit = "O15 - Trusted IP range: http://" & sIPRange & " (HKLM)"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
            End If
        Next i
    End If
    
'======================= REPEAT FOR ESC =======================
    'enum all subkeys (i.e. all domains)
    sDomains = Split(RegEnumSubkeys(HKEY_CURRENT_USER, sZoneMapEscDomains), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            If bIgnoreSafe Then
                For j = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sSafeRegDomains(j), CStr(sDomains(i)), vbTextCompare) > 0 Then
                        If InStr(sDomains(i), "msn.com") = 0 Then
                            'it's a safe domain!
                            GoTo NextEscDomain
                        Else
                            Exit For
                        End If
                    End If
                Next j
            End If
            sSubDomains = Split(RegEnumSubkeys(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i)), "|")
            If UBound(sSubDomains) <> -1 Then
                'list any trusted subdomains for main domain
                For j = 0 To UBound(sSubDomains)
                    If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i) & "\" & sSubDomains(j), "*") = 2 Then
                        'entire subdomain is trusted
                        sHit = "O15 - ESC Trusted Zone: " & sSubDomains(j) & "." & sDomains(i)
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                    If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i) & "\" & sSubDomains(j), "http") = 2 Then
                        'only http on subdomain is trusted
                        sHit = "O15 - ESC Trusted Zone: http://" & sSubDomains(j) & "." & sDomains(i)
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                Next j
            End If
            'list main domain as well if that's trusted too (*grumble*)
            If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i), "*") = 2 Then
                'entire domain is trusted
                sHit = "O15 - ESC Trusted Zone: *." & sDomains(i)
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i), "http") = 2 Then
                'only http on domain is trusted
                sHit = "O15 - ESC Trusted Zone: http://*." & sDomains(i)
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
NextEscDomain:
        Next i
    End If
    
    'repeat for HKLM (domains)
    sDomains = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sZoneMapEscDomains), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            If bIgnoreSafe Then
                For j = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sSafeRegDomains(j), CStr(sDomains(i)), vbTextCompare) > 0 Then
                        If InStr(sDomains(i), "msn.com") = 0 Then
                            'it's a safe domain!
                            GoTo NextEscDomain2
                        Else
                            Exit For
                        End If
                    End If
                Next j
            End If
            sSubDomains = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i)), "|")
            If UBound(sSubDomains) <> -1 Then
                'list any trusted subdomains for main domain
                For j = 0 To UBound(sSubDomains)
                    If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i) & "\" & sSubDomains(j), "*") = 2 Then
                        'entire subdomain is trusted
                        sHit = "O15 - ESC Trusted Zone: " & sSubDomains(j) & "." & sDomains(i) & " (HKLM)"
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                    If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i) & "\" & sSubDomains(j), "http") = 2 Then
                        'only http on subdomain is trusted
                        sHit = "O15 - ESC Trusted Zone: http://" & sSubDomains(j) & "." & sDomains(i) & " (HKLM)"
                        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                    End If
                Next j
            End If
            'list main domain as well, if applicable
            If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i), "*") = 2 Then
                'entire domain is trusted
                sHit = "O15 - ESC Trusted Zone: *." & sDomains(i) & " (HKLM)"
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i), "http") = 2 Then
                'only http on domain is trusted
                sHit = "O15 - ESC Trusted Zone: http://*." & sDomains(i) & " (HKLM)"
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
NextEscDomain2:
        Next i
    End If
    
    'enum all IP ranges
    sDomains = Split(RegEnumSubkeys(HKEY_CURRENT_USER, sZoneMapEscRanges), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            sIPRange = RegGetString(HKEY_CURRENT_USER, sZoneMapEscRanges & "\" & sDomains(i), ":Range")
            If Left(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscRanges & "\" & sDomains(i), "*") = 2 Then
                    'all protocols for this ip range is trusted
                    sHit = "O15 - ESC Trusted IP range: " & sIPRange
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
                If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscRanges & "\" & sDomains(i), "http") = 2 Then
                    'only http protocol for this ip range is trusted
                    sHit = "O15 - ESC Trusted IP range: http://" & sIPRange
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
            End If
        Next i
    End If
    
    'repeat for HKLM (ip ranges)
    sDomains = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sZoneMapEscRanges), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            sIPRange = RegGetString(HKEY_LOCAL_MACHINE, sZoneMapEscRanges & "\" & sDomains(i), ":Range")
            If Left(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscRanges & "\" & sDomains(i), "*") = 2 Then
                    'all protocols for this ip range is trusted
                    sHit = "O15 - ESC Trusted IP range: " & sIPRange & " (HKLM)"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
                If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscRanges & "\" & sDomains(i), "http") = 2 Then
                    'only http protocol for this ip range is trusted
                    sHit = "O15 - ESC Trusted IP range: http://" & sIPRange & " (HKLM)"
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
            End If
        Next i
    End If
    '=============================================================

    
    'check all ProtocolDefaults values
ZoneMapProtDefsHKCU:
    Dim sZoneNames$(), sProtVals$(), lProtZoneDefs&(5), lProtZones&(5)
    'sZoneNames = Split("MY COMPUTER|INTRANET|TRUSTED|INTERNET|RESTRICTED|UNKNOWN", "|")
    sZoneNames = Split("My Computer|Intranet|Trusted|Internet|Restricted|Unknown", "|")
    sProtVals = Split("@ivt|file|ftp|http|https|shell", "|")
    lProtZoneDefs(0) = 1
    lProtZoneDefs(1) = 3
    lProtZoneDefs(2) = 3
    lProtZoneDefs(3) = 3
    lProtZoneDefs(4) = 3
    lProtZoneDefs(5) = 0
    
    For i = 0 To 5
        lProtZones(i) = RegGetDword(HKEY_CURRENT_USER, sZoneMapProtDefs, sProtVals(i))
        If lProtZones(i) < 0 Or lProtZones(i) > 5 Then lProtZones(i) = 5
        If lProtZones(i) <> lProtZoneDefs(i) Then
            sHit = "O15 - ProtocolDefaults: '" & sProtVals(i) & "' protocol is in " & sZoneNames(lProtZones(i)) & " Zone, should be " & sZoneNames(lProtZoneDefs(i)) & " Zone"
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
    Next i
    
ZoneMapProtDefsHKLM:
    For i = 0 To 5
        lProtZones(i) = RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapProtDefs, sProtVals(i))
        If lProtZones(i) < 0 Or lProtZones(i) > 5 Then lProtZones(i) = 5
        If lProtZones(i) <> lProtZoneDefs(i) Then
            sHit = "O15 - ProtocolDefaults: '" & sProtVals(i) & "' protocol is in " & sZoneNames(lProtZones(i)) & " Zone, should be " & sZoneNames(lProtZoneDefs(i)) & " Zone (HKLM)"
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
    Next i
    
CleanUp:
    For i = 0 To UBound(sSafeRegDomains)
        sSafeRegDomains(i) = Crypt(sSafeRegDomains(i), sProgramVersion, True)
    Next i
End Sub

Public Sub FixOther15Item(sItem$)
    'O15 - Trusted Zone: free.aol.com (HKLM)
    'O15 - Trusted Zone: http://free.aol.com
    'O15 - Trusted IP range: 66.66.66.66 (HKLM)
    'O15 - Trusted IP range: http://66.66.66.*
    'O15 - ESC Trusted Zone: free.aol.com (HKLM)
    'O15 - ESC Trusted IP range: 66.66.66.66
    'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)
    '* other domains are now listed since 1.95.1 *
    '* retarded hijackers use wrong format for trusted sites - 1.99.2 *
    
    Dim lHive&, sKey1$, sKey2$, sKey3$, sValue$
    Dim sZoneMapDomains$, sZoneMapRanges$, sZoneMapProtDefs$
    Dim sZoneMapEscDomains$, sZoneMapEscRanges$
    Dim i%, sDummy$, vRanges As Variant
    On Error GoTo Error:
    sZoneMapDomains = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\"
    sZoneMapRanges = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\"
    sZoneMapEscDomains = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\"
    sZoneMapEscRanges = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscRanges\"
    sZoneMapProtDefs = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\ProtocolDefaults"
    
    If InStr(sItem, " (HKLM)") Then
        lHive = HKEY_LOCAL_MACHINE
    Else
        lHive = HKEY_CURRENT_USER
    End If
    
    If InStr(sItem, "http://") Then
        sValue = "http"
    Else
        sValue = "*"
    End If
    
    If InStr(sItem, "Trusted IP range:") > 0 Then GoTo IPRange:
    If InStr(sItem, "ProtocolDefaults") > 0 Then GoTo ProtDefs:
    
    'sKey1 = subdomain regkey   (e.g. aol.com\free)
    'sKey2 = root domain regkey (e.g. aol.com)
    
    'O15 : *.domain.com     -> domain.com regkey
    'O15 : sub.domain.com   -> domain.com\sub regkey
    'O15 : *.sub.domain.com -> domain.com\*.sub regkey (WTF)
    
    'strip domain from rest
    sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
    If InStr(sDummy, " (HKLM)") > 0 Then sDummy = Left(sDummy, InStr(sDummy, " (HKLM)") - 1)
    'strip protocol (if any) from domain
    If InStr(sDummy, "//") > 0 Then sDummy = Mid(sDummy, InStr(sDummy, "//") + 2)
    If InStr(sDummy, "*.") > 0 Then
        sDummy = Mid(sDummy, InStr(sDummy, "*.") + 2)
        'stupid 3rd case
        If InStr(sDummy, ".") <> InStrRev(sDummy, ".") Then sDummy = "*." & sDummy
    End If
    
    'sub.domain.com or domain.com
    'if domain is ip (1.1.1.1) watch out
    If InStr(sDummy, ".") = InStrRev(sDummy, ".") Or IsIPAddress(sDummy) Then
        'domain.com
        sKey2 = sDummy
        sKey1 = vbNullString
    Else
        'sub.domain.com
        i = InStrRev(sDummy, ".")
        i = InStrRev(sDummy, ".", i - 1)
        If DomainHasDoubleTLD(sDummy) Then i = InStrRev(sDummy, ".", i - 1)
        'If InStr(sDummy, ".co.uk") = Len(sDummy) - 5 Then i = InStrRev(sDummy, ".", i - 1)
        'If InStr(sDummy, ".ac.uk") = Len(sDummy) - 5 Then i = InStrRev(sDummy, ".", i - 1)
        sKey2 = Mid(sDummy, i + 1)
        sKey1 = sKey2 & "\" & Left(sDummy, i - 1)
        sKey3 = Mid(sDummy, 3)
    End If
    
    'relevant value should be deleted, and if no
    'other value is present, subkey as well.
    'if main key has no other subkeys, delete that also.
    If InStr(sItem, "ESC Trusted") = 0 Then
        If sKey1 = vbNullString Then
            RegDelVal lHive, sZoneMapDomains & sKey2, sValue
            If Not RegKeyHasValues(lHive, sZoneMapDomains & sKey2) Then
                RegDelKey lHive, sZoneMapDomains & sKey2
            End If
        Else
            RegDelVal lHive, sZoneMapDomains & sKey1, sValue
            If Not RegKeyHasValues(lHive, sZoneMapDomains & sKey1) Then
                RegDelKey lHive, sZoneMapDomains & sKey1
                If Not RegKeyHasSubKeys(lHive, sZoneMapDomains & sKey2) And _
                   Not RegKeyHasValues(lHive, sZoneMapDomains & sKey2) Then
                    RegDelKey lHive, sZoneMapDomains & sKey2
                End If
            End If
            '1.99.2 - fix for retarded hijackers like *.frame.crazywinnings.com
            RegDelVal lHive, sZoneMapDomains & sKey3, sValue
            If Not RegKeyHasValues(lHive, sZoneMapDomains & sKey3) Then
                RegDelKey lHive, sZoneMapDomains & sKey3
            End If
        End If
    Else '1.99.2: added EscDomains
        If sKey1 = vbNullString Then
            RegDelVal lHive, sZoneMapEscDomains & sKey2, sValue
            If Not RegKeyHasValues(lHive, sZoneMapEscDomains & sKey2) Then
                RegDelKey lHive, sZoneMapEscDomains & sKey2
            End If
        Else
            RegDelVal lHive, sZoneMapEscDomains & sKey1, sValue
            If Not RegKeyHasValues(lHive, sZoneMapEscDomains & sKey1) Then
                RegDelKey lHive, sZoneMapEscDomains & sKey1
                If Not RegKeyHasSubKeys(lHive, sZoneMapEscDomains & sKey2) And _
                   Not RegKeyHasValues(lHive, sZoneMapEscDomains & sKey2) Then
                    RegDelKey lHive, sZoneMapEscDomains & sKey2
                End If
            End If
            '1.99.2 - fix for retarded hijackers like *.frame.crazywinnings.com
            RegDelVal lHive, sZoneMapEscDomains & sKey3, sValue
            If Not RegKeyHasValues(lHive, sZoneMapEscDomains & sKey3) Then
                RegDelKey lHive, sZoneMapEscDomains & sKey3
            End If
        End If
    End If
    Exit Sub
    
IPRange:
    'O15 - Trusted IP range: 66.66.66.66 (HKLM)
    'O15 - Trusted IP range: http://66.66.66.*
    'O15 - ESC Trusted IP range: 66.66.66.66
    'enum subkeys of ZoneMap\Ranges, find key that holds IP range, kill it
    
    'strip IP range from rest
    sDummy = Mid(sItem, InStr(sItem, ":") + 2)
    If InStr(sDummy, " (HKLM)") > 0 Then sDummy = Left(sDummy, InStr(sDummy, " (HKLM)") - 1)
    If InStr(sDummy, "//") > 0 Then sDummy = Mid(sDummy, InStr(sDummy, "//") + 2)
    sKey2 = sDummy
    If InStr(sItem, "ESC Trusted") = 0 Then
        vRanges = Split(RegEnumSubkeys(lHive, sZoneMapRanges), "|")
        If UBound(vRanges) <> -1 Then
            For i = 0 To UBound(vRanges)
                sKey1 = RegGetString(lHive, sZoneMapRanges & "\" & vRanges(i), ":Range")
                If InStr(sKey1, sKey2) = 1 Then
                    RegDelKey lHive, sZoneMapRanges & "\" & vRanges(i)
                    Exit For
                End If
            Next i
        End If
    Else
        vRanges = Split(RegEnumSubkeys(lHive, sZoneMapEscRanges), "|")
        If UBound(vRanges) <> -1 Then
            For i = 0 To UBound(vRanges)
                sKey1 = RegGetString(lHive, sZoneMapEscRanges & "\" & vRanges(i), ":Range")
                If InStr(sKey1, sKey2) = 1 Then
                    RegDelKey lHive, sZoneMapEscRanges & "\" & vRanges(i)
                    Exit For
                End If
            Next i
        End If
    End If
    Exit Sub
    
ProtDefs:
    'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)
    sDummy = Mid(sItem, InStr(sItem, ": ") + 3)
    sDummy = Left(sDummy, InStr(sDummy, "'") - 1)
    If InStr(sItem, "(HKLM)") > 0 Then
        lHive = HKEY_LOCAL_MACHINE
    Else
        lHive = HKEY_CURRENT_USER
    End If
    Select Case sDummy
        Case "@ivt": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 1
        Case "file": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 3
        Case "ftp": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 3
        Case "http": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 3
        Case "https": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 3
        Case "shell": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 0
    End Select
    
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther15Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckNetscapeMozilla()
    Dim sDummy$, sNSVer$, sMailKey$, sPrefsJs$, sUserName$
    On Error GoTo Error:
        
    sUserName = GetUser
    
    If RegKeyExists(HKEY_CURRENT_USER, "Software\Netscape\Netscape Navigator\Main") Then
        'netscape 4.x is installed
        
        'get "popstatePath" - only way to find Users folder
        'I really hate Netscape
        sMailKey = "Software\Netscape\Netscape Navigator\biff\users"
        sDummy = RegGetFirstSubKey(HKEY_CURRENT_USER, sMailKey)
        sMailKey = sMailKey & "\" & sDummy & "\servers"
        sDummy = RegGetFirstSubKey(HKEY_CURRENT_USER, sMailKey)
        sMailKey = sMailKey & "\" & sDummy
        sDummy = RegGetString(HKEY_CURRENT_USER, sMailKey, "popstatePath")
        If sDummy <> vbNullString Then
            'cut off \mail\popstate.dat
            sDummy = Left(sDummy, InStrRev(sDummy, "\") - 6)
            sPrefsJs = sDummy & "\prefs.js"
            If FileExists(sPrefsJs) Then
                If FileLen(sPrefsJs) > 0 Then
                    Open sPrefsJs For Input As #1
                        Do
                            Line Input #1, sDummy
                            If InStr(sDummy, "user_pref(""browser.startup.homepage"",") > 0 Then
                                frmMain.lstResults.AddItem "N1 - Netscape 4: " & sDummy & " (" & sPrefsJs & ")"
                                Exit Do
                            End If
                        Loop Until EOF(1)
                    Close #1
                End If
            End If
        End If
    End If

    sDummy = vbNullString
    'moz/ns6/ns7 all use similar regkeys
    'moz uses \mozilla\currentversion or \seamonkey\currentversion
    'ns6 uses \netscape\netscape 6\currentversion
    'ns7 uses \netscape\currentversion or \netscape\netscape 6\currentversion
    'they all use the same place to store PREFS.JS though
    sDummy = RegGetString(HKEY_CURRENT_USER, "Software\Mozilla\Mozilla Firefox", "CurrentVersion")
    If sDummy = vbNullString Then sDummy = RegGetString(HKEY_CURRENT_USER, "Software\Netscape\Netscape 6", "CurrentVersion")
    If sDummy = vbNullString Then sDummy = RegGetString(HKEY_CURRENT_USER, "Software\Netscape\Netscape", "CurrentVersion")
    If sDummy <> vbNullString Then
        'mozilla, netscape 6 or netscape 7 is installed
        
        'sDummy is something like "1.2b" [moz],
        '"6.2.3 (en)" [ns6], or "7.0 (en)" [ns7]
        If Left(sDummy, 1) = "6" Then
            sNSVer = "N2 - Netscape 6: "
        ElseIf Left(sDummy, 1) = "7" Then
            sNSVer = "N3 - Netscape 7: "
        Else
            sNSVer = "N4 - Mozilla: "
        End If
        
        'prefs.js is stored in the insane location of
        '%APPLICATIONDATA%\Mozilla\Profiles\default\
        '     [random string].slt\prefs.js
        '%APPLICDATA% also varies per Windows version
        If Not bIsWinNT Then
            sPrefsJs = sWinDir & "\Application Data"
        Else
            sPrefsJs = Left(sWinDir, 2) & "\Documents and Settings\" & sUserName & "\Application Data"
        End If
        sPrefsJs = sPrefsJs & "\Mozilla\Profiles\default"
        sDummy = GetFirstSubFolder(sPrefsJs)
        sPrefsJs = sPrefsJs & "\" & sDummy & "\prefs.js"
        If FileExists(sPrefsJs) Then
            If FileLen(sPrefsJs) > 0 Then
                Open sPrefsJs For Input As #1
                    Do
                        Line Input #1, sDummy
                        If InStr(sDummy, "user_pref(""browser.startup.homepage"",") > 0 Then
                            frmMain.lstResults.AddItem sNSVer & sDummy & " (" & sPrefsJs & ")"
                            Exit Do
                        End If
                    Loop Until EOF(1)
                Close #1
                Open sPrefsJs For Input As #1
                    Do
                        Line Input #1, sDummy
                        If InStr(sDummy, "user_pref(""browser.search.defaultengine"",") > 0 Then
                            frmMain.lstResults.AddItem sNSVer & sDummy & " (" & sPrefsJs & ")"
                            Exit Do
                        End If
                    Loop Until EOF(1)
                Close #1
            End If
        End If
    End If
    Exit Sub
    
Error:
    Close #1
    ErrorMsg "modMain_CheckNetscapeMozilla", Err.Number, Err.Description
End Sub

Public Sub FixNetscapeMozilla(sItem$)
    'N1 - Netscape 4: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
    'N2 - Netscape 6: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
    'N3 - Netscape 7: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
    'N4 - Mozilla: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
    '               user_pref("browser.search.defaultengine", "http://url"); (c:\..\prefs.js)
    
    Dim sPrefsJs$, sDummy$
    On Error GoTo Error:
    sPrefsJs = Mid(sItem, InStrRev(sItem, "(") + 1)
    sPrefsJs = Left(sPrefsJs, Len(sPrefsJs) - 1)
    If FileExists(sPrefsJs) Then
        Open sPrefsJs For Input As #1
        Open sPrefsJs & ".new" For Output As #2
            Do
                Line Input #1, sDummy
                If InStr(sDummy, "user_pref(""browser.startup.homepage"",") > 0 And _
                   InStr(sItem, "user_pref(""browser.startup.homepage"",") > 0 Then
                    Print #2, "user_pref(""browser.startup.homepage"", ""http://home.netscape.com/"");"
                ElseIf InStr(sDummy, "user_pref(""browser.search.defaultengine"",") > 0 And _
                   InStr(sItem, "user_pref(""browser.search.defaultengine"",") > 0 Then
                    Print #2, "user_pref(""browser.search.defaultengine"", ""http://www.google.com/"");"
                Else
                    Print #2, sDummy
                End If
            Loop Until EOF(1)
        Close #1
        Close #2
        DeleteFile sPrefsJs
        Name sPrefsJs & ".new" As sPrefsJs
    End If
    Exit Sub
    
Error:
    Close #1
    Close #2
    ErrorMsg "modMain_FixNetscapeMozilla", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckRegistry3Item()
    Dim sURLHook$, hKey&, i&, sName$, uData() As Byte
    Dim sHit$, sCLSID$, sFile$
    sURLHook = "Software\Microsoft\Internet Explorer\URLSearchHooks"
    If RegOpenKeyEx(HKEY_CURRENT_USER, sURLHook, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        sCLSID = String(lEnumBufSize, 0)
        ReDim uData(lEnumBufSize)
        If RegEnumValue(hKey, 0, sCLSID, Len(sCLSID), 0, ByVal 0, uData(0), UBound(uData)) <> 0 Then
            'default URLSearchHook is missing!
            sHit = "R3 - Default URLSearchHook is missing"
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            RegCloseKey hKey
            Exit Sub
        End If
        
        Do
            sCLSID = Left(sCLSID, InStr(sCLSID, Chr(0)) - 1)
            sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, "")
            If sCLSID <> "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}" Then
                'found a new urlsearchhook!
                If sName = vbNullString Then sName = "(no name)"
                sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", "")
                If sFile = vbNullString Then sFile = "(no file)"
                If sFile <> "(no file)" And Not FileExists(sFile) Then sFile = sFile & " (file missing)"
                
                sHit = "R3 - URLSearchHook: " & sName & " - " & sCLSID & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            
            i = i + 1
            sCLSID = String(lEnumBufSize, 0)
            ReDim uData(lEnumBufSize)
        Loop Until RegEnumValue(hKey, i, sCLSID, Len(sCLSID), 0, ByVal 0, uData(0), UBound(uData)) <> 0
        RegCloseKey hKey
    Else
        'default URLSearchHook is missing!
        sHit = "R3 - Default URLSearchHook is missing"
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modMain_CheckRegistry3Item", Err.Number, Err.Description
End Sub

Public Sub FixRegistry3Item(sItem$)
    'R3 - Shitty search hook - {00000000} - c:\windows\bho.dll"
    'R3 - Default URLSearchHook is missing
    Dim sDummy$
    On Error GoTo Error:
    If sItem = "R3 - Default URLSearchHook is missing" Then
        RegCreateKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks"
        RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks", "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}", ""
        Exit Sub
    End If
    
    sDummy = Mid(sItem, InStr(6, sItem, " - ") + 3)
    sDummy = Left(sDummy, InStr(sDummy, " - ") - 1)
    
    'If InStr(sItem, "- _{") > 0 Then
    '    sDummy = Mid(sItem, InStr(sItem, "- _{") + 2)
    'Else
    '    sDummy = Mid(sItem, InStr(sItem, "- {") + 2)
    'End If
    'sDummy = Left(sDummy, InStrRev(sDummy, " - ") - 1)
    
    RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks", sDummy
    'just in case
    RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks", "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}", ""
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixRegistry3Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther16Item()
    'O16 - Downloaded Program Files
    Dim sDPFKey$, sName$, sFriendlyName$, sCodeBase$, i&, hKey&, sHit$
    
    'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ActiveXCache
    'is location of actual %WINDIR%\DPF\ folder
    sDPFKey = "Software\Microsoft\Code Store Database\Distribution Units"
    On Error GoTo Error:
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sDPFKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        'key doesn't exist
        Exit Sub
    End If
    
    sName = String(255, 0)
    If RegEnumKeyEx(hKey, 0, sName, 255, 0, vbNullString, 0, ByVal 0) <> 0 Then
        'no subkeys
        RegCloseKey hKey
        Exit Sub
    End If
    
    Do
        sName = Left(sName, InStr(sName, Chr(0)) - 1)
        If Left(sName, 1) = "{" And Right(sName, 1) = "}" Then
            sFriendlyName = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sName, "")
            If sFriendlyName = vbNullString Then
                sFriendlyName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName, "")
            End If
        End If
        sCodeBase = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sName & "\DownloadInformation", "CODEBASE")
        
        If InStr(sCodeBase, "http://java.sun.com") <> 1 And _
           InStr(sCodeBase, "http://www.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://webresponse.one.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://rtc.webresponse.one.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://office.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://officeupdate.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://protect.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://dql.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://codecs.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://download.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://windowsupdate.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://v4.windowsupdate.microsoft.com") <> 1 And _
           InStr(sCodeBase, "http://download.macromedia.com") <> 1 And _
           InStr(sCodeBase, "http://fpdownload.macromedia.com") <> 1 And _
           InStr(sCodeBase, "http://active.macromedia.com") <> 1 And _
           InStr(sCodeBase, "http://www.apple.com") <> 1 And _
           InStr(sCodeBase, "http://http://security.symantec.com") <> 1 And _
           InStr(sCodeBase, "http://download.yahoo.com") <> 1 And _
           InStr(sName, "Microsoft XML Parser") = 0 And _
           InStr(sName, "Java Classes") = 0 And _
           InStr(sName, "Classes for Java") = 0 And _
           InStr(sName, "Java Runtime Environment") = 0 Or _
           bIgnoreAllWhitelists Then
           
            sHit = "O16 - DPF: " & sName & IIf(sFriendlyName <> vbNullString, " (" & sFriendlyName & ")", "") & " - " & sCodeBase
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        
        i = i + 1
        sName = String(255, 0)
        sFriendlyName = vbNullString
    Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, 0, ByVal 0) <> 0
    RegCloseKey hKey
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modMain_CheckOther16Item", Err.Number, Err.Description
End Sub

Public Sub FixOther16Item(sItem$)
    'O16 - DPF: {0000000} (shit toolbar) - http://bla.com/bla.dll
    'O16 - DPF: Plugin - http://bla.com/bla.dll
    
    Dim sDPFKey$, hKey&, sDummy$, sName$, sOSD$, sINF$, sInProcServer32$
    On Error GoTo Error:
    sDummy = Mid(sItem, 12)
    If Left(sDummy, 1) = "{" Then
        sName = Left(sDummy, InStr(sDummy, "}"))
    Else
        sName = Left(sDummy, InStr(sDummy, " - ") - 1)
        'experimental - bugfix for when item is
        'O16 - DPF: sName (sFriendlyName) - [file]
        'WHICH IS NOT EVEN POSSIBLE!!!
        'WHERE THE HELL DID THIS BUG CAME FROM????
        If InStr(sName, " (") > 0 Then
            sName = Left(sName, InStr(sName, " (") - 1)
        End If
    End If
    sDPFKey = "Software\Microsoft\Code Store Database\Distribution Units\" & sName
    
    If Not RegKeyExists(HKEY_LOCAL_MACHINE, sDPFKey) Then
        'unable to find that key
        'MsgBox "Could not delete '" & sItem & "' because it doesn't exist anymore.", vbExclamation
        Exit Sub
    End If
    
    'a DPF object can consist of:
    '* DPF regkey           -> sDPFKey
    '* CLSID regkey         -> CLSID\ & sName
    '* OSD file             -> sOSD = RegGetString
    '* INF file             -> sINF = RegGetString
    '* InProcServer32 file  -> sIPS = RegGetString
    
    sOSD = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\DownloadInformation", "OSD")
    sINF = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\DownloadInformation", "INF")
    If Left(sName, 1) = "{" Then
        sInProcServer32 = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName & "\InProcServer32", "")
        'maybe the file error is caused by this line?
        On Error Resume Next
        Shell "regsvr32 /u /s """ & sInProcServer32 & """"
        DoEvents
        On Error GoTo Error:
    End If
    
    'RegDelSubKeys HKEY_LOCAL_MACHINE, sDPFKey & "\Contains"
    'RegDelSubKeys HKEY_LOCAL_MACHINE, sDPFKey
    RegDelKey HKEY_LOCAL_MACHINE, sDPFKey
    If Left(sName, 1) = "{" Then
        'RegDelSubKeys HKEY_CLASSES_ROOT, "CLSID\" & sName
        RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & sName
    End If
    
    On Error Resume Next
    DeleteFile sInProcServer32
    DeleteFile sOSD
    DeleteFile sINF
    On Error GoTo Error:
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther16Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther17Item()
    'check 'domain' and 'domainname' values in:
    'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters
    'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\*
    'HKLM\Software\Microsoft\Windows\CurrentVersion\Telephony
    'HKLM\System\CurrentControlSet\Services\VxD\MSTCP
    'and all values in other ControlSet's as well
    '
    'new one from UltimateSearch: value 'SearchList' in
    'HKLM\System\CurrentControlSet\Services\VxD\MSTCP
    '
    'just in case: NameServer as well, CoolWebSearch
    'may be using this
    
    Dim hKey&, i&, j&, sName$, sDomain$, sKeyDomain$()
    Dim sHit$, sSearchList$, sNameServer$
    On Error GoTo Error:
    ReDim sKeyDomain(0 To 3)
    sKeyDomain(0) = "System\CurrentControlSet\Services\Tcpip\Parameters"
    sKeyDomain(1) = "System\CurrentControlSet\Services\VxD\MSTCP"
    sKeyDomain(2) = "Software\Microsoft\Windows\CurrentVersion\Telephony"
    'sKeyDomain(3) is used below, for CS1 etc
    
    'HKLM\System\CCS\Services\Tcpip\Parameters,Domain
    sDomain = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(0), "Domain")
    If sDomain <> vbNullString Then
        sHit = "O17 - HKLM\System\CCS\Services\Tcpip\Parameters: Domain = " & sDomain
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    'HKLM\System\CCS\Services\Tcpip\Parameters,DomainName
    sDomain = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(0), "DomainName")
    If sDomain <> vbNullString Then
        sHit = "O17 - HKLM\System\CCS\Services\Tcpip\Parameters: DomainName = " & sDomain
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'HKLM\System\CCS\Services\VxD\MSTCP,Domain
    sDomain = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(1), "Domain")
    If sDomain <> vbNullString Then
        sHit = "O17 - HKLM\System\CCS\Services\VxD\MSTCP: Domain = " & sDomain
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    'HKLM\System\CCS\Services\VxD\MSTCP,DomainName
    sDomain = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(1), "DomainName")
    If sDomain <> vbNullString Then
        sHit = "O17 - HKLM\System\CCS\Services\VxD\MSTCP: DomainName = " & sDomain
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'HKLM\Software\MS\Windows\CurVer\Telephony,Domain
    sDomain = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(2), "Domain")
    If sDomain <> vbNullString Then
        sHit = "O17 - HKLM\Software\..\Telephony: Domain = " & sDomain
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    'HKLM\Software\MS\Windows\CurVer\Telephony,DomainName
    sDomain = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(2), "DomainName")
    If sDomain <> vbNullString Then
        sHit = "O17 - HKLM\Software\..\Telephony: DomainName = " & sDomain
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'HKLM\System\CCS\Services\Tcpip\..\ subkeys
    RegOpenKeyEx HKEY_LOCAL_MACHINE, sKeyDomain(0) & "\Interfaces", 0, KEY_ENUMERATE_SUB_KEYS, hKey
    sName = String(255, 0)
    If RegEnumKeyEx(hKey, 0, sName, 255, 0, vbNullString, 0, ByVal 0) = 0 Then
        Do
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
            
            'HKLM\System\CCS\Services\Tcpip\Param\Int\*,Domain
            sDomain = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & sName, "Domain")
            If sDomain <> vbNullString Then
                sHit = "O17 - HKLM\System\CCS\Services\Tcpip\..\" & sName & ": Domain = " & sDomain
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            'HKLM\System\CCS\Services\Tcpip\Param\Int\*,DomainName
            sDomain = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & sName, "DomainName")
            If sDomain <> vbNullString Then
                sHit = "O17 - HKLM\System\CCS\Services\Tcpip\..\" & sName & ": DomainName = " & sDomain
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            
            'HKLM\System\CCS\Services\Tcpip\Param\Int\*,SearchList
            sSearchList = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & sName, "SearchList")
            If sSearchList <> vbNullString Then
                sHit = "O17 - HKLM\System\CCS\Services\Tcpip\..\" & sName & ": SearchList = " & sSearchList
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            
            'HKLM\System\CCS\Services\Tcpip\Param\Int\*,NameServer
            sNameServer = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & sName, "NameServer")
            If sNameServer <> vbNullString Then
                sHit = "O17 - HKLM\System\CCS\Services\Tcpip\..\" & sName & ": NameServer = " & sNameServer
                If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
            End If
            
            sName = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    'HKLM\System\CS[1-999]\Services\Tcpip\Parameters
    'HKLM\System\CS[1-999]\Services\Tcpip\Parameters\Interfaces\*
    For j = 1 To 999
        If Not RegKeyExists(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000")) Then Exit For
        
        'HKLM\System\CS*\Services\Tcpip\Parameters,Domain
        sDomain = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters", "Domain")
        If sDomain <> vbNullString Then
            sHit = "O17 - HKLM\System\CS" & j & "\Services\Tcpip\Parameters: Domain = " & sDomain
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        'HKLM\System\CS*\Services\Tcpip\Parameters,DomainName
        sDomain = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters", "DomainName")
        If sDomain <> vbNullString Then
            sHit = "O17 - HKLM\System\CS" & j & "\Services\Tcpip\Parameters: DomainName = " & sDomain
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        
        'HKLM\System\CS*\Services\VxD\MSTCP,Domain
        sDomain = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\VxD\MSTCP", "Domain")
        If sDomain <> vbNullString Then
            sHit = "O17 - HKLM\System\CS" & j & "\Services\VxD\MSTCP: Domain = " & sDomain
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        'HKLM\System\CS*\Services\VxD\MSTCP,DomainName
        sDomain = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\VxD\MSTCP", "DomainName")
        If sDomain <> vbNullString Then
            sHit = "O17 - HKLM\System\CS" & j & "\Services\VxD\MSTCP: DomainName = " & sDomain
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        
        'HKLM\System\CS*\Services\Tcpip\Parameters,SearchList
        sSearchList = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters", "SearchList")
        If sSearchList <> vbNullString Then
            sHit = "O17 - HKLM\System\CS" & j & "\Services\Tcpip\Parameters: SearchList = " & sSearchList
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        
        'HKLM\System\CS*\Services\VxD\MSTCP,SearchList
        sSearchList = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\VxD\MSTCP", "SearchList")
        If sSearchList <> vbNullString Then
            sHit = "O17 - HKLM\System\CS" & j & "\Services\VxD\MSTCP: SearchList = " & sSearchList
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        
        'HKLM\System\CS*\Services\Tcpip\Parameters,NameServer
        sNameServer = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters", "NameServer")
        If sNameServer <> vbNullString Then
            sHit = "O17 - HKLM\System\CS" & j & "\Services\Tcpip\Parameters: NameServer = " & sNameServer
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        
        'HKLM\System\CS*\Services\VxD\MSTCP,NameServer
        sNameServer = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\VxD\MSTCP", "NameServer")
        If sNameServer <> vbNullString Then
            sHit = "O17 - HKLM\System\CS" & j & "\Services\VxD\MSTCP: NameServer = " & sNameServer
            If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
        End If
        
        
        'HKLM\System\CS*\Services\Tcpip\Parameters\Interfaces\*
        hKey = 0
        RegOpenKeyEx HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters\Interfaces", 0, KEY_ENUMERATE_SUB_KEYS, hKey
        sName = String(255, 0)
        If RegEnumKeyEx(hKey, 0, sName, 255, 0, vbNullString, 0, ByVal 0) = 0 Then
            Do
                sName = Left(sName, InStr(sName, Chr(0)) - 1)
                
                'HKLM\System\..\Interfaces\*,Domain
                sDomain = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters\Interfaces\" & sName, "Domain")
                If sDomain <> vbNullString Then
                    sHit = "O17 - HKLM\System\CS" & j & "\Services\Tcpip\..\" & sName & ": Domain = " & sDomain
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
                'HKLM\System\..\Interfaces\*,DomainName
                sDomain = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters\Interfaces\" & sName, "DomainName")
                If sDomain <> vbNullString Then
                    sHit = "O17 - HKLM\System\CS" & j & "\Services\Tcpip\..\" & sName & ": DomainName = " & sDomain
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
                
                'HKLM\System\..\Interfaces\*,SearchList
                sSearchList = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters\Interfaces\" & sName, "SearchList")
                If sSearchList <> vbNullString Then
                    sHit = "O17 - HKLM\System\CS" & j & "\Services\Tcpip\..\" & sName & ": SearchList = " & sSearchList
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
                
                'HKLM\System\..\Interfaces\*,SearchList
                sNameServer = RegGetString(HKEY_LOCAL_MACHINE, "System\ControlSet" & Format(j, "000") & "\Services\Tcpip\Parameters\Interfaces\" & sName, "NameServer")
                If sNameServer <> vbNullString Then
                    sHit = "O17 - HKLM\System\CS" & j & "\Services\Tcpip\..\" & sName & ": NameServer = " & sNameServer
                    If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
                End If
                
                sName = String(255, 0)
                i = i + 1
            Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, 0, ByVal 0) <> 0
            RegCloseKey hKey
        End If
    Next j
    
    'new one from UltimateSearch!
    'HKLM\System\CCS\Services\VxD\MSTCP,SearchList
    sSearchList = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(1), "SearchList")
    If sSearchList <> vbNullString Then
        sHit = "O17 - HKLM\System\CCS\Services\VxD\MSTCP: SearchList = " & sSearchList
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'HKLM\System\CCS\Services\Tcpip\Parameters,SearchList
    sSearchList = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(0), "SearchList")
    If sSearchList <> vbNullString Then
        sHit = "O17 - HKLM\System\CCS\Services\Tcpip\Parameters: SearchList = " & sSearchList
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'HKLM\System\CCS\Services\VxD\MSTCP,SearchList
    sNameServer = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(1), "NameServer")
    If sNameServer <> vbNullString Then
        sHit = "O17 - HKLM\System\CCS\Services\VxD\MSTCP: NameServer = " & sNameServer
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'HKLM\System\CCS\Services\Tcpip\Parameters,NameServer
    sNameServer = RegGetString(HKEY_LOCAL_MACHINE, sKeyDomain(0), "NameServer")
    If sNameServer <> vbNullString Then
        sHit = "O17 - HKLM\System\CCS\Services\Tcpip\Parameters: NameServer = " & sNameServer
        If Not IsOnIgnoreList(sHit) Then frmMain.lstResults.AddItem sHit
    End If
    
    'for the list of SearchList in:
    'HKLM\System\CCS\Services\Tcpip\Parameters\Interfaces\*
    'HKLM\System\CS[1-999]\Services\Tcpip\Parameters
    'HKLM\System\CS[1-999]\Services\Tcpip\Parameters\Interfaces\*
    'see the loop above
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modMain_CheckOther17Item", Err.Number, Err.Description
End Sub

Public Sub FixOther17Item(sItem$)
    'O17 - Domain hijack
    'O17 - HKLM\System\CCS\Services\VxD\MSTCP: Domain[Name] = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\Parameters: Domain[Name] = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\..\{0000}: Domain[Name] = blah
    '                  CS1
    '                  CS2
    '                  ...
    'O17 - HKLM\Software\..\Telephony: SearchList = blah
    'O17 - HKLM\System\CCS\Services\VxD\MSTCP: SearchList = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\Parameters: SearchList = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\..\{0000}: SearchList = blah
    '                  CS1
    '                  CS2
    '                  ...
    'ditto for NameServer
    
    Dim sKey$, sValue$, sDummy$, i%, j%
    On Error GoTo Error:
    sDummy = Mid(sItem, 7)
    sKey = Left(sDummy, InStr(sDummy, ":") - 1)
    If InStr(sKey, "\..\") > 0 Then
        'expand \..\
        If InStr(sKey, "Telephony") > 0 Then
            sKey = Replace(sKey, "\..\", "\Microsoft\Windows\CurrentVersion\", , 1)
        End If
        If InStr(sKey, "Tcpip") > 0 Then
            sKey = Replace(sKey, "\..\", "\Parameters\Interfaces\", , 1)
        End If
    End If
    If InStr(sKey, "\CCS\") > 0 Then
        sKey = Replace(sKey, "\CCS\", "\CurrentControlSet\", , 1)
    End If
    
    'expand CCS/CS1/CS2/..
    i = InStr(sKey, "\CS")
    If i > 0 And i < 20 Then
        '<20 just in case a domain with \cs comes up
        '\CS1\   or   \CS11\
        j = InStr(i + 3, sKey, "\") - i - 3
        sKey = Replace(sKey, "\CS", "\ControlSet" & String(3 - j, "0"), , 1)
    End If
    
    'get value
    If InStr(sItem, ": DomainName ") > 0 Then
        sValue = "DomainName"
    End If
    If InStr(sItem, ": Domain ") > 0 Then
        sValue = "Domain"
    End If
    If InStr(sItem, ": SearchList ") > 0 Then
        sValue = "SearchList"
    End If
    If InStr(sItem, ": NameServer ") > 0 Then
        sValue = "NameServer"
    End If
    
    'delete the shit!
    'don't need to get root key - it's always HKLM
    sKey = Mid(sKey, 6)
    RegDelVal HKEY_LOCAL_MACHINE, sKey, sValue
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther17Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther18Item()
    'enumerate everything in HKCR\Protocols\Handler
    'enumerate everything in HKCR\Protocols\Filters (section 2)
    
    Dim hKey&, i&, sName$, sCLSID$, sFile$, sHit$
    On Error GoTo Error:
    
    If RegOpenKeyEx(HKEY_CLASSES_ROOT, "Protocols\Handler", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        'key not found
        GoTo Filters:
    End If
    
    sName = String(255, 0)
    If RegEnumKeyEx(hKey, 0, sName, 255, 0, vbNullString, 0, ByVal 0) <> 0 Then
        'no subkeys
        RegCloseKey hKey
        GoTo Filters:
    End If
    
    'decrypt stuff
    For i = 0 To UBound(sSafeProtocols)
        If sSafeProtocols(i) = vbNullString Then Exit For
        sSafeProtocols(i) = Crypt(sSafeProtocols(i), sProgramVersion)
    Next i
    
    i = 0
    Do
        sName = TrimNull(sName)
        sCLSID = UCase(RegGetString(HKEY_CLASSES_ROOT, "Protocols\Handler\" & sName, "CLSID"))
        If sCLSID = vbNullString Then sCLSID = "(no CLSID)"
        If sCLSID <> "(no CLSID)" Then
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", "")
            sFile = Replace(sFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
            If sFile = vbNullString Then
                sFile = "(no file)"
            Else
                If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
            End If
        Else
            sFile = "(no file)"
        End If
        
        'for each protocol, check if name is on safe list
        If InStr(1, Join(sSafeProtocols, vbCrLf), sName, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
            sHit = "O18 - Protocol: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                frmMain.lstResults.AddItem sHit
            End If
        Else
            'and if so, check if CLSID is also on safe list
            '(no hijacker would hijack a protocol by
            'changing the CLSID to another safe one, right?)
            If InStr(1, Join(sSafeProtocols, vbCrLf), sCLSID, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                If sCLSID <> "(no CLSID)" Then
                     sHit = "O18 - Protocol hijack: " & sName & " - " & sCLSID
                     If Not IsOnIgnoreList(sHit) Then
                         If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                         frmMain.lstResults.AddItem sHit
                     End If
                End If
            End If
        End If
        
        sName = String(255, 0)
        i = i + 1
    Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, 0, ByVal 0) <> 0
    RegCloseKey hKey
    
    '-------------------
Filters:
    
    hKey = 0
    i = 0
    sCLSID = vbNullString
    sFile = vbNullString
    If RegOpenKeyEx(HKEY_CLASSES_ROOT, "PROTOCOLS\Filter", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        Exit Sub
    End If
    sName = String(255, 0)
    If RegEnumKeyEx(hKey, 0, sName, 255, 0, vbNullString, 0, ByVal 0) <> 0 Then
        RegCloseKey hKey
        GoTo CleanUp:
    End If
    
    Do
        sName = TrimNull(sName)
        sCLSID = RegGetString(HKEY_CLASSES_ROOT, "PROTOCOLS\Filter\" & sName, "CLSID")
        If sCLSID = vbNullString Then
            sCLSID = "(no CLSID)"
            sFile = "(no file)"
        Else
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", "")
            sFile = Replace(sFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
            If sFile = vbNullString Then sFile = "(no file)"
        End If
        
        If InStr(1, Join(sSafeFilters, vbCrLf), sName, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
            'add to results list
            sHit = "O18 - Filter: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                frmMain.lstResults.AddItem sHit
            End If
        Else
            If InStr(1, Join(sSafeFilters, vbCrLf), sCLSID, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                'add to results list
                sHit = "O18 - Filter hijack: " & sName & " - " & sCLSID & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        End If
        
        sName = String(255, 0)
        i = i + 1
    Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, 0, ByVal 0) <> 0
    RegCloseKey hKey

    '---------------------
CleanUp:
    're-encrypt stuff
    For i = 0 To UBound(sSafeProtocols)
        If sSafeProtocols(i) = vbNullString Then Exit For
        sSafeProtocols(i) = Crypt(sSafeProtocols(i), sProgramVersion, True)
    Next i
    
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modMain_CheckOther18Item", Err.Number, Err.Description
    GoTo CleanUp
End Sub

Public Sub FixOther18Item(sItem$)
    'O18 - Protocol: cn
    
    Dim sDummy$, i%, sCLSID$ ', sProtCLSIDs$()
    On Error GoTo Error:
    If InStr(sItem, "Filter: ") > 0 Then GoTo FixFilter:
    
    'decrypt stuff
    For i = 0 To UBound(sSafeProtocols)
        If sSafeProtocols(i) = vbNullString Then Exit For
        sSafeProtocols(i) = Crypt(sSafeProtocols(i), sProgramVersion)
    Next i
    
    'get protocol name
    sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
    sDummy = Left(sDummy, InStr(sDummy, " - ") - 1)
    
    If InStr(sItem, "Protocol hijack: ") > 0 Then GoTo FixProtHijack:
    
    If InStr(sItem, "(no CLSID)") = 0 Then
        'RegDelSubKeys HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy
        RegDelKey HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy
    End If
    
    're-encrypt stuff
    For i = 0 To UBound(sSafeProtocols)
        If sSafeProtocols(i) = vbNullString Then Exit For
        sSafeProtocols(i) = Crypt(sSafeProtocols(i), sProgramVersion, True)
    Next i
    Exit Sub
    
FixProtHijack:
    For i = 0 To UBound(sSafeProtocols)
        'find CLSID for protocol name
        If sSafeProtocols(i) = vbNullString Then Exit For
        If InStr(1, sSafeProtocols(i), sDummy) > 0 Then
            sCLSID = Mid(sSafeProtocols(i), InStr(sSafeProtocols(i), "|") + 1)
            Exit For
        End If
    Next i
    RegSetStringVal HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy, "CLSID", sCLSID
    
    're-encrypt stuff
    For i = 0 To UBound(sSafeProtocols)
        If sSafeProtocols(i) = vbNullString Then Exit For
        sSafeProtocols(i) = Crypt(sSafeProtocols(i), sProgramVersion, True)
    Next i
    Exit Sub
    
FixFilter:
    'O18 - Filter: text/blah - {0} - c:\file.dll
    sDummy = Mid(sItem, InStr(sItem, ": ") + 2)
    'why the hell did I use InstrRev here first? bugfix 1.98.1
    sDummy = Left(sDummy, InStr(sDummy, " - ") - 1)
    
    If InStr(sItem, "Filter hijack: ") > 0 Then GoTo FixFilterHijack:
    
    RegDelKey HKEY_CLASSES_ROOT, "PROTOCOLS\Filter\" & sDummy
    Exit Sub
    
FixFilterHijack:
    For i = 0 To UBound(sSafeFilters)
        If sSafeFilters(i) = vbNullString Then Exit For
        If InStr(1, sSafeFilters(i), sDummy) > 0 Then
            sCLSID = Mid(sSafeFilters(i), InStr(sSafeFilters(i), "|") + 1)
            Exit For
        End If
    Next i
    RegSetStringVal HKEY_CLASSES_ROOT, "PROTOCOLS\Filter\" & sDummy, "CLSID", sCLSID
    Exit Sub

Error:
    ErrorMsg "modMain_FixOther18Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther19Item()
    On Error GoTo Error:
    'HKCU\Software\Microsoft\Internet Explorer\Styles,Use My Stylesheet
    'HKCU\Software\Microsoft\Internet Explorer\Styles,User Stylesheet
    'this hijack doesn't work for HKLM
    
    Dim lUseMySS&, sUserSS$, sHit$
    lUseMySS = RegGetDword(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Styles", "Use My Stylesheet")
    sUserSS = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Styles", "User Stylesheet")
    If Not FileExists(sUserSS) Then sUserSS = sUserSS & " (file missing)"
    If lUseMySS = 1 And sUserSS <> vbNullString Then
        sHit = "O19 - User stylesheet: " & sUserSS
        If Not IsOnIgnoreList(sHit) Then
            'md5 doesn't seem useful here
            'If bMD5 Then sHit = sHit & getfilemd5(sUserSS)
            frmMain.lstResults.AddItem sHit
        End If
    End If
    
    lUseMySS = RegGetDword(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Styles", "Use My Stylesheet")
    sUserSS = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Styles", "User Stylesheet")
    If Not FileExists(sUserSS) Then sUserSS = sUserSS & " (file missing)"
    If lUseMySS = 1 And sUserSS <> vbNullString Then
        sHit = "O19 - User stylesheet: " & sUserSS & " (HKLM)"
        If Not IsOnIgnoreList(sHit) Then
            frmMain.lstResults.AddItem sHit
        End If
    End If
    
    Exit Sub
Error:
    ErrorMsg "modMain_CheckOther19Item", Err.Number, Err.Description
End Sub

Public Sub FixOther19Item(sItem$)
    On Error GoTo Error:
    'O19 - User stylesheet: c:\file.css (file missing)
    
    If InStr(sItem, " (HKLM)") = 0 Then
        RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Styles", "Use My Stylesheet"
        RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Styles", "User Stylesheet"
    Else
        RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Styles", "Use My Stylesheet"
        RegDelVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Styles", "User Stylesheet"
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther19Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther20Item()
    'appinit_dlls + winlogon notify
    Dim sAppInit$, sFile$, sHit$
    sAppInit = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
    sFile = RegGetString(HKEY_LOCAL_MACHINE, sAppInit, "AppInit_DLLs")
    If sFile <> vbNullString Then
        sFile = Replace(sFile, Chr(0), "|")
        If InStr(1, sSafeAppInit, sFile, vbTextCompare) = 0 Or _
           bIgnoreAllWhitelists Then
            'item is not on whitelist
            sHit = "O20 - AppInit_DLLs: " & sFile
            
            If bIgnoreAllWhitelists = True Then
                frmMain.lstResults.AddItem sHit
            ElseIf Not IsOnIgnoreList(sHit) Then
                 frmMain.lstResults.AddItem sHit
            End If
        End If
    End If
    
    Dim sSubKeys$(), i&, sWinlogon$, ss$
    sWinlogon = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify"
    sSubKeys = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sWinlogon), "|")
    If UBound(sSubKeys) <> -1 Then
        For i = 0 To UBound(sSubKeys)
            If InStr(1, "*" & sSafeWinlogonNotify & "*", "*" & sSubKeys(i) & "*", vbTextCompare) = 0 Then
                sFile = RegGetString(HKEY_LOCAL_MACHINE, sWinlogon & "\" & sSubKeys(i), "DllName")
                
                If Len(sFile) = 0 Then
                    sFile = "Invalid registry found"
                Else
                    If StrComp(Mid(sFile, 1, 1), "\", vbTextCompare) = 0 Then
                        If FileExists(sWinDir & "\" & sFile) Then sFile = sWinDir & "\" & sFile
                        If FileExists(sWinSysDir & "\" & sFile) Then sFile = sWinSysDir & "\" & sFile
                    End If
                    
                    sFile = NormalizePath(sFile)
                    
                    If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
                    If FileExists(sFile) And bMD5 Then
                        sFile = sFile & GetFileFromAutostart(sFile)
                    End If
                    
'                    If InStr(1, sFile, "%", vbTextCompare) = 1 Then
'                       sFile = "Suspicious registry value"
'                  End If
                  
                End If
                
                sHit = "O20 - Winlogon Notify: " & sSubKeys(i) & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        Next i
    End If
End Sub

Public Sub FixOther20Item(sItem$)
    'O20 - AppInit_DLLs: file.dll
    'O20 - Winlogon Notify: bladibla - c:\file.dll
    'to do:
    '* clear appinit regval (don't delete it)
    '* kill regkey (for winlogon notify)
    Dim sAppInit$, sNotify$
    On Error GoTo Error:
    If InStr(sItem, "AppInit_DLLs") > 0 Then
        sAppInit = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
        RegSetStringVal HKEY_LOCAL_MACHINE, sAppInit, "AppInit_DLLs", ""
    Else
        sNotify = Mid(sItem, InStr(sItem, ":") + 2)
        sNotify = Left(sNotify, InStr(sNotify, " - ") - 1)
        RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify\" & sNotify
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther20Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther21Item()
    'shellserviceobjectdelayload
    Dim sSSODL$, sHit$, sFile$, j&, bOnWhiteList As Boolean
    Dim hKey&, i&, sName$, lNameLen&, sCLSID$, uData() As Byte, lDataLen&
    sSSODL = "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sSSODL, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        'key doesn't exist, or failed to open
        Exit Sub
    End If
    
    lNameLen = lEnumBufSize
    sName = String(lNameLen, 0)
    lDataLen = lEnumBufSize
    ReDim uData(lDataLen)
    If RegEnumValue(hKey, 0, sName, lNameLen, 0, REG_SZ, uData(0), lDataLen) <> 0 Then
        'no values, or enum failed
        RegCloseKey hKey
        Exit Sub
    End If
    
    Do
        sName = Left(sName, lNameLen)
        If sName = vbNullString Then
            sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, "")
            If sName = vbNullString Then sName = "(no name)"
        End If
        sCLSID = StrConv(uData, vbUnicode)
        sCLSID = TrimNull(Left(sCLSID, lDataLen))
        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", "")
        sFile = Replace(sFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
        If sFile = vbNullString Then
            sFile = "(no file)"
        Else
            If Not FileExists(sFile) Then
                sFile = sFile & " (file missing)"
            End If
        End If
        
        bOnWhiteList = False
        For j = 0 To UBound(sSafeSSODL)
            If Trim(sSafeSSODL(i)) = vbNullString Then Exit For
            If InStr(1, sSafeSSODL(j), sCLSID, vbTextCompare) > 0 Then
                bOnWhiteList = True
                If bIgnoreAllWhitelists Then bOnWhiteList = False
                Exit For
            End If
        Next j
        
        sHit = "O21 - SSODL: " & sName & " - " & sCLSID & " - " & sFile
        If Not IsOnIgnoreList(sHit) And Not bOnWhiteList Then
            If bMD5 Then sHit = sHit & GetFileMD5(sFile)
            frmMain.lstResults.AddItem sHit
        End If
        
        i = i + 1
        lNameLen = lEnumBufSize
        sName = String(lNameLen, 0)
        lDataLen = lEnumBufSize
        ReDim uData(lDataLen)
    Loop Until RegEnumValue(hKey, i, sName, lNameLen, 0, REG_SZ, uData(0), lDataLen) <> 0
    RegCloseKey hKey
End Sub

Public Sub FixOther21Item(sItem$)
    'O21 - SSODL: webcheck - {000....000} - c:\file.dll (file missing)
    'actions to take:
    '* kill regval
    '* kill clsid regkey
    Dim sName$, sCLSID$, sFile$, sSSODL$
    On Error GoTo Error:
    sSSODL = "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
   
    sName = Mid(sItem, 14)
    sCLSID = Mid(sName, InStr(sName, " - ") + 3)
    sFile = Mid(sCLSID, InStr(sCLSID, " - ") + 3)
    sName = Left(sName, InStr(sName, " - ") - 1)
    sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
    If InStr(sFile, " ( file missing)") > 0 Then
        sFile = Left(sFile, InStr(sFile, " (file missing)") - 1)
    End If
    If sFile = "(no file)" Then sFile = vbNullString
    
    RegDelVal HKEY_LOCAL_MACHINE, sSSODL, sName
    RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther21Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther22Item()
    'sharedtaskscheduler
    Dim sSTS$, hKey&, i&, sCLSID$, lCLSIDLen&, uData() As Byte, lDataLen&
    Dim sFile$, sName$, sHit$
    sSTS = "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler"
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sSTS, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        'regkey doesn't exist, or failed to open
        Exit Sub
    End If
    
    lCLSIDLen = lEnumBufSize
    sCLSID = String(lCLSIDLen, 0)
    lDataLen = lEnumBufSize
    ReDim uData(lDataLen)
    If RegEnumValue(hKey, 0, sCLSID, lCLSIDLen, 0, REG_SZ, uData(0), lDataLen) <> 0 Then
        'no values, or enum failed
        RegCloseKey hKey
        Exit Sub
    End If
    
    Do
        sCLSID = Left(sCLSID, lCLSIDLen)
        sName = StrConv(uData, vbUnicode)
        sName = TrimNull(Left(sName, lDataLen))
        If sName = vbNullString Then sName = "(no name)"
        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", "")
        sFile = Replace(sFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
        If sFile = vbNullString Then
            sFile = "(no file)"
        Else
            If Not FileExists(sFile) Then
                sFile = sFile & " (file missing)"
            End If
        End If
        
        sHit = "O22 - SharedTaskScheduler: " & sName & " - " & sCLSID & " - " & sFile
        If Not IsOnIgnoreList(sHit) Then
            If bMD5 Then sHit = sHit & GetFileMD5(sFile)
            frmMain.lstResults.AddItem sHit
        End If
        
        i = i + 1
        lCLSIDLen = lEnumBufSize
        sCLSID = String(lCLSIDLen, 0)
        lDataLen = lEnumBufSize
        ReDim uData(lDataLen)
    Loop Until RegEnumValue(hKey, i, sCLSID, lCLSIDLen, 0, REG_SZ, uData(0), lDataLen) <> 0
    RegCloseKey hKey
End Sub

Public Sub FixOther22Item(sItem$)
    'O22 - SharedTaskScheduler: blah - {000...000} - file.dll
    'todo:
    '* kill regval
    '* kill clsid regkey
    Dim sCLSID$, sSTS$
    sSTS = "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler"
    On Error GoTo Error:
    
    sCLSID = Mid(sItem, InStr(sItem, ": ") + 2)
    sCLSID = Mid(sCLSID, InStr(sCLSID, " - ") + 3)
    sCLSID = Left(sCLSID, InStr(sCLSID, " - ") - 1)
    
    RegDelVal HKEY_LOCAL_MACHINE, sSTS, sCLSID
    RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther22Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther23Item()
    'enum NT services
    Dim sServices$(), i%, j%, sName$, sDisplayName$
    Dim lStart&, lType&, sFile$, sCompany$, sHit$
    Dim bHideDisabled As Boolean, bHideMicrosoft As Boolean
    If Not bIsWinNT Then Exit Sub
    
    bHideDisabled = True
    bHideMicrosoft = True
    
    If bIgnoreAllWhitelists Then
        bHideMicrosoft = False
    End If
    
    sServices = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    If UBound(sServices) = 0 Or UBound(sServices) = -1 Then Exit Sub
    For i = 0 To UBound(sServices)
        sName = sServices(i)
        lStart = RegGetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start")
        lType = RegGetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Type")
        sDisplayName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DisplayName")
        If sDisplayName = vbNullString Then sDisplayName = sName
        sFile = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ImagePath")
        
        'cleanup filename
        If sFile = vbNullString Then
            sFile = "(no file)"
        Else
            If Left(sFile, 1) = """" Then
                'fix bug when e.g. ["c:\file.exe" -option]
                sFile = Mid(sFile, 2)
                sFile = Left(sFile, InStr(sFile, """") - 1)
            End If
            'If Right(sFile, 1) = """" Then sFile = Left(sFile, Len(sFile) - 1)
            
            'expand aliases
            sFile = ExpandEnvironmentVars(sFile)
            sFile = Replace(sFile, "%systemroot%", sWinDir, , , vbTextCompare)
            sFile = Replace(sFile, "\systemroot", sWinDir, , , vbTextCompare)
            sFile = Replace(sFile, "systemroot", sWinDir, , , vbTextCompare)
            
            'prefix windows folder if not specified
            If InStr(1, sFile, "system32\", vbTextCompare) = 1 Then
                sFile = sWinDir & "\" & sFile
            End If
            
            'sometimes the damn path isn't there AT ALL >_<
            If InStr(sFile, "\") = 0 Then
                If FileExists(sWinDir & "\" & sFile) Then sFile = sWinDir & "\" & sFile
                If FileExists(sWinSysDir & "\" & sFile) Then sFile = sWinSysDir & "\" & sFile
                If FileExists(Left(sWinDir, 3) & sFile) Then sFile = Left(sWinDir, 3) & sFile
            End If
            
            'remove parameters (and double filenames)
            'j = InStrRev(sFile, ".exe", , vbTextCompare) + 3
            j = InStr(1, sFile, ".exe ", vbTextCompare) + 3
            If j < Len(sFile) And j > 3 Then sFile = Left(sFile, j)
            
            'add .exe if not specified
            If InStr(1, sFile, ".exe", vbTextCompare) = 0 And _
               InStr(1, sFile, ".sys", vbTextCompare) = 0 Then
                If InStr(sFile, " ") > 0 Then
                    sFile = Left(sFile, InStr(sFile, " ") - 1)
                    sFile = sFile & ".exe"
                End If
            End If
            
            sFile = Trim(sFile)
            sCompany = GetFilePropCompany(sFile)
            If sCompany = vbNullString Then sCompany = "Unknown owner" '"?"
            
            If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
        End If
        
        If lStart <> 0 And lStart <> 1 And lType >= 16 Then
            If Not (lStart = 4 And bHideDisabled) And _
               Not (InStr(sCompany, "Microsoft") > 0 And bHideMicrosoft) Then
                If bMD5 Then sFile = sFile & GetFileFromAutostart(sFile)
                sHit = "O23 - Service: " & sDisplayName & IIf(sName <> sDisplayName, " (" & sName & ") - ", " - ") & sCompany & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        End If
    Next i
End Sub

Public Sub FixOther23Item(sItem$)
    'stop & disable NT service - DON'T delete it
    'O23 - Service: <displayname> - <company> - <file>
    ' (file missing) or (filesize .., MD5 ..) can be appended
    If Not bIsWinNT Then Exit Sub
    On Error GoTo Error:
    
    Dim sServices$(), i%, sName$, sDisplayName$
    sServices = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    If UBound(sServices) = 0 Or UBound(sServices) = -1 Then Exit Sub
    sDisplayName = Mid(sItem, InStr(sItem, ": ") + 2)
    sDisplayName = Left(sDisplayName, InStr(sDisplayName, " - ") - 1)
    For i = 0 To UBound(sServices)
        If sDisplayName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServices(i), "DisplayName") Then
            sName = sServices(i)
            
            RegSetDwordVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start", 4
            
            'this does the same as AboutBuster: run NET STOP on the
            'service. if the API way wouldn't crash VB everytime, I'd
            'use that. :/
            Shell sWinSysDir & "\NET.exe STOP """ & sName & """", vbHide
            'better do the display name too in case the regkey name
            'has funky characters (res://dll or temp\sp.html parasites)
            Shell sWinSysDir & "\NET.exe StOP """ & sDisplayName & """", vbHide
            Sleep 1000
            DoEvents
            
            RegSetDwordVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start", 4
            bRebootNeeded = True
            Exit For
        End If
    Next i
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther23Item", Err.Number, Err.Description, "sItem=" & sItem
End Sub

Public Sub CheckOther24Item()
    'activex desktop components
    Dim sDCKey$, sComponents$(), i&
    Dim sSource$, sSubscr$, sName$, sHit$
    sDCKey = "Software\Microsoft\Internet Explorer\Desktop\Components"
    sComponents = Split(RegEnumSubkeys(HKEY_CURRENT_USER, sDCKey), "|")
    
    For i = 0 To UBound(sComponents)
        If RegKeyExists(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i)) Then
            sSource = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "Source")
            sSubscr = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "SubscribedURL")
            sName = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "FriendlyName")
            If sName = vbNullString Then sName = "(no name)"
            If Not (LCase(sSource) = "about:home" And LCase(sSubscr) = "about:home") And _
               Not (UCase(sSource) = "131A6951-7F78-11D0-A979-00C04FD705A2" And UCase(sSubscr) = "131A6951-7F78-11D0-A979-00C04FD705A2") Then
                If sSource <> vbNullString Then
                    sHit = "O24 - Desktop Component " & sComponents(i) & ": " & sName & " - " & sSource
                Else
                    If sSubscr <> vbNullString Then
                        sHit = "O24 - Desktop Component " & sComponents(i) & ": " & sName & " - " & sSubscr
                    Else
                        sHit = "O24 - Desktop Component " & sComponents(i) & ": " & sName & " - (no file)"
                    End If
                End If
                
                If Not IsOnIgnoreList(sHit) Then
                    frmMain.lstResults.AddItem sHit
                End If
            End If
        End If
    Next i
End Sub

Public Sub FixOther24Item(sItem$)
    'delete the entire registry key
    'O24 - Desktop Component 1: Internet Explorer Channel Bar - 131A6951-7F78-11D0-A979-00C04FD705A2
    'O24 - Desktop Component 2: Security - %windir%\index.html
    
    Dim sDCKey$, sNum$, sName$, sURL$, sComponents$(), i&, sTestName$, sTestURL1$, sTestURL2$
    sDCKey = "Software\Microsoft\Internet Explorer\Desktop\Components"
    
    sNum = Mid(sItem, InStr(sItem, ":") - 1, 1)
    sName = Mid(sItem, InStr(sItem, ":") + 2)
    sURL = Mid(sName, InStr(sName, " - ") + 3)
    sName = Left(sName, InStr(sName, " - ") - 1)
    If "(no name)" = sName Then
        sName = ""
    End If
    
    sTestName = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sNum, "FriendlyName")
    sTestURL1 = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sNum, "Source")
    sTestURL2 = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sNum, "SubscribedURL")
    If sName = sTestName And (sURL = sTestURL1 Or sURL = sTestURL2) Then
        'found it!
        RegDelKey HKEY_CURRENT_USER, sDCKey & "\" & sNum
        If FileExists(sTestURL1) Then DeleteFile sTestURL1
        If FileExists(sTestURL2) Then DeleteFile sTestURL2
        Dim x As Long
        SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, "", 1 'SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    End If
End Sub

Public Sub FixUNIXHostsFile()
    'unix-style = hosts file has inproper linebreaks
    'Win32 linebreak: chr(13) + chr(10)
    'UNIX  linebreak: chr(10)
    'Mac   linebreak: chr(13)
    On Error GoTo Error:
    If Not FileExists(sHostsFile) Then Exit Sub
    If FileLen(sHostsFile) = 0 Then Exit Sub
    
    Dim sLine$, sFile$, sNewFile$, iAttr%, vContent As Variant
    iAttr = GetAttr(sHostsFile)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetAttr sHostsFile, vbNormal
    Open sHostsFile For Binary As #1
        sFile = Input(FileLen(sHostsFile), #1)
    Close #1
    
    'temp rename all proper linebreaks, replace unix-style
    'linebreaks with proper linebreaks, rename back
    sNewFile = sFile
    sNewFile = Replace(sNewFile, vbCrLf, "/|\|/")
    sNewFile = Replace(sNewFile, Chr(10), vbCrLf)
    'sNewFile = Replace(sNewFile, vbCrLf, "/|\|/")
    'sNewFile = Replace(sNewFile, Chr(13), vbCrLf)
    sNewFile = Replace(sNewFile, "/|\|/", vbCrLf)
    If sNewFile <> sFile Then
        DeleteFile sHostsFile
        Open sHostsFile For Output As #1
            Print #1, sNewFile
        Close #1
    End If
    SetAttr sHostsFile, iAttr
    Exit Sub
    
Error:
    Close #1
    ErrorMsg "modMain_FixUNIXHostsFile", Err.Number, Err.Description
End Sub

Public Function IsOnIgnoreList(sHit$) As Boolean
    Dim i%
    On Error GoTo Error:
    If sHit = vbNullString Then
        'load ignore list
        Dim iIgnoreNum%
        iIgnoreNum = CInt(RegRead("IgnoreNum", "0"))
        If iIgnoreNum <= 0 Then
            ReDim sIgnoreList(1 To 1)
            Exit Function
        End If
        
        ReDim sIgnoreList(1 To iIgnoreNum)
        For i = 1 To iIgnoreNum
            sIgnoreList(i) = RegRead("Ignore" & CStr(i), "")
            If sIgnoreList(i) = vbNullString Then
                'premature end of ignore list, halt
                ReDim Preserve sIgnoreList(1 To i - 1)
                Exit For
            End If
        Next i
        Exit Function
    End If
    
    On Error Resume Next
    If sIgnoreList(1) = vbNullString Then
        'ignore list empty, nothing is on ignorelist
        IsOnIgnoreList = False
        Exit Function
    End If
    If Err Then
        'ignore list empty, nothing is on ignorelist
        IsOnIgnoreList = False
        Exit Function
    End If
    On Error GoTo Error:
    
    IsOnIgnoreList = False
    For i = 1 To UBound(sIgnoreList)
        If sHit = sIgnoreList(i) Then
            IsOnIgnoreList = True
            Exit Function
        End If
    Next i
    Exit Function
    
Error:
    ErrorMsg "modMain_IsOnIgnoreList", Err.Number, Err.Description, sHit
    If Err.Number = 9 Then
        'clear ignorelist, it has been meddled with
        RegDel "IgnoreNum"
        For i = 1 To 99
            RegDel "Ignore" & CStr(i)
        Next i
    End If
End Function

Public Sub ErrorMsg(sProcedure$, iErrNum%, sErrDesc$, Optional sParameters$)
    Dim sMsg$
    Close
    If iErrNum = 0 Then Exit Sub
    'sMsg = "An unexpected error has occurred at procedure: " & _
           sProcedure & "(" & sParameters & ")" & vbCrLf & _
           "Error #" & CStr(iErrNum) & " - " & sErrDesc & vbCrLf & vbCrLf & _
           "Please email me at www.merijn.org/contact.html, reporting the following:" & vbCrLf & _
           "* What you were trying to fix when the error occurred, if applicable" & vbCrLf & _
           "* How you can reproduce the error" & vbCrLf & _
           "* A complete HijackThis scan log, if possible" & vbCrLf & vbCrLf & _
           "Windows version: " & sWinVersion & vbCrLf & _
           "MSIE version: " & sMSIEVersion & vbCrLf & _
           "HijackThis version: " & App.Major & "." & App.Minor & "." & App.Revision & _
           vbCrLf & vbCrLf & "This message has been copied to your clipboard." & _
           vbCrLf & "Click OK to continue the rest of the scan."
    
    sMsg = "Please help us improve HijackThis by reporting this error" & _
    vbCrLf & vbCrLf & "Click 'Yes' to submit" & _
    vbCrLf & vbCrLf & "Error Details: " & _
    vbCrLf & vbCrLf & "An unexpected error has occurred at procedure: " & _
    sProcedure & "(" & sParameters & ")" & vbCrLf & _
    "Error #" & CStr(iErrNum) & " - " & sErrDesc & _
     vbCrLf & vbCrLf & "Windows version: " & sWinVersion & vbCrLf & _
    "MSIE version: " & sMSIEVersion & vbCrLf & _
    "HijackThis version: " & App.Major & "." & App.Minor & "." & App.Revision
    
    Clipboard.Clear
    Clipboard.SetText sMsg
    If vbYes = MsgBox(sMsg, vbCritical + vbYesNo) Then
    Dim szParams As String
    Dim szCrashUrl As String
    szCrashUrl = "http://www.trendmicro.com/go/hjt/error/?"
    szParams = "function=" & sProcedure
    szParams = szParams & "&params=" & sParameters
    szParams = szParams & "&errorno=" & CStr(iErrNum)
    szParams = szParams & "&errortxt" & sErrDesc
    szParams = szParams & "&winver=" & sWinVersion
    szParams = szParams & "&iever=" & sMSIEVersion
    szParams = szParams & "&hjtver=" & App.Major & "." & App.Minor & "." & App.Revision
    szCrashUrl = szCrashUrl & URLEncode(szParams)
    If True = IsOnline Then
        ShellExecute 0&, "open", szCrashUrl, vbNullString, vbNullString, vbNormalFocus
    Else
        MsgBox "No Internet Connection Available"
    End If
    End If
    
End Sub

Public Sub CheckDateFormat()
    Dim sBuffer$, uST As SYSTEMTIME
    With uST
        .wDay = 10
        .wMonth = 11
        .wYear = 2003
    End With
    sBuffer = String(255, 0)
    GetDateFormat ByVal 0, 0, uST, vbNullString, sBuffer, 255
    sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    
    'last try with GetLocaleInfo didn't work on Win2k/XP
    If InStr(sBuffer, "10") < InStr(sBuffer, "11") Then
        bIsUSADateFormat = False
        'MsgBox "sBuffer = " & sBuffer & vbCrLf & "10 < 11, so bIsUSADateFormat False"
    Else
        bIsUSADateFormat = True
        'MsgBox sBuffer & vbCrLf & "10 !< 11, so bIsUSADateFormat True"
    End If
    
    'Dim lLndID&, sDateFormat$
    'lLndID = GetSystemDefaultLCID()
    'sDateFormat = String(255, 0)
    'GetLocaleInfo lLndID, LOCALE_SSHORTDATE, sDateFormat, 255
    'sDateFormat = Left(sDateFormat, InStr(sDateFormat, Chr(0)) - 1)
    'If sDateFormat = vbNullString Then Exit Sub
    ''sDateFormat = "dd-MM-yy" or "M/d/yy"
    ''I hope this works - dunno what happens in
    ''yyyy-mm-dd or yyyy-dd-mm format
    'If InStr(1, sDateFormat, "d", vbTextCompare) < _
    '   InStr(1, sDateFormat, "m", vbTextCompare) Then
    '    bIsUSADateFormat = False
    'Else
    '    bIsUSADateFormat = True
    'End If
End Sub

Public Function Unescape$(sURL$)
    Dim i%, sDummy$, sHex$
    
    'replace hex codes with proper character
    sDummy = sURL
    
    'don't need entire ascii range, 32-126
    'is all readable characters (I think)
    'For i = 1 To 255
    For i = 32 To 126
        sHex = Hex(i)
        If Len(sHex) = 1 Then sHex = "0" & sHex
        sDummy = Replace(sDummy, "%" & sHex, Chr(i), , , vbTextCompare)
    Next i
    
    Unescape = sDummy & " (obfuscated)"
End Function

Public Function HasSpecialCharacters(sName$) As Boolean
    'function checks for special characters in string,
    'like Chinese or Japanese.
    'Used in CheckOther3Item (IE Toolbar)
    HasSpecialCharacters = False
    
    'function disabled because of proper DBCS support
    Exit Function
    
    If Len(sName) <> lstrlen(sName) Then
        HasSpecialCharacters = True
        Exit Function
    End If
    
    If Len(sName) <> LenB(StrConv(sName, vbFromUnicode)) Then
        HasSpecialCharacters = True
        Exit Function
    End If
End Function

Public Sub CheckForReadOnlyMedia()
    Dim sMsg$
    On Error Resume Next
    Open BuildPath(App.Path, "~dummy.tmp") For Output As #1
        Print #1, "."
    Close #1
    
    If Err Then
        'damn, got no write access
        bNoWriteAccess = True
        sMsg = Translate(340)
'        sMsg = "It looks like you're running HijackThis from " & _
'               "a read-only device like a CD or locked floppy disk." & _
'               "If you want to make backups of items you fix, " & _
'               "you must copy HijackThis.exe to your hard disk " & _
'               "first, and run it from there." & vbCrLf & vbCrLf & _
'               "If you continue, you might get 'Path/File Access' " & _
'               "errors - do NOT email me those please."
        MsgBox sMsg, vbExclamation
        
    End If
    DeleteFile BuildPath(App.Path, "~dummy.tmp")
    On Error GoTo 0:
End Sub

Public Sub SetAllFontCharset()
    With frmMain
        SetFontCharSet .txtCheckUpdateProxy.Font
        SetFontCharSet .txtDefSearchAss.Font
        SetFontCharSet .txtDefSearchCust.Font
        SetFontCharSet .txtDefSearchPage.Font
        SetFontCharSet .txtDefStartPage.Font
        SetFontCharSet .txtHelp.Font
        SetFontCharSet .txtNothing.Font
        
        SetFontCharSet .lstBackups.Font
        SetFontCharSet .lstIgnore.Font
        SetFontCharSet .lstResults.Font
    End With
End Sub

Private Sub SetFontCharSet(objTxtboxFont As Object)
    'A big thanks to 'Gun' and 'Adult', two Japanese users
    'who helped me greatly with this
    Dim bNonUsCharset As Boolean
    On Error Resume Next
    bNonUsCharset = True
    Select Case GetUserDefaultLCID
         Case &H404
            objTxtboxFont.Charset = CHINESEBIG5_CHARSET
            objTxtboxFont.Name = ChrW(&H65B0) + ChrW(&H7D30) + ChrW(&H660E) _
                       + ChrW(&H9AD4)   'New Ming-Li
         Case &H411
            objTxtboxFont.Charset = SHIFTJIS_CHARSET
            objTxtboxFont.Name = ChrW(&HFF2D) + ChrW(&HFF33) + ChrW(&H20) + _
                       ChrW(&HFF30) + ChrW(&H30B4) + ChrW(&H30B7) + _
                       ChrW(&H30C3) + ChrW(&H30AF)
         Case &H412
            objTxtboxFont.Charset = HANGEUL_CHARSET
            objTxtboxFont.Name = ChrW(&HAD74) + ChrW(&HB9BC)
         Case &H804
            objTxtboxFont.Charset = CHINESESIMPLIFIED_CHARSET
            objTxtboxFont.Name = ChrW(&H5B8B) + ChrW(&H4F53)
         Case Else
            objTxtboxFont.Charset = DEFAULT_CHARSET
            'objTxtboxFont.Name = ""
            bNonUsCharset = False
    End Select
    
    If bNonUsCharset Then objTxtboxFont.Size = 9
End Sub

Public Function TrimNull$(s$)
    If InStr(s, Chr(0)) = 0 Then
        TrimNull = s
    Else
        TrimNull = Left(s, InStr(s, Chr(0)) - 1)
    End If
End Function

Public Sub CheckForStartedFromTempDir()
    'if user picks 'run from current location when downloading HijackThis.exe,
    'or runs file directly from zip file, exe will be ran from temp folder,
    'meaning a reboot or cache clean could delete it, as well any backups
    'made. Also the user won't be able to find the exe anymore :P
    
    Dim sAppPath$, bWeDontLikeThisPath As Boolean, sMsg$
    sAppPath = App.Path
    
    'started ok, from seperate folder
    If InStr(1, sAppPath, "Hijack", vbTextCompare) > 0 And _
       InStr(1, sAppPath, "HijackThis.zip", vbTextCompare) = 0 Then Exit Sub
    If InStr(1, sAppPath, "Spyware", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, sAppPath, "Security", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, sAppPath, "Program Files", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, sAppPath, "Desktop", vbTextCompare) > 0 Then Exit Sub
    
    If (Right(sAppPath, 1) = "\" And InStr(1, sAppPath, Left(sWinDir, 3)) > 0) Or _
       InStr(1, sAppPath, sWinDir, vbTextCompare) > 0 Or _
       InStr(1, sAppPath, sWinSysDir, vbTextCompare) > 0 Then
        'started from root folder or some system folder
        bWeDontLikeThisPath = True
    End If
    If InStr(InStrRev(sAppPath, "\"), sAppPath, "My Documents") > 0 Then
        'started directly from my documents folder
        bWeDontLikeThisPath = True
    End If
    If InStr(1, sAppPath, "Temporary", vbTextCompare) > 0 Then
        'started from IE cache
        bWeDontLikeThisPath = True
    End If
    If InStr(1, sAppPath, "Local Settings\Temp", vbTextCompare) > 0 Then
        'started from NT/2000/XP temp folder
        bWeDontLikeThisPath = True
    End If
    If InStr(1, sAppPath, sWinDir & "\temp", vbTextCompare) > 0 Then
        'started from 9x/ME temp folder
        bWeDontLikeThisPath = True
    End If
    
sHit:
    If bWeDontLikeThisPath Then
        sMsg = "HijackThis appears to have been started from a temporary " & _
               "folder. Since temp folders tend to be be emptied regularly, " & _
               "it's wise to copy HijackThis.exe to a folder of its own, " & _
               "for instance C:\Program Files\HijackThis." & vbCrLf & _
               "This way, any backups that will be made of fixed items " & _
               "won't be lost." & vbCrLf & vbCrLf & _
               "Please quit HijackThis and copy it to a separate folder " & _
               "first before fixing any items."
        
        MsgBox sMsg, vbExclamation + vbOKOnly
    End If
End Sub

Public Sub ShowFileProperties(sFile$)
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_NOCLOSEPROCESS
        .hwnd = frmMain.hwnd
        .lpFile = sFile
        .lpVerb = "properties"
        .nShow = 1
    End With
    ShellExecuteEx uSEI
End Sub

Public Sub RestartSystem(Optional sExtraPrompt$)
    If bIsWinNT Then
        SHRestartSystemMB frmMain.hwnd, StrConv(sExtraPrompt & IIf(sExtraPrompt <> vbNullString, vbCrLf & vbCrLf, vbNullString), vbUnicode), 2
    Else
        SHRestartSystemMB frmMain.hwnd, sExtraPrompt & IIf(sExtraPrompt <> vbNullString, vbCrLf & vbCrLf, vbNullString), 0
    End If
End Sub

Public Sub DeleteFileOnReboot(sFile$, Optional bDeleteBlindly As Boolean = False)
    'If Not bIsWinNT Then Exit Sub
    If Not FileExists(sFile) And Not bDeleteBlindly Then Exit Sub
    If bIsWinNT Then
        MoveFileEx sFile, vbNullString, MOVEFILE_DELAY_UNTIL_REBOOT
    Else
        Dim sDummy$
        On Error Resume Next
        Open sWinDir & "\wininit.ini" For Append As #1
            Print #1, "[rename]"
            Print #1, "NUL=" & GetDOSFilename(sFile)
            Print #1,
        Close #1
    End If
    RestartSystem Replace(Translate(342), "[]", sFile)
    'RestartSystem "The file '" & sFile & "' will be deleted by Windows when the system restarts."
End Sub

Public Function IsIPAddress(sIP$) As Boolean
    'IsIPAddress = IIf(inet_addr(sIP) <> -1, True, False)
    'can't really trust this API, sometimes it bails when the fourth
    'octet is >127
    Dim sOctets$()
    sOctets = Split(sIP, ".")
    If UBound(sOctets) = 3 Then
        If IsNumeric(sOctets(0)) And _
           IsNumeric(sOctets(1)) And _
           IsNumeric(sOctets(2)) And _
           IsNumeric(sOctets(3)) Then
            If (sOctets(0) >= 0 And sOctets(0) <= 255) And _
               (sOctets(1) >= 0 And sOctets(1) <= 255) And _
               (sOctets(2) >= 0 And sOctets(2) <= 255) And _
               (sOctets(3) >= 0 And sOctets(3) <= 255) Then
                IsIPAddress = True
            End If
        End If
    End If
End Function

Public Function DomainHasDoubleTLD(sDomain$) As Boolean
    Dim sDoubleTLDs$(), i%
    sDoubleTLDs = Split(".co.uk|" & _
                        ".da.ru|" & _
                        ".h1.ru|" & _
                        ".me.uk|" & _
                        ".ss.ru|" & _
                        ".xu.pl", "|")
                        '".com.au|" & _
                        ".com.br|" & _
                        ".1gb.ru|" & _
                        ".biz.ua|" & _
                        ".jps.ru|" & _
                        ".psn.cn|" & _
                        ".spb.ru|" & _
                        'above stuff somehow isn't recognized by IE
                        'as a double TLD - it's not a bug, it's a feature!

    For i = 0 To UBound(sDoubleTLDs)
        If InStr(sDomain, sDoubleTLDs(i)) = Len(sDomain) - Len(sDoubleTLDs(i)) + 1 Then
            DomainHasDoubleTLD = True
            Exit Function
        End If
    Next i
End Function

Public Function GetDOSFilename$(sFile$, Optional bReverse As Boolean = False)
    'works for folders too btw
    If bReverse Then GoTo Expand:
    Dim sBuffer$
    If Not FileExists(sFile) Then
        GetDOSFilename = sFile
        Exit Function
    End If
    sBuffer = String(260, 0)
    GetDOSFilename = Left(sBuffer, GetShortPathName(sFile, sBuffer, Len(sBuffer)))
    Exit Function
    
Expand:
    Dim hFind, uWFD As WIN32_FIND_DATA, sPath$, sDir$, sExe$, sFullPath$
    hFind = FindFirstFile(sFile, uWFD)
    If hFind <> 0 Then
        sExe = Left(uWFD.cFileName, InStr(uWFD.cFileName, Chr(0)) - 1)
    End If
    FindClose hFind
    sFullPath = sExe
    
    sDir = Left(sFile, InStrRev(sFile, "\") - 1)
    If InStr(sDir, ":") <> Len(sDir) Then
        Do
            hFind = FindFirstFile(sDir, uWFD)
            If hFind <> 0 Then
                sPath = Left(uWFD.cFileName, InStr(uWFD.cFileName, Chr(0)) - 1)
            End If
            FindClose hFind
            sFullPath = sPath & "\" & sFullPath
        sDir = Left(sFile, InStrRev(sDir, "\") - 1)
        Loop Until InStr(sDir, ":") = Len(sDir)
    End If
    sFullPath = sDir & "\" & sFullPath
    GetDOSFilename = sFullPath
End Function

Public Sub DeleteNTService(sServiceName$)
    'I wish everything this hard was this simple :/
    Dim hSCManager&, hService&
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    If hSCManager > 0 Then
        hService = OpenService(hSCManager, sServiceName, SERVICE_ALL_ACCESS)
        If hService > 0 Then
            If DeleteService(hService) > 0 Then
                bRebootNeeded = True
            End If
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSCManager
    End If
    
    'MSDN says all hooks to service must be closed to allow SCManager to
    'delete it. if this doesn't happen, the service is deleted on reboot
    If bRebootNeeded Then
        RestartSystem Replace(Translate(343), "[]", sServiceName)
        'RestartSystem "The service '" & sServiceName & "'  has been marked for deletion."
    Else
        MsgBox Replace(Translate(344), "[]", sServiceName), vbExclamation
'        MsgBox "Unable to delete the service '" & sServiceName & "'. Make sure the name " & _
'               "name is correct and the service is not running.", vbExclamation
    End If
End Sub

Public Function RunningInIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then RunningInIDE = True
End Function

Public Function GetSpybotVersion$()
    If RegKeyExists(HKEY_CURRENT_USER, "Software\PepiMK Software\SpybotSnD") Then
        GetSpybotVersion = RegGetString(HKEY_CURRENT_USER, "Software\PepiMK Software\SpybotSnD", "Version")
    Else
        GetSpybotVersion = "not installed"
    End If
End Function

Public Function GetAdAwareVersion$()
    If RegKeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Ad-Aware SE Personal") Then
        Dim s$
        s = RegGetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Ad-Aware SE Personal", "DisplayIcon")
        If s <> vbNullString Then
            s = Left(s, InStrRev(s, ".exe") + 3)
            GetAdAwareVersion = GetFilePropVersion(s)
        Else
            GetAdAwareVersion = "not installed"
        End If
    Else
        If RegKeyExists(HKEY_CURRENT_USER, "Software\Lavasoft\AD-Aware") Then
            'only works for 5.x and lower, 6.x uses stupid tricks
            GetAdAwareVersion = RegGetString(HKEY_CURRENT_USER, "Software\Lavasoft\AD-Aware", "Version")
        Else
            GetAdAwareVersion = "not installed"
        End If
    End If
End Function

Public Function GetUser$(Optional bCheckAdmin As Boolean = False)
    Dim sUserName$
    sUserName = String(255, 0)
    GetUserName sUserName, 255
    sUserName = Left(sUserName, InStr(sUserName, Chr(0)) - 1)
    GetUser = UCase(sUserName)
End Function

Public Function GetComputer$()
    Dim sComputerName$
    sComputerName = String(255, 0)
    GetComputerName sComputerName, 255
    sComputerName = Left(sComputerName, InStr(sComputerName, Chr(0)) - 1)
    GetComputer = UCase(sComputerName)
End Function

Public Sub CheckOther4ItemUsers()
    'list autostart entries from HKEY_USERS subkeys
    Dim sUsers$(), i%, sKeys$(4), sUserName$
    sKeys(0) = "Software\Microsoft\Windows\CurrentVersion\Run"
    sKeys(1) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
    sKeys(2) = "Software\Microsoft\Windows\CurrentVersion\RunServices"
    sKeys(3) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    sKeys(4) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
    
    'get all users' SID and map it to the corresponding username
    'not all users visible in User Accounts screen have a SID though,
    'they probably get this when logging in for the first time
    sUsers = Split(RegEnumSubkeys(HKEY_USERS, ""), "|")
    For i = 0 To UBound(sUsers)
        If Left(sUsers(i), 1) = "S" And InStr(sUsers(i), "_Classes") = 0 Then
            sUserName = MapSIDToUsername(sUsers(i))
            If sUserName = vbNullString Then sUserName = "?"
            If UCase(sUserName) <> GetUser() Then
                sUsers(i) = sUsers(i) & "|" & sUserName
            Else
                'filter out current user
                sUsers(i) = vbNullString
            End If
        Else
            sUsers(i) = vbNullString
        End If
    Next i
    
    'add default user
    ReDim Preserve sUsers(UBound(sUsers) + 1)
    sUsers(UBound(sUsers)) = ".DEFAULT|Default user"
        
    Dim j%, k&, hKey&, sValue$, uData() As Byte, sData$, sHit$, sSID$
    For i = 0 To UBound(sUsers)
        If InStr(sUsers(i), "|") > 0 Then
            sSID = Left(sUsers(i), InStr(sUsers(i), "|") - 1)
            sUserName = Mid(sUsers(i), InStr(sUsers(i), "|") + 1)
            For j = 0 To UBound(sKeys)
                If RegOpenKeyEx(HKEY_USERS, sSID & "\" & sKeys(j), 0, KEY_QUERY_VALUE, hKey) = 0 Then
                    sValue = String(lEnumBufSize, 0)
                    ReDim uData(lEnumBufSize)
                    If RegEnumValue(hKey, 0, sValue, Len(sValue), 0, ByVal 0, uData(0), UBound(uData)) = 0 Then
                        Do
                            sValue = TrimNull(sValue)
                            sData = TrimNull(StrConv(uData, vbUnicode))
                            Select Case j
                                Case 0: sHit = "O4 - HKUS\" & sSID & "\..\Run: [" & sValue & "] " & sData & " (User '" & sUserName & "')"
                                Case 1: sHit = "O4 - HKUS\" & sSID & "\..\RunOnce: [" & sValue & "] " & sData & " (User '" & sUserName & "')"
                                Case 2: sHit = "O4 - HKUS\" & sSID & "\..\RunServices: [" & sValue & "] " & sData & " (User '" & sUserName & "')"
                                Case 3: sHit = "O4 - HKUS\" & sSID & "\..\RunServicesOnce: [" & sValue & "] " & sData & " (User '" & sUserName & "')"
                                Case 4: sHit = "O4 - HKUS\" & sSID & "\..\Policies\Explorer\Run: [" & sValue & "] " & sData & " (User '" & sUserName & "')"
                            End Select
                            If Not IsOnIgnoreList(sHit) Then
                                frmMain.lstResults.AddItem sHit
                            End If
                            k = k + 1
                            sValue = String(lEnumBufSize, 0)
                            ReDim uData(lEnumBufSize)
                        Loop Until RegEnumValue(hKey, k, sValue, Len(sValue), 0, ByVal 0, uData(0), UBound(uData)) <> 0
                    End If
                    RegCloseKey hKey
                End If
            Next j
        End If
    Next i
    
    'repeat for startup group folders - straight copy/paste from O4 /w fixes
    Dim sAutostartFolder$(1 To 4), sFile$, sShortCut$
    For i = 0 To UBound(sUsers)
        If InStr(sUsers(i), "|") > 0 Then
            sSID = Left(sUsers(i), InStr(sUsers(i), "|") - 1)
            sUserName = Mid(sUsers(i), InStr(sUsers(i), "|") + 1)
        
            sAutostartFolder(1) = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
            sAutostartFolder(2) = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AltStartup")
            sAutostartFolder(3) = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
            sAutostartFolder(4) = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AltStartup")
        
            For k = 1 To UBound(sAutostartFolder)
                If bIsWinNT And sAutostartFolder(k) <> vbNullString Then
                    sAutostartFolder(k) = Replace(sAutostartFolder(k), "%USERPROFILE%", "C:\Documents And Settings\" & sUserName)
                Else
                    sAutostartFolder(k) = Replace(sAutostartFolder(k), "%USERPROFILE%", sWinDir & "\" & sUserName)
                End If
                If sAutostartFolder(k) <> vbNullString And _
                   FolderExists(sAutostartFolder(k)) Then
                    sShortCut = Dir(sAutostartFolder(k) & "\*.*", vbArchive + vbHidden + vbReadOnly + vbSystem + vbDirectory)
                    If sShortCut <> vbNullString Then
                        Do
                            Select Case k
                                Case 1: sHit = "O4 - " & sSID & " Startup: "
                                Case 2: sHit = "O4 - " & sSID & " AltStartup: "
                                Case 3: sHit = "O4 - " & sSID & " User Startup: "
                                Case 4: sHit = "O4 - " & sSID & " User AltStartup: "
                            End Select
                            sFile = GetFileFromShortCut(sAutostartFolder(k) & "\" & sShortCut)
                            sHit = sHit & sShortCut & sFile & " (User '" & sUserName & "')"
                            If LCase(sShortCut) <> "desktop.ini" And _
                               sShortCut <> "." And sShortCut <> ".." And _
                               Not IsOnIgnoreList(sHit) Then
                                If bMD5 And sFile <> vbNullString And sFile <> " = ?" Then
                                    sHit = sHit & GetFileMD5(Mid(sFile, 4))
                                End If
                                frmMain.lstResults.AddItem sHit
                            End If
                            
                            sShortCut = Dir
                        Loop Until sShortCut = vbNullString
                    End If
                End If
            Next k
        End If
    Next i

End Sub

Public Function MapSIDToUsername$(sSID$)
    Dim objWMI As Object, objSID As Object
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:{impersonationLevel=Impersonate}")
    Set objSID = objWMI.Get("Win32_SID.SID='" & sSID & "'")
    MapSIDToUsername = objSID.AccountName
    Set objSID = Nothing
    Set objWMI = Nothing
End Function

Public Sub FixOther4ItemUsers(sItem$)
    'O4 - Enumeration of autoloading Regedit entries
    'O4 - HKUS\S-1-5-19\..\Run: [blah] program.exe (Username 'Joe')

    On Error GoTo Error:
    Dim lHive&, sKey$, sVal$, sData$, sSID$, sUserName$
    If InStr(sItem, "[") = 0 Then GoTo FixShortCut
    sItem = Mid(sItem, 6)
    lHive = HKEY_USERS
    sSID = Mid(sItem, InStr(sItem, "\") + 1)
    sSID = Left(sSID, InStr(sSID, "\") - 1)
    
    If InStr(sItem, "\RunServices:") > 0 Then
        sKey = "Software\Microsoft\Windows\CurrentVersion\RunServices"
    ElseIf InStr(sItem, "\RunOnce:") > 0 Then
        sKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
    ElseIf InStr(sItem, "\RunServicesOnce:") > 0 Then
        sKey = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    Else
        If InStr(1, sItem, "\Policies\", vbTextCompare) = 0 Then
            sKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        Else
            sKey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
        End If
    End If

    sVal = Mid(sItem, InStr(sItem, "[") + 1)
    sData = Mid(sVal, InStrRev(sVal, "]") + 2)
    sData = Left(sData, InStr(sData, " (User '") - 1)
    KillProcessByFile GetFileFromAutostart(sData, False)
    'some wankers used a garbled value name with a ']' in it.
    'assuming no one ever uses a filename with a ']' in it in the
    'future, this workaround should work (InStrRev instead of InStr)
    'update: autorun with sol[1].exe - doh!
    sVal = Left(sVal, InStrRev(sVal, "] ") - 1)
    
    RegDelVal lHive, sSID & "\" & sKey, sVal
    Exit Sub
    
FixShortCut:
    'O4 - SID Startup: bla.lnk = c:\bla.exe (User 'blah')
    Dim sPath$, sFile$
    sPath = Mid(sItem, 6)
    If InStr(sPath, ": ") = 0 Then Exit Sub
    sSID = Left(sPath, InStr(sPath, " ") - 1)
    sUserName = MapSIDToUsername(sSID)
    sPath = Mid(sPath, InStr(sPath, " ") + 1)
    sPath = Left(sPath, InStr(sPath, ": ") - 1)
    Select Case sPath
        Case "Startup":                sPath = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
        Case "User Startup":           sPath = RegGetString(HKEY_USERS, sSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup")
    End Select
    If sPath = vbNullString Then Exit Sub
    If bIsWinNT Then
        sPath = Replace(sPath, "%USERPROFILE%", "C:\Documents And Settings\" & sUserName)
    Else
        sPath = Replace(sPath, "%USERPROFILE%", sWinDir & "\" & sUserName)
    End If
    sFile = Mid(sItem, InStr(sItem, ": ") + 2)
    If InStr(sFile, " = ") > 0 Then
        sData = Mid(sFile, InStr(sFile, " = ") + 3)
        sFile = Left(sFile, InStr(sFile, " = ") - 1)
    Else
        sData = sPath & "\" & sFile
    End If
    sFile = sPath & IIf(Right(sPath, 1) = "\", "", "\") & sFile
    If FileExists(sFile) Then
        On Error Resume Next
        KillProcessByFile GetFileFromAutostart(sData)
        DeleteFile sFile
        If Err Then
            MsgBox Replace(Translate(320), "[]", sItem) & " " & _
                   IIf(bIsWinNT, Translate(321), Translate(322)) & _
                   " " & Translate(323), vbExclamation
'            MsgBox "Unable to delete the file:" & vbCrLf & _
'                   sItem & vbCrLf & vbCrLf & "The file " & _
'                   "may be in use. Use " & IIf(bIsWinNT, _
'                   "Task Manager", "a process killer like " & _
'                   "ProcView") & " to shutdown the program " & _
'                   "and run HijackThis again to delete the file.", vbExclamation
        End If
        On Error GoTo Error:
    End If
    Exit Sub
    
Error:
    ErrorMsg "modMain_FixOther4Item", Err.Number, Err.Description, "sItem=" & sItem
    
End Sub

Public Function GetBootMode$()
    Dim lRet&
    lRet = GetSystemMetrics(SM_CLEANBOOT)
    With frmMain
        Select Case lRet
            Case 1 'safe mode
                GetBootMode = "Safe mode"
            Case 2 ' safe mode with network support
                GetBootMode = "Safe mode with network support"
            Case Else 'normal boot mode
                GetBootMode = "Normal"
        End Select
    End With
End Function

Public Sub CopyFolder(sFolder$, sTo$)
    Dim uFOS As SHFILEOPSTRUCT
    With uFOS
        .wFunc = FO_COPY
        .pFrom = sFolder
        .pTo = sTo
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    MsgBox SHFileOperation(uFOS)
End Sub

Public Sub DeleteFolder(sFolder$)
    Dim uFOS As SHFILEOPSTRUCT
    With uFOS
        .wFunc = FO_DELETE
        .pFrom = sFolder
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    SHFileOperation uFOS
End Sub

Public Sub MoveFolder(sFolder$, sTo$)
    Dim uFOS As SHFILEOPSTRUCT
    With uFOS
        .wFunc = FO_MOVE
        .pFrom = sFolder
        .pTo = sTo
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    SHFileOperation uFOS
End Sub

Public Function ExpandEnvironmentVars$(s$)
    Dim sDummy$, lLen&
    If InStr(s, "%") = 0 Then
        ExpandEnvironmentVars = s
        Exit Function
    End If
    lLen = ExpandEnvironmentStrings(s, ByVal 0, 0)
    If lLen > 0 Then
        sDummy = String(lLen, 0)
        ExpandEnvironmentStrings s, sDummy, Len(sDummy)
        sDummy = TrimNull(sDummy)
        
        If InStr(sDummy, "%") = 0 Then
            ExpandEnvironmentVars = sDummy
            Exit Function
        End If
    Else
        sDummy = s
    End If
End Function

Public Function GetUserType$()
    'based on OpenProcessToken API example from API-Guide
    Dim hProcessToken&
    Dim BufferSize&
    Dim psidAdmin&, psidPower&, psidUser&, psidGuest&
    Dim lResult&
    Dim i%
    Dim tpTokens As TOKEN_GROUPS
    Dim tpSidAuth As SID_IDENTIFIER_AUTHORITY
    
    If Not bIsWinNT Then
        GetUserType = "Administrator"
        Exit Function
    End If
    
    GetUserType = "unknown"
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
    
    ' Obtain current process token
    If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) Then
        Call OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken)
    End If
    If hProcessToken Then

        ' Determine the buffer size required
        Call GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) ' Determine required buffer size
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
            
            ' Retrieve your token information
            If GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize) <> 1 Then
                Call CloseHandle(hProcessToken)
                Exit Function
            End If
            
            ' Move it from memory into the token structure
            Call CopyMemory(tpTokens, InfoBuffer(0), Len(tpTokens))
            
            ' Retreive the builtin sid pointers
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_POWER_USERS, 0, 0, 0, 0, 0, 0, psidPower)
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_USERS, 0, 0, 0, 0, 0, 0, psidUser)
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_GUESTS, 0, 0, 0, 0, 0, 0, psidGuest)
            
            If IsValidSid(psidAdmin) And IsValidSid(psidPower) And _
               IsValidSid(psidUser) And IsValidSid(psidGuest) Then
                For i = 0 To tpTokens.GroupCount
                
                    ' Run through your token sid pointers
                    If IsValidSid(tpTokens.Groups(i).Sid) Then
                    
                        ' Test for a match between the admin sid equalling your sid's
                        If EqualSid(ByVal tpTokens.Groups(i).Sid, ByVal psidAdmin) Then
                            GetUserType = "Administrator"
                            Exit For
                        End If
                        If EqualSid(ByVal tpTokens.Groups(i).Sid, ByVal psidPower) Then
                            GetUserType = "Power User"
                            Exit For
                        End If
                        If EqualSid(ByVal tpTokens.Groups(i).Sid, ByVal psidUser) Then
                            GetUserType = "Limited User"
                            Exit For
                        End If
                        If EqualSid(ByVal tpTokens.Groups(i).Sid, ByVal psidGuest) Then
                            GetUserType = "Guest"
                            Exit For
                        End If
                    End If
                Next
            End If
            If psidAdmin Then Call FreeSid(psidAdmin)
            If psidPower Then Call FreeSid(psidPower)
            If psidUser Then Call FreeSid(psidUser)
            If psidGuest Then Call FreeSid(psidGuest)
        End If
        Call CloseHandle(hProcessToken)
    End If
End Function

Public Sub ToggleWow64FSRedirection(bEnable As Boolean)
    Dim lIsWow64&
    Call IsWow64Process(GetCurrentProcess, lIsWow64)
    If Not lIsWow64 Then Exit Sub
    
    If bEnable Then
        Wow64RevertWow64FsRedirection lWow64Old
        lWow64Old = 0
    Else
        lWow64Old = 0
        Wow64DisableWow64FsRedirection lWow64Old
    End If
End Sub

Public Sub SilentDeleteOnReboot(sCmd$)
    Dim sDummy$, sFileName$
    'sCmd is all command-line parameters, like this
    '/param1 /deleteonreboot c:\progra~1\bla\bla.exe /param3
    '/param1 /deleteonreboot "c:\program files\bla\bla.exe" /param3
    
    sDummy = Mid(sCmd, InStr(sCmd, "/deleteonreboot") + Len("/deleteonreboot") + 1)
    If InStr(sDummy, """") = 1 Then
        'enclosed in quotes, chop off at next quote
        sFileName = Mid(sDummy, 2)
        sFileName = Left(sFileName, InStr(sFileName, """") - 1)
    Else
        'no quotes, chop off at next space if present
        If InStr(sDummy, " ") > 0 Then
            sFileName = Left(sDummy, InStr(sDummy, " ") - 1)
        Else
            sFileName = sDummy
        End If
    End If
    DeleteFileOnReboot sFileName, True
End Sub

Public Sub DeleteFile(sFile$)
    If Not FileExists(sFile) Then Exit Sub
    Dim uSFO As SHFILEOPSTRUCT
    With uSFO
        .pFrom = sFile
        .wFunc = FO_DELETE
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    SHFileOperation uSFO
End Sub

Function IsProcedureAvail(ByVal ProcedureName As String, ByVal DllFilename As String) As Boolean

    Dim hModule As Long, procAddr As Long
    hModule = LoadLibrary(DllFilename)
    If hModule Then
        procAddr = GetProcAddress(hModule, ProcedureName)
        FreeLibrary hModule
    End If
    IsProcedureAvail = (procAddr <> 0)
End Function

