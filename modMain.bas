Attribute VB_Name = "modMain"
'R0 - Changed Registry value (MSIE)
'R1 - Created Registry value
'R2 - Created Registry key
'R3 - Created extra value in regkey where only one should be
'F0 - Changed inifile value (system.ini)
'F1 - Created inifile value (win.ini)
'N1 (removed in 2.0.7) - Changed NS4.x homepage
'N2 (removed in 2.0.7) - Changed NS6 homepage
'N3 (removed in 2.0.7) - Changed NS7 homepage/searchpage
'N4 (removed in 2.0.7) - Changed Moz homepage/searchpage
'O1 - Hosts / hosts.ics / DNSApi hijackers
'O2 - BHO
'O3 - IE Toolbar
'O4 - Reg. autorun entry / msconfig disabled items
'O5 - Control.ini IE Options block
'O6 - Policies IE Options/Control Panel block
'O7 - Policies: Regedit block / IPSec
'O8 - IE Context menuitem
'O9 - IE Tools menuitem/button
'O10 - Winsock hijack
'O11 - IE Advanced Options group
'O12 - IE Plugin
'O13 - IE DefaultPrefix hijack
'O14 - IERESET.INF hijack
'O15 - Trusted Zone autoadd
'O16 - Downloaded Program Files
'O17 - Domain hijacks / DHCP DNS
'O18 - Protocol & Filter enum
'O19 - User style sheet hijack
'O20 - AppInit_DLLs registry value + Winlogon Notify subkeys
'O21 - ShellServiceObjectDelayLoad enum
'O22 - SharedTaskScheduler enum
'O23 - Windows Services
'O24 - Active desktop components
'O25 - Windows Management Instrumentation (WMI) event consumers

'// Do not forget to add prefix to Backup module (2 times) and Fix, and increase count on Sections' sort function
'// or you will be shot :)

'Next possible methods:
'* SearchAccurates 'URL' method in a InitPropertyBag (??)
'* HKLM\..\CurrentVersion\ModuleUsage
'* HKLM\..\CurrentVersion\Explorer\ShellExecuteHooks (eudora)
'* HKLM\..\Internet Explorer\SafeSites (searchaccurate)

Option Explicit

Public Const MAX_NAME = 256&
Private Const LB_SETHORIZONTALEXTENT    As Long = &H194

Public Enum ENUM_Cure_type
    FILE_BASED = 0              ' if need to cure .RunObject
    REGISTRY_KEY_BASED = 1      ' if need to cure .RegKey
    REGISTRY_PARAM_BASED = 2    ' if need to cure .RegParam inside the .RegKey
    AUTORUN_BASED = 3           ' if need to cure .AutoRunObject
End Enum

Private Type O25_Info
    sScriptFile     As String
    '-------------------------
    sTimerClassName As String
    TimerID         As String
    '-------------------------
    ConsumerName    As String
    ConsumerNameSpace As String
    ConsumerPath    As String
    '-------------------------
    FilterName      As String
    FilterNameSpace As String
    FilterPath      As String
End Type

Public Type TYPE_Scan_Results
    O25             As O25_Info
    HitLineW        As String
    HitLineA        As String
    Section         As String
    Alias           As String
    lHive           As Long
    RegKey()        As String
    RegParam        As String
    DefaultData     As Variant
    CLSID           As String
    AutoRunObject   As String
    RunObject       As String
    RunObjectArgs   As String
    ExpandedTarget  As String   ' target, expanded from .RunObject
    CureType        As ENUM_Cure_type
    Redirected      As Boolean  'is key under Wow64
End Type

Type Perfomance_TYPE
    'BeginExecution  As Date ' время, когда утилита реально была запусщена (определяется только при запуске из-под AutoLogger-а)
    'AVSandbox       As Long ' время (в секундах), проведенное в песочнице антивируса
    StartTime       As Date ' время начала работы программы
    EndTime         As Date ' время завершения работы программы
    'SearchTime      As Date ' время, затраченное на поиск файлов ярлыков
    'MAX_TimeOut     As Long ' Максимальное кол-во времени, которое разрешено работать программе (в секундах, с учетом времени, проведенного в песочнице антивируса)
End Type

Private Type TASK_WHITELIST_ENTRY
    OSver As Single
    Name As String
    Directory As String
    RunObj As String
    Args As String
End Type

Private Type DICTIONARIES
    TaskWL_ID  As clsTrickHashTable
End Type

Private Type SHELLEXECUTEINFO
    cbSize          As Long
    fMask           As Long
    hWnd            As Long
    lpVerb          As Long
    lpFile          As Long
    lpParameters    As Long
    lpDirectory     As Long
    nShow           As Long
    hInstApp        As Long
    lpIDList        As Long
    lpClass         As Long
    hkeyClass       As Long
    dwHotKey        As Long
    hIcon           As Long
    hProcess        As Long
End Type

Private Type TYPE_MY_OSVERSION
    OSName              As String
    SPVer               As Single
    Bitness             As String
    Major               As Long
    Minor               As Long
    MajorMinor          As Single
    Build               As Single
    Edition             As String
    Platform            As String
    bIsVistaOrLater     As Boolean
    bIsWin64            As Boolean
    bIsSafeBoot         As Boolean
    BootMode            As String
    bIsAdmin            As Boolean
    LangSystemName      As String
    LangSystemCode      As Long
    LangDisplayName     As String
    LangDisplayCode     As Long
    LangNonUnicodeName  As String
    LangNonUnicodeCode  As Long
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

Private Type SHFILEOPSTRUCT
    hWnd    As Long
    wFunc   As Long
    pFrom   As Long
    pTo     As Long
    fFlags  As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear       As Integer
    wMonth      As Integer
    wDayOfWeek  As Integer
    wDay        As Integer
    wHour       As Integer
    wMinute     As Integer
    wSecond     As Integer
    wMilliseconds As Integer
End Type

Private Type OPENFILENAME
    lStructSize         As Long
    hWndOwner           As Long
    hInstance           As Long
    lpstrFilter         As Long
    lpstrCustomFilter   As Long
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As Long
    nMaxFile            As Long
    lpstrFileTitle      As Long
    nMaxFileTitle       As Long
    lpstrInitialDir     As Long
    lpstrTitle          As Long
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As Long
    pvReserved          As Long
    dwReserved          As Long
    FlagsEx             As Long
End Type

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
'Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameW" (ByVal lpBuffer As Long, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameW" (ByVal lpBuffer As Long, nSize As Long) As Long

Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatW" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As Long, ByVal lpDateStr As Long, ByVal cchDate As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExW" (SEI As SHELLEXECUTEINFO) As Long

Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameW" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Public Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bDontOverwrite As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationW" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathW" (ByVal hWndOwner As Long, ByVal CSIDL As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As Long) As Long

Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal lSize As Long)
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function OpenThreadToken Lib "advapi32" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As Long, TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function IsValidSid Lib "advapi32" (ByVal pSid As Long) As Long
Private Declare Function EqualSid Lib "advapi32" (ByVal pSid1 As Long, ByVal pSid2 As Long) As Long
Private Declare Function EqualPrefixSid Lib "advapi32" (ByVal pSid1 As Long, ByVal pSid2 As Long) As Long
Private Declare Sub FreeSid Lib "advapi32" (ByVal pSid As Long)

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxW" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, ByVal Arguments As Long) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValueW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function PathFindOnPath Lib "Shlwapi" Alias "PathFindOnPathW" (ByVal pszFile As Long, ppszOtherDirs As Any) As Boolean
Private Declare Function PathRemoveFileSpec Lib "Shlwapi.dll" Alias "PathRemoveFileSpecW" (ByVal pszPath As Long) As Long
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long

Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal iMaxLength As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long

Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidW" (ByVal lpSystemName As Long, ByVal lpSid As Long, ByVal lpName As Long, cchName As Long, ByVal lpReferencedDomainName As Long, cchReferencedDomainName As Long, eUse As Long) As Long
Private Declare Function ConvertStringSidToSid Lib "advapi32.dll" Alias "ConvertStringSidToSidW" (ByVal StringSid As Long, pSid As Long) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (psa() As Any) As Long


Public Const CREATE_NEW                As Long = 1&
Public Const OPEN_EXISTING             As Long = 3&
Public Const CREATE_ALWAYS             As Long = 2&
Public Const GENERIC_READ              As Long = &H80000000
Public Const GENERIC_WRITE             As Long = &H40000000
Public Const FILE_SHARE_READ           As Long = &H1&
Public Const FILE_SHARE_WRITE          As Long = &H2&
Public Const FILE_SHARE_DELETE         As Long = 4&

Private Const SPI_SETDESKWALLPAPER  As Long = 20&
Private Const SPIF_SENDWININICHANGE As Long = &H2&
Private Const SPIF_UPDATEINIFILE    As Long = &H1&

Public Const CF_UNICODETEXT    As Long = 13&
Public Const GMEM_MOVEABLE     As Long = &H2&
Public Const CF_LOCALE         As Long = 16

Private Const MAX_PATH As Long = 260&

Private Const SECURITY_NT_AUTHORITY         As Long = &H5&
Private Const TOKEN_QUERY                   As Long = &H8&
Private Const TokenGroups                   As Long = 2&
Private Const SECURITY_BUILTIN_DOMAIN_RID   As Long = &H20&
Private Const DOMAIN_ALIAS_RID_ADMINS       As Long = &H220&
Private Const DOMAIN_ALIAS_RID_USERS        As Long = &H221&
Private Const DOMAIN_ALIAS_RID_GUESTS       As Long = &H222&
Private Const DOMAIN_ALIAS_RID_POWER_USERS  As Long = &H223&
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS  As Long = 548&
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS   As Long = 549&
Private Const DOMAIN_ALIAS_RID_PRINT_OPS    As Long = 550&
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS   As Long = 551&

Private Const ERROR_NONE_MAPPED As Long = 1332&

Private Const FO_MOVE               As Long = &H1&
Private Const FO_COPY               As Long = &H2&
Private Const FO_DELETE             As Long = &H3&
Private Const FOF_NOCONFIRMATION    As Long = &H10&
Private Const FOF_SILENT            As Long = &H4&

Private Const SM_CLEANBOOT          As Long = &H43&

Private Const MOVEFILE_DELAY_UNTIL_REBOOT As Long = &H4&

Private Const FILE_ATTRIBUTE_READONLY   As Long = 1&

Private Const SEE_MASK_INVOKEIDLIST     As Long = &HC&
Private Const SEE_MASK_NOCLOSEPROCESS   As Long = &H40&
'Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Const DEFAULT_CHARSET           As Long = 1&
Private Const SYMBOL_CHARSET            As Long = 2&
Private Const SHIFTJIS_CHARSET          As Long = 128&
Private Const HANGEUL_CHARSET           As Long = 129&
Private Const CHINESEBIG5_CHARSET       As Long = 136&
Private Const CHINESESIMPLIFIED_CHARSET As Long = 134&

Private Const VER_PLATFORM_WIN32s        As Long = 0&
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Private Const VER_PLATFORM_WIN32_NT      As Long = 2&

Private Const OFN_HIDEREADONLY      As Long = &H4&
Private Const OFN_NONETWORKBUTTON   As Long = &H20000
Private Const OFN_PATHMUSTEXIST     As Long = &H800&
Private Const OFN_FILEMUSTEXIST     As Long = &H1000&
Private Const OFN_OVERWRITEPROMPT   As Long = &H2&

Private Const SM_CXFULLSCREEN           As Long = 16&
Private Const SM_CYFULLSCREEN           As Long = 17&

Public sRegVals()   As String
Public sFileVals()  As String

Public bAutoSelect      As Boolean
Public bConfirm         As Boolean
Public bMakeBackup      As Boolean
Public bIgnoreSafe      As Boolean
Public bLogProcesses    As Boolean
Public bSkipErrorMsg    As Boolean
Public bMinToTray       As Boolean
Public bStartupList     As Boolean
Public bStartupListSilent As Boolean
Public sHostsFile$
Public bIsWin9x As Boolean
Public bIsWinNT As Boolean
Public bIsWinME As Boolean
Public bIsWinVistaOrLater As Boolean

Public bIsWin64 As Boolean
Public bIsWin32 As Boolean
Public inIDE    As Boolean
Public bForceRU As Boolean
Public bForceEN As Boolean

Public SysDisk          As String
Public sWinDir          As String
Public sWinSysDir       As String
Public sWinSysDirWow64  As String
Public PF_32            As String
Public PF_64            As String
Public PF_32_Common     As String
Public PF_64_Common     As String
Public AppData          As String
Public LocalAppData     As String
Public Desktop          As String
Public UserProfile      As String
Public AllUsersProfile  As String
Public TempCU           As String
Public colProfiles      As Collection

Private sIgnoreList()   As String
Public bDebugMode       As Boolean
Public sWinVersion      As String
Public bRebootNeeded    As Boolean
Public bUpdatePolicyNeeded As Boolean
Public DisableSubclassing As Boolean

Public bIsUSADateFormat As Boolean
Public bNoWriteAccess   As Boolean
Public bSeenLSPWarning  As Boolean

Public sSafeLSPFiles        As String
Public sSafeProtocols()     As String
Public sSafeRegDomains()    As String
Public sSafeSSODL()         As String
Public sSafeFilters()       As String
Public sSafeAppInit         As String
Public sSafeWinlogonNotify  As String

Public AppVer               As String
Public ForkVer              As String
Public AppVerString         As String
Public StartupListVer       As String
Public ADSspyVer            As String
Public ProcManVer           As String
Public sProgramVersion      As String  'encryption phrase
Public ErrReport            As String  'report of all errors during scan process

Public bShownBHOWarning     As Boolean
Public bShownToolbarWarning As Boolean

Public bMD5                 As Boolean
Public bIgnoreAllWhitelists As Boolean
Public bHideMicrosoft       As Boolean
Public bAutoLog             As Boolean
Public bAutoLogSilent       As Boolean
Public bLogEnvVars          As Boolean

Public bSeenHostsFileAccessDeniedWarning As Boolean
Public bGlobalDontFocusListBox As Boolean

Public g_DEFSTARTPAGE       As String
Public g_DEFSEARCHPAGE      As String
Public g_DEFSEARCHASS       As String
Public g_DEFSEARCHCUST      As String
Public g_UninstallState     As Boolean  'HJT is beeing uninstalled
Public g_ProgressMaxTags    As Long     'last progressbar tag number (count of items)
Public g_HJT_Items_Count    As Long

Public FixLog               As String   'future use

Public gProcess()           As MY_PROC_ENTRY
Public g_TasksWL()          As TASK_WHITELIST_ENTRY
Public oDict                As DICTIONARIES

Public oScrTaskWL_ID        As clsTrickHashTable

Public Scan()   As TYPE_Scan_Results    '// Dragokas - plan to use it instead of parsing lines from result screen.
                                        '// User type structures of arrays will be filled together with using of method frmMain.lstResults.AddItem
                                        '// It is much effectively and will have Unicode support (native vb6 ListBox is ANSI only.)
                                        '// Should replace it with Forms 2.0 or throught the API CreateWindowEx) in future.
Public OSver    As TYPE_MY_OSVERSION
Public OSInfo   As clsOSInfo
Public Proc     As clsProcess
Public Perf     As Perfomance_TYPE

'it map ANSI scan result string from ListBox to Unicode string that is stored in memory (TYPE_Scan_Results structure)
Public Function GetScanResults(HitLineA As String, Result As TYPE_Scan_Results) As Boolean
    Dim i As Long
    For i = 1 To UBound(Scan)
        If HitLineA = Scan(i).HitLineA Then
            Result = Scan(i)
            GetScanResults = True
            Exit Function
        End If
    Next
    'Cannot find appropriate cure item for:, "Error"
    MsgBoxW Translate(592) & vbCrLf & HitLineA, vbCritical, Translate(591)
End Function

' it add Unicode TYPE_Scan_Results structure to shared array
Public Sub AddToScanResults(Result As TYPE_Scan_Results, Optional DoNotAddToListBox As Boolean)
    'LockWindowUpdate frmMain.lstResults.hwnd
    DoEvents
    If Not DoNotAddToListBox Then
        frmMain.lstResults.AddItem Result.HitLineW
        'select the last added line
        frmMain.lstResults.ListIndex = frmMain.lstResults.ListCount - 1
    End If
    ReDim Preserve Scan(UBound(Scan) + 1)
    'Unicode to ANSI mapping (dirty hack)
    Result.HitLineA = frmMain.lstResults.List(frmMain.lstResults.ListCount - 1)
    Scan(UBound(Scan)) = Result
    'Sleep 5
    'LockWindowUpdate False
End Sub

Public Sub AddToScanResultsSimple(Section As String, HitLine As String, Optional DoNotAddToListBox As Boolean)
    Dim Result As TYPE_Scan_Results
    With Result
        .Section = Section
        .HitLineW = HitLine
    End With
    AddToScanResults Result, DoNotAddToListBox
End Sub

Public Sub GetHosts()
    If bIsWinNT Then
        sHostsFile = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DataBasePath")
        'sHostsFile = replace$(sHostsFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
        sHostsFile = EnvironW(sHostsFile) & "\hosts"
    End If
End Sub

Public Sub LoadStuff()
    On Error GoTo ErrorHandler:
    
    Dim i As Long
    
    '=== LOAD FILEVALS ===
    'syntax:
    ' inifile,section,value,resetdata,baddata
    ' |       |       |     |         |
    ' |       |       |     |         5) data that shouldn't be (never used)
    ' |       |       |     4) data to reset to
    ' |       |       |        (delete all if empty)
    ' |       |       3) value to check
    ' |       2) section to check
    ' 1) file to check
    
    Dim colFileVals As Collection
    Set colFileVals = New Collection
    
    With colFileVals
        .Add "system.ini,boot,Shell,explorer.exe,"
        .Add "win.ini,windows,load,,"
        .Add "win.ini,windows,run,,"
        .Add "REG:system.ini,boot,Shell,explorer.exe|%WINDIR%\explorer.exe,"
        .Add "REG:win.ini,windows,load,,"
        .Add "REG:win.ini,windows,run,,"
        .Add "REG:system.ini,boot,UserInit,%WINDIR%\System32\UserInit.exe,"
    End With
    ReDim sFileVals(colFileVals.Count - 1)
    For i = 1 To colFileVals.Count
        sFileVals(i - 1) = colFileVals.Item(i)
    Next


    '=== LOAD REGVALS ===
    'syntax:
    '  regkey,regvalue,resetdata,baddata
    '  |      |        |          |
    '  |      |        |          data that shouldn't be (never used)
    '  |      |        R0 - data to reset to
    '  |      R1 - value to check
    '  R2 - regkey to check
    '
    'when empty:
    'R0 - everything is considered bad (always used), change to resetdata
    'R1 - value being present is considered bad, delete value
    'R2 - key being present is considered bad, delete key (not used)
    
    Dim colRegIE As Collection
    Set colRegIE = New Collection
    
    Dim Hive
    Dim Default_Page_URL$: Default_Page_URL = "http://go.microsoft.com/fwlink/p/?LinkId=255141"
    Dim Default_Search_URL$: Default_Search_URL = "http://go.microsoft.com/fwlink/?LinkId=54896"
    
    With colRegIE
      For Each Hive In Array("HKCU", "HKLM", "HKU\.DEFAULT")
    
        .Add Hive & "\Software\Microsoft\Internet Explorer,Default_Page_URL," & Default_Page_URL & "|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Default_Page_URL," & Default_Page_URL & "|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Search,Default_Page_URL," & Default_Page_URL & "|,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer,Default_Search_URL," & Default_Search_URL & "|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Default_Search_URL," & Default_Search_URL & "|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Search,Default_Search_URL," & Default_Search_URL & "|,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer,SearchAssistant,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer,CustomizeSearch,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer,Search,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer,Search Bar,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer,Search Page,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer,Start Page,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer,SearchURL,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer,(Default),,"
        .Add Hive & "\Software\Microsoft\Internet Explorer,www,,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,SearchAssistant,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,CustomizeSearch,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Search Bar,http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Search Page,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Start Page,$DEFSTARTPAGE|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,SearchURL,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Start Page Redirect Cache,http://ru.msn.com/?ocid=iehp|,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer\Search,SearchAssistant,$DEFSEARCHASS|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Search,CustomizeSearch,$DEFSEARCHCUST|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Search,(Default),,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer\SearchURL,(Default),,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\SearchURL,SearchURL,,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Startpagina,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,First Home Page,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Local Page,%SystemRoot%\System32\blank.htm|%SystemRoot%\SysWOW64\blank.htm|%11%\blank.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Start Page_bak,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,HomeOldSP,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,YAHOOSubst,,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Window Title,,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Extensions Off Page,about:NoAdd-ons|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\Main,Security Risk Page,about:SecurityRisk|,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,blank,res://mshtml.dll/blank.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,DesktopItemNavigationFailure,res://ieframe.dll/navcancl.htm|res://shdoclc.dll/navcancl.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,InPrivate,res://ieframe.dll/inprivate.htm|res://ieframe.dll/inprivate_win7.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,NavigationCanceled,res://ieframe.dll/navcancl.htm|res://shdoclc.dll/navcancl.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,NavigationFailure,res://ieframe.dll/navcancl.htm|res://shdoclc.dll/navcancl.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,NoAdd-ons,res://ieframe.dll/noaddon.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,NoAdd-onsInfo,res://ieframe.dll/noaddoninfo.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,PostNotCached,res://ieframe.dll/repost.htm|res://mshtml.dll/repost.htm|,"
        .Add Hive & "\Software\Microsoft\Internet Explorer\AboutURLs,SecurityRisk,res://ieframe.dll/securityatrisk.htm|,"

        .Add Hive & "\Software\Microsoft\Internet Connection Wizard,ShellNext,,"
        
        .Add Hive & "\Software\Microsoft\Internet Explorer\Toolbar,LinksFolderName,Links|Ссылки|,"
        
      Next
        
        'Only short hive names permitted here !
        
        .Add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings,AutoConfigURL,,"
        .Add "HKLM\Software\Microsoft\Windows\CurrentVersion\Internet Settings,AutoConfigURL,,"
        .Add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ProxyServer,,"
        .Add "HKLM\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ProxyServer,,"
        .Add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ProxyOverride,,"
        .Add "HKLM\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ProxyOverride,,"
        
        .Add "HKLM\System\CurrentControlSet\services\NlaSvc\Parameters\Internet\ManualProxies,(Default),,"
        .Add "HKLM\SOFTWARE\Clients\StartMenuInternet\IEXPLORE.EXE\shell\open\command,(Default)," & _
            IIf(bIsWin64, "%ProgramW6432%", "%ProgramFiles%") & "\Internet Explorer\iexplore.exe" & _
            IIf(bIsWin64, "|%ProgramFiles(x86)%\Internet Explorer\iexplore.exe", "") & _
            "|" & """" & "%ProgramFiles%\Internet Explorer\iexplore.exe" & """" & ","
        
        'Note: if you would like to add x64 shared key here (which do not redirected), you should specify it directly on CheckO1item function (look at example: HKLM\SOFTWARE\Clients).
        
    End With
    ReDim sRegVals(colRegIE.Count - 1)
    For i = 1 To colRegIE.Count
        sRegVals(i - 1) = colRegIE.Item(i)
    Next
    
    
    ' === LOAD NONSTANDARD-BUT-SAFE-DOMAINS LIST ===
    
    Dim colSafeRegDomains As Collection
    Set colSafeRegDomains = New Collection
    
    With colSafeRegDomains
        .Add "http://www.microsoft.com"
        .Add "http://home.microsoft.com"
        .Add "http://www.msn.com"
        .Add "http://search.msn.com"
        .Add "http://ie.search.msn.com"
        .Add "ie.search.msn.com"
        .Add "<local>"
        .Add "http://www.google.com"
        .Add "127.0.0.1;localhost"
        .Add "about:blank"
        .Add "http://go.microsoft.com/"
        .Add "www.microsoft.com/"
        ' "iexplore"
        ' "http://www.aol.com"
    End With
    ReDim sSafeRegDomains(colSafeRegDomains.Count - 1)
    For i = 1 To colSafeRegDomains.Count
        sSafeRegDomains(i - 1) = colSafeRegDomains.Item(i)
    Next


    ' === LOAD LSP PROVIDERS SAFELIST ===
    'asterisk is used for filename separation, because.
    'did you ever see a filename with an asterisk?
    sSafeLSPFiles = "*A2antispamlsp.dll*Adlsp.dll*Agbfilt.dll*Antiyfilter.dll*Ao2lsp.dll*Aphish.dll*Asdns.dll*Aslsp.dll*Asnsp.dll*Avgfwafu.dll*Avsda.dll*Betsp.dll*Biolsp.dll*Bmi_lsp.dll*Caslsp.dll*Cavemlsp.dll*Cdnns.dll*Connwsp.dll*Cplsp.dll*Csesck32.dll*Cslsp.dll*Cssp.al*Ctxlsp.dll*Ctxnsp.dll*Cwhook.dll*Cwlsp.dll*Dcsws2.dll*Disksearchservicestub.dll*Drwebsp.dll*Drwhook.dll*Espsock2.dll*Farlsp.dll*Fbm.dll*Fbm_lsp.dll*Fortilsp.dll*Fslsp.dll*Fwcwsp.dll*Fwtunnellsp.dll*Gapsp.dll*Googledesktopnetwork1.dll*Hclsock5.dll*Iapplsp.dll*Iapp_lsp.dll*Ickgw32i.dll*Ictload.dll*Idmmbc.dll*Iga.dll*Imon.dll*Imslsp.dll*Inetcntrl.dll*Ippsp.dll*Ipsp.dll*Iss_clsp.dll*Iss_slsp.dll*Kvwsp.dll*Kvwspxp.dll*Lslsimon.dll*Lsp32.dll*" & _
        "Lspcs.dll*Mclsp.dll*Mdnsnsp.dll*Msafd.dll*Msniffer.dll*Mswsock.dll*Mswsosp.dll*Mwtsp.dll*Mxavlsp.dll*Napinsp.dll*Nblsp.dll*Ndpwsspr.dll*Netd.dll*Nihlsp.dll*Nlaapi.dll*Nl_lsp.dll*Nnsp.dll*Normanpf.dll*Nutafun4.dll*Nvappfilter.dll*Nwws2nds.dll*Nwws2sap.dll*Nwws2slp.dll*Odsp.dll*Pavlsp.dll*Pclsp.dll*Pctlsp.dll*Pfftsp.dll*Pgplsp.dll*Pidlsp.dll*Pnrpnsp.dll*Prifw.dll*Proxy.dll*Prplsf.dll*Pxlsp.dll*Rnr20.dll*Rsvpsp.dll*S5spi.dll*Samnsp.dll*Sarah.dll*Scopinet.dll*Skysocks.dll*Sliplsp.dll*Smnsp.dll*Spacklsp.dll*Spampallsp.dll*Spi.dll*Spidll.dll*Spishare.dll*Spsublsp.dll*Sselsp.dll*Stplayer.dll*Syspy.dll*Tasi.dll*Tasp.dll*Tcpspylsp.dll*Ua_lsp.dll*Ufilter.dll*Vblsp.dll*Vetredir.dll*Vlsp.dll*Vnsp.dll*" & _
        "Wglsp.dll*Whllsp.dll*Whlnsp.dll*Winrnr.dll*Wins4f.dll*Winsflt.dll*WinSysAM.dll*Wps.dll*Wshbth.dll*Wspirda.dll*Wspwsp.dll*Xfilter.dll*xfire_lsp.dll*Xnetlsp.dll*Ypclsp.dll*Zklspr.dll*_Easywall.dll*_Handywall.dll*vsocklib.dll*"
    
    ' === LOAD PROTOCOL SAFELIST === (O18)
    
    Dim colSafeProtocols As Collection
    Set colSafeProtocols = New Collection
        
    With colSafeProtocols
        .Add "about|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "belarc|{6318E0AB-2E93-11D1-B8ED-00608CC9A71F}"
        .Add "BPC|{3A1096B3-9BFA-11D1-AE77-00C04FBBDEBC}"
        .Add "CDL|{3DD53D40-7B8B-11D0-B013-00AA0059CE02}"
        .Add "cdo|{CD00020A-8B95-11D1-82DB-00C04FB1625D}"
        .Add "copernicagentcache|{AAC34CFD-274D-4A9D-B0DC-C74C05A67E1D}"
        .Add "copernicagent|{A979B6BD-E40B-4A07-ABDD-A62C64A4EBF6}"
        .Add "dodots|{9446C008-3810-11D4-901D-00B0D04158D2}"
        .Add "DVD|{12D51199-0DB5-46FE-A120-47A3D7D937CC}"
        .Add "file|{79EAC9E7-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ftp|{79EAC9E3-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "gopher|{79EAC9E4-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "https|{79EAC9E5-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "http|{79EAC9E2-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ic32pp|{BBCA9F81-8F4F-11D2-90FF-0080C83D3571}"
        .Add "ipp|"
        .Add "its|{9D148291-B9C8-11D0-A4CC-0000F80149F6}"
        .Add "javascript|{3050F3B2-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "junomsg|{C4D10830-379D-11D4-9B2D-00C04F1579A5}"
        .Add "lid|{5C135180-9973-46D9-ABF4-148267CBB8BF}"
        .Add "local|{79EAC9E7-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "mailto|{3050F3DA-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "mctp|{D7B95390-B1C5-11D0-B111-0080C712FE82}"
        .Add "mhtml|{05300401-BCBC-11D0-85E3-00C04FD85AB4}"
        .Add "mk|{79EAC9E6-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ms-its50|{F8606A00-F5CF-11D1-B6BB-0000F80149F6}"
        .Add "ms-its51|{F6F1E82D-DE4D-11D2-875C-0000F8105754}"
        .Add "ms-itss|{0A9007C0-4076-11D3-8789-0000F8105754}"
        .Add "ms-its|{9D148291-B9C8-11D0-A4CC-0000F80149F6}"
        .Add "msdaipp|"
        .Add "mso-offdap|{3D9F03FA-7A94-11D3-BE81-0050048385D1}"
        .Add "ndwiat|{13F3EA8B-91D7-4F0A-AD76-D2853AC8BECE}"
        .Add "res|{3050F3BC-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "sysimage|{76E67A63-06E9-11D2-A840-006008059382}"
        .Add "tve-trigger|{CBD30859-AF45-11D2-B6D6-00C04FBBDE6E}"
        .Add "tv|{CBD30858-AF45-11D2-B6D6-00C04FBBDE6E}"
        .Add "vbscript|{3050F3B2-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "vnd.ms.radio|{3DA2AA3B-3D96-11D2-9BD2-204C4F4F5020}"
        .Add "wia|{13F3EA8B-91D7-4F0A-AD76-D2853AC8BECE}"
        .Add "mso-offdap11|{32505114-5902-49B2-880A-1F7738E5A384}"
        .Add "DirectDVD|{85A81A02-336B-43FF-998B-FE8E194FBA4D}"
        .Add "pcn|{D540F040-F3D9-11D0-95BE-00C04FD93CA5}"
        .Add "msencarta|{74D92DF3-6D9D-11D1-8B38-006097DBED7A}"
        .Add "msero|{B0D92A71-886B-453B-A649-1B91F93801E7}"
        .Add "msref|{74D92DF3-6D9D-11D1-8B38-006097DBED7A}"
        .Add "df2|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df3|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df4|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df5|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df23chat|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df5demo|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "ofpjoin|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "saphtmlp|{D1F8BD1E-7967-11D2-B43A-006094B9EADB}"
        .Add "sapr3|{D1F8BD1E-7967-11D2-B43A-006094B9EADB}"
        .Add "lbxfile|{56831180-F115-11D2-B6AA-00104B2B9943}"
        .Add "lbxres|{24508F1B-9E94-40EE-9759-9AF5795ADF52}"
        .Add "cetihpz|{CF184AD3-CDCB-4168-A3F7-8E447D129300}"
        .Add "aim|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "shell|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "asp|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "hsp|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-asp|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-hsp|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-zip|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "zip|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "bega|{A57721C9-B905-49B3-8BCA-B99FBB8C627E}"
        .Add "bt2|{1730B77B-F429-498F-9B15-4514D83C8294}"
        .Add "cetihpz|{CF184AD3-CDCB-4168-A3F7-8E447D129300}"
        .Add "copernicdesktopsearch|{D9656C75-5090-45C3-B27E-436FBC7ACFA7}"
        .Add "crick|{B861500A-A326-11D3-A248-0080C8F7DE1E}"
        .Add "dadb|{82D6F09F-4AC2-11D3-8BD9-0080ADB8683C}"
        .Add "dialux|{8352FA4C-39C6-11D3-ADBA-00A0244FB1A2}"
        .Add "emistp|{0EFAEA2E-11C9-11D3-88E3-0000E867A001}"
        .Add "ezstor|{6344A3A0-96A7-11D4-88CC-000000000000}"
        .Add "flowto|{C7101FB0-28FB-11D5-883A-204C4F4F5020}"
        .Add "g7ps|{9EACF0FB-4FC7-436E-989B-3197142AD979}"
        .Add "intu-res|{9CE7D474-16F9-4889-9BB9-53E2008EAE8A}"
        .Add "iwd|{EA5F5649-A6C7-11D4-9E3C-0020AF0FFB56}"
        .Add "mavencache|{DB47FDC2-8C38-4413-9C78-D1A68BF24EED}"
        .Add "ms-help|{314111C7-A502-11D2-BBCA-00C04F8EC294}"
        .Add "msnim|{828030A1-22C1-4009-854F-8E305202313F}"
        .Add "myrm|{4D034FC3-013F-4B95-B544-44D49ABE3E76}"
        .Add "nbso|{DF700763-3EAD-4B64-9626-22BEEFF3EA47}"
        .Add "nim|{3D206AE2-3039-413B-B748-3ACC562EC22A}"
        .Add "OWC11.mso-offdap|{32505114-5902-49B2-880A-1F7738E5A384}"
        .Add "pcl|{182D0C85-206F-4103-B4FA-DCC1FB0A0A44}"
        .Add "pure-go|{4746C79A-2042-4332-8650-48966E44ABA8}"
        .Add "qrev|{9DE24BAC-FC3C-42C4-9FC4-76B3FAFDBD90}"
        .Add "rmh|{23C585BB-48FF-4865-8934-185F0A7EB84C}"
        .Add "SafeAuthenticate|{8125919B-9BE9-4213-A1D6-75188A22D21E}"
        .Add "sds|{79E0F14C-9C52-4218-89A7-7C4B0563D121}"
        .Add "siteadvisor|{3A5DC592-7723-4EAA-9EE6-AF4222BCF879}"
        .Add "smscrd|{FA3F5003-93D4-11D2-8E48-00A0C98BD8C3}"
        .Add "stibo|{FFAD3420-6D61-44F6-BA25-293F17152D79}"
        .Add "textwareilluminatorbase|{CE5CD329-1650-414A-8DB0-4CBF72FAED87}"
        .Add "widimg|{EE7C2AFF-5742-44FF-BD0E-E521B0D3C3BA}"
        .Add "wlmailhtml|{03C514A3-1EFB-4856-9F99-10D7BE1653C0}"
        .Add "x-atng|{7E8717B0-D862-11D5-8C9E-00010304F989}"
        .Add "x-excid|{9D6CC632-1337-4A33-9214-2DA092E776F4}"
        .Add "x-mem1|{C3719F83-7EF8-4BA0-89B0-3360C7AFB7CC}"
        .Add "x-mem3|{4F6D06DD-44AB-4F89-BF13-9027B505B15A}"
        .Add "ct|{774E529C-2458-48A2-8F57-3ED3105D8612}"
        .Add "cw|{774E529C-2458-48A2-8F57-3ED3105D8612}"
        .Add "eti|{3AAE7392-E7AA-11D2-969E-00105A088846}"
        .Add "livecall|{828030A1-22C1-4009-854F-8E305202313F}"
        .Add "tbauth|{14654CA6-5711-491D-B89A-58E571679951}"
        .Add "windows.tbauth|{14654CA6-5711-491D-B89A-58E571679951}"
    End With
    ReDim sSafeProtocols(colSafeProtocols.Count - 1)
    For i = 1 To colSafeProtocols.Count
        sSafeProtocols(i - 1) = colSafeProtocols.Item(i)
    Next
        
    
    Dim colSafeFilters As Collection    '(O18)
    Set colSafeFilters = New Collection
        
    With colSafeFilters
        .Add "application/octet-stream|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "application/x-complus|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "application/x-msdownload|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "Class Install Handler|{32B533BB-EDAE-11d0-BD5A-00AA00B92AF1}"
        .Add "deflate|{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "gzip|{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "lzdhtml|{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "text/webviewhtml|{733AC4CB-F1A4-11d0-B951-00A0C90312E1}"
        .Add "text/xml|{807553E5-5146-11D5-A672-00B0D022E945}"
        .Add "application/x-icq|{db40c160-09a1-11d3-baf2-000000000000}"
        'added in HJT 1.99.2 final
        .Add "application/msword|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd.ms-excel|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd.ms-powerpoint|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/x-microsoft-rpmsg-message|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd-backup-octet-stream|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "application/vnd-viewer|{CD4527E8-4FC7-48DB-9806-10537B501237}"
        .Add "application/x-bt2|{6E1DDCE8-76BC-4390-9488-806E8FB1AD77}"
        .Add "application/x-internet-signup|{A173B69A-1F9B-4823-9FDA-412F641E65D6}"
        .Add "text/html|{8D42AD12-D7A1-4797-BCB7-AD89E5FCE4F7}"
        .Add "text/html|{F79B2338-A6E7-46D4-9201-422AA6E74F43}"
        .Add "text/x-mrml|{C51721BE-858B-4A66-A8BF-D2882FF49820}"
        .Add "text/xml|{807563E5-5146-11D5-A672-00B0D022E945}"
        .Add "application/octet-stream|{F969FE8E-1937-45AD-AF42-8A4D11CBDC2A}"
        .Add "application/xhtml+xml|{32F66A26-7614-11D4-BD11-00104BD3F987}"
        .Add "text/xml|{32F66A26-7614-11D4-BD11-00104BD3F987}"
    End With
    ReDim sSafeFilters(colSafeFilters.Count - 1)
    For i = 1 To colSafeFilters.Count
        sSafeFilters(i - 1) = colSafeFilters.Item(i)
    Next


    'LOAD APPINIT_DLLS SAFELIST (O20)
    sSafeAppInit = "*aakah.dll*akdllnt.dll*ROUSRNT.DLL*ssohook*KATRACK.DLL*APITRAP.DLL*UmxSbxExw.dll*sockspy.dll*scorillont.dll*wbsys.dll*NVDESK32.DLL*hplun.dll*mfaphook.dll*PAVWAIT.DLL*OCMAPIHK.DLL*MsgPlusLoader.dll*IconCodecService.dll*wl_hook.dll*Google\GOOGLE~1\GOEC62~1.DLL*adialhk.dll*wmfhotfix.dll*interceptor.dll*qaphooks.dll*RMProcessLink.dll*msgrmate.dll*wxvault.dll*ctu33.dll*ati2evxx.dll*vsmvhk.dll*"
    
    'LOAD SSODL SAFELIST (O21)
    
    Dim colSafeSSODL As Collection    '(O18)
    Set colSafeSSODL = New Collection
        
    With colSafeSSODL
        .Add "{E6FB5E20-DE35-11CF-9C87-00AA005127ED}"  'WebCheck: E:\WINDOWS\System32\webcheck.dll (WinAll)
        .Add "{35CEC8A3-2BE6-11D2-8773-92E220524153}"  'SysTray: E:\WINDOWS\System32\stobject.dll (Win2k/XP)
        .Add "{7849596a-48ea-486e-8937-a2a3009f31a9}"  'PostBootReminder: E:\WINDOWS\system32\SHELL32.dll (WinXP)
        .Add "{fbeb8a05-beee-4442-804e-409d6c4515e9}"  'CDBurn: E:\WINDOWS\system32\SHELL32.dll (WinXP)
        .Add "{11566B38-955B-4549-930F-7B7482668782}"  'AUHook: C:\WINDOWS\SYSTEM\AUHOOK.DLL (WinME)
        .Add "{7007ACCF-3202-11D1-AAD2-00805FC1270E}"  'Network.ConnectionTray: C:\WINNT\system32\NETSHELL.dll (Win2k)
        .Add "{e57ce738-33e8-4c51-8354-bb4de9d215d1}"  'UPnPMonitor: C:\WINDOWS\SYSTEM\UPNPUI.DLL (WinME/XP)
        .Add "{BCBCD383-3E06-11D3-91A9-00C04F68105C}"  'AUHook: C:\WINDOWS\SYSTEM\AUHOOK.DLL (WinME)
        .Add "{F5DF91F9-15E9-416B-A7C3-7519B11ECBFC}"  '0aMCPClient: C:\Program Files\StarDock\MCPCore.dll
        .Add "{AAA288BA-9A4C-45B0-95D7-94D524869DB5}"  'WPDShServiceObj   WPDShServiceObj.dll Windows Portable Device Shell Service Object
        .Add "{1799460C-0BC8-4865-B9DF-4A36CD703FF0}" 'IconPackager Repair  iprepair.dll    Stardock\Object Desktop\ ThemeManager
        .Add "{6D972050-A934-44D7-AC67-7C9E0B264220}" 'EnhancedDialog   enhdlginit.dll  EnhancedDialog by Stardock
    End With
    ReDim sSafeSSODL(colSafeSSODL.Count - 1)
    For i = 1 To colSafeSSODL.Count
        sSafeSSODL(i - 1) = colSafeSSODL.Item(i)
    Next
    
    
    'LOAD WINLOGON NOTIFY SAFELIST (O20)
    'second line added in HJT 1.99.2 final
    sSafeWinlogonNotify = "crypt32chain*cryptnet*cscdll*ScCertProp*Schedule*SensLogn*termsrv*wlballoon*igfxcui*AtiExtEvent*wzcnotif*" & _
                          "ActiveSync*atmgrtok*avldr*Caveo*ckpNotify*Command AntiVirus Download*ComPlusSetup*CwWLEvent*dimsntfy*DPWLN*EFS*FolderGuard*GoToMyPC*IfxWlxEN*igfxcui*IntelWireless*klogon*LBTServ*LBTWlgn*LMIinit*loginkey*MCPClient*MetaFrame*NavLogon*NetIdentity Notification*nwprovau*OdysseyClient*OPXPGina*PCANotify*pcsinst*PFW*PixVue*ppeclt*PRISMAPI.DLL*PRISMGNA.DLL*psfus*QConGina*RAinit*RegCompact*SABWinLogon*SDNotify*Sebring*STOPzilla*sunotify*SymcEventMonitors*T3Notify*TabBtnWL*Timbuktu Pro*tpfnf2*tpgwlnotify*tphotkey*VESWinlogon*WB*WBSrv*WgaLogon*wintask*WLogon*WRNotifier*Zboard*zsnotify*sclgntfy"
    
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_LoadStuff"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub StartScan()
    On Error GoTo ErrorHandler:
    
    frmMain.txtNothing.Visible = False
    'frmMain.shpBackground.Tag = iItems
    SetProgressBar g_HJT_Items_Count   'R + F + O25
    
    Call GetProcesses_Zw(gProcess)
    
    Dim sRule$, i&
    'load ignore list
    IsOnIgnoreList ""
    
    frmMain.lstResults.Clear
    
    'Registry
    
    UpdateProgressBar "R"
    For i = 0 To UBound(sRegVals)
        'If sRegVals(i) = "" Then Exit For
        ProcessRuleReg sRegVals(i)
    Next i
    
    CheckRegistry3Item
    
    UpdateProgressBar "F"
    'File
    For i = 0 To UBound(sFileVals)
        If sFileVals(i) <> "" Then
            CheckFileItems sFileVals(i)
        End If
    Next i
    
    'Netscape/Mozilla stuff
    'CheckNetscapeMozilla        'N1-4
    
    
    'Other options
    UpdateProgressBar "O1"
    CheckO1Item
    CheckO1Item_ICS
    CheckO1Item_DNSApi
    UpdateProgressBar "O2"
    CheckO2Item
    UpdateProgressBar "O3"
    CheckO3Item
    UpdateProgressBar "O4"
    CheckO4Item
    UpdateProgressBar "O5"
    CheckO5Item
    UpdateProgressBar "O6"
    CheckO6Item
    UpdateProgressBar "O7"
    CheckO7Item
    UpdateProgressBar "O8"
    CheckO8Item
    UpdateProgressBar "O9"
    CheckO9Item
    UpdateProgressBar "O10"
    CheckO10Item
    UpdateProgressBar "O11"
    CheckO11Item
    UpdateProgressBar "O12"
    CheckO12Item
    UpdateProgressBar "O13"
    CheckO13Item
    UpdateProgressBar "O14"
    CheckO14Item
    UpdateProgressBar "O15"
    CheckO15Item
    UpdateProgressBar "O16"
    CheckO16Item
    UpdateProgressBar "O17"
    CheckO17Item
    UpdateProgressBar "O18"
    CheckO18Item
    UpdateProgressBar "O19"
    CheckO19Item
    UpdateProgressBar "O20"
    CheckO20Item
    UpdateProgressBar "O21"
    CheckO21Item
    UpdateProgressBar "O22"
    CheckO22Item
    UpdateProgressBar "O23"
    CheckO23Item
    'added in HJT 1.99.2: Desktop Components
    UpdateProgressBar "O24"
    CheckO24Item
    '2.0.7 - WMI Events
    UpdateProgressBar "O25"
    CheckO25Item
    UpdateProgressBar "ProcList"
    
    
    With frmMain
        .lblMD5.Visible = False
        '.lblInfo(1).Visible = True
        '.picPaypal.Visible = True
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
    
ErrorHandler:
    ErrorMsg err, "modMain_StartScan"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub SetProgressBar(lMaxTags As Long)
    
    'ProgressBar label settings
    frmMain.lblStatus.Visible = True
    frmMain.lblStatus.Caption = ""
    frmMain.lblStatus.ForeColor = &HFFFF&   'Yellow
    frmMain.lblStatus.ZOrder 0 'on top
    frmMain.lblStatus.Left = 400
    
    'results label -> off
    frmMain.lblInfo(0).Visible = False
    
    'program description label -> off
    frmMain.lblInfo(1).Visible = False
    
    frmMain.shpBackground.Visible = True
    With frmMain.shpProgress
        .Tag = "0"
        .Visible = True
    End With
    frmMain.shpProgress.Width = 255 ' default
    frmMain.shpProgress.ZOrder 1
    frmMain.shpBackground.ZOrder 1
    g_ProgressMaxTags = lMaxTags
End Sub

Public Sub CloseProgressbar()
    frmMain.shpBackground.Visible = False
    'frmMain.lblInfo(0).Visible = True
    frmMain.lblInfo(1).Visible = True
    frmMain.shpProgress.Visible = False
    frmMain.lblStatus.Visible = False
End Sub

Public Sub UpdateProgressBar(Section As String)
    On Error GoTo ErrorHandler:
    
    Dim lTag As Long
    
    With frmMain
    
        If Not IsNumeric(.shpProgress.Tag) Then .shpProgress.Tag = "0"
        lTag = .shpProgress.Tag
        lTag = lTag + 1
        .shpProgress.Tag = lTag
        
        Select Case Section
            Case "R", "R0", "R1", "R2", "R3": .lblStatus.Caption = Translate(230) & "..."
            Case "F", "F1", "F2", "F3": .lblStatus.Caption = Translate(231) & "..."
            'Case 3: .lblStatus.Caption = Translate(232) & "..."
            Case "O1": .lblStatus.Caption = Translate(233) & "..."
            Case "O2": .lblStatus.Caption = Translate(234) & "..."
            Case "O3": .lblStatus.Caption = Translate(235) & "..."
            Case "O4": .lblStatus.Caption = Translate(236) & "..."
            Case "O5": .lblStatus.Caption = Translate(237) & "..."
            Case "O6": .lblStatus.Caption = Translate(238) & "..."
            Case "O7": .lblStatus.Caption = Translate(239) & "..."
            Case "O8": .lblStatus.Caption = Translate(240) & "..."
            Case "O9": .lblStatus.Caption = Translate(241) & "..."
            Case "O10": .lblStatus.Caption = Translate(242) & "..."
            Case "O11": .lblStatus.Caption = Translate(243) & "..."
            Case "O12": .lblStatus.Caption = Translate(244) & "..."
            Case "O13": .lblStatus.Caption = Translate(245) & "..."
            Case "O14": .lblStatus.Caption = Translate(246) & "..."
            Case "O15": .lblStatus.Caption = Translate(247) & "..."
            Case "O16": .lblStatus.Caption = Translate(248) & "..."
            Case "O17": .lblStatus.Caption = Translate(249) & "..."
            Case "O18": .lblStatus.Caption = Translate(250) & "..."
            Case "O19": .lblStatus.Caption = Translate(251) & "..."
            Case "O20": .lblStatus.Caption = Translate(252) & "..."
            Case "O21": .lblStatus.Caption = Translate(253) & "..."
            Case "O22": .lblStatus.Caption = Translate(254) & "..."
            Case "O23": .lblStatus.Caption = Translate(255) & "..."
            Case "O24": .lblStatus.Caption = Translate(257) & "..."
            Case "O25": .lblStatus.Caption = Translate(258) & "..."
            
            Case "ProcList": .lblStatus.Caption = Translate(260) & "...": Exit Sub
            Case "Backup":   .lblStatus.Caption = Translate(259) & "...": .shpProgress.Width = 255: Exit Sub
            Case "Finish":   .lblStatus.Caption = Translate(256): .shpProgress.Width = .shpBackground.Width + .shpBackground.Left - .shpProgress.Left: Exit Sub
        End Select
        
        .shpProgress.Width = .shpBackground.Width * (lTag / g_ProgressMaxTags)  'g_ProgressMaxTags = items to check or fix -1
        
        '.lblStatus.Refresh
        '.Refresh
    End With
    DoEvents
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_UpdateProgressBar", "shpProgress.Tag=", frmMain.shpProgress.Tag
    If inIDE Then Stop: Resume Next
End Sub


'CheckR0item
'CheckR1item
'CheckR2item
Private Sub ProcessRuleReg(ByVal sRule$)
    On Error GoTo ErrorHandler:
    Dim vRule As Variant, iMode&, i&, bIsNSBSD As Boolean, Result As TYPE_Scan_Results
    Dim sHit$, sKey$, sParam$, sData$, sDefDataStrings$, Wow6432Redir As Boolean, UseWow, aDef() As String
    
    'If InStr(1, sRule, "HKLM\Software\Microsoft\Internet Explorer\Main,Local Page", 1) <> 0 Then Stop
    
    'Registry rule syntax:
    '[regkey],[regvalue],[infected data],[default data]
    '* [regkey]           = "" -> abort - no way man!
    ' * [regvalue]        = "" -> delete entire key
    '  * [default data]   = "" -> delete value
    '   * [infected data] = "" -> any value (other than default) is considered infected
    vRule = Split(sRule, ",")
    
    If UBound(vRule) <> 3 Or Left$(CStr(vRule(0)), 2) <> "HK" Then
        'decryption failed or spelling error
        Stop 'debug.Assert
        Exit Sub
    End If
    
    ' iMode = 0 -> check if value is infected
    ' iMode = 1 -> check if value is present
    ' iMode = 2 -> check if regkey is present
    If CStr(vRule(0)) = vbNullString Then Exit Sub
    If CStr(vRule(3)) = vbNullString Then iMode = 0
    If CStr(vRule(2)) = vbNullString Then iMode = 1
    If CStr(vRule(1)) = vbNullString Then iMode = 2
    
    sKey = vRule(0)
    sParam = vRule(1)
    If sParam = "(Default)" Then sParam = vbNullString
    sDefDataStrings = vRule(2)
    
    'x32 -> 1 cycle
    'x64 + HKCU -> 1 cycle
    'x64 + HKU  -> 1 cycle
    'x64 + HKLM\System -> 1 cycle (shared keys)
    'x64 + HKLM\SOFTWARE\Clients -> 1 cycle (shared key)
    For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If (bIsWin32 And Wow6432Redir) _
            Or bIsWin64 And Wow6432Redir And _
            (Left$(sKey, 4) = "HKCU" _
            Or Left$(sKey, 4) = "HKU\" _
            Or StrBeginWith(sKey, "HKLM\System\") _
            Or StrBeginWith(sKey, "HKLM\SOFTWARE\Clients")) Then Exit For
    
        Select Case iMode
        
        Case 0 'check for incorrect value
            sData = RegGetString(0&, sKey, sParam, Wow6432Redir)
            sData = EnvironW(sData)

            If Not inArraySerialized(sData, sDefDataStrings, "|", , , 1) Then
                bIsNSBSD = False
                If bIgnoreSafe Then bIsNSBSD = StrBeginWithArray(sData, sSafeRegDomains)
                If Not bIsNSBSD Then
                    If InStr(1, sData, "%2e", 1) > 0 Then sData = UnEscape(sData)
                    
                    sHit = IIf(bIsWin32, "R0 - ", IIf(Wow6432Redir, "R0-32 - ", "R0 - ")) & sKey & "," & sParam & " = " & sData
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "R0"
                            .HitLineW = sHit
                            ReDim .RegKey(0)
                            .RegKey(0) = sKey
                            .RegParam = sParam
                            .DefaultData = SplitSafe(sDefDataStrings, "|")(0)
                            .Redirected = Wow6432Redir
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
            
        Case 1  'check for present value
            sData = RegGetString(0&, sKey, sParam, Wow6432Redir)
            If 0 <> Len(sData) Then
                'check if domain is on safe list
                bIsNSBSD = False
                If bIgnoreSafe Then bIsNSBSD = StrBeginWithArray(sData, sSafeRegDomains)
                'make hit
                If Not bIsNSBSD Then
                    If InStr(1, sData, "%2e", 1) > 0 Then sData = UnEscape(sData)
                    
                    sHit = IIf(bIsWin32, "R1 - ", IIf(Wow6432Redir, "R1-32 - ", "R1 - ")) & _
                        sKey & "," & IIf(sParam = "", "(Default)", sParam) & IIf(sData <> "", " = " & sData, "")
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "R1"
                            .HitLineW = sHit
                            ReDim .RegKey(0)
                            .RegKey(0) = sKey
                            .RegParam = sParam
                            .Redirected = Wow6432Redir
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
            
        Case 2 'check if regkey is present
            If RegKeyExists(0&, sKey, Wow6432Redir) Then
            
                sHit = IIf(bIsWin32, "R2 - ", IIf(Wow6432Redir, "R2-32 - ", "R2 - ")) & sKey
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "R2"
                            .HitLineW = sHit
                            ReDim .RegKey(0)
                            .RegKey(0) = sKey
                            .Redirected = Wow6432Redir
                        End With
                        AddToScanResults Result
                    End If
            End If
        End Select
    Next
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_ProcessRuleReg", "sRule=", sRule
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixRegItem(sItem$)
    'R0 - HKCU\Software\..\Main,Window Title
    'R1 - HKCU\Software\..\Main,Window Title=MSIE 5.01
    'R2 - HKCU\Software\..\Main
    
    On Error GoTo ErrorHandler:
    Dim lHive&, sKey$, sValue$, i&, sFixed$, sDummy$, Result As TYPE_Scan_Results
    
    If Not GetScanResults(sItem, Result) Then Exit Sub
    
    With Result
      Select Case .Section
      Case "R0"
        'restore value
        RegSetStringVal 0&, .RegKey(0), .RegParam, CStr(.DefaultData), .Redirected
      Case "R1"
        'delete value
        RegDelVal 0&, .RegKey(0), .RegParam, .Redirected
      Case "R2"
        'delete key
        RegDelKey 0&, .RegKey(0), .Redirected
      End Select
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixRegItem", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckFileItems(ByVal sRule$)
    Dim vRule As Variant, iMode&, sHit$, Result As TYPE_Scan_Results
    Dim sFile$, sSection$, sParam$, sData$, sLegitData$, UseWow, Wow6432Redir As Boolean
    Dim aHive
    On Error GoTo ErrorHandler:
    'IniFile rule syntax:
    '[inifile],[section],[value],[default data],[infected data]
    '* [inifile]          = "" -> abort
    ' * [section]         = "" -> abort
    '  * [value]          = "" -> abort
    '   * [default data]  = "" -> delete if found
    '    * [infected data]= "" -> fix if infected
    
    'decrypt rule
    'sRule = Crypt(sRule, sProgramVersion)
    
    'Checking white list rules
    '1-st token should contains .ini
    'total number of tokens should be 5 (0 to 4)
    
    vRule = Split(sRule, ",")
    If UBound(vRule) <> 4 Or InStr(CStr(vRule(0)), ".ini") = 0 Then
        'spelling error or decrypting error
        Exit Sub
    End If
    
    '1,2,3 tokens should not be empty
    '4-th token is empty -> check if value is present     (F1)
    '4-th token is present -> check if value is infected  (F0)
    
    'File checking rules:
    '
    'example:
    '--------------
    '1. system.ini    (file)
    '2. boot          (section)
    '3. Shell         (parameter)
    '4. explorer.exe  (data / value)
    
    If CStr(vRule(0)) = vbNullString Then Exit Sub
    If CStr(vRule(1)) = vbNullString Then Exit Sub
    If CStr(vRule(2)) = vbNullString Then Exit Sub
    If CStr(vRule(4)) = vbNullString Then iMode = 0
    If CStr(vRule(3)) = vbNullString Then iMode = 1
    
    sFile = vRule(0)
    sSection = vRule(1)
    sParam = vRule(2)
    sLegitData = vRule(3)
    
    'Registry checking rules (prefix REG: on 1-st token)
    '
    'example:
    '1. REG:system.ini ()
    '2. boot           (section)
    '3. Shell          (parameter)
    '4. explorer.exe   (data / value)
    
    'if 4-th token is empty -> check if value is present, in the Registry      (F3)
    'if 4-th token is present -> check if value is infected, in the Registry   (F2)
    
    ' adding char "," to value 'UserInit'
    If InStr(1, sLegitData, "UserInit", 1) <> 0 Then sLegitData = sLegitData & "|" & sLegitData & ","
    
    If Left$(sFile, 3) = "REG" Then
        'skip Win9x
        If Not bIsWinNT Then Exit Sub
        If CStr(vRule(4)) = vbNullString Then iMode = 2
        If CStr(vRule(3)) = vbNullString Then iMode = 3
    End If
    
    'iMode:
    ' F0 = check if value is infected (file)
    ' F1 = check if value is present (file)
    ' F2 = check if value is infected, in the Registry
    ' F3 = check if value is present, in the Registry
    
    Select Case iMode
        Case 0
            'F0 = check if value is infected (file)
            'sValue = String$(255, " ")
            'GetPrivateProfileString CStr(vRule(1)), CStr(vRule(2)), "", sValue, 255, CStr(vRule(0))
            'sValue = Rtrim$(sValue)
            sData = IniGetString$(sFile, sSection, sParam)
            sData = RTrimNull(sData)
            
            If Not inArraySerialized(sData, sLegitData, "|", , , vbTextCompare) Then
                If bIsWinNT And Trim$(sData) <> vbNullString Then
                    sHit = "F0 - " & sFile & ": " & sParam & "=" & sData
                    If Not IsOnIgnoreList(sHit) Then
                        If bMD5 Then sHit = sHit & " (" & GetFileFromAutostart(sData) & ")"
                        With Result
                            .Section = "F0"
                            .HitLineW = sHit
                            .RunObject = sFile
                            ReDim .RegKey(0)
                            .RegKey(0) = sSection
                            .RegParam = sParam
                            .DefaultData = SplitSafe(sLegitData, "|")(0)
                            .CureType = FILE_BASED
                        End With
                    AddToScanResults Result
                    End If
                End If
            End If
            
        Case 1
            'F1 = check if value is present (file)
            'sValue = String$(255, " ")
            'GetPrivateProfileString CStr(vRule(1)), CStr(vRule(2)), "", sValue, 255, CStr(vRule(0))
            'sValue = Rtrim$(sValue)
            sData = IniGetString$(sFile, sSection, sParam)
            sData = RTrimNull(sData)
            
            If Trim$(sData) <> vbNullString Then
                sHit = "F1 - " & sFile & ": " & sParam & "=" & sData
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & " (" & GetFileFromAutostart(sData) & ")"
                    With Result
                        .Section = "F1"
                        .HitLineW = sHit
                        .RunObject = sFile
                        ReDim .RegKey(0)
                        .RegKey(0) = sSection
                        .RegParam = sParam
                        .DefaultData = ""
                        .CureType = FILE_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
            
        Case 2
          'F2 = check if value is infected, in the Registry
          'so far F2 is only reg:Shell and reg:UserInit

'        For Each aHive In Array("HKLM", "HKCU")
'
'          For Each UseWow In Array(False, True)
'            Wow6432Redir = UseWow
'            If bIsWin32 And Wow6432Redir Then Exit For
'            If aHive = "HKCU" And Wow6432Redir Then Exit For
            
            aHive = "HKLM"
            Wow6432Redir = False
            
            sData = RegGetString(0&, aHive & "\Software\Microsoft\Windows NT\CurrentVersion\WinLogon", sParam, Wow6432Redir)
            
            If Not inArraySerialized(sData, sLegitData, "|", , , vbTextCompare) Then
                sHit = IIf(bIsWin32, "F2 - ", IIf(Wow6432Redir, "F2-32 - ", "F2 - ")) & sFile & ": " & aHive & "\..\" & sParam & "=" & sData
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & " (" & GetFileFromAutostart(sData) & ")"
                    With Result
                        .Section = "F2"
                        .HitLineW = sHit
                        .RunObject = sFile
                        ReDim .RegKey(0)
                        .RegKey(0) = aHive & "\Software\Microsoft\Windows NT\CurrentVersion\WinLogon"
                        .RegParam = sParam
                        .DefaultData = SplitSafe(sLegitData, "|")(0)
                        '.RunObject
                        .CureType = REGISTRY_PARAM_BASED
                        .Redirected = Wow6432Redir
                    End With
                    AddToScanResults Result
                End If
            End If
          
          'Next
        'Next
            
        Case 3
          'F3 = check if value is present, in the Registry
          'this is not really smart when more INIFile items get
          'added, but so far F3 is only reg:load and reg:run
        
        For Each aHive In Array("HKLM", "HKCU")
          For Each UseWow In Array(False, True)    ' HKCU (independent key)
            Wow6432Redir = UseWow
            If bIsWin32 And Wow6432Redir Then Exit For
            If aHive = "HKCU" And Wow6432Redir Then Exit For
            
            sData = RegGetString(0&, aHive & "\Software\Microsoft\Windows NT\CurrentVersion\Windows", sParam, Wow6432Redir)
            If sData <> vbNullString Then
                sHit = IIf(bIsWin32, "F3 - ", IIf(Wow6432Redir, "F3-32 - ", "F3 - ")) & sFile & ": " & aHive & "\..\" & sParam & "=" & sData
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & " (" & GetFileFromAutostart(sData) & ")"
                     With Result
                        .Section = "F3"
                        .HitLineW = sHit
                        .RunObject = sFile
                        ReDim .RegKey(0)
                        .RegKey(0) = aHive & "\Software\Microsoft\Windows NT\CurrentVersion\Windows"
                        .RegParam = sParam
                        .DefaultData = ""
                        .CureType = REGISTRY_PARAM_BASED
                        .Redirected = Wow6432Redir
                    End With
                    AddToScanResults Result
                End If
            End If
          Next
        Next
    End Select
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_ProcessRuleIniFile", "sRule=", sRule
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixFileItem(sItem$)
    'F0 - system.ini: Shell=c:\win98\explorer.exe openme.exe
    'F1 - win.ini: load=hpfsch
    On Error GoTo ErrorHandler:
    'coding is easy if you cheat :)
    
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    
    With Result
      Select Case .Section
      
      Case "F0"
        'restore value
        If StrComp(.RunObject, "system.ini", 1) = 0 And StrComp(.RegParam, "Shell", 1) = 0 Then
            'WritePrivateProfileString "boot", "shell", "explorer.exe", "system.ini"
            IniSetString "system.ini", "boot", "shell", "explorer.exe"
        End If
        
      Case "F1"
        'delete value
        If StrComp(.RunObject, "win.ini", 1) = 0 And StrComp(.RegParam, "load", 1) = 0 Then
            'WritePrivateProfileString "windows", "load", "", "win.ini"
            IniSetString "win.ini", "windows", "load", vbNullString
        ElseIf StrComp(.RunObject, "win.ini", 1) = 0 And StrComp(.RegParam, "run", 1) = 0 Then
            'WritePrivateProfileString "windows", "run", "", "win.ini"
            IniSetString "win.ini", "windows", "run", vbNullString
        End If
        
      Case "F2"
        'restore registry value
        If StrComp(.RegParam, "userinit", 1) = 0 Then .DefaultData = .DefaultData & ","
        RegSetStringVal 0&, .RegKey(0), .RegParam, CStr(.DefaultData), .Redirected
      Case "F3"
        'delete registry value
        RegDelVal 0&, .RegKey(0), .RegParam, .Redirected
      End Select
    End With
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_FixFileItem", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

'CheckR3item
Public Sub CheckRegistry3Item()
    On Error GoTo ErrorHandler:

    Dim sURLHook$, hKey&, i&, sName$, sHit$, sCLSID$, sFile$, Result As TYPE_Scan_Results
    Dim vHive As Variant, sKey$, UseWow As Variant, Wow6432Redir As Boolean
    
    sURLHook = "Software\Microsoft\Internet Explorer\URLSearchHooks"
    
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(sURLHook), 0&, KEY_QUERY_VALUE, hKey) = 0 Then
        i = 0
        sCLSID = String$(MAX_VALUENAME, 0&)
        If RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&) <> 0 Then
            sHit = "R3 - Default URLSearchHook is missing"
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "R3", sHit
            RegCloseKey hKey
        End If
    Else
        sHit = "R3 - Default URLSearchHook is missing"
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "R3", sHit
    End If
    
    For Each vHive In Array("HKLM", "HKCU", "HKU\.DEFAULT")
    
      sKey = vHive & "\" & sURLHook
    
      For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If (bIsWin32 And Wow6432Redir) _
            Or bIsWin64 And Wow6432Redir And _
            (Left$(sKey, 4) = "HKCU" _
            Or Left$(sKey, 4) = "HKU\") Then Exit For
    
        If RegOpenKeyExW(0&, StrPtr(sKey), 0&, KEY_QUERY_VALUE, hKey) = 0 Then
        
          i = 0
          Do
            sCLSID = String$(MAX_VALUENAME, 0&)
        
            If RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&) <> 0 Then
                Exit Do
            End If
        
            sCLSID = TrimNull(sCLSID)
            
            If sCLSID <> "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}" Then
                'found a new urlsearchhook!
                sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, False)
                If 0 = Len(sFile) Then sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, True)
                If 0 = Len(sFile) Then sName = "(no name)"
                sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", vbNullString, False)
                If 0 = Len(sFile) Then sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", vbNullString, True)
                If 0 = Len(sFile) Then sFile = "(no file)"
                If sFile <> "(no file)" Then
                    If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
                End If
                
                sHit = IIf(bIsWin32, "R3 - ", IIf(Wow6432Redir, "R3-32 - ", "R3 - ")) & vHive & "\..\URLSearchHooks: " & _
                    sName & " - " & sCLSID & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "R3"
                        .HitLineW = sHit
                        ReDim .RegKey(0)
                        .RegKey(0) = sKey
                        .RegParam = sCLSID
                        .Redirected = Wow6432Redir
                    End With
                    AddToScanResults Result
                End If
            End If
            
            i = i + 1
          Loop
          RegCloseKey hKey
        End If
      Next
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckRegistry3Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixRegistry3Item(sItem$)
    'R3 - Shitty search hook - {00000000} - c:\windows\bho.dll"
    'R3 - Default URLSearchHook is missing
    
    On Error GoTo ErrorHandler:
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    
    If sItem = "R3 - Default URLSearchHook is missing" Then
        RegCreateKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks"
        RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks", "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}", vbNullString
        Exit Sub
    End If
    
    With Result
        RegDelVal 0&, .RegKey(0), .RegParam, .Redirected
    End With
    
    'just in case
    RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\URLSearchHooks", "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}", vbNullString
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_FixRegistry3Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckO1Item_DNSApi()
    On Error GoTo ErrorHandler:
    
    Const MaxSize As Long = 5242880 ' 5 MB.
    
    Dim vFile As Variant, ff As Long, i As Long, size As Currency, p As Long, buf() As Byte, sHit As String, Result As TYPE_Scan_Results
    Dim bufExample() As Byte
    Dim bufExample_2() As Byte
    
    bufExample = StrConv(LCase$("\drivers\etc\hosts"), vbFromUnicode)
    bufExample_2 = StrConv(UCase$("\drivers\etc\hosts"), vbFromUnicode)
    
    ToggleWow64FSRedirection False
    
    For Each vFile In Array(sWinDir & "\system32\dnsapi.dll", sWinDir & "\syswow64\dnsapi.dll")
    
        If OSver.Bitness = "x32" And InStr(1, vFile, "syswow64", 1) <> 0 Then Exit For

        If OpenW(CStr(vFile), FOR_READ, ff) Then
            
            size = LOFW(ff)
            
            If size > MaxSize Then
                ErrorMsg err, "modMain_CheckO1Item_DNSApi", "File is too big: " & vFile & " (Allowed: " & MaxSize & " byte max., current is: " & size & "byte.)"
            ElseIf size > 0 Then
                
                ReDim buf(size - 1)
                
                If GetW(ff, 1, , VarPtr(buf(0)), CLng(size)) Then
                
                    p = InArrSign_NoCase(buf, bufExample, bufExample_2)
                    
                    If p = -1 Then                      'add isMicrosoftFile() ?
                        ' if signature not found
                        sHit = "O1 - DNSApi: File is patched - " & vFile
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
                    End If
                End If
            End If
            CloseW ff
        End If
    Next
    ToggleWow64FSRedirection True
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO1Item_DNSApi"
    If inIDE Then Stop: Resume Next
End Sub

Private Function InArrSign(ArrSrc() As Byte, ArrEx() As Byte) As Long
    Dim i As Long, J As Long, p As Long, Found As Boolean
    InArrSign = -1
    For i = 0 To UBound(ArrSrc) - UBound(ArrEx)
        p = i
        Found = True
        For J = 0 To UBound(ArrEx)
            If ArrSrc(p) <> ArrEx(J) Then Found = False: Exit For
            p = p + 1
        Next
        If Found Then InArrSign = p - UBound(ArrEx) - 1: Exit For
    Next
End Function

Private Function InArrSign_NoCase(ArrSrc() As Byte, ArrEx() As Byte, ArrEx_2() As Byte) As Long
    'ArrEx - all lcase
    'ArrEx_2 - all Ucase
    Dim i As Long, J As Long, p As Long, Found As Boolean
    InArrSign_NoCase = -1
    For i = 0 To UBound(ArrSrc) - UBound(ArrEx)
        p = i
        Found = True
        For J = 0 To UBound(ArrEx)
            If ArrSrc(p) <> ArrEx(J) And ArrSrc(p) <> ArrEx_2(J) Then Found = False: Exit For
            p = p + 1
        Next
        If Found Then InArrSign_NoCase = p - UBound(ArrEx) - 1: Exit For
    Next
End Function

Private Sub CheckO1Item_ICS()
    ' hosts.ics
    'https://support.microsoft.com/ru-ru/kb/309642
    
    On Error GoTo ErrorHandler:
    Dim ff%, sHostsFileICS$, sHostsFileICS_Default$, sHostsICS_Default$, sHit$
    Dim sLines$, sLine As Variant, NonDefaultPath As Boolean, cFileSize As Currency, hFile As Long
    
    If bIsWin9x Then sHostsFileICS_Default = sWinDir & "\hosts.ics"
    If bIsWinNT Then sHostsFileICS_Default = sWinDir & "\System32\drivers\etc\hosts.ics"
    
    sHostsFileICS = sHostsFile & ".ics"
    
    If StrComp(sHostsFileICS, sHostsFileICS_Default) <> 0 Then
        NonDefaultPath = True
    End If
    
    ff = FreeFile()
    
    If NonDefaultPath Then                              'Note: \System32\drivers\etc is not under Wow6432 redirection
        ToggleWow64FSRedirection False, sHostsFileICS
    End If
    
    cFileSize = FileLenW(sHostsFileICS)
    
    ' Size = 0 or just not exists
    If cFileSize = 0 Then
        ToggleWow64FSRedirection True
        If NonDefaultPath Then
            GoTo CheckHostsICS_Default:
        Else
            Exit Sub
        End If
    End If
    
'    On Error Resume Next
'    Open sHostsFileICS For Binary Access Read As #ff
'    If err.Number <> 0 Then
'        sHit = "O1 - Unable to read Hosts.ICS file"
'        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
'        ToggleWow64FSRedirection True
'        If NonDefaultPath Then
'            GoTo CheckHostsICS_Default:
'        Else
'            Exit Sub
'        End If
'    End If
'
'    On Error GoTo ErrorHandler:
'    sLines = String$(cFileSize, 0)
'    Get #ff, , sLines
'    Close #ff
    
    If OpenW(sHostsFileICS, FOR_READ, hFile) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
        ToggleWow64FSRedirection True
    Else
        sHit = "O1 - Unable to read Hosts.ICS file"
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
        ToggleWow64FSRedirection True
        If NonDefaultPath Then
            GoTo CheckHostsICS_Default:
        Else
            Exit Sub
        End If
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    For Each sLine In Split(sLines, vbLf)
        sLine = Replace$(sLine, vbTab, " ")
        sLine = Replace$(sLine, vbCr, "")
        sLine = Trim$(sLine)
        
        If sLine <> vbNullString Then
            If Left$(sLine, 1) <> "#" Then
                Do
                    sLine = Replace$(sLine, "  ", " ")
                Loop Until InStr(sLine, "  ") = 0
                    
                sHit = "O1 - Hosts.ICS: " & sLine
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
            End If
        End If
    Next
    
CheckHostsICS_Default:

    ToggleWow64FSRedirection True

    If Not NonDefaultPath Then Exit Sub
    
    ff = FreeFile()
    
    cFileSize = FileLenW(sHostsFileICS_Default)
    
    ' Size = 0 or just not exists
    If cFileSize = 0 Then Exit Sub
    
'    On Error Resume Next
'    Open sHostsFileICS_Default For Binary Access Read As #ff
'    If err.Number <> 0 Then
'        sHit = "O1 - Unable to read Hosts.ICS default file"
'        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
'        Exit Sub
'    End If
'
'    On Error GoTo ErrorHandler:
'    sLines = String$(cFileSize, 0)
'    Get #ff, , sLines
'    Close #ff
    
    If OpenW(sHostsFileICS_Default, FOR_READ, hFile) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
    Else
        sHit = "O1 - Unable to read Hosts.ICS default file"
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
        Exit Sub
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    For Each sLine In Split(sLines, vbLf)
        sLine = Replace$(sLine, vbTab, " ")
        sLine = Replace$(sLine, vbCr, "")
        sLine = Trim$(sLine)
        
        If sLine <> vbNullString Then
            If Left$(sLine, 1) <> "#" Then
                Do
                    sLine = Replace$(sLine, "  ", " ")
                Loop Until InStr(sLine, "  ") = 0
                
                sHit = "O1 - Hosts.ICS default: " & sLine
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
            End If
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO1Item"
    Close #ff
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub


Private Sub CheckO1Item()
    On Error GoTo ErrorHandler:
    Dim sFile$, sHit$, sDomains$(), i&, ff%, HostsDefaultFile$, NonDefaultPath As Boolean, bResetOptAdded As Boolean
    Dim iAttr&, sLine As Variant, sLines$, cFileSize@
    Dim aHits() As String, J As Long, hFile As Long
    ReDim aHits(0)
    
    '// TODO: Add UTF8.
    'http://serverfault.com/questions/452268/hosts-file-ignored-how-to-troubleshoot
    
    GetHosts
    
    If bIsWin9x Then HostsDefaultFile = sWinDir & "\hosts"
    If bIsWinNT Then HostsDefaultFile = sWinDir & "\System32\drivers\etc\hosts"
    
    If StrComp(sHostsFile, HostsDefaultFile) <> 0 Then
        'sHit = "O1 - Hosts file is located at: " & sHostsFile
        sHit = "O1 - " & Translate(271) & ": " & sHostsFile
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
        NonDefaultPath = True
    End If
    
    'If Not FileExists(sHostsFile) Then Exit Sub
    
'    On Error Resume Next
'    iAttr = GetFileAttributes(StrPtr(sHostsFile))
'    If (iAttr And 2048) Then iAttr = iAttr - 2048
'
'    SetFileAttributes StrPtr(sHostsFile), vbNormal
'    SetFileAttributes StrPtr(sHostsFile), vbArchive
'
'    If Err.Number And Not inIDE And Not bAutoLogSilent Then  ' tired to see this warning from IDE
'        MsgBoxW replace$(Translate(300), "[]", sHostsFile), vbExclamation
''        msgboxW "For some reason your system denied write " & _
''        "access to the Hosts file." & vbCrLf & "If any hijacked domains " & _
''        "are in this file, HiJackThis may NOT be able to " & _
''        "fix this." & vbCrLf & vbCrLf & "If that happens, you need " & _
''        "to edit the file yourself. To do this, click " & _
''        "Start, Run and type:" & vbCrLf & vbCrLf & _
''        "   notepad """ & sHostsFile & """" & vbCrLf & vbCrLf & _
''        "and press Enter. Find the line(s) HiJackThis " & _
''        "reports and delete them. Save the file as " & _
''        """hosts."" (with quotes), and reboot.", vbExclamation
'    End If
'    SetFileAttributes StrPtr(sHostsFile), iAttr
    
    'ff = FreeFile()
    
    If NonDefaultPath Then                              'Note: \System32\drivers\etc is not under Wow6432 redirection
        ToggleWow64FSRedirection False, sHostsFile
    End If
    
    cFileSize = FileLenW(sHostsFile)
    
    If cFileSize = 0 Then
        If NonDefaultPath Then
            'Check default path also
            GoTo CheckHostsDefault:
        Else
            sHit = "O1 - Hosts: Reset contents to default"
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
            ToggleWow64FSRedirection True
            Exit Sub
        End If
    End If
    
'    On Error Resume Next
'    'Open sHostsFile For Input As #ff
'    Open sHostsFile For Binary Access Read As #ff
'    If err.Number <> 0 Then
'        sHit = "O1 - Unable to read Hosts file"
'        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
'        ToggleWow64FSRedirection True
'        If NonDefaultPath Then
'            GoTo CheckHostsDefault:
'        Else
'            Exit Sub
'        End If
'    End If
    
'    On Error GoTo ErrorHandler:
'    sLines = String$(cFileSize, 0)
'    Get #ff, , sLines
'    Close #ff
    
    If OpenW(sHostsFile, FOR_READ, hFile) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
        ToggleWow64FSRedirection True
    Else
        sHit = "O1 - Unable to read Hosts file"
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
        ToggleWow64FSRedirection True
        If NonDefaultPath Then
            GoTo CheckHostsDefault:
        Else
            Exit Sub
        End If
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    i = 0
    
    For Each sLine In Split(sLines, vbLf)
    
'        Do Until EOF(ff)
'            Line Input #ff, sLine
'            If InStr(sLine, Chr$(10)) > 0 Then
'                'hosts file has line delimiters
'                'that confuse Line Input - so
'                'convert them to vbCrLf :)
'                Close #ff
'                If Not bTriedFixUnixHostsFile Then
'                    FixUNIXHostsFile
'                    bTriedFixUnixHostsFile = True
'                    CheckO1Item
'                Else
'                    If Not bAutoLogSilent Then
'                        MsgBoxW Translate(301), vbExclamation
''                       msgboxW "Your hosts file has invalid linebreaks and " & _
''                           "HiJackThis is unable to fix this. O1 items will " & _
''                           "not be displayed." & vbCrLf & vbCrLf & _
''                           "Click OK to continue the rest of the scan.", vbExclamation
'                    End If
'                End If
'                ToggleWow64FSRedirection True
'                Exit Sub
'            End If
            
            'ignore all lines that start with loopback
            '(127.0.0.1), null (0.0.0.0) and private IPs
            '(192.168. / 10.)
            sLine = Replace$(sLine, vbTab, " ")
            sLine = Replace$(sLine, vbCr, "")
            sLine = Trim$(sLine)
            
            If sLine <> vbNullString Then
                'If InStr(sLine, "127.0.0.1") <> 1 And _
                '   InStr(sLine, "0.0.0.0") <> 1 And _
                '   InStr(sLine, "192.168.") <> 1 And _
                '   InStr(sLine, "10.") <> 1 And _
                '   InStr(sLine, "#") <> 1 And _
                '   Not (bIgnoreSafe And InStr(sLine, "216.239.37.101") > 0) Or _
                '   bIgnoreAllWhitelists Then
                    '216.239.37.101 = google.com
                    
                '::1 - default for Vista
                If Left$(sLine, 1) <> "#" And _
                  StrComp(sLine, "127.0.0.1       localhost", 1) <> 0 And _
                  StrComp(sLine, "::1             localhost", 1) <> 0 And _
                  StrComp(sLine, "127.0.0.1 localhost", 1) <> 0 Then
                  
                    Do
                        sLine = Replace$(sLine, "  ", " ")
                    Loop Until InStr(sLine, "  ") = 0
                    
                    sHit = "O1 - Hosts: " & sLine
                    If Not IsOnIgnoreList(sHit) Then
                        'AddToScanResultsSimple "O1", sHit
                        If UBound(aHits) < i Then ReDim Preserve aHits(UBound(aHits) + 100)
                        aHits(i) = sHit
                        i = i + 1
                    End If
                    
'                    If i = 10 And Not NonDefaultPath And Not bResetOptAdded Then
'                        sHit = "O1 - Hosts: Reset contents to default"
'                        If Not IsOnIgnoreList(sHit) Then
'                            frmMain.lstResults.AddItem sHit, frmMain.lstResults.ListCount - 10
'                            AddToScanResultsSimple "O1", sHit, DoNotAddToListBox:=True
'                        End If
'                        bResetOptAdded = True
'                    End If
                    
                    'I don't plan to fix Hosts file on hijacked location for now.
                    
'                    If i > 100 Then
'                        If Not bAutoLogSilent Then
'                            MsgBoxW replace$(Translate(302), "[]", sHostsFile), vbExclamation
''                           msgboxW "You have an particularly large " & _
''                            "amount of hijacked domains. It's probably " & _
''                            "better to delete the file itself then to " & _
''                            "fix each item (and create a backup)." & vbCrLf & _
''                            vbCrLf & "If you see the same IP address in all " & _
''                            "the reported O1 items, consider deleting your " & _
''                            "Hosts file, which is located at " & sHostsFile & _
''                           ".", vbExclamation
'                        End If
'                        'Close #ff
'                        ToggleWow64FSRedirection True
'                        Exit For
'                    End If
                End If
            End If
        'Loop
    Next
    'Close #ff

    If i > 0 Then
        If i >= 10 Then
            If Not NonDefaultPath Then
                sHit = "O1 - Hosts: Reset contents to default"
                If Not IsOnIgnoreList(sHit) Then
                    AddToScanResultsSimple "O1", sHit
                End If
            End If
        End If
'        'maximum 100 hosts entries
'        If i <= 100 Then
'            For j = 0 To i - 1
'                AddToScanResultsSimple "O1", aHits(j)
'            Next
'        Else
'            sHit = "O1 - Hosts: has " & i & " entries"
'        End If
        For J = 0 To i - 1
            AddToScanResultsSimple "O1", aHits(J)
        Next
    End If
    
    ReDim aHits(0)

CheckHostsDefault:
    'if Hosts was redirected -> checking records on default hosts also. ( Prefix "O1 - Hosts default: " )
    
    i = 0
    
    ToggleWow64FSRedirection True
    
    If NonDefaultPath Then
        
        If FileExists(HostsDefaultFile) Then
            
            cFileSize = FileLenW(HostsDefaultFile)
            If cFileSize <> 0 Then
            
'                ff = FreeFile()
                
'                On Error Resume Next
'                Open HostsDefaultFile For Binary Access Read As #ff
'                If err.Number <> 0 Then
'                    sHit = "O1 - Unable to read Default Hosts file"
'                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
'                End If
'                On Error GoTo ErrorHandler:
        
'                sLines = String$(cFileSize, 0)
'                Get #ff, , sLines
'                Close #ff

                If OpenW(HostsDefaultFile, FOR_READ, hFile) Then
                    sLines = String$(cFileSize, vbNullChar)
                    GetW hFile, 1, sLines
                    CloseW hFile
                Else
                    sHit = "O1 - Unable to read Default Hosts file"
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
                    Exit Sub
                End If
                
                sLines = Replace$(sLines, vbCrLf, vbLf)

                For Each sLine In Split(sLines, vbLf)
                
                    sLine = Replace$(sLine, vbTab, " ")
                    sLine = Replace$(sLine, vbCr, "")
                    sLine = Trim$(sLine)
                    
                    If sLine <> vbNullString Then
                    
                        If Left$(sLine, 1) <> "#" And _
                          StrComp(sLine, "127.0.0.1       localhost", 1) <> 0 And _
                          StrComp(sLine, "::1             localhost", 1) <> 0 Then    '::1 - default for Vista
                            Do
                                sLine = Replace$(sLine, "  ", " ")
                            Loop Until InStr(sLine, "  ") = 0
                    
                            sHit = "O1 - Hosts default: " & sLine
                            If Not IsOnIgnoreList(sHit) Then
                                'AddToScanResultsSimple "O1", sHit
                                If UBound(aHits) < i Then ReDim Preserve aHits(UBound(aHits) + 100)
                                aHits(i) = sHit
                                i = i + 1
                            End If
                    
'                            If i = 10 And Not bResetOptAdded Then
'                                sHit = "O1 - Hosts default: Reset contents to default"
'                                If Not IsOnIgnoreList(sHit) Then
'                                    frmMain.lstResults.AddItem sHit, frmMain.lstResults.ListCount - 10
'                                    AddToScanResultsSimple "O1", sHit, DoNotAddToListBox:=True
'                                End If
'                                bResetOptAdded = True
'                            End If
'
'                            If i > 100 Then
'                                If Not bAutoLogSilent Then
'                                    If vbYes = MsgBoxW(replace$(Translate(302), "[]", sHostsFile), vbExclamation Or vbYesNo) Then
'                                        Shell "explorer.exe /select," & """" & sHostsFile & """", vbNormalFocus
'                                    End If
'        '                           msgboxW "You have an particularly large " & _
'        '                            "amount of hijacked domains. It's probably " & _
'        '                            "better to delete the file itself then to " & _
'        '                            "fix each item (and create a backup)." & vbCrLf & _
'        '                            vbCrLf & "If you see the same IP address in all " & _
'        '                            "the reported O1 items, consider deleting your " & _
'        '                            "Hosts file, which is located at " & sHostsFile & _
'        '                           "." & vbcrlf & vbcrlf & "Would you like to open its folder now?", vbExclamation or vbyesno
'                                End If
'                                Exit Sub
'                            End If
                        End If
                    End If
                Next
            End If
        End If
        
        If i > 0 Then
            If i >= 10 Then
                sHit = "O1 - Hosts default: Reset contents to default"
                If Not IsOnIgnoreList(sHit) Then
                    AddToScanResultsSimple "O1", sHit
                End If
            End If
'            'maximum 100 hosts entries
'            If i <= 100 Then
'                 For j = 0 To i - 1
'                    AddToScanResultsSimple "O1", aHits(j)
'                 Next
'            Else
'                sHit = "O1 - Hosts default: has " & i & " entries"
'            End If
            For J = 0 To i - 1
                AddToScanResultsSimple "O1", aHits(J)
            Next
        End If
    End If

    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO1Item"
    Close #ff
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub

Public Function CheckAccessWrite(Path As String, Optional bDeleteFile As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hFile As Long
    
    hFile = CreateFile(StrPtr(Path), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, CREATE_NEW, ByVal 0&, ByVal 0&)
    
    If hFile <> 0 Then
        CloseHandle hFile
        CheckAccessWrite = True
    End If
    
    If bDeleteFile Then
        If FileExists(Path) Then
            DeleteFileWEx StrPtr(Path)
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "modMain_CheckAccessWrite"
    If inIDE Then Stop: Resume Next
End Function

Public Sub FixO1Item(sItem$)
    'O1 - Hijack of auto.search.msn.com etc with Hosts file
    On Error GoTo ErrorHandler:
    Dim sLine As Variant, sHijacker$, i&, iAttr&, ff1%, ff2%, HostsDefaultPath$, sLines$, HostsDefaultFile$, cFileSize@, sHosts$
    Dim sHostsTemp$, bResetHosts As Boolean, aLines() As String, isICS As Boolean, SFC As String
    
    'If StrBeginWith(sItem, "O1 - Hosts default: has ") Or _
    '    StrBeginWith(sItem, "O1 - Hosts: has ") Then Exit Sub
    
    If InStr(1, sItem, "O1 - DNSApi:", 1) <> 0 Then
        sHijacker = Mid$(sItem, InStr(sItem, ":") + 2)
        sHijacker = Mid$(sHijacker, InStr(sHijacker, " - ") + 3)
        If OSver.Bitness = "x64" And FolderExists(sWinDir & "\sysnative") Then 'Vista+
            SFC = EnvironW("%SystemRoot%") & "\sysnative\sfc.exe"
        Else
            SFC = EnvironW("%SystemRoot%") & "\System32\sfc.exe"
        End If
        If FileExists(SFC) Then
            'TryUnlock sHijacker
            Proc.ProcessRun SFC, "/SCANFILE=" & """" & sHijacker & """", , 0
            If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 15000) Then
                Proc.ProcessClose , , True
            End If
        End If
        
        FlushDNS
        
        Exit Sub
    End If
    
    If bIsWin9x Then HostsDefaultPath = sWinDir
    If bIsWinNT Then HostsDefaultPath = "%SystemRoot%\System32\drivers\etc"
    
    HostsDefaultFile = EnvironW(HostsDefaultPath & "\" & "hosts")
    
    'If InStr(sItem, "Hosts file is located at") > 0 Then
    If InStr(sItem, Translate(271)) > 0 Then
        'hosts file relocation - always bad
        RegSetExpandStringVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", HostsDefaultPath
        GetHosts    'reload var. 'sHostsFile'
        
        FlushDNS
        
        Exit Sub
    End If
    
    If StrComp(sItem, "O1 - Hosts: Reset contents to default", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts default: Reset contents to default", 1) = 0 Then
        bResetHosts = True
    End If
    
    If StrBeginWith(sItem, "O1 - Hosts default: ") Or bResetHosts Then
        
        sHosts = HostsDefaultFile   'default hosts path
    Else
        sHosts = sHostsFile         'path that may be redirected
    End If
    
    If StrBeginWith(sItem, "O1 - Hosts.ICS: ") Then
        sHosts = sHostsFile & ".ics"
        isICS = True
    ElseIf StrBeginWith(sItem, "O1 - Hosts.ICS default: ") Then
        sHosts = HostsDefaultFile & ".ics"
        isICS = True
    End If
    
    sHostsTemp = TempCU & "\" & "hosts.new"
    If Not CheckAccessWrite(sHostsTemp, True) Then
        sHostsTemp = BuildPath(AppPath(), "hosts.new")
    End If
    
    If FileExists(sHostsTemp) Then
        DeleteFileWEx StrPtr(sHostsTemp)
    End If
    
    ToggleWow64FSRedirection False
    
    cFileSize = FileLenW(sHosts)
    
    ToggleWow64FSRedirection False
    
    If cFileSize = 0 Or bResetHosts Then
        'no reset for ICS for now
        If isICS Then Exit Sub
        '2.0.7. - Reset Hosts to its default contents
        ff2 = FreeFile()
        Open sHostsTemp For Output As #ff2
            Print #ff2, GetDefaultHostsContents()
        Close #ff2
        GoTo Replace
    End If
    
    'If Not StrBeginWith(sItem, "O1 - Hosts: ") Then Exit Sub
    
    'parse to server name
    ' Example: 127.0.0.1 my.dragokas.com -> var. 'sHijacker' = "my.dragokas.com"
    sHijacker = Mid$(sItem, InStr(sItem, ":") + 2)
    sHijacker = Trim$(sHijacker)
    If Not isICS Then
        If InStr(sHijacker, " ") > 0 Then
            Dim sTemp$
            sTemp = Mid$(sHijacker, InStr(sHijacker, " ") + 1)
            If 0 <> Len(sTemp) Then sHijacker = sTemp
        End If
    End If
    
    ToggleWow64FSRedirection False
    
    'Reset attributes (and save old one in var. 'iAttr')
    iAttr = GetFileAttributes(StrPtr(sHosts))
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetFileAttributes StrPtr(sHosts), vbNormal
    
    'read current hosts file
    ff1 = FreeFile()
    Open sHosts For Binary Access Read As #ff1
    sLines = String$(LOF(ff1), 0)
    Get #ff1, , sLines
    Close #ff1
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    'build new hosts file (exclude bad lines)
    ff2 = FreeFile()
    Open sHostsTemp For Output As #ff2
        aLines = Split(sLines, vbLf)
        If SafeArrayGetDim(aLines) Then
          For i = 0 To UBound(aLines)
            sLine = aLines(i)
            sLine = Replace$(sLine, vbTab, " ")
            sLine = Replace$(sLine, vbCr, "")
            Do
                sLine = Replace$(sLine, "  ", " ")
            Loop Until InStr(sLine, "  ") = 0
            If InStr(1, sLine, sHijacker, 1) <> 0 Then
                'don't write line to hosts file
            Else
                'skip last empty line
                If 0 <> Len(sLine) Or (0 = Len(sLine) And i < UBound(aLines)) Then Print #ff2, aLines(i)
            End If
          Next
        End If
    Close #ff2
    
Replace:
    If DeleteFileForce(sHosts) Then
        ToggleWow64FSRedirection False
        If 0 = MoveFile(StrPtr(sHostsTemp), StrPtr(sHosts)) Then
            If err.LastDllError = 5 Then err.Raise 70
        End If
        'Recover old one attrib.
        SetFileAttributes StrPtr(sHosts), iAttr
    Else
        err.Raise 70
    End If
    
    ToggleWow64FSRedirection True
    
    FlushDNS
    
    '//TODO:
    'clear cache
    
    '1. Mozilla Firefox
    '%LocalAppData%\Mozilla\Firefox\Profiles\<Name>\cache2 -> rename to *.bak
    
    '2. Microsoft Internet Explorer
    
    '3. Google Chrome
    
    '4. Yandex Browser
    
    '5.1. Opera Presto
    
    '5.2. (Chromo) Opera
    
    '6. Edge
    '...
    Exit Sub
    
ErrorHandler:
    If err.Number = 70 And Not bSeenHostsFileAccessDeniedWarning Then
        'permission denied
        MsgBoxW Translate(303), vbExclamation
'        msgboxW "HiJackThis could not write the selected changes to your " & _
'               "hosts file. The probably cause is that some program is " & _
'               "denying access to it, or that your user account doesn't have " & _
'               "the rights to write to it.", vbExclamation
        bSeenHostsFileAccessDeniedWarning = True
    Else
        ErrorMsg err, "modMain_FixO1Item", "sItem=", sItem
    End If
    Close #ff1, #ff2
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub

Sub FlushDNS()
    If GetServiceRunState("dnscache") = SERVICE_RUNNING Then
        Shell "ipconfig.exe /flushdns", vbHide
    End If
End Sub


Public Sub CheckO2Item()
  On Error GoTo ErrorHandler:
  Dim hKey&, i&, J&, sName$, sCLSID$, lpcName&, sFile$, sHit$, BHO_key$, sFileExisted$, Result As TYPE_Scan_Results
  Dim Wow6432Redir As Boolean, UseWow, sBuf$
    
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
  
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr("Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects"), 0&, _
        KEY_ENUMERATE_SUB_KEYS Or (KEY_WOW64_64KEY And Not Wow6432Redir), hKey) <> 0 Then Exit Sub
        
    i = 0
    Do
        sCLSID = String$(MAX_KEYNAME, vbNullChar)
        lpcName = Len(sCLSID)
        If RegEnumKeyExW(hKey, i, StrPtr(sCLSID), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
        
        sCLSID = Left$(sCLSID, lstrlen(StrPtr(sCLSID)))
        
        If sCLSID <> vbNullString And Not StrBeginWith(sCLSID, "MSHist") Then
            
            BHO_key = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID
            
            If InStr(sCLSID, "}}") > 0 Then
                'the new searchwww.com trick - use a double
                '}} in the IE toolbar registration, reg the toolbar
                'with only one } - IE ignores the double }}, but
                'HT didn't. It does now!
                sCLSID = Left$(sCLSID, Len(sCLSID) - 1)
            End If
            
            'get filename from HKCR\CLSID\sName
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Wow6432Redir)
            
            sFileExisted = ""
            
            If InStr(sFile, "__BHODemonDisabled") > 0 Then
                sFile = Left$(sFile, InStr(sFile, "__BHODemonDisabled") - 1) & " (disabled by BHODemon)"
            Else
                If 0 <> Len(sFile) Then
                    If FileExists(sFile) Then
                        sFileExisted = sFile
                    Else
                        sFile = sFile & " (file missing)"
                    End If
                End If
            End If
            If 0 = Len(sFile) Then sFile = "(no file)"
            
            'get bho name from BHO regkey
            sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID, vbNullString, Wow6432Redir)
            If sName = vbNullString Then
                'get BHO name from CLSID regkey
                sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, True)
                If sName = vbNullString Then sName = "(no name)"
            End If
            
            If Left$(sName, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sName)
                If 0 <> Len(sBuf) Then sName = sBuf
            End If
            
            sHit = IIf(bIsWin32, "O2", IIf(Wow6432Redir, "O2-32", "O2")) & " - BHO: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                
                With Result
                    .Section = "O2"
                    .HitLineW = sHit
                    '.Alias = sAlias
                    ReDim .RegKey(17)
                    .RegKey(0) = BHO_key
                    .RegKey(1) = "HKCR\CLSID\" & sCLSID
                    .RegKey(2) = "HKCU\Software\Microsoft\Internet Explorer\Extension Compatibility\" & sCLSID
                    .RegKey(3) = "HKLM\Software\Microsoft\Internet Explorer\Extension Compatibility\" & sCLSID
                    .RegKey(4) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Ext\Stats\" & sCLSID
                    .RegKey(5) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Ext\Stats\" & sCLSID
                    .RegKey(6) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & sCLSID
                    .RegKey(7) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & sCLSID
                    .RegKey(8) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved\" & sCLSID
                    .RegKey(9) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved\" & sCLSID
                    .RegKey(10) = "HKCU\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration\" & sCLSID
                    .RegKey(11) = "HKLM\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration\" & sCLSID
                    .RegKey(12) = "HKCU\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration" & sCLSID
                    .RegKey(13) = "HKLM\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration" & sCLSID
                    .RegKey(14) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID" 'param
                    .RegKey(15) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID" 'param
                    .RegKey(16) = "HKCU\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration" 'param
                    .RegKey(17) = "HKLM\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration" 'param
                    .RegParam = sCLSID
                    .RunObject = sFileExisted
                    '.CureType = REGISTRY_KEY_BASED
                    .Redirected = Wow6432Redir
                End With
                AddToScanResults Result
            End If
        End If
        i = i + 1
    Loop
    RegCloseKey hKey
  Next
  Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_CheckO2Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO2Item(sItem$)
    'O2 - Enumeration of existing MSIE BHO's
    'O2 - BHO: AcroIEHlprObj Class - {00000...000} - C:\PROGRAM FILES\ADOBE\ACROBAT 5.0\ACROBAT\ACTIVEX\ACROIEHELPER.OCX
    'O2 - BHO: ... (no file)
    'O2 - BHO: ... c:\bla.dll (file missing)
    'O2 - BHO: ... c:\bla.dll (disabled by BHODemon)
    
    On Error GoTo ErrorHandler:
    
    If Not bShownBHOWarning And ProcessExist("iexplore.exe") Then
        MsgBoxW Translate(310), vbExclamation
'        msgboxW "HiJackThis is about to remove a " & _
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
    'On Error GoTo ErrorHandler:
    
    '//TODO: Add:
    'HKLM\SOFTWARE\WOW6432NODE\MICROSOFT\INTERNET EXPLORER\LOW RIGHTS\ELEVATIONPOLICY\{CLSID}
    'HKLM\SOFTWARE\CLASSES\APPID\{Name}
    'HKLM\SOFTWARE\CLASSES\APPID\{GUID}
    'HKLM\SOFTWARE\WOW6432NODE\CLASSES\APPID\{Name}
    'HKLM\SOFTWARE\WOW6432NODE\CLASSES\APPID\{GUID}
    'HKLM\SOFTWARE\CLASSES\INTERFACE\{GUID}
    'HKLM\SOFTWARE\CLASSES\TYPELIB\{GUID}
    
    Dim Result As TYPE_Scan_Results, i As Long
    If Not GetScanResults(sItem, Result) Then Exit Sub
    Dim UseWow
    
    With Result
        For i = 0 To 13
            For Each UseWow In Array(False, True)
                RegDelKey 0&, .RegKey(i), CBool(UseWow)
            Next
        Next

        RegDelVal 0&, .RegKey(14), .RegParam
        RegDelVal 0&, .RegKey(15), .RegParam
        RegDelVal 0&, .RegKey(16), .RegParam
        
        For Each UseWow In Array(False, True)
            RegDelVal 0&, .RegKey(17), .RegParam, CBool(UseWow)
        Next
        If 0 <> Len(.RunObject) Then
            If FileExists(.RunObject) Then DeleteFileWEx StrPtr(.RunObject)
        End If
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO2Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO3Item()
  'HKLM\Software\Microsoft\Internet Explorer\Toolbar
  On Error GoTo ErrorHandler:
    
  Dim hKey&, hKey2&, i&, J&, sCLSID$, sName$, Result As TYPE_Scan_Results
  Dim sFile$, sHit$, SearchwwwTrick As Boolean, sBuf$
  Dim UseWow, Wow6432Redir As Boolean
    
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr("Software\Microsoft\Internet Explorer\Toolbar"), 0, _
      KEY_QUERY_VALUE Or (KEY_WOW64_64KEY And Not Wow6432Redir), hKey) <> 0 Then Exit Sub
    
    i = 0
    Do
        sCLSID = String$(MAX_VALUENAME, 0)
        ReDim uData(MAX_VALUENAME)
        
        'enumerate MSIE toolbars
        If RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&) <> 0 Then Exit Do
        sCLSID = TrimNull(sCLSID)
        
        If InStr(sCLSID, "}}") > 0 Then
            'the new searchwww.com trick - use a double
            '}} in the IE toolbar registration, reg the toolbar
            'with only one } - IE ignores the double }}, but
            'HT didn't. It does now!
            
            sCLSID = Left$(sCLSID, Len(sCLSID) - 1)
            SearchwwwTrick = True
        Else
            SearchwwwTrick = False
        End If
        
        'found one? then check corresponding HKCR key
        sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, Wow6432Redir)
        If 0 = Len(sName) Then
            sName = "(no name)"
        End If
        If Left$(sName, 1) = "@" Then
            sBuf = GetStringFromBinary(, , sName)
            If 0 <> Len(sBuf) Then sName = sBuf
        End If
        
        'If HasSpecialCharacters(sName) Then
            'when japanese characters are in toolbar name,
            'it tends to screw up things
        '    sName = "?????"
        'End If
        
        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Wow6432Redir)
        If 0 = Len(sFile) Then
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
        
        If 0 <> Len(sName) And InStr(sCLSID, "{") > 0 Then
        
          If Not SearchwwwTrick Or _
            (SearchwwwTrick And (sCLSID <> "BrandBitmap" And sCLSID <> "SmBrandBitmap")) Then
        
            sHit = IIf(bIsWin32, "O3", IIf(Wow6432Redir, "O3-32", "O3")) & " - Toolbar: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                With Result
                    .Section = "O3"
                    .HitLineW = sHit
                    ReDim .RegKey(17)
                    .RegKey(0) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility\" & sCLSID
                    .RegKey(1) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Toolbar" 'param
                    
                    .RegKey(4) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Ext\Stats\" & sCLSID
                    .RegKey(5) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Ext\Stats\" & sCLSID
                    .RegKey(6) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & sCLSID
                    .RegKey(7) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\" & sCLSID
                    .RegKey(8) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved\" & sCLSID
                    .RegKey(9) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved\" & sCLSID
                    .RegKey(10) = "HKCU\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration\" & sCLSID
                    .RegKey(11) = "HKLM\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration\" & sCLSID
                    .RegKey(12) = "HKCU\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration" & sCLSID
                    .RegKey(13) = "HKLM\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration" & sCLSID
                    .RegKey(14) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID" 'param
                    .RegKey(15) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID" 'param
                    .RegKey(16) = "HKCU\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration"  'param
                    .RegKey(17) = "HKLM\Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration"  'param
                    
                    .RegParam = sCLSID
                    .Redirected = Wow6432Redir
                End With
                AddToScanResults Result
            End If
          End If
        End If
        i = i + 1
    Loop
    RegCloseKey hKey
  Next

  Exit Sub
  
ErrorHandler:
    ErrorMsg err, "modMain_CheckO3Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO3Item(sItem$)
    'O3 - Enumeration of existing MSIE toolbars
    
    On Error GoTo ErrorHandler:
    Dim Result As TYPE_Scan_Results, i As Long
    If Not GetScanResults(sItem, Result) Then Exit Sub
    Dim UseWow
    
    With Result
        For Each UseWow In Array(False, True)
            RegDelKey 0&, .RegKey(0), CBool(UseWow)
            RegDelVal 0&, .RegKey(1), .RegParam, CBool(UseWow)
        Next
        RegDelVal 0&, .RegKey(14), .RegParam
        RegDelVal 0&, .RegKey(15), .RegParam
        RegDelVal 0&, .RegKey(16), .RegParam
        
        For Each UseWow In Array(False, True)
            RegDelVal 0&, .RegKey(17), .RegParam, CBool(UseWow)
        
            For i = 4 To 13
                RegDelKey 0&, .RegKey(i), CBool(UseWow)
            Next
        Next
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO3Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub


'returns array of SID strings, except of current user
Sub GetUserNamesAndSids(aSID() As String, aUser() As String)
    On Error GoTo ErrorHandler:
    
    'get all users' SID and map it to the corresponding username
    'not all users visible in User Accounts screen have a SID though,
    'they probably get this when logging in for the first time

    Dim CurUserName$, i&, k&, sUsername$, aTmpSID() As String, aTmpUser() As String

    CurUserName = GetUser()
    
    aTmpSID = SplitSafe(RegEnumSubKeys(HKEY_USERS, vbNullString), "|")
    ReDim aTmpUser(UBound(aTmpSID))
    For i = 0 To UBound(aTmpSID)
        If Left$(aTmpSID(i), 1) = "S" And InStr(aTmpSID(i), "_Classes") = 0 Then
            sUsername = MapSIDToUsername(aTmpSID(i))
            If 0 = Len(sUsername) Then sUsername = "?"
            If StrComp(sUsername, CurUserName, 1) <> 0 Then
                aTmpUser(i) = sUsername
            Else
                'filter current user key with HKCU
                aTmpSID(i) = ""
                aTmpUser(i) = ""
            End If
        Else
            aTmpSID(i) = ""
            aTmpUser(i) = ""
        End If
    Next i
    
    'compress array
    k = 0
    ReDim aSID(UBound(aTmpSID))
    ReDim aUser(UBound(aTmpSID))
    
    For i = 0 To UBound(aTmpSID)
        If 0 <> Len(aTmpSID(i)) Then
            aSID(k) = aTmpSID(i)
            aUser(k) = aTmpUser(i)
            k = k + 1
        End If
    Next
    If k > 0 Then
        ReDim Preserve aSID(k - 1)
        ReDim Preserve aUser(k - 1)
    Else
        ReDim Preserve aSID(0)
        ReDim Preserve aUser(0)
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain.GetUserNamesAndSids"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_RegRuns(aHives() As String, aUser() As String)
    On Error GoTo ErrorHandler:
    
    '// TODO: add RunOnceEx, RunServicesOnceEx
    'https://support.microsoft.com/en-us/kb/310593
    'http://www.oszone.net/2762
    '" DllFileName | FunctionName | CommandLineArguements "
    
    Dim sRegRuns() As String, sDes() As String, Wow6432Redir As Boolean, UseWow, Result As TYPE_Scan_Results
    Dim sHive$, i&, k&, sKey$, sData$, sHit$, sAlias$, Param As Variant, sMD5$
    Dim bData() As Byte, isDisabledWin8 As Boolean, isDisabledWinXP As Boolean, flagDisabled As Long, sKeyDisable As String
    
    ReDim sRegRuns(1 To 10)
    ReDim sDes(UBound(sRegRuns))
    
    sRegRuns(1) = "Software\Microsoft\Windows\CurrentVersion\Run"
    sDes(1) = "Run"
    
    sRegRuns(2) = "Software\Microsoft\Windows\CurrentVersion\RunServices"
    sDes(2) = "RunServices"
    
    sRegRuns(3) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
    sDes(3) = "RunOnce"
    
    sRegRuns(4) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    sDes(4) = "RunServicesOnce"
    
    sRegRuns(5) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
    sDes(5) = "Policies\Explorer\Run"
    
    sRegRuns(6) = "Software\Microsoft\Windows\CurrentVersion\Run-"
    sDes(6) = "Run-"

    sRegRuns(7) = "Software\Microsoft\Windows\CurrentVersion\RunServices-"
    sDes(7) = "RunServices-"

    sRegRuns(8) = "Software\Microsoft\Windows\CurrentVersion\RunOnce-"
    sDes(8) = "RunOnce-"

    sRegRuns(9) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce-"
    sDes(9) = "RunServicesOnce-"
    
    For i = 0 To UBound(aHives) 'HKLM, HKCU, HKU()

        sHive = aHives(i)

        For Each UseWow In Array(False, True)
    
            Wow6432Redir = UseWow
  
            If (bIsWin32 And Wow6432Redir) _
              Or bIsWin64 And Wow6432Redir And _
              (sHive = "HKCU" _
              Or StrBeginWith(sHive, "HKU\")) Then Exit For
  
            For k = LBound(sRegRuns) To UBound(sRegRuns)
            
                If sHive = "HKCU" And StrBeginWith(sRegRuns(k), "SYSTEM\") Then GoTo Continue:  'skip hkcu\system
      
                sKey = sHive & "\" & sRegRuns(k)
        
                For Each Param In Split(GetEnumValues(0&, sKey, Wow6432Redir), "|")
        
                    isDisabledWin8 = False
                    
                    isDisabledWinXP = (Right$(sRegRuns(k), 1) = "-")    ' Run- e.t.c.
                    
                    sData = GetRegData(0&, sKey, CStr(Param), Wow6432Redir)
            
                    If OSver.MajorMinor >= 6.2 Then  ' Win 8+
                      
                      If StrComp(sRegRuns(k), "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 1) = 0 Then
                    
                        'Param. name is always "Run" on x32 bit. OS.
                        sKeyDisable = sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                            IIf(bIsWin64 And Wow6432Redir, "Run32", "Run")
                    
                        If RegValueExists(0&, sKeyDisable, CStr(Param)) Then
                            
                            ReDim bData(0)
                            bData() = RegGetBinary(0&, sKeyDisable, CStr(Param))
            
                            If UBound(bData) >= 11 Then
                                
                                GetMem4 ByVal VarPtr(bData(0)), flagDisabled
                    
                                If flagDisabled <> 2 Then isDisabledWin8 = True
                            End If
                        End If
                      End If
                    End If
            
                    If Len(sData) <> 0 And Not isDisabledWin8 Then
                
                        'Example:
                        '"O4 - HKLM\..\Run: "
                        '"O4 - HKU\S-1-5-19\..\Run: "
                        sAlias = IIf(bIsWin32, "O4", IIf(Wow6432Redir, "O4-32", "O4")) & _
                          " - " & IIf(isDisabledWinXP, "(disabled) ", "") & sHive & "\..\" & sDes(k) & ": "
                
                        sHit = sAlias & "[" & Param & "] " & sData
                
                        If Not IsOnIgnoreList(sHit) Then
                
                            If bMD5 Then sMD5 = GetFileFromAutostart(sData): sHit = sHit & sMD5
                    
                            With Result
                                .Section = "O4"
                                .HitLineW = sHit
                                .Alias = sAlias
                                ReDim .RegKey(0)
                                .RegKey(0) = sKey
                                .RegParam = Param
                                .RunObject = sData
                                .CureType = REGISTRY_PARAM_BASED
                                .Redirected = Wow6432Redir
                            End With
                            AddToScanResults Result
                        End If
                    End If
                Next Param
Continue:
            Next k
        Next UseWow
    Next i
    
    'Certain param based checkings
    
    ReDim aRegKey(1 To 2) As String         'key
    ReDim aRegParam(1 To UBound(aRegKey))   'param
    ReDim aDefData(1 To UBound(aRegKey))    'data
    ReDim sDes(1 To UBound(aRegKey))        'description
   
    aRegKey(1) = "Software\Microsoft\Command Processor"
    aRegParam(1) = "Autorun"
    aDefData(1) = ""
    sDes(1) = "Command Processor\Autorun"
    
    aRegKey(2) = "SYSTEM\CurrentControlSet\Control\BootVerificationProgram"
    aRegParam(2) = "ImagePath"
    aDefData(2) = ""
    sDes(2) = "BootVerificationProgram"
    
    For i = 0 To UBound(aHives) 'HKLM, HKCU, HKU()

        sHive = aHives(i)

        For Each UseWow In Array(False, True)
    
            Wow6432Redir = UseWow
  
            If (bIsWin32 And Wow6432Redir) _
              Or bIsWin64 And Wow6432Redir And _
              (sHive = "HKCU" _
              Or StrBeginWith(sHive, "HKU\")) Then Exit For
  
            For k = LBound(aRegKey) To UBound(aRegKey)
    
                If (sHive = "HKCU" Or StrBeginWith(sHive, "HKU\")) And StrBeginWith(aRegKey(k), "SYSTEM\") Then GoTo Continue2:  'skip hkcu\system
                If StrBeginWith(aRegKey(k), "SYSTEM\") And Wow6432Redir Then GoTo Continue2:
                
                sKey = sHive & "\" & aRegKey(k)
                Param = aRegParam(k)
                
                'Debug.Print IIf(Wow6432Redir, replace$(sKey, "Software\", "Software\Wow6432node\"), sKey) & "," & aRegParam(k)
    
                sData = GetRegData(0&, sKey, aRegParam(k), Wow6432Redir)

                If sData <> aDefData(k) Then

                    'HKLM\..\Command Processor\Autorun:
                    sAlias = IIf(bIsWin32, "O4", IIf(Wow6432Redir, "O4-32", "O4")) & _
                          " - " & sHive & "\..\" & sDes(k) & ": "

                    sHit = sAlias & sData

                    If Not IsOnIgnoreList(sHit) Then

                        If bMD5 Then sMD5 = GetFileFromAutostart(sData): sHit = sHit & sMD5

                        With Result
                            .Section = "O4"
                            .HitLineW = sHit
                            .Alias = sAlias
                            ReDim .RegKey(0)
                            .RegKey(0) = sKey
                            .RegParam = Param
                            .RunObject = sData
                            .CureType = REGISTRY_PARAM_BASED
                            .Redirected = Wow6432Redir
                        End With
                        AddToScanResults Result
                    End If
                End If
Continue2:
            Next
        Next
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain.CheckO4_RegRuns"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_MSConfig(aHives() As String, aUser() As String)
    On Error GoTo ErrorHandler:
    
    'HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupreg
    'HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupfolder
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder -> checked in CheckO4_AutostartFolder()
    
    Dim sHive$, i&, J&, sAlias$, sMD5$, Result As TYPE_Scan_Results
    Dim SubKey As Variant, sDay$, sMonth$, sYear$, sKey$, sFile$, stime$, sHit$, SourceHive$
    Dim Values$(), bData() As Byte, flagDisabled As Long, dDate As Date, UseWow As Variant, Wow6432Redir As Boolean
    
    If OSver.MajorMinor >= 6.2 Then ' Win 8+
    
        For i = 0 To UBound(aHives) 'HKLM, HKCU, HKU()

            sHive = aHives(i)
            
            For Each UseWow In Array(False, True)
    
                Wow6432Redir = UseWow
  
                If (bIsWin32 And Wow6432Redir) _
                  Or bIsWin64 And Wow6432Redir And (sHive = "HKCU" Or StrBeginWith(sHive, "HKU\")) Then
                    Exit For
                End If
            
                For J = 1 To GetEnumValuesToArray(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values())
            
                    flagDisabled = 2
                    ReDim bData(0)
                    
                    bData() = RegGetBinary(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values(J))
            
                    If UBound(bData) >= 11 Then
                        GetMem4 ByVal VarPtr(bData(0)), flagDisabled
                    End If
                    
                    If flagDisabled <> 2 Then   'is Disabled ?
                    
                        dDate = ConvertFileTimeToLocalDate(VarPtr(bData(4)))
                        
                        If RegValueExists(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(J), Wow6432Redir) Then
                        
                            sFile = RegGetString(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(J), Wow6432Redir)
                        
                            'if you change it, change fix appropriate !!!
                            sAlias = "O4 - " & sHive & "\..\StartupApproved\" & IIf(bIsWin64 And Wow6432Redir, "Run32", "Run") & ": "
            
                            sHit = sAlias & "[" & Values(J) & "] " & sFile & " (" & Format$(dDate, "yyyy\/mm\/dd") & ")" & _
                                IIf(sFile = "", " (no file)", "") ', IIf(Not FileExists(sFile), " (file missing)", ""))
                        
                            If Not IsOnIgnoreList(sHit) Then
                            
                                If bMD5 Then sMD5 = GetFileFromAutostart(sFile): sHit = sHit & sMD5
                
                                With Result
                                    .Section = "O4"
                                    .HitLineW = sHit
                                    .Alias = sAlias
                                    ReDim .RegKey(1)
                                    .RegKey(0) = sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run")
                                    .RegKey(1) = sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                                    .RegParam = Values(J)
                                    .RunObject = sFile
                                    .Redirected = Wow6432Redir
                                    .CureType = REGISTRY_PARAM_BASED
                                End With
                                AddToScanResults Result
                            End If
                        End If
                    End If
                Next
            Next
        Next
        
    Else
    
        sHive = "HKLM"
        sKey = sHive & "\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupreg"
        
        For Each SubKey In Split(RegEnumSubKeys(0&, sKey), "|")
        
            sFile = GetRegData(0&, sKey & "\" & SubKey, "command")
            
            sYear = GetRegData(0&, sKey & "\" & SubKey, "YEAR")
            sMonth = Right$("0" & GetRegData(0&, sKey & "\" & SubKey, "MONTH"), 2)
            sDay = Right$("0" & GetRegData(0&, sKey & "\" & SubKey, "DAY"), 2)
            
            If Val(sYear) = 0 Or Val(sMonth) = 0 Or Val(sDay) = 0 Then
                stime = Format$(GetRegKeyTime(0&, sKey & "\" & SubKey), "yyyy\/mm\/dd")
            Else
                stime = sYear & "/" & sMonth & "/" & sDay
            End If
            
            'SourceHive = GetRegData(0&, sKey & "\" & SubKey, "hkey")
            'If SourceHive <> "HKLM" And SourceHive <> "HKCU" Then SourceHive = ""
            
            '"O4 - MSConfig: "
            sAlias = "O4 - MSConfig\startupreg: " '& SourceHive & ": "
            
            sHit = sAlias & "[" & SubKey & "] " & sFile & " (" & stime & ")" & IIf(sFile = "", " (no file)", "") ', IIf(Not FileExists(sFile), " (file missing)", ""))
            
            If Not IsOnIgnoreList(sHit) Then
                
                If bMD5 Then sMD5 = GetFileFromAutostart(sFile): sHit = sHit & sMD5
                
                With Result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    ReDim .RegKey(0)
                    .RegKey(0) = sKey & "\" & SubKey
                    .RunObject = sFile
                    .CureType = REGISTRY_KEY_BASED
                End With
                AddToScanResults Result
            End If
        Next
        
        'Startup folder items
        
        sKey = "HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupfolder"
        
        For Each SubKey In Split(RegEnumSubKeys(0&, sKey), "|")
        
            sFile = GetRegData(0&, sKey & "\" & SubKey, "backup")
            
            stime = Format$(GetRegKeyTime(0&, sKey & "\" & SubKey), "yyyy\/mm\/dd")
        
            sAlias = "O4 - MSConfig\startupfolder: "    'if you change it, change fix appropriate !!!
            
            sHit = sAlias & SubKey & " - " & sFile & " (" & stime & ")" & IIf(sFile = "", " (no file)", IIf(Not FileExists(sFile), " (file missing)", ""))
            
            If Not IsOnIgnoreList(sHit) Then
                
                If bMD5 Then sMD5 = GetFileFromAutostart(sFile): sHit = sHit & sMD5
                
                With Result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    ReDim .RegKey(0)
                    .RegKey(0) = sKey & "\" & SubKey
                    .RunObject = sFile
                    .CureType = REGISTRY_KEY_BASED
                End With
                AddToScanResults Result
            End If
        Next
    End If

    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain.CheckO4_MSConfig"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_AutostartFolder(aSID() As String, aUser() As String)
    On Error GoTo ErrorHandler:

    Dim aRegKeys() As String, aParams() As String, aDes() As String, aDesConst() As String, Result As TYPE_Scan_Results
    Dim sAutostartFolder$(), sShortCut$, i&, k&, Wow6432Redir As Boolean, UseWow, sFolder$, sHit$
    Dim FldCnt&, sKey$, sSID$, sFile$, sLinkPath$, sLinkExt$, sTarget$, bLink As Boolean, bPE_EXE As Boolean
    Dim bData() As Byte, isDisabled As Boolean, flagDisabled As Long, sKeyDisable As String, sHive As String, dDate As Date
    Dim StartupCU As String, aFiles() As String, sArguments As String
    
    ReDim aRegKeys(1 To 8)
    ReDim aParams(1 To UBound(aRegKeys))
    ReDim aDesConst(1 To UBound(aRegKeys))
    
    ReDim sAutostartFolder(100) ' HKCU + HKLM + Wow64 + HKU
    ReDim aDes(100)
    
    
    'aRegKeys  - Key
    'aParams   - Value
    'aDesConst - Description for HJT Section
    
    'HKLM (HKLM hives should go first)
    aRegKeys(1) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(1) = "Common Startup"
    aDesConst(1) = "Global Startup"
    
    aRegKeys(2) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(2) = "Common AltStartup"
    aDesConst(2) = "Global AltStartup"
    
    aRegKeys(3) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(3) = "Common Startup"
    aDesConst(3) = "Global User Startup"
    
    aRegKeys(4) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(4) = "Common AltStartup"
    aDesConst(4) = "Global User AltStartup"
    
    'HKCU
    aRegKeys(5) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(5) = "Startup"
    aDesConst(5) = "Startup"
    
    aRegKeys(6) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(6) = "AltStartup"
    aDesConst(6) = "AltStartup"
    
    aRegKeys(7) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(7) = "Startup"
    aDesConst(7) = "User Startup"
    
    aRegKeys(8) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(8) = "AltStartup"
    aDesConst(8) = "User AltStartup"

    
    
    FldCnt = 0
    
    ' Get folder pathes
    For k = 1 To UBound(aRegKeys)
    
        For Each UseWow In Array(False, True)
            
            Wow6432Redir = UseWow
        
            'skip HKCU Wow64
            If (bIsWin32 And Wow6432Redir) _
              Or bIsWin64 And Wow6432Redir And StrBeginWith(aRegKeys(k), "HKCU") Then Exit For
    
            FldCnt = FldCnt + 1
            sAutostartFolder(FldCnt) = RegGetString(0&, aRegKeys(k), aParams(k), Wow6432Redir)
            aDes(FldCnt) = aDesConst(k)
            
            'save path of Startup for current user to substitute other user names
            If aParams(k) = "Startup" Then
                If Len(sAutostartFolder(FldCnt)) <> 0 Then
                    StartupCU = EnvironW(sAutostartFolder(FldCnt))
                End If
            End If
        Next
    Next
    
    '+ HKU pathes
    For i = 0 To UBound(aSID)
        If Len(aSID(i)) <> 0 Then
            sSID = aSID(i)
            'sUserName = aUser(i)
            
            For k = 1 To UBound(aRegKeys)
            
                'only HKCU keys
                If StrBeginWith(aRegKeys(k), "HKCU") Then
                
                    ' Convert HKCU -> HKU
                    sKey = Replace$(aRegKeys(k), "HKCU\", "HKU\" & sSID)
                
                    FldCnt = FldCnt + 1
                    If UBound(sAutostartFolder) < FldCnt Then
                        ReDim Preserve sAutostartFolder(UBound(sAutostartFolder) + 100)
                        ReDim Preserve aDes(UBound(aDes) + 100)
                    End If
            
                    sAutostartFolder(FldCnt) = RegGetString(0&, sKey, aParams(k))
                    aDes(FldCnt) = sSID & " " & aDesConst(k)
                End If
            Next
        End If
    Next
    
    ReDim Preserve sAutostartFolder(FldCnt)
    ReDim Preserve aDes(FldCnt)
    
    For k = 1 To UBound(sAutostartFolder)
        sAutostartFolder(k) = EnvironW(sAutostartFolder(k))
    Next
    
    ' adding all similar folders in c:\users (in case user isn't logged - so HKU\SID willn't be exist for him, cos his hive is not mounted)
    
    For i = 1 To colProfiles.Count
        'not current user
        If StrComp(colProfiles(i), UserProfile, 1) <> 0 Then
            If Len(colProfiles(i)) <> 0 Then
                ReDim Preserve sAutostartFolder(UBound(sAutostartFolder) + 1)
                ReDim Preserve aDes(UBound(aDes) + 1)
                sAutostartFolder(UBound(sAutostartFolder)) = Replace$(StartupCU, UserProfile, colProfiles(i), 1, 1, 1)
                aDes(UBound(aDes)) = "Startup other users"
            End If
        End If
    Next
    
    DeleteDuplicatesInArray sAutostartFolder, vbTextCompare, DontCompress:=True
    
    For k = 1 To UBound(sAutostartFolder)
        
        sFolder = sAutostartFolder(k)
        
        If 0 <> Len(sFolder) Then
          If FolderExists(sFolder) Then
            
            Erase aFiles
            aFiles = ListFiles(sFolder)
            
            If SafeArrayGetDim(aFiles) Then
            
              For i = 0 To UBound(aFiles)
            
                sShortCut = GetFileNameAndExt(aFiles(i))

                If LCase$(sShortCut) <> "desktop.ini" Then

                  If Not FolderExists(sFolder & "\" & sShortCut) Then
                  
                    isDisabled = False
              
                    If OSver.MajorMinor >= 6.2 Then  ' Win 8+

                        If StrInParamArray(aDes(k), "Startup", "User Startup", "Global Startup", "Global User Startup") Then

                            Select Case aDes(k)
                                Case "Startup": sHive = "HKCU"
                                Case "User Startup": sHive = "HKCU"
                                Case "Global Startup": sHive = "HKLM"
                                Case "Global User Startup": sHive = "HKLM"
                            End Select

                            sKeyDisable = sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder"

                            If RegValueExists(0&, sKeyDisable, sShortCut) Then

                                ReDim bData(0)
                                bData() = RegGetBinary(0&, sKeyDisable, sShortCut)
                        
                                If UBound(bData) >= 11 Then
                            
                                    GetMem4 ByVal VarPtr(bData(0)), flagDisabled

                                    If flagDisabled <> 2 Then
                        
                                        isDisabled = True
                                        dDate = ConvertFileTimeToLocalDate(VarPtr(bData(4)))
                                    End If
                                End If
                            End If
                        End If
                    End If
                  
                  
                    sFile = ""
                    bLink = False
                    bPE_EXE = False
                    
                    sLinkPath = sFolder & "\" & sShortCut
                    sLinkExt = UCase$(GetExtensionName(sShortCut))
                    
                    'Example:
                    '"O4 - Global User AltStartup: "
                    '"O4 - S-1-5-19 User AltStartup: "
                    If isDisabled Then
                        sHit = "O4 - " & sHive & "\..\StartupApproved\StartupFolder: " 'if you change it, change fix also !!!
                    Else
                        sHit = "O4 - " & aDes(k) & ": "
                    End If
                    
                    If StrInParamArray(sLinkExt, ".LNK", ".URL", ".WEBSITE", ".PIF") Then bLink = True
                    
                    If Not bLink Or sLinkExt = ".PIF" Then  'not a Shortcut ?
                        bPE_EXE = isPE_EXE(sLinkPath)       'PE EXE ?
                    End If
                    
                    If bLink Then
                        sTarget = GetFileFromShortcut(sLinkPath, sArguments)
                            
                        sHit = sHit & sShortCut & "    ->    " & sTarget & " " & sArguments '& " (User '" & sUserName & "')"
                    Else
                        sHit = sHit & sShortCut & IIf(bPE_EXE, "    ->    (PE EXE)", "")
                    End If
                    
                    If isDisabled Then sHit = sHit & " (" & Format$(dDate, "yyyy\/mm\/dd") & ")"
                    
                    If Not IsOnIgnoreList(sHit) Then
                               
                        If bMD5 Then
                            If Not bLink Or bPE_EXE Then
                                sHit = sHit & GetFileMD5(sLinkPath)
                            Else
                                If 0 <> Len(sTarget) Then
                                    sHit = sHit & GetFileMD5(sTarget)
                                End If
                            End If
                        End If
                        
                        With Result
                          If isDisabled Then
                            .Section = "O4"
                            .HitLineW = sHit
                            .Alias = sHive & "\..\StartupApproved\StartupFolder:"
                            .RunObject = sLinkPath
                            ReDim .RegKey(0)
                            .RegKey(0) = sKeyDisable
                            .RegParam = sShortCut
                            .CureType = FILE_BASED Or REGISTRY_PARAM_BASED
                          Else
                            .Section = "O4"
                            .HitLineW = sHit
                            .Alias = aDes(k)
                            .RunObject = sLinkPath
                            .ExpandedTarget = sTarget
                            .CureType = FILE_BASED
                          End If
                        End With
                        AddToScanResults Result
                    End If
                  End If
                End If
              Next
            End If
          End If
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain.CheckO4_AutostartFolder"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckO4Item()
    On Error GoTo ErrorHandler:
    
    'Alpha 1.0. // Dragokas. Reworked. Bugs fix. Deleted x64/x32 views shared keys.
    'Added support of msconfig disabled items. Unicode support.
    
    '2.6.1.25 [05.06.16] // Dragokas. Full revision, simplifying, merging CheckO4ItemX86, CheckO4ItemUsers to 1 func.
    
    ' look at keys affected by wow64 redirector
    ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa384253(v=vs.85).aspx
    ' http://safezone.cc/threads/27567/
    
    Dim aSID() As String, aUser() As String, aHives() As String, i&
    
    GetUserNamesAndSids aSID(), aUser()
    
    ReDim aHives(UBound(aSID) + 2)  '+ HKLM, HKCU
    ReDim Preserve aUser(UBound(aHives))
    
    'Convert SID -> to hive
    For i = 0 To UBound(aSID)
        aHives(i) = "HKU\" & aSID(i)
    Next
    'Add HKLM, HKCU
    aHives(UBound(aHives) - 1) = "HKLM"
    aUser(UBound(aHives) - 1) = "All users"
    
    aHives(UBound(aHives)) = "HKCU"
    aUser(UBound(aHives)) = GetUser()
    
    'Scanning routines
    
    CheckO4_RegRuns aHives(), aUser()
    
    CheckO4_MSConfig aHives(), aUser()
    
    CheckO4_AutostartFolder aSID(), aUser()
    
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_CheckO4Item"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub FixO4Item(sItem$)
    'O4 - Enumeration of autoloading Regedit entries
    'O4 - HKLM\..\Run: [blah] program.exe
    'O4 - Startup: bla.lnk = c:\bla.exe
    'O4 - HKU\S-1-5-19\..\Run: [blah] program.exe (Username 'Joe')
    'O4 - Startup: bla.exe
    
    On Error GoTo ErrorHandler:
    
    Dim Result As TYPE_Scan_Results
    Dim sFile$, i&
    
    If Not GetScanResults(sItem, Result) Then Exit Sub

    With Result
    
        If InStr(sItem, "StartupApproved\StartupFolder") <> 0 Then
            If FileExists(.ExpandedTarget) Then KillProcessByFile .ExpandedTarget
            If FileExists(.RunObject) Then
                If DeleteFileForce(.RunObject) Then
                    RegDelVal 0&, .RegKey(0), .RegParam 'del param. only if file successfully deleted
                End If
            Else
                RegDelVal 0&, .RegKey(0), .RegParam
            End If
            Exit Sub
        End If
        
        If StrBeginWith(sItem, "O4 - MSConfig\startupfolder: ") Then
            RegDelKey 0&, .RegKey(0)
            If FileExists(.RunObject) Then
                DeleteFileForce .RunObject
            End If
            Exit Sub
        End If
    
        Select Case .CureType
        
            Case REGISTRY_KEY_BASED
                RegDelKey 0&, .RegKey(0), .Redirected
                
            Case REGISTRY_PARAM_BASED
                If 0 <> Len(.RunObject) Then
                    sFile = GetLongPath(.RunObject)
                Else
                    sFile = GetLongPath(.RegParam)
                End If
                If FileExists(sFile) Then KillProcessByFile sFile
                
                If InStr(sItem, "StartupApproved\Run") <> 0 Then
                    RegDelVal 0&, .RegKey(0), .RegParam
                    RegDelVal 0&, .RegKey(1), .RegParam, .Redirected
                Else
                    For i = 0 To UBound(.RegKey)
                        RegDelVal 0&, .RegKey(i), .RegParam, .Redirected
                    Next
                End If
                
            Case FILE_BASED
                If FileExists(.ExpandedTarget) Then KillProcessByFile .ExpandedTarget
                If FileExists(.RunObject) Then
                    If Not DeleteFileForce(.RunObject) Then
                        MsgBoxW Replace$(Translate(320), "[]", sItem) & " " & _
                           IIf(bIsWinNT, Translate(321), Translate(322)) & _
                           " " & Translate(323), vbExclamation
            '            msgboxW "Unable to delete the file:" & vbCrLf & _
            '                   sItem & vbCrLf & vbCrLf & "The file " & _
            '                   "may be in use. Use " & IIf(bIsWinNT, _
            '                   "Task Manager", "a process killer like " & _
            '                   "ProcView") & " to shutdown the program " & _
            '                   "and run HiJackThis again to delete the file.", vbExclamation
                    End If
                End If
        End Select
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO4Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckO5Item()
    Dim sControlIni$, sDummy$, sHit$
    On Error GoTo ErrorHandler:
    
    sControlIni = sWinDir & "\control.ini"
    If Dir$(sControlIni) = vbNullString Then Exit Sub
    
    'sDummy = String$(5, " ")
    'GetPrivateProfileString "don't load", "inetcpl.cpl", "", sDummy, 5, sControlIni
    'sDummy = RTrim$(sDummy)
    
    IniGetString sControlIni, "don't load", "inetcpl.cpl"
    sDummy = RTrimNull(sDummy)
    
    If sDummy <> vbNullString Then
        sHit = "O5 - control.ini: inetcpl.cpl=" & sDummy
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O5", sHit
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO5Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO5Item(sItem$)
    'O5 - Blocking of loading Internet Options in Control Panel
    'WritePrivateProfileString "don't load", "inetcpl.cpl", vbNullString, "control.ini"
    On Error GoTo ErrorHandler:
    IniSetString "control.ini", "don't load", "inetcpl.cpl", vbNullString
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO5Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckO6Item()
    On Error GoTo ErrorHandler:
    
    'HKEY_CURRENT_USER\ software\ policies\ microsoft\
    'internet explorer. If there are sub folders called
    '"restrictions" and/or "control panel", delete them
    
    Dim sHit$, Hive, i&, Key$(2), Des$(2), Result As TYPE_Scan_Results
    'keys 0,1,2 - are x6432 shared.
    
    Key(0) = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    Des(0) = "Software\Policies\Microsoft\Internet Explorer\Restrictions present"
    
    Key(1) = "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions"
    Des(1) = "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions present"
    
    Key(2) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
    Des(2) = "Software\Policies\Microsoft\Internet Explorer\Control Panel present"
    
    For Each Hive In Array("HKCU", "HKLM")
        For i = 0 To UBound(Key)
            If RegKeyHasValues(0&, Hive & "\" & Key(i)) Then
                sHit = "O6 - " & Hive & "\" & Des(i) '& " " & IIf(Hive = "HKCU", "(HKCU)", "(HKLM)")
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O6"
                        .HitLineW = sHit
                        ReDim .RegKey(0)
                        .RegKey(0) = Hive & "\" & Key(i)
                    End With
                End If
                AddToScanResults Result
            End If
        Next
    Next
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO6Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO6Item(sItem$)
    On Error GoTo ErrorHandler:
    'O6 - Disabling of Internet Options' Main tab with Policies
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    RegDelKey 0&, Result.RegKey(0), True
    RegDelKey 0&, Result.RegKey(0), False
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO6Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckO7Item()
    On Error GoTo ErrorHandler:
    Dim lData&, sHit$, Hive, iPos As Long, Result As TYPE_Scan_Results

    'http://www.oszone.net/11424

    '//TODO:
    '%WinDir%\System32\GroupPolicyUsers"
    '%WinDir%\System32\GroupPolicy"
    'HKEY_CURRENT_USER\Software\Policies\Microsoft
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies

    For Each Hive In Array("HKCU", "HKLM")
        'key - x64 Shared
        lData = RegGetDword(0&, Hive & "\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools")
        If lData <> 0 Then
            'sHit = "O7 - " & Hive & "\Software\Microsoft\Windows\CurrentVersion\Policies\System, DisableRegedit=" & lData
            sHit = "O7 - Policy:" & " (" & Hive & ") " & "DisableRegedit=" & lData
            
            If Not IsOnIgnoreList(sHit) Then
                With Result
                    .Section = "O7"
                    .HitLineW = sHit
                    ReDim .RegKey(0)
                    .RegKey(0) = Hive & "\Software\Microsoft\Windows\CurrentVersion\Policies\System"
                    .RegParam = "DisableRegistryTools"
                    .CureType = REGISTRY_PARAM_BASED
                End With
                AddToScanResults Result
            End If
        End If
        
    Next
    
    'IPSec policy
    'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\
    'secpol.msc
    
    'ipsecPolicy{GUID}                  'example: 5d57bbac-8464-48b2-a731-9dd7e6f65c9f
    
    '\ipsecName                         -> Name of policy
    '\whenChanged                       -> Date in Unix format ( ConvertUnixTimeToLocalDate )
    '\ipsecNFAReference [REG_MULTI_SZ]  -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNFA{GUID_1} '96372b24-f2bf-4f50-a036-5897aac92f2f
                                                   'SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNFA{GUID_2} '8c676c64-306c-47db-ab50-e0108a1621dd
    
    '\ipsecISAKMPReference              -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecISAKMPPolicy{GUID} '738d84c5-d070-4c6c-9468-12b171cfd10e
    
    '--------------------------------
    'ipsecNFA{GUID}
    
    '\ipsecNegotiationPolicyReference     -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNegotiationPolicy{GUID} '7c5a4ff0-ae4b-47aa-a2b6-9a72d2d6374c
    '\ipsecFilterReference [REG_MULTI_SZ] -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecFilter{GUID_1} 'c73baa5d-71a6-4533-bf7d-f640b1ff2eb8
    
    '--------------------------------
    'ipsecNegotiationPolicy{GUID}
    
    Dim i As Long, KeyPolicy() As String, IPSecName$, KeyNFA() As String, KeyNegotiation As String, dModify As Date, lModify As Long, IPSecID As String
    Dim KeyISAKMP As String, J As Long, KeyFilter() As String, k As Long, NegAction As String, NegType As String, bEnabled As Boolean, sActPolicy As String
    Dim bFilterData() As Byte, IP(1) As String, RuleAction As String, bMirror As Boolean, DataSerialized As String
    Dim IP_Type(1) As String, M As Long, n As Long, PortNum(1) As Long, PortType As String
    
    For i = 1 To RegEnumSubkeysToArray(0&, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local", KeyPolicy())
        
      If StrBeginWith(KeyPolicy(i), "ipsecPolicy{") Then
        
        'what policy is currently active?
        sActPolicy = RegGetString(0&, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local", "ActivePolicy")
        
        bEnabled = (StrComp(sActPolicy, "SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\" & KeyPolicy(i), 1) = 0)
        
        'add prefix
        KeyPolicy(i) = "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyPolicy(i)
        
        bMirror = False
        RuleAction = ""
        
        IPSecID = Mid$(KeyPolicy(i), InStrRev(KeyPolicy(i), "{"))
        
        IPSecName = RegGetString(0&, KeyPolicy(i), "ipsecName")
        
        lModify = modRegistry.RegGetDword(0&, KeyPolicy(i), "whenChanged")
            
        dModify = ConvertUnixTimeToLocalDate(lModify)
        
        KeyISAKMP = RegGetString(0&, KeyPolicy(i), "ipsecISAKMPReference")
        KeyISAKMP = MidFromCharRev(KeyISAKMP, "\")
        KeyISAKMP = IIf(KeyISAKMP = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyISAKMP)
        
        Erase KeyNFA
        KeyNFA() = RegGetMultiSZ(0&, KeyPolicy(i), "ipsecNFAReference")
        '() -> ipsecNegotiationPolicy
        '() -> ipsecFilter (optional)
            
        If SafeArrayGetDim(KeyNFA) Then
            
          For J = 0 To UBound(KeyNFA)
            KeyNFA(J) = MidFromCharRev(KeyNFA(J), "\")
            KeyNFA(J) = IIf(KeyNFA(J) = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyNFA(J))
          Next
            
          For J = 0 To UBound(KeyNFA)

            KeyNegotiation = RegGetString(0&, KeyNFA(J), "ipsecNegotiationPolicyReference")
            KeyNegotiation = MidFromCharRev(KeyNegotiation, "\")
            KeyNegotiation = IIf(KeyNegotiation = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyNegotiation)
            
            NegType = RegGetString(0&, KeyNegotiation, "ipsecNegotiationPolicyType")
            NegAction = RegGetString(0&, KeyNegotiation, "ipsecNegotiationPolicyAction")
            
            If StrComp(NegType, "{62f49e10-6c37-11d1-864c-14a300000000}", 1) = 0 Then
                If StrComp(NegAction, "{8a171dd2-77e3-11d1-8659-a04f00000000}", 1) = 0 Then
                    RuleAction = "Allow"
                ElseIf StrComp(NegAction, "{3f91a819-7647-11d1-864d-d46a00000000}", 1) = 0 Then
                    RuleAction = "Block"
                Else
                    RuleAction = "Unknown"
                End If
            Else
                RuleAction = "Unknown"
            End If
            
            Erase KeyFilter
            KeyFilter() = RegGetMultiSZ(0&, KeyNFA(J), "ipsecFilterReference")
            
            If SafeArrayGetDim(KeyFilter) Then
                
                For k = 0 To UBound(KeyFilter)
                    KeyFilter(k) = MidFromCharRev(KeyFilter(k), "\")
                    KeyFilter(k) = IIf(KeyFilter(k) = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyFilter(k))
                Next
                
                For k = 0 To UBound(KeyFilter)
                    
                    bFilterData() = RegGetBinary(0&, KeyFilter(k), "ipsecData")
                    
                    If SafeArrayGetDim(bFilterData) Then
                      If UBound(bFilterData) = &H71 Then

                        '00,00,00,00,00,00,00,00 -> any IP
                        'xx,xx,xx,xx,ff,ff,ff,ff -> specified IP / subnet
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 0 -> my IP
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 0x81 -> DNS-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 0x82 -> WINS-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 0x83 -> DHCP-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 0x84 -> Gateway
                        '
                        '[0x4E] == 1 -> mirrored
                        '
                        '[0x66] -> port type
                        '[0x6A] -> port number (source)
                        '[0x6C] -> port number (destination)
                        
                        bMirror = (bFilterData(&H4E) = 1)
                        PortNum(0) = bFilterData(&H6A)
                        PortNum(1) = bFilterData(&H6C)
                        
                        Select Case bFilterData(&H66)
                            Case 0: PortType = "Any"
                            Case 6: PortType = "TCP"
                            Case 17: PortType = "UDP"
                            Case 1: PortType = "ICMP"
                            Case 27: PortType = "RDP"
                            Case 8: PortType = "EGP"
                            Case 20: PortType = "HMP"
                            Case 255: PortType = "RAW"
                            Case 66: PortType = "RVD"
                            Case 22: PortType = "XNS-IDP"
                            Case Else: PortType = "type: " & CLng(bFilterData(&H66))
                        End Select
                        
                        For M = 0 To 1
                        
                            IP(M) = bFilterData(&H52 + 8 * M) & "." & _
                                bFilterData(&H52 + 1 + 8 * M) & "." & _
                                bFilterData(&H52 + 2 + 8 * M) & "." & _
                                bFilterData(&H52 + 3 + 8 * M)
                        
                            If IP(M) = "0.0.0.0" Then IP(M) = ""
                            DataSerialized = ""
                            
                            For n = &H52 + 8 * M To &H52 + 7 + 8 * M
                                DataSerialized = DataSerialized & Right$("0" & Hex(bFilterData(n)), 2) & ","
                            Next
                            DataSerialized = LCase$(Left$(DataSerialized, Len(DataSerialized) - 1))
                            
                            Select Case DataSerialized
                                Case "00,00,00,00,00,00,00,00": IP_Type(M) = "Any IP"
                                Case "00,00,00,00,ff,ff,ff,ff"
                                    Select Case bFilterData(&H6F)
                                        Case 0: IP_Type(M) = "my IP"
                                        Case &H81: IP_Type(M) = "DNS-servers"
                                        Case &H82: IP_Type(M) = "WINS-servers"
                                        Case &H83: IP_Type(M) = "DHCP-servers"
                                        Case &H84: IP_Type(M) = "Gateway"
                                        Case Else: IP_Type(M) = "Unknown"
                                    End Select
                                Case Else
                                    If StrEndWith(DataSerialized, "ff,ff,ff,ff") Then
                                        IP_Type(M) = "IP: "
                                    Else
                                        IP_Type(M) = "Unknown"
                                    End If
                            End Select
                        Next
                        
                        'keys:
                        'KeyPolicy(i) - 1
                        'KeyISAKMP - 1
                        'KeyNFA(j) - 0 to ...
                        'KeyNegotiation - 1
                        'KeyFilter(k) - 0 to ...
                        
                        'flags:
                        'bEnabled - policy enabled ?
                        'bMirror - true, if rule also applies to reverse direction: from destination to source
                        
                        'Other:
                        'IPSecName - name of policy
                        'IPSecID - identifier in registry
                        'dModify - date last modified
                        'RuleAction - action for filter
                        'PortNum()
                        'PortType
                        
                        'example:
'O7 - IPSec: (Enabled) IP_Policy_Name [yyyy/mm/dd] - {5d57bbac-8464-48b2-a731-9dd7e6f65c9f} - Source: My IP - Destination: 8.8.8.8 (Port 80 TCP) - (mirrored) Action: Block
                        
                        sHit = "O7 - IPSec: " & IIf(bEnabled, "(enabled)", "(disabled)") & " " & IPSecName & " " & _
                            "[" & Format(dModify, "yyyy\/mm\/dd") & "]" & " - " & IPSecID & " - " & _
                            "Source: " & IP_Type(0) & IP(0) & _
                            IIf((PortType = "TCP" Or PortType = "UDP") And PortNum(0) <> 0, " (Port " & PortNum(0) & " " & PortType & ")", "") & " - " & _
                            "Destination: " & IP_Type(1) & IP(1) & _
                            IIf((PortType = "TCP" Or PortType = "UDP") And PortNum(1) <> 0, " (Port " & PortNum(1) & " " & PortType & ")", "") & " - " & _
                            IIf(bMirror, "(mirrored)", "") & " " & "Action: " & RuleAction

                        If Not IsOnIgnoreList(sHit) Then
                            With Result
                                .Section = "O7"
                                .HitLineW = sHit
                                ReDim .RegKey(2 + UBound(KeyNFA) + 1 + UBound(KeyFilter) + 1)
                                .RegKey(0) = KeyPolicy(i)
                                If KeyISAKMP <> "" Then .RegKey(1) = KeyISAKMP
                                n = 1
                                For M = 0 To UBound(KeyNFA)
                                    If KeyNFA(M) <> "" Then
                                        n = n + 1
                                        .RegKey(n) = KeyNFA(M)
                                    End If
                                Next
                                n = n + 1
                                If KeyNegotiation <> "" Then .RegKey(n) = KeyNegotiation
                                For M = 0 To UBound(KeyFilter)
                                    If KeyFilter(M) <> "" Then
                                        n = n + 1
                                        .RegKey(n) = KeyFilter(M)
                                    End If
                                Next
                                .CureType = REGISTRY_KEY_BASED
                            End With
                        End If
                        AddToScanResults Result
                        
                      End If
                    End If
                Next
            End If
          Next
        End If
      End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO7Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO7Item(sItem$)
    'O7 - Disabling of Policies
    On Error GoTo ErrorHandler:
    
    Dim Result As TYPE_Scan_Results, i As Long
    If Not GetScanResults(sItem, Result) Then Exit Sub
    
    With Result
        If .CureType = REGISTRY_KEY_BASED Then
            For i = 0 To UBound(.RegKey)
                RegDelKey 0&, .RegKey(i)
            Next
        ElseIf .CureType = REGISTRY_PARAM_BASED Then
            RegDelVal 0&, .RegKey(0), .RegParam
        End If
    End With
    
    Call UpdatePolicy
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO7Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO8Item()
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt
    
    On Error GoTo ErrorHandler:
    Dim hKey&, hKey2&, i&, sName$, lpcName&, sData$, sHit$, Result As TYPE_Scan_Results
    
    'HKCU key is not redirected here
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr("Software\Microsoft\Internet Explorer\MenuExt"), 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        i = 0
        sName = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sName)
        
        Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
            sName = RTrimNull(sName)
            sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & sName, vbNullString)
            If sData <> vbNullString Then
                sHit = "O8 - Extra context menu item: " & sName & " - " & sData
                If Not IsOnIgnoreList(sHit) Then
                    'md5 doesn't seem useful here
                    If bMD5 Then sHit = sHit & GetFileMD5(sData)
                    With Result
                        .Section = "O8"
                        .HitLineW = sHit
                        ReDim .RegKey(0)
                        .RegKey(0) = "HKCU\" & "Software\Microsoft\Internet Explorer\MenuExt\" & sName
                    End With
                    AddToScanResults Result
                End If
            End If
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            i = i + 1
        Loop
        RegCloseKey hKey
    End If
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_CheckO8Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO8Item(sItem$)
    'O8 - Extra context menu items
    'O8 - Extra context menu item: [name] - html file
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt
    
    On Error GoTo ErrorHandler:
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    With Result
        'RegDelKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & sName
        RegDelKey 0&, .RegKey(0)
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO8Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO9Item()
    'HKLM\Software\Microsoft\Internet Explorer\Extensions
    'HKCU\..\etc
    
    On Error GoTo ErrorHandler:
    Dim hKey&, hKey2&, i&, sData$, sCLSID$, sCLSID2$, lpcName&, sFile$, sHit$, sBuf$, IsInfected As Boolean, Result As TYPE_Scan_Results
    Dim Wow6432Redir As Boolean, UseWow, vHive, lHive&
    
  For Each vHive In Array(HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER)
  For Each UseWow In Array(False, True)
  
    lHive = vHive
    Wow6432Redir = UseWow
    If Wow6432Redir And (bIsWin32 Or lHive = HKEY_CURRENT_USER) Then Exit For
    
    'open root key
    If RegOpenKeyExW(lHive, StrPtr("Software\Microsoft\Internet Explorer\Extensions"), 0, _
      KEY_ENUMERATE_SUB_KEYS Or (KEY_WOW64_64KEY And Not Wow6432Redir), hKey) = 0 Then
        i = 0
        sCLSID = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sCLSID)
        'start enum of root key subkeys (i.e., extensions)
        Do While RegEnumKeyExW(hKey, i, StrPtr(sCLSID), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
            sCLSID = TrimNull(sCLSID)
            If sCLSID = "CmdMapping" Then GoTo NextExt:
            
            'check for 'MenuText' or 'ButtonText'
            sData = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "ButtonText", Wow6432Redir)
            
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
            sFile = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "Exec", Wow6432Redir)
            If sFile = vbNullString Then
                sFile = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "Script", Wow6432Redir)
                If sFile = vbNullString Then
                    sCLSID2 = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "BandCLSID", Wow6432Redir)
                    sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, Wow6432Redir)
                    If sFile = vbNullString Then
                        sCLSID2 = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "CLSIDExtension", Wow6432Redir)
                        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, Wow6432Redir)
                        If sFile = vbNullString Then
                            sCLSID2 = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\TreatAs", vbNullString, Wow6432Redir)
                            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, Wow6432Redir)
                            If sFile = vbNullString Then
                                sCLSID2 = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "CLSID", Wow6432Redir)
                                sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, Wow6432Redir)
                            End If
                        End If
                    End If
                End If
            End If
            
            If sFile <> vbNullString Then

                'expand %systemroot% var
                'sFile = replace$(sFile, "%systemroot%", sWinDir, , , vbTextCompare)
                sFile = UnQuote(EnvironW(sFile))
                
                'strip stuff from res://[dll]/page.htm to just [dll]
                If InStr(1, sFile, "res://", vbTextCompare) = 1 And _
                   (LCase$(Right$(sFile, 4)) = ".htm" Or LCase$(Right$(sFile, 4)) = "html") Then
                    sFile = Mid$(sFile, 7)
                    sFile = Left$(sFile, InStrRev(sFile, "/") - 1)
                End If
                
                'remove other stupid prefixes
                If InStr(1, sFile, "file://", 1) = 1 Then
                    If Not FileExists(Mid$(sFile, 8)) Then sFile = sFile & " (file missing)"
                Else
                    If InStr(1, sFile, "http:", 1) <> 1 And _
                      InStr(1, sFile, "https:", 1) <> 1 Then
                        If Not FileExists(sFile) Then
                            sFile = sFile & " (file missing)"
                        End If
                    End If
                End If
            Else
                sFile = "(no file)"
            End If
            
            IsInfected = True
            If sFile = PF_64 & "\Messenger\msmsgs.exe" Then
                If IsMicrosoftFile(sFile) And Not bIgnoreAllWhitelists Then IsInfected = False
            End If
            
            If IsInfected Then
            
              If sData = vbNullString Then sData = "(no name)"
              If Left$(sData, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sData)
                If 0 <> Len(sBuf) Then sData = sBuf
              End If
            
              'O9 - Extra button:
              'O9-32 - Extra button:
              sHit = IIf(bIsWin32 Or lHive = HKEY_CURRENT_USER, "O9", IIf(Wow6432Redir, "O9-32", "O9")) & _
                " - Extra button: " & sData & " - " & sCLSID & " - " & sFile & IIf(lHive = HKEY_LOCAL_MACHINE, " (HKLM)", "")
              
              If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                With Result
                    .Section = "O9"
                    .HitLineW = sHit
                    .lHive = lHive
                    .CLSID = sCLSID
                    .RunObject = sFile
                End With
                AddToScanResults Result
              End If
            
              sData = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Extensions\" & sCLSID, "MenuText", Wow6432Redir)
            
              If Left$(sData, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sData)
                If 0 <> Len(sBuf) Then sData = sBuf
              End If
                
              'don't show it again in case sdata=null
              If sData <> vbNullString Then
                'O9 - Extra 'Tools' menuitem:
                'O9-32 - Extra 'Tools' menuitem:
                sHit = IIf(bIsWin32 Or lHive = HKEY_CURRENT_USER, "O9", IIf(Wow6432Redir, "O9-32", "O9")) & _
                  " - Extra 'Tools' menuitem: " & sData & " - " & sCLSID & " - " & sFile & IIf(lHive = HKEY_LOCAL_MACHINE, " (HKLM)", "")
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    With Result
                        .Section = "O9"
                        .HitLineW = sHit
                        .lHive = lHive
                        .CLSID = sCLSID
                        .RunObject = sFile
                        .Redirected = Wow6432Redir
                    End With
                    AddToScanResults Result
                End If
              End If
            End If
NextExt:
            sCLSID = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sCLSID)
            i = i + 1
        Loop
        RegCloseKey hKey
    End If
  Next
  Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO9Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO9Item(sItem$)
    'O9 - Extra buttons/Tools menu items
    'O9 - Extra button: [name] - [CLSID] - [file] [(HKCU)]
    
    On Error GoTo ErrorHandler:
    Dim Result As TYPE_Scan_Results

    If Not GetScanResults(sItem, Result) Then Exit Sub
    With Result
        RegDelKey .lHive, "Software\Microsoft\Internet Explorer\Extensions\" & .CLSID, .Redirected
        RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", .CLSID
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO9Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO10Item()
    CheckLSP
End Sub

Public Sub CheckO11Item()
    On Error GoTo ErrorHandler:
    
    'HKLM\Software\Microsoft\Internet Explorer\AdvancedOptions
    Dim hKey&, i&, sKey$, sName$, lpcName&, sHit$, UseWow, Wow6432Redir As Boolean, Result As TYPE_Scan_Results
    
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr("Software\Microsoft\Internet Explorer\AdvancedOptions"), 0, _
      KEY_ENUMERATE_SUB_KEYS Or (KEY_WOW64_64KEY And Not Wow6432Redir), hKey) = 0 Then
        
        sKey = String$(MAX_KEYNAME, 0)
        lpcName = Len(sKey)
        i = 0
        Do While RegEnumKeyExW(hKey, i, StrPtr(sKey), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
            sKey = Left$(sKey, InStr(sKey, vbNullChar) - 1)
            sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\AdvancedOptions\" & sKey, "Text", Wow6432Redir)
            If InStr("JAVA_VM.JAVA_SUN.BROWSE.ACCESSIBILITY.SEARCHING." & _
                     "HTTP1.1.MULTIMEDIA.Multimedia.CRYPTO.PRINT." & _
                     "TOEGANKELIJKHEID.TABS.INTERNATIONAL*.ACCELERATED_GRAPHICS", sKey) = 0 And _
               sName <> vbNullString Then
               
               'O11 - Options group:
               'O11-32 - Options group:
               sHit = IIf(bIsWin32, "O11", IIf(Wow6432Redir, "O11-32", "O11")) & " - Options group: [" & sKey & "] " & sName
                
                If bIgnoreAllWhitelists Or Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O11"
                        .HitLineW = sHit
                        ReDim .RegKey(0)
                        .RegKey(0) = "HKLM\" & sKey
                        .Redirected = Wow6432Redir
                    End With
                    AddToScanResults Result
                End If
            End If
            sKey = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sKey)
            i = i + 1
        Loop
        RegCloseKey hKey
    End If
  Next
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO11Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO11Item(sItem$)
    'O11 - Options group: [BLA] Blah"
    
    On Error GoTo ErrorHandler:
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    With Result
        'RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\AdvancedOptions\" & sKey
        RegDelKey 0&, .RegKey(0), .Redirected
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO11Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO12Item()
    'HKLM\Software\Microsoft\Internet Explorer\Plugins\Extensions
    'HKLM\Software\Microsoft\Internet Explorer\Plugins\MIME
    
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, sName$, sFile$, sHit$, lpcName&, Key, sKey$, UseWow, Wow6432Redir As Boolean, Result As TYPE_Scan_Results
    
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    For Each Key In Array( _
      "Software\Microsoft\Internet Explorer\Plugins\Extension", _
      "Software\Microsoft\Internet Explorer\Plugins\MIME")
      
      sKey = Key
      
      If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr(sKey), 0, KEY_ENUMERATE_SUB_KEYS Or (KEY_WOW64_64KEY And Not Wow6432Redir), hKey) = 0 Then
      
        sName = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sName)
        i = 0
        
        Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
            sName = Left$(sName, InStr(sName, vbNullChar) - 1)
            sFile = RegGetString(HKEY_LOCAL_MACHINE, sKey & "\" & sName, "Location", Wow6432Redir)
            If sFile <> vbNullString Then
                'O12 - Plugin
                'O12-32 - Plugin
                sHit = IIf(bIsWin32, "O12", IIf(Wow6432Redir, "O12-32", "O12")) & " - Plugin for " & sName & ": " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    With Result
                        .Section = "O12"
                        .HitLineW = sHit
                        ReDim .RegKey(0)
                        .RegKey(0) = "HKLM\" & sKey & "\" & sName
                        .RunObject = sFile
                        .Redirected = Wow6432Redir
                    End With
                    AddToScanResults Result
                End If
            End If
            
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            i = i + 1
        Loop
        RegCloseKey hKey
      End If
    Next
  Next
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO12Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO12Item(sItem$)
    'O12 - Plugin for .ofb: C:\Win98\blah.dll
    'O12 - Plugin for text/blah: C:\Win98\blah.dll
    
    On Error GoTo ErrorHandler:
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    
    If Not bShownToolbarWarning And ProcessExist("iexplore.exe") Then
        MsgBoxW Translate(330), vbExclamation
'        msgboxW "HiJackThis is about to remove a " & _
'               "plugin from " & _
'               "your system. Close all Internet " & _
'               "Explorer windows before continuing for " & _
'               "the best chance of success.", vbExclamation
        bShownToolbarWarning = True
    End If
    
    With Result
        'RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Plugins\" & sType & sKey
        RegDelKey 0&, .RegKey(0), .Redirected
        DeleteFileWEx StrPtr(.RunObject)
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO12Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO13Item()
    On Error GoTo ErrorHandler:
    Dim sDummy$, sKeyURL$, sHit$, UseWow, Wow6432Redir As Boolean, Result As TYPE_Scan_Results
    Dim aKey() As String, aVal() As String, aExa() As String, aDes() As String, i As Long
    
    ReDim aKey(6)
    ReDim aVal(UBound(aKey))
    ReDim aExa(UBound(aKey))
    ReDim aDes(UBound(aKey))
    
    aKey(0) = "DefaultPrefix"
    aVal(0) = ""
    aExa(0) = "http://"
    aDes(0) = "DefaultPrefix"
    
    aKey(1) = "Prefixes"
    aVal(1) = "www"
    aExa(1) = "http://"
    aDes(1) = "WWW Prefix"
    
    aKey(2) = "Prefixes"
    aVal(2) = "www."
    aExa(2) = ""
    aDes(2) = "WWW. Prefix"
    
    aKey(3) = "Prefixes"
    aVal(3) = "home"
    aExa(3) = "http://"
    aDes(3) = "Home Prefix"
    
    aKey(4) = "Prefixes"
    aVal(4) = "mosaic"
    aExa(4) = "http://"
    aDes(4) = "Mosaic Prefix"
    
    aKey(5) = "Prefixes"
    aVal(5) = "ftp"
    aExa(5) = "ftp://"
    aDes(5) = "FTP Prefix"
    
    aKey(6) = "Prefixes"
    aVal(6) = "gopher"
    aExa(6) = "gopher://|"
    aDes(6) = "Gopher Prefix"
    
    sKeyURL = "Software\Microsoft\Windows\CurrentVersion\URL"

    For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If bIsWin32 And Wow6432Redir Then Exit For
    
        For i = 0 To UBound(aKey)
        
            sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeyURL & "\" & aKey(i), aVal(i), Wow6432Redir)
            If Not inArraySerialized(sDummy, aExa(i), "|", , , vbBinaryCompare) Then
                'infected!
                
                sHit = IIf(bIsWin32, "O13", IIf(Wow6432Redir, "O13-32", "O13")) & " - " & aDes(i) & ": " & sDummy
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O13"
                        .HitLineW = sHit
                        ReDim .RegKey(0)
                        .RegKey(0) = "HKLM\" & sKeyURL & "\" & aKey(i)
                        .RegParam = aVal(i)
                        .DefaultData = aExa(i)
                        .Redirected = Wow6432Redir
                    End With
                    AddToScanResults Result
                End If
            End If
        Next
    Next
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO13Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO13Item(sItem$)
    'defaultprefix fix
    'O13 - DefaultPrefix: http://www.hijacker.com/redir.cgi?
    'O13 - [WWW/Home/Mosaic/FTP/Gopher] Prefix: ..
    
    On Error GoTo ErrorHandler:
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    With Result
        RegSetStringVal 0&, .RegKey(0), .RegParam, CStr(.DefaultData), .Redirected
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO13Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO14Item()
    'O14 - Reset Websettings check
    On Error GoTo ErrorHandler:
    Dim sLine$, sHit$, ff%
    Dim sStartPage$, sSearchPage$, sMsStartPage$
    Dim sSearchAssis$, sCustSearch$
    Dim sFile$, lBOM&, aLogStrings() As String, i&
    
    sFile = sWinDir & "\inf\iereset.inf"
    
    If Not FileExists(sFile) Then Exit Sub
    If FileLenW(sFile) = 0 Then Exit Sub
    
    Dim b() As Byte
    ReDim b(1)
    ff = FreeFile()
    Open sFile For Binary Access Read As #ff
    Get #ff, 1, b()
    Close #ff
    aLogStrings = ReadFileToArray(sFile, IIf(b(0) = &HFF& And b(1) = &HFE&, True, False))
    
    For i = 0 To UBound(aLogStrings)
        sLine = aLogStrings(i)
        
            If InStr(sLine, "SearchAssistant") > 0 Then
                sSearchAssis = Mid$(sLine, InStr(sLine, "http://"))
                sSearchAssis = Left$(sSearchAssis, Len(sSearchAssis) - 1)
            End If
            If InStr(sLine, "CustomizeSearch") > 0 Then
                sCustSearch = Mid$(sLine, InStr(sLine, "http://"))
                sCustSearch = Left$(sCustSearch, Len(sCustSearch) - 1)
            End If
            If InStr(sLine, "START_PAGE_URL=") = 1 And _
               InStr(sLine, "MS_START_PAGE_URL") = 0 Then
                sStartPage = Mid$(sLine, InStr(sLine, "=") + 1)
                sStartPage = UnQuote(sStartPage)
            End If
            If InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sSearchPage = Mid$(sLine, InStr(sLine, "=") + 1)
                sSearchPage = UnQuote(sSearchPage)
            End If
            If InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sMsStartPage = Mid$(sLine, InStr(sLine, "=") + 1)
                sMsStartPage = UnQuote(sMsStartPage)
            End If
    Next
    
    'SearchAssistant = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm
    If sSearchAssis <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm" And _
      sSearchAssis <> g_DEFSEARCHASS Then
        sHit = "O14 - IERESET.INF: SearchAssistant=" & sSearchAssis
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'CustomizeSearch = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm
    If sCustSearch <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm" And _
      sCustSearch <> g_DEFSEARCHCUST Then
        sHit = "O14 - IERESET.INF: CustomizeSearch=" & sCustSearch
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'SEARCH_PAGE_URL = http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch
    If sSearchPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch" And _
      sSearchPage <> g_DEFSEARCHPAGE Then
        sHit = "O14 - IERESET.INF: SEARCH_PAGE_URL=" & sSearchPage
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'START_PAGE_URL  = http://www.msn.com
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If sStartPage <> "http://www.msn.com" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome" And _
       sStartPage <> g_DEFSTARTPAGE Then
        sHit = "O14 - IERESET.INF: START_PAGE_URL=" & sStartPage
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'MS_START_PAGE_URL=http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '(=START_PAGE_URL) http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If sMsStartPage <> vbNullString Then
        If sMsStartPage <> "http://www.msn.com" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome" And _
           sMsStartPage <> g_DEFSTARTPAGE Then
            sHit = "O14 - IERESET.INF: MS_START_PAGE_URL=" & sMsStartPage
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO14Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Function ReadFileToArray(sFile As String, Optional isUnicode As Boolean) As String()
    On Error GoTo ErrorHandler:
    Dim ff   As Integer
    Dim b()  As Byte
    Dim Text As String
    ff = FreeFile()
    Open sFile For Binary Access Read As #ff
        ReDim b(LOF(ff) - 1)
        Get #ff, 1, b()
    Close #ff
    If isUnicode Then
        Text = b()
        If b(0) = &HFF& And b(1) = &HFE& Then Text = Mid$(Text, 2)
    Else
        Text = StrConv(b(), vbUnicode, &H419&)
    End If
    Text = Replace$(Text, vbCr, vbNullString)
    ReadFileToArray = SplitSafe(Text, vbLf)
    Exit Function
ErrorHandler:
    ErrorMsg err, "modMain_ReadFileToArray"
    If inIDE Then Stop: Resume Next
End Function

Public Function ReadFileContents(sFile As String, isUnicode As Boolean) As String
    On Error GoTo ErrorHandler:
    Dim ff   As Integer
    Dim b()  As Byte
    Dim Text As String
    ff = FreeFile()
    Open sFile For Binary Access Read As #ff
        ReDim b(LOF(ff) - 1)
        Get #ff, 1, b()
    Close #ff
    If isUnicode Then
        Text = b()
        If b(0) = &HFF& And b(1) = &HFE& Then Text = Mid$(Text, 2)  ' - BOM UTF16-LE
    Else
        Text = StrConv(b(), vbUnicode, &H419&)
        If b(0) = &HEF& And b(1) = &HBB& And b(2) = &HBF& Then      ' - BOM UTF-8
            Text = Mid$(Text, 4)
        End If
    End If
    ReadFileContents = Text
    Exit Function
ErrorHandler:
    ErrorMsg err, "modMain_ReadFileContents"
    If inIDE Then Stop: Resume Next
End Function

Public Sub FixO14Item(sItem$)
    'resetwebsettings fix
    'O14 - IERESET.INF: [item]=[URL]
    
    On Error GoTo ErrorHandler:
    'sItem - not used
    Dim sLine$, sFixedIeResetInf$, ff%
    Dim i&, aLogStrings() As String, sFile$, isUnicode As Boolean
    
    sFile = sWinDir & "\INF\iereset.inf"
    
    If Not FileExists(sFile) Then Exit Sub
    ff = FreeFile()
    
    Dim b() As Byte
    ReDim b(1)
    ff = FreeFile()
    Open sFile For Binary Access Read As #ff
    Get #ff, 1, b()
    Close #ff
    If b(0) = &HFF& And b(1) = &HFE& Then isUnicode = True
    aLogStrings = ReadFileToArray(sFile, IIf(isUnicode, True, False))
    
    For i = 0 To UBound(aLogStrings)
        sLine = aLogStrings(i)

            If InStr(sLine, "SearchAssistant") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""SearchAssistant"",0,""" & _
                    IIf(g_DEFSEARCHASS <> "", g_DEFSEARCHASS, "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm") & """" & vbCrLf
            ElseIf InStr(sLine, "CustomizeSearch") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""CustomizeSearch"",0,""" & _
                    IIf(g_DEFSEARCHCUST <> "", g_DEFSEARCHCUST, "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm") & """" & vbCrLf
            ElseIf InStr(sLine, "START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "START_PAGE_URL=""" & _
                    IIf(g_DEFSTARTPAGE <> "", g_DEFSTARTPAGE, "http://www.msn.com") & """" & vbCrLf
            ElseIf InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "SEARCH_PAGE_URL=""" & _
                    IIf(g_DEFSEARCHPAGE <> "", g_DEFSEARCHPAGE, "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch") & """" & vbCrLf
            ElseIf InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "MS_START_PAGE_URL=""" & _
                    IIf(g_DEFSTARTPAGE <> "", g_DEFSTARTPAGE, "http://www.msn.com") & """" & vbCrLf
            Else
                sFixedIeResetInf = sFixedIeResetInf & sLine & vbCrLf
            End If
        
    Next
    sFixedIeResetInf = Left$(sFixedIeResetInf, Len(sFixedIeResetInf) - 2)   '-CrLf
    
    'SetFileAttributes StrPtr(sWinDir & "\INF\iereset.inf"), vbArchive  '???
    DeleteFileWEx (StrPtr(sFile))
    
    ff = FreeFile()
    
    If isUnicode Then
        b() = ChrW(-257) & sFixedIeResetInf
        Open sFile For Binary Access Write As #ff
        Put #ff, , b()
    Else
        Open sFile For Output As #ff
        Print #ff, sFixedIeResetInf
    End If
    
    Close #ff
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO14Item", "sItem=", sItem
    Close #ff
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO15Item()
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
    
    On Error GoTo ErrorHandler:
    Dim sZoneMapDomains$, sZoneMapRanges$, sZoneMapProtDefs$
    Dim sZoneMapEscDomains$, sZoneMapEscRanges$
    Dim sDomains$(), sSubDomains$()
    Dim i&, J&, sHit$, sIPRange$, UseWow, Wow6432Redir As Boolean
    sZoneMapDomains = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"
    sZoneMapRanges = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    sZoneMapProtDefs = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\ProtocolDefaults"
    sZoneMapEscDomains = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains"
    sZoneMapEscRanges = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscRanges"
    
    'enum all subkeys (i.e. all domains)
    sDomains = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sZoneMapDomains), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            If bIgnoreSafe Then
                For J = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sSafeRegDomains(J), CStr(sDomains(i)), vbTextCompare) > 0 Then
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
                Next J
            End If
            sSubDomains = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i)), "|")
            If UBound(sSubDomains) <> -1 Then
                'list any trusted subdomains for main domain
                For J = 0 To UBound(sSubDomains)
                    If RegGetDword(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i) & "\" & sSubDomains(J), "*") = 2 Then
                        'entire subdomain is trusted
                        sHit = "O15 - Trusted Zone: " & sSubDomains(J) & "." & sDomains(i)
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                    End If
                    If RegGetDword(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i) & "\" & sSubDomains(J), "http") = 2 Then
                        'only http on subdomain is trusted
                        sHit = "O15 - Trusted Zone: http://" & sSubDomains(J) & "." & sDomains(i)
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                    End If
                Next J
            End If
            'list main domain as well if that's trusted too (*grumble*)
            If RegGetDword(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i), "*") = 2 Then
                'entire domain is trusted
                sHit = "O15 - Trusted Zone: *." & sDomains(i)
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
            End If
            If RegGetDword(HKEY_CURRENT_USER, sZoneMapDomains & "\" & sDomains(i), "http") = 2 Then
                'only http on domain is trusted
                sHit = "O15 - Trusted Zone: http://*." & sDomains(i)
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
            End If
NextDomain:
        Next i
    End If

  'repeat for HKLM (domains)
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    sDomains = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sZoneMapDomains, Wow6432Redir), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            If bIgnoreSafe Then
                For J = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sSafeRegDomains(J), CStr(sDomains(i)), vbTextCompare) > 0 Then
                        If InStr(sDomains(i), "msn.com") = 0 Then
                            'it's a safe domain!
                            GoTo NextDomain2
                        Else
                            Exit For
                        End If
                    End If
                Next J
            End If
            sSubDomains = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i), Wow6432Redir), "|")
            If UBound(sSubDomains) <> -1 Then
                'list any trusted subdomains for main domain
                For J = 0 To UBound(sSubDomains)
                    If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i) & "\" & sSubDomains(J), "*", Wow6432Redir) = 2 Then
                        'entire subdomain is trusted
                        sHit = "O15 - Trusted Zone: " & sSubDomains(J) & "." & sDomains(i) & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                    End If
                    If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i) & "\" & sSubDomains(J), "http", Wow6432Redir) = 2 Then
                        'only http on subdomain is trusted
                        sHit = "O15 - Trusted Zone: http://" & sSubDomains(J) & "." & sDomains(i) & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                    End If
                Next J
            End If
            'list main domain as well, if applicable
            If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i), "*", Wow6432Redir) = 2 Then
                'entire domain is trusted
                sHit = "O15 - Trusted Zone: *." & sDomains(i) & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
            End If
            If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapDomains & "\" & sDomains(i), "http", Wow6432Redir) = 2 Then
                'only http on domain is trusted
                sHit = "O15 - Trusted Zone: http://*." & sDomains(i) & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
            End If
NextDomain2:
        Next i
    End If
  Next
  
    'enum all IP ranges
    sDomains = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sZoneMapRanges), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            sIPRange = RegGetString(HKEY_CURRENT_USER, sZoneMapRanges & "\" & sDomains(i), ":Range")
            If Left$(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                If RegGetDword(HKEY_CURRENT_USER, sZoneMapRanges & "\" & sDomains(i), "*") = 2 Then
                    'all protocols for this ip range is trusted
                    sHit = "O15 - Trusted IP range: " & sIPRange
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                End If
                If RegGetDword(HKEY_CURRENT_USER, sZoneMapRanges & "\" & sDomains(i), "http") = 2 Then
                    'only http protocol for this ip range is trusted
                    sHit = "O15 - Trusted IP range: http://" & sIPRange
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                End If
            End If
        Next i
    End If

  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    'repeat for HKLM (ip ranges)
    sDomains = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sZoneMapRanges, Wow6432Redir), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            sIPRange = RegGetString(HKEY_LOCAL_MACHINE, sZoneMapRanges & "\" & sDomains(i), ":Range", Wow6432Redir)
            If Left$(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapRanges & "\" & sDomains(i), "*", Wow6432Redir) = 2 Then
                    'all protocols for this ip range is trusted
                    sHit = "O15 - Trusted IP range: " & sIPRange & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                End If
                If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapRanges & "\" & sDomains(i), "http", Wow6432Redir) = 2 Then
                    'only http protocol for this ip range is trusted
                    sHit = "O15 - Trusted IP range: http://" & sIPRange & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                End If
            End If
        Next i
    End If
  Next
    
'======================= REPEAT FOR ESC =======================
    'enum all subkeys (i.e. all domains)
    sDomains = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sZoneMapEscDomains), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            If bIgnoreSafe Then
                For J = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sSafeRegDomains(J), CStr(sDomains(i)), vbTextCompare) > 0 Then
                        If InStr(sDomains(i), "msn.com") = 0 Then
                            'it's a safe domain!
                            GoTo NextEscDomain
                        Else
                            Exit For
                        End If
                    End If
                Next J
            End If
            sSubDomains = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i)), "|")
            If UBound(sSubDomains) <> -1 Then
                'list any trusted subdomains for main domain
                For J = 0 To UBound(sSubDomains)
                    If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i) & "\" & sSubDomains(J), "*") = 2 Then
                        'entire subdomain is trusted
                        sHit = "O15 - ESC Trusted Zone: " & sSubDomains(J) & "." & sDomains(i)
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                    End If
                    If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i) & "\" & sSubDomains(J), "http") = 2 Then
                        'only http on subdomain is trusted
                        sHit = "O15 - ESC Trusted Zone: http://" & sSubDomains(J) & "." & sDomains(i)
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                    End If
                Next J
            End If
            'list main domain as well if that's trusted too (*grumble*)
            If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i), "*") = 2 Then
                'entire domain is trusted
                sHit = "O15 - ESC Trusted Zone: *." & sDomains(i)
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
            End If
            If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscDomains & "\" & sDomains(i), "http") = 2 Then
                'only http on domain is trusted
                sHit = "O15 - ESC Trusted Zone: http://*." & sDomains(i)
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
            End If
NextEscDomain:
        Next i
    End If
    
  'repeat for HKLM (domains)
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    sDomains = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sZoneMapEscDomains, Wow6432Redir), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            If bIgnoreSafe Then
                For J = 0 To UBound(sSafeRegDomains)
                    If InStr(1, sSafeRegDomains(J), CStr(sDomains(i)), vbTextCompare) > 0 Then
                        If InStr(sDomains(i), "msn.com") = 0 Then
                            'it's a safe domain!
                            GoTo NextEscDomain2
                        Else
                            Exit For
                        End If
                    End If
                Next J
            End If
            sSubDomains = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i), Wow6432Redir), "|")
            If UBound(sSubDomains) <> -1 Then
                'list any trusted subdomains for main domain
                For J = 0 To UBound(sSubDomains)
                    If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i) & "\" & sSubDomains(J), "*", Wow6432Redir) = 2 Then
                        'entire subdomain is trusted
                        sHit = "O15 - ESC Trusted Zone: " & sSubDomains(J) & "." & sDomains(i) & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                    End If
                    If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i) & "\" & sSubDomains(J), "http", Wow6432Redir) = 2 Then
                        'only http on subdomain is trusted
                        sHit = "O15 - ESC Trusted Zone: http://" & sSubDomains(J) & "." & sDomains(i) & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                    End If
                Next J
            End If
            'list main domain as well, if applicable
            If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i), "*", Wow6432Redir) = 2 Then
                'entire domain is trusted
                sHit = "O15 - ESC Trusted Zone: *." & sDomains(i) & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
            End If
            If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscDomains & "\" & sDomains(i), "http", Wow6432Redir) = 2 Then
                'only http on domain is trusted
                sHit = "O15 - ESC Trusted Zone: http://*." & sDomains(i) & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
            End If
NextEscDomain2:
        Next i
    End If
  Next
    
    'enum all IP ranges
    sDomains = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sZoneMapEscRanges), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            sIPRange = RegGetString(HKEY_CURRENT_USER, sZoneMapEscRanges & "\" & sDomains(i), ":Range")
            If Left$(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscRanges & "\" & sDomains(i), "*") = 2 Then
                    'all protocols for this ip range is trusted
                    sHit = "O15 - ESC Trusted IP range: " & sIPRange
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                End If
                If RegGetDword(HKEY_CURRENT_USER, sZoneMapEscRanges & "\" & sDomains(i), "http") = 2 Then
                    'only http protocol for this ip range is trusted
                    sHit = "O15 - ESC Trusted IP range: http://" & sIPRange
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                End If
            End If
        Next i
    End If
    
  'repeat for HKLM (ip ranges)
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
      
    sDomains = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sZoneMapEscRanges, Wow6432Redir), "|")
    If UBound(sDomains) > -1 Then
        For i = 0 To UBound(sDomains)
            sIPRange = RegGetString(HKEY_LOCAL_MACHINE, sZoneMapEscRanges & "\" & sDomains(i), ":Range", Wow6432Redir)
            If Left$(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscRanges & "\" & sDomains(i), "*", Wow6432Redir) = 2 Then
                    'all protocols for this ip range is trusted
                    sHit = "O15 - ESC Trusted IP range: " & sIPRange & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                End If
                If RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapEscRanges & "\" & sDomains(i), "http", Wow6432Redir) = 2 Then
                    'only http protocol for this ip range is trusted
                    sHit = "O15 - ESC Trusted IP range: http://" & sIPRange & " (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
                End If
            End If
        Next i
    End If
  Next
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
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
        End If
    Next i
    
ZoneMapProtDefsHKLM:
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    For i = 0 To 5
        lProtZones(i) = RegGetDword(HKEY_LOCAL_MACHINE, sZoneMapProtDefs, sProtVals(i), Wow6432Redir)
        If lProtZones(i) < 0 Or lProtZones(i) > 5 Then lProtZones(i) = 5
        If lProtZones(i) <> lProtZoneDefs(i) Then
            sHit = "O15 - ProtocolDefaults: '" & sProtVals(i) & "' protocol is in " & sZoneNames(lProtZones(i)) & " Zone, should be " & sZoneNames(lProtZoneDefs(i)) & " Zone (HKLM)" & IIf(bIsWin64 And Wow6432Redir, "(32)", "")
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O15", sHit
        End If
    Next i
  Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO15Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO15Item(sItem$)
    'O15 - Trusted Zone: free.aol.com (HKLM)
    'O15 - Trusted Zone: http://free.aol.com
    'O15 - Trusted IP range: 66.66.66.66 (HKLM)
    'O15 - Trusted IP range: http://66.66.66.*
    'O15 - ESC Trusted Zone: free.aol.com (HKLM)
    'O15 - ESC Trusted IP range: 66.66.66.66
    'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)
    '* other domains are now listed since 1.95.1 *
    '* retarded hijackers use wrong format for trusted sites - 1.99.2 *
    
    On Error GoTo ErrorHandler:
    Dim lHive&, sKey1$, sKey2$, sKey3$, sValue$
    Dim sZoneMapDomains$, sZoneMapRanges$, sZoneMapProtDefs$
    Dim sZoneMapEscDomains$, sZoneMapEscRanges$
    Dim i&, sDummy$, vRanges As Variant, Wow6432Redir As Boolean
    
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
    
    If InStr(sItem, "(HKLM)(32)") <> 0 Then Wow6432Redir = True
    
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
    sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
    If InStr(sDummy, " (HKLM)") > 0 Then sDummy = Left$(sDummy, InStr(sDummy, " (HKLM)") - 1)
    'strip protocol (if any) from domain
    If InStr(sDummy, "//") > 0 Then sDummy = Mid$(sDummy, InStr(sDummy, "//") + 2)
    If InStr(sDummy, "*.") > 0 Then
        sDummy = Mid$(sDummy, InStr(sDummy, "*.") + 2)
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
        sKey2 = Mid$(sDummy, i + 1)
        sKey1 = sKey2 & "\" & Left$(sDummy, i - 1)
        sKey3 = Mid$(sDummy, 3)
    End If
    
    'relevant value should be deleted, and if no
    'other value is present, subkey as well.
    'if main key has no other subkeys, delete that also.
    If InStr(sItem, "ESC Trusted") = 0 Then
        If sKey1 = vbNullString Then
            RegDelVal lHive, sZoneMapDomains & sKey2, sValue, Wow6432Redir
            If Not RegKeyHasValues(lHive, sZoneMapDomains & sKey2, Wow6432Redir) Then
                RegDelKey lHive, sZoneMapDomains & sKey2, Wow6432Redir
            End If
        Else
            RegDelVal lHive, sZoneMapDomains & sKey1, sValue, Wow6432Redir
            If Not RegKeyHasValues(lHive, sZoneMapDomains & sKey1, Wow6432Redir) Then
                RegDelKey lHive, sZoneMapDomains & sKey1, Wow6432Redir
                If Not RegKeyHasSubKeys(lHive, sZoneMapDomains & sKey2, Wow6432Redir) And _
                   Not RegKeyHasValues(lHive, sZoneMapDomains & sKey2, Wow6432Redir) Then
                    RegDelKey lHive, sZoneMapDomains & sKey2, Wow6432Redir
                End If
            End If
            '1.99.2 - fix for retarded hijackers like *.frame.crazywinnings.com
            RegDelVal lHive, sZoneMapDomains & sKey3, sValue, Wow6432Redir
            If Not RegKeyHasValues(lHive, sZoneMapDomains & sKey3, Wow6432Redir) Then
                RegDelKey lHive, sZoneMapDomains & sKey3, Wow6432Redir
            End If
        End If
    Else '1.99.2: added EscDomains
        If sKey1 = vbNullString Then
            RegDelVal lHive, sZoneMapEscDomains & sKey2, sValue, Wow6432Redir
            If Not RegKeyHasValues(lHive, sZoneMapEscDomains & sKey2, Wow6432Redir) Then
                RegDelKey lHive, sZoneMapEscDomains & sKey2, Wow6432Redir
            End If
        Else
            RegDelVal lHive, sZoneMapEscDomains & sKey1, sValue, Wow6432Redir
            If Not RegKeyHasValues(lHive, sZoneMapEscDomains & sKey1, Wow6432Redir) Then
                RegDelKey lHive, sZoneMapEscDomains & sKey1, Wow6432Redir
                If Not RegKeyHasSubKeys(lHive, sZoneMapEscDomains & sKey2, Wow6432Redir) And _
                   Not RegKeyHasValues(lHive, sZoneMapEscDomains & sKey2, Wow6432Redir) Then
                    RegDelKey lHive, sZoneMapEscDomains & sKey2, Wow6432Redir
                End If
            End If
            '1.99.2 - fix for retarded hijackers like *.frame.crazywinnings.com
            RegDelVal lHive, sZoneMapEscDomains & sKey3, sValue, Wow6432Redir
            If Not RegKeyHasValues(lHive, sZoneMapEscDomains & sKey3, Wow6432Redir) Then
                RegDelKey lHive, sZoneMapEscDomains & sKey3, Wow6432Redir
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
    sDummy = Mid$(sItem, InStr(sItem, ":") + 2)
    If InStr(sDummy, " (HKLM)") > 0 Then sDummy = Left$(sDummy, InStr(sDummy, " (HKLM)") - 1)
    If InStr(sDummy, "//") > 0 Then sDummy = Mid$(sDummy, InStr(sDummy, "//") + 2)
    sKey2 = sDummy
    If InStr(sItem, "ESC Trusted") = 0 Then
        vRanges = Split(RegEnumSubKeys(lHive, sZoneMapRanges, Wow6432Redir), "|")
        If UBound(vRanges) <> -1 Then
            For i = 0 To UBound(vRanges)
                sKey1 = RegGetString(lHive, sZoneMapRanges & "\" & vRanges(i), ":Range", Wow6432Redir)
                If InStr(sKey1, sKey2) = 1 Then
                    RegDelKey lHive, sZoneMapRanges & "\" & vRanges(i), Wow6432Redir
                    Exit For
                End If
            Next i
        End If
    Else
        vRanges = Split(RegEnumSubKeys(lHive, sZoneMapEscRanges, Wow6432Redir), "|")
        If UBound(vRanges) <> -1 Then
            For i = 0 To UBound(vRanges)
                sKey1 = RegGetString(lHive, sZoneMapEscRanges & "\" & vRanges(i), ":Range", Wow6432Redir)
                If InStr(sKey1, sKey2) = 1 Then
                    RegDelKey lHive, sZoneMapEscRanges & "\" & vRanges(i), Wow6432Redir
                    Exit For
                End If
            Next i
        End If
    End If
    Exit Sub
    
ProtDefs:
    'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)
    sDummy = Mid$(sItem, InStr(sItem, ": ") + 3)
    sDummy = Left$(sDummy, InStr(sDummy, "'") - 1)

    Select Case sDummy
        Case "@ivt": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 1, Wow6432Redir
        Case "file": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 3, Wow6432Redir
        Case "ftp": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 3, Wow6432Redir
        Case "http": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 3, Wow6432Redir
        Case "https": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 3, Wow6432Redir
        Case "shell": RegSetDwordVal lHive, sZoneMapProtDefs, sDummy, 0, Wow6432Redir
    End Select
    
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_FixO15Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

'Public Sub CheckNetscapeMozilla()
'    Dim sDummy$, sNSVer$, sMailKey$, sPrefsJs$, sUserName$, ff%
'    On Error GoTo ErrorHandler:
'
'    sUserName = GetUser
'
'    If RegKeyExists(HKEY_CURRENT_USER, "Software\Netscape\Netscape Navigator\Main") Then
'        'netscape 4.x is installed
'
'        'get "popstatePath" - only way to find Users folder
'        'I really hate Netscape
'        sMailKey = "Software\Netscape\Netscape Navigator\biff\users"
'        sDummy = RegGetFirstSubKey(HKEY_CURRENT_USER, sMailKey)
'        sMailKey = sMailKey & "\" & sDummy & "\servers"
'        sDummy = RegGetFirstSubKey(HKEY_CURRENT_USER, sMailKey)
'        sMailKey = sMailKey & "\" & sDummy
'        sDummy = RegGetString(HKEY_CURRENT_USER, sMailKey, "popstatePath")
'        If sDummy <> vbNullString Then
'            'cut off \mail\popstate.dat
'            sDummy = Left$(sDummy, InStrRev(sDummy, "\") - 6)
'            sPrefsJs = sDummy & "\prefs.js"
'            If FileExists(sPrefsJs) Then
'                If FileLenW(sPrefsJs) > 0 Then
'                    ff = FreeFile()
'                    Open sPrefsJs For Input As #ff
'                        Do
'                            Line Input #ff, sDummy
'                            If InStr(sDummy, "user_pref(""browser.startup.homepage"",") > 0 Then
'                                frmMain.lstResults.AddItem "N1 - Netscape 4: " & sDummy & " (" & sPrefsJs & ")"
'                                Exit Do
'                            End If
'                        Loop Until EOF(ff)
'                    Close #ff
'                End If
'            End If
'        End If
'    End If
'
'    sDummy = vbNullString
'    'moz/ns6/ns7 all use similar regkeys
'    'moz uses \mozilla\currentversion or \seamonkey\currentversion
'    'ns6 uses \netscape\netscape 6\currentversion
'    'ns7 uses \netscape\currentversion or \netscape\netscape 6\currentversion
'    'they all use the same place to store PREFS.JS though
'    sDummy = RegGetString(HKEY_CURRENT_USER, "Software\Mozilla\Mozilla Firefox", "CurrentVersion")
'    If sDummy = vbNullString Then sDummy = RegGetString(HKEY_CURRENT_USER, "Software\Netscape\Netscape 6", "CurrentVersion")
'    If sDummy = vbNullString Then sDummy = RegGetString(HKEY_CURRENT_USER, "Software\Netscape\Netscape", "CurrentVersion")
'    If sDummy <> vbNullString Then
'        'mozilla, netscape 6 or netscape 7 is installed
'
'        'sDummy is something like "1.2b" [moz],
'        '"6.2.3 (en)" [ns6], or "7.0 (en)" [ns7]
'        If Left$(sDummy, 1) = "6" Then
'            sNSVer = "N2 - Netscape 6: "
'        ElseIf Left$(sDummy, 1) = "7" Then
'            sNSVer = "N3 - Netscape 7: "
'        Else
'            sNSVer = "N4 - Mozilla: "
'        End If
'
'        'prefs.js is stored in the insane location of
'        '%APPLICATIONDATA%\Mozilla\Profiles\default\
'        '     [random string].slt\prefs.js
'        '%APPLICDATA% also varies per Windows version
'        If Not bIsWinNT Then
'            sPrefsJs = sWinDir & "\Application Data"
'        Else
'            sPrefsJs = Left$(sWinDir, 2) & "\Documents and Settings\" & sUserName & "\Application Data"
'        End If
'        sPrefsJs = sPrefsJs & "\Mozilla\Profiles\default"
'        sDummy = GetFirstSubFolder(sPrefsJs)
'        sPrefsJs = sPrefsJs & "\" & sDummy & "\prefs.js"
'        If FileExists(sPrefsJs) Then
'            If FileLenW(sPrefsJs) > 0 Then
'                ff = FreeFile()
'                Open sPrefsJs For Input As #ff
'                    Do
'                        Line Input #ff, sDummy
'                        If InStr(sDummy, "user_pref(""browser.startup.homepage"",") > 0 Then
'                            frmMain.lstResults.AddItem sNSVer & sDummy & " (" & sPrefsJs & ")"
'                            Exit Do
'                        End If
'                    Loop Until EOF(ff)
'                Close #ff
'                ff = FreeFile()
'                Open sPrefsJs For Input As #ff
'                    Do
'                        Line Input #ff, sDummy
'                        If InStr(sDummy, "user_pref(""browser.search.defaultengine"",") > 0 Then
'                            frmMain.lstResults.AddItem sNSVer & sDummy & " (" & sPrefsJs & ")"
'                            Exit Do
'                        End If
'                    Loop Until EOF(ff)
'                Close #ff
'            End If
'        End If
'    End If
'    Exit Sub
'
'ErrorHandler:
'    Close #ff
'    ErrorMsg Err, "modMain_CheckNetscapeMozilla"
'    If inIDE Then Stop: Resume Next
'End Sub
'
'Public Sub FixNetscapeMozilla(sItem$)
'    'N1 - Netscape 4: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'    'N2 - Netscape 6: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'    'N3 - Netscape 7: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'    'N4 - Mozilla: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'    '               user_pref("browser.search.defaultengine", "http://url"); (c:\..\prefs.js)
'
'    Dim sPrefsJs$, sDummy$, ff1%, ff2%
'    On Error GoTo ErrorHandler:
'    sPrefsJs = Mid$(sItem, InStrRev(sItem, "(") + 1)
'    sPrefsJs = Left$(sPrefsJs, Len(sPrefsJs) - 1)
'    If FileExists(sPrefsJs) Then
'        ff1 = FreeFile()
'        Open sPrefsJs For Input As #ff1
'        ff2 = FreeFile()
'        Open sPrefsJs & ".new" For Output As #ff2
'            Do
'                Line Input #ff1, sDummy
'                If InStr(sDummy, "user_pref(""browser.startup.homepage"",") > 0 And _
'                   InStr(sItem, "user_pref(""browser.startup.homepage"",") > 0 Then
'                    Print #ff2, "user_pref(""browser.startup.homepage"", ""http://home.netscape.com/"");"
'                ElseIf InStr(sDummy, "user_pref(""browser.search.defaultengine"",") > 0 And _
'                   InStr(sItem, "user_pref(""browser.search.defaultengine"",") > 0 Then
'                    Print #ff2, "user_pref(""browser.search.defaultengine"", ""http://www.google.com/"");"
'                Else
'                    Print #ff2, sDummy
'                End If
'            Loop Until EOF(ff1)
'        Close #ff1
'        Close #ff2
'        deletefileWEx (StrPtr(sPrefsJs))
'        Name sPrefsJs & ".new" As sPrefsJs
'    End If
'    Exit Sub
'
'ErrorHandler:
'    Close #ff1
'    Close #ff2
'    ErrorMsg Err, "modMain_FixNetscapeMozilla", "sItem=", sItem
'    If inIDE Then Stop: Resume Next
'End Sub

Public Sub CheckO16Item()
    'O16 - Downloaded Program Files
    
    'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ActiveXCache
    'is location of actual %WINDIR%\DPF\ folder
    
    On Error GoTo ErrorHandler:
    Dim sDPFKey$, sName$, sFriendlyName$, sCodebase$, i&, hKey&, lpcName&, sHit$, UseWow, Wow6432Redir As Boolean, Result As TYPE_Scan_Results
    
    sDPFKey = "Software\Microsoft\Code Store Database\Distribution Units"
    
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr(sDPFKey), 0, KEY_ENUMERATE_SUB_KEYS Or (KEY_WOW64_64KEY And Not Wow6432Redir), hKey) = 0 Then
    
      sName = String$(MAX_KEYNAME, 0)
      lpcName = Len(sName)
      i = 0
    
      Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
        sName = Left$(sName, InStr(sName, vbNullChar) - 1)
        If Left$(sName, 1) = "{" And Right$(sName, 1) = "}" Then
            sFriendlyName = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sName, vbNullString, Wow6432Redir)
            If sFriendlyName = vbNullString Then
                sFriendlyName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName, vbNullString, Wow6432Redir)
            End If
        End If
        sCodebase = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\" & sName & "\DownloadInformation", "CODEBASE", Wow6432Redir)
        
        If (InStr(sCodebase, "http://www.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://webresponse.one.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://rtc.webresponse.one.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://office.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://officeupdate.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://protect.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://dql.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://codecs.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://download.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://windowsupdate.microsoft.com") <> 1 And _
           InStr(sCodebase, "http://v4.windowsupdate.microsoft.com") <> 1) _
           Or bIgnoreAllWhitelists Then
           
           'InStr(sCodeBase, "http://java.sun.com") <> 1 And _
           'InStr(sCodeBase, "http://download.macromedia.com") <> 1 And _
           'InStr(sCodeBase, "http://fpdownload.macromedia.com") <> 1 And _
           'InStr(sCodeBase, "http://active.macromedia.com") <> 1 And _
           'InStr(sCodeBase, "http://www.apple.com") <> 1 And _
           'InStr(sCodeBase, "http://http://security.symantec.com") <> 1 And _
           'InStr(sCodeBase, "http://download.yahoo.com") <> 1 And _
           'InStr(sName, "Microsoft XML Parser") = 0 And _
           'InStr(sName, "Java Classes") = 0 And _
           'InStr(sName, "Classes for Java") = 0 And _
           'InStr(sName, "Java Runtime Environment") = 0 Or _

           ' "O16 - DPF: "
            sHit = IIf(bIsWin32, "O16", IIf(Wow6432Redir, "O16-32", "O16")) & " - DPF: " & sName & IIf(sFriendlyName <> vbNullString, " (" & sFriendlyName & ")", vbNullString) & " - " & sCodebase
            If Not IsOnIgnoreList(sHit) Then
                With Result
                    .Section = "O16"
                    .HitLineW = sHit
                    .CLSID = sName
                    .Redirected = Wow6432Redir
                End With
                AddToScanResults Result
            End If
        End If
        
        i = i + 1
        sName = String$(MAX_KEYNAME, 0)
        lpcName = Len(sName)
        
      Loop
      RegCloseKey hKey
    End If
  Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO16Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO16Item(sItem$)
    'O16 - DPF: {0000000} (shit toolbar) - http://bla.com/bla.dll
    'O16 - DPF: Plugin - http://bla.com/bla.dll
    
    On Error GoTo ErrorHandler:
    Dim sDPFKey$, hKey&, sDummy$, sName$, sOSD$, sInf$, sInProcServer32$, Wow6432Redir As Boolean
    
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub

    With Result
        sName = .CLSID
        Wow6432Redir = .Redirected
    End With

    sDPFKey = "Software\Microsoft\Code Store Database\Distribution Units\" & sName
    
    If Not RegKeyExists(HKEY_LOCAL_MACHINE, sDPFKey, Wow6432Redir) Then
        'unable to find that key
        'msgboxW "Could not delete '" & sItem & "' because it doesn't exist anymore.", vbExclamation
        Exit Sub
    End If
    
    'a DPF object can consist of:
    '* DPF regkey           -> sDPFKey
    '* CLSID regkey         -> CLSID\ & sName
    '* OSD file             -> sOSD = RegGetString
    '* INF file             -> sINF = RegGetString
    '* InProcServer32 file  -> sIPS = RegGetString
    
    sOSD = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\DownloadInformation", "OSD", Wow6432Redir)
    sInf = RegGetString(HKEY_LOCAL_MACHINE, sDPFKey & "\DownloadInformation", "INF", Wow6432Redir)
    If Left$(sName, 1) = "{" Then
        sInProcServer32 = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName & "\InProcServer32", vbNullString, Wow6432Redir)
        If sInProcServer32 = "" Then
            sInProcServer32 = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName & "\InProcServer32", vbNullString, Not Wow6432Redir)
        End If
        If sInProcServer32 <> "" Then
            If FileExists(sInProcServer32) Then
                Shell "regsvr32.exe /u /s """ & sInProcServer32 & """", vbHide
                DoEvents
            End If
        End If
    End If
    
    RegDelKey HKEY_LOCAL_MACHINE, sDPFKey, Wow6432Redir
    If Left$(sName, 1) = "{" Then
        RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & sName, True
        RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & sName, False
    End If

    DeleteFileWEx (StrPtr(sInProcServer32))
    DeleteFileWEx (StrPtr(sOSD))
    DeleteFileWEx (StrPtr(sInf))
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_FixO16Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO17Item()
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
    
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, J&, sName$, sDomain$, sHit$, sParam$, Param, CSKey$, n&, sData$, aNames() As String
    Dim UseWow, Wow6432Redir As Boolean, Result As TYPE_Scan_Results, Data() As String
    Dim TcpIpNameServers() As String: ReDim TcpIpNameServers(0)
    ReDim sKeyDomain(0 To 1) As String
    'these keys are x64 shared
    sKeyDomain(0) = "Services\Tcpip\Parameters"
    sKeyDomain(1) = "Services\VxD\MSTCP"
    
    For J = 0 To 999    ' 0 - is CSS
    
        CSKey = IIf(J = 0, "System\CurrentControlSet", "System\ControlSet" & Format(J, "000"))
    
        For Each Param In Array("Domain", "DomainName", "SearchList", "NameServer")
            sParam = Param
            
            For n = 0 To UBound(sKeyDomain)
                'HKLM\System\CCS\Services\Tcpip\Parameters,Domain
                'HKLM\System\CCS\Services\Tcpip\Parameters,DomainName
                'HKLM\System\CCS\Services\VxD\MSTCP,Domain
                'HKLM\System\CCS\Services\VxD\MSTCP,DomainName
                'new one from UltimateSearch!
                'HKLM\System\CCS\Services\VxD\MSTCP,SearchList
                'HKLM\System\CCS\Services\VxD\MSTCP,SearchList
                'HKLM\System\CCS\Services\Tcpip\Parameters,SearchList
                'HKLM\System\CCS\Services\Tcpip\Parameters,NameServer
                sData = RegGetString(HKEY_LOCAL_MACHINE, CSKey & "\" & sKeyDomain(n), sParam)
                If sData <> vbNullString Then
                    sHit = "O17 - HKLM\" & IIf(J = 0, "System\CSS", CSKey) & "\" & sKeyDomain(n) & ": " & sParam & " = " & sData
                    If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O17", sHit
                End If
            Next
            
            'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\.. subkeys
            'HKLM\System\CS*\Services\Tcpip\Parameters\Interfaces\.. subkeys
            For n = 1 To RegEnumSubkeysToArray(HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces", aNames)
                
                sData = RegGetString(HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces\" & aNames(n), sParam)
                If sData <> vbNullString Then
                
                    ReDim Data(0)
                    Data(0) = sData
                
                    If sParam = "NameServer" Then
                        
                        'Split lines like:
                        'O17 - HKLM\System\CCS\Services\Tcpip\..\{19B2C21E-CA09-48A1-9456-E4191BE91F00}: NameServer = 89.20.100.53 83.219.25.69
                        'O17 - HKLM\System\CCS\Services\Tcpip\..\{2A220B45-7A12-4A0B-92F0-00254794215A}: NameServer = 192.168.1.1,8.8.8.8
                        'into several separate
                        sData = Trim$(sData)
                        If InStr(sData, " ") <> 0 Then
                            Data = Split(sData)
                        ElseIf InStr(sData, ",") <> 0 Then
                            Data = Split(sData, ",")
                        End If
                    
                        For i = 0 To UBound(Data)
                            ReDim Preserve TcpIpNameServers(UBound(TcpIpNameServers) + 1)
                            TcpIpNameServers(UBound(TcpIpNameServers)) = Data(i)
                        Next
                    End If
                    
                    For i = 0 To UBound(Data)
                        sHit = "O17 - HKLM\" & IIf(J = 0, "System\CSS", CSKey) & "\Services\Tcpip\..\" & aNames(n) & ": " & sParam & " = " & Data(i)
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O17", sHit
                    Next
                End If
            Next
        Next
    Next
    
    Dim sTelephonyDomain$
    sTelephonyDomain = "Software\Microsoft\Windows\CurrentVersion\Telephony"
    
    For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If bIsWin32 And Wow6432Redir Then Exit For
    
        'HKLM\Software\MS\Windows\CurVer\Telephony,Domain
        'HKLM\Software\MS\Windows\CurVer\Telephony,DomainName
        For Each Param In Array("Domain", "DomainName")
            sParam = Param
            sDomain = RegGetString(HKEY_LOCAL_MACHINE, sTelephonyDomain, sParam, Wow6432Redir)
            If sDomain <> vbNullString Then
                'O17 - HKLM\Software\..\Telephony:
                sHit = IIf(bIsWin32, "O17", IIf(Wow6432Redir, "O17-32", "O17")) & " - HKLM\Software\..\Telephony: " & sParam & " = " & sDomain
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O17", sHit
            End If
        Next
    Next
    
    '------------------------------------------------------------
    
    Dim DNS() As String
    
    If GetDNS(DNS) Then
        For i = 0 To UBound(DNS)
            If Len(DNS(i)) <> 0 Then
                'If Not (DNS(i) = "192.168.0.1" Or DNS(i) = "192.168.1.1") Then
                    If Not inArray(DNS(i), TcpIpNameServers, , , vbTextCompare) Then
                        sHit = "O17 - DHCP DNS - " & i + 1 & ": " & DNS(i)
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O17", sHit
                    End If
                'End If
            End If
        Next
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO17Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO17Item(sItem$)
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
    
    On Error GoTo ErrorHandler:
    Dim sKey$, sValue$, sDummy$, i&, J&, Wow6432Redir As Boolean
    
    If StrBeginWith(sItem, "O17 - DHCP DNS:") Then
        'Cure for this object is not provided: []
        'You need to manually set the DNS address on the router, which is issued to you by provider.
        MsgBoxW Replace$(TranslateNative(349), "[]", sItem), vbExclamation
        FlushDNS
        Exit Sub
    End If
    
    If StrBeginWith(sItem, "O17") Then Wow6432Redir = True
    
    sDummy = Mid$(sItem, InStr(sItem, " - ") + 3)
    sKey = Left$(sDummy, InStr(sDummy, ":") - 1)
    If InStr(sKey, "\..\") > 0 Then
        'expand \..\
        If InStr(sKey, "Telephony") > 0 Then
            sKey = Replace$(sKey, "\..\", "\Microsoft\Windows\CurrentVersion\", , 1)
        End If
        If InStr(sKey, "Tcpip") > 0 Then
            sKey = Replace$(sKey, "\..\", "\Parameters\Interfaces\", , 1)
        End If
    End If
    If InStr(sKey, "\CCS\") > 0 Then
        sKey = Replace$(sKey, "\CCS\", "\CurrentControlSet\", , 1)
    End If
    
    'expand CCS/CS1/CS2/..
    i = InStr(sKey, "\CS")
    If i > 0 And i < 20 Then
        '<20 just in case a domain with \cs comes up
        '\CS1\   or   \CS11\
        J = InStr(i + 3, sKey, "\") - i - 3
        sKey = Replace$(sKey, "\CS", "\ControlSet" & String$(3 - J, "0"), , 1)
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
    sKey = Mid$(sKey, 6)
    RegDelVal HKEY_LOCAL_MACHINE, sKey, sValue, Wow6432Redir
    FlushDNS
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_FixO17Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO18Item()
    'enumerate everything in HKCR\Protocols\Handler
    'enumerate everything in HKCR\Protocols\Filters (section 2)
    'keys are x64 shared
    
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, sName$, sCLSID$, sFile$, lpcName&, sHit$, Wow6432Redir As Boolean
    
    Wow6432Redir = False
    
    If RegOpenKeyExW(HKEY_CLASSES_ROOT, StrPtr("Protocols\Handler"), 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
      sName = String$(MAX_KEYNAME, 0&)
      lpcName = Len(sName)
      i = 0
      Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
        sName = TrimNull(sName)
        sCLSID = UCase$(RegGetString(HKEY_CLASSES_ROOT, "Protocols\Handler\" & sName, "CLSID", Wow6432Redir))
        If sCLSID = vbNullString Then sCLSID = "(no CLSID)"
        If sCLSID <> "(no CLSID)" Then
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Wow6432Redir)
            sFile = Replace$(sFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
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
                 AddToScanResultsSimple "O18", sHit
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
                         AddToScanResultsSimple "O18", sHit
                     End If
                End If
            End If
        End If
        
        sName = String$(MAX_KEYNAME, 0)
        lpcName = Len(sName)
        i = i + 1
      Loop
      RegCloseKey hKey
    End If
    
    '-------------------
    'Filters:
    
    hKey = 0
    sCLSID = vbNullString
    sFile = vbNullString
    
    If RegOpenKeyExW(HKEY_CLASSES_ROOT, StrPtr("PROTOCOLS\Filter"), 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
      sName = String$(MAX_KEYNAME, 0&)
      lpcName = Len(sName)
      i = 0
      Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
        sName = TrimNull(sName)
        sCLSID = RegGetString(HKEY_CLASSES_ROOT, "PROTOCOLS\Filter\" & sName, "CLSID", Wow6432Redir)
        If sCLSID = vbNullString Then
            sCLSID = "(no CLSID)"
            sFile = "(no file)"
        Else
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Wow6432Redir)
            If sFile = vbNullString Then
                sFile = "(no file)"
            Else
                sFile = EnvironW(sFile)
            End If
        End If
        
        If InStr(1, Join(sSafeFilters, vbCrLf), sName, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
            'add to results list
            sHit = "O18 - Filter: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                AddToScanResultsSimple "O18", sHit
            End If
        Else
            If InStr(1, Join(sSafeFilters, vbCrLf), sCLSID, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                'add to results list
                sHit = "O18 - Filter hijack: " & sName & " - " & sCLSID & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    AddToScanResultsSimple "O18", sHit
                End If
            End If
        End If
        
        sName = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sName)
        i = i + 1
      Loop
      RegCloseKey hKey
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_CheckO18Item"
    RegCloseKey hKey
End Sub

Public Sub FixO18Item(sItem$)
    'O18 - Protocol: cn
    
    Dim sDummy$, i&, sCLSID$ ', sProtCLSIDs$()
    On Error GoTo ErrorHandler:
    If InStr(sItem, "Filter: ") > 0 Then GoTo FixFilter:
       
    'get protocol name
    sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
    sDummy = Left$(sDummy, InStr(sDummy, " - ") - 1)
    
    If InStr(sItem, "Protocol hijack: ") > 0 Then GoTo FixProtHijack:
    
    If InStr(sItem, "(no CLSID)") = 0 Then
        'RegDelSubKeys HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy
        RegDelKey HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy
    End If
    
    Exit Sub
    
FixProtHijack:
    For i = 0 To UBound(sSafeProtocols)
        'find CLSID for protocol name
        If sSafeProtocols(i) = vbNullString Then Exit For
        If InStr(1, sSafeProtocols(i), sDummy) > 0 Then
            sCLSID = Mid$(sSafeProtocols(i), InStr(sSafeProtocols(i), "|") + 1)
            Exit For
        End If
    Next i
    RegSetStringVal HKEY_CLASSES_ROOT, "Protocols\Handler\" & sDummy, "CLSID", sCLSID
    
    Exit Sub
    
FixFilter:
    'O18 - Filter: text/blah - {0} - c:\file.dll
    sDummy = Mid$(sItem, InStr(sItem, ": ") + 2)
    'why the hell did I use InstrRev here first? bugfix 1.98.1
    sDummy = Left$(sDummy, InStr(sDummy, " - ") - 1)
    
    If InStr(sItem, "Filter hijack: ") > 0 Then GoTo FixFilterHijack:
    
    RegDelKey HKEY_CLASSES_ROOT, "PROTOCOLS\Filter\" & sDummy
    Exit Sub
    
FixFilterHijack:
    For i = 0 To UBound(sSafeFilters)
        If sSafeFilters(i) = vbNullString Then Exit For
        If InStr(1, sSafeFilters(i), sDummy) > 0 Then
            sCLSID = Mid$(sSafeFilters(i), InStr(sSafeFilters(i), "|") + 1)
            Exit For
        End If
    Next i
    RegSetStringVal HKEY_CLASSES_ROOT, "PROTOCOLS\Filter\" & sDummy, "CLSID", sCLSID
    Exit Sub

ErrorHandler:
    ErrorMsg err, "modMain_FixO18Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO19Item()
    On Error GoTo ErrorHandler:
    'HKCU\Software\Microsoft\Internet Explorer\Styles,Use My Stylesheet
    'HKCU\Software\Microsoft\Internet Explorer\Styles,User Stylesheet
    'this hijack doesn't work for HKLM
    
    Dim lUseMySS&, sUserSS$, sHit$, UseWow, Wow6432Redir As Boolean, Result As TYPE_Scan_Results, vHive, lHive&
    
  For Each vHive In Array(HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER)
  For Each UseWow In Array(False, True)
    lHive = vHive
    Wow6432Redir = UseWow
    If Wow6432Redir And (bIsWin32 Or lHive = HKEY_CURRENT_USER) Then Exit For
    
    lUseMySS = RegGetDword(lHive, "Software\Microsoft\Internet Explorer\Styles", "Use My Stylesheet", Wow6432Redir)
    sUserSS = RegGetString(lHive, "Software\Microsoft\Internet Explorer\Styles", "User Stylesheet", Wow6432Redir)
    If Not FileExists(sUserSS) Then sUserSS = sUserSS & " (file missing)"
    If lUseMySS = 1 And sUserSS <> vbNullString Then
        'O19 - User stylesheet:
        'O19-32 - User stylesheet:
        sHit = IIf(bIsWin32 Or lHive = HKEY_CURRENT_USER, "O19", IIf(Wow6432Redir, "O19-32", "O19")) & _
          " - User stylesheet: " & sUserSS & IIf(lHive = HKEY_LOCAL_MACHINE, " (HKLM)", "")
        If Not IsOnIgnoreList(sHit) Then
            'md5 doesn't seem useful here
            'If bMD5 Then sHit = sHit & getfilemd5(sUserSS)
            With Result
                .Section = "O19"
                .HitLineW = sHit
                .lHive = lHive
                .Redirected = Wow6432Redir
            End With
            AddToScanResults Result
        End If
    End If
  Next
  Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO19Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO19Item(sItem$)
    On Error GoTo ErrorHandler:
    'O19 - User stylesheet: c:\file.css (file missing)
    
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    With Result
        RegDelVal .lHive, "Software\Microsoft\Internet Explorer\Styles", "Use My Stylesheet", .Redirected
        RegDelVal .lHive, "Software\Microsoft\Internet Explorer\Styles", "User Stylesheet", .Redirected
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO19Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO20Item()
    'AppInit_DLLs - https://support.microsoft.com/ru-ru/kb/197571
    
    'modules are delimited by spaces or commas
    'long file names are not permitted
    
    On Error GoTo ErrorHandler:
    'appinit_dlls + winlogon notify
    Dim sAppInit$, sFile$, sHit$, UseWow, Wow6432Redir As Boolean, Result As TYPE_Scan_Results
    
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    sAppInit = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
    sFile = RegGetString(HKEY_LOCAL_MACHINE, sAppInit, "AppInit_DLLs", Wow6432Redir)
    If sFile <> vbNullString Then
        sFile = Replace$(sFile, vbNullChar, "|")                        '// TODO: !!!
        If InStr(1, sSafeAppInit, sFile, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
            'item is not on whitelist
            'O20 - AppInit_DLLs
            'O20-32 - AppInit_DLLs
            sHit = IIf(bIsWin32, "O20", IIf(Wow6432Redir, "O20-32", "O20")) & " - AppInit_DLLs: " & sFile
            
            If Not IsOnIgnoreList(sHit) Or bIgnoreAllWhitelists Then
                With Result
                    .Section = "O20"
                    .HitLineW = sHit
                    ReDim .RegKey(0)
                    .RegKey(0) = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Windows"
                    .RegParam = "AppInit_DLLs"
                    .Redirected = Wow6432Redir
                End With
                AddToScanResults Result     'Action -> Clear Data in 'AppInit_DLLs'
            End If
        End If
    End If
    
    Dim sSubkeys$(), i&, sWinLogon$, SS$
    sWinLogon = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify"
    sSubkeys = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sWinLogon, Wow6432Redir), "|")
    If UBound(sSubkeys) <> -1 Then
        For i = 0 To UBound(sSubkeys)
            If InStr(1, "*" & sSafeWinlogonNotify & "*", "*" & sSubkeys(i) & "*", vbTextCompare) = 0 Then
                sFile = RegGetString(HKEY_LOCAL_MACHINE, sWinLogon & "\" & sSubkeys(i), "DllName", Wow6432Redir)
                
                If Len(sFile) = 0 Then
                    sFile = "Invalid registry found"
                Else
                    If Left$(sFile, 1) = "\" Then
                        If FileExists(sWinSysDir & "\" & sFile) Then
                            sFile = sWinSysDir & "\" & sFile
                        ElseIf FileExists(sWinDir & "\" & sFile) Then
                            sFile = sWinDir & "\" & sFile
                        End If
                    End If
                    
                    sFile = GetLongPath(EnvironW(sFile))
                    
                    If Not FileExists(sFile) Then
                        sFile = sFile & " (file missing)"
                    ElseIf bMD5 Then
                        sFile = sFile & GetFileFromAutostart(sFile)
                    End If
                    
'                    If InStr(1, sFile, "%", vbTextCompare) = 1 Then
'                       sFile = "Suspicious registry value"
'                    End If
                  
                End If
                'O20 - Winlogon Notify:
                'O20-32 - Winlogon Notify:
                sHit = IIf(bIsWin32, "O20", IIf(Wow6432Redir, "O20-32", "O20")) & " - Winlogon Notify: " & sSubkeys(i) & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O20"
                        .HitLineW = sHit
                        ReDim .RegKey(0)
                        .RegKey(0) = sSubkeys(i)
                        .Redirected = Wow6432Redir
                    End With
                    AddToScanResults Result     'Action -> Remove Key inside 'Notify'
                End If
            End If
        Next i
    End If
  Next

    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO20Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO20Item(sItem$)
    On Error GoTo ErrorHandler:
    
    'O20 - AppInit_DLLs: file.dll
    'O20 - Winlogon Notify: bladibla - c:\file.dll
    'to do:
    '* clear appinit regval (don't delete it)
    '* kill regkey (for winlogon notify)
    
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    Dim sAppInit$, sNotify$
    
    With Result
      If .RegParam = "AppInit_DLLs" Then
        sAppInit = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
        RegSetStringVal HKEY_LOCAL_MACHINE, sAppInit, "AppInit_DLLs", vbNullString, .Redirected
      ElseIf .RegParam = "" Then
        sNotify = .RegKey(0)
        RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify\" & sNotify, .Redirected
      End If
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO20Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO21Item()
    On Error GoTo ErrorHandler:
    
    'shellserviceobjectdelayload
    Dim sSSODL$, sHit$, sFile$, J&, bOnWhiteList As Boolean
    Dim hKey&, i&, sName$, lNameLen&, sCLSID$, lDataLen&
    Dim UseWow, Wow6432Redir As Boolean, Result As TYPE_Scan_Results

    sSSODL = "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
    
  For Each UseWow In Array(False, True)
    Wow6432Redir = UseWow
    If bIsWin32 And Wow6432Redir Then Exit For
    
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr(sSSODL), 0, KEY_QUERY_VALUE Or (KEY_WOW64_64KEY And Not Wow6432Redir), hKey) = 0 Then
    
      Do
        lNameLen = MAX_VALUENAME
        sName = String$(lNameLen, 0&)
        lDataLen = MAX_VALUENAME
        sCLSID = String$(lDataLen, 0&)
    
        If RegEnumValueW(hKey, i, StrPtr(sName), lNameLen, 0&, REG_SZ, StrPtr(sCLSID), lDataLen) <> 0 Then Exit Do
    
        sName = Left$(sName, lNameLen)
        sCLSID = TrimNull(sCLSID)
        If 0 = Len(sName) Then
            sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, Wow6432Redir)
            If sName = vbNullString Then sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString, Not Wow6432Redir)
            If sName = vbNullString Then sName = "(no name)"
        End If
        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Wow6432Redir)
        If sFile = vbNullString Then sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Not Wow6432Redir)
        If sFile = vbNullString Then
            sFile = "(no file)"
        Else
            sFile = EnvironW(sFile)
            If Not FileExists(sFile) Then
                sFile = sFile & " (file missing)"
            End If
        End If
        
        bOnWhiteList = inArray(sCLSID, sSafeSSODL, , , vbTextCompare)
        If bIgnoreAllWhitelists Then bOnWhiteList = False
        
        sHit = "O21 - SSODL: " & sName & " - " & sCLSID & " - " & sFile
        If Not IsOnIgnoreList(sHit) And Not bOnWhiteList Then
            If bMD5 Then sHit = sHit & GetFileMD5(sFile)
            With Result
                .Section = "O21"
                .HitLineW = sHit
                .CLSID = sCLSID
                ReDim .RegKey(0)
                .RegKey(0) = "HKLM\Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
                .RegParam = sName
                .Redirected = Wow6432Redir
            End With
            AddToScanResults Result
        End If
        
        i = i + 1
      Loop
      RegCloseKey hKey
    End If
  Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO21Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO21Item(sItem$)
    On Error GoTo ErrorHandler:
    
    'O21 - SSODL: webcheck - {000....000} - c:\file.dll (file missing)
    'actions to take:
    '* kill regval
    '* kill clsid regkey
    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub
    With Result
        RegDelVal 0&, .RegKey(0), .RegParam, .Redirected
        RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & .CLSID, True
        RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & .CLSID, False
    End With
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO21Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO22Item()
    On Error GoTo ErrorHandler:
    'ScheduledTask
    
    If OSver.bIsVistaOrLater Then
        EnumTasks   '<--- New routine
        Exit Sub
    End If
    
    '//TODO: Add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\
    'for Windows Vista and higher.
    
    'Win XP / Server 2003
    
    Dim sSTS$, hKey&, i&, sCLSID$, lCLSIDLen&, lDataLen&
    Dim sFile$, sName$, sHit$, isSafe As Boolean, WL_ID&
    Dim Wow6432Redir As Boolean
    
    Wow6432Redir = True
    
    sSTS = "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler"
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr(sSTS), 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        'regkey doesn't exist, or failed to open
        Exit Sub
    End If
    
    Do
        lCLSIDLen = MAX_VALUENAME
        sCLSID = String$(lCLSIDLen, 0&)
        lDataLen = MAX_VALUENAME
        sName = String$(lDataLen, 0&)
    
        If RegEnumValueW(hKey, i, StrPtr(sCLSID), lCLSIDLen, 0&, REG_SZ, StrPtr(sName), lDataLen) <> 0 Then Exit Do
    
        sCLSID = Left$(sCLSID, lCLSIDLen)
        sName = TrimNull(sName)
        If sName = vbNullString Then sName = "(no name)"
        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Wow6432Redir)
        sFile = Replace$(sFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
        If sFile = vbNullString Then
            sFile = "(no file)"
        Else
            If Not FileExists(sFile) Then
                sFile = sFile & " (file missing)"
            End If
        End If
        
        'whitelist
        isSafe = isInTasksWhiteList(sCLSID & "\" & sName, sFile, "")
        
        If Not isSafe Then
            sHit = "O22 - ScheduledTask: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                AddToScanResultsSimple "O22", sHit
            End If
        End If
        i = i + 1
    Loop
    RegCloseKey hKey
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_CheckO22Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO22Item(sItem$)
    On Error GoTo ErrorHandler:

    'O22 - ScheduledTask: blah - {000...000} - file.dll
    'todo:
    '* kill regval
    '* kill clsid regkey
    
    If OSver.bIsVistaOrLater Then
        Dim Result As TYPE_Scan_Results
        
        If Not GetScanResults(sItem, Result) Then Exit Sub
        
        KillTask Result.AutoRunObject
        Exit Sub
    End If
    
    Dim sCLSID$, sSTS$
    sSTS = "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler"
    
    sCLSID = Mid$(sItem, InStr(sItem, ": ") + 2)
    sCLSID = Mid$(sCLSID, InStr(sCLSID, " - ") + 3)
    sCLSID = Left$(sCLSID, InStr(sCLSID, " - ") - 1)
    
    RegDelVal HKEY_LOCAL_MACHINE, sSTS, sCLSID
    RegDelKey HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_FixO22Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO23Item()   '2.0.7 fixed
    On Error GoTo ErrorHandler:
    'enum NT services
    Dim sServices$(), i&, J&, sName$, sDisplayName$, ext$
    Dim lStart&, lType&, sFile$, sCompany$, sHit$, sBuf$, sTmp$, IsCompositeCmd As Boolean, sCompositeFile$, Arr
    Dim bHideDisabled As Boolean, NoFile As Boolean, Stady As Long
    Dim ServState As SERVICE_STATE
    
    If Not bIsWinNT Then Exit Sub
    
    bHideDisabled = True
    
    If bIgnoreAllWhitelists Then bHideMicrosoft = False
    
    Stady = 0
    
    sServices = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    Stady = 1
    If UBound(sServices) = -1 Then Exit Sub
    
    Stady = 2
    
    For i = 0 To UBound(sServices)
        sName = sServices(i)
        Stady = 3
        lStart = RegGetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start")
        Stady = 4
        lType = RegGetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Type")
        
        If lType < 16 Or (lStart = 4 And bHideDisabled) Then GoTo Continue
        
        Stady = 5
        sDisplayName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DisplayName")
        Stady = 6
        If Len(sDisplayName) = 0 Then
            sDisplayName = sName
        Else
            If Left$(sDisplayName, 1) = "@" Then                    'extract string resource from file
                Stady = 7
                sBuf = GetStringFromBinary(, , sDisplayName)
                Stady = 8
                If 0 <> Len(sBuf) Then sDisplayName = sBuf
            End If
        End If
        
        Stady = 9
        sFile = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ImagePath")
        Stady = 10
        
        'If sDisplayName = "DokanMounter" Then Stop
        
        'cleanup filename
        If Len(sFile) <> 0 Then
            Stady = 11
            'remove arguments e.g. ["c:\file.exe" -option]
            If Left$(sFile, 1) = """" Then
                J = InStr(2, sFile, """") - 2
                Stady = 12
                If J > 0 Then
                    sFile = Mid$(sFile, 2, J)
                Else
                    sFile = Mid$(sFile, 2)
                End If
            End If
            Stady = 13
            'expand aliases
            sFile = EnvironW(sFile)
            Stady = 14
            'sFile = replace$(sFile, "%systemroot%", sWinDir, , 1, vbTextCompare)
            sFile = Replace$(sFile, "\systemroot", sWinDir, , 1, vbTextCompare)
            sFile = Replace$(sFile, "systemroot", sWinDir, , 1, vbTextCompare)
            
            Stady = 15
            'prefix for windows folder if not specified?
            If StrComp("system32\", Left$(sFile, 9), 1) = 0 Then
                sFile = sWinDir & "\" & sFile
            End If
            
            Stady = 16
            'remove parameters (and double filenames)
            J = InStr(1, sFile, ".exe ", vbTextCompare) + 3 ' mark -> '.exe' + space
            If J < Len(sFile) And J > 3 Then sFile = Left$(sFile, J)
            
            Stady = 17
            'add .exe if not specified
            If Len(sFile) > 3 Then ext = Mid$(sFile, Len(sFile) - 3)
            If StrComp(ext, ".exe", 1) <> 0 And _
                StrComp(ext, ".sys", 1) <> 0 Then
                  sFile = sFile & ".exe"
            End If
            
            Stady = 18
            'wow64 correction
            If IsServiceWow64(sName) Then
                sFile = Replace$(sFile, sWinSysDir, sWinSysDirWow64, , , vbTextCompare)
            End If
            
            Stady = 19
            If Mid$(sFile, 2, 1) <> ":" Then 'if not fully qualified path
                If InStr(sFile, "\") = 0 Then
                    'sFile = sFile & String$(MAX_PATH - Len(sFile), vbNullChar)
                    'PathFindOnPath StrPtr(sFile), 0&
                    'sFile = TrimNull(sFile)
                    Stady = 20
                    sBuf = FindOnPath(sFile)
                    Stady = 21
                    If 0 <> Len(sBuf) Then sFile = sBuf
                End If
            End If
            
            'NoFile = False
            
'            If Not FileExists(sFile) Then
'                'sometimes the damn path isn't there AT ALL >_<
'                If Mid$(sFile, 2, 1) <> ":" Then
'                    If FileExists(Left$(sWinDir, 3) & sFile) Then
'                        sFile = Left$(sWinDir, 3) & sFile
'                    ElseIf FileExists(sWinDir & "\" & sFile) Then
'                        sFile = sWinDir & "\" & sFile
'                    ElseIf FileExists(sWinSysDir & "\" & sFile) Then
'                        sFile = sWinSysDir & "\" & sFile
'                    Else
'                        NoFile = True
'                    End If
'                Else
'                    NoFile = True
'                End If
'            End If
            
            'If NoFile Then
            
        End If
        
        '//// TODO: Check this !!!
        
        'https://technet.microsoft.com/en-us/library/cc959922.aspx
        'https://support.microsoft.com/en-us/kb/103000
        
        'Start
        '0 - Boot
        '1 - System
        '2 - Automatic
        '3 - Manual
        '4 - Disabled
        
        'Type
        '1 - Kernel device driver
        '2 - File System driver
        '4 - A set of arguments for an adapter
        '8 - File System driver service
        '16 - A Win32 program that runs in a process by itself. This type of Win32 service can be started by the service controller.
        '32 - A Win32 service that can share a process with other Win32 services
        '272 - A Win32 program that runs in a process by itself (like Type16) and that can interact with users.
        '288 - A Win32 program that shares a process and that can interact with users.
        
        ServState = GetServiceRunState(sName)
        
        If lType >= 16 Then
        If Not (lStart = 4 And bHideDisabled) Then  ' _
           'and Not (InStr(sCompany, "Microsoft") > 0 Or InStr(sCompany, "Корпорация Майкрософт") > 0 And bHideMicrosoft) Then
               
            Stady = 22
               
            Dim argc As Long
            Dim argv() As String
            Dim isSafeMSCmdLine As Boolean
            Dim SignResult As SignResult_TYPE
            Dim FoundFile As String
            Dim IsMSCert As Boolean
            
            IsCompositeCmd = False
            isSafeMSCmdLine = False
            
            If Not FileExists(sFile) And sFile <> "" Then
            
                ' Дальше идут процедуры парсинга командной строки и проверки сертиката для каждого файла из этой цепочки
                ' Если любой файл из цепочки не проходит проверку, строка считается небезопасной
            
                Stady = 23
            
                ParseCommandLine sFile, argc, argv
                
                'если файл в составе коммандной строки, например: C:\WINDOWS\system32\svchost -k rpcss.exe
                
                If argc > 2 Then        ' 1 -> app exe self, 2 -> actual cmd, 3 -> arg
                
                  Stady = 24
                
                  If Not FileExists(argv(1)) Then   ' если запускающий файл не существует -> ищем его
                    Stady = 25
                    FoundFile = FindOnPath(argv(1))
                    argv(1) = FoundFile
                  Else
                    FoundFile = argv(1)
                  End If
                
                  Stady = 26
                
                  ' если запускающий файл существует (иначе, нет смысла проверять остальные аргументы)
                  If 0 <> Len(FoundFile) Then
                    
                    'флаг о том, что служба запускает составную командную строку, в которой как минимум первый (запускающий файл) существует
                    IsCompositeCmd = True
                
                    isSafeMSCmdLine = True
                
                    Stady = 27
                 
                    For J = 1 To UBound(argv) ' argv[1] -> запускающий файл в цепочке
                    
                        ' проверяем хеш корневого сертификата каждого из элементов командной строки, если он был найден по известным путям Path
                        
                        FoundFile = FindOnPath(argv(J))
                        
                        Stady = 28
                        
                        If 0 <> Len(FoundFile) Then
                        
                            Stady = 29
                        
                            SignVerify FoundFile, 0&, SignResult
                            
                            Stady = 30
                            
                            IsMSCert = IsMicrosoftCertHash(SignResult.RootCertHash) And SignResult.isLegit
                            
                            DoEvents
                            
                            Stady = 31
                            
                            If Not IsMSCert Then isSafeMSCmdLine = False: Exit For
                            
                        End If
                        
                    Next
                
                  End If
                End If
            
            End If
            
            Stady = 32
            
            If 0 = Len(sFile) Then
                sFile = "(no file)"
            Else
                If (Not FileExists(sFile)) And (Not IsCompositeCmd) Then
                    sFile = sFile & " (file missing)"
                Else
                    If IsCompositeCmd Then
                        FoundFile = argv(1)
                    Else
                        FoundFile = sFile
                    End If
                    Stady = 33
                    sCompany = GetFilePropCompany(FoundFile)
                    If Len(sCompany) = 0 Then sCompany = "Unknown owner"
                End If
            End If
            
            If Not IsCompositeCmd And sFile <> "(no file)" Then    'иначе, такая проверка уже выполнена ранее
                Stady = 34
                'If sFile = "C:\Windows\system32\SearchIndexer.exe" Then Stop
                SignVerify sFile, 0&, SignResult
                DoEvents
            End If
            
            Stady = 35
            
            With SignResult
                ' если корневой сертификат цепочки доверия принадлежит Майкрософт, то исключаем службу из лога
                If Not (IsMicrosoftCertHash(.RootCertHash) And .isLegit And bHideMicrosoft) Then
                    Stady = 36
                    If bMD5 Then
                        If sFile <> "(no file)" Then sFile = sFile & GetFileFromAutostart(sFile)
                    End If
                    Stady = 37
                    sHit = "O23 - Service " & IIf(ServState <> SERVICE_STOPPED, "R", "S") & lStart & _
                        ": " & sDisplayName & " - (" & sName & ")" & " - " & sCompany & " - " & sFile
                        
                    If Not IsOnIgnoreList(sHit) Then
                        Stady = 38
                        AddToScanResultsSimple "O23", sHit
                    End If
                End If
            End With
        End If
        End If
Continue:
    Next i
    
    Exit Sub
        
ErrorHandler:
    ErrorMsg err, "modMain_CheckO23Item", "Service=", sDisplayName, "Stady=", Stady
    If inIDE Then Stop: Resume Next
End Sub

Public Function IsMicrosoftFile(sFile As String) As Boolean
    Dim SignResult As SignResult_TYPE
    SignVerify sFile, 0&, SignResult
    IsMicrosoftFile = IsMicrosoftCertHash(SignResult.RootCertHash) And SignResult.isLegit
End Function

Public Sub FixO23Item(sItem$)
    'stop & disable NT service - DON'T delete it
    'O23 - Service: <displayname> - <company> - <file>
    ' (file missing) or (filesize .., MD5 ..) can be appended
    If Not bIsWinNT Then Exit Sub
    On Error GoTo ErrorHandler:
    
    Dim sServices$(), i&, sName$, sDisplayName$
    sServices = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    If UBound(sServices) = 0 Or UBound(sServices) = -1 Then Exit Sub
    sDisplayName = Mid$(sItem, InStr(sItem, ": ") + 2)
    sDisplayName = Left$(sDisplayName, InStr(sDisplayName, " - ") - 1)
    For i = 0 To UBound(sServices)
        If sDisplayName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServices(i), "DisplayName") Then
            sName = sServices(i)
            
            RegSetDwordVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start", 4
            
            'this does the same as AboutBuster: run NET STOP on the
            'service. if the API way wouldn't crash VB everytime, I'd
            'use that. :/
            Shell sWinSysDir & "\NET.exe STOP """ & sName & """ /y", vbHide
            'better do the display name too in case the regkey name
            'has funky characters (res://dll or temp\sp.html parasites)
            Shell sWinSysDir & "\NET.exe StOP """ & sDisplayName & """ /y", vbHide
            Sleep 1000
            DoEvents
            
            RegSetDwordVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start", 4
            
            '// TODO: Check it!
            DeleteNTService sName
            If sName <> sDisplayName Then DeleteNTService sDisplayName
            
            'bRebootNeeded = True
        End If
    Next i
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "modMain_FixO23Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO24Item()
    On Error GoTo ErrorHandler:
    'activex desktop components
    Dim sDCKey$, sComponents$(), i&
    Dim sSource$, sSubscr$, sName$, sHit$, Wow64key As Boolean
    
    Wow64key = False
    
    sDCKey = "Software\Microsoft\Internet Explorer\Desktop\Components"
    sComponents = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sDCKey, Wow64key), "|")
    
    For i = 0 To UBound(sComponents)
        If RegKeyExists(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), Wow64key) Then
            sSource = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "Source", Wow64key)
            sSubscr = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "SubscribedURL", Wow64key)
            sName = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "FriendlyName", Wow64key)
            If sName = vbNullString Then sName = "(no name)"
            If Not (LCase$(sSource) = "about:home" And LCase$(sSubscr) = "about:home") And _
               Not (UCase$(sSource) = "131A6951-7F78-11D0-A979-00C04FD705A2" And UCase$(sSubscr) = "131A6951-7F78-11D0-A979-00C04FD705A2") Then
               
                sHit = "O24 - Desktop Component " & sComponents(i) & ": " & sName & " - " & IIf(sSource <> "", sSource, IIf(sSubscr <> "", sSubscr, "(no file)"))
                
                If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O24", sHit
            End If
        End If
    Next i
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_CheckO24Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO24Item(sItem$)
    On Error GoTo ErrorHandler:

    'delete the entire registry key
    'O24 - Desktop Component 1: Internet Explorer Channel Bar - 131A6951-7F78-11D0-A979-00C04FD705A2
    'O24 - Desktop Component 2: Security - %windir%\index.html
    
    Const SPIF_UPDATEINIFILE As Long = 1&
    
    Dim sDCKey$, sNum$, sName$, sURL$, sComponents$(), i&, sTestName$, sTestURL1$, sTestURL2$
    sDCKey = "Software\Microsoft\Internet Explorer\Desktop\Components"
    
    sNum = Mid$(sItem, InStr(sItem, ":") - 1, 1)
    sName = Mid$(sItem, InStr(sItem, ":") + 2)
    sURL = Mid$(sName, InStr(sName, " - ") + 3)
    sName = Left$(sName, InStr(sName, " - ") - 1)
    If "(no name)" = sName Then
        sName = vbNullString
    End If
    
    sTestName = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sNum, "FriendlyName")
    sTestURL1 = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sNum, "Source")
    sTestURL2 = RegGetString(HKEY_CURRENT_USER, sDCKey & "\" & sNum, "SubscribedURL")
    If sName = sTestName And (sURL = sTestURL1 Or sURL = sTestURL2) Then
        'found it!
        RegDelKey HKEY_CURRENT_USER, sDCKey & "\" & sNum
        If FileExists(sTestURL1) Then DeleteFileWEx (StrPtr(sTestURL1))
        If FileExists(sTestURL2) Then DeleteFileWEx (StrPtr(sTestURL2))
        
        SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, 0&, SPIF_UPDATEINIFILE 'SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_FixO23Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub
  
    
Public Sub FixUNIXHostsFile()
    'unix-style = hosts file has inproper linebreaks
    'Win32 linebreak: chr(13) + chr(10)
    'UNIX  linebreak: chr(10)
    'Mac   linebreak: chr(13)
    On Error GoTo ErrorHandler:
    If Not FileExists(sHostsFile) Then Exit Sub
    If FileLenW(sHostsFile) = 0 Then Exit Sub
    
    Dim sLine$, sFile$, sNewFile$, iAttr&, vContent As Variant, ff%
    iAttr = GetFileAttributes(StrPtr(sHostsFile))
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetFileAttributes StrPtr(sHostsFile), vbNormal
    
    ff = FreeFile()
    Open sHostsFile For Binary As #ff
        sFile = Input(FileLenW(sHostsFile), #ff)
    Close #ff
    
    'temp rename all proper linebreaks, replace unix-style
    'linebreaks with proper linebreaks, rename back
    sNewFile = sFile
    sNewFile = Replace$(sNewFile, vbCrLf, "/|\|/")
    sNewFile = Replace$(sNewFile, Chr(10), vbCrLf)
    'sNewFile = replace$(sNewFile, vbCrLf, "/|\|/")
    'sNewFile = replace$(sNewFile, Chr(13), vbCrLf)
    sNewFile = Replace$(sNewFile, "/|\|/", vbCrLf)
    If sNewFile <> sFile Then
        DeleteFileWEx (StrPtr(sHostsFile))
        ff = FreeFile()
        Open sHostsFile For Output As #ff
            Print #ff, sNewFile
        Close #ff
    End If
    SetFileAttributes StrPtr(sHostsFile), iAttr
    Exit Sub
    
ErrorHandler:
    Close #ff
    ErrorMsg err, "modMain_FixUNIXHostsFile"
    If inIDE Then Stop: Resume Next
End Sub

Public Function IsOnIgnoreList(sHit$, Optional UpdateList As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    
    Static IsInit As Boolean
    Static sIgnoreList() As String
    
    If IsInit And Not UpdateList Then
        If inArray(sHit, sIgnoreList) Then IsOnIgnoreList = True
    Else
        Dim iIgnoreNum&, i&
        
        IsInit = True
        ReDim sIgnoreList(0)
        
        iIgnoreNum = Val(RegReadHJT("IgnoreNum", "0"))
        If iIgnoreNum > 0 Then ReDim sIgnoreList(iIgnoreNum)
        
        For i = 1 To iIgnoreNum
            sIgnoreList(i) = Crypt(RegReadHJT("Ignore" & i, vbNullString), sProgramVersion, doCrypt:=False)
        Next
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "modMain_IsOnIgnoreList", sHit
    If inIDE Then Stop: Resume Next
End Function

Public Sub ErrorMsg(ErrObj As ErrObject, sProcedure$, ParamArray CodeModule())
    Dim sMsg$, sParameters$, HRESULT$, sErrDesc$, iErrNum&, iErrLastDll&, i&
    Dim hWnd As Long, ptr As Long, hMem As Long
    Dim DateTime As String, curTime As Date, ErrText$
    'If iErrNum = 0 Then Exit Sub
    'sMsg = "An unexpected error has occurred at procedure: " & _
           sProcedure & "(" & sParameters & ")" & vbCrLf & _
           "Error #" & CStr(iErrNum) & " - " & sErrDesc & vbCrLf & vbCrLf & _
           "Please email me at www.merijn.org/contact.html, reporting the following:" & vbCrLf & _
           "* What you were trying to fix when the error occurred, if applicable" & vbCrLf & _
           "* How you can reproduce the error" & vbCrLf & _
           "* A complete HiJackThis scan log, if possible" & vbCrLf & vbCrLf & _
           "Windows version: " & sWinVersion & vbCrLf & _
           "MSIE version: " & sMSIEVersion & vbCrLf & _
           "HiJackThis version: " & App.Major & "." & App.Minor & "." & App.Revision & _
           vbCrLf & vbCrLf & "This message has been copied to your clipboard." & _
           vbCrLf & "Click OK to continue the rest of the scan."
    
    sErrDesc = ErrObj.Description
    iErrNum = ErrObj.Number
    iErrLastDll = ErrObj.LastDllError
    
    'If iErrNum = 0 Then Exit Sub
    
    If iErrNum <> 33333 Then    'error defined by HJT
        HRESULT = MessageText(CLng(iErrNum))
    End If
    
    For i = 0 To UBound(CodeModule)
        sParameters = sParameters & CodeModule(i) & " "
    Next
    
    ' Emergency mode (if translation module is not initialized yet)
    If 0 = Len(Translate(590)) Then
        Translate(590) = "Please help us improve HiJackThis by reporting this error." & _
            vbCrLf & vbCrLf & "Error message has been copied to clipboard." & _
            vbCrLf & "Click 'Yes' to submit." & _
            vbCrLf & vbCrLf & "Error Details: " & _
            vbCrLf & vbCrLf & "An unexpected error has occurred at function: "
        TranslateNative(590) = Translate(590)
    End If
    
    sMsg = TranslateNative(590) & " " & _
    sProcedure & " (" & sParameters & ")" & vbCrLf & _
    "Error # " & iErrNum & " - " & sErrDesc & _
    vbCrLf & "LastDllError # " & iErrLastDll & _
    vbCrLf & "HRESULT: " & HRESULT & vbCrLf & vbCrLf & _
    "Windows version: " & OSInfo.Bitness & " " & OSInfo.OSName & " (" & OSInfo.Edition & "), " & _
    OSInfo.Major & "." & OSInfo.Minor & "." & OSInfo.Build & ", " & _
    "Service Pack: " & OSInfo.SPVer & "" & IIf(OSInfo.IsSafeBoot, " (Safe Boot)", "") & _
    vbCrLf & vbCrLf & AppVer
    
    '"Windows version: " & sWinVersion & vbCrLf & vbCrLf & AppVer
    
    If Not bAutoLogSilent Then
    
      Clipboard.Clear
      Clipboard.SetText sMsg
      
      If OpenClipboard(hWnd) Then
        hMem = GlobalAlloc(GMEM_MOVEABLE, 4)
        If hMem <> 0 Then
            ptr = GlobalLock(hMem)
            If ptr <> 0 Then
                GetMem4 &H419, ByVal ptr
                GlobalUnlock hMem
                SetClipboardData CF_LOCALE, hMem
            End If
        End If
        hMem = GlobalAlloc(GMEM_MOVEABLE, LenB(sMsg))
        If hMem <> 0 Then
            ptr = GlobalLock(hMem)
            If ptr <> 0 Then
                lstrcpyn ByVal ptr, ByVal StrPtr(sMsg), LenB(sMsg)
                'CopyMemory ByVal ptr, ByVal StrPtr(sMsg), LenB(sMsg)
                GlobalUnlock hMem
                SetClipboardData CF_UNICODETEXT, hMem
            End If
        End If
        CloseClipboard
      End If
    End If
    
    ' Append error log
    
    curTime = Now()
    
    DateTime = Right$("0" & Day(curTime), 2) & _
        "." & Right$("0" & Month(curTime), 2) & _
        "." & Year(curTime) & _
        " " & Right$("0" & Hour(curTime), 2) & _
        ":" & Right$("0" & Minute(curTime), 2) & _
        ":" & Right$("0" & Second(curTime), 2)
    
    ErrText = " - " & sProcedure & _
        " - #" & iErrNum & " " & _
        sErrDesc & _
        ". LastDllError = " & iErrLastDll & _
        IIf(Len(HRESULT), " (" & HRESULT & ")", "") & " " & _
        IIf(Len(sParameters), "" & sParameters, "")
    
    Debug.Print ErrText
    
    ErrReport = ErrReport & vbCrLf & _
        "- " & DateTime & ErrText
    
    'If Not bAutoLogSilent Then
    
    If Not bAutoLog And Not bSkipErrorMsg Then
        frmError.Show vbModeless
        frmError.Label1.Caption = sMsg
        frmError.Hide
        frmError.Show vbModal
        
'        If vbYes = MsgBoxW(sMsg, vbCritical Or vbYesNo, Translate(591)) Then
'            Dim szParams As String
'            Dim szCrashUrl As String
'            szCrashUrl = "http://safezone.cc/threads/25222/" 'https://sourceforge.net/p/hjt/_list/tickets"
'
'            'szParams = "function=" & sProcedure
'            'szParams = szParams & "&params=" & sParameters
'            'szParams = szParams & "&errorno=" & iErrNum
'            'szParams = szParams & "&errorlastdll=" & iErrLastDll
'            'szParams = szParams & "&errortxt" & sErrDesc
'            'szParams = szParams & "&winver=" & sWinVersion
'            'szParams = szParams & "&hjtver=" & App.Major & "." & App.Minor & "." & App.Revision
'            'szCrashUrl = szCrashUrl & URLEncode(szParams)
'            If True = IsOnline Then
'                ShellExecute 0&, "open", szCrashUrl, vbNullString, vbNullString, vbNormalFocus
'            Else
'                'MsgBoxW "No Internet Connection Available"
'                MsgBoxW Translate(560)
'            End If
'        End If
    End If
End Sub

Public Function MessageText(lCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
    
    Dim sRtrnMsg   As String
    Dim lret        As Long

    sRtrnMsg = Space$(MAX_PATH)
    lret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, lCode, 0&, StrPtr(sRtrnMsg), MAX_PATH, 0&)
    If lret > 0 Then
        MessageText = Left$(sRtrnMsg, lret)
        MessageText = Replace$(MessageText, vbCrLf, vbNullString)
    End If
End Function

Public Sub CheckDateFormat()
    Dim sBuffer$, uST As SYSTEMTIME
    With uST
        .wDay = 10
        .wMonth = 11
        .wYear = 2003
    End With
    sBuffer = String$(255, 0)
    GetDateFormat 0&, 0&, uST, 0&, StrPtr(sBuffer), 255&
    sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    
    'last try with GetLocaleInfo didn't work on Win2k/XP
    If InStr(sBuffer, "10") < InStr(sBuffer, "11") Then
        bIsUSADateFormat = False
        'msgboxW "sBuffer = " & sBuffer & vbCrLf & "10 < 11, so bIsUSADateFormat False"
    Else
        bIsUSADateFormat = True
        'msgboxW sBuffer & vbCrLf & "10 !< 11, so bIsUSADateFormat True"
    End If
    
    'Dim lLndID&, sDateFormat$
    'lLndID = GetSystemDefaultLCID()
    'sDateFormat = String$(255, 0)
    'GetLocaleInfo lLndID, LOCALE_SSHORTDATE, sDateFormat, 255
    'sDateFormat = left$(sDateFormat, InStr(sDateFormat, vbnullchar) - 1)
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

'Public Function UnEscape(sURL As String) As String
'    Dim i&, sDummy$, sHex$
'
'    'replace hex codes with proper character
'    sDummy = sURL
'
'    'don't need entire ascii range, 32-126
'    'is all readable characters (I think)
'    'For i = 1 To 255
'    For i = 32 To 126
'        sHex = Hex(i)
'        If Len(sHex) = 1 Then sHex = "0" & sHex
'        sDummy = replace$(sDummy, "%" & sHex, Chr(i), , , vbTextCompare)
'    Next i
'
'    UnEscape = sDummy & " (obfuscated)"
'End Function

Public Function UnEscape(ByVal StringToDecode As String) As String
    Dim i As Long
    Dim acode As Integer, lTmp As Long, HexChar As String

    On Error GoTo ErrorHandler

'    Set scr = CreateObject("MSScriptControl.ScriptControl")
'    scr.Language = "VBScript"
'    scr.Reset
'    Escape = scr.Eval("unescape(""" & s & """)")

    UnEscape = StringToDecode

    If InStr(UnEscape, "%") = 0 Then
         Exit Function
    End If
    For i = Len(UnEscape) To 1 Step -1
        acode = Asc(Mid$(UnEscape, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars

            Case 37
                ' Decode % value
                HexChar = UCase$(Mid$(UnEscape, i + 1, 2))
                If HexChar Like "[0123456789ABCDEF][0123456789ABCDEF]" Then
                    lTmp = CLng("&H" & HexChar)
                    UnEscape = Left$(UnEscape, i - 1) & Chr$(lTmp) & Mid$(UnEscape, i + 3)
                End If
        End Select
    Next
    Exit Function
ErrorHandler:
    ErrorMsg err, "UnEscape", "string:", StringToDecode
End Function

Public Function HasSpecialCharacters(sName$) As Boolean
    'function checks for special characters in string,
    'like Chinese or Japanese.
    'Used in CheckO3Item (IE Toolbar)
    HasSpecialCharacters = False
    
    'function disabled because of proper DBCS support
    Exit Function
    
    If Len(sName) <> lstrlen(StrPtr(sName)) Then
        HasSpecialCharacters = True
        Exit Function
    End If
    
    If Len(sName) <> LenB(StrConv(sName, vbFromUnicode)) Then
        HasSpecialCharacters = True
        Exit Function
    End If
End Function

Public Sub CheckForReadOnlyMedia()
    Dim sMsg$, ff%
    On Error Resume Next
    
    ff = FreeFile()
    Open BuildPath(AppPath(), "~dummy.tmp") For Output As #ff
        Print #ff, "."
    Close #ff
    
    If err.Number Then     'Some strange error happens here, if we delete .Number property
        'damn, got no write access
        bNoWriteAccess = True
        sMsg = Translate(7)
'        sMsg = "It looks like you're running HiJackThis from " & _
'               "a read-only device like a CD or locked floppy disk." & _
'               "If you want to make backups of items you fix, " & _
'               "you must copy HiJackThis.exe to your hard disk " & _
'               "first, and run it from there." & vbCrLf & vbCrLf & _
'               "If you continue, you might get 'Path/File Access' " & _
'               "errors - do NOT email me those please."
        MsgBoxW sMsg, vbExclamation
        
    End If
    DeleteFileWEx (StrPtr(BuildPath(AppPath(), "~dummy.tmp")))
End Sub

Public Sub SetAllFontCharset()
    On Error GoTo ErrorHandler:

    Dim ctl         As Control
    Dim ctlBtn      As CommandButton
    Dim ctlCheckBox As CheckBox
    Dim ctlTxtBox   As TextBox
    Dim ctlLstBox   As ListBox
    Dim CtlLbl      As Label

    For Each ctl In frmMain.Controls
        Select Case TypeName(ctl)
            Case "CommandButton"
                Set ctlBtn = ctl
                SetFontCharSet ctlBtn.Font
            Case "TextBox"
                Set ctlTxtBox = ctl
                SetFontCharSet ctlTxtBox.Font
            Case "ListBox"
                Set ctlLstBox = ctl
                SetFontCharSet ctlLstBox.Font
            Case "Label"
                Set CtlLbl = ctl
                SetFontCharSet CtlLbl.Font
            Case "CheckBox"
                Set ctlCheckBox = ctl
                If ctlCheckBox.Name <> "chkConfigTabs" Then
                    SetFontCharSet ctlCheckBox.Font
                End If
        End Select
    Next ctl

'    With frmMain
'        SetFontCharSet .txtCheckUpdateProxy.Font
'        SetFontCharSet .txtDefSearchAss.Font
'        SetFontCharSet .txtDefSearchCust.Font
'        SetFontCharSet .txtDefSearchPage.Font
'        SetFontCharSet .txtDefStartPage.Font
'        SetFontCharSet .txtHelp.Font
'        SetFontCharSet .txtNothing.Font
'
'        SetFontCharSet .lstBackups.Font
'        SetFontCharSet .lstIgnore.Font
'        SetFontCharSet .lstResults.Font
'    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_SetAllFontCharset"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub SetFontCharSet(objTxtboxFont As Object)
    On Error GoTo ErrorHandler:
    
    'A big thanks to 'Gun' and 'Adult', two Japanese users
    'who helped me greatly with this
    
    'https://msdn.microsoft.com/en-us/library/aa241713(v=vs.60).aspx
    
    Static IsInit As Boolean
    Static lLCID As Long
    Dim bNonUsCharset As Boolean
    
    bNonUsCharset = True
    
    If Not IsInit Then
        lLCID = GetUserDefaultLCID()
        IsInit = True
    End If
    
    Select Case lLCID
         Case &H404
            objTxtboxFont.Charset = CHINESEBIG5_CHARSET
            objTxtboxFont.Name = ChrW(&H65B0) & ChrW(&H7D30) & ChrW(&H660E) & ChrW(&H9AD4)   'New Ming-Li
         Case &H411
            objTxtboxFont.Charset = SHIFTJIS_CHARSET
            objTxtboxFont.Name = ChrW(&HFF2D) & ChrW(&HFF33) & ChrW(&H20) & ChrW(&HFF30) & ChrW(&H30B4) & ChrW(&H30B7) & ChrW(&H30C3) & ChrW(&H30AF)
         Case &H412
            objTxtboxFont.Charset = HANGEUL_CHARSET
            objTxtboxFont.Name = ChrW(&HAD74) & ChrW(&HB9BC)
         Case &H804
            objTxtboxFont.Charset = CHINESESIMPLIFIED_CHARSET
            objTxtboxFont.Name = ChrW(&H5B8B) & ChrW(&H4F53)
         Case Else
            objTxtboxFont.Charset = DEFAULT_CHARSET
            'objTxtboxFont.Name = ""
            bNonUsCharset = False
    End Select
    
    If bNonUsCharset Then objTxtboxFont.size = 9
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modMain_SetFontCharSet"
    If inIDE Then Stop: Resume Next
End Sub

Public Function TrimNull(s$) As String
    TrimNull = Left$(s, lstrlen(StrPtr(s)))
End Function

Public Sub CheckForStartedFromTempDir()
    'if user picks 'run from current location when downloading HiJackThis.exe,
    'or runs file directly from zip file, exe will be ran from temp folder,
    'meaning a reboot or cache clean could delete it, as well any backups
    'made. Also the user won't be able to find the exe anymore :P
    
    'fixed - 2.0.7
    
    Dim v1          As String
    Dim v2          As String
    Dim Cnt         As Long
    Dim sBuffer     As String
    Dim RunFromTemp As Boolean
    Dim sMsg        As String
    
'    sMsg = "HiJackThis appears to have been started from a temporary " & _
'               "folder. Since temp folders tend to be be emptied regularly, " & _
'               "it's wise to copy HiJackThis.exe to a folder of its own, " & _
'               "for instance C:\Program Files\HiJackThis." & vbCrLf & _
'               "This way, any backups that will be made of fixed items " & _
'               "won't be lost." & vbCrLf & vbCrLf & _
'               "May I unpack HJT to desktop for you ?"
'               '"Please quit HiJackThis and copy it to a separate folder " & _
'               '"first before fixing any items."

    'Just too many words
    'User can be shocked and he will close this program immediately and forewer :)
    'l'll try this simple (just only this time):
    
    'Launch from the archive is forbidden !" & vbCrLf & vbCrLf & "May I unzip to desktop for you ?"
    sMsg = TranslateNative(8)
    
    ' проверка на запуск из архива
    If Len(TempCU) <> 0& Then
    
        If StrBeginWith(AppPath(), TempCU) Then RunFromTemp = True
        If Not RunFromTemp Then         'fix, когда app.path раскрывается в стиле 8.3
            'v1 = PathToDOS(AppPath(), Force:=True)
            'v2 = PathToDOS(TempCU, Force:=True)
            
            'If Len(v1) = 0 Then v1 = App.Path
            'If Len(v2) = 0 Then v2 = TempCU
            
            sBuffer = String$(MAX_PATH, vbNullChar)
            Cnt = GetLongPathName(StrPtr(AppPath()), StrPtr(sBuffer), Len(sBuffer))
            If Cnt Then
                v1 = Left$(sBuffer, Cnt)
            Else
                v1 = AppPath()
            End If

            sBuffer = String$(MAX_PATH, vbNullChar)
            Cnt = GetLongPathName(StrPtr(TempCU), StrPtr(sBuffer), Len(sBuffer))
            If Cnt Then
                v2 = Left$(sBuffer, Cnt)
            Else
                v2 = TempCU
            End If
            
            If Len(v1) <> 0 And Len(v2) <> 0 And StrBeginWith(v1, v2) Then RunFromTemp = True
        End If
        
        If RunFromTemp Then
            'msgboxW "Запуск из архива запрещен !" & vbCrLf & "Распаковать на рабочий стол для Вас ?", vbExclamation, AppName
            If MsgBoxW(sMsg, vbExclamation Or vbYesNo, "HiJackThis") = vbYes Then
                Dim NewFile As String
                NewFile = Desktop & "\" & AppExeName(True)
                If FileExists(NewFile) Then     ', Cache:=NO_CACHE
                    SetFileAttributes StrPtr(NewFile), GetFileAttributes(StrPtr(NewFile)) And Not FILE_ATTRIBUTE_READONLY
                    DeleteFileWEx StrPtr(NewFile)
                End If
                CopyFile StrPtr(AppPath(True)), StrPtr(NewFile), ByVal 0&
                If FileExists(NewFile) Then     ', Cache:=NO_CACHE
                    frmMain.ReleaseMutex
                    Proc.ProcessRun NewFile     ', "/twice"
                    Unload frmMain
                    End
                Else
                    'Could not unzip file to Desktop! Please, unzip it manually.
                    MsgBoxW Translate(1007), vbCritical
                    Unload frmMain
                    End
                End If
            Else
                Unload frmMain
                End
            End If
        End If
    End If
End Sub

Public Sub ShowFileProperties(sFile$, Handle As Long)
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_DOENVSUBST Or SEE_MASK_FLAG_NO_UI
        .hWnd = Handle
        .lpFile = StrPtr(sFile)
        .lpVerb = StrPtr("properties")
        .nShow = 1
    End With
    ShellExecuteEx uSEI
End Sub

Public Sub RestartSystem(Optional sExtraPrompt$)
    If bIsWinNT Then
        SHRestartSystemMB frmMain.hWnd, StrConv(sExtraPrompt & IIf(sExtraPrompt <> vbNullString, vbCrLf & vbCrLf, vbNullString), vbUnicode), 2
    Else
        SHRestartSystemMB frmMain.hWnd, sExtraPrompt & IIf(sExtraPrompt <> vbNullString, vbCrLf & vbCrLf, vbNullString), 0
    End If
End Sub

Public Sub DeleteFileOnReboot(sFile$, Optional bDeleteBlindly As Boolean = False)
    'If Not bIsWinNT Then Exit Sub
    If Not FileExists(sFile) And Not bDeleteBlindly Then Exit Sub
    If bIsWinNT Then
        MoveFileEx StrPtr(sFile), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    Else
        Dim sDummy$, ff%
        On Error Resume Next
        ff = FreeFile()
        TryUnlock sWinDir & "\wininit.ini"
        Open sWinDir & "\wininit.ini" For Append As #ff
            Print #ff, "[rename]"
            Print #ff, "NUL=" & GetDOSFilename(sFile)
            Print #ff,
        Close #ff
    End If
    RestartSystem Replace$(Translate(342), "[]", sFile)
    'RestartSystem "The file '" & sFile & "' will be deleted by Windows when the system restarts."
    
    'Windows Server 2003 Note:
    'https://support.microsoft.com/en-us/kb/948601
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
    Dim sDoubleTLDs$(), i&
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

'if short name is unavailable, it returns source string anyway
Public Function GetDOSFilename$(sFile$, Optional bReverse As Boolean = False)
    'works for folders too btw
    Dim Cnt&, sBuffer$
    If bReverse Then
        sBuffer = Space$(MAX_PATH_W)
        Cnt = GetLongPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
        If Cnt Then
            GetDOSFilename = Left$(sBuffer, Cnt)
        Else
            GetDOSFilename = sFile
        End If
    Else
        sBuffer = Space$(MAX_PATH)
        Cnt = GetShortPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
        If Cnt Then
            GetDOSFilename = Left$(sBuffer, Cnt)
        Else
            GetDOSFilename = sFile
        End If
    End If
End Function

Public Function GetUser() As String
    Dim sUsername$
    sUsername = String$(MAX_PATH, vbNullChar)
    If 0 <> GetUserName(StrPtr(sUsername), MAX_PATH) Then
        sUsername = Left$(sUsername, lstrlen(StrPtr(sUsername)))
    End If
    GetUser = sUsername 'UCase$(sUserName)
End Function

Public Function GetComputer() As String
    Dim sComputerName$
    sComputerName = String$(MAX_PATH, vbNullChar)
    If 0 <> GetComputerName(StrPtr(sComputerName), MAX_PATH) Then
        sComputerName = Left$(sComputerName, lstrlen(StrPtr(sComputerName)))
    End If
    GetComputer = sComputerName 'UCase$(sComputerName)
End Function

Public Sub CopyFolder(sFolder$, sTo$)
    Dim uFOS As SHFILEOPSTRUCT
    With uFOS
        .wFunc = FO_COPY
        .pFrom = StrPtr(sFolder)
        .pTo = StrPtr(sTo)
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    MsgBoxW SHFileOperation(uFOS)
End Sub

Public Sub DeleteFolder(sFolder$)
    Dim uFOS As SHFILEOPSTRUCT
    With uFOS
        .wFunc = FO_DELETE
        .pFrom = StrPtr(sFolder)
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    SHFileOperation uFOS
End Sub

Public Sub MoveFolder(sFolder$, sTo$)
    Dim uFOS As SHFILEOPSTRUCT
    With uFOS
        .wFunc = FO_MOVE
        .pFrom = StrPtr(sFolder)
        .pTo = StrPtr(sTo)
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    SHFileOperation uFOS
End Sub

'Public Function ExpandEnvironmentVars$(s$)
'    Dim sDummy$, lLen&
'    If InStr(s, "%") = 0 Then
'        ExpandEnvironmentVars = s
'        Exit Function
'    End If
'    lLen = ExpandEnvironmentStrings(StrPtr(s), 0&, 0&)
'    If lLen > 0 Then
'        sDummy = String$(lLen, 0)
'        ExpandEnvironmentStrings StrPtr(s), StrPtr(sDummy), Len(sDummy)
'        sDummy = TrimNull(sDummy)
'
'        If InStr(sDummy, "%") = 0 Then
'            ExpandEnvironmentVars = sDummy
'            Exit Function
'        End If
'    Else
'        sDummy = s
'    End If
'End Function

Public Function GetUserType$()
    'based on OpenProcessToken API example from API-Guide
    Dim hProcessToken&
    Dim BufferSize&
    Dim psidAdmin&, psidPower&, psidUser&, psidGuest&
    Dim lResult&
    Dim i&
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
                CloseHandle hProcessToken
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
                    If IsValidSid(tpTokens.Groups(i).SID) Then
                    
                        ' Test for a match between the admin sid equalling your sid's
                        If EqualSid(tpTokens.Groups(i).SID, psidAdmin) Then
                            GetUserType = "Administrator"
                            Exit For
                        End If
                        If EqualSid(tpTokens.Groups(i).SID, psidPower) Then
                            GetUserType = "Power User"
                            Exit For
                        End If
                        If EqualSid(tpTokens.Groups(i).SID, psidUser) Then
                            GetUserType = "Limited User"
                            Exit For
                        End If
                        If EqualSid(tpTokens.Groups(i).SID, psidGuest) Then
                            GetUserType = "Guest"
                            Exit For
                        End If
                    End If
                Next
            End If
            If psidAdmin Then FreeSid psidAdmin
            If psidPower Then FreeSid psidPower
            If psidUser Then FreeSid psidUser
            If psidGuest Then FreeSid psidGuest
        End If
        CloseHandle hProcessToken
    End If
End Function

Public Function MapSIDToUsername(sSID As String) As String
    
    On Error GoTo ErrorHandler:
    
'    Dim objWMI As Object, objSID As Object
'    Set objWMI = GetObject("winmgmts:{impersonationLevel=Impersonate}")
'    Set objSID = objWMI.Get("Win32_SID.SID='" & sSID & "'")
'    MapSIDToUsername = objSID.AccountName
'    Set objSID = Nothing
'    Set objWMI = Nothing
    
    '   PURPOSE: there are certain builtin accounts on Windows NT which do not have a mapped
    '   account name. LookupAccountSid will return the error ERROR_NONE_MAPPED.  This function
    '   generates SIDs for the following accounts that are not mapped:
    '    * ACCOUNT OPERATORS
    '    * SYSTEM OPERATORS
    '    * PRINTER OPERATORS
    '    * BACKUP OPERATORS
    '   the other SID it creates is a LOGON SID, it has a prefix of S-1-5-5.  a LOGON SID is a
    '   unique identifier for a user's logon session.
    
    Dim bufSid() As Byte
    Dim AccName As String
    Dim AccDomain As String
    Dim AccType As Long
    Dim ccAccName As Long
    Dim ccAccDomain As Long
    Dim OtherName()
    Dim lret As Long
    Dim tpSidAuth As SID_IDENTIFIER_AUTHORITY
    Dim pSid(3) As Long
    Dim psidLogonSid As Long
    Dim psidCheck As Long
    Dim i As Long
    
    If UCase$(sSID) = ".DEFAULT" Then
        MapSIDToUsername = "Default user"
        Exit Function
    End If
    
    MapSIDToUsername = "unknown"
    
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
    
    OtherName = Array("Account operators", "Server operators", "Printer operators", "Backup operators")
    
    bufSid = CreateBufferedSID(sSID)
    
    If SafeArrayGetDim(bufSid) Then
    
        AccName = String$(MAX_NAME, 0)
        AccDomain = String$(MAX_NAME, 0)
        ccAccName = Len(AccName)
        ccAccDomain = Len(AccDomain)
        psidCheck = VarPtr(bufSid(0))
    
        If 0 <> LookupAccountSid(0&, psidCheck, StrPtr(AccName), ccAccName, StrPtr(AccDomain), ccAccDomain, AccType) Then
        
            MapSIDToUsername = Left$(AccName, ccAccName)
            
        Else
        
            If err.LastDllError = ERROR_NONE_MAPPED Then
            
                ' Create account operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ACCOUNT_OPS, 0, 0, 0, 0, 0, 0, pSid(0))

                ' Create system operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_SYSTEM_OPS, 0, 0, 0, 0, 0, 0, pSid(1))
        
                ' Create printer operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_PRINT_OPS, 0, 0, 0, 0, 0, 0, pSid(2))
        
                ' Create backup operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_BACKUP_OPS, 0, 0, 0, 0, 0, 0, pSid(3))

                ' Create a logon SID.
                Call AllocateAndInitializeSid(tpSidAuth, 2, 5, 0, 0, 0, 0, 0, 0, 0, psidLogonSid)
    
                '*psnu =  SidTypeAlias;

                If EqualPrefixSid(psidCheck, psidLogonSid) Then
                    MapSIDToUsername = "LOGON SID"
                Else
                    For i = 0 To 3
                        If EqualSid(psidCheck, pSid(i)) Then
                            MapSIDToUsername = OtherName(i)
                            Exit For
                        End If
                    Next
                End If

                For i = 0 To 3
                    FreeSid pSid(i)
                Next
                FreeSid psidLogonSid
            End If
        End If
    End If
  Exit Function
ErrorHandler:
    ErrorMsg err, "MapSIDToUsername", "SID: ", sSID
    If inIDE Then Stop: Resume Next
End Function

'Public Sub ToggleWow64FSRedirection(bEnable As Boolean)
'    Static lWow64Old&
'    If bIsWin64 Then    'mean cheking of Wow64Process, but VB6 is always x86, so it doesn't matter
'        If bEnable Then
'            If lWow64Old <> 0 Then
'                Wow64RevertWow64FsRedirection lWow64Old
'                lWow64Old = 0
'            End If
'        Else
'            If lWow64Old = 0 Then Wow64DisableWow64FsRedirection lWow64Old
'        End If
'    End If
'End Sub

Public Sub SilentDeleteOnReboot(sCmd$)
    Dim sDummy$, sFileName$
    'sCmd is all command-line parameters, like this
    '/param1 /deleteonreboot c:\progra~1\bla\bla.exe /param3
    '/param1 /deleteonreboot "c:\program files\bla\bla.exe" /param3
    
    sDummy = Mid$(sCmd, InStr(sCmd, "/deleteonreboot") + Len("/deleteonreboot") + 1)
    If InStr(sDummy, """") = 1 Then
        'enclosed in quotes, chop off at next quote
        sFileName = Mid$(sDummy, 2)
        sFileName = Left$(sFileName, InStr(sFileName, """") - 1)
    Else
        'no quotes, chop off at next space if present
        If InStr(sDummy, " ") > 0 Then
            sFileName = Left$(sDummy, InStr(sDummy, " ") - 1)
        Else
            sFileName = sDummy
        End If
    End If
    DeleteFileOnReboot sFileName, True
End Sub

Public Sub DeleteFileShell(sFile$)
    If Not FileExists(sFile) Then Exit Sub
    Dim uSFO As SHFILEOPSTRUCT
    With uSFO
        .pFrom = StrPtr(sFile)
        .wFunc = FO_DELETE
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    SHFileOperation uSFO
End Sub

Function IsProcedureAvail(ByVal ProcedureName As String, ByVal DllFilename As String) As Boolean
    Dim hModule As Long, procAddr As Long
    hModule = LoadLibrary(StrPtr(DllFilename))
    If hModule Then
        procAddr = GetProcAddress(hModule, ProcedureName)
        FreeLibrary hModule
    End If
    IsProcedureAvail = (procAddr <> 0)
End Function


Public Function CmnDlgSaveFile(sTitle$, sFilter$, Optional sDefFile$)
    Dim uOFN As OPENFILENAME, sFile$
    On Error GoTo ErrorHandler:
    
    Const OFN_ENABLESIZING As Long = &H800000
    
    sFile = String$(256, 0)
    LSet sFile = sDefFile
    With uOFN
        .lStructSize = Len(uOFN)
        If InStr(sFilter, "|") > 0 Then sFilter = Replace$(sFilter, "|", vbNullChar)
        If Right$(sFilter, 2) <> vbNullChar & vbNullChar Then sFilter = sFilter & vbNullChar & vbNullChar
        .lpstrFilter = StrPtr(sFilter)
        .lpstrFile = StrPtr(sFile)
        .lpstrTitle = StrPtr(sTitle)
        .nMaxFile = Len(sFile)
        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_OVERWRITEPROMPT Or OFN_ENABLESIZING
    End With
    If GetSaveFileName(uOFN) = 0 Then Exit Function
    sFile = TrimNull(sFile)
    CmnDlgSaveFile = sFile
    Exit Function
    
ErrorHandler:
    ErrorMsg err, "modMain_CmnDlgSaveFile", "sTitle=", sTitle, "sFilter=", sFilter, "sDefFile=", sDefFile
    If inIDE Then Stop: Resume Next
End Function

'Public Function CmnDlgOpenFile(sTitle$, sFilter$, Optional sDefFile$)
'    Dim uOFN As OPENFILENAME, sFile$
'    On Error GoTo ErrorHandler:
'
'    sFile = sDefFile & String$(256 - Len(sDefFile), 0)
'    With uOFN
'        .lStructSize = Len(uOFN)
'        If InStr(sFilter, "|") > 0 Then sFilter = replace$(sFilter, "|", vbNullChar)
'        If Right$(sFilter, 2) <> vbNullChar & vbNullChar Then sFilter = sFilter & vbNullChar & vbNullChar
'        .lpstrFilter = sFilter
'        .lpstrFile = sFile
'        .lpstrTitle = sTitle
'        .nMaxFile = 256
'        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_PATHMUSTEXIST
'    End With
'    If GetOpenFileName(uOFN) = 0 Then Exit Function
'    sFile = TrimNull(uOFN.lpstrFile)
'    CmnDlgOpenFile = sFile
'    Exit Function
'
'ErrorHandler:
'    ErrorMsg err, "modMain_CmnDlgOpenFile", "sTitle=", sTitle, "sFilter=", sFilter, "sDefFile=", sDefFile
'    If inIDE Then Stop: Resume Next
'End Function

Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String = " ") As VbMsgBoxResult
    Dim hActiveWnd As Long, hMyWnd As Long, frm As Form
    If inIDE Then
        MsgBoxW = MsgBox(Prompt, Buttons, Title)
    Else
        hActiveWnd = GetForegroundWindow()
        For Each frm In Forms
            If frm.hWnd = hActiveWnd Then hMyWnd = hActiveWnd: Exit For
        Next
        MsgBoxW = MessageBox(IIf(hMyWnd <> 0, hMyWnd, frmMain.hWnd), StrPtr(Prompt), StrPtr(Title), ByVal Buttons)
    End If
End Function

Public Function UnQuote(str As String) As String   ' Trim quotes
    Const QT = """"
    Dim s As String: s = str
    Do While Left$(s, 1&) = QT
        s = Mid$(s, 2&)
    Loop
    Do While Right$(s, 1&) = QT
        s = Left$(s, Len(s) - 1&)
    Loop
    UnQuote = s
End Function

Public Sub ReInitScanResults()  'Global results structure will be cleaned

    'ReDim Scan.Globals(0)
    ReDim Scan(0)

End Sub

Public Sub InitVariables()
    'SysDisk
    'sWinDir
    'sWinSysDir
    'sSysDir (the same as sWinSysDir)
    'sWinSysDirWow64
    'PF_32
    'PF_64
    'AppData
    'LocalAppData
    'Desktop
    'UserProfile
    'AllUsersProfile
    'TempCU

    Const CSIDL_DESKTOP = 0&

    CRCinit

    'Init user type arrays of scan results
    ReInitScanResults

    Dim lr As Long
    SysDisk = Space$(MAX_PATH)
    lr = GetWindowsDirectory(StrPtr(SysDisk), MAX_PATH)
    If lr Then
        sWinDir = Left$(SysDisk, lr)
        SysDisk = Left$(SysDisk, 2)
    Else
        sWinDir = EnvironW("%SystemRoot%")
        SysDisk = EnvironW("%SystemDrive%")
    End If
    sWinSysDir = sWinDir & "\" & IIf(bIsWinNT, "system32", "system")
    sSysDir = sWinSysDir
    sWinSysDirWow64 = sWinDir & "\SysWow64"
    
    If bIsWin64 Then
        PF_32 = EnvironW("%ProgramFiles%")
        If OSver.MajorMinor >= 6.1 Then     'Win 7 and later
            PF_64 = EnvironW("%ProgramW6432%")
        Else
            PF_64 = SysDisk & "\Program Files"
        End If
    Else
        PF_32 = EnvironW("%ProgramFiles%")
        PF_64 = PF_32
    End If
    
    PF_32_Common = PF_32 & "\Common Files"
    PF_64_Common = PF_64 & "\Common Files"
    
    AppData = EnvironW("%AppData%")
    If OSver.bIsVistaOrLater Then
        LocalAppData = EnvironW("%LocalAppData%")
    Else
        LocalAppData = GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)
        If Len(LocalAppData) = 0 Then LocalAppData = EnvironW("%USERPROFILE%") & "\Local Settings\Application Data"
    End If
    
    Desktop = GetSpecialFolderPath(CSIDL_DESKTOP)
    UserProfile = EnvironW("%UserProfile%")
    AllUsersProfile = EnvironW("%ALLUSERSPROFILE%")
    
    'TempCU = Environ("temp") ' will return path in format 8.3 on XP
    TempCU = GetRegData(HKEY_CURRENT_USER, "Environment", "Temp")
    ' if REG_EXPAND_SZ is missing
    If InStr(TempCU, "%") <> 0 Then
        TempCU = EnvironW(TempCU)
    End If
    If Len(TempCU) = 0 Or InStr(TempCU, "%") <> 0 Then ' if there TEMP is not defined
        If OSver.bIsVistaOrLater Then
            TempCU = UserProfile & "\Local\Temp"
        Else
            TempCU = UserProfile & "\Local Settings\Temp"
        End If
    End If
    
    ' Shortcut interfaces initialization
    'IURL_Init
    ISL_Init
    
    With oDict
        Set .TaskWL_ID = New clsTrickHashTable
    End With
    
    Set colProfiles = New Collection
    GetProfiles
    
End Sub

Public Function EnvironW(ByVal SrcEnv As String) As String
    Dim lr As Long
    Dim buf As String
    If Len(SrcEnv) = 0 Then Exit Function
    If InStr(SrcEnv, "%") = 0 Then
        EnvironW = SrcEnv
    Else
        'redirector correction
        If OSver.bIsWin64 Then
            If InStr(1, SrcEnv, "%PROGRAMFILES%", 1) <> 0 Then
                SrcEnv = Replace$(SrcEnv, "%PROGRAMFILES%", PF_64, 1, 1, 1)
            End If
            If InStr(1, SrcEnv, "%COMMONPROGRAMFILES%", 1) <> 0 Then
                SrcEnv = Replace$(SrcEnv, "%COMMONPROGRAMFILES%", PF_64_Common, 1, 1, 1)
            End If
        End If
        buf = String$(MAX_PATH, vbNullChar)
        lr = ExpandEnvironmentStrings(StrPtr(SrcEnv), StrPtr(buf), MAX_PATH + 1)
        
        If lr Then
            EnvironW = Left$(buf, lr - 1)
        Else
            EnvironW = SrcEnv
        End If
        
        If InStr(EnvironW, "%") <> 0 Then
            If OSver.MajorMinor <= 6 Then
                If InStr(1, EnvironW, "%ProgramW6432%", 1) <> 0 Then
                    EnvironW = Replace$(EnvironW, "%ProgramW6432%", SysDisk & "\Program Files", 1, -1, 1)
                End If
            End If
        End If
    End If
End Function

Public Function GetSpecialFolderPath(CSIDL As Long, Optional hToken As Long = 0&) As String
    Const SHGFP_TYPE_CURRENT As Long = &H0&
    Const SHGFP_TYPE_DEFAULT As Long = &H1&
    Dim lr      As Long
    Dim sPath   As String
    sPath = String$(MAX_PATH, 0&)
    ' 3-th parameter - is a token of user
    lr = SHGetFolderPath(0&, CSIDL, hToken, SHGFP_TYPE_CURRENT, StrPtr(sPath))
    If lr = 0 Then GetSpecialFolderPath = Left$(sPath, lstrlen(StrPtr(sPath)))
End Function

Public Function StrInParamArray(Stri As String, ParamArray Etalon()) As Boolean
    Dim i As Long
    For i = 0 To UBound(Etalon)
        If StrComp(Stri, Etalon(i), 1) = 0 Then StrInParamArray = True: Exit For
    Next
End Function

Public Function GetParentDir(sPath As String) As String
    Dim pos As Long
    pos = InStrRev(sPath, "\")
    If pos <> 0 Then
        GetParentDir = Left$(sPath, pos - 1)
    End If
End Function

' Возвращает true, если искомое значение найдено в одном из элементов массива (lB, uB ограничивает просматриваемый диапазон индексов)
Public Function inArray( _
    Stri As String, _
    MyArray() As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    If lB = -2147483647 Then lB = LBound(MyArray)   'some trick
    If uB = 2147483647 Then uB = UBound(MyArray)    'Thanks to Казанский :)
    Dim i As Long
    For i = lB To uB
        If StrComp(Stri, MyArray(i), CompareMethod) = 0 Then inArray = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg err, "inArray"
    If inIDE Then Stop: Resume Next
End Function

'Note: Serialized array - it is a string which stores all items of array delimited by some character (default delimiter in HJT is '|' and '*' chars)
'Example 1: "string1*string2*string3"
'Example 2: "string1|string2|string3" and so.

'this function returns true, if any of items in serialized array has exact match with 'Stri' variable
'you can restrict search with LBound and UBound items only.
Public Function inArraySerialized( _
    Stri As String, _
    SerializedArray As String, _
    Delimiter As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    Dim MyArray() As String
    If 0 = Len(SerializedArray) Then
        If 0 = Len(Stri) Then inArraySerialized = True
        Exit Function
    End If
    MyArray = SplitSafe(SerializedArray, Delimiter)
    If lB = -2147483647 Then lB = LBound(MyArray)   'some trick
    If uB = 2147483647 Then uB = UBound(MyArray)    'Thanks to Казанский :)
    Dim i As Long
    For i = lB To uB
        If StrComp(Stri, MyArray(i), CompareMethod) = 0 Then inArraySerialized = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg err, "inArraySerialized", "SerializedString: ", SerializedArray, "delim: ", Delimiter
    If inIDE Then Stop: Resume Next
End Function

'The same as Split(), except of proper error handling when source data is empty string and you assign result to variable defined as array.
'So, in case of empty string it return array with 0 items.
'Also: return type is 'string()' instead of 'variant()'
'
'Warning note: Do not use this function in For each statement !!! - use default Split() instead:
'Differences in behavior:
'Split() with empty string cause 'For each' to not execute any its cycles at all.
'SplitSafe() cause to execute 'For Each' for a 1 cycle with empty value.
Public Function SplitSafe(sComplexString As String, Optional Delimiter As String = " ") As String()
    If 0 = Len(sComplexString) Then
        ReDim ret(0) As String
        SplitSafe = ret
    Else
        SplitSafe = Split(sComplexString, Delimiter)
    End If
End Function

'get the first item of serilized array
Public Function SplitExGetFirst(sSerializedArray As String, Optional Delimiter As String = " ") As String
    SplitExGetFirst = SplitSafe(sSerializedArray, Delimiter)(0)
End Function

'get the last item of serialized array
Public Function SplitExGetLast(sSerializedArray As String, Optional Delimiter As String = " ") As String
    Dim ret() As String
    ret = SplitSafe(sSerializedArray, Delimiter)
    SplitExGetLast = ret(UBound(ret))
End Function

Private Sub DeleteDuplicatesInArray(Arr() As String, CompareMethod As VbCompareMethod, Optional DontCompress As Boolean)
    On Error GoTo ErrorHandler:
    
    'DontCompress:
    'if true, do not move items:
    'function will return array with empty items in places where duplicate match were found
    'so, its structure will be similar to the source array
    
    'if false, returns new reconstructed array:
    'all subsequent array items are shifted to the item where duplicate was found.
    
    Dim i   As Long
    
    If DontCompress Then
        For i = LBound(Arr) To UBound(Arr)
            If inArray(Arr(i), Arr, i + 1, UBound(Arr), CompareMethod) Then
                Arr(i) = vbNullString
            End If
        Next
    Else
        Dim TmpArr() As String
        ReDim TmpArr(LBound(Arr) To UBound(Arr))
        Dim Cnt As Long
        Cnt = LBound(Arr)
        For i = LBound(Arr) To UBound(Arr)
            If Not inArray(Arr(i), Arr, i + 1, UBound(Arr), CompareMethod) Then
                TmpArr(Cnt) = Arr(i)
                Cnt = Cnt + 1
            End If
        Next
        ReDim Preserve TmpArr(LBound(TmpArr) To Cnt - 1)
        Arr = TmpArr
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg err, "DeleteDuplicatesInArray"
    If inIDE Then Stop: Resume Next
End Sub

Public Function StrBeginWith(Text As String, BeginPart As String) As Boolean
    StrBeginWith = (StrComp(Left$(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function

Public Function StrEndWith(Text As String, LastPart As String) As Boolean
    StrEndWith = (StrComp(Right$(Text, Len(LastPart)), LastPart, 1) = 0)
End Function

Public Function StrEndWithParamArray(Text As String, ParamArray LastPart()) As Boolean
    Dim i As Long
    For i = 0 To UBound(LastPart)
        If Len(LastPart(i)) <> 0 Then
            If StrComp(Right$(Text, Len(LastPart(i))), LastPart(i), 1) = 0 Then
                StrEndWithParamArray = True
                Exit For
            End If
        End If
    Next
End Function

Public Function StrBeginWithArray(Text As String, BeginPart() As String) As Boolean
    Dim i As Long
    For i = 0 To UBound(BeginPart)
        If Len(BeginPart(i)) <> 0 Then
            If StrComp(Left$(Text, Len(BeginPart(i))), BeginPart(i), 1) = 0 Then
                StrBeginWithArray = True
                Exit For
            End If
        End If
    Next
End Function

Public Function FindOnPath(sAppName As String) As String
    On Error GoTo ErrorHandler:

    Static Exts
    Static IsInit As Boolean
    Dim ProcPath$
    Dim sFile As String
    Dim sFolder As String
    Dim pos As Long
    Dim i As Long
    Dim FoundFile As Boolean
    Dim sFileTry As String
    
    If Not IsInit Then
        IsInit = True
        Exts = Split(EnvironW("%PathExt%"), ";")
        For i = 0 To UBound(Exts)
            Exts(i) = LCase(Exts(i))
        Next
    End If
    
    If FileExists(sAppName) Then
        FindOnPath = sAppName
        Exit Function
    End If
    
    pos = InStrRev(sAppName, "\")
    
    If pos <> 0 Then
        sFolder = Left$(sAppName, pos - 1)
        sFile = Mid$(sAppName, pos + 1)
        
        For i = 0 To UBound(Exts)
            sFileTry = sFolder & "\" & sFile & Exts(i)
            
            If FileExists(sFileTry) Then
                FindOnPath = sFileTry
                Exit Function
            End If
        Next
    Else
        ToggleWow64FSRedirection False
    
        ProcPath = Space$(MAX_PATH)
        LSet ProcPath = sAppName & vbNullChar
        
        If CBool(PathFindOnPath(StrPtr(ProcPath), 0&)) Then
            FindOnPath = TrimNull(ProcPath)
        Else
            'go through the extensions list
            
            For i = 0 To UBound(Exts)
                sFileTry = sAppName & Exts(i)
            
                ProcPath = Space$(MAX_PATH)
                LSet ProcPath = sFileTry & vbNullChar
            
                If CBool(PathFindOnPath(StrPtr(ProcPath), 0&)) Then
                    FindOnPath = TrimNull(ProcPath)
                    Exit For
                End If
            
            Next
            
        End If
        
        ToggleWow64FSRedirection True
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "FindOnPath", "AppName: ", sAppName
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Function

Public Sub CenterForm(myForm As Form) ' Центрирование формы на экране с учетом системных панелей
    On Error Resume Next
    Dim Left    As Long
    Dim Top     As Long
    Left = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN) / 2 - myForm.Width / 2
    Top = Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYFULLSCREEN) / 2 - myForm.Height / 2
    myForm.Move Left, Top
End Sub

Public Function GetFileNameAndExt(Path As String) As String ' вернет только имя файла вместе с расширением
    Dim pos As Long
    pos = InStrRev(Path, "\")
    If pos <> 0 Then
        GetFileNameAndExt = Mid$(Path, pos + 1)
    Else
        GetFileNameAndExt = Path
    End If
End Function

' Получить только имя файла (без расширения имени)
Public Function GetFileName(Path As String) As String
    On Error GoTo ErrorHandler
    Dim posDot      As Long
    Dim posSl       As Long
    
    posSl = InStrRev(Path, "\")
    If posSl <> 0 Then
        posDot = InStrRev(Path, ".")
        If posDot < posSl Then posDot = 0
    Else
        posDot = InStrRev(Path, ".")
    End If
    If posDot = 0 Then posDot = Len(Path) + 1
    
    GetFileName = Mid$(Path, posSl + 1, posDot - posSl - 1)
    Exit Function
ErrorHandler:
    ErrorMsg err, "Parser.GetFileName", "Path: ", Path
End Function

'true if success
Public Function FileCopyW(FileSource As String, FileDestination As String, Optional bOverwrite As Boolean = True) As Boolean
    ToggleWow64FSRedirection False, FileSource
    ToggleWow64FSRedirection False, FileDestination
    FileCopyW = CopyFile(StrPtr(FileSource), StrPtr(FileDestination), Not bOverwrite)
    ToggleWow64FSRedirection True
End Function

Public Function ConvertVersionToNumber(sVersion As String) As Long  '"1.1.1.1" -> 1 number (all fields should be < 100)
    On Error Resume Next
    Dim Ver() As String
    
    If 0 = Len(sVersion) Then Exit Function
    
    Ver = Split(sVersion, ".")
    If UBound(Ver) = 3 Then
        ConvertVersionToNumber = Ver(3) + Ver(2) * 100& + Ver(1) * 10000& + Ver(0) * CLng(1000000)
    End If
End Function

Public Sub UpdatePolicy(Optional noWait As Boolean)
    Dim GPUpdatePath$
    If bIsWin64 And FolderExists(sWinDir & "\sysnative") And OSver.MajorMinor >= 6 Then
        GPUpdatePath = sWinDir & "\sysnative\gpupdate.exe"
    Else
        GPUpdatePath = sWinDir & "\system32\gpupdate.exe"
    End If
    If Proc.ProcessRun(GPUpdatePath, "/force", , vbHide) Then
        If Not noWait Then
            Proc.WaitForTerminate , , , 15000
        End If
    End If
End Sub

Public Sub ConcatArrays(DestArray() As String, AddArray() As String)
    'Appends AddArray() to the end of DestArray.
    'DestArray() should be declared as dynamic
    
    'UnInitialized arrays are permitted
    'Warning: if both arrays is uninitialized - DestArray() will remain the same (with uninitialized state)
    
    On Error GoTo ErrorHandler
    
    Dim i&, idx&
    
    If Not CBool(SafeArrayGetDim(AddArray)) Then Exit Sub
    If Not CBool(SafeArrayGetDim(DestArray)) Then
        idx = -1
        ReDim DestArray(UBound(AddArray) - LBound(AddArray))
    Else
        idx = UBound(DestArray)
        ReDim Preserve DestArray(UBound(DestArray) + (UBound(AddArray) - LBound(AddArray)) + 1)
    End If
    
    For i = LBound(AddArray) To UBound(AddArray)
        idx = idx + 1
        DestArray(idx) = AddArray(i)
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "Parser.ConcatArrays"
End Sub

Public Sub QuickSort(J, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler:
    Dim i As Long, l As Long, M As String, wsp As String
    i = low: l = high: M = J((i + l) \ 2)
    Do Until i > l: Do While J(i) < M: i = i + 1: Loop: Do While J(l) > M: l = l - 1: Loop
        If (i <= l) Then wsp = J(i): J(i) = J(l): J(l) = wsp: i = i + 1: l = l - 1
    Loop
    If low < l Then QuickSort J, low, l
    If i < high Then QuickSort J, i, high
    Exit Sub
ErrorHandler:
    ErrorMsg err, "QuickSort"
    If inIDE Then Stop: Resume Next
End Sub

' exclude items from ArraySrc() that is not match 'Mask' and save to 'ArrayDest()'
' return value is a number of items in 'ArrayDest'
' if number of items is 0, ArrayDest() will have 1 empty item.
Public Function FilterArray(ArraySrc() As String, ArrayDest() As String, Mask As String) As Long
    On Error GoTo ErrorHandler:
    Dim i As Long, J As Long
    ReDim ArrayDest(LBound(ArraySrc) To UBound(ArraySrc))
    For i = LBound(ArraySrc) To UBound(ArraySrc)
        If ArraySrc(i) Like Mask Then
            J = J + 1
            ArrayDest(LBound(ArraySrc) + J - 1) = ArraySrc(i)
        End If
    Next
    If J = 0 Then
        ReDim ArrayDest(LBound(ArraySrc) To LBound(ArraySrc))
    Else
        ReDim Preserve ArrayDest(LBound(ArraySrc) To LBound(ArraySrc) + J - 1)
    End If
    FilterArray = J
    Exit Function
ErrorHandler:
    ErrorMsg err, "FilterArray"
    If inIDE Then Stop: Resume Next
End Function

'get a substring starting at the specified character (search begins with the end of the line)
Public Function MidFromCharRev(sText As String, Delimiter As String) As String
    On Error GoTo ErrorHandler:
    Dim iPos As Long
    If 0 <> Len(sText) Then
        iPos = InStrRev(sText, Delimiter)
        If iPos <> 0 Then
            MidFromCharRev = Mid$(sText, iPos + 1)
        Else
            MidFromCharRev = ""
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "MidFromCharRev"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetCollectionKey(ByVal Index As Long, Col As Collection) As String ' Thanks to 'The Trick' (А. Кривоус) for this code
    On Error GoTo ErrorHandler:
    Dim lpSTR As Long, ptr As Long, Key As String
    If Col Is Nothing Then Exit Function
    Select Case Index
    Case Is < 1, Is > Col.Count: Exit Function
    Case Else
        ptr = ObjPtr(Col)
        Do While Index
            GetMem4 ByVal ptr + 24, ptr
            Index = Index - 1
        Loop
    End Select
    lpSTR = StrPtr(Key)
    GetMem4 ByVal ptr + 16, ByVal VarPtr(Key)
    GetCollectionKey = Key
    GetMem4 lpSTR, ByVal VarPtr(Key)
    Exit Function
ErrorHandler:
    ErrorMsg err, "GetCollectionKey"
    If inIDE Then Stop: Resume Next
End Function

Public Function isCollectionKeyExists(Key As String, Col As Collection) As Boolean
    Dim i As Long
    For i = 1 To Col.Count
        If GetCollectionKey(i, Col) = Key Then isCollectionKeyExists = True: Exit For
    Next
End Function

Public Sub GetProfiles()    'result -> in global variable 'colProfiles' (collection)
    On Error GoTo ErrorHandler:
    
    'include all folders inside <c:\users>
    'without 'Public'
    
    Dim ProfileListKey      As String
    Dim ProfilesDirectory   As String
    Dim ProfileSubKey()     As String
    Dim ProfilePath         As String
    Dim SubFolders()        As String
    'Dim UserProfile         As String
    Dim i                   As Long
    Dim lr                  As Long
    Dim Path                As String
    Dim objFolder           As Variant
    
    ProfileListKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    ProfilesDirectory = GetRegData(0&, ProfileListKey, "ProfilesDirectory")

    If RegEnumSubkeysToArray(0&, ProfileListKey, ProfileSubKey()) > 0 Then
        For i = 1 To UBound(ProfileSubKey)
            If Not (ProfileSubKey(i) = "S-1-5-18" Or _
                    ProfileSubKey(i) = "S-1-5-19" Or _
                    ProfileSubKey(i) = "S-1-5-20") Then
                
                ProfilePath = GetRegData(0&, ProfileListKey & "\" & ProfileSubKey(i), "ProfileImagePath")
                
                If Len(ProfilePath) <> 0 Then
                    If FolderExists(ProfilePath) Then
                        If Not isCollectionKeyExists(ProfilePath, colProfiles) Then
                            colProfiles.Add ProfilePath, ProfilePath
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    'UserProfile = EnvironW("%UserProfile%")
    
    'добавляю папки, которые находятся в подкаталоге (на 1 уровень ниже) профиля текущего пользователя
    
    If Len(UserProfile) <> 0 Then
        If FolderExists(UserProfile) Then
            Path = UserProfile
            lr = PathRemoveFileSpec(StrPtr(Path))   ' get Parent directory
            If lr Then Path = Left$(Path, lstrlen(StrPtr(Path)))

            SubFolders() = ListSubfolders(Path)

            If CBool(SafeArrayGetDim(SubFolders)) Then
                For Each objFolder In SubFolders()
                    If Len(objFolder) <> 0 And Not (StrEndWith(CStr(objFolder), "\Public") And OSver.MajorMinor >= 6) Then
                        If FolderExists(CStr(objFolder)) Then
                            If Not isCollectionKeyExists(CStr(objFolder), colProfiles) Then
                                colProfiles.Add CStr(objFolder), CStr(objFolder)
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "GetProfiles"
    If inIDE Then Stop: Resume Next
End Sub

Public Function UnpackResource(ResourceID As Long, DestinationPath As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim ff      As Integer
    Dim b()     As Byte
    UnpackResource = True
    b = LoadResData(ResourceID, "CUSTOM")
    ff = FreeFile
    Open DestinationPath For Binary Access Write As #ff
        Put #ff, , b
    Close #ff
    Exit Function
ErrorHandler:
    ErrorMsg err, "UnpackResource", "ID: " & ResourceID, "Destination path: " & DestinationPath
    UnpackResource = False
    If inIDE Then Stop: Resume Next
End Function

Public Sub Terminate_HJT()
    Unload frmMain
    End
End Sub

Public Sub AddHorizontalScrollBarToResults(lstControl As ListBox)
    'Adds a horizontal scrollbar to the results display if it is needed (after the scan)
    Dim X As Long, s$
    Dim listLength As Long
    With lstControl
        For listLength = 0 To .ListCount - 1
            s = Replace$(.List(listLength), vbTab, "12345678")
            If .Width < frmMain.TextWidth(s) + 1000 And X < frmMain.TextWidth(s) + 1000 Then
                X = frmMain.TextWidth(.List(listLength)) + 500
            End If
        Next
        If frmMain.ScaleMode = vbTwips Then X = X / Screen.TwipsPerPixelX + 50  ' if twips change to pixels (+50 to account for the width of the vertical scrollbar
        SendMessage .hWnd, LB_SETHORIZONTALEXTENT, X, ByVal 0&
    End With
End Sub
