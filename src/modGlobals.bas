Attribute VB_Name = "modGlobals"
Option Explicit

Public Const MAX_TIMEOUT_DEFAULT As Long = 180 'Standart timeout

Public Const g_AppName As String = "HiJackThis Fork"

Public Const g_Backup_Do_Every_Days As Long = 7
Public Const g_Backup_Erase_Every_Days As Long = 28

Public Const MAX_MODULE_NAME32 As Long = 255&

Public TaskBar As TaskbarLib.TaskbarList

Public Enum HE_HIVE
    HE_HIVE_ALL = 7
    HE_HIVE_HKLM = 1
    HE_HIVE_HKCU = 2
    HE_HIVE_HKU = 4
End Enum
#If False Then
    Dim HE_HIVE_ALL, HE_HIVE_HKLM, HE_HIVE_HKCU, HE_HIVE_HKU
#End If

Public Enum HE_SID
    HE_SID_ALL = 7
    HE_SID_DEFAULT = 1
    HE_SID_SERVICE = 2
    HE_SID_USER = 4
End Enum
#If False Then
    Dim HE_SID_ALL, HE_SID_DEFAULT, HE_SID_SERVICE, HE_SID_USER
#End If

Public Enum HE_REDIRECTION
    HE_REDIR_BOTH = 3
    HE_REDIR_NO_WOW = 1
    HE_REDIR_WOW = 2
    HE_REDIR_DONT_IGNORE_SHARED = 4
End Enum
#If False Then
    Dim HE_REDIR_BOTH, HE_REDIR_WOW, HE_REDIR_NO_WOW
#End If

Public Enum BUTTON_ALIGNMENT
    BS_CENTER = &H300&
    BS_LEFT = &H100&
    BS_RIGHT = &H200&
    BS_TOP = &H400&
End Enum

Public Type LVITEMW
    Mask        As Long
    iItem       As Long
    iSubItem    As Long
    State       As Long
    stateMask   As Long
    pszText     As Long
    cchTextMax  As Long
    iImage      As Long
    lParam      As Long
    iIndent     As Long
End Type

'Public Type LVITEMW_64
'    mask        As Long
'    iItem       As Long
'    iSubItem    As Long
'    state       As Long
'    stateMask   As Long
'    align1      As Long
'    pszText     As Currency
'    cchTextMax  As Long
'    iImage      As Long
'    lParam      As Currency
'    iIndent     As Long
'    align2      As Long
'End Type

'frmEULA

Public Type tagINITCOMMONCONTROLSEX
    dwSize  As Long
    dwICC   As Long
End Type

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Public Declare Function SetCurrentProcessExplicitAppUserModelID Lib "shell32.dll" (ByVal pAppID As Long) As Long
'Public Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
'Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
'Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Public Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

'frmMain
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Public Declare Function CreateMutex Lib "kernel32.dll" Alias "CreateMutexW" (ByVal lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As Long) As Long
Public Declare Function SetWindowTheme Lib "UxTheme.dll" (ByVal hwnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Public Declare Function MessageBeep Lib "user32.dll" (ByVal uType As Long) As Long
'Public Declare Sub CloseHandle Lib "kernel32.dll" (ByVal Handle As Long)
'Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function ILCreateFromPath Lib "shell32.dll" Alias "ILCreateFromPathW" (ByVal pszPath As Long) As Long
Public Declare Function SHOpenFolderAndSelectItems Lib "shell32.dll" (ByVal pidlFolder As Long, ByVal cidl As Long, ByVal apidl As Long, ByVal dwFlags As Long) As Long
Public Declare Sub ILFree Lib "shell32.dll" (ByVal pIDL As Long)
'Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal p_Hmodule As Long) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32.dll" Alias "SetCurrentDirectoryW" (ByVal lpPathName As Long) As Long

'modmain
Public Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)
Public Declare Function GetEnvironmentStrings Lib "kernel32.dll" Alias "GetEnvironmentStringsW" () As Long
Public Declare Function FreeEnvironmentStrings Lib "kernel32.dll" Alias "FreeEnvironmentStringsW" (ByVal lpszEnvironmentBlock As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal uCmd As Long) As Long
Public Declare Function LoadImageW Lib "user32.dll" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Public Declare Function SendMessageW Lib "user32.dll" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SfcIsFileProtected Lib "Sfc.dll" (ByVal RpcHandle As Long, ByVal ProtFileName As Long) As Long

Public Const GW_OWNER          As Long = 4
Public Const IMAGE_ICON        As Long = 1
Public Const ICON_SMALL        As Long = 0
Public Const ICON_BIG          As Long = 1
Public Const LR_DEFAULTSIZE    As Long = &H40
Public Const SM_CXICON         As Long = 11
Public Const SM_CYICON         As Long = 12
Public Const SM_CXSMICON       As Long = 49
Public Const SM_CYSMICON       As Long = 50
Public Const WM_SETICON        As Long = &H80

Public HE           As clsHiveEnum
Public Reg          As clsRegistry

Public colSafeDNS   As New Collection
Public colSafeCert  As New Collection
Public colBadCert   As New Collection
Public cReg4vals    As New Collection
Public sRegVals()   As String
Public sFileVals()  As String

Public bAutoSelect      As Boolean
Public bConfirm         As Boolean
Public bMakeBackup      As Boolean
Public bAdditional      As Boolean
Public bShowSRP         As Boolean
'Public bIgnoreSafeDomains      As Boolean
Public bLogProcesses    As Boolean
Public bSkipErrorMsg    As Boolean
Public bMinToTray       As Boolean
Public bStartupListSilent As Boolean
Public bScanExecuted    As Boolean
Public bCryptDisable    As Boolean
Public bPolymorph       As Boolean
Public bCheckForUpdates As Boolean
Public bFirstRun        As Boolean
Public bFirstRebootScan As Boolean
Public gNotUserClick    As Boolean
Public gNoGUI           As Boolean
Public g_WER_Disabled   As Boolean
Public g_HwndMain       As Long
Public g_NeedTerminate  As Boolean

Public sHostsFile$

Public bIsWin9x As Boolean
Public bIsWinNT As Boolean
Public bIsWinME As Boolean
Public bIsWin2k As Boolean
Public bIsWinXP As Boolean
Public bIsWinVistaAndNewer As Boolean
Public bIsWin7AndNewer As Boolean

Public bIsWin64 As Boolean
Public bIsWOW64 As Boolean
Public bIsWin32 As Boolean
Public inIDE    As Boolean
Public bForceRU As Boolean
Public bForceEN As Boolean

Public SysDisk          As String
Public sWinDir          As String
Public sSysNativeDir    As String
Public sSysDir          As String
Public sWinSysDir       As String
Public sWinSysDirWow64  As String
Public PF_32            As String
Public PF_64            As String
Public PF_32_Common     As String
Public PF_64_Common     As String
Public StartMenuPrograms As String
Public AppData          As String
Public AppDataLocalLow  As String
Public LocalAppData     As String
Public Desktop          As String
Public UserProfile      As String
Public AllUsersProfile  As String
Public TempCU           As String
Public envCurUser       As String
Public ProgramData      As String
Public colProfiles      As Collection

Public sWinVersion      As String
Public bRebootRequired  As Boolean
Public bUpdatePolicyNeeded As Boolean
Public DisableSubclassing As Boolean
Public isRanHJT_Scan    As Boolean
Public bmnuExit_Clicked As Boolean

Public bIsUSADateFormat As Boolean
Public bNoWriteAccess   As Boolean
Public bSeenLSPWarning  As Boolean

Public sSafeLSPFiles        As String
Public aSafeProtocols()     As String
Public sSafeProtocols       As String
Public aSafeRegDomains()    As String
Public aSafeSSODL()         As String
Public aSafeSIOI()          As String
Public aSafeSEH()           As String
Public aSafeFilters()       As String
Public sSafeFilters         As String
Public sSafeAppInit         As String
Public sSafeWinlogonNotify  As String
Public sSafeIfeVerifier     As String

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
Public g_CurrentLang        As String

Public ErrLogCustomText As clsStringBuilder
Public bDebugMode As Boolean
Public bDebugToFile As Boolean
Public bScanMode As Boolean
Public hDebugLog As Long
Public hLog     As Long

Public gSIDs() As String, gUsers() As String, gHives() As String

Public tim() As clsTimer

Public Const LB_ITEMFROMPOINT  As Long = &H1A9&

'
' ---------------------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------
'

'modFile

Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
 
Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
 
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    lpszFileName(MAX_PATH - 1) As Integer
    lpszAlternate(13) As Integer
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Public Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Public Type SHFILEOPSTRUCT
    hwnd                    As Long
    wFunc                   As Long
    pFrom                   As Long
    pTo                     As Long
    fFlags                  As Integer
    fAnyOperationsAborted   As Long
    hNameMappings           As Long
    lpszProgressTitle       As Long
End Type

Public Type SHELLEXECUTEINFO
    cbSize          As Long
    fMask           As Long
    hwnd            As Long
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

Public Declare Function CreateTransaction Lib "KtmW32.dll" (ByVal lpTransactionAttributes As Long, ByVal UOW As Long, ByVal CreateOptions As Long, ByVal IsolationLevel As Long, ByVal IsolationFlags As Long, ByVal Timeout As Long, ByVal Description As Long) As Long
Public Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingW" (ByVal hFile As Long, ByVal lpAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
Public Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Long
Public Declare Function PathFileExists Lib "Shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Public Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CreateFileTransacted Lib "kernel32.dll" Alias "CreateFileTransactedW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long, ByVal hTransaction As Long, ByVal pusMiniVersion As Long, ByVal pExtendedParameter As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function SHFileExists Lib "shell32.dll" Alias "#45" (ByVal szPath As String) As Long
Public Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
Public Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Long
Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeW" (ByVal nDrive As Long) As Long
Public Declare Function GetLogicalDrives Lib "kernel32.dll" () As Long
Public Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Public Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStringDest As Long, ByVal lpStringSrc As Long) As Long
Public Declare Function GetLongPathNameW Lib "kernel32.dll" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As Long, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Long, puLen As Long) As Long
Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
'Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal lSize As Long)
Public Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Public Declare Function PathIsNetworkPath Lib "Shlwapi.dll" Alias "PathIsNetworkPathW" (ByVal pszPath As Long) As Long
Public Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, ByVal lpOutBuffer As Long, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bDontOverwrite As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationW" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameW" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExW" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function MoveFileEx Lib "kernel32.dll" Alias "MoveFileExW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
Public Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathW" (ByVal hWndOwner As Long, ByVal CSIDL As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As Long) As Long
Public Declare Function SHGetKnownFolderPath Lib "shell32.dll" (rfid As UUID, ByVal dwFlags As Long, ByVal hToken As Long, ppszPath As Long) As Long
Public Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal iMaxLength As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszGuid As Long, pGuid As UUID) As Long
Public Declare Function PathFindOnPath Lib "Shlwapi.dll" Alias "PathFindOnPathW" (ByVal pszFile As Long, ppszOtherDirs As Any) As Boolean

Public Const FILE_SHARE_READ           As Long = &H1&
Public Const FILE_SHARE_WRITE          As Long = &H2&
Public Const FILE_SHARE_DELETE         As Long = 4&
Public Const FILE_READ_ATTRIBUTES      As Long = &H80&
Public Const OPEN_EXISTING             As Long = 3&
Public Const CREATE_ALWAYS             As Long = 2&
Public Const GENERIC_READ              As Long = &H80000000
Public Const GENERIC_WRITE             As Long = &H40000000
Public Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&
Public Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF
Public Const ERROR_SUCCESS             As Long = 0&
Public Const INVALID_FILE_ATTRIBUTES   As Long = -1&
Public Const NO_ERROR                  As Long = 0&
Public Const FILE_BEGIN                As Long = 0&
Public Const FILE_CURRENT              As Long = 1&
Public Const FILE_END                  As Long = 2&
Public Const INVALID_SET_FILE_POINTER  As Long = &HFFFFFFFF
Public Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Public Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400&
Public Const ERROR_HANDLE_EOF          As Long = 38&
Public Const SEC_IMAGE                 As Long = &H1000000
Public Const PAGE_READONLY             As Long = 2&
Public Const FILE_MAP_READ             As Long = 4&

Public Const KEY_QUERY_VALUE           As Long = &H1&
Public Const RegType_DWord             As Long = 4&

Public Const MOVEFILE_DELAY_UNTIL_REBOOT As Long = &H4&

Public Const IOCTL_STORAGE_CHECK_VERIFY2   As Long = &H2D0800
Public Const IOCTL_STORAGE_CHECK_VERIFY    As Long = &H2D4800

'modHash
Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextW" (ByRef phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptHashData_Array Lib "advapi32.dll" Alias "CryptHashData" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptHashData_Str Lib "advapi32.dll" Alias "CryptHashData" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, ByVal pbData As Long, pdwDataLen As Long, ByVal dwFlags As Long) As Long

Public Const ALG_TYPE_ANY As Long = 0
Public Const ALG_SID_MD5 As Long = 3
Public Const ALG_SID_SHA1 As Long = 4
Public Const ALG_CLASS_HASH As Long = 32768

Public Const HP_HASHVAL As Long = 2
Public Const HP_HASHSIZE As Long = 4

Public Const CRYPT_VERIFYCONTEXT = &HF0000000

Public Const PROV_RSA_FULL As Long = 1
Public Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"

'modInternet

Public Const MAX_HOSTNAME_LEN = 132&
Public Const MAX_DOMAIN_NAME_LEN = 132&
Public Const MAX_SCOPE_ID_LEN = 260&

Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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
    Flags As Long
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

Public Type IP_ADDR_STRING
    Next As Long
    IpAddress(15) As Byte
    IpMask(15) As Byte
    Context As Long
End Type

Public Type FIXED_INFO
    HostName(MAX_HOSTNAME_LEN - 1) As Byte
    DomainName(MAX_DOMAIN_NAME_LEN - 1) As Byte
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId(MAX_SCOPE_ID_LEN - 1) As Byte
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type

Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectW" (ByVal InternetSession As Long, ByVal sServerName As Long, ByVal nServerPort As Integer, ByVal sUsername As Long, ByVal sPassword As Long, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenW" (ByVal sAgent As Long, ByVal lAccessType As Long, ByVal sProxyName As Long, ByVal sProxyBypass As Long, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlW" (ByVal hInternetSession As Long, ByVal sURL As Long, ByVal sHeaders As Long, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, Buffer As Any, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Public Declare Function InternetReadFileString Lib "wininet.dll" Alias "InternetReadFile" (ByVal hFile As Long, ByVal Buffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestW" (ByVal hHttpSession As Long, ByVal sVerb As Long, ByVal sObjectName As Long, ByVal sVersion As Long, ByVal sReferer As Long, lplpszAcceptTypes As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestW" (ByVal hHttpRequest As Long, ByVal sHeaders As Long, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Long
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoW" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Public Declare Function GetNetworkParams Lib "IPHlpApi.dll" (FixedInfo As Any, pOutBufLen As Long) As Long

Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_OVERWRITEPROMPT = &H2

Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_FLAG_RELOAD = &H80000000

Public Const INTERNET_SERVICE_HTTP = 3
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000

Public Const ERROR_BUFFER_OVERFLOW = 111&

'modLSP


Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(257) As Byte
    szSystemStatus(129) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type WSANAMESPACE_INFO
    NSProviderId   As UUID
    dwNameSpace    As Long
    fActive        As Long
    dwVersion      As Long
    lpszIdentifier As Long
End Type

Public Type WSAPROTOCOLCHAIN
    ChainLen As Long
    ChainEntries(6) As Long
End Type

Public Type WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As UUID
    dwCatalogEntryId As Long
    ProtocolChain As WSAPROTOCOLCHAIN
    iVersion As Long
    iAddressFamily As Long
    iMaxSockAddr As Long
    iMinSockAddr As Long
    iSocketType As Long
    iProtocol As Long
    iProtocolMaxOffset As Long
    iNetworkByteOrder As Long
    iSecurityScheme As Long
    dwMessageSize As Long
    dwProviderReserved As Long
    szProtocol As String * 256
End Type

Public Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumValueW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegEnumKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegDeleteKeyW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
Public Declare Function RegCreateKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegQueryValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function SHRestartSystemMB Lib "shell32.dll" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Integer, ByVal lpWSAD As Long) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAEnumProtocols Lib "ws2_32.dll" Alias "WSAEnumProtocolsW" (ByVal lpiProtocols As Long, ByVal lpProtocolBuffer As Long, lpdwBufferLength As Long) As Long
Public Declare Function WSAEnumNameSpaceProviders Lib "ws2_32.dll" Alias "WSAEnumNameSpaceProvidersW" (lpdwBufferLength As Long, ByVal lpnspBuffer As Long) As Long
Public Declare Function WSCGetProviderPath Lib "ws2_32.dll" (ByVal lpProviderId As Long, ByVal lpszProviderDllPath As Long, ByVal lpProviderDllPathLen As Long, ByVal lpErrno As Long) As Long
Public Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As UUID, ByVal lpsz As Long, ByVal cchMax As Long) As Long

Public Const SOCKET_ERROR As Long = -1
Public Const REG_OPTION_NON_VOLATILE As Long = 0

'modMain

Public Const MAX_NAME = 256&
Public Const LB_SETHORIZONTALEXTENT    As Long = &H194
Public Const HWND_TOPMOST As Long = -1&
Public Const HWND_NOTOPMOST As Long = -2&
Public Const SWP_NOMOVE As Long = 2&
Public Const SWP_NOSIZE As Long = 1&

Public Type SAFEARRAY
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
End Type

Public Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

Public Type SID_AND_ATTRIBUTES
    SID         As Long
    Attributes  As Long
End Type

Public Type TOKEN_GROUPS
    GroupCount As Long
    Groups(20) As SID_AND_ATTRIBUTES
End Type

Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameW" (ByVal lpBuffer As Long, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameW" (ByVal lpBuffer As Long, nSize As Long) As Long
Public Declare Function GetDateFormat Lib "kernel32.dll" Alias "GetDateFormatW" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As Long, ByVal lpDateStr As Long, ByVal cchDate As Long) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (lpFrequency As Any) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Any) As Long
Public Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function GetUserDefaultLCID Lib "kernel32.dll" () As Long
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
'Public Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Public Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
Public Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Public Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Public Declare Function IsValidSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
Public Declare Function EqualSid Lib "advapi32.dll" (ByVal pSid1 As Long, ByVal pSid2 As Long) As Long
Public Declare Function EqualPrefixSid Lib "advapi32.dll" (ByVal pSid1 As Long, ByVal pSid2 As Long) As Long
Public Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSid As Long)
Public Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
'Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, ByVal Arguments As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long
Public Declare Function PathRemoveFileSpec Lib "Shlwapi.dll" Alias "PathRemoveFileSpecW" (ByVal pszPath As Long) As Long
Public Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Public Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function EmptyClipboard Lib "user32.dll" () As Long
Public Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Public Declare Function GetMem2 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Public Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidW" (ByVal lpSystemName As Long, ByVal lpSid As Long, ByVal lpName As Long, cchName As Long, ByVal lpReferencedDomainName As Long, cchReferencedDomainName As Long, eUse As Long) As Long
Public Declare Function ConvertStringSidToSid Lib "advapi32.dll" Alias "ConvertStringSidToSidW" (ByVal StringSid As Long, pSid As Long) As Long
Public Declare Function IsBadReadPtr Lib "kernel32.dll" (ByVal lp As Long, ByVal ucb As Long) As Long


Public Const CREATE_NEW                As Long = 1&

Public Const SPI_SETDESKWALLPAPER  As Long = 20&
Public Const SPIF_SENDWININICHANGE As Long = &H2&
Public Const SPIF_UPDATEINIFILE    As Long = &H1&

Public Const CF_UNICODETEXT    As Long = 13&
Public Const GMEM_MOVEABLE     As Long = &H2&
Public Const CF_LOCALE         As Long = 16

Public Const SECURITY_NT_AUTHORITY         As Long = &H5&
Public Const TOKEN_QUERY                   As Long = &H8&
Public Const TokenGroups                   As Long = 2&
Public Const SECURITY_BUILTIN_DOMAIN_RID   As Long = &H20&
Public Const DOMAIN_ALIAS_RID_ADMINS       As Long = &H220&
Public Const DOMAIN_ALIAS_RID_USERS        As Long = &H221&
Public Const DOMAIN_ALIAS_RID_GUESTS       As Long = &H222&
Public Const DOMAIN_ALIAS_RID_POWER_USERS  As Long = &H223&
Public Const DOMAIN_ALIAS_RID_ACCOUNT_OPS  As Long = 548&
Public Const DOMAIN_ALIAS_RID_SYSTEM_OPS   As Long = 549&
Public Const DOMAIN_ALIAS_RID_PRINT_OPS    As Long = 550&
Public Const DOMAIN_ALIAS_RID_BACKUP_OPS   As Long = 551&

Public Const ERROR_NONE_MAPPED As Long = 1332&

Public Const FO_MOVE               As Long = &H1&
Public Const FO_COPY               As Long = &H2&
Public Const FO_DELETE             As Long = &H3&
Public Const FOF_NOCONFIRMATION    As Long = &H10&
Public Const FOF_SILENT            As Long = &H4&

Public Const SM_CLEANBOOT          As Long = &H43&

Public Const FILE_ATTRIBUTE_READONLY   As Long = 1&

Public Const SEE_MASK_INVOKEIDLIST     As Long = &HC&
Public Const SEE_MASK_NOCLOSEPROCESS   As Long = &H40&
Public Const SEE_MASK_FLAG_NO_UI       As Long = &H400

Public Const DEFAULT_CHARSET           As Long = 1&
Public Const SYMBOL_CHARSET            As Long = 2&
Public Const SHIFTJIS_CHARSET          As Long = 128&
Public Const HANGEUL_CHARSET           As Long = 129&
Public Const CHINESEBIG5_CHARSET       As Long = 136&
Public Const CHINESESIMPLIFIED_CHARSET As Long = 134&

Public Const VER_PLATFORM_WIN32s        As Long = 0&
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Public Const VER_PLATFORM_WIN32_NT      As Long = 2&

Public Const SM_CXFULLSCREEN       As Long = 16&
Public Const SM_CYFULLSCREEN       As Long = 17&

Public Const KEY_WOW64_64KEY       As Long = &H100&
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8

'modPermissions

Public Type TOKEN_PRIVILEGES
    PrivilegeCount  As Long
    LuidLowPart     As Long
    LuidHighPart    As Long
    Attributes      As Long
End Type

Public Enum SECURITY_INFORMATION                       'required access - to query / to set info:
    ATTRIBUTE_SECURITY_INFORMATION = &H20&              'query: READ_CONTROL; set: WRITE_DAC
    BACKUP_SECURITY_INFORMATION = &H10000               'query: READ_CONTROL and ACCESS_SYSTEM_SECURITY; set: WRITE_DAC and WRITE_OWNER and ACCESS_SYSTEM_SECURITY
    DACL_SECURITY_INFORMATION = 4                       'query: READ_CONTROL; set: WRITE_DAC
    GROUP_SECURITY_INFORMATION = 2                      'query: READ_CONTROL; set: WRITE_OWNER
    LABEL_SECURITY_INFORMATION = 16                     'query: READ_CONTROL; set: WRITE_OWNER
    OWNER_SECURITY_INFORMATION = 1                      'query: READ_CONTROL; set: WRITE_OWNER
    PROTECTED_DACL_SECURITY_INFORMATION = &H80000000    'query: -; set: WRITE_DAC
    PROTECTED_SACL_SECURITY_INFORMATION = &H40000000    'query: -; set: ACCESS_SYSTEM_SECURITY
    SACL_SECURITY_INFORMATION = 8                       'query: ACCESS_SYSTEM_SECURITY; set: ACCESS_SYSTEM_SECURITY
    SCOPE_SECURITY_INFORMATION = &H40&                  'query: READ_CONTROL; set: ACCESS_SYSTEM_SECURITY
    UNPROTECTED_DACL_SECURITY_INFORMATION = &H20000000  'query: -; set: WRITE_DAC
    UNPROTECTED_SACL_SECURITY_INFORMATION = &H10000000  'query: -; set: ACCESS_SYSTEM_SECURITY
End Enum

Public Enum SE_OBJECT_TYPE
    SE_UNKNOWN_OBJECT_TYPE = 0
    SE_FILE_OBJECT
    SE_SERVICE
    SE_PRINTER
    SE_REGISTRY_KEY
    SE_LMSHARE
    SE_KERNEL_OBJECT
    SE_WINDOW_OBJECT
    SE_DS_OBJECT
    SE_DS_OBJECT_ALL
    SE_PROVIDER_DEFINED_OBJECT
    SE_WMIGUID_OBJECT
    SE_REGISTRY_WOW64_32KEY
End Enum

Public Enum ACL_INFORMATION_CLASS
    AclRevisionInformation = 1
    AclSizeInformation
End Enum

Public Declare Sub GetNativeSystemInfo Lib "kernel32.dll" (ByVal lpSystemInfo As Long)
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueW" (ByVal lpSystemName As Long, ByVal lpName As Long, lpLuid As Long) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function CopySid Lib "advapi32.dll" (ByVal nDestinationSidLength As Long, ByVal pDestinationSid As Long, ByVal pSourceSid As Long) As Long
Public Declare Function GetLengthSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
Public Declare Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal Handle As Long, ByVal RequestedInformation As SECURITY_INFORMATION, ByVal pSecurityDescriptor As Long, ByVal nLength As Long, ByVal lpnLengthNeeded As Long) As Long
Public Declare Function MakeAbsoluteSD Lib "advapi32.dll" (ByVal pSelfRelativeSD As Long, ByVal pAbsoluteSD As Long, ByVal lpdwAbsoluteSDSize As Long, ByVal pDacl As Long, ByVal lpdwDaclSize As Long, ByVal pSacl As Long, ByVal lpdwSaclSize As Long, ByVal pOwner As Long, ByVal lpdwOwnerSize As Long, ByVal pPrimaryGroup As Long, ByVal lpdwPrimaryGroupSize As Long) As Long
Public Declare Function IsValidSecurityDescriptor Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long) As Long
Public Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclW" (ByVal cCountOfExplicitEntries As Long, ByVal pListOfExplicitEntries As Long, ByVal pOldAcl As Long, NewAcl As Long) As Long
Public Declare Function SetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As SECURITY_INFORMATION, ByVal psidOwner As Long, ByVal psidGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long) As Long
Public Declare Function SetNamedSecurityInfo Lib "advapi32.dll" Alias "SetNamedSecurityInfoW" (ByVal pObjectName As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ByVal psidOwner As Long, ByVal psidGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long) As Long
Public Declare Function GetAclInformation Lib "advapi32.dll" (ByVal pAcl As Long, ByVal pAclInformation As Long, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As ACL_INFORMATION_CLASS) As Long
Public Declare Function GetAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceIndex As Long, pAce As Long) As Long
Public Declare Function GetExplicitEntriesFromAcl Lib "advapi32.dll" Alias "GetExplicitEntriesFromAclW" (ByVal pAcl As Long, pcCountOfExplicitEntries As Long, pListOfExplicitEntries As Long) As Long
Public Declare Function DeleteAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceIndex As Long) As Long
Public Declare Function InitializeAcl Lib "advapi32.dll" (ByVal pAcl As Long, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Public Declare Function LocalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function IsValidAcl Lib "advapi32.dll" (ByVal pAcl As Long) As Long
Public Declare Function TreeResetNamedSecurityInfo Lib "advapi32.dll" Alias "TreeResetNamedSecurityInfoW" (ByVal pObjectName As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As SECURITY_INFORMATION, ByVal pOwner As Long, ByVal pGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long, ByVal KeepExplicit As Long, ByVal fnProgress As Long, ByVal ProgressInvokeSetting As Long, ByVal Args As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long

Public Const MAX_KEYNAME            As Long = 255&

Public Const REG_OPTION_BACKUP_RESTORE As Long = 4&
Public Const GENERIC_ALL            As Long = &H10000000
Public Const WRITE_DAC              As Long = &H40000
Public Const WRITE_OWNER            As Long = &H80000
Public Const READ_CONTROL           As Long = &H20000
Public Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Public Const SE_PRIVILEGE_ENABLED    As Long = 2&

Public Const OBJECT_INHERIT_ACE     As Long = 1&
Public Const CONTAINER_INHERIT_ACE  As Long = 2&

Public Const NO_MULTIPLE_TRUSTEE    As Long = 0&

Public Const ACCESS_DENIED_ACE_TYPE As Long = 1&

Public Const REG_CREATED_NEW_KEY    As Long = 1&

Public Const ERROR_MORE_DATA        As Long = 234&
Public Const ERROR_NO_TOKEN         As Long = 1008&

Public Const LMEM_FIXED             As Long = 0&
Public Const LMEM_ZEROINIT          As Long = &H40&

Public Const ACL_REVISION           As Long = 2&

'modProcess

Public Type CLIENT_ID
    UniqueProcess   As Long  ' HANDLE
    UniqueThread    As Long  ' HANDLE
End Type

Public Type UNICODE_STRING
    Length          As Integer
    MaximumLength   As Integer
    Buffer          As Long
End Type

Public Type VM_COUNTERS
    PeakVirtualSize             As Long
    VirtualSize                 As Long
    PageFaultCount              As Long
    PeakWorkingSetSize          As Long
    WorkingSetSize              As Long
    QuotaPeakPagedPoolUsage     As Long
    QuotaPagedPoolUsage         As Long
    QuotaPeakNonPagedPoolUsage  As Long
    QuotaNonPagedPoolUsage      As Long
    PagefileUsage               As Long
    PeakPagefileUsage           As Long
End Type

Public Type IO_COUNTERS
    ReadOperationCount      As Currency 'ULONGLONG
    WriteOperationCount     As Currency
    OtherOperationCount     As Currency
    ReadTransferCount       As Currency
    WriteTransferCount      As Currency
    OtherTransferCount      As Currency
End Type

Public Enum KWAIT_REASON
    Executive = 0
    FreePage = 1
    PageIn = 2
    PoolAllocation = 3
    DelayExecution = 4
    Suspended = 5
    UserRequest = 6
    WrExecutive = 7
    WrFreePage = 8
    WrPageIn = 9
    WrPoolAllocation = 10
    WrDelayExecution = 11
    WrSuspended = 12
    WrUserRequest = 13
    WrEventPair = 14
    WrQueue = 15
    WrLpcReceive = 16
    WrLpcReply = 17
    WrVirtualMemory = 18
    WrPageOut = 19
    WrRendezvous = 20
    Spare2 = 21
    Spare3 = 22
    Spare4 = 23
    Spare5 = 24
    WrCalloutStack = 25
    WrKernel = 26
    WrResource = 27
    WrPushLock = 28
    WrMutex = 29
    WrQuantumEnd = 30
    WrDispatchInt = 31
    WrPreempted = 32
    WrYieldExecution = 33
    WrFastMutex = 34
    WrGuardedMutex = 35
    WrRundown = 36
    MaximumWaitReason = 37
End Enum

Public Enum KTHREAD_STATE
    Initialized = 0
    Ready = 1
    Running = 2
    Standby = 3
    Terminated = 4
    Waiting = 5
    Transition = 6
    DeferredReady = 7
    GateWait = 8
End Enum

Public Type SYSTEM_THREAD
    KernelTime          As LARGE_INTEGER
    UserTime            As LARGE_INTEGER
    CreateTime          As LARGE_INTEGER
    WaitTime            As Long
    StartAddress        As Long
    ClientId            As CLIENT_ID
    Priority            As Long
    BasePriority        As Long
    ContextSwitchCount  As Long
    State               As KTHREAD_STATE
    WaitReason          As KWAIT_REASON
    dReserved01         As Long
End Type

Public Type SYSTEM_PROCESS_INFORMATION
    NextEntryOffset         As Long
    NumberOfThreads         As Long
    SpareLi1                As LARGE_INTEGER
    SpareLi2                As LARGE_INTEGER
    SpareLi3                As LARGE_INTEGER
    CreateTime              As LARGE_INTEGER
    UserTime                As LARGE_INTEGER
    KernelTime              As LARGE_INTEGER
    ImageName               As UNICODE_STRING
    BasePriority            As Long
    ProcessID               As Long
    InheritedFromProcessId  As Long
    HandleCount             As Long
    SessionID               As Long
    pPageDirectoryBase      As Long '_PTR
    VirtualMemoryCounters   As VM_COUNTERS
    PrivatePageCount        As Long
    IoCounters              As IO_COUNTERS
    Threads()               As SYSTEM_THREAD
End Type

Public Type THREADENTRY32
    dwSize As Long
    dwRefCount As Long
    th32ThreadID As Long
    th32ProcessID As Long
    dwBasePriority As Long
    dwCurrentPriority As Long
    dwFlags As Long
End Type

Public Type PROCESSENTRY32W
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile(MAX_PATH - 1) As Integer
End Type

Public Type MODULEENTRY32W
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule(MAX_MODULE_NAME32) As Integer
    szExePath(MAX_PATH - 1) As Integer
End Type

Public Declare Function NtQuerySystemInformation Lib "NTDLL.DLL" (ByVal infoClass As Long, Buffer As Any, ByVal BufferSize As Long, Ret As Long) As Long
Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Public Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Public Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, lpFilePart As Long) As Long
Public Declare Function QueryFullProcessImageName Lib "kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As Long, ByVal lpdwSize As Long) As Long
Public Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceW" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32.dll" Alias "Process32FirstW" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32W) As Long
Public Declare Function Process32Next Lib "kernel32.dll" Alias "Process32NextW" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32W) As Long
Public Declare Function Module32First Lib "kernel32.dll" Alias "Module32FirstW" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32W) As Long
Public Declare Function Module32Next Lib "kernel32.dll" Alias "Module32NextW" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32W) As Long
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapshot As Long, uThread As THREADENTRY32) As Long
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef ThreadEntry As THREADENTRY32) As Long
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function NtSuspendProcess Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
Public Declare Function NtResumeProcess Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function ZwSetInformationProcess Lib "NTDLL.DLL" (ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function SHRunDialog Lib "shell32.dll" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPTHREAD = &H4
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_QUERY_LIMITED_INFORMATION = &H1000
Public Const PROCESS_VM_READ = 16
Public Const THREAD_SUSPEND_RESUME = &H2
Public Const PROCESS_SUSPEND_RESUME As Long = &H800&

Public Const SystemProcessInformation      As Long = &H5&
Public Const STATUS_INFO_LENGTH_MISMATCH   As Long = &HC0000004
Public Const STATUS_SUCCESS                As Long = 0&
Public Const ERROR_PARTIAL_COPY            As Long = 299&
Public Const ERROR_ACCESS_DENIED           As Long = 5&

'modRegistry

Public Const MAX_PATH_W     As Long = 32767&
Public Const MAX_VALUENAME  As Long = 32767&

Public Enum ENUM_REG_HIVE
    HKEY_USER_SPECIFIED = 0
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    HKCR = &H80000000
    HKCU = &H80000001
    HKLM = &H80000002
    HKU = &H80000003
    HKPD = &H80000004
    HKCC = &H80000005
    HKDD = &H80000006
End Enum
#If False Then
    Dim HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS, HKEY_USER_SPECIFIED
    Dim HKCR, HKCU, HKLM, HKU
#End If

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
#If False Then
    Dim REG_NONE, REG_SZ, REG_EXPAND_SZ, REG_BINARY, REG_DWORD, REG_DWORDLittleEndian, REG_DWORDBigEndian, REG_LINK, REG_MULTI_SZ, REG_ResourceList
    Dim REG_FullResourceDescriptor, REG_ResourceRequirementsList, REG_QWORD, REG_QWORD_LITTLE_ENDIAN
#End If

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
#If False Then
    Dim FLAG_REG_ALL, FLAG_REG_NONE, FLAG_REG_SZ, FLAG_REG_EXPAND_SZ, FLAG_REG_BINARY, FLAG_REG_DWORD, FLAG_REG_DWORDLittleEndian, FLAG_REG_DWORDBigEndian
    Dim FLAG_REG_LINK, FLAG_REG_MULTI_SZ, FLAG_REG_ResourceList, FLAG_REG_FullResourceDescriptor, FLAG_REG_ResourceRequirementsList, FLAG_REG_QWORD, FLAG_REG_QWORD_LITTLE_ENDIAN
#End If

Public Enum KEY_VIRTUAL_TYPE
    KEY_VIRTUAL_NOT_EXIST = 1
    KEY_VIRTUAL_USUAL = 2
    KEY_VIRTUAL_SHARED = 4
    KEY_VIRTUAL_REDIRECTED = 8
End Enum
#If False Then
    Dim KEY_VIRTUAL_NOT_EXIST, KEY_VIRTUAL_USUAL, KEY_VIRTUAL_SHARED, KEY_VIRTUAL_REDIRECTED, KEY_VIRTUAL_SYMLINK
#End If

Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyW" (ByVal hKey As Long, ByVal lpClass As Long, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
Public Declare Function RegGetValue Lib "advapi32.dll" Alias "RegGetValueW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal lpValue As Long, ByVal dwFlags As Long, pdwType As Long, ByVal pvData As Long, pcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueW" (ByVal hKey As Long, ByVal lpValueName As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
Public Declare Function RegDeleteKeyEx Lib "advapi32.dll" Alias "RegDeleteKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal samDesired As Long, ByVal Reserved As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function SHDeleteKey Lib "Shlwapi.dll" Alias "SHDeleteKeyW" (ByVal lRootKey As Long, ByVal szKeyToDelete As Long) As Long
Public Declare Function RegSaveKeyEx Lib "advapi32.dll" Alias "RegSaveKeyExW" (ByVal hKey As Long, ByVal lpFile As Long, ByVal lpSecurityAttributes As Long, ByVal Flags As Long) As Long
Public Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" (lpSystemTime As SYSTEMTIME, vtime As Date) As Long
Public Declare Function GetMem8 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Public Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal PPV As Long, ByVal oVft As Long, ByVal cc As Long, ByVal rtTYP As VbVarType, ByVal paCNT As Long, paTypes As Any, paValues As Any, ByRef fuReturn As Variant) As Long

Public Const CC_STDCALL As Long = 4

Public Const KEY_CREATE_SUB_KEY     As Long = &H4
Public Const KEY_SET_VALUE          As Long = &H2
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const SYNCHRONIZE            As Long = &H100000
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Public Const REG_STANDARD_FORMAT   As Long = 1&
Public Const REG_LATEST_FORMAT     As Long = 2&
Public Const RRF_RT_ANY            As Long = &HFFFF&
Public Const RRF_NOEXPAND          As Long = &H10000000

'modService

Public Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerW" (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceW" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal dwDesiredAccess As Long) As Long
Public Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Public Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Public Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As Any) As Long

Public Const SC_MANAGER_CREATE_SERVICE     As Long = &H2&
Public Const SC_MANAGER_ENUMERATE_SERVICE  As Long = &H4&
Public Const SERVICE_QUERY_CONFIG          As Long = &H1&
Public Const SERVICE_CHANGE_CONFIG         As Long = &H2&
Public Const SERVICE_QUERY_STATUS          As Long = &H4&
Public Const SERVICE_ENUMERATE_DEPENDENTS  As Long = &H8&
Public Const SERVICE_START                 As Long = &H10&
Public Const SERVICE_STOP                  As Long = &H20&
Public Const SERVICE_PAUSE_CONTINUE        As Long = &H40&
Public Const SERVICE_INTERROGATE           As Long = &H80&
Public Const SERVICE_USER_DEFINED_CONTROL  As Long = &H100&
Public Const STANDARD_RIGHTS_REQUIRED      As Long = &HF0000
Public Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Public Const SERVICE_ACCESS_DELETE         As Long = &H10000

Public Const ERROR_SERVICE_SPECIFIC_ERROR  As Long = 1066&
Public Const ERROR_SERVICE_MARKED_FOR_DELETE As Long = 1072&
Public Const ERROR_INVALID_HANDLE          As Long = 6&

'modShortcut

Public Declare Function CoCreateInstance Lib "ole32.dll" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, pvarResult As Object) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, ByVal pszStrPtr As Long) As Long
Public Declare Function CallWindowProcA Lib "user32.dll" (ByVal pFunc As Long, ByVal pESL As Long, ByVal pStrOut As Long, Optional ByVal Reserved1 As Long, Optional ByVal Reserved2 As Long) As Long
'Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long

'modTranslation

Public Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Const LOCALE_SENGLANGUAGE = &H1001&

'modUtils

Public Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Public Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(255) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function PtInRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExW" (ByVal lpFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
'Public Declare Function LoadString Lib "user32.dll" Alias "LoadStringW" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As Long, ByVal nBufferMax As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function VariantTimeToSystemTime Lib "oleaut32.dll" (ByVal vtime As Date, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32.dll" (ByVal lpTimeZone As Any, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function GetTimeZoneInformation Lib "kernel32.dll" (ByVal lpTimeZoneInformation As Long) As Long
Public Declare Function IsWow64Process Lib "kernel32.dll" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Public Const GWL_STYLE As Long = -16&

Public Const TIME_ZONE_ID_INVALID As Long = -1&
Public Const TIME_ZONE_ID_DAYLIGHT As Long = 2
Public Const TIME_ZONE_ID_STANDARD As Long = 1
Public Const TIME_ZONE_ID_UNKNOWN As Long = 0

Public Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2

Public Const GWL_WNDPROC    As Long = &HFFFFFFFC
Public Const WM_MOUSEWHEEL  As Long = &H20A&

Public Const SHCNE_DELETE       As Long = 4&
Public Const SHCNF_PATH         As Long = 1&
Public Const SHCNF_FLUSHNOWAIT  As Long = &H2000&
Public Const SHCNE_CREATE       As Long = 2&
Public Const SHCNE_RENAMEITEM   As Long = 1&
Public Const SHCNE_ATTRIBUTES   As Long = &H800&

Public Const ERROR_FILE_NOT_FOUND      As Long = 2&

Public Const RGN_OR            As Long = 2

'modDigiSign

Public Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Public Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Public Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (arr() As Any) As Long
Public Declare Function GetMem1 Lib "msvbvm60.dll" (pSrc As Any, pDst As Any) As Long
Public Declare Function lstrcpynA Lib "kernel32.dll" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal iMaxLength As Long) As Long

'modWFP

'Public Declare Function VirtualProtect Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Sub EbGetExecutingProj Lib "vba6.dll" (hProject As Long)
Public Declare Function TipGetFunctionId Lib "vba6.dll" (ByVal hProj As Long, ByVal bstrName As Long, ByRef bstrId As Long) As Long
Public Declare Function TipGetLpfnOfFunctionId Lib "vba6.dll" (ByVal hProject As Long, ByVal bstrId As Long, ByRef lpAddress As Long) As Long
'Public Declare Sub SysFreeString Lib "oleaut32.dll" (ByVal lpbstr As Long)
Public Declare Function GetProcAddressByOrd Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As Long) As Long

'Other
Public Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReason As Long) As Long
Public Declare Function GetAllUsersProfileDirectory Lib "Userenv.dll" Alias "GetAllUsersProfileDirectoryW" (ByVal lpProfileDir As Long, lpcchSize As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExW" (ByVal lpDirectoryName As Long, ByVal lpFreeBytesAvailable As Long, ByVal lpTotalNumberOfBytes As Long, ByVal lpTotalNumberOfFreeBytes As Long) As Long
Public Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryW" (ByVal lpPathName As Long) As Long
Public Declare Function AssocQueryString Lib "Shlwapi.dll" Alias "AssocQueryStringW" (ByVal Flags As Long, ByVal str As Long, ByVal pszAssoc As Long, ByVal pszExtra As Long, ByVal pszOut As Long, pcchOut As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringW" (ByVal lpAppName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long
Public Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Public Declare Function GetTickCount64 Lib "kernel32" () As Currency
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function WaitForInputIdle Lib "user32.dll" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemoryStr Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function AttachThreadInput Lib "user32.dll" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function SetFocus2 Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Const EWX_RESTARTAPPS As Long = &H40&
Public Const EWX_REBOOT As Long = 2&
Public Const EWX_FORCEIFHUNG As Long = &H10&
Public Const EWX_FORCE As Long = 4&
Public Const SHTDN_REASON_MAJOR_APPLICATION As Long = &H40000
Public Const SHTDN_REASON_MINOR_INSTALLATION As Long = 2&
Public Const SHTDN_REASON_FLAG_PLANNED As Long = &H80000000

Public Const LB_SETTABSTOPS As Long = &H192&

Public Type SYSTEM_MODULE
    Reserved1           As Long
    Reserved2           As Long
    ImageBaseAddress    As Long
    ImageSize           As Long
    Flags               As Long
    ID                  As Integer
    Rank                As Integer
    w018                As Integer
    NameOffset          As Integer
    Name                As String * 256
End Type
 
Public Type SYSTEM_MODULE_INFORMATION
    ModulesCount        As Long
    Modules()           As SYSTEM_MODULE
End Type

Public Const WM_KEYDOWN = &H100
Public Const WM_CHAR = &H102
Public Const LVM_GETITEMCOUNT = 4100
Public Const LVM_GETITEMTEXTW = 4211
Public Const LVM_GETITEMSTATE = 4140
Public Const LVIS_SELECTED = 2
Public Const LVM_SETITEMSTATE = 4139
Public Const LVIF_TEXT = 1
Public Const LVIF_STATE = 8
Public Const PROCESS_VM_OPERATION = &H8
Public Const PROCESS_VM_WRITE = &H20
Public Const MEM_COMMIT = &H1000
Public Const MEM_DECOMMIT = &H4000
Public Const MEM_RELEASE = &H8000
Public Const PAGE_READWRITE = &H4
Public Const LVIS_FOCUSED = 1
