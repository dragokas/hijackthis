Attribute VB_Name = "modFile"
'[modFile.bas]

'
' modFile module by Alex Dragokas
'

Option Explicit

Const MAX_PATH As Long = 260&
Const MAX_FILE_SIZE As Currency = 104857600@

'Public Enum VbFileAttributeExtended
'    vbAll = -1&
'    vbDirectory = 16& ' mean - include folders also
'    vbFile = vbAll And Not vbDirectory
'    vbSystem = 4&
'    vbHidden = 2&
'    vbReadOnly = 1
'    vbNormal = 0&
'    vbReparse = 1024& 'symlinks / junctions (not include hardlink to file; they reflect attributes of the target)
'End Enum
'#If False Then
'    Dim vbAll, vbFile, vbReparse 'case sensitive protection against modification (for non-overloaded enum variables only)
'#End If
'
'Public Enum VB_FILE_ACCESS_MODE
'    FOR_READ = 1
'    FOR_READ_WRITE = 2
'    FOR_OVERWRITE_CREATE = 4
'End Enum
'#If False Then
'    Dim FOR_READ, FOR_READ_WRITE, FOR_OVERWRITE_CREATE
'#End If
'
'Public Enum ENUM_FILE_DATE_TYPE
'    DATE_CREATED = 1
'    DATE_MODIFIED = 2
'    DATE_ACCESSED = 3
'End Enum
'#If False Then
'    Dim DATE_CREATED, DATE_MODIFIED, DATE_ACCESSED
'#End If
 
'Private Type LARGE_INTEGER
'    LowPart As Long
'    HighPart As Long
'End Type
'
'Private Type FILETIME
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type
'
'Private Type WIN32_FIND_DATA
'    dwFileAttributes As Long
'    ftCreationTime As FILETIME
'    ftLastAccessTime As FILETIME
'    ftLastWriteTime As FILETIME
'    nFileSizeHigh As Long
'    nFileSizeLow As Long
'    dwReserved0 As Long
'    dwReserved1 As Long
'    lpszFileName(MAX_PATH - 1) As Integer
'    lpszAlternate(13) As Integer
'End Type
'
'Private Type SECURITY_ATTRIBUTES
'    nLength As Long
'    lpSecurityDescriptor As Long
'    bInheritHandle As Long
'End Type
'
'Private Type VS_FIXEDFILEINFO
'    dwSignature As Long
'    dwStrucVersionl As Integer
'    dwStrucVersionh As Integer
'    dwFileVersionMSl As Integer
'    dwFileVersionMSh As Integer
'    dwFileVersionLSl As Integer
'    dwFileVersionLSh As Integer
'    dwProductVersionMSl As Integer
'    dwProductVersionMSh As Integer
'    dwProductVersionLSl As Integer
'    dwProductVersionLSh As Integer
'    dwFileFlagsMask As Long
'    dwFileFlags As Long
'    dwFileOS As Long
'    dwFileType As Long
'    dwFileSubtype As Long
'    dwFileDateMS As Long
'    dwFileDateLS As Long
'End Type
'
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
'Public Enum DRIVE_TYPE
'    DRIVE_UNKNOWN = 0
'    DRIVE_NO_ROOT_DIR
'    DRIVE_REMOVABLE
'    DRIVE_FIXED
'    DRIVE_REMOTE
'    DRIVE_CDROM
'    DRIVE_RAMDISK
'    DRIVE_ANY
'End Enum
'
'Public Enum DRIVE_TYPE_BIT
'    DRIVE_BIT_UNKNOWN = 1
'    DRIVE_BIT_NO_ROOT_DIR = 2
'    DRIVE_BIT_REMOVABLE = 4
'    DRIVE_BIT_FIXED = 8
'    DRIVE_BIT_REMOTE = 16
'    DRIVE_BIT_CDROM = 32
'    DRIVE_BIT_RAMDISK = 64
'    DRIVE_BIT_ANY = 128
'End Enum
'
'Private Type SHFILEOPSTRUCT
'    hWnd                    As Long
'    wFunc                   As Long
'    pFrom                   As Long
'    pTo                     As Long
'    fFlags                  As Integer
'    fAnyOperationsAborted   As Long
'    hNameMappings           As Long
'    lpszProgressTitle       As Long '  only used if FOF_SIMPLEPROGRESS
'End Type
'
'Private Type SHELLEXECUTEINFO
'    cbSize          As Long
'    fMask           As Long
'    hWnd            As Long
'    lpVerb          As Long
'    lpFile          As Long
'    lpParameters    As Long
'    lpDirectory     As Long
'    nShow           As Long
'    hInstApp        As Long
'    lpIDList        As Long
'    lpClass         As Long
'    hkeyClass       As Long
'    dwHotKey        As Long
'    hIcon           As Long
'    hProcess        As Long
'End Type
'
'Private Type UUID
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data4(0 To 7) As Byte
'End Type

'Public Type MOUNTMGR_TARGET_NAME
'    DeviceNameLength As Integer
'    DeviceName(MAX_PATH) As Integer 'WCHAR DeviceName[1] 'MAX_PATH + NUL
'End Type
'
'Public Type MOUNTMGR_VOLUME_PATHS
'    MultiSzLength As Long
'    MultiSz(MAX_PATH) As Integer 'WCHAR MultiSz[1] 'MAX_PATH + NUL
'End Type
'
'Public Type FILE_NAME_INFORMATION
'    FileNameLength As Long
'    FileName(MAX_PATH) As Integer 'WCHAR FileName[1] 'MAX_PATH + NUL
'End Type
'
'Public Type MOUNTMGR_BUFER
'    TargetName As MOUNTMGR_TARGET_NAME
'    TargetPaths As MOUNTMGR_VOLUME_PATHS
'    NameInfo As FILE_NAME_INFORMATION
'    UnicodeString As UNICODE_STRING
'    Buffer(MAX_PATH) As Integer
'End Type
'
'Public Enum OBJECT_INFORMATION_CLASS
'    ObjectBasicInformation = 0
'    ObjectNameInformation
'    ObjectTypeInformation
'    ObjectAllTypesInformation
'    ObjectHandleInformation
'    ObjectSessionInformation
'End Enum
'
'Public Enum VOLUME_INFO_FLAGS
'    FILE_CASE_PRESERVED_NAMES = 2
'    FILE_CASE_SENSITIVE_SEARCH = 1
'    FILE_DAX_VOLUME = &H20000000 ' introduced in Windows 10, version 1607.
'    FILE_FILE_COMPRESSION = &H10&
'    FILE_NAMED_STREAMS = &H40000
'    FILE_PERSISTENT_ACLS = 8
'    FILE_READ_ONLY_VOLUME = &H80000
'    FILE_SEQUENTIAL_WRITE_ONCE = &H100000
'    FILE_SUPPORTS_ENCRYPTION = &H20000
'    FILE_SUPPORTS_EXTENDED_ATTRIBUTES = &H800000    'value is not supported until Windows Server 2008 R2 and Windows 7.
'    FILE_SUPPORTS_HARD_LINKS = &H400000             'value is not supported until Windows Server 2008 R2 and Windows 7.
'    FILE_SUPPORTS_OBJECT_IDS = &H10000
'    FILE_SUPPORTS_OPEN_BY_FILE_ID = &H1000000       'value is not supported until Windows Server 2008 R2 and Windows 7.
'    FILE_SUPPORTS_REPARSE_POINTS = &H80&            'Note: ReFS can't enum them with FindFirstVolumeMountPoint / FindNextVolumeMountPoint
'    FILE_SUPPORTS_SPARSE_FILES = &H40&
'    FILE_SUPPORTS_TRANSACTIONS = &H200000
'    FILE_SUPPORTS_USN_JOURNAL = &H2000000           'value is not supported until Windows Server 2008 R2 and Windows 7.
'    FILE_UNICODE_ON_DISK = 4
'    FILE_VOLUME_IS_COMPRESSED = &H8000&
'    FILE_VOLUME_QUOTAS = &H20&
'End Enum

'
'Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingW" (ByVal hFile As Long, ByVal lpAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
'Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
'Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Long
'Private Declare Function PathFileExists Lib "Shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
'Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
'Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
'Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
'Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
''Private Declare Function SHFileExists Lib "shell32.dll" Alias "#45" (ByVal szPath As String) As Long
'Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
'Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Long
'Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeW" (ByVal nDrive As Long) As Long
'Private Declare Function GetLogicalDrives Lib "kernel32.dll" () As Long
'Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
'Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long) As Long
''Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
'Private Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
'Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
'Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStringDest As Long, ByVal lpStringSrc As Long) As Long
'Private Declare Function GetLongPathNameW Lib "kernel32.dll" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
'Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
'Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As Long, lpdwHandle As Long) As Long
'Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Long, puLen As Long) As Long
'Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
'Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal lSize As Long)
'Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
'Private Declare Function PathIsNetworkPath Lib "Shlwapi.dll" Alias "PathIsNetworkPathW" (ByVal pszPath As Long) As Long
'Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, ByVal lpOutBuffer As Long, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bDontOverwrite As Long) As Long
'Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationW" (lpFileOp As SHFILEOPSTRUCT) As Long
'Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
'Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameW" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
'Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExW" (SEI As SHELLEXECUTEINFO) As Long
'Private Declare Function MoveFileEx Lib "kernel32.dll" Alias "MoveFileExW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
'Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathW" (ByVal hWndOwner As Long, ByVal CSIDL As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As Long) As Long
'Private Declare Function SHGetKnownFolderPath Lib "shell32.dll" (rfid As UUID, ByVal dwFlags As Long, ByVal hToken As Long, ppszPath As Long) As Long
'Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal iMaxLength As Long) As Long
'Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszGuid As Long, pGuid As UUID) As Long
'Private Declare Function PathFindOnPath Lib "Shlwapi.dll" Alias "PathFindOnPathW" (ByVal pszFile As Long, ppszOtherDirs As Any) As Long
Private Declare Function NtQueryObject Lib "ntdll.dll" (ByVal Handle As Long, ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, ObjectInformation As Any, ByVal ObjectInformationLength As Long, ReturnLength As Long) As Long
'
'Const FILE_SHARE_READ           As Long = &H1&
'Const FILE_SHARE_WRITE          As Long = &H2&
'Const FILE_SHARE_DELETE         As Long = 4&
'Const FILE_READ_ATTRIBUTES      As Long = &H80&
'Const OPEN_EXISTING             As Long = 3&
'Const CREATE_ALWAYS             As Long = 2&
'Const GENERIC_READ              As Long = &H80000000
'Const GENERIC_WRITE             As Long = &H40000000
'Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&
'Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF
'Const ERROR_SUCCESS             As Long = 0&
'Const INVALID_FILE_ATTRIBUTES   As Long = -1&
'Const NO_ERROR                  As Long = 0&
'Const FILE_BEGIN                As Long = 0&
'Const FILE_CURRENT              As Long = 1&
'Const FILE_END                  As Long = 2&
'Const INVALID_SET_FILE_POINTER  As Long = &HFFFFFFFF
'Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
'Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400&
'Const ERROR_HANDLE_EOF          As Long = 38&
'Const SEC_IMAGE                 As Long = &H1000000
'Const PAGE_READONLY             As Long = 2&
'Const FILE_MAP_READ             As Long = 4&
'
'Const HKEY_LOCAL_MACHINE        As Long = &H80000002
'Const KEY_QUERY_VALUE           As Long = &H1&
'Const RegType_DWord             As Long = 4&
'
'Const MOVEFILE_DELAY_UNTIL_REBOOT As Long = &H4&
'
'Private Const IOCTL_STORAGE_CHECK_VERIFY2   As Long = &H2D0800
'Private Const IOCTL_STORAGE_CHECK_VERIFY    As Long = &H2D4800

Const FileNameInformation       As Long = 9&

Const ch_Dot                    As String = "."
Const ch_DotDot                 As String = ".."
Const ch_Slash                  As String = "\"
Const ch_SlashAsterisk          As String = "\*"

Private lWow64Old               As Long
Private DriveTypeName           As New Collection
Private arrPathFolders()        As String
Private arrPathFiles()          As String
Private Total_Files             As Long
Private Total_Folders           As Long


Public Function FileExists(ByVal sFile As String, Optional bUseWow64 As Boolean, Optional bAllowNetwork As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    
    Static bLastFile(2) As String, bLastStatus(2) As Boolean
    Dim bIsWinSysDir As Boolean
    Dim pos As Long
    
    AppendErrorLogCustom "FileExists - Begin", "File: " & sFile
    
    '\\?\ \\.\
    If Left$(sFile, 4) = "\\?\" Then sFile = Mid$(sFile, 5)
    If Left$(sFile, 4) = "\\.\" Then sFile = Mid$(sFile, 5)
    
    'ADS?
    pos = InStr(4, sFile, ":")
    If pos Then sFile = Left$(sFile, pos - 1)
    '// TODO add checking if stream exists
    
    sFile = EnvironW(Trim$(sFile))
    If Len(sFile) = 0 Then GoTo ExitFunc
    If Left$(sFile, 2) = "\\" Then
        If Not bAllowNetwork Then
            GoTo ExitFunc 'DriveType = "REMOTE"
        End If
    End If
    
    'little cache stack :)
    If bScanMode Then ' used only in HJT Checking mode. This flag has set in "StartScan" function
        If StrComp(sFile, bLastFile(2), 1) = 0 Then FileExists = bLastStatus(2): GoTo ExitFunc
        If StrComp(sFile, bLastFile(1), 1) = 0 Then FileExists = bLastStatus(1): GoTo ExitFunc
        If StrComp(sFile, bLastFile(0), 1) = 0 Then FileExists = bLastStatus(0): GoTo ExitFunc
        
        'advanced cache - to minimize future numbers of file system redirector calls

        If StrComp(Left$(sFile, Len(sWinSysDir)), sWinSysDir, vbTextCompare) = 0 Then
            bIsWinSysDir = True
            If oDictFileExist.Exists(sFile) Then
                FileExists = oDictFileExist(sFile)
                GoTo Finalize
            End If
        End If
    End If
    
    ' use 2 methods for reliability reason (both supported unicode pathes)
    Dim Ex(1) As Boolean
    Dim ret As Long
    Dim Redirect As Boolean, bOldStatus As Boolean
    Dim WFD     As WIN32_FIND_DATA
    Dim hFile   As Long
    
    If Not bUseWow64 Then Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    ret = GetFileAttributes(StrPtr(sFile))
    If Err.LastDllError = 5 Or (ret <> INVALID_FILE_ATTRIBUTES And (0 = (ret And FILE_ATTRIBUTE_DIRECTORY))) Then
        Ex(0) = True
    End If
    
    If Not bAutoLogSilent Then
        hFile = FindFirstFile(StrPtr(sFile), WFD)
        
        If hFile <> INVALID_HANDLE_VALUE Then
            If Not CBool(WFD.dwFileAttributes And vbDirectory) And WFD.dwFileAttributes <> INVALID_FILE_ATTRIBUTES Then Ex(1) = True
            FindClose hFile
        End If
    End If
    
    '// FS redirection reverting if need
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    FileExists = Ex(0) Or Ex(1)

    If bIsWinSysDir Then
        oDictFileExist.Add sFile, FileExists
    End If

Finalize:

    'shift cache stack
    bLastFile(0) = bLastFile(1)
    bLastFile(1) = bLastFile(2)
    bLastFile(2) = sFile
    
    bLastStatus(0) = bLastStatus(1)
    bLastStatus(1) = bLastStatus(2)
    bLastStatus(2) = FileExists
    
ExitFunc:
    AppendErrorLogCustom "FileExists - End", "File: " & sFile, "bUseWow64: " & bUseWow64, "Exists: " & FileExists
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FileExists", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function

Public Function FolderExists(ByVal sFolder$, Optional bUseWow64 As Boolean, Optional bAllowNetwork As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "FolderExists - Begin", "Folder: " & sFolder, "bUseWow64: " & bUseWow64
    
    Dim ret As Long, Redirect As Boolean, bOldStatus As Boolean
        
    sFolder = Trim$(sFolder)
    If Len(sFolder) = 0 Then Exit Function
    
    If Left$(sFolder, 2) = "\\" Then
        If Not bAllowNetwork Then
            Exit Function 'network path
        End If
    End If
    
    If Not bUseWow64 Then Redirect = ToggleWow64FSRedirection(False, sFolder, bOldStatus)
    
    ret = GetFileAttributes(StrPtr(sFolder))
    If Err.LastDllError = 5 Or (CBool(ret And vbDirectory) And (ret <> INVALID_FILE_ATTRIBUTES)) Then
        FolderExists = True
    End If
    
    '// FS redirection reverting if need
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "FolderExists - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FolderExists", "Folder:", sFolder$, "Redirection: ", bUseWow64
    If inIDE Then Stop: Resume Next
End Function


'Public Sub GetDriveTypeNames()
'    On Error GoTo ErrorHandler
'    Dim lr As Long
'    Dim i  As Long
'    Dim DT As String
'
'    For i = 65& To 90&
'
'      lr = GetDriveType(StrPtr(Chr$(i) & ":\"))
'
'      Select Case lr
'        Case 3&
'            DT = "FIXED"
'        Case 2&
'            DT = "REMOVABLE"
'        Case 5&
'            DT = "CDROM"
'        Case 4&
'            DT = "REMOTE"
'        Case 0&
'            DT = "UNKNOWN"
'        Case 1&
'            DT = "DISCONNECTED" '"NO_ROOT_DIR"
'        Case 6&
'            DT = "RAMDISK"
'        Case Else
'            DT = "UNKNOWN"
'      End Select
'
'      DriveTypeName.Add DT, Chr$(i)
'
'    Next
'
'    Exit Sub
'ErrorHandler:
'    ErrorMsg err, "modFile.GetDriveTypeNames", "Drive:", Chr$(i)
'End Sub


Function FileLenW(Optional Path As String, Optional hFileHandle As Long) As Currency ', Optional DoNotUseCache As Boolean
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FileLenW - Begin", "Path: " & Path, "Handle: " & hFileHandle
    
'    ' Last cached File
'    Static CachedFile As String
'    Static CachedSize As Currency
    
    Dim lr          As Long
    Dim hFile       As Long
    Dim FileSize    As Currency
    Dim Redirect    As Boolean, bOldStatus As Boolean

'    If Not DoNotUseCache Then
'        If StrComp(Path, CachedFile, 1) = 0 Then
'            FileLenW = CachedSize
'            Exit Function
'        End If
'    End If
    
    If hFileHandle = 0 Then
        Redirect = ToggleWow64FSRedirection(False, Path, bOldStatus)
        hFile = CreateFile(StrPtr(Path), 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        hFile = hFileHandle
    End If
    
    If hFile > 0 Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then FileLenW = FileSize * 10000&
        End If
'        If Not DoNotUseCache Then
'            CachedFile = Path
'            CachedSize = FileLenW
'        End If
        If hFileHandle = 0 Then CloseHandle hFile: hFile = 0&
    End If
    
    If hFileHandle = 0 Then
        If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    End If
    
    AppendErrorLogCustom "FileLenW - End", "Size: " & FileSize
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FileLenW", "File:", Path, "hFile:", hFile, "FileSize:", FileLenW, "Return:", lr
End Function



Public Function OpenW(FileName As String, Access As VB_FILE_ACCESS_MODE, retHandle As Long, Optional Flags As Long) As Boolean
    
    AppendErrorLogCustom "OpenW - Begin", "File: " & FileName, "Access: " & Access
    
    Dim FSize As Currency
    
    If Access And FOR_READ Then
        If Not FileExists(FileName) Then
            retHandle = INVALID_HANDLE_VALUE
            Exit Function
        End If
    End If
    
    'For read operation we are applying (GENERIC_READ - EA - SYNCHRONIZE) access rights,
    'because if DACL block GENERIC_WRITE, it also block SYNCHRONIZE, such a way it also block GENERIC_READ (info from MSDN)
    '-EA, just because I don't need it.
    
    If Access = FOR_READ Then
        retHandle = CreateFile(StrPtr(FileName), FILE_READ_ATTRIBUTES Or FILE_READ_DATA Or STANDARD_RIGHTS_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, Flags Or FILE_FLAG_SEQUENTIAL_SCAN, ByVal 0&)
    ElseIf Access = FOR_OVERWRITE_CREATE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_WRITE, 0&, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE Or Flags, ByVal 0&)
    ElseIf Access = FOR_READ_WRITE Then
        If Not FileExists(FileName) Then
            retHandle = CreateFile(StrPtr(FileName), GENERIC_WRITE, 0&, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE Or Flags, ByVal 0&)
            If retHandle <> INVALID_HANDLE_VALUE Then
                CloseHandle retHandle
            End If
        End If
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, Flags, ByVal 0&)
    Else
        'WriteCon "Wrong access mode!", cErr
    End If

    OpenW = (INVALID_HANDLE_VALUE <> retHandle)
    
    ' ограничение на максимально возможный файл дл€ открыти€ ( > 100 ћЅ )
    If OpenW Then
        If Access And (FOR_READ Or FOR_READ_WRITE) Then
            FSize = LOFW(retHandle)
            If FSize > MAX_FILE_SIZE Then
                CloseHandle retHandle
                retHandle = INVALID_HANDLE_VALUE
                OpenW = False
                '"Ќе хочу и не буду открывать этот файл, потому что его размер превышает безопасный максимум"
                Err.Clear: AppendErrorLogNoErr Err, "modFile.OpenW", "Trying to open too big file" & ": (" & (FSize \ 1024 \ 1024) & " MB.) " & FileName
            End If
        End If
    Else
        AppendErrorLogNoErr Err, "modFile.OpenW", "Cannot open file: " & FileName
        'Err.Raise 75 ' Path/File Access error
    End If

    AppendErrorLogCustom "OpenW - End", "Handle: " & retHandle
End Function

                                                                  'do not change Variant type at all or you will die ^_^
Public Function GetW(hFile As Long, Optional vPos As Variant, Optional vOut As Variant, Optional vOutPtr As Long, Optional cbToRead As Long) As Boolean
    
    'On Error GoTo ErrorHandler
    AppendErrorLogCustom "GetW - Begin", "Handle: " & hFile, "pos: " & vPos, "cbToRead: " & cbToRead

    Dim lBytesRead  As Long
    Dim lr          As Long
    Dim ptr         As Long
    Dim vType       As Long
    Dim UnknType    As Boolean
    Dim oldPos      As Long
    Dim pos         As Long
    
    'modMain_CheckO1Item - #52 (Bad file name or number) LastDllError = 126 (Ќе найден указанный модуль.)
    If Not IsMissing(vPos) Then
      pos = CLng(vPos)
      If pos >= 1 Then
        pos = pos - 1   ' VB's Get & SetFilePointer difference correction
        
        oldPos = SetFilePointer(hFile, 0&, ByVal 0&, FILE_CURRENT)
        If oldPos = INVALID_SET_FILE_POINTER And NO_ERROR <> Err.LastDllError Then
            Err.Clear: ErrorMsg Err, "Cannot get file pointer! LastDllErr = " & Err.LastDllError: Err.Raise 52
        End If
        
        If pos <> oldPos Then
            If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then
                If NO_ERROR <> Err.LastDllError Then
                    Err.Clear: ErrorMsg Err, "Cannot set file pointer! LastDllErr = " & Err.LastDllError: Err.Raise 52
                End If
            End If
        End If
      End If
    End If
    
    If Not IsMissing(vOut) Then
        vType = VarType(vOut)
    End If
    
    If 0 <> cbToRead Then   'vbError = vType
        lr = ReadFile(hFile, vOutPtr, cbToRead, lBytesRead, 0&)
                
    ElseIf vbString = vType Then
        lr = ReadFile(hFile, StrPtr(vOut), Len(vOut), lBytesRead, 0&)
        If Err.LastDllError <> 0 Or lr = 0 Then Err.Raise 52
        
        vOut = StrConv(vOut, vbUnicode)
        If Len(vOut) <> 0 Then vOut = Left$(vOut, Len(vOut) \ 2)
    Else
        'do a bit of magik :)
        memcpy ptr, ByVal VarPtr(vOut) + 8, 4&  'VT_BYREF
        Select Case vType
        Case vbByte
            lr = ReadFile(hFile, ptr, 1&, lBytesRead, 0&)
        Case vbInteger
            lr = ReadFile(hFile, ptr, 2&, lBytesRead, 0&)
        Case vbLong
            lr = ReadFile(hFile, ptr, 4&, lBytesRead, 0&)
        Case vbCurrency
            lr = ReadFile(hFile, ptr, 8&, lBytesRead, 0&)
        Case Else
            UnknType = True
            ErrorMsg Err, "modFile.GetW. type #" & VarType(vOut) & " of buffer is not supported.": Err.Raise 52
        End Select
    End If
    GetW = (0 <> lr)
    If 0 = lr And Not UnknType Then ErrorMsg Err, "Cannot read file. LastDllErr = " & Err.LastDllError: Err.Raise 52
    
    AppendErrorLogCustom "GetW - End", "BytesRead: " & lBytesRead
'    Exit Function
'ErrorHandler:
'    AppendErrorLogFormat Now, err, "modFile.GetW"
'    Resume Next
End Function

Public Function PutStringW(hFile As Long, Optional pos As Long, Optional sStr As String) As Boolean
    Dim doAppend As Boolean
    If Len(sStr) <> 0 Then
        If pos = 0 Then doAppend = True
        PutStringW = PutW(hFile, pos, StrPtr(sStr), LenB(sStr), doAppend)
    Else
        PutStringW = True
    End If
End Function

Public Function PutW(hFile As Long, pos As Long, vInPtr As Long, cbToWrite As Long, Optional doAppend As Boolean) As Boolean
    On Error GoTo ErrorHandler
    'don't uncomment it -> recurse on bDebugToFile !!!
    'AppendErrorLogCustom "PutW - Begin", "Handle: " & hFile, "pos: " & pos, "Bytes: " & cbToWrite
    
    Dim lBytesWrote  As Long
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If doAppend Then
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) Then Exit Function
    Else
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then Exit Function
    End If
    
    If NO_ERROR = Err.LastDllError Then
    
        If WriteFile(hFile, vInPtr, cbToWrite, lBytesWrote, 0&) Then PutW = True
        
    End If
    
    'AppendErrorLogCustom "PutW - End"
    Exit Function
ErrorHandler:
    'don't change/append this identifier !!! -> can cause recurse on bDebugToFile !!!
    ErrorMsg Err, "modFile.PutW"
End Function

Public Function LOFW(hFile As Long) As Currency
    On Error GoTo ErrorHandler
    Dim lr          As Long
    Dim FileSize    As Currency
    
    AppendErrorLogCustom "LOFW - Begin", "Handle: " & hFile
    
    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then
                LOFW = FileSize * 10000&
            Else
                Err.Clear
                ErrorMsg Now, "File is too big. Size: " & FileSize
            End If
        End If
    End If
    
    AppendErrorLogCustom "LOFW - End", "Size: " & LOFW
ErrorHandler:
End Function

Public Function PrintW(hFile As Long, sStr As String, Optional bUnicode As Boolean) As Boolean
    Dim bSuccess As Boolean
    If hFile > 0 Then
        If Len(sStr) <> 0 Then
            If bUnicode Then
                bSuccess = PutW(hFile, 0, StrPtr(sStr), LenB(sStr), True)
            Else
                bSuccess = PutW(hFile, 0, StrPtr(StrConv(sStr, vbFromUnicode)), Len(sStr), True)
            End If
        Else
            bSuccess = True
        End If
        ' + CrLf
        If bUnicode Then
            bSuccess = bSuccess And PutW(hFile, 0, StrPtr(vbCrLf), 4, True)
        Else
            bSuccess = bSuccess And PutW(hFile, 0, StrPtr(ChrW(&HA0D&)), 2, True)
        End If
    End If
End Function

Public Function PrintBOM(hFile As Long) As Boolean
    Dim BOM As String
    BOM = ChrW$(-257)
    PrintBOM = PutW(hFile, 0, StrPtr(BOM), 2, True)
End Function

Public Function CloseW(hFile As Long, Optional bFlushBuffers As Boolean) As Long
    AppendErrorLogCustom "CloseW", "Handle: " & hFile
    If hFile = 0 Then Exit Function
    
'    Dim AccessMode As Long
'    Dim IO As IO_STATUS_BLOCK
'    Dim accInf As FILE_ACCESS_INFORMATION
'
'    If STATUS_SUCCESS = NtQueryInformationFile(hFile, IO, accInf, Len(accInf), FileAccessInformation) Then
'        If AccessMode And FILE_WRITE_DATA Then
'            FlushFileBuffers hFile
'        End If
'    End If
    
    If bFlushBuffers Then
        FlushFileBuffers hFile
    End If

    CloseW = CloseHandle(hFile)
    If CloseW Then hFile = 0
End Function

'// TODO. I don't like it. Re-check it !!!
Public Function LineInputW(hFile As Long, sLine As String) As Boolean
    Dim ch$, lBytesRead&, lr&
    sLine = vbNullString
    Do
        ch = vbNullChar

        lr = ReadFile(hFile, StrPtr(ch), 1, lBytesRead, 0&)
        
        If lr = 0 Or lBytesRead = 0 Or AscW(ch) = 10 Then
            If Right$(sLine, 1) = vbCr Then sLine = Left$(sLine, Len(sLine) - 1)
            Exit Do
        Else
            LineInputW = True
            sLine = sLine & ch
        End If
    Loop
End Function

Public Function ToggleWow64FSRedirection(bEnable As Boolean, Optional PathNecessity As String, Optional OldStatus As Boolean) As Boolean
    'Static lWow64Old        As Long    'Warning: do not use initialized variables for this API !
                                        'Static variables is not allowed !
                                        'lWow64Old is now declared globally
    'True - enable redirector
    'False - disable redirector

    'OldStatus: current state of redirection
    'True - redirector was enabled
    'False - redirector was disabled

    'Return value is:
    'true if success

    Static IsNotRedirected  As Boolean
    Dim lr                  As Long
    
    OldStatus = Not IsNotRedirected
    
    If Not bIsWin64 Then Exit Function
    
    If Len(PathNecessity) <> 0 Then
        If StrComp(Left$(Replace(Replace(PathNecessity, "/", "\"), "\\", "\"), Len(sWinSysDir)), sWinSysDir, vbTextCompare) <> 0 Then Exit Function
    End If
    
    If bEnable Then
        If IsNotRedirected Then
            lr = Wow64RevertWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = False
        End If
    Else
        If Not IsNotRedirected Then
            lr = Wow64DisableWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = True
        End If
    End If
    
    If ToggleWow64FSRedirection Then
        If OldStatus <> bEnable Then
            AppendErrorLogCustom "ToggleWow64FSRedirection - End", "Path: " & PathNecessity, _
                "Old State: " & OldStatus, "New State: " & bEnable
        End If
    End If
    
End Function


Public Function GetExtensionName(Path As String) As String  'вернет .ext
    Dim pos As Long
    pos = InStrRev(Path, ".")
    If pos <> 0 Then GetExtensionName = Mid$(Path, pos)
End Function

'' явл€етс€ ли файл форматом PE EXE
'Public Function isPE_EXE(Optional FileName As String, Optional FileHandle As Long) As Boolean
'    On Error GoTo ErrorHandler
'    AppendErrorLogCustom "isPE_EXE - Begin", "File: " & FileName
'
''    #If UseHashTable Then
''        Static PE_EXE_Cache As clsTrickHashTable
''    #Else
''        Static PE_EXE_Cache As Object
''    #End If
''
''    If 0 = ObjPtr(PE_EXE_Cache) Then
''        #If UseHashTable Then
''            Set PE_EXE_Cache = New clsTrickHashTable
''        #Else
''            Set PE_EXE_Cache = CreateObject("Scripting.Dictionary")
''        #End If
''        PE_EXE_Cache.CompareMode = vbTextCompare
''    Else
''        If Len(FileName) <> 0& Then
''            If PE_EXE_Cache.Exists(FileName) Then
''                isPE_EXE = PE_EXE_Cache(FileName)
''                Exit Function
''            End If
''        End If
''    End If
'
'    'Static PE_EXE_Cache    As New Collection ' value = true, если файл €вл€етс€ форматом PE EXE
'
'    'If Len(FileName) <> 0 Then
'    '    If isCollectionKeyExists(FileName, PE_EXE_Cache) Then
'    '        isPE_EXE = PE_EXE_Cache(FileName)
'    '        Exit Function
'    '    End If
'    'End If
'
'    Dim hFile          As Long
'    Dim PE_offset      As Long
'    Dim MZ(1)          As Byte
'    Dim pe(3)          As Byte
'    Dim FSize          As Currency
'
'    If FileHandle = 0& Then
'        OpenW FileName, FOR_READ, hFile
'    Else
'        hFile = FileHandle
'    End If
'    If hFile <> INVALID_HANDLE_VALUE Then
'        FSize = LOFW(hFile)
'        If FSize >= &H3C& + 4& Then
'            GetW hFile, 1&, , VarPtr(MZ(0)), ((UBound(MZ) + 1&) * CLng(LenB(MZ(0))))
'            If (MZ(0) = 77& And MZ(1) = 90&) Or (MZ(1) = 77& And MZ(0) = 90&) Then  'MZ or ZM
'                GetW hFile, &H3C& + 1&, PE_offset
'                If PE_offset And FSize >= PE_offset + 4 Then
'                    GetW hFile, PE_offset + 1&, , VarPtr(pe(0)), ((UBound(pe) + 1&) * CLng(LenB(pe(0))))
'                    If pe(0) = 80& And pe(1) = 69& And pe(2) = 0& And pe(3) = 0& Then isPE_EXE = True   'PE NUL NUL
'                End If
'            End If
'        End If
'        If FileHandle = 0& Then CloseW hFile: hFile = 0&
'    End If
'
'    'If Len(FileName) <> 0& Then PE_EXE_Cache.Add FileName, isPE_EXE
'
'    AppendErrorLogCustom "isPE_EXE - End"
'    Exit Function
'
'ErrorHandler:
'    ErrorMsg Err, "Parser.isPE_EXE", "File:", FileName
'    'On Error Resume Next
'    'If Len(FileName) <> 0& Then PE_EXE_Cache.Add FileName, isPE_EXE
'    If FileHandle = 0& Then
'        If hFile <> 0 Then CloseW hFile: hFile = 0&
'    End If
'End Function

'main function to list folders

' ¬озвращает массив путей.
' ≈сли ничего не найдено - возвращаетс€ неинициализированный массив.
Public Function ListSubfolders(Path As String, Optional Recursively As Boolean = False) As String()
    On Error GoTo ErrorHandler

    AppendErrorLogCustom "ListSubfolders - Begin", "Path:", Path, "Recur:", Recursively

    Dim bRedirStateChanged As Boolean, bOldState As Boolean
    
    'прежде, чем использовать ListSubfolders_Ex, нужно инициализировать глобальные массивы.
    ReDim arrPathFolders(100) As String
    'при каждом вызове ListSubfolders_Ex следует обнулить глобальный счетчик файлов
    Total_Folders = 0&
    
    If bIsWin64 Then
        If StrBeginWith(Path, sWinDir) Then
            bRedirStateChanged = ToggleWow64FSRedirection(False, , bOldState)
        End If
    End If
    
    'вызов тушки
    Call ListSubfolders_Ex(Path, Recursively)
    If Total_Folders > 0 Then
        Total_Folders = Total_Folders - 1
        ReDim Preserve arrPathFolders(Total_Folders)      '0 to Max -1
        ListSubfolders = arrPathFolders
    End If
    
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
    
    AppendErrorLogCustom "ListSubfolders - End"
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.ListSubfolders", "Path:", Path, "Recur:", Recursively
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
    If inIDE Then Stop: Resume Next
End Function


Private Sub ListSubfolders_Ex(Path As String, Optional Recursively As Boolean = False)
    On Error GoTo ErrorHandler
    
    Dim SubPathName     As String
    Dim PathName        As String
    Dim hFind           As Long
    Dim L               As Long
    Dim lpSTR           As Long
    Dim fd              As WIN32_FIND_DATA
    
    'Local module variables:
    '
    ' Total_Folders as long
    ' arrPathFolders() as string
    
    Do
        If hFind <> 0& Then
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: Exit Do
        Else
            hFind = FindFirstFile(StrPtr(Path & ch_SlashAsterisk), fd)  '"\*"
            If hFind = INVALID_HANDLE_VALUE Then Exit Do
        End If
        
        L = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT ' мимо симлинков
        Do While L <> 0&
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0: Exit Do
            L = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT
        Loop
    
        If hFind <> 0& Then
            lpSTR = VarPtr(fd.dwReserved1) + 4&
            PathName = Space$(lstrlen(lpSTR))
            lstrcpy StrPtr(PathName), lpSTR
        
            If fd.dwFileAttributes And vbDirectory Then
                If PathName <> ch_Dot Then  '"."
                    If PathName <> ch_DotDot Then '".."
                        SubPathName = Path & "\" & PathName
                        If UBound(arrPathFolders) < Total_Folders Then ReDim Preserve arrPathFolders(UBound(arrPathFolders) + 100&) As String
                        arrPathFolders(Total_Folders) = SubPathName
                        Total_Folders = Total_Folders + 1&
                        If Recursively Then
                            Call ListSubfolders_Ex(SubPathName, Recursively)
                        End If
                    End If
                End If
            End If
        End If
        
    Loop While hFind
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modFile.ListSubfolders", "Folder:", Path
    Resume Next
End Sub


'main function to list files
'
'ret - string array( 0 to MAX-1 ) or non-touched array if none.

Public Function ListFiles(Path As String, Optional Extension As String = "", Optional Recursively As Boolean = False) As String()
    On Error GoTo ErrorHandler

    AppendErrorLogCustom "ListFiles - Begin", "Path: " & Path, "Ext-s: " & Extension, "Recur: " & Recursively

    Dim bRedirStateChanged As Boolean, bOldState As Boolean
    'прежде, чем использовать ListFiles_Ex, нужно инициализировать глобальные массивы.
    ReDim arrPathFiles(100) As String
    'при каждом вызове ListFiles_Ex следует обнулить глобальный счетчик файлов
    Total_Files = 0&
    
    If bIsWin64 Then
        If StrBeginWith(Path, sWinDir) Then
            bRedirStateChanged = ToggleWow64FSRedirection(False, , bOldState)
        End If
    End If
    
    'вызов тушки
    Call ListFiles_Ex(Path, Extension, Recursively)
    If Total_Files > 0 Then
        Total_Files = Total_Files - 1
        ReDim Preserve arrPathFiles(Total_Files)      '0 to Max -1
        ListFiles = arrPathFiles
    End If
    
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
    
    AppendErrorLogCustom "ListFiles - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.ListFiles", "Path:", Path, "Ext-s:", Extension, "Recur:", Recursively
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
    If inIDE Then Stop: Resume Next
End Function


Private Sub ListFiles_Ex(Path As String, Optional Extension As String = "", Optional Recursively As Boolean = False)
    'Example of Extension:
    '".txt" - txt files
    'empty line - all files (by default)

    On Error GoTo ErrorHandler
    
    Dim SubPathName     As String
    Dim PathName        As String
    Dim hFind           As Long
    Dim L               As Long
    Dim lpSTR           As Long
    Dim fd              As WIN32_FIND_DATA
    
    'Local module variables:
    '
    ' Total_Files as long
    ' arrPathFiles() as string
    
    Do
        If hFind <> 0& Then
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: Exit Do
        Else
            hFind = FindFirstFile(StrPtr(Path & ch_SlashAsterisk), fd)  '"\*"
            If hFind = INVALID_HANDLE_VALUE Then Exit Do
        End If
        
        L = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT ' мимо симлинков
        Do While L <> 0&
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0: Exit Do
            L = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT
        Loop
    
        If hFind <> 0& Then
            lpSTR = VarPtr(fd.dwReserved1) + 4&
            PathName = Space$(lstrlen(lpSTR))
            lstrcpy StrPtr(PathName), lpSTR
        
            If fd.dwFileAttributes And vbDirectory Then
                If PathName <> ch_Dot Then  '"."
                    If PathName <> ch_DotDot Then '".."
                        SubPathName = Path & "\" & PathName
                        If Recursively Then
                            Call ListFiles_Ex(SubPathName, Extension, Recursively)
                        End If
                    End If
                End If
            Else
                If inArray(GetExtensionName(PathName), SplitSafe(Extension, ";"), , , 1) Or Len(Extension) = 0 Then
                    SubPathName = Path & "\" & PathName
                    If UBound(arrPathFiles) < Total_Files Then ReDim Preserve arrPathFiles(UBound(arrPathFiles) + 100&) As String
                    arrPathFiles(Total_Files) = SubPathName
                    Total_Files = Total_Files + 1&
                    If Not bAutoLogSilent Then
                        If Total_Files Mod 1000 = 0 Then DoEvents
                    End If
                End If
            End If
        End If
    Loop While hFind
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modFile.ListFiles_Ex", "Folder:", Path
    Resume Next
End Sub

Public Function GetLocalDisks$()
    Dim lDrives&, i&, sDrive$, sLocalDrives$
    lDrives = GetLogicalDrives()
    For i = 0 To 26
        If (lDrives And 2 ^ i) Then
            sDrive = Chr$(65 + i) & ":\"
            Select Case GetDriveType(StrPtr(sDrive))
                Case DRIVE_FIXED, DRIVE_RAMDISK: sLocalDrives = sLocalDrives & Chr$(65 + i) & " "
            End Select
        End If
    Next i
    GetLocalDisks = Trim$(sLocalDrives)
End Function

Public Function EnumFiles$(sFolder$)    'returns list of files divided by |
    Dim hFind&, sFile$, uWFD As WIN32_FIND_DATA, sList$, lpSTR&, bRedirStateChanged As Boolean, bOldState As Boolean
    
    If Not FolderExists(sFolder) Then Exit Function
    
    bRedirStateChanged = ToggleWow64FSRedirection(False, sFolder, bOldState)
    
    hFind = FindFirstFile(StrPtr(BuildPath(sFolder, "*.*")), uWFD)
    If hFind <> INVALID_HANDLE_VALUE Then
        Do
            lpSTR = VarPtr(uWFD.lpszFileName(0))
            sFile = Space$(lstrlen(lpSTR))
            lstrcpy StrPtr(sFile), lpSTR
            
            If sFile <> "." And sFile <> ".." Then
                sList = sList & "|" & sFile
            End If
            If bSL_Abort Then
                FindClose hFind
                GoTo Finalize
            End If
        Loop Until FindNextFile(hFind, uWFD) = 0
        FindClose hFind
        If sList <> vbNullString Then EnumFiles = Mid$(sList, 2)
    End If
    
Finalize:
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
End Function

Public Function GetLongFilename$(sFilename$)
    Dim sLongFilename$
    If InStr(sFilename, "~") = 0 Then
        GetLongFilename = sFilename
        Exit Function
    End If
    sLongFilename = String$(MAX_PATH, 0)
    GetLongPathNameW StrPtr(sFilename), StrPtr(sLongFilename), Len(sLongFilename)
    GetLongFilename = TrimNull(sLongFilename)
End Function

Public Function GetFilePropVersion(sFilename As String) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetFilePropVersion - Begin", "File: " & sFilename
    
    Dim hData&, lDataLen&, uBuf() As Byte
    Dim uVFFI As VS_FIXEDFILEINFO, sVersion$, Redirect As Boolean, bOldStatus As Boolean
    
    If Not FileExists(sFilename) Then
        AppendErrorLogCustom sFilename & " is not found. Err = " & Err.LastDllError
        Exit Function
    End If
    
    Redirect = ToggleWow64FSRedirection(False, sFilename, bOldStatus)
    
    lDataLen = GetFileVersionInfoSize(StrPtr(sFilename), ByVal 0&)
    If lDataLen = 0 Then
        AppendErrorLogCustom "lDataLen = 0. Err = " & Err.LastDllError
        GoTo Finalize
    End If
    
    ReDim uBuf(0 To lDataLen - 1)
    If 0 <> GetFileVersionInfo(StrPtr(sFilename), 0&, lDataLen, uBuf(0)) Then
    
        If 0 <> VerQueryValue(uBuf(0), StrPtr("\"), hData, lDataLen) Then
        
            If hData <> 0 Then
        
                CopyMemory uVFFI, ByVal hData, Len(uVFFI)
    
                With uVFFI
                    sVersion = .dwFileVersionMSh & "." & _
                        .dwFileVersionMSl & "." & _
                        .dwFileVersionLSh & "." & _
                        .dwFileVersionLSl
                End With
            Else
                AppendErrorLogCustom "hData = 0"
            End If
        Else
            AppendErrorLogCustom "VerQueryValue = 0. Err = " & Err.LastDllError
        End If
    Else
        AppendErrorLogCustom "GetFileVersionInfo = 0. Err = " & Err.LastDllError
    End If
    GetFilePropVersion = sVersion
    
Finalize:
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "GetFilePropVersion - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePropVersion", sFilename
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function GetVersionFromVBP(sFilename As String) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetVersionFromVBP - Begin", "File: " & sFilename
    
    Dim hFile As Long, sLine As String, arr() As String
    Dim MajorVer As Long, MinorVer As Long, BuildVer As Long, RevisionVer As Long
    
    OpenW sFilename, FOR_READ, hFile
    
    If hFile > 0 Then
        Do While LineInputW(hFile, sLine)
            arr = SplitSafe(sLine, "=")
            If UBound(arr) = 1 Then
                If StrComp(arr(0), "MajorVer", 1) = 0 Then
                    MajorVer = Val(arr(1))
                ElseIf StrComp(arr(0), "MinorVer", 1) = 0 Then
                    MinorVer = Val(arr(1))
                ElseIf StrComp(arr(0), "BuildVer", 1) = 0 Then
                    BuildVer = Val(arr(1))
                ElseIf StrComp(arr(0), "RevisionVer", 1) = 0 Then
                    RevisionVer = Val(arr(1))
                End If
            End If
        Loop
        CloseW hFile
    End If
    
    GetVersionFromVBP = MajorVer & "." & MinorVer & "." & BuildVer & "." & RevisionVer
    
    AppendErrorLogCustom "GetVersionFromVBP - End", "File: " & sFilename
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetVersionFromVBP", sFilename
    If inIDE Then Stop: Resume Next
End Function

Public Function GetValueFromVBP(sFilename As String, sParameterName As String) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetValueFromVBP - Begin", "File: " & sFilename
    
    Dim hFile As Long, sLine As String, arr() As String
    
    OpenW sFilename, FOR_READ, hFile
    
    If hFile > 0 Then
        Do While LineInputW(hFile, sLine)
            arr = SplitSafe(sLine, "=")
            If UBound(arr) = 1 Then
                If StrComp(arr(0), sParameterName, 1) = 0 Then
                    GetValueFromVBP = UnQuote(arr(1))
                    Exit Do
                End If
            End If
        Loop
        CloseW hFile
    End If
    
    AppendErrorLogCustom "GetValueFromVBP - End", "File: " & sFilename
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetValueFromVBP", sFilename
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFilePropCompany(sFilename As String) As String
    On Error GoTo ErrorHandler:
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, Stady&, Redirect As Boolean, bOldStatus As Boolean
    
    If Not FileExists(sFilename) Then Exit Function
    
    Redirect = ToggleWow64FSRedirection(False, sFilename, bOldStatus)
    
    Stady = 1
    lDataLen = GetFileVersionInfoSize(StrPtr(sFilename), ByVal 0&)
    If lDataLen = 0 Then GoTo Finalize
    
    Stady = 2
    ReDim uBuf(0 To lDataLen - 1)
    
    Stady = 3
    If 0 <> GetFileVersionInfo(StrPtr(sFilename), 0&, lDataLen, uBuf(0)) Then
        
        Stady = 4
        VerQueryValue uBuf(0), StrPtr("\VarFileInfo\Translation"), hData, lDataLen
        If lDataLen = 0 Then GoTo Finalize
        
        Stady = 5
        CopyMemory uCodePage(0), ByVal hData, 4
        
        Stady = 6
        sCodePage = Right$("0" & Hex$(uCodePage(1)), 2) & _
                Right$("0" & Hex$(uCodePage(0)), 2) & _
                Right$("0" & Hex$(uCodePage(3)), 2) & _
                Right$("0" & Hex$(uCodePage(2)), 2)
        
        'get CompanyName string
        Stady = 7
        If VerQueryValue(uBuf(0), StrPtr("\StringFileInfo\" & sCodePage & "\CompanyName"), hData, lDataLen) = 0 Then GoTo Finalize
    
        If lDataLen > 0 And hData <> 0 Then
            Stady = 8
            sCompanyName = String$(lDataLen, 0)
            
            Stady = 9
            lstrcpy ByVal StrPtr(sCompanyName), ByVal hData
        End If
        
        Stady = 10
        GetFilePropCompany = RTrimNull(sCompanyName)
    End If
    
Finalize:
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePropCompany", sFilename, "DataLen: ", lDataLen, "hData: ", hData, "sCodePage: ", sCodePage, _
        "Buf: ", uCodePage(0), uCodePage(1), uCodePage(2), uCodePage(3), "Stady: ", Stady
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

'get File's Property string, like:
'
'CompanyName
'FileDescription
'FileVersion
'InternalName
'LegalCopyright
'OriginalFilename
'ProductName
'ProductVersion
'
Public Function GetFileProperty(sFilename As String, sPropertyName As String) As String
    On Error GoTo ErrorHandler:
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sPropValue$, Redirect As Boolean, bOldStatus As Boolean
    
    If Not FileExists(sFilename) Then Exit Function
    
    Redirect = ToggleWow64FSRedirection(False, sFilename, bOldStatus)
    
    lDataLen = GetFileVersionInfoSize(StrPtr(sFilename), ByVal 0&)
    If lDataLen = 0 Then GoTo Finalize
    
    ReDim uBuf(0 To lDataLen - 1)
    
    If 0 <> GetFileVersionInfo(StrPtr(sFilename), 0&, lDataLen, uBuf(0)) Then
        
        VerQueryValue uBuf(0), StrPtr("\VarFileInfo\Translation"), hData, lDataLen
        If lDataLen = 0 Then GoTo Finalize
        
        CopyMemory uCodePage(0), ByVal hData, 4
        
        sCodePage = Right$("0" & Hex$(uCodePage(1)), 2) & _
                Right$("0" & Hex$(uCodePage(0)), 2) & _
                Right$("0" & Hex$(uCodePage(3)), 2) & _
                Right$("0" & Hex$(uCodePage(2)), 2)
        
        If VerQueryValue(uBuf(0), StrPtr("\StringFileInfo\" & sCodePage & "\" & sPropertyName), hData, lDataLen) = 0 Then GoTo Finalize
    
        If lDataLen > 0 And hData <> 0 Then
            sPropValue = String$(lDataLen, 0)
            lstrcpy ByVal StrPtr(sPropValue), ByVal hData
        End If

        GetFileProperty = RTrimNull(sPropValue)
    End If
    
Finalize:
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePropCompany", sFilename, sPropertyName
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function DirW( _
    Optional ByVal PathMaskOrFolderWithSlash As String, _
    Optional AllowedAttributes As VbFileAttributeExtended = vbNormal, _
    Optional FoldersOnly As Boolean) As String
    
    On Error GoTo ErrorHandler
    
    'WARNING note:
    'Original VB DirW$ contains bug: ReadOnly attribute incorrectly handled, so it always is in results
    'This sub properly handles 'RO' and also contains one extra flag: FILE_ATTRIBUTE_REPARSE_POINT (vbReparse)
    'Doesn't return "." and ".." folders.
    'Unicode aware
    
    Const MeaningfulBits As Long = &H417&   'D + H + R + S + Reparse
                                            '(to revert to default VB Dir behaviour, replace it by &H16 value)
    
    Dim fd      As WIN32_FIND_DATA
    Dim lpSTR   As Long
    Dim Mask    As Long
    
    Static hFind        As Long
    Static lFlags       As VbFileAttributeExtended
    Static bFoldersOnly As Boolean
    
    If hFind <> 0& And Len(PathMaskOrFolderWithSlash) = 0& Then
        If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0&: Exit Function
    Else
        If hFind Then FindClose hFind: hFind = 0&
        PathMaskOrFolderWithSlash = Trim$(PathMaskOrFolderWithSlash)
        lFlags = AllowedAttributes 'cache
        bFoldersOnly = FoldersOnly 'cache
        
        Select Case Right$(PathMaskOrFolderWithSlash, 1&)
        Case "", ":", "/"
            PathMaskOrFolderWithSlash = PathMaskOrFolderWithSlash & "*.*"
        End Select
        
        hFind = FindFirstFile(StrPtr(PathMaskOrFolderWithSlash), fd)
        
        If hFind = INVALID_HANDLE_VALUE Then
            If (Err.LastDllError) > 12& Then hFind = 0&: Err.Raise 52&
            Exit Function
        End If
    End If
    
    Do
        If fd.dwFileAttributes = FILE_ATTRIBUTE_NORMAL Then
            Mask = 0& 'found
        Else
            Mask = fd.dwFileAttributes And (Not lFlags) And MeaningfulBits
        End If
        If bFoldersOnly Then
            If Not CBool(fd.dwFileAttributes And vbDirectory) Then
                Mask = 1 'continue enum
            End If
        End If
    
        If Mask = 0 Then
            lpSTR = VarPtr(fd.lpszFileName(0))
            DirW = String$(lstrlen(lpSTR), 0&)
            lstrcpy StrPtr(DirW), lpSTR
            If fd.dwFileAttributes And vbDirectory Then
                If DirW <> "." And DirW <> ".." Then Exit Do 'exclude self and relative paths aliases
            Else
                Exit Do
            End If
        End If
    
        If FindNextFile(hFind, fd) = 0 Then FindClose hFind: hFind = 0: Exit Function
    Loop
    
    Exit Function
ErrorHandler:
    Debug.Print Err; Err.Description; "DirW"
End Function

Public Function GetEmptyName(ByVal sFullPath As String) As String

    Dim sExt As String
    Dim sName As String
    Dim sPath As String
    Dim i As Long

    If Not FileExists(sFullPath) Then
        GetEmptyName = sFullPath
    Else
        sExt = GetExtensionName(sFullPath)
        sPath = GetPathName(sFullPath)
        sName = GetFileName(sFullPath)
        Do
            i = i + 1
            sFullPath = BuildPath(sPath, sName & "(" & i & ")" & sExt)
        Loop While FileExists(sFullPath)
        
        GetEmptyName = sFullPath
    End If
End Function

Public Function GetFileDate(Optional File As String, Optional Date_Type As ENUM_FILE_DATE_TYPE, Optional hFile As Long) As Date
    On Error GoTo ErrorHandler
    
    Dim rval        As Long
    Dim ctime       As FILETIME
    Dim atime       As FILETIME
    Dim wtime       As FILETIME
    Dim sTime       As SYSTEMTIME
    Dim bOldRedir   As Boolean
    Dim bExternalHandle As Boolean
    
    AppendErrorLogCustom "Parser.GetFileDate - Begin: " & File
    
    If hFile <= 0 Then
        ToggleWow64FSRedirection False, File, bOldRedir
    
        hFile = CreateFile(StrPtr(File), ByVal 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    
        ToggleWow64FSRedirection bOldRedir
    Else
        bExternalHandle = True
    End If
    
    If hFile <> INVALID_HANDLE_VALUE Then
        rval = GetFileTime(hFile, ctime, atime, wtime)
        Select Case Date_Type
        Case DATE_MODIFIED
            rval = FileTimeToLocalFileTime(wtime, wtime)
            rval = FileTimeToSystemTime(wtime, sTime)
        Case DATE_CREATED
            rval = FileTimeToLocalFileTime(ctime, ctime)
            rval = FileTimeToSystemTime(ctime, sTime)
        Case DATE_ACCESSED
            rval = FileTimeToLocalFileTime(atime, atime)
            rval = FileTimeToSystemTime(atime, sTime)
        End Select
        SystemTimeToVariantTime sTime, GetFileDate
        'GetFileDate = DateSerial(ftime.wYear, ftime.wMonth, ftime.wDay) + TimeSerial(ftime.wHour, ftime.wMinute, ftime.wSecond)
        If Not bExternalHandle Then
            CloseHandle hFile
        End If
    End If
    
    AppendErrorLogCustom "Parser.GetFileDate - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFileDate", "File: " & File
    If inIDE Then Stop: Resume Next
End Function

Public Function IsFileOneMonthModified(sFile As String) As Boolean

    Dim bOldRedir   As Boolean
    Dim hFile       As Long

    ToggleWow64FSRedirection False, sFile, bOldRedir
    
    hFile = CreateFile(StrPtr(sFile), ByVal 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    
    If DateDiff("d", GetFileDate(, DATE_CREATED, hFile), Now) < 31 Then
        IsFileOneMonthModified = True
    ElseIf DateDiff("d", GetFileDate(, DATE_MODIFIED, hFile), Now) < 31 Then
        IsFileOneMonthModified = True
    End If
    
    ToggleWow64FSRedirection bOldRedir
End Function

'check file on Portable Executable
Public Function isPE(sFile As String) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim bOldRedir As Boolean
    Dim hFile As Long
    Dim hMapping As Long
    Dim pBuf As Long
    
    ToggleWow64FSRedirection False, sFile, bOldRedir
    
    hFile = CreateFile(StrPtr(sFile), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    
    ToggleWow64FSRedirection bOldRedir
    
    If hFile <> INVALID_HANDLE_VALUE Then
    
        hMapping = CreateFileMapping(hFile, 0&, PAGE_READONLY Or SEC_IMAGE, 0&, 0&, 0&)
        
        CloseHandle hFile
        
        If hMapping <> 0 Then
            
            pBuf = MapViewOfFile(hMapping, FILE_MAP_READ, 0&, 0&, 0&)
            
            If pBuf <> 0 Then
            
                isPE = True
                UnmapViewOfFile pBuf
            End If
            
            CloseHandle hMapping
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "isPE"
    If inIDE Then Stop: Resume Next
End Function

Private Function StrBeginWith(Text As String, BeginPart As String) As Boolean
    StrBeginWith = (StrComp(Left$(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function

Private Function SplitSafe(sComplexString As String, Optional Delimiter As String = " ") As String()
    If 0 = Len(sComplexString) Then
        ReDim arr(0) As String
        SplitSafe = arr
    Else
        SplitSafe = Split(sComplexString, Delimiter)
    End If
End Function

' ¬озвращает true, если искомое значение найдено в одном из элементов массива (lB, uB ограничивает просматриваемый диапазон индексов)
Private Function inArray( _
    Stri As String, _
    MyArray() As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    If lB = -2147483647 Then lB = LBound(MyArray)   'some trick
    If uB = 2147483647 Then uB = UBound(MyArray)    'Thanks to  азанский :)
    Dim i As Long
    For i = lB To uB
        If StrComp(Stri, MyArray(i), CompareMethod) = 0 Then inArray = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "inArray"
    If inIDE Then Stop: Resume Next
End Function

'Public Function BuildPath$(sPath$, sFile$)
'    BuildPath = sPath & IIf(Right$(sPath, 1) = "\", vbNullString, IIf(Len(sFile) = 0, vbNullString, "\")) & sFile
'End Function

Public Function BuildPath(ParamArray Paths()) As String
    Dim i As Long
    For i = 0 To UBound(Paths)
        BuildPath = BuildPath & IIf(Right$(BuildPath, 1) = "\", vbNullString, "\") & Paths(i)
    Next
    BuildPath = Mid$(BuildPath, 2)
End Function

'To stop on the first NULL char occurrence
Public Function TrimNull(s$) As String
    TrimNull = Left$(s, lstrlen(StrPtr(s)))
End Function

'To check for NULL char from right to left
Public Function RTrimNull(s$) As String
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) <> vbNullChar Then
            RTrimNull = Left$(s, i)
            Exit Function
        End If
    Next
End Function

Public Function GetRootPath(Path As String) As String
    Dim pos As Long
    If InStr(Path, ":") <> 0 Then
        GetRootPath = Left$(Path, 2)
    Else
        pos = InStr(Path, "\")
        If pos <> 0 Then
            GetRootPath = Left(Path, pos - 1)
        Else
            GetRootPath = Path
        End If
    End If
End Function

Public Function GetPathName(Path As String) As String   ' получить родительский каталог
    Dim pos As Long
    pos = InStrRev(Path, "\")
    If pos <> 0 Then GetPathName = Left$(Path, pos - 1)
End Function

' ѕолучить им€ файла
Public Function GetFileName(ByVal Path As String, Optional bWithExtension As Boolean) As String
    On Error GoTo ErrorHandler
    Dim posDot      As Long
    Dim posSl       As Long
    Dim posColon    As Long
    
    'trim ADS
    posColon = InStr(4, Path, ":")
    If posColon Then
        'not URL ?
        If Mid$(Path, posColon, 3) <> ":\\" Then
            Path = Left$(Path, posColon - 1)
        End If
    End If
    
    posSl = InStrRev(Path, "\")
    If Not bWithExtension Then
        If posSl <> 0 Then
            posDot = InStrRev(Path, ".")
            If posDot < posSl Then posDot = 0
        Else
            posDot = InStrRev(Path, ".")
        End If
    End If
    If posDot = 0 Then posDot = Len(Path) + 1
    
    GetFileName = Mid$(Path, posSl + 1, posDot - posSl - 1)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFileName", "Path: ", Path
End Function

Private Function MapDriveTypeToDriveTypeBit(DriveType As DRIVE_TYPE) As DRIVE_TYPE_BIT
    Dim dtb As DRIVE_TYPE_BIT
    
    Select Case DriveType
    Case DRIVE_UNKNOWN = 0
        dtb = DRIVE_BIT_UNKNOWN
    Case DRIVE_NO_ROOT_DIR
        dtb = DRIVE_BIT_NO_ROOT_DIR
    Case DRIVE_REMOVABLE
        dtb = DRIVE_BIT_REMOVABLE
    Case DRIVE_FIXED
        dtb = DRIVE_BIT_FIXED
    Case DRIVE_REMOTE
        dtb = DRIVE_BIT_REMOTE
    Case DRIVE_CDROM
        dtb = DRIVE_BIT_CDROM
    Case DRIVE_RAMDISK
        dtb = DRIVE_BIT_RAMDISK
    Case DRIVE_ANY
        dtb = DRIVE_BIT_ANY
    End Select
    
    MapDriveTypeToDriveTypeBit = dtb
End Function

Public Function GetDrives(Optional DriveTypeBit As DRIVE_TYPE_BIT = DRIVE_BIT_ANY) As String()
    On Error GoTo ErrorHandler
   
    Dim BufLen          As Long
    Dim buf             As String
    Dim i               As Long
    Dim Drives()        As String
    Dim ReadyDrives()   As String
    Dim idx             As Long
    Dim hDevice         As Long
    Dim cbBytesReturned As Long
    Dim curDriveType    As Long
    Dim lControlCode    As Long

    buf = String$(MAX_PATH, 0)

    'получаем список всех букв дисков в системе
    BufLen = GetLogicalDriveStrings(MAX_PATH, StrPtr(buf))
   
    If BufLen <> 0 Then
        buf = Left$(buf, BufLen - 1)
        Drives = Split(buf, vbNullChar)
       
        ReDim ReadyDrives(UBound(Drives) + 1)
       
        For i = 0 To UBound(Drives)
            Drives(i) = Left$(Drives(i), 2)
        
            hDevice = CreateFile(StrPtr("\\.\" & Drives(i)), _
                             FILE_READ_ATTRIBUTES, _
                             FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                             ByVal 0&, OPEN_EXISTING, 0&, 0&)
       
            If hDevice <> 0 Then
           
                If StrComp(Drives(i), "A:", 1) = 0 Or StrComp(Drives(i), "B:", 1) = 0 Then
                    lControlCode = IOCTL_STORAGE_CHECK_VERIFY
                Else
                    lControlCode = IOCTL_STORAGE_CHECK_VERIFY2
                End If
           
                'провер€ем готово ли устройство (вставлен ли диск)
                If DeviceIoControl(hDevice, _
                            lControlCode, _
                             ByVal 0&, 0&, _
                            0&, 0&, _
                            cbBytesReturned, _
                            0&) Then
                    
                    curDriveType = GetDriveType(StrPtr(Drives(i)))
                    
                    If (DriveTypeBit And DRIVE_BIT_ANY) Or (DriveTypeBit And MapDriveTypeToDriveTypeBit(curDriveType)) Then
                    
                        idx = idx + 1
                        ReadyDrives(idx) = Drives(i)
                    ElseIf (DriveTypeBit And DRIVE_BIT_REMOTE) And curDriveType = DRIVE_NO_ROOT_DIR Then
                    
                        If PathIsNetworkPath(StrPtr(Drives(i))) Then    'WARNING: this doesn't work for cmd's subst
                            idx = idx + 1
                            ReadyDrives(idx) = Drives(i)
                        Else
                            'disconnected disk
                        End If
                    End If
                End If
               
                CloseHandle hDevice
            End If
        Next
    End If
   
    If idx > 0 Then
        ReDim Preserve ReadyDrives(idx)
    Else
        ReDim ReadyDrives(0)
    End If
   
    GetDrives = ReadyDrives
    
    Exit Function
ErrorHandler:
    Debug.Print Now, Err, "modFile.GetDrives"
End Function



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
    Dim s$, pos&
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

Public Function ReadFileToArray(sFile As String, Optional isUnicode As Boolean) As String()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ReadFileToArray - Begin", "File: " & sFile
    Dim Text  As String
    Text = Replace$(ReadFileContents(sFile, isUnicode), vbCr, vbNullString)
    If Len(Text) <> 0 Then
        ReadFileToArray = SplitSafe(Text, vbLf)
    End If
    AppendErrorLogCustom "ReadFileToArray - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ReadFileToArray"
    If inIDE Then Stop: Resume Next
End Function

Public Function ReadFileContents(sFile As String, isUnicode As Boolean) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ReadFileContents - Begin", "File: " & sFile
    Dim hFile   As Long
    Dim b()     As Byte
    Dim Text    As String
    Dim lSize   As Currency
    Dim Redirect As Boolean, bOldStatus As Boolean
    If Not FileExists(sFile) Then Exit Function
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    OpenW sFile, FOR_READ, hFile, g_FileBackupFlag
    If hFile <= 0 Then Exit Function
    lSize = LOFW(hFile)
    If lSize = 0 Then CloseW hFile: Exit Function
    ReDim b(lSize - 1)
    GetW hFile, 1, , VarPtr(b(0)), UBound(b) + 1
    CloseW hFile
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If isUnicode Then
        ReadFileContents = b()
        If UBound(b) >= 1 Then
            If b(0) = &HFF& And b(1) = &HFE& Then ReadFileContents = Mid$(ReadFileContents, 2)  ' - BOM UTF16-LE
        End If
    Else
        ReadFileContents = StrConv(b(), vbUnicode, OSver.LangNonUnicodeCode)
        If UBound(b) >= 2 Then
            If b(0) = &HEF& And b(1) = &HBB& And b(2) = &HBF& Then      ' - BOM UTF-8
                ReadFileContents = Mid$(ReadFileContents, 4)
            End If
        End If
    End If
    AppendErrorLogCustom "ReadFileContents - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ReadFileContents"
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function


Public Function IniGetString( _
    sFile As String, _
    sSection As String, _
    sParameter As String, _
    Optional vDefault As Variant, _
    Optional isUnicode As Variant, _
    Optional bMultiple As Boolean = False) As String
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "IniGetString - Begin", "File: " & sFile
    
    Dim sIniFile$, i&, aContents() As String, sData$, hFile As Long, sText As String
    Dim Redirect As Boolean, bOldStatus As Boolean
    
    'if bMultiple == true, get several "values | values" from the same parameter's names
    
    If Not IsMissing(vDefault) Then
        IniGetString = vDefault
    End If
    
    If Not FileExists(sFile) Then
        sIniFile = FindOnPath(sFile, False)

'        'absolute path -> exit
'        If InStr(sFile, "\") <> 0 Then Exit Function
'        If FileExists(sWinSysDir & "\" & sFile) Then
'            sIniFile = sWinSysDir & "\" & sFile
'        ElseIf FileExists(sWinDir & "\" & sFile) Then
'            sIniFile = sWinDir & "\" & sFile
'        End If
        If 0 = Len(sIniFile) Then Exit Function
    Else
        sIniFile = sFile
    End If
    
    If FileLenW(sIniFile) = 0 Then Exit Function 'size == 0
    
    If IsMissing(isUnicode) Then
        isUnicode = (FileGetTypeBOM(sIniFile) = 1200)
    End If
    
    aContents = ReadFileToArray(sIniFile, CBool(isUnicode))
    If 0 = AryItems(aContents) Then Exit Function 'file is empty
    
    'find index of section
    Do Until StrComp(Trim$(aContents(i)), "[" & sSection & "]", vbTextCompare) = 0
        i = i + 1
        If i > UBound(aContents) Then Exit Function
    Loop
    'next line
    i = i + 1
    If i <= UBound(aContents) Then
      'within current section
      Do Until Left$(LTrim$(aContents(i)), 1) = "["
        'if string begin with our parameter
        If InStr(1, aContents(i), sParameter, vbTextCompare) = 1 Then
            'if next char is =, excluding space characters after parameter's name
            If Left$(LTrim$(Mid$(aContents(i), Len(sParameter) + 1)), 1) = "=" Then
                'appending sData with value
                If bMultiple Then
                    sData = sData & "|" & Mid$(aContents(i), InStr(aContents(i), "=") + 1)
                Else
                    IniGetString = Mid$(aContents(i), InStr(aContents(i), "=") + 1)
                    Exit Function
                End If
            End If
        End If
        'next line
        i = i + 1
        'eof
        If i > UBound(aContents) Then Exit Do
      Loop
    End If
    
    If Len(sData) <> 0 Then
        IniGetString = Mid$(sData, 2)
    End If
    
    AppendErrorLogCustom "IniGetString - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IniGetString", "sFile=", sFile, "sSection=", sSection, "sParameter=", sParameter
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function


Public Function IniSetString( _
    sFile As String, _
    sSection As String, _
    sParameter As String, _
    vData As Variant, _
    Optional ByVal isUnicode As Variant) As Boolean
    
    On Error GoTo ErrorHandler:
    Dim sIniFile$, i&, iSect&, aContents() As String, sNewData$
    Dim Redirect As Boolean, bOldStatus As Boolean, sDir As String
    
    sIniFile = sFile
    
    If Not FileExists(sFile) Then
        'relative path -> search %windir%, %windir%\System32
        If InStr(sFile, "\") = 0 Then
            If FileExists(sWinSysDir & "\" & sFile) Then
                sIniFile = sWinSysDir & "\" & sFile
            ElseIf FileExists(sWinDir & "\" & sFile) Then
                sIniFile = sWinDir & "\" & sFile
            End If
        End If
    End If
    
    sDir = GetParentDir(sIniFile)
    
    If Not FolderExists(sDir) Then
        If Not MkDirW(sDir) Then
            TryUnlock sDir
            If Not MkDirW(sDir) Then
                If Not bAutoLogSilent Then
                    'Could not create folder '[]'. Please verify that write access is allowed to this location.
                    MsgBoxW Replace$(Translate(1022), "[]", sDir), vbCritical
                End If
                Exit Function
            End If
        End If
    End If
    
    If Not CheckAccessWrite(sIniFile) Then
        TryUnlock sIniFile
        If Not CheckAccessWrite(sIniFile) Then
            If Not bAutoLogSilent Then
                'The value '[*]' could not be written to the settings file '[**]'. Please verify that write access is allowed to that file.
                MsgBoxW Replace$(Replace$(Translate(1008), "[*]", vData), "[**]", sIniFile), vbCritical
            End If
            Exit Function
        End If
    End If
    
    If IsMissing(isUnicode) Then
        isUnicode = (FileGetTypeBOM(sIniFile) = 1200)
    End If
    
    aContents = ReadFileToArray(sIniFile, CBool(isUnicode))
    
    sNewData = sParameter & "=" & vData
    
    If 0 = AryItems(aContents) Then  'file is empty
        sNewData = "[" & sSection & "]" & vbCrLf & sNewData
        IniSetString = WriteDataToFile(sIniFile, sNewData, CBool(isUnicode), True)
        Exit Function
    End If
    
    For i = 0 To UBound(aContents)
        If Len(Trim$(aContents(i))) <> 0 Then
            If StrComp(aContents(i), "[" & sSection & "]", vbTextCompare) = 0 Then
                'found the correct section
                iSect = i
                Exit For
            End If
        End If
    Next i
    If i = UBound(aContents) + 1 Then   'section not found
        sNewData = "[" & sSection & "]" & vbCrLf & sParameter & "=" & vData
        IniSetString = WriteDataToFile(sIniFile, Join(aContents, vbCrLf) & vbCrLf & sNewData, CBool(isUnicode), True)
        Exit Function
    End If
    
    For i = iSect + 1 To UBound(aContents)
        
        'begin new section?
        If InStr(aContents(i), "[") = 1 Then
            'parameter not found ("[" - mean next section)
            'inserting two lines into one
            aContents(i) = sNewData & vbCrLf & aContents(i)
        
            'input new data, replace file
            IniSetString = WriteDataToFile(sIniFile, Join(aContents, vbCrLf), CBool(isUnicode), True)
            Exit Function
        End If
        
        'found parameter?
        If InStr(1, aContents(i), sParameter, vbTextCompare) = 1 Then
            
            'if next char is =, excluding space characters after parameter's name
            If Left$(LTrim$(Mid$(aContents(i), Len(sParameter) + 1)), 1) = "=" Then
            
                aContents(i) = sNewData
                'input new data, replace file
                IniSetString = WriteDataToFile(sIniFile, Join(aContents, vbCrLf), CBool(isUnicode), True)
                Exit Function
            End If
        End If
    Next i
    
    'It is a last section, but no value in it -> just add a value
    IniSetString = WriteDataToFile(sIniFile, Join(aContents, vbCrLf) & vbCrLf & sNewData, CBool(isUnicode), True)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IniSetString", "sFile=", sFile, "sSection=", sSection, "sParameter=", sParameter, "sData=", vData
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function IniRemoveString( _
    sFile As String, _
    sSection As String, _
    sParameter As String, _
    Optional ByVal isUnicode As Variant) As Boolean
    
    On Error GoTo ErrorHandler:
    Dim sIniFile$, i&, iSect&, aContents() As String
    Dim Redirect As Boolean, bOldStatus As Boolean
    
    sIniFile = sFile
    
    If Not FileExists(sFile) Then
        'relative path -> search %windir%, %windir%\System32
        If InStr(sFile, "\") = 0 Then
            If FileExists(sWinSysDir & "\" & sFile) Then
                sIniFile = sWinSysDir & "\" & sFile
            ElseIf FileExists(sWinDir & "\" & sFile) Then
                sIniFile = sWinDir & "\" & sFile
            Else
                IniRemoveString = True
                Exit Function
            End If
        Else
            IniRemoveString = True
            Exit Function
        End If
    End If

    If Not CheckAccessWrite(sIniFile) Then
        TryUnlock sIniFile
        If Not CheckAccessWrite(sIniFile) Then
            If Not bAutoLogSilent Then
                'The parameter '[*]' could not be removed from the settings file '[**]'. Please verify that write access is allowed for that file.
                MsgBoxW Replace$(Replace$(Translate(1009), "[*]", sParameter), "[**]", sIniFile), vbCritical
            End If
            Exit Function
        End If
    End If
    
    If FileLenW(sIniFile) = 0 Then
        IniRemoveString = True
        Exit Function
    End If
    
    If IsMissing(isUnicode) Then
        isUnicode = (FileGetTypeBOM(sIniFile) = 1200)
    End If
    
    aContents = ReadFileToArray(sIniFile, CBool(isUnicode))
    
    If 0 = AryItems(aContents) Then  'file is empty
        IniRemoveString = True
        Exit Function
    End If
    
    For i = 0 To UBound(aContents)
        If Len(Trim$(aContents(i))) <> 0 Then
            If StrComp(aContents(i), "[" & sSection & "]", vbTextCompare) = 0 Then
                'found the correct section
                iSect = i
                Exit For
            End If
        End If
    Next i
    If i = UBound(aContents) + 1 Then   'section not found
        IniRemoveString = True
        Exit Function
    End If
    
    For i = iSect + 1 To UBound(aContents)
        
        'begin new section?
        If InStr(aContents(i), "[") = 1 Then
            'parameter not found ("[" - mean next section)
            IniRemoveString = True
            Exit Function
        End If
        
        'found parameter?
        If InStr(1, aContents(i), sParameter, vbTextCompare) = 1 Then
            
            'if next char is =, excluding space characters after parameter's name
            If Left$(LTrim$(Mid$(aContents(i), Len(sParameter) + 1)), 1) = "=" Then
                'erase parameter
                aContents(i) = ""
                'replace file
                IniRemoveString = WriteDataToFile(sIniFile, Join(aContents, vbCrLf), CBool(isUnicode), True)
                Exit Function
            End If
        End If
    Next i
    
    'It is a last section, but no value in it -> do nothing
    IniRemoveString = True
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IniRemoveString", "sFile=", sFile, "sSection=", sSection, "sParameter=", sParameter
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function WriteDataToFile(sFile As String, sContents As String, Optional isUnicode As Boolean, Optional bShowWarning As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim hFile As Long
    Dim Redirect As Boolean, bOldStatus As Boolean
    Dim iAttr As Long
    
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    iAttr = GetFileAttributes(StrPtr(sFile))
    If (iAttr And 2048) Then iAttr = iAttr And Not 2048
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    If 0 = DeleteFileWEx(StrPtr(sFile)) Then
        If (Not bAutoLogSilent) And bShowWarning Then
            'The value '[*]' could not be written to the settings file '[**]'. Please verify that write access is allowed to that file.
            MsgBoxW Replace$(Replace$(Translate(1008), "[*]", sContents), "[**]", sFile), vbCritical
        End If
        Exit Function
    End If
    
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    OpenW sFile, FOR_OVERWRITE_CREATE, hFile, g_FileBackupFlag
    If hFile > 0 Then
        If isUnicode Then
            PutW hFile, 1, StrPtr(ChrW$(-257)), 2
            WriteDataToFile = PutW(hFile, 3, StrPtr(sContents), LenB(sContents), True)
        Else
            WriteDataToFile = PutW(hFile, 1, StrPtr(StrConv(sContents, vbFromUnicode)), Len(sContents))
        End If
        CloseW hFile, True
    End If
    
    If iAttr <> 0 Then
        SetFileAttributes StrPtr(sFile), iAttr
    End If
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "WriteDataToFile", "sFile=", sFile, "sContents=", sContents, "iAttr=", iAttr, "bShowWarning=", bShowWarning, "isUnicode=", isUnicode
    CloseW hFile, True
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function ReadIniA(sFile As String, sSectionNoBrackets As String, sParameter As String, Optional vDefault As Variant = vbNullString) As String
    On Error GoTo ErrorHandler

    'NOTE: sSection must be specified without [ ]

    Dim buf As String
    Dim lr As Long
    
    buf = String$(256&, 0)
    lr = GetPrivateProfileString(StrPtr(sSectionNoBrackets), StrPtr(sParameter), StrPtr(CStr(vDefault)), StrPtr(buf), Len(buf), StrPtr(sFile))
    If Err.LastDllError = ERROR_MORE_DATA Then
        buf = String$(1001&, 0)
        lr = GetPrivateProfileString(StrPtr(sSectionNoBrackets), StrPtr(sParameter), StrPtr(CStr(vDefault)), StrPtr(buf), Len(buf), StrPtr(sFile))
        If Err.LastDllError = ERROR_MORE_DATA Then
            buf = String$(10001&, 0)
            lr = GetPrivateProfileString(StrPtr(sSectionNoBrackets), StrPtr(sParameter), StrPtr(CStr(vDefault)), StrPtr(buf), Len(buf), StrPtr(sFile))
        End If
    End If
    If lr = 0 Then
        ReadIniA = vDefault
    Else
        ReadIniA = Left$(buf, lr)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ReadIni", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function

Public Function WriteIniA(sFile As String, sSection As String, sParameter As String, vData As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sFolder As String
    sFolder = GetParentDir(sFile)
    If Not FolderExists(sFolder) Then MkDirW sFolder
    
    WriteIniA = WritePrivateProfileString(StrPtr(sSection), StrPtr(sParameter), StrPtr(CStr(vData)), StrPtr(sFile))
    Exit Function
ErrorHandler:
    ErrorMsg Err, "WriteIni", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFileNameAndExt(ByVal Path As String) As String ' вернет только им€ файла вместе с расширением
    Dim pos As Long
    Dim posColon    As Long
    
    'trim ADS
    posColon = InStr(4, Path, ":")
    If posColon Then
        Path = Left$(Path, posColon - 1)
    End If
    
    pos = InStrRev(Path, "\")
    If pos <> 0 Then
        GetFileNameAndExt = Mid$(Path, pos + 1)
    Else
        GetFileNameAndExt = Path
    End If
End Function

Public Function PathX64(sPath As String) As String
    If OSver.IsWin32 Then
        PathX64 = sPath
    Else
        If StrBeginWith(sPath, sWinSysDir) Then
            PathX64 = Replace$(sPath, sWinSysDir, sSysNativeDir, , 1, 1)
        Else
            PathX64 = sPath
        End If
    End If
End Function

'true if success
Public Function FileCopyW(FileSource As String, FileDestination As String, Optional bOverwrite As Boolean = True) As Boolean
    On Error GoTo ErrorHandler:
    Dim bOldRedir As Boolean
    Dim sFolder As String
    
    If Not FileExists(FileSource) Then Exit Function
    ToggleWow64FSRedirection False, FileSource, bOldRedir
    ToggleWow64FSRedirection False, FileDestination
    
    sFolder = GetParentDir(FileDestination)
    
    If Not FolderExists(sFolder) Then
        If Not MkDirW(sFolder) Then Exit Function
    End If
    
    FileCopyW = CopyFile(StrPtr(FileSource), StrPtr(FileDestination), Not bOverwrite)
    If Not FileCopyW Then
        TryUnlock FileDestination
        FileCopyW = CopyFile(StrPtr(FileSource), StrPtr(FileDestination), Not bOverwrite)
        If Not FileCopyW Then
            If DeleteFileWEx(StrPtr(FileDestination), , True) Then
                FileCopyW = CopyFile(StrPtr(FileSource), StrPtr(FileDestination), Not bOverwrite)
            End If
        End If
    End If
    ToggleWow64FSRedirection bOldRedir
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FileCopyW", "Source:", FileSource, "Destination:", FileDestination
    ToggleWow64FSRedirection bOldRedir
    If inIDE Then Stop: Resume Next
End Function



Public Function FindOnPath(ByVal sAppName As String, Optional bUseSourceValueOnFailure As Boolean, Optional sAdditionalDir As String) As String
    On Error GoTo ErrorHandler:
    
    '// TODO:
    'replace PathFindOnPath API by manual parsing of %PATH%, because PathFindOnPath includes folders while searching, that is unappropriate

    AppendErrorLogCustom "FindOnPath - Begin"

    Static Exts() As String
    Static isInit As Boolean
    Dim ProcPath$
    Dim sFile As String
    Dim sFolder As String
    Dim pos As Long
    Dim i As Long
    Dim sFileTry As String
    Dim bFullPath As Boolean
    Dim bSuccess As Boolean

    If Not isInit Then
        isInit = True
        Exts = Split(EnvironW("%PathExt%"), ";")
        For i = 0 To UBound(Exts)
            Exts(i) = LCase$(Exts(i))
        Next
    End If

    If Len(sAppName) = 0 Then Exit Function

    If Left$(sAppName, 1) = """" Then
        If Right$(sAppName, 1) = """" Then
            sAppName = UnQuote(sAppName)
        End If
    End If

    If Mid$(sAppName, 2, 1) = ":" Then bFullPath = True

    If bFullPath Then
        If FileExists(sAppName) Then
            FindOnPath = sAppName
            Exit Function
        End If
    Else
        If 0 <> Len(sAdditionalDir) Then
            sFileTry = BuildPath(sAdditionalDir, sAppName)
            If FileExists(sFileTry) Then
                FindOnPath = sFileTry
                Exit Function
            End If
        End If
    End If

    pos = InStrRev(sAppName, "\")

    If bFullPath And pos <> 0 Then
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

        If InStr(sAppName, ".") <> 0 Then
            ProcPath = Space$(MAX_PATH)
            LSet ProcPath = sAppName & vbNullChar
        
            If CBool(PathFindOnPath(StrPtr(ProcPath), 0&)) Then
                FindOnPath = TrimNull(ProcPath)
                If FileExists(FindOnPath) Then 'if not a folder
                    bSuccess = True
                End If
            End If
        End If
        
        If Not bSuccess Then
            'go through the extensions list

            For i = 0 To UBound(Exts)
                sFileTry = sAppName & Exts(i)

                ProcPath = String$(MAX_PATH, 0&)
                LSet ProcPath = sFileTry & vbNullChar

                If CBool(PathFindOnPath(StrPtr(ProcPath), 0&)) Then
                    FindOnPath = TrimNull(ProcPath)
                    Exit For
                End If
                
            Next
            
        End If

        ToggleWow64FSRedirection True
    End If
    
    AppendErrorLogCustom "FindOnPath - App Paths"
    
    If Len(FindOnPath) = 0 And Not bFullPath Then
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" & sAppName, vbNullString)
        If 0 <> Len(sFile) Then
            If FileExists(sFile) Then
                FindOnPath = sFile
            End If
        End If
    End If
    
    If Len(FindOnPath) = 0 Then
        sFile = vbNullString
        If StrBeginWith(sAppName, "\systemroot") Then
            sFile = Replace$(sAppName, "\systemroot", sWinDir, , 1, vbTextCompare)
        ElseIf StrBeginWith(sAppName, "system32\") Then
            sFile = sWinDir & "\" & sAppName
        ElseIf StrBeginWith(sAppName, "SysWOW64\") Then
            sFile = sWinDir & "\" & sAppName
        End If
        
        If FileExists(sFile) Then
            FindOnPath = sFile
        End If
    End If

    If Len(FindOnPath) = 0 And bUseSourceValueOnFailure Then
        FindOnPath = sAppName
    End If
    
    AppendErrorLogCustom "FindOnPath - End"

    Exit Function
ErrorHandler:
    ErrorMsg Err, "FindOnPath", "AppName: ", sAppName
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Function

Public Sub SplitIntoPathAndArgs(ByVal InLine As String, Path As String, Optional Args As String, Optional bIsRegistryData As Boolean)
    On Error GoTo ErrorHandler
    Dim pos As Long
    Dim sTmp As String
    Dim bFail As Boolean
    
    Path = vbNullString
    Args = vbNullString
    If Len(InLine) = 0& Then Exit Sub
    
    InLine = Trim$(InLine)
    If Left$(InLine, 1) = """" Then
        pos = InStr(2, InLine, """")
        If pos <> 0 Then
            Path = Mid$(InLine, 2, pos - 2)
            Args = Trim$(Mid$(InLine, pos + 1))
        Else
            Path = Mid$(InLine, 2)
        End If
    Else
        '//TODO: Check correct system behaviour: maybe it uses number of 'space' characters, like, if more than 1 'space', exec bIsRegistryData routine.
    
        If bIsRegistryData Then
            If FileExists(InLine) Then
                Path = InLine
                Exit Sub
            End If
            'Expanding paths like: C:\Program Files (x86)\Download Master\dmaster.exe -autorun
            pos = InStrRev(InLine, ".exe", -1, 1)
            If pos <> 0 Then
                Path = Left$(InLine, pos + 3)
                Args = LTrim(Mid$(InLine, pos + 4))
                If Not FileExists(Path) Then bFail = True
            End If
        Else
            bFail = True
        End If
        
        If bFail Or Len(Path) = 0 Then
            pos = InStr(InLine, " ")
            If pos <> 0 Then
                Path = Left$(InLine, pos - 1)
                Args = Mid$(InLine, pos + 1)
            Else
                Path = InLine
            End If
        End If
    End If
    If Len(Path) <> 0 Then
        If Not FileExists(Path) Then  'find on %PATH%
            sTmp = FindOnPath(Path)
            If Len(sTmp) <> 0 Then
                Path = sTmp
            End If
        End If
    End If
    
    'Anti-HJT-hijack :)
    If InStr(1, Args, "(Microsoft)", 1) <> 0 Then
        If Not IsMicrosoftFile(Path) Then
            Args = Args & " <== not a Microsoft !!!"
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Parser.SplitIntoPathAndArgs", "In Line:", InLine
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetParentDir(sPath As String) As String
    Dim pos As Long
    pos = InStrRev(sPath, "\")
    If pos <> 0 Then
        GetParentDir = Left$(sPath, pos - 1)
    End If
End Function

Public Function CopyFolder(ByVal sFolder$, ByVal sTo$) As Boolean
    On Error GoTo ErrorHandler:
    Dim uFOS As SHFILEOPSTRUCT
    sFolder = sFolder & Chr$(0)
    sTo = sTo & Chr$(0)
    With uFOS
        .wFunc = FO_COPY
        .pFrom = StrPtr(sFolder)
        .pTo = StrPtr(sTo)
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI Or FOF_NOCONFIRMMKDIR
    End With
    CopyFolder = (0 = SHFileOperation(uFOS))
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CopyFolder", "Folder:", sFolder
    If inIDE Then Stop: Resume Next
End Function

Public Function DeleteFolderForce(sFolder As String, Optional bForceDeleteMicrosoft As Boolean) As Boolean
    On Error GoTo ErrorHandler:

    Dim aFiles() As String
    Dim i As Long
    Dim iAttr As Long
    Dim bRedirect As Boolean
    Dim bOldStatus As Boolean
    
    If Not FolderExists(sFolder) Then
        DeleteFolderForce = True
    Else
        DeleteFolderForce = True
        bRedirect = ToggleWow64FSRedirection(False, sFolder, bOldStatus)
        iAttr = GetFileAttributes(StrPtr(sFolder))
        If (iAttr And 2048) Then iAttr = iAttr - 2048
        If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes StrPtr(sFolder), iAttr And Not FILE_ATTRIBUTE_READONLY
        If Not DeleteFolder(sFolder) Then
            TryUnlock sFolder
            If Not DeleteFolder(sFolder) Then
                If Not MoveFolder(sFolder, sFolder & ".bak") Then
                    DeleteFolderForce = False
                    aFiles = ListFiles(sFolder)
                    If AryItems(aFiles) Then
                        For i = 0 To UBound(aFiles)
                            DeleteFileWEx StrPtr(aFiles(i)), bForceDeleteMicrosoft
                        Next
                    End If
                End If
            End If
        End If
        If bRedirect Then Call ToggleWow64FSRedirection(bOldStatus)
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DeleteFolderForce", "Folder:", sFolder
    If bRedirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function DeleteFolder(ByVal sFolder$) As Boolean
    On Error GoTo ErrorHandler
    Dim uFOS As SHFILEOPSTRUCT
    sFolder = sFolder & Chr$(0)
    With uFOS
        .wFunc = FO_DELETE
        .pFrom = StrPtr(sFolder)
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI Or FOF_NOCONFIRMMKDIR
    End With
    DeleteFolder = (0 = SHFileOperation(uFOS))
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DeleteFolder", "Folder:", sFolder
    If inIDE Then Stop: Resume Next
End Function

Public Function MoveFolder(ByVal sFolder$, ByVal sTo$) As Boolean
    On Error GoTo ErrorHandler
    Dim uFOS As SHFILEOPSTRUCT
    sFolder = sFolder & Chr$(0)
    sTo = sTo & Chr$(0)
    With uFOS
        .wFunc = FO_MOVE
        .pFrom = StrPtr(sFolder)
        .pTo = StrPtr(sTo)
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI Or FOF_NOCONFIRMMKDIR
    End With
    MoveFolder = (0 = SHFileOperation(uFOS))
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MoveFolder", "Source:", sFolder, "Destination:" & sTo
    If inIDE Then Stop: Resume Next
End Function

'if short name is unavailable, it returns source string anyway
Public Function GetDOSFilename(sFile$, Optional bReverse As Boolean = False) As String
    On Error GoTo ErrorHandler
    'works for folders too btw
    Dim cnt&, sBuffer$
    If bReverse Then
        sBuffer = Space$(MAX_PATH_W)
        cnt = GetLongPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
        If cnt Then
            GetDOSFilename = Left$(sBuffer, cnt)
        Else
            GetDOSFilename = sFile
        End If
    Else
        sBuffer = Space$(MAX_PATH)
        cnt = GetShortPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
        If cnt Then
            GetDOSFilename = Left$(sBuffer, cnt)
        Else
            GetDOSFilename = sFile
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetDOSFilename", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function

Public Function GetLongPath(sFile As String) As String '8.3 -> to Full name
    On Error GoTo ErrorHandler
    If InStr(sFile, "~") = 0 Then
        GetLongPath = sFile
        Exit Function
    End If
    Dim sBuffer As String, cnt As Long, pos As Long, sFolder As String
    
    If Not FileExists(sFile) And Not FolderExists(sFile) Then
        'try to convert folder struct instead, like C:\PROGRA~1\MICROS~1\Office15\ONBttnIE.dll (file missing)
        pos = InStrRev(sFile, "\", -1)
        If pos <> 0 Then
            Do
                sFolder = Left$(sFile, pos - 1)
                
                If InStr(sFolder, "~") = 0 Then Exit Do
                
                If FolderExists(sFolder) Then
                    GetLongPath = GetLongPath(sFolder) & "\" & Mid$(sFile, pos + 1)
                    Exit Do
                End If
                
                pos = pos - 1
                If pos <> 0 Then
                    pos = InStrRev(sFile, "\", pos)
                End If
                
            Loop While pos <> 0
        End If
        If GetLongPath = "" Then GetLongPath = sFile
        Exit Function
    End If
    
    sBuffer = String$(MAX_PATH_W, 0&)
    cnt = GetLongPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
    If cnt Then
        GetLongPath = Left$(sBuffer, cnt)
    Else
        GetLongPath = sFile
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetLongPath", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function

Public Function ShowFileProperties(sFile$, Handle As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_DOENVSUBST Or SEE_MASK_FLAG_NO_UI
        .hwnd = Handle
        .lpFile = StrPtr(PathX64(sFile))
        .lpVerb = StrPtr("properties")
        .nShow = 1
    End With
    ShowFileProperties = (ShellExecuteEx(uSEI) <> 0)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ShowFileProperties", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function

Public Sub DeleteFileOnReboot(sFile$, Optional bDeleteBlindly As Boolean = False, Optional bNoReboot As Boolean = False)
    On Error GoTo ErrorHandler:

    'If Not bIsWinNT Then Exit Sub
    If Not FileExists(sFile) And Not bDeleteBlindly Then Exit Sub
    If bIsWinNT Then
        MoveFileEx StrPtr(sFile), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    Else
        Dim ff%
        On Error Resume Next
        ff = FreeFile()
        Open sWinDir & "\wininit.ini" For Append As #ff
        If Err.Number = 5 Then
            Err.Clear
            TryUnlock sWinDir & "\wininit.ini"
            Open sWinDir & "\wininit.ini" For Append As #ff
            If Err.Number <> 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Print #ff, "[rename]"
        Print #ff, "NUL=" & GetDOSFilename(sFile)
        Print #ff,
        Close #ff
    End If
    
    If Not bNoReboot Then
        'RestartSystem "The file '" & sFile & "' will be deleted by Windows when the system restarts."
        RestartSystem Replace$(Translate(342), "[]", sFile)
    End If
    
    '// TODO:
    'Windows Server 2003 Note:
    'https://support.microsoft.com/en-us/kb/948601
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "DeleteFileOnReboot", "File:", sFile
    If inIDE Then Stop: Resume Next
End Sub

'Support Win 2000+
Public Function GetSpecialFolderPath(CSIDL As Long, Optional hToken As Long = 0&) As String
    On Error GoTo ErrorHandler:

    Const SHGFP_TYPE_CURRENT As Long = &H0&
    Const SHGFP_TYPE_DEFAULT As Long = &H1&
    Const CSIDL_FLAG_DONT_UNEXPAND As Long = &H2000&
    Const CSIDL_FLAG_DONT_VERIFY As Long = &H4000&
    Dim lr      As Long
    Dim sPath   As String
    sPath = String$(MAX_PATH, 0&)
    ' 3-th parameter - is a token of user
    lr = SHGetFolderPath(0&, CSIDL Or CSIDL_FLAG_DONT_UNEXPAND Or CSIDL_FLAG_DONT_VERIFY, hToken, SHGFP_TYPE_CURRENT, StrPtr(sPath))
    If lr = 0 Then GetSpecialFolderPath = Left$(sPath, lstrlen(StrPtr(sPath)))
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetSpecialFolderPath", "CSIDL:", CSIDL
    If inIDE Then Stop: Resume Next
End Function

'Support Win Vista+
Public Function GetKnownFolderPath(ByVal KnownFolderID As String) As String
    On Error GoTo ErrorHandler
    'https://msdn.microsoft.com/en-us/library/windows/desktop/dd378457(v=vs.85).aspx
    
    Const KF_FLAG_NOT_PARENT_RELATIVE   As Long = &H200&
    Const KF_FLAG_DEFAULT_PATH          As Long = &H400&
    Const KF_FLAG_CREATE                As Long = &H8000&
    Const KF_FLAG_DONT_VERIFY           As Long = &H4000&
    
    If OSver.MajorMinor < 6 Then Exit Function
    
    Dim rfid    As UUID
    Dim lr      As Long
    Dim sPath   As String
    Dim ptr     As Long
    Dim strLen  As Long
    
    CLSIDFromString StrPtr(KnownFolderID), rfid
    
    lr = SHGetKnownFolderPath(rfid, KF_FLAG_NOT_PARENT_RELATIVE Or KF_FLAG_DEFAULT_PATH Or KF_FLAG_DONT_VERIFY, 0&, ptr)
    
    If 0 = lr Then
        strLen = lstrlen(ptr)
        sPath = String$(strLen, vbNullChar)
        lstrcpyn StrPtr(sPath), ptr, strLen + 1&
    
        GetKnownFolderPath = sPath
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.GetKnownFolderPath", "KnownFolderID:", KnownFolderID
    If inIDE Then Stop: Resume Next
End Function

'Support Win Vista+
Public Function GetKnownFolderPath_GUID(iid As UUID) As String
    On Error GoTo ErrorHandler
    'https://msdn.microsoft.com/en-us/library/windows/desktop/dd378457(v=vs.85).aspx
    
    Const KF_FLAG_NOT_PARENT_RELATIVE   As Long = &H200&
    Const KF_FLAG_DEFAULT_PATH          As Long = &H400&
    Const KF_FLAG_CREATE                As Long = &H8000&
    Const KF_FLAG_DONT_VERIFY           As Long = &H4000&
    
    If OSver.MajorMinor < 6 Then Exit Function
    
    Dim lr      As Long
    Dim sPath   As String
    Dim ptr     As Long
    Dim strLen  As Long
    
    lr = SHGetKnownFolderPath(iid, KF_FLAG_NOT_PARENT_RELATIVE Or KF_FLAG_DEFAULT_PATH Or KF_FLAG_DONT_VERIFY, 0&, ptr)
    
    If 0 = lr Then
        strLen = lstrlen(ptr)
        sPath = String$(strLen, vbNullChar)
        lstrcpyn StrPtr(sPath), ptr, strLen + 1&
        
        GetKnownFolderPath_GUID = sPath
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.GetKnownFolderPath_GUID"
    If inIDE Then Stop: Resume Next
End Function

Public Function MkDirW(ByVal Path As String, Optional ByVal LastComponentIsFile As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    ' Create folders struct
    ' LastComponentIsFile - true, if you specify filename as a last part of path component
    ' Return value: true, if successfully created or if folder is already exists
    Dim FC As String, lr As Boolean, pos As Long
    Dim bRedirect As Boolean, bOldStatus As Boolean
    If LastComponentIsFile Then Path = Left(Path, InStrRev(Path, "\") - 1) ' cut off file name
    If InStr(Path, ":") = 0 And Not StrBeginWith(Path, "\\") Then 'if relative path
        Dim sCurDir$, nChar As Long
        sCurDir = String$(MAX_PATH, 0&)
        nChar = GetCurrentDirectory(MAX_PATH + 1, StrPtr(sCurDir))
        sCurDir = Left$(sCurDir, nChar)
        If Right$(sCurDir, 1) <> "\" Then sCurDir = sCurDir & "\"
        Path = sCurDir & Path
    End If
    If FolderExists(Path, , True) Then
        MkDirW = True
        Exit Function
    End If
    bRedirect = ToggleWow64FSRedirection(False, Path, bOldStatus)
    If StrBeginWith(Path, "\\") Then
        pos = InStr(3, Path, "\")
        If pos = 0 Then Exit Function
    End If
    Do 'looping through each path component
        pos = pos + 1
        pos = InStr(pos, Path, "\")
        If pos Then FC = Left(Path, pos - 1) Else FC = Path
        If FolderExists(FC, , True) Then
            lr = True 'if folder is already created
        Else
            lr = CBool(CreateDirectory(StrPtr(FC), ByVal 0&))
            If lr = 0 Then Exit Do
        End If
    Loop While (pos <> 0) And (lr <> 0)
    MkDirW = lr
    If bRedirect Then ToggleWow64FSRedirection bOldStatus
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.MkDirW", "Path:", Path, "IsFile:", LastComponentIsFile
    If inIDE Then Stop: Resume Next
End Function

'// replace path with environment variables if possible
Public Function EnvironUnexpand(ByVal p_Path As String, Optional bFirstOccurenceOnly As Boolean = True) As String
    On Error GoTo ErrorHandler
    
    Static bInit As Boolean
    Static aSrc(9) As String
    Static aDst(9) As String
    
    If Not bInit Then
        bInit = True
        
        aSrc(0) = sWinDir: aDst(0) = "%SystemRoot%"
        aSrc(1) = PF_32: aDst(1) = "%ProgramFiles(x86)%"
        aSrc(2) = PF_64: aDst(2) = "%ProgramFiles%"
        aSrc(3) = ProgramData: aDst(3) = "%ProgramData%"
        aSrc(4) = LocalAppData: aDst(4) = "%LOCALAPPDATA%"
        aSrc(5) = AppData: aDst(5) = "%APPDATA%"
        aSrc(6) = TempCU: aDst(6) = "%TEMP%"
        aSrc(7) = AllUsersProfile: aDst(7) = "%PUBLIC%"
        aSrc(8) = UserProfile: aDst(8) = "%USERPROFILE%"
        aSrc(9) = SysDisk: aDst(9) = "%SystemDrive%"
    End If
    
    Dim i As Long
    For i = 0 To UBound(aSrc)
        If ReplaceEV(p_Path, aSrc(i), aDst(i)) Then
            If bFirstOccurenceOnly Then
                EnvironUnexpand = p_Path
                Exit Function
            End If
        End If
    Next
    
    EnvironUnexpand = p_Path
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.EnvironUnexpand", "Path:", p_Path
    If inIDE Then Stop: Resume Next
End Function

Function ReplaceEV(p_Path As String, p_What As String, p_Into As String) As Boolean
  If StrBeginWith(p_Path, p_What) Then p_Path = p_Into & Mid(p_Path, Len(p_What) + 1): ReplaceEV = True
End Function

Public Function GetFreeDiscSpace(sRoot As String, bForCurrentUser As Boolean) As Currency ' result = Int64
    On Error GoTo ErrorHandler
    If IsProcedureAvail("GetDiskFreeSpaceExW", "kernel32.dll") Then
        If bForCurrentUser Then
            If GetDiskFreeSpaceEx(StrPtr(sRoot), VarPtr(GetFreeDiscSpace), 0&, 0&) = 0 Then Dbg "GetDiskFreeSpaceEx is failed with: " & Err.LastDllError
        Else
            If GetDiskFreeSpaceEx(StrPtr(sRoot), 0&, 0&, VarPtr(GetFreeDiscSpace)) = 0 Then Dbg "GetDiskFreeSpaceEx is failed with: " & Err.LastDllError
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.GetFreeDiscSpace", "Root:", sRoot
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFileSymlinkTarget(sPath As String) As String
    On Error GoTo ErrorHandler
    'also see:
    'https://msdn.microsoft.com/en-us/library/windows/desktop/aa366789(v=vs.85).aspx
    'https://blez.wordpress.com/2012/09/17/enumerating-opened-handles-from-a-process/
    'https://stackoverflow.com/questions/65170/how-to-get-name-associated-with-open-handle/
    
    Dim hFile           As Long
    Dim returnedLength  As Long
    Dim Status          As Long
    Dim DeviceObjName   As String
    Dim objName(1000)   As Integer
    
    hFile = CreateFile(StrPtr(sPath), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, _
        OPEN_EXISTING, g_FileBackupFlag, 0&)
    
    If hFile <> INVALID_HANDLE_VALUE Then
        
        'OBJECT_NAME_INFORMATION {
        '    UNICODE_STRING          Name;
        '    WCHAR                   NameBuffer[0];
        '}
        
        Status = NtQueryObject(hFile, ObjectNameInformation, objName(0), UBound(objName) * 2, returnedLength)
        
        If NT_SUCCESS(Status) And objName(0) > 0 Then
            DeviceObjName = StringFromPtrW(VarPtr(objName(4)))
            GetFileSymlinkTarget = GetDOSFilename(ConvertDosDeviceToDriveName(DeviceObjName), True)
        End If
        
        CloseHandle hFile
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.GetFileSymlinkTarget", "Path:", sPath
    If inIDE Then Stop: Resume Next
End Function

Public Function NT_SUCCESS(NT_Code As Long) As Boolean
    NT_SUCCESS = (NT_Code >= 0)
End Function

Public Function GetEmptyDriveNames() As String()
    Dim buf As String
    Dim BufLen As Long
    Dim Letters As String
    Dim Drives() As String
    Dim EmptyDrive() As String
    Dim i As Long

    Letters = StrReverse("ABCDEFGHIJKLMNOPQRSTUVWXYZ")

    buf = String$(MAX_PATH, 0)

    BufLen = GetLogicalDriveStrings(MAX_PATH, StrPtr(buf))
   
    If BufLen <> 0 Then
        buf = Left$(buf, BufLen - 1)
        
        Drives = Split(buf, vbNullChar)
       
        ReDim ReadyDrives(UBound(Drives) + 1)
       
        For i = 0 To UBound(Drives)
            Letters = Replace(Letters, Left$(Drives(i), 1), "", , , vbTextCompare)
        Next
        
        If Len(Letters) <> 0 Then
            ReDim EmptyDrive(Len(Letters) - 1) As String
        
            For i = 0 To Len(Letters) - 1
                EmptyDrive(i) = Mid$(Letters, i + 1, 1) & ":"
            Next
            GetEmptyDriveNames = EmptyDrive
        End If
    End If
End Function

Public Function GetVolumeFlags(ByVal sVolume) As VOLUME_INFO_FLAGS
    Dim Flags As Long
    Dim lMaxCompLength As Long
    Dim sFS As String
    Dim sVolName As String
    Dim lVolSN As Long
    
    If Len(sVolume) > 3 Then sVolume = Left$(sVolume, 3)
    If Right$(sVolume, 1) <> "\" Then sVolume = sVolume & "\"
    
    sFS = String(MAX_PATH, 0)
    sVolName = String(MAX_PATH, 0)
    
    If GetVolumeInformation(sVolume, sVolName, Len(sVolName), lVolSN, lMaxCompLength, Flags, sFS, Len(sFS)) Then
        GetVolumeFlags = Flags
    End If
End Function

Public Function FileHasBOM_UTF16(sFile As String) As Boolean
    Dim Redirect As Boolean, bOldStatus As Boolean, hFile As Long, lSize As Long
    If Not FileExists(sFile) Then Exit Function
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    OpenW sFile, FOR_READ, hFile, g_FileBackupFlag
    If hFile <= 0 Then Exit Function
    lSize = LOFW(hFile)
    If lSize < 2 Then CloseW hFile: Exit Function
    ReDim b(1) As Byte
    GetW hFile, 1, , VarPtr(b(0)), UBound(b) + 1
    CloseW hFile
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If b(0) = &HFF& And b(1) = &HFE& Then FileHasBOM_UTF16 = True
End Function

Public Function FileHasBOM_UTF8(sFile As String) As Boolean
    Dim Redirect As Boolean, bOldStatus As Boolean, hFile As Long, lSize As Long
    If Not FileExists(sFile) Then Exit Function
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    OpenW sFile, FOR_READ, hFile, g_FileBackupFlag
    If hFile <= 0 Then Exit Function
    lSize = LOFW(hFile)
    If lSize < 3 Then CloseW hFile: Exit Function
    ReDim b(2) As Byte
    GetW hFile, 1, , VarPtr(b(0)), UBound(b) + 1
    CloseW hFile
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If b(0) = &HEF& And b(1) = &HBB& And b(2) = &HBF& Then FileHasBOM_UTF8 = True
End Function

Public Function FileGetTypeBOM(sFile As String) As Long
    Dim Redirect As Boolean, bOldStatus As Boolean, hFile As Long, lSize As Long
    If Not FileExists(sFile) Then Exit Function
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    OpenW sFile, FOR_READ, hFile, g_FileBackupFlag
    If hFile <= 0 Then Exit Function
    lSize = LOFW(hFile)
    If lSize < 2 Then CloseW hFile: Exit Function
    ReDim b(IIf(lSize < 3, 2, 3)) As Byte
    GetW hFile, 1, , VarPtr(b(0)), UBound(b) + 1
    CloseW hFile
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If b(0) = &HFF& And b(1) = &HFE& Then
        FileGetTypeBOM = 1200
        Exit Function
    End If
    If UBound(b) > 1 Then
        If b(0) = &HEF& And b(1) = &HBB& And b(2) = &HBF& Then
            FileGetTypeBOM = 65001
        End If
    End If
End Function
