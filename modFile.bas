Attribute VB_Name = "modFile"
'
' modFile module by Alex Dragokas
'

Option Explicit

Const MAX_PATH As Long = 260&
Const MAX_FILE_SIZE As Currency = 104857600@

Enum VB_FILE_ACCESS_MODE
    FOR_READ = 1
    FOR_READ_WRITE = 2
    FOR_OVERWRITE_CREATE = 4
End Enum

Enum CACHE_TYPE
    USE_CACHE
    NO_CACHE
End Enum
 
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
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
    lpszFileName(MAX_PATH) As Integer
    lpszAlternate(14) As Integer
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function PathFileExists Lib "Shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function SHFileExists Lib "shell32.dll" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Long
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeW" (ByVal nDrive As Long) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long

Const FILE_SHARE_READ           As Long = &H1&
Const FILE_SHARE_WRITE          As Long = &H2&
Const FILE_SHARE_DELETE         As Long = 4&
Const FILE_READ_ATTRIBUTES      As Long = &H80&
Const OPEN_EXISTING             As Long = 3&
Const CREATE_ALWAYS             As Long = 2&
Const GENERIC_READ              As Long = &H80000000
Const GENERIC_WRITE             As Long = &H40000000
Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&
Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF
Const ERROR_SUCCESS             As Long = 0&
Const INVALID_FILE_ATTRIBUTES   As Long = -1&
Const NO_ERROR                  As Long = 0&
Const FILE_BEGIN                As Long = 0&
Const FILE_CURRENT              As Long = 1&
Const FILE_END                  As Long = 2&
Const INVALID_SET_FILE_POINTER  As Long = &HFFFFFFFF

Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Const KEY_QUERY_VALUE           As Long = &H1&
Const RegType_DWord             As Long = 4&

Private lWow64Old               As Long
Private DriveTypeName           As New Collection



Public Function FileExists(ByVal sFile$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim Redirect As Boolean
    
    sFile = Trim$(sFile)
    If Len(sFile) = 0 Then Exit Function
    If Left$(sFile, 2) = "\\" Then Exit Function 'DriveType = "REMOTE"
    
    ' use 2 methods for reliability reason (both supported unicode pathes)
    Dim Ex(1) As Boolean
    Dim ret As Long
    
    Dim WFD     As WIN32_FIND_DATA
    Dim hFile   As Long
    
    If Not bUseWow64 Then Redirect = ToggleWow64FSRedirection(False, sFile)
    
    ret = GetFileAttributes(StrPtr(sFile))
    If ret <> INVALID_HANDLE_VALUE And (0 = (ret And FILE_ATTRIBUTE_DIRECTORY)) Then Ex(0) = True
 
    hFile = FindFirstFile(StrPtr(sFile), WFD)
    Ex(1) = (hFile <> INVALID_HANDLE_VALUE) And Not CBool(WFD.dwFileAttributes And vbDirectory)
    FindClose hFile

    ' // here must be enabling of FS redirector
    If Redirect Then Call ToggleWow64FSRedirection(True)

    FileExists = Ex(0) Or Ex(1)
    Exit Function
ErrorHandler:
    ErrorMsg err, "modFile.FileExists", "File:", sFile$
    If inIDE Then Stop: Resume Next
End Function

Public Function FolderExists(ByVal sFolder$, Optional ForceUnderRedirection As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim ret As Long
    sFolder = Trim$(sFolder)
    If Len(sFolder) = 0 Then Exit Function
    If Left$(sFolder, 2) = "\\" Then Exit Function 'network path
    
    '// FS redirection checking
    
    ret = GetFileAttributes(StrPtr(sFolder))
    FolderExists = CBool(ret And vbDirectory) And (ret <> INVALID_FILE_ATTRIBUTES)
    
    '// FS redirection enambling
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "modFile.FolderExists", "Folder:", sFolder$, "Redirection: ", ForceUnderRedirection
    If inIDE Then Stop: Resume Next
End Function


Public Sub GetDriveTypeNames()
    On Error GoTo ErrorHandler
    Dim lr As Long
    Dim i  As Long
    Dim DT As String
    
    For i = 65& To 90&
    
      lr = GetDriveType(StrPtr(Chr$(i) & ":\"))
    
      Select Case lr
        Case 3&
            DT = "FIXED"
        Case 2&
            DT = "REMOVABLE"
        Case 5&
            DT = "CDROM"
        Case 4&
            DT = "REMOTE"
        Case 0&
            DT = "UNKNOWN"
        Case 1&
            DT = "DISCONNECTED" '"NO_ROOT_DIR"
        Case 6&
            DT = "RAMDISK"
        Case Else
            DT = "UNKNOWN"
      End Select
      
      DriveTypeName.Add DT, Chr$(i)
      
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "modFile.GetDriveTypeNames", "Drive:", Chr$(i)
End Sub


Function FileLenW(Path As String) As Currency ', Optional DoNotUseCache As Boolean
    On Error GoTo ErrorHandler
'    ' Last cached File
'    Static CachedFile As String
'    Static CachedSize As Currency
    
    Dim lr          As Long
    Dim hFile       As Long
    Dim FileSize    As Currency

'    If Not DoNotUseCache Then
'        If StrComp(Path, CachedFile, 1) = 0 Then
'            FileLenW = CachedSize
'            Exit Function
'        End If
'    End If

    hFile = CreateFile(StrPtr(Path), FILE_READ_ATTRIBUTES, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then FileLenW = FileSize * 10000&
        End If
'        If Not DoNotUseCache Then
'            CachedFile = Path
'            CachedSize = FileLenW
'        End If
        CloseHandle hFile: hFile = 0&
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "modFile.FileLenW", "File:", Path, "hFile:", hFile, "FileSize:", FileSize, "Return:", lr
End Function



Public Function OpenW(FileName As String, Access As VB_FILE_ACCESS_MODE, retHandle As Long, Optional MountToMemory As Boolean) As Boolean '// TODO: MountToMemory
    
    Dim FSize As Currency
    
    'Print #ffOpened, FileName
    
    If Access And (FOR_READ Or FOR_READ_WRITE) Then
        If Not FileExists(FileName) Then
            retHandle = INVALID_HANDLE_VALUE
            Exit Function
        End If
    End If
        
    If Access = FOR_READ Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_OVERWRITE_CREATE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_READ_WRITE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        'WriteCon "Wrong access mode!", cErr
    End If

    OpenW = (INVALID_HANDLE_VALUE <> retHandle)
    
    ' ограничение на максимально возможный файл для открытия ( > 100 МБ )
    If OpenW Then
        If Access And (FOR_READ Or FOR_READ_WRITE) Then
            FSize = LOFW(retHandle)
            If FSize > MAX_FILE_SIZE Then
                CloseHandle retHandle
                retHandle = INVALID_HANDLE_VALUE
                OpenW = False
                '"Не хочу и не буду открывать этот файл, потому что его размер превышает безопасный максимум"
                err.Clear: ErrorMsg err, "modFile.OpenW: " & "Trying to open too big file" & ": (" & (FSize \ 1024 \ 1024) & " MB.) " & FileName
            End If
        End If
    Else
        err.Clear: ErrorMsg err, "modFile.OpenW: Cannot open file: " & FileName
        err.Raise 75 ' Path/File Access error
    End If

End Function

                                                                  'do not change Variant type at all or you will die ^_^
Public Function GetW(hFile As Long, pos As Long, Optional vOut As Variant, Optional vOutPtr As Long, Optional cbToRead As Long) As Boolean
                                                                  
    'On Error GoTo ErrorHandler
    
    Dim lBytesRead  As Long
    Dim lr          As Long
    Dim ptr         As Long
    Dim vType       As Long
    Dim UnknType    As Boolean
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If INVALID_SET_FILE_POINTER <> SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then
        If NO_ERROR = err.LastDllError Then
            vType = VarType(vOut)
            
            If 0 <> cbToRead Then   'vbError = vType
                lr = ReadFile(hFile, vOutPtr, cbToRead, lBytesRead, 0&)
                
            ElseIf vbString = vType Then
                lr = ReadFile(hFile, StrPtr(vOut), Len(vOut), lBytesRead, 0&)
                If err.LastDllError <> 0 Or lr = 0 Then err.Raise 52
                
                vOut = StrConv(vOut, vbUnicode)
                If Len(vOut) <> 0 Then vOut = Left$(vOut, Len(vOut) \ 2)
            Else
                'do a bit of magik :)
                memcpy ptr, ByVal VarPtr(vOut) + 8, 4& 'VT_BYREF
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
                    err.Clear: ErrorMsg err, "modFile.GetW. type #" & VarType(vOut) & " of buffer is not supported.": err.Raise 52
                End Select
            End If
            GetW = (0 <> lr)
            If 0 = lr And Not UnknType Then err.Clear: ErrorMsg err, "Cannot read file!": err.Raise 52
        Else
            err.Clear: ErrorMsg err, "Cannot set file pointer!": err.Raise 52
        End If
    Else
        err.Clear: ErrorMsg err, "Cannot set file pointer!": err.Raise 52
    End If
    
'    Exit Function
'ErrorHandler:
'    AppendErrorLogFormat Now, err, "modFile.GetW"
'    Resume Next
End Function

Public Function PutW(hFile As Long, pos As Long, vInPtr As Long, cbToWrite As Long, Optional doAppend As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lBytesWrote  As Long
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If doAppend Then
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) Then Exit Function
    Else
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then Exit Function
    End If
    
    If NO_ERROR = err.LastDllError Then
    
        If WriteFile(hFile, vInPtr, cbToWrite, lBytesWrote, 0&) Then PutW = True
        
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "modFile.PutW"
End Function

Public Function LOFW(hFile As Long) As Currency
    On Error GoTo ErrorHandler
    Dim lr          As Long
    Dim FileSize    As Currency
    
    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then
                LOFW = FileSize * 10000&
            Else
                err.Clear
                ErrorMsg Now, "File is too big. Size: " & FileSize
            End If
        End If
    End If
ErrorHandler:
End Function

Public Function CloseW(hFile As Long) As Long
    CloseW = CloseHandle(hFile)
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
        If StrComp(Left$(PathNecessity, Len(sWinDir)), sWinDir, vbTextCompare) <> 0 Then Exit Function
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
End Function


Public Function GetExtensionName(Path As String) As String  'вернет .EXT
    Dim pos As Long
    pos = InStrRev(Path, ".")
    If pos <> 0 Then GetExtensionName = Mid$(Path, pos)
End Function

' Является ли файл форматом PE EXE
Public Function isPE_EXE(Optional FileName As String, Optional FileHandle As Long) As Boolean
    On Error GoTo ErrorHandler

'    #If UseHashTable Then
'        Static PE_EXE_Cache As clsTrickHashTable
'    #Else
'        Static PE_EXE_Cache As Object
'    #End If
'
'    If 0 = ObjPtr(PE_EXE_Cache) Then
'        #If UseHashTable Then
'            Set PE_EXE_Cache = New clsTrickHashTable
'        #Else
'            Set PE_EXE_Cache = CreateObject("Scripting.Dictionary")
'        #End If
'        PE_EXE_Cache.CompareMode = vbTextCompare
'    Else
'        If Len(FileName) <> 0& Then
'            If PE_EXE_Cache.Exists(FileName) Then
'                isPE_EXE = PE_EXE_Cache(FileName)
'                Exit Function
'            End If
'        End If
'    End If

    'Static PE_EXE_Cache    As New Collection ' value = true, если файл является форматом PE EXE

    'If Len(FileName) <> 0 Then
    '    If isCollectionKeyExists(FileName, PE_EXE_Cache) Then
    '        isPE_EXE = PE_EXE_Cache(FileName)
    '        Exit Function
    '    End If
    'End If

    Dim hFile          As Long
    Dim PE_offset      As Long
    Dim MZ(1)          As Byte
    Dim pe(3)          As Byte
    Dim FSize          As Currency
  
    If FileHandle = 0& Then
        OpenW FileName, FOR_READ, hFile
    Else
        hFile = FileHandle
    End If
    If hFile <> INVALID_HANDLE_VALUE Then
        FSize = LOFW(hFile)
        If FSize >= &H3C& + 4& Then
            GetW hFile, 1&, , VarPtr(MZ(0)), ((UBound(MZ) + 1&) * CLng(LenB(MZ(0))))
            If (MZ(0) = 77& And MZ(1) = 90&) Or (MZ(1) = 77& And MZ(0) = 90&) Then  'MZ or ZM
                GetW hFile, &H3C& + 1&, PE_offset
                If PE_offset And FSize >= PE_offset + 4 Then
                    GetW hFile, PE_offset + 1&, , VarPtr(pe(0)), ((UBound(pe) + 1&) * CLng(LenB(pe(0))))
                    If pe(0) = 80& And pe(1) = 69& And pe(2) = 0& And pe(3) = 0& Then isPE_EXE = True   'PE NUL NUL
                End If
            End If
        End If
        If FileHandle = 0& Then CloseW hFile: hFile = 0&
    End If
    
    'If Len(FileName) <> 0& Then PE_EXE_Cache.Add FileName, isPE_EXE
    Exit Function
    
ErrorHandler:
    ErrorMsg err, "Parser.isPE_EXE", "File:", FileName
    On Error Resume Next
    'If Len(FileName) <> 0& Then PE_EXE_Cache.Add FileName, isPE_EXE
    If FileHandle = 0& Then
        If hFile <> 0 Then CloseW hFile: hFile = 0&
    End If
End Function
