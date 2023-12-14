Attribute VB_Name = "modPeParser"
' [modPeParser.bas]
'
' Pe Parser by The Trick & Dragokas
'
' Thanks to fafalone for improvements
'
Option Explicit

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private Type IMAGE_DOS_HEADER
    e_magic_e_cblp                  As Long
    e_cp                            As Integer
    e_crlc                          As Integer
    e_cparhdr                       As Integer
    e_minalloc                      As Integer
    e_maxalloc                      As Integer
    e_ss                            As Integer
    e_sp                            As Integer
    e_csum                          As Integer
    e_ip                            As Integer
    e_cs                            As Integer
    e_lfarlc                        As Integer
    e_ovno                          As Integer
    e_res(0 To 3)                   As Integer
    e_oemid                         As Integer
    e_oeminfo                       As Integer
    e_res2(0 To 9)                  As Integer
    e_lfanew                        As Long
End Type
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress                  As Long
    Size                            As Long
End Type
Private Type IMAGE_OPTIONAL_HEADER
    Magic                           As Integer
    MajorLinkerVersion              As Byte
    MinorLinkerVersion              As Byte
    SizeOfCode                      As Long
    SizeOfInitializedData           As Long
    SizeOfUnitializedData           As Long
    AddressOfEntryPoint             As Long
    BaseOfCode                      As Long
    BaseOfData                      As Long
    ImageBase                       As Long
    SectionAlignment                As Long
    FileAlignment                   As Long
    MajorOperatingSystemVersion     As Integer
    MinorOperatingSystemVersion     As Integer
    MajorImageVersion               As Integer
    MinorImageVersion               As Integer
    MajorSubsystemVersion           As Integer
    MinorSubsystemVersion           As Integer
    W32VersionValue                 As Long
    SizeOfImage                     As Long
    SizeOfHeaders                   As Long
    CheckSum                        As Long
    SubSystem                       As Integer
    DllCharacteristics              As Integer
    SizeOfStackReserve              As Long
    SizeOfStackCommit               As Long
    SizeOfHeapReserve               As Long
    SizeOfHeapCommit                As Long
    LoaderFlags                     As Long
    NumberOfRvaAndSizes             As Long
    DataDirectory(15)               As IMAGE_DATA_DIRECTORY
End Type
Private Type IMAGE_FILE_HEADER
    Machine                         As Integer
    NumberOfSections                As Integer
    TimeDateStamp                   As Long
    PointerToSymbolTable            As Long
    NumberOfSymbols                 As Long
    SizeOfOptionalHeader            As Integer
    Characteristics                 As Integer
End Type
Private Type IMAGE_NT_HEADERS
    Signature                       As Long
    FileHeader                      As IMAGE_FILE_HEADER
    OptionalHeader                  As IMAGE_OPTIONAL_HEADER
End Type
Private Type IMAGE_SECTION_HEADER
    SectionName(1)                  As Long
    VirtualSize                     As Long
    VirtualAddress                  As Long
    SizeOfRawData                   As Long
    PointerToRawData                As Long
    PointerToRelocations            As Long
    PointerToLinenumbers            As Long
    NumberOfRelocations             As Integer
    NumberOfLinenumbers             As Integer
    Characteristics                 As Long
End Type
Private Type IMAGE_IMPORT_DESCRIPTOR
    Characteristics                 As Long
    TimeDateStamp                   As Long
    ForwarderChain                  As Long
    pName                           As Long
    FirstThunk                      As Long
End Type

Private Type IMAGE_BASE_RELOCATION
    VirtualAddress                  As Long
    SizeOfBlock                     As Long
End Type

Private Type UNICODE_STRING
    Length                          As Integer
    MaxLength                       As Integer
    lpBuffer                        As Long
End Type
Private Type PROCESS_BASIC_INFORMATION
    ExitStatus                      As Long
    PebBaseAddress                  As Long
    AffinityMask                    As Long
    BasePriority                    As Long
    UniqueProcessId                 As Long
    InheritedFromUniqueProcessId    As Long
End Type
Public Type LIST_ENTRY
    Flink                           As Long
    Blink                           As Long
End Type
Public Type PEB_LDR_DATA
    Length                          As Long
    Initialized                     As Long
    SsHandle                        As Long
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
End Type
Public Type LDR_MODULE
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitOrderModuleList           As LIST_ENTRY
    BaseAddress                     As Long
    EntryPoint                      As Long
    SizeOfImage                     As Long
    FullDllName                     As UNICODE_STRING
    BaseDllName                     As UNICODE_STRING
    Flags                           As Long
    LoadCount                       As Integer
    TlsIndex                        As Integer
    HashTableEntry                  As LIST_ENTRY
    TimeDateStamp                   As Long
End Type

Public Enum NTGLB_Flags
    FLG_STOP_ON_EXCEPTION = &H1
    FLG_SHOW_LDR_SNAPS = &H2
    FLG_DEBUG_INITIAL_COMMAND = &H4
    FLG_STOP_ON_HUNG_GUI = &H8
    FLG_HEAP_ENABLE_TAIL_CHECK = &H10
    FLG_HEAP_ENABLE_FREE_CHECK = &H20
    FLG_HEAP_VALIDATE_PARAMETERS = &H40
    FLG_HEAP_VALIDATE_ALL = &H80
    FLG_POOL_ENABLE_TAIL_CHECK = &H100           '3.51 to 5.0
    FLG_APPLICATION_VERIFIER = &H100             '5.1+
    FLG_MONITOR_SILENT_PROCESS_EXIT = &H200      '6.1+ only
    FLG_POOL_ENABLE_TAGGING = &H400
    FLG_HEAP_ENABLE_TAGGING = &H800
    FLG_USER_STACK_TRACE_DB = &H1000
    FLG_KERNEL_STACK_TRACE_DB = &H2000
    FLG_MAINTAIN_OBJECT_TYPELIST = &H4000
    FLG_HEAP_ENABLE_TAG_BY_DLL = &H8000&
    FLG_IGNORE_DEBUG_PRIV = &H10000              '3.51 to 4.0
    FLG_DISABLE_STACK_EXTENSION = &H10000        '5.1+(5.0 is undef)
    FLG_ENABLE_CSRDEBUG = &H20000
    FLG_ENABLE_KDEBUG_SYMBOL_LOAD = &H40000
    FLG_DISABLE_PAGE_KERNEL_STACKS = &H80000
    FLG_HEAP_ENABLE_CALL_TRACING = &H100000      '3.51 to 4.0
    FLG_ENABLE_SYSTEM_CRIT_BREAKS = &H100000     '5.1+ (5.0 is undef)
    FLG_HEAP_DISABLE_COALESCING = &H200000
    FLG_ENABLE_CLOSE_EXCEPTIONS = &H400000       '4.0+
    FLG_ENABLE_EXCEPTION_LOGGING = &H800000      '4.0+
    FLG_ENABLE_HANDLE_TYPE_TAGGING = &H1000000   '4.0+
    FLG_HEAP_PAGE_ALLOCS = &H2000000             '4.0+
    FLG_DEBUG_INITIAL_COMMAND_EX = &H4000000     '4.0+
    FLG_DISABLE_DBGPRINT = &H8000000             '5.0+
    FLG_CRITSEC_EVENT_CREATION = &H10000000      '5.0+
    FLG_LDR_TOP_DOWN = &H20000000                '5.1-6.2
    FLG_STOP_ON_UNHANDLED_EXCEPTION = &H20000000 '6.3+
    FLG_ENABLE_HANDLE_EXCEPTIONS = &H40000000    '5.1+
    FLG_DISABLE_PROTDLLS = &H80000000             '5.0+
End Enum

Public Enum APP_COMPAT_FLAGS
    KACF_OLDGETSHORTPATHNAME = &H1
    KACF_VERSIONLIE_NOT_USED = &H2
    KACF_GETDISKFREESPACE = &H8
    KACF_FTMFROMCURRENTAPT = &H20
    KACF_DISALLOWORBINDINGCHANGES = &H40
    KACF_OLE32VALIDATEPTRS = &H80
    KACF_DISABLECICERO = &H100
    KACF_OLE32ENABLEASYNCDOCFILE = &H200
    KACF_OLE32ENABLELEGACYEXCEPTIONHANDLING = &H400
    KACF_RPCDISABLENDRCLIENTHARDENING = &H800
    KACF_RPCDISABLENDRMAYBENULL_SIZEIS = &H1000
    KACF_DISABLEALLDDEHACK_NOT_USED = &H2000
    KACF_RPCDISABLENDR61_RANGE = &H4000
    KACF_RPC32ENABLELEGACYEXCEPTIONHANDLING = &H8000&
    KACF_OLE32DOCFILEUSELEGACYNTFSFLAGS = &H10000
    KACF_RPCDISABLENDRCONSTIIDCHECK = &H20000
    KACF_USERDISABLEFORWARDERPATCH = &H40000
    KACF_OLE32DISABLENEW_WMPAINT_DISPATCH = &H100000
    KACF_ADDRESTRICTEDSIDINCOINITIALIZESECURITY = &H200000
    KACF_ALLOCDEBUGINFOFORCRITSECTIONS = &H400000
    KACF_OLEAUT32ENABLEUNSAFELOADTYPELIBRELATIVE = &H800000
    KACF_ALLOWMAXIMIZEDWINDOWGAMMA = &H1000000
    KACF_DONOTADDTOCACHE = &H80000000
End Enum

'Generally, Win XP - 8
'https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/pebteb/peb/bitfield.htm
Public Enum PEB_BITFIELD_OLD
    PebImageUsedLargePages = &H1
    PebIsProtectedProcess = &H2  'V+
    PebIsLegacyProcess = &H4  'V-8
    PebIsImageDynamicallyRelocated = &H8  'V+
    PebSkipPatchingUser32Forwarders = &H10
    PebIsPackagedProcess = &H20
    PebIsAppContainer = &H40
    PebIsProtectedProcessLight = &H80
End Enum

'Generally, Windows 10-11
Public Enum PEB_BITFIELD_NEW
    PebNImageUsedLargePages = &H1
    PebNIsProtectedProcess = &H2  'V+
    PebNIsImageDynamicallyRelocated = &H4  'V+
    PebNSkipPatchingUser32Forwarders = &H8
    PebNIsPackagedProcess = &H10
    PebNIsAppContainer = &H20
    PebNIsProtectedProcessLight = &H40
    PebNIsLongPathAwareProcess = &H80
End Enum

Public Type QLARGE_INTEGER
    #If (TWINBASIC = 1) Or (Win64 = 1) Then
    QuadPart As LongLong
    #Else
    lowpart As Long
    highpart As Long
    #End If
End Type

Public Enum ImageSubsystemType
    IMAGE_SUBSYSTEM_UNKNOWN = 0   ' Unknown subsystem.
    IMAGE_SUBSYSTEM_NATIVE = 1   ' Image doesn't require a subsystem (e.g. kernel mode drivers).
    IMAGE_SUBSYSTEM_WINDOWS_GUI = 2   ' Image runs in the Windows GUI subsystem.
    IMAGE_SUBSYSTEM_WINDOWS_CUI = 3   ' Image runs in the Windows character subsystem.
    IMAGE_SUBSYSTEM_OS2_CUI = 5   ' image runs in the OS/2 character subsystem.
    IMAGE_SUBSYSTEM_POSIX_CUI = 7   ' image runs in the Posix character subsystem.
    IMAGE_SUBSYSTEM_NATIVE_WINDOWS = 8   ' image is a native Win9x driver.
    IMAGE_SUBSYSTEM_WINDOWS_CE_GUI = 9   ' Image runs in the Windows CE subsystem.
    IMAGE_SUBSYSTEM_EFI_APPLICATION = 10   '
    IMAGE_SUBSYSTEM_EFI_BOOT_SERVICE_DRIVER = 11   '
    IMAGE_SUBSYSTEM_EFI_RUNTIME_DRIVER = 12   '
    IMAGE_SUBSYSTEM_EFI_ROM = 13
    IMAGE_SUBSYSTEM_XBOX = 14
    IMAGE_SUBSYSTEM_WINDOWS_BOOT_APPLICATION = 16
    IMAGE_SUBSYSTEM_XBOX_CODE_CATALOG = 17
End Enum

'See also: https://www.nirsoft.net/kernel_struct/vista/PEB.html
Public Type PEB                                                         'thanks to fafalone
    InheritedAddressSpace As Byte
    ReadImageFileExecOptions As Byte
    BeingDebugged As Byte
    BitField As Byte '// PEB_BITFIELD_OLD on XP, https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/pebteb/peb/bitfield.htm
    Mutant As LongPtr
    ImageBaseAddress As LongPtr
    Ldr As LongPtr '// Pointer to PEB_LDR_DATA
    ProcessParameters As LongPtr '// RTL_USER_PROCESS_PARAMETERS
    SubSystemData As LongPtr
    ProcessHeap As LongPtr
    FastPebLock As LongPtr
    AtlThunkSListPtr As LongPtr
    SparePtr2 As LongPtr
    EnvironmentUpdateCount As Long
    KernelCallbackTable As LongPtr
    SystemReserved(0) As Long
    SpareUlong As Long
    FreeList As LongPtr
    TlsExpansionCounter As Long
    TlsBitmap As LongPtr
    TlsBitmapBits(1) As Long
    ReadOnlySharedMemoryBase As LongPtr
    ReadOnlySharedMemoryHeap As LongPtr
    ReadOnlyStaticServerData As LongPtr
    AnsiCodePageData As LongPtr
    OemCodePageData As LongPtr
    UnicodeCaseTableData As LongPtr
    NumberOfProcessors As Long
    NtGlobalFlag As NTGLB_Flags
    #If (TWINBASIC = 0) And (Win64 = 0) Then
    pad(3) As Byte
    #End If
    CriticalSectionTimeout As QLARGE_INTEGER
    HeapSegmentReserve As LongPtr
    HeapSegmentCommit As LongPtr
    HeapDeCommitTotalFreeThreshold As LongPtr
    HeapDeCommitFreeBlockThreshold As LongPtr
    NumberOfHeaps As Long
    MaximumNumberOfHeaps As Long
    ProcessHeaps As LongPtr
    GdiSharedHandleTable As LongPtr
    ProcessStarterHelper As LongPtr
    GdiDCAttributeList As Long
    LoaderLock As LongPtr
    OSMajorVersion As Long
    OSMinorVersion As Long
    OSBuildNumber As Integer
    OSCSDVersion As Integer
    OSPlatformId As Long
    ImageSubsystem As ImageSubsystemType
    ImageSubsystemMajorVersion As Long
    ImageSubsystemMinorVersion As Long
    ImageProcessAffinityMask As LongPtr
    #If Win64 Then
    GdiHandleBuffer(59) As Long
    #Else
    GdiHandleBuffer(33) As Long
    #End If
    PostProcessInitRoutine As LongPtr
    TlsExpansionBitmap As LongPtr
    TlsExpansionBitmapBits(31) As Long
    SessionId As Long
    AppCompatFlagsHi As Long
    AppCompatFlags As APP_COMPAT_FLAGS 'ULARGE_INTEGER
    AppCompatFlagUser As QLARGE_INTEGER
    pShimData As LongPtr
    AppCompatInfo As LongPtr
    CSDVersion As UNICODE_STRING
    ActivationContextData As LongPtr
    ProcessAssemblyStorageMap As LongPtr
    SystemDefaultActivationContextData As LongPtr
    SystemAssemblyStorageMap As LongPtr
    MinimumStackCommit As LongPtr
    #If (TWINBASIC = 0) And (Win64 = 0) Then
    pad2(3) As Byte
    #End If
End Type

Private Const IMAGE_FILE_MACHINE_I386               As Long = &H14C
Private Const IMAGE_DOS_SIGNATURE                   As Long = &H5A4D
Private Const IMAGE_NT_SIGNATURE                    As Long = &H4550&
Private Const IMAGE_NT_OPTIONAL_HDR32_MAGIC         As Long = &H10B&
Private Const IMAGE_FILE_RELOCS_STRIPPED            As Long = &H1
Private Const IMAGE_FILE_EXECUTABLE_IMAGE           As Long = &H2
Private Const IMAGE_FILE_32BIT_MACHINE              As Long = &H100
Private Const IMAGE_DIRECTORY_ENTRY_IMPORT          As Long = 1
Private Const IMAGE_DIRECTORY_ENTRY_BASERELOC       As Long = 5
Private Const IMAGE_SCN_MEM_EXECUTE                 As Long = &H20000000
Private Const IMAGE_SCN_MEM_READ                    As Long = &H40000000
Private Const IMAGE_SCN_MEM_WRITE                   As Long = &H80000000
Private Const IMAGE_REL_BASED_HIGHLOW               As Long = 3
Private Const HEAP_NO_SERIALIZE                     As Long = &H1
Private Const STATUS_SUCCESS                        As Long = 0
Private Const STATUS_INFO_LENGTH_MISMATCH           As Long = &HC0000004
Private Const ProcessBasicInformation               As Long = 0
Private Const INVALID_HANDLE_VALUE                  As Long = -1
Private Const FILE_MAP_READ                         As Long = &H4
Private Const PAGE_READONLY                         As Long = 2&
Private Const GENERIC_READ                          As Long = &H80000000
Private Const OPEN_EXISTING                         As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL                 As Long = &H80

Private Declare Function CreateFile Lib "kernel32" _
                         Alias "CreateFileW" ( _
                         ByVal lpFileName As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwShareMode As Long, _
                         ByRef lpSecurityAttributes As Any, _
                         ByVal dwCreationDisposition As Long, _
                         ByVal dwFlagsAndAttributes As Long, _
                         ByVal hTemplateFile As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" ( _
                         ByVal hFileMappingObject As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwFileOffsetHigh As Long, _
                         ByVal dwFileOffsetLow As Long, _
                         ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" ( _
                         ByVal lpBaseAddress As Long) As Long
Private Declare Function OpenFileMapping Lib "kernel32" _
                         Alias "OpenFileMappingW" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" _
                         Alias "CreateFileMappingW" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileMappingAttributes As Any, _
                         ByVal flProtect As Long, _
                         ByVal dwMaximumSizeHigh As Long, _
                         ByVal dwMaximumSizeLow As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileSize As Any) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As Any, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As Any, lpLocalFileTime As FILETIME) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" (lpSystemTime As SYSTEMTIME, vtime As Date) As Long
Private Declare Function RtlGetCurrentPeb Lib "ntdll" () As Long

Private Function GetBaseAddress() As LongPtr
    Dim lpPeb As LongPtr
    Dim tPeb As PEB
    lpPeb = RtlGetCurrentPeb()
    If lpPeb Then
        CopyMemory tPeb, ByVal lpPeb, LenB(tPeb)
        GetBaseAddress = tPeb.ImageBaseAddress
    End If
End Function

Private Sub GetFileMapping(sPath As String, out_hFile As Long, out_hMap As Long, out_pView As Long)
    
    Dim TSize As Currency
    Dim addrHigh As Long
    Dim addrLow As Long
    
    out_hFile = CreateFile(StrPtr(sPath), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, _
        FILE_FLAG_NO_BUFFERING Or FILE_FLAG_SEQUENTIAL_SCAN, 0)
    
    If out_hFile = INVALID_HANDLE_VALUE Then
        Exit Sub
    End If
    
    If GetFileSizeEx(out_hFile, TSize) = 0 Or TSize = 0 Then
        Exit Sub
    End If
    
    out_hMap = CreateFileMapping(out_hFile, ByVal 0&, PAGE_READONLY, 0, 0, 0)
    If out_hMap = 0 Then
        CloseHandle out_hFile
        Exit Sub
    End If
    
    addrHigh = 0
    addrLow = 0
    out_pView = MapViewOfFile(out_hMap, FILE_MAP_READ, addrHigh, addrLow, TSize)
    
End Sub

Private Sub CloseFileMapping(hFile As Long, hMap As Long, pView As Long)
    
    If pView Then
        UnmapViewOfFile pView
    End If
    
    If hMap Then
        CloseHandle hMap
    End If
    
    If hFile Then
        CloseHandle hFile
    End If
    
End Sub

Public Function GetOwnCompilationDate() As String
    Dim dDate As Date
    If inIDE Then
        GetOwnCompilationDate = "IDE"
    Else
        If GetPeCompilationTime(vbNullString, dDate) Then
            GetOwnCompilationDate = Year(dDate) & "-" & Right$("0" & Month(dDate), 2) & "-" & Right$("0" & Day(dDate), 2)
        Else
            GetOwnCompilationDate = "?"
        End If
    End If
End Function

Public Function GetPeCompilationTime(sPath As String, out_Date As Date) As Boolean

    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetPeCompilationTime - Begin"

    Dim NtHdr       As IMAGE_NT_HEADERS
    Dim pBase       As Long
    Dim pRawData    As Long
    Dim hFile       As Long
    Dim hMap        As Long
    Dim pView       As Long
    Dim dwUtcOffset As Long
    
    If Len(sPath) <> 0 Then
        GetFileMapping sPath, hFile, hMap, pView
    Else
        pView = GetBaseAddress()
    End If
    
    If pView = 0 Then
        GoTo clean
    End If
    
    ' // Get IMAGE_NT_HEADERS
    If GetImageNtHeaders(pView, NtHdr) = 0 Then
        GoTo clean
    End If
    
    ' // Check flags
    If NtHdr.FileHeader.Machine <> IMAGE_FILE_MACHINE_I386 Or _
       (NtHdr.FileHeader.Characteristics And IMAGE_FILE_EXECUTABLE_IMAGE) = 0 Or _
       (NtHdr.FileHeader.Characteristics And IMAGE_FILE_32BIT_MACHINE) = 0 Then GoTo clean
    
    Call GetTimeZoneOffset(dwUtcOffset)
    out_Date = DateAdd("s", NtHdr.FileHeader.TimeDateStamp, #1/1/1970#)
    out_Date = DateAdd("n", -dwUtcOffset, out_Date)
    
    GetPeCompilationTime = True
    
clean:
    If Len(sPath) <> 0 Then
        CloseFileMapping hFile, hMap, pView
    End If
    
    AppendErrorLogCustom "GetPeCompilationTime - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetPeCompilationTime"
    If inIDE Then Stop: Resume Next
End Function

Private Function GetImageNtHeaders( _
                 ByVal pBase As Long, _
                 ByRef NtHdr As IMAGE_NT_HEADERS) As Long
    Dim dosHdr  As IMAGE_DOS_HEADER
    Dim pNtHdr  As Long
    
    ' // Get DOS header
    CopyMemory ByVal VarPtr(dosHdr), ByVal pBase, Len(dosHdr)
    
    ' // Check MZ signature and alignment
    If (dosHdr.e_magic_e_cblp And &HFFFF&) <> IMAGE_DOS_SIGNATURE Or _
       (dosHdr.e_lfanew And &H3) <> 0 Then
        Exit Function
    End If
    
    ' // Get pointer to NT headers
    pNtHdr = pBase + dosHdr.e_lfanew
    
    ' // Get NT headers
    CopyMemory ByVal VarPtr(NtHdr), ByVal pNtHdr, Len(NtHdr)
    
    ' // Check NT signature
    If (NtHdr.Signature <> IMAGE_NT_SIGNATURE) Or _
        NtHdr.OptionalHeader.Magic <> IMAGE_NT_OPTIONAL_HDR32_MAGIC Or _
        NtHdr.FileHeader.SizeOfOptionalHeader <> Len(NtHdr.OptionalHeader) Then
        Exit Function
    End If
    
    GetImageNtHeaders = VarPtr(NtHdr)
    
End Function
