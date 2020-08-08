Attribute VB_Name = "modProcess"
'[modProcess.bas]

'
' Windows Process module by Alex Dragokas
'

Option Explicit

'Public Enum SHOWWINDOW_FLAGS
'    SW_HIDE = 0
'    SW_SHOWNORMAL = 1
'    SW_SHOWMINIMIZED = 2
'    SW_MAXIMIZE = 3
'    SW_SHOWMAXIMIZED = 3
'    SW_SHOWNOACTIVATE = 4
'    SW_SHOW = 5
'    SW_MINIMIZE = 6
'    SW_SHOWMINNOACTIVE = 7
'    SW_SHOWNA = 8
'    SW_RESTORE = 9
'    SW_SHOWDEFAULT = 10
'    SW_FORCEMINIMIZE = 11
'End Enum
'
'Public Type MY_PROC_ENTRY
'    Name        As String
'    Path        As String
'    PID         As Long
'    Threads     As Long
'    Priority    As Long
'    SessionID   As Long
'End Type
'
'Public Enum PROCESS_PRIORITY
'    ABOVE_NORMAL_PRIORITY_CLASS = &H8000&
'    BELOW_NORMAL_PRIORITY_CLASS = &H4000&
'    HIGH_PRIORITY_CLASS = &H80&
'    IDLE_PRIORITY_CLASS = &H40&
'    NORMAL_PRIORITY_CLASS = &H20&
'    PROCESS_MODE_BACKGROUND_BEGIN = &H100000
'    PROCESS_MODE_BACKGROUND_END = &H200000
'    REALTIME_PRIORITY_CLASS = &H100&
'End Enum
'
'Public Enum THREAD_PRIORITY
'    THREAD_MODE_BACKGROUND_BEGIN = &H10000
'    THREAD_MODE_BACKGROUND_END = &H20000
'    THREAD_PRIORITY_ABOVE_NORMAL = 1&
'    THREAD_PRIORITY_BELOW_NORMAL = -1&
'    THREAD_PRIORITY_HIGHEST = 2&
'    THREAD_PRIORITY_IDLE = -15&
'    THREAD_PRIORITY_LOWEST = -2&
'    THREAD_PRIORITY_NORMAL = 0&
'    THREAD_PRIORITY_TIME_CRITICAL = 15&
'End Enum
'
'Public Enum SECURITY_IMPERSONATION_LEVEL
'    SecurityAnonymous
'    SecurityIdentification
'    SecurityImpersonation
'    SecurityDelegation
'End Enum
'
'Public Enum TOKEN_TYPE
'    TokenPrimary = 1
'    TokenImpersonation
'End Enum
'
'Public Type PROCESS_INFORMATION
'    hProcess As Long
'    hThread As Long
'    dwProcessId As Long
'    dwThreadId As Long
'End Type
'
'Public Type STARTUPINFO
'    cb As Long
'    lpReserved As Long
'    lpDesktop As Long
'    lpTitle As Long
'    dwX As Long
'    dwY As Long
'    dwXSize As Long
'    dwYSize As Long
'    dwXCountChars As Long
'    dwYCountChars As Long
'    dwFillAttribute As Long
'    dwFlags As Long
'    wShowWindow As Integer
'    cbReserved2 As Integer
'    lpReserved2 As Byte
'    hStdInput As Long
'    hStdOutput As Long
'    hStdError As Long
'End Type
'
'Public Const PROCESS_SET_INFORMATION As Long = &H200&

'Private Type LARGE_INTEGER
'    LowPart     As Long
'    HighPart    As Long
'End Type
'
'Private Type CLIENT_ID
'    UniqueProcess   As Long  ' HANDLE
'    UniqueThread    As Long  ' HANDLE
'End Type
'
'Private Type UNICODE_STRING
'    length      As Integer
'    MaxLength   As Integer
'    lpBuffer    As Long
'End Type
'
'Private Type VM_COUNTERS
'    PeakVirtualSize             As Long
'    VirtualSize                 As Long
'    PageFaultCount              As Long
'    PeakWorkingSetSize          As Long
'    WorkingSetSize              As Long
'    QuotaPeakPagedPoolUsage     As Long
'    QuotaPagedPoolUsage         As Long
'    QuotaPeakNonPagedPoolUsage  As Long
'    QuotaNonPagedPoolUsage      As Long
'    PagefileUsage               As Long
'    PeakPagefileUsage           As Long
'End Type
'
'Private Type IO_COUNTERS
'    ReadOperationCount      As Currency 'ULONGLONG
'    WriteOperationCount     As Currency
'    OtherOperationCount     As Currency
'    ReadTransferCount       As Currency
'    WriteTransferCount      As Currency
'    OtherTransferCount      As Currency
'End Type
'
'Private Type SYSTEM_THREAD
'    KernelTime          As LARGE_INTEGER
'    UserTime            As LARGE_INTEGER
'    CreateTime          As LARGE_INTEGER
'    WaitTime            As Long
'    StartAddress        As Long
'    ClientId            As CLIENT_ID
'    Priority            As Long
'    BasePriority        As Long
'    ContextSwitchCount  As Long
'    State               As Long 'enum KTHREAD_STATE
'    WaitReason          As Long 'enum KWAIT_REASON
'    dReserved01         As Long
'End Type
'
'Private Type SYSTEM_PROCESS_INFORMATION
'    NextEntryOffset         As Long
'    NumberOfThreads         As Long
'    SpareLi1                As LARGE_INTEGER
'    SpareLi2                As LARGE_INTEGER
'    SpareLi3                As LARGE_INTEGER
'    CreateTime              As LARGE_INTEGER
'    UserTime                As LARGE_INTEGER
'    KernelTime              As LARGE_INTEGER
'    ImageName               As UNICODE_STRING
'    BasePriority            As Long
'    ProcessID               As Long
'    InheritedFromProcessId  As Long
'    HandleCount             As Long
'    SessionID               As Long
'    pPageDirectoryBase      As Long '_PTR
'    VirtualMemoryCounters   As VM_COUNTERS
'    PrivatePageCount        As Long
'    IoCounters              As IO_COUNTERS
'    Threads()               As SYSTEM_THREAD
'End Type
'
'Private Type PROCESSENTRY32
'    dwSize As Long
'    cntUsage As Long
'    th32ProcessID As Long
'    th32DefaultHeapID As Long
'    th32ModuleID As Long
'    cntThreads As Long
'    th32ParentProcessID As Long
'    pcPriClassBase As Long
'    dwFlags As Long
'    szExeFile As String * 260
'End Type
'
'Private Type MODULEENTRY32
'    dwSize As Long
'    th32ModuleID As Long
'    th32ProcessID As Long
'    GlblcntUsage As Long
'    ProccntUsage As Long
'    modBaseAddr As Long
'    modBaseSize As Long
'    hModule As Long
'    szModule  As String * 256
'    szExePath As String * 260
'End Type
'
'Private Type THREADENTRY32
'    dwSize As Long
'    dwRefCount As Long
'    th32ThreadID As Long
'    th32ProcessID As Long
'    dwBasePriority As Long
'    dwCurrentPriority As Long
'    dwFlags As Long
'End Type
'

Private Enum PROCESS_INFORMATION_CLASS
    ProcessBasicInformation
    ProcessQuotaLimits
    ProcessIoCounters
    ProcessVmCounters
    ProcessTimes
    ProcessBasePriority
    ProcessRaisePriority
    ProcessDebugPort
    ProcessExceptionPort
    ProcessAccessToken
    ProcessLdtInformation
    ProcessLdtSize
    ProcessDefaultHardErrorMode
    ProcessIoPortHandlers
    ProcessPooledUsageAndLimits
    ProcessWorkingSetWatch
    ProcessUserModeIOPL
    ProcessEnableAlignmentFaultFixup
    ProcessPriorityClass
    ProcessWx86Information
    ProcessHandleCount
    ProcessAffinityMask
    ProcessPriorityBoost
    MaxProcessInfoClass
End Enum

Private Type PROCESS_BASIC_INFORMATION64
    Reserved1       As Currency
    PebBaseAddress  As Currency
    Reserved2(1)    As Currency
    UniqueProcessId As Currency
    Reserved3       As Currency
End Type

Private Type LIST_ENTRY64
    Flink As Currency
    Blink As Currency
End Type

Private Type PEB_LDR_DATA64
    Length                          As Long
    Initialized                     As Byte
    SsHandle                        As Currency
    InLoadOrderModuleList           As LIST_ENTRY64
    InMemoryOrderModuleList         As LIST_ENTRY64
    InInitializationOrderModuleList As LIST_ENTRY64
End Type

'Structure is cut down to ProcessHeap.
Private Type PEB64
    InheritedAddressSpace       As Byte
    ReadImageFileExecOptions    As Byte
    BeingDebugged               As Byte
    Spare                       As Byte
    Align                       As Long
    Mutant                      As Currency
    ImageBaseAddress            As Currency
    LoaderData                  As Currency
    ProcessParameters           As Currency
    SubSystemData               As Currency
    ProcessHeap                 As Currency
End Type

Private Type UNICODE_STRING64
    Length          As Integer
    MaximumLength   As Integer
    Align           As Long
    Buffer          As Currency
End Type

Private Type LDR_DATA_TABLE_ENTRY64
    InLoadOrderModuleList           As LIST_ENTRY64
    InMemoryOrderModuleList         As LIST_ENTRY64
    InInitializationOrderModuleList As LIST_ENTRY64
    BaseAddress                     As Currency
    EntryPoint                      As Currency
    SizeOfImage                     As Currency
    FullDllName                     As UNICODE_STRING64
    BaseDllName                     As UNICODE_STRING64
    Flags                           As Long
    LoadCount                       As Integer
    TlsIndex                        As Integer
    HashTableEntry                  As LIST_ENTRY64
    TimeDateStamp                   As Currency
End Type

Private Type TwoLongs
    HiLong As Long
    LowLong As Long
End Type

Public Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long
    PebBaseAddress As Long
    AffinityMask As Long
    BasePriority As Long
    UniqueProcessId As Long
    InheritedFromUniqueProcessId As Long
End Type

'Private Declare Function NtQuerySystemInformation Lib "ntdll.dll" (ByVal infoClass As Long, Buffer As Any, ByVal BufferSize As Long, ret As Long) As Long
'Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
'Private Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
'Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, lpFilePart As Long) As Long
'Private Declare Function QueryFullProcessImageName Lib "kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As Long, ByVal lpdwSize As Long) As Long
'Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
'Private Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceW" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long
'Private Declare Sub memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
'
'Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
'Private Declare Function Process32First Lib "kernel32.dll" Alias "Process32FirstW" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
'Private Declare Function Process32Next Lib "kernel32.dll" Alias "Process32NextW" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
'
'Private Declare Function Module32First Lib "kernel32.dll" Alias "Module32FirstW" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
'Private Declare Function Module32Next Lib "kernel32.dll" Alias "Module32NextW" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
'Private Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapshot As Long, uThread As THREADENTRY32) As Long
'Private Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef ThreadEntry As THREADENTRY32) As Long
'Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'
'Private Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProcess As Long) As Long
'Private Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProcess As Long) As Long
'Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
'Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
'Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
'Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
'Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
'Private Declare Function ZwSetInformationProcess Lib "ntdll.dll" (ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
'
'Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
'Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'
'Private Declare Function SHRunDialog Lib "shell32.dll" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
'
'Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStrDest As Long, ByVal lpStrSrc As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal length As Long)
'
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'

'Public Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
'Public Declare Function GetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long) As Long
'Public Declare Function SetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long, ByVal nPriority As Long) As Long
'Public Declare Function GetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long) As Long
'Public Declare Function SetProcessPriorityBoost Lib "kernel32.dll" (ByVal hProcess As Long, ByVal DisablePriorityBoost As Long) As Long
'Public Declare Function GetProcessPriorityBoost Lib "kernel32.dll" (ByVal hThread As Long, pDisablePriorityBoost As Long) As Long
'Public Declare Function SetThreadPriorityBoost Lib "kernel32.dll" (ByVal hThread As Long, ByVal DisablePriorityBoost As Long) As Long
'Public Declare Function GetThreadPriorityBoost Lib "kernel32.dll" (ByVal hThread As Long, pDisablePriorityBoost As Long) As Long
'Public Declare Function GetProcessID Lib "kernel32.dll" (ByVal Process As Long) As Long
'
'Public Declare Function CreateProcessWithTokenW Lib "Advapi32.dll" (ByVal hToken As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
'Public Declare Function OpenThreadToken Lib "Advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
'Public Declare Function DuplicateTokenEx  Lib "Advapi32.dll" (byval hExistingToken as Long, byval dwDesiredAccess as Long, lpTokenAttributes as SECURITY_ATTRIBUTES , byval ImpersonationLevel as SECURITY_IMPERSONATION_LEVEL
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByVal lpFileTime As Long, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32.dll" (ByVal lpTimeZone As Any, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long

Private Declare Function NtWow64QueryInformationProcess64 Lib "ntdll.dll" ( _
    ByVal ProcessHandle As Long, _
    ByVal ProcessInformationClass As PROCESS_INFORMATION_CLASS, _
    ByVal ProcessInformation As Long, _
    ByVal ProcessInformationLength As Long, _
    ByVal ReturnLength As Long) As Long

Private Declare Function NtWow64ReadVirtualMemory64 Lib "ntdll.dll" ( _
    ByVal ProcessHandle As Long, _
    ByVal BaseAddress As Currency, _
    ByVal Buffer As Long, _
    ByVal Size As Currency, _
    ByVal NumberOfBytesRead As Long) As Long

'typedef NTSTATUS(NTAPI *_NtWow64QueryInformationProcess64)(
'    IN HANDLE ProcessHandle,
'    ULONG ProcessInformationClass,
'    OUT PVOID ProcessInformation,
'    IN ULONG ProcessInformationLength,
'    OUT PULONG ReturnLength OPTIONAL);
'
'typedef NTSTATUS(NTAPI *_NtWow64ReadVirtualMemory64)(
'    IN HANDLE ProcessHandle,
'    IN DWORD64 BaseAddress,
'    OUT PVOID Buffer,
'    IN ULONG64 Size,
'    OUT PDWORD64 NumberOfBytesRead);


'Private Const TH32CS_SNAPPROCESS = &H2
'Private Const TH32CS_SNAPMODULE = &H8
'Private Const TH32CS_SNAPTHREAD = &H4
'Private Const PROCESS_TERMINATE = &H1
'Private Const PROCESS_QUERY_INFORMATION = 1024
'Private Const PROCESS_QUERY_LIMITED_INFORMATION = &H1000
'Private Const PROCESS_VM_READ = 16
'Private Const THREAD_SUSPEND_RESUME = &H2
'Private Const PROCESS_SUSPEND_RESUME As Long = &H800&
'
'Private Const SystemProcessInformation      As Long = &H5&
'Private Const STATUS_INFO_LENGTH_MISMATCH   As Long = &HC0000004
'Private Const STATUS_SUCCESS                As Long = 0&
'Private Const ERROR_PARTIAL_COPY            As Long = 299&
'Private Const INVALID_HANDLE_VALUE          As Long = &HFFFFFFFF
'Private Const ERROR_ACCESS_DENIED           As Long = 5&

Public Const THREAD_SET_INFORMATION As Long = &H20&
Public Const THREAD_SET_LIMITED_INFORMATION  As Long = &H400&

Public Sub AcquirePrivileges()
    'List of privileges: https://msdn.microsoft.com/en-us/library/windows/desktop/bb530716(v=vs.85).aspx
    '                    https://msdn.microsoft.com/en-us/library/windows/desktop/ee695867(v=vs.85).aspx
    SetCurrentProcessPrivileges "SeDebugPrivilege"
    If (SetCurrentProcessPrivileges("SeBackupPrivilege") And _
        SetCurrentProcessPrivileges("SeRestorePrivilege")) Then
        g_FileBackupFlag = FILE_FLAG_BACKUP_SEMANTICS
    End If
    SetCurrentProcessPrivileges "SeTakeOwnershipPrivilege"
    SetCurrentProcessPrivileges "SeSecurityPrivilege"       'SACL
    'SetCurrentProcessPrivileges "SeAssignPrimaryTokenPrivilege" '(SYSTEM, LocalService или NetworkService) only
    SetCurrentProcessPrivileges "SeIncreaseQuotaPrivilege"  'CreateProcessWithTokenW
    SetCurrentProcessPrivileges "SeImpersonatePrivilege"    'CreateProcessWithTokenW
    SetCurrentProcessPrivileges "SeChangeNotifyPrivilege"   'NtQueryInformationFile
    SetCurrentProcessPrivileges "SeIncreaseBasePriorityPrivilege" 'required by SetProcessIOPriority()
End Sub

Public Function KillProcess(lPID&) As Boolean
    Dim hProcess&, lCriticalFlag&
    If lPID = 0 Then Exit Function
    
    Dim sTaskKill As String
    If OSver.Bitness = "x64" And FolderExists(sWinDir & "\sysnative") Then
        sTaskKill = EnvironW("%SystemRoot%") & "\Sysnative\taskkill.exe"
    Else
        sTaskKill = EnvironW("%SystemRoot%") & "\System32\taskkill.exe"
    End If
    
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProcess <> 0 Then
        lCriticalFlag = 0
        If GetProcessCriticalFlag(lPID, lCriticalFlag) Then
            If lCriticalFlag <> 0 Then
                Call SetProcessCriticalFlag(lPID, False)
            End If
            Call GetProcessCriticalFlag(lPID, lCriticalFlag)
        End If
        
        If lCriticalFlag = 0 Then
            If TerminateProcess(hProcess, 0) = 0 Then
                If FileExists(sTaskKill) Then
                    Proc.ProcessRun sTaskKill, "/F /PID " & lPID, , 0
                    Proc.WaitForTerminate , , , 10000
                End If
            End If
        End If
        CloseHandle hProcess
    Else
        If FileExists(sTaskKill) Then
            Proc.ProcessRun sTaskKill, "/F /PID " & lPID, , 0
            Proc.WaitForTerminate , , , 10000
        End If
    End If
    
    'SleepNoLock 500
    If Proc.IsRunned(, lPID) Then
        If OSver.MajorMinor >= 6 Then
            'The selected process could not be killed. It may have already closed, or it may be protected by Windows.
            'This process might be a service, which you can stop from the Services applet in Control Panel -> Admin Tools.
            '(To load this window, click 'Win + R' and enter 'services.msc')
            If Not g_bNoGUI Then
                MsgBoxW Translate(1654), vbCritical
            End If
        Else
            'The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows.
            If Not g_bNoGUI Then
                MsgBoxW Translate(1652), vbCritical
            End If
        End If
    Else
        KillProcess = True
    End If
End Function

Public Function PauseProcess(lPID As Long) As Boolean
    On Error GoTo ErrorHandler:

    Dim hThread&, hProc&, i&, SysThread() As SYSTEM_THREAD
    
    If Not bIsWinNT And Not bIsWinME Then Exit Function
    If lPID = 0 Or lPID = GetCurrentProcessId Then Exit Function
    If lPID = MyParentProc.pid Then Exit Function
    
    If IsProcedureAvail("NtSuspendProcess", "ntdll.dll") Then
        hProc = OpenProcess(PROCESS_SUSPEND_RESUME, 0, lPID)
    
        If hProc <> 0 Then
            Call NtSuspendProcess(hProc)
            CloseHandle hProc
            
            If IsProcessSuspended(lPID) Then
                PauseProcess = True
                Exit Function
            End If
        End If
    End If
    
    For i = 0 To GetThreads_Zw(lPID, SysThread) - 1
        
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, SysThread(i).ClientId.UniqueThread)
        If hThread <> 0 Then
            Call SuspendThread(hThread)
            CloseHandle hThread
        End If
    Next
    
    PauseProcess = IsProcessSuspended(lPID)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "PauseProcess. PID:", lPID
    If inIDE Then Stop: Resume Next
End Function

Public Function ResumeProcess(lPID As Long) As Boolean
    On Error GoTo ErrorHandler:

    Dim hThread&, hProc&, i&, SysThread() As SYSTEM_THREAD
    
    If Not (bIsWinNT Or bIsWinME) Then Exit Function
    If lPID = 0 Or lPID = GetCurrentProcessId Then Exit Function
    
    If IsProcedureAvail("NtResumeProcess", "ntdll.dll") Then
        hProc = OpenProcess(PROCESS_SUSPEND_RESUME, 0, lPID)
        
        If hProc <> 0 Then
            Call NtResumeProcess(hProc)
            CloseHandle hProc
            
            If IsProcessResumed(lPID) Then
                ResumeProcess = True
                Exit Function
            End If
        End If
    End If
    
    For i = 0 To GetThreads_Zw(lPID, SysThread) - 1
        
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, SysThread(i).ClientId.UniqueThread)
        If hThread <> 0 Then
            Call ResumeThread(hThread)
            CloseHandle hThread
        End If
    Next
    
    ResumeProcess = IsProcessResumed(lPID)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ResumeProcess. PID:", lPID
    If inIDE Then Stop: Resume Next
End Function

Public Function KillProcessByFile(sPath$, Optional bForceMicrosoft As Boolean) As Boolean
    Dim hProcess&, i&, sTaskKill As String, lCriticalFlag As Long
    Dim aPID() As Long, bKilled As Boolean
    'Note: this sub is silent - it displays no errors !
    
    If sPath = vbNullString Then Exit Function
    
    sPath = FindOnPath(sPath, True)
    
    If Not bForceMicrosoft Then
        If IsMicrosoftFile(sPath, True) Then Exit Function
    End If
    
    If IsSystemCriticalProcessPath(sPath) Then Exit Function
    
    If StrComp(sPath, MyParentProc.Path, 1) = 0 And Not StrEndWith(sPath, "explorer.exe") Then Exit Function
    
    If Not bIsWinNT Then
        KillProcessByFile = KillProcess9xByFile(sPath)
        Exit Function
    End If
    
    If OSver.Bitness = "x64" And FolderExists(sWinDir & "\sysnative") Then
        sTaskKill = EnvironW("%SystemRoot%") & "\Sysnative\taskkill.exe"
    Else
        sTaskKill = EnvironW("%SystemRoot%") & "\System32\taskkill.exe"
    End If
    
    Dim lNumProcesses As Long
    Dim Process() As MY_PROC_ENTRY
    
    lNumProcesses = GetProcesses(Process)
    
    If lNumProcesses Then
        
        For i = 0 To UBound(Process)
        
            If StrComp(sPath, Process(i).Path, 1) = 0 Then
                KillProcessByFile = False
                
                lCriticalFlag = 0
                If GetProcessCriticalFlag(Process(i).pid, lCriticalFlag) Then
                    If lCriticalFlag <> 0 Then
                        Call SetProcessCriticalFlag(Process(i).pid, False)
                    End If
                    Call GetProcessCriticalFlag(Process(i).pid, lCriticalFlag)
                End If
                
                If lCriticalFlag = 0 Then
                    PauseProcess Process(i).pid
                    hProcess = OpenProcess(PROCESS_TERMINATE, 0, Process(i).pid)
                    AddToArrayLong aPID, Process(i).pid
                    bKilled = False
                    If hProcess <> 0 Then
                        If TerminateProcess(hProcess, 0) <> 0 Then
                            bKilled = True
                            KillProcessByFile = True
                        End If
                        CloseHandle hProcess
                    End If
                    If Not bKilled Then
                        If FileExists(sTaskKill) Then
                            Proc.ProcessRun sTaskKill, "/F /PID " & Process(i).pid, , 0
                            Proc.WaitForTerminate , , , 10000
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    'get killing confirmation
    If AryPtr(aPID) Then
        'SleepNoLock 500
        For i = 0 To UBound(aPID)
            If Proc.IsRunned(, aPID(i)) Then
                KillProcessByFile = False
                Exit For
            End If
            DoEvents
        Next
    End If
End Function

Public Function PauseProcessByFile(sPath$) As Boolean
    Dim i&
    
    If StrComp(sPath, MyParentProc.Path, 1) = 0 Then
        PauseProcessByFile = True
        Exit Function
    End If
    
    'Note: this sub is silent - it displays no errors !
    If sPath = vbNullString Then Exit Function
    If Not bIsWinNT Then
        KillProcess9xByFile sPath
        Exit Function
    End If
    
    Dim lNumProcesses As Long
    Dim Process() As MY_PROC_ENTRY
    
    lNumProcesses = GetProcesses(Process)
        
    If lNumProcesses Then
        
        For i = 0 To UBound(Process)
        
            If StrComp(sPath, Process(i).Path, 1) = 0 Then
            
                PauseProcessByFile = PauseProcess(Process(i).pid)
            End If
        Next
    End If
End Function

Public Function KillProcess9xByFile(sPath$) As Boolean
    'Note: this sub is silent - it displays no errors!
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i&
    On Error Resume Next
    If sPath = vbNullString Then Exit Function
    If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
        'no PSAPI.DLL file or wrong version
        Exit Function
    End If

    lNumProcesses = lNeeded / 4
    For i = 1 To lNumProcesses
        hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
        If hProc <> 0 Then
            'Openprocess can return 0 but we ignore this since
            'system processes are somehow protected, further
            'processes CAN be opened.... silly windows

            lNeeded = 0
            sProcessName = String$(MAX_PATH, 0)
            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                sProcessName = TrimNull(sProcessName)
                If sProcessName <> vbNullString Then
                    If Left$(sProcessName, 1) = "\" Then sProcessName = Mid$(sProcessName, 2)
                    If Left$(sProcessName, 3) = "??\" Then sProcessName = Mid$(sProcessName, 4)
                    If InStr(1, sProcessName, "%Systemroot%", vbTextCompare) > 0 Then sProcessName = Replace$(sProcessName, "%Systemroot%", sWinDir, , , vbTextCompare)
                    If InStr(1, sProcessName, "Systemroot", vbTextCompare) > 0 Then sProcessName = Replace$(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)

                    If InStr(1, sProcessName, sPath, vbTextCompare) > 0 Then

                        'found the process!
                        PauseProcess lProcesses(i)
                        If TerminateProcess(hProc, 0) <> 0 Then
                            DoEvents
                            KillProcess9xByFile = True
                        End If
                        'CloseHandle hProc
                        'Exit Function
                    End If
                End If
            End If
            CloseHandle hProc
        End If
    Next i
End Function

Public Function GetProcesses(ProcList() As MY_PROC_ENTRY) As Long
    If OSver.MajorMinor >= 5.1 Then
        GetProcesses = GetProcesses_Zw(ProcList)
    Else
        GetProcesses = GetProcesses_2k(ProcList)
    End If
End Function

Public Function GetProcesses_2k(ProcList() As MY_PROC_ENTRY) As Long
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetProcesses_2k - Begin"
    
    Dim hSnap As Long
    Dim cnt As Long
    Dim uProcess As PROCESSENTRY32W
    
    ReDim ProcList(100)
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    
    If hSnap <> INVALID_HANDLE_VALUE Then
            
        uProcess.dwSize = LenB(uProcess)
        
        If Process32First(hSnap, uProcess) <> 0 Then
            Do
                ProcList(cnt).Name = StringFromPtrW(VarPtr(uProcess.szExeFile(0)))
                If ProcList(cnt).Name <> "[System Process]" And ProcList(cnt).Name <> "System" Then
                    ProcList(cnt).pid = uProcess.th32ProcessID
                    If 0 <> uProcess.th32ProcessID Then
                        ProcList(cnt).Path = TrimNull(GetFilePathByPID(uProcess.th32ProcessID))
                    End If
                    cnt = cnt + 1
                    If cnt > UBound(ProcList) Then ReDim Preserve ProcList(UBound(ProcList) + 100)
                End If
            Loop Until Process32Next(hSnap, uProcess) = 0
        End If
        
        CloseHandle hSnap: hSnap = 0
    End If

    If cnt > 1 Then
        ReDim Preserve ProcList(cnt - 1)
    End If
    GetProcesses_2k = cnt

    AppendErrorLogCustom "GetProcesses_2k - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetProcesses_2k"
    If inIDE Then Stop: Resume Next
End Function

Public Function IsProcessSuspended(ProcessID As Long) As Boolean
    Dim SysThread() As SYSTEM_THREAD
    Dim i As Long
    IsProcessSuspended = True
    
    For i = 0 To GetThreads_Zw(ProcessID, SysThread) - 1
        If SysThread(i).WaitReason <> Suspended Then
            IsProcessSuspended = False
            Exit Function
        End If
    Next
End Function

Public Function IsProcessResumed(ProcessID As Long) As Boolean
    Dim SysThread() As SYSTEM_THREAD
    Dim i As Long
    IsProcessResumed = True
    
    For i = 0 To GetThreads_Zw(ProcessID, SysThread) - 1
        If SysThread(i).WaitReason = Suspended Then
            IsProcessResumed = False
            Exit Function
        End If
    Next
End Function

Public Function GetThreads_Zw(ProcessID As Long, ThreadList() As SYSTEM_THREAD) As Long
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetThreads_Zw - Begin"
    
    Const SPI_SIZE      As Long = &HB8&                                 'SPI struct: http://www.informit.com/articles/article.aspx?p=22442&seqNum=5
    Const THREAD_SIZE   As Long = &H40&
    
    Dim cnt         As Long
    Dim ret         As Long
    Dim buf()       As Byte
    Dim Offset      As Long
    Dim Process     As SYSTEM_PROCESS_INFORMATION
    Dim i           As Long
    
    ReDim ProcList(200)
    
    SetCurrentProcessPrivileges "SeDebugPrivilege"
    
    If NtQuerySystemInformation(SystemProcessInformation, ByVal 0&, 0&, ret) = STATUS_INFO_LENGTH_MISMATCH Then
    
        ReDim buf(ret - 1)
        
        If NtQuerySystemInformation(SystemProcessInformation, buf(0), ret, ret) = STATUS_SUCCESS Then
        
            With Process
            
                Do
                    memcpy Process, buf(Offset), SPI_SIZE
                    
                    If .ProcessID = ProcessID Then

                        ReDim ThreadList(0 To .NumberOfThreads - 1)
                    
                        For i = 0 To .NumberOfThreads - 1
                            memcpy ThreadList(i), buf(Offset + SPI_SIZE + i * THREAD_SIZE), THREAD_SIZE
                        Next
                        
                        cnt = .NumberOfThreads
                        Exit Do
                    End If
                    
                    Offset = Offset + .NextEntryOffset
                    
                Loop While .NextEntryOffset
                
            End With
            
        End If
        
    End If
    
    GetThreads_Zw = cnt
    
    AppendErrorLogCustom "GetThreads_Zw - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetThreads_Zw"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetProcesses_Zw(ProcList() As MY_PROC_ENTRY) As Long    'Return -> Count of processes
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetProcesses_Zw - Begin"

    Const SPI_SIZE      As Long = &HB8&                                 'SPI struct: http://www.informit.com/articles/article.aspx?p=22442&seqNum=5
    Const THREAD_SIZE   As Long = &H40&
    
    Dim cnt         As Long
    Dim ret         As Long
    Dim buf()       As Byte
    Dim Offset      As Long
    Dim Process     As SYSTEM_PROCESS_INFORMATION
    Dim ProcName    As String
    Dim ProcPath    As String
    Dim sTime       As SYSTEMTIME
    Dim TimeZoneInfo(171)   As Byte
    
    GetTimeZoneInformation VarPtr(TimeZoneInfo(0))
    
    ReDim ProcList(200)
    
    SetCurrentProcessPrivileges "SeDebugPrivilege"
    
    If NtQuerySystemInformation(SystemProcessInformation, ByVal 0&, 0&, ret) = STATUS_INFO_LENGTH_MISMATCH Then
    
        ReDim buf(ret - 1)
        
        If NtQuerySystemInformation(SystemProcessInformation, buf(0), ret, ret) = STATUS_SUCCESS Then
        
            With Process
            
                Do
                    memcpy Process, buf(Offset), SPI_SIZE
                    
                    'ReDim .Threads(0 To .NumberOfThreads - 1)
                    
                    'For i = 0 To .NumberOfThreads - 1
                    '    memcpy .Threads(i), buf(Offset + SPI_SIZE + i * THREAD_SIZE), THREAD_SIZE
                    'Next
                    
                    If .ProcessID = 0 Then
                        ProcName = "System Idle Process"
                    ElseIf .ProcessID = 4 Then
                        ProcName = "System"
                    Else
                        ProcName = Space$(.ImageName.Length \ 2)
                        memcpy ByVal StrPtr(ProcName), ByVal .ImageName.Buffer, .ImageName.Length
                        ProcPath = GetFilePathByPID(.ProcessID)
                        
                        If Len(ProcPath) = 0 Then
                            ProcPath = FindOnPath(ProcName)
                        End If
                    End If
                    
                    If UBound(ProcList) < cnt Then ReDim Preserve ProcList(UBound(ProcList) + 100)
                    
                    With ProcList(cnt)
                        .Name = ProcName
                        .Path = ProcPath
                        .pid = Process.ProcessID
                        '.ParentPID = process.
                        .Priority = Process.BasePriority
                        .Threads = Process.NumberOfThreads
                        .SessionID = Process.SessionID
                        FileTimeToSystemTime VarPtr(Process.CreateTime), sTime
                        SystemTimeToTzSpecificLocalTime VarPtr(TimeZoneInfo(0)), sTime, sTime
                        SystemTimeToVariantTime sTime, .CreationTime
                    End With
                    
                    Offset = Offset + .NextEntryOffset
                    cnt = cnt + 1
                    
                Loop While .NextEntryOffset
                
            End With
            
        End If
        
    End If
    
    If cnt > 1 Then
        ReDim Preserve ProcList(cnt - 1)
    End If
    GetProcesses_Zw = cnt
    
    AppendErrorLogCustom "GetProcesses_Zw - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetProcesses_Zw"
    If inIDE Then Stop: Resume Next
End Function


Function GetFilePathByPID(pid As Long) As String
    On Error GoTo ErrorHandler:

    Const MAX_PATH_W                        As Long = 32767&
    Const PROCESS_VM_READ                   As Long = 16&
    Const PROCESS_QUERY_INFORMATION         As Long = 1024&
    Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
    
    Dim ProcPath    As String
    Dim hProc       As Long
    Dim cnt         As Long
    Dim pos         As Long
    Dim FullPath    As String

    hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0&, pid)
    
    If hProc = 0 Then
        If Err.LastDllError = ERROR_ACCESS_DENIED Then
            hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0&, pid)
        End If
    End If
    
    If hProc <> 0 Then
    
        If bIsWinVistaAndNewer Then
            cnt = MAX_PATH_W + 1
            ProcPath = Space$(cnt)
            Call QueryFullProcessImageName(hProc, 0&, StrPtr(ProcPath), VarPtr(cnt))
        End If
        
        If 0 <> Err.LastDllError Or Not bIsWinVistaAndNewer Then     'Win 2008 Server (x64) can cause Error 128 if path contains space characters
        
            ProcPath = Space$(MAX_PATH)
            cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
        
            If cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
                ProcPath = Space$(MAX_PATH_W)
                cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
            End If
        End If
        
        If cnt <> 0 Then                          'clear path
            ProcPath = Left$(ProcPath, cnt)
            ProcPath = PathNormalize(ProcPath)
        Else
            ProcPath = ""
        End If
        
        If ERROR_PARTIAL_COPY = Err.LastDllError Or cnt = 0 Then     'because GetModuleFileNameEx cannot access to that information for 64-bit processes on WOW64
            ProcPath = Space$(MAX_PATH)
            cnt = GetProcessImageFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
            
            If cnt <> 0 Then
                ProcPath = Left$(ProcPath, cnt)
                
                ' Convert DosDevice format to Disk drive format
                If StrComp(Left$(ProcPath, 8), "\Device\", 1) = 0 Then
                    pos = InStr(9, ProcPath, "\")
                    If pos <> 0 Then
                        FullPath = ConvertDosDeviceToDriveName(Left$(ProcPath, pos - 1))
                        If Len(FullPath) <> 0 Then
                            ProcPath = FullPath & Mid$(ProcPath, pos + 1)
                        End If
                    End If
                End If
            Else
                ProcPath = vbNullString
            End If
            
        End If
        
        If Len(ProcPath) <> 0 Then    'if process ran with 8.3 style, GetModuleFileNameEx will return 8.3 style on x64 and full pathname on x86
                                      'so wee need to expand it ourself
            
            'ProcPath = GetFullPath(ProcPath)
            ProcPath = GetLongPath(ProcPath)
            
            If InStr(ProcPath, "\") <> 0 Then
                GetFilePathByPID = ProcPath
            End If
        End If
        
        CloseHandle hProc
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePathByPID"
    If inIDE Then Stop: Resume Next
End Function

Public Function ConvertDosDeviceToDriveName(inDosDeviceName As String) As String
    On Error GoTo ErrorHandler:
    'alternatives:
    'RtlNtPathNameToDosPathName (XP+)
    'RtlVolumeDeviceToDosName
    'IOCTL_MOUNTMGR_QUERY_DOS_VOLUME_PATH
    
    Static DosDevices   As New Collection
    Static bInit       As Boolean
    
    If bInit Then
        If DosDevices.Count Then GoTo GetFromCollection
        Exit Function
    End If
    
    Dim aDrive()        As String
    Dim sDrives         As String
    Dim cnt             As Long
    Dim i               As Long
    Dim DosDeviceName   As String
    
    bInit = True
    
    cnt = GetLogicalDriveStrings(0&, StrPtr(sDrives))
    
    sDrives = Space$(cnt)
    
    cnt = GetLogicalDriveStrings(Len(sDrives), StrPtr(sDrives))
    
    If 0 = Err.LastDllError Then
    
        aDrive = Split(Left$(sDrives, cnt - 1), vbNullChar)
    
        For i = 0 To UBound(aDrive)
            
            DosDeviceName = Space$(MAX_PATH)
            
            cnt = QueryDosDevice(StrPtr(Left$(aDrive(i), 2)), StrPtr(DosDeviceName), Len(DosDeviceName))
            
            If cnt <> 0 Then
            
                DosDeviceName = Left$(DosDeviceName, InStr(DosDeviceName, vbNullChar) - 1)

                If Not isCollectionKeyExists(DosDeviceName, DosDevices) Then
                    DosDevices.Add aDrive(i), DosDeviceName
                End If

            End If
            
        Next
    
    End If

GetFromCollection:

    Dim pos As Long
    Dim sDrivePart As String
    Dim sOtherPart As String

    'Extract drive part
    If StrComp(Left$(inDosDeviceName, 8), "\Device\", 1) = 0 Then
        pos = InStr(9, inDosDeviceName, "\")
        If pos = 0 Then
            sDrivePart = inDosDeviceName
        Else
            sDrivePart = Left$(inDosDeviceName, pos - 1)
            sOtherPart = Mid$(inDosDeviceName, pos + 1)
        End If
        If isCollectionKeyExists(sDrivePart, DosDevices) Then
            ConvertDosDeviceToDriveName = BuildPath(DosDevices(sDrivePart), sOtherPart)
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ConvertDosDeviceToDriveName"
    If inIDE Then Stop: Resume Next
End Function

Public Function ProcessExist(NameOrPath As Variant, bGetNewListOfProcesses As Boolean) As Boolean
    Dim i As Long
    
    If AryPtr(gProcess) = 0 Or bGetNewListOfProcesses Then Call GetProcesses(gProcess)
    
    If InStr(NameOrPath, "\") <> 0 Then
        'by path
        For i = 0 To UBound(gProcess)
            If StrComp(NameOrPath, gProcess(i).Path, 1) = 0 Then ProcessExist = True: Exit For
        Next
    Else
        'by name
        For i = 0 To UBound(gProcess)
            If StrComp(NameOrPath, gProcess(i).Name, 1) = 0 Then ProcessExist = True: Exit For
        Next
    End If
End Function

'Public Sub RefreshDLLListNT(lPID&, objList As ListBox)
'    Dim arList() As String, i&
'    objList.Clear
'    GetDLLList lPID, arList()
'    If AryItems(arList) Then
'        For i = 0 To UBound(arList)
'            objList.AddItem arList(i)
'        Next
'    End If
'End Sub



' ---------------------------------------------------------------------------------------------------
' StartupList2 routine
' ---------------------------------------------------------------------------------------------------

Public Function GetRunningProcesses$()
    Dim aProcess() As MY_PROC_ENTRY
    Dim i&, sProc$
    If GetProcesses(aProcess) Then
        For i = 0 To UBound(aProcess)
            ' PID=Full Process Path|...
            With aProcess(i)
                If Not IsDefaultSystemProcess(.pid, .Name, .Path) Then
                    sProc = sProc & "|" & .pid & "=" & .Path
                End If
            End With
        Next
        GetRunningProcesses = Mid$(sProc, 2)
    End If
End Function

'For StartupList2
Public Function GetLoadedModules(lPID As Long, sProcess As String) As String
    '//TODO: sProcess
    
    Dim DLLList() As String
    
    DLLList = EnumModules64(lPID)

    If AryItems(DLLList) Then
        GetLoadedModules = Join(DLLList, "|")
    End If
End Function

Public Function GetPriorityProcess(Optional hProcess As Long, Optional lPID As Long) As PROCESS_PRIORITY
    If hProcess <> 0 Then
        GetPriorityProcess = GetPriorityClass(hProcess)
    ElseIf lPID <> 0 Then
        hProcess = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0&, lPID)
        If hProcess <> 0 Then
            GetPriorityProcess = GetPriorityClass(hProcess)
            CloseHandle hProcess
        End If
    End If
End Function

Public Function SetPriorityProcess(hProcess As Long, ePriorityProcess As PROCESS_PRIORITY) As Boolean
    'see table:   https://msdn.microsoft.com/ru-ru/library/windows/desktop/ms685100(v=vs.85).aspx
    'see remarks: https://msdn.microsoft.com/ru-ru/library/windows/desktop/ms686219(v=vs.85).aspx
    
    SetPriorityProcess = SetPriorityClass(hProcess, ePriorityProcess)
End Function

Public Function SetPriorityThread(hThread As Long, ePriorityThread As THREAD_PRIORITY, Optional bAllThreads As Boolean) As Boolean
    'see table:   https://msdn.microsoft.com/ru-ru/library/windows/desktop/ms685100(v=vs.85).aspx
    'see remarks: https://msdn.microsoft.com/ru-ru/library/windows/desktop/ms686277(v=vs.85).aspx
    
    SetPriorityThread = SetThreadPriority(hThread, ePriorityThread)
End Function

Public Function SetPriorityAllThreads(hProcess As Long, ePriorityThread As THREAD_PRIORITY) As Boolean
    Dim thrID() As Long
    Dim hThread As Long
    Dim i As Long

    thrID = GetProcessThreadIDs(hProcess)
    
    If AryItems(thrID) Then
        SetPriorityAllThreads = True
    
        For i = 0 To UBound(thrID)
            hThread = OpenThread(THREAD_SET_INFORMATION, False, thrID(i))
            
            If (hThread = 0) And (Err.LastDllError = ERROR_ACCESS_DENIED) Then
                hThread = OpenThread(THREAD_SET_LIMITED_INFORMATION, False, thrID(i))
            End If
            
            If (hThread <> 0) Then
                SetPriorityAllThreads = SetPriorityAllThreads And SetThreadPriority(hThread, ePriorityThread)
                CloseHandle hThread
            Else
                SetPriorityAllThreads = False
            End If
        Next
    End If
End Function

Public Function GetProcessThreadIDs(Optional hProcess As Long, Optional pid As Long) As Long()
    
    On Error GoTo ErrorHandler
    
    Dim hSnap As Long
    Dim te As THREADENTRY32
    Dim thrID() As Long
    Dim nThreads As Long
    
    ReDim thrID(0)
    
    If hProcess <> 0 Then
        If hProcess = GetCurrentProcess() Then
            pid = GetCurrentProcessId
        Else
            pid = GetProcessID(hProcess)
            If pid = 0 Then Exit Function
        End If
    Else
        If pid = 0 Then Exit Function
    End If
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0)
    
    If hSnap <> INVALID_HANDLE_VALUE Then
        te.dwSize = Len(te)
        
        If Thread32First(hSnap, te) Then
            Do
                If te.th32ProcessID = pid Then
                    ReDim Preserve thrID(nThreads)
                    thrID(nThreads) = te.th32ThreadID
                    nThreads = nThreads + 1
                End If
                
                If te.dwSize <> Len(te) Then te.dwSize = Len(te)
                
            Loop While Thread32Next(hSnap, te)
        End If
        CloseHandle hSnap
    End If
    
    If nThreads > 0 Then
        GetProcessThreadIDs = thrID
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "GetProcessThreadIDs", "Error = " & Err.Number, "LastDllError = " & Err.LastDllError
    If inIDE Then Stop: Resume Next
End Function

Public Function IsSystemCriticalProcessPath(sPath As String) As Boolean
    Static dPath As clsTrickHashTable
    
    If dPath Is Nothing Then
        Set dPath = New clsTrickHashTable
        dPath.CompareMode = TextCompare
        
        dPath.Add sWinSysDir & "\smss.exe", 0
        dPath.Add sWinSysDir & "\csrss.exe", 0
        dPath.Add sWinSysDir & "\svchost.exe", 0
        dPath.Add sWinSysDir & "\winlogon.exe", 0
        dPath.Add sWinSysDir & "\wininit.exe", 0
        dPath.Add sWinSysDir & "\lsm.exe", 0
        dPath.Add sWinSysDir & "\services.exe", 0
        dPath.Add sWinSysDir & "\lsass.exe", 0
        dPath.Add sWinSysDir & "\msdtc.exe", 0 'database / file / message queue transactions
    End If
    
    IsSystemCriticalProcessPath = dPath.Exists(sPath)
End Function

Public Sub SystemPriorityDowngrade(bState As Boolean)
    On Error GoTo ErrorHandler
    
    Static dPath As clsTrickHashTable
    Static dPrior As clsTrickHashTable
    Dim i As Long
    
    If dPath Is Nothing Then
        If bState = False Then
            MsgBoxW ("Invalid using of SystemPriorityDowngrade")
            Exit Sub
        End If
        
        Set dPath = New clsTrickHashTable
        Set dPrior = New clsTrickHashTable
        
        dPath.CompareMode = TextCompare
        dPrior.CompareMode = TextCompare
        
        'Critical processes and processes that are important for normal operation of own software
        dPath.Add sWinSysDir & "\alg.exe", 0
        dPath.Add sWinSysDir & "\smss.exe", 0
        dPath.Add sWinSysDir & "\csrss.exe", 0
        dPath.Add sWinSysDir & "\ctfmon.exe", 0
        dPath.Add sWinSysDir & "\lsass.exe", 0
        dPath.Add sWinSysDir & "\msdtc.exe", 0
        dPath.Add sWinSysDir & "\services.exe", 0
        dPath.Add sWinSysDir & "\svchost.exe", 0
        dPath.Add sWinSysDir & "\winlogon.exe", 0
        dPath.Add sWinSysDir & "\wininit.exe", 0
        dPath.Add sWinSysDir & "\lsm.exe", 0
        'dllhost.exe ?
        
        'Access denied
        dPath.Add sWinSysDir & "\audiodg.exe", 0
        dPath.Add sWinSysDir & "\SecurityHealthService.exe", 0
        'SearchFilter
        'SearchProtocolHost
        'RuntimeBroker
        'dllHost
    End If
    
    If bState = True Then 'do downgrade
    
        Dim hProc As Long
        Dim dwSelfPID As Long
        Dim Priority As PROCESS_PRIORITY
        
        dwSelfPID = GetCurrentProcessId()
        
        If Not bAutoLogSilent Then
            Call GetProcesses(gProcess)
        Else
            If AryPtr(gProcess) = 0 Then
                Call GetProcesses(gProcess)
            End If
        End If
        
        If AryPtr(gProcess) <> 0 Then
            
            For i = 0 To UBound(gProcess)
                
                If Not dPath.Exists(gProcess(i).Path) Then
                
                    If StrComp(gProcess(i).Name, "avz.exe", 1) <> 0 And _
                        StrComp(gProcess(i).Name, "avz5.exe", 1) <> 0 And _
                        gProcess(i).pid <> dwSelfPID And _
                        gProcess(i).pid <> 0 And _
                        InStr(1, gProcess(i).Path, "Windows Defender", 1) = 0 And _
                        Not IsDefaultSystemProcess(gProcess(i).pid, gProcess(i).Name, gProcess(i).Path) _
                        Then
                        
                        hProc = OpenProcess(PROCESS_SET_INFORMATION, 0, gProcess(i).pid)
                        
                        If hProc <> 0 Then
                            
                            Priority = GetPriorityProcess(, gProcess(i).pid)
                            
                            '//TODO:
                            'try also low mem. priority?
                            'https://docs.microsoft.com/en-us/windows/desktop/api/processthreadsapi/ns-processthreadsapi-_memory_priority_information
                            
                            'downgrade
                            If SetPriorityProcess(hProc, IDLE_PRIORITY_CLASS) Then
                                'save old state
                                dPrior.Add hProc, CLng(Priority)
                            Else
                                Debug.Print "Can't set priority: " & gProcess(i).Path
                            
                                'on failure
                                CloseHandle hProc
                            End If
                        Else
                            Debug.Print "Can't open: " & gProcess(i).Path & " (PID = " & gProcess(i).pid & ")"
                        End If
                        
                    End If
                
                End If
                
            Next
        End If
        
    Else 'revert changes
        
        For i = 0 To dPrior.Count - 1
            
            hProc = dPrior.Keys(i)
            Priority = dPrior.Items(i)
            
            If Priority = 0 Then
                Priority = NORMAL_PRIORITY_CLASS
            End If
            
            If Not SetPriorityProcess(hProc, Priority) Then
                Debug.Print "Can't restore priority: " & gProcess(i).Path
            End If
            
            CloseHandle hProc
        Next
        
        Set dPath = Nothing
        Set dPrior = Nothing
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "GetProcessThreadIDs", "Error = " & Err.Number, "LastDllError = " & Err.LastDllError
    If inIDE Then Stop: Resume Next
End Sub

Public Function EnumModules64(pid As Long) As String()
    On Error GoTo ErrorHandler:
    
    'Const PROCESS_ALL_ACCESS As Long = &H1FFFFF
    
    Dim hProc As Long
    Dim ID As Long
    Dim Wow64Process As Long
    Dim PEB As PEB64
    Dim PEB_LDR_DATA As PEB_LDR_DATA64
    Dim ldr_table_entry As LDR_DATA_TABLE_ENTRY64
    Dim buf() As Byte
    Dim cchBuf As Long
    Dim address As Currency
    Dim pFirst_ldr_module As Currency
    Dim sMainModule As String
    Dim i As Long
    
    Static bInit As Boolean
    Static bFuncWow64ProcessExist As Boolean
    Static bFuncNtWow64QueryInformationProcess64Exist As Boolean
    Static bFuncNtWow64ReadVirtualMemory64Exist As Boolean
    
    If pid = 0 Then Exit Function
    
    If Not OSver.IsWin64 Then
        EnumModules64 = EnumModules32(pid)
        Exit Function
    End If
    
    If Not bInit Then
        bInit = True
        bFuncWow64ProcessExist = IsProcedureAvail("IsWow64Process", "kernel32.dll")
        bFuncNtWow64QueryInformationProcess64Exist = IsProcedureAvail("NtWow64QueryInformationProcess64", "ntdll.dll")
        bFuncNtWow64ReadVirtualMemory64Exist = IsProcedureAvail("NtWow64ReadVirtualMemory64", "ntdll.dll")
    End If
    
    If Not bFuncWow64ProcessExist Then
        EnumModules64 = EnumModules32(pid)
        Exit Function
    End If
    
    hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0&, pid)
    
    If hProc <> 0 Then
        If IsWow64Process(hProc, Wow64Process) Then
            If 0 <> Wow64Process Then
                EnumModules64 = EnumModules32(pid)
                Exit Function
            Else
                
                If Not bFuncNtWow64QueryInformationProcess64Exist Then Exit Function
                If Not bFuncNtWow64ReadVirtualMemory64Exist Then Exit Function
                
                sMainModule = GetFilePathByPID(pid)
                
                Dim aModule() As String
                ReDim aModule(100)
                
                If read_peb_data(hProc, PEB) Then
                
                    ReDim buf(LenB(PEB_LDR_DATA) - 1)
                    
                    If read_mem64(hProc, PEB.LoaderData, LenB(PEB_LDR_DATA), buf) Then
                    
                        memcpy PEB_LDR_DATA, buf(0), LenB(PEB_LDR_DATA)
                        
                        address = PEB_LDR_DATA.InLoadOrderModuleList.Flink
                        
                        pFirst_ldr_module = PEB.LoaderData + cMath.IntToInt64(VarPtr(PEB_LDR_DATA.InLoadOrderModuleList) - VarPtr(PEB_LDR_DATA.Length))
                        
                        Do
                            ReDim buf(LenB(ldr_table_entry) - 1)
                        
                            If read_mem64(hProc, address, LenB(ldr_table_entry), buf) Then
                            
                                memcpy ldr_table_entry, buf(0), LenB(ldr_table_entry)
                                
                                cchBuf = ldr_table_entry.FullDllName.MaximumLength
                                
                                If cchBuf > 0 And ldr_table_entry.FullDllName.Buffer <> 0 Then
                                    ReDim buf(cchBuf - 1)
                                
                                    If read_mem64(hProc, ldr_table_entry.FullDllName.Buffer, cchBuf, buf) Then
                                        
                                        aModule(i) = PathNormalize(StringFromPtrW(VarPtr(buf(0))))
                                        
                                        If StrComp(aModule(i), sMainModule, 1) <> 0 Then
                                            i = i + 1
                                            If i > UBound(aModule) Then ReDim Preserve aModule(UBound(aModule) + 100)
                                        End If
                                    End If
                                Else
                                    Exit Do
                                End If
                                
                                address = ldr_table_entry.InLoadOrderModuleList.Flink
                            Else
                                Exit Do
                            End If
                        Loop Until address = pFirst_ldr_module
                    End If
                End If
            End If
        End If
        CloseHandle hProc
    End If

    If i > 0 Then
        ReDim Preserve aModule(i - 1)
        EnumModules64 = aModule
    End If

    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumModules64"
    If inIDE Then Stop: Resume Next
End Function

Private Function read_mem64(Handle As Long, address As Currency, Length As Long, buf() As Byte) As Boolean
    
    Dim cbret As Currency
    ReDim buf(Length - 1) As Byte
    Dim HRes As Long
    
    HRes = NtWow64ReadVirtualMemory64(Handle, address, VarPtr(buf(0)), cMath.IntToInt64(Length), VarPtr(cbret))
    
    If NT_SUCCESS(HRes) Then
        read_mem64 = True
    Else
        Debug.Print "NtWow64ReadVirtualMemory64 failed with code: 0x" & Hex(HRes)
    End If
End Function

Private Function read_pbi(Handle As Long, PBI As PROCESS_BASIC_INFORMATION64) As Boolean
    Dim cbret As Long
    Dim HRes As Long
    
    HRes = NtWow64QueryInformationProcess64(Handle, ProcessBasicInformation, VarPtr(PBI), LenB(PBI), VarPtr(cbret))
    
    If NT_SUCCESS(HRes) Then
        read_pbi = True
    Else
        Debug.Print "NtWow64QueryInformationProcess64 failed with code: 0x" & Hex(HRes)
    End If
End Function

Private Function read_peb_data(Handle As Long, PEB As PEB64) As Boolean

    Dim PBI As PROCESS_BASIC_INFORMATION64
    ReDim buf(LenB(PEB) - 1) As Byte

    If read_pbi(Handle, PBI) Then
        If read_mem64(Handle, PBI.PebBaseAddress, LenB(PEB), buf) Then
            memcpy PEB, buf(0), LenB(PEB)
            read_peb_data = True
        End If
    End If
End Function

Public Function EnumModules32(pid As Long) As String()
    On Error GoTo ErrorHandler:
    
    Dim hSnap&, uME32 As MODULEENTRY32W
    Dim sDllFile$
    Dim aModule() As String
    Dim hModules(1 To 1024) As Long
    Dim i&, j&, cnt&, lNeeded&, hProc&
    
    ReDim aModule(50)
    
    If OSver.MajorMinor < 5 Then

        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, pid)
        
        If hSnap <> INVALID_HANDLE_VALUE Then
            uME32.dwSize = Len(uME32)
            If Module32First(hSnap, uME32) = 0 Then
                CloseHandle hSnap
                Exit Function
            End If
        
            Do
                aModule(i) = StringFromPtrW(VarPtr(uME32.szExePath(0)))
                i = i + 1
                If i > UBound(aModule) Then ReDim Preserve aModule(UBound(aModule) + 50)
            Loop Until Module32Next(hSnap, uME32) = 0
            CloseHandle hSnap
        End If
    
    Else
        hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0, pid)
    
        If hProc <> 0 Then
            lNeeded = 0
            If EnumProcessModules(hProc, hModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                For j = 2 To 1024
                    If hModules(j) = 0 Then Exit For
                    
                    aModule(i) = String$(MAX_PATH, 0)
                    
                    cnt = GetModuleFileNameEx(hProc, hModules(j), StrPtr(aModule(i)), Len(aModule(i)))
            
                    If cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
                        aModule(i) = String$(MAX_PATH_W, 0)
                        cnt = GetModuleFileNameEx(hProc, hModules(j), StrPtr(aModule(i)), Len(aModule(i)))
                    End If
                    If cnt <> 0 Then
                        aModule(i) = PathNormalize(Left$(aModule(i), cnt))
                    End If
                    
                    i = i + 1
                    If i > UBound(aModule) Then ReDim Preserve aModule(UBound(aModule) + 50)
                Next j
            End If
            CloseHandle hProc
        End If
        
    End If
    
    If i > 0 Then
        ReDim Preserve aModule(i - 1)
        EnumModules32 = aModule
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumModules32"
    If inIDE Then Stop: Resume Next
End Function

Public Function SetProcessIOPriority(lPID As Long, dwPriority As Long) 'required SeIncreaseBasePriorityPrivilege 'XP SP3+
    Dim hProc&, lret&, bRequirement As Boolean
    If lPID = 0 Or lPID = 4 Then Exit Function
    
    If OSver.IsWindowsXPOrGreater Then
        bRequirement = True
        If OSver.MajorMinor = 5.1 And OSver.SPVer < 3 Then bRequirement = False
    End If
    
    If bRequirement Then
        hProc = OpenProcess(PROCESS_SET_INFORMATION, 0, lPID)
        If hProc <> 0 Then
            lret = NtSetInformationProcess(hProc, ProcessIoPriority, VarPtr(dwPriority), 4&)
            If 0 = lret Then
                SetProcessIOPriority = True
            Else
                Debug.Print "Failed in NtSetInformationProcess (ProcessIoPriority) with error = 0x" & Hex(lret)
            End If
            CloseHandle hProc
        Else
            Debug.Print "Failed in OpenProcess with error = 0x" & Hex(Err.LastDllError) & ", PID = " & lPID
        End If
    End If
End Function

Public Function GetProcessPagePriority(lPID As Long) As Long
    Dim hProc&, lret&, dwPrio&
    If lPID = 0 Or lPID = 4 Then Exit Function
    
    hProc = OpenProcess(IIf(OSver.IsWindowsVistaOrGreater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0, lPID)
    If hProc <> 0 Then
        lret = NtQueryInformationProcess(hProc, ProcessPagePriority, VarPtr(dwPrio), 4&, 0&)
        If 0 = lret Then
            GetProcessPagePriority = dwPrio
        Else
            Debug.Print "Failed in NtQueryInformationProcess (ProcessPagePriority) with error = 0x" & Hex(lret) & ", PID = " & lPID
        End If
        CloseHandle hProc
    End If
End Function

Public Function GetParentPID(lPID As Long) As Long
    Dim hProc&, lret&
    If lPID = 0 Or lPID = 4 Then Exit Function
    Dim PBI As PROCESS_BASIC_INFORMATION
    
    hProc = OpenProcess(IIf(OSver.IsWindowsVistaOrGreater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0, lPID)
    If hProc <> 0 Then
        lret = NtQueryInformationProcess(hProc, ProcessBasicInformation, VarPtr(PBI), LenB(PBI), 0&)
        If 0 = lret Then
            GetParentPID = PBI.InheritedFromUniqueProcessId
        Else
            Debug.Print "Failed in NtQueryInformationProcess (ProcessBasicInformation) with error = 0x" & Hex(lret) & ", PID = " & lPID
        End If
        CloseHandle hProc
    End If
End Function

Public Function IsDefaultSystemProcess(pid As Long, sName As String, sPath As String) As Boolean
    If Len(sPath) = 0 Then
        If IsSystemIdleProcess(pid, sName, sPath) Then IsDefaultSystemProcess = True: Exit Function
        If IsSystemProcess(pid, sName, sPath) Then IsDefaultSystemProcess = True: Exit Function
        If IsMinimalProcess(pid, sName, sPath) Then IsDefaultSystemProcess = True: Exit Function
    End If
End Function

Public Function IsSystemIdleProcess(pid As Long, sName As String, sPath As String) As Boolean
    If Len(sPath) = 0 Then
        If sName = "System Idle Process" Then
            If pid = 0 Then IsSystemIdleProcess = True
        End If
    End If
End Function

Public Function IsSystemProcess(pid As Long, sName As String, sPath As String) As Boolean
    If Len(sPath) = 0 Then
        If sName = "System" Then
            If pid = 4 Then IsSystemProcess = True
        End If
    End If
End Function

Public Function IsMinimalProcess(pid As Long, sName As String, sPath As String) As Boolean
    Dim bComply As Boolean
    
    If Len(sPath) = 0 Then
        bComply = True
        
        Select Case sName
        Case "Registry"
        Case "Memory Compression"
        Case "MemCompression"
        Case "Secure System"
        Case Else: bComply = False
        End Select
        
        If bComply Then
            If GetParentPID(pid) = 4 Then IsMinimalProcess = True
        End If
    End If
End Function

Public Function GetProcessCriticalFlag(lPID&, l_OutFlag As Long) As Boolean
    Dim hProc&, Flag&, lret&
    If lPID = 0 Or lPID = 4 Then Exit Function
    
    'hProc = OpenProcess(IIf(OSver.IsWindowsVistaOrGreater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0, lPid)
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION, 0, lPID)
    If hProc <> 0 Then
        lret = NtQueryInformationProcess(hProc, ProcessBreakOnTermination, VarPtr(Flag), 4&, 0&)
        If 0 = lret Then
            l_OutFlag = Flag
            GetProcessCriticalFlag = True
        Else
            Debug.Print "Failed in NtQueryInformationProcess (ProcessBreakOnTermination) with error = 0x" & Hex(lret) & ", PID = " & lPID
        End If
        CloseHandle hProc
    Else
        Debug.Print "Failed in OpenProcess with error = 0x" & Hex(Err.LastDllError) & ", PID = " & lPID
    End If
End Function

Public Function SetProcessCriticalFlag(lPID&, bEnable As Boolean) As Boolean 'required SeDebugPrivilege
    Dim hProc&, lret&
    If lPID = 0 Or lPID = 4 Then Exit Function
    
    hProc = OpenProcess(PROCESS_SET_INFORMATION, 0, lPID)
    If hProc <> 0 Then
        lret = NtSetInformationProcess(hProc, ProcessBreakOnTermination, IIf(bEnable, VarPtr(1&), VarPtr(0&)), 4&)
        If 0 = lret Then
            SetProcessCriticalFlag = True
        Else
            Debug.Print "Failed in NtSetInformationProcess (ProcessBreakOnTermination) with error = 0x" & Hex(lret)
        End If
        CloseHandle hProc
    Else
        Debug.Print "Failed in OpenProcess with error = 0x" & Hex(Err.LastDllError) & ", PID = " & lPID
    End If
End Function

Public Sub KillOtherHJTInstances(Optional sAdditionalDir As String)
    Dim lNumProcesses As Long
    Dim i As Long
    Dim Process() As MY_PROC_ENTRY
    
    lNumProcesses = GetProcesses(Process)
    
    For i = 0 To lNumProcesses - 1
        If StrComp(Process(i).Path, AppPath(True), 1) = 0 Then
            If Process(i).pid <> GetCurrentProcessId() Then
                KillProcess Process(i).pid
            End If
        End If
        If Len(sAdditionalDir) <> 0 Then
            If StrComp(Process(i).Path, sAdditionalDir, 1) = 0 Then
                If Process(i).pid <> GetCurrentProcessId() Then
                    KillProcess Process(i).pid
                End If
            End If
        End If
    Next
End Sub

Public Function IsMicrosoftHostFile(sPath As String) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "IsMicrosoftHostFile - Begin"
    
    Dim sName As String
    Dim i As Long
    sName = GetFileName(sPath, True)
    
    Dim aHost(9) As String
    aHost(0) = "cmd.exe"
    aHost(1) = "rundll32.exe"
    aHost(2) = "wmic.exe"
    aHost(3) = "cscript.exe"
    aHost(4) = "wscript.exe"
    aHost(5) = "powershell.exe"
    aHost(6) = "sc.exe"
    aHost(7) = "mshta.exe"
    aHost(8) = "pcalua.exe"
    aHost(9) = "msiexec.exe"
    
    For i = 0 To UBound(aHost)
        If StrComp(sName, aHost(i), 1) = 0 Then
            IsMicrosoftHostFile = True
            Exit Function
        End If
    Next
    
    AppendErrorLogCustom "IsMicrosoftHostFile - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IsMicrosoftHostFile"
    If inIDE Then Stop: Resume Next
End Function

Public Sub KillMicrosoftHostProcesses()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "FreezeCustomProcesses - Begin"
    
    Dim aProcess() As MY_PROC_ENTRY
    Dim i&, sProc$
    If GetProcesses(aProcess) Then
        For i = 0 To UBound(aProcess)
            With aProcess(i)
                If Not IsDefaultSystemProcess(.pid, .Name, .Path) Then
                    If IsMicrosoftHostFile(.Path) Then
                        KillProcess .pid
                    End If
                End If
            End With
        Next
    End If
    
    AppendErrorLogCustom "FreezeCustomProcesses - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FreezeCustomProcesses"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FreezeCustomProcesses()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "FreezeCustomProcesses - Begin"
    
    Dim aProcess() As MY_PROC_ENTRY
    Dim i&, sProc$
    If GetProcesses(aProcess) Then
        For i = 0 To UBound(aProcess)
            With aProcess(i)
                If Not IsDefaultSystemProcess(.pid, .Name, .Path) Then
                    If Not IsMicrosoftFile(.Path) Then
                        PauseProcess .pid
                    ElseIf IsMicrosoftHostFile(.Path) Then
                        PauseProcess .pid
                    End If
                End If
            End With
        Next
    End If
    AppendErrorLogCustom "FreezeCustomProcesses - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FreezeCustomProcesses"
    If inIDE Then Stop: Resume Next
End Sub
