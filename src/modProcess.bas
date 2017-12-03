Attribute VB_Name = "modProcess"
'
' Windows Process module by Alex Dragokas
'

Option Explicit

Public Type MY_PROC_ENTRY
    Name        As String
    Path        As String
    PID         As Long
    Threads     As Long
    Priority    As Long
    SessionID   As Long
End Type

Public Enum PROCESS_PRIORITY
    ABOVE_NORMAL_PRIORITY_CLASS = &H8000&
    BELOW_NORMAL_PRIORITY_CLASS = &H4000&
    HIGH_PRIORITY_CLASS = &H80&
    IDLE_PRIORITY_CLASS = &H40&
    NORMAL_PRIORITY_CLASS = &H20&
    PROCESS_MODE_BACKGROUND_BEGIN = &H100000
    PROCESS_MODE_BACKGROUND_END = &H200000
    REALTIME_PRIORITY_CLASS = &H100&
End Enum

Public Enum THREAD_PRIORITY
    THREAD_MODE_BACKGROUND_BEGIN = &H10000
    THREAD_MODE_BACKGROUND_END = &H20000
    THREAD_PRIORITY_ABOVE_NORMAL = 1&
    THREAD_PRIORITY_BELOW_NORMAL = -1&
    THREAD_PRIORITY_HIGHEST = 2&
    THREAD_PRIORITY_IDLE = -15&
    THREAD_PRIORITY_LOWEST = -2&
    THREAD_PRIORITY_NORMAL = 0&
    THREAD_PRIORITY_TIME_CRITICAL = 15&
End Enum

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
'Private Declare Function NtQuerySystemInformation Lib "NTDLL.DLL" (ByVal infoClass As Long, Buffer As Any, ByVal BufferSize As Long, ret As Long) As Long
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
'Private Declare Function NtSuspendProcess Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
'Private Declare Function NtResumeProcess Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
'Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
'Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
'Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
'Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
'Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
'Private Declare Function ZwSetInformationProcess Lib "NTDLL.DLL" (ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
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

Public Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function SetProcessPriorityBoost Lib "kernel32.dll" (ByVal hProcess As Long, ByVal DisablePriorityBoost As Long) As Long
Public Declare Function GetProcessPriorityBoost Lib "kernel32.dll" (ByVal hThread As Long, pDisablePriorityBoost As Long) As Long
Public Declare Function SetThreadPriorityBoost Lib "kernel32.dll" (ByVal hThread As Long, ByVal DisablePriorityBoost As Long) As Long
Public Declare Function GetThreadPriorityBoost Lib "kernel32.dll" (ByVal hThread As Long, pDisablePriorityBoost As Long) As Long
Public Declare Function GetProcessID Lib "kernel32.dll" (ByVal Process As Long) As Long


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


Public Function KillProcess(lPID&) As Boolean
    Dim hProcess&
    If lPID = 0 Then Exit Function
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProcess = 0 Then
        'The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows.
        MsgBoxW Translate(1652), vbCritical
    Else
        If TerminateProcess(hProcess, 0) = 0 Then
            'The selected process could not be killed." & _
                   " It may be protected by Windows.
            MsgBoxW Translate(1653), vbCritical
        Else
            DoEvents
            KillProcess = True
        End If
        CloseHandle hProcess
    End If
End Function

Public Function KillProcessNT(lPID&) As Boolean
    Dim hProc&
    If lPID = 0 Then Exit Function
    
    hProc = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProc <> 0 Then
        If TerminateProcess(hProc, 0) = 0 Then
        'The selected process could not be killed." & _
                   " It may be protected by Windows.
            MsgBoxW Translate(1653), vbCritical
        Else
            DoEvents
            KillProcessNT = True
        End If
        CloseHandle hProc
    Else
        'The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows." & vbCrLf & vbCrLf & _
               "This process might be a service, which you can " & _
               "stop from the Services applet in Admin Tools." & vbCrLf & _
               "(To load this window, click Start, Run and enter 'services.msc')
        MsgBoxW Translate(1654), vbCritical
    End If
End Function

Public Function PauseProcess(lPID As Long) As Boolean
    On Error GoTo ErrorHandler:

    '//TODO: add process status checking after Nt

    Dim hSnap&, uTE32 As THREADENTRY32, hThread&, IsFailed As Boolean, hProc&
    
    If Not bIsWinNT And Not bIsWinME Then Exit Function
    If lPID = 0 Or lPID = GetCurrentProcessId Then Exit Function
    
    If IsProcedureAvail("NtSuspendProcess", "ntdll.dll") Then
        hProc = OpenProcess(PROCESS_SUSPEND_RESUME, 0, lPID)
    
        If hProc <> 0 Then
            Call NtSuspendProcess(hProc)
            CloseHandle hProc
        End If
    End If
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lPID)
    If hSnap = -1 Then Exit Function
    
    uTE32.dwSize = Len(uTE32)
    If Thread32First(hSnap, uTE32) = 0 Then
        CloseHandle hSnap
        Exit Function
    End If
    
    Do
        If uTE32.th32ProcessID = lPID Then
            PauseProcess = True
            
            hThread = OpenThread(THREAD_SUSPEND_RESUME, False, uTE32.th32ThreadID)
            If hThread = 0 Then
                IsFailed = True
            Else
                If SuspendThread(hThread) = -1 Then IsFailed = True
                CloseHandle hThread
            End If
        End If
    Loop Until Thread32Next(hSnap, uTE32) = 0
    CloseHandle hSnap
    
    If IsFailed Then PauseProcess = False
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "PauseProcess. PID:", lPID
    If inIDE Then Stop: Resume Next
End Function

Public Function ResumeProcess(lPID As Long) As Boolean
    On Error GoTo ErrorHandler:

    '//TODO: add process status checking after Nt

    Dim hSnap&, uTE32 As THREADENTRY32, hThread&, IsFailed As Boolean, hProc&
    
    If Not (bIsWinNT Or bIsWinME) Then Exit Function
    If lPID = 0 Or lPID = GetCurrentProcessId Then Exit Function
    
    If IsProcedureAvail("NtResumeProcess", "ntdll.dll") Then
        hProc = OpenProcess(PROCESS_SUSPEND_RESUME, 0, lPID)
        
        If hProc <> 0 Then
            Call NtResumeProcess(hProc)
            CloseHandle hProc
        End If
    End If
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lPID)
    If hSnap = -1 Then Exit Function
    
    uTE32.dwSize = Len(uTE32)
    If Thread32First(hSnap, uTE32) = 0 Then
        CloseHandle hSnap
        Exit Function
    End If
    
    Do
        If uTE32.th32ProcessID = lPID Then
            ResumeProcess = True
            
            hThread = OpenThread(THREAD_SUSPEND_RESUME, False, uTE32.th32ThreadID)
            If hThread = 0 Then
                IsFailed = True
            Else
                If ResumeThread(hThread) = -1 Then IsFailed = True
                CloseHandle hThread
            End If
        End If
    Loop Until Thread32Next(hSnap, uTE32) = 0
    CloseHandle hSnap
    
    If IsFailed Then ResumeProcess = False
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ResumeProcess. PID:", lPID
    If inIDE Then Stop: Resume Next
End Function

Public Function KillProcessByFile(sPath$, Optional bForceMicrosoft As Boolean) As Boolean
    Dim hProcess&, i&, sTaskKill As String
    'Note: this sub is silent - it displays no errors !
    
    If sPath = vbNullString Then Exit Function
    
    sPath = FindOnPath(sPath, True)
    
    If Not bForceMicrosoft Then
        If IsMicrosoftFile(sPath) Then Exit Function
    End If
    
    If bIsWinNT Then
        KillProcessByFile = KillProcessNTByFile(sPath)
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
                PauseProcess Process(i).PID
                hProcess = OpenProcess(PROCESS_TERMINATE, 0, Process(i).PID)
                If hProcess <> 0 Then
                    If TerminateProcess(hProcess, 0) <> 0 Then
                        'Success
                        DoEvents
                        KillProcessByFile = True
                    End If
                    CloseHandle hProcess
                End If
                If KillProcessByFile = False Then
                    If FileExists(sTaskKill) Then
                        Proc.ProcessRun sTaskKill, "/F /PID " & Process(i).PID, , 0
                        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
                            Proc.ProcessClose , , True
                        End If
                    End If
                End If
            End If
        Next
    End If
End Function

Public Function PauseProcessByFile(sPath$) As Boolean
    Dim i&
    'Note: this sub is silent - it displays no errors !
    If sPath = vbNullString Then Exit Function
    If bIsWinNT Then
        KillProcessNTByFile sPath
        Exit Function
    End If
    
    Dim lNumProcesses As Long
    Dim Process() As MY_PROC_ENTRY
    
    lNumProcesses = GetProcesses(Process)
        
    If lNumProcesses Then
        
        For i = 0 To UBound(Process)
        
            If StrComp(sPath, Process(i).Path, 1) = 0 Then
            
                PauseProcessByFile = PauseProcess(Process(i).PID)
            End If
        Next
    End If
End Function

Public Function KillProcessNTByFile(sPath$) As Boolean
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
                            KillProcessNTByFile = True
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
                    ProcList(cnt).PID = uProcess.th32ProcessID
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
                        .PID = Process.ProcessID
                        .Priority = Process.BasePriority
                        .Threads = Process.NumberOfThreads
                        .SessionID = Process.SessionID
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


Function GetFilePathByPID(PID As Long) As String
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

    hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0&, PID)
    
    If hProc = 0 Then
        If Err.LastDllError = ERROR_ACCESS_DENIED Then
            hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0&, PID)
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
            If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = sWinDir & Mid$(ProcPath, 12)
            If "\??\" = Left$(ProcPath, 4) Then ProcPath = Mid$(ProcPath, 5)
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
                ProcPath = ""
            End If
            
        End If
        
        If Len(ProcPath) <> 0 Then    'if process ran with 8.3 style, GetModuleFileNameEx will return 8.3 style on x64 and full pathname on x86
                                      'so wee need to expand it ourself
            
            ProcPath = GetFullPath(ProcPath)
            GetFilePathByPID = GetLongPath(ProcPath)
        End If
        
        CloseHandle hProc
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePathByPID"
    If inIDE Then Stop: Resume Next
End Function

Public Function ConvertDosDeviceToDriveName(inDosDeviceName As String) As String
    '//TODO: fix #5 (Invalid procedure call or argument)

    On Error GoTo ErrorHandler:

    Static DosDevices   As New Collection
    
    If DosDevices.Count Then
        ConvertDosDeviceToDriveName = DosDevices(inDosDeviceName)
        Exit Function
    End If
    
    Dim aDrive()        As String
    Dim sDrives         As String
    Dim cnt             As Long
    Dim i               As Long
    Dim DosDeviceName   As String
    
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

                DosDevices.Add aDrive(i), DosDeviceName

            End If
            
        Next
    
    End If
    
    ConvertDosDeviceToDriveName = DosDevices(inDosDeviceName)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ConvertDosDeviceToDriveName"
    If inIDE Then Stop: Resume Next
End Function

Public Function ProcessExist(NameOrPath As String, bDoRefresh As Boolean) As Boolean
    Dim i As Long
    
    If AryPtr(gProcess) = 0 Or bDoRefresh Then Call GetProcesses(gProcess)
    
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

Public Sub GetDLLList(lPID As Long, arList() As String)
    On Error GoTo ErrorHandler:

    Erase arList

    Dim lNeeded&
    Dim hProc&, lModules&(1 To 1024)
    Dim sModuleName$, j&, cnt&, myDLLs() As String
    
    '//TODO: add support for x64 processes:
    'replace by WMI CIM_ProcessExecutable: http://2011sg.poshcode.org/94.html
    'or Win32_ModuleLoadTrace
    'or RTL_PROCESS_MODULE_INFORMATION: http://www.rohitab.com/discuss/topic/40696-list-loaded-drivers-with-ntquerysysteminformation/
    'https://doxygen.reactos.org/d7/d55/ldrapi_8c_source.html#l00972
    
    ReDim myDLLs(1024): cnt = 0

'    If OSver.MajorMinor >= 6 Then 'Vista+
'        If GetServiceRunState("winmgmt") = SERVICE_RUNNING Then
'
'            Dim oWMI As Object, colMod As Object, oMod As Object
'
'            Set oWMI = CreateObject("winmgmts:{impersonationLevel=Impersonate}!\\.\root\cimv2")
'            Set colMod = oWMI.ExecQuery("Select * from CIM_ProcessExecutable")
'
'            For Each oMod In colMod
'                With oMod.Dependent ' CIM_Process
'                    Debug.Print .Handle
'                    Debug.Print .Caption
'                End With
'                With oMod.Antecedent ' CIM_DataFile
'                    Debug.Print .Path
'                End With
'            Next
'
'            Exit Function
'        End If
'    End If
    
    hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0, lPID)
    If hProc <> 0 Then
        lNeeded = 0
        If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
            For j = 2 To 1024
                If lModules(j) = 0 Then Exit For
                sModuleName = String$(MAX_PATH, 0)
                GetModuleFileNameExA hProc, lModules(j), sModuleName, Len(sModuleName)
                sModuleName = TrimNull(sModuleName)
                If sModuleName <> vbNullString And _
                   sModuleName <> "?" Then
                    myDLLs(cnt) = sModuleName
                    cnt = cnt + 1
                End If
            Next j
        End If
        CloseHandle hProc
    End If
    
'    hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lPID)
'        If hProc > 0 Then
'            lNeeded = 0
'            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) > 0 Then
'                For i = 2 To 1024
'                    If lModules(i) = 0 Then Exit For
'                    sModuleName = String$(MAX_PATH, 0)
'                    GetModuleFileNameExA hProc, lModules(i), sModuleName, Len(sModuleName)
'                    sModuleName = TrimNull(sModuleName)
'                    If sModuleName <> vbNullString And sModuleName <> "?" Then
'                        myDLLs(Cnt) = sModuleName
'                        Cnt = Cnt + 1
'                    End If
'                Next i
'            End If
'            CloseHandle hProc
'        End If

    If cnt > 0 Then
        cnt = cnt - 1
        ReDim Preserve myDLLs(cnt)
        arList() = myDLLs()
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetDLLList"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RefreshDLLListNT(lPID&, objList As ListBox)
    Dim arList() As String, i&
    objList.Clear
    GetDLLList lPID, arList()
    If IsArrDimmed(arList) Then
        For i = 0 To UBound(arList)
            objList.AddItem arList(i)
        Next
    End If
End Sub



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
            
                If Not ((StrComp(.Name, "System Idle Process", 1) = 0 And .PID = 0) _
                        Or (StrComp(.Name, "System", 1) = 0 And .PID = 4) _
                        Or (StrComp(.Name, "Memory Compression", 1) = 0) _
                        Or (StrComp(.Name, "Secure System", 1) = 0)) Then

                    sProc = sProc & "|" & .PID & "=" & .Path
                End If
            End With
        Next
        GetRunningProcesses = Mid$(sProc, 2)
    End If
End Function

Public Function GetLoadedModules(lPID As Long, sProcess As String) As String
    '//TODO: sProcess
    
    Dim DLLList() As String
    
    Call GetDLLList(lPID, DLLList)

    If IsArrDimmed(DLLList) Then
        GetLoadedModules = Join(DLLList, "|")
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
    
    If AryPtr(thrID) Then
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

Public Function GetProcessThreadIDs(hProcess As Long) As Long()
    
    On Error GoTo ErrorHandler
    
    Dim hSnap As Long
    Dim PID As Long
    Dim te As THREADENTRY32
    Dim thrID() As Long
    Dim nThreads As Long
    
    ReDim thrID(0)
    
    If hProcess = GetCurrentProcess() Then
        PID = GetCurrentProcessId
    Else
        PID = GetProcessID(hProcess)
        If PID = 0 Then Exit Function
    End If
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0)
    
    If hSnap <> INVALID_HANDLE_VALUE Then
        te.dwSize = Len(te)
        
        If Thread32First(hSnap, te) Then
            Do
                If te.th32ProcessID = PID Then
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
