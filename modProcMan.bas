Attribute VB_Name = "modProcMan"
Option Explicit

Private Declare Function NtQuerySystemInformation Lib "NTDLL.DLL" (ByVal infoClass As Long, Buffer As Any, ByVal BufferSize As Long, ret As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, lpFilePart As Long) As Long
Private Declare Function QueryFullProcessImageName Lib "kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As Long, ByVal lpdwSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceW" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long
Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, uThread As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef ThreadEntry As THREADENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueW" (ByVal lpSystemName As Long, ByVal lpName As Long, lpLuid As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long

Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As Long, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStrDest As Long, ByVal lpStrSrc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal length As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type VS_FIXEDFILEINFO
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

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type

Private Type THREADENTRY32
    dwSize As Long
    dwRefCount As Long
    th32ThreadID As Long
    th32ProcessID As Long
    dwBasePriority As Long
    dwCurrentPriority As Long
    dwFlags As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount  As Long
    LuidLowPart     As Long
    LuidHighPart    As Long
    Attributes      As Long
End Type

Public Type MY_PROC_ENTRY
    Name        As String
    Path        As String
    PID         As Long
    Threads     As Long
    Priority    As Long
    SessionID   As Long
End Type

Private Type LARGE_INTEGER
    LowPart     As Long
    HighPart    As Long
End Type

Private Type CLIENT_ID
    UniqueProcess   As Long  ' HANDLE
    UniqueThread    As Long  ' HANDLE
End Type

Private Type UNICODE_STRING
    length      As Integer
    MaxLength   As Integer
    lpBuffer    As Long
End Type

Private Type VM_COUNTERS
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

Private Type IO_COUNTERS
    ReadOperationCount      As Currency 'ULONGLONG
    WriteOperationCount     As Currency
    OtherOperationCount     As Currency
    ReadTransferCount       As Currency
    WriteTransferCount      As Currency
    OtherTransferCount      As Currency
End Type

Private Type SYSTEM_THREAD
    KernelTime          As LARGE_INTEGER
    UserTime            As LARGE_INTEGER
    CreateTime          As LARGE_INTEGER
    WaitTime            As Long
    StartAddress        As Long
    ClientId            As CLIENT_ID
    Priority            As Long
    BasePriority        As Long
    ContextSwitchCount  As Long
    State               As Long 'enum KTHREAD_STATE
    WaitReason          As Long 'enum KWAIT_REASON
    dReserved01         As Long
End Type

Private Type SYSTEM_PROCESS_INFORMATION
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

Private Const SystemProcessInformation      As Long = &H5&
Private Const STATUS_INFO_LENGTH_MISMATCH   As Long = &HC0000004
Private Const STATUS_SUCCESS                As Long = 0&
Private Const ERROR_PARTIAL_COPY            As Long = 299&

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPTHREAD = &H4
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_QUERY_LIMITED_INFORMATION = &H1000
Public Const PROCESS_VM_READ = 16
Public Const THREAD_SUSPEND_RESUME = &H2

Private Const LB_SETTABSTOPS = &H192

Public ffErr As Integer

'Public Function SetCurrentProcessPrivileges(PrivilegeName As String) As Boolean
'    Const TOKEN_ADJUST_PRIVILEGES   As Long = &H20
'    Const SE_PRIVILEGE_ENABLED      As Long = &H2
'    Dim tp As TOKEN_PRIVILEGES, hToken&
'    If LookupPrivilegeValue(0&, StrPtr(PrivilegeName), tp.LuidLowPart) Then   'i.e. "SeDebugPrivilege"
'        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES, hToken) Then
'            tp.PrivilegeCount = 1
'            tp.Attributes = SE_PRIVILEGE_ENABLED
'            SetCurrentProcessPrivileges = AdjustTokenPrivileges(hToken, 0&, tp, 0&, 0&, 0&)
'            CloseHandle hToken
'        End If
'    End If
'End Function

Public Sub RefreshProcessList(objList As ListBox)
    Dim hSnap&, uPE32 As PROCESSENTRY32, i&
    Dim sExeFile$, hProcess&

    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    
    uPE32.dwSize = Len(uPE32)
    If Process32First(hSnap, uPE32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    objList.Clear
    Do
        sExeFile = TrimNull(uPE32.szExeFile)
        objList.AddItem uPE32.th32ProcessID & vbTab & sExeFile
    Loop Until Process32Next(hSnap, uPE32) = 0
    CloseHandle hSnap
End Sub

Public Sub RefreshProcessListNT(objList As ListBox)
        Dim lNumProcesses As Long, i As Long
        Dim sProcessName As String
        Dim Process() As MY_PROC_ENTRY
        
        lNumProcesses = GetProcesses_Zw(Process)
        
        If lNumProcesses Then
        
            For i = 0 To UBound(Process)
        
                sProcessName = Process(i).Path
                
                If Len(Process(i).Path) = 0 Then
                    If StrComp(Process(i).Name, "System Idle Process", 1) <> 0 _
                        And StrComp(Process(i).Name, "System", 1) <> 0 Then
                            sProcessName = Process(i).Name '& " (cannot get Process Path)"
                    End If
                End If
                
                If Len(sProcessName) <> 0 Then
                    'objList.AddItem Process(i).PID & vbTab & Process(i).SessionID & vbTab & sProcessName
                    objList.AddItem Process(i).PID & vbTab & sProcessName
                    
                End If
                
            Next
        End If

'    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
'    Dim hProc&, sProcessName$, lModules&(1 To 1024), i&
'    On Error Resume Next
'
'    If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
'        'no PSAPI.DLL file or wrong version
'        Exit Sub
'    End If
'
'    objList.Clear
'    lNumProcesses = lNeeded / 4
'    For i = 1 To lNumProcesses
'        'hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
'        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0, lProcesses(i))
'        If hProc <> 0 Then
'            'Openprocess can return 0 but we ignore this since
'            'system processes are somehow protected, further
'            'processes CAN be opened.... silly windows
'
'            lNeeded = 0
'            sProcessName = String$(260, 0)
'            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
'                GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
'                sProcessName = TrimNull(sProcessName)
'                If sProcessName <> vbNullString Then
'                    If Left$(sProcessName, 1) = "\" Then sProcessName = Mid$(sProcessName, 2)
'                    If Left$(sProcessName, 3) = "??\" Then sProcessName = Mid$(sProcessName, 4)
'                    If InStr(1, sProcessName, "%Systemroot%", vbTextCompare) > 0 Then sProcessName = Replace$(sProcessName, "%Systemroot%", sWinDir, , , vbTextCompare)
'                    If InStr(1, sProcessName, "Systemroot", vbTextCompare) > 0 Then sProcessName = Replace$(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)
'
'                    objList.AddItem lProcesses(i) & vbTab & sProcessName
'                End If
'            End If
'        End If
'        CloseHandle hProc
'    Next i
End Sub

Public Sub KillProcess(lPID&)
    Dim hProcess&
    If lPID = 0 Then Exit Sub
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProcess = 0 Then
        MsgBoxW "The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows.", vbCritical
    Else
        If TerminateProcess(hProcess, 0) = 0 Then
            MsgBoxW "The selected process could not be killed." & _
                   " It may be protected by Windows.", vbCritical
        Else
            CloseHandle hProcess
            DoEvents
        End If
    End If
End Sub

Public Sub KillProcessNT(lPID&)
    Dim hProc&
    On Error Resume Next
    If lPID = 0 Then Exit Sub
    hProc = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProc <> 0 Then
        If TerminateProcess(hProc, 0) = 0 Then
            MsgBoxW "The selected process could not be killed." & _
                   " It may be protected by Windows.", vbCritical
        Else
            CloseHandle hProc
            DoEvents
        End If
    Else
        MsgBoxW "The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows." & vbCrLf & vbCrLf & _
               "This process might be a service, which you can " & _
               "stop from the Services applet in Admin Tools." & vbCrLf & _
               "(To load this window, click Start, Run and enter 'services.msc')", vbCritical
    End If
End Sub

Public Sub RefreshDLLListNT(lPID&, objList As ListBox)
    Dim arList() As String, i&
    objList.Clear
    GetDLLList lPID, arList()
    For i = 0 To UBound(arList)
        objList.AddItem arList(i)
    Next
End Sub

Public Function GetDLLList(lPID&, arList() As String)
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024)
    Dim sModuleName$, J&, cnt&, myDLLs() As String
    On Error Resume Next
    
    ReDim myDLLs(1024): cnt = 0
    
    hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0, lPID)
    If hProc <> 0 Then
        lNeeded = 0
        If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
            For J = 2 To 1024
                If lModules(J) = 0 Then Exit For
                sModuleName = String$(260, 0)
                GetModuleFileNameExA hProc, lModules(J), sModuleName, Len(sModuleName)
                sModuleName = TrimNull(sModuleName)
                If sModuleName <> vbNullString And _
                   sModuleName <> "?" Then
                    myDLLs(cnt) = sModuleName
                    cnt = cnt + 1
                End If
            Next J
        End If
        CloseHandle hProc
    End If
    If cnt > 0 Then cnt = cnt - 1
    ReDim Preserve myDLLs(cnt)
    arList() = myDLLs()
End Function

Public Sub RefreshDLLList(lPID&, objList As ListBox)
    Dim hSnap&, uME32 As MODULEENTRY32
    Dim sDllFile$
    objList.Clear
    If lPID = 0 Then Exit Sub
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, lPID)
    uME32.dwSize = Len(uME32)
    If Module32First(hSnap, uME32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    Do
        sDllFile = TrimNull(uME32.szExePath)
        objList.AddItem sDllFile
    Loop Until Module32Next(hSnap, uME32) = 0
    CloseHandle hSnap
End Sub

Public Sub PauseProcess(lPID&, Optional bPauseOrResume As Boolean = True)
    Dim hSnap&, uTE32 As THREADENTRY32, hThread&
    If Not bIsWinNT And Not bIsWinME Then Exit Sub
    If lPID = GetCurrentProcessId Then Exit Sub
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lPID)
    If hSnap = -1 Then Exit Sub
    
    uTE32.dwSize = Len(uTE32)
    If Thread32First(hSnap, uTE32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    Do
        If uTE32.th32ProcessID = lPID Then
            hThread = OpenThread(THREAD_SUSPEND_RESUME, False, uTE32.th32ThreadID)
            If bPauseOrResume Then
                SuspendThread hThread
            Else
                ResumeThread hThread
            End If
            CloseHandle hThread
        End If
    Loop Until Thread32Next(hSnap, uTE32) = 0
    CloseHandle hSnap
End Sub

Public Sub SaveProcessList(objProcess As ListBox, objDLL As ListBox, Optional bDoDLLs As Boolean = False)
    Dim sFileName$, i&, sProcess$, sModule$, ff%
    sFileName = CmnDlgSaveFile("Save process list to file..", "Text files (*.txt)|*.txt|All files (*.*)|*.*", "processlist.txt")
    If sFileName = vbNullString Then Exit Sub
    
    On Error Resume Next
    ff = FreeFile()
    Open sFileName For Output As #ff
        Print #ff, "Process list saved on " & Format(Time, "Long Time") & ", on " & Format(Date, "Short Date")
        Print #ff, "Platform: " & sWinVersion & vbCrLf
        Print #ff, "[pid]" & vbTab & "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]"
        For i = 0 To objProcess.ListCount - 1
            sProcess = objProcess.List(i)
            Print #ff, sProcess & vbTab & vbTab & _
                      GetFilePropVersion(Mid$(sProcess, InStr(sProcess, vbTab) + 1)) & vbTab & _
                      GetFilePropCompany(Mid$(sProcess, InStr(sProcess, vbTab) + 1))
        Next i
    
'        If bDoDLLs Then
'            sProcess = objProcess.List(objProcess.ListIndex)
'            sProcess = mid$(sProcess, InStr(sProcess, vbTab) + 1)
'            Print #ff, vbCrLf & vbCrLf & "DLLs loaded by process " & sProcess & ":" & vbCrLf
'            Print #ff, "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]"
'            For i = 0 To objDLL.ListCount - 1
'                sModule = objDLL.List(i)
'                Print #ff, sModule & vbTab & vbTab & GetFilePropVersion(sModule) & vbTab & GetFilePropCompany(sModule)
'            Next i
'        End If
    
        If bDoDLLs Then
            Dim arList() As String, J&, lPID&       'Full image. DLLs of ALL processes.
            
            For i = 0 To objProcess.ListCount - 1
                sProcess = objProcess.List(i)
                lPID = CLng(Left$(sProcess, InStr(sProcess, vbTab) - 1))
                sProcess = Mid$(sProcess, InStr(sProcess, vbTab) + 1)
                GetDLLList lPID, arList()
                Print #ff, vbCrLf & vbCrLf & "DLLs loaded by process [" & lPID & "] " & sProcess & ":" & vbCrLf
                Print #ff, "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]"
                For J = 0 To UBound(arList)
                    sModule = arList(J)
                    Print #ff, sModule & vbTab & vbTab & GetFilePropVersion(sModule) & vbTab & GetFilePropCompany(sModule)
                Next
                Print #ff, vbNullString
                DoEvents
            Next
        End If

    Close #ff
    
    ShellExecute 0&, StrPtr("open"), StrPtr(sFileName), 0&, 0&, 1&
End Sub

Public Sub KillProcessByFile(sPath$)
    'Dim hSnap&, uPE32 As PROCESSENTRY32
    Dim sExeFile$, hProcess&, i&
    'Note: this sub is silent - it displays no errors !
    If sPath = vbNullString Then Exit Sub
    If bIsWinNT Then
        KillProcessNTByFile sPath
        Exit Sub
    End If
    
'    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
'
'    uPE32.dwSize = Len(uPE32)
'    If Process32First(hSnap, uPE32) = 0 Then
'        CloseHandle hSnap
'        Exit Sub
'    End If
'
'    Do
'        sExeFile = TrimNull(uPE32.szExeFile)
'        If InStr(1, sExeFile, sPath, vbTextCompare) > 0 Then
'            CloseHandle hSnap: hSnap = 0
'
'            'found the process!
'            PauseProcess uPE32.th32ProcessID
'            hProcess = OpenProcess(PROCESS_TERMINATE, 0, uPE32.th32ProcessID)
'            If hProcess <> 0 Then
'                If TerminateProcess(hProcess, 0) <> 0 Then
'                    CloseHandle hProcess
'                    DoEvents
'                End If
'            End If
'            Exit Do
'        End If
'    Loop Until Process32Next(hSnap, uPE32) = 0
'    If hSnap <> 0 Then CloseHandle hSnap

    Dim lNumProcesses As Long
    Dim sProcessPath As String
    Dim Process() As MY_PROC_ENTRY
    
    lNumProcesses = GetProcesses_Zw(Process)
        
    If lNumProcesses Then
        
        For i = 0 To UBound(Process)
        
            If StrComp(sPath, Process(i).Path, 1) = 0 Then
            
                PauseProcess Process(i).PID
                hProcess = OpenProcess(PROCESS_TERMINATE, 0, Process(i).PID)
                If hProcess <> 0 Then
                    If TerminateProcess(hProcess, 0) <> 0 Then
                        'Success
                        DoEvents
                    End If
                    CloseHandle hProcess
                End If
            End If
        Next
    End If
End Sub

Public Sub KillProcessNTByFile(sPath$)
    'Note: this sub is silent - it displays no errors!
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i&
    On Error Resume Next
    If sPath = vbNullString Then Exit Sub
    If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
        'no PSAPI.DLL file or wrong version
        Exit Sub
    End If

    lNumProcesses = lNeeded / 4
    For i = 1 To lNumProcesses
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
        If hProc <> 0 Then
            'Openprocess can return 0 but we ignore this since
            'system processes are somehow protected, further
            'processes CAN be opened.... silly windows

            lNeeded = 0
            sProcessName = String$(260, 0)
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
                            CloseHandle hProc
                            DoEvents
                        End If

                        Exit Sub
                    End If
                End If
            End If
        End If
        CloseHandle hProc
    Next i
End Sub

Public Sub CopyProcessList(objProcess As ListBox, objDLL As ListBox, Optional bDoDLLs As Boolean = False)
    Dim i&, sList$, sProcess$, sModule$
    
    On Error Resume Next
    sList = "Process list saved on " & Format(Time, "Long Time") & ", on " & Format(Date, "Short Date") & vbCrLf & _
            "Platform: " & sWinVersion & vbCrLf & vbCrLf & _
            "[pid]" & vbTab & "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]" & vbCrLf
    For i = 0 To objProcess.ListCount - 1
        sProcess = objProcess.List(i)
        sList = sList & sProcess & vbTab & vbTab & _
                GetFilePropVersion(Mid$(sProcess, InStr(sProcess, vbTab) + 1)) & vbTab & _
                GetFilePropCompany(Mid$(sProcess, InStr(sProcess, vbTab) + 1)) & vbCrLf
    Next i
    
    If bDoDLLs Then
        sProcess = objProcess.List(objProcess.ListIndex)
        sProcess = Mid$(sProcess, InStr(sProcess, vbTab) + 1)
        sList = sList & vbCrLf & vbCrLf & "DLLs loaded by process " & sProcess & ":" & vbCrLf & vbCrLf & _
                "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]" & vbCrLf
        For i = 0 To objDLL.ListCount - 1
            sModule = objDLL.List(i)
            sList = sList & sModule & vbTab & vbTab & GetFilePropVersion(sModule) & vbTab & GetFilePropCompany(sModule) & vbCrLf
        Next i
    End If
    
    Clipboard.Clear
    Clipboard.SetText sList
    If bDoDLLs Then
        MsgBoxW "The process list and dll list have been copied to your clipboard.", vbInformation
    Else
        MsgBoxW "The process list has been copied to your clipboard.", vbInformation
    End If
End Sub

Public Function GetFilePropVersion(sFileName As String) As String
    On Error GoTo ErrorHandler:
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, uVFFI As VS_FIXEDFILEINFO, sVersion$
    
    If Not FileExists(sFileName) Then Exit Function
    
    lDataLen = GetFileVersionInfoSize(StrPtr(sFileName), ByVal 0&)
    If lDataLen = 0 Then Exit Function
    
    ReDim uBuf(0 To lDataLen - 1)
    If 0 <> GetFileVersionInfo(StrPtr(sFileName), 0&, lDataLen, uBuf(0)) Then
    
        If 0 <> VerQueryValue(uBuf(0), StrPtr("\"), hData, lDataLen) Then
        
            If hData <> 0 Then
        
                CopyMemory uVFFI, ByVal hData, Len(uVFFI)
    
                With uVFFI
                    sVersion = .dwFileVersionMSh & "." & _
                        .dwFileVersionMSl & "." & _
                        .dwFileVersionLSh & "." & _
                        .dwFileVersionLSl
                End With
            End If
        End If
    End If
    GetFilePropVersion = sVersion
    Exit Function
ErrorHandler:
    ErrorMsg err, "GetFilePropVersion", sFileName
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFilePropCompany(sFileName As String) As String
    On Error GoTo ErrorHandler:
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, Stady&
    
    If Not FileExists(sFileName) Then Exit Function
    
    Stady = 1
    lDataLen = GetFileVersionInfoSize(StrPtr(sFileName), ByVal 0&)
    If lDataLen = 0 Then Exit Function
    
    Stady = 2
    ReDim uBuf(0 To lDataLen - 1)
    
    Stady = 3
    If 0 <> GetFileVersionInfo(StrPtr(sFileName), 0&, lDataLen, uBuf(0)) Then
        
        Stady = 4
        VerQueryValue uBuf(0), StrPtr("\VarFileInfo\Translation"), hData, lDataLen
        If lDataLen = 0 Then Exit Function
        
        Stady = 5
        CopyMemory uCodePage(0), ByVal hData, 4
        
        Stady = 6
        sCodePage = Right$("0" & Hex(uCodePage(1)), 2) & _
                Right$("0" & Hex(uCodePage(0)), 2) & _
                Right$("0" & Hex(uCodePage(3)), 2) & _
                Right$("0" & Hex(uCodePage(2)), 2)
        
        'get CompanyName string
        Stady = 7
        If VerQueryValue(uBuf(0), StrPtr("\StringFileInfo\" & sCodePage & "\CompanyName"), hData, lDataLen) = 0 Then Exit Function
    
        If lDataLen > 0 And hData <> 0 Then
            Stady = 8
            sCompanyName = String$(lDataLen, 0)
            
            Stady = 9
            lstrcpy ByVal StrPtr(sCompanyName), ByVal hData
        End If
        
        Stady = 10
        GetFilePropCompany = RTrimNull(sCompanyName)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "GetFilePropCompany", sFileName, "DataLen: ", lDataLen, "hData: ", hData, "sCodePage: ", sCodePage, _
        "Buf: ", uCodePage(0), uCodePage(1), uCodePage(2), uCodePage(3), "Stady: ", Stady
    If inIDE Then Stop: Resume Next
End Function

Public Sub SetListBoxColumns(objListBox As ListBox)
    Dim lTabStop&(1)
    On Error GoTo 0:
    lTabStop(0) = 70
    lTabStop(1) = 0
    SendMessage objListBox.hwnd, LB_SETTABSTOPS, UBound(lTabStop), lTabStop(0)
End Sub


Public Function GetProcesses_Zw(ProcList() As MY_PROC_ENTRY) As Long    'Return -> Count of processes
    On Error GoTo ErrorHandler:

    Const SPI_SIZE      As Long = &HB8&                                 'SPI struct: http://www.informit.com/articles/article.aspx?p=22442&seqNum=5
    Const THREAD_SIZE   As Long = &H40&
    
    Dim i           As Long
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
                        ProcName = Space$(.ImageName.length \ 2)
                        memcpy ByVal StrPtr(ProcName), ByVal .ImageName.lpBuffer, .ImageName.length
                        ProcPath = GetFilePathByPID(.ProcessID)
                        
                        If Len(ProcPath) = 0 Then
                            ProcPath = GetLongPath(ProcName)
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
    
    ReDim Preserve ProcList(cnt - 1)
    GetProcesses_Zw = cnt
    Exit Function
ErrorHandler:
    ErrorMsg err, "GetProcesses_Zw"
    If inIDE Then Stop: Resume Next
End Function

Public Sub LogError(ParamArray ErrText())
    Static Init As Boolean
    
    If Not Init Then
        Init = True
        ffErr = FreeFile()
        Open BuildPath(AppPath(), "\_Errors.log") For Output As #ffErr
    End If

    Dim i  As Long
    Dim sTotal As String
    For i = 0 To UBound(ErrText)
        sTotal = sTotal & ErrText(i) & " "
    Next
    Print #ffErr, sTotal
End Sub

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
    Dim SizeOfPath  As Long
    Dim lpFilePart  As Long

    hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0&, PID)
    
    If hProc = 0 Then
        If err.LastDllError = ERROR_ACCESS_DENIED Then
            hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0&, PID)
        End If
    End If
    
    If hProc <> 0 Then
    
        If bIsWinVistaOrLater Then
            cnt = MAX_PATH_W + 1
            ProcPath = Space$(cnt)
            Call QueryFullProcessImageName(hProc, 0&, StrPtr(ProcPath), VarPtr(cnt))
        End If
        
        If 0 <> err.LastDllError Or Not bIsWinVistaOrLater Then     'Win 2008 Server (x64) can cause Error 128 if path contains space characters
        
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
        End If
        
        If ERROR_PARTIAL_COPY = err.LastDllError Or cnt = 0 Then     'because GetModuleFileNameEx cannot access to that information for 64-bit processes on WOW64
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
                
            End If
            
        End If
        
        If cnt <> 0 Then    'if process ran with 8.3 style, GetModuleFileNameEx will return 8.3 style on x64 and full pathname on x86
                            'so wee need to expand it ourself
        
            FullPath = Space$(MAX_PATH)
            SizeOfPath = GetFullPathName(StrPtr(ProcPath), MAX_PATH, StrPtr(FullPath), lpFilePart)
            If SizeOfPath <> 0& Then
                GetFilePathByPID = Left$(FullPath, SizeOfPath)
            Else
                GetFilePathByPID = ProcPath
            End If
            
        End If
        
        CloseHandle hProc
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "GetFilePathByPID"
    If inIDE Then Stop: Resume Next
End Function

Public Function ConvertDosDeviceToDriveName(inDosDeviceName As String) As String
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
    
    sDrives = Space(cnt)
    
    cnt = GetLogicalDriveStrings(Len(sDrives), StrPtr(sDrives))

    If 0 = err.LastDllError Then
    
        aDrive = Split(Left$(sDrives, cnt - 1), vbNullChar)
    
        For i = 0 To UBound(aDrive)
            
            DosDeviceName = Space(MAX_PATH)
            
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
    ErrorMsg err, "ConvertDosDeviceToDriveName"
    If inIDE Then Stop: Resume Next
End Function

Public Function ProcessExist(NameOrPath As String) As Boolean
    Dim i As Long
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
