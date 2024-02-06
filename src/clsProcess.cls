VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'[clsProcess.cls]

'
' Process Class Module by Alex Dragokas
'
' Thanks to Catstail (Boris Fayfel), The Trick (Anatolii Krivous) and Victor Bravo VI for improvements.
'

Option Explicit

Private Const MAX_PATH As Long = 260&

Private Type PROCESSENTRY32
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

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
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

Private Type SAFEARRAY
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
End Type

Private Enum TOKEN_TYPE
    TokenPrimary = 1
    TokenImpersonation
End Enum

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Enum SECURITY_IMPERSONATION_LEVEL
    SecurityAnonymous
    SecurityIdentification
    SecurityImpersonation
    SecurityDelegation
End Enum

Private Type PIPE
    hRead As Long
    hWrite As Long
End Type

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32.dll" Alias "Process32FirstW" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" Alias "Process32NextW" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessW" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CommandLineToArgvW Lib "shell32.dll" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
'Private Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Private Declare Function GetMem2 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
'Private Declare Function SetEnvironmentVariable Lib "kernel32.dll" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long
'Private Declare Function GetEnvironmentVariable Lib "kernel32.dll" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Function DuplicateTokenEx Lib "Advapi32.dll" (ByVal hExistingToken As Long, ByVal dwDesiredAccess As Long, ByVal lpTokenAttributes As Long, ByVal ImpersonationLevel As Long, ByVal TokenType As Long, phNewToken As Long) As Long
Private Declare Function GetShellWindow Lib "user32.dll" () As Long
Private Declare Function CreateProcessWithTokenW Lib "Advapi32.dll" (ByVal hToken As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
Private Declare Function IUnknown_QueryService Lib "Shlwapi.dll" (ByVal pUnk As IUnknown, ByRef guidService As Any, ByRef riid As Any, ByRef ppvOut As Any) As Long
Private Declare Function EndTask Lib "user32.dll" (ByVal hWnd As Long, ByVal fShutDown As Long, ByVal fForce As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Private Declare Sub GetStartupInfo Lib "kernel32.dll" Alias "GetStartupInfoW" (lpStartupInfo As STARTUPINFO)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Const INVALID_HANDLE_VALUE      As Long = -1
'Private Const PROCESS_ALL_ACCESS        As Long = &H1F0FFF
Private Const INFINITE                  As Long = -1
Private Const STILL_ACTIVE              As Long = &H103
Private Const WAIT_FAILED               As Long = -1
Private Const TH32CS_SNAPPROCESS        As Long = 2
Private Const TH32CS_SNAPNOHEAPS        As Long = &H40000000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const PROCESS_QUERY_INFORMATION As Long = 1024&
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
Private Const WM_CLOSE                  As Long = 16&
Private Const WM_QUIT                   As Long = &H12  ' wparam - exitcode
Private Const PROCESS_TERMINATE         As Long = &H1
Private Const ERROR_SUCCESS             As Long = 0&
Private Const SW_SHOWNORMAL             As Long = 1&
Private Const CREATE_NEW_CONSOLE        As Long = &H10&
'Private Const STARTF_USESHOWWINDOW      As Long = &H1&
Private Const STARTF_USESTDHANDLES      As Long = &H100&
'Private Const SW_HIDE                   As Long = 0&

Private si      As STARTUPINFO
Private pi      As PROCESS_INFORMATION
Private inIDE   As Boolean
Private PipeIn  As PIPE
Private PipeOut As PIPE
Private mIdent  As String
Private m_StatusCode As Long

Public Property Get pid() As Long
    pid = pi.dwProcessId
End Property

Public Property Let pid(lPID As Long)
    pi.dwProcessId = lPID
End Property

Public Property Get ThreadId() As Long
    ThreadId = pi.dwThreadID
End Property

Public Property Let ThreadId(lTID As Long)
    pi.dwThreadID = lTID
End Property

Public Property Get HandleProc() As Long
    HandleProc = pi.hProcess
End Property

Public Property Let HandleProc(hProcess As Long)
    pi.hProcess = hProcess
End Property

Public Property Get HandleThread() As Long
    HandleThread = pi.hThread
End Property

Public Property Let HandleThread(hThread As Long)
    pi.hThread = hThread
End Property

Public Property Get Identifier() As String
    Identifier = mIdent
End Property

Public Property Let Identifier(sIdent As String)
    mIdent = sIdent
End Property

Public Function IsAlive() As Boolean
    If ExitCode() = STILL_ACTIVE Then
        IsAlive = True
    End If
End Function

Public Function ExitCode() As Long
    Dim iExit As Long
    If pi.hProcess <> 0 Then
        GetExitCodeProcess pi.hProcess, iExit
        ExitCode = iExit
    Else
        ExitCode = -1
    End If
End Function

Public Sub CloseStdIn()
    If PipeIn.hWrite <> 0 Then CloseHandle (PipeIn.hWrite): PipeIn.hWrite = 0
    If PipeIn.hRead <> 0 Then CloseHandle (PipeIn.hRead): PipeIn.hRead = 0
End Sub

Public Sub CloseStdOut()
    If PipeOut.hWrite <> 0 Then CloseHandle (PipeOut.hWrite): PipeOut.hWrite = 0
    If PipeOut.hRead <> 0 Then CloseHandle (PipeOut.hRead): PipeOut.hRead = 0
End Sub

Function IsRunned(Optional ByRef ProcessName As String, Optional ByRef ProcessID As Long) As Boolean
    ' Запущен ли процесс, указанный по имени или ProcessID
    ' (регистр символов не учитывается)
    ' Если процесс найден, в переменную ProcessID передается PID найденного процесса (если поиск велся по имени),
    ' либо в переменную ProcessName возвращается имя найденного процесса (если поиск велся по PID).

    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "clsProcess.IsRunned - Begin", "ProcName: " & ProcessName, "PID: " & ProcessID

    Dim hSnap   As Long
    Dim pe      As PROCESSENTRY32

    If ProcessName = vbNullString And ProcessID = 0 Then
        ProcessID = pid()
    End If
    
    If ProcessName = vbNullString And ProcessID = 0 Then Exit Function

    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS Or TH32CS_SNAPNOHEAPS, 0)

    If hSnap = INVALID_HANDLE_VALUE Then Exit Function

    pe.dwSize = Len(pe)

    If Process32First(hSnap, pe) Then
        Do
            If Len(ProcessName) = 0 Then
                If pe.th32ProcessID = ProcessID Then
                    IsRunned = True
                    ProcessName = StringFromPtrW(VarPtr(pe.szExeFile(0)))
                    'ProcessName = TrimNull(StrConv(pe.szExeFile, vbFromUnicode))
                    Exit Do
                End If
            Else
                If StrComp(StringFromPtrW(VarPtr(pe.szExeFile(0))), ProcessName, 1) = 0 Then
                    IsRunned = True
                    ProcessID = pe.th32ProcessID
                    Exit Do
                End If
            End If
        Loop While Process32Next(hSnap, pe)
    End If
    CloseHandle hSnap
    
    AppendErrorLogCustom "clsProcess.IsRunned - End"
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_IsRunned"
    If inIDE Then Stop: Resume Next
End Function


Function WaitForTerminate( _
                        Optional ByVal WindowHandle As Long, _
                        Optional ByVal ProcessHandle As Long, _
                        Optional ByVal FreezeThread As Boolean, _
                        Optional ByVal TimeoutMs As Long, _
                        Optional ByVal ProcessID As Long) As Long
                        ' Ожидает завершения процесса                           'Thanks for help to Catstail (Б. Файфель)
    
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "WaitForTerminate - Begin", "WndHandle: " & WindowHandle, "ProcessHandle: " & ProcessHandle, "FreezeThread? " & FreezeThread, _
        "TimeOutMs: " & TimeoutMs, "PID: " & ProcessID
    
    ' FreezeThread, если true - заморозит текущий поток на время, пока отслеживаемый процесс не будет закрыт
    ' TimeOut, если FreezeThread = false, задает время таймаута ожидания завершения процесса.
    ' Возвращаемое значение: ExitCode процесса, или -1, если вышел таймаут ожидания или WaitForSingleObject вернул ошибку (WAIT_FAILED)
    
    Dim iExit       As Long
    Dim hProc       As Long
    Dim iResult     As Long
    Dim dwThreadID  As Long
    Dim hThread     As Long
    Dim InitialTime As Single
    Dim DeadTime    As Single
    Dim bTimeOutON  As Boolean
    Dim bNewHandle  As Boolean
    
    Static isInit As Boolean
    Static bIsWinVistaOrLater As Boolean
    
    If Not isInit Then
        isInit = True
        Dim inf(68) As Long
        inf(0) = 276: GetVersionEx inf(0): bIsWinVistaOrLater = (inf(1) >= 6)
    End If
    
    If ProcessID <> 0 Then
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or SYNCHRONIZE, False, ProcessID)
        bNewHandle = True
        
    ElseIf WindowHandle <> 0 Then
        ' получаем PID процесса
        dwThreadID = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        '::: Получаем хэндл процесса
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or SYNCHRONIZE, False, ProcessID)
        bNewHandle = True
    ElseIf ProcessHandle <> 0 Then
        hProc = ProcessHandle
    Else
        hProc = HandleProc()
        hThread = HandleThread()
    End If
    
    If hProc = 0 Then Exit Function
    
    If FreezeThread Then
       '::: Ждем до победного конца
       iResult = WaitForSingleObject(hProc, INFINITE)
       '::: Если ожидание завершилось неудачно - вернем код -1
       If (iResult = WAIT_FAILED) Then
           iExit = -1
       Else
           '::: Получаем код завершения
           GetExitCodeProcess hProc, iExit
       End If
    Else
       '::: Пробуем получить код завершения
       GetExitCodeProcess hProc, iExit
       InitialTime = Timer
       DeadTime = InitialTime + TimeoutMs / 1000
       bTimeOutON = (TimeoutMs <> INFINITE)
       '::: Если процесс все еще в состоянии выполнения,
       '::: не дадим системе зависнуть
       Do While (iExit = STILL_ACTIVE) And IIf(bTimeOutON, DeadTime > Timer, True)
          DoEvents
          Sleep 100
          GetExitCodeProcess hProc, iExit
       Loop
       If bTimeOutON And (Timer > DeadTime) Then iExit = -1
    End If
    '::: Закроем хендл при успешном завершении работы приложения, иначе он может пригодиться для принудительного завершения процесса
    If iExit = ERROR_SUCCESS Then
        CloseHandle hThread: hThread = 0
        CloseHandle hProc: hProc = 0
        If pi.hProcess = hProc Then pi.hProcess = 0
        If pi.hThread = hThread Then pi.hThread = 0
    End If
    If bNewHandle Then
        '::: Если открывали хендл самостоятельно в этой процедуре, не забываем закрыть его
        If hProc <> 0 Then CloseHandle hProc
    End If
    
    WaitForTerminate = iExit
    
    AppendErrorLogCustom "WaitForTerminate - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_WaitForTerminate"
    If inIDE Then Stop: Resume Next
End Function

Private Function GetPIDsByName(ProcessName As String) As Long()
    On Error GoTo ErrorHandler:
    Dim pe               As PROCESSENTRY32
    Dim hSnap            As Long
    Dim pid()            As Long
    ReDim pid(0) As Long
    
    AppendErrorLogCustom "GetPIDsByName - Begin", "ProcessName: " & ProcessName
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS Or TH32CS_SNAPNOHEAPS, 0)

    If hSnap = INVALID_HANDLE_VALUE Then Exit Function

    pe.dwSize = Len(pe)
    
    If Process32First(hSnap, pe) Then
        Do
            If StrComp(StringFromPtrW(VarPtr(pe.szExeFile(0))), ProcessName, 1) = 0 Then
                
                ReDim Preserve pid(UBound(pid) + 1) As Long
                pid(UBound(pid)) = pe.th32ProcessID
            End If
        Loop While Process32Next(hSnap, pe)
    End If
    CloseHandle hSnap
    
    GetPIDsByName = pid
    
    AppendErrorLogCustom "GetPIDsByName - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_GetPIDsByName"
    If inIDE Then Stop: Resume Next
End Function

Public Function ProcessClose( _
                            Optional ByVal WindowHandle As Long, _
                            Optional ProcessName As String, _
                            Optional Async As Boolean, _
                            Optional TimeoutMs As Long = INFINITE, _
                            Optional ByVal ProcessID As Long, _
                            Optional SendCloseMsg As Boolean) As Long
                            'Возвращаемое значение = ExitCode процесса, или -1 в случае неудачи
    
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "ProcessClose - Begin", "WindowHandle: " & WindowHandle, "ProcessName: " & ProcessName, "Async? " & Async, _
        "TimeOutMs: " & TimeoutMs, "PID: " & ProcessID, "SendCloseMsg? " & SendCloseMsg

    Dim lr               As Long
    Dim hThread          As Long
    Dim hProc            As Long
    Dim dwThreadID       As Long
    Dim pid()            As Long
    Dim i                As Long
    Dim ProcNameCur      As String

    If Len(ProcessName) <> 0 Then

        pid() = GetPIDsByName(ProcessName)
        
        If IsArrDimmed(pid) Then
            For i = 1 To UBound(pid)
                ProcessClose , , True, 0, pid(i), SendCloseMsg 'Async mode
            Next
        End If
        
        If Not Async Then
            If TimeoutMs = INFINITE Then
                For i = 1 To UBound(pid)
                    ProcNameCur = vbNullString
                    Call IsRunned(ProcNameCur, pid(i)) 'get name by PID
                    If StrComp(ProcNameCur, ProcessName, vbTextCompare) = 0 Then 'not changed ?
                        Call WaitForTerminate(, , False, INFINITE, pid(i))
                    End If
                Next
            Else
                For i = 1 To UBound(pid)
                    ProcNameCur = vbNullString
                    Call IsRunned(ProcNameCur, pid(i)) 'get name by PID
                    If StrComp(ProcNameCur, ProcessName, vbTextCompare) = 0 Then 'not changed ?
                        ProcessID = pid(i)
                        Exit For ' get only 1 process
                    End If
                Next
                If ProcessID <> 0 Then Call WaitForTerminate(, , False, TimeoutMs, ProcessID)
            End If
        End If
        
        ProcessID = 0
        ProcessClose = IsRunned(ProcessName, ProcessID)
        Exit Function
        
    ElseIf WindowHandle <> 0 Then
        dwThreadID = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        lr = PostMessage(WindowHandle, WM_CLOSE, 0, 0)                           ' WM_CLOSE главному окну
        Call WaitForTerminate(, hProc, False, 1000)                              ' подождём 1 сек.
        lr = PostMessage(WindowHandle, WM_QUIT, 0, 0)                            ' WM_QUIT главному окну
        If Not Async Then
            lr = WaitForTerminate(WindowHandle, , False, TimeoutMs)              ' ожидание, без замораживания текущего потока
        Else
            lr = WaitForTerminate(WindowHandle, , False, 500)
        End If
        If IsProcedureAvail("EndTask", "user32.dll") Then
            EndTask WindowHandle, False, True
        End If
        If IsRunned(, ProcessID) Then
            hProc = OpenProcess(PROCESS_TERMINATE Or SYNCHRONIZE, False, ProcessID)
            If hProc = 0 Then hProc = OpenProcess(PROCESS_TERMINATE, False, ProcessID)
        End If
        
    ElseIf ProcessID <> 0 Then
        hProc = OpenProcess(PROCESS_TERMINATE Or SYNCHRONIZE, False, ProcessID)
        If hProc = 0 Then hProc = OpenProcess(PROCESS_TERMINATE, False, ProcessID)
    
    Else
        hThread = pi.hThread
        ProcessID = pi.dwProcessId
        hProc = pi.hProcess
    End If
    
    If SendCloseMsg Then
        If (ProcessID <> 0) Then
            hThread = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
            If (WindowHandle <> 0) Then
                lr = PostMessage(WindowHandle, WM_CLOSE, 0, 0)                           ' WM_CLOSE главному окну
                If (hProc = 0) Then
                    hProc = OpenProcess(SYNCHRONIZE, False, ProcessID)
                End If
                If (hProc <> 0) Then
                    Call WaitForTerminate(, hProc, False, 1000)
                End If
            End If
        End If
    End If
    
    If hProc <> 0 Then
        TerminateProcess hProc, 0
        If Not (Async) Then
            ProcessClose = WaitForTerminate(, hProc, False, TimeoutMs)
        End If
    End If
    
    If (ProcessID <> 0) Then
        ProcessClose = Not IsRunned(, ProcessID)                           ' Проверю наличие в списке процессов
    End If
    
    If ProcessClose Then
        If hThread <> 0 Then CloseHandle hThread
        If hProc <> 0 Then CloseHandle hProc
        If pi.hProcess = hProc Then pi.hProcess = 0
        If pi.hThread = hThread Then pi.hThread = 0
    End If
    
    AppendErrorLogCustom "ProcessClose - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_ProcessClose", "ProcessName:", ProcessName
    If inIDE Then Stop: Resume Next
End Function

Public Function RunPowershell( _
    ByVal sCmd As String, _
    Optional bWait As Boolean = False, _
    Optional iTimeoutMs As Long = 30000, _
    Optional WindowStyle As SHOWWINDOW_FLAGS = SW_HIDE) As Boolean

    sCmd = "-ExecutionPolicy UnRestricted -c " & """" & sCmd & """"
    RunPowershell = Proc.ProcessRun(BuildPath(sWinSysDir, "WindowsPowerShell\v1.0\powershell.exe"), sCmd, , WindowStyle)
    If RunPowershell And bWait Then
        Me.WaitForTerminate , , , iTimeoutMs
    End If
    
End Function

Public Function ProcessRun( _
                            ByVal FileName As String, _
                            Optional ByVal CommandLine As String, _
                            Optional ByVal CurrentDirectory As String = vbNullString, _
                            Optional WindowStyle As SHOWWINDOW_FLAGS = SW_SHOWNORMAL, _
                            Optional CloseHandles As Boolean = False, _
                            Optional bReadConsole As Boolean) As Boolean
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ProcessRun - Begin", "FileName: " & FileName, "CommandLine: " & CommandLine, "CurDir: " & CurrentDirectory, _
        "WindowStyle: " & WindowStyle, "CloseHandles? " & CloseHandles
    
    Const STARTF_USESHOWWINDOW      As Long = 1
    Const NORMAL_PRIORITY_CLASS     As Long = &H20
    Const PIPE_BUFFER_SIZE          As Long = 64 * 1024&
    
    Dim N           As Long
    Dim lr          As Long
    Dim CMDLine     As String
    Dim argc        As Long
    Dim argv()      As String
    Dim iLastErr    As Long
    
    AppendErrorLogCustom "Parsing command line"
    
    If Len(FileName) <> 0 Then
    
        CMDLine = """" & FileName & """"                    ' Имя файла в кавычки
        
        If Len(CommandLine) <> 0 Then                       ' Если есть аргументы, их тоже добавляем в кавычках
        
            ParseCommandLine CommandLine, argc, argv        ' Разбиваем аргументы командной строки на составляющие
            
            For N = 1 To argc
                CMDLine = CMDLine & " " & argv(N)
            Next
        End If
        
        'если командная строка подготовлена для cmd.exe, не трогаем аргументы
        If StrComp(FileName, Environ$("ComSpec"), 1) = 0 Or StrComp(FileName, "schtasks.exe", 1) = 0 Then
            CMDLine = """" & FileName & """" & " " & CommandLine
        End If
    Else
        CMDLine = CommandLine
    End If
    
    si.cb = Len(si)
    GetStartupInfo si                                   ' клонируем структуру текущего процесса
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = WindowStyle                        ' Стиль окна
    
    If bReadConsole Then
        Dim Attr As SECURITY_ATTRIBUTES
        Attr.nLength = Len(Attr)
        Attr.lpSecurityDescriptor = 0
        Attr.bInheritHandle = True
        
        AppendErrorLogCustom "CreatePipe1"
        
        If CreatePipe(PipeIn.hRead, PipeIn.hWrite, Attr, PIPE_BUFFER_SIZE) <> 0 Then
    
            AppendErrorLogCustom "CreatePipe2"
    
            If CreatePipe(PipeOut.hRead, PipeOut.hWrite, Attr, PIPE_BUFFER_SIZE) <> 0 Then
      
                si.hStdInput = PipeIn.hRead
                si.hStdOutput = PipeOut.hWrite
                si.hStdError = PipeOut.hWrite
                si.dwFlags = si.dwFlags Or STARTF_USESTDHANDLES
            End If
        End If
    End If
    
    AppendErrorLogCustom "CreateProcess"
    
    If si.hStdOutput <> 0 Then
    
        lr = CreateProcess(ByVal 0&, _
                StrPtr(CMDLine), _
                ByVal 0&, _
                ByVal 0&, _
                1&, _
                NORMAL_PRIORITY_CLASS Or CREATE_NEW_CONSOLE, _
                ByVal 0&, _
                StrPtr(CurrentDirectory), _
                si, _
                pi)
        
        iLastErr = Err.LastDllError
        
        If PipeOut.hWrite <> 0 Then
            CloseHandle PipeOut.hWrite: PipeOut.hWrite = 0
        End If
    Else

        lr = CreateProcess(ByVal 0&, _
                    StrPtr(CMDLine), _
                    ByVal 0&, _
                    ByVal 0&, _
                    False, _
                    NORMAL_PRIORITY_CLASS, _
                    ByVal 0&, _
                    StrPtr(CurrentDirectory), _
                    si, _
                    pi)
                    
        iLastErr = Err.LastDllError
    End If
    
    ProcessRun = CBool(lr)    ' not 0 is SUCCESS
    m_StatusCode = Err.LastDllError
    
    If CloseHandles Then
        If pi.hProcess <> 0 Then CloseHandle pi.hProcess: pi.hProcess = 0
        If pi.hThread <> 0 Then CloseHandle pi.hThread: pi.hThread = 0
    End If
    
    AppendErrorLogCustom "ProcessRun - End. Return: " & lr & ". LastError = " & iLastErr
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_ProcessRun", "FileName:", FileName, "CommandLine:", CommandLine
    If inIDE Then Stop: Resume Next
End Function

Public Function ProcessRunUnelevated( _
                            ByVal FileName As String, _
                            Optional ByVal CommandLine As String, _
                            Optional ByVal CurrentDirectory As String = vbNullString, _
                            Optional WindowStyle As SHOWWINDOW_FLAGS = SW_SHOWNORMAL, _
                            Optional CloseHandles As Boolean = False) As Long
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ProcessRunUnelevated - Begin", "FileName: " & FileName, "CommandLine: " & CommandLine, "CurDir: " & CurrentDirectory, _
        "WindowStyle: " & WindowStyle, "CloseHandles? " & CloseHandles
    
    'This code launch program via filtered token context; parent process handle is current process
    
    Const STARTF_USESHOWWINDOW      As Long = 1&
    Const TOKEN_QUERY               As Long = 8&
    Const TOKEN_ASSIGN_PRIMARY      As Long = 1&
    Const TOKEN_DUPLICATE           As Long = 2&
    Const TOKEN_ADJUST_DEFAULT      As Long = &H80&
    Const TOKEN_ADJUST_SESSIONID    As Long = &H100&
    
    Const TOKEN_RIGHTS As Long = TOKEN_QUERY Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_ADJUST_DEFAULT Or TOKEN_ADJUST_SESSIONID
    
    Dim hShellProcessToken As Long
    Dim hPrimaryToken As Long
    Dim lShellPID   As Long
    Dim hWndShell   As Long
    Dim hProcShell  As Long
    Dim N           As Long
    Dim lr          As Long
    Dim CMDLine     As String
    Dim argc        As Long
    Dim argv()      As String
    
    CMDLine = """" & FileName & """"                    ' Имя файла в кавычки
    
    If CommandLine <> vbNullString Then                 ' Если есть аргументы, их тоже добавляем в кавычках
    
        ParseCommandLine CommandLine, argc, argv        ' Разбиваем аргументы командной строки на составляющие
        For N = 1 To argc
            CMDLine = CMDLine & " """ & argv(N) & """"
        Next
    End If
    
    'если командная строка подготовлена для cmd.exe, не трогаем аргументы
    If StrComp(FileName, Environ$("ComSpec"), 1) = 0 Or StrComp(FileName, "schtasks.exe", 1) = 0 Then
        CMDLine = """" & FileName & """" & " " & CommandLine
    End If
    
    si.cb = Len(si)
    GetStartupInfo si                                   ' клонируем структуру текущего процесса
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = WindowStyle                        ' Стиль окна
    
    'SetCurrentProcessPrivileges "SeIncreaseQuotaPrivilege"  'CreateProcessWithTokenW
    'SetCurrentProcessPrivileges "SeImpersonatePrivilege"    'CreateProcessWithTokenW
    
    If Not OSver.IsElevated Then
        ProcessRunUnelevated = ProcessRun(FileName, CommandLine, CurrentDirectory, WindowStyle, CloseHandles)
        Exit Function
    End If
    
    If Not ProcessExist("explorer.exe", True) Then
        If inIDE Then Debug.Print "ProcessRunUnelevated. No explorer process found."
        Exit Function
    End If
    
    hWndShell = GetShellWindow()
    
    If hWndShell <> 0 Then
        GetWindowThreadProcessId hWndShell, lShellPID
        
        If lShellPID <> 0 Then
            
            hProcShell = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), _
                False, lShellPID)
            
            If hProcShell <> 0 Then
                
                If OpenProcessToken(hProcShell, TOKEN_DUPLICATE, hShellProcessToken) Then
                    
                    If DuplicateTokenEx(hShellProcessToken, TOKEN_RIGHTS, 0&, SecurityImpersonation, TokenPrimary, hPrimaryToken) Then
                        
                        lr = CreateProcessWithTokenW(hPrimaryToken, 0&, _
                            ByVal 0&, _
                            StrPtr(CMDLine), _
                            ByVal 0&, _
                            ByVal 0&, _
                            StrPtr(CurrentDirectory), _
                            si, _
                            pi)
                            
                        CloseHandle hPrimaryToken
                    End If
                    CloseHandle hShellProcessToken
                End If
                CloseHandle hProcShell
            End If
        End If
    End If
    
    ProcessRunUnelevated = lr    ' not 0 is SUCCESS
    
    If CloseHandles Then
        If pi.hProcess <> 0 Then CloseHandle pi.hProcess: pi.hProcess = 0
        If pi.hThread <> 0 Then CloseHandle pi.hThread: pi.hThread = 0
    End If
    
    AppendErrorLogCustom "ProcessRunUnelevated - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_ProcessRunUnelevated", "FileName:", FileName, "CommandLine:", CommandLine
    If inIDE Then Stop: Resume Next
End Function

Public Function ProcessRunUnelevated2( _
                            ByVal FileName As String, _
                            Optional ByVal CommandLine As String, _
                            Optional ByVal CurrentDirectory As String = vbNullString, _
                            Optional WindowStyle As SHOWWINDOW_FLAGS = SW_SHOWNORMAL, _
                            Optional Verb As String = vbNullString) As Boolean
    
    On Error GoTo ErrorHandler:
    'This code launch program via filtered token context; parent process handle is 'explorer.exe' process
    'Thanks to Victor Bravo VI (vbforums) for this code
    
    AppendErrorLogCustom "ProcessRunUnelevated2 - Begin", "FileName: " & FileName, "CommandLine: " & CommandLine, "CurDir: " & CurrentDirectory, _
        "WindowStyle: " & WindowStyle
    
    Const CC_STDCALL = 4&, QueryActiveShellView_Offset = 15& * 4&, GetItemObject_Offset = 15& * 4&
    Const NOERROR = 0&, S_OK = 0&, SVGIO_BACKGROUND = 0&, SWC_DESKTOP = 8&, SWFO_NEEDDISPATCH = 1&

    If OSver.MajorMinor <= 5.2 Then
        ProcessRunUnelevated2 = ProcessRun(FileName, CommandLine, CurrentDirectory, WindowStyle, True)
        Exit Function
    End If

    Dim IShellBrowser As IUnknown, IShellView As IUnknown, Obj As Object, ShellFolderView As Object
    Dim IID_IDispatch(3&) As Long, IID_IShellBrowser(3&) As Long, SID_STopLevelBrowser(3&) As Long
    Dim hWnd As Long, RV As Variant, v(2&) As Variant, VT(2&) As Integer, pv(2&) As Long
    
    Dim vOperation As Variant:          vOperation = Verb 'by default - open
    Dim vCommandLine As Variant:        vCommandLine = CommandLine
    Dim vCurrentDirectory As Variant:   vCurrentDirectory = CurrentDirectory
    Dim vWindowStyle As Variant:        vWindowStyle = WindowStyle
    
    On Error Resume Next
    Set Obj = CreateObject("Shell.Application").Windows.FindWindowSW(Empty, Empty, SWC_DESKTOP, hWnd, SWFO_NEEDDISPATCH)
    On Error GoTo ErrorHandler
    If Not (Obj Is Nothing) Then
      If CLSIDFromString(StrPtr("{4C96BE40-915C-11CF-99D3-00AA004AE837}"), SID_STopLevelBrowser(0&)) = NOERROR Then
        If CLSIDFromString(StrPtr("{000214E2-0000-0000-C000-000000000046}"), IID_IShellBrowser(0&)) = NOERROR Then
          If IUnknown_QueryService(Obj, SID_STopLevelBrowser(0&), IID_IShellBrowser(0&), IShellBrowser) = S_OK Then
            If DispCallFunc(ObjPtr(IShellBrowser), QueryActiveShellView_Offset, CC_STDCALL, vbLong, 1&, vbLong, VarPtr(CVar(VarPtr(IShellView))), RV) = S_OK And RV = S_OK Then
              If CLSIDFromString(StrPtr("{00020400-0000-0000-C000-000000000046}"), IID_IDispatch(0&)) = NOERROR Then
    
                v(0&) = SVGIO_BACKGROUND:          VT(0&) = vbLong:  pv(0&) = VarPtr(v(0&))
                v(1&) = VarPtr(IID_IDispatch(0&)): VT(1&) = vbLong:  pv(1&) = VarPtr(v(1&))
                v(2&) = VarPtr(ShellFolderView):   VT(2&) = vbLong:  pv(2&) = VarPtr(v(2&))
    
                If DispCallFunc(ObjPtr(IShellView), GetItemObject_Offset, CC_STDCALL, vbLong, 3&, VT(0&), pv(0&), RV) = S_OK And RV = S_OK Then
    
                  ShellFolderView.Application.ShellExecute FileName, vCommandLine, vCurrentDirectory, vOperation, vWindowStyle
                  
                  ProcessRunUnelevated2 = True
                End If
              End If
            End If
          End If
        End If
      End If
    End If
    
    If Not ProcessRunUnelevated2 Then
        ProcessRunUnelevated2 = ProcessRun(FileName, CommandLine, CurrentDirectory, WindowStyle, True)
    End If
    
    AppendErrorLogCustom "ProcessRunUnelevated - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ProcessRunUnelevated2", "FileName:", FileName, "CommandLine:", CommandLine
    If inIDE Then Stop: Resume Next
End Function

'true - to switch thread content to a filtered token
'false - to revert thread to original context
'@return: true - on success
Public Function Unelevate() As Boolean
    '// TODO

End Function

Public Function ConsoleWrite(ByVal sData As String) As Boolean
    Dim lBytesWrote As Long
 
    sData = StrConv(sData, vbFromUnicode)
 
    If WriteFile(PipeIn.hWrite, StrPtr(sData), LenB(sData), lBytesWrote, 0&) Then
    
        If lBytesWrote = LenB(sData) Then ConsoleWrite = True
    End If
End Function

Public Function GetConsoleReadSize() As Long
    PeekNamedPipe PipeOut.hRead, ByVal 0&, 0&, ByVal 0&, GetConsoleReadSize, ByVal 0&
End Function

Public Function ConsoleRead() As String
    Dim buf As String, Count As Long
    
    PeekNamedPipe PipeOut.hRead, ByVal 0&, 0&, ByVal 0&, Count, ByVal 0&
    If Count > 0 Then
        buf = String$((Count + 1) \ 2, 0&)
        If ReadFile(PipeOut.hRead, ByVal StrPtr(buf), Count, Count, ByVal 0&) Then
            ConsoleRead = ConvertCodePage(StrPtr(StrConv(buf, vbUnicode)), CP_DOS, CP_UTF16LE)
        End If
    End If
End Function

Public Function ConsoleReadUntilDeath(Optional iTimeoutMs As Long = INFINITE) As String
    Const DELAY_INTERVAL As Long = 400
    Dim sText As String
    Dim iElapsed As Long
    Me.CloseStdIn
    Do
        While GetConsoleReadSize() <> 0
            sText = sText & ConsoleRead()
        Wend
        Sleep 100 'gives system a time to dealloc process
        If Not IsAlive() Then Exit Do
        SleepNoLock DELAY_INTERVAL
        iElapsed = iElapsed + DELAY_INTERVAL
    Loop While (iElapsed < iTimeoutMs) Or (iTimeoutMs = INFINITE)
    ConsoleReadUntilDeath = sText
End Function

Private Sub ParseCommandLine(Line As String, argc As Long, argv() As String)
    On Error GoTo ErrorHandler:
    Dim ptr     As Long
    Dim Index   As Long
    Dim strLen  As Long
    Dim strAdr  As Long
    
    ptr = CommandLineToArgvW(StrPtr(Line), argc)
    ReDim argv(argc)
    argv(0) = App.path & "\" & App.ExeName & ".exe"
    If argc = 0 Then Exit Sub
    
    For Index = 1 To argc
        GetMem4 ByVal ptr + (Index - 1) * 4, strAdr
        strLen = lstrlen(strAdr)
        argv(Index) = Space$(strLen)
        lstrcpyn StrPtr(argv(Index)), strAdr, strLen + 1
    Next
    LocalFree ptr
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsProcess_ParseCommandLine", "Line:", Line
    If inIDE Then Stop: Resume Next
End Sub

Private Function TrimNull(s$) As String
    TrimNull = Left$(s, lstrlen(StrPtr(s)))
End Function

Private Sub Class_Initialize()
    Debug.Assert MakeTrue(inIDE)
End Sub

Public Function IsArrDimmed(vArray As Variant) As Boolean
    IsArrDimmed = (GetArrDims(vArray) > 0)
End Function

Public Function GetArrDims(vArray As Variant) As Integer
    Dim ppSA As Long
    Dim pSA As Long
    Dim VT As Long
    Dim sa As SAFEARRAY
    Const vbByRef As Integer = 16384

    If IsArray(vArray) Then
        GetMem4 ByVal VarPtr(vArray) + 8, ppSA      ' pV -> ppSA (pSA)
        If ppSA <> 0 Then
            GetMem2 vArray, VT
            If VT And vbByRef Then
                GetMem4 ByVal ppSA, pSA                 ' ppSA -> pSA
            Else
                pSA = ppSA
            End If
            If pSA <> 0 Then
                memcpy sa, ByVal pSA, LenB(sa)
                If sa.pvData <> 0 Then
                    GetArrDims = sa.cDims
                End If
            End If
        End If
    End If
End Function

Private Sub Class_Terminate()
    If PipeIn.hRead <> 0 Then CloseHandle PipeIn.hRead
    If PipeIn.hWrite <> 0 Then CloseHandle PipeIn.hWrite
    If PipeOut.hRead <> 0 Then CloseHandle PipeOut.hRead
    If PipeOut.hWrite <> 0 Then CloseHandle PipeOut.hWrite
    If pi.hThread <> 0 Then CloseHandle pi.hThread
    If pi.hProcess <> 0 Then CloseHandle pi.hProcess
End Sub

Private Function StringFromPtrW(ptr As Long) As String
    Dim strSize As Long
    If 0 <> ptr Then
        strSize = lstrlen(ptr)
        If 0 <> strSize Then
            StringFromPtrW = String$(strSize, 0&)
            lstrcpyn StrPtr(StringFromPtrW), ptr, strSize + 1&
        End If
    End If
End Function

'returns Win API Status code of 'ProcessRun' requested in this class
'
Public Property Get StatusCode() As Long
    StatusCode = m_StatusCode
End Property

Private Function MakeTrue(ByRef bValue As Boolean) As Boolean
    MakeTrue = True
    bValue = True
End Function
