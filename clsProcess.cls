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
'
' Process Class Module by Alex Dragokas
'
' Thanks to Catstail (Boris Fayfel) & The Trick (Anatolii Krivous) for improvements.
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
    szExeFile As Long 'As String * MAX_PATH
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

Enum Win_Style
    SW_HIDE = 0&
    SW_MINIMIZE = 6&
    SW_SHOWNORMAL = 1&
End Enum

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
Private Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32.dll" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32.dll" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long

Private Declare Sub GetStartupInfo Lib "kernel32.dll" Alias "GetStartupInfoW" (lpStartupInfo As STARTUPINFO)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const PROCESS_ALL_ACCESS        As Long = &H1F0FFF
Private Const INFINITE                  As Long = -1
Private Const STILL_ACTIVE              As Long = &H103
Private Const WAIT_FAILED               As Long = -1
Private Const TH32CS_SNAPPROCESS        As Long = 2
Private Const SYNCHRONIZE               As Long = &H100000
Private Const PROCESS_QUERY_INFORMATION As Long = 1024&
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
Private Const WM_CLOSE                  As Long = 16&
Private Const WM_QUIT                   As Long = &H12  ' wparam - exitcode
Private Const PROCESS_TERMINATE         As Long = &H1

Private si    As STARTUPINFO
Private pi    As PROCESS_INFORMATION


Public Property Get GetProcessID() As Long
    GetProcessID = pi.dwProcessId
End Property

Public Property Get GetThreadId() As Long
    GetThreadId = pi.dwThreadId
End Property

Public Property Get GetProcessHandle() As Long
    GetProcessHandle = pi.hProcess
End Property

Public Property Get GetThreadHandle() As Long
    GetThreadHandle = pi.hThread
End Property

Function IsRunned(Optional ByRef ProcessName As String, Optional ByRef ProcessID As Long) As Boolean
    ' Запущен ли процесс, указанный по имени или ProcessID
    ' (регистр символов не учитывается)
    ' Если процесс найден, в переменную ProcessID передается PID найденного процесса (если поиск велся по имени),
    ' либо в переменную ProcessName возвращается имя найденного процесса (если поиск велся по PID).

    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "clsProcess.IsRunned - Begin", "ProcName: " & ProcessName, "PID: " & ProcessID

    Dim hSnap   As Long
    Dim pe      As PROCESSENTRY32

    If ProcessName = vbNullString And ProcessID = 0 Then Exit Function

    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

    If hSnap = INVALID_HANDLE_VALUE Then Exit Function

    pe.dwSize = Len(pe)

    If Process32First(hSnap, pe) Then
        Do
            If ProcessName = vbNullString Then
                If pe.th32ProcessID = ProcessID Then
                    IsRunned = True
                    ProcessName = TrimNull(StrConv(pe.szExeFile, vbFromUnicode))
                    Exit Do
                End If
            Else
                If StrComp(TrimNull(StrConv(pe.szExeFile, vbFromUnicode)), ProcessName, 1) = 0 Then
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
                        Optional ByVal TimeOutMs As Long, _
                        Optional ByVal ProcessID As Long) As Long
                        ' Ожидает завершения процесса                           'Thanks for help to Catstail (Б. Файфель)
    
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "WaitForTerminate - Begin", "WndHandle: " & WindowHandle, "ProcessHandle: " & ProcessHandle, "FreezeThread? " & FreezeThread, _
        "TimeOutMs: " & TimeOutMs, "PID: " & ProcessID
    
    ' FreezeThread, если true - заморозит текущий поток на время, пока отслеживаемый процесс не будет закрыт
    ' TimeOut, если FreezeThread = false, задает время таймаута ожидания завершения процесса.
    ' Возвращаемое значение: ExitCode процесса, или -1, если вышел таймаут ожидания или WaitForSingleObject вернул ошибку (WAIT_FAILED)
    
    Dim cRead       As Long
    Dim iExit       As Long
    Dim hProc       As Long
    Dim iResult     As Long
    Dim dwThreadId  As Long
    Dim hThread     As Long
    Dim InitialTime As Single
    Dim DeadTime    As Single
    Dim bTimeOutON  As Boolean
    
    Static IsInit As Boolean
    Static bIsWinVistaOrLater As Boolean
    
    If Not IsInit Then
        IsInit = True
        Dim inf(68) As Long
        inf(0) = 276: GetVersionEx inf(0): bIsWinVistaOrLater = (inf(1) >= 6)
    End If
    
    If ProcessID <> 0 Then
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or SYNCHRONIZE, False, ProcessID)
        
    ElseIf WindowHandle <> 0 Then
        ' получаем PID процесса
        dwThreadId = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        '::: Получаем хэндл процесса
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or SYNCHRONIZE, False, ProcessID)
    ElseIf ProcessHandle <> 0 Then
        hProc = ProcessHandle
    Else
        hProc = GetProcessHandle()
        hThread = GetThreadHandle()
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
       DeadTime = InitialTime + TimeOutMs / 1000
       bTimeOutON = (TimeOutMs <> INFINITE)
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
        CloseHandle hThread
        CloseHandle hProc
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
    Dim PID()            As Long
    ReDim PID(0)
    
    AppendErrorLogCustom "GetPIDsByName - Begin", "ProcessName: " & ProcessName
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

    If hSnap = INVALID_HANDLE_VALUE Then Exit Function

    pe.dwSize = Len(pe)

    If Process32First(hSnap, pe) Then
        Do
            If StrComp(TrimNull(StrConv(pe.szExeFile, vbFromUnicode)), ProcessName, 1) = 0 Then
                
                ReDim Preserve PID(UBound(PID) + 1)
                PID(UBound(PID)) = pe.th32ProcessID
            End If
        Loop While Process32Next(hSnap, pe)
    End If
    CloseHandle hSnap
    
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
                            Optional TimeOutMs As Long = INFINITE, _
                            Optional ByVal ProcessID As Long, _
                            Optional SendCloseMsg As Boolean) As Long
                            'Возвращаемое значение = ExitCode процесса, или -1 в случае неудачи
    
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "ProcessClose - Begin", "WindowHandle: " & WindowHandle, "ProcessName: " & ProcessName, "Async? " & Async, _
        "TimeOutMs: " & TimeOutMs, "PID: " & ProcessID, "SendCloseMsg? " & SendCloseMsg

    Dim lr               As Long
    Dim hProc            As Long
    Dim hThread          As Long
    Dim dwThreadId       As Long
    Dim PID()            As Long
    Dim i                As Long
    Dim ProcNameCur      As String

    If Len(ProcessName) <> 0 Then

        PID() = GetPIDsByName(ProcessName)
        
        If IsArrDimmed(PID) Then
            For i = 1 To UBound(PID)
                ProcessClose , , True, 0, PID(i), SendCloseMsg 'Async mode
            Next
        End If
        
        If Not Async Then
            If TimeOutMs = INFINITE Then
                For i = 1 To UBound(PID)
                    ProcNameCur = ""
                    Call IsRunned(ProcNameCur, PID(i)) 'get name by PID
                    If StrComp(ProcNameCur, ProcessName, vbTextCompare) = 0 Then 'not changed ?
                        Call WaitForTerminate(, , False, INFINITE, PID(i))
                    End If
                Next
            Else
                For i = 1 To UBound(PID)
                    ProcNameCur = ""
                    Call IsRunned(ProcNameCur, PID(i)) 'get name by PID
                    If StrComp(ProcNameCur, ProcessName, vbTextCompare) = 0 Then 'not changed ?
                        ProcessID = PID(i)
                        Exit For ' get only 1 process
                    End If
                Next
                If ProcessID <> 0 Then Call WaitForTerminate(, , False, TimeOutMs, ProcessID)
            End If
        End If
        
        ProcessID = 0
        ProcessClose = IsRunned(ProcessName, ProcessID)
        Exit Function
        
    ElseIf WindowHandle <> 0 Then
        dwThreadId = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        lr = PostMessage(WindowHandle, WM_CLOSE, 0, 0)                           ' WM_CLOSE главному окну
        Call WaitForTerminate(, hProc, False, 1000)                              ' подождём 1 сек.
        lr = PostMessage(WindowHandle, WM_QUIT, 0, 0)                            ' WM_QUIT главному окну
        If Not Async Then
            lr = WaitForTerminate(WindowHandle, , False, TimeOutMs)              ' ожидание, без замораживания текущего потока
        Else
            lr = WaitForTerminate(WindowHandle, , False, 500)
        End If
        If IsRunned(, ProcessID) Then
            hProc = OpenProcess(PROCESS_TERMINATE, False, ProcessID)
        End If
        
    ElseIf ProcessID <> 0 Then
        hProc = OpenProcess(PROCESS_TERMINATE, False, ProcessID)
    
    Else
        hThread = pi.hThread
        ProcessID = pi.dwProcessId
        hProc = pi.hProcess
    End If
    
    If SendCloseMsg Then
        hThread = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        lr = PostMessage(WindowHandle, WM_CLOSE, 0, 0)                           ' WM_CLOSE главному окну
        Call WaitForTerminate(, hProc, False, 1000)
    End If
    
    If hProc <> 0 Then
        TerminateProcess hProc, 0
        If Not (Async) Then
            ProcessClose = WaitForTerminate(, hProc, False, TimeOutMs)
        End If
    End If
    
    ProcessClose = Not IsRunned(, ProcessID)                           ' Проверю наличие в списке процессов
    
    If ProcessClose Then
        If hThread <> 0 Then CloseHandle hThread
        If hProc <> 0 Then CloseHandle hProc
    End If
    
    AppendErrorLogCustom "ProcessClose - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_ProcessClose", "ProcessName:", ProcessName
    If inIDE Then Stop: Resume Next
End Function


Public Function ProcessRun( _
                            ByVal FileName As String, _
                            Optional ByVal CommandLine As String, _
                            Optional ByVal CurrentDirectory As String = vbNullString, _
                            Optional WindowStyle As Long = 1&, _
                            Optional CloseHandles As Boolean = False)
                            ' По-умолчанию стиль окна - SW_SHOWNORMAL
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ProcessRun - Begin", "FileName: " & FileName, "CommandLine: " & CommandLine, "CurDir: " & CurrentDirectory, _
        "WindowStyle: " & WindowStyle, "CloseHandles? " & CloseHandles
    
    Const STARTF_USESHOWWINDOW      As Long = 1
    Const NORMAL_PRIORITY_CLASS     As Long = &H20
    
    Dim n           As Long
    Dim lr          As Long
    Dim CMDLine     As String
    Dim argc        As Long
    Dim argv()      As String
    
    CMDLine = """" & FileName & """"                    ' Имя файла в кавычки
    
    If CommandLine <> vbNullString Then                 ' Если есть аргументы, их тоже добавляем в кавычках
    
        ParseCommandLine CommandLine, argc, argv        ' Разбиваем аргументы командной строки на составляющие
        For n = 1 To argc
            CMDLine = CMDLine & " """ & argv(n) & """"
        Next
    End If
    
    si.cb = Len(si)
    GetStartupInfo si                                   ' клонируем структуру текущего процесса
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = WindowStyle                        ' Стиль окна
    
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

    ProcessRun = lr    ' not 0 is SUCCESS
    
    If CloseHandles Then
        CloseHandle pi.hProcess
        CloseHandle pi.hThread
    End If
    
    AppendErrorLogCustom "ProcessRun - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_ProcessRun", "FileName:", FileName, "CommandLine:", CommandLine
    If inIDE Then Stop: Resume Next
End Function

Private Sub ParseCommandLine(Line As String, argc As Long, argv() As String)
    On Error GoTo ErrorHandler:
    Dim Ptr     As Long
    Dim index   As Long
    Dim strLen  As Long
    Dim strAdr  As Long
    
    Ptr = CommandLineToArgvW(StrPtr(Line), argc)
    ReDim argv(argc)
    argv(0) = App.Path & "\" & App.EXEName & ".exe"
    If argc = 0 Then Exit Sub
    
    For index = 1 To argc
        GetMem4 ByVal Ptr + (index - 1) * 4, strAdr
        strLen = lstrlen(strAdr)
        argv(index) = Space(strLen)
        lstrcpyn StrPtr(argv(index)), strAdr, strLen + 1
    Next
    GlobalFree Ptr
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsProcess_ParseCommandLine", "Line:", Line
    If inIDE Then Stop: Resume Next
End Sub

