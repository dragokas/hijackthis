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
    szExeFile As String * MAX_PATH
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

Private Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
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

Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoW" (lpStartupInfo As STARTUPINFO)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const PROCESS_ALL_ACCESS        As Long = &H1F0FFF
Private Const INFINITE                  As Long = -1
Private Const STILL_ACTIVE              As Long = &H103
Private Const WAIT_FAILED               As Long = -1
Private Const TH32CS_SNAPPROCESS        As Long = 2
Private Const SYNCHRONIZE               As Long = &H100000
Private Const PROCESS_QUERY_INFORMATION As Long = 1024&
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&

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
                    ProcessName = Left$(pe.szExeFile, InStr(pe.szExeFile, Chr(0)) - 1)
                    Exit Do
                End If
            Else
                If lstrcmpi(StrPtr(pe.szExeFile), StrPtr(ProcessName)) = 0 Then
                    IsRunned = True
                    ProcessID = pe.th32ProcessID
                    Exit Do
                End If
            End If
        Loop While Process32Next(hSnap, pe)
    End If
    CloseHandle hSnap
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsProcess_IsRunned"
    If inIDE Then Stop: Resume Next
End Function


Function WaitForTerminate( _
                        Optional ByVal WindowHandle As Long, _
                        Optional ByVal ProcessHandle As Long, _
                        Optional ByVal FreezeThread As Boolean, _
                        Optional ByVal TimeOut As Single) As Long 'msec
                        ' Ожидает завершения процесса                           'Thanks for help to Catstail (Б. Файфель)
    
    On Error GoTo ErrorHandler:
    
    ' FreezeThread, если true - заморозит текущий поток на время, пока отслеживаемый процесс не будет закрыт
    ' TimeOut, если FreezeThread = false, задает время таймаута ожидания завершения процесса.
    ' Возвращаемое значение: ExitCode процесса, или -1, если вышел таймаут ожидания или WaitForSingleObject вернул ошибку (WAIT_FAILED)
    
    Dim cRead       As Long
    Dim iExit       As Long
    Dim hProc       As Long
    Dim ProcessID   As Long
    Dim iResult     As Long
    Dim hThreadId   As Long
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
    
    If WindowHandle = 0 Then
        If ProcessHandle = 0 Then
            hProc = GetProcessHandle()
            hThread = GetThreadHandle()
        Else
            hProc = ProcessHandle
        End If
    Else
        '//TODO: get handle of thread
    
        ' получаем PID процесса
        hThreadId = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        '::: Получаем хэндл процесса
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or SYNCHRONIZE, False, ProcessID)
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
       DeadTime = InitialTime + TimeOut / 1000
       bTimeOutON = (TimeOut <> INFINITE)
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
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsProcess_WaitForTerminate"
    If inIDE Then Stop: Resume Next
End Function


Public Function ProcessClose( _
                            Optional WindowHandle As Long, _
                            Optional ProcessName As String, _
                            Optional Async As Boolean, _
                            Optional TimeOut As Single = INFINITE) As Long
                            'Возвращаемое значение = ExitCode процесса, или -1 в случае неудачи
    
    On Error GoTo ErrorHandler:

    'Const WM_CLOSE         As Long = &H10
    'Const WM_DESTROY       As Long = &H2
    Const WM_QUIT           As Long = &H12  ' wparam - exitcode
    Const PROCESS_TERMINATE As Long = &H1   ' access right

    Dim lr               As Long
    Dim hProc            As Long
    Dim ProcessID        As Long
    Dim hThreadId        As Long

    If ProcessName <> vbNullString Then
        If IsRunned(ProcessName, ProcessID) Then
            hProc = OpenProcess(PROCESS_TERMINATE, False, ProcessID)
            If hProc <> 0 Then
                TerminateProcess hProc, 0
                If Not (Async) Then
                    ProcessClose = WaitForTerminate(, hProc, False, TimeOut)
                End If
            End If
        End If
        Exit Function
    End If

    If WindowHandle <> 0 Then
        hThreadId = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        lr = PostMessage(WindowHandle, WM_QUIT, 0, 0)                            ' WM_QUIT главному окну
        If Not Async Then lr = WaitForTerminate(WindowHandle, , False, TimeOut)  ' ожидание, без замораживания текущего потока / Timeout в 1 сек.
    Else
        ProcessID = pi.dwProcessId
        TerminateProcess pi.hProcess, 0
    End If

    If lr = -1 Then                                                              ' если по хорошему закрыть не получилось
        If ProcessID <> 0 Then
            hProc = OpenProcess(PROCESS_TERMINATE, False, ProcessID)             ' получим Handle процесса
        End If
        If hProc <> 0 Then
            TerminateProcess hProc, 0                                            ' Terminate
        End If
        If Not Async Then lr = WaitForTerminate(WindowHandle, , False, TimeOut)  ' Снова подожду
    End If
    If lr <> -1 Then
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
    End If
    ProcessClose = IIf(IsRunned(, ProcessID), -1, lr)                            ' Проверю наличие в списке процессов
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsProcess_ProcessClose", "ProcessName:", ProcessName
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
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "clsProcess_ProcessRun", "FileName:", FileName, "CommandLine:", CommandLine
    If inIDE Then Stop: Resume Next
End Function

Private Sub ParseCommandLine(Line As String, argc As Long, argv() As String)
    On Error GoTo ErrorHandler:
    Dim ptr     As Long
    Dim Index   As Long
    Dim strLen  As Long
    Dim strAdr  As Long
    
    ptr = CommandLineToArgvW(StrPtr(Line), argc)
    ReDim argv(argc)
    argv(0) = App.Path & "\" & App.EXEName & ".exe"
    If argc = 0 Then Exit Sub
    
    For Index = 1 To argc
        GetMem4 ByVal ptr + (Index - 1) * 4, strAdr
        strLen = lstrlen(strAdr)
        argv(Index) = Space(strLen)
        lstrcpyn StrPtr(argv(Index)), strAdr, strLen + 1
    Next
    GlobalFree ptr
    Exit Sub
ErrorHandler:
    ErrorMsg err, "clsProcess_ParseCommandLine", "Line:", Line
    If inIDE Then Stop: Resume Next
End Sub

