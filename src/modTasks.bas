Attribute VB_Name = "modTasks"
'[modTasks.bas]

Option Explicit

' Windows Scheduled Tasks Enumerator/Killer by Alex Dragokas

' To add early binding, set reference to taskschd.dll (not applicable to EnumTasks2)

' keys
' Vista+: HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks
' XP:     HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler

'regexp replacing (for CLSID):
'^(.*?): (\(disabled\) )?(.*?) - (.*?) - (.*?)( \(Microsoft\))?
' \3; \4; \5

'https://docs.microsoft.com/en-us/windows/desktop/taskschd/task-scheduler-objects

'Suspected in high CPU loading and other problems:
'
'\Microsoft\Windows\TabletPC\InputPersonalization - C:\Program Files\Common Files\Microsoft Shared\Ink\InputPersonalization.exe (Microsoft)
'\Microsoft\Windows Live\SOXE\Extractor Definitions Update Task - {3519154C-227E-47F3-9CC9-12C3F05817F1} - C:\Program Files\Windows Live\SOXE\wlsoxe.dll (Microsoft)
'\Microsoft\Windows\rempl\shell - C:\Program Files\rempl\sedlauncher.exe - many complaints about waking up PC

' Action type constants
Private Enum TASK_ACTION_TYPE
    TASK_ACTION_EXEC = 0
    TASK_ACTION_COM_HANDLER = 5&
    TASK_ACTION_SEND_EMAIL = 6&
    TASK_ACTION_SHOW_MESSAGE = 7&
End Enum

Private Type TASK_ENTRY
    ActionType  As TASK_ACTION_TYPE
    RunObj      As String 'optional
    RunObjExpanded As String 'optional
    RunArgs     As String 'optional
    ClassID     As String 'optional
    ClassData   As String 'optional
    RunObjCom   As String 'optional
    WorkDir     As String 'optional
    Enabled     As Boolean
    FileMissing As Boolean
    RegID       As String
    GroupId     As String 'ComputerName\GroupName or GroupName or SID
    UserId      As String 'ComputerName\UserName or UserName or SID
End Type

Private Enum JOB_PRIORITY_CLASS
    JOB_REALTIME_PRIORITY = &H800000
    JOB_HIGH_PRIORITY = &H1000000
    JOB_IDLE_PRIORITY = &H2000000
    JOB_NORMAL_PRIORITY = &H4000000
End Enum

Private Enum JOB_SCHED_STATUS
    SCHED_S_TASK_READY = &H41300
    SCHED_S_TASK_RUNNING = &H41301
    SCHED_S_TASK_NOT_SCHEDULED = &H41305
End Enum

Private Type JOB_HEADER
    ProductVer As Integer
    FormatVer As Integer
    ID As UUID
    AppNameOffset As Integer
    TriggerOffset As Integer
    ErrRetryCount As Integer
    ErrRetryInterval As Integer
    IdleDeadline As Integer
    IdleWait As Integer
    Priority As JOB_PRIORITY_CLASS
    MaxRunTime As Long
    ExitCode As Long
    Status As JOB_SCHED_STATUS
    Flags As Long
    LastRunTime As SYSTEMTIME
End Type

Private Type JOB_UNICODE_STRING
    Length As Integer
    Data As String
End Type

Private Type JOB_USER_DATA
    Size As Integer
    Data() As Byte
End Type

Private Type JOB_RESERVED_DATA
    Size As Integer
    StartError As Long
    TaskFlags As Long
End Type

Private Enum JOB_TRIGGER_TYPE
    ONCE = 0
    DAILY = 1
    WEEKLY = 2
    MONTHLYDATE = 3
    MONTHLYDOW = 4
    EVENT_ON_IDLE = 5
    EVENT_AT_SYSTEMSTART = 6
    EVENT_AT_LOGON = 7
End Enum

Private Enum JOB_TRIGGER_FLAGS
    TASK_TRIGGER_FLAG_HAS_END_DATE = &H80000000
    TASK_TRIGGER_FLAG_KILL_AT_DURATION_END = &H40000000
    TASK_TRIGGER_FLAG_DISABLED = &H20000000
End Enum

Private Type JOB_TRIGGER
    TriggerSize As Integer
    Reserved1 As Integer
    BeginYear As Integer
    BeginMonth As Integer
    BeginDay As Integer
    EndYear As Integer
    EndMonth As Integer
    EndDay As Integer
    StartHour As Integer
    StartMinute As Integer
    MinutesDuration As Long
    MinutesInterval As Long
    Flags As Long ' sum of JOB_TRIGGER_FLAGS bits
    TriggerType As JOB_TRIGGER_TYPE
    TriggerSpecific0 As Integer
    TriggerSpecific1 As Integer
    TriggerSpecific2 As Integer
    Padding As Integer
    Reserved2 As Integer
    Reserved3 As Integer
End Type

Private Type JOB_TRIGGERS
    ccTriggers As Integer
    aTrigger() As JOB_TRIGGER
End Type

Private Type JOB_SIGNATURE
    SignVer As Integer
    MinClientVer As Integer
    Signature(63) As Byte
End Type

Private Type JOB_PROPERTY
    ccRunInstance As Integer
    AppName As JOB_UNICODE_STRING
    Parameters As JOB_UNICODE_STRING
    WorkDir As JOB_UNICODE_STRING
    Author  As JOB_UNICODE_STRING
    Comment  As JOB_UNICODE_STRING
    UserData As JOB_USER_DATA
    ReservedData As JOB_RESERVED_DATA
    Triggers As JOB_TRIGGERS
    JobSignature As JOB_SIGNATURE
End Type

Private Type JOB_FILE
    head As JOB_HEADER
    prop As JOB_PROPERTY
End Type

' Task state
Private Const TASK_STATE_RUNNING        As Long = 4&
Private Const TASK_STATE_QUEUED         As Long = 2&

' Include hidden tasks enumeration
Private Const TASK_ENUM_HIDDEN          As Long = 1&

Private LogHandle As Integer
Private BitsAdminExecuted As Boolean
Private cProcBitsAdmin As clsProcess

Public Function CreateTask(TaskName As String, FullPath As String, Arguments As String, Optional Description As String, Optional DelaySec As Long = 0) As Boolean
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "CreateTask - Begin"
    
    Dim pService            As Object
    Dim pRegInfo            As IRegistrationInfo
    Dim pRootFolder         As ITaskFolder
    Dim pTask               As ITaskDefinition
    Dim pPrincipal          As IPrincipal
    Dim pSettings           As ITaskSettings
    Dim pTriggerCollection  As ITriggerCollection
    Dim pTrigger            As ITrigger
    Dim pRegistrationTrigger As IRegistrationTrigger
    Dim pLogonTrigger       As ILogonTrigger
    Dim pActionCollection   As IActionCollection
    Dim pAction             As IAction
    Dim pExecAction         As IExecAction
    Dim pRegisteredTask     As IRegisteredTask
    
    Set pService = CreateObject("Schedule.Service")
    pService.Connect

    If (pService.Connected) Then
        
        Set pRootFolder = pService.GetFolder("\")
        On Error Resume Next
        pRootFolder.DeleteTask TaskName, 0 'delete task if it's already exist
        On Error GoTo ErrorHandler
        
        Set pTask = pService.NewTask(0)
        
        If Not (pTask Is Nothing) Then
        
            Set pRegInfo = pTask.RegistrationInfo
            pRegInfo.Author = GetCompName(ComputerNameNetBIOS) & "\" & OSver.UserName 'Username or %UserDomain%\%UserName%
            pRegInfo.Description = Description
            Set pRegInfo = Nothing
            
            Set pPrincipal = pTask.Principal
            pPrincipal.RunLevel = TASK_RUNLEVEL_HIGHEST
            Set pPrincipal = Nothing
            
            Set pSettings = pTask.Settings
            pSettings.Enabled = True
            'https://docs.microsoft.com/en-us/windows/desktop/taskschd/tasksettings-priority
            pSettings.Priority = 8 'BELOW_NORMAL_PRIORITY_CLASS
            Set pSettings = Nothing
            
            Set pTriggerCollection = pTask.Triggers
            Set pTrigger = pTriggerCollection.Create(TASK_TRIGGER_LOGON)
            Set pLogonTrigger = pTrigger
            
            'https://docs.microsoft.com/en-us/windows/desktop/TaskSchd/logontrigger-delay
            pLogonTrigger.Delay = "PT" & CStr(DelaySec) & "S" 'S - mean sec. delay 'format is PnYnMnDTnHnMnS, where T - is date/time separator
            pLogonTrigger.Enabled = True
            pLogonTrigger.UserId = GetCompName(ComputerNameNetBIOS) & "\" & OSver.UserName  'UserDomain\UserName, like "Alex-PC\Alex"
            Set pTrigger = Nothing
            Set pTriggerCollection = Nothing
            
            Set pActionCollection = pTask.Actions
            Set pAction = pActionCollection.Create(TASK_ACTION_EXEC)
            Set pExecAction = pAction
            pExecAction.Path = FullPath
            pExecAction.Arguments = Arguments
            pExecAction.WorkingDirectory = GetParentDir(FullPath)
            Set pAction = Nothing
            Set pActionCollection = Nothing
            
            Set pRegisteredTask = pRootFolder.RegisterTaskDefinition( _
                TaskName, pTask, TASK_CREATE_OR_UPDATE, "", "", TASK_LOGON_INTERACTIVE_TOKEN, "")
            
            If Not (pRegisteredTask Is Nothing) Then
                CreateTask = True
            End If
            
            Set pRegisteredTask = Nothing
            Set pTask = Nothing
        End If
        
        Set pService = Nothing
    End If
    
    AppendErrorLogCustom "CreateTask - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CreateTask"
    If inIDE Then Stop: Resume Next
End Function

'Public Sub EnumTasks(Optional MakeCSV As Boolean)
'    On Error GoTo ErrorHandler
'    AppendErrorLogCustom "EnumTasks - Begin"
'
'    Dim Stady As Long
'    Dim sLogFile As String
'
'    If GetServiceRunState("Schedule") <> SERVICE_RUNNING Then
'        Err.Raise 33333, , "Task scheduler service is not running!"
'        Exit Sub
'    End If
'
'    'Compatibility: Vista+
'
'    ' Create the TaskService object.
'    Dim Service As Object
'    Set Service = CreateObject("Schedule.Service")
'    Stady = 1
'    Service.Connect
'    Stady = 2
'
'    ' Get the root task folder that contains the tasks.
'    Dim rootFolder As ITaskFolder
'    Set rootFolder = Service.GetFolder("\")
'
'    bCreateLogFile = MakeCSV
'
'    If MakeCSV Then
'        LogHandle = FreeFile()
'        sLogFile = BuildPath(AppPath(), "Tasks.csv")
'        Open sLogFile For Output As #LogHandle
'        Print #LogHandle, "OSver" & ";" & "State" & ";" & "Name" & ";" & "Dir" & ";" & "RunObj" & ";" & "Args" & ";" & "Note" & ";" & "Error"
'    End If
'
'    Stady = 3
'    ' Recursively call for enumeration of current folder and all subfolders
'    EnumTasksInITaskFolder rootFolder
'
'    Set rootFolder = Nothing
'    Set Service = Nothing
'
'    If MakeCSV Then
'        Close #LogHandle
'        Shell "rundll32.exe shell32.dll,ShellExec_RunDLL " & """" & sLogFile & """", vbNormalFocus
'    End If
'
'    AppendErrorLogCustom "EnumTasks - End"
'    Exit Sub
'
'ErrorHandler:
'    ErrorMsg Err, "EnumTasks. Stady: " & Stady
'    If inIDE Then Stop: Resume Next
'End Sub
'
'Sub EnumTasksInITaskFolder(rootFolder As ITaskFolder)
'    On Error GoTo ErrorHandler:
'    AppendErrorLogCustom "EnumTasksInITaskFolder - Begin"
'
'    Dim Result      As SCAN_RESULT
'    Dim taskState   As String
'    Dim RunObj      As String
'    Dim RunObjExpanded As String
'    Dim RunArgs     As String
'    Dim DirParent   As String
'    Dim DirFull     As String
'    Dim sHit        As String
'    Dim NoFile      As Boolean
'    Dim isSafe      As Boolean
'    Dim ActionType  As Long
'    Dim taskFolder  As ITaskFolder
'    Dim SignResult  As SignResult_TYPE
'    Dim bIsMicrosoftFile As Boolean
'    Dim sWorkDir    As String
'
'    Dim nTask           As Long
'    Dim RunObjLast      As String
'    Dim RunArgsLast     As String
'    Dim taskStateLast   As String
'    'Dim DirParentLast   As String
'    Dim DirFullLast     As String
'    Dim lTaskState      As Long
'    Dim bTaskEnabled    As Boolean
'    '------------------------------
'    'Dim ComeBack        As Boolean
'    Dim Stady           As Single
'    Dim HRESULT         As String
'    Dim errN            As Long
'    Dim StadyLast       As Single
'    Dim RunObjCom       As String
'
'
'    'Debug.Print "Folder Name: " & rootFolder.Name
'    'Debug.Print "Folder Path: " & rootFolder.Path
'
'    Dim taskCollection As Object
'    Set taskCollection = rootFolder.GetTasks(TASK_ENUM_HIDDEN)
'    Stady = 1
'    AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'    Dim numberOfTasks As Long
'    numberOfTasks = taskCollection.Count
'    Stady = 2
'    AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'    Dim registeredTask  As IRegisteredTask
'    Dim taskDefinition  As ITaskDefinition
'    Dim taskAction      As IAction
'    Dim taskActionExec  As IExecAction
'    Dim taskActionEmail As IEmailAction
'    Dim taskActionMsg   As IShowMessageAction
'    Dim taskActionCOM   As IComHandlerAction
'    Dim taskActions     As IActionCollection
'
'    'Dim taskSettings3   As ITaskSettings3  'Win8+
'
'    On Error Resume Next
'
'    If numberOfTasks = 0 Then
'        'Debug.Print "No tasks are registered."
'        Stady = 3
'    Else
'        'Debug.Print "Number of tasks registered: " & numberOfTasks
'        Stady = 4
'
'        nTask = nTask + 1
'        For Each registeredTask In taskCollection
'
'            If Not bAutoLogSilent Then DoEvents
'
'            Err.Clear
'            Call LogError(Err, Stady, ClearAll:=True)
'
'            NoFile = False
'            isSafe = False
'
'            RunObjLast = RunObj
'            RunArgsLast = RunArgs
'            taskStateLast = taskState
'            'DirParentLast = DirParent
'            DirFullLast = DirFull
'            RunObj = ""
'            RunObjExpanded = ""
'            RunArgs = ""
'            taskState = "Unknown"
'            DirParent = ""
'            Stady = 5
'            lTaskState = 0
'            bTaskEnabled = False
'            RunObjCom = ""
'            sWorkDir = ""
'
'
'            DirFull = registeredTask.Path
'            Call LogError(Err, Stady)
'
'            'If DirFull = "\klcp_update" Then Stop
'
'            DirParent = GetParentDir(DirFull)
'            If 0 = Len(DirParent) Then DirParent = "{root}"
'
'            With registeredTask
'                'Debug.Print "Task Name: " & .Name
'                'Debug.Print "Task Path: " & .Path
'                Stady = 6
'                AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'                Err.Clear
'                Set taskDefinition = .Definition
'                Call LogError(Err, Stady)
'
'                If Err.Number = 0 Then
'
'                  Stady = 7
'                  AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'                  Set taskActions = taskDefinition.Actions
'                  Call LogError(Err, Stady)
'
'                  For Each taskAction In taskActions
'
'                    Stady = 8
'                    AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'                    ActionType = taskAction.type
'                    Call LogError(Err, Stady)
'
'                    Select Case ActionType
'
'                        Case TASK_ACTION_EXEC
'                            Stady = 9
'                            Set taskActionExec = taskAction
'                            'Debug.Print " Type: Executable"
'                            'Debug.Print "  Exec Path: " & taskActionExec.Path
'                            'Debug.Print "  Exec Args: " & taskActionExec.Arguments
'                            'Debug.Print "  Exec Type: " & taskActionExec.Type
'                            Stady = 10
'                            RunObj = taskActionExec.Path
'                            sWorkDir = EnvironW(taskActionExec.WorkingDirectory)
'                            Call LogError(Err, Stady)
'
'                            'RunObj = EnvironW(RunObj)
'                            RunObjExpanded = UnQuote(EnvironW(RunObj))
'
'                            If Mid$(RunObjExpanded, 2, 1) <> ":" Then
'                                If sWorkDir <> "" Then
'                                    RunObjExpanded = BuildPath(sWorkDir, RunObjExpanded)
'                                End If
'                            End If
'
'                            Stady = 11
'                            RunArgs = taskActionExec.Arguments
'                            Call LogError(Err, Stady)
'
'                            NoFile = (0 = Len(FindOnPath(RunObjExpanded)))
'
'                        Case TASK_ACTION_SEND_EMAIL
'                            Stady = 12
'                            'Debug.Print " Type: Email"
'                            Set taskActionEmail = taskAction
'                            'Debug.Print "  Recepient: " & taskActionEmail.To
'                            'Debug.Print "  Subject:   " & taskActionEmail.Subject
'                            RunObj = taskActionEmail.To & ", " & taskActionEmail.Subject
'                            Call LogError(Err, Stady)
'
'                        Case TASK_ACTION_SHOW_MESSAGE
'                            Stady = 13
'                            'Debug.Print " Type: Message Box"
'                            Set taskActionMsg = taskAction
'                            'Debug.Print "  Title: " & taskActionMsg.Title
'                            RunObj = taskActionMsg.Title
'                            Call LogError(Err, Stady)
'
'                        Case TASK_ACTION_COM_HANDLER
'                            Stady = 14
'                            'Debug.Print " Type: COM Handler"
'                            Set taskActionCOM = taskAction
'                            'Debug.Print "  ClassID: " & taskActionCOM.ClassId
'                            'Debug.Print "  Data:    " & taskActionCOM.Data
'                            RunObj = taskActionCOM.ClassID & IIf(Len(taskActionCOM.Data) <> 0, "," & taskActionCOM.Data, "")
'                            Call LogError(Err, Stady)
'
'                            'If InStr(taskActionCOM.ClassId, "{DE434264-8FE9-4C0B-A83B-89EBEEBFF78E}") <> 0 Then Stop
'
'                            RunObjCom = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & taskActionCOM.ClassID & "\InprocServer32", vbNullString)
'
'                            If RunObjCom = "" Then
'                                RunObjCom = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & taskActionCOM.ClassID & "\InprocServer32", vbNullString, True)
'                            End If
'
'                            If RunObjCom <> "" Then
'                                RunObjCom = FindOnPath(UnQuote(EnvironW(RunObjCom)), True)
'                            End If
'                    End Select
'
'                  Next
'                End If
'
'            End With
'
'            'BrokenTask will be under error ignor mode until log line on this cycle
'
'            Stady = 15
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'            Select Case registeredTask.state
'                Case "0"
'                    taskState = "Unknown"
'                Case "1"
'                    taskState = "Disabled"
'                Case "2"
'                    taskState = "Queued"
'                Case "3"
'                    taskState = "Ready"
'                Case "4"
'                    taskState = "Running"
'            End Select
'            Call LogError(Err, Stady)
'
'            Err.Clear
'            Stady = 16
'            lTaskState = registeredTask.state
'            Call LogError(Err, Stady)
'
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'            Stady = 17
'            bTaskEnabled = registeredTask.Enabled
'            Call LogError(Err, Stady)
'
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'            If Err.Number <> 0 Then
'                If taskState <> "Unknown" Then
'                    taskState = taskState & ", Unknown"
'                End If
'            Else
'                If lTaskState <> TASK_STATE_DISABLED _
'                    And bTaskEnabled = False Then
'                        taskState = taskState & ", Disabled"
'                End If
'            End If
'
'            Stady = 18
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'            'get last saved error
'            Call LogError(Err, StadyLast, errN, False)
'
'            HRESULT = ""
'            If errN <> 0 Then HRESULT = ErrMessageText(errN)
'
'            If bCreateLogFile Then
'                'taskState
'                Print #LogHandle, OSver.MajorMinor & ";" & "" & ";" & ScreenChar(registeredTask.Name) & ";" & ScreenChar(DirParent) & ";" & _
'                    ScreenChar(RunObj) & ";" & ScreenChar(RunArgs) & ";" & _
'                    IIf(NoFile, "(file missing)", "") & ";" & _
'                    IIf(0 <> Len(HRESULT), "(" & HRESULT & ", idx: " & StadyLast & ")", "") '& _
'                    'IIf(NoFile Or 0 <> Len(HRESULT), " <==== ATTENTION", "")
'            End If
'
'            Stady = 19
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'            If Len(RunObjExpanded) <> 0 Then RunObj = RunObjExpanded
'            RunObj = Replace$(RunObj, "\\", "\")
'
'            AppendErrorLogCustom "EnumTasksInITaskFolder: Checking - " & DirParent & "\" & registeredTask.Name
'
'            isSafe = isInTasksWhiteList(DirParent & "\" & registeredTask.Name, RunObj, RunArgs)
'
'            Stady = 19.1
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'            If ActionType = TASK_ACTION_EXEC Then
'                RunObj = PathNormalize(RunObj)
'                If isSafe Then
'                    SignVerify RunObj, SV_LightCheck Or SV_PreferInternalSign, SignResult
'                End If
'            Else
'                WipeSignResult SignResult
'            End If
'
'            Stady = 19.2
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'            ' Digital signature checking
'            If isSafe Then
'                AppendErrorLogCustom "[OK] EnumTasksInITaskFolder: WhiteListed."
'
'                'If Left$(RunObj, 1) <> "{" Then 'not CLSID-based task
'                If ActionType = TASK_ACTION_EXEC Then
'
'                    bIsMicrosoftFile = (SignResult.isMicrosoftSign And SignResult.isLegit)
'
'                    isSafe = (bIsMicrosoftFile And bHideMicrosoft And Not bIgnoreAllWhitelists)
'
'                    If Not isSafe Then
'                        If Not bIsMicrosoftFile Then
'                            AppendErrorLogCustom "[Failed] EnumTasksInITaskFolder: File - " & RunObj & " => is not Microsoft EDS !!! <======"
'                            Debug.Print "Task MS file has wrong EDS: " & RunObj
'                        End If
'
'                        Stady = 19.3
'                        AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'                        If FileExists(RunObj) Then NoFile = False
'
'                    End If
'                ElseIf ActionType = TASK_ACTION_COM_HANDLER And Len(RunObjCom) <> 0 Then
'
'                    SignVerify RunObjCom, SV_LightCheck Or SV_PreferInternalSign, SignResult
'
'                    bIsMicrosoftFile = (SignResult.isMicrosoftSign And SignResult.isLegit)
'
'                    isSafe = (bIsMicrosoftFile And bHideMicrosoft And Not bIgnoreAllWhitelists)
'
'                    If Not isSafe Then
'                        If Not bIsMicrosoftFile Then
'                            AppendErrorLogCustom "[Failed] EnumTasksInITaskFolder: File - " & RunObjCom & " => is not Microsoft EDS !!! <======"
'                            Debug.Print "Task MS file has wrong EDS: " & RunObjCom
'                        End If
'                    End If
'                End If
'            Else
'                AppendErrorLogCustom "[Failed] EnumTasksInITaskFolder: NOT WhiteListed !!! <======"
'            End If
'
'            If ActionType = TASK_ACTION_COM_HANDLER Then
'                If Len(RunObjCom) = 0 Then
'                    RunObj = RunObj & " - (no file)"
'                Else
'                    RunObj = RunObj & " - " & RunObjCom
'                    NoFile = Not FileExists(RunObjCom)
'                End If
'            End If
'
'            Stady = 19.4
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'            If Not isSafe Then
'
'              'sHit = "O22 - ScheduledTask: " & "(" & taskState & ") " & registeredTask.Name & " - " & DirParent & " - " & RunObj & _
'              '  IIf(Len(RunArgs) <> 0, " " & RunArgs, "") & _
'              '  IIf(NoFile, " (file missing)", "") & _
'              '  IIf(0 <> Len(SignResult.SubjectName) And SignResult.isLegit, " (" & SignResult.SubjectName & ")", "") & _
'              '  IIf(0 <> Len(HRESULT), " (" & HRESULT & ", idx: " & StadyLast & ")", "") '& _
'              '  'IIf(NoFile Or 0 <> Len(HRESULT), " <==== ATTENTION", "")
'
'              sHit = "O22 - Task " & "(" & taskState & "): " & _
'                IIf(DirParent = "{root}", registeredTask.Name, DirParent & "\" & registeredTask.Name)
'
''I temporarily remove EDS name in log
''              sHit = sHit & " - " & RunObj & _
''                IIf(Len(RunArgs) <> 0, " " & RunArgs, "") & _
''                IIf(NoFile, " (file missing)", "") & _
''                IIf(0 <> Len(SignResult.SubjectName) And SignResult.isLegit, " (" & SignResult.SubjectName & ")", "") & _
''                IIf(0 <> Len(HRESULT), " (" & HRESULT & ", idx: " & StadyLast & ")", "")
'
'              sHit = sHit & " - " & RunObj & _
'                IIf(Len(RunArgs) <> 0, " " & RunArgs, "") & _
'                IIf(NoFile, " (file missing)", "") & _
'                IIf(0 <> Len(HRESULT), " (" & HRESULT & ", idx: " & StadyLast & ")", "")
'
'              If Not IsOnIgnoreList(sHit) Then
'
'                Stady = 19.5
'                AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'                If g_bCheckSum Then
'                    If FileExists(RunObj) Then
'                        sHit = sHit & GetFileCheckSum(RunObj)
'                    End If
'                End If
'
'                Stady = 19.6
'                AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'                With Result
'                    .Section = "O22"
'                    .HitLineW = sHit
'                    '.RunObject = RunObj
'                    '.RunObjectArgs = RunArgs
'                    '.AutoRunObject = DirFull
'                    AddFileToFix .File, REMOVE_TASK, DirFull
'                    .CureType = CUSTOM_BASED
'                End With
'                AddToScanResults Result
'              End If
'            End If
'
'            Stady = 19.7
'            AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'        Next
'    End If
'
'    On Error GoTo ErrorHandler:
'
'    Stady = 20
'    AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'    Set taskActionExec = Nothing
'    Set taskActionEmail = Nothing
'    Set taskActionMsg = Nothing
'    Set taskActionCOM = Nothing
'    Set taskAction = Nothing
'    Set taskDefinition = Nothing
'    Set registeredTask = Nothing
'    Set taskCollection = Nothing
'
'    Stady = 21
'    AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'    Dim taskFolderCollection As ITaskFolderCollection
'    Set taskFolderCollection = rootFolder.GetFolders(0&)
'
'    Stady = 22
'    AppendErrorLogCustom "EnumTasksInITaskFolder", "Stady: " & Stady
'
'    For Each taskFolder In taskFolderCollection 'deep to subfolders
'        EnumTasksInITaskFolder taskFolder
'    Next
'
'    Set taskFolder = Nothing
'    Set taskFolderCollection = Nothing
'
'    AppendErrorLogCustom "EnumTasksInITaskFolder - End"
'    Exit Sub
'ErrorHandler:
'    ErrorMsg Err, "EnumTasksInITaskFolder. Stady: " & Stady & ". Number of tasks: " & numberOfTasks & ". Curr. task # " & nTask & ": " & DirFull & ", " & _
'        "RunObj = " & RunObj & ", RunArgs = " & RunArgs & ", taskState = " & taskState
'        '& ". ____Last task Data:___ " & DirFullLast & ", " & _
'        '"RunObjLast = " & RunObjLast & ", RunArgsLast = " & RunArgsLast & ", taskStateLast = " & taskStateLast
''    If ComeBack Then
''        ComeBack = False
''        If inIDE Then Stop
''        Return
''    End If
'    If inIDE Then Stop: Resume Next
'End Sub

Public Function PathNormalize(ByVal sPath As String) As String
    
    AppendErrorLogCustom "PathNormalize - Begin", "File: " & sPath
    
    Dim sTmp As String, bShouldSeek As Boolean
    
    sPath = UnQuote(sPath)
    
    If StrBeginWith(sPath, "!\??\") Then sPath = Mid$(sPath, 6)
    If StrBeginWith(sPath, "\??\") Then sPath = Mid$(sPath, 5)
    If StrBeginWith(sPath, "\\?\") Then sPath = Mid$(sPath, 5)
    If StrBeginWith(sPath, "\\.\") Then sPath = Mid$(sPath, 5)
    If StrBeginWith(sPath, "file:///") Then sPath = Mid$(sPath, 9)
    If StrBeginWith(sPath, "\SystemRoot\") Then sPath = sWinDir & Mid$(sPath, 12)
    
    If StrBeginWith(sPath, "\\") Then
        PathNormalize = sPath
        Exit Function
    End If
    
    If StrBeginWith(sPath, "\") Then sPath = SysDisk & sPath
    
    '???
    'sPath = Replace(sPath, "/", "\")
    
    If Mid$(sPath, 2, 1) <> ":" Then
        bShouldSeek = True  'relative or on the %PATH%
    Else
        If Not FileExists(sPath) Then bShouldSeek = True 'e.g. no extension
        sPath = GetLongPath(sPath)
    End If
    
    If bShouldSeek Then
        sTmp = FindOnPath(sPath)
        If Len(sTmp) <> 0 Then
            sPath = sTmp
        End If
    End If
    
    PathNormalize = sPath
    
    AppendErrorLogCustom "PathNormalize - End"
End Function

Public Function isInTasksWhiteList(sPathName As String, sTargetFile As String, Optional sArguments As String) As Boolean
    On Error GoTo ErrorHandler
    Dim WL_ID As Long

    If Not bHideMicrosoft Then Exit Function
    If Not oDict.TaskWL_ID.Exists(sPathName) Then
    
        'O22 - ScheduledTask: (Ready) User_Feed_Synchronization-{826B43E8-D4FF-4589-B639-B04CB653CCC1} - {root} - C:\Windows\system32\msfeedssync.exe sync
        
        If sPathName Like "{root}\User_Feed_Synchronization-{????????-????-????-????-????????????}" Then
            If StrComp(sTargetFile, sWinSysDir & "\msfeedssync.exe", 1) = 0 And sArguments = "sync" Then
                isInTasksWhiteList = True
            End If
        ElseIf sPathName Like "{root}\Optimize Start Menu Cache Files-S-1-5-21-*" Then
            If StrComp(sTargetFile, "{2D3F8A1B-6DCD-4ED5-BDBA-A096594B98EF},$(Arg0)", 1) = 0 Then
                isInTasksWhiteList = True
            End If
        ElseIf sPathName Like "\WPD\SqmUpload_S-1-5-21-*" Then
            If StrComp(sTargetFile, sWinDir & "\system32\rundll32.exe", 1) = 0 And sArguments = "portabledeviceapi.dll,#1" Then
                isInTasksWhiteList = True
            End If
        ElseIf sPathName Like "{root}\OneDrive Standalone Update Task-S-1-5-21-*" Then
            If StrComp(sTargetFile, LocalAppData & "\Microsoft\OneDrive\OneDriveStandaloneUpdater.exe", 1) = 0 And sArguments = "" Then
                isInTasksWhiteList = True
            End If
        ElseIf sPathName Like "\Microsoft\VisualStudio\Updates\UpdateConfiguration_S*" Then
            If StrComp(sTargetFile, PF_32 & "\Microsoft Visual Studio\Installer\resources\app\ServiceHub\Services\Microsoft.VisualStudio.Setup.Service\VSIXConfigurationUpdater.exe", 1) = 0 And sArguments = "" Then
                isInTasksWhiteList = True
            End If
        End If
        Exit Function
    End If
    
    WL_ID = oDict.TaskWL_ID(sPathName)
    
    Dim MyArray() As String
    Dim i As Long
    Dim bSafe As Boolean
    
    With g_TasksWL(WL_ID)
        'verifying all components
        
        MyArray = Split(.RunObj, "|")
        
        For i = 0 To UBound(MyArray)
            If Left$(MyArray(i), 1) <> "{" Then 'not CLSID-based ?
                If InStr(MyArray(i), "\") = 0 Then
                    'if full path not set in database, comparing by filename only
                    If StrComp(GetFileNameAndExt(sTargetFile), MyArray(i), vbTextCompare) = 0 Then bSafe = True: Exit For
                Else
                    If StrComp(sTargetFile, MyArray(i), vbTextCompare) = 0 Then bSafe = True: Exit For
                End If
            Else
                If StrComp(sTargetFile, MyArray(i), vbTextCompare) = 0 Then bSafe = True: Exit For
            End If
        Next
        
        If bSafe Then
            If inArraySerialized(sArguments, .Args, "|", , , 1) Then
                isInTasksWhiteList = True
            End If
        End If
    End With
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modTasks.isInTasksWhiteList"
    If inIDE Then Stop: Resume Next
End Function

' replace ; -> \; (for parsing purposes)
Function ScreenChar(sText As String) As String
    ScreenChar = " " & Replace$(sText, ";", "\;")
End Function

'Sub LogError(objError As ErrObject, in_out_Stady As Single, Optional out_LastLoggedErrorNumber As Long, Optional in_ActionPut As Boolean = True, Optional ClearAll As Boolean)
'    '
'    'in_ActionPut - if false, this function fill _out_ parameters with last saved error number and Stady position
'
'    'Purpose of function:
'    'Log first error and do not overwrite it until 'ClearAll' parameter become true
'
'    Static Stady As Long, ErrNum As Long
'
'    If ClearAll = True Then
'        Stady = 0
'        ErrNum = 0
'        Exit Sub
'    End If
'
'    If in_ActionPut = False Then
'        'get
'        in_out_Stady = Stady
'        out_LastLoggedErrorNumber = ErrNum
'    Else
'        'put
'        If ErrNum = 0 Then
'            ErrNum = objError.Number
'            Stady = in_out_Stady
'        End If
'    End If
'End Sub

'Public Function KillTask(TaskFullPath As String) As Boolean
'
'    '// TODO: Replace KillProcess by FreezeProcess
'
'    On Error GoTo ErrorHandler
'    Dim TaskPath As String
'    Dim TaskName As String
'    Dim pos As Long
'    Dim Stady As Long
'    'Dim ComeBack As Boolean
'    Dim lTaskState As Long
'    Dim BrokenTask As Boolean
'
'    'Compatibility: Vista+
'
'    pos = InStrRev(TaskFullPath, "\")
'    If pos <> 0 Then
'        Stady = 1
'        TaskPath = Left$(TaskFullPath, pos)
'        If Len(TaskPath) > 1 Then TaskPath = Left$(TaskPath, Len(TaskPath) - 1) 'trim last backslash
'        TaskName = Mid$(TaskFullPath, pos + 1)
'    Else
'        Exit Function
'    End If
'
'    On Error Resume Next
'    Stady = 2
'    ' Create the TaskService object.
'    Dim Service As Object
'    Set Service = CreateObject("Schedule.Service")
'    If Err.Number <> 0 Then Exit Function
'
'    Stady = 3
'    Service.Connect
'    If Err.Number <> 0 Then Exit Function
'
'    Stady = 4
'    ' Get the root task folder that contains the tasks.
'    Dim rootFolder As ITaskFolder
'    Set rootFolder = Service.GetFolder(TaskPath)
'    If Err.Number <> 0 Then Exit Function
'
'    Stady = 5
'    Dim registeredTask  As IRegisteredTask
'    Set registeredTask = rootFolder.GetTask(TaskName)
'    If Err.Number <> 0 Then Exit Function
'
'    'Dim taskCollection As Object
'    'Set taskCollection = rootFolder.GetTasks(TASK_ENUM_HIDDEN)
'    '
'    '    For Each registeredTask In taskCollection
'    '        With registeredTask
'    '          If InStr(1, .Path, TaskName, 1) <> 0 Then
'    '            Debug.Print "Task Name: " & .Name
'    '            Debug.Print "Task Path: " & .Path
'    '          End If
'    '
'    '        End With
'    '    Next
'    'Stop
'
'    ' Stop the task
'
'    'I should insert here this strange error handling routines because ITask Schedule interfaces with different
'    'kinds of possible errors in XML structures and unsufficient access rights caused by malware is a very poor stuff for developers,
'    'because it can produce so many unexpected errors. So, we are full of trubles.
'
'    'Maybe, later I rewrite it into manual parsing.
'
'    On Error Resume Next
'    Stady = 6
'    lTaskState = registeredTask.state
'    If Err.Number <> 0 Then
'        'ErrorMsg err, "KillTask. Stady: " & Stady
'        BrokenTask = True
'    End If
'
''    If err.Number <> 0 Then
''        ComeBack = True
''        GoSub ErrorHandler
''        registeredTask.Stop 0&
''        Sleep 2000&
''        BrokenTask = True
''    ElseIf lTaskState = TASK_STATE_RUNNING Or lTaskState = TASK_STATE_QUEUED Then
''        registeredTask.Stop 0&
''        Sleep 2000&
''    End If
'
''    If BrokenTask Or lTaskState = TASK_STATE_RUNNING Or lTaskState = TASK_STATE_QUEUED Then
''        registeredTask.Stop 0&
''        Sleep 2000&
''    End If
'
'    Stady = 7
'    If registeredTask.Enabled Then registeredTask.Enabled = False
'
'    Dim taskDefinition  As ITaskDefinition
'    Dim taskAction      As IAction
'    Dim taskActionExec  As IExecAction
'
'    Stady = 8
'
'    ' Kill process
'    Err.Clear
'    Set taskDefinition = registeredTask.Definition
'
'    If Err.Number <> 0 Then
'        If Not BrokenTask Then
'            'ErrorMsg err, "KillTask. Stady: " & Stady
'        End If
'    Else
'      Stady = 9
'      For Each taskAction In taskDefinition.Actions
'        Stady = 10
'        If TASK_ACTION_EXEC = taskAction.type Then
'            Stady = 11
'            Set taskActionExec = taskAction
'            'Debug.Print taskActionExec.Path
'            If FileExists(taskActionExec.Path) Then
'                KillProcessByFile taskActionExec.Path
'            End If
'        End If
'      Next
'    End If
'
'    'On Error GoTo ErrorHandler
'    On Error Resume Next
'    Err.Clear
'    Stady = 12
'    ' Remove the Job
'    rootFolder.DeleteTask TaskName, 0&
'    If Err.Number = 0 Then
'        Sleep 1000&
'        KillTask = True
'    End If
'
'    Stady = 13
'    Set taskActionExec = Nothing
'    Set taskAction = Nothing
'    Set taskDefinition = Nothing
'    Set registeredTask = Nothing
'    Set rootFolder = Nothing
'    Set Service = Nothing
'    Exit Function
'ErrorHandler:
'    ErrorMsg Err, "KillTask. Stady: " & Stady
'    If inIDE Then Stop: Resume Next
'End Function


Public Function DisableTask(TaskFullPath As String) As Boolean
    
    SetTaskState TaskFullPath, False
End Function

Public Function EnableTask(TaskFullPath As String) As Boolean
    
    SetTaskState TaskFullPath, True
End Function

Function SetTaskState(TaskFullPath As String, bEnable As Boolean)
    
    On Error GoTo ErrorHandler
    Dim TaskPath As String
    Dim TaskName As String
    Dim pos As Long
    
    'Compatibility: Vista+
    
    pos = InStrRev(TaskFullPath, "\")
    If pos <> 0 Then
        TaskPath = Left$(TaskFullPath, pos)
        If Len(TaskPath) > 1 Then TaskPath = Left$(TaskPath, Len(TaskPath) - 1) 'trim last backslash
        TaskName = Mid$(TaskFullPath, pos + 1)
    Else
        Exit Function
    End If
    
    On Error Resume Next
    ' Create the TaskService object.
    Dim Service As Object
    Set Service = CreateObject("Schedule.Service")
    If Err.Number <> 0 Then Exit Function

    Service.Connect
    If Err.Number <> 0 Then Exit Function

    ' Get the root task folder that contains the tasks.
    Dim rootFolder As ITaskFolder
    Set rootFolder = Service.GetFolder(TaskPath)
    If Err.Number <> 0 Then Exit Function

    Dim registeredTask  As IRegisteredTask
    Set registeredTask = rootFolder.GetTask(TaskName)
    If Err.Number <> 0 Then Exit Function
    
    If bEnable Then
        If Not registeredTask.Enabled Then
            Err.Clear
            registeredTask.Enabled = True
            If Err.Number = 0 Then SetTaskState = True
        Else
            SetTaskState = True
        End If
        
        Dim NewTask As IRegisteredTask
        Set NewTask = registeredTask.Run(Null)
    Else
        If registeredTask.Enabled Then
            Err.Clear
            registeredTask.Enabled = False
            If Err.Number = 0 Then SetTaskState = True
        Else
            SetTaskState = True
        End If
        
        registeredTask.Stop 0&
    End If
    
    Set registeredTask = Nothing
    Set rootFolder = Nothing
    Set Service = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SetTaskState. Enable? " & bEnable
    If inIDE Then Stop: Resume Next
End Function


' ####################################################################
' ####################################################################

'// Alternate version based on manual XML parsing
'// (no windows service involved)

Public Sub EnumTasksVista(Optional MakeCSV As Boolean)
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "EnumTasksVista - Begin"
    
    Dim dXmlPathFromDisk    As clsTrickHashTable
    Dim dTasksTreeGuid      As clsTrickHashTable
    Dim dTasksEmpty         As clsTrickHashTable
    
    Dim sLogFile        As String
    Dim i               As Long
    Dim j               As Long
    Dim result          As SCAN_RESULT
    Dim DirParent       As String
    Dim DirXml          As String
    Dim DirXml_2        As String
    Dim TaskName        As String
    Dim sHit            As String
    Dim NoFile          As Boolean
    Dim isSafe          As Boolean
    'Dim SignResult      As SignResult_TYPE
    Dim bIsMicrosoftFile As Boolean
    Dim aFiles()        As String
    Dim te()            As TASK_ENTRY
    Dim numTasks        As Long
    Dim sWinTasksFolder As String
    Dim bNoFile         As Boolean
    Dim aSubKeys()      As String
    Dim oKey            As Variant
    Dim ID              As String
    Dim bTelemetry      As Boolean
    Dim sRunFilename    As String
    Dim bActivation     As Boolean
    Dim bUpdate         As Boolean
    Dim sDllFile        As String
    
    '// TODO: Add record: "O22 - Task: 'Task scheduler' service is disabled!"
    
    'If GetServiceRunState("Schedule") <> SERVICE_RUNNING Then
        'Err.Raise 33333, , "Task scheduler service is not running!"
    'End If
    
    Set dXmlPathFromDisk = New clsTrickHashTable
    Set dTasksTreeGuid = New clsTrickHashTable
    Set dTasksEmpty = New clsTrickHashTable
    
    dXmlPathFromDisk.CompareMode = 1
    dTasksTreeGuid.CompareMode = 1
    dTasksEmpty.CompareMode = 1
    
    If MakeCSV Then
        LogHandle = FreeFile()
        sLogFile = BuildPath(AppPath(), "Tasks.csv")
        Open sLogFile For Output As #LogHandle
        Print #LogHandle, "OSver" & ";" & "State" & ";" & "Name" & ";" & "Dir" & ";" & "RunObj" & ";" & "Args" & ";" & "Note" & ";" & "Error"
    End If
    
    sWinTasksFolder = BuildPath(sWinSysDir, "Tasks")
    
    aFiles = ListFiles(sWinTasksFolder, "", True)
    
    If AryItems(aFiles) Then
      For i = 0 To UBound(aFiles)
      
        AppendErrorLogCustom "Task: " & aFiles(i)
        
        numTasks = AnalyzeTask(aFiles(i), te())
        
        DirXml = Mid$(aFiles(i), Len(sWinTasksFolder) + 1)
        
        'cache xml path
        dXmlPathFromDisk.Add DirXml, 0
        
        DirParent = GetParentDir(DirXml)
        If 0 = Len(DirParent) Then DirParent = "{root}"
        
        TaskName = GetFileNameAndExt(DirXml)
        
        UpdateProgressBar "O22", DirXml
        
        For j = 0 To numTasks - 1
            
            'WipeSignResult SignResult
            
            bIsMicrosoftFile = False
            
            If MakeCSV Then
                'taskState
                Print #LogHandle, OSver.MajorMinor & ";" & ScreenChar(DirParent & "\" & TaskName) & ";" & _
                    ScreenChar(te(j).RunObj) & ";" & _
                    IIf(te(j).ActionType = TASK_ACTION_EXEC, ScreenChar(te(j).RunArgs), ScreenChar(te(j).RunObjCom)) & ";" & _
                    IIf(te(j).FileMissing, "(file missing)", "") & ";"
            End If
            
            If Len(te(j).RunObjExpanded) <> 0 Then te(j).RunObj = te(j).RunObjExpanded
            te(j).RunObj = Replace$(te(j).RunObj, "\\", "\")
            
            'database checking
            
            If te(j).ActionType = TASK_ACTION_EXEC Then
                isSafe = isInTasksWhiteList(DirParent & "\" & TaskName, te(j).RunObj, te(j).RunArgs)
                
            ElseIf te(j).ActionType = TASK_ACTION_COM_HANDLER Then
                isSafe = isInTasksWhiteList(DirParent & "\" & TaskName, te(j).RunObj, te(j).RunObjCom)
            End If
            
            'If DirXml = "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" Then Stop
            
            'EDS checking for records, that doesn't exist in database
            
            If isSafe Or StrBeginWith(DirXml, "\Microsoft\") Then
            
                If te(j).ActionType = TASK_ACTION_EXEC Then
                
                    te(j).RunObj = PathNormalize(te(j).RunObj)
                    bIsMicrosoftFile = IsMicrosoftFile(te(j).RunObj)
                    
                ElseIf te(j).ActionType = TASK_ACTION_COM_HANDLER Then
                
                    te(j).RunObjCom = PathNormalize(te(j).RunObjCom)
                    bIsMicrosoftFile = IsMicrosoftFile(te(j).RunObjCom)
                End If
            End If
            
            ' EDS checking for items in database
            If isSafe Then
                AppendErrorLogCustom "[OK] EnumTasksInITaskFolder: WhiteListed."
                
                If te(j).ActionType = TASK_ACTION_EXEC Then
                    
                    isSafe = (bIsMicrosoftFile And bHideMicrosoft And Not bIgnoreAllWhitelists)
                    
                    If Not bIsMicrosoftFile Then
                        AppendErrorLogCustom "[Failed] EnumTasksInITaskFolder: File - " & te(j).RunObj & " => is not Microsoft EDS !!! <======"
                        Debug.Print "Task MS file has wrong EDS: " & te(j).RunObj
                    End If
                ElseIf te(j).ActionType = TASK_ACTION_COM_HANDLER And Len(te(j).RunObjCom) <> 0 Then
                
                  'for some reason, part of CLSID records on Win10 is not registered
                  If te(j).RunObjCom <> "(no file)" Then
                
                    bIsMicrosoftFile = IsMicrosoftFile(te(j).RunObjCom)
                    
                    isSafe = (bIsMicrosoftFile And bHideMicrosoft And Not bIgnoreAllWhitelists)
                    
                    If Not bIsMicrosoftFile Then
                        AppendErrorLogCustom "[Failed] EnumTasksInITaskFolder: File - " & te(j).RunObjCom & " => is not Microsoft EDS !!! <======"
                        Debug.Print "Task MS file has wrong EDS: " & te(j).RunObjCom
                    End If
                  End If
                End If
            Else
                AppendErrorLogCustom "[Failed] EnumTasksInITaskFolder: NOT WhiteListed !!! <======"
            End If
            
            bNoFile = False
            If te(j).ActionType = TASK_ACTION_COM_HANDLER Then
                If FileMissing(te(j).RunObjCom) Then
                    te(j).RunObj = te(j).RunObj & " - (no file)"
                    bNoFile = True
                Else
                    te(j).RunObj = te(j).RunObj & " - " & te(j).RunObjCom
                    te(j).FileMissing = Not FileExists(te(j).RunObjCom)
                End If
            End If
            
            '
            ' Broken tasks scan
            '
            ' Stady 1. Task is impersonated with missing user/group
            '
            If Not IsValidTaskUserId(te(j).UserId) And Not IsValidTaskGroupId(te(j).GroupId) Then
            
                sHit = "O22 - Task: (damaged) " & IIf(te(j).Enabled, "", "(disabled) ") & _
                    IIf(DirParent = "{root}", TaskName, DirParent & "\" & TaskName)

                sHit = sHit & " - " & te(j).RunObj & _
                    IIf(Len(te(j).RunArgs) <> 0, " " & te(j).RunArgs, "") & _
                    IIf(te(j).FileMissing And Not bNoFile, " (file missing)", "") & _
                    IIf(bIsMicrosoftFile, " (Microsoft)", "") & _
                    " (user missing)"
                  
                  If Not IsOnIgnoreList(sHit) Then
                    
                      If g_bCheckSum Then
                          If FileExists(te(j).RunObj) Then
                              sHit = sHit & GetFileCheckSum(te(j).RunObj)
                          End If
                      End If
                      
                      With result
                          .Section = "O22"
                          .HitLineW = sHit
                          .Name = DirXml 'used in "Disable" stuff
                          .State = IIf(te(j).Enabled, ITEM_STATE_ENABLED, ITEM_STATE_DISABLED)
                          
                          AddFileToFix .File, REMOVE_FILE, aFiles(i)
                          
                          te(j).RegID = Reg.GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirXml, "Id")
                          If te(j).RegID <> "" Then
                              AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & te(j).RegID
                              AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & te(j).RegID
                              AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & te(j).RegID
                              AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & te(j).RegID
                          End If
                          AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirXml
                          
                          .CureType = REGISTRY_BASED Or FILE_BASED
                      End With
                      AddToScanResults result
                  End If
                  'continue scan
            End If
            
            If Not isSafe Then
                
                bTelemetry = False
                bActivation = False
                bUpdate = False
                
                sRunFilename = GetFileName(te(j).RunObj, True)

                If InStr(1, DirParent, "Customer Experience", 1) <> 0 Then
                    bTelemetry = True
                ElseIf InStr(1, DirParent, "Application Experience", 1) <> 0 Then
                    bTelemetry = True
                ElseIf InStr(1, DirParent, "telemetry", 1) <> 0 Then
                    bTelemetry = True
                ElseIf InStr(1, TaskName, "telemetry", 1) <> 0 Then
                    bTelemetry = True
                ElseIf StrComp(sRunFilename, "NvTmMon.exe", 1) = 0 Then
                    bTelemetry = True
                ElseIf StrComp(sRunFilename, "OLicenseHeartbeat.exe", 1) = 0 Then
                    bTelemetry = True
                ElseIf StrComp(DirXml, "\Microsoft\Windows\IME\SQM data sender", 1) = 0 Then
                    bTelemetry = True
                ElseIf StrComp(sRunFilename, "NvTmRep.exe", 1) = 0 Then
                    bTelemetry = True
                ElseIf StrComp(sRunFilename, "NvTmMon.exe", 1) = 0 Then
                    bTelemetry = True
                End If
                
                If bIsMicrosoftFile Then
                
                    If InStr(1, DirParent, "Activation", 1) <> 0 Then
    
                        '\Microsoft\Windows\Windows Activation Technologies\ValidationTask - C:\Windows\system32\Wat\WatAdminSvc.exe /run (Microsoft)
                        bActivation = True
                        
                        'ElseIf StrComp(sRunFilename, "schtasks.exe", 1) = 0 Then
                        '    'C:\Windows\system32\schtasks.exe /run /I /TN "\Microsoft\Windows\Windows Activation Technologies\ValidationTask"
                        '    bActivation = True
                        
                    ElseIf InStr(1, DirParent, "gwx", 1) <> 0 Then '(Get Windows 10)
                        '\Microsoft\Windows\Setup\gwx\refreshgwxconfig - C:\Windows\system32\GWX\GWXConfigManager.exe /RefreshConfig (Microsoft)
                        '\Microsoft\Windows\Setup\gwx\refreshgwxcontent - C:\Windows\system32\GWX\GWXConfigManager.exe /RefreshContent (Microsoft)
                        '\Microsoft\Windows\Setup\gwx\runappraiser - C:\Windows\system32\GWX\GWXConfigManager.exe /RunAppraiser (Microsoft)
                        
                        bUpdate = True
                        
                    ElseIf StrComp(DirParent, "\Microsoft\Windows\Setup", 1) = 0 Then
                        If StrComp(sRunFilename, "EOSNotify.exe", 1) = 0 Then
                            bUpdate = True
                            'native OS feature to disable EOS notification
                            If 1 = Reg.GetDword(0, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\EOSNotify", "DiscontinueEOS") Then
                                isSafe = True
                            ElseIf OSver.MajorMinor <= 6.1 Then '(Windows 7 End of Support)
                                'always report
                            ElseIf OSver.MajorMinor <= 6.3 Then '(Windows 8 / 8.1 End of Support)
                                If OSver.IsEmbedded Then
                                    If Now() < #11/7/2023# Then 'November 7, 2023 - Win 8.1 Embedded
                                        isSafe = True
                                    End If
                                Else
                                    If Now() < #1/10/2023# Then 'January 10, 2023 - Win 8.1
                                        isSafe = True
                                    End If
                                End If
                            End If
                        End If
                    ElseIf StrComp(DirParent, "\Microsoft\Windows\End Of Support", 1) = 0 Then
                        If StrComp(sRunFilename, "sipnotify.exe", 1) = 0 Then
                            bUpdate = True
                            'native OS feature to disable EOS notification
                            If 1 = Reg.GetDword(0, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\SipNotify", "DontRemindMe") Then
                                isSafe = True
                            ElseIf OSver.MajorMinor <= 6.1 Then '(Windows 7 End of Support)
                                'always report
                            ElseIf OSver.MajorMinor <= 6.3 Then '(Windows 8 / 8.1 End of Support)
                                If OSver.IsEmbedded Then
                                    If Now() < #11/7/2023# Then 'November 7, 2023 - Win 8.1 Embedded
                                        isSafe = True
                                    End If
                                Else
                                    If Now() < #1/10/2023# Then 'January 10, 2023 - Win 8.1
                                        isSafe = True
                                    End If
                                End If
                            End If
                        End If
                    ElseIf StrComp(sRunFilename, "MusNotification.exe", 1) = 0 Then
                        '"Updates are available" window
                        'https://superuser.com/questions/972038/how-to-get-rid-of-updates-are-available-message-in-windows-10
                        bUpdate = True
                    ElseIf StrComp(sRunFilename, "UpdateAssistant.exe", 1) = 0 Then
                        bUpdate = True
                    End If
                
                End If
                
                If Not isSafe Then
                    
                    If StrBeginWith(DirXml, "\MicrosoftEdgeUpdateTaskMachine") Then
                        If StrComp(te(j).RunObj, PF_32 & "\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe", 1) = 0 Then
                            If IsMicrosoftFile(te(j).RunObj) Then
                                isSafe = True
                                bIsMicrosoftFile = True
                            End If
                        End If
                    ElseIf StrComp(DirXml, "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan", 1) = 0 Then
                        If StrComp(sRunFilename, "MpCmdRun.exe", 1) = 0 Then
                            If IsMicrosoftFile(te(j).RunObj) Then
                                isSafe = True
                                bIsMicrosoftFile = True
                            End If
                        End If
                    End If
                    
                End If
                
                'If InStr(1, DirParent, "Setup", 1) <> 0 Then
                '    Debug.Print "Folder = " & DirParent
                '    Debug.Print "File = " & sRunFilename
                'End If
                
                If (Not isSafe) Or (Not bHideMicrosoft) Then
                  'skip signature mark for LoLBin processes
                  If bIsMicrosoftFile Then
                  
                      If IsLoLBin(sRunFilename) Then
                      
                        bIsMicrosoftFile = False
                      
                        If StrComp(sRunFilename, "rundll32.exe", 1) = 0 Then
                            
                            sDllFile = GetRundllFile(te(j).RunArgs)
                            
                            If Len(sDllFile) = 0 Then te(j).FileMissing = True
                            
                            bIsMicrosoftFile = IsMicrosoftFile(sDllFile)
                            
                        ElseIf StrComp(sRunFilename, "cmd.exe", 1) = 0 Then
                            
                            sDllFile = GetWScriptFile(te(j).RunArgs)
                            
                            If StrComp(sDllFile, sWinSysDir & "\silcollector.cmd", 1) = 0 Then
                                
                                If GetFileSHA1(sDllFile, , True) = "22E925EBDAD83C6B1D1928DBB7369DA5E80A4DEB" Then bIsMicrosoftFile = True
                            End If
                            
                        ElseIf StrComp(sRunFilename, "cscript.exe", 1) = 0 Then
                            
                            If StrComp(sDllFile, sWinSysDir & "\calluxxprovider.vbs", 1) = 0 Then
                            
                                If GetFileSHA1(sDllFile, , True) = "5FFF9BFB3C918BC04D9BFAAB63989198FE57BDDA" Then bIsMicrosoftFile = True
                            End If
                            
                        End If
                      End If
                      
                      'If SignResult.isMicrosoftSign And Len(te(j).RunArgs) <> 0 Then
                      If bIsMicrosoftFile And Len(te(j).RunArgs) <> 0 Then
                          If InStr(1, te(j).RunArgs, "http", 1) <> 0 Then
                              'SignResult.isMicrosoftSign = False
                              bIsMicrosoftFile = False
                          ElseIf InStr(1, te(j).RunArgs, "ftp", 1) <> 0 Then
                              'SignResult.isMicrosoftSign = False
                              bIsMicrosoftFile = False
                          ElseIf InStr(1, EnvironW(te(j).RunArgs), "http", 1) <> 0 Then
                              'SignResult.isMicrosoftSign = False
                              bIsMicrosoftFile = False
                          ElseIf InStr(1, EnvironW(te(j).RunArgs), "ftp", 1) <> 0 Then
                              'SignResult.isMicrosoftSign = False
                              bIsMicrosoftFile = False
                          End If
                      End If
                  End If
                  
                  sHit = "O22 - Task: " & IIf(te(j).Enabled, "", "(disabled) ") & _
                    IIf(bTelemetry, "(telemetry) ", "") & _
                    IIf(bActivation, "(activation) ", "") & _
                    IIf(bUpdate, "(update) ", "") & _
                    IIf(DirParent = "{root}", TaskName, DirParent & "\" & TaskName)
                  
                  sHit = sHit & " - " & te(j).RunObj & _
                    IIf(Len(te(j).RunArgs) <> 0, " " & te(j).RunArgs, "") & _
                    IIf(te(j).FileMissing And Not bNoFile, " (file missing)", "") & _
                    IIf(bIsMicrosoftFile, " (Microsoft)", "")
                  
                  If Not IsOnIgnoreList(sHit) Then
                
                      If g_bCheckSum Then
                          If FileExists(te(j).RunObj) Then
                              sHit = sHit & GetFileCheckSum(te(j).RunObj)
                          End If
                      End If
                      
                      With result
                          .Section = "O22"
                          .HitLineW = sHit
                          .Name = DirXml 'used in "Disable" stuff
                          .State = IIf(te(j).Enabled, ITEM_STATE_ENABLED, ITEM_STATE_DISABLED)
                          
                          AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, te(j).RunObj 'should go first! Uses by VT.
                          AddFileToFix .File, REMOVE_FILE, aFiles(i)
                          
                          te(j).RegID = Reg.GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirXml, "Id")
                          If te(j).RegID <> "" Then
                              AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & te(j).RegID
                              AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & te(j).RegID
                              AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & te(j).RegID
                              AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & te(j).RegID
                          End If
                          AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirXml
                          
                          AddProcessToFix .Process, FREEZE_OR_KILL_PROCESS, aFiles(i)
                          
                          .CureType = FILE_BASED Or REGISTRY_BASED Or PROCESS_BASED
                      End With
                      AddToScanResults result
                  End If
                End If
            End If
            
        Next
        
      Next
    End If
    
    ' Searching for missing or damaged tasks:
    
    ' Notice:
    '
    ' dXmlPathFromDisk - is a real list of all xml Paths, gathered with FS function
    '
    
    'Stady 2
    '
    ' TaskCache\Tasks\{GUID} => "Path" pointing to missing xml
    '
    Dim aID() As String
    
    For i = 1 To Reg.EnumSubKeysToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks", aID())
    
        DirXml = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & aID(i), "Path")
        
        If Not dXmlPathFromDisk.Exists(DirXml) Then
            
            sHit = "O22 - Task: (damaged) HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & aID(i) & _
                " - " & DirXml & " (no xml)"
            
            If Not IsOnIgnoreList(sHit) Then
            
                With result
                    .Section = "O22"
                    .HitLineW = sHit
                    
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & aID(i)
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & aID(i)
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & aID(i)
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & aID(i)
                    
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirXml
                    
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    'Stady 3
    '
    ' \Tree\ has missing TaskCache\Tasks\{GUID} record
    '
    Dim nIndex As Long
    Erase aSubKeys

    For i = 1 To Reg.EnumSubKeysToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree", aSubKeys(), , , True)
        
        ID = Reg.GetString(HKEY_LOCAL_MACHINE, aSubKeys(i), "Id")
        
        If Not dTasksTreeGuid.Exists(ID) Then dTasksTreeGuid.Add ID, 0&
        
        'if a final task key (not a subkey)
        If Len(ID) <> 0 Then
            
            nIndex = -1

            If Reg.ValueExists(HKEY_LOCAL_MACHINE, aSubKeys(i), "Index") Then

                nIndex = Reg.GetDword(HKEY_LOCAL_MACHINE, aSubKeys(i), "Index")

            End If
            
            '
            ' "Index" is undoc.
            ' 0 - means that the task does not point to TaskCache\Tasks\{GUID}, and does not have xml
            '
            If nIndex <> 0 Then
                
                If Not Reg.KeyExists(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & ID) Then
                
                    sHit = "O22 - Task: (damaged) HKLM\" & aSubKeys(i) & " (key missing)"
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O22"
                            .HitLineW = sHit
        
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & ID
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & ID
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & ID
                            
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, aSubKeys(i)
                            
                            .CureType = REGISTRY_BASED
                        End With
                        
                        AddToScanResults result
                    End If
                End If
            End If
        End If
    Next
    
    'Stady 4
    '
    ' \TaskCache\Tasks\{GUID} doesn't have any corresponding records in \TaskCache\Tree pointing to
    '
    Erase aSubKeys
    
    For i = 1 To Reg.EnumSubKeysToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks", aSubKeys(), , , False)
        
        If Not dTasksTreeGuid.Exists(aSubKeys(i)) Then
        
            ID = aSubKeys(i)
        
            DirXml = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & ID, "Path")
            
            sHit = "O22 - Task: (damaged) HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & ID & _
                " (no key)"
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O22"
                    .HitLineW = sHit
                    
                    If Len(DirXml) <> 0 Then
                        AddFileToFix .File, REMOVE_FILE, BuildPath(sWinSysDir, "Tasks", DirXml)
                    End If
                    
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & ID
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & ID
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & ID
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & ID
                    
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                
                AddToScanResults result
            End If
            
        End If
    Next
    
    'Stady 5a
    '
    ' \TaskCache\Tree\ {root} has ghosts (empty) keys
    '
    Erase aSubKeys
    
    For i = 1 To Reg.EnumSubKeysToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree", aSubKeys(), , , True)
        
        DirXml = Mid$(aSubKeys(i), Len("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree") + 1)

        If Not StrBeginWith(DirXml, "\Microsoft") Then ' just in case
        
            If Not Reg.KeyHasSubKeys(HKLM, aSubKeys(i)) Then 'is root?
        
                If Not Reg.ValueExists(HKLM, aSubKeys(i), "Id") Then

                    dTasksEmpty.Add DirXml, 0&
                
                    sHit = "O22 - Task: (damaged) HKLM\" & aSubKeys(i) & _
                        " (empty)"
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O22"
                            .HitLineW = sHit
                            
                            AddFileToFix .File, REMOVE_FILE, BuildPath(sWinSysDir, "Tasks", DirXml)
                            AddFileToFix .File, REMOVE_FOLDER, BuildPath(sWinSysDir, "Tasks", DirXml)
                            
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, aSubKeys(i)
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        
                        AddToScanResults result
                    End If
                End If
            End If
        End If
    Next
    
    'Stady 5b
    '
    ' Empty folders in %SystemRoot%\System32\Tasks\
    '
    Dim aFolders() As String
    Dim sMicrosoft$: sMicrosoft = BuildPath(sWinSysDir, "Tasks\Microsoft")
    
    aFolders = ListSubfolders(BuildPath(sWinSysDir, "Tasks"), True)
    
    For i = 0 To UBoundSafe(aFolders)
        
        If Not StrBeginWith(aFolders(i), sMicrosoft) Then 'just in case
        
            If Not FolderHasSubfolder(aFolders(i)) Then 'is root?
            
                If Not FolderHasFile(aFolders(i)) Then 'no tasks?
                    
                    DirXml = Mid$(aFolders(i), Len(BuildPath(sWinSysDir, "Tasks")) + 1)
                    
                    If Not dTasksEmpty.Exists(DirXml) Then 'filter the record that is already in results list
                        
                        If GetFileName(aFolders(i), True) <> "WPD" Then
                        
                            sHit = "O22 - Task: (damaged) " & aFolders(i) & " (empty)"
                        
                            If Not IsOnIgnoreList(sHit) Then
                                With result
                                    .Section = "O22"
                                    .HitLineW = sHit
                                    
                                    AddFileToFix .File, REMOVE_FOLDER, aFolders(i)
                                    
                                    .CureType = FILE_BASED
                                End With
                                
                                AddToScanResults result
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    'Stady X
    '
    ' Check "Hash" key
    '
    ' info from @Leonid Makarov:
    ' Contains a SHA-256 or CRC32, before KB2305420. A byte-order-mark at beginning of the file is not included in the calculation of the hash.
    '
    ' See: https://github.com/libyal/winreg-kb/blob/main/documentation/Task%20Scheduler%20Keys.asciidoc
    '
    ' TODO: temporarily delayed. Reason: CRC check will spend a time
    
    'Stady X
    '
    ' Find "Triggers" inconsistency
    '
    ' TODO: undoc. key
    
    If MakeCSV Then
        Close #LogHandle
        Shell "rundll32.exe shell32.dll,ShellExec_RunDLL " & """" & sLogFile & """", vbNormalFocus
    End If
    
    Set dXmlPathFromDisk = Nothing
    Set dTasksTreeGuid = Nothing
    Set dTasksEmpty = Nothing
    
    AppendErrorLogCustom "EnumTasksVista - End"
    Exit Sub

ErrorHandler:
    ErrorMsg Err, "EnumTasksVista"
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetRundllFile(ByVal sArg As String) As String

    Dim pos&
    
    If Left$(sArg, 1) = "/" Then 'begin with switch?
    
        pos = InStr(sArg, " ")
        If pos = 0 Then Exit Function
        
        sArg = Mid$(sArg, pos + 1)
    End If
    
    pos = InStr(sArg, ",") ' skip the verb
    If pos <> 0 Then sArg = Left$(sArg, pos - 1)
    
    GetRundllFile = FindOnPath(sArg)

End Function

' Simple parsing of: WScrpt's arguments, like: //nologo "C:\path\some.js" -arg
'
' @result: C:\path\some.js
'
Public Function GetWScriptFile(ByVal sArguments As String) As String

    Dim pos As Long, argc As Long, i As Long
    Dim argv() As String
    
    ParseCommandLine sArguments, argc, argv
    
    For i = 1 To argc
        argv(i) = UnQuote(argv(i))
        If Mid$(argv(i), 2, 2) = ":\" Then
            GetWScriptFile = argv(i)
            Exit For
        End If
    Next

End Function

Private Function AnalyzeTask(sFilename As String, te() As TASK_ENTRY) As Long
    'additional info:
    'https://github.com/libyal/winreg-kb/blob/master/documentation/Task%20Scheduler%20Keys.asciidoc
    
    On Error GoTo ErrorHandler:
    
    Dim xmlDoc          As XMLDocument
    Dim xmlElement      As CXmlElement
    'Dim xmlAttribute    As CXmlAttribute
    Dim bExecBased      As Boolean
    Dim sFileData       As String
    Dim sTmp            As String
    Dim sTaskStatus     As String
    Dim i As Long
    Dim j As Long
    'Dim k As Long
    
    ReDim te(0)
    
    sFileData = ReadFileContents(sFilename, FileGetTypeBOM(sFilename) = 1200)
    
    If Len(sFileData) = 0 Then Exit Function
    
    Set xmlDoc = New XMLDocument
    Call xmlDoc.LoadData(sFileData)
    
    Set xmlElement = xmlDoc.NodeByName("Principals")
    
    If Not (xmlElement Is Nothing) Then
        For i = 1 To xmlElement.NodeCount
            If 0 = StrComp(xmlElement.Node(i).Name, "Principal", 1) Then
'                te(j).PrincipalId = vbNullString
'                For k = 1 To xmlElement.Node(i).AttributeCount
'                    Set xmlAttribute = xmlElement.Node(i).ElementAttribute(k)
'                    If 0 = StrComp(xmlAttribute.KeyWord, "id", 1) Then
'                        te(j).PrincipalId = xmlAttribute.Value
'                    End If
'                Next
                te(j).GroupId = xmlElement.Node(i).NodeValueByName("GroupId")
                te(j).UserId = xmlElement.Node(i).NodeValueByName("UserId")
                
                'Debug.Print te(j).PrincipalId & " = " & te(j).UserId & "( " & te(j).GroupId & " )"
            End If
        Next
    End If
    
    Set xmlElement = xmlDoc.NodeByName("Actions")
    
    If Not (xmlElement Is Nothing) Then
        For i = 1 To xmlElement.NodeCount
            If 0 = StrComp(xmlElement.Node(i).Name, "Exec", 1) Then
                te(j).RunObj = xmlElement.Node(i).NodeValueByName("Command")
                If te(j).RunObj <> "" Then
                    te(j).WorkDir = xmlElement.Node(i).NodeValueByName("WorkingDirectory")
                    te(j).RunArgs = xmlElement.Node(i).NodeValueByName("Arguments")
                    te(j).RunObjExpanded = UnQuote(EnvironW(te(j).RunObj))
                    te(j).RunArgs = EnvironW(te(j).RunArgs)
                    te(j).WorkDir = EnvironW(te(j).WorkDir)
                    
                    If Len(te(j).RunObjExpanded) = 0 Then
                        te(j).FileMissing = True
                    Else
                        If Mid$(te(j).RunObjExpanded, 2, 1) <> ":" Then
                            If te(j).WorkDir <> "" Then
                                sTmp = BuildPath(te(j).WorkDir, te(j).RunObjExpanded)
                                If FileExists(sTmp) Then te(j).RunObjExpanded = sTmp 'if only file exists in this work. folder
                            End If
                        End If
                        
                        te(j).RunObjExpanded = FindOnPath(te(j).RunObjExpanded, True) 'otherwise, search on %PATH%
                        te(j).FileMissing = Not FileExists(te(j).RunObjExpanded)
                    
                        If te(j).FileMissing And sTmp <> "" Then 'not exists at all? -> use initial Work.Dir + File, if present one
                            te(j).RunObjExpanded = sTmp
                        End If
                    End If
                    
                    te(j).ActionType = TASK_ACTION_EXEC
                    bExecBased = True
                    j = j + 1
                    ReDim Preserve te(j)
                End If
            End If
        Next
        
        If Not bExecBased Then
            For i = 1 To xmlElement.NodeCount
                If 0 = StrComp(xmlElement.Node(i).Name, "ComHandler", 1) Then
                    te(j).ClassID = xmlElement.Node(i).NodeValueByName("ClassId")
                    te(j).ClassData = xmlElement.Node(i).NodeValueByName("Data")
                    If te(j).ClassID <> "" Then
                        te(j).RunObj = te(j).ClassID & IIf(Len(te(j).ClassData) <> 0, "," & te(j).ClassData, "")
                        
                        Call GetFileByCLSID(te(j).ClassID, te(j).RunObjCom)
                        te(j).ActionType = TASK_ACTION_COM_HANDLER
                        j = j + 1
                        ReDim Preserve te(j)
                    End If
                End If
            Next
            
            'If te(j).ActionType <> TASK_ACTION_COM_HANDLER Then
                '// TODO email / msg

            'End If
        End If
    End If
    
    Set xmlElement = Nothing
    
    sTaskStatus = xmlDoc.NodeValueByName("Settings\Enabled")
    
    If Len(sTaskStatus) = 0 Then
        te(0).Enabled = True
    Else
        te(0).Enabled = (0 = StrComp(sTaskStatus, "true", 1))
    End If
    
    If j > 0 Then ReDim Preserve te(j - 1)
    
    For i = 1 To UBound(te)
        te(i).Enabled = te(0).Enabled
    Next
    
    AnalyzeTask = j
    
    Set xmlDoc = Nothing
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "AnalyzeTask"
    If inIDE Then Stop: Resume Next
End Function

'// intended for deletion task without SCAN_RESULT info
Public Function KillTask2(TaskFullPath As String) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim sWinTasksFolder As String
    Dim ID As String
    
    If Len(TaskFullPath) = 0 Then Exit Function
    
    sWinTasksFolder = BuildPath(sWinSysDir, "Tasks")
    
    KillTask2 = DeleteFileWEx(StrPtr(BuildPath(sWinTasksFolder, TaskFullPath)))
    
    ID = Reg.GetString(HKEY_LOCAL_MACHINE, BuildPath("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree", TaskFullPath), "Id")
    
    If ID <> "" Then
        Reg.DelKey HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & ID
        Reg.DelKey HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & ID
        Reg.DelKey HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & ID
        Reg.DelKey HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & ID
    End If
    
    KillTask2 = KillTask2 And Reg.DelKey(HKLM, BuildPath("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree", TaskFullPath))
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "KillTask2"
    If inIDE Then Stop: Resume Next
End Function

Public Sub EnumJobs()
    On Error GoTo ErrorHandler:
    
    'http://www.forensicswiki.org/wiki/Windows_Job_File_Format
    
    Dim Job As JOB_FILE
    Dim aFiles() As String
    Dim TaskFolder As String
    Dim i As Long
    Dim j As Long
    Dim sTmp As String
    Dim sHit As String
    Dim result As SCAN_RESULT
    Dim sJobName As String
    Dim bEnabled As Boolean
    Dim sRunState As String
    Dim bUpdate As Boolean
    Dim bActivation As Boolean
    
    TaskFolder = BuildPath(sWinDir, "Tasks")
    
    aFiles = ListFiles(TaskFolder, ".job", False)
    
    If AryItems(aFiles) Then
        For i = 0 To UBound(aFiles)
            
            Call ParseJob(aFiles(i), Job)
            
            sJobName = GetFileName(aFiles(i), True)
            
            Job.prop.AppName.Data = EnvironW(PathNormalize(Job.prop.AppName.Data))
            
            If Mid$(Job.prop.AppName.Data, 2, 1) <> ":" Then
                If Job.prop.WorkDir.Data <> "" Then
                    sTmp = BuildPath(Job.prop.WorkDir.Data, Job.prop.AppName.Data)
                    If FileExists(sTmp) Then Job.prop.AppName.Data = sTmp 'if only file exists in this work. folder
                End If
            End If
            
            bEnabled = False
            
            'task is enabled if there are at least 1 trigger without TASK_TRIGGER_FLAG_DISABLED flag
            For j = 0 To Job.prop.Triggers.ccTriggers - 1
                bEnabled = bEnabled Or Not CBool(Job.prop.Triggers.aTrigger(j).Flags And TASK_TRIGGER_FLAG_DISABLED)
            Next
            
            sRunState = ""
            Select Case Job.head.Status
                Case SCHED_S_TASK_READY
                    sRunState = "Ready"
                Case SCHED_S_TASK_RUNNING
                    sRunState = "Running"
                Case SCHED_S_TASK_NOT_SCHEDULED
                    sRunState = "Not scheduled"
                'case 0x41302: is some undoc. state (possible, disabled)
            End Select
            
            bUpdate = False
            bActivation = False
            If StrEndWith(Job.prop.AppName.Data, "xp_eos.exe") Then 'End of Windows XP support
                If IsMicrosoftFile(Job.prop.AppName.Data) Then bUpdate = True
            ElseIf StrEndWith(Job.prop.AppName.Data, "wgasetup.exe") Then
                If IsMicrosoftFile(Job.prop.AppName.Data) Then bActivation = True
            End If
            
            Job.prop.AppName.Data = FormatFileMissing(Job.prop.AppName.Data, Job.prop.Parameters.Data)
            
            sHit = "O22 - Task (.job): " & IIf(bEnabled, "", "(disabled) ") & IIf(sRunState <> "", "(" & sRunState & ") ", "") & _
                IIf(bUpdate, "(update) ", "") & _
                IIf(bActivation, "(activation) ", "") & _
                sJobName & " - " & _
                Job.prop.AppName.Data

            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(Job.prop.AppName.Data)

            If Not IsOnIgnoreList(sHit) Then
            
                With result
                    .Section = "O22"
                    .HitLineW = sHit
                    .State = IIf(bEnabled, ITEM_STATE_ENABLED, ITEM_STATE_DISABLED)
                    .Name = aFiles(i)
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, Job.prop.AppName.Data 'should go first! Uses by VT.
                    AddFileToFix .File, REMOVE_FILE, aFiles(i)
                    AddRegToFix .Reg, REMOVE_VALUE, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\CompatibilityAdapter\Signatures", sJobName
                    AddRegToFix .Reg, REMOVE_VALUE, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\CompatibilityAdapter\Signatures", sJobName & ".fp"
                    .CureType = FILE_BASED Or REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        Next
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumJobs"
    If inIDE Then Stop: Resume Next
End Sub

Private Function ParseJob(sFile As String, Job As JOB_FILE) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim EmptyJob As JOB_FILE
    
    Dim cStream As clsStream
    Set cStream = New clsStream
    
    Dim i As Long
    
    Job = EmptyJob
    
    cStream.LoadFileToStream sFile, cStream
    
    If cStream.Size <= 0 Then
        ErrorMsg Err, "ParseJob. Unable to open file: " & sFile
        ParseJob = False
        Exit Function
    Else
        cStream.BufferPointer = 0
        cStream.ReadData VarPtr(Job.head), LenB(Job.head)
        cStream.ReadData VarPtr(Job.prop.ccRunInstance), 2
        Read_Job_String cStream, Job.prop.AppName
        Read_Job_String cStream, Job.prop.Parameters
        Read_Job_String cStream, Job.prop.WorkDir
        Read_Job_String cStream, Job.prop.Author
        Read_Job_String cStream, Job.prop.Comment
        cStream.ReadData VarPtr(Job.prop.UserData.Size), 2
        If Job.prop.UserData.Size > 0 Then
            ReDim Job.prop.UserData.Data(Job.prop.UserData.Size - 1)
            cStream.ReadData VarPtr(Job.prop.UserData.Data(0)), Job.prop.UserData.Size
        End If
        cStream.ReadData VarPtr(Job.prop.ReservedData.Size), 2
        If Job.prop.ReservedData.Size = 8 Then
            cStream.ReadData VarPtr(Job.prop.ReservedData.StartError), 4
            cStream.ReadData VarPtr(Job.prop.ReservedData.TaskFlags), 4
        End If
        cStream.ReadData VarPtr(Job.prop.Triggers.ccTriggers), 2
        If Job.prop.Triggers.ccTriggers > 0 Then
            ReDim Job.prop.Triggers.aTrigger(Job.prop.Triggers.ccTriggers - 1)
            For i = 0 To Job.prop.Triggers.ccTriggers - 1
                cStream.ReadData VarPtr(Job.prop.Triggers.aTrigger(i)), LenB(Job.prop.Triggers.aTrigger(i))
            Next
        End If
        cStream.ReadData VarPtr(Job.prop.JobSignature), LenB(Job.prop.JobSignature)
    End If
    
    ParseJob = True
    Set cStream = Nothing
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ParseJob", sFile
    If inIDE Then Stop: Resume Next
End Function

Private Sub Read_Job_String(cStream As clsStream, JobUniStr As JOB_UNICODE_STRING)
    On Error GoTo ErrorHandler:
    
    Dim cchText As Long
    
    JobUniStr.Data = vbNullString
    
    cStream.ReadData VarPtr(JobUniStr.Length), 2
    
    If JobUniStr.Length > 1 Then
    
        cchText = JobUniStr.Length
        
        If cchText > 300 Then cchText = 300
    
        JobUniStr.Data = String$(cchText - 1, 0&)
        
        cStream.ReadData StrPtr(JobUniStr.Data), (cchText - 1) * 2& 'minus null terminator
        
        cStream.BufferPointer = cStream.BufferPointer + 2
        
        If cchText < JobUniStr.Length Then 'correct ptr
            cStream.BufferPointer = cStream.BufferPointer + (JobUniStr.Length - cchText)
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Read_Job_String"
    If inIDE Then Stop: Resume Next
End Sub

Sub EnumTasksXP() 'Win XP / Server 2003
   
    On Error GoTo ErrorHandler:
   
    Dim sSTS$, hKey&, i&, sCLSID$, lCLSIDLen&, lDataLen&
    Dim sFile$, sName$, sHit$, isSafe As Boolean
    Dim Wow6432Redir As Boolean, result As SCAN_RESULT
    
    sSTS = "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler"
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr(sSTS), 0, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not Wow6432Redir), hKey) <> 0 Then
        'regkey doesn't exist, or failed to open
        Exit Sub
    End If
    
    Do
        lCLSIDLen = MAX_VALUENAME
        sCLSID = String$(lCLSIDLen, 0&)
        lDataLen = MAX_VALUENAME
        sName = String$(lDataLen, 0&)
    
        If RegEnumValueW(hKey, i, StrPtr(sCLSID), lCLSIDLen, 0&, REG_SZ, StrPtr(sName), lDataLen) <> 0 Then Exit Do
    
        sCLSID = Left$(sCLSID, lCLSIDLen)
        sName = TrimNull(sName)
        If sName = vbNullString Then sName = "(no name)"
        sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Wow6432Redir)
        sFile = UnQuote(EnvironW(sFile))
        If sFile = vbNullString Then
            sFile = "(no file)"
        Else
            If Not FileExists(sFile) Then
                sFile = sFile & " (file missing)"
            End If
        End If
        
        'whitelist
        isSafe = isInTasksWhiteList(sCLSID & "\" & sName, sFile)
        
        If Not isSafe Then
            sHit = "O22 - ScheduledTask: " & sName & " - " & sCLSID & " - " & sFile
            
            If Not IsOnIgnoreList(sHit) Then
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                With result
                    .Section = "O22"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler", sCLSID
                    AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
            
        End If
        i = i + 1
    Loop
    RegCloseKey hKey

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumTasksXP"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub EnumTasks()
    
    If OSver.IsWindowsVistaOrGreater Then
        EnumTasksVista
    Else
        EnumTasksXP
    End If
    
End Sub

Public Sub EnumBITS_Stage1()

    Dim BitsAdmin As String
    
    If Not OSver.IsWindowsVistaOrGreater Then Exit Sub
    
    If GetServiceStartMode("bits") = SERVICE_MODE_DISABLED Then
        If GetServiceRunState("bits") <> SERVICE_RUNNING Then Exit Sub
    End If
    
    BitsAdmin = PathX64(BuildPath(sWinSysDir, "bitsadmin.exe"))
    
    Set cProcBitsAdmin = New clsProcess
    
    If cProcBitsAdmin.ProcessRun(BitsAdmin, "/list /allusers /verbose", , vbHide, False, True) Then
    
        BitsAdminExecuted = True
    End If
    
End Sub

Public Sub EnumBITS_Stage2()
    On Error GoTo ErrorHandler:
    
    Dim sLog As String, aLog() As String, sGUID As String, aURL() As String, sNotify As String, sName As String
    Dim bSectFile As Boolean, bSafe As Boolean, sHit As String, result As SCAN_RESULT, resultAll As SCAN_RESULT
    Dim sType As String, sURL As String, sDestination As String, pos As Long
    
    If Not BitsAdminExecuted Then Exit Sub
    
    BitsAdminExecuted = False
    
    cProcBitsAdmin.WaitForTerminate , , False, 15000
    sLog = cProcBitsAdmin.ConsoleRead()
    Set cProcBitsAdmin = Nothing
    
    Dim n As Long, i As Long
    
    If Len(sLog) <> 0 Then
        aLog = Split(sLog, vbCrLf)
        
        Do While n <= UBound(aLog)
            
            If Not bSectFile Then
                
                If StrBeginWith(aLog(n), "GUID:") Then
                
                    Erase aURL
                    sNotify = vbNullString
                    sGUID = GetStringToken(aLog(n), 2)
                    sName = GetStringToken(aLog(n), 4, -1)
                    If Len(sName) <> 0 Then
                        If Left$(sName, 1) = "'" Then sName = Mid$(sName, 2)
                        If Right$(sName, 1) = "'" And Len(sName) > 1 Then sName = Left$(sName, Len(sName) - 1)
                    End If
                    
                ElseIf StrBeginWith(aLog(n), "TYPE:") Then
                    
                    sType = LCase$(GetStringToken(aLog(n), 2))
                    
                ElseIf StrBeginWith(aLog(n), "JOB FILES:") Then
                
                    bSectFile = True
                End If
            Else
                If Not StrBeginWith(aLog(n), "NOTIFICATION COMMAND LINE:") Then
                    
                    ArrayAddStr aURL, GetStringToken(aLog(n), 5, -1)
                Else
                    sNotify = GetStringToken(aLog(n), 4, -1)
                    bSectFile = False
                    
                    For i = 0 To UBound(aURL)
                        
                        bSafe = False
                        
                        If Not bIgnoreAllWhitelists And bHideMicrosoft Then
                            If Not bSafe Then If StrBeginWith(aURL(i), "https://g.live.com/") Then bSafe = True
                            If Not bSafe Then If StrBeginWith(aURL(i), "http://download.windowsupdate.com/") Then bSafe = True
                            If Not bSafe Then If StrBeginWith(aURL(i), "http://au.download.windowsupdate.com/") Then bSafe = True
                            If Not bSafe Then If StrBeginWith(aURL(i), "http://bg.v4.emdl.ws.microsoft.com/") Then bSafe = True
                        End If
                        
                        If Not bSafe Then
                            
                            If sNotify = "none" Then sNotify = vbNullString
                            
                            sHit = "O22 - BITS Job: (" & sType & ") " & sGUID & " - " & aURL(i) & IIf(Len(sNotify) <> 0, " - " & sNotify, vbNullString)
                            
                            pos = InStr(aURL(i), "->")
                            If pos <> 0 Then
                                sURL = Trim$(Left$(aURL(i), pos - 1))
                                sDestination = Trim$(Mid$(aURL(i), pos + 2))
                            End If
                            
                            If Not IsOnIgnoreList(sHit) Then
                                
                                With result
                                    .Section = "O22"
                                    .HitLineW = sHit
                                    AddCustomToFix .Custom, CUSTOM_ACTION_BITS, sName, sGUID, sURL, sDestination, sNotify
                                    'AddFileToFix .File, BACKUP_FILE, BuildPath(AllUsersProfile, "Microsoft\Network\Downloader", "qmgr0.dat")
                                    'AddFileToFix .File, BACKUP_FILE, BuildPath(AllUsersProfile, "Microsoft\Network\Downloader", "qmgr1.dat")
                                    'AddFileToFix .File, BACKUP_FILE, BuildPath(AllUsersProfile, "Microsoft\Network\Downloader", "qmgr.db")
                                    .CureType = CUSTOM_BASED
                                End With
                                ConcatScanResults resultAll, result
                                AddToScanResults result
                            End If
                        End If
                    Next
                End If
            End If
            
            n = n + 1
        Loop
    End If
    
    If resultAll.CureType <> 0 Then
        resultAll.HitLineW = "O22 - BITS Job: Fix all (including legit)"
        resultAll.FixAll = True
        AddToScanResults resultAll
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumBITS"
    If inIDE Then Stop: Resume Next
End Sub

' Split string by tokens, retrieve one or several tokens (in CMD/"for" manner)
'
' tok* values:
' > 0 - concrete token number (begins from 1)
' 0 - not required (allowed for tokEnd only)
' -1 - all tokens
'
Public Function GetStringToken( _
    ByVal str As String, _
    Optional tokStart As Long = 1, _
    Optional tokEnd As Long = 0, _
    Optional delim As String = " ") As String
    
    On Error GoTo ErrorHandler:
    
    If Len(str) <> 0 Then
        Dim Tok() As String
        Dim i As Long
        Dim ret As String
        Dim n As Long
        Dim ch As Long
        Dim Length As Long
        
        str = Replace$(str, vbTab, " ")
        
        Do
            n = n + 1
            If n > Len(str) Then Exit Function
            ch = Asc(Mid$(str, n, 1))
        Loop While ch = 32 'skip leading spaces
        
        Do 'remove double delims
            Length = Len(str)
            str = Replace$(str, delim & delim, delim)
        Loop While Length <> Len(str)
        
        If n = 1 Then
            Tok = Split(str, delim)
        Else
            Tok = Split(Mid$(str, n), delim)
        End If
        
        If tokEnd = 0 Then
            If UBound(Tok) >= tokStart - 1 Then GetStringToken = Tok(tokStart - 1)
        Else
            For i = tokStart - 1 To IIf(tokEnd = -1, UBound(Tok), tokEnd)
                ret = ret & Tok(i) & delim
            Next
            If Len(ret) <> 0 Then GetStringToken = Left$(ret, Len(ret) - 1)
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetStringToken", str, tokStart, tokEnd, "'" & delim & "'"
    If inIDE Then Stop: Resume Next
End Function

Public Sub RemoveBitsJob(sGUID As String, Optional bRemoveAll As Boolean)
    On Error GoTo ErrorHandler:
    
    Dim BitsAdmin As String
    
    BitsAdmin = PathX64(BuildPath(sWinSysDir, "bitsadmin.exe"))
    
    Set Proc = New clsProcess
    
    If bRemoveAll Then
        If Proc.ProcessRun(BitsAdmin, "/RESET /ALLUSERS", , vbHide, True) Then
            Proc.WaitForTerminate , , False, 15000
        End If
    End If
    If Proc.ProcessRun(BitsAdmin, "/TAKEOWNERSHIP " & sGUID, , vbHide, True) Then
        Proc.WaitForTerminate , , False, 15000
    End If
    If Proc.ProcessRun(BitsAdmin, "/SETACLFLAGS " & sGUID & " OGDS", , vbHide, True) Then
        Proc.WaitForTerminate , , False, 15000
    End If
    If Proc.ProcessRun(BitsAdmin, "/CANCEL " & sGUID, , vbHide, True) Then
        Proc.WaitForTerminate , , False, 15000
    End If
    
    Set Proc = New clsProcess
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RemoveBitsJob"
    If inIDE Then Stop: Resume Next
End Sub

Public Function RestoreBitsJob(sName As String, sURL As String, sDestination As String, sNotify As String) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim BitsAdmin As String, sGUID As String, sLog As String
    Dim pos As Long
    
    If Not OSver.IsWindowsVistaOrGreater Then Exit Function
    
    If GetServiceStartMode("bits") = SERVICE_MODE_DISABLED Then SetServiceStartMode "bits", SERVICE_MODE_MANUAL
    If GetServiceRunState("bits") <> SERVICE_RUNNING Then StartService "bits", True, False
    
    BitsAdmin = PathX64(BuildPath(sWinSysDir, "bitsadmin.exe"))
    
    Set Proc = New clsProcess
    
    If Proc.ProcessRun(BitsAdmin, "/create /download " & """" & sName & """", , vbHide, False, True) Then
        Proc.WaitForTerminate , , False, 15000
        sLog = Proc.ConsoleRead()
        Set Proc = New clsProcess
        
        If Len(sLog) <> 0 Then
            pos = InStr(1, sLog, "Created job", 1)
            If pos <> 0 Then
                sGUID = Trim$(GetStringToken(Mid$(sLog, pos), 3))
                pos = InStr(sGUID, ".")
                If pos <> 0 Then sGUID = Left$(sGUID, pos - 1)
                
                If Proc.ProcessRun(BitsAdmin, "/addfile " & sGUID & " " & sURL & " " & """" & sDestination & """", , vbHide, True) Then
                    Proc.WaitForTerminate , , False, 15000
                End If
                
                If Len(sNotify) <> 0 Then
                    If Proc.ProcessRun(BitsAdmin, "/setnotifycmdline " & sGUID & " " & Replace$(sNotify, "'", """"), , vbHide, True) Then
                        Proc.WaitForTerminate , , False, 15000
                    End If
                End If
                
                RestoreBitsJob = (Proc.pid <> 0)
            End If
        End If
    End If
    
    Set Proc = New clsProcess
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RestoreBitsJob"
    If inIDE Then Stop: Resume Next
End Function
