Attribute VB_Name = "modTasks"
Option Explicit

' Windows Scheduled Tasks Enumerator/Killer by Alex Dragokas

' To add early binding, set reference to taskschd.dll (not applicable to EnumTasks2)

' keys
' Vista+: HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks
' XP:     HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler

'regexp replacing (for CLSID):
'^(.*?): (\(disabled\) )?(.*?) - (.*?) - (.*?)( \(Microsoft\))?
' \3; \4; \5

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
End Type

' Task state
Private Const TASK_STATE_RUNNING        As Long = 4&
Private Const TASK_STATE_QUEUED         As Long = 2&

' Include hidden tasks enumeration
Private Const TASK_ENUM_HIDDEN          As Long = 1&

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private CreateLogFile As Boolean
Private LogHandle As Integer


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
'    CreateLogFile = MakeCSV
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
'            If CreateLogFile Then
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
'                If bMD5 Then
'                    If FileExists(RunObj) Then
'                        sHit = sHit & GetFileMD5(RunObj)
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

Public Function PathNormalize(ByVal sFileName As String) As String
    
    AppendErrorLogCustom "PathNormalize - Begin", "File: " & sFileName
    
    Dim sTmp As String, bShouldSeek As Boolean
    
    sFileName = UnQuote(sFileName)
    
    If Mid$(sFileName, 2, 1) <> ":" Then
        bShouldSeek = True  'relative or on the %PATH%
    Else
        If Not FileExists(sFileName) Then bShouldSeek = True 'e.g. no extension
    End If
    
    If bShouldSeek Then
        sTmp = FindOnPath(sFileName)
        If Len(sTmp) <> 0 Then
            sFileName = sTmp
        End If
    End If
    
    PathNormalize = sFileName
    
    AppendErrorLogCustom "PathNormalize - End"
End Function

Public Function isInTasksWhiteList(sPathName As String, sTargetFile As String, sArguments As String) As Boolean
    On Error GoTo ErrorHandler
    Dim WL_ID As Long

    If bIgnoreAllWhitelists Then Exit Function
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
        End If
        Exit Function
    End If
    
    WL_ID = oDict.TaskWL_ID(sPathName)
    
    With g_TasksWL(WL_ID)
        'verifying all components
        If inArraySerialized(sTargetFile, .RunObj, "|", , , 1) And _
          inArraySerialized(sArguments, .Args, "|", , , 1) Then
            isInTasksWhiteList = True
        End If
    End With
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modTasks.isInTasksWhiteList"
    If inIDE Then Stop: Resume Next
End Function

' replace ; -> \\\, (for CSV)
' adds 1 space (Excel support)
Function ScreenChar(sText As String) As String
    ScreenChar = " " & Replace$(sText, ";", "\\\,")
End Function
' replace \\\, -> ; (to interpret CSV)
' remove 1 space from left side
Public Function UnScreenChar(sText As String) As String
    UnScreenChar = LTrim$(Replace$(sText, "\\\,", ";"))
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


' ####################################################################
' ####################################################################

'// Alternate version based on manual XML parsing
'// (no windows service involved)

Public Sub EnumTasks2(Optional MakeCSV As Boolean)
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "EnumTasks2 - Begin"
    
    Dim sLogFile        As String
    Dim odFileTasks     As clsTrickHashTable
    Dim odRegTasks      As clsTrickHashTable
    Dim i               As Long
    Dim j               As Long
    Dim Result          As SCAN_RESULT
    Dim DirParent       As String
    Dim DirFull         As String
    Dim TaskName        As String
    Dim sHit            As String
    Dim NoFile          As Boolean
    Dim isSafe          As Boolean
    Dim SignResult      As SignResult_TYPE
    Dim bIsMicrosoftFile As Boolean
    Dim aFiles()        As String
    Dim te()            As TASK_ENTRY
    Dim numTasks        As Long
    Dim sWinTasksFolder As String
    Dim bNoFile         As Boolean
    Dim aSubKeys()      As String
    Dim oKey            As Variant
    Dim ID              As String
    Dim aID()           As String
    Dim DirFull_2       As String
    Dim bTelemetry      As Boolean
    Dim sRunFilename    As String
    
    '// Todo: Add analysis for dangerous host exe, like cmd.exe. Don't try to show 'Microsoft' EDS sign in log for them,
    'exception: if no arguments specified.
    'powershell.exe
    'cmd.exe.
    'wscript.exe
    'mshta.exe
    'svchost.exe
    'rundll32.exe
    'pcalua.exe
    'schtasks.exe
    'sc.exe
    
    '// TODO: Add record: "O22 - Task: 'Task scheduler' service is disabled!"
    
    If GetServiceRunState("Schedule") <> SERVICE_RUNNING Then
        'Err.Raise 33333, , "Task scheduler service is not running!"
    End If
    
'    Set odFileTasks = New clsTrickHashTable
'    Set odRegTasks = New clsTrickHashTable
'    odFileTasks.CompareMode = 1
'    odRegTasks.CompareMode = 1
    
    If MakeCSV Then
        LogHandle = FreeFile()
        sLogFile = BuildPath(AppPath(), "Tasks.csv")
        Open sLogFile For Output As #LogHandle
        Print #LogHandle, "OSver" & ";" & "State" & ";" & "Name" & ";" & "Dir" & ";" & "RunObj" & ";" & "Args" & ";" & "Note" & ";" & "Error"
    End If
    
'    'enum registry info first
'    Erase aID
'    For i = 1 To Reg.EnumSubKeysToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks", aID())
'        DirFull = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & aID(i), "Path")
'        'cache full names that points to xml file
'        If DirFull <> "" Then
'            If Not odRegTasks.Exists(DirFull) Then odRegTasks.Add DirFull, aID(i) 'key = \ + relative xml path; value = {CLSID}
'        End If
'    Next
    
    sWinTasksFolder = BuildPath(sWinSysDir, "Tasks")
    
    aFiles = ListFiles(sWinTasksFolder, "", True)
    
    If AryPtr(aFiles) Then
      For i = 0 To UBound(aFiles)
      
        AppendErrorLogCustom "Task: " & aFiles(i)
        
        numTasks = AnalyzeTask(aFiles(i), te())
        
        DirFull = Mid$(aFiles(i), Len(sWinTasksFolder) + 1)
        
        'cache xml path
        'odFileTasks.Add DirFull, 0
        
        DirParent = GetParentDir(DirFull)
        If 0 = Len(DirParent) Then DirParent = "{root}"
        
        TaskName = GetFileNameAndExt(DirFull)
        
        UpdateProgressBar "O22", TaskName
        
        For j = 0 To numTasks - 1
            
            WipeSignResult SignResult
            
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
            
            'If DirFull = "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" Then Stop
            
            'EDS checking for records, that doesn't exist in database
            
            If isSafe Or StrBeginWith(DirFull, "\Microsoft\") Then
                If te(j).ActionType = TASK_ACTION_EXEC Then
                    te(j).RunObj = PathNormalize(te(j).RunObj)
                    SignVerify te(j).RunObj, SV_LightCheck Or SV_PreferInternalSign, SignResult
                    
                ElseIf te(j).ActionType = TASK_ACTION_COM_HANDLER Then
                    te(j).RunObjCom = PathNormalize(te(j).RunObjCom)
                    SignVerify te(j).RunObjCom, SV_LightCheck Or SV_PreferInternalSign, SignResult
                End If
            End If
            
            ' EDS checking for items in database
            If isSafe Then
                AppendErrorLogCustom "[OK] EnumTasksInITaskFolder: WhiteListed."
                
                If te(j).ActionType = TASK_ACTION_EXEC Then
                    
                    bIsMicrosoftFile = (SignResult.isMicrosoftSign And SignResult.isLegit)
                    
                    isSafe = (bIsMicrosoftFile And bHideMicrosoft And Not bIgnoreAllWhitelists)
                    
                    If Not bIsMicrosoftFile Then
                        AppendErrorLogCustom "[Failed] EnumTasksInITaskFolder: File - " & te(j).RunObj & " => is not Microsoft EDS !!! <======"
                        Debug.Print "Task MS file has wrong EDS: " & te(j).RunObj
                    End If
                ElseIf te(j).ActionType = TASK_ACTION_COM_HANDLER And Len(te(j).RunObjCom) <> 0 Then
                
                  'for some reason, part of CLSID records on Win10 is not registered
                  If te(j).RunObjCom <> "(no file)" Then
                
                    SignVerify te(j).RunObjCom, SV_LightCheck Or SV_PreferInternalSign, SignResult
                
                    bIsMicrosoftFile = (SignResult.isMicrosoftSign And SignResult.isLegit)
                    
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
            
            If Not isSafe Then
                
                bTelemetry = False
                
                sRunFilename = GetFileName(te(j).RunObj, True)
                
                If InStr(1, DirParent, "Customer Experience", 1) <> 0 Then bTelemetry = True
                If InStr(1, DirParent, "Application Experience", 1) <> 0 Then bTelemetry = True
                If InStr(1, DirParent, "telemetry", 1) <> 0 Then bTelemetry = True
                If InStr(1, TaskName, "telemetry", 1) <> 0 Then bTelemetry = True
                If StrComp(sRunFilename, "OLicenseHeartbeat.exe", 1) = 0 Then bTelemetry = True
                
                'skip signature mark for host processes
                If SignResult.isMicrosoftSign Then
                    If inArraySerialized(sRunFilename, "rundll32.exe|schtasks.exe|sc.exe|cmd.exe|wscript.exe|" & _
                      "mshta.exe|pcalua.exe|powershell.exe|svchost.exe", "|") Then SignResult.isMicrosoftSign = False
                End If
                
                sHit = "O22 - Task: " & IIf(te(j).Enabled, "", "(disabled) ") & _
                  IIf(bTelemetry, "(telemetry) ", "") & _
                  IIf(DirParent = "{root}", TaskName, DirParent & "\" & TaskName)
                
                sHit = sHit & " - " & te(j).RunObj & _
                  IIf(Len(te(j).RunArgs) <> 0, " " & te(j).RunArgs, "") & _
                  IIf(te(j).FileMissing And Not bNoFile, " (file missing)", "") & _
                  IIf(SignResult.isMicrosoftSign, " (Microsoft)", "")
                
                If Not IsOnIgnoreList(sHit) Then
              
                    If bMD5 Then
                        If FileExists(te(j).RunObj) Then
                            sHit = sHit & GetFileMD5(te(j).RunObj)
                        End If
                    End If
                    
                    With Result
                        .Section = "O22"
                        .HitLineW = sHit
                        AddFileToFix .File, REMOVE_FILE, aFiles(i)
                        
                        te(j).RegID = Reg.GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirFull, "Id")
                        If te(j).RegID <> "" Then
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & te(j).RegID
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & te(j).RegID
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & te(j).RegID
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & te(j).RegID
                        End If
                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirFull
                        
                        .CureType = FILE_BASED Or REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
            
        Next
        
      Next
    End If
    
'    'check registry entries for compliance with xml files
'    'Stady 1
'    For Each oKey In odRegTasks.Keys
'
'        If Not odFileTasks.Exists(oKey) Then
'
'            DirFull = oKey
'            ID = odRegTasks(oKey)
'
'            sHit = "O22 - Task: (damaged) " & DirFull
'
'            With Result
'                .Section = "O22"
'                .HitLineW = sHit
'
'                'if there are another ids with the same path name
'                For i = 1 To UBound(aID)
'                    DirFull_2 = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & aID(i), "Path")
'
'                    If StrComp(DirFull, DirFull_2, 1) = 0 Then
'
'                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & aID(i)
'                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & aID(i)
'                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & aID(i)
'                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & aID(i)
'
'                    End If
'                Next
'
'                AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirFull
'
'            End With
'            AddToScanResults Result
'
'            odFileTasks.Add DirFull, 0
'        End If
'    Next
    
'    'Stady 2
'    'if \Tree\ still contains leftovers
'    Erase aSubKeys
'    For i = 1 To Reg.EnumSubKeysToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree", aSubKeys(), , , True)
'
'        DirFull = Mid$(aSubKeys(i), Len("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree") + 1)
'
'        ID = Reg.GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirFull, "Id")
'
'        'if real task (not a subdir)
'        If ID <> "" Then
'
'            If Not odFileTasks.Exists(DirFull) Then
'
'                sHit = "O22 - Task: (damaged) " & DirFull
'
'                With Result
'                    .Section = "O22"
'                    .HitLineW = sHit
'
'                    If ID <> "" Then
'                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\" & ID
'                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\" & ID
'                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\" & ID
'                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\" & ID
'                    End If
'                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree" & DirFull
'                End With
'                AddToScanResults Result
'            End If
'        End If
'    Next
    
    If MakeCSV Then
        Close #LogHandle
        Shell "rundll32.exe shell32.dll,ShellExec_RunDLL " & """" & sLogFile & """", vbNormalFocus
    End If
    
    Set odFileTasks = Nothing
    Set odRegTasks = Nothing
    
    AppendErrorLogCustom "EnumTasks2 - End"
    Exit Sub

ErrorHandler:
    ErrorMsg Err, "EnumTasks2"
    If inIDE Then Stop: Resume Next
End Sub

Private Function AnalyzeTask(sFileName As String, te() As TASK_ENTRY) As Long
    'additional info:
    'https://github.com/libyal/winreg-kb/blob/master/documentation/Task%20Scheduler%20Keys.asciidoc
    
    On Error GoTo ErrorHandler:
    
    Dim xmlDoc          As XMLDocument
    Dim xmlElement      As CXmlElement
    Dim bExecBased      As Boolean
    Dim sFileData       As String
    Dim sTmp            As String
    Dim i As Long
    Dim j As Long
    
    ReDim te(0)
    
    sFileData = ReadFileContents(sFileName, False)
    
    If Len(sFileData) = 0 Then Exit Function
    
    Set xmlDoc = New XMLDocument
    Call xmlDoc.LoadData(sFileData)
    
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
            
            If te(j).ActionType <> TASK_ACTION_COM_HANDLER Then
                '// TODO email / msg
                
                
            End If
        End If
    End If
    
    te(0).Enabled = (0 = StrComp(xmlDoc.NodeValueByName("Settings\Enabled"), "true", 1))
    
    If j > 0 Then ReDim Preserve te(j - 1)
    
    For i = 1 To UBound(te)
        te(i).Enabled = te(0).Enabled
    Next
    
    AnalyzeTask = j
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "AnalyzeTask"
    If inIDE Then Stop: Resume Next
End Function

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
