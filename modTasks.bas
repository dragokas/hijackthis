Attribute VB_Name = "modTasks"
Option Explicit

' Windows Scheduled Tasks Enumerator/Killer by Alex Dragokas

' You must add reference to c:\windows\system32 (SysWow64) \taskschd.dll
' Otherwise, specify variables with Object type instead of exact types of ITask group interfaces.
' Current decision - extracted Microsoft Type Library. You should register it first.
' Type Library registration tool by Steve McMahon may help.

' Action type constants
Private Const TASK_ACTION_EXEC          As Long = 0
Private Const TASK_ACTION_COM_HANDLER   As Long = 5&
Private Const TASK_ACTION_SEND_EMAIL    As Long = 6&
Private Const TASK_ACTION_SHOW_MESSAGE  As Long = 7&

' Task state
Private Const TASK_STATE_RUNNING        As Long = 4&
Private Const TASK_STATE_QUEUED         As Long = 2&

' Include hidden tasks enumeration
Private Const TASK_ENUM_HIDDEN          As Long = 1&

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private CreateLogFile As Boolean
Private LogHandle As Integer


Public Sub EnumTasks(Optional MakeCSV As Boolean)
    On Error GoTo ErrorHandler
    Dim Stady As Long
    Dim sLogFile As String
    
    If GetServiceRunState("Schedule") <> SERVICE_RUNNING Then
        err.Raise 33333, , "Task scheduler service is not running!"
        Exit Sub
    End If
    
    'Compatibility: Vista+
    
    ' Create the TaskService object.
    Dim Service As Object
    Set Service = CreateObject("Schedule.Service")
    Stady = 1
    Service.Connect
    Stady = 2
    
    ' Get the root task folder that contains the tasks.
    Dim rootFolder As ITaskFolder
    Set rootFolder = Service.GetFolder("\")
    
    CreateLogFile = MakeCSV
    
    If MakeCSV Then
        LogHandle = FreeFile()
        sLogFile = BuildPath(AppPath(), "Tasks.csv")
        Open sLogFile For Output As #LogHandle
        Print #LogHandle, "OSver" & ";" & "State" & ";" & "Name" & ";" & "Dir" & ";" & "RunObj" & ";" & "Args" & ";" & "Note" & ";" & "Error"
    End If
    
    Stady = 3
    ' Recursively call for enumeration of current folder and all subfolders
    EnumTasksInITaskFolder rootFolder
    
    Set rootFolder = Nothing
    Set Service = Nothing
    
    If MakeCSV Then
        Close #LogHandle
        Shell "rundll32.exe shell32.dll,ShellExec_RunDLL " & """" & sLogFile & """", vbNormalFocus
    End If
    
    Exit Sub

ErrorHandler:
    ErrorMsg err, "EnumTasks. Stady: " & Stady
    If inIDE Then Stop: Resume Next
End Sub

Sub EnumTasksInITaskFolder(rootFolder As ITaskFolder, Optional isRecursiveState As Boolean)
    On Error GoTo ErrorHandler:
    
    Dim Result      As TYPE_Scan_Results
    Dim taskState   As String
    Dim RunObj      As String
    Dim RunObjExpanded As String
    Dim RunArgs     As String
    Dim DirParent   As String
    Dim DirFull     As String
    Dim sHit        As String
    Dim NoFile      As Boolean
    Dim isSafe      As Boolean
    Dim WL_ID       As Long
    Dim ActionType  As Long
    Dim taskFolder  As ITaskFolder
    
    Dim nTask           As Long
    Dim RunObjLast      As String
    Dim RunArgsLast     As String
    Dim taskStateLast   As String
    'Dim DirParentLast   As String
    Dim DirFullLast     As String
    Dim lTaskState      As Long
    Dim bTaskEnabled    As Boolean
    '------------------------------
    'Dim ComeBack        As Boolean
    Dim Stady           As Long
    Dim HRESULT         As String
    Dim errN            As Long
    Dim StadyLast       As Long
    
    
    'Debug.Print "Folder Name: " & rootFolder.Name
    'Debug.Print "Folder Path: " & rootFolder.Path
 
    Dim taskCollection As Object
    Set taskCollection = rootFolder.GetTasks(TASK_ENUM_HIDDEN)
    Stady = 1

    Dim numberOfTasks As Long
    numberOfTasks = taskCollection.Count
    Stady = 2

    Dim registeredTask  As IRegisteredTask
    Dim taskDefinition  As ITaskDefinition
    Dim taskAction      As IAction
    Dim taskActionExec  As IExecAction
    Dim taskActionEmail As IEmailAction
    Dim taskActionMsg   As IShowMessageAction
    Dim taskActionCOM   As IComHandlerAction
    Dim taskActions     As IActionCollection
    
    'Dim taskSettings3   As ITaskSettings3  'Win8+

    On Error Resume Next

    If numberOfTasks = 0 Then
        'Debug.Print "No tasks are registered."
        Stady = 3
    Else
        'Debug.Print "Number of tasks registered: " & numberOfTasks
        Stady = 4
        
        nTask = nTask + 1
        For Each registeredTask In taskCollection
        
            DoEvents
            
            err.Clear
            Call LogError(err, Stady, ClearAll:=True)
        
            NoFile = False
            isSafe = False
            
            RunObjLast = RunObj
            RunArgsLast = RunArgs
            taskStateLast = taskState
            'DirParentLast = DirParent
            DirFullLast = DirFull
            RunObj = ""
            RunObjExpanded = ""
            RunArgs = ""
            taskState = "Unknown"
            DirParent = ""
            Stady = 5
            
            
            DirFull = registeredTask.Path
            Call LogError(err, Stady)
            
            DirParent = GetParentDir$(DirFull)
            If 0 = Len(DirParent) Then DirParent = "{root}"
            
            With registeredTask
                'Debug.Print "Task Name: " & .Name
                'Debug.Print "Task Path: " & .Path
                Stady = 6

                err.Clear
                Set taskDefinition = .Definition
                Call LogError(err, Stady)
                
                If err.Number = 0 Then
                  
                  Stady = 7
                  
                  Set taskActions = taskDefinition.Actions
                  Call LogError(err, Stady)
                  
                  For Each taskAction In taskActions
                    
                    Stady = 8
                    
                    ActionType = taskAction.Type
                    Call LogError(err, Stady)
                    
                    Select Case ActionType
                    
                        Case TASK_ACTION_EXEC
                            Stady = 9
                            Set taskActionExec = taskAction
                            'Debug.Print " Type: Executable"
                            'Debug.Print "  Exec Path: " & taskActionExec.Path
                            'Debug.Print "  Exec Args: " & taskActionExec.Arguments
                            'Debug.Print "  Exec Type: " & taskActionExec.Type
                            Stady = 10
                            RunObj = taskActionExec.Path
                            Call LogError(err, Stady)
                            
                            'RunObj = EnvironW(RunObj)
                            RunObjExpanded = EnvironW(RunObj)
                            
                            Stady = 11
                            RunArgs = taskActionExec.Arguments
                            Call LogError(err, Stady)
                            
                            NoFile = Not FileExists(GetLongPath(RunObjExpanded))
                            
                        Case TASK_ACTION_SEND_EMAIL
                            Stady = 12
                            'Debug.Print " Type: Email"
                            Set taskActionEmail = taskAction
                            'Debug.Print "  Recepient: " & taskActionEmail.To
                            'Debug.Print "  Subject:   " & taskActionEmail.Subject
                            RunObj = taskActionEmail.To & ", " & taskActionEmail.Subject
                            Call LogError(err, Stady)
                        
                        Case TASK_ACTION_SHOW_MESSAGE
                            Stady = 13
                            'Debug.Print " Type: Message Box"
                            Set taskActionMsg = taskAction
                            'Debug.Print "  Title: " & taskActionMsg.Title
                            RunObj = taskActionMsg.Title
                            Call LogError(err, Stady)
                            
                        Case TASK_ACTION_COM_HANDLER
                            Stady = 14
                            'Debug.Print " Type: COM Handler"
                            Set taskActionCOM = taskAction
                            'Debug.Print "  ClassID: " & taskActionCOM.ClassId
                            'Debug.Print "  Data:    " & taskActionCOM.Data
                            RunObj = taskActionCOM.ClassId & IIf(Len(taskActionCOM.Data) <> 0, "," & taskActionCOM.Data, "")
                            Call LogError(err, Stady)
                    End Select
                    
                  Next
                End If
                
            End With
            
            'BrokenTask will be under error ignor mode until log line on this cycle
            
            Stady = 15
            Select Case registeredTask.State
                Case "0"
                    taskState = "Unknown"
                Case "1"
                    taskState = "Disabled"
                Case "2"
                    taskState = "Queued"
                Case "3"
                    taskState = "Ready"
                Case "4"
                    taskState = "Running"
            End Select
            Call LogError(err, Stady)
            
            err.Clear
            Stady = 16
            lTaskState = registeredTask.State
            Call LogError(err, Stady)
            
            Stady = 17
            bTaskEnabled = registeredTask.Enabled
            Call LogError(err, Stady)
            
            If err.Number <> 0 Then
                taskState = taskState & " (Unknown)"
            Else
                If lTaskState <> TASK_STATE_DISABLED _
                    And bTaskEnabled = False Then
                        taskState = taskState & " (Disabled)"
                End If
            End If
            
            Stady = 18
            
            'get last saved error
            Call LogError(err, StadyLast, errN, False)
            
            HRESULT = ""
            If errN <> 0 Then HRESULT = MessageText(errN)
            
            If CreateLogFile Then
                'taskState
                Print #LogHandle, OSver.MajorMinor & ";" & "" & ";" & ScreenChar(registeredTask.Name) & ";" & ScreenChar(DirParent) & ";" & _
                    ScreenChar(RunObj) & ";" & ScreenChar(RunArgs) & ";" & _
                    IIf(NoFile, "(file missing)", "") & ";" & _
                    IIf(0 <> Len(HRESULT), "(" & HRESULT & ", idx: " & StadyLast & ")", "")
            End If
            
            If Len(RunObjExpanded) <> 0 Then RunObj = RunObjExpanded
            
            sHit = "O22 - ScheduledTask: " & "(" & taskState & ") " & registeredTask.Name & " - " & DirParent & " - " & RunObj & _
                IIf(Len(RunArgs) <> 0, " " & RunArgs, "") & _
                IIf(NoFile, " (file missing)", "") & _
                IIf(0 <> Len(HRESULT), " (" & HRESULT & ", idx: " & StadyLast & ")", "")
            
            isSafe = isInTasksWhiteList(DirParent & "\" & registeredTask.Name, RunObj, RunArgs)
            
            'do not log subfolder yet (just check it)
            'If Not isRecursiveState Then
            
            If Not isSafe Then
              If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then
                    If FileExists(RunObj) Then
                        sHit = sHit & GetFileMD5(RunObj)
                    End If
                End If
                
                With Result
                    .Section = "O22"
                    .HitLineW = sHit
                    .RunObject = RunObj
                    .RunObjectArgs = RunArgs
                    .AutoRunObject = DirFull
                    .CureType = AUTORUN_BASED
                End With
                AddToScanResults Result
              End If
            End If
              
            'End If

            'Debug.Print "    Task State: " & taskState
        Next
    End If
    
    On Error GoTo ErrorHandler:
    
    Stady = 20
    Set taskActionExec = Nothing
    Set taskActionEmail = Nothing
    Set taskActionMsg = Nothing
    Set taskActionCOM = Nothing
    Set taskAction = Nothing
    Set taskDefinition = Nothing
    Set registeredTask = Nothing
    Set taskCollection = Nothing
    
    Stady = 21
    Dim taskFolderCollection As ITaskFolderCollection
    Set taskFolderCollection = rootFolder.GetFolders(0&)
    
    For Each taskFolder In taskFolderCollection 'deep to subfolders
        EnumTasksInITaskFolder taskFolder, True
    Next
    
    Set taskFolder = Nothing
    Set taskFolderCollection = Nothing
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "EnumTasksInITaskFolder. Stady: " & Stady & ". Number of tasks: " & numberOfTasks & ". Curr. task # " & nTask & ": " & DirFull & ", " & _
        "RunObj = " & RunObj & ", RunArgs = " & RunArgs & ", taskState = " & taskState
        '& ". ____Last task Data:___ " & DirFullLast & ", " & _
        '"RunObjLast = " & RunObjLast & ", RunArgsLast = " & RunArgsLast & ", taskStateLast = " & taskStateLast
'    If ComeBack Then
'        ComeBack = False
'        If inIDE Then Stop
'        Return
'    End If
    If inIDE Then Stop: Resume Next
End Sub

Public Function isInTasksWhiteList(sPathName As String, sTargetFile As String, sArguments As String) As Boolean
    On Error GoTo ErrorHandler
    Dim WL_ID As Long

    If bIgnoreAllWhitelists Then Exit Function
    If Not oDict.TaskWL_ID.Exists(sPathName) Then Exit Function
    
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
    ErrorMsg err, "modTasks.isInTasksWhiteList"
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
    UnScreenChar = LTrim(Replace$(sText, "\\\,", ";"))
End Function

Sub LogError(objError As ErrObject, in_out_Stady As Long, Optional out_LastLoggedErrorNumber As Long, Optional in_ActionPut As Boolean = True, Optional ClearAll As Boolean)
    '
    'in_ActionPut - if false, this function fill _out_ parameters with last saved error number and Stady position
    
    'Purpose of function:
    'Log first error and do not overwrite it until 'ClearAll' parameter become true
    
    Static Stady As Long, ErrNum As Long
    
    If ClearAll = True Then
        Stady = 0
        ErrNum = 0
        Exit Sub
    End If
    
    If in_ActionPut = False Then
        'get
        in_out_Stady = Stady
        out_LastLoggedErrorNumber = ErrNum
    Else
        'put
        If ErrNum = 0 Then
            ErrNum = objError.Number
            Stady = in_out_Stady
        End If
    End If
End Sub

Public Function KillTask(TaskFullPath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim TaskPath As String
    Dim TaskName As String
    Dim pos As Long
    Dim Stady As Long
    Dim ComeBack As Boolean
    Dim lTaskState As Long
    Dim BrokenTask As Boolean
    
    'Compatibility: Vista+
    
    pos = InStrRev(TaskFullPath, "\")
    If pos <> 0 Then
        Stady = 1
        TaskPath = Left$(TaskFullPath, pos)
        If Len(TaskPath) > 1 Then TaskPath = Left$(TaskPath, Len(TaskPath) - 1) 'trim last backslash
        TaskName = Mid$(TaskFullPath, pos + 1)
    Else
        Exit Function
    End If

    Stady = 2
    ' Create the TaskService object.
    Dim Service As Object
    Set Service = CreateObject("Schedule.Service")
    Stady = 3
    Service.Connect
    
    Stady = 4
    ' Get the root task folder that contains the tasks.
    Dim rootFolder As ITaskFolder
    Set rootFolder = Service.GetFolder(TaskPath)
    
    Stady = 5
    Dim registeredTask  As IRegisteredTask
    Set registeredTask = rootFolder.GetTask(TaskName)
    
    'Dim taskCollection As Object
    'Set taskCollection = rootFolder.GetTasks(TASK_ENUM_HIDDEN)
    '
    '    For Each registeredTask In taskCollection
    '        With registeredTask
    '          If InStr(1, .Path, TaskName, 1) <> 0 Then
    '            Debug.Print "Task Name: " & .Name
    '            Debug.Print "Task Path: " & .Path
    '          End If
    '
    '        End With
    '    Next
    'Stop
    
    ' Stop the task
    
    'I should insert here this strange error handling routines because ITask Schedule interfaces with different
    'kinds of possible errors in XML structures and unsufficient access rights caused by malware is a very poor stuff for developers,
    'because it can produce so many unexpected errors. So, we are full of trubles.
    
    'Maybe, later I rewrite it into manual parsing.
    
    On Error Resume Next
    Stady = 6
    lTaskState = registeredTask.State
    If err.Number <> 0 Then
        ComeBack = True
        GoSub ErrorHandler
        registeredTask.Stop 0&
        Sleep 2000&
        BrokenTask = True
    ElseIf lTaskState = TASK_STATE_RUNNING Or lTaskState = TASK_STATE_QUEUED Then
        registeredTask.Stop 0&
        Sleep 2000&
    End If
    
    Stady = 7
    If registeredTask.Enabled Then registeredTask.Enabled = False
    
    Dim taskDefinition  As ITaskDefinition
    Dim taskAction      As IAction
    Dim taskActionExec  As IExecAction
    
    Stady = 8
    
    ' Kill process
    err.Clear
    Set taskDefinition = registeredTask.Definition
    
    If err.Number <> 0 Then
        If Not BrokenTask Then
            ComeBack = True
            GoSub ErrorHandler
        End If
    Else
      Stady = 9
      For Each taskAction In taskDefinition.Actions
        Stady = 10
        If TASK_ACTION_EXEC = taskAction.Type Then
            Stady = 11
            Set taskActionExec = taskAction
            'Debug.Print taskActionExec.Path
            If FileExists(taskActionExec.Path) Then
                KillProcessByFile taskActionExec.Path
            End If
        End If
      Next
    End If
    
    On Error GoTo ErrorHandler
    
    Stady = 12
    ' Remove the Job
    rootFolder.DeleteTask TaskName, 0&
    Sleep 1000&
    
    KillTask = True
    
    Stady = 13
    Set taskActionExec = Nothing
    Set taskAction = Nothing
    Set taskDefinition = Nothing
    Set registeredTask = Nothing
    Set rootFolder = Nothing
    Set Service = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg err, "KillTask. Stady: " & Stady
    If ComeBack Then
        ComeBack = False
        If inIDE Then Stop
        Return
    End If
    If inIDE Then Stop: Resume Next
End Function

