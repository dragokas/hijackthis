Attribute VB_Name = "modMain_2"
'[modMain_2.bas]

'
' Core check / Fix Engine
'
' (part 2: O25 - O26)

'
' O25, O26 by Alex Dragokas
'

'O25 - Windows Management Instrumentation (WMI) event consumers
'O26 - Image File Execution Options (IFEO) and System Tools hijack

Option Explicit

Private Const MAX_CODE_LENGTH As Long = 300&

Private Declare Function VerifierIsPerUserSettingsEnabled Lib "Verifier.dll" () As Long

' Get array of namespaces
' Note: you should initialize aNameSpaces variable-length array with '0' element first.

Sub WMI_GetNamespaces(sNamespace As String, aNameSpaces() As String)

    On Error GoTo ErrorHandler:

    Dim objService As Object, objNamespace As Object, colNamespaces As Object, SubNameSpace As String
    
    On Error Resume Next
    Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & sNamespace)
    
    If Err.Number <> 0 Then
        ErrorMsg Err, "modMain2_WMI_GetNamespaces", "Namespace: ", sNamespace
    End If
    On Error GoTo ErrorHandler:
    
    If Not (objService Is Nothing) And InStr(1, sNamespace, "Root\WMI", 1) <> 1 Then
        
        Set colNamespaces = objService.InstancesOf("__NAMESPACE")
        
        For Each objNamespace In colNamespaces
            
            SubNameSpace = sNamespace & "\" & objNamespace.Name
            
            'do not query AD
            
            If InStr(1, SubNameSpace, "Root\directory\LDAP", 1) <> 1 Then
            
                If Not bAutoLogSilent Then DoEvents
            
                ReDim Preserve aNameSpaces(UBound(aNameSpaces) + 1)
                aNameSpaces(UBound(aNameSpaces)) = SubNameSpace
            
                Call WMI_GetNamespaces(SubNameSpace, aNameSpaces)
            
            End If

        Next
        
        Set colNamespaces = Nothing: Set objNamespace = Nothing
        
    End If
    
    Set objService = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain2_WMI_GetNamespaces", "Namespace: ", sNamespace
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckO25Item()
    'WMI_Get_Event_Consumers()
    '
    'http://www.trendmicro.com/cloud-content/us/pdfs/security-intelligence/white-papers/wp__understanding-wmi-malware.pdf
    '
    'thanks to Julius Dizon, Lennard Galang, Marvin Cruz
    '

    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO25Item - Begin"

    Dim objService As Object, colBindings As Object, objBinding As Object, objSWbemDateTime As Object
    Dim FilterName As String, ConsumerName As String, i As Long
    Dim FilterPath As String, ConsumerPath As String
    Dim FilterNameSpace As String, ConsumerNameSpace As String, ConsumerClassName As String
    Dim objServiceConsumer As Object, objConsumer As Object
    Dim objServiceFilter As Object, objFilter As Object, sFilterQuery As String
    Dim objTimerNamespace As Object, objTimer As Object, sTimerClassName As String, sTimerName As String, lTimerInterval As Long, lKillTimeout As Long
    Dim sHit As String, sScriptFile As String, sAdditionalInfo As String, sEventName As String, sScriptText As String, sScriptEngine As String
    Dim cmdExecute As String, cmdWorkDir As String, cmdArguments As String, bRunInteractively As Boolean
    
    Dim result As SCAN_RESULT
    Dim Stady As Single, ComeBack As Boolean, NoConsumer As Boolean, NoFilter As Boolean, bOtherConsumerClass As Boolean
    Dim bDangerScript As Boolean
    
    If GetServiceRunState("winmgmt") <> SERVICE_RUNNING Then
        If Not bAdditional Then
            Exit Sub
        Else
            If Not RunWMI_Service(True, False, bAutoLogSilent) Then Exit Sub
        End If
    End If
    
    If OSver.MajorMinor <= 5 Then Exit Sub 'XP+ only
    
    Dim aNameSpaces() As String
    ReDim aNameSpaces(0)
    
    'connecting to namespace 'root\subscription' for future use
    Stady = 0
    Set objTimerNamespace = CreateObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\root\subscription")
    
    Stady = 1
    Set objTimerNamespace = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\root\subscription")
    
    'get all namespaces for current machine
    
    Stady = 2
    'Call WMI_GetNamespaces("Root", aNameSpaces)
    'let's concentrate on actual malware method
    ReDim aNameSpaces(1)
    aNameSpaces(1) = "root\subscription"
    
    For i = 1 To UBound(aNameSpaces)

        'connecting to namespace

        Stady = 3
        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & aNameSpaces(i))

        If Not bAutoLogSilent Then DoEvents
    
        'get binding info ( Filter <-> Consumer )
        Stady = 4
        Set colBindings = objService.ExecQuery("SELECT * FROM __FilterToConsumerBinding", "WQL", 16 + 32)
        
        For Each objBinding In colBindings
            
            If Not IsNull(objBinding.Filter) Then FilterPath = objBinding.Filter
            If Not IsNull(objBinding.Consumer) Then ConsumerPath = objBinding.Consumer
            
            'split into components
            
            FilterName = GetStringInsideQt(FilterPath)
            ConsumerName = GetStringInsideQt(ConsumerPath)
            
            Call ExtractNameSpaceAndClassNameFromString(FilterPath, FilterNameSpace)
            Call ExtractNameSpaceAndClassNameFromString(ConsumerPath, ConsumerNameSpace, ConsumerClassName)
            
            If 0 = Len(FilterNameSpace) Then FilterNameSpace = aNameSpaces(i)
            If 0 = Len(ConsumerNameSpace) Then ConsumerNameSpace = aNameSpaces(i)
            
            'Debug.Print FilterPath
            'Debug.Print ConsumerPath
            
            If 0 <> Len(FilterName) And 0 <> Len(ConsumerName) Then
            
                'connecting to consumer's own namespace
                Stady = 5
                If StrComp(ConsumerNameSpace, aNameSpaces(i), 1) = 0 Then
                    'if consumer's namespace is a same
                    Set objServiceConsumer = objService
                Else
                    On Error Resume Next
                    Set objServiceConsumer = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & ConsumerNameSpace)
                End If
                
                Stady = 6
                On Error Resume Next
                Set objConsumer = objServiceConsumer.Get(ConsumerPath)
                On Error GoTo ErrorHandler:
                
                cmdExecute = ""
                cmdWorkDir = ""
                cmdArguments = ""
                sScriptFile = ""
                sScriptText = ""
                sAdditionalInfo = ""
                sScriptEngine = ""
                sTimerName = ""
                lTimerInterval = 0
                lKillTimeout = 0
                bRunInteractively = False
                
                If Not (objConsumer Is Nothing) Then
                    
                    'Checking several known classes on: root\subscription
                    'to provide a bit more information to log
                    
                    bOtherConsumerClass = False
                    
                    Stady = 7
                    If StrComp(ConsumerClassName, "ActiveScriptEventConsumer", 1) = 0 Then
                    
                        result.O25.Consumer.Type = O25_CONSUMER_ACTIVE_SCRIPT
                    
                        'Debug.Print objConsumer.ScriptingEngine    'language (engine)
                        'Debug.Print objConsumer.ScriptFileName     'external file
                        'Debug.Print objConsumer.ScriptText         'embedded script code
                        Stady = 8
                        If Not IsNull(objConsumer.ScriptFilename) Then sScriptFile = objConsumer.ScriptFilename
                        If Not IsNull(objConsumer.ScriptText) Then sScriptText = objConsumer.ScriptText
                        If Not IsNull(objConsumer.ScriptingEngine) Then sScriptEngine = objConsumer.ScriptingEngine
                        If Not IsNull(objConsumer.KillTimeout) Then lKillTimeout = CLng(Val(objConsumer.KillTimeout))
                        
'                        If Not IsNull(objConsumer.ScriptingEngine) Then
'                            sAdditionalInfo = "Lang=" & """" & objConsumer.ScriptingEngine & """" & ", "
'                        End If
                        If 0 <> Len(sScriptFile) Then
                            'sAdditionalInfo = sAdditionalInfo & "ScriptFileName=" & """" & sScriptFile & """"
                            sAdditionalInfo = sAdditionalInfo & sScriptFile
                        End If
                        If 0 <> Len(sScriptText) Then
                            'sAdditionalInfo = sAdditionalInfo & "ScriptCode=" & """" & StripCode(sScriptText) & """"
                            sAdditionalInfo = sAdditionalInfo & IIf(sAdditionalInfo <> "", " / ", "") & StripCode(sScriptText)
                        End If

                    ElseIf StrComp(ConsumerClassName, "CommandLineEventConsumer", 1) = 0 Then
                        Stady = 9
                        
                        result.O25.Consumer.Type = O25_CONSUMER_COMMAND_LINE
                        
                        'Example:
                        'kernrate: O25 - WMI Event: BVTConsumer / BVTFilter -> Executable="", Arguments="cscript KernCap.vbs"
                        'https://pcsxcetrasupport3.wordpress.com/2011/10/23/event-10-mystery-solved/
                        
                        'Debug.Print objConsumer.ExecutablePath         'main execution module
                        'Debug.Print objConsumer.CommandLineTemplate    'arguments
                        'debug.print objConsumer.WorkingDirectory       'Work Dir.
                        ComeBack = True
                        If Not IsNull(objConsumer.ExecutablePath) Then cmdExecute = objConsumer.ExecutablePath
                        Stady = 9.1
                        If Not IsNull(objConsumer.WorkingDirectory) Then cmdWorkDir = objConsumer.WorkingDirectory
                        Stady = 9.2
                        If Not IsNull(objConsumer.CommandLineTemplate) Then cmdArguments = objConsumer.CommandLineTemplate
                        Stady = 9.3
                        If Not IsNull(objConsumer.RunInteractively) Then bRunInteractively = objConsumer.RunInteractively
                        If Not IsNull(objConsumer.KillTimeout) Then lKillTimeout = CLng(Val(objConsumer.KillTimeout))
                        
                        'sAdditionalInfo = "Executable=" & """" & cmdExecute & """" & _
                        '    ", WorkDir=" & """" & cmdWorkDir & """" & _
                        '    ", Arguments=" & """" & StripCode(cmdArguments) & """"
                           
                        sAdditionalInfo = cmdExecute & " " & cmdArguments & IIf(cmdWorkDir <> "", " (WorkDir = " & cmdWorkDir & ")", "")
                        
                        ComeBack = False
                        
'                    ElseIf StrComp(ConsumerClassName, "LogFileEventConsumer", 1) = 0 Then
'                        Stady = 10
'                        'Debug.Print objConsumer.FileName    'Where information logged
'                        'Debug.Print objConsumer.Text        'What kind of information logged
'
'                        If Not IsNull(objConsumer.FileName) Then sConsumerFileName = objConsumer.FileName
'                        If Not IsNull(objConsumer.Text) Then sConsumerText = objConsumer.Text
'
'                        sAdditionalInfo = "LogFileName=" & """" & sConsumerFileName & """" & _
'                            ", InfoType=" & """" & sConsumerText & """"
'
'                    ElseIf StrComp(ConsumerClassName, "NTEventLogEventConsumer", 1) = 0 Then
'                        Stady = 11
'                        'Debug.Print objConsumer.SourceName
'                        If Not IsNull(objConsumer.SourceName) Then
'                            sAdditionalInfo = "LogSourceName=" & """" & objConsumer.SourceName & """"
'                        End If
'
'                    ElseIf StrComp(ConsumerClassName, "SMTPEventConsumer", 1) = 0 Then
'                        'Debug.Print objConsumer.BccLine
'                        'Debug.Print objConsumer.CcLine
'                        'Debug.Print objConsumer.CreatorSID
'                        'Debug.Print objConsumer.FromLine
'                        'Debug.Print objConsumer.HeaderFields
'                        'Debug.Print objConsumer.MachineName
'                        'Debug.Print objConsumer.MaximumQueueSize
'                        'Debug.Print objConsumer.Message
'                        'Debug.Print objConsumer.Name
'                        'Debug.Print objConsumer.ReplyToLine
'                        'Debug.Print objConsumer.SMTPServer
'                        'Debug.Print objConsumer.Subject
'                        'Debug.Print objConsumer.ToLine
'                    Else
'                        Stady = 12
'                        'other consumers -> Show Namespace + ClassName
'                        sAdditionalInfo = "ClassName=" & """" & ConsumerNameSpace & ":" & ConsumerClassName & """"
                    Else
                        bOtherConsumerClass = True
                    End If
                    
                End If
                
                'Trying to find associated timer inside the filter
                
                'connecting to filter's own namespace
                
                Stady = 13
                If StrComp(FilterNameSpace, aNameSpaces(i), 1) = 0 Then
                    'if consumer's namespace is a same
                    Set objServiceFilter = objService
                Else
                    On Error Resume Next
                    Set objServiceFilter = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & FilterNameSpace)
                End If
                
                Stady = 14
                On Error Resume Next
                Set objFilter = objServiceFilter.Get(FilterPath)
                On Error GoTo ErrorHandler:
                
                If Not (objFilter Is Nothing) Then
                
                    Stady = 15
                    If Not IsNull(objFilter.Query) Then sFilterQuery = objFilter.Query
                
                    'receives events from timer ?
                    If InStr(1, sFilterQuery, "__timerevent", 1) <> 0 Then
                    
                        'SELECT * FROM __timerevent WHERE timerid="Dragokas_WMITimer2"
                        sTimerName = GetStringInsideQt(sFilterQuery)
                    
                        If 0 <> Len(sTimerName) Then
                            'searching timer's Class name (2 options)
                            
                            Set objTimer = Nothing
                            
                            Stady = 16
                            On Error Resume Next
                            sTimerClassName = "__IntervalTimerInstruction"
                            Set objTimer = objTimerNamespace.Get(sTimerClassName & ".TimerId=" & """" & sTimerName & """")
                            
                            If Not (objTimer Is Nothing) Then
                                result.O25.Timer.Type = O25_TIMER_INTERVAL
                                
                                If Not IsNull(objTimer.IntervalBetweenEvents) Then
                                    lTimerInterval = CLng(Val(objTimer.IntervalBetweenEvents))
                                End If
                            Else
                                sTimerClassName = "__AbsoluteTimerInstruction"
                                Set objTimer = objTimerNamespace.Get(sTimerClassName & ".TimerId=" & """" & sTimerName & """")
                                result.O25.Timer.Type = O25_TIMER_ABSOLUTE
                                
                                Set objSWbemDateTime = CreateObject("WbemScripting.SWbemDateTime")
                                objSWbemDateTime.Value = objTimer.EventDateTime
                                result.O25.Timer.EventDateTime = objSWbemDateTime.GetVarDate(True)
                                Set objSWbemDateTime = Nothing
                            End If
                            On Error GoTo ErrorHandler:
                            
                            If objTimer Is Nothing Then
                                sTimerClassName = ""
                                sTimerName = ""
                            Else
                                Set objTimer = Nothing
                            End If
                        
                        End If
                    
                    Else
                        Stady = 17
                        'if another event source -> print its name
                        sEventName = ExtractEventName(sFilterQuery)
                        If 0 <> Len(sEventName) Then
                            sAdditionalInfo = "Event=" & """" & sEventName & """" & ", " & sAdditionalInfo
                        End If
                    End If
                
                End If
            
                'WhiteList
                Stady = 18
                'If Not (StrComp(ConsumerClassName, "NTEventLogEventConsumer", 1) = 0 And StrComp(FilterName, "SCM Event Log Filter", 1) = 0) Then
                
                bDangerScript = True
                
                If Not bOtherConsumerClass Then
                    If Not bIgnoreAllWhitelists Then
                        If bHideMicrosoft Then
                            If ConsumerName = "BVTConsumer" And cmdExecute = "" And cmdArguments = "cscript KernCap.vbs" Then
                                If cmdWorkDir <> "" Then
                                    If Not FileExists(BuildPath(cmdWorkDir, "KernCap.vbs")) Then bDangerScript = False
                                Else
                                    If FindOnPath("KernCap.vbs") = "" Then bDangerScript = False
                                End If
                            
                            ElseIf FilterName = "BVTFilter" And sEventName = "__InstanceModificationEvent WITHIN 60 WHERE TargetInstance ISA ""Win32_Processor"" AND TargetInstance.LoadPercentage > 99" Then
                        
                                bDangerScript = False
                            End If
                        End If
                    End If
                End If
                
                If bDangerScript And Not bOtherConsumerClass Then 'skip other consumer classes, except "ActiveScriptEventConsumer" and "CommandLineEventConsumer"
                
                    'added more safely cheking
                    NoConsumer = False: NoFilter = False
                    
                    If IsNull(objConsumer) Then
                        NoConsumer = True
                    Else
                        If objConsumer Is Nothing Then NoConsumer = True
                    End If
                    If IsNull(objFilter) Then
                        NoFilter = True
                    Else
                        If objFilter Is Nothing Then NoFilter = True
                    End If
                    
                    sHit = "O25 - WMI Event: " & _
                            IIf(NoConsumer, " (no consumer)", ConsumerName) & " - " & _
                            IIf(NoFilter, " (no filter)", FilterName) & " - " & _
                            sAdditionalInfo
                        
                    If Not IsOnIgnoreList(sHit) Then
                        If g_bCheckSum And 0 <> Len(sScriptFile) Then sHit = sHit & GetFileCheckSum(sScriptFile)
                        
                        With result
                            .Section = "O25"
                            .HitLineW = sHit
                            With .O25
                            
                                .Consumer.Script.File = sScriptFile
                                .Consumer.Script.Text = sScriptText
                                .Consumer.Script.Engine = sScriptEngine
                                
                                .Consumer.Cmd.CommandLine = cmdArguments
                                .Consumer.Cmd.ExecPath = cmdExecute
                                .Consumer.Cmd.WorkDir = cmdWorkDir
                                .Consumer.Cmd.Interactive = bRunInteractively
                                
                                .Consumer.Name = ConsumerName
                                .Consumer.NameSpace = ConsumerNameSpace
                                .Consumer.Path = ConsumerPath
                                .Consumer.KillTimeout = lKillTimeout
                                
                                .Timer.ID = sTimerName
                                .Timer.className = sTimerClassName
                                .Timer.Interval = lTimerInterval
                                
                                .Filter.Query = sFilterQuery
                                .Filter.Name = FilterName
                                .Filter.NameSpace = FilterNameSpace
                                .Filter.Path = FilterPath
                                
                            End With
                            AddCustomToFix .Custom, CUSTOM_ACTION_O25
                            
                            If .O25.Consumer.Type = O25_CONSUMER_ACTIVE_SCRIPT Then
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, .O25.Consumer.Script.File
                            End If
                            
                            'Jump to ...
                            AddJumpFile .Jump, JUMP_FILE, sScriptFile
                            AddJumpFile .Jump, JUMP_FILE, cmdExecute
                            AddJumpFile .Jump, JUMP_FILE, BuildPath(cmdWorkDir, cmdExecute)
                            
                            .CureType = CUSTOM_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                
                End If
                
                Set objConsumer = Nothing
                Set objFilter = Nothing
            
            End If
        Next
        
        Set objBinding = Nothing: Set colBindings = Nothing: Set objService = Nothing
    Next
    
    Set objTimerNamespace = Nothing

    AppendErrorLogCustom "CheckO25Item - End"
    Exit Sub
ErrorHandler:
    If i >= 1 And i <= UBound(aNameSpaces) Then
        ErrorMsg Err, "modMain2_CheckO25Item", "Namespace: " & aNameSpaces(i), "Stady: " & Stady
    Else
        ErrorMsg Err, "modMain2_CheckO25Item", "Stady: " & Stady
    End If
    If inIDE Then Stop: Resume Next
    If ComeBack Then Resume Next
End Sub

'select * from MSFT_SCMEventLogEvent WHERE ...
' -> MSFT_SCMEventLogEvent WHERE ...
Function ExtractEventName(sQuery As String) As String
    Dim pos As Long
    pos = InStr(1, sQuery, "from", 1)
    If pos <> 0 Then
        ExtractEventName = Mid$(sQuery, pos + 5)
    End If
End Function

Private Sub ShutdownScriptEngine()
    
    Dim vProc As Variant
    
    For Each vProc In Array("wscript.exe", "cscript.exe", "mshta.exe", "powershell.exe")
        If ProcessExist(vProc, True) Then
            Proc.ProcessClose ProcessName:=CStr(vProc), Async:=False, TimeOutMs:=1000, SendCloseMsg:=True
        End If
    Next
End Sub

Public Sub FixO25Item(sItem$, result As SCAN_RESULT)
    FixIt result
End Sub
    
'cure WMI infection
Public Sub RemoveSubscriptionWMI(result As SCAN_RESULT)
    
    On Error GoTo ErrorHandler
    
    Dim objService As Object, Finish As Boolean, i As Long
    Dim colBindings As Object, objBinding As Object, objBindingToDelete As Object
    
    ShutdownScriptEngine
    
    With result.O25
    
        On Error Resume Next
        'filter
        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & .Filter.NameSpace)
        objService.Get(.Filter.Path).Delete_
        
        'consumer
        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & .Consumer.NameSpace)
        objService.Get(.Consumer.Path).Delete_
        
        'timer
        If 0 <> Len(.Timer.ID) Then
        
            Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\root\subscription")
            objService.Get(.Timer.className & ".TimerId=" & """" & .Timer.ID & """").Delete_
        
        End If
        
        On Error GoTo ErrorHandler
        
        'remove binding
        
        Dim aNameSpaces() As String
        ReDim aNameSpaces(0)
    
        'get all namespaces for current machine
    
        Call WMI_GetNamespaces("Root", aNameSpaces)
        
        For i = 1 To UBound(aNameSpaces)

            'connecting to namespace
            
            Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & aNameSpaces(i))
            
            DoEvents
            
            'get binding info ( Filter <-> Consumer )
            
            Set colBindings = objService.ExecQuery("SELECT * FROM __FilterToConsumerBinding", "WQL", 16 + 32)
            
            For Each objBinding In colBindings
        
                If Not IsNull(objBinding.Filter) And Not IsNull(objBinding.Consumer) Then
                
                    If objBinding.Filter = .Filter.Path And objBinding.Consumer = .Consumer.Path Then
                
                        Set objBindingToDelete = objBinding
                        Finish = True
                        Exit For
                    End If
                End If
            Next
            
            If Finish Then Exit For
            
        Next
        
        If Not (objBindingToDelete Is Nothing) Then objBindingToDelete.Delete_
        
        Set objBinding = Nothing: Set colBindings = Nothing: Set objService = Nothing
        
    End With
    
    ShutdownScriptEngine
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RemoveSubscriptionWMI", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

'restore O25 backup
Public Function RecoverO25Item(O25 As O25_ENTRY) As Boolean
    On Error GoTo ErrorHandler
    
    Dim objSWbemDateTime As Object, objService As Object, objFilter As Object, objConsumer As Object, objBinding As Object, objTimer As Object
    
    With O25
        'connecting to the namespace
        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\root\subscription")
        
        'create filter class instance __EventFilter (events filter)
        If (0 <> Len(.Filter.Name)) Then
            Set objFilter = objService.Get("__EventFilter").SpawnInstance_()
            'set up filter properties
            objFilter.Name = .Filter.Name
            objFilter.QueryLanguage = "WQL"
            objFilter.Query = .Filter.Query
            'save filter
            objFilter.Put_
        End If
        
        'creating class instance ActiveScriptEventConsumer / CommandLineEventConsumer (events consumer)
        ' & set up consumer properties
        
        If .Consumer.Type = O25_CONSUMER_ACTIVE_SCRIPT Then
        
            Set objConsumer = objService.Get("ActiveScriptEventConsumer").SpawnInstance_()
            
            objConsumer.ScriptingEngine = .Consumer.Script.Engine
            If 0 <> Len(.Consumer.Script.Text) Then ' this check is required !!!
                objConsumer.ScriptText = .Consumer.Script.Text
            End If
            If 0 <> Len(.Consumer.Script.File) Then
                objConsumer.ScriptFilename = .Consumer.Script.File
            End If
            
        ElseIf .Consumer.Type = O25_CONSUMER_COMMAND_LINE Then
        
            Set objConsumer = objService.Get("CommandLineEventConsumer").SpawnInstance_()
        
            If 0 <> Len(.Consumer.Cmd.CommandLine) Then
                objConsumer.CommandLineTemplate = .Consumer.Cmd.CommandLine
            End If
            If 0 <> Len(.Consumer.Cmd.ExecPath) Then
                objConsumer.ExecutablePath = .Consumer.Cmd.ExecPath
            End If
            If 0 <> Len(.Consumer.Cmd.WorkDir) Then
                objConsumer.WorkingDirectory = .Consumer.Cmd.WorkDir
            End If
            objConsumer.RunInteractively = .Consumer.Cmd.Interactive
            
        End If
        
        If Not (objConsumer Is Nothing) Then
            objConsumer.KillTimeout = .Consumer.KillTimeout
            objConsumer.Name = .Consumer.Name
            'save consumer
            objConsumer.Put_
        End If
        
        'binding
        If (0 <> Len(.Filter.Name)) Then
            Set objFilter = objService.Get("__EventFilter.Name=""" & .Filter.Name & """")
        End If
            
        If .Consumer.Type = O25_CONSUMER_ACTIVE_SCRIPT Then
            Set objConsumer = objService.Get("ActiveScriptEventConsumer.Name=""" & .Consumer.Name & """")

        ElseIf .Consumer.Type = O25_CONSUMER_COMMAND_LINE Then
            Set objConsumer = objService.Get("CommandLineEventConsumer.Name=""" & .Consumer.Name & """")
        End If
        
        'creating class instance __FilterToConsumerBinding (binding)
        
        If Not (objFilter Is Nothing) And Not (objConsumer Is Nothing) Then
            Set objBinding = objService.Get("__FilterToConsumerBinding").SpawnInstance_()
            'set up binding properties
            objBinding.Filter = objFilter.Path_
            objBinding.Consumer = objConsumer.Path_
            'save binding
            objBinding.Put_
        End If
        
        'create timer
        'creating timer class instance & set up timer properties
        
        If .Timer.Type = O25_TIMER_ABSOLUTE Then
            Set objTimer = objService.Get("__AbsoluteTimerInstruction").SpawnInstance_()
            
            Set objSWbemDateTime = CreateObject("WbemScripting.SWbemDateTime")
            objSWbemDateTime.SetVarDate .Timer.EventDateTime, True 'true - convert to local time
            objTimer.EventDateTime = objSWbemDateTime.Value
            Set objSWbemDateTime = Nothing
        
        ElseIf .Timer.Type = O25_TIMER_INTERVAL Then
            Set objTimer = objService.Get("__IntervalTimerInstruction").SpawnInstance_()
            objTimer.IntervalBetweenEvents = .Timer.Interval
        End If
        
        If Not (objTimer Is Nothing) Then
            objTimer.SkipIfPassed = False
            objTimer.TimerId = .Timer.ID
            'save timer
            objTimer.Put_
        End If
        
        RecoverO25Item = True
        
        Set objTimer = Nothing
        Set objBinding = Nothing
        Set objConsumer = Nothing
        Set objFilter = Nothing
        Set objService = Nothing
    End With
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modMain2_RecoverO25Item"
    If inIDE Then Stop: Resume Next
End Function


' strip string to length defined
Function StripCode(ByVal sCode As String, Optional Max_Characters As Long = MAX_CODE_LENGTH, Optional AddActualLength As Boolean = True) As String
    On Error GoTo ErrorHandler

    sCode = Replace$(sCode, vbCr, "")
    sCode = Replace$(sCode, vbLf, ChrW$(182) & Space$(1))

    If Len(sCode) <= Max_Characters Then
        StripCode = sCode
    Else
        StripCode = Left$(sCode, Max_Characters) & IIf(AddActualLength, "(" & Len(sCode) & " bytes" & ")", "")
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modMain2_StripCode", sCode
    If inIDE Then Stop: Resume Next
End Function


'Example:
' \\ALEX-PC\ROOT\subscription:ActiveScriptEventConsumer.Name="Dragokas_consumer"
' out_NameSpace <- ROOT\subscription
' out_ClassName <- ActiveScriptEventConsumer
Sub ExtractNameSpaceAndClassNameFromString(sComplexString As String, out_NameSpace As String, Optional out_ClassName As String)
    On Error GoTo ErrorHandler
    Dim pos As Long, pos2 As Long, pos3 As Long
    out_NameSpace = ""
    out_ClassName = ""
    If InStr(1, sComplexString, "\\") = 1 Then
        pos = InStr(3, sComplexString, "\")
        If pos <> 0 Then
            pos2 = InStr(pos, sComplexString, ":")
            If pos2 <> 0 Then
                out_NameSpace = Mid$(sComplexString, pos + 1, pos2 - pos - 1)
                pos3 = InStr(pos2, sComplexString, ".Name", 1)
                If pos3 <> 0 Then
                    out_ClassName = Mid$(sComplexString, pos2 + 1, pos3 - pos2 - 1)
                End If
            End If
        End If
    Else
        pos3 = InStr(1, sComplexString, ".Name", 1)
        If pos3 <> 0 Then
            out_ClassName = Left$(sComplexString, pos3 - 1)
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain2_ExtractNameSpaceAndClassNameFromString", sComplexString
    If inIDE Then Stop: Resume Next
End Sub

'Example:
'__EventFilter.Name="SCM Event Log Filter" -> SCM Event Log Filter
Function GetStringInsideQt(sStr As String) As String
    On Error GoTo ErrorHandler
    Dim pos As Long, pos2 As Long
    pos = InStr(1, sStr, """")
    If pos <> 0 Then
        pos2 = InStr(pos + 1, sStr, """")
        If pos = 0 Then
            GetStringInsideQt = Mid$(sStr, pos + 1)
        Else
            GetStringInsideQt = Mid$(sStr, pos + 1, pos2 - pos - 1)
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modMain2_GetStringInsideQt", sStr
    If inIDE Then Stop: Resume Next
End Function


Public Sub CheckO26Item()
    'O26 - Image File Execution Options:
    
    'Area:
    'HKLM\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options
    'HKLM\Software\Microsoft\Windows\CurrentVersion\PackagedAppXDebug
    '
    
    'Articles:
    'https://docs.microsoft.com/en-us/windows-hardware/drivers/debugger/gflags-overview
    'http://www.alex-ionescu.com/Estoteric%20Hooks.pdf
    '
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO26Item - Begin"
    
    Const FLG_APPLICATION_VERIFIER As Long = &H100&
    
    Dim sKeys$(), sSubkeys$(), i&, j&, sFile$, sHit$, result As SCAN_RESULT
    Dim bDisabled As Boolean, vGFlag As Variant
    Dim bPerUser As Boolean, bIsSafe As Boolean, aTmp() As String, sNonSafe As String, bMissing As Boolean, sAlias As String
    Dim bSafe As Boolean, bShared As Boolean, sOrigLine As String
    
    If bIsWinVistaAndNewer Then
        If IsProcedureAvail("VerifierIsPerUserSettingsEnabled", "Verifier.dll") Then
            bPerUser = VerifierIsPerUserSettingsEnabled()
        Else
            bPerUser = CBool(1 And Reg.GetDword(0&, "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager", "ImageExecutionOptions"))
        End If
    End If
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    If bPerUser Then
        HE.Init HE_HIVE_ALL
    Else
        HE.Init HE_HIVE_HKLM
    End If
    HE.AddKey "Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
    'key is redirected (XP-Vista)
    'key is x64-shared (Win 7+)
    
    Do While HE.MoveNext
            
        sAlias = IIf(bIsWin32, "O26", IIf(HE.Redirected, "O26-32", "O26"))
        
        sKeys = Split(Reg.EnumSubKeys(HE.Hive, HE.Key, HE.Redirected), "|")    'for each image
        
        For i = 0 To UBound(sKeys)

            sFile = Reg.GetString(HE.Hive, HE.Key & "\" & sKeys(i), "Debugger", HE.Redirected)
    
            If sFile <> vbNullString Then
                sFile = FindOnPath(UnQuote(EnvironW(sFile)), True)
        
                If FileExists(sFile) Then
                    sFile = GetLongPath(sFile) '8.3 -> Full
                    bMissing = False
                Else
                    bMissing = True
                End If
                
                bSafe = False
                
                'check by safe list
                If bHideMicrosoft Then
                    If StrComp(sKeys(i), "taskmgr.exe", 1) = 0 Then
                        If InStr(1, GetFileProperty(sFile, "FileDescription"), "Process Explorer", 1) <> 0 Then
                            If IsMicrosoftFile(sFile) Then
                                bSafe = True
                            End If
                        End If
                    End If

                    'exclude default line for WinXP:
                    'O26 - Image File Execution Options: Your Image File Name Here without a path - ntsd -d (file missing)
                    If OSver.MajorMinor <= 5.2 Then
                        If sFile = "ntsd -d" Then
                            If Not FileExists("ntsd") Then
                                bSafe = True
                            End If
                        End If
                    End If
                End If
                
                If (Not bSafe) Or bIgnoreAllWhitelists Then
                    
                    sHit = sAlias & " - Debugger: " & HE.HiveNameAndSID & "\..\" & sKeys(i) & ": [Debugger] = " & sFile & IIf(bMissing, " (file missing)", "")
        
                    If Not IsOnIgnoreList(sHit) Then
          
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                        With result
                            .Section = "O26"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\" & sKeys(i), "Debugger", , HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
                
            End If
            
            'Detecting AVRF Hook
            
            sFile = Reg.GetString(HE.Hive, HE.Key & "\" & sKeys(i), "VerifierDlls", HE.Redirected)
            
            If Len(sFile) <> 0 Then
                
                bDisabled = False
                vGFlag = Reg.GetString(HE.Hive, HE.Key & "\" & sKeys(i), "GlobalFlag", HE.Redirected)
        
                If IsNumeric(vGFlag) Then
                    If Not CBool(CLng(vGFlag) And FLG_APPLICATION_VERIFIER) Then bDisabled = True
                Else
                    If CStr(vGFlag) <> "0x100" Then bDisabled = True
                End If
                
                sFile = FindOnPath(UnQuote(EnvironW(sFile)), True)
                
                If FileExists(sFile) Then
                    sFile = GetLongPath(sFile) '8.3 -> Full
                    bMissing = False
                Else
                    bMissing = True
                End If
                
                sHit = sAlias & " - Debugger: " & HE.HiveNameAndSID & "\..\" & _
                    sKeys(i) & ": [VerifierDlls] = " & sFile & IIf(bMissing, " (file missing)", "") & IIf(bDisabled, " (disabled)", "")
                
                If Not IsOnIgnoreList(sHit) Then
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    
                    With result
                        .Section = "O26"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\" & sKeys(i), "VerifierDlls", , HE.Redirected
                        If Not bDisabled Then
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sKeys(i), , , HE.Redirected
                        End If
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            
            End If
            
        Next i
        
        'AVRF Global Hook
        
        sFile = Reg.GetString(HE.Hive, HE.Key & "\" & "{ApplicationVerifierGlobalSettings}", "VerifierProviders", HE.Redirected)
        
        If sFile <> "" Then
            
            aTmp = SplitSafe(sFile)
            ArrayRemoveEmptyItems aTmp
            
            For i = 0 To UBound(aTmp)
                sFile = aTmp(i)
                sOrigLine = sFile
                
                If InStr(1, "*" & sSafeIfeVerifier & "*", "*" & sFile & "*", 1) = 0 Or bIgnoreAllWhitelists Or (Not bHideMicrosoft) Then
                    
                    sFile = FormatFileMissing(sFile)
                    
                    sHit = sAlias & " - Debugger Global hook: [VerifierProviders] = " & sFile
                    
                    If Not IsOnIgnoreList(sHit) Then
                        
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                        
                        With result
                            .Section = "O26"
                            .HitLineW = sHit
                            
                            AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE, _
                                HE.Hive, HE.Key & "\" & "{ApplicationVerifierGlobalSettings}", "VerifierProviders", , HE.Redirected, , _
                                sOrigLine, "", " "
                            
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        End If
    Loop
    
    'Set HE = Nothing
    
    'Check for UWP debuggers
    
    If OSver.IsWindows10OrGreater Then
    
        'Win10 only (WinRT apps are not supported)
    
        For i = 1 To Reg.EnumSubKeysToArray(HKCU, "Software\Microsoft\Windows\CurrentVersion\PackagedAppXDebug", sKeys())
            
            sFile = Reg.GetString(HKCU, "Software\Microsoft\Windows\CurrentVersion\PackagedAppXDebug\" & sKeys(i), "")
            
            sFile = FormatFileMissing(sFile)
                    
            sHit = sAlias & " - UWP Debugger: " & sKeys(i) & " (default) = " & sFile
            
            If Not IsOnIgnoreList(sHit) Then
                
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                With result
                    .Section = "O26"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HKCU, "Software\Microsoft\Windows\CurrentVersion\PackagedAppXDebug\" & sKeys(i)
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        Next
        
        For i = 1 To Reg.EnumSubKeysToArray(HKCU, "Software\Classes\ActivatableClasses\Package", sKeys())
            
            For j = 1 To Reg.EnumSubKeysToArray(HKCU, "Software\Classes\ActivatableClasses\Package\" & sKeys(i) & "\DebugInformation", sSubkeys())
            
                sFile = Reg.GetString(HKCU, "Software\Classes\ActivatableClasses\Package\" & sKeys(i) & "\DebugInformation\" & sSubkeys(j), "DebugPath")
                
                sFile = FormatFileMissing(sFile)
                        
                sHit = sAlias & " - UWP Debugger: " & sKeys(i) & " [DebugPath] = " & sFile
                
                If Not IsOnIgnoreList(sHit) Then
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    
                    With result
                        .Section = "O26"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HKCU, "Software\Classes\ActivatableClasses\Package\" & sKeys(i) & "\DebugInformation"
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            Next
        Next
    End If
    
    CheckO26ToolsHiJack
    
    AppendErrorLogCustom "CheckO26Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain2_CheckO26Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO26ToolsHiJack()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO26ToolsHiJack - Begin"
    
    'Area:
    'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Accessibility\ATs
    'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer
    '
    
    'Articles:
    'https://malware.news/t/too-accessible-how-crowdstrike-falcon-detects-and-prevents-windows-logon-bypasses/18539
    '
    
    Dim sFile$, sHit$, result As SCAN_RESULT
    Dim aKey$(), sRestore$
    Dim i As Long, j As Long, bSafe As Boolean, bUseSFC As Boolean
    
    Dim dSafe As clsTrickHashTable
    Set dSafe = New clsTrickHashTable
    dSafe.CompareMode = 1
    
    'LogonScreen Hijack
    
    dSafe.Add "magnifierpane", "%SystemRoot%\System32\Magnify.exe"
    dSafe.Add "Narrator", "%SystemRoot%\System32\Narrator.exe"
    dSafe.Add "osk", "%SystemRoot%\System32\osk.exe"
    dSafe.Add "Oracle_JavaAccessBridge", "*" 'any
    dSafe.Add "animations", "13"
    dSafe.Add "audiodescription", "12"
    dSafe.Add "caretbrowsing", "21"
    dSafe.Add "caretwidth", "8"
    dSafe.Add "colorfiltering", "22"
    dSafe.Add "cursorscheme", ""
    dSafe.Add "filterkeys", "0"
    dSafe.Add "focusborderheight", "6"
    dSafe.Add "focusborderwidth", "7"
    dSafe.Add "highcontrast", "1"
    dSafe.Add "keyboardcues", "9"
    dSafe.Add "keyboardpref", "10"
    dSafe.Add "messageduration", "17"
    dSafe.Add "minimumhitradius", "18"
    dSafe.Add "mousekeys", "2"
    dSafe.Add "overlappedcontent", "11"
    dSafe.Add "showsounds", "19"
    dSafe.Add "soundsentry", "3"
    dSafe.Add "stickykeys", "4"
    dSafe.Add "togglekeys", "5"
    dSafe.Add "windowarranging", "20"
    dSafe.Add "windowtracking", "14"
    dSafe.Add "windowtrackingtimeout", "16"
    dSafe.Add "windowtrackingzorder", "15"
    
    'HKCU is not affected
    'Other subkeys (with numeric StartExe) is affected!
    
    For i = 1 To Reg.EnumSubKeysToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Accessibility\ATs", aKey())
        
        bSafe = False
        bUseSFC = False
        sRestore = vbNullString
        
        sFile = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Accessibility\ATs\" & aKey(i), "StartExe")
        
        If dSafe.Exists(aKey(i)) Then
        
            sRestore = dSafe(aKey(i))
            
            If sRestore = "*" Then
            
                bSafe = True
                sRestore = ""
                
            ElseIf StrComp(sFile, EnvironW(sRestore), 1) = 0 Then
                
                If IsNumeric(sFile) Or Len(sFile) = 0 Then
                
                    bSafe = True
                
                Else
                    bUseSFC = True
                    
                    If IsMicrosoftFile(sFile) Then
                        bSafe = True
                    End If
                End If
            End If
        Else
            'not in database
            
            If IsNumeric(sFile) Then
                
                bSafe = True
            
            Else
                bUseSFC = True
                
                If IsMicrosoftFile(sFile) Then
            
                    bSafe = True
                End If
            End If
        End If
        
        If Not bSafe Then

            sFile = FormatFileMissing(sFile)
            
            sHit = "O26 - Tools: " & "HKLM\" & "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Accessibility\ATs\" & aKey(i) & _
                " [StartExe] = " & sFile
            
            If Not IsOnIgnoreList(sHit) Then
                
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                With result
                    .Section = "O26"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Accessibility\ATs\" & aKey(i), _
                        "StartExe", sRestore, , IIf(IsNumeric(sRestore), REG_RESTORE_SZ, REG_RESTORE_EXPAND_SZ)
                    If bUseSFC Then
                        AddFileToFix .File, RESTORE_FILE_SFC, EnvironW(sRestore)
                    End If
                    
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    Set dSafe = Nothing
    
    Dim sBackupPath As String
    Dim sCleanupPath As String
    Dim sDefragPath As String
    Dim sRootFile As String
    
    If OSver.IsWindows10OrGreater Then
        sBackupPath = vbNullString
        sCleanupPath = "%SystemRoot%\System32\cleanmgr.exe /D %c"
        sDefragPath = "%systemroot%\system32\dfrgui.exe"
    ElseIf OSver.IsWindows8OrGreater Then
        sBackupPath = vbNullString
        sCleanupPath = "%SystemRoot%\System32\cleanmgr.exe /D %c"
        sDefragPath = "%systemroot%\system32\dfrgui.exe"
    ElseIf OSver.IsWindows7OrGreater Then
        sBackupPath = "%SystemRoot%\system32\sdclt.exe"
        sCleanupPath = "%SystemRoot%\System32\cleanmgr.exe /D %c"
        sDefragPath = "%systemroot%\system32\dfrgui.exe"
    ElseIf OSver.IsWindowsXPOrGreater Then
        sBackupPath = "%SystemRoot%\system32\ntbackup.exe"
        sCleanupPath = "%SystemRoot%\System32\cleanmgr.exe /D %c"
        sDefragPath = "%SystemRoot%\system32\dfrg.msc %c:"
    Else
        sBackupPath = vbNullString
        sCleanupPath = "%SystemRoot%\System32\cleanmgr.exe /D %c"
        sDefragPath = "%SystemRoot%\system32\dfrg.msc %c:"
    End If
    
    Dim sData As String, sArgs As String, sKey As String
    
    Set dSafe = New clsTrickHashTable
    dSafe.CompareMode = 1
    
    dSafe.Add "BackupPath", sBackupPath
    dSafe.Add "cleanuppath", sCleanupPath
    dSafe.Add "DefragPath", sDefragPath
    
    For i = 0 To dSafe.Count - 1
        
        sKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\" & dSafe.Keys(i)
        sData = Reg.GetString(HKLM, sKey, "")
        
        SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
        sRootFile = sFile
        sFile = FormatFileMissing(sFile, sArgs)
        
        bSafe = True
        If StrComp(sData, EnvironW(dSafe.Items(i)), 1) <> 0 Then
            bSafe = False
        Else
            If Len(sRootFile) <> 0 Then
                If FileMissing(sFile) Then bSafe = False
            End If
            If StrEndWith(sRootFile, ".exe") Then
                If Not IsMicrosoftFile(sRootFile) Then bSafe = False
            End If
        End If
        
        If Not bSafe Then
            
            sHit = "O26 - Tools: " & "HKLM\" & sKey & " (default) = " & ConcatFileArg(sFile, sArgs)
            
            If Not IsOnIgnoreList(sHit) Then
            
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                With result
                    .Section = "O26"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HKLM, sKey, "", dSafe.Items(i), , REG_RESTORE_EXPAND_SZ
                    
                    If FileMissing(sFile) And Len(sRootFile) <> 0 Then
                        AddFileToFix .File, RESTORE_FILE_SFC, EnvironW(sRootFile)
                    End If
                    
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    Set dSafe = Nothing
    
    AppendErrorLogCustom "CheckO26ToolsHiJack - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain2_CheckO26ToolsHiJack"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO26Item(sItem$, result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    FixRegistryHandler result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain2_FixO26Item", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub
