Attribute VB_Name = "modMain_2"
'[modMain_2.bas]

'
' Core check / Fix Engine
'
' (part 2: O25 - O27)

'
' O25, O26, O27 by Alex Dragokas
'

'O25 - Windows Management Instrumentation (WMI) event consumers
'O26 - Image File Execution Options (IFEO) and System Tools hijack
'O27 - Account & Remote Desktop Protocol

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
    Dim ComeBack As Boolean, NoConsumer As Boolean, NoFilter As Boolean, bOtherConsumerClass As Boolean
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
    On Error Resume Next
    Set objTimerNamespace = CreateObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\root\subscription")
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler:
        Set objTimerNamespace = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\root\subscription")
    End If
    On Error GoTo ErrorHandler:
    
    'get all namespaces for current machine
    
    'Call WMI_GetNamespaces("Root", aNameSpaces)
    'let's concentrate on actual malware method
    ReDim aNameSpaces(1)
    aNameSpaces(1) = "root\subscription"
    
    For i = 1 To UBound(aNameSpaces)

        'connecting to namespace

        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & aNameSpaces(i))

        If Not bAutoLogSilent Then DoEvents
    
        'get binding info ( Filter <-> Consumer )
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
                If StrComp(ConsumerNameSpace, aNameSpaces(i), 1) = 0 Then
                    'if consumer's namespace is a same
                    Set objServiceConsumer = objService
                Else
                    On Error Resume Next
                    Set objServiceConsumer = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & ConsumerNameSpace)
                End If
                
                On Error Resume Next
                Set objConsumer = objServiceConsumer.Get(ConsumerPath)
                On Error GoTo ErrorHandler:
                
                cmdExecute = vbNullString
                cmdWorkDir = vbNullString
                cmdArguments = vbNullString
                sScriptFile = vbNullString
                sScriptText = vbNullString
                sAdditionalInfo = vbNullString
                sScriptEngine = vbNullString
                sTimerName = vbNullString
                lTimerInterval = 0
                lKillTimeout = 0
                bRunInteractively = False
                
                If Not (objConsumer Is Nothing) Then
                    
                    'Checking several known classes on: root\subscription
                    'to provide a bit more information to log
                    
                    bOtherConsumerClass = False
                    
                    If StrComp(ConsumerClassName, "ActiveScriptEventConsumer", 1) = 0 Then
                    
                        result.O25.Consumer.Type = O25_CONSUMER_ACTIVE_SCRIPT
                    
                        'Debug.Print objConsumer.ScriptingEngine    'language (engine)
                        'Debug.Print objConsumer.ScriptFileName     'external file
                        'Debug.Print objConsumer.ScriptText         'embedded script code
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
                            sAdditionalInfo = sAdditionalInfo & IIf(Len(sAdditionalInfo) <> 0, " / ", vbNullString) & StripCode(sScriptText)
                        End If

                    ElseIf StrComp(ConsumerClassName, "CommandLineEventConsumer", 1) = 0 Then
                    
                        result.O25.Consumer.Type = O25_CONSUMER_COMMAND_LINE
                        
                        'Example:
                        'kernrate: O25 - WMI Event: BVTConsumer / BVTFilter -> Executable="", Arguments="cscript KernCap.vbs"
                        'https://pcsxcetrasupport3.wordpress.com/2011/10/23/event-10-mystery-solved/
                        
                        'Debug.Print objConsumer.ExecutablePath         'main execution module
                        'Debug.Print objConsumer.CommandLineTemplate    'arguments
                        'debug.print objConsumer.WorkingDirectory       'Work Dir.
                        ComeBack = True
                        If Not IsNull(objConsumer.ExecutablePath) Then cmdExecute = objConsumer.ExecutablePath
                        If Not IsNull(objConsumer.WorkingDirectory) Then cmdWorkDir = objConsumer.WorkingDirectory
                        If Not IsNull(objConsumer.CommandLineTemplate) Then cmdArguments = objConsumer.CommandLineTemplate
                        If Not IsNull(objConsumer.RunInteractively) Then bRunInteractively = objConsumer.RunInteractively
                        If Not IsNull(objConsumer.KillTimeout) Then lKillTimeout = CLng(Val(objConsumer.KillTimeout))
                        
                        'sAdditionalInfo = "Executable=" & """" & cmdExecute & """" & _
                        '    ", WorkDir=" & """" & cmdWorkDir & """" & _
                        '    ", Arguments=" & """" & StripCode(cmdArguments) & """"
                           
                        sAdditionalInfo = cmdExecute & " " & cmdArguments & IIf(Len(cmdWorkDir) <> 0, " (WorkDir = " & cmdWorkDir & ")", vbNullString)
                        
                        ComeBack = False
                        
'                    ElseIf StrComp(ConsumerClassName, "LogFileEventConsumer", 1) = 0 Then
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
'                        'other consumers -> Show Namespace + ClassName
'                        sAdditionalInfo = "ClassName=" & """" & ConsumerNameSpace & ":" & ConsumerClassName & """"
                    Else
                        bOtherConsumerClass = True
                    End If
                    
                End If
                
                'Trying to find associated timer inside the filter
                
                'connecting to filter's own namespace
                
                If StrComp(FilterNameSpace, aNameSpaces(i), 1) = 0 Then
                    'if consumer's namespace is a same
                    Set objServiceFilter = objService
                Else
                    On Error Resume Next
                    Set objServiceFilter = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & FilterNameSpace)
                End If
                
                On Error Resume Next
                Set objFilter = objServiceFilter.Get(FilterPath)
                On Error GoTo ErrorHandler:
                
                If Not (objFilter Is Nothing) Then
                
                    If Not IsNull(objFilter.Query) Then sFilterQuery = objFilter.Query
                
                    'receives events from timer ?
                    If InStr(1, sFilterQuery, "__timerevent", 1) <> 0 Then
                    
                        'SELECT * FROM __timerevent WHERE timerid="Dragokas_WMITimer2"
                        sTimerName = GetStringInsideQt(sFilterQuery)
                    
                        If 0 <> Len(sTimerName) Then
                            'searching timer's Class name (2 options)
                            
                            Set objTimer = Nothing
                            
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
                                sTimerClassName = vbNullString
                                sTimerName = vbNullString
                            Else
                                Set objTimer = Nothing
                            End If
                        
                        End If
                    
                    Else
                        'if another event source -> print its name
                        sEventName = ExtractEventName(sFilterQuery)
                        If 0 <> Len(sEventName) Then
                            sAdditionalInfo = "Event=" & """" & sEventName & """" & ", " & sAdditionalInfo
                        End If
                    End If
                
                End If
            
                'WhiteList
                'If Not (StrComp(ConsumerClassName, "NTEventLogEventConsumer", 1) = 0 And StrComp(FilterName, "SCM Event Log Filter", 1) = 0) Then
                
                bDangerScript = True
                
                If Not bOtherConsumerClass Then
                    If Not bIgnoreAllWhitelists Then
                        If bHideMicrosoft Then
                            If ConsumerName = "BVTConsumer" And Len(cmdExecute) = 0 And cmdArguments = "cscript KernCap.vbs" Then
                                If Len(cmdWorkDir) <> 0 Then
                                    If Not FileExists(BuildPath(cmdWorkDir, "KernCap.vbs")) Then bDangerScript = False
                                Else
                                    If Len(FindOnPath("KernCap.vbs")) = 0 Then bDangerScript = False
                                End If
                            
                            'ElseIf FilterName = "BVTFilter" And sEventName = "__InstanceModificationEvent WITHIN 60 WHERE TargetInstance ISA ""Win32_Processor"" AND TargetInstance.LoadPercentage > 99" Then
                            ElseIf FilterName = "BVTFilter" And sEventName = Caes_Decode("`bNuBEnCtxbLCJINJJ_V^_rk\go VJWMPW 73 j]\k` sH[RRctahkZi`d LXH ""fzG01xkUTJN^`^c"" tIA UdwnnEVCJMvKBF.kVJOwTcVZem\dd > 80") Then
                        
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
                        
                    If 0 <> Len(sScriptFile) Then
                    
                        SignVerifyJack sScriptFile, result.SignResult
                        
                        sHit = sHit & FormatSign(result.SignResult)
                        
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sScriptFile)
                    End If
                        
                    If Not IsOnIgnoreList(sHit) Then
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
                                
                                .Timer.id = sTimerName
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
                            AddJumpFile .Jump, JUMP_FILE, cmdExecute, cmdArguments
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
        ErrorMsg Err, "modMain2_CheckO25Item", "Namespace: " & aNameSpaces(i)
    Else
        ErrorMsg Err, "modMain2_CheckO25Item"
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
        ExtractEventName = mid$(sQuery, pos + 5)
    End If
End Function

Private Sub ShutdownScriptEngine()
    
    Dim vProc As Variant
    
    For Each vProc In Array("cmd.exe", "wscript.exe", "cscript.exe", "mshta.exe", "powershell.exe")
        If ProcessExist(vProc, True) Then
            Proc.ProcessClose ProcessName:=CStr(vProc), Async:=False, TimeoutMs:=1000, SendCloseMsg:=True
        End If
    Next
End Sub

Public Sub FixO25Item(sItem$, result As SCAN_RESULT)
    FixIt result
End Sub
    
'cure WMI infection
Public Sub RemoveSubscriptionWMI(O25 As O25_ENTRY)
    
    On Error GoTo ErrorHandler
    
    Dim objService As Object, Finish As Boolean, i As Long
    Dim colBindings As Object, objBinding As Object, objBindingToDelete As Object
    
    ShutdownScriptEngine
    
    With O25
    
        On Error Resume Next
        'filter
        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & .Filter.NameSpace)
        objService.Get(.Filter.Path).Delete_
        
        'consumer
        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & .Consumer.NameSpace)
        objService.Get(.Consumer.Path).Delete_
        
        'timer
        If 0 <> Len(.Timer.id) Then
        
            Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\root\subscription")
            objService.Get(.Timer.className & ".TimerId=" & """" & .Timer.id & """").Delete_
        
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
    ErrorMsg Err, "RemoveSubscriptionWMI", O25.Consumer.Name
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
            objTimer.TimerId = .Timer.id
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

    sCode = Replace$(sCode, vbCr, vbNullString)
    sCode = Replace$(sCode, vbLf, ChrW$(182) & Space$(1))

    If Len(sCode) <= Max_Characters Then
        StripCode = sCode
    Else
        StripCode = Left$(sCode, Max_Characters) & IIf(AddActualLength, "(" & Len(sCode) & " bytes" & ")", vbNullString)
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
    out_NameSpace = vbNullString
    out_ClassName = vbNullString
    If InStr(1, sComplexString, "\\") = 1 Then
        pos = InStr(3, sComplexString, "\")
        If pos <> 0 Then
            pos2 = InStr(pos, sComplexString, ":")
            If pos2 <> 0 Then
                out_NameSpace = mid$(sComplexString, pos + 1, pos2 - pos - 1)
                pos3 = InStr(pos2, sComplexString, ".Name", 1)
                If pos3 <> 0 Then
                    out_ClassName = mid$(sComplexString, pos2 + 1, pos3 - pos2 - 1)
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
            GetStringInsideQt = mid$(sStr, pos + 1)
        Else
            GetStringInsideQt = mid$(sStr, pos + 1, pos2 - pos - 1)
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
    
    Dim sKeys$(), sSubkeys$(), i&, j&, sFile$, sArgs$, sHit$, sData$, result As SCAN_RESULT
    Dim bDisabled As Boolean, sGFlag As String
    Dim bPerUser As Boolean, aTmp() As String, sAlias As String
    Dim bSafe As Boolean, sOrigLine As String
    
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
            
        sAlias = BitPrefix("O26", HE)
        
        sKeys = Split(Reg.EnumSubKeys(HE.Hive, HE.Key, HE.Redirected), "|")    'for each image
        
        For i = 0 To UBound(sKeys)

            sData = Reg.GetString(HE.Hive, HE.Key & "\" & sKeys(i), "Debugger", HE.Redirected)
            
            If Len(sData) <> 0 Then
            
                SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                sFile = FormatFileMissing(sFile, sArgs)
                
                bSafe = False
                SignVerifyJack sFile, result.SignResult
                
                'check by safe list
                If bHideMicrosoft Then
                    If StrComp(sKeys(i), "taskmgr.exe", 1) = 0 Then
                        If InStr(1, GetFileProperty(sFile, "FileDescription"), "Process Explorer", 1) <> 0 Then
                            If result.SignResult.isMicrosoftSign Then
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
                    
                    SignVerifyJack sFile, result.SignResult
                    
                    sHit = sAlias & " - Debugger: " & HE.HiveNameAndSID & "\..\" & sKeys(i) & ": [Debugger] = " & _
                        ConcatFileArg(sFile, sArgs) & FormatSign(result.SignResult)
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O26"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\" & sKeys(i), "Debugger", , HE.Redirected
                            AddJumpFile .Jump, JUMP_FILE, sFile
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
                
            End If
            
            'Detecting AVRF Hook
            
            sData = Reg.GetString(HE.Hive, HE.Key & "\" & sKeys(i), "VerifierDlls", HE.Redirected)
            
            If Len(sData) <> 0 Then
                
                bDisabled = False
                sGFlag = Reg.GetString(HE.Hive, HE.Key & "\" & sKeys(i), "GlobalFlag", HE.Redirected)
                
                If 0 = (HexStringToNumber(sGFlag) And FLG_APPLICATION_VERIFIER) Then bDisabled = True

                SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                sFile = FormatFileMissing(sFile, sArgs)
                
                SignVerifyJack sFile, result.SignResult
                
                sHit = sAlias & " - Debugger: " & HE.HiveNameAndSID & "\..\" & _
                    sKeys(i) & ": [VerifierDlls] = " & ConcatFileArg(sFile, sArgs) & FormatSign(result.SignResult) & _
                    IIf(bDisabled, " (disabled)", vbNullString)
                
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O26"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\" & sKeys(i), "VerifierDlls", , HE.Redirected
                        If Not bDisabled Then
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sKeys(i), , , HE.Redirected
                        End If
                        AddJumpFile .Jump, JUMP_FILE, sFile
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            
            End If
            
        Next i
        
        'AVRF Global Hook
        
        sFile = Reg.GetString(HE.Hive, HE.Key & "\" & "{ApplicationVerifierGlobalSettings}", "VerifierProviders", HE.Redirected)
        
        If Len(sFile) <> 0 Then
            
            aTmp = SplitSafe(sFile)
            ArrayRemoveEmptyItems aTmp
            
            For i = 0 To UBound(aTmp)
                sFile = aTmp(i)
                sOrigLine = sFile
                
                If InStr(1, "*" & sSafeIfeVerifier & "*", "*" & sFile & "*", 1) = 0 Or bIgnoreAllWhitelists Or (Not bHideMicrosoft) Then
                    
                    sFile = FormatFileMissing(sFile)
                    
                    SignVerifyJack sFile, result.SignResult
                    
                    sHit = sAlias & " - Debugger Global hook: [VerifierProviders] = " & sFile & FormatSign(result.SignResult)
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O26"
                            .HitLineW = sHit
                            
                            AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE, _
                                HE.Hive, HE.Key & "\" & "{ApplicationVerifierGlobalSettings}", "VerifierProviders", , HE.Redirected, , _
                                sOrigLine, vbNullString, " "
                            
                            AddJumpFile .Jump, JUMP_FILE, sFile
                            
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
            
            sFile = Reg.GetString(HKCU, "Software\Microsoft\Windows\CurrentVersion\PackagedAppXDebug\" & sKeys(i), vbNullString)
            
            sFile = FormatFileMissing(sFile)
            
            SignVerifyJack sFile, result.SignResult
            
            sHit = sAlias & " - UWP Debugger: " & sKeys(i) & " (default) = " & sFile & FormatSign(result.SignResult)
            
            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O26"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HKCU, "Software\Microsoft\Windows\CurrentVersion\PackagedAppXDebug\" & sKeys(i)
                    AddJumpFile .Jump, JUMP_FILE, sFile
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        Next
        
        For i = 1 To Reg.EnumSubKeysToArray(HKCU, "Software\Classes\ActivatableClasses\Package", sKeys())
            
            For j = 1 To Reg.EnumSubKeysToArray(HKCU, "Software\Classes\ActivatableClasses\Package\" & sKeys(i) & "\DebugInformation", sSubkeys())
            
                sFile = Reg.GetString(HKCU, "Software\Classes\ActivatableClasses\Package\" & sKeys(i) & "\DebugInformation\" & sSubkeys(j), "DebugPath")
                
                sFile = FormatFileMissing(sFile)
                
                SignVerifyJack sFile, result.SignResult
                
                sHit = sAlias & " - UWP Debugger: " & sKeys(i) & " [DebugPath] = " & sFile & FormatSign(result.SignResult)
                
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O26"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HKCU, "Software\Classes\ActivatableClasses\Package\" & sKeys(i) & "\DebugInformation"
                        AddJumpFile .Jump, JUMP_FILE, sFile
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
    Dim i As Long, bSafe As Boolean, bUseSFC As Boolean
    
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
    dSafe.Add "cursorscheme", vbNullString
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
                sRestore = vbNullString
                
            ElseIf StrComp(sFile, EnvironW(sRestore), 1) = 0 Then
                
                If IsNumeric(sFile) Or Len(sFile) = 0 Then
                
                    bSafe = True
                
                Else
                    bUseSFC = True
                    
                    If SignVerifyJack(sFile, result.SignResult) And result.SignResult.isMicrosoftSign Then
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
                
                If SignVerifyJack(sFile, result.SignResult) And result.SignResult.isMicrosoftSign Then
            
                    bSafe = True
                End If
            End If
        End If
        
        If Not bSafe Then

            sFile = FormatFileMissing(sFile)
            
            sHit = "O26 - Tools: " & "HKLM\" & "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Accessibility\ATs\" & aKey(i) & _
                " [StartExe] = " & sFile & FormatSign(result.SignResult)
            
            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
            
            If Not IsOnIgnoreList(sHit) Then
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
    ElseIf OSver.IsWindowsVistaOrGreater Then
        sBackupPath = "%SystemRoot%\system32\sdclt.exe"
        sCleanupPath = "%SystemRoot%\System32\cleanmgr.exe /D %c"
        sDefragPath = "%systemroot%\system32\dfrgui.exe"
    ElseIf OSver.IsWindowsXPOrGreater Then
        sBackupPath = "%SystemRoot%\system32\sdclt.exe"
        sCleanupPath = "%SystemRoot%\System32\cleanmgr.exe /D %c"
        sDefragPath = "%SystemRoot%\system32\dfrg.msc %c:"
    Else
        sBackupPath = vbNullString
        sCleanupPath = "%SystemRoot%\System32\cleanmgr.exe /D %c"
        sDefragPath = "%SystemRoot%\system32\dfrg.msc %c:"
    End If
    
    If OSver.IsServer Then
        If OSver.MajorMinor >= 6.1 Then
            sBackupPath = sWinSysDir & "\wbadmin.msc"
        End If
    End If
    
    Dim sData As String, sArgs As String, sKey As String
    
    Set dSafe = New clsTrickHashTable
    dSafe.CompareMode = 1
    
    dSafe.Add "BackupPath", sBackupPath
    dSafe.Add "cleanuppath", sCleanupPath
    dSafe.Add "DefragPath", sDefragPath
    
    For i = 0 To dSafe.Count - 1
        
        sKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\" & dSafe.Keys(i)
        sData = Reg.GetString(HKLM, sKey, vbNullString)
        
        SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
        sFile = FormatFileMissing(sFile, sArgs)
        
        WipeSignResult result.SignResult
        bSafe = True
        
        If StrComp(EnvironW(sData), EnvironW(dSafe.Items(i)), 1) <> 0 Then bSafe = False
        
        If bSafe And Len(dSafe.Items(i)) <> 0 Then
            If FileMissing(sFile) Then
                bSafe = False
            Else
                If StrEndWith(sFile, ".exe") Then
                    SignVerifyJack sFile, result.SignResult
                    If Not result.SignResult.isMicrosoftSign Then bSafe = False
                End If
            End If
        End If
        
        If Not bSafe Then
            
            sHit = "O26 - Tools: " & "HKLM\" & sKey & " (default) = " & ConcatFileArg(sFile, sArgs) & FormatSign(result.SignResult)
            
            If bDebugMode Then sHit = sHit & " (expected: " & dSafe.Items(i) & ")"
            
            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O26"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HKLM, sKey, vbNullString, dSafe.Items(i), , REG_RESTORE_EXPAND_SZ
                    AddFileToFix .File, RESTORE_FILE_SFC, EnvironW(RemoveArguments(dSafe.Items(i))), sArgs
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
    FixIt result
End Sub

Public Sub CheckO27Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO27Item - Begin"
    
    CheckO27Item_RDP
    
    Dim lData&
    Dim sHit$, result As SCAN_RESULT
    Dim HE As clsHiveEnum:      Set HE = New clsHiveEnum
    Dim DC As clsDataChecker:   Set DC = New clsDataChecker
    
    HE.Init HE_HIVE_HKLM, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    
    DC.AddValueData "dontdisplaylastusername", 0
    
    'O27 - Account: (Other)
    Do While HE.MoveNext
        Do While DC.MoveNext
            lData = Reg.GetDword(HE.Hive, HE.Key, DC.ValueName, HE.Redirected)
            
            If Not DC.ContainsData(lData) Then
                sHit = "O27 - Account: (Other) " & HE.KeyAndHivePhysical & ": " & "[" & DC.ValueName & "] = " & Reg.StatusCodeDescOnFail(lData)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O27"
                        .HitLineW = sHit
                        AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, DC.ValueName, DC.DataLong, HE.Redirected, REG_RESTORE_DWORD
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Loop
    Loop
    
    'O27 - Account: (RDP Group)
    Dim sRdpSid As String, sRdpGroup As String
    sRdpSid = "S-1-5-32-555"
    sRdpGroup = MapSIDToUsername(sRdpSid)
    Dim i As Long
    For i = 0 To UBound(g_LocalUserNames)
        If IsUserMembershipInGroup(g_LocalUserNames(i), sRdpGroup) Then
            sHit = "O27 - Account: (RDP Group) User '" & g_LocalUserNames(i) & "' is a member of Remote desktop group"
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O27"
                    .HitLineW = sHit
                    AddCustomToFix .Custom, CUSTOM_ACTION_REMOVE_GROUP_MEMBERSHIP, sRdpGroup, , , g_LocalUserNames(i)
                    .CureType = CUSTOM_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    'O27 - Account: (AutoLogon)
    Call CheckAutoLogon
    
    'O27 - Account: (Missing)
    Dim sidList() As String
    Dim aProfiles() As String
    Dim sProfilePath As String
    Dim sKey As String
    
    sKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    
    For i = 1 To Reg.EnumSubKeysToArray(HKLM, sKey, sidList)
        
        sProfilePath = Reg.GetString(HKLM, sKey & "\" & sidList(i), "ProfileImagePath")
        ArrayAddStr aProfiles, sProfilePath
        
        If Not FolderExists(sProfilePath) Then
            sHit = "O27 - Account: (Missing) HKLM\..\ProfileList\" & sidList(i) & " [ProfileImagePath] = " & sProfilePath & " " & STR_FOLDER_MISSING
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O27"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, sKey & "\" & sidList(i)
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    'O27 - Account: (Bad profile)
    If Len(ProfilesDir) <> 0 Then
        Dim aFolders() As String
        Dim sName As String
        Dim aWhiteUsers(3) As String
        aWhiteUsers(0) = "Default"
        aWhiteUsers(1) = "Public"
        aWhiteUsers(2) = "All Users"
        aWhiteUsers(3) = "Default User"
        
        aFolders = ListSubfolders(ProfilesDir) '%SystemDrive%\Users\*
        
        For i = 0 To UBound(aFolders)
            sName = GetFileName(aFolders(i), True)
            If Not InArray(sName, aWhiteUsers, , , vbTextCompare) Then
                If Not InArray(aFolders(i), aProfiles, , , vbTextCompare) Then
                    sHit = "O27 - Account: (Bad profile) Folder is not referenced by any of user SIDs: " & aFolders(i)
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O27"
                            .HitLineW = sHit
                            AddFileToFix .File, REMOVE_FOLDER, aFolders(i)
                            .CureType = FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
        Next
    End If
    
    'O27 - Account: (Hidden)
    Dim aValue() As String
    sKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\" & Caes_Decode("TsjjrlyPtvJRMUVAv\P_uZfi") 'SpecialAccounts\UserList"
    For i = 1 To Reg.EnumValuesToArray(HKLM, sKey, aValue(), False)
        sHit = "O27 - Account: (Hidden) User '" & aValue(i) & "' is invisible on logon screen"
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O27"
                .HitLineW = sHit
                AddRegToFix .Reg, REMOVE_VALUE, HKLM, sKey, aValue(i)
                .CureType = REGISTRY_BASED
            End With
            AddToScanResults result
        End If
    Next
    
    AppendErrorLogCustom "CheckO27Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO27Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO27Item_RDP()
    On Error GoTo ErrorHandler:
    
    Dim lData&
    Dim sHit$, result As SCAN_RESULT
    Dim HE As clsHiveEnum:      Set HE = New clsHiveEnum
    Dim DC As clsDataChecker:   Set DC = New clsDataChecker
    
    HE.Init HE_HIVE_HKLM, , HE_REDIR_NO_WOW
    HE.AddKey "SYSTEM\CurrentControlSet\Control\Terminal Server"
    
    DC.AddValueData "AllowRemoteRPC", 0
    DC.AddValueData "fDenyTSConnections", 1
    If (Not OSver.IsServer) And OSver.IsWindows7OrGreater Then
        DC.AddValueData "fSingleSessionPerUser", 1
    End If
    
    'O27 - RDP: (Other)
    Do While HE.MoveNext
        Do While DC.MoveNext
            lData = Reg.GetDword(HE.Hive, HE.Key, DC.ValueName, HE.Redirected)
            
            If Not DC.ContainsData(lData) Then
                sHit = "O27 - RDP: (Other) " & HE.KeyAndHivePhysical & ": " & "[" & DC.ValueName & "] = " & Reg.StatusCodeDescOnFail(lData)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O27"
                        .HitLineW = sHit
                        AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, DC.ValueName, DC.DataLong, HE.Redirected, REG_RESTORE_DWORD
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Loop
    Loop
    
    'https://learn.microsoft.com/en-us/troubleshoot/windows-server/remote/shadow-terminal-server-session
    
    HE.Clear: DC.Clear
    
    HE.AddKey "SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"
    HE.AddKey "SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"
    
    '0 - Disable shadow
    '1 - Full access with user's permission
    DC.AddValueData "Shadow", Array(0, 1)
    
    Do While HE.MoveNext
        Do While DC.MoveNext
            lData = Reg.GetDword(HE.Hive, HE.Key, DC.ValueName, HE.Redirected)
            
            If Not DC.ContainsData(lData) Then
                sHit = "O27 - RDP: (Other) " & HE.KeyAndHivePhysical & ": " & "[" & DC.ValueName & "] = " & Reg.StatusCodeDescOnFail(lData)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O27"
                        .HitLineW = sHit
                        AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, DC.ValueName, DC.DataLong, HE.Redirected, REG_RESTORE_DWORD
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Loop
    Loop
    
    If GetServiceRunState("MpsSvc") = SERVICE_RUNNING Then
        CheckO27Item_Firewall
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO27Item_RDP"
    If inIDE Then Stop: Resume Next
End Sub
    
Public Sub CheckO27Item_Firewall()
    On Error GoTo ErrorHandler:
    
    'O27 - RDP: (Port)
    Dim sHit$, result As SCAN_RESULT
    Dim pFwNetFwPolicy2 As New NetFwPolicy2
    Dim pFwRules As INetFwRules
    Dim pFwRule As NetFwRule
    Dim sProtocol As String
    Dim sService As String
    Dim sApp As String
    Dim Stage As Long
    
    Stage = 1
    Set pFwRules = pFwNetFwPolicy2.Rules
    
    Stage = 2
    For Each pFwRule In pFwRules
        Stage = 3
        If pFwRule.Enabled Then
            Stage = 4
            If pFwRule.Action = NET_FW_ACTION_ALLOW Then
                Stage = 5
                If pFwRule.direction = NET_FW_RULE_DIR_IN Then
                    Stage = 6
                    'Win11 has RdpSa.exe
                    'Also, sometimes 3389 opened system-wide
                    'If StrComp(GetFileName(pFwRule.ApplicationName, True), "svchost.exe", 1) = 0 Then
                        'If FW_IsPortInRange(3389, pFwRule.LocalPorts) Then
                        If StrComp("3389", pFwRule.LocalPorts) = 0 Then
                            Stage = 7
                            sService = pFwRule.serviceName
                            sApp = pFwRule.ApplicationName
                            
                            If Len(sService) = 0 Then sService = "(no service)"
                            If Len(sApp) = 0 Then sApp = "(all applications)"
                            
                            sHit = "O27 - RDP: (Port) 3389 " & FW_GetProtocolName(pFwRule.Protocol) & " opened as inbound - " & _
                                sService & " - (" & pFwRule.Name & ") - " & sApp
                            
                            If Not IsOnIgnoreList(sHit) Then
                                With result
                                    .Section = "O27"
                                    .HitLineW = sHit
                                    AddCustomToFix .Custom, CUSTOM_ACTION_FIREWALL_RULE, pFwRule.Name
                                    .CureType = CUSTOM_BASED
                                End With
                                AddToScanResults result
                            End If
                        End If
                    'End If
                End If
            End If
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO27Item_Firewall", "Stage: " & Stage
    If inIDE Then Stop: Resume Next
End Sub

Private Function FW_GetProtocolName(iProtocol As Long) As String
    Select Case iProtocol
    Case 6
        FW_GetProtocolName = "TCP"
    Case 17
        FW_GetProtocolName = "UDP"
    Case Else
        FW_GetProtocolName = "Other"
    End Select
End Function

'strPorts example: "3306,3310,3315-3320", "*", "3306"
Private Function FW_IsPortInRange(port As Long, strPorts As String) As Boolean
    On Error GoTo ErrorHandler:
    '// TODO: support Port name aliases, e.g. RPC-EPMap
    If Len(strPorts) = 0 Then Exit Function
    If strPorts = "*" Then FW_IsPortInRange = True: Exit Function
    Dim vRange, pos As Long, minValue As Long, maxValue As Long, tmp1 As String, tmp2 As String
    For Each vRange In Split(strPorts, ",")
        pos = InStr(1, vRange, "-")
        If pos = 0 Then
            If IsNumeric(vRange) Then
                If CLng(vRange) = port Then FW_IsPortInRange = True: Exit Function
            End If
        Else
            tmp1 = Left$(vRange, pos - 1)
            tmp2 = mid$(vRange, pos + 1)
            If IsNumeric(tmp1) And IsNumeric(tmp1) Then
                minValue = CLng(tmp1)
                maxValue = CLng(tmp2)
                If port >= minValue And port <= maxValue Then FW_IsPortInRange = True: Exit Function
            End If
        End If
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FW_IsPortInRange", strPorts
    If inIDE Then Stop: Resume Next
End Function

Public Function FW_RuleSetState(sRuleName As String, bEnabled As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim pFwNetFwPolicy2 As New NetFwPolicy2
    Dim pFwRules As INetFwRules
    Dim pFwRule As NetFwRule
    
    Set pFwRules = pFwNetFwPolicy2.Rules
    
    For Each pFwRule In pFwRules
        If StrComp(pFwRule.Name, sRuleName, vbTextCompare) = 0 Then
            pFwRule.Enabled = bEnabled
            FW_RuleSetState = True
        End If
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FW_RuleSetState"
End Function

Public Sub FixO27Item(sItem$, result As SCAN_RESULT)
    FixIt result
End Sub
