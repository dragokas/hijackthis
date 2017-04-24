Attribute VB_Name = "modMain_2"
Option Explicit

Const MAX_CODE_LENGTH As Long = 300&

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
            
                DoEvents
            
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

    Dim objService As Object, colBindings As Object, objBinding As Object
    Dim FilterName As String, ConsumerName As String, i As Long
    Dim FilterPath As String, ConsumerPath As String
    Dim FilterNameSpace As String, ConsumerNameSpace As String, ConsumerClassName As String
    Dim objServiceConsumer As Object, objConsumer As Object
    Dim objServiceFilter As Object, objFilter As Object, sFilterQuery As String
    Dim objTimerNamespace As Object, objTimer As Object, sTimerClassName As String, sTimerName As String
    Dim sHit As String, sScriptFile As String, sAdditionalInfo As String, sEventName As String, sScriptText As String
    Dim cmdExecute As String, cmdWorkDir As String, cmdArguments As String, sConsumerFileName As String, sConsumerText As String
    'Dim O25() As O25_Info: ReDim O25(0)
    Dim Result As TYPE_Scan_Results
    Dim Stady As Single, ComeBack As Boolean, NoConsumer As Boolean, NoFilter As Boolean, bOtherConsumerClass As Boolean
    
    If GetServiceRunState("winmgmt") <> SERVICE_RUNNING Then Exit Sub
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

        DoEvents
    
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
                sAdditionalInfo = ""
                
                If Not (objConsumer Is Nothing) Then
                    
                    'Checking several known classes on: root\subscription
                    'to provide a bit more information to log
                    
                    bOtherConsumerClass = False
                    
                    Stady = 7
                    If StrComp(ConsumerClassName, "ActiveScriptEventConsumer", 1) = 0 Then
                    
                        'Debug.Print objConsumer.ScriptingEngine    'language (engine)
                        'Debug.Print objConsumer.ScriptFileName     'external file
                        'Debug.Print objConsumer.ScriptText         'embedded script code
                        Stady = 8
                        If Not IsNull(objConsumer.ScriptFileName) Then sScriptFile = objConsumer.ScriptFileName
                        If Not IsNull(objConsumer.ScriptText) Then sScriptText = objConsumer.ScriptText
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
                            
                            If objTimer Is Nothing Then
                                sTimerClassName = "__AbsoluteTimerInstruction"
                                Set objTimer = objTimerNamespace.Get(sTimerClassName & ".TimerId=" & """" & sTimerName & """")
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
                If Not bOtherConsumerClass Then
                If bIgnoreAllWhitelists Or _
                    (Not (StrComp(ConsumerName, "BVTConsumer", 1) = 0 And cmdExecute = "" And cmdWorkDir = SysDisk & "\\tools\\kernrate" And cmdArguments = "cscript KernCap.vbs" And Not FileExists(SysDisk & "\tools\kernrate\KernCap.vbs")) _
                    And Not (FilterName = "BVTFilter") And sEventName = "__InstanceModificationEvent WITHIN 60 WHERE TargetInstance ISA ""Win32_Processor"" AND TargetInstance.LoadPercentage > 99") Then
                
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
                            If bMD5 And 0 <> Len(sScriptFile) Then sHit = sHit & GetFileMD5(sScriptFile)
                
                            With Result
                                .Section = "O25"
                                .HitLineW = sHit
                                With .O25
                                    .sScriptFile = sScriptFile
                                    .ConsumerName = ConsumerName
                                    .ConsumerNameSpace = ConsumerNameSpace
                                    .ConsumerPath = ConsumerPath
                                    .FilterName = FilterName
                                    .FilterNameSpace = FilterNameSpace
                                    .FilterPath = FilterPath
                                    .sTimerClassName = sTimerClassName
                                    .TimerID = sTimerName
                                End With
                            End With
                            AddToScanResults Result
                        End If
                End If
                End If
                'End If
                
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


'cure WMI infection
Public Sub FixO25Item(sItem$)
    
    On Error GoTo ErrorHandler
    
    Dim objService As Object, objInstance As Object, Finish As Boolean, i As Long
    Dim colBindings As Object, objBinding As Object, objBindingToDelete As Object
    Dim Result As TYPE_Scan_Results
    
    If Not GetScanResults(sItem, Result) Then Exit Sub
    
    With Result.O25
    
        On Error Resume Next
        'filter
        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & .FilterNameSpace)
        objService.Get(.FilterPath).Delete_
        
        'consumer
        Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\" & .ConsumerNameSpace)
        objService.Get(.ConsumerPath).Delete_
        
        'timer
        If 0 <> Len(.TimerID) Then
        
            Set objService = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security, Backup)}!\\.\root\subscription")
            objService.Get(.sTimerClassName & ".TimerId=" & """" & .TimerID & """").Delete_
        
        End If
        
        On Error GoTo ErrorHandler
        
        If FileExists(.sScriptFile) Then
            DeleteFileWEx StrPtr(.sScriptFile)
        End If
        
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
                
                    If objBinding.Filter = .FilterPath And objBinding.Consumer = .ConsumerPath Then
                
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
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain2_FixO25Item"
    If inIDE Then Stop: Resume Next
End Sub

' strip string to length defined
Function StripCode(ByVal sCode, Optional Max_Characters As Long = MAX_CODE_LENGTH, Optional AddActualLength As Boolean = True) As String
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
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO26Item - Begin"
    
    Dim sKeys$(), i&, sIFE$, sFile$, sHit$, Result As TYPE_Scan_Results
    
    sIFE = "Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
    'key is x64-shared
    'no HKCU
    
    sKeys = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sIFE), "|")
    For i = 0 To UBound(sKeys)
        sFile = RegGetString(HKEY_LOCAL_MACHINE, sIFE & "\" & sKeys(i), "Debugger")
        If sFile <> vbNullString Then
            sFile = FindOnPath(UnQuote(EnvironW(sFile)), True)
            
            sHit = "O26 - Image File Execution Options: " & sKeys(i) & " - " & sFile
            
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                
                If FileExists(sFile) Then
                    sFile = GetLongPath(sFile) '8.3 -> Full
                Else
                    sHit = sHit & " (file missing)"
                End If
                
                With Result
                    .Section = "O26"
                    .HitLineW = sHit
                    ReDim .RegKey(0)
                    .RegKey(0) = "HKLM\" & sIFE & "\" & sKeys(i)
                    .RegParam = "Debugger"
                End With
                AddToScanResults Result
            End If
            
        End If
    Next i
    
    AppendErrorLogCustom "CheckO26Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain2_CheckO26Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO26Item(sItem$)
    
    On Error GoTo ErrorHandler

    Dim Result As TYPE_Scan_Results
    If Not GetScanResults(sItem, Result) Then Exit Sub

    With Result
        RegDelVal 0, .RegKey(0), .RegParam
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain2_FixO26Item"
    If inIDE Then Stop: Resume Next
End Sub
