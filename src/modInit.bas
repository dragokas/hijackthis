Attribute VB_Name = "modInit"
Option Explicit
'
' (main entry point)
'

'by default = false
#Const bDebugMode = False       ' /debug key analogue
#Const bDebugToFile = False     ' /DebugToFile key analogue
#Const SilentAutoLog = False    ' /silentautolog key analogue
#Const DoCrash = False          ' crash the program (test reason)
#Const DoFreeze = False         ' emulate app freeze at initialization stage
#Const CryptDisable = False     ' disable encryption of ignore list and several other settings
#Const NoSelfSignTest = False   ' whether I should disable the checking of own signature
#Const AutologgerMode = False    ' keys combination used within AutoLogger (which is: /accepteula /silentautolog /default /skipIgnoreList /timeout:120 /debugtofile)

'by default = true
#Const AUTOLOGGER_DEBUG_TO_FILE = True 'auto-activate /DebugToFile if app ran within Autologger

Public Script As clsScript

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long

Public Sub Main()
    On Error GoTo ErrorHandler:
    
    If Not App.TaskVisible Then Exit Sub ' Additional threads are re-entrant, meaning they will always execute the "Sub Main" so just exit the sub
    
    PreInit
    
    If Not bAcceptEula Then
        frmEULA.Show vbModal
    End If
    
    If Not bAcceptEula Then Exit Sub
    
    PostInit
    
    Set Script = New clsScript
    
    If bPolymorph And Not bAutoLog Then
        '// TODO
        'If Script.HasFixInClipboard() Then
        '    Script.ExecuteFixFromClipboard false
        '    Exit Sub
        'end if
    End If
    
    If FileExists(BuildPath(AppPath, "apps\VBCCR17.OCX")) Then
        If FixTempFolder() Then
            frmMain.Show vbModeless
        Else
            MsgBox "Cannot run HijackThis" & vbNewLine & _
            "Please, check that you have write access to the temp folder: " & Environ("Temp")
        End If
    Else
        MsgBox "Cannot run HijackThis" & vbNewLine & _
            "Required file is missing: apps\VBCCR17.OCX" & vbNewLine & _
            "Ensure you unpacked archive completely!"
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Main"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub PreInit()
    On Error GoTo ErrorHandler:
    
    Debug.Assert MakeTrue(inIDE)
    
    MAX_PATH_W_BUF = String$(MAX_PATH_W, 0&)
    
    SetCurrentDirectory StrPtr(AppPath())
    
    EnableVisualStyles

    bPolymorph = (InStr(1, AppExeName(), "_poly", 1) <> 0) Or (StrComp(GetExtensionName(AppExeName(True)), ".pif", 1) = 0)
    
    If Not bPolymorph Then
        If HasExport("SetCurrentProcessExplicitAppUserModelID", "shell32.dll") Then ' Windows 7 and Later
            Call SetCurrentProcessExplicitAppUserModelID(StrPtr("Alex.Dragokas.HiJackThis"))
        End If
    End If
    
    If IsWow64() Then bIsWin64 = True
    bIsWOW64 = bIsWin64 ' mean VB6 app-s are always x32 bit.
    bIsWin32 = Not bIsWin64

    #If AutologgerMode Then
        bAcceptEula = True
    #End If
    
    g_SettingsRegKey = "Software\HijackThis+"

    Dim sExeName As String
    Dim argc As Long
    Dim aCmdArgs() As String
    
    g_sCommandLine = Command$()
    ParseCommandLine g_sCommandLine, argc, aCmdArgs(), False
    
    sExeName = AppExeName()
    
    bForceRU = InStr(1, sExeName, "_RU", 1) Or HasSwitch("langRU", aCmdArgs)  '/langRU
    bForceEN = InStr(1, sExeName, "_EN", 1) Or HasSwitch("langEN", aCmdArgs)  '/langEN
    bForceUA = InStr(1, sExeName, "_UA", 1) Or HasSwitch("langUA", aCmdArgs)  '/langUA
    bForceFR = InStr(1, sExeName, "_FR", 1) Or HasSwitch("langFR", aCmdArgs)  '/langUA
    bForceSP = InStr(1, sExeName, "_SP", 1) Or HasSwitch("langSP", aCmdArgs)  '/langSP
    
    Dim hasEulaKey As Boolean
    
    If Reg_KeyExists(HKEY_LOCAL_MACHINE, g_SettingsRegKey) Then
        hasEulaKey = True
        bAcceptEula = True
    Else
        '/accepteula /uninstall
        If HasSwitch("accepteula", aCmdArgs) Or _
            HasSwitch("uninstall", aCmdArgs) Then

            bAcceptEula = True
        End If
    End If
    
    If bAcceptEula Then
        If Not hasEulaKey Then
            EULA_Accept
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "PreInit"
    If inIDE Then Stop: Resume Next
End Sub

Private Function FolderExistsNAC(sFolder As String) As Boolean
    Dim status As Long
    status = GetFileAttributes(StrPtr(sFolder))
    If (CBool(status And vbDirectory) And (status <> INVALID_FILE_ATTRIBUTES)) Then
        FolderExistsNAC = True
    End If
End Function

Public Function FixTempFolder() As Boolean
    
    'Fix against Error 481 - Invalid picture
    
    Dim sTempF      As String
    Dim bUnlockDone As Boolean
    Dim bSuccess    As Boolean
    
    sTempF = String$(MAX_PATH, 0)
    
    If GetTempPath(Len(sTempF), StrPtr(sTempF)) = 0 Then
        sTempF = Environ("TMP")
    End If
    
    If Len(sTempF) = 0 Then
        sTempF = Environ("USERPROFILE") & "\AppData\Local\Temp"
        SetEnvironmentVariable StrPtr("TMP"), StrPtr(sTempF)
    End If
    
    If Not FolderExistsNAC(sTempF) Then
        If CreateDirectory(StrPtr(sTempF), ByVal 0&) = 0 Then
            If Err.LastDllError = ERROR_PATH_NOT_FOUND Then
                'create folder hierarchy
                MkDirW sTempF
                'refresh error code
                Call CreateDirectory(StrPtr(sTempF), ByVal 0&)
            End If
            If Err.LastDllError = ERROR_ACCESS_DENIED Then
                If TryUnlock(sTempF, False) Then
                    MkDirW sTempF
                End If
                bUnlockDone = True
            End If
        End If
    End If
    
    If FolderExistsNAC(sTempF) Then
        'doing AccessCheck() without direct I/O
        If Not CheckFileAccess(sTempF, GENERIC_READ Or GENERIC_WRITE) Then
            If Not bUnlockDone Then
                Call TryUnlock(sTempF, False)
            End If
        End If
        
        'need to do CheckFileAccess once more, but we'll not going to:
        'because, CheckFileAccess may returned incorrect result in case some functions were intercepted
        bSuccess = True
    End If
    
    FixTempFolder = bSuccess
    
End Function

Public Sub EULA_Accept()
    Reg_CreateKey HKEY_LOCAL_MACHINE, g_SettingsRegKey
End Sub

Private Sub PostInit()
    On Error GoTo ErrorHandler:
    
    Call AcquirePrivileges
    
    Set OSver = New clsOSInfo
    Set Reg = New clsRegistry
    
    InitVariables
    
    MigrateSettings
    
    InitBackupIni
    
    ReInitScanResults
    
    InitVerifyDigiSign
    
    ProcessCommandLine
    
    FixPermissions
    
    If inIDE Then
        AppVerString = GetVersionFromVBP(BuildPath(AppPath(), App.ExeName & ".vbp"))
    Else
        AppVerString = GetFilePropVersion(AppPath(True))
    End If
    
    AppendErrorLogCustom "Logfile ( tracing ) of HijackThis+ v." & AppVerString & vbCrLf & vbCrLf & _
        "Command line: " & AppPath(True) & " " & g_sCommandLine & vbCrLf & vbCrLf & MakeLogHeader() & vbCrLf
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "PostInit"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub FixPermissions()
    On Error GoTo ErrorHandler:
    Dim Blocked As Boolean
    
    If Not CheckKeyAccess(HKLM, g_SettingsRegKey, KEY_READ Or IIf(OSver.IsElevated, KEY_WRITE, 0)) Then
        Blocked = True
    ElseIf Reg.HasDenyACL(HKLM, g_SettingsRegKey) Then
        Blocked = True
    End If
    
    If Blocked Then
        Call modPermissions.RegKeyResetDACL(HKLM, g_SettingsRegKey, False, True)
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixPermissions"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub ProcessCommandLine()
    On Error GoTo ErrorHandler:
    
    Dim argc As Long
    If Not inIDE Then
        g_sCommandLine = StringFromPtrW(GetCommandLine())
    End If
    ParseCommandLine g_sCommandLine, argc, g_sCommandLineArg(), True
    
    Dim pos         As Long
    Dim sTime       As String
    Dim sCmdLine    As String
    Dim sCustomPath As String
    Dim ExeName     As String
    Dim bLogBusy    As Boolean
    
    If HasCommandLineKey("StartupScan") Then '/StartupScan
        bStartupScan = True
    End If
    
    If bStartupScan Then
        Call SetPriorityProcess(OSver.CurrentProcessId, BELOW_NORMAL_PRIORITY_CLASS)
        Call SetProcessIOPriority(OSver.CurrentProcessId, IO_PRIORITY_LOW)
    Else
        Call SetPriorityProcess(OSver.CurrentProcessId, HIGH_PRIORITY_CLASS)
        Call SetProcessIOPriority(OSver.CurrentProcessId, IO_PRIORITY_HIGH)
        Call SetProcessPagePriority(OSver.CurrentProcessId, MEMORY_PRIORITY_NORMAL)
    End If
    
    Perf.MAX_TimeOut = MAX_TIMEOUT_DEFAULT
    '/timeout
    pos = InStr(1, g_sCommandLine, "timeout", 1)
    If pos <> 0 Then
        sTime = mid$(g_sCommandLine, pos + Len("timeout") + 1)
        pos = InStr(sTime, " ")
        If pos <> 0 Then
            sTime = Left$(sTime, pos - 1)
        End If
        If IsNumeric(sTime) Then
            Perf.MAX_TimeOut = CLng(sTime)
            If Perf.MAX_TimeOut < 0 Then Perf.MAX_TimeOut = MAX_TIMEOUT_DEFAULT
        End If
    End If
    
    #If bDebugMode Then
        bDebugMode = True
    #End If
    #If bDebugToFile Then
        bDebugToFile = True
    #End If
    #If SilentAutoLog Then
        bAutoLog = True: bAutoLogSilent = True
    #End If
    #If CryptDisable Then
        bCryptDisable = True
    #End If
    
    '/autolog
    If HasCommandLineKey("autolog") Then bAutoLog = True
    '/silentautolog
    If HasCommandLineKey("silentautolog") Then bAutoLog = True: bAutoLogSilent = True
    
    Set ErrLogCustomText = New clsStringBuilder 'tracing
    
    '/debug
    If HasCommandLineKey("debug") Or _
        InStr(1, AppExeName(), "_debug", 1) <> 0 Or _
        InStr(1, AppExeName(), "_dbg", 1) <> 0 Then
        bDebugMode = True
    End If
    
    If HasCommandLineKey("days:") <> 0 Then '/days:
        ABR_RunBackup
        End 'do crash ^_^
    End If
    
    '/nogui /StartupScan /install
    If HasCommandLineKey("nogui") _
      Or HasCommandLineKey("StartupScan") _
      Or HasCommandLineKey("install") Then
        gNoGUI = True
    End If
    
    sCmdLine = Replace$(g_sCommandLine, ":", "+")
    
    '/Tool:xxx
    If Len(sCmdLine) <> 0 Then
        If InStr(1, sCmdLine, "tool+StartupList", 1) <> 0 Then bRunToolStartupList = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+UninstMan", 1) <> 0 Then bRunToolUninstMan = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+DigiSign", 1) <> 0 Then bRunToolEDS = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+RegUnlocker", 1) <> 0 Then bRunToolRegUnlocker = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+RegTypeChecker", 1) <> 0 Then bRunToolRegTypeChecker = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+ADSSpy", 1) <> 0 Then bRunToolADSSpy = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+Hosts", 1) <> 0 Then bRunToolHosts = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+ProcMan", 1) <> 0 Then bRunToolProcMan = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+CheckLNK", 1) <> 0 Then bRunToolCBL = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+ClearLNK", 1) <> 0 Then bRunToolClearLNK = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+Auto" & "runs", 1) <> 0 Then bRunToolAutoruns = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+Executed", 1) <> 0 Then bRunToolExecuted = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+LastActivity", 1) <> 0 Then bRunToolLastActivity = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+ServiWin", 1) <> 0 Then bRunToolServiWin = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+TaskScheduler", 1) <> 0 Then bRunToolTaskScheduler = True: gNoGUI = True
    End If
    
    ExeName = GetFileName(AppPath(True))
    
    If StrBeginWith(ExeName, "StartupList") Then bRunToolStartupList = True: gNoGUI = True
    If StrBeginWith(ExeName, "UninstMan") Then bRunToolUninstMan = True: gNoGUI = True
    If StrBeginWith(ExeName, "DigiSignChecker") Then bRunToolEDS = True: gNoGUI = True
    If StrBeginWith(ExeName, "RegUnlocker") Then bRunToolRegUnlocker = True: gNoGUI = True
    If StrBeginWith(ExeName, "RegTypeChecker") Then bRunToolRegTypeChecker = True: gNoGUI = True
    If StrBeginWith(ExeName, "ADSSpy") Then bRunToolADSSpy = True: gNoGUI = True
    If StrBeginWith(ExeName, "HostsMan") Then bRunToolHosts = True: gNoGUI = True
    If StrBeginWith(ExeName, "ProcMan") Then bRunToolProcMan = True: gNoGUI = True
    
    '/saveLog "c:\LogPath"
    '/saveLog "c:\LogPath\LogName.log"
    If InStr(1, Command$, "saveLog", vbTextCompare) > 0 Then
        'path to save logfile to
        sCustomPath = mid$(Command$, InStr(1, Command$, "saveLog", 1) + 8)
        If Left$(sCustomPath, 1) = """" Then
            'path enclosed in quotes, get what's between
            sCustomPath = mid$(sCustomPath, 2)
            If InStr(sCustomPath, """") > 0 Then
                sCustomPath = Left$(sCustomPath, InStr(sCustomPath, """") - 1)
            Else
                'no closing quote
                sCustomPath = vbNullString
            End If
        Else
            'path has no quotes, stop at first space
            If InStr(sCustomPath, " ") > 0 Then
                sCustomPath = Left$(sCustomPath, InStr(sCustomPath, " ") - 1)
            End If
        End If
    End If
    If Len(sCustomPath) <> 0 Then
        If Not FolderExists(sCustomPath, , True) Then
            If Not MkDirW(sCustomPath, StrEndWith(sCustomPath, ".log")) Then
                sCustomPath = vbNullString
            End If
        End If
    End If
    If StrEndWith(sCustomPath, ".log") Then
        g_sLogFile = sCustomPath
        g_sDebugLogFile = BuildPath(GetParentDir(sCustomPath), "HiJackThis_debug.log")
        
        If FileExists(g_sLogFile, , True) Then
            If CheckFileAccessWrite_Physically(g_sLogFile, False) Then
                DeleteFileForce g_sLogFile, , True
            Else
                bLogBusy = True
            End If
        Else
            If Not CheckFileAccess(GetParentDir(g_sLogFile), GENERIC_WRITE) Then
                bLogBusy = True
            End If
        End If
    End If
    If Len(sCustomPath) = 0 Or bLogBusy Then
        g_sLogFile = BuildPath(AppPath(), "HiJackThis_.log")
        g_sDebugLogFile = BuildPath(AppPath(), "HiJackThis_debug.log")
    End If
    
    If bAutoLog Then
        OpenLogHandle
    End If
    
    '/DebugToFile
    If HasCommandLineKey("DebugToFile") Then
        bDebugToFile = True
    End If
    
    #If AUTOLOGGER_DEBUG_TO_FILE Then
        If InStr(1, AppPath(), "\AutoLogger\", 1) <> 0 Then
            bDebugToFile = True
        End If
    #End If
    
    #If AutologgerMode Then
        bAutoLogSilent = True
        bAutoLog = True
        bLoadDefaults = True
        bSkipIgnoreList = True
        bDebugToFile = True
        Perf.MAX_TimeOut = 120
        bAcceptEula = True
    #End If
    
    If bDebugMode Or bDebugToFile Then
        bDebugToFile = True ' /debug also initiate /bDebugToFile
        OpenDebugLogHandle
    End If
    
    If bAutoLogSilent Then
        DisableSubclassing = True
    End If
    
    AppendErrorLogCustom "Command line: " & g_sCommandLine
    
    #If DoFreeze Then
        Do: DoEvents: Loop
    #End If
    
    #If DoCrash Then
        DoCrash
    #End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ProcessCommandLine"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnableVisualStyles()
    On Error GoTo ErrorHandler:

    Dim ICC         As tagINITCOMMONCONTROLSEX
    Dim lr          As Long
    Dim hModShell   As Long

    If Not inIDE Then
        hModShell = LoadLibrary(StrPtr("shell32.dll"))
    End If
    
    If hModShell <> 0 And Not inIDE Then
    
        With ICC
            .dwSize = Len(ICC)
            .dwICC = ICC_STANDARD_CLASSES 'http://www.geoffchappell.com/studies/windows/shell/comctl32/api/commctrl/initcommoncontrolsex.htm
        End With

        lr = InitCommonControlsEx(ICC)
        
        If lr = 0 Or Err.Number <> 0 Then
            InitCommonControls ' 9x version
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnableVisualStyles"
    If inIDE Then Stop: Resume Next
End Sub

Private Function IsWow64() As Boolean
    Dim hModule As Long, procAddr As Long, lIsWin64 As Long
    Static isInit As Boolean, result As Boolean
    
    If isInit Then
        IsWow64 = result
    Else
        isInit = True
        hModule = LoadLibrary(StrPtr("kernel32.dll"))
        If hModule Then
            procAddr = GetProcAddress(hModule, "IsWow64Process")
            If procAddr <> 0 Then
                IsWow64Process GetCurrentProcess(), lIsWin64
                result = CBool(lIsWin64)
                IsWow64 = result
            End If
            FreeLibrary hModule
        End If
    End If
End Function

Private Function Reg_CreateKey(hHive As ENUM_REG_HIVE, ByVal sKey$) As Boolean
    Dim hKey&, lret&
    lret = RegCreateKeyEx(hHive, StrPtr(sKey), 0&, ByVal 0&, 0&, KEY_CREATE_SUB_KEY Or (bIsWOW64 And KEY_WOW64_64KEY), ByVal 0&, hKey, ByVal 0&)
    Reg_CreateKey = (ERROR_SUCCESS = lret)
    If hKey <> 0 Then RegCloseKey hKey
End Function

Private Function Reg_KeyExists(hHive As ENUM_REG_HIVE, ByVal sKey$) As Boolean
    Dim hKey&, lStatus&
    lStatus = RegOpenKeyEx(hHive, StrPtr(sKey), 0&, WRITE_OWNER Or (bIsWOW64 And KEY_WOW64_64KEY), hKey)
    If lStatus = ERROR_SUCCESS Then
        Reg_KeyExists = True
        RegCloseKey hKey
    ElseIf lStatus = ERROR_ACCESS_DENIED Then
        'for 'Limited User'
        lStatus = RegOpenKeyEx(hHive, StrPtr(sKey), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY), hKey)
        If ERROR_SUCCESS = lStatus Then
            Reg_KeyExists = True
            RegCloseKey hKey
        End If
    End If
End Function

Private Function HasSwitch(ByVal sKey As String, aCmdArgs() As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim offset As Long
    Dim bHasKey As Boolean
    If UBound(aCmdArgs) > 0 Then
        For i = 1 To UBound(aCmdArgs)
            bHasKey = False
            If StrBeginWith(aCmdArgs(i), "/" & sKey) Then
                bHasKey = True
            ElseIf StrBeginWith(aCmdArgs(i), "-" & sKey) Then
                bHasKey = True
            End If
            If bHasKey Then
                If Right$(sKey, 1) = ":" Then offset = -1
                ch = mid$(aCmdArgs(i), Len(sKey) + 2 + offset, 1)
                If (Len(ch) = 0 Or ch = ":") Then
                    HasSwitch = True
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Function MakeTrue(Value As Boolean) As Boolean
    Value = True
    MakeTrue = True
End Function

Private Function HasExport(ByVal ProcedureName As String, ByVal DllFilename As String) As Boolean
    Dim hModule As Long, procAddr As Long
    hModule = LoadLibrary(StrPtr(DllFilename))
    If hModule Then
        procAddr = GetProcAddress(hModule, StrPtr(StrConv(ProcedureName, vbFromUnicode)))
        FreeLibrary hModule
    End If
    HasExport = (procAddr <> 0)
End Function

Public Sub MigrateSettings()
    Dim oldKey As String
    'Software\TrendMicro\HiJackThisFork
    oldKey = Caes_Decode("TrkAFlEtmgMBMEjNJ[ZIqZwVZdOehtItyt")
    If Reg.KeyExists(HKLM, oldKey) Then
        Reg.MoveKey HKLM, oldKey, False, _
            HKLM, g_SettingsRegKey, False, False
        'Software\TrendMicro\HiJackThis
        Reg.DelKey HKLM, Caes_Decode("TrkAFlEtmgMBMEjNJ[ZIqZwVZdOeht"), False
    End If
End Sub
