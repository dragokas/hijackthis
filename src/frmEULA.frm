VERSION 5.00
Begin VB.Form frmEULA 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7200
   ClientLeft      =   4788
   ClientTop       =   4836
   ClientWidth     =   6744
   Icon            =   "frmEULA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6744
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEULA 
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1920
      Width           =   6615
   End
   Begin VB.CommandButton cmdNotAgree 
      Caption         =   "I Do Not Accept"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAgree 
      Caption         =   "I Accept"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblAware 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEULA.frx":4B2A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6450
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to HiJackThis Fork 3"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1185
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmEULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmEULA.frm]

'
' License agreement form
'
' (main entry point)
'

Option Explicit

#Const bDebugMode = False       ' /debug key analogue
#Const bDebugToFile = False     ' /bDebugToFile key analogue
#Const SilentAutoLog = False    ' /silentautolog key analogue
#Const DoCrash = False          ' crash the program (test reason)
#Const CryptDisable = False     ' disable encryption of ignore list and several other settings
#Const NoSelfSignTest = False   ' whether I should disable the checking of own signature

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Const ICC_STANDARD_CLASSES As Long = &H4000&

Private ControlsEvent As clsEvents

Private Sub Form_Initialize()
    On Error GoTo ErrorHandler:
    
    MAX_PATH_W_BUF = String$(MAX_PATH_W, 0&)
    
    SetCurrentDirectory StrPtr(AppPath())
    
    bPolymorph = (InStr(1, AppExeName(), "_poly", 1) <> 0) Or (StrComp(GetExtensionName(AppExeName(True)), ".pif", 1) = 0)
    
    Dim argc As Long
    g_sCommandLine = Command$()
    ParseCommandLine g_sCommandLine, argc, g_sCommandLineArg()

    Dim ICC         As tagINITCOMMONCONTROLSEX
    Dim lr          As Long
    Dim hModShell   As Long
    Dim pos         As Long
    Dim sTime       As String
    Dim sCMDLine    As String
    Dim sPath       As String
    Dim ExeName     As String
    
    ' Code launched from IDE ?
    Debug.Assert CheckIDE(inIDE)
    
    If HasCommandLineKey("release") Then Exit Sub '/release
    
    If IsWow64() Then bIsWin64 = True
    bIsWOW64 = bIsWin64 ' mean VB6 app-s are always x32 bit.
    bIsWin32 = Not bIsWin64
    
    Call AcquirePrivileges
    
    Set OSver = New clsOSInfo
    
    ' boost priority
    If HasCommandLineKey("StartupScan") Then '/StartupScan
        bStartupScan = True
        Call SetPriorityProcess(OSver.CurrentProcessId, BELOW_NORMAL_PRIORITY_CLASS)
        Call SetProcessIOPriority(OSver.CurrentProcessId, IO_PRIORITY_LOW)
    Else
        Call SetPriorityProcess(OSver.CurrentProcessId, HIGH_PRIORITY_CLASS)
        Call SetProcessIOPriority(OSver.CurrentProcessId, IO_PRIORITY_HIGH)
        Call SetProcessPagePriority(OSver.CurrentProcessId, MEMORY_PRIORITY_NORMAL)
    End If
    
    ' Enable visual styles
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
    
    Perf.MAX_TimeOut = MAX_TIMEOUT_DEFAULT
    '/timeout
    pos = InStr(1, g_sCommandLine, "timeout", 1)
    If pos <> 0 Then
        sTime = Mid$(g_sCommandLine, pos + Len("timeout") + 1)
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
    
    If bAutoLog Then Perf.StartTime = GetTickCount()
    
    Set Reg = New clsRegistry
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
    
    If OSver.MajorMinor >= 6.1 And Not bPolymorph Then ' Windows 7 and Later
        lr = SetCurrentProcessExplicitAppUserModelID(StrPtr("Alex.Dragokas.HiJackThis"))
    End If
    
    Set oDictFileExist = New clsTrickHashTable  'file exists cache
    oDictFileExist.CompareMode = 1
    
    '/nogui /StartupScan /install
    If HasCommandLineKey("nogui") _
      Or HasCommandLineKey("StartupScan") _
      Or HasCommandLineKey("install") Then
        gNoGUI = True
    End If
    
    sCMDLine = Replace$(g_sCommandLine, ":", "+")
    
    '/Tool:xxx
    If Len(sCMDLine) <> 0 Then
        If InStr(1, sCMDLine, "tool+StartupList", 1) <> 0 Then bRunToolStartupList = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+UninstMan", 1) <> 0 Then bRunToolUninstMan = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+DigiSign", 1) <> 0 Then bRunToolEDS = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+RegUnlocker", 1) <> 0 Then bRunToolRegUnlocker = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+RegTypeChecker", 1) <> 0 Then bRunToolRegTypeChecker = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+ADSSpy", 1) <> 0 Then bRunToolADSSpy = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+Hosts", 1) <> 0 Then bRunToolHosts = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+ProcMan", 1) <> 0 Then bRunToolProcMan = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+CheckLNK", 1) <> 0 Then bRunToolCBL = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+ClearLNK", 1) <> 0 Then bRunToolClearLNK = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+Autoruns", 1) <> 0 Then bRunToolAutoruns = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+Executed", 1) <> 0 Then bRunToolExecuted = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+LastActivity", 1) <> 0 Then bRunToolLastActivity = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+ServiWin", 1) <> 0 Then bRunToolServiWin = True: gNoGUI = True
        If InStr(1, sCMDLine, "tool+TaskScheduler", 1) <> 0 Then bRunToolTaskScheduler = True: gNoGUI = True
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
        sPath = Mid$(Command$, InStr(1, Command$, "saveLog", 1) + 8)
        If Left$(sPath, 1) = """" Then
            'path enclosed in quotes, get what's between
            sPath = Mid$(sPath, 2)
            If InStr(sPath, """") > 0 Then
                sPath = Left$(sPath, InStr(sPath, """") - 1)
            Else
                'no closing quote
                sPath = vbNullString
            End If
        Else
            'path has no quotes, stop at first space
            If InStr(sPath, " ") > 0 Then
                sPath = Left$(sPath, InStr(sPath, " ") - 1)
            End If
        End If
    End If
    
    If Len(sPath) <> 0 Then
        If Not FolderExists(sPath, , True) Then
            If Not MkDirW(sPath, StrEndWith(sPath, ".log")) Then
                sPath = vbNullString
            End If
        End If
    End If
    If Len(sPath) <> 0 Then
        If StrEndWith(sPath, ".log") Then
            g_sLogFile = sPath
            g_sDebugLogFile = BuildPath(GetParentDir(sPath), "HiJackThis_debug.log")
        Else
            g_sLogFile = BuildPath(sPath, "HiJackThis.log")
            g_sDebugLogFile = BuildPath(sPath, "HiJackThis_debug.log")
        End If
        'check write access
        If Not CheckAccessWrite(g_sLogFile, True) Then sPath = vbNullString
    End If
    If Len(sPath) = 0 Then
        g_sLogFile = BuildPath(AppPath(), "HiJackThis.log")
        g_sDebugLogFile = BuildPath(AppPath(), "HiJackThis_debug.log")
    End If
    
    If bAutoLog Then
        OpenLogHandle
    End If
    
    '/DebugToFile
    If HasCommandLineKey("DebugToFile") Then
        bDebugToFile = True
    End If
    If bDebugMode Then
        OpenDebugLogHandle
    End If

    Exit Sub
ErrorHandler:
    If InStr(1, g_sCommandLine, "silentautolog", 1) = 0 Then
        MsgBoxW "Error in frmEULA.Form_Initialize. Err. number = " & Err.Number & " - " & Err.Description & ". LastDllErr = " & Err.LastDllError
    End If
    If inIDE Then Stop
    Resume Next
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "frmEULA.Form_Load - Begin"

    If HasCommandLineKey("release") Then '/release
        'Если мы - клон, ожидать завершения PID, переданного через аргументы командной строки, после чего удалить файлы
        WatchForProcess
        Unload Me
        Exit Sub
    End If

    If inIDE Then
        AppVerString = GetVersionFromVBP(BuildPath(AppPath(), App.ExeName & ".vbp")) '_HijackThis.vbp"
    Else
        AppVerString = GetFilePropVersion(AppPath(True))
    End If
    
    'header of tracing log
    AppendErrorLogCustom vbCrLf & vbCrLf & "Logfile ( tracing ) of HiJackThis+ v." & AppVerString & vbCrLf & vbCrLf & _
        "Command line: " & AppPath(True) & " " & g_sCommandLine & vbCrLf & vbCrLf & MakeLogHeader()
    
    'Me.Caption = Me.Caption & AppVerString

    bForceRU = InStr(1, AppExeName(), "_RU", 1) Or HasCommandLineKey("langRU")  '/langRU
    bForceEN = InStr(1, AppExeName(), "_EN", 1) Or HasCommandLineKey("langEN")  '/langEN
    bForceUA = InStr(1, AppExeName(), "_UA", 1) Or HasCommandLineKey("langUA")  '/langUA
    bForceFR = InStr(1, AppExeName(), "_FR", 1) Or HasCommandLineKey("langFR")  '/langUA
    bForceSP = InStr(1, AppExeName(), "_SP", 1) Or HasCommandLineKey("langSP")  '/langSP

    bIsWOW64 = IsWow64()

    #If DoCrash Then
        DoCrash
    #End If

    '/accepteula /uninstall
    If HasCommandLineKey("accepteula") Or _
        HasCommandLineKey("uninstall") Or _
        Reg.KeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThisFork") Then
            EULA_Agree
            Me.Hide
            'frmMain.Show vbModeless
            Load frmMain
            Unload Me
    Else
        Localize
        txtEULA.Text = GetEULA()
        Set ControlsEvent = New clsEvents
        Set ControlsEvent.txtBoxInArr = txtEULA   'focus on txtbox to add scrolling support
        bFirstRun = True
    End If
    
    AppendErrorLogCustom "frmEULA.Form_Load - End"
    Exit Sub
ErrorHandler:
    MsgBoxW "Error in frmEULA.Form_Load. Err. number = " & Err.Number & " - " & Err.Description & ". LastDllErr = " & Err.LastDllError
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdAgree_Click()
    EULA_Agree
    Me.Hide
    Set ControlsEvent = Nothing
    'frmMain.Show
    Load frmMain
    Unload Me
End Sub

Private Sub cmdNotAgree_Click()
    Set ControlsEvent = Nothing
    Unload Me
End Sub

Function CheckIDE(Value As Boolean) As Boolean: Value = True: CheckIDE = True: End Function

Sub WatchForProcess()   'waiting for process completion to release unpacked resources
    On Error GoTo ErrorHandler
    Const INFINITE                  As Long = -1
    Const SYNCHRONIZE               As Long = &H100000
    Const PROCESS_QUERY_INFORMATION As Long = 1024&
    
    Dim ProcessID As Long
    Dim hProc As Long
    Dim lret As Long
    Dim sPathComCtl1 As String
    Dim sPathComCtl2 As String
    
    sPathComCtl1 = BuildPath(AppPath(), "MSComCtl.ocx")
    sPathComCtl2 = BuildPath(AppPath(), "MSComCtl.oca")
    
    ProcessID = Val(Mid$(g_sCommandLine, InStr(1, g_sCommandLine, "/release:", 1) + Len("/release:")))
    
    If ProcessID <> 0 Then
        
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or SYNCHRONIZE, False, ProcessID)
    
        If hProc <> 0 Then
            Call WaitForSingleObject(hProc, INFINITE)
            CloseHandle hProc
        End If
        
        If Not FileExists(BuildPath(AppPath(), "_HijackThis.vbp")) Then
            lret = DeleteFileW(StrPtr(sPathComCtl1))
        End If
        
        lret = DeleteFileW(StrPtr(sPathComCtl2))
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "WatchForProcess", "PID:", ProcessID
    If inIDE Then Stop: Resume Next
End Sub

Sub EULA_Agree()
    Reg.CreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", False
    Reg.CreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", True
    Reg.CreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThisFork"
End Sub

Sub Localize()
    On Error GoTo ErrorHandler

    'pre-loading native OS UI language
    If bForceEN Then
        LoadLanguage &H409, True, PreLoadNativeLang:=True
    ElseIf bForceRU Then
        LoadLanguage &H419, True, PreLoadNativeLang:=True
    ElseIf bForceUA Then
        LoadLanguage &H422, True, PreLoadNativeLang:=True
    ElseIf bForceFR Then
        LoadLanguage &H40C, True, PreLoadNativeLang:=True
    ElseIf bForceSP Then
        LoadLanguage &H40A, True, PreLoadNativeLang:=True
    Else
        LoadLanguage 0, False, PreLoadNativeLang:=True
    End If
    
    ' Trend Micro HiJackThis - License Agreement
    Me.Caption = TranslateNative(1092)
    
    ' Welcome to HiJackThis 3
    lblWelcome.Caption = TranslateNative(1090)
    
    ' HiJackThis is free and open source program. It is provided "AS IS" without warranty of any kind. You may use this software at your own risk.
    ' This software is not permitted for commercial purposes.
    ' You need to read and accept license agreement to continue:
    lblAware.Caption = TranslateNative(1091)
    
    ' I Accept
    cmdAgree.Caption = TranslateNative(1093)
    
    ' I Do Not Accept
    cmdNotAgree.Caption = TranslateNative(1094)
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Localize"
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

Function GetEULA() As String
    Dim sText$
    sText = sText & "Program is licensed under GNU GENERAL PUBLIC LICENSE Version 2, June 1991 \n"
    sText = sText & "The source code is available at: https://github.com/dragokas/hijackthis"
    
    GetEULA = Replace$(sText, "\n", vbCrLf)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set ControlsEvent = Nothing
End Sub
