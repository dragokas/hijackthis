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

Private Const ICC_STANDARD_CLASSES As Long = &H4000&

Private ControlsEvent As clsEvents

Private Sub Form_Initialize()
    On Error GoTo ErrorHandler:
    
    Dim argc As Long
    g_sCommandLine = Command$()
    ParseCommandLine g_sCommandLine, argc, g_sCommandLineArg()

    Dim ICC         As tagINITCOMMONCONTROLSEX
    Dim lr          As Long
    Dim hModShell   As Long
    Dim pos         As Long
    Dim sTime       As String
    Dim sCmdLine    As String
    Dim sPath       As String
    Dim sFile       As String
    Dim ExeName     As String
    
    ' Code launched from IDE ?
    Debug.Assert CheckIDE(inIDE)
    
    If HasCommandLineKey("release") Then Exit Sub '/release
    
    Call AcquirePrivileges
    
    Set OSver = New clsOSInfo
    
    ' boost priority
    If HasCommandLineKey("StartupScan") Then '/StartupScan
        bStartupScan = True
        Call SetPriorityProcess(GetCurrentProcess(), BELOW_NORMAL_PRIORITY_CLASS)
        Call SetProcessIOPriority(GetCurrentProcessId(), 1) 'Low
    Else
        Call SetPriorityProcess(GetCurrentProcess(), HIGH_PRIORITY_CLASS)
        Call SetProcessIOPriority(GetCurrentProcessId(), 3) 'High
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
    
    If hModShell <> 0 Then
        FreeLibrary hModShell
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
    
    If OSver.MajorMinor >= 6.1 Then ' Windows 7 and Later
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
    
    sCmdLine = Replace$(g_sCommandLine, ":", "+")
    
    '/Tool:xxx
    If Len(sCmdLine) <> 0 Then
        If InStr(1, sCmdLine, "tool+StartupList", 1) <> 0 Then bRunToolStartupList = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+UninstMan", 1) <> 0 Then bRunToolUninstMan = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+DigiSign", 1) <> 0 Then bRunToolEDS = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+RegUnlocker", 1) <> 0 Then bRunToolRegUnlocker = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+ADSSpy", 1) <> 0 Then bRunToolADSSpy = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+Hosts", 1) <> 0 Then bRunToolHosts = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+ProcMan", 1) <> 0 Then bRunToolProcMan = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+CheckLNK", 1) <> 0 Then bRunToolCBL = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+ClearLNK", 1) <> 0 Then bRunToolClearLNK = True: gNoGUI = True
        If InStr(1, sCmdLine, "tool+Autoruns", 1) <> 0 Then bRunToolAutoruns = True: gNoGUI = True
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
                sPath = ""
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
        If Not CheckAccessWrite(g_sLogFile, True) Then sPath = ""
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
    AppendErrorLogCustom vbCrLf & vbCrLf & "Logfile ( tracing ) of HiJackThis Fork v." & AppVerString & vbCrLf & vbCrLf & _
        "Command line: " & AppPath(True) & " " & g_sCommandLine & vbCrLf & vbCrLf & MakeLogHeader()
    
    'Me.Caption = Me.Caption & AppVerString

    bForceRU = InStr(1, AppExeName(), "_RU", 1) Or HasCommandLineKey("langRU")  '/langRU
    bForceEN = InStr(1, AppExeName(), "_EN", 1) Or HasCommandLineKey("langEN")  '/langEN
    bForceUA = InStr(1, AppExeName(), "_UA", 1) Or HasCommandLineKey("langUA")  '/langUA
    bForceFR = InStr(1, AppExeName(), "_FR", 1) Or HasCommandLineKey("langFR")  '/langUA

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
    Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
    
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
        
        lret = DeleteFileW(StrPtr(sPathComCtl1))
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

Function GetEULA() As String
    Dim sText$
    sText = sText & " \n"
    sText = sText & "                                            GNU GENERAL PUBLIC LICENSE \n"
    sText = sText & "                                                       Version 2, June 1991 \n"
    sText = sText & " \n"
    sText = sText & " Copyright (C) 1989, 1991 Free Software Foundation, Inc., \n"
    sText = sText & " 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA \n"
    sText = sText & " Everyone is permitted to copy and distribute verbatim copies \n"
    sText = sText & " of this license document, but changing it is not allowed. \n"
    sText = sText & " \n"
    sText = sText & "                            Preamble \n"
    sText = sText & " \n"
    sText = sText & "  The licenses for most software are designed to take away your \n"
    sText = sText & "freedom to share and change it.  By contrast, the GNU General Public \n"
    sText = sText & "License is intended to guarantee your freedom to share and change free \n"
    sText = sText & "software--to make sure the software is free for all its users.  This \n"
    sText = sText & "General Public License applies to most of the Free Software \n"
    sText = sText & "Foundation's software and to any other program whose authors commit to \n"
    sText = sText & "using it.  (Some other Free Software Foundation software is covered by \n"
    sText = sText & "the GNU Lesser General Public License instead.)  You can apply it to \n"
    sText = sText & "your programs, too. \n"
    sText = sText & " \n"
    sText = sText & "  When we speak of free software, we are referring to freedom, not \n"
    sText = sText & "price.  Our General Public Licenses are designed to make sure that you \n"
    sText = sText & "have the freedom to distribute copies of free software (and charge for \n"
    sText = sText & "this service if you wish), that you receive source code or can get it \n"
    sText = sText & "if you want it, that you can change the software or use pieces of it \n"
    sText = sText & "in new free programs; and that you know you can do these things. \n"
    sText = sText & " \n"
    sText = sText & "  To protect your rights, we need to make restrictions that forbid \n"
    sText = sText & "anyone to deny you these rights or to ask you to surrender the rights. \n"
    sText = sText & "These restrictions translate to certain responsibilities for you if you \n"
    sText = sText & "distribute copies of the software, or if you modify it. \n"
    sText = sText & " \n"
    sText = sText & "  For example, if you distribute copies of such a program, whether \n"
    sText = sText & "gratis or for a fee, you must give the recipients all the rights that \n"
    sText = sText & "you have.  You must make sure that they, too, receive or can get the \n"
    sText = sText & "source code.  And you must show them these terms so they know their \n"
    sText = sText & "rights. \n"
    sText = sText & " \n"
    sText = sText & "  We protect your rights with two steps: (1) copyright the software, and \n"
    sText = sText & "(2) offer you this license which gives you legal permission to copy, \n"
    sText = sText & "distribute and/or modify the software. \n"
    sText = sText & " \n"
    sText = sText & "  Also, for each author's protection and ours, we want to make certain \n"
    sText = sText & "that everyone understands that there is no warranty for this free \n"
    sText = sText & "software.  If the software is modified by someone else and passed on, we \n"
    sText = sText & "want its recipients to know that what they have is not the original, so \n"
    sText = sText & "that any problems introduced by others will not reflect on the original \n"
    sText = sText & "authors' reputations. \n"
    sText = sText & " \n"
    sText = sText & "  Finally, any free program is threatened constantly by software \n"
    sText = sText & "patents.  We wish to avoid the danger that redistributors of a free \n"
    sText = sText & "program will individually obtain patent licenses, in effect making the \n"
    sText = sText & "program proprietary.  To prevent this, we have made it clear that any \n"
    sText = sText & "patent must be licensed for everyone's free use or not licensed at all. \n"
    sText = sText & " \n"
    sText = sText & "  The precise terms and conditions for copying, distribution and \n"
    sText = sText & "modification follow. \n"
    sText = sText & " \n"
    sText = sText & "                    GNU GENERAL PUBLIC LICENSE \n"
    sText = sText & "   TERMS AND CONDITIONS FOR COPYING, DISTRIBUTION AND MODIFICATION \n"
    sText = sText & " \n"
    sText = sText & "  0. This License applies to any program or other work which contains \n"
    sText = sText & "a notice placed by the copyright holder saying it may be distributed \n"
    sText = sText & "under the terms of this General Public License.  The ""Program"", below, \n"
    sText = sText & "refers to any such program or work, and a ""work based on the Program"" \n"
    sText = sText & "means either the Program or any derivative work under copyright law: \n"
    sText = sText & "that is to say, a work containing the Program or a portion of it, \n"
    sText = sText & "either verbatim or with modifications and/or translated into another \n"
    sText = sText & "language.  (Hereinafter, translation is included without limitation in \n"
    sText = sText & "the term ""modification"".)  Each licensee is addressed as ""you"". \n"
    sText = sText & " \n"
    sText = sText & "Activities other than copying, distribution and modification are not \n"
    sText = sText & "covered by this License; they are outside its scope.  The act of \n"
    sText = sText & "running the Program is not restricted, and the output from the Program \n"
    sText = sText & "is covered only if its contents constitute a work based on the \n"
    sText = sText & "Program (independent of having been made by running the Program). \n"
    sText = sText & "Whether that is true depends on what the Program does. \n"
    sText = sText & " \n"
    sText = sText & "  1. You may copy and distribute verbatim copies of the Program's \n"
    sText = sText & "source code as you receive it, in any medium, provided that you \n"
    sText = sText & "conspicuously and appropriately publish on each copy an appropriate \n"
    sText = sText & "copyright notice and disclaimer of warranty; keep intact all the \n"
    sText = sText & "notices that refer to this License and to the absence of any warranty; \n"
    sText = sText & "and give any other recipients of the Program a copy of this License \n"
    sText = sText & "along with the Program. \n"
    sText = sText & " \n"
    sText = sText & "You may charge a fee for the physical act of transferring a copy, and \n"
    sText = sText & "you may at your option offer warranty protection in exchange for a fee. \n"
    sText = sText & " \n"
    sText = sText & "  2. You may modify your copy or copies of the Program or any portion \n"
    sText = sText & "of it, thus forming a work based on the Program, and copy and \n"
    sText = sText & "distribute such modifications or work under the terms of Section 1 \n"
    sText = sText & "above, provided that you also meet all of these conditions: \n"
    sText = sText & " \n"
    sText = sText & "    a) You must cause the modified files to carry prominent notices \n"
    sText = sText & "    stating that you changed the files and the date of any change. \n"
    sText = sText & " \n"
    sText = sText & "    b) You must cause any work that you distribute or publish, that in \n"
    sText = sText & "    whole or in part contains or is derived from the Program or any \n"
    sText = sText & "    part thereof, to be licensed as a whole at no charge to all third \n"
    sText = sText & "    parties under the terms of this License. \n"
    sText = sText & " \n"
    sText = sText & "    c) If the modified program normally reads commands interactively \n"
    sText = sText & "    when run, you must cause it, when started running for such \n"
    sText = sText & "    interactive use in the most ordinary way, to print or display an \n"
    sText = sText & "    announcement including an appropriate copyright notice and a \n"
    sText = sText & "    notice that there is no warranty (or else, saying that you provide \n"
    sText = sText & "    a warranty) and that users may redistribute the program under \n"
    sText = sText & "    these conditions, and telling the user how to view a copy of this \n"
    sText = sText & "    License.  (Exception: if the Program itself is interactive but \n"
    sText = sText & "    does not normally print such an announcement, your work based on \n"
    sText = sText & "    the Program is not required to print an announcement.) \n"
    sText = sText & " \n"
    sText = sText & "These requirements apply to the modified work as a whole.  If \n"
    sText = sText & "identifiable sections of that work are not derived from the Program, \n"
    sText = sText & "and can be reasonably considered independent and separate works in \n"
    sText = sText & "themselves, then this License, and its terms, do not apply to those \n"
    sText = sText & "sections when you distribute them as separate works.  But when you \n"
    sText = sText & "distribute the same sections as part of a whole which is a work based \n"
    sText = sText & "on the Program, the distribution of the whole must be on the terms of \n"
    sText = sText & "this License, whose permissions for other licensees extend to the \n"
    sText = sText & "entire whole, and thus to each and every part regardless of who wrote it. \n"
    sText = sText & " \n"
    sText = sText & "Thus, it is not the intent of this section to claim rights or contest \n"
    sText = sText & "your rights to work written entirely by you; rather, the intent is to \n"
    sText = sText & "exercise the right to control the distribution of derivative or \n"
    sText = sText & "collective works based on the Program. \n"
    sText = sText & " \n"
    sText = sText & "In addition, mere aggregation of another work not based on the Program \n"
    sText = sText & "with the Program (or with a work based on the Program) on a volume of \n"
    sText = sText & "a storage or distribution medium does not bring the other work under \n"
    sText = sText & "the scope of this License. \n"
    sText = sText & " \n"
    sText = sText & "  3. You may copy and distribute the Program (or a work based on it, \n"
    sText = sText & "under Section 2) in object code or executable form under the terms of \n"
    sText = sText & "Sections 1 and 2 above provided that you also do one of the following: \n"
    sText = sText & " \n"
    sText = sText & "    a) Accompany it with the complete corresponding machine-readable \n"
    sText = sText & "    source code, which must be distributed under the terms of Sections \n"
    sText = sText & "    1 and 2 above on a medium customarily used for software interchange; or, \n"
    sText = sText & " \n"
    sText = sText & "    b) Accompany it with a written offer, valid for at least three \n"
    sText = sText & "    years, to give any third party, for a charge no more than your \n"
    sText = sText & "    cost of physically performing source distribution, a complete \n"
    sText = sText & "    machine-readable copy of the corresponding source code, to be \n"
    sText = sText & "    distributed under the terms of Sections 1 and 2 above on a medium \n"
    sText = sText & "    customarily used for software interchange; or, \n"
    sText = sText & " \n"
    sText = sText & "    c) Accompany it with the information you received as to the offer \n"
    sText = sText & "    to distribute corresponding source code.  (This alternative is \n"
    sText = sText & "    allowed only for noncommercial distribution and only if you \n"
    sText = sText & "    received the program in object code or executable form with such \n"
    sText = sText & "    an offer, in accord with Subsection b above.) \n"
    sText = sText & " \n"
    sText = sText & "The source code for a work means the preferred form of the work for \n"
    sText = sText & "making modifications to it.  For an executable work, complete source \n"
    sText = sText & "code means all the source code for all modules it contains, plus any \n"
    sText = sText & "associated interface definition files, plus the scripts used to \n"
    sText = sText & "control compilation and installation of the executable.  However, as a \n"
    sText = sText & "special exception, the source code distributed need not include \n"
    sText = sText & "anything that is normally distributed (in either source or binary \n"
    sText = sText & "form) with the major components (compiler, kernel, and so on) of the \n"
    sText = sText & "operating system on which the executable runs, unless that component \n"
    sText = sText & "itself accompanies the executable. \n"
    sText = sText & " \n"
    sText = sText & "If distribution of executable or object code is made by offering \n"
    sText = sText & "access to copy from a designated place, then offering equivalent \n"
    sText = sText & "access to copy the source code from the same place counts as \n"
    sText = sText & "distribution of the source code, even though third parties are not \n"
    sText = sText & "compelled to copy the source along with the object code. \n"
    sText = sText & " \n"
    sText = sText & "  4. You may not copy, modify, sublicense, or distribute the Program \n"
    sText = sText & "except as expressly provided under this License.  Any attempt \n"
    sText = sText & "otherwise to copy, modify, sublicense or distribute the Program is \n"
    sText = sText & "void, and will automatically terminate your rights under this License. \n"
    sText = sText & "However, parties who have received copies, or rights, from you under \n"
    sText = sText & "this License will not have their licenses terminated so long as such \n"
    sText = sText & "parties remain in full compliance. \n"
    sText = sText & " \n"
    sText = sText & "  5. You are not required to accept this License, since you have not \n"
    sText = sText & "signed it.  However, nothing else grants you permission to modify or \n"
    sText = sText & "distribute the Program or its derivative works.  These actions are \n"
    sText = sText & "prohibited by law if you do not accept this License.  Therefore, by \n"
    sText = sText & "modifying or distributing the Program (or any work based on the \n"
    sText = sText & "Program), you indicate your acceptance of this License to do so, and \n"
    sText = sText & "all its terms and conditions for copying, distributing or modifying \n"
    sText = sText & "the Program or works based on it. \n"
    sText = sText & " \n"
    sText = sText & "  6. Each time you redistribute the Program (or any work based on the \n"
    sText = sText & "Program), the recipient automatically receives a license from the \n"
    sText = sText & "original licensor to copy, distribute or modify the Program subject to \n"
    sText = sText & "these terms and conditions.  You may not impose any further \n"
    sText = sText & "restrictions on the recipients' exercise of the rights granted herein. \n"
    sText = sText & "You are not responsible for enforcing compliance by third parties to \n"
    sText = sText & "this License. \n"
    sText = sText & " \n"
    sText = sText & "  7. If, as a consequence of a court judgment or allegation of patent \n"
    sText = sText & "infringement or for any other reason (not limited to patent issues), \n"
    sText = sText & "conditions are imposed on you (whether by court order, agreement or \n"
    sText = sText & "otherwise) that contradict the conditions of this License, they do not \n"
    sText = sText & "excuse you from the conditions of this License.  If you cannot \n"
    sText = sText & "distribute so as to satisfy simultaneously your obligations under this \n"
    sText = sText & "License and any other pertinent obligations, then as a consequence you \n"
    sText = sText & "may not distribute the Program at all.  For example, if a patent \n"
    sText = sText & "license would not permit royalty-free redistribution of the Program by \n"
    sText = sText & "all those who receive copies directly or indirectly through you, then \n"
    sText = sText & "the only way you could satisfy both it and this License would be to \n"
    sText = sText & "refrain entirely from distribution of the Program. \n"
    sText = sText & " \n"
    sText = sText & "If any portion of this section is held invalid or unenforceable under \n"
    sText = sText & "any particular circumstance, the balance of the section is intended to \n"
    sText = sText & "apply and the section as a whole is intended to apply in other \n"
    sText = sText & "circumstances. \n"
    sText = sText & " \n"
    sText = sText & "It is not the purpose of this section to induce you to infringe any \n"
    sText = sText & "patents or other property right claims or to contest validity of any \n"
    sText = sText & "such claims; this section has the sole purpose of protecting the \n"
    sText = sText & "integrity of the free software distribution system, which is \n"
    sText = sText & "implemented by public license practices.  Many people have made \n"
    sText = sText & "generous contributions to the wide range of software distributed \n"
    sText = sText & "through that system in reliance on consistent application of that \n"
    sText = sText & "system; it is up to the author/donor to decide if he or she is willing \n"
    sText = sText & "to distribute software through any other system and a licensee cannot \n"
    sText = sText & "impose that choice. \n"
    sText = sText & " \n"
    sText = sText & "This section is intended to make thoroughly clear what is believed to \n"
    sText = sText & "be a consequence of the rest of this License. \n"
    sText = sText & " \n"
    sText = sText & "  8. If the distribution and/or use of the Program is restricted in \n"
    sText = sText & "certain countries either by patents or by copyrighted interfaces, the \n"
    sText = sText & "original copyright holder who places the Program under this License \n"
    sText = sText & "may add an explicit geographical distribution limitation excluding \n"
    sText = sText & "those countries, so that distribution is permitted only in or among \n"
    sText = sText & "countries not thus excluded.  In such case, this License incorporates \n"
    sText = sText & "the limitation as if written in the body of this License. \n"
    sText = sText & " \n"
    sText = sText & "  9. The Free Software Foundation may publish revised and/or new versions \n"
    sText = sText & "of the General Public License from time to time.  Such new versions will \n"
    sText = sText & "be similar in spirit to the present version, but may differ in detail to \n"
    sText = sText & "address new problems or concerns. \n"
    sText = sText & " \n"
    sText = sText & "Each version is given a distinguishing version number.  If the Program \n"
    sText = sText & "specifies a version number of this License which applies to it and ""any \n"
    sText = sText & "later version"", you have the option of following the terms and conditions \n"
    sText = sText & "either of that version or of any later version published by the Free \n"
    sText = sText & "Software Foundation.  If the Program does not specify a version number of \n"
    sText = sText & "this License, you may choose any version ever published by the Free Software \n"
    sText = sText & "Foundation. \n"
    sText = sText & " \n"
    sText = sText & "  10. If you wish to incorporate parts of the Program into other free \n"
    sText = sText & "programs whose distribution conditions are different, write to the author \n"
    sText = sText & "to ask for permission.  For software which is copyrighted by the Free \n"
    sText = sText & "Software Foundation, write to the Free Software Foundation; we sometimes \n"
    sText = sText & "make exceptions for this.  Our decision will be guided by the two goals \n"
    sText = sText & "of preserving the free status of all derivatives of our free software and \n"
    sText = sText & "of promoting the sharing and reuse of software generally. \n"
    sText = sText & " \n"
    sText = sText & "                            NO WARRANTY \n"
    sText = sText & " \n"
    sText = sText & "  11. BECAUSE THE PROGRAM IS LICENSED FREE OF CHARGE, THERE IS NO  \n"
    sText = sText & "WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY APPLICABLE  \n"
    sText = sText & "LAW.  EXCEPT WHEN OTHERWISE STATED IN WRITING THE COPYRIGHT  \n"
    sText = sText & "HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM ""AS IS"" WITHOUT  \n"
    sText = sText & "WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT  \n"
    sText = sText & "NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND  \n"
    sText = sText & "FITNESS FOR A PARTICULAR PURPOSE.  THE ENTIRE RISK AS TO THE QUALITY  \n"
    sText = sText & "AND PERFORMANCE OF THE PROGRAM IS WITH YOU.  SHOULD THE PROGRAM  \n"
    sText = sText & "PROVE DEFECTIVE, YOU ASSUME THE COST OF ALL NECESSARY SERVICING,  \n"
    sText = sText & "REPAIR OR CORRECTION. \n"
    sText = sText & " \n"
    sText = sText & "  12. IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN  \n"
    sText = sText & "WRITING WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MAY  \n"
    sText = sText & "MODIFY AND/OR REDISTRIBUTE THE PROGRAM AS PERMITTED ABOVE, BE  \n"
    sText = sText & "LIABLE TO YOU FOR DAMAGES, INCLUDING ANY GENERAL, SPECIAL,  \n"
    sText = sText & "INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE USE OR  \n"
    sText = sText & "INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED TO LOSS OF  \n"
    sText = sText & "DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY  \n"
    sText = sText & "YOU OR THIRD PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH  \n"
    sText = sText & "ANY OTHER PROGRAMS), EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN  \n"
    sText = sText & "ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. \n"
    sText = sText & " \n"
    sText = sText & "                     END OF TERMS AND CONDITIONS \n"
    sText = sText & " \n"
    sText = sText & "            How to Apply These Terms to Your New Programs \n"
    sText = sText & " \n"
    sText = sText & "  If you develop a new program, and you want it to be of the greatest \n"
    sText = sText & "possible use to the public, the best way to achieve this is to make it \n"
    sText = sText & "free software which everyone can redistribute and change under these terms. \n"
    sText = sText & " \n"
    sText = sText & "  To do so, attach the following notices to the program.  It is safest \n"
    sText = sText & "to attach them to the start of each source file to most effectively \n"
    sText = sText & "convey the exclusion of warranty; and each file should have at least \n"
    sText = sText & "the ""copyright"" line and a pointer to where the full notice is found. \n"
    sText = sText & " \n"
    sText = sText & "    <one line to give the program's name and a brief idea of what it does.> \n"
    sText = sText & "    Copyright (C) <year>  <name of author> \n"
    sText = sText & " \n"
    sText = sText & "    This program is free software; you can redistribute it and/or modify \n"
    sText = sText & "    it under the terms of the GNU General Public License as published by \n"
    sText = sText & "    the Free Software Foundation; either version 2 of the License, or \n"
    sText = sText & "    (at your option) any later version. \n"
    sText = sText & " \n"
    sText = sText & "    This program is distributed in the hope that it will be useful, \n"
    sText = sText & "    but WITHOUT ANY WARRANTY; without even the implied warranty of \n"
    sText = sText & "    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the \n"
    sText = sText & "    GNU General Public License for more details. \n"
    sText = sText & " \n"
    sText = sText & "    You should have received a copy of the GNU General Public License along \n"
    sText = sText & "    with this program; if not, write to the Free Software Foundation, Inc., \n"
    sText = sText & "    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA. \n"
    sText = sText & " \n"
    sText = sText & "Also add information on how to contact you by electronic and paper mail. \n"
    sText = sText & " \n"
    sText = sText & "If the program is interactive, make it output a short notice like this \n"
    sText = sText & "when it starts in an interactive mode: \n"
    sText = sText & " \n"
    sText = sText & "    Gnomovision version 69, Copyright (C) year name of author \n"
    sText = sText & "    Gnomovision comes with ABSOLUTELY NO WARRANTY; for details type `show w'. \n"
    sText = sText & "    This is free software, and you are welcome to redistribute it \n"
    sText = sText & "    under certain conditions; type `show c' for details. \n"
    sText = sText & " \n"
    sText = sText & "The hypothetical commands `show w' and `show c' should show the appropriate \n"
    sText = sText & "parts of the General Public License.  Of course, the commands you use may \n"
    sText = sText & "be called something other than `show w' and `show c'; they could even be \n"
    sText = sText & "mouse-clicks or menu items--whatever suits your program. \n"
    sText = sText & " \n"
    sText = sText & "You should also get your employer (if you work as a programmer) or your \n"
    sText = sText & "school, if any, to sign a ""copyright disclaimer"" for the program, if \n"
    sText = sText & "necessary.  Here is a sample; alter the names: \n"
    sText = sText & " \n"
    sText = sText & "  Yoyodyne, Inc., hereby disclaims all copyright interest in the program \n"
    sText = sText & "  `Gnomovision' (which makes passes at compilers) written by James Hacker. \n"
    sText = sText & " \n"
    sText = sText & "  <signature of Ty Coon>, 1 April 1989 \n"
    sText = sText & "  Ty Coon, President of Vice \n"
    sText = sText & " \n"
    sText = sText & "This General Public License does not permit incorporating your program into \n"
    sText = sText & "proprietary programs.  If your program is a subroutine library, you may \n"
    sText = sText & "consider it more useful to permit linking proprietary applications with the \n"
    sText = sText & "library.  If this is what you want to do, use the GNU Lesser General \n"
    sText = sText & "Public License instead of this License. \n"
    
    GetEULA = Replace$(sText, "\n", vbCrLf)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set ControlsEvent = Nothing
End Sub
