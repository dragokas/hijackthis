VERSION 5.00
Begin VB.Form frmUnlockRegKey 
   Caption         =   "Registry Key Unlocker"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox chkRecur 
      Caption         =   "Recursively (process keys and all subkeys)"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Width           =   8055
   End
   Begin VB.Label lblWhatToDo 
      Caption         =   "Enter Registry Key(s) to unlock and reset access:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmUnlockRegKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Registry key unlocker by Alex Dragokas
'

Option Explicit

Private Sub cmdGo_Click()
    On Error GoTo ErrorHandler:

    Dim sKeys       As String
    Dim aKeys()     As String
    Dim vKey
    Dim Recursively As Boolean
    Dim hFile       As Long
    Dim FixLines    As String
    Dim sHeader     As String
    Dim sLogPath    As String
    Dim TimeStarted As String
    Dim TimeFinished As String
    
    sLogPath = BuildPath(AppPath(), "FixReg.log")
    
    sKeys = Text1.Text
    
    If sKeys = "" Then
        'You should enter at least one key!
        MsgBoxW Translate(1905), vbExclamation
        Exit Sub
    End If
    
    TimeStarted = GetTime()
    
    Recursively = (chkRecur.Value = 1)
    
    sKeys = Replace$(sKeys, vbCr, "")
    aKeys = Split(sKeys, vbLf)
    
    For Each vKey In aKeys
        If Len(vKey) <> 0 Then
            If True = modPermissions.RegKeyResetDACL(0&, CStr(vKey), False, Recursively) Then
                '[OK]
                '(recursively)
                FixLines = FixLines & Translate(1906) & " - " & vKey & IIf(Recursively, " " & Translate(1907), "") & vbCrLf
            Else
                '[Fail]
                FixLines = FixLines & Translate(1908) & " - " & vKey & vbCrLf
            End If
        End If
    Next
    
    'If Not FileExists(sLogPath) Then
        sHeader = "Logfile of Registry Key Unlocker (HJT v." & AppVerString & ")" & vbCrLf & vbCrLf
    
        sHeader = sHeader & "Platform:  " & OSver.Bitness & " " & OSver.OSName & " (" & OSver.Edition & "), " & _
            OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & _
            IIf(OSver.ReleaseId <> 0, " (ReleaseId: " & OSver.ReleaseId & ")", "") & ", " & _
            "Service Pack: " & OSver.SPVer & "" & IIf(OSver.IsSafeBoot, " (Safe Boot)", "") & vbCrLf
        sHeader = sHeader & "Language:  " & "OS: " & OSver.LangSystemNameFull & " (" & "0x" & Hex$(OSver.LangSystemCode) & "). " & _
            "Display: " & OSver.LangDisplayNameFull & " (" & "0x" & Hex$(OSver.LangDisplayCode) & "). " & _
            "Non-Unicode: " & OSver.LangNonUnicodeNameFull & " (" & "0x" & Hex$(OSver.LangNonUnicodeCode) & ")" & vbCrLf
        
        If OSver.MajorMinor >= 6 Then
            sHeader = sHeader & "Elevated:  " & IIf(OSver.IsElevated, "Yes", "No") & vbCrLf
        End If
    
        sHeader = sHeader & "Ran by:    " & GetUser() & vbTab & "(group: " & OSver.UserType & ") on " & GetComputer()
    'End If
    
    OpenW sLogPath, FOR_OVERWRITE_CREATE, hFile
    If hFile > 0 Then
        'If Len(sHeader) <> 0 Then Print #ff, sHeader
        PrintW hFile, sHeader
        PrintW hFile, ""
        PrintW hFile, "Logging started at:      " & TimeStarted
        PrintW hFile, ""
        PrintW hFile, FixLines
        TimeFinished = GetTime()
        PrintW hFile, "Logging finished at:     " & TimeFinished & vbCrLf
        CloseW hFile
    End If
    
    Shell "notepad.exe" & " " & """" & sLogPath & """", vbNormalFocus
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmUnlockRegKey.cmdGo_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Function GetTime() As String
    GetTime = Right$("0" & Day(Now), 2) & "." & Right$("0" & Month(Now), 2) & "." & Year(Now) & " - " & _
            Right$("0" & Hour(Now), 2) & ":" & Right$("0" & Minute(Now), 2)
End Function

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub

Private Sub Form_Load()
    ReloadLanguage
    CenterForm Me
    Me.Icon = frmMain.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Width < 7860 Then Me.Width = 7860
        If Me.Height < 2570 Then Me.Height = 2570
    End If
    Text1.Width = Me.Width - 630
    Text1.Height = Me.Height - 2010
    chkRecur.Top = Me.Height - 1300
    cmdGo.Top = Me.Height - 1300
    cmdExit.Top = Me.Height - 1300
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub
