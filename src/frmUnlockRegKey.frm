VERSION 5.00
Object = "{317589D1-37C8-47D9-B5B0-1C995741F353}#1.0#0"; "VBCCR17.OCX"
Begin VB.Form frmUnlockRegKey 
   Caption         =   "Registry Key Unlocker"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmUnlockRegKey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   8445
   Begin VBCCR17.CommandButtonW cmdJump 
      Height          =   450
      Left            =   6600
      TabIndex        =   4
      Top             =   60
      Width           =   1692
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Open in Regedit"
   End
   Begin VBCCR17.CommandButtonW cmdGo 
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Go"
   End
   Begin VBCCR17.CheckBoxW chkRecur 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
      _ExtentX        =   0
      _ExtentY        =   0
      Value           =   1
      Caption         =   "Recursively (process keys and all subkeys)"
   End
   Begin VBCCR17.TextBoxW txtKeys 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   8055
      _ExtentX        =   0
      _ExtentY        =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3
   End
   Begin VBCCR17.LabelW lblWhatToDo 
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6132
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Enter Registry Key(s) to unlock and reset access:"
   End
End
Attribute VB_Name = "frmUnlockRegKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmUnlockRegKey.frm]

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
    Dim sLogPath    As String
    Dim TimeStarted As String
    Dim TimeFinished As String
    Dim sList As clsStringBuilder
    
    sLogPath = BuildPath(AppPath(), "FixReg.log")
    
    sKeys = txtKeys.Text
    
    If Len(sKeys) = 0 Then
        'You should enter at least one key!
        MsgBoxW Translate(1905), vbExclamation
        Exit Sub
    End If
    
    TimeStarted = GetTime()
    
    Set sList = New clsStringBuilder
    sList.Append ChrW$(-257)
    sList.AppendLine "Logfile of Registry Key Unlocker (HJT+ v." & AppVerString & ")"
    sList.AppendLine
    sList.AppendLine MakeLogHeader()
    sList.AppendLine
    sList.AppendLine "Logging started at:      " & TimeStarted
    sList.AppendLine
    
    TimeStarted = GetTime()
    
    Recursively = (chkRecur.Value = 1)
    
    sKeys = Replace$(sKeys, vbCr, vbNullString)
    aKeys = Split(sKeys, vbLf)
    
    For Each vKey In aKeys
        If Len(vKey) <> 0 Then
            ' // TODO: log each reg key separately (only [Failed] events)
            '
            If True = modPermissions.RegKeyResetDACL(0&, CStr(vKey), False, Recursively) Then
                '[OK]
                '(recursively)
                sList.AppendLine Translate(1906) & " - " & vKey & IIf(Recursively, " " & Translate(1907), vbNullString)
            Else
                '[Fail]
                sList.AppendLine Translate(1908) & " - " & vKey
            End If
        End If
    Next
    
    sList.AppendLine
    TimeFinished = GetTime()
    sList.AppendLine "Logging finished at:     " & TimeFinished
    sList.AppendLine
    sList.Append "--" & vbCrLf & "End of file"
    
    If OpenW(sLogPath, FOR_OVERWRITE_CREATE, hFile, g_FileBackupFlag) Then
        PutW hFile, 1, StrPtr(sList.ToString), sList.Length * 2
        CloseW hFile, True
    End If
    
    OpenLogFile sLogPath
    
    Set sList = Nothing
    
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

Private Sub cmdJump_Click()
    Dim sKeys As String
    Dim aKeys() As String
    
    sKeys = txtKeys.Text
    
    If Len(sKeys) = 0 Then
        'You should enter at least one key!
        MsgBoxW Translate(1905), vbExclamation
        Exit Sub
    End If
    
    sKeys = Replace$(sKeys, vbCr, vbNullString)
    aKeys = Split(sKeys, vbLf)
    
    Reg.Jump 0, aKeys(0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
    ProcessHotkey KeyCode, Me
End Sub

Private Sub Form_Load()
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    ReloadLanguage True
    LoadWindowPos Me, SETTINGS_SECTION_REGUNLOCKER
    SubClassTextbox Me.txtKeys.hWnd, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    SaveWindowPos Me, SETTINGS_SECTION_REGUNLOCKER

    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        Cancel = True
        Me.Hide
    Else
        SubClassTextbox Me.txtKeys.hWnd, False
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Width < 7860 Then Me.Width = 7860
        If Me.Height < 2570 Then Me.Height = 2570
    End If
    txtKeys.Width = Me.Width - 630
    txtKeys.Height = Me.Height - 2010
    chkRecur.Top = Me.Height - 1300
    cmdGo.Top = Me.Height - 1300
End Sub

Private Sub txtKeys_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub
