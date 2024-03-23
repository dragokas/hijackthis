VERSION 5.00
Object = "{317589D1-37C8-47D9-B5B0-1C995741F353}#1.0#0"; "VBCCR17.OCX"
Begin VB.Form frmUnlockRegKey 
   Caption         =   "Registry Key Unlocker"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmUnlockRegKey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8445
   Begin VBCCR17.CommandButtonW cmdJump 
      Height          =   450
      Left            =   6600
      TabIndex        =   2
      Top             =   60
      Width           =   1692
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Open in Regedit"
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
   Begin VBCCR17.FrameW FramePerm 
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2990
      Begin VBCCR17.OptionButtonW optPermCustom 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Caption         =   "Custom SDDL:"
      End
      Begin VBCCR17.OptionButtonW optPermDefault 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         Value           =   -1  'True
         Caption         =   "Default permissions"
      End
      Begin VBCCR17.TextBoxW txtSDDL 
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1085
         Enabled         =   0   'False
         MultiLine       =   -1  'True
         ScrollBars      =   1
      End
      Begin VBCCR17.CommandButtonW cmdPickSDDL 
         Height          =   330
         Left            =   4080
         TabIndex        =   7
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         Caption         =   "Pick from key..."
      End
   End
   Begin VBCCR17.FrameW FrameGo 
      Height          =   735
      Left            =   960
      TabIndex        =   8
      Top             =   4320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1296
      BorderStyle     =   0
      Caption         =   "FrameW2"
      Begin VBCCR17.CommandButtonW cmdGo 
         Height          =   495
         Left            =   5640
         TabIndex        =   9
         Top             =   120
         Width           =   1575
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   12648384
         Caption         =   "Go"
      End
      Begin VBCCR17.CheckBoxW chkRecur 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   873
         Value           =   1
         Caption         =   "Recursively (including files and subfolders)"
      End
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
    Dim sKey        As String
    Dim Recursively As Boolean
    Dim hFile       As Long
    Dim sLogPath    As String
    Dim TimeStarted As String
    Dim TimeFinished As String
    Dim sList As clsStringBuilder
    Dim SDDL        As String
    Dim bCustomPerm As Boolean
    Dim bSuccess    As Boolean
    Dim SDDL_Before As String
    Dim SDDL_After  As String
    
    If Me.optPermCustom.Value Then
        SDDL = Me.txtSDDL.Text
        
        If Not IsValidSDDL(SDDL) Then
            'Invalid SDDL specified!
            MsgBox Translate(2416), vbExclamation
            Exit Sub
        End If
        
        bCustomPerm = True
    End If
    
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
            sKey = Reg.Normalize(CStr(vKey))
            
            If Not Reg.KeyExists(0&, sKey) Then
                '[Not found]
                sList.AppendLine Translate(1912) & " - " & sKey
                GoTo Continue
            End If
            
            SDDL_Before = GetKeyStringSD(0&, sKey)
            
            If bCustomPerm Then
                bSuccess = SetRegKeyStringSD(0&, sKey, SDDL, False, Recursively)
            Else
                bSuccess = modPermissions.RegKeyResetDACL(0&, sKey, False, Recursively)
            End If
            
            SDDL_After = GetKeyStringSD(0&, sKey)
            
            ' // TODO: log each reg key separately (only [Failed] events)
            '
            If bSuccess Then
                '[OK]
                '(recursively)
                sList.AppendLine Translate(1906) & " - " & sKey & IIf(Recursively, " " & Translate(1907), vbNullString)
            Else
                '[Fail]
                sList.AppendLine Translate(1908) & " - " & sKey
            End If
            
            sList.AppendLine "." & vbCrLf & "  " & SDDL_Before & vbCrLf & "=>" & SDDL_After & vbCrLf
        End If
Continue:
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
        If Me.Width < 8505 Then Me.Width = 8505
        If Me.Height < 5505 Then Me.Height = 5505
    End If
    txtKeys.Width = Me.Width - 630
    txtKeys.Height = Me.Height - 3690
    FramePerm.Top = txtKeys.Top + txtKeys.Height + 50
    FrameGo.Top = FramePerm.Top + FramePerm.Height + 50
    Me.cmdJump.Left = txtKeys.Left + txtKeys.Width - cmdJump.Width
End Sub

Private Sub optPermCustom_Click()
    txtSDDL.Enabled = True
End Sub

Private Sub optPermDefault_Click()
    txtSDDL.Enabled = False
End Sub

Private Sub cmdPickSDDL_Click()
    Dim sKey As String
    sKey = InputBox(Translate(1910), "", "HKEY_LOCAL_MACHINE\")
    If StrPtr(sKey) = 0 Then Exit Sub
    If Len(sKey) = 0 Then
        txtSDDL.Text = vbNullString
        Exit Sub
    End If
    sKey = Reg.Normalize(sKey)
    If Reg.KeyExists(0, sKey) Then
        txtSDDL.Text = GetKeyStringSD(0, sKey)
        If Len(txtSDDL.Text) <> 0 Then
            optPermCustom.Value = True
        End If
    Else
        MsgBox Translate(1911) & vbNewLine & vbNewLine & sKey, vbExclamation
    End If
End Sub

Private Sub txtKeys_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub
