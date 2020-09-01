VERSION 5.00
Begin VB.Form frmUnlockRegKey 
   Caption         =   "Registry Key Unlocker"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8448
   Icon            =   "frmUnlockRegKey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   8448
   Begin VB.CommandButton cmdJump 
      Caption         =   "Open in Regedit"
      Height          =   450
      Left            =   6480
      TabIndex        =   5
      Top             =   60
      Width           =   1572
   End
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
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox txtKeys 
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
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6132
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
    Dim FixLines    As String
    Dim sHeader     As String
    Dim sLogPath    As String
    Dim TimeStarted As String
    Dim TimeFinished As String
    Dim sList As clsStringBuilder
    
    sLogPath = BuildPath(AppPath(), "FixReg.log")
    
    sKeys = txtKeys.Text
    
    If sKeys = "" Then
        'You should enter at least one key!
        MsgBoxW Translate(1905), vbExclamation
        Exit Sub
    End If
    
    TimeStarted = GetTime()
    
    Set sList = New clsStringBuilder
    sList.Append ChrW$(-257)
    sList.AppendLine "Logfile of Registry Key Unlocker (HJT v." & AppVerString & ")"
    sList.AppendLine
    sList.AppendLine MakeLogHeader()
    sList.AppendLine "Logging started at:      " & TimeStarted
    sList.AppendLine
    
    TimeStarted = GetTime()
    
    Recursively = (chkRecur.Value = 1)
    
    sKeys = Replace$(sKeys, vbCr, "")
    aKeys = Split(sKeys, vbLf)
    
    For Each vKey In aKeys
        If Len(vKey) <> 0 Then
            If True = modPermissions.RegKeyResetDACL(0&, CStr(vKey), False, Recursively) Then
                '[OK]
                '(recursively)
                sList.AppendLine Translate(1906) & " - " & vKey & IIf(Recursively, " " & Translate(1907), "")
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
    
    If sKeys = "" Then
        'You should enter at least one key!
        MsgBoxW Translate(1905), vbExclamation
        Exit Sub
    End If
    
    sKeys = Replace$(sKeys, vbCr, "")
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    SaveWindowPos Me, SETTINGS_SECTION_REGUNLOCKER

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
    txtKeys.Width = Me.Width - 630
    txtKeys.Height = Me.Height - 2010
    chkRecur.Top = Me.Height - 1300
    cmdGo.Top = Me.Height - 1300
    CmdExit.Top = Me.Height - 1300
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub
