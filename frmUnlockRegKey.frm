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
Option Explicit

Private Sub cmdGo_Click()
    On Error GoTo ErrorHandler:

    Dim sKeys As String
    Dim aKeys, Key
    Dim Recursively As Boolean
    Dim ff As Long
    Dim FixLines As String
    
    sKeys = Text1.Text
    
    If sKeys = "" Then
        'You should enter at least one key!
        MsgBoxW Translate(1905), vbExclamation
        Exit Sub
    End If
    
    Recursively = (chkRecur.value = 1)
    
    sKeys = Replace$(sKeys, vbCr, "")
    aKeys = Split(sKeys, vbLf)
    
    For Each Key In aKeys
        If Len(Key) <> 0 Then
            If True = modPermissions.RegKeyResetDACL(0&, CStr(Key), False, Recursively) Then
                '[OK]
                '(recursively)
                FixLines = FixLines & Translate(1906) & " - " & Key & IIf(Recursively, " " & Translate(1907), "") & vbCrLf
            Else
                '[Fail]
                FixLines = FixLines & Translate(1908) & " - " & Key & vbCrLf
            End If
        End If
    Next
    
    '// TODO: add unicode support
    
    ff = FreeFile()
    FixLog = BuildPath(AppPath(), "FixReg.log")
    Open FixLog For Append As #ff
    Print #ff, FixLines
    Close ff
    Shell "notepad.exe" & " " & """" & FixLog & """", vbNormalFocus
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmUnlockRegKey.cmdGo_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CmdExit_Click()
    Me.Hide
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
    If Me.Width < 7860 Then Me.Width = 7860
    If Me.Height < 2570 Then Me.Height = 2570
    Text1.Width = Me.Width - 630
    Text1.Height = Me.Height - 2010
    chkRecur.Top = Me.Height - 1300
    cmdGo.Top = Me.Height - 1300
    CmdExit.Top = Me.Height - 1300
End Sub
