VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Смена иконы у EXE"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3360
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtIco 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox txtExe 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
   Begin VB.PictureBox dlg2 
      Height          =   480
      Left            =   1800
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   1080
      Width           =   1200
   End
   Begin VB.CommandButton cmdChangeIcon 
      Caption         =   "Сменить"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowseIco 
      Caption         =   "Выбрать икону"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowseExe 
      Caption         =   "Выбрать файл"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CommandLineToArgvW Lib "Shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Sub cmdBrowseExe_Click()
 With dlg
 .DialogTitle = "Select Exe File..."
 .Filter = "Executable Files (*.exe)|*.exe"
 .ShowOpen
 End With

 txtExe.Text = dlg.FileName
 End Sub

 Private Sub cmdBrowseIco_Click()
 With dlg
 .DialogTitle = "Select Icon File..."
 .Filter = "Icons (*.ico)|*.ico"
 .ShowOpen
 End With

 txtIco.Text = dlg.FileName
 End Sub

 Private Sub cmdChangeIcon_Click()
 If ChangeIcon(txtExe.Text, txtIco.Text) Then
 MsgBox "Done"
 Else
 MsgBox "Error Occurred."
 End If
 End Sub

Private Sub Form_Load()
    Dim arg() As String
    If Command() <> "" Then
        Me.Hide
        ParseCommandLine Command(), arg
        If UBound(arg) > 0 Then
            ChangeIcon arg(0), arg(1)
        Else
            MsgBox "Использование: " & App.EXEName & " [Исходный EXE] [Файл .ico]"
        End If
        End
    End If
End Sub

Private Function ParseCommandLine(Line As String, Out() As String) As Boolean
    Dim ptr     As Long
    Dim count   As Long
    Dim index   As Long
    Dim strLen  As Long
    Dim strAdr  As Long
    
    ptr = CommandLineToArgvW(StrPtr(Line), count)
    
    If count < 1 Then Exit Function
    
    ReDim Out(count - 1)
    
    For index = 0 To count - 1
        GetMem4 ByVal ptr + index * 4, strAdr
        strLen = lstrlen(strAdr)
        Out(index) = Space(strLen)
        lstrcpyn StrPtr(Out(index)), strAdr, strLen + 1
    Next
    
    GlobalFree ptr
    
    ParseCommandLine = True
    
End Function
