VERSION 5.00
Begin VB.Form frmError 
   Caption         =   "ERROR"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox chkNoMoreErrors 
      Caption         =   "Do not show this message again"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Height          =   3975
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Private Declare Function MessageBeep Lib "user32.dll" (ByVal uType As Long) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconW" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Private Const IDI_ASTERISK      As Long = 32516&    'Information
Private Const IDI_EXCLAMATION   As Long = 32515&    'Exclamation
Private Const IDI_HAND          As Long = 32513&    'Critical Stop
Private Const IDI_QUESTION      As Long = 32514&    'Question Mark

Private Const MB_ICONERROR      As Long = &H10&


Private Sub chkNoMoreErrors_Click()
    frmMain.chkSkipErrorMsg.value = chkNoMoreErrors.value
    bSkipErrorMsg = (chkNoMoreErrors.value = 1)
End Sub

Private Sub cmdNo_Click()
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    Dim szCrashUrl As String
    szCrashUrl = "http://safezone.cc/threads/25222/" 'https://sourceforge.net/p/hjt/_list/tickets"
    ShellExecute 0&, StrPtr("open"), StrPtr(szCrashUrl), 0&, 0&, vbNormalFocus
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim Icon As Long
    
    CenterForm Me
    
    With Picture1
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .BorderStyle = 0
    End With
    
    Icon = LoadIcon(0&, IDI_HAND)
    DrawIcon Picture1.hdc, 0&, 0&, Icon
    
    MessageBeep MB_ICONERROR
    
    Me.Caption = Translate(591)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        Cancel = True
        Me.Hide
    End If
End Sub
