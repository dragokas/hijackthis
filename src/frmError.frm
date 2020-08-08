VERSION 5.00
Begin VB.Form frmError 
   Caption         =   "ERROR"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   7080
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   120
      ScaleHeight     =   684
      ScaleWidth      =   684
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   495
      Left            =   5400
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
'[frmError.frm]

'
' Errors window by Alex Dragokas
'

Option Explicit

Private Declare Function MessageBeep Lib "user32.dll" (ByVal uType As Long) As Long
Private Declare Function LoadIcon Lib "user32.dll" Alias "LoadIconW" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function DrawIcon Lib "user32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

Private Const IDI_ASTERISK      As Long = 32516&    'Information
Private Const IDI_EXCLAMATION   As Long = 32515&    'Exclamation
Private Const IDI_HAND          As Long = 32513&    'Critical Stop
Private Const IDI_QUESTION      As Long = 32514&    'Question Mark

Private Const MB_ICONERROR      As Long = &H10&


Private Sub chkNoMoreErrors_Click()
    frmMain.chkSkipErrorMsg.Value = chkNoMoreErrors.Value
    bSkipErrorMsg = (chkNoMoreErrors.Value = 1)
End Sub

Private Sub cmdNo_Click()
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    'https://sourceforge.net/p/hjt/_list/tickets"
    OpenURL "https://github.com/dragokas/hijackthis/issues/4", "https://safezone.cc/threads/28770/"
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub

Private Sub Form_Load()
    Dim Icon As Long
    
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    'ReloadLanguage
    
    With Me
        If AryItems(TranslateNative) Then
            .Caption = TranslateNative(550)
            .chkNoMoreErrors.Caption = TranslateNative(551)
            .cmdYes.Caption = TranslateNative(552)
            .cmdNo.Caption = TranslateNative(553)
        Else
            .Caption = "ERROR"
            .chkNoMoreErrors.Caption = "Do not show this message again"
            .cmdYes.Caption = "Yes"
            .cmdNo.Caption = "No"
        End If
    End With
    
    CenterForm Me
    'Me.Icon = frmMain.Icon 'main form may not be initialized yet, so skip this line!
    
    With Picture1
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .BorderStyle = 0
    End With
    
    Icon = LoadIcon(0&, IDI_HAND)
    DrawIcon Picture1.hdc, 0&, 0&, Icon
    
    MessageBeep MB_ICONERROR
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        Cancel = True
        Me.Hide
    End If
End Sub
