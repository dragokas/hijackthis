VERSION 5.00
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkEscSeq 
      Caption         =   "Esc-sequences"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkRegExp 
      Caption         =   "Regular expressions"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "Whole word"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.Frame frmDir 
      Caption         =   "Direction"
      Height          =   1695
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   2535
      Begin VB.OptionButton optDirEnd 
         Caption         =   "Ending"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optDirBegin 
         Caption         =   "Beginning"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optDirUp 
         Caption         =   "Up"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optDirDown 
         Caption         =   "Down"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match case"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "Find next"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblEscSeq 
      Caption         =   "\[0020], \\,  \n,  \t"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblWhat 
      Caption         =   "What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bEscSeq As Boolean
Private bRegExp As Boolean
Private bMatchCase As Boolean
Private bWholeWord As Boolean


Private Sub Form_Load()
    Dim OptB As OptionButton
    Dim Ctl As Control
    
    Me.Icon = frmMain.Icon
    CenterForm Me
    SetAllFontCharset Me, g_FontName, g_FontSize
    Call ReloadLanguage(True)
    
    ' if Win XP -> disable all window styles from option buttons
    If bIsWinXP Then
        For Each Ctl In Me.Controls
            If TypeName(Ctl) = "OptionButton" Then
                Set OptB = Ctl
                SetWindowTheme OptB.hwnd, StrPtr(" "), StrPtr(" ")
            End If
        Next
        Set OptB = Nothing
    End If
End Sub

'======== Options

Private Sub chkEscSeq_Click()
    If chkEscSeq.Value = vbChecked Then chkRegExp.Value = vbUnchecked
    bEscSeq = chkEscSeq.Value
End Sub

Private Sub chkRegExp_Click()
    If chkRegExp.Value = vbChecked Then chkEscSeq.Value = vbUnchecked
    bRegExp = chkRegExp.Value
End Sub

Private Sub chkMatchCase_Click()
    bMatchCase = chkMatchCase.Value
End Sub

Private Sub chkWholeWord_Click()
    bWholeWord = chkWholeWord.Value
End Sub

'======= Buttons

Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub CmdFind_Click()
    '
End Sub

'======= Direction

Private Sub optDirDown_Click()
    '
End Sub

Private Sub optDirUp_Click()
    '
End Sub

Private Sub optDirBegin_Click()
    '
End Sub

Private Sub optDirEnd_Click()
    '
End Sub
