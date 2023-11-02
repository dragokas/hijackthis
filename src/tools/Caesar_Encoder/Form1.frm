VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Caesar encoder"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOnTop 
      Caption         =   "On top"
      Height          =   195
      Left            =   4440
      TabIndex        =   15
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkSubstituteCmd 
      Caption         =   "Substitute VB command"
      Height          =   195
      Left            =   2280
      TabIndex        =   14
      Top             =   120
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton btnClearToDecode 
      Caption         =   "Clear"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton btnClearOriginal 
      Caption         =   "Clear"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnCopyDecoded 
      Caption         =   "Copy"
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtDecoded 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   5295
   End
   Begin VB.CommandButton btnPasteToDecode 
      Caption         =   "Paste"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtToDecode 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5295
   End
   Begin VB.CommandButton btnCopyEncoded 
      Caption         =   "Copy"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton btnPasteOriginal 
      Caption         =   "Paste"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtEncoded 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox txtOriginal 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "Original"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Decode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7320
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Encoded"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Original"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Form_Initialize()
    Call InitCommonControls
End Sub

Private Sub Form_Load()
    SetWindowAlwaysOnTop Me.hWnd, True
End Sub

'------ ENCODE --------

Private Sub btnClearOriginal_Click()
    txtOriginal.Text = ""
    txtEncoded.Text = ""
End Sub

Private Sub btnCopyEncoded_Click()
    ClipboardSetText txtEncoded.Text
End Sub

Private Sub btnPasteOriginal_Click()
    txtOriginal.Text = ClipboardGetText()
End Sub

Private Sub chkOnTop_Click()
    SetWindowAlwaysOnTop Me.hWnd, (chkOnTop.Value = vbChecked)
End Sub

Private Sub txtOriginal_Change()
    If Len(txtOriginal.Text) <> 0 Then
        Dim sEncoded As String
        sEncoded = Caesar_Encode(txtOriginal.Text)
        If chkSubstituteCmd.Value = vbChecked Then sEncoded = "Caes_Decode(""" & sEncoded & """)"
        txtEncoded.Text = sEncoded
    End If
End Sub

'----- DECODE -------

Public Function TrimCmd(ByVal sCmd As String) As String
    If InStr(1, sCmd, "Caes_Decode(""", 1) <> 0 Then
        sCmd = Mid$(sCmd, Len("Caes_Decode(""") + 1)
        If Right$(sCmd, 1) = ")" Then sCmd = Left$(sCmd, Len(sCmd) - 1)
        If Right$(sCmd, 1) = """" Then sCmd = Left$(sCmd, Len(sCmd) - 1)
        TrimCmd = sCmd
    Else
        TrimCmd = sCmd
    End If
End Function

Private Sub btnClearToDecode_Click()
    txtToDecode.Text = ""
    txtDecoded.Text = ""
End Sub

Private Sub btnPasteToDecode_Click()
    txtToDecode.Text = ClipboardGetText()
End Sub

Private Sub txtToDecode_Change()
    If Len(txtToDecode.Text) <> 0 Then
        txtDecoded.Text = Caesar_Decode(TrimCmd(txtToDecode.Text))
    End If
End Sub

Private Sub btnCopyDecoded_Click()
    ClipboardSetText txtDecoded.Text
End Sub
