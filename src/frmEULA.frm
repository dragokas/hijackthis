VERSION 5.00
Begin VB.Form frmEULA 
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   4740
   ClientTop       =   4380
   ClientWidth     =   7800
   Icon            =   "frmEULA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgree 
      Caption         =   "I Accept"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtEULA 
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2160
      Width           =   7575
   End
   Begin VB.CommandButton cmdNotAgree 
      Caption         =   "I Do Not Accept"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblAware 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEULA.frx":000C
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7530
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   7485
   End
End
Attribute VB_Name = "frmEULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmEULA.frm]

'
' License agreement form
'

Option Explicit

Private Sub Form_Load()
    
    Localize
    HighlightButton cmdAgree, 20

End Sub

Private Sub cmdAgree_Click()
    bAcceptEula = True
    'Me.Hide
    Unload Me
End Sub

Private Sub cmdNotAgree_Click()
    Unload Me
End Sub

Private Sub Localize()
    On Error GoTo ErrorHandler
    
    PreloadNativeLanguage
    
    ' Trend Micro HiJackThis - License Agreement
    'Me.Caption = TranslateNative(1092)
    
    ' Welcome to HiJackThis+
    lblWelcome.Caption = TranslateNative(1090)
    
    ' HiJackThis is free and open source program. It is provided "AS IS" without warranty of any kind. You may use this software at your own risk.
    ' This software is not permitted for commercial purposes.
    ' You need to read and accept license agreement to continue:
    lblAware.Caption = TranslateNative(1091)
    
    ' I Accept
    cmdAgree.Caption = TranslateNative(1093)
    
    ' I Do Not Accept
    cmdNotAgree.Caption = TranslateNative(1094)
    
    ' UELA text
    txtEULA.Text = TranslateNative(1095)
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Localize"
    If inIDE Then Stop: Resume Next
End Sub

Private Function GetEULA() As String
    Dim sText$
    sText = sText & "Program is licensed under GNU GENERAL PUBLIC LICENSE Version 2, June 1991 \n"
    sText = sText & "The source code is available at: https://github.com/dragokas/hijackthis"
    
    GetEULA = Replace$(sText, "\n", vbCrLf)
End Function

Private Function HighlightButton(btn As VB.CommandButton, thickness As Long)
    Dim box As VB.TextBox
    Set box = Controls.Add("VB.TextBox", "box1")
    With box
        .BorderStyle = 0
        .BackColor = vbGreen
        .Top = btn.Top - thickness
        .Left = btn.Left - thickness
        .Height = btn.Height + thickness * 2
        .Width = btn.Width + thickness * 2
        .Enabled = False
        .Visible = True
    End With
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then Cancel = 1
End Sub
