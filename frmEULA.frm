VERSION 5.00
Begin VB.Form frmEULA 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TrendMicro HijackThis"
   ClientHeight    =   7200
   ClientLeft      =   4785
   ClientTop       =   4830
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   6495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmEULAgpl.frx":0000
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "I Do Not Accept"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I Accept"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   6720
      Width           =   1575
   End
End
Attribute VB_Name = "frmEULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RegCreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis"
frmMain.Show
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

If RegKeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis") Then
''If True Then
    frmMain.Show
    Unload Me
End If
End Sub

