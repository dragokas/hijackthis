VERSION 5.00
Begin VB.Form frmEULA 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Trend Micro HiJackThis - License Agreement"
   ClientHeight    =   7200
   ClientLeft      =   4785
   ClientTop       =   4830
   ClientWidth     =   6750
   Icon            =   "frmEULA.frx":0000
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmEULA.frx":0B3A
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
Option Explicit

Private Type tagINITCOMMONCONTROLSEX
    dwSize  As Long
    dwICC   As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Function SetCurrentProcessExplicitAppUserModelID Lib "shell32.dll" (ByVal pAppID As Long) As Long

Private ControlsEvent As New clsEvents


Private Sub Form_Initialize()
    On Error Resume Next
    Dim ICC As tagINITCOMMONCONTROLSEX
    Dim lr As Long
    Dim inf(68) As Long: inf(0) = 276: GetVersionEx inf(0)
    If inf(1) + inf(2) / 10 >= 6.1 Then ' Windows 7 and Later
        lr = SetCurrentProcessExplicitAppUserModelID(StrPtr("TrendMicro.HiJackThis"))
    End If
    
    ' Enable visual styles
    With ICC
        .dwSize = Len(ICC)
        .dwICC = &HFF& 'http://www.geoffchappell.com/studies/windows/shell/comctl32/api/commctrl/initcommoncontrolsex.htm
    End With
    InitCommonControlsEx ICC
End Sub

Private Sub Form_Load()
    If InStr(1, Command$, "/accepteula", 1) <> 0 Or _
        RegKeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", False) Or _
        RegKeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", True) Then
            EULA_Agree
            Me.Hide
            frmMain.Show
    Else
        Set ControlsEvent.txtBoxInArr = Text1   'focus on txtbox to add scrolling support
    End If
End Sub

Private Sub Command1_Click()
    EULA_Agree
    Me.Hide
    Set ControlsEvent = Nothing
    frmMain.Show
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Sub EULA_Agree()
    RegCreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", False
    RegCreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", True
End Sub
