VERSION 5.00
Object = "{317589D1-37C8-47D9-B5B0-1C995741F353}#1.0#0"; "VBCCR17.OCX"
Begin VB.Form frmHostsMan 
   AutoRedraw      =   -1  'True
   Caption         =   "Hosts File Manager"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8595
   Icon            =   "frmHostsMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VBCCR17.FrameW fraHostsMan 
      Height          =   3900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VBCCR17.FrameW fraButtons 
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1508
         BorderStyle     =   0
         Begin VBCCR17.CommandButtonW cmdHostsManRefreshList 
            Height          =   420
            Left            =   6600
            TabIndex        =   2
            Top             =   360
            Width           =   1335
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Refresh list"
         End
         Begin VBCCR17.CommandButtonW cmdHostsManOpen 
            Height          =   420
            Left            =   3600
            TabIndex        =   3
            Top             =   360
            Width           =   1455
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Open in editor"
         End
         Begin VBCCR17.CommandButtonW cmdHostsManReset 
            Height          =   420
            Left            =   5160
            TabIndex        =   4
            Top             =   360
            Width           =   1335
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Reset"
         End
         Begin VBCCR17.CommandButtonW cmdHostsManToggle 
            Height          =   420
            Left            =   1800
            TabIndex        =   5
            Top             =   360
            Width           =   1695
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Toggle line(s)"
         End
         Begin VBCCR17.CommandButtonW cmdHostsManDel 
            Height          =   420
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1575
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Delete line(s)"
         End
         Begin VBCCR17.LabelW lblHostsTip2 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   45
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   450
            Caption         =   "Note: changes to the hosts file take effect when you restart your browser."
         End
      End
      Begin VBCCR17.ListBoxW lstHostsMan 
         Height          =   2340
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   8175
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         MouseTrack      =   -1  'True
         IntegralHeight  =   0   'False
         MultiSelect     =   2
      End
      Begin VBCCR17.LabelW lblHostsTip1 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   200
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   450
         Caption         =   "Hosts file is located at: C:\ ..."
      End
   End
End
Attribute VB_Name = "frmHostsMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Hosts -> Delete line
Private Sub cmdHostsManDel_Click()
    If lstHostsMan.ListIndex <> -1 And lstHostsMan.ListCount > 0 Then
        HostsDeleteLine lstHostsMan
    End If
End Sub

'Hosts -> Open in editor
Private Sub cmdHostsManOpen_Click()
    If FileExists(g_HostsFile) Then
        OpenInTextEditor g_HostsFile
    Else
         MsgBoxW Translate(281), vbExclamation '"No hosts file found."
    End If
End Sub

'Hosts -> Toggle line
Private Sub cmdHostsManToggle_Click()
    If lstHostsMan.ListIndex <> -1 And lstHostsMan.ListCount > 0 Then
        HostsToggleLine lstHostsMan
    End If
End Sub

'Hosts -> Reset
Private Sub cmdHostsManReset_Click()
    If HostsReset() Then
        ListHostsFile lstHostsMan
    End If
End Sub

'Hosts -> Refresh List
Private Sub cmdHostsManRefreshList_Click()
    ListHostsFile lstHostsMan
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
    ProcessHotkey KeyCode, Me
End Sub

Private Sub Form_Load()
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    ReloadLanguage True
    LoadWindowPos Me, SETTINGS_SECTION_HOSTSMAN

    If Not OSver.IsElevated Then
        cmdHostsManDel.Enabled = False
        cmdHostsManToggle.Enabled = False
    End If
    ListHostsFile lstHostsMan
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    SaveWindowPos Me, SETTINGS_SECTION_HOSTSMAN

    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Width < 8970 Then Me.Width = 8970
        If Me.Height < 4650 Then Me.Height = 4650
    End If
    fraHostsMan.Width = Me.ScaleWidth - 190
    fraHostsMan.Height = Me.ScaleHeight - 190
    lstHostsMan.Width = Me.ScaleWidth - 720
    lstHostsMan.Height = Me.ScaleHeight - 1750
    fraButtons.Top = fraHostsMan.Top + lstHostsMan.Top + lstHostsMan.Height - 100
End Sub

Private Sub lstHostsMan_MouseEnter()
    lstHostsMan.SetFocus 'to allow scrolling
End Sub
