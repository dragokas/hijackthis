VERSION 5.00
Object = "{317589D1-37C8-47D9-B5B0-1C995741F353}#1.0#0"; "VBCCR17.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   7380
   ClientLeft      =   4365
   ClientTop       =   1500
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   8835
   Begin VB.Timer tmrVTProgress 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   480
   End
   Begin VB.Timer tmrRunScan 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   6840
      Top             =   480
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Left            =   6120
      Top             =   480
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6480
      Top             =   480
   End
   Begin VB.PictureBox pictLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6720
      Picture         =   "frmMain.frx":4B2A
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   240
      Width           =   2460
   End
   Begin VBCCR17.CommandButtonW cmdHidden 
      Default         =   -1  'True
      Height          =   195
      Left            =   24960
      TabIndex        =   76
      Top             =   14760
      Width           =   75
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin VBCCR17.TextBoxW txtNothing 
      Height          =   285
      Left            =   1080
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BorderStyle     =   0
      Text            =   "frmMain.frx":966C
      Alignment       =   2
      Locked          =   -1  'True
   End
   Begin VBCCR17.FrameW fraHelp 
      Height          =   4095
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Help"
      Begin VBCCR17.CheckBoxW chkHelp 
         Height          =   375
         Index           =   4
         Left            =   5400
         TabIndex        =   140
         Top             =   240
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Fix"
         Style           =   1
      End
      Begin VBCCR17.CheckBoxW chkHelp 
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   83
         Top             =   240
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "History"
         Style           =   1
      End
      Begin VBCCR17.CheckBoxW chkHelp 
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   82
         Top             =   240
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Purpose"
         Style           =   1
      End
      Begin VBCCR17.CheckBoxW chkHelp 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   81
         Top             =   240
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Keys"
         Style           =   1
      End
      Begin VBCCR17.CheckBoxW chkHelp 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Caption         =   "Sections"
         Style           =   1
      End
      Begin VBCCR17.TextBoxW txtHelp 
         Height          =   3375
         Left            =   120
         TabIndex        =   128
         Top             =   600
         Width           =   5895
         _ExtentX        =   0
         _ExtentY        =   0
         Locked          =   -1  'True
         HideSelection   =   0   'False
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
   End
   Begin VBCCR17.ListBoxW lstResults 
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6135
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      IntegralHeight  =   0   'False
      Style           =   1
   End
   Begin VBCCR17.FrameW fraConfig 
      Height          =   5895
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VBCCR17.CheckBoxW chkConfigTabs 
         Height          =   555
         Index           =   3
         Left            =   4800
         TabIndex        =   11
         Top             =   180
         Width           =   1455
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Misc Tools"
         Style           =   1
      End
      Begin VBCCR17.CheckBoxW chkConfigTabs 
         Height          =   555
         Index           =   2
         Left            =   3360
         TabIndex        =   10
         Top             =   180
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Backups"
         Style           =   1
      End
      Begin VBCCR17.CheckBoxW chkConfigTabs 
         Height          =   555
         Index           =   1
         Left            =   1800
         TabIndex        =   9
         Top             =   180
         Width           =   1455
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ignorelist"
         Style           =   1
      End
      Begin VBCCR17.CheckBoxW chkConfigTabs 
         Height          =   555
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   180
         Width           =   1455
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Value           =   1
         Caption         =   "Settings"
         Style           =   1
      End
      Begin VBCCR17.FrameW fraConfigTabs 
         Height          =   4215
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "fraConfigBackup"
         Begin VBCCR17.CheckBoxW chkShowSRP 
            Height          =   375
            Left            =   120
            TabIndex        =   87
            Top             =   3840
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   661
            Caption         =   "Show System Restore Points"
         End
         Begin VBCCR17.CommandButtonW cmdConfigBackupCreateSRP 
            Height          =   720
            Left            =   7440
            TabIndex        =   86
            Top             =   3600
            Width           =   990
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Create restore point"
         End
         Begin VBCCR17.CommandButtonW cmdConfigBackupCreateRegBackup 
            Height          =   720
            Left            =   7440
            TabIndex        =   85
            Top             =   2760
            Width           =   990
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Create registry backup"
         End
         Begin VBCCR17.CommandButtonW cmdConfigBackupDeleteAll 
            Height          =   495
            Left            =   7440
            TabIndex        =   25
            Top             =   1920
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Delete all"
         End
         Begin VBCCR17.CommandButtonW cmdConfigBackupDelete 
            Height          =   495
            Left            =   7440
            TabIndex        =   24
            Top             =   1320
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Delete"
         End
         Begin VBCCR17.CommandButtonW cmdConfigBackupRestore 
            Height          =   495
            Left            =   7440
            TabIndex        =   20
            Top             =   720
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Restore"
         End
         Begin VBCCR17.ListBoxW lstBackups 
            Height          =   2385
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   7215
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            IntegralHeight  =   0   'False
            Style           =   1
         End
         Begin VBCCR17.LabelW lblBackupTip 
            Height          =   612
            Left            =   120
            TabIndex        =   36
            Top             =   0
            Width           =   8250
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   $"frmMain.frx":96C0
            WordWrap        =   -1  'True
         End
      End
      Begin VBCCR17.FrameW fraConfigTabs 
         Height          =   3135
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "fraConfigIgnorelist"
         Begin VBCCR17.CommandButtonW cmdConfigIgnoreDelSel 
            Height          =   495
            Left            =   7440
            TabIndex        =   22
            Top             =   480
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Remove"
         End
         Begin VBCCR17.CommandButtonW cmdConfigIgnoreDelAll 
            Height          =   495
            Left            =   7440
            TabIndex        =   23
            Top             =   1080
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Clear all"
         End
         Begin VBCCR17.ListBoxW lstIgnore 
            Height          =   2625
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   7215
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            IntegralHeight  =   0   'False
            Style           =   1
         End
         Begin VBCCR17.LabelW lblIgnoreTip 
            Height          =   252
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   7212
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "The following items will be ignored when scanning: "
            WordWrap        =   -1  'True
         End
      End
      Begin VBCCR17.FrameW fraHostsMan 
         Height          =   3735
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Hosts file manager"
         Begin VBCCR17.CommandButtonW cmdHostsManRefreshList 
            Height          =   425
            Left            =   6600
            TabIndex        =   141
            Top             =   3240
            Width           =   1335
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Refresh list"
         End
         Begin VBCCR17.CommandButtonW cmdHostsManOpen 
            Height          =   425
            Left            =   3600
            TabIndex        =   49
            Top             =   3240
            Width           =   1455
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Open in editor"
         End
         Begin VBCCR17.CommandButtonW cmdHostsManReset 
            Height          =   425
            Left            =   5160
            TabIndex        =   48
            Top             =   3240
            Width           =   1335
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Reset"
         End
         Begin VBCCR17.CommandButtonW cmdHostsManToggle 
            Height          =   425
            Left            =   1800
            TabIndex        =   47
            Top             =   3240
            Width           =   1695
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Toggle line(s)"
         End
         Begin VBCCR17.CommandButtonW cmdHostsManDel 
            Height          =   425
            Left            =   120
            TabIndex        =   46
            Top             =   3240
            Width           =   1575
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Delete line(s)"
         End
         Begin VBCCR17.ListBoxW lstHostsMan 
            Height          =   2340
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   8175
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            IntegralHeight  =   0   'False
            MultiSelect     =   2
         End
         Begin VBCCR17.LabelW lblHostsTip2 
            Height          =   252
            Left            =   120
            TabIndex        =   37
            Top             =   3000
            Width           =   8052
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Note: changes to the hosts file take effect when you restart your browser."
         End
         Begin VBCCR17.LabelW lblHostsTip1 
            Height          =   252
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   8052
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Hosts file is located at: C:\ ..."
         End
      End
      Begin VBCCR17.FrameW fraConfigTabs 
         Height          =   9120
         Index           =   3
         Left            =   120
         TabIndex        =   41
         Top             =   -4080
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   0
         _ExtentY        =   0
         BorderStyle     =   0
         Begin VB.VScrollBar vscMiscTools 
            Height          =   4095
            LargeChange     =   20
            Left            =   7680
            Max             =   100
            SmallChange     =   20
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
         Begin VBCCR17.FrameW fraMiscToolsScroll 
            Height          =   12015
            Left            =   0
            TabIndex        =   50
            Top             =   2000
            Width           =   7695
            _ExtentX        =   0
            _ExtentY        =   0
            BorderStyle     =   0
            Begin VBCCR17.FrameW FraRemoveHJT 
               Height          =   855
               Left            =   120
               TabIndex        =   129
               Top             =   10440
               Width           =   7335
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   0
               Caption         =   "Uninstall"
               Begin VBCCR17.CommandButtonW cmdUninstall 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   130
                  Top             =   360
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Uninstall HiJackThis"
               End
               Begin VBCCR17.LabelW lblUninstallHJT 
                  Height          =   444
                  Left            =   2640
                  TabIndex        =   63
                  Top             =   348
                  Width           =   4548
                  _ExtentX        =   0
                  _ExtentY        =   0
                  ForeColor       =   255
                  BackStyle       =   0
                  Caption         =   "Remove all HiJackThis Registry entries, backups and quit"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
            End
            Begin VBCCR17.FrameW FraPlugins 
               Height          =   1455
               Left            =   120
               TabIndex        =   122
               Top             =   6240
               Width           =   7335
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
               Caption         =   "Plugins"
               Begin VBCCR17.CommandButtonW cmdLnkCleaner 
                  Height          =   480
                  Left            =   120
                  TabIndex        =   124
                  Top             =   840
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "ClearLNK"
               End
               Begin VBCCR17.CommandButtonW cmdLnkChecker 
                  Height          =   480
                  Left            =   120
                  TabIndex        =   123
                  Top             =   240
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Check Browsers' LNK"
               End
               Begin VBCCR17.LabelW lblLnkCleanerAbout 
                  Height          =   615
                  Left            =   2520
                  TabIndex        =   126
                  Top             =   800
                  Width           =   4650
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "Clean and restore list of infected shortcuts (.LNK), found via Check Browsers' LNK plugin"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
               Begin VBCCR17.LabelW lblLnkCheckerAbout 
                  Height          =   390
                  Left            =   2520
                  TabIndex        =   125
                  Top             =   230
                  Width           =   4650
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "Identify infected shortcuts (.LNK) which cause unwanted advertisement in browsers"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
            End
            Begin VBCCR17.FrameW FraSysTools 
               Height          =   4695
               Left            =   120
               TabIndex        =   105
               Top             =   1440
               Width           =   7335
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
               Caption         =   "System tools"
               Begin VBCCR17.CommandButtonW cmdDigiSigChecker 
                  Height          =   480
                  Left            =   120
                  TabIndex        =   120
                  Top             =   4080
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Digital signature checker"
               End
               Begin VBCCR17.CommandButtonW cmdRegKeyUnlocker 
                  Height          =   480
                  Left            =   120
                  TabIndex        =   117
                  Top             =   3480
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Registry Key Unlocker"
               End
               Begin VBCCR17.CommandButtonW cmdARSMan 
                  Height          =   480
                  Left            =   120
                  TabIndex        =   116
                  Top             =   2880
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Uninstall Manager..."
               End
               Begin VBCCR17.CommandButtonW cmdADSSpy 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   113
                  Top             =   2400
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "ADS Spy..."
               End
               Begin VBCCR17.CommandButtonW cmdDeleteService 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   112
                  Top             =   1920
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Delete a Windows service..."
               End
               Begin VBCCR17.CommandButtonW cmdDelOnReboot 
                  Height          =   480
                  Left            =   120
                  TabIndex        =   109
                  Top             =   1320
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Delete a file on reboot..."
               End
               Begin VBCCR17.CommandButtonW cmdHostsManager 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   108
                  Top             =   840
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Hosts file manager"
               End
               Begin VBCCR17.CommandButtonW cmdProcessManager 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   106
                  Top             =   360
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Process manager"
               End
               Begin VBCCR17.LabelW lblDigiSigCheckerAbout 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   121
                  Top             =   4120
                  Width           =   4650
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "Verify authenticode digital signatures on the given list of files"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
               Begin VBCCR17.LabelW lblRegKeyUnlockerAbout 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   119
                  Top             =   3540
                  Width           =   4650
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "Reset permissions on the given registry keys list"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
               Begin VBCCR17.LabelW lblARSManAbout 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   118
                  Top             =   2930
                  Width           =   4410
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "Managing items in the Add/Remove Software list"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
               Begin VBCCR17.LabelW lblADSSpyAbout 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   115
                  Top             =   2460
                  Width           =   4665
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "Scan for hidden data streams"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
               Begin VBCCR17.LabelW lblDeleteServiceAbout 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   114
                  Top             =   1900
                  Width           =   4660
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "Delete a Windows Service (O23). USE WITH CAUTION!"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
               Begin VBCCR17.LabelW lblHostsManagerAbout 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   110
                  Top             =   900
                  Width           =   4650
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "Editor for the 'hosts' file"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
               Begin VBCCR17.LabelW lblProcessManagerAbout 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   107
                  Top             =   360
                  Width           =   4320
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Small process manager, working much like the Task Manager"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
               Begin VBCCR17.LabelW lblDelOnRebootAbout 
                  Height          =   390
                  Left            =   2520
                  TabIndex        =   111
                  Top             =   1320
                  Width           =   4695
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   "If a file cannot be removed from the disk, Windows can be setup to delete it when the system is restarted"
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
            End
            Begin VBCCR17.FrameW FraStartupList 
               Height          =   1335
               Left            =   120
               TabIndex        =   102
               Top             =   0
               Width           =   7335
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "StartupList"
               Begin VBCCR17.CommandButtonW cmdStartupList 
                  Height          =   465
                  Left            =   120
                  TabIndex        =   103
                  Top             =   480
                  Width           =   2295
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "StartupList scan"
               End
               Begin VBCCR17.LabelW lblStartupListAbout 
                  Height          =   1032
                  Left            =   2520
                  TabIndex        =   104
                  Top             =   240
                  Width           =   4632
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackStyle       =   0
                  Caption         =   $"frmMain.frx":97A5
                  AutoSize        =   -1  'True
                  WordWrap        =   -1  'True
               End
            End
            Begin VBCCR17.FrameW FraUpdateCheck 
               Height          =   2532
               Left            =   120
               TabIndex        =   98
               Top             =   7800
               Width           =   7335
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
               Caption         =   "Update check"
               Begin VBCCR17.CheckBoxW chkUpdateSilently 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   138
                  Top             =   1080
                  Width           =   4695
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Update in silent mode"
               End
               Begin VBCCR17.CheckBoxW chkUpdateToTest 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   137
                  Top             =   740
                  Width           =   4575
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Update to test versions"
               End
               Begin VBCCR17.CheckBoxW chkCheckUpdatesOnStart 
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   136
                  Top             =   390
                  Width           =   4695
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Check updates automatically on program startup"
               End
               Begin VBCCR17.OptionButtonW OptProxyDirect 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   64
                  Top             =   960
                  Width           =   2175
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Direct connection"
               End
               Begin VBCCR17.CheckBoxW chkSocks4 
                  Height          =   195
                  Left            =   240
                  TabIndex        =   65
                  Top             =   2160
                  Width           =   855
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
                  Caption         =   "Socks4"
               End
               Begin VBCCR17.OptionButtonW optProxyManual 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   66
                  Top             =   1440
                  Width           =   2172
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Proxy"
               End
               Begin VBCCR17.OptionButtonW optProxyIE 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   67
                  ToolTipText     =   "(inetcpl.cpl)"
                  Top             =   1200
                  Width           =   2172
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Value           =   -1  'True
                  Caption         =   "IE settings"
               End
               Begin VBCCR17.TextBoxW txtUpdateProxyPass 
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   68
                  Top             =   1800
                  Width           =   1455
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
                  PasswordChar    =   42
                  IMEMode         =   3
               End
               Begin VBCCR17.TextBoxW txtUpdateProxyLogin 
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   69
                  Top             =   1800
                  Width           =   1335
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
               End
               Begin VBCCR17.CheckBoxW chkUpdateUseProxyAuth 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   70
                  Top             =   1800
                  Width           =   2172
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Use authorization"
               End
               Begin VBCCR17.TextBoxW txtUpdateProxyPort 
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   71
                  Top             =   1440
                  Width           =   1455
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
                  Text            =   "frmMain.frx":986D
               End
               Begin VBCCR17.TextBoxW txtUpdateProxyHost 
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   75
                  Top             =   1440
                  Width           =   1335
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
                  Text            =   "frmMain.frx":9895
               End
               Begin VBCCR17.CommandButtonW cmdCheckUpdate 
                  Height          =   480
                  Left            =   240
                  TabIndex        =   99
                  Top             =   360
                  Width           =   2055
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Check for update online"
               End
               Begin VBCCR17.LabelW lblUpdatePass 
                  Height          =   252
                  Left            =   4800
                  TabIndex        =   7
                  Top             =   1824
                  Width           =   852
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
                  Caption         =   "Password"
               End
               Begin VBCCR17.LabelW lblUpdateLogin 
                  Height          =   252
                  Left            =   2520
                  TabIndex        =   100
                  Top             =   1800
                  Width           =   612
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
                  Caption         =   "Login"
               End
               Begin VBCCR17.LabelW lblUpdatePort 
                  Height          =   252
                  Left            =   4800
                  TabIndex        =   101
                  Top             =   1476
                  Width           =   612
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
                  Caption         =   "Port"
               End
               Begin VBCCR17.LabelW lblUpdateServer 
                  Height          =   252
                  Left            =   2520
                  TabIndex        =   127
                  Top             =   1476
                  Width           =   612
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Enabled         =   0   'False
                  Caption         =   "Server"
               End
            End
            Begin VBCCR17.FrameW FraTestStaff 
               Height          =   855
               Left            =   120
               TabIndex        =   96
               Top             =   11280
               Visible         =   0   'False
               Width           =   7335
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Testing staff"
               Begin VBCCR17.CommandButtonW cmdTaskScheduler 
                  Height          =   345
                  Left            =   240
                  TabIndex        =   97
                  Top             =   360
                  Width           =   2055
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Task Scheduler Log"
               End
            End
         End
      End
      Begin VBCCR17.FrameW fraConfigTabs 
         Height          =   4250
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   8440
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "fraConfigMain"
         Begin VB.VScrollBar vscSettings 
            Height          =   4160
            LargeChange     =   20
            Left            =   8040
            Max             =   100
            TabIndex        =   59
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VBCCR17.FrameW fraConfigTabsNested 
            Height          =   7815
            Left            =   0
            TabIndex        =   60
            Top             =   -120
            Width           =   8055
            _ExtentX        =   0
            _ExtentY        =   0
            BorderStyle     =   0
            Begin VBCCR17.FrameW FraInterface 
               Height          =   1800
               Left            =   0
               TabIndex        =   61
               Top             =   3120
               Width           =   7935
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Interface"
               Begin VBCCR17.CheckBoxW chkFontBold 
                  Height          =   280
                  Left            =   3050
                  TabIndex        =   139
                  Top             =   1380
                  Width           =   280
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   7.5
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "B"
                  Style           =   1
               End
               Begin VBCCR17.CheckBoxW chkFontWholeInterface 
                  Height          =   255
                  Left            =   3480
                  TabIndex        =   135
                  Top             =   1400
                  Width           =   4332
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Apply selected font on whole interface"
               End
               Begin VBCCR17.ComboBoxW cmbFontSize 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   134
                  Top             =   1380
                  Width           =   735
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Style           =   2
               End
               Begin VBCCR17.ComboBoxW cmbFont 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   131
                  Top             =   1380
                  Width           =   2055
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Style           =   2
               End
               Begin VBCCR17.CheckBoxW chkConfigMinimizeToTray 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   79
                  Top             =   840
                  Width           =   6015
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Minimize program to system tray when clicking _ button"
               End
               Begin VBCCR17.CheckBoxW chkSkipErrorMsg 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   78
                  Top             =   600
                  Width           =   4695
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Do not show error messages"
               End
               Begin VBCCR17.CheckBoxW chkSkipIntroFrameSettings 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   58
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   4575
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Do not show main menu at startup"
               End
               Begin VBCCR17.LabelW lblFontSize 
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   133
                  Top             =   1140
                  Width           =   975
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Size"
               End
               Begin VBCCR17.LabelW lblFont 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   132
                  Top             =   1140
                  Width           =   1935
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Font"
               End
            End
            Begin VBCCR17.FrameW FraIncludeSections 
               Height          =   1575
               Left            =   0
               TabIndex        =   88
               Top             =   120
               Width           =   3372
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Scan area"
               Begin VBCCR17.CheckBoxW chkAdditionalScan 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   92
                  ToolTipText     =   "Include specific sections, like O4 - RenameOperations, O21 - Column Hanlders / Context menu, O23 - Drivers e.t.c."
                  Top             =   1080
                  Width           =   3015
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Additional scan"
               End
               Begin VBCCR17.CheckBoxW chkAdvLogEnvVars 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   91
                  ToolTipText     =   "Include environment variables in logfile"
                  Top             =   720
                  Width           =   3015
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Environment variables"
               End
               Begin VBCCR17.CheckBoxW chkLogProcesses 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   90
                  ToolTipText     =   "Include list of running processes in logfiles"
                  Top             =   360
                  Width           =   3015
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Value           =   1
                  Caption         =   "Processes"
               End
            End
            Begin VBCCR17.FrameW FraFixing 
               Height          =   1215
               Left            =   0
               TabIndex        =   62
               Top             =   1800
               Width           =   7935
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Fix && Backup"
               Begin VBCCR17.TextBoxW txtDefStartPage 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   15
                  Top             =   1560
                  Width           =   5175
                  _ExtentX        =   0
                  _ExtentY        =   0
               End
               Begin VBCCR17.TextBoxW txtDefSearchPage 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   16
                  Top             =   1920
                  Width           =   5175
                  _ExtentX        =   0
                  _ExtentY        =   0
               End
               Begin VBCCR17.TextBoxW txtDefSearchAss 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   17
                  Top             =   2280
                  Width           =   5175
                  _ExtentX        =   0
                  _ExtentY        =   0
               End
               Begin VBCCR17.TextBoxW txtDefSearchCust 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   18
                  Top             =   2640
                  Width           =   5175
                  _ExtentX        =   0
                  _ExtentY        =   0
               End
               Begin VBCCR17.CheckBoxW chkConfirm 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   14
                  Top             =   600
                  Width           =   7455
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Value           =   1
                  Caption         =   "Confirm fixing && ignoring of items (safe mode)"
               End
               Begin VBCCR17.CheckBoxW chkBackup 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   13
                  Top             =   360
                  Width           =   7335
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Value           =   1
                  Caption         =   "Make backups before fixing items"
               End
               Begin VBCCR17.CheckBoxW chkAutoMark 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   12
                  Top             =   840
                  Width           =   7335
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Mark everything found for fixing after scan (DANGEROUS !!!)"
               End
            End
            Begin VBCCR17.FrameW fraScanOpt 
               Height          =   1575
               Left            =   3480
               TabIndex        =   89
               Top             =   120
               Width           =   4452
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Scan options"
               Begin VBCCR17.ComboBoxW cmbHashType 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   40
                  TabStop         =   0   'False
                  Top             =   840
                  Width           =   1212
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Style           =   2
               End
               Begin VBCCR17.CheckBoxW chkConfigStartupScan 
                  Height          =   270
                  Left            =   120
                  TabIndex        =   72
                  ToolTipText     =   "Run HiJackThis scan at Windows startup and show results (if only items are found)"
                  Top             =   1120
                  Width           =   3972
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Add HiJackThis to startup"
               End
               Begin VBCCR17.CheckBoxW chkDoCheckSum 
                  Height          =   195
                  Left            =   120
                  TabIndex        =   95
                  ToolTipText     =   "Calculate checksum of files if possible"
                  Top             =   900
                  Width           =   2892
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Calculate Checksum"
               End
               Begin VBCCR17.CheckBoxW chkIgnoreAll 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   94
                  ToolTipText     =   "Include in log any entries regardless whitelist"
                  Top             =   610
                  Width           =   3972
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Caption         =   "Ignore ALL Whitelists"
               End
               Begin VBCCR17.CheckBoxW chkIgnoreMicrosoft 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   93
                  ToolTipText     =   "Do not include in log files and registry related to Microsoft"
                  Top             =   360
                  Width           =   3972
                  _ExtentX        =   0
                  _ExtentY        =   0
                  Value           =   1
                  Caption         =   "Hide Microsoft entries"
               End
            End
         End
      End
   End
   Begin VBCCR17.FrameW fraScan 
      Height          =   1455
      Left            =   120
      TabIndex        =   30
      Top             =   5880
      Width           =   2775
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "   Scan && fix stuff"
      Begin VBCCR17.CommandButtonW CmdHidden2 
         Height          =   425
         Left            =   240
         TabIndex        =   77
         Top             =   1560
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Focus"
      End
      Begin VBCCR17.CommandButtonW cmdInfo 
         Height          =   450
         Left            =   240
         TabIndex        =   3
         Top             =   850
         Width           =   2340
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Info on selected item..."
      End
      Begin VBCCR17.CommandButtonW cmdScan 
         Height          =   450
         Left            =   240
         TabIndex        =   1
         Tag             =   "1"
         Top             =   300
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Scan"
      End
      Begin VBCCR17.CommandButtonW cmdFix 
         Height          =   450
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   1140
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         Caption         =   "Fix checked"
      End
   End
   Begin VBCCR17.FrameW fraSubmit 
      Height          =   1455
      Left            =   3000
      TabIndex        =   51
      Top             =   5880
      Width           =   2885
      _ExtentX        =   0
      _ExtentY        =   0
      Begin VBCCR17.CommandButtonW cmdAnalyze 
         Height          =   450
         Left            =   480
         TabIndex        =   52
         Top             =   300
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         Caption         =   "Analyze report"
      End
      Begin VBCCR17.CommandButtonW cmdMainMenu 
         Height          =   450
         Left            =   720
         TabIndex        =   54
         Top             =   850
         Width           =   1455
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Main Menu"
      End
   End
   Begin VBCCR17.FrameW fraOther 
      Height          =   1455
      Left            =   6000
      TabIndex        =   31
      Top             =   5880
      Width           =   2775
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "   Other stuff"
      Begin VBCCR17.CommandButtonW cmdSaveDef 
         Height          =   450
         Left            =   120
         TabIndex        =   6
         Top             =   850
         Width           =   2532
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         Caption         =   "Add checked to ignorelist"
      End
      Begin VBCCR17.CommandButtonW cmdConfig 
         Height          =   450
         Left            =   1320
         TabIndex        =   5
         Tag             =   "0"
         Top             =   300
         Width           =   1332
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Settings"
      End
      Begin VBCCR17.CommandButtonW cmdHelp 
         Height          =   450
         Left            =   120
         TabIndex        =   4
         Tag             =   "0"
         Top             =   300
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Help"
      End
   End
   Begin VBCCR17.FrameW fraN00b 
      Height          =   6255
      Left            =   120
      TabIndex        =   55
      Top             =   960
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Caption         =   "Main menu"
      Begin VBCCR17.CommandButtonW cmdN00bHJTQuickStart 
         Height          =   495
         Left            =   480
         TabIndex        =   147
         Top             =   4200
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Tutorial && Support"
      End
      Begin VBCCR17.CommandButtonW cmdSettings 
         Height          =   495
         Left            =   480
         TabIndex        =   146
         Top             =   3480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Settings"
      End
      Begin VBCCR17.CommandButtonW cmdN00bBackups 
         Height          =   495
         Left            =   480
         TabIndex        =   145
         Top             =   2880
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Backups"
      End
      Begin VBCCR17.CommandButtonW cmdFixing 
         Height          =   495
         Left            =   480
         TabIndex        =   144
         Top             =   2160
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Fixing"
      End
      Begin VBCCR17.CommandButtonW cmdN00bScan 
         Height          =   495
         Left            =   480
         TabIndex        =   143
         Top             =   1440
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Do a system scan only"
      End
      Begin VBCCR17.CommandButtonW cmdN00bLog 
         Height          =   495
         Left            =   480
         TabIndex        =   142
         Top             =   840
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Do a system scan and save a logfile"
      End
      Begin VBCCR17.ComboBoxW cboN00bLanguage 
         Height          =   315
         Left            =   6120
         TabIndex        =   74
         Top             =   720
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         Style           =   2
      End
      Begin VBCCR17.CheckBoxW chkSkipIntroFrame 
         Height          =   255
         Left            =   360
         TabIndex        =   57
         Top             =   5520
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Do not show this menu after starting the program"
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   480
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VBCCR17.LabelW lblInfo 
         Height          =   195
         Index           =   9
         Left            =   6480
         TabIndex        =   73
         Top             =   360
         Width           =   1320
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Change language:"
         AutoSize        =   -1  'True
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   10
         X1              =   480
         X2              =   4200
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   8
         X1              =   480
         X2              =   4560
         Y1              =   2760
         Y2              =   2760
      End
      Begin VBCCR17.LabelW lblInfo 
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   56
         Top             =   480
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "What would you like to do?"
         AutoSize        =   -1  'True
      End
   End
   Begin VBCCR17.LabelW lblInfo 
      Height          =   852
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   36
      Width           =   6492
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   $"frmMain.frx":98C7
      WordWrap        =   -1  'True
   End
   Begin VBCCR17.LabelW lblMD5 
      Height          =   195
      Left            =   600
      TabIndex        =   42
      Top             =   690
      Visible         =   0   'False
      Width           =   8275
      _ExtentX        =   0
      _ExtentY        =   0
      ForeColor       =   16711680
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Calculating MD5 checksum of [file]..."
   End
   Begin VB.Shape shpMD5Progress 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   120
      Left            =   120
      Tag             =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VBCCR17.LabelW lblStatus 
      Height          =   195
      Left            =   400
      TabIndex        =   43
      Top             =   330
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   0
      _ExtentY        =   0
      ForeColor       =   65535
      BackStyle       =   0
      Caption         =   "Status"
      AutoSize        =   -1  'True
   End
   Begin VB.Shape shpProgress 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Top             =   240
      Visible         =   0   'False
      Width           =   8275
   End
   Begin VB.Shape shpMD5Background 
      BackStyle       =   1  'Opaque
      Height          =   120
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   8275
   End
   Begin VBCCR17.LabelW lblInfo 
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   45
      Visible         =   0   'False
      Width           =   8500
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   $"frmMain.frx":99B9
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuFileInstallHJT 
         Caption         =   "Install HJT"
      End
      Begin VB.Menu mnuFileUninstHJT 
         Caption         =   "Uninstall HJT"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsReg 
         Caption         =   "Registry"
         Begin VB.Menu mnuToolsRegUnlockKey 
            Caption         =   "Key Unlocker"
         End
         Begin VB.Menu mnuToolsRegTypeChecker 
            Caption         =   "Key Type Analyzer"
         End
      End
      Begin VB.Menu mnuToolsFiles 
         Caption         =   "Files"
         Begin VB.Menu mnuToolsDigiSign 
            Caption         =   "Digital Signature checker"
         End
         Begin VB.Menu mnuToolsADSSpy 
            Caption         =   "Alternative Data Streams Spy"
         End
         Begin VB.Menu mnuToolsHosts 
            Caption         =   "Hosts file Manager"
         End
         Begin VB.Menu mnuToolsUnlockFiles 
            Caption         =   "Unlock File / Folder"
         End
         Begin VB.Menu mnuToolsDelFileOnReboot 
            Caption         =   "Plan to Delete File on Reboot..."
         End
      End
      Begin VB.Menu mnuToolsService 
         Caption         =   "Services"
         Begin VB.Menu mnuToolsDelServ 
            Caption         =   "Delete Service"
         End
      End
      Begin VB.Menu mnuToolsShortcuts 
         Caption         =   "Shortcuts"
         Begin VB.Menu mnuToolsShortcutsChecker 
            Caption         =   "Check Browsers' LNK"
         End
         Begin VB.Menu mnuToolsShortcutsFixer 
            Caption         =   "ClearLNK"
         End
      End
      Begin VB.Menu mnuToolsUninst 
         Caption         =   "Uninstall Manager"
      End
      Begin VB.Menu mnuToolsProcMan 
         Caption         =   "Process Manager"
      End
      Begin VB.Menu mnuToolsStartupList 
         Caption         =   "StartupList"
      End
   End
   Begin VB.Menu mnuBasicManual 
      Caption         =   "Basic manual 2"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpManualRussian 
         Caption         =   "Russian"
      End
      Begin VB.Menu mnuHelpManualEnglish 
         Caption         =   "English (outdated)"
      End
      Begin VB.Menu mnuHelpManualFrench 
         Caption         =   "French (outdated)"
      End
      Begin VB.Menu mnuHelpManualGerman 
         Caption         =   "German (outdated)"
      End
      Begin VB.Menu mnuHelpManualSpanish 
         Caption         =   "Spanish (outdated)"
      End
      Begin VB.Menu mnuHelpManualPortuguese 
         Caption         =   "Portuguese (outdated)"
      End
      Begin VB.Menu mnuHelpManualDutch 
         Caption         =   "Dutch (outdated)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "User's Manual"
         Begin VB.Menu mnuHelpManualBasic 
            Caption         =   "Basic manual"
         End
         Begin VB.Menu mnuHelpManualSections 
            Caption         =   "Sections' description"
         End
         Begin VB.Menu mnuHelpManualCmdKeys 
            Caption         =   "Command line keys"
         End
      End
      Begin VB.Menu mnuHelpReportBug 
         Caption         =   "Report a bug"
      End
      Begin VB.Menu mnuHelpUpdate 
         Caption         =   "Check for updates"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About HJT"
      End
   End
   Begin VB.Menu mnuResultList 
      Caption         =   "Result List Popup"
      Begin VB.Menu mnuResultFix 
         Caption         =   "Fix checked"
      End
      Begin VB.Menu mnuResultAddToIgnore 
         Caption         =   "Add to ignore list"
      End
      Begin VB.Menu mnuResultAddALLToIgnore 
         Caption         =   "Add ALL to ignore list"
      End
      Begin VB.Menu mnuResultDisable 
         Caption         =   "Disable"
      End
      Begin VB.Menu mnuResultDelim1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResultCopy 
         Caption         =   "Copy"
         Begin VB.Menu mnuResultCopyLine 
            Caption         =   "Whole Line"
         End
         Begin VB.Menu mnuResultCopyRegKey 
            Caption         =   "Registry Key"
         End
         Begin VB.Menu mnuResultCopyRegParam 
            Caption         =   "Registry Parameter"
         End
         Begin VB.Menu mnuResultCopyFilePath 
            Caption         =   "File Path"
         End
         Begin VB.Menu mnuResultCopyFileName 
            Caption         =   "File Name"
         End
         Begin VB.Menu mnuResultCopyFileArguments 
            Caption         =   "File Arguments"
         End
         Begin VB.Menu mnuResultCopyFileObject 
            Caption         =   "File (as Object)"
         End
         Begin VB.Menu mnuResultCopyFileHash 
            Caption         =   "File Hash"
            Begin VB.Menu mnuResultCopyFileHashMD5 
               Caption         =   "MD5"
            End
            Begin VB.Menu mnuResultCopyFileHashSHA1 
               Caption         =   "SHA1"
            End
            Begin VB.Menu mnuResultCopyFileHashSHA256 
               Caption         =   "SHA256"
            End
         End
         Begin VB.Menu mnuResultCopyValue 
            Caption         =   "Value"
         End
      End
      Begin VB.Menu mnuResultVT 
         Caption         =   "VirusTotal"
         Begin VB.Menu mnuResultVTHash 
            Caption         =   "Search by Hash"
         End
         Begin VB.Menu mnuResultVTSubmit 
            Caption         =   "Submit with 'Autoruns'"
         End
      End
      Begin VB.Menu mnuResultInfo 
         Caption         =   "Info on selected"
      End
      Begin VB.Menu mnuResultSearch 
         Caption         =   "Search on Google"
      End
      Begin VB.Menu mnuResultJump 
         Caption         =   "Jump to Registry / File"
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry1"
            Index           =   0
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry2"
            Index           =   1
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry3"
            Index           =   2
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry4"
            Index           =   3
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry5"
            Index           =   4
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry6"
            Index           =   5
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry7"
            Index           =   6
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry8"
            Index           =   7
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry9"
            Index           =   8
         End
         Begin VB.Menu mnuResultJumpReg 
            Caption         =   "Reg.Entry10"
            Index           =   9
         End
         Begin VB.Menu mnuResultJumpDelim 
            Caption         =   "-"
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry1"
            Index           =   0
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry2"
            Index           =   1
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry3"
            Index           =   2
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry4"
            Index           =   3
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry5"
            Index           =   4
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry6"
            Index           =   5
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry7"
            Index           =   6
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry8"
            Index           =   7
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry9"
            Index           =   8
         End
         Begin VB.Menu mnuResultJumpFile 
            Caption         =   "File.Entry10"
            Index           =   9
         End
      End
      Begin VB.Menu mnuResultDelim2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveReport 
         Caption         =   "Save Report..."
      End
      Begin VB.Menu mnuResultReScan 
         Caption         =   "ReScan"
      End
   End
   Begin VB.Menu mnuDelChoose 
      Caption         =   "Choosing file or folder to unlock/delete"
      Visible         =   0   'False
      Begin VB.Menu mnuDelChooseFile 
         Caption         =   "Choose File..."
      End
      Begin VB.Menu mnuDelChooseFolder 
         Caption         =   "Choose Folder..."
      End
   End
   Begin VB.Menu mnuSupportContext 
      Caption         =   "Support popup"
      Begin VB.Menu mnuSupportOnline 
         Caption         =   "Online tutorial"
      End
      Begin VB.Menu mnuSupportOffline 
         Caption         =   "Offline tutorial"
      End
      Begin VB.Menu mnuSupportCure 
         Caption         =   "Cure forums"
         Begin VB.Menu mnuSupportCureGithub 
            Caption         =   "GitHub (EN)"
         End
         Begin VB.Menu mnuSupportSafezone 
            Caption         =   "SafeZone (RU)"
         End
         Begin VB.Menu mnuSupportKaspersky 
            Caption         =   "KasClub (RU)"
         End
         Begin VB.Menu mnuSupportOszone 
            Caption         =   "OSZone (RU)"
         End
         Begin VB.Menu mnuSupportSoftboard 
            Caption         =   "SoftBoard (RU)"
         End
         Begin VB.Menu mnuSupportTHG 
            Caption         =   "THG (RU)"
         End
         Begin VB.Menu mnuSupportCyberforum 
            Caption         =   "CyberForum (RU)"
         End
         Begin VB.Menu mnuSupportVirusinfo 
            Caption         =   "VI (RU)"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmMain.frm]

'
' HJT Main form
'

' Call stack note:
'
' "Do a system scan and save log file" button calls:
'    -> cmdN00bLog_Click -> cmdScan_Click -> StartScan -> HJT_SaveReport -> CreateLogFile (process list)
'
' App key:     HKLM\Software\HijackThis+
' Old App key: HKLM\Software\TrendMicro\HiJackThisFork

Option Explicit

#Const SCRIPT_FIX = False

Private Const HJT_ALPHA             As Boolean = True
Private Const HJT_BETA              As Boolean = False

Private Const ADS_SPY_VERSION       As String = "1.14"
Private Const STARTUP_LIST_VERSION  As String = "2.13"
Private Const PROC_MAN_VERSION      As String = "1.07"
Private Const UNINST_MAN_VERSION    As String = "2.1"

Private Const MAX_JUMP_LIST_ITEMS   As Long = 10

Private ControlsEvent() As New clsEvents
Private WithEvents FormSys As frmSysTray
Attribute FormSys.VB_VarHelpID = -1

Private bSwitchingTabs  As Boolean
Private bIsBeta         As Boolean
Private bIsAlpha        As Boolean
Private lToolsHeight    As Long
Private bLockResize     As Boolean
Private iPrevListIndex  As Long

Private JumpFileCache() As FIX_FILE
Private JumpRegCache()  As FIX_REG_KEY


Public Sub Test()
    On Error GoTo ErrorHandler:
    
    'If you need something to test after program started and initialized all required variables, please use this sub.
    
'    If inIDE Then
'        Dim SignResult As SignResult_TYPE, sFile As String: sFile = "C:\Windows\system32\DRIVERS\klgse.sys"
'        Debug.Print SignVerify(sFile, 0, SignResult)
'        'Debug.Print SignVerifyJack(sFile, SignResult)
'        'Debug.Print IsMicrosoftDriverFileEx(sFile, SignResult)
'        Debug.Print "Api err code: 0x" & Hex(SignResult.ApiErrorCode) & " - " & ErrMessageText(SignResult.ApiErrorCode)
'        Debug.Print "Signer: " & SignResult.SubjectName
'        Debug.Print "Email: " & SignResult.SubjectEmail
'        Debug.Print "HashFinalCert: " & SignResult.HashFinalCert
'    End If
    
    'DownloadAndUpdateSelf "https://dragokas.com/tools/HiJackThis.zip", False
    '// TODO
    'UnpackZIP = make recursive
    
    'LoadDatabase
    
    chkHelp(4).Visible = False
    cmdFixing.Visible = False
    
    #If SCRIPT_FIX Then
        Script.ExecuteFixFromFile "C:\Users\Alex\Desktop\script.txt"
    #End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.Test"
    If inIDE Then Stop: Resume Next
End Sub

' Tips on functions:

'1. Use AddWarning() to append text to the end of the log, before debugging info.

Private Sub Form_Load()
    Static bInit As Boolean
    
    g_HwndMain = Me.hWnd
    
    Perf.StartTime = GetTickCount()
    
    pvSetFormIcon Me

    If Not bAutoLogSilent Then
        ScalePictureDPI pictLogo
        Call PictureBoxRgn(pictLogo, RGB(255, 255, 255))
    End If
    
    If Not (OSver.IsElevated Or OSver.IsLocalSystemContext) Then
        cmdDelOnReboot.Enabled = False
        mnuToolsDelFileOnReboot.Enabled = False
    End If
    
    If Not OSver.IsElevated Then
        mnuFileInstallHJT.Enabled = False
        mnuToolsRegUnlockKey.Enabled = False
        mnuToolsRegTypeChecker.Enabled = False
        mnuToolsUnlockFiles.Enabled = False
        cmdHostsManDel.Enabled = False
        cmdHostsManToggle.Enabled = False
        mnuToolsDelServ.Enabled = False
        mnuToolsShortcutsChecker.Enabled = False
        mnuToolsShortcutsFixer.Enabled = False
        cmdDeleteService.Enabled = False
        cmdRegKeyUnlocker.Enabled = False
        cmdLnkChecker.Enabled = False
        cmdLnkCleaner.Enabled = False
        cmdConfigIgnoreDelSel.Enabled = False
        cmdConfigIgnoreDelAll.Enabled = False
        chkSkipIntroFrame.Enabled = False
        chkConfigStartupScan.Enabled = False
        chkSkipIntroFrameSettings.Enabled = False
        cmdConfigBackupCreateSRP.Enabled = False
        chkShowSRP.Enabled = False
        cmdConfigBackupCreateRegBackup.Enabled = False
    End If
    
    If bInit Then
        If Not bAutoLogSilent Then
            MsgBoxW "Critical error. Main form is initialized twice!"
            End
        End If
    Else
        bInit = True
        mnuResultList.Visible = False
        mnuSupportContext.Visible = False
        If gNoGUI Then Me.Hide
        
        cmbHashType.AddItem "MD5"
        cmbHashType.AddItem "SHA1"
        If OS_SupportSHA2() Then
            cmbHashType.AddItem "SHA256"
        Else
            mnuResultVT.Visible = False
        End If
        
        FormStart_Stage1
        If g_NeedTerminate Then
            Me.WindowState = vbMinimized
        End If
        tmrStart.Enabled = True
    End If

End Sub

Private Sub lstResults_ItemCheck(ByVal Item As Long)
    If bScanMode Then
        lstResults.ItemChecked(Item) = False 'forbid marking item during the scan
    End If
End Sub

Private Sub lstResults_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = VK_APPS Then ' context menu key
        lstResults_MouseUp 2, 0, -1, -1
    End If
End Sub

Private Sub tmrShutdown_Timer()
    If (GetTickCount() - Perf.StartTime) / 1000 > Perf.MAX_TimeOut Then
        HJT_Shutdown
    End If
End Sub

Private Sub tmrRunScan_Timer()
    tmrRunScan.Enabled = False
    cmdScan_Click
End Sub

Private Sub tmrStart_Timer()
    tmrStart.Enabled = False
    If Not gNoGUI Then Me.Show vbModeless
    FormStart_Stage2
End Sub

Private Sub tmrVTProgress_Timer()
    
    Static sPath As String
    
    If Len(sPath) = 0 Then
        sPath = GetAutorunsPath()
    End If
    
    If Not ProcessExist(sPath, True) Then
    
        tmrVTProgress.Enabled = False
        
        If Not g_bScanInProgress Or Not g_bCheckSum Then UpdateVTProgressbar True
        
        g_bVTScanInProgress = False
        g_bVTScanned = True
    End If
End Sub

Private Sub FormStart_Stage1()

    On Error GoTo ErrorHandler:
    
    Dim Ctl   As Control
    Dim ChkB  As VBCCR17.CheckBoxW
    Dim i     As Long
    Dim sCmdLine As String
    
    AppendErrorLogCustom "FormStart_Stage1 - Begin"
    
    If HJT_ALPHA Then bIsAlpha = True
    If HJT_BETA Then bIsBeta = True
    
    StartupListVer = STARTUP_LIST_VERSION
    ADSspyVer = ADS_SPY_VERSION
    ProcManVer = PROC_MAN_VERSION
    UninstManVer = UNINST_MAN_VERSION
    
    g_HJT_Items_Count = 38 'R + F + B + O1-...-O27 + Subsections (for progressbar)

    If inIDE Then DisableSubclassing = True
    
    If bAutoLogSilent Then 'timeout timer
        If Perf.MAX_TimeOut <> 0 Then
            tmrShutdown.Interval = 1000
            tmrShutdown.Enabled = True
        End If
    End If
    
    If (Not DisableSubclassing) And (Not bAutoLogSilent) Then
        SubClassScroll True
    End If
    
    AppVerPlusName = g_AppName & " (Plus) build " & GetOwnCompilationDate() & " " & _
        IIf(bIsAlpha, "Alpha", IIf(bIsBeta, "Beta", "Stable")) & " v." & AppVerString
    
    If Not bPolymorph Then
        SetWindowTitleText Me.hWnd, AppVerPlusName
    End If
    
    bFirstRebootScan = ScanAfterReboot(False)
    If bFirstRebootScan Then
        RegSaveHJT "RebootRequired", 0
    End If
    
    fraMiscToolsScroll.Height = 400 + FraTestStaff.Top
    
    'testing stuff
    If inIDE Or InStr(1, AppExeName(), "test", 1) <> 0 Or bDebugMode Then
        fraMiscToolsScroll.Height = fraMiscToolsScroll.Height + FraTestStaff.Height
        'Task scheduler jobs log on 'misc section'.
        Me.FraTestStaff.Visible = True
        'cmdTaskScheduler.Visible = True
        lToolsHeight = 0
        'added autoadjustment depending on the top of the most bottom frame
        lToolsHeight = 850 - (FraRemoveHJT.Top + FraRemoveHJT.Height - (9600 + 855))
    Else
        lToolsHeight = 850 - (FraTestStaff.Top - 10560)
    End If
    
    MenuIcons_Initialize Me
    LoadLanguageList
    LoadResources
    
    lblMD5.Caption = vbNullString
    txtNothing.ZOrder 1
    
    'move frame with "AnalyzeThis" button to the left a little bit (Vista+)
    If OSver.IsWindowsVistaOrGreater Then
        fraSubmit.Left = fraSubmit.Left - 65
    End If
    
    If OSver.IsLocalSystemContext Then
        'block some tools to prevent damage to system or output of inaccurate data
        mnuFileInstallHJT.Enabled = False
        mnuToolsRegUnlockKey.Enabled = False
        mnuToolsRegTypeChecker.Enabled = False
        mnuToolsUninst.Enabled = False
        mnuToolsStartupList.Enabled = False
    End If
    
    ' Set common events for controls
    ReDim ControlsEvent(0)
    'Set ControlsEvent(0).FrmInArr = Me
    For Each Ctl In Me.Controls
        i = i + 1
        ReDim Preserve ControlsEvent(0 To i)
        Select Case TypeName(Ctl)
            Case "CommandButtonW"
                Set ControlsEvent(i).BtnInArr = Ctl
            Case "TextBoxW"
                Set ControlsEvent(i).txtBoxInArr = Ctl
            Case "ListBoxW"
                Set ControlsEvent(i).lstBoxInArr = Ctl
            'Case "LabelW"
            '    'Set ControlsEvent(i).LblInArr = ctl
            Case "CheckBoxW"
                Set ChkB = Ctl
                'CheckBoxes in array dosn't support this type of events
                If ChkB.Name <> "chkConfigTabs" And ChkB.Name <> "chkHelp" Then
                    Set ControlsEvent(i).chkBoxInArr = Ctl
                End If
        End Select
    Next Ctl
    
    GetHosts
    GetBrowsersInfo
    
    Set Proc = New clsProcess
    
    cryptInit
    B64_Init
    
    'header of tracing log
    AppendErrorLogCustom vbCrLf & vbCrLf & "Logfile ( tracing ) of HijackThis+ v." & AppVerString & vbCrLf & vbCrLf & _
        "Command line: " & AppPath(True) & " " & g_sCommandLine & vbCrLf & vbCrLf & MakeLogHeader() & vbCrLf
    
    LoadLoLBinList
    LoadSettings
    
    Dim aFont() As String
    
    If Not bAutoLogSilent Then
        cmbFont.AddItem "Automatic" 'to use settings according to LCID -> see: SetFontCharSet() sub
        
        ReDim aFont(Screen.FontCount - 1)
        For i = 0 To Screen.FontCount - 1
            'exclude vertical fonts
            If Left$(Screen.Fonts(i), 1) <> "@" Then aFont(i) = Screen.Fonts(i)
        Next i
        'Sort the list
        QuickSort aFont, 0, UBound(aFont)
        
        For i = 0 To UBound(aFont)
            If Len(aFont(i)) <> 0 Then cmbFont.AddItem aFont(i)
        Next
        
        For i = 0 To cmbFont.ListCount - 1
            If cmbFont.List(i) = g_FontName Then
                cmbFont.ListIndex = i
                Exit For
            End If
        Next
        If cmbFont.ListIndex = -1 Then cmbFont.ListIndex = 0
        
        cmbFontSize.AddItem "Auto"
        For i = 6 To 14
            cmbFontSize.AddItem CStr(i)
        Next
        
        For i = 0 To cmbFontSize.ListCount - 1
            If cmbFontSize.List(i) = g_FontSize Then
                cmbFontSize.ListIndex = i
                Exit For
            End If
        Next
        'SetAllFontCharset Me, g_FontName, g_FontSize '(already raised by ListIndex change event)
    End If
    
    '/ihatewhitelists
    If HasCommandLineKey("ihatewhitelists") Then bIgnoreAllWhitelists = True: bHideMicrosoft = False 'must go after LoadSettings !!!
    '/default
    If HasCommandLineKey("default") Then bLoadDefaults = True
    If bLoadDefaults Then
        bAutoSelect = False
        bConfirm = True
        bMakeBackup = True
        bLogProcesses = True
        bLogModules = False
        bLogEnvVars = False
        bAdditional = True
        bSkipErrorMsg = False
        bMinToTray = False
        bCheckForUpdates = False
        bHideMicrosoft = True
        bIgnoreAllWhitelists = False
        g_bCheckSum = False
    End If
    '/skipIgnoreList
    If HasCommandLineKey("skipIgnoreList") Or bSkipIgnoreList Then
        AppendErrorLogCustom "Skipping ignore list..."
        bSkipIgnoreList = True
        IsOnIgnoreList vbNullString, EraseList:=True
    End If
    '/skipErrors
    If HasCommandLineKey("skipErrors") Then
        bSkipErrorMsg = True
    End If
    
    If HasCommandLineKey("delmode:disable") Then g_bDelmodeDisabling = True
    
    '/Area:xxx
    sCmdLine = Replace$(g_sCommandLine, ":", "+")
    
    '/Area:Processes
    If InStr(1, sCmdLine, "Area+Process", 1) > 0 Then bLogProcesses = True
    If InStr(1, sCmdLine, "Area-Process", 1) > 0 Then bLogProcesses = False
    '/Area:Modules
    If InStr(1, sCmdLine, "Area+Modules", 1) > 0 Then bLogModules = True
    If InStr(1, sCmdLine, "Area-Modules", 1) > 0 Then bLogModules = False
    '/Area:Environment
    If InStr(1, sCmdLine, "Area+Environment", 1) > 0 Then bLogEnvVars = True
    If InStr(1, sCmdLine, "Area-Environment", 1) > 0 Then bLogEnvVars = False
    '/Area:Additional
    If InStr(1, sCmdLine, "Area+Additional", 1) > 0 Then bAdditional = True
    If InStr(1, sCmdLine, "Area-Additional", 1) > 0 Then bAdditional = False
    
    fraConfig.Left = 120
    fraHelp.Left = 120
    fraConfig.Top = 120
    fraHelp.Top = 120
    fraMiscToolsScroll.Top = 0
    fraConfigTabs(0).Top = 830
    fraHostsMan.Top = 840
    fraConfigTabs(1).Top = 840
    fraConfigTabs(2).Top = 840
    fraConfigTabs(3).Top = 840
    
    If bAutoLogSilent Then
        Me.Height = 1800
        Me.Enabled = False 'Decrease message queue
        DisableProcessWindowsGhosting 'Disable queue
        'Revert (Core only): NtUserCallOneParam(1, ONEPARAM_ROUTINE_ENABLEPROCWNDGHSTING)
    Else
        bLockResize = True
        
        If Screen.Height >= 9000 Then
            Me.Height = CLng(RegReadHJT("WinHeight", "8355"))
            If Me.Height < 8355 Then Me.Height = 8355 'old = 8000
        Else
            Me.Height = CLng(RegReadHJT("WinHeight", "6600"))
            If Me.Height < 6600 Then Me.Height = 6600
        End If
        Me.Width = CLng(RegReadHJT("WinWidth", "9000"))
        If Me.Width < 9000 Then Me.Width = 9000
        
        bLockResize = False
        
        LoadWindowPos Me, SETTINGS_SECTION_MAIN
        
    End If
    
    'If RegReadHJT("SkipIntroFrame", "0") = "0" Or (ConvertVersionToNumber(RegReadHJT("Version", vbNullString)) < ConvertVersionToNumber("2.7.0.11")) Then
        fraN00b.Visible = True
        fraScan.Visible = False
        fraOther.Visible = False
        lstResults.Visible = False
        fraSubmit.Visible = False
        NotifyChangeFrame FRAME_ALIAS_MAIN
        
    '    Call RegSaveHJT("SkipIntroFrame", "0")
        
    'Else
    '    chkSkipIntroFrame.Value = 1
    '    pictLogo.Visible = False
    '    NotifyChangeFrame FRAME_ALIAS_SCAN
    'End If
    
    If Not bAutoLogSilent Then
        If Not CheckForReadOnlyMedia() Then
            g_NeedTerminate = True
        End If
    End If
    
    If Not bAutoLogSilent Then
        If CheckForStartedFromTempDir() Then
            g_NeedTerminate = True
            Exit Sub
        End If
    End If
    
    If Not bIsWinNT Then cmdDeleteService.Enabled = False
    
    If Not bAutoLogSilent Then
        SetMenuIcons Me
    End If
    
    AppendErrorLogCustom "FormStart_Stage1 - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FormStart_Stage1"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub FormStart_Stage2()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "FormStart_Stage2 - Begin"
    
    Static bInit As Boolean
    Dim bSilentUninst As Boolean
    Dim aKey() As String
    Dim aValue() As String
    Dim lTotal As Long
    Dim i As Long
    Dim sMsg As String
    
    If bInit Then
        Exit Sub
    Else
        bInit = True
    End If
    
    '/silentuninstall
    bSilentUninst = HasCommandLineKey("silentuninstall")
    
    '/uninstall
    If HasCommandLineKey("uninstall") Or bSilentUninst Then
        Me.Hide
        If Not HJT_Uninstall(bSilentUninst) Then
            g_ExitCodeProcess = 1
        End If
        Unload Me
        Exit Sub
    End If
    
    If g_NeedTerminate Then Unload Me: Exit Sub

    '/md5 /sha1 /sha256
    If HasCommandLineKey("md5") Then g_bCheckSum = True: g_eUseHashType = HASH_TYPE_MD5
    If HasCommandLineKey("sha1") Then g_bCheckSum = True: g_eUseHashType = HASH_TYPE_SHA1
    If HasCommandLineKey("sha256") Then g_bCheckSum = True: g_eUseHashType = HASH_TYPE_SHA256
    '/deleteonreboot
    If HasCommandLineKey("deleteonreboot") Then
        SilentDeleteOnReboot UnQuote(g_sCommandLine)
        Unload Me
        Exit Sub
    End If
    
    If (Not inIDE) And (Not bPolymorph) Then
        Err.Clear
        g_hMutex = CreateMutex(0&, 1&, StrPtr("mutex_HijackThisPlus"))
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) And 0 = Len(g_sCommandLine) Then
            If Not bAutoLogSilent Then
                If MsgBoxW(Translate(2), vbExclamation Or vbYesNo, g_AppName) = vbNo Then Unload Me: Exit Sub
            End If
        End If
    End If
    
    If bCheckForUpdates Then
        If Not bAutoLogSilent Then
            CheckForUpdate True, bUpdateSilently, bUpdateToTest
            If g_NeedTerminate Then Unload Me: Exit Sub
        End If
    End If
    
    Dim sCmdLine$
    sCmdLine = g_sCommandLine
    
    '/install
    If HasCommandLineKey("install") Then
        
        '/autostart
        If HasCommandLineKey("autostart") Then
            
            'check /autostart d:X
            
            Dim bSetDelay As Boolean
            Dim lDelay As Long
            
            lTotal = ParseSubCmdLine(sCmdLine, "autostart", aKey(), aValue())
            For i = 0 To lTotal - 1
                Select Case UCase$(aKey(i))
                Case "D"
                    If IsNumeric(aValue(i)) Then
                        lDelay = CLng(aValue(i))
                        bSetDelay = True
                    End If
                End Select
            Next
            
            If (bSetDelay) Then
                InstallAutorunHJT True, lDelay, True
            Else
                InstallAutorunHJT True, , True
            End If
        Else
            InstallHJT True, HasCommandLineKey("noGUI")   '/noGUI
        End If
        Unload Me
        Exit Sub
    End If
    
    If (Not bAutoLog) And (Not inIDE) Then
        CheckInstalledVersionHJT
    End If
    
    If bDebugMode Or bDebugToFile Then
        'checking is EDS machanism working correclty
        Dim SignResult As SignResult_TYPE
    
        'check sign. of core dll
        SignVerify BuildPath(sWinDir, "system32\ntdll.dll"), SV_LightCheck Or SV_SelfTest, SignResult
        Dbg "Fingerprint should be: CDD4EEAE6000AC7F40C3802C171E30148030C072 or 3B1EFD3A66EA28B16697394703A72CA340A05BD5"
        If StrComp(SignResult.HashRootCert, "CDD4EEAE6000AC7F40C3802C171E30148030C072", 1) = 0 Or _
          StrComp(SignResult.HashRootCert, "3B1EFD3A66EA28B16697394703A72CA340A05BD5", 1) = 0 Then
            Dbg "Fingerprint is matched (OK)."
        Else
            Dbg "Fingerprint is NOT matched (FAILED)."
        End If
        'check sign of self
        SignVerify AppPath(True), SV_SelfTest Or SV_PreferInternalSign Or SV_AllowExpired, SignResult
        
        If Not IsDragokasSign(SignResult) Then
            Dbg "HJT internal signature is INVALID."
        End If
    End If
    
    MyParentProc.pid = GetParentPID(GetCurrentProcessId())
    If (MyParentProc.pid <> 0) Then
        MyParentProc.Path = GetFilePathByPID(MyParentProc.pid)
    End If
    
    Test 'for all of my tests
    
    If HasCommandLineKey("noBackup") Then bMakeBackup = False
    If HasCommandLineKey("noGUI") Then g_bNoGUI = True
    
    DoHotFixes
    
    #If Not NoSelfSignTest Then
        If (Not inIDE) And Len(g_sCommandLine) = 0 And Not bPolymorph Then
            If Not CheckIntegrityHJT() Then
                If Not bAutoLogSilent Then
                    'Warning! Integrity of HiJackThis program is corrupted. Perhaps, file is patched or infected by file virus.
                    sMsg = Translate(1023)
                    If IsRussianLangCode(OSver.LangSystemCode) Or IsRussianLangCode(OSver.LangDisplayCode) Then
                        'https://support.kaspersky.ru/14226
                        'https://free.drweb.ru/aid_admin
                        sMsg = sMsg & Caes_Decode("iwywB://JNKMNSW.RJ^]Tcf`p.mr/05573") & " / "
                        sMsg = sMsg & Caes_Decode("iwywB://wKzB.EU\LK._d/T^[X\aljq")
                    Else
                        'https://www.kaspersky.com/downloads/free-rescue-disk
                        'https://free.drweb.com/aid_admin/?lng=en
                        sMsg = sMsg & Caes_Decode("iwywB://NPR.JBVUL[^Xh.Vdd/_lvoothmD/uIxz-QFVH\N-QXd^") & " / "
                        sMsg = sMsg & Caes_Decode("iwywB://wKzB.EU\LK.P^^/V`]Z^cnls/?wAv=xI")
                    End If
                    ErrReport = ErrReport & vbCrLf & sMsg & vbCrLf
                    MsgBoxW sMsg, vbExclamation
                End If
            End If
        End If
    #End If
    
    '/tool:xxx
    If bRunToolStartupList Then
        'vbModal is not working here -> walkaround
        g_bStartupListTerminateOnExit = True
        RunStartupList False 'True
        'frmStartupList2.Show vbModal
        'Unload Me: Exit Sub
    End If
    If bRunToolUninstMan Then
        frmUninstMan.Show vbModal
        Unload Me: Exit Sub
    End If
    If bRunToolEDS Then
        frmCheckDigiSign.Show vbModal
        Unload Me: Exit Sub
    End If
    If bRunToolRegUnlocker Then
        frmUnlockRegKey.Show vbModal
        Unload Me: Exit Sub
    End If
    If bRunToolRegTypeChecker Then
        frmRegTypeChecker.Show vbModal
        Unload Me: Exit Sub
    End If
    If bRunToolADSSpy Then
        frmADSspy.Show vbModal
        Unload Me: Exit Sub
    End If
    If bRunToolHosts Then
        Me.Show
        mnuToolsHosts_Click
    End If
    If bRunToolProcMan Then
        frmProcMan.Show vbModal
        Unload Me: Exit Sub
    End If
    If bRunToolCBL Then
        mnuToolsShortcutsChecker_Click
        Unload Me: Exit Sub
    End If
    If bRunToolClearLNK Then
        mnuToolsShortcutsFixer_Click
        Unload Me: Exit Sub
    End If
    
    FormStart_Stage3 'private switches
    
    If HasCommandLineKey("Area:None") Then
        DeleteFile StrPtr(g_sLogFile)
    Else
        CheckAutoLog 'UI & Autolog stuff if required
    End If
    
    #If SCRIPT_FIX Then
    If Not bAutoLog Then
        If Script.HasFixInClipboard() Then
            'Execute fix from clipboard?
            If MsgBox(Translate(2500), vbYesNo Or vbQuestion, g_AppName) = vbYes Then
                Script.ExecuteFixFromClipboard False
            End If
        End If
    End If
    #End If
    
    AppendErrorLogCustom "FormStart_Stage2 - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FormStart_Stage2"
    If inIDE Then Stop: Resume Next
End Sub

Sub DoHotFixes()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "frmMain.DoHotFixes - Begin"

    Dim i As Long
    
    'g_bBackupMade = False
    
    If InStr(1, g_sCommandLine, "/Fix", 1) <> 0 Or _
        InStr(1, g_sCommandLine, "-Fix", 1) <> 0 Then
        
        g_bFixArg = True
        lstResults.Clear
    End If
    
    If HasCommandLineKey("Disinfect") Then
        g_bFixHosts = True
        g_bFixO4 = True
        g_bFixPolicy = True
        g_bFixCert = True
        g_bFixIpSec = True
        g_bFixEnvVar = True
        g_bFixO20 = True
        g_bFixO21 = True
        g_bFixTasks = True
        g_bFixServices = True
        g_bFixWMIJob = True
        g_bFixIFEO = True
    End If
    
    If HasCommandLineKey("FixHosts") Then g_bFixHosts = True    'O1
    If HasCommandLineKey("FixO4") Then g_bFixO4 = True          'O4
    If HasCommandLineKey("FixPolicy") Then g_bFixPolicy = True  'O7
    If HasCommandLineKey("FixCert") Then g_bFixCert = True      'O7
    If HasCommandLineKey("FixIpSec") Then g_bFixIpSec = True    'O7
    If HasCommandLineKey("FixEnvVar") Then g_bFixEnvVar = True  'O7
    If HasCommandLineKey("FixO20") Then g_bFixO20 = True        'O20
    If HasCommandLineKey("FixO21") Then g_bFixO21 = True        'O21
    If HasCommandLineKey("FixTasks") Then g_bFixTasks = True    'O22
    If HasCommandLineKey("FixServices") Then g_bFixServices = True  'O23
    If HasCommandLineKey("FixWMIJob") Then g_bFixWMIJob = True      'O25
    If HasCommandLineKey("FixIFEO") Then g_bFixIFEO = True          'O26
    
    If g_bFixHosts Then CheckO1Item
    If g_bFixO4 Then CheckO4Item
    If g_bFixPolicy Then
        CheckPolicies
        CheckPolicyACL
    End If
    If g_bFixCert Then CheckCertificatesEDS
    If g_bFixIpSec Then CheckIPSec
    If g_bFixEnvVar Then CheckSystemProblemsEnvVars
    If g_bFixO20 Then CheckO20Item
    If g_bFixO21 Then CheckO21Item
    If g_bFixTasks Then CheckO22Item
    If g_bFixServices Then CheckO23Item
    If g_bFixWMIJob Then CheckO25Item
    If g_bFixIFEO Then CheckO26Item
    
    If HasCommandLineKey("FreezeProcess") Then
        Call FreezeCustomProcesses
    End If
    
    If lstResults.ListCount <> 0 And Not inIDE Then
        For i = 0 To lstResults.ListCount - 1
            lstResults.ItemChecked(i) = True
        Next
        cmdFix_Click
        
        If HasCommandLineKey("LockPoints") Then
            Call LockAutorunPoints
        End If
        
        Call Kill_LOLBIN
        
        For i = 0 To lstResults.ListCount - 1
            lstResults.ItemChecked(i) = True
        Next
        cmdFix_Click
    End If
    
    If HasCommandLineKey("LockPoints") Then
        Call LockAutorunPoints
    End If
    
'    '/FixHosts
'    If g_bFixHosts Then 'O1
'        lstResults.Clear
'        CheckO1Item
'        If lstResults.ListCount <> 0 Then
'            sItem = lstResults.List(0)
'            If GetScanResults(sItem, result) Then
'                result.HitLineW = "O1 - Hosts: is empty"
'                MakeBackupEx result
'                FixO1Item result.HitLineW, result
'                FlushDNS
'            End If
'        End If
'    End If
'
'    If g_bFixO4 Then
'        lstResults.Clear
'        CheckO4Item
'        If lstResults.ListCount <> 0 Then
'            sItem = lstResults.List(0)
'            If GetScanResults(sItem, result) Then
'                MakeBackupEx result
'                FixO4Item result.HitLineW, result
'            End If
'        End If
'    End If
'
'    If g_bFixPolicy Or g_bFixCert Or g_bFixIpSec Or g_bFixEnvVar Then 'O7
'        lstResults.Clear
'        CheckO7Item
'        For i = 0 To lstResults.ListCount - 1
'            sItem = lstResults.List(i)
'            If GetScanResults(sItem, result) Then
'                If g_bFixPolicy Then
'                    If StrBeginWith(sItem, "O7 - Policy:") Then MakeBackupEx result: FixO7Item sItem, result
'                    If StrBeginWith(sItem, "O7 - Taskbar policy:") Then MakeBackupEx result: FixO7Item sItem, result
'                    If StrBeginWith(sItem, "O7 - Explorer Policy:") Then MakeBackupEx result: FixO7Item sItem, result
'                End If
'                If g_bFixCert Then
'                    If StrBeginWith(sItem, "O7 - Policy: [Untrusted Certificate]") Then
'                        If InStr(1, sItem, "Fix all items") = 0 Then
'                            MakeBackupEx result: FixO7Item sItem, result
'                        End If
'                    End If
'                End If
'                If g_bFixIpSec Then
'                    If StrBeginWith(sItem, "O7 - IPSec:") Then MakeBackupEx result: FixO7Item sItem, result
'                End If
'                If g_bFixEnvVar Then
'                    If StrBeginWith(sItem, "O7 - TroubleShooting: (EV)") Then MakeBackupEx result: FixO7Item sItem, result
'                End If
'            End If
'        Next
'        If lstResults.ListCount <> 0 Then
'            If bUpdatePolicyNeeded Then UpdatePolicy
'        End If
'    End If
'
'    If g_bFixWMIJob Then 'O25
'        lstResults.Clear
'        CheckO25Item
'        For i = 0 To lstResults.ListCount - 1
'            sItem = lstResults.List(i)
'            If GetScanResults(sItem, result) Then MakeBackupEx result: FixO25Item sItem, result
'        Next
'    End If
'
'    If g_bFixIFEO Then 'O26
'        lstResults.Clear
'        CheckO26Item
'        For i = 0 To lstResults.ListCount - 1
'            sItem = lstResults.List(i)
'            If GetScanResults(sItem, result) Then MakeBackupEx result: FixO26Item sItem, result
'        Next
'    End If
'
'    If g_bBackupMade Then
'        IncreaseNumberOfFixes
'    End If
    
    AppendErrorLogCustom "frmMain.DoHotFixes - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FormStart_Stady2"
    If inIDE Then Stop: Resume Next
End Sub

Sub AddFirewallAllowed(sFile As String)
    Dim sFilename As String
    sFilename = GetFileName(sFile, True)
    
    'netsh advfirewall firewall add rule name="Autoruns_in" dir=in action=allow program=".....\autorunsc.exe" enable=yes
    
    If Proc.ProcessRun(BuildPath(sWinSysDir, "netsh.exe"), _
          "advfirewall firewall add rule name=""" & sFilename & "_in" & """ dir=in action=allow program=""" & sFile & """ enable=yes", , vbHide) Then
            Proc.WaitForTerminate , , True
    End If
    If Proc.ProcessRun(BuildPath(sWinSysDir, "netsh.exe"), _
          "advfirewall firewall add rule name=""" & sFilename & "_out" & """ dir=out action=allow program=""" & sFile & """ enable=yes", , vbHide) Then
            Proc.WaitForTerminate , , True
    End If
End Sub

Private Sub FormStart_Stage3()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "FormStart_Stage3 - Begin"
    
    Dim i As Long, j As Long, k As Long
    Dim sLogDir As String
    'Dim sToolsDir As String
    Dim sFile As String
    Dim bSigSystemOK As Boolean
    Dim sAutorunsExe As String
    Dim hFile As Long, hFile1 As Long, hFile2 As Long, hFile3 As Long
    Dim sLine As String
    Dim sLog As String
    Dim sLogToCheck As String
    Dim sLogClear As String
    Dim sLogSuspicious As String
    Dim sLogFailed As String
    Dim sWhiteListPath As String
    Dim bToolsExecuted As Boolean
    
    'sToolsDir = BuildPath(AppPath(), "tools")
    sLogDir = BuildPath(AppPath(), "LOG")
    
    sLogToCheck = BuildPath(sLogDir, "files_to_check.txt")
    sLogClear = BuildPath(sLogDir, "files_clear.txt")
    sLogSuspicious = BuildPath(sLogDir, "files_suspicious.txt")
    sLogFailed = BuildPath(sLogDir, "files_failed.txt")
    sWhiteListPath = BuildPath(AppPath(), "whitelists.txt")
    
    If (OSver.IsWin64) Then
        sAutorunsExe = "auto64.exe"
    Else
        sAutorunsExe = "auto.exe"
    End If

    If HasCommandLineKey("addfirewall") Then

        If bRunToolAutoruns Then
            AddFirewallAllowed BuildPath(AppPath(), "tools\Scan\" & sAutorunsExe)
        End If
        
    End If
   
'    bRunToolAutoruns = True
'    bRunToolExecuted = True
'    bRunToolLastActivity = True
'    bRunToolServiWin = True
'    bRunToolTaskScheduler = True
'    g_bVTCheck = True
'    g_bSigCheck = True
'    g_bRawIgnoreList = True
    
    If bRunToolAutoruns Then
    
        MkDirW sLogDir
        SetCurrentDirectory StrPtr(sLogDir)
        
        '-a * -m -s -u -v -vt -nobanner -o results.xml
        'bug: -x doesn't redirect to file, so we need to intercept console by handles
'        If Proc.ProcessRun(BuildPath(AppPath(), "tools\Scan\" & sEXE), "-a i -nobanner -x", , vbNormalFocus, , True) Then
'            '// TODO: true if only silent mode, otherwise set value from /timeout:X
'            Debug.Print Proc.ConsoleRead()
'            Proc.WaitForTerminate , , True
'            Set Proc = New clsProcess
'        End If

        If Proc.ProcessRun(Environ$("ComSpec"), "/c """"" & BuildPath(AppPath(), "tools\Scan\" & sAutorunsExe) & """ -accepteula -a * -nobanner -x > """ _
            & BuildPath(sLogDir, "results.xml") & """""", , vbHide, , False) Then
            
            Proc.WaitForTerminate , , True
        End If
        bToolsExecuted = True
    End If
    
    If bRunToolExecuted Then
        MkDirW sLogDir
        SetCurrentDirectory StrPtr(sLogDir)
        If Proc.ProcessRun(BuildPath(AppPath(), "tools\Scan\executed.exe"), "/scomma executed.csv", , vbHide) Then
            Proc.WaitForTerminate , , True
        End If
        bToolsExecuted = True
    End If
    If bRunToolLastActivity Then
        MkDirW sLogDir
        SetCurrentDirectory StrPtr(sLogDir)
        If Proc.ProcessRun(BuildPath(AppPath(), "tools\Scan\lastactivity.exe"), "/scomma lastactivity.csv", , vbHide) Then
            Proc.WaitForTerminate , , True
        End If
        bToolsExecuted = True
    End If
    If bRunToolServiWin Then
        MkDirW sLogDir
        SetCurrentDirectory StrPtr(sLogDir)
        If Proc.ProcessRun(BuildPath(AppPath(), "tools\Scan\serwin.exe"), "/scomma services services.csv /status ""~started""", , vbHide) Then
            Proc.WaitForTerminate , , True
        End If
        bToolsExecuted = True
    End If
    If bRunToolTaskScheduler Then
        MkDirW sLogDir
        SetCurrentDirectory StrPtr(sLogDir)
        If Proc.ProcessRun(BuildPath(AppPath(), "tools\Scan\sheduler.exe"), "/sort ~status /scomma sheduler.csv", , vbHide) Then
            Proc.WaitForTerminate , , True
        End If
        bToolsExecuted = True
    End If
    
    SetCurrentDirectory StrPtr(AppPath())
    
    If HasCommandLineKey("delmode:pending") Then
        g_bDelModePending = True
    End If
    
    If HasCommandLineKey("autofix:vt") Then
        g_bAutoFixVT = True
    End If
    
    If HasCommandLineKey("vtcheck") Then
        g_bVTCheck = True
    End If
    
    If HasCommandLineKey("rawIgnoreList") Then
        g_bRawIgnoreList = True
    End If
    
    If HasCommandLineKey("sigcheck") Then
        g_bSigCheck = True
    End If
    
    ' ------------------------------
    '  !!! EXIT if NONE selected !!!
    ' ------------------------------
    
    If Not (g_bAutoFixVT Or g_bVTCheck) Then
'        If bAutoLogSilent Then
'            Unload Me
'        End If
        Exit Sub
    End If
    
    bSigSystemOK = isEDS_Work()
    
    'sigcheck
    Dim dRunFiles As clsTrickHashTable
    Dim dClearFiles As clsTrickHashTable
    Dim dSuspFiles As clsTrickHashTable

    Set dRunFiles = New clsTrickHashTable
    Set dClearFiles = New clsTrickHashTable
    Set dSuspFiles = New clsTrickHashTable

    dRunFiles.CompareMode = vbTextCompare

    If bRunToolAutoruns Then
        ParseFilesXML dRunFiles, BuildPath(sLogDir, "results.xml")
    End If
    If bRunToolExecuted Then
        ParseFilesCSV dRunFiles, BuildPath(sLogDir, "executed.csv"), 1
    End If
    If bRunToolLastActivity Then
        ParseFilesCSV dRunFiles, BuildPath(sLogDir, "lastactivity.csv"), 4, "exe,dll"
    End If
    If bRunToolServiWin Then
        ParseFilesCSV dRunFiles, BuildPath(sLogDir, "services.csv"), 13
    End If
    If bRunToolTaskScheduler Then
        ParseFilesCSV dRunFiles, BuildPath(sLogDir, "sheduler.csv"), 20
    End If

    ' WhiteListing

    If g_bRawIgnoreList Then

        OpenW sWhiteListPath, FOR_READ, hFile

        If hFile > 0 Then
            Do While LineInputW(hFile, sLine)
                If dRunFiles.Exists(sLine) Then dRunFiles.Remove sLine
            Loop
            CloseW hFile
        End If
    End If
    
    If g_bSigCheck Then
        Dim dTemp As clsTrickHashTable
        Set dTemp = New clsTrickHashTable
        dTemp.CompareMode = vbTextCompare

        For i = 0 To dRunFiles.Count - 1
            sFile = dRunFiles.Keys(i)
            DoEvents
            If Not IsMicrosoftFile(sFile, Not bSigSystemOK, True) Then
                dTemp.Add sFile, 0
                If inIDE Then Debug.Print sFile
            End If
        Next
        Set dRunFiles = dTemp
        Set dTemp = Nothing
    End If

    'Logging

    If dRunFiles.Count Then

        If OpenW(sLogToCheck, FOR_OVERWRITE_CREATE, hFile) Then

            For i = 0 To dRunFiles.Count - 1
                sFile = dRunFiles.Keys(i)
                PrintLineW hFile, sFile
            Next

            CloseW hFile, True
        End If
    End If

    ' VIRUSTOTAL check

    If inIDE Then Debug.Print "Run VT check ..."

    DeleteFile StrPtr(sLogClear)
    DeleteFile StrPtr(sLogSuspicious)
    DeleteFile StrPtr(sLogFailed)

    OpenW sLogClear, FOR_READ_WRITE, hFile1
    OpenW sLogSuspicious, FOR_READ_WRITE, hFile2
    OpenW sLogFailed, FOR_READ_WRITE, hFile3

    Const MAX_CHECK_ATTEMPT As Long = 20
    Const PROC_TIMEOUT_SEC As Long = 60

    Dim cProc() As clsProcess
    Dim AppInitBak As String
    Dim nDetects As Long
    Dim sURL As String
    Dim sCheckedFile As String
    Dim dTime() As Date
    Dim bRunned As Boolean
    Dim bInProgress As Boolean
    Dim nCheckIdx As Long
    Dim nPrevIdx As Long
    Dim bNextIdx As Boolean
    Dim eProc1() As MY_PROC_ENTRY
    Dim eProc2() As MY_PROC_ENTRY
    Dim sAutorunsExePath As String
    Dim bPidMatch As Boolean
    Dim dNull As Date
    
    sAutorunsExePath = BuildPath(AppPath(), "tools\Scan\" & sAutorunsExe)

    Dim lThreads As Long
    lThreads = 10

    ReDim Preserve dTime(lThreads - 1)
    ReDim Preserve cProc(lThreads - 1)
    For i = 0 To UBound(cProc)
        Set cProc(i) = New clsProcess
    Next

    Do While dRunFiles.Count Or bInProgress

        DoEvents
        bInProgress = False
        
        For i = 0 To UBound(cProc)

            bRunned = cProc(i).IsRunned

            If bRunned Then
                'check for timeout
                bInProgress = True ' at least one file in runned process
                
                If DateDiff("s", dTime(i), Now()) > PROC_TIMEOUT_SEC Then
                    If cProc(i).pid <> 0 Then
                        cProc(i).ProcessClose
                        cProc(i).Identifier = vbNullString
                        bRunned = False
                        
                        sLog = BuildPath(sLogDir, "vt_result_" & CStr(i) & ".xml")
                        DeleteFile StrPtr(sLog)
                    End If
                    If inIDE Then Debug.Print "[Thread:" & i & "] - TIMEOUT !"
                End If
            End If
            
            If Not bRunned Then
                cProc(i).Identifier = vbNullString
                
                'get previous result
                sLog = BuildPath(sLogDir, "vt_result_" & CStr(i) & ".xml")
                
                If (ParseVTResult(sLog, sCheckedFile, nDetects, sURL)) Then
                    
                    If nDetects = 0 Then
                        If sURL = "n/a" Then
                            If (hFile3 > 0) Then
                                PrintLineW hFile3, sCheckedFile
                            End If
                        ElseIf Not dClearFiles.Exists(sCheckedFile) Then
                            dClearFiles.Add sCheckedFile, 0

                            sLine = sCheckedFile & " - " & sURL

                            If (hFile1 > 0) Then
                                PrintLineW hFile1, sLine
                            End If
                        End If
                    Else
                        If Not dSuspFiles.Exists(sCheckedFile) Then
                            dSuspFiles.Add sCheckedFile, nDetects

                            sLine = sCheckedFile & " - [" & CStr(nDetects) & "] - " & sURL

                            If (hFile2 > 0) Then
                                PrintLineW hFile2, sLine
                            End If
                        End If
                    End If

                    If dRunFiles.Exists(sCheckedFile) Then
                        dRunFiles.Remove sCheckedFile
                    End If

                    If inIDE Then Debug.Print "[Thread:" & i & "] - Result: Detects - " & nDetects & " - " & sCheckedFile & " - " & sURL

                End If

                DeleteFile StrPtr(sLog)

                If dRunFiles.Count Then

                    AppInitBak = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs")

                    If nCheckIdx > (dRunFiles.Count - 1) Then
                        nCheckIdx = 0
                    End If

                    ' exclude files already been checked in runned processes
                    nPrevIdx = nCheckIdx
                    Do
                        sFile = dRunFiles.Keys(nCheckIdx)

                        bNextIdx = False
                        For j = 0 To UBound(cProc)
                            If StrComp(cProc(j).Identifier, sFile, 1) = 0 Then
                                bNextIdx = True
                                Exit For
                            End If
                        Next
                        
                        If bNextIdx Then
                            nCheckIdx = nCheckIdx + 1

                            If nCheckIdx > (dRunFiles.Count - 1) Then
                                nCheckIdx = 0
                            End If

                            If (nCheckIdx = nPrevIdx) Then
                                sFile = vbNullString
                                Exit Do 'loop is exceeded
                            End If
                        End If
                    Loop While bNextIdx

                    If Len(sFile) <> 0 Then
                        'substitute the file we need to check
                        Reg.SetStringVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", sFile

                        If dRunFiles(sFile) >= MAX_CHECK_ATTEMPT Then

                            dRunFiles.Remove sFile
                            If hFile3 > 0 Then
                                PrintLineW hFile3, sFile
                            End If
                        Else

                            dRunFiles(sFile) = dRunFiles(sFile) + 1

                            If inIDE Then Debug.Print "[Thread:" & i & "] Attempt #" & dRunFiles(sFile) & ". VT Checking file: " & sFile
                            If inIDE Then Debug.Print "Files left: " & dRunFiles.Count
                            
                            GetProcesses eProc1
                            
                            dTime(i) = Now()
                            cProc(i).Identifier = sFile
                            '-accepteula -a d -vs -vt -nobanner -x
                            cProc(i).ProcessRun Environ$("ComSpec"), "/c """"" & sAutorunsExePath & """ " & Caes_Decode("-dhjnAGtLEv -B I -a` -ii -ilabqslA -M") & " > """ _
                                & sLog & """""", , vbHide, , False

                            SleepNoLock 200

                            cProc(i).pid = 0
                            cProc(i).ThreadId = 0
                            cProc(i).HandleProc = 0
                            cProc(i).HandleThread = 0
                            
                            GetProcesses eProc2
                            
                            'substitute autorunsc PID instead of cmd.exe and set it for current process class
                            Do
                                For k = 0 To UBound(eProc2)
                                    If StrComp(sAutorunsExePath, eProc2(k).Path, 1) = 0 Then
                                        bPidMatch = False
                                        For j = 0 To UBound(eProc1)
                                            If StrComp(sAutorunsExePath, eProc1(j).Path, 1) = 0 Then
                                                If eProc1(j).pid = eProc2(k).pid Then
                                                    bPidMatch = True
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        If Not bPidMatch Then
                                            cProc(i).pid = eProc2(k).pid
                                            Exit Do
                                        End If
                                    End If
                                Next
                            Loop While 0
                            
                            'just in case: ensure no other autoruns process (runned from installation folder) has timed out
                            For k = 0 To UBound(eProc2)
                                If StrComp(sAutorunsExePath, eProc2(k).Path, 1) = 0 Then
                                    'not null date
                                    If eProc2(k).CreationTime <> dNull Then
                                        If DateDiff("s", eProc2(k).CreationTime, Now()) > PROC_TIMEOUT_SEC Then
                                            Proc.ProcessClose , , True, , eProc2(k).pid
                                        End If
                                    End If
                                End If
                            Next
                            
                            'restore
                            Reg.SetStringVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", AppInitBak

                            nCheckIdx = nCheckIdx + 1
                        End If
                    End If
                End If
            End If

        Next

        DoEvents
        SleepNoLock 1000

    Loop
    
    For i = 0 To UBound(cProc)
        sLog = BuildPath(sLogDir, "vt_result_" & CStr(i) & ".xml")
        DeleteFile StrPtr(sLog)
    Next
    
    CloseW hFile1, True
    CloseW hFile2, True
    CloseW hFile3, True
        
    '=============
    '   DO A FIX
    '=============
    If g_bAutoFixVT Then

        For i = 0 To dSuspFiles.Count - 1

            sFile = dSuspFiles.Keys(i)
            DeleteFileForce sFile, True

        Next

        If HasCommandLineKey("reboot") Then '/reboot
            If dSuspFiles.Count Then
                RestartSystem , True, True
            End If
        End If
    End If
    
    If bAutoLogSilent Then
        ExitProcess 0
        'Unload Me
        Exit Sub
    End If
    
    AppendErrorLogCustom "FormStart_Stage3 - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FormStart_Stage3"
    If inIDE Then Stop: Resume Next
End Sub

Function ParseVTResult(sLog As String, sFile As String, nDetects As Long, sURL As String) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "frmMain.ParseVTResult - Begin"

    Dim hFile As Long
    Dim sLine As String
    Dim aLine() As String
    Dim sContent As String
    Dim i As Long
    Dim pos As Long

    sFile = vbNullString
    nDetects = 0
    sURL = vbNullString

    OpenW sLog, FOR_READ, hFile
    
    If hFile > 0 Then
        CloseW hFile
        sContent = ReadFileContents(sLog, FileGetTypeBOM(sLog) = CP_UTF16LE)
        
        If Len(sContent) < 2 Then Exit Function
        
        sContent = Replace(sContent, vbCr, vbNullString)
        aLine = Split(sContent, vbLf)
        
        For i = 0 To UBoundSafe(aLine)
            sLine = aLine(i)
            
            If (StrBeginWith(sLine, "<imagepath>")) Then
                sLine = mid$(sLine, Len("<imagepath>") + 1)
                pos = InStr(1, sLine, "</imagepath>", 1)
                If pos <> 0 Then
                    sFile = Left$(sLine, pos - 1)
                    sFile = Replace(sFile, "&#34;", """")
                End If
            ElseIf (StrBeginWith(sLine, "<vt-detection>")) Then
                sLine = mid$(sLine, Len("<vt-detection>") + 1)
                pos = InStr(1, sLine, "</vt-detection>", 1)
                If pos <> 0 Then
                    sLine = Left$(sLine, pos - 1)
                    If sLine = "Unknown" Then
                        nDetects = 0
                        ParseVTResult = True
                    Else
                        pos = InStr(sLine, "/")
                        If pos <> 0 Then
                            sLine = Left$(sLine, pos - 1)
                            If IsNumeric(sLine) Then
                                nDetects = CLng(sLine)
                                ParseVTResult = True
                            End If
                        End If
                    End If
                End If
            ElseIf (StrBeginWith(sLine, "<vt-permalink>")) Then
                sLine = mid$(sLine, Len("<vt-permalink>") + 1)
                pos = InStr(1, sLine, "</vt-permalink>", 1)
                If pos <> 0 Then
                    sURL = Left$(sLine, pos - 1)
                End If
            End If
        Next
    End If
    
    AppendErrorLogCustom "frmMain.ParseVTResult - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FormStart_Stady3"
    If inIDE Then Stop: Resume Next
End Function

Sub ParseFilesCSV(dRunFiles As clsTrickHashTable, sLog As String, lColumnPos As Long, Optional sExtensions As String = vbNullString)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "frmMain.ParseFilesXML - Begin"
    
    Dim hFile As Long
    Dim sLine As String
    Dim bUseFilter As Boolean
    Dim aExt() As String
    Dim i As Long
    Dim sFile As String
    Dim aTok() As String
    Dim bComply As Boolean
    
    If (lColumnPos < 1) Then Exit Sub 'failure
    
    If Len(sExtensions) <> 0 Then
        bUseFilter = True
        aExt = Split(sExtensions, ",")
        For i = 0 To UBound(aExt)
            If (Left$(aExt(i), 1) <> ".") Then aExt(i) = "." & aExt(i)
        Next
    End If
    
    OpenW sLog, FOR_READ, hFile
    
    If hFile > 0 Then
        Do While LineInputW(hFile, sLine)
            sFile = vbNullString
            aTok = Split(sLine, ",", lColumnPos + 1)
            If AryPtr(aTok) Then
                If UBound(aTok) + 1 >= lColumnPos Then
                    sFile = aTok(lColumnPos - 1)
                End If
            Else
                If lColumnPos = 1 Then
                    sFile = sLine
                End If
            End If
            If Len(sFile) <> 0 Then
                sFile = UnQuote(sFile)
                bComply = False
                If (bUseFilter) Then
                    If InArray(GetExtensionName(sFile), aExt, , , vbTextCompare) Then
                        bComply = True
                    End If
                Else
                    bComply = True
                End If
                If bComply Then
                    sFile = FindOnPath(sFile)
                    If Len(sFile) <> 0 Then
                        sFile = GetLongPath(sFile)
                        sFile = GetFullPath(sFile)
                        If Not dRunFiles.Exists(sFile) Then dRunFiles.Add sFile, 0
                    End If
                End If
            End If
        Loop
        CloseW hFile
    End If
    
    DoEvents
    AppendErrorLogCustom "frmMain.ParseFilesXML - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_CheckAutoLog"
    If inIDE Then Stop: Resume Next
End Sub

Sub ParseFilesXML(dRunFiles As clsTrickHashTable, sLog As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "frmMain.ParseFilesXML - Begin"
    
    Dim hFile As Long
    Dim sLine As String
    Dim aLine() As String
    Dim sFile As String
    Dim sContent As String
    Dim pos As Long
    Dim i As Long
    
    OpenW sLog, FOR_READ, hFile
    
    If hFile > 0 Then
        CloseW hFile
        sContent = ReadFileContents(sLog, FileGetTypeBOM(sLog) = CP_UTF16LE)
        
        If Len(sContent) < 2 Then Exit Sub
        
        sContent = Replace(sContent, vbCr, vbNullString)
        aLine = Split(sContent, vbLf)
        
        For i = 0 To UBoundSafe(aLine)
            sLine = aLine(i)
        
            sFile = vbNullString
            If (StrBeginWith(sLine, "<imagepath>")) Then
                sLine = mid$(sLine, Len("<imagepath>") + 1)
                pos = InStr(1, sLine, "</imagepath>", 1)
                If pos <> 0 Then
                    sFile = Left$(sLine, pos - 1)
                End If
            ElseIf (StrBeginWith(sLine, "<launchstring>")) Then
                sLine = mid$(sLine, Len("<launchstring>") + 1)
                pos = InStr(1, sLine, "</launchstring>", 1)
                If pos <> 0 Then
                    sFile = Left$(sLine, pos - 1)
                    sFile = Replace(sFile, "&#34;", """")
                    SplitIntoPathAndArgs sFile, sFile, , True
                    sFile = UnQuote(sFile)
                    sFile = EnvironW(sFile)
                End If
            End If
            If Len(sFile) <> 0 Then
                sFile = FindOnPath(sFile)
                If Len(sFile) <> 0 Then
                    sFile = GetLongPath(sFile)
                    sFile = GetFullPath(sFile)
                    If Not dRunFiles.Exists(sFile) Then dRunFiles.Add sFile, 0
                End If
            End If
        Next
    End If
    
    DoEvents
    AppendErrorLogCustom "frmMain.ParseFilesXML - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_CheckAutoLog"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckAutoLog()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.CheckAutoLog - Begin"
    
    If Not bAutoLogSilent Then LockInterfaceMain bDoUnlock:=True
    If Not bAutoLogSilent Then DoEvents
    
    If Not gNoGUI Then
        If Not bAutoLogSilent Then DoEvents
        Me.Show
        If Not bAutoLogSilent Then DoEvents
        Me.Refresh
        DoEvents
        
        If Not bAutoLogSilent Then
            If (chkSkipIntroFrame.Value = 1) Then
                If cmdScan.Visible And cmdScan.Enabled Then
                    cmdScan.SetFocus
                End If
            Else
                If cmdN00bLog.Visible And cmdN00bLog.Enabled Then
                    cmdN00bLog.SetFocus
                End If
            End If
        End If
    Else
        Me.Hide
    End If
    
    If bAutoLog Then
        cmdN00bClose_Click
        cmdScan_Click
        If Not bAutoLogSilent Then DoEvents
        If bAutoLogSilent Then Unload Me: Exit Sub
    End If
    
    If bStartupScan Then
        cmdN00bClose_Click
        cmdScan_Click
        If lstResults.ListCount = 0 Then
            Unload Me: Exit Sub
        Else
            Me.Show
            Call pvSetVisionForLabelResults
        End If
    End If
    
    '/StartupList
    If HasCommandLineKey("StartupList") Then
        bStartupListSilent = True
        cmdN00bTools_Click
        Call chkConfigTabs_Click(3)
        cmdStartupList_Click
    End If
    
    '/SysTray
    If HasCommandLineKey("SysTray") Then
        bMinToTray = True
        Me.WindowState = vbMinimized
    End If
    
    AppendErrorLogCustom "frmMain.CheckAutoLog - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_CheckAutoLog"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub LoadResources()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.LoadResources - Begin"
    
    Dim Lines()     As String
    Dim sBuf        As String
    Dim i           As Long
    Dim Columns()   As String
    Dim id          As Long
    
    'Task Scheduler white list
    sBuf = LoadEncryptedResFile("database\TasksWhite.csv", 101, Not inIDE)
    sBuf = Replace$(sBuf, vbCr, vbNullString)
    
    Lines = Split(sBuf, vbLf)
    ReDim g_TasksWL(UBound(Lines))
    
    For i = 1 To UBound(Lines)  'skip header
    
        If 0 <> Len(Lines(i)) Then

            Lines(i) = Replace$(Lines(i), "\;", "\\\\")
            Columns = SplitSafe(Lines(i), ";")
            '---------------------------
            'Columns (0) 'OSver
            'Columns (1) 'Dir\Name
            'Columns (2) 'RunObj
            'Columns (3) 'Args
            'Columns (4) 'Note      (not used)
            '---------------------------
            
            'If InStr(Lines(i), "RegIdleBackup") Then Stop
            
            With g_TasksWL(i)
                .OSver = Val(Columns(0))
                If .OSver = 11 Then .OSver = 10

                'select appropriate version from DB
                If .OSver = OSver.MajorMinor Then
                    
                    .Path = Trim$(Columns(1))
                    If UBound(Columns) > 1 Then
                        .RunObj = EnvironW(Replace$(Trim$(Columns(2)), "\\\\", ";"))
                        If Not isCLSID(.RunObj) Then
                            If InStr(.RunObj, "\") = 0 Then
                                'find full path for relative name
                                'filename without full path can be used in database to do comparision by filename only (see: isInTasksWhiteList())
                                .RunObj = FindOnPath(.RunObj, True)
                            End If
                        End If
                    End If
                    
                    If UBound(Columns) > 2 Then .Args = EnvironW(Replace$(Trim$(Columns(3)), "\\\\", ";"))
                    
                    'Dictonary 'oDict.TaskWL_ID':
                    'value -> (dir + name of task)
                    'data -> id to 'g_TasksWL' user type array

                    If Not oDict.TaskWL_ID.Exists(.Path) Then
                        oDict.TaskWL_ID.Add .Path, i
                    Else 'append several lines with same paths
                        id = oDict.TaskWL_ID(.Path)
                        
                        'additional check in case 'FindOnPath' didn't find executable
                        g_TasksWL(id).RunObj = g_TasksWL(id).RunObj & IIf(Len(g_TasksWL(id).RunObj) = 0, vbNullString, "|") & .RunObj
                        g_TasksWL(id).Args = g_TasksWL(id).Args & "|" & .Args
                    End If
                End If
            End With
        End If
    Next
    
    AppendErrorLogCustom "frmMain.LoadResources - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.LoadResources"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    If bAutoLogSilent And UnloadMode <> 1 Then
        Cancel = True
        Exit Sub
    End If
    
    Dim sReason As String
    
    Select Case UnloadMode
    
        Case 0
            sReason = "The user choose the Close command from the Control menu on the form."
        Case 1
            sReason = "The Unload statement is invoked from code."
        Case 2
            sReason = "The current Microsoft Windows operating environment session is ending."
        Case 3
            sReason = "The Microsoft Windows Task Manager is closing the application."
        Case 4
            sReason = "An MDI child form is closing because the MDI form is closing."
        Case 5
            sReason = "A form is closing because its owner is closing."
    End Select
    
    AppendErrorLogCustom "(!!!) Form_QueryUnload initiated (!!!) - Reason: " & UnloadMode & " - " & sReason
    
    If (UnloadMode = 0 Or bmnuExit_Clicked) Then 'initiated by user (clicking 'X')
        If isRanHJT_Scan Then
            'Scanning is not finished yet! Are you really sure you want to forcibly close the program?
            If MsgBoxW(Translate(1010), vbExclamation Or vbYesNo) = vbNo Then
                Cancel = True
                Exit Sub
            End If
            AppendErrorLogCustom "User clicked 'X' while scanning and agreed with forced closing of the program."
            g_ExitCodeProcess = 1067
        End If
    End If
    g_bAppShutdown = True
    BackupFlush
    If g_WER_Disabled Then DisableWER bRevert:=True
    
    Dim frm As Form
    ToggleWow64FSRedirection True
    If Not g_UninstallState Then
        SaveSettings
        SaveWindowPos Me, SETTINGS_SECTION_MAIN
        
        RegSaveHJT "Version", AppVerString
    End If
    SubClassScroll False
    'UnregisterHotKey Me.hwnd, HOTKEY_ID_CTRL_A
    'UnregisterHotKey Me.hwnd, HOTKEY_ID_CTRL_F
    For Each frm In Forms
        If Not (frm Is Me) And Not (frm.Name = "frmEULA") Then
            Unload frm
            Set frm = Nothing
        End If
    Next
    
    If (UnloadMode = 0 Or bmnuExit_Clicked) And isRanHJT_Scan Then End
    If hLibPcre2 <> 0 Then FreeLibrary hLibPcre2: hLibPcre2 = 0
    With oDict
        Set .TaskWL_ID = Nothing
        Set .dSafeProtocols = Nothing
        Set .dSafeFilters = Nothing
    End With
    MenuReleaseIcons
    'Set HE = Nothing
    
    'because can still be used by StartupList2
    'Set Reg = Nothing
    
    g_HwndMain = 0
End Sub

Public Sub ReleaseMutex()
    If g_hMutex <> 0 Then CloseHandle g_hMutex: g_hMutex = 0
End Sub

Private Sub Form_Terminate()
    Set frmStartupList2 = Nothing
    Set ErrLogCustomText = Nothing
    Set oDictFileExist = Nothing
    
    If g_ExitCodeProcess <> 0 Then
        ExitProcess g_ExitCodeProcess
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim s$
    ReleaseMutex
    ISL_Dispatch
    Close
    SetFontDefaults Nothing, True
    If g_hDebugLog <> 0 Then
        s = vbCrLf & "--" & vbCrLf & "Debug log closed because main form is terminated (!!!)"
        PutW_NoLog g_hDebugLog, 1, StrPtr(s), LenB(s), True
        CloseHandle g_hDebugLog: g_hDebugLog = 0
    End If
    g_HwndMain = 0
End Sub

Private Sub cmdADSSpy_Click() 'Misc Tools -> ADS Spy
    frmADSspy.Show
End Sub

Private Sub mnuHelpManualBasic_Click()  'Help -> User's manual -> Basic manual
    'cmdN00bHJTQuickStart_Click
    PopupMenu mnuBasicManual
End Sub

Private Sub mnuHelpManualCmdKeys_Click()   'Help -> User's manual -> Command line keys
    cmdN00bClose_Click
    ' Закрыть окно "Инструменты"
    'If cmdConfig.Caption = Translate(19) Then cmdConfig_Click
    If cmdConfig.Tag = "1" Then cmdConfig_Click
    'If cmdHelp.Caption = Translate(16) Then cmdHelp_Click
    If cmdHelp.Tag = "0" Then cmdHelp_Click
    fraHelp.Visible = True
    fraHelp.ZOrder 0
    chkHelp_Click 1
End Sub

Private Sub mnuHelpManualSections_Click()   'Help -> User's manual -> Sections' description
    cmdN00bClose_Click
    ' Закрыть окно "Инструменты"
    'If cmdConfig.Caption = Translate(19) Then cmdConfig_Click
    If cmdConfig.Tag = "1" Then cmdConfig_Click
    'If cmdHelp.Caption = Translate(16) Then cmdHelp_Click
    If cmdHelp.Tag = "0" Then cmdHelp_Click
    fraHelp.Visible = True
    fraHelp.ZOrder 0
    chkHelp_Click 0
End Sub

'Help - Report a bug
Private Sub mnuHelpReportBug_Click()
    'HiJackThis Fork и вопросы к разработчикам
    'HJT: Main discussion thread - improvement & development & news
    OpenURL "https://github.com/dragokas/hijackthis/issues/4", "https://safezone.cc/threads/28770/"
End Sub

Private Sub pictLogo_Click()
    'Visit product description page?
    If MsgBox(Translate(1016), vbQuestion Or vbYesNo, g_AppName) = vbYes Then
        OpenURL "https://github.com/dragokas/hijackthis", "https://safezone.cc/resources/hijackthis-fork.201/", True 'by current lang.
    End If
End Sub

'AnalyzeThis
Private Sub cmdAnalyze_Click()
    'open instruction on how to prepare logs for 'Issue' section on GitHub to ask for help in PC cure
    OpenURL "https://github.com/dragokas/hijackthis/wiki/How-to-make-a-request-for-help-in-the-PC-cure-section%3F", "https://safezone.cc/pravila/"
End Sub

Private Sub cmdARSMan_Click() 'Misc Tools -> Uninstall Manager
    frmUninstMan.Show
End Sub

Private Sub cmdDigiSigChecker_Click() 'Misc Tools -> Digital signature checker
    frmCheckDigiSign.Show
End Sub

Private Sub cmdLnkChecker_Click() 'Misc Tools -> Check Browsers' LNK
    mnuToolsShortcutsChecker_Click
End Sub

Private Sub cmdLnkCleaner_Click() 'Misc Tools -> ClearLNK
    mnuToolsShortcutsFixer_Click
End Sub

Private Sub cmdRegKeyUnlocker_Click() 'Misc Tools -> Registry Keys unlocker
    frmUnlockRegKey.Show
End Sub

Private Sub cmdDeleteService_Click() 'Misc Tools -> Delete service ...
    If Not bIsWinNT Then Exit Sub
    Dim sServiceName$, sDisplayName$, sFile$, sCompany$, sTmp$, sDllPath$, sBuf$
    Dim result As SCAN_RESULT
    
'    sServiceName = InputBox("Enter the exact service name as it appears in the scan results,
'    or the short name between brackets if that is listed.", "Delete a Windows NT Service")
    sServiceName = InputBox(Translate(113), Translate(114))

    If Len(sServiceName) = 0 Then Exit Sub
    
    If Not IsServiceExists(sServiceName) Then
        sTmp = GetServiceNameByDisplayName(sServiceName)
        If Len(sTmp) <> 0 Then
            sServiceName = sTmp
        Else
            MsgBoxW Replace$(Translate(115), "[]", sServiceName), vbExclamation
'           msgboxW "Service '" & sServiceName & "' was not found in the Registry." & vbCrLf & _
'               "Make sure you entered the name of the service correctly.", vbExclamation
            Exit Sub
        End If
    End If
    
    sFile = CleanServiceFileName(GetServiceImagePath(sServiceName), sServiceName)
    sDllPath = GetServiceDllPath(sServiceName)
    sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "DisplayName")
    If Left$(sDisplayName, 1) = "@" Then 'extract string resource from file
        sBuf = GetStringFromBinary(, , sDisplayName)
        If 0 <> Len(sBuf) Then sDisplayName = sBuf
    End If
    
    sCompany = GetFilePropCompany(IIf(Len(sDllPath) <> 0, sDllPath, sFile))
    If sCompany = vbNullString Then sCompany = Translate(502) '"Unknown owner" '"?"
    
    If Not FileExists(sFile) Then sFile = sFile & " (" & Translate(503) & ")"  '" " & STR_FILE_MISSING
    
    If MsgBoxW(Translate(117) & vbCrLf & _
              Translate(505) & ": " & sServiceName & vbCrLf & _
              Translate(506) & ": " & sDisplayName & vbCrLf & _
              Translate(507) & ": " & sFile & IIf(Len(sDllPath) <> 0, " -> " & sDllPath, vbNullString) & vbCrLf & _
              Translate(508) & ": " & sCompany & vbCrLf & vbCrLf & _
              Translate(118), vbYesNo Or vbDefaultButton2 Or vbExclamation) = vbYes Then
'    If msgboxW("The following service was found:" & vbCrLf & _
'              "Short name: " & sServiceName & vbCrLf & _
'              "Full name: " & sDisplayName & vbCrLf & _
'              "File: " & sFile & vbCrLf & _
'              "Owner: " & sCompany & vbCrLf & vbCrLf & _
'              "Are you absolutely sure you want to delete this service?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
        
        With result
            .Section = "O23"
            .HitLineW = "O23 - Service: " & sServiceName & " (" & sDisplayName & ")"
            AddServiceToFix .Service, DELETE_SERVICE, sServiceName
            
            If Len(sDllPath) = 0 Then
                AddFileToFix .File, BACKUP_FILE, sFile 'WARNING !!! Do not delete file (like "sweeping") here, because "ForceMicrosoft" mode is activated! You can remove svchost.exe file by mistake !!!
            Else
                AddFileToFix .File, BACKUP_FILE, sDllPath
            End If
            AddRegToFix .Reg, BACKUP_KEY, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName
            
            .ForceMicrosoft = True
            .Reboot = True
            .CureType = SERVICE_BASED Or FILE_BASED Or REGISTRY_BASED
        End With
        
        LockInterface bAllowInfoButtons:=False, bDoUnlock:=False
        IncreaseNumberOfFixes
        IncreaseFixID
        MakeBackup result
        BackupFlush
        
        FixIt result
        
        BackupFlush
        LockInterface False, True
        
        bRebootRequired = True
        RestartSystem
    End If
End Sub

Private Sub cmdDelOnReboot_Click() 'Misc Tools -> Delete on reboot ...
    Dim sFilename$
'    'Enter file name:, Delete on Reboot
'    sFilename = InputBox(Translate(1950), Translate(1951))
'    If StrPtr(sFilename) = 0 Then Exit Sub
    
    'Delete on Reboot
    sFilename = OpenFileDialog(Translate(1951), Desktop, _
        Translate(1003) & " (*.*)|*.*|" & Translate(1956) & " (*.dll)|*.dll|" & Translate(1957) & " (*.exe)|*.exe", Me.hWnd)
    If Len(sFilename) = 0 Then Exit Sub
    
    DeleteFileOnReboot sFilename, True, True
End Sub

Private Sub cmdHostsManager_Click() 'Misc Tools -> 'Hosts' file manager
    fraConfigTabs(3).Visible = False
    fraHostsMan.Visible = True
    NotifyChangeFrame FRAME_ALIAS_HOSTS
    ListHostsFile lstHostsMan
End Sub

'Hosts -> Back
Private Sub cmdHostsManBack_Click()
    fraHostsMan.Visible = False
    fraConfigTabs(3).Visible = True
    NotifyChangeFrame FRAME_ALIAS_MISC_TOOLS
End Sub

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

'Scan results window - switch to "Main menu"
Private Sub cmdMainMenu_Click()

    txtNothing.Visible = False
    NotifyChangeFrame FRAME_ALIAS_MAIN

    CloseProgressbar
    
    frmMain.pictLogo.Visible = True
    'If cmdConfig.Caption = Translate(19) Then 'Report
    
    If cmdConfig.Tag = "1" Then ' Moved from 'Settings' or custom frame (not 'Scan Results')
    
        AppendErrorLogCustom "SaveSettings initiated by clicking 'Main menu'."
        SaveSettings
        
        fraConfig.Visible = False
        fraHostsMan.Visible = False
        If chkConfigTabs(3).Value = 1 Then fraConfigTabs(3).Visible = True
        cmdConfig.Caption = Translate(1089): cmdConfig.Tag = "0" 'Settings
        'cmdHelp.Enabled = True
        cmdSaveDef.Enabled = True
        fraScan.Enabled = True
        cmdScan.Enabled = True
        cmdFix.Enabled = True
        cmdInfo.Enabled = True
    End If
    
    fraHelp.Visible = False
    fraN00b.Visible = True
    fraScan.Visible = False
    fraOther.Visible = False
    lstResults.Visible = False
    fraSubmit.Visible = False
    'cmdScan.Caption = Translate(11) ' don't touch it !!!
    'cmdScan.Tag = "1"
    'cmdHelp.Caption = Translate(16)
    lblInfo(0).Visible = True
    lblInfo(1).Visible = False
    shpMD5Progress.Visible = False
    shpMD5Background.Visible = False
    lblMD5.Visible = False
    'chkSkipIntroFrame.Value = RegReadHJT("SkipIntroFrame", "0")
End Sub

'// Main menu - List of Backups
Private Sub cmdN00bBackups_Click()
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    cmdConfig_Click
    chkConfigTabs_Click 2
End Sub

'// Main menu - None of above, just start the program (deprecated)
Private Sub cmdN00bClose_Click()
    NotifyChangeFrame FRAME_ALIAS_SCAN
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    'If cmdHelp.Caption = Translate(17) Then 'Back
    If cmdHelp.Tag = "1" Then 'Back
        cmdHelp_Click
    End If
    Call pvSetVisionForLabelResults '"Welcome to HJT" / or "Below are the results..."
    If cmdScan.Visible And cmdScan.Enabled Then
        cmdScan.SetFocus
    End If
    If g_bCalcHashInProgress Then
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        ResumeHashProgressbar
    End If
End Sub

'// Main menu - Settings
Private Sub cmdSettings_Click()
    mnuFileSettings_Click
End Sub

'// Main menu - Tutorial & Support
Private Sub cmdN00bHJTQuickStart_Click()
    PopupMenu mnuSupportContext
End Sub

'// Tutorial & Support - Offline tutorial
Private Sub mnuSupportOffline_Click()
    cmdN00bClose_Click
    cmdHelp_Click
End Sub

'// Tutorial & Support - Online tutorial
Private Sub mnuSupportOnline_Click()
    OpenURL "https://dragokas.com/tools/help/hjt_tutorial.html", "https://regist.safezone.cc/hijackthis_help/hijackthis.html", True
End Sub

'// Tutorial & Support - Cure forums
Private Sub mnuSupportCureGithub_Click()
    OpenURL "https://github.com/dragokas/hijackthis/wiki/How-to-make-a-request-for-help-in-the-PC-cure-section%3F"
End Sub
Private Sub mnuSupportCyberforum_Click()
    'https://www.cyberforum.ru
    OpenURL Caes_Decode("iwywB://NPR.BZEJYOZ_d^.gl") & "/viruses/thread49792.html"
End Sub
Private Sub mnuSupportKaspersky_Click()
    ' https://forum.kasperskyclub.ru
    OpenURL Caes_Decode("iwywB://wHMRL.NFZYP_b\lXcn].qv") & "/topic/43640-porjadok-oformlenija-zaprosa-o-pomoshhi/"
End Sub
Private Sub mnuSupportOszone_Click()
    'http://forum.oszone.net
    OpenURL Caes_Decode("iwyw://uFKPJ.PV_VWP.]Vg") & "/thread-98169.html"
End Sub
Private Sub mnuSupportSafezone_Click()
    OpenURL "https://www.safezone.cc/pravila/"
End Sub
Private Sub mnuSupportSoftboard_Click()
    'https://softboard.ru
    OpenURL Caes_Decode("iwywB://JHAQAPDWK.]b") & "/topic/51343-%D0%BF%D1%80%D0%B0%D0%B2%D0%B8%D0%BB%D0%B0-%D0%BF%D0%BE%D0%B4%D1%80%D0%B0%D0%B7%D0%B4%D0%B5%D0%BB%D0%B0/"
End Sub
Private Sub mnuSupportTHG_Click()
    'http://www.thg.ru
    OpenURL Caes_Decode("iwyw://LNP.QGH.W\") & "/forum/showthread.php?t=92236"
End Sub
Private Sub mnuSupportVirusinfo_Click()
    'https://virusinfo.info
    OpenURL Caes_Decode("iwywB://MBMRRJQKV.T[U`") & "/showthread.php?t=1235"
End Sub

Private Sub MoveToScanResultsFrame()
    NotifyChangeFrame FRAME_ALIAS_SCAN
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
End Sub

'// Main menu - Do a system scan and save a log file
Private Sub cmdN00bLog_Click()
    
    MoveToScanResultsFrame
    
    If isRanHJT_Scan Then
        SetProgressBarOnFront
        Exit Sub
    End If
    
    cmdScan.Caption = Translate(11): cmdScan.Tag = "1" 'don't touch this!!!
    
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    bAutoLog = True
    cmdScan_Click
End Sub

'// Main menu - Do a system scan only
Private Sub cmdN00bScan_Click()

    MoveToScanResultsFrame

    If isRanHJT_Scan Then
        SetProgressBarOnFront
        Exit Sub
    End If
    
    cmdScan.Caption = Translate(11): cmdScan.Tag = "1" 'don't touch this!!!
    
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    cmdScan_Click
End Sub

'// Main menu - Misc Tools
Private Sub cmdN00bTools_Click()
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    
    'lstResults.Visible = True
    
    'If cmdConfig.Caption = Translate(1089) Then cmdConfig_Click
    
    cmdConfig.Caption = Translate(1089): cmdConfig.Tag = "0"
    cmdConfig_Click
    chkConfigTabs_Click 3
End Sub

Private Sub chkConfigTabs_Click(Index As Integer)

    On Error GoTo ErrorHandler:
    
    Static idxLastTab As Long
    Static isInit As Boolean
    
    Dim i           As Long
    Dim iIgnoreNum  As Long
    Dim sIgnore     As String
    
    If bSwitchingTabs Then Exit Sub
    If frmMain.cmdHidden.Visible And frmMain.cmdHidden.Enabled Then
        frmMain.cmdHidden.SetFocus
    End If
    bSwitchingTabs = True
    
    If Not isInit Then isInit = True
    
    chkConfigTabs(0).Value = 0
    chkConfigTabs(1).Value = 0
    chkConfigTabs(2).Value = 0
    chkConfigTabs(3).Value = 0
    chkConfigTabs(Index).Value = 1
    
    fraConfigTabs(0).Visible = False
    fraConfigTabs(1).Visible = False
    fraConfigTabs(2).Visible = False
    fraConfigTabs(3).Visible = False
    fraConfigTabs(Index).Visible = True
    
    fraHostsMan.Visible = False
    
    bSwitchingTabs = False
    fraConfig.Visible = True
    
    Select Case Index
    
    Case 0 'main settings
        NotifyChangeFrame FRAME_ALIAS_SETTING
        
    Case 1 'ignore list
        NotifyChangeFrame FRAME_ALIAS_IGNORE_LIST
        
        lstIgnore.Clear
        iIgnoreNum = CInt(RegReadHJT("IgnoreNum", "0"))
        If iIgnoreNum > 0 Then
            For i = 1 To iIgnoreNum
                sIgnore = DeCrypt(RegReadHJT("Ignore" & CStr(i), vbNullString))
                If sIgnore <> vbNullString Then
                    lstIgnore.AddItem sIgnore
                Else
                    Exit For
                End If
            Next i
        End If
        lstIgnore.ListIndex = -1
        AddHorizontalScrollBarToResults lstIgnore
        
    Case 2 'backups
        NotifyChangeFrame FRAME_ALIAS_BACKUPS
        ListBackups
        
    Case 3 'Misc tools
        NotifyChangeFrame FRAME_ALIAS_MISC_TOOLS
        
    End Select
    
    idxLastTab = Index
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "chkConfigTabs_Click", "idx:" & Index
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdConfig_Click()
    On Error GoTo ErrorHandler:

    'сперва выйти из фрейма "Help"
    'If cmdHelp.Caption = Translate(17) Then cmdHelp_Click
    If cmdHelp.Tag = "1" Then cmdHelp_Click
    
    CloseProgressbar
    
    'If cmdConfig.Caption = Translate(1089) Then   'Settings
    If cmdConfig.Tag = "0" Then
        NotifyChangeFrame FRAME_ALIAS_SETTING
        pictLogo.Visible = False
        'chkSkipIntroFrameSettings.Value = CLng(RegReadHJT("SkipIntroFrame", "0"))
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        lstResults.Visible = False
        cmdConfig.Caption = Translate(19): cmdConfig.Tag = "1"
        cmdSaveDef.Enabled = False
        fraScan.Enabled = False
        cmdScan.Enabled = False
        cmdFix.Enabled = False
        cmdInfo.Enabled = False
        txtNothing.ZOrder 1
        txtNothing.Visible = False
        cmdAnalyze.Enabled = False
        
        'fraConfigTabs(0).Visible = True
        'fraConfig.Visible = True
        'chkConfigTabs(0).Value = 1
        
        chkConfigTabs_Click 0
        
    Else                            'Back
        
        NotifyChangeFrame FRAME_ALIAS_SCAN
        Call pvSetVisionForLabelResults '"Welcome to HJT" / or "Below are the results..."
        lstResults.Visible = True
        fraHostsMan.Visible = False
        If chkConfigTabs(3).Value = 1 Then fraConfigTabs(3).Visible = True
        cmdConfig.Caption = Translate(1089): cmdConfig.Tag = "0"
        cmdSaveDef.Enabled = True
        cmdInfo.Enabled = True
        fraConfig.Visible = False
        fraScan.Enabled = True
        
        If Not isRanHJT_Scan Then
            cmdScan.Enabled = True

            If lstResults.ListCount > 0 Then
                cmdAnalyze.Enabled = True
                cmdFix.Enabled = True
            End If
        End If
        
        AppendErrorLogCustom "SaveSettings initiated by clicking 'Back' button."
        
        SaveSettings
    End If

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfig_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdConfigBackupDeleteAll_Click()
    If lstBackups.ListCount = 0 Then Exit Sub
    'If msgboxW("Are you sure you want to delete ALL backups?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    If MsgBoxW(Translate(88), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'    If msgboxW("Delete all backups? Are you sure? I mean, " & _
'    "like, ALL of them will be gone! I didn't put in " & _
'    "this backup thingy just for fun, I never use this " & _
'    "kind of stuff. But hey, if _you_ want to do it, just " & _
'    "click Yes. But you never know when you're going to " & _
'    "need them - maybe a day or two from now you think " & _
'    "'I'm sure I had a sample of that bugger, if only I " & _
'    "could find it and email it to McAfee, since it has " & _
'    "now been classified a virus'. Or you meet someone on " & _
'    "SpywareInfo.com that wants to take that porn DLL " & _
'    "apart and see what makes it tick." & vbCrLf & vbCrLf & _
'    "Ah crap. I get carried away and look what I did. " & _
'    "Never mind." & vbCrLf & vbCrLf & "Are you sure you " & _
'    "want to delete all backups?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    DeleteBackup vbNullString, True
    lstBackups.Clear
    ListBackups
End Sub

Private Sub cmdConfigBackupDelete_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    If lstBackups.ListIndex = -1 Then Exit Sub
    If lstBackups.CheckedIndices.Count = 0 Then
        'First you have to mark a checkbox next to at least one item!
        MsgBox Translate(554), vbInformation
        Exit Sub
    End If
    If lstBackups.CheckedIndices.Count = 1 Then
        If MsgBoxW(Translate(84), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    '    If msgboxW("Are you sure you want to delete this backup?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        If MsgBoxW(Replace$(Translate(85), "[]", lstBackups.CheckedIndices.Count), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If msgboxW("Are you sure you want to delete these " & lstBackups.CheckedIndices.Count & " backups?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    For i = lstBackups.ListCount - 1 To 0 Step -1
        If lstBackups.ItemChecked(i) Then
            DeleteBackup lstBackups.List(i)
            lstBackups.RemoveItem i
        End If
    Next i
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfigBackupDelete_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub chkShowSRP_Click()
    bShowSRP = chkShowSRP.Value
    RegSaveHJT "ShowSRP", Abs(CLng(bShowSRP))
    ListBackups 'update list
End Sub

Private Sub cmdConfigBackupRestore_Click()
    On Error GoTo ErrorHandler:
    Dim i&, j&
    Dim sDecription As String
    Dim sBackupLine As String
    Dim lBackupID As Long
    Dim lstIdx As Long
    Dim aLines() As String
    
    If lstBackups.ListIndex = -1 Then Exit Sub
    If lstBackups.CheckedIndices.Count = 0 Then
        'First you have to mark a checkbox next to at least one item!
        MsgBox Translate(554), vbInformation
        Exit Sub
    End If
    
    If bRebootRequired Or ("1" = RegReadHJT("RebootRequired", "0")) Then
        'Cannot start restoring until the system will be rebooted!
        MsgBoxW Translate(1577), vbExclamation
        RestartSystem
        Exit Sub
    End If
    
    If lstBackups.CheckedIndices.Count = 1 Then
        'exclude question for ABR / SRP backups (it has inividual message)
        BackupSplitLine lstBackups.List(GetListBoxSelectedItemID(lstBackups)), , , , sDecription
        If sDecription <> ABR_BACKUP_TITLE _
          And Not StrBeginWith(sDecription, SRP_BACKUP_TITLE) Then
            If MsgBoxW(Translate(86), vbQuestion + vbYesNo) = vbNo Then Exit Sub
            'If msgboxW("Restore this item?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    Else
        If MsgBoxW(Replace$(Translate(87), "[]", lstBackups.CheckedIndices.Count), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If msgboxW("Restore these " & lstBackups.CheckedIndices.Count & " items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    'cache selected lines (to account for the shifting of elements in the list)
    ReDim Preserve aLines(lstBackups.ListCount - 1)
    j = 0
    For i = 0 To lstBackups.ListCount - 1   'vice versa order (list is already grouped vice versa)
        'only marked with checkbox
        If lstBackups.ItemChecked(i) Then
            aLines(j) = lstBackups.List(i)
            j = j + 1
        End If
    Next
    ReDim Preserve aLines(j - 1)
    
    'list cached lines
    For i = 0 To UBound(aLines)
        sBackupLine = aLines(i)
        'if restore is success
        If RestoreBackup(sBackupLine) Then  '<<< ACTUAL restoring
            'not ABR / SRP backups ?
            BackupSplitLine sBackupLine, lBackupID, , , sDecription
            If sDecription <> ABR_BACKUP_TITLE _
              And Not StrBeginWith(sDecription, SRP_BACKUP_TITLE) Then
                DeleteBackup sBackupLine
                lstIdx = GetListIndexByBackupID(lBackupID)
                If lstIdx <> -1 Then
                    lstBackups.RemoveItem lstIdx
                End If
            End If
        Else
            'Unknown error happened during restore item: []. Item is restored partially only.
            MsgBoxW Replace$(Translate(1574), "[]", sBackupLine), vbExclamation
        End If
    Next i
    
    BackupFlush
    'result list need to be flushed, because new "infected" items will appear after restoring from backup
    lstResults.Clear
    cmdScan.Caption = Translate(11): cmdScan.Tag = "1" 'Scan
    
    If bRebootRequired Then RestartSystem
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfigBackupRestore_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Function GetListBoxSelectedItemID(lst As ListBox) As Long
    Dim i&
    For i = 0 To lst.ListCount - 1
        If lstBackups.ItemChecked(i) Then
            GetListBoxSelectedItemID = i
        End If
    Next i
End Function

Private Sub cmdConfigIgnoreDelAll_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    If lstIgnore.ListCount = 0 Then Exit Sub
    'Are you sure?" & vbCrLf & "This will delete ALL ignore list.
    If vbNo = MsgBoxW(Translate(73), vbYesNo) Then Exit Sub
    RegSaveHJT "IgnoreNum", 0
    For i = 0 To lstIgnore.ListCount - 1
        RegDelHJT "Ignore" & CStr(i + 1)
    Next i
    lstIgnore.Clear
    IsOnIgnoreList vbNullString, UpdateList:=True
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfigIgnoreDelAll_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdConfigBackupCreateRegBackup_Click()
    'Run ABR to backup FULL registry
    cmdConfigBackupCreateRegBackup.Enabled = False
    If ABR_CreateBackup(True) Then
        'Full registry backup is successfully created.
        MsgBoxW Translate(1567), vbInformation
    End If
    cmdConfigBackupCreateRegBackup.Enabled = True
End Sub

Private Sub cmdConfigBackupCreateSRP_Click()
    'Create System Restore Point
    Dim nSeqNum As Long
    cmdConfigBackupCreateSRP.Enabled = False
    nSeqNum = SRP_Create_API()
    If nSeqNum <> 0 And bShowSRP Then
        frmMain.lstBackups.AddItem _
            BackupConcatLine(0&, 0&, BackupFormatDate(Now()), SRP_BACKUP_TITLE & " - " & nSeqNum & " - " & "Restore Point by HijackThis+"), 0
    End If
    cmdConfigBackupCreateSRP.Enabled = True
    'Note: that actual restore point record will appear in the WMI list after ~ 15 sec.
End Sub

Private Sub cmdConfigIgnoreDelSel_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    If lstIgnore.ListIndex = -1 Then Exit Sub
    If lstIgnore.CheckedIndices.Count = 0 Then
        'First you have to mark a checkbox next to at least one item!
        MsgBox Translate(554), vbInformation
        Exit Sub
    End If
    For i = 0 To lstIgnore.ListCount - 1
        RegDelHJT "Ignore" & CStr(i + 1)
    Next i
    For i = lstIgnore.ListCount - 1 To 0 Step -1
        If lstIgnore.ItemChecked(i) Then lstIgnore.RemoveItem i
    Next i
    RegSaveHJT "IgnoreNum", lstIgnore.ListCount
    For i = 0 To lstIgnore.ListCount - 1
        RegSaveHJT "Ignore" & CStr(i + 1), Crypt(lstIgnore.List(i))
    Next i
    IsOnIgnoreList vbNullString, UpdateList:=True
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfigIgnoreDelSel_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub IncreaseNumberOfFixes()
    On Error GoTo ErrorHandler:

    Dim dLastFix    As Date
    Dim dNow        As Date
    Dim dMidNight   As Date
    Dim lNumFixes   As Long
    Dim sTime       As String
    
    dNow = Now()
    dMidNight = GetDateAtMidnight(dNow)
    
    sTime = RegReadHJT("DateLastFix", vbNullString)
    
    If Len(sTime) <> 0 Then
        If StrBeginWith(sTime, "HJT:") Then
            sTime = mid$(sTime, 6)
            dLastFix = CDateEx(sTime, 1, 6, 9, 12, 15, 18)
        Else 'backward support
            If IsDate(sTime) Then
                dLastFix = CDate(sTime)
            End If
        End If
    End If
    
    lNumFixes = CLng(RegReadHJT("FixesToday", "0"))
    
    If lNumFixes = 0 Then
        RegSaveHJT "FixesToday", CStr(1)
    ElseIf dLastFix < dMidNight Then
        RegSaveHJT "FixesToday", CStr(1)
    Else
        RegSaveHJT "FixesToday", CStr(lNumFixes + 1)
    End If
    
    RegSaveHJT "DateLastFix", "HJT: " & Format$(dNow, "yyyy\/MM\/dd HH:nn:ss")
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "IncreaseNumberOfFixes"
    If inIDE Then Stop: Resume Next
End Sub

'// Fix Checked
Private Sub cmdFix_Click()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.cmdFix_Click - Begin"

    Dim i&, j&, sPrefix$, pos&, sItem$, sSubSection$
    Dim bFlushDNS As Boolean
    Dim bO24Fixed As Boolean
    Dim bO14Fixed As Boolean
    Dim bRestartExplorer As Boolean

    Dim result As SCAN_RESULT
    
    If lstResults.CheckedIndices.Count = 0 Then
'        If MsgBoxW(Translate(344), vbQuestion + vbYesNo) = vbNo Then
'        'If msgboxW("Nothing selected! Continue?", vbQuestion + vbYesNo) = vbNo Then
'            Exit Sub
'        Else
'            lstResults.Clear
'            cmdFix.FontBold = False
'            cmdFix.Enabled = False
'            'cmdScan.Caption = "Scan"
'            cmdScan.Caption = Translate(11)
'            cmdScan.FontBold = True
'            Exit Sub
'        End If

        'First you have to mark a checkbox next to at least one item!
        MsgBox Translate(554), vbInformation
        Exit Sub
    End If
    
    If lstResults.ListCount = 0 Then Exit Sub
    
    '/startupscan
    '/noGUI
    If (lstResults.ListCount = lstResults.CheckedIndices.Count) And (Not HasCommandLineKey("StartupScan")) And (lstResults.CheckedIndices.Count > 5) And _
        (Not g_bNoGUI) Then
        
        If MsgBoxW(Translate(345), vbExclamation Or vbYesNo) = vbNo Then Exit Sub
'        If msgboxW("You selected to fix everything HiJackThis has found. " & _
'                  "This could mean items important to your system " & _
'                  "will be deleted and the full functionality of your " & _
'                  "system will degrade." & vbCrLf & vbCrLf & _
'                  "If you aren't sure how to use HiJackThis, you should " & _
'                  "ask for help, not blindly fix things. The SpywareInfo " & _
'                  "forums will gladly help you with your log." & vbCrLf & vbCrLf & _
'                  "Are you sure you want to fix all items in your scan " & _
'                  "results?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    End If
    
    If bConfirm And Not HasCommandLineKey("noGUI") Then
        'lstResults.ListIndex = -1
        If MsgBoxW(Replace$(Translate(346), "[]", lstResults.CheckedIndices.Count) & _
           IIf(bMakeBackup, ".", ", " & Translate(347)), vbQuestion + vbYesNo, g_AppName) = vbNo Then Exit Sub
'        If msgboxW("Fix " & lstResults.CheckedIndices.Count & _
'         " selected items? This will permanently " & _
'         "delete and/or repair what you selected" & _
'         IIf(bMakeBackup, ".", ", unless you enable backups."), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    If Not g_bNoGUI Then
        LockInterface bAllowInfoButtons:=False, bDoUnlock:=False
    End If
    
    IncreaseNumberOfFixes 'save number of fixes for today in registry
    
    IncreaseFixID 'to track same items
    
    If Not g_bNoGUI Then
        SetProgressBar lstResults.CheckedIndices.Count + 1
        UpdateProgressBar "Backup"
    End If
    
    If bMakeBackup Then
        'Creating FULL registry backup
        ABR_CreateBackup False
    End If
    
    'shpBackground.Tag = lstResults.CheckedIndices.Count
    'shpProgress.Tag = "0"
    
    'shpProgress.Width = 15
    'bRebootRequired = False
    bUpdatePolicyNeeded = False
    bShownBHOWarning = False
    bShownToolbarWarning = False
    bSeenHostsFileAccessDeniedWarning = False
    bNeedRebuildPolicyChain = False
    
    Call GetProcesses(gProcess)
    
    For j = 0 To 1
      '0 - do backup only
      '1 - do fix
      
      If j = 1 Then BackupFlush
    
      For i = 0 To lstResults.ListCount - 1
        If lstResults.ItemChecked(i) = True Then
            lstResults.ListIndex = i
            
            sPrefix = vbNullString
            sItem = lstResults.List(i)
            pos = InStr(sItem, "-")
            If pos <> 0 Then
                sPrefix = Trim$(Left$(sItem, pos - 1))
                sSubSection = Trim$(mid$(sItem, pos + 1))
            End If
            
            If j = 0 Then
                AppendErrorLogCustom "Backup: " & sItem
            Else
                AppendErrorLogCustom "Fixing: " & sItem
            End If
            
            If GetScanResults(sItem, result) Then 'map ANSI string to Unicode
            
              If j = 0 Then
                MakeBackup result
              Else
                If Not g_bNoGUI Then
                    UpdateProgressBar sPrefix
                End If
                
                bRebootRequired = bRebootRequired Or result.Reboot
            
                Select Case sPrefix
                Case "R0", "R1", "R2": FixRegItem sItem, result
                Case "R3":             FixR3Item sItem, result
                Case "R4":             FixR4Item sItem, result
                Case "F0", "F1":       FixFileItem sItem, result
                Case "F2", "F3":       FixFileItem sItem, result
                'Case "N1", "N2", "N3", "N4": FixNetscapeMozilla sItem,Result
                Case "B":              FixBrowserItem sItem, result
                Case "O1":             FixO1Item sItem, result: bFlushDNS = True
                Case "O2":             FixO2Item sItem, result
                Case "O3":             FixO3Item sItem, result
                Case "O4":             FixO4Item sItem, result
                Case "O5":             FixO5Item sItem, result
                Case "O6":             FixO6Item sItem, result
                Case "O7":             FixO7Item sItem, result
                Case "O8":             FixO8Item sItem, result
                Case "O9":             FixO9Item sItem, result
                Case "O10":            FixLSP
                Case "O11":            FixO11Item sItem, result
                Case "O12":            FixO12Item sItem, result
                Case "O13":            FixO13Item sItem, result
                Case "O14":            If Not bO14Fixed Then FixO14Item sItem, result: bO14Fixed = True 'O14 fix uses only once
                Case "O15":            FixO15Item sItem, result
                Case "O16":            FixO16Item sItem, result
                Case "O17":            FixO17Item sItem, result: bFlushDNS = True
                Case "O18":            FixO18Item sItem, result
                Case "O19":            FixO19Item sItem, result
                Case "O20":            FixO20Item sItem, result
                Case "O21":            FixO21Item sItem, result: bRestartExplorer = True
                Case "O22":            FixO22Item sItem, result
                Case "O23":            FixO23Item sItem, result
                Case "O24":            FixO24Item sItem, result: bO24Fixed = True
                Case "O25":            FixO25Item sItem, result
                Case "O26":            FixO26Item sItem, result
                Case "O27":            FixO27Item sItem, result
                Case Else
                   ' msgboxW "Fixing of " & Rtrim$(left$(lstResults.List(i), 3)) & _
                           " is not implemented yet. Bug me about it at " & _
                           "www.merijn.org/contact.html, because I should have done it.", _
                           vbInformation, "bad coder - no donuts"
                           
                    'Fixing of [] is not implemented yet."
                    If Not g_bNoGUI Then
                        MsgBoxW Replace$(Translate(348), "[]", sPrefix, vbInformation)
                    End If
                End Select
              End If
            End If
            
        End If
      Next
    Next
    
    BackupFlush
    Reg.FlushAll
    
    If bRestartExplorer Then RestartExplorer
    If bFlushDNS Then FlushDNS
    If bNeedRebuildPolicyChain Then PolicyScripts_RebuildChain
    If bUpdatePolicyNeeded Then
        If OSver.IsWindows8OrGreater Then ' prevents gpupdate on Win 8/10 due to BSOD
            bRebootRequired = True
        Else
            UpdatePolicy
        End If
    End If
    If bO24Fixed Then FixO24Item_Post ' restart shell
    
    If Not g_bNoGUI Then
        UpdateProgressBar "Finish"
    End If
    lstResults.Clear
    
    g_bScanInProgress = False
    g_bGeneralScanned = False
    
    'if somewhere explorer.exe has been killed, but not launched
    If Not ProcessExist("explorer.exe", True) Then
        RestartExplorer
    End If
    
    Dim bShouldReboot As Boolean
    Dim bSilentReboot As Boolean
    Dim bServerReboot As Boolean
    
    If bRebootRequired Then
        If Not g_bNoGUI Then
            bShouldReboot = True
            bSilentReboot = False
            bServerReboot = False
        Else
            If HasCommandLineKey("reboot") Then
                bShouldReboot = True
                bSilentReboot = True
                bServerReboot = True
            End If
        End If
    End If
    
    If bShouldReboot Then
        RegSaveHJT "RebootRequired", 1
        RestartSystem vbNullString, bSilentReboot, bServerReboot
    End If
    
    'CloseProgressbar 'leave progressBar visible to ensure the user saw completion of cure
    
    If (Not inIDE) And (Not g_bNoGUI) Then MessageBeep MB_ICONINFORMATION
    
    If Not g_bNoGUI Then
        LockInterface bAllowInfoButtons:=True, bDoUnlock:=True
    End If
    
    cmdFix.Enabled = False
    cmdFix.Font.Bold = False
    cmdScan.Caption = Translate(11): cmdScan.Tag = "1"
    cmdScan.Font.Bold = True
    
    If cmdScan.Visible Then
        cmdScan.Enabled = True
        cmdScan.SetFocus
    End If
    
    AppendErrorLogCustom "frmMain.cmdFix_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdFix_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdHelp_Click()
    'Back
    'If cmdConfig.Caption = Translate(19) Then
    If cmdConfig.Tag = "1" Then
        cmdConfig_Click
    End If

    'If cmdHelp.Caption = Translate(16) Then 'Help
    If cmdHelp.Tag = "0" Then
        NotifyChangeFrame FRAME_ALIAS_HELP_SECTIONS
        cmdInfo.Enabled = False
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        lstResults.Visible = False
        fraHelp.Visible = True
        cmdHelp.Caption = Translate(17): cmdHelp.Tag = "1" 'Back
        'cmdConfig.Enabled = False
        cmdSaveDef.Enabled = False
        cmdScan.Enabled = False
        cmdFix.Enabled = False
        cmdAnalyze.Enabled = False
        txtNothing.ZOrder 1
        txtNothing.Visible = False
        
        fraHelp.Visible = True
        fraHelp.ZOrder 0
        chkHelp_Click 0 'help on section
    Else
        NotifyChangeFrame FRAME_ALIAS_SCAN
        Call pvSetVisionForLabelResults '"Welcome to HJT" / or "Below are the results..."
        cmdInfo.Enabled = True
        lstResults.Visible = True
        fraHelp.Visible = False
        cmdHelp.Caption = Translate(16): cmdHelp.Tag = "0" ' Info...
        'cmdConfig.Enabled = True
        cmdSaveDef.Enabled = True
        If Not isRanHJT_Scan Then
            cmdScan.Enabled = True
            If lstResults.ListCount > 0 Then
                cmdFix.Enabled = True
                cmdAnalyze.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cmdInfo_Click()
    If lstResults.Visible Then
        If lstResults.CheckedIndices.Count = 0 And lstResults.ListIndex = -1 Then
            'First you have to mark a checkbox next to at least one item!
            MsgBox Translate(554), vbInformation
            Exit Sub
        End If
        GetInfo GetSelected_OrCheckedItem
    ElseIf txtHelp.Visible Then
        GetInfo LTrim$(txtHelp.SelText)
    End If
End Sub

Private Sub cmdSaveDef_Click()
    On Error GoTo ErrorHandler:
    If lstResults.CheckedIndices.Count = 0 Then
        'First you have to mark a checkbox next to at least one item!
        MsgBox Translate(554), vbInformation
        Exit Sub
    End If
    If bConfirm Then
        If MsgBoxW(Translate(25), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'        If msgboxW("This will set HiJackThis to ignore the " & _
'                  "checked items, unless they change. Cont" & _
'                  "inue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    Dim i&, j&
    i = CInt(RegReadHJT("IgnoreNum", "0"))
    RegSaveHJT "IgnoreNum", CStr(i + lstResults.CheckedIndices.Count)
    j = i + 1
    For i = 0 To lstResults.ListCount - 1
        If lstResults.ItemChecked(i) Then
            RegSaveHJT "Ignore" & CStr(j), Crypt(lstResults.List(i))
            j = j + 1
            'sync listbox records with the RAM
            RemoveFromScanResults lstResults.List(i)
        End If
    Next i
    IsOnIgnoreList vbNullString, UpdateList:=True
    
    For i = lstResults.ListCount - 1 To 0 Step -1
        If lstResults.ItemChecked(i) Then lstResults.RemoveItem i
    Next i
    If lstResults.ListCount = 0 Then
        txtNothing.Visible = True
        txtNothing.ZOrder 0
        cmdFix.Font.Bold = False
        'cmdScan.Caption = "Scan"
        'cmdScan.Caption = Translate(11)
        'cmdScan.Tag = "1"
        'cmdScan.FontBold = True
        'If cmdScan.Visible Then
        '    If cmdScan.Enabled Then
        '        cmdScan.SetFocus
        '    End If
        'End If
    End If
    SortSectionsOfResultList
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdSaveDef_Click"
    If inIDE Then Stop: Resume Next
End Sub

' "Scan" button on "Scan results" window
Private Sub cmdScan_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    
    AppendErrorLogCustom "frmMain.cmdScan_Click - Begin"
    
    If bAutoLogSilent Then
        'LockInterface bAllowInfoButtons:=False, bDoUnlock:=False
        LockMenu bDoUnlock:=False
    End If
    
    If isRanHJT_Scan Then
        Exit Sub
    Else
        isRanHJT_Scan = True
    End If
    cmdScan.Enabled = False
    
    'cmdN00bLog.Enabled = False
    'cmdN00bScan.Enabled = False
    
    FraIncludeSections.Enabled = False
    fraScanOpt.Enabled = False
    cmdFix.Enabled = False
    cmdSaveDef.Enabled = False
    
    If cmdScan.Tag = "1" Then 'Button text == "Scan"
        
        g_bScanInProgress = True
        
        If bAutoLogSilent And Not bStartupScan Then
            Call SystemPriorityDowngrade(True)
        End If
        
        If bAutoLog Then
            If Len(g_sLogFile) = 0 Then
                g_sLogFile = BuildPath(AppPath(), "HiJackThis.log")
            End If
            DeleteFileEx g_sLogFile
        End If
        
        'first scan after rebooting ?
        bFirstRebootScan = ScanAfterReboot()
        
        ' Erase main W array of scan results
        ReInitScanResults
        
        cmdAnalyze.Enabled = False
    
        ' Clear Error Log
        ErrReport = vbNullString
        
        CheckIntegrityHJT
        
        'pre-adding horizontal scrollbar
        If Not bAutoLogSilent Then
            SendMessage frmMain.lstResults.hWnd, LB_SETHORIZONTALEXTENT, 1500&, ByVal 0&
        End If
        
        ' *******************************************************************
        
        iPrevListIndex = 0
        
        ' pre-cache to prevent CPU indicators distortion by own process
        g_iCpuUsage = CLng(OSver.CpuUsage)
        
        StartScan '<<<<<<<-------- Main scan routine
        
        If txtNothing.Visible Or Not bAutoLog Then UpdateProgressBar "Finish"
        
        SortSectionsOfResultList
        
        'add the horizontal scrollbar if needed
        If Not bAutoLogSilent Then
            AddHorizontalScrollBarToResults lstResults
        End If
        
        If Not bAutoLogSilent Then
            'in case user moved to other frame -> return to scan results frame
            If cmdHelp.Tag = "1" Then cmdHelp_Click
            If cmdConfig.Tag = "1" Then cmdConfig_Click
            MoveToScanResultsFrame
            SetProgressBarOnFront
        End If
        
        If frmMain.lstResults.ListCount > 0 And Not bAutoLogSilent Then
            If bAutoSelect Then
                For i = 0 To frmMain.lstResults.ListCount - 1
                    frmMain.lstResults.ItemChecked(i) = True
                Next i
            End If
        End If
        
        'we are on the results window frame? - why did I check it ???
        'If lstResults.Visible Then
            cmdScan.Enabled = True
            cmdN00bLog.Enabled = True
            cmdN00bScan.Enabled = True
            FraIncludeSections.Enabled = True
            fraScanOpt.Enabled = True
            cmdFix.Enabled = True
            cmdAnalyze.Enabled = True
            cmdSaveDef.Enabled = True
        'End If
        
        If Not bAutoLog Then
            If Not IsFormForeground(frmSearch) Then
        
                If frmMain.WindowState <> vbMinimized Then
                    SetWindowPos g_HwndMain, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                    SetWindowPos g_HwndMain, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                    SetForegroundWindow g_HwndMain
                    SetActiveWindow g_HwndMain
                    SetFocus2 g_HwndMain
                End If
                
                If cmdFix.Visible And cmdFix.Enabled Then
                    cmdFix.SetFocus
                End If
            End If
        End If
        
        cmdScan.Caption = Translate(12): cmdScan.Tag = "2"
        cmdScan.Font.Bold = False
        
        g_bGeneralScanned = True ' should go before HJT_SaveReport!
        
        If bAutoLog Then
            If Not bAutoLogSilent Then DoEvents
            HJT_SaveReport '<<<<<< ------ Saving report
        End If
        
        g_bScanInProgress = False

        CloseProgressbar True
        
        bAutoLog = False
        
        If bAutoLogSilent And Not bStartupScan Then
            Call SystemPriorityDowngrade(False)
        End If
        
    Else    'Caption = Save...

        Call HJT_SaveReport
        
        UpdateProgressBar "Finish"
        
        cmdScan.Enabled = True
        cmdN00bLog.Enabled = True
        cmdN00bScan.Enabled = True
        FraIncludeSections.Enabled = True
        fraScanOpt.Enabled = True
        cmdFix.Enabled = True
    End If
    
    'focus on 1-st element of list
    'If Me.lstResults.Visible Then Me.lstResults.SetFocus
    
    isRanHJT_Scan = False
    
    If bStartupScan Then
        LockInterface bAllowInfoButtons:=True, bDoUnlock:=True
    Else
        If bAutoLogSilent Then
            'LockMenu bDoUnlock:=True
        End If
    End If
    
    If Not bAutoLogSilent Then
        'set bold type and disable items that's not need when 0 items
        If lstResults.ListCount = 0 Then
            cmdFix.Font.Bold = False
            cmdFix.Enabled = False
            cmdSaveDef.Enabled = False
        Else
            cmdFix.Font.Bold = True
        End If
    End If
    
    AppendErrorLogCustom "frmMain.cmdScan_Click - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdScan_Click", "(" & cmdScan.Caption & ")"
    If bAutoLogSilent Then LockInterface bAllowInfoButtons:=True, bDoUnlock:=True
    cmdScan.Enabled = True
    cmdN00bLog.Enabled = True
    cmdN00bScan.Enabled = True
    FraIncludeSections.Enabled = True
    fraScanOpt.Enabled = True
    cmdFix.Enabled = True
    isRanHJT_Scan = False
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdStartupList_Click() 'Misc Tools -> StartupList scan
    RunStartupList False
End Sub

Private Sub RunStartupList(bModal As Boolean)
    'On Error Resume Next
    bSL_Abort = False
    bSL_Terminate = False
    '!!! vbModal is not working here !!!
    frmStartupList2.Show IIf(bModal, vbModal, vbModeless)
End Sub

Private Sub cmdUninstall_Click() 'Misc Tools -> Uninstall HiJackThis
    HJT_Uninstall False
End Sub

Private Function HJT_Uninstall(bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim HJT_Install_Path As String
    Dim HJT_Location As String
    HJT_Location = BuildPath(PF_32, "HiJackThis Fork\HiJackThis.exe")
    
    If Not bSilent Then
        If StrComp(AppPath(True), HJT_Location, 1) = 0 Then
            'This will completely remove HiJackThis, including settings and backups. Continue?
            If MsgBoxW(Translate(154), vbQuestion Or vbYesNo) = vbNo Then Exit Function
        Else
    '    If msgboxW("This will remove HiJackThis' settings from the Registry " & _
    '              "and exit. Note that you will have to delete the " & _
    '              "HiJackThis.exe file manually." & vbCrLf & vbCrLf & _
    '              "Continue with uninstall?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            If MsgBoxW(Translate(153), vbQuestion Or vbYesNo, g_AppName) = vbNo Then Exit Function
        End If
    End If
    
    KillOtherHJTInstances HJT_Location
    
    Reg.DelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe"
    'Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis", False
    'Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis", True
    'Reg.DelKey HKEY_LOCAL_MACHINE, "Software\Soeperman Enterprises Ltd.\HijackThis", True
    'Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThisFork"
    Reg.DelKey HKEY_LOCAL_MACHINE, g_SettingsRegKey
    'If Not Reg.KeyHasSubKeys(HKEY_LOCAL_MACHINE, "Software\TrendMicro", False) Then
    '    Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro", False
    'End If
    'If Not Reg.KeyHasSubKeys(HKEY_LOCAL_MACHINE, "Software\TrendMicro", True) Then
    '    Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro", True
    'End If
    'Reg.DelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan", False 'don't change!
    'Reg.DelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan", True 'it's the name of 2.0.5 version
    CreateUninstallKey False
    DeleteBackup vbNullString, True
    ABR_RemoveBackupALL True
    RemoveHJTShortcuts
    
    RemoveAutorunHJT
    
    SetCurrentDirectory StrPtr(SysDisk)
    HJT_Install_Path = BuildPath(PF_32, "HiJackThis Fork")
    
    If FolderExists(HJT_Install_Path) Then
        If StrComp(AppPath(True), HJT_Install_Path & "\HiJackThis.exe", 1) = 0 Then
        'delayed removing of HJT installation folder via cmd.exe, if it is launched from there
            Dim sCmd As String
            '"/v /d /c (cd\& for /L %+ in (1,1,10) do ((timeout /t 1|| ping 127.1 -n 2)& rd /s /q ""{0}""&& exit))"
            sCmd = Caes_Decode("/y /m /r (xAu& KV[ /u %+ bi (0,4,89) uH ((ULRLX`a /g 8|| ojql 696.4 -K 5)& [O /d /h ""{5}""&& tOBO))")
            sCmd = Replace$(sCmd, "{0}", HJT_Install_Path)
            Proc.ProcessRun Environ$("ComSpec"), sCmd, SysDisk, vbHide, True
        Else
            DeleteFolderForce HJT_Install_Path
            RemoveDirectory StrPtr(HJT_Install_Path)
        End If
    End If
    
    Close
    g_UninstallState = True
    HJT_Uninstall = True
    Unload Me
    Exit Function
ErrorHandler:
    ErrorMsg Err, "HJT_Uninstall"
    If inIDE Then Stop: Resume Next
End Function

Private Sub Form_Resize()
    
    If bLockResize Or bAutoLogSilent Then Exit Sub
    
    If Me.WindowState = vbMinimized Then
        If (bMinToTray) Then
            If FormSys Is Nothing Then
                Set FormSys = New frmSysTray
                Load FormSys
                Set FormSys.FSys = Me
                FormSys.TrayIcon = Me
            End If
            frmSysTray.MeResize Me
        End If
    Else
        If Not (FormSys Is Nothing) Then
            Unload FormSys
            Set FormSys = Nothing
        End If
    End If
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleHeight < 5800 Then Exit Sub
    If Me.ScaleWidth < 6560 Then Exit Sub
    
    '== width ==
    ' - main -
    
    lstResults.Width = Me.ScaleWidth - 240
    shpBackground.Width = Me.ScaleWidth - 500
    shpMD5Progress.Width = (Me.ScaleWidth - 500) * (CLng(shpMD5Progress.Tag) / 100)
    shpMD5Background.Width = Me.ScaleWidth - 500
    lblMD5.Width = Me.ScaleWidth - 500
    
    txtNothing.Left = (Me.Width - txtNothing.Width) / 2
    fraOther.Left = Me.ScaleWidth - 2895
    
    ' - help
    fraHelp.Width = Me.ScaleWidth - 240
    txtHelp.Width = Me.ScaleWidth - 480
    
    ' - config -
    fraConfig.Width = Me.ScaleWidth - 240
    fraConfigTabs(0).Width = Me.ScaleWidth - 480
    fraConfigTabs(1).Width = Me.ScaleWidth - 480
    fraConfigTabs(2).Width = Me.ScaleWidth - 480
    fraConfigTabs(3).Width = Me.ScaleWidth - 480
    '(ignorelist)
    lstIgnore.Width = Me.ScaleWidth - 1800
    cmdConfigIgnoreDelSel.Left = Me.ScaleWidth - 1575
    cmdConfigIgnoreDelAll.Left = Me.ScaleWidth - 1575
    '(backups)
    lstBackups.Width = Me.ScaleWidth - 1800
    cmdConfigBackupRestore.Left = Me.ScaleWidth - 1575
    cmdConfigBackupDelete.Left = Me.ScaleWidth - 1575
    cmdConfigBackupDeleteAll.Left = Me.ScaleWidth - 1575
    cmdConfigBackupCreateRegBackup.Left = Me.ScaleWidth - 1575
    cmdConfigBackupCreateSRP.Left = Me.ScaleWidth - 1575
    
    '(misc)
    fraHostsMan.Width = Me.ScaleWidth - 480
    lstHostsMan.Width = Me.ScaleWidth - 720
    
    fraN00b.Width = Me.ScaleWidth - 195
    
    '== height ==
    ' - main -
    lstResults.Height = Me.ScaleHeight - 2490
    fraScan.Top = Me.ScaleHeight - 1530
    fraOther.Top = Me.ScaleHeight - 1530
    fraSubmit.Top = Me.ScaleHeight - 1530
    txtNothing.Top = lstResults.Top + (lstResults.Height - txtNothing.Height) / 2
    ' - help -
    fraHelp.Height = Me.ScaleHeight - 1650
    txtHelp.Height = Me.ScaleHeight - 2320
    ' - config -
    fraConfig.Height = Me.ScaleHeight - 1755
    'fraConfigTabs(0).Height = Me.ScaleHeight - 2805
    fraConfigTabs(0).Height = Me.ScaleHeight - 2700
    fraConfigTabs(1).Height = Me.ScaleHeight - 2805
    fraConfigTabs(2).Height = Me.ScaleHeight - 2805
    fraConfigTabs(3).Height = Me.ScaleHeight - 2750
    '(main)
    '(ignorelist)
    lstIgnore.Height = Me.ScaleHeight - 3250
    '(backups)
    lstBackups.Height = Me.ScaleHeight - 3850
    chkShowSRP.Top = lstBackups.Top + lstBackups.Height + 10
    '(misc)
    
    fraHostsMan.Height = Me.ScaleHeight - 2805
    lstHostsMan.Height = Me.ScaleHeight - 4035 - 240
    lblHostsTip2.Top = Me.ScaleHeight - 3300 - 300
    cmdHostsManDel.Top = Me.ScaleHeight - 3300
    cmdHostsManToggle.Top = Me.ScaleHeight - 3300
    cmdHostsManOpen.Top = Me.ScaleHeight - 3300
    cmdHostsManReset.Top = Me.ScaleHeight - 3300
    cmdHostsManRefreshList.Top = Me.ScaleHeight - 3300
    vscMiscTools.Height = fraConfigTabs(3).Height
    fraN00b.Height = Me.ScaleHeight - 1175
    
    'scrolling bar for misc tools frame
    'imgMiscToolsDown.Top = fraConfigTabs(3).Height - 255
    'imgMiscToolsDown2.Top = fraConfigTabs(3).Height - 255
    If fraConfig.Height < fraMiscToolsScroll.Height + 1050 Then
        'imgMiscToolsUp.Visible = True
        'imgMiscToolsDown.Visible = True
        vscMiscTools.Visible = True
    Else
        'imgMiscToolsUp.Visible = False
        'imgMiscToolsUp2.Visible = False
        'imgMiscToolsDown.Visible = False
        'imgMiscToolsDown2.Visible = False
        vscMiscTools.Visible = False
    End If
    
    'add the horizontal scrollbar to the results display if needed
    If Not bAutoLogSilent Then
        AddHorizontalScrollBarToResults lstResults
    End If
End Sub

Private Sub LoadSettings(Optional nRun As Long)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "frmMain.LoadSettings - Begin"
    
    Dim sCurLang$, WinHeight&, WinWidth&, lProxyType&, lSocksVer&
    
    ' Scan area
    
    chkLogProcesses.Value = CInt(RegReadHJT("LogProcesses", "1"))
    chkAdvLogEnvVars.Value = CInt(RegReadHJT("LogEnvVars", "0"))
    chkAdditionalScan.Value = CInt(RegReadHJT("LogAdditional", "1"))
    
    bLogProcesses = chkLogProcesses.Value
    bLogEnvVars = chkAdvLogEnvVars.Value
    bAdditional = chkAdditionalScan.Value

    ' Scan options
    
    chkIgnoreMicrosoft.Value = CInt(RegReadHJT("HideMicrosoft", "1"))
    chkIgnoreAll.Value = CInt(RegReadHJT("IgnoreAllWhiteList", "0"))
    
    Dim sHashType As String
    chkDoCheckSum.Value = CInt(RegReadHJT("CalcMD5", "0")) ' Calc CheckSum ?
    cmbHashType.Enabled = CBool(chkDoCheckSum.Value)
    sHashType = RegReadHJT("HashType", vbNullString)
    If Len(sHashType) = 0 Then
        sHashType = "Newest"
        RegSaveHJT "HashType", sHashType
    End If
    Select Case sHashType
    Case "Newest"
        If OS_SupportSHA2 Then
            ComboSetValue cmbHashType, "SHA256"
            g_eUseHashType = HASH_TYPE_SHA256
        Else
            ComboSetValue cmbHashType, "SHA1"
            g_eUseHashType = HASH_TYPE_SHA1
        End If
    Case "SHA256"
        ComboSetValue cmbHashType, "SHA256"
        g_eUseHashType = HASH_TYPE_SHA256
    Case "SHA1"
        ComboSetValue cmbHashType, "SHA1"
        g_eUseHashType = HASH_TYPE_SHA1
    Case "MD5"
        ComboSetValue cmbHashType, "MD5"
        g_eUseHashType = HASH_TYPE_MD5
    End Select
    
    bHideMicrosoft = chkIgnoreMicrosoft.Value
    bIgnoreAllWhitelists = chkIgnoreAll.Value
    g_bCheckSum = chkDoCheckSum.Value
    
    ' Fix & Backup
    
    chkBackup.Value = CInt(RegReadHJT("MakeBackup", "1"))
    chkConfirm.Value = CInt(RegReadHJT("Confirm", "1"))
    chkAutoMark.Value = CInt(RegReadHJT("AutoSelect", "0"))
    
    bMakeBackup = chkBackup.Value
    bConfirm = chkConfirm.Value
    bAutoSelect = chkAutoMark.Value

    ' Interface
    
    'chkSkipIntroFrameSettings.Value = CInt(RegReadHJT("SkipIntroFrame", "0", bUseOldKey))
    'chkSkipIntroFrame.Value = CInt(RegReadHJT("SkipIntroFrame", "0", bUseOldKey))
    chkSkipErrorMsg.Value = CInt(RegReadHJT("SkipErrorMsg", "0"))
    chkConfigMinimizeToTray.Value = CInt(RegReadHJT("MinToTray", "0"))
    
    bSkipErrorMsg = chkSkipErrorMsg.Value
    bMinToTray = chkConfigMinimizeToTray.Value
    
    g_FontName = RegReadHJT("FontName", vbNullString)
    g_FontSize = RegReadHJT("FontSize", "Auto")
    g_bFontBold = CInt(RegReadHJT("FontBold", "1"))
    
    If Len(g_FontName) = 0 Then
        If FontExist("MS Sans Serif") Then
            g_FontName = "MS Sans Serif"
            g_FontSize = "9"
        Else
            g_FontName = "Automatic"
        End If
        g_bFontBold = True
    End If
    Me.chkFontBold.Value = Abs(g_bFontBold)
    
    chkFontWholeInterface.Value = CInt(RegReadHJT("FontOnInterface", "0"))
    
    sCurLang = RegReadHJT("LanguageFile", "English")
    
    ' Updates

    chkCheckUpdatesOnStart.Value = CInt(RegReadHJT("CheckForUpdates", "0"))
    chkUpdateToTest.Value = CInt(RegReadHJT("UpdateToTest", "0"))
    chkUpdateSilently.Value = CInt(RegReadHJT("UpdateSilently", "0"))
    lProxyType = CInt(RegReadHJT("ProxyType", "1")) '0 - Direct, 1 - IE, 2 - Manual proxy
    OptProxyDirect.Value = Abs(lProxyType = 0)
    optProxyIE.Value = Abs(lProxyType = 1)
    optProxyManual.Value = Abs(lProxyType = 2)
    
    lSocksVer = CInt(RegReadHJT("ProxySocksVer", "0"))
    chkSocks4.Value = Abs(lSocksVer = 4)
    
    chkUpdateUseProxyAuth.Value = CInt(RegReadHJT("ProxyUseAuth", "0"))
    
    txtUpdateProxyHost.Text = RegReadHJT("ProxyServer", vbNullString)
    txtUpdateProxyPort.Text = RegReadHJT("ProxyPort", vbNullString)
    txtUpdateProxyLogin.Text = RegReadHJT("ProxyLogin", vbNullString)
    txtUpdateProxyPass.Text = DeCrypt(RegReadHJT("ProxyPass", vbNullString))
    
    bCheckForUpdates = chkCheckUpdatesOnStart.Value
    bUpdateToTest = chkUpdateToTest.Value
    bUpdateSilently = chkUpdateSilently.Value

    ' Backup (restore point)
    
    chkShowSRP.Value = CInt(RegReadHJT("ShowSRP", "0"))
    
    gNotUserClick = True
    If OSver.IsWindowsVistaOrGreater Then
        If FileExists(BuildPath(sWinSysDir, "Tasks\HiJackThis Autostart Scan")) Then
            chkConfigStartupScan.Value = 1
        Else
            chkConfigStartupScan.Value = 0
        End If
    Else
        If Reg.ValueExists(HKLM, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis Autostart Scan") Then
            chkConfigStartupScan.Value = 1
        Else
            chkConfigStartupScan.Value = 0
        End If
    End If
    gNotUserClick = False
    
    Dim sData$, LastVerLaunched$, isEncodedVer As Boolean
    
    LastVerLaunched = RegReadHJT("Version", vbNullString)
    If ConvertVersionToNumber(LastVerLaunched) < ConvertVersionToNumber("2.6.1.21") Then isEncodedVer = True
    If Len(LastVerLaunched) <> 0 Then
        'force "additional" log as default if previous version was launch before
        If ConvertVersionToNumber(LastVerLaunched) <= ConvertVersionToNumber("3.1.0.1") Then
            chkAdditionalScan.Value = vbChecked
            RegSaveHJT "LogAdditional", "1"
        End If
    End If
    
    Dim iIgnoreNum As Long, i As Long
    
    If CryptVer < 2 Then
        RegSaveHJT "CryptVer", 2
    End If
    
    If CryptVer = 1 Then 'not supported anymore
        iIgnoreNum = Val(RegReadHJT("IgnoreNum", "0", True))
        If iIgnoreNum > 0 Then
            For i = 1 To iIgnoreNum
                RegDelHJT "Ignore" & CStr(i)
            Next
        End If
    End If
    
    g_sLastSearch = RegReadHJT("LastSearch", vbNullString)
    
    IsOnIgnoreList vbNullString, UpdateList:=True
    
    If CryptVer = 2 And OSver.IsElevated And nRun = 0 Then 'nRun - surely prevents infinite recurse call
        'need to reEncode
        iIgnoreNum = Val(RegReadHJT("IgnoreNum", "0"))
        If iIgnoreNum > 0 Then
            Dim aIgnoreList() As String
            ReDim aIgnoreList(iIgnoreNum) As String
            For i = 1 To iIgnoreNum
                aIgnoreList(i) = DeCrypt(RegReadHJT("Ignore" & i, vbNullString))
            Next
        End If
        
        RegSaveHJT "CryptVer", 3: CryptVer = 3
        cryptInit
        
        If iIgnoreNum > 0 Then
            For i = 1 To iIgnoreNum
                RegSaveHJT "Ignore" & CStr(i), Crypt(aIgnoreList(i))
            Next i
        End If
        
        SaveSettings
        LoadSettings nRun:=1
    End If
    
    AppendErrorLogCustom "frmMain.LoadSettings - End"
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "frmMain_LoadSettings"
    If inIDE Then Stop: Resume Next
    Resume Next
End Sub

Private Sub mnuToolsDigiSign_Click()        'Tools -> Files -> Digital signature checker
    frmCheckDigiSign.Show vbModeless
End Sub

Private Sub mnuToolsRegUnlockKey_Click()    'Tools -> Registry -> Key unlocker
    frmUnlockRegKey.Show vbModeless
End Sub

Private Sub mnuToolsRegTypeChecker_Click()  'Tools -> Registry -> Key Type Checker
    frmRegTypeChecker.Show vbModeless
End Sub

Private Sub mnuToolsStartupList_Click()     'Tools -> StartupList
    cmdStartupList_Click
End Sub

Private Sub vscMiscTools_Change()
    'lToolsHeight = 2200 ' decrease this value if you would like more space inside scroll of last config tab
    'note: this value is redefined in "FormStart_Stady1" (separately, for IDE and release)
    fraMiscToolsScroll.Top = -vscMiscTools.Value * (fraMiscToolsScroll.Height - (fraConfigTabs(3).Height + lToolsHeight)) / 100
    DoEvents
End Sub

Private Sub vscMiscTools_Scroll()
    Call vscMiscTools_Change
End Sub

Private Sub LoadLanguageList()
    On Error GoTo ErrorHandler:
    Dim sFile$, sCurLang$, i&, LangID&
    
    AppendErrorLogCustom "frmMain.LoadLanguageList - Begin"
    
    cboN00bLanguage.AddItem "English"
    cboN00bLanguage.AddItem "French"
    cboN00bLanguage.AddItem "Russian"
    cboN00bLanguage.AddItem "Spanish"
    cboN00bLanguage.AddItem "Ukrainian"
    
    sFile = DirW$(BuildPath(AppPath(), "*.lng"), vbFile)
    
    Do While Len(sFile)
        If sFile <> "_Lang_EN.lng" And _
            sFile <> "_Lang_FR.lng" And _
            sFile <> "_Lang_RU.lng" And _
            sFile <> "_Lang_UA.lng" And _
            sFile <> "_Lang_SP.lng" Then
                cboN00bLanguage.AddItem sFile
        End If
        sFile = DirW$()
    Loop
    
    sCurLang = RegReadHJT("LanguageFile", "English")  'HJT settings
    If bForceFR Then sCurLang = "French"
    If bForceRU Then sCurLang = "Russian"
    If bForceUA Then sCurLang = "Ukrainian"
    If bForceEN Then sCurLang = "English"
    If bForceSP Then sCurLang = "Spanish"
    
    LangID = -1
    For i = 0 To cboN00bLanguage.ListCount - 1
        If sCurLang = cboN00bLanguage.List(i) Then LangID = i
    Next
    
    If LangID = -1 Then LangID = 0 'default language - English
    
    cboN00bLanguage.ListIndex = LangID
    
    AppendErrorLogCustom "frmMain.LoadLanguageList - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_LoadLanguageList"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cboN00bLanguage_Click()
    On Error GoTo ErrorHandler:
    Dim sFile$
    
    AppendErrorLogCustom "frmMain.cboN00bLanguage_Click - Begin"
    
    'Lang IDs
    'https://docs.microsoft.com/en-US/windows/win32/intl/language-identifier-constants-and-strings
    
    sFile = cboN00bLanguage.List(cboN00bLanguage.ListIndex)
    
    If HasCommandLineKey("default") Then sFile = "English"
    
    If Len(sFile) = 0 Then Exit Sub
    If sFile = "English" Then
        'LoadDefaultLanguage
        LoadLanguage &H409, bForceEN
        g_CurrentLangID = &H409
        g_CurrentLang = sFile
    ElseIf sFile = "Russian" Then
        LoadLanguage &H419, bForceRU
        g_CurrentLangID = &H419
        g_CurrentLang = sFile
    ElseIf sFile = "Ukrainian" Then
        LoadLanguage &H422, bForceUA
        g_CurrentLangID = &H422
        g_CurrentLang = "Russian" 'magik
    ElseIf sFile = "French" Then
        LoadLanguage &H40C, bForceFR
        g_CurrentLangID = &H40C
        g_CurrentLang = sFile
    ElseIf sFile = "Spanish" Then
        LoadLanguage &H40A, bForceSP
        g_CurrentLangID = &H40A
        g_CurrentLang = sFile
    Else
        LoadEncryptedLangFile sFile
        ReloadLanguageNative
        ReloadLanguage
        g_CurrentLangID = &H409
        g_CurrentLang = "English"
    End If
    
    ' Do not save force mode state!
    If Not (bForceRU Or bForceEN Or bForceUA Or bForceFR Or bForceSP) Then RegSaveHJT "LanguageFile", sFile
    
    If cmdN00bScan.Enabled And cmdN00bScan.Visible Then cmdN00bScan.SetFocus
    AppendErrorLogCustom "frmMain.cboN00bLanguage_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_cboN00bLanguage_Click"
    If inIDE Then Stop: Resume Next
End Sub


' =========== Menus ==============
'
'

Private Sub mnuFileSettings_Click()     'File -> Settings
    cmdN00bTools_Click
    Call chkConfigTabs_Click(0)
End Sub

Private Sub mnuFileInstallHJT_Click()   'File -> Install HJT
    InstallHJT True
End Sub

Private Sub mnuFileUninstHJT_Click()    'File -> Uninstall HJT
    If isRanHJT_Scan Then Exit Sub
    cmdUninstall_Click
End Sub

Private Sub mnuFileExit_Click()         'File -> Exit
    bmnuExit_Clicked = True
    Unload Me
    bmnuExit_Clicked = False
End Sub

Private Sub mnuToolsADSSpy_Click()      'Tools -> ADS Spy
    'cmdN00bTools_Click
    cmdADSSpy_Click
End Sub

Private Sub mnuToolsDelFileOnReboot_Click()     'Tools -> Files -> Delete a file on reboot...
    cmdDelOnReboot_Click
End Sub

Private Sub mnuToolsUnlockFiles_Click()    'Tools -> Files -> Unlock Folder...
    frmUnlockFile.Show
End Sub

Private Sub mnuToolsDelServ_Click()     'Tools -> Delete Service
    cmdDeleteService_Click
End Sub

Private Sub mnuToolsHosts_Click()       'Tools -> Hosts file manager
    cmdN00bTools_Click
    cmdHostsManager_Click
End Sub

Private Sub mnuToolsProcMan_Click()     'Tools -> Process Manager
    frmProcMan.Show
End Sub

Private Sub mnuToolsUninst_Click()      'Tools -> Uninstall manager
    cmdARSMan_Click
End Sub

Private Sub mnuToolsShortcutsChecker_Click()    'Tools -> Shortcuts -> Check Browsers' LNK
    'Download Check Browsers' LNK by Dragokas & regist
    'and ask to run
    Dim sTool As String: sTool = BuildPath(GetToolsDir(), "Check Browsers LNK.exe")
    Dim bRequireDL As Boolean
    If FileExists(sTool) Then
        If bCheckForUpdates Then
            If DateDiff("d", GetFileDate(sTool, DATE_CREATED), Now()) > 30 Then ' 1 month elapsed
                bRequireDL = True
            End If
        End If
        If Not bRequireDL Then
            If OSver.IsWindowsVistaOrGreater Then
                If Not IsDragokasFile(sTool) Then bRequireDL = True
            End If
        End If
    Else
        bRequireDL = True
    End If
    If Not bRequireDL Then
        Proc.ProcessRun sTool, vbNullString, AppPath(False), 1, True
    Else
        DownloadUnzipAndRun "https://dragokas.com/tools/" & Caes_Decode("DkjjtMEDNLzORgkj.cT]"), "Check Browsers LNK.exe", bUpdateSilently, True 'CheckBrowsersLNK.zip
    End If
End Sub
Private Sub mnuToolsShortcutsFixer_Click()      'Tools -> Shortcuts -> ClearLNK
    'Download ClearLNK by Dragokas
    'and ask to run
    Dim sTool As String: sTool = BuildPath(GetToolsDir(), "ClearLNK.exe")
    Dim bRequireDL As Boolean
    If FileExists(sTool) Then
        If bCheckForUpdates Then
            If DateDiff("d", GetFileDate(sTool, DATE_CREATED), Now()) > 30 Then ' 1 month elapsed
                bRequireDL = True
            End If
        End If
        If Not bRequireDL Then
            If OSver.IsWindowsVistaOrGreater Then
                If Not IsDragokasFile(sTool) Then bRequireDL = True
            End If
        End If
    Else
        bRequireDL = True
    End If
    If Not bRequireDL Then
        Proc.ProcessRun sTool, vbNullString, AppPath(False), 1, True
    Else
        DownloadUnzipAndRun "https://dragokas.com/tools/" & Caes_Decode("DojhAW[Z.SDM"), "ClearLNK.exe", bUpdateSilently, True 'ClearLNK.zip
    End If
End Sub

Private Sub mnuHelpManualEnglish_Click()
    Dim szQSUrl$: szQSUrl = "https://dragokas.com/tools/help/hjt_tutorial.html"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualRussian_Click()
    Dim szQSUrl$
    'szQSUrl = "https://safezone.cc/threads/25184/"
    szQSUrl = "https://regist.safezone.cc/hijackthis_help/hijackthis.html"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualFrench_Click()
    'https://www.bleepingcomputer.com
    Dim szQSUrl$: szQSUrl = Caes_Decode("iwywB://NPR.AMHJWRYTR``elm`o.drr") & "/tutorials/comment-utiliser-hijackthis/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualGerman_Click()
    'https://www.bleepingcomputer.com
    Dim szQSUrl$: szQSUrl = Caes_Decode("iwywB://NPR.AMHJWRYTR``elm`o.drr") & "/tutorials/wie-hijackthis-genutzt-wird-um/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualSpanish_Click()
    'https://www.bleepingcomputer.com
    Dim szQSUrl$: szQSUrl = Caes_Decode("iwywB://NPR.AMHJWRYTR``elm`o.drr") & "/tutorials/como-usar-hijackthis/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualPortuguese_Click()
    'https://www.linhadefensiva.org
    Dim szQSUrl$: szQSUrl = Caes_Decode("iwywB://NPR.KJQMHMPST_f^mZ.lqh") & "/2005/06/hijackthis-completo/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualDutch_Click()
    'https://www.bleepingcomputer.com
    Dim szQSUrl$: szQSUrl = Caes_Decode("iwywB://NPR.AMHJWRYTR``elm`o.drr") & "/tutorials/hoe-gebruik-je-hijackthis/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub

Private Sub mnuHelpUpdate_Click()       'Help -> Download new version
    CheckForUpdate False, bUpdateSilently, bUpdateToTest
    If g_NeedTerminate Then Unload Me
End Sub

Private Sub mnuHelpAbout_Click()        'Help -> About HJT
    cmdN00bClose_Click
    ' Закрыть окно "Инструменты"
    'If cmdConfig.Caption = Translate(19) Then cmdConfig_Click
    If cmdConfig.Tag = "1" Then cmdConfig_Click
    'If cmdHelp.Caption = Translate(16) Then cmdHelp_Click
    If cmdHelp.Tag = "0" Then cmdHelp_Click
    fraHelp.Visible = True
    fraHelp.ZOrder 0
    'chkHelp(2).value = 1
    chkHelp_Click 2
End Sub

' --------------------------------------

'Private Sub txtHelp_LostFocus()
'    txtHelpHasFocus = False
'End Sub
'Private Sub txtHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Not txtHelpHasFocus Then
'        If GetForegroundWindow() = txtHelp.Parent.hwnd Then
'            txtHelpHasFocus = True
'            If txtHelp.Visible Then
'                txtHelp.SetFocus
'            End If
'        End If
'    End If
'End Sub

Sub SaveSettings()
    
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.SaveSettings - Begin"
    
    bAutoSelect = IIf(chkAutoMark.Value = 1, True, False)
    bConfirm = IIf(chkConfirm.Value = 1, True, False)
    bMakeBackup = IIf(chkBackup.Value = 1, True, False)
    'bIgnoreSafeDomains = IIf(chkIgnoreSafeDomains.Value = 1, True, False)
    
    bLogProcesses = IIf(chkLogProcesses.Value = 1, True, False)
    bAdditional = IIf(chkAdditionalScan.Value = 1, True, False)
    bSkipErrorMsg = IIf(chkSkipErrorMsg.Value = 1, True, False)
    bMinToTray = IIf(chkConfigMinimizeToTray.Value = 1, True, False)
    bLogEnvVars = (chkAdvLogEnvVars.Value = 1)
    g_bCheckSum = (chkDoCheckSum.Value = 1)
    bCheckForUpdates = IIf(chkCheckUpdatesOnStart.Value = 1, True, False)
    
    RegSaveHJT "AutoSelect", CStr(Abs(CInt(bAutoSelect)))
    RegSaveHJT "Confirm", CStr(Abs(CInt(bConfirm)))
    RegSaveHJT "MakeBackup", CStr(Abs(CInt(bMakeBackup)))
    'RegSaveHJT "IgnoreSafe", CStr(Abs(CInt(bIgnoreSafeDomains)))
    RegSaveHJT "LogProcesses", CStr(Abs(CInt(bLogProcesses)))
    RegSaveHJT "LogAdditional", CStr(Abs(CInt(bAdditional)))
    'RegSaveHJT "SkipIntroFrame", CStr(chkSkipIntroFrameSettings.Value)
    RegSaveHJT "SkipErrorMsg", CStr(Abs(CInt(bSkipErrorMsg)))
    RegSaveHJT "MinToTray", CStr(Abs(CInt(bMinToTray)))
    RegSaveHJT "LastSearch", g_sLastSearch
    RegSaveHJT "LogEnvVars", Abs(CLng(bLogEnvVars))
    RegSaveHJT "CalcMD5", Abs(CLng(g_bCheckSum)) 'CalcMD5 - for backward compatibility, actual meaning is "Calc CheckSum"
    Select Case g_eUseHashType
    Case HASH_TYPE_MD5:     RegSaveHJT "HashType", "MD5"
    Case HASH_TYPE_SHA1:    RegSaveHJT "HashType", "SHA1"
    Case HASH_TYPE_SHA256:  RegSaveHJT "HashType", "SHA256"
    End Select
    RegSaveHJT "CheckForUpdates", CStr(Abs(CInt(bCheckForUpdates)))
    RegSaveHJT "UpdateToTest", CStr(Abs(CInt(bUpdateToTest)))
    RegSaveHJT "UpdateSilently", CStr(Abs(CInt(bUpdateSilently)))
    RegSaveProxySettings
    
    AppendErrorLogCustom "frmMain.SaveSettings - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SaveSettings"
    If inIDE Then Stop: Resume Next
End Sub


'Context menu in result list of scan:

Private Sub lstResults_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo ErrorHandler:
    
    Dim result As SCAN_RESULT
    Dim sItem As String
    Dim i As Long, j As Long
    Dim RegItems As Long
    Dim FileItems As Long
    Dim idx As Long, XY As Long, XPix As Long, YPix As Long
    Dim iMin As Long, iMax As Long, iStep As Long
    
    If Button = 1 Then
        If Shift = 1 And Not bScanMode Then 'Mark (check) multiple items with 'Shift + Click'
        
            If lstResults.ListIndex = -1 Then Exit Sub
            
            iMin = iPrevListIndex
            iMax = lstResults.ListIndex
            iStep = IIf(iMax > iMin, 1, -1)
            
            If Abs(iMax - iMin) > 0 Then 'don't double mark the same single item

                For i = iMin To iMax Step iStep
                    lstResults.ItemChecked(i) = Not lstResults.ItemChecked(i)
                Next
            End If
            
            lstResults.ListIndex = iMax
        End If
        
        iPrevListIndex = lstResults.ListIndex
    End If
    
    'select item by right click
    If Button = 2 And Not (x = -1 And y = -1) Then '-1, -1 is magik from 'KeyDown' event
        XPix = x / Screen.TwipsPerPixelX
        YPix = y / Screen.TwipsPerPixelY
        XY = YPix * 65536 + XPix
        idx = SendMessage(lstResults.hWnd, LB_ITEMFROMPOINT, 0&, ByVal XY)
        If idx >= 0 And idx <= (lstResults.ListCount - 1) Then
            lstResults.ListIndex = idx
        End If
    End If
    
    If Button = 2 And Not (isRanHJT_Scan And bAutoLogSilent) Then
        If lstResults.CheckedIndices.Count = 0 Then     'items not checked ?
            mnuResultFix.Enabled = False
            mnuResultAddToIgnore.Enabled = False
        Else
            mnuResultFix.Enabled = True
            mnuResultAddToIgnore.Enabled = True
        End If
        If lstResults.ListIndex = -1 Then   'item not selected ?
            mnuResultInfo.Enabled = False
            mnuResultSearch.Enabled = False
            If mnuResultVT.Visible Then mnuResultVT.Enabled = False
            mnuResultCopy.Enabled = False
            On Error Resume Next
            mnuResultDelim1.Enabled = False
            On Error GoTo 0
        Else
            mnuResultInfo.Enabled = True
            mnuResultSearch.Enabled = True
            If mnuResultVT.Visible Then mnuResultVT.Enabled = True
            mnuResultCopy.Enabled = True
            On Error Resume Next
            mnuResultDelim1.Enabled = True ' wtf this one is rarely returns "'Enabled' property can't be set on this control"
            On Error GoTo 0
        End If
        If lstResults.ListCount = 0 Then    'no items
            mnuResultAddALLToIgnore.Enabled = False
        Else
            mnuResultAddALLToIgnore.Enabled = True
        End If
        mnuSaveReport.Enabled = True
        mnuResultReScan.Enabled = True
        
        If isRanHJT_Scan Then
            mnuResultFix.Enabled = False
            mnuSaveReport.Enabled = False
            mnuResultReScan.Enabled = False
        End If
        
        Erase JumpFileCache
        Erase JumpRegCache
        
        'building the jump list
        mnuResultJump.Enabled = False
        
        sItem = GetSelected_OrCheckedItem()
        
        If Len(sItem) <> 0 Then
        
            If GetScanResults(sItem, result) Then
                
                If (result.Section = "O22" Or result.Section = "O23") And InStr(result.HitLineW, "O22 - BITS") = 0 Then 'services & tasks
                    mnuResultDisable.Visible = True
                    
                    If result.State = ITEM_STATE_ENABLED Then
                        SetMenuCaptionByMenu mnuResultDisable, Translate(1168) ' "Disable"
                    Else
                        SetMenuCaptionByMenu mnuResultDisable, Translate(1169) ' "Enable"
                    End If
                    
                    If result.Section = "O22" Then
                    
                        If StrBeginWith(sItem, "O22 - Task (.job)") Then
                        
                            mnuResultDisable.Enabled = False
                        Else
                            mnuResultDisable.Enabled = True
                        End If
                        
                    ElseIf result.Section = "O23" Then
                        
                        If StrBeginWith(sItem, "O23 - Service") Then
                        
                            mnuResultDisable.Enabled = True
                        Else
                            mnuResultDisable.Enabled = False
                        End If
                    End If
                    
                Else
                    mnuResultDisable.Visible = False
                End If
                
                If AryPtr(result.File) Or AryPtr(result.Reg) Or AryPtr(result.Jump) Then
                    mnuResultJump.Enabled = True
                    
                    If CBool(AryPtr(result.File)) And CBool(AryPtr(result.Reg)) Then
                        mnuResultJumpDelim.Visible = True
                    Else
                        mnuResultJumpDelim.Visible = False
                    End If
                    
                    For j = 1 To MAX_JUMP_LIST_ITEMS
                        mnuResultJumpFile(j - 1).Visible = True
                    Next
                    For j = 1 To MAX_JUMP_LIST_ITEMS
                        mnuResultJumpReg(j - 1).Visible = True
                    Next
                    
                    'list of cure files
                    JumpListExtractFiles result.File, FileItems
                    
                    'list of cure reg. entries
                    JumpListExtractRegistry result.Reg, FileItems, RegItems

                    'list of files and reg. entries added just for jumping
                    If AryPtr(result.Jump) Then
                        For j = 0 To UBound(result.Jump)
                            With result.Jump(j)
                                If (.Type And JUMP_ENTRY_FILE) Then
                                    JumpListExtractFiles .File, FileItems
                                ElseIf (.Type And JUMP_ENTRY_REGISTRY) Then
                                    JumpListExtractRegistry .Registry, FileItems, RegItems
                                End If
                            End With
                        Next
                    End If
                    
                    For j = FileItems + 1 To MAX_JUMP_LIST_ITEMS
                        mnuResultJumpFile(j - 1).Visible = False
                    Next
                    For j = RegItems + 1 To MAX_JUMP_LIST_ITEMS
                        mnuResultJumpReg(j - 1).Visible = False
                    Next
                    
                End If
            End If
        Else
            mnuResultDisable.Visible = False
        End If
        PopupMenu mnuResultList
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.lstResults_MouseUp"
    If inIDE Then Stop: Resume Next
End Sub

'FileItems++
'FIX_FILE -> mnuResultJumpFile
'FIX_FILE -> JumpFileCache()
'consider MAX_JUMP_LIST_ITEMS
Private Sub JumpListExtractFiles(aFixFile() As FIX_FILE, in_out_FileItems As Long)
    Dim bExists As Boolean
    Dim j As Long
    If AryPtr(aFixFile) Then
        For j = 0 To UBound(aFixFile)
            If in_out_FileItems >= MAX_JUMP_LIST_ITEMS Then Exit For
            
            bExists = FileExists(aFixFile(j).Path)
            
            If bExists Then
                in_out_FileItems = in_out_FileItems + 1
                mnuResultJumpFile(in_out_FileItems - 1).Caption = aFixFile(j).Path
                
                If AryPtr(JumpFileCache) Then
                    ReDim Preserve JumpFileCache(UBound(JumpFileCache) + 1)
                Else
                    ReDim JumpFileCache(0)
                End If
                
                JumpFileCache(UBound(JumpFileCache)) = aFixFile(j)
            End If
        Next
    End If
End Sub

'RegItems++ (FileItems++)
'FIX_REG_KEY -> mnuResultJumpFile
'FIX_REG_KEY -> JumpRegCache()
'consider MAX_JUMP_LIST_ITEMS
Private Sub JumpListExtractRegistry(aFixReg() As FIX_REG_KEY, in_out_FileItems As Long, in_out_RegItems As Long)
    Dim bNoValue As Boolean
    Dim bExists As Boolean
    Dim j As Long
    If AryPtr(aFixReg) Then
        For j = 0 To UBound(aFixReg)
            With aFixReg(j)
                If Len(.IniFile) <> 0 Then
                    
                    If in_out_FileItems < MAX_JUMP_LIST_ITEMS Then
                        in_out_FileItems = in_out_FileItems + 1
                        bExists = FileExists(.IniFile)
                        mnuResultJumpFile(in_out_FileItems - 1).Caption = .IniFile & " => [" & .Key & "], " & .Param & IIf(bExists, vbNullString, " (no file)")
                        
                        If AryPtr(JumpFileCache) Then
                            ReDim Preserve JumpFileCache(UBound(JumpFileCache) + 1)
                        Else
                            ReDim JumpFileCache(0)
                        End If
                        
                        JumpFileCache(UBound(JumpFileCache)).Path = .IniFile
                    End If
                Else
                    If in_out_RegItems < MAX_JUMP_LIST_ITEMS Then
                        in_out_RegItems = in_out_RegItems + 1
                        bExists = Reg.KeyExists(.Hive, .Key, .Redirected)
                        bNoValue = False
                        If (.ActionType And BACKUP_KEY) Or (.ActionType And REMOVE_KEY) Then
                        Else
                            bNoValue = Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected)
                        End If
                          
                        mnuResultJumpReg(in_out_RegItems - 1).Caption = _
                          Reg.GetPhysicalNameWin32(.Hive, .Key, .Redirected) & ", " & .Param & _
                          IIf(.Redirected, " (x32)", vbNullString) & IIf(bExists, vbNullString, " (no key)") & IIf(bNoValue, " (no value)", vbNullString)
                        
                        If AryPtr(JumpRegCache) Then
                            ReDim Preserve JumpRegCache(UBound(JumpRegCache) + 1)
                        Else
                            ReDim JumpRegCache(0)
                        End If
                        
                        JumpRegCache(UBound(JumpRegCache)) = aFixReg(j)
                    End If
                End If
            End With
        Next
    End If
End Sub

Private Sub mnuResultJumpFile_Click(Index As Integer)   'Context => Jump to ... => File
    Dim sItem As String
    Dim sFile As String
    Dim sFolder As String
    Dim result As SCAN_RESULT
    
    sItem = GetSelected_OrCheckedItem()
    
    If GetScanResults(sItem, result) Then
        If AryPtr(JumpFileCache) Then
            If UBound(JumpFileCache) >= Index Then
                sFile = JumpFileCache(Index).Path
                sFile = GetFullPath(sFile)
                sFolder = GetParentDir(sFile)
                If FileExists(sFile) Then
                    OpenAndSelectFile sFile
                ElseIf FolderExists(sFile) Then
                    OpenAndSelectFile sFile
                ElseIf FolderExists(sFolder) Then
                    OpenAndSelectFile sFolder
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuResultJumpReg_Click(Index As Integer)   'Context => Jump to ... => Registry
    Dim sItem As String
    Dim result As SCAN_RESULT
    
    sItem = GetSelected_OrCheckedItem()
    
    If GetScanResults(sItem, result) Then
        If AryPtr(JumpRegCache) Then
            If UBound(JumpRegCache) >= Index Then
                With JumpRegCache(Index)
                    Reg.Jump .Hive, .Key, .Param, .Redirected
                End With
            End If
        End If
    End If
End Sub

Private Function GetSelected_OrCheckedItem() As String
    Dim i As Long
    If (lstResults.ListIndex <> -1) Then  'selection
        GetSelected_OrCheckedItem = lstResults.List(lstResults.ListIndex)
    ElseIf (lstResults.CheckedIndices.Count >= 1) Then  'checkbox
        For i = 0 To lstResults.ListCount - 1
            If lstResults.ItemChecked(i) = True Then
                GetSelected_OrCheckedItem = lstResults.List(i)
                Exit For
            End If
        Next
    End If
End Function

Private Function GetSelected_OrCheckedItemIndex() As Long
    Dim i As Long
    GetSelected_OrCheckedItemIndex = -1
    If (lstResults.ListIndex <> -1) Then  'selection
        GetSelected_OrCheckedItemIndex = lstResults.ListIndex
    ElseIf (lstResults.CheckedIndices.Count >= 1) Then 'checkbox
        For i = 0 To lstResults.ListCount - 1
            If lstResults.ItemChecked(i) = True Then
                GetSelected_OrCheckedItemIndex = i
                Exit For
            End If
        Next
    End If
End Function

Private Function GetSelected_OrCheckedItemResult() As SCAN_RESULT
    Dim idx As Long
    idx = GetSelected_OrCheckedItemIndex()
    If idx <> -1 Then
        Dim sItem As String
        sItem = lstResults.List(idx)
        Call GetScanResults(sItem, GetSelected_OrCheckedItemResult)
    End If
End Function

Private Sub mnuResultFix_Click()          'Context menu => Fix checked
    cmdFix_Click
End Sub

Private Sub mnuResultInfo_Click()         'Context menu => Info on selected
    cmdInfo_Click
End Sub

Private Sub mnuResultAddToIgnore_Click()  'Context menu => Add to ignore list
    cmdSaveDef_Click
End Sub

Private Sub mnuResultAddALLToIgnore_Click()  'Context menu => Add ALL to ignore list
    Dim i As Long
    For i = 0 To lstResults.ListCount - 1
        lstResults.ItemChecked(i) = True
    Next
    cmdSaveDef_Click
    If lstResults.ListCount > 0 Then
        For i = 0 To lstResults.ListCount - 1
            lstResults.ItemChecked(i) = False
        Next
    End If
End Sub

Private Sub mnuResultDisable_Click() 'Context menu => Disable (Tasks & Services)
    
    Dim result As SCAN_RESULT
    Dim sItem As String
    Dim idx As Long
    
    sItem = GetSelected_OrCheckedItem()
    
    If Len(sItem) <> 0 Then
        If GetScanResults(sItem, result, idx) Then
            
            If result.Section = "O22" Then
                
                If result.State = ITEM_STATE_ENABLED Then
                
                    DisableTask result.Name
                    Scan(idx).State = ITEM_STATE_DISABLED
                Else
                    EnableTask result.Name
                    Scan(idx).State = ITEM_STATE_ENABLED
                End If
            
            ElseIf result.Section = "O23" Then
            
                If result.State = ITEM_STATE_ENABLED Then
            
                    SetServiceStartMode result.Name, SERVICE_MODE_DISABLED
                    StopService result.Name
                    Scan(idx).State = ITEM_STATE_DISABLED
                Else
                    SetServiceStartMode result.Name, SERVICE_MODE_AUTOMATIC
                    StartService result.Name
                    Scan(idx).State = ITEM_STATE_ENABLED
                End If
            End If
            
            If result.State = ITEM_STATE_ENABLED Then
            
                MsgBoxW Replace$(Translate(353), "[]", result.Name), vbInformation
            Else
                MsgBoxW Replace$(Translate(354), "[]", result.Name), vbInformation
            End If
            
        End If
    End If
    
End Sub

Private Sub mnuResultCopyLine_Click() ' Context menu => Copy Whole Line
    ClipboardSetText GetSelected_OrCheckedItemResult().HitLineW
End Sub

Private Sub mnuResultCopyRegKey_Click() ' Context menu => Copy Registry Key
    Dim result As SCAN_RESULT
    result = GetSelected_OrCheckedItemResult()
    If AryPtr(result.Reg) Then
        ClipboardSetText BuildPath(Reg.GetHiveNameByHandle(result.Reg(0).Hive), result.Reg(0).Key)
    ElseIf AryPtr(result.Jump) Then
        If AryPtr(result.Jump(0).Registry) Then
            ClipboardSetText BuildPath(Reg.GetHiveNameByHandle(result.Jump(0).Registry(0).Hive), result.Jump(0).Registry(0).Key)
        End If
    End If
End Sub

Private Sub mnuResultCopyRegParam_Click() ' Context menu => Copy Registry Parameter
    Dim result As SCAN_RESULT
    result = GetSelected_OrCheckedItemResult()
    If AryPtr(result.Reg) Then
        ClipboardSetText result.Reg(0).Param
    ElseIf AryPtr(result.Jump) Then
        If AryPtr(result.Jump(0).Registry) Then
            ClipboardSetText result.Jump(0).Registry(0).Param
        End If
    End If
End Sub

Private Function GetSelectedFilePath() As String
    Dim result As SCAN_RESULT
    Dim sFile As String
    result = GetSelected_OrCheckedItemResult()
    If AryPtr(result.File) Then
        GetSelectedFilePath = result.File(0).Path
    ElseIf AryPtr(result.Jump) Then
        If AryPtr(result.Jump(0).File) Then
            GetSelectedFilePath = result.Jump(0).File(0).Path
        End If
    End If
End Function

Private Sub mnuResultCopyFilePath_Click() ' Context menu => Copy File Path
    ClipboardSetText GetSelectedFilePath()
End Sub

Private Sub mnuResultCopyFileArguments_Click() ' Context menu => Copy File Argument
    Dim result As SCAN_RESULT
    result = GetSelected_OrCheckedItemResult()
    If AryPtr(result.File) Then
        ClipboardSetText result.File(0).Arguments
    ElseIf AryPtr(result.Jump) Then
        If AryPtr(result.Jump(0).File) Then
            ClipboardSetText result.Jump(0).File(0).Arguments
        End If
    End If
End Sub

Private Sub mnuResultCopyFileName_Click() ' Context menu => Copy File Name
    ClipboardSetText GetFileName(GetSelectedFilePath(), True)
End Sub

Private Sub mnuResultCopyFileObject_Click() ' Context menu => Copy File (as Object)
    Dim sFile As String: sFile = GetSelectedFilePath()
    If Len(sFile) <> 0 Then
        If FileExists(sFile) Then
            Call ShellExecute(g_HwndMain, StrPtr("copy"), StrPtr(sFile), 0&, 0&, 1)
        End If
    End If
End Sub

Private Sub mnuResultCopyFileHashMD5_Click() ' Context menu => Copy Hash MD5
    Dim sFile As String: sFile = GetSelectedFilePath()
    If Len(sFile) <> 0 Then
        If FileExists(sFile) Then ClipboardSetText GetFileMD5(sFile, , True, False)
    End If
End Sub

Private Sub mnuResultCopyFileHashSHA1_Click() ' Context menu => Copy Hash SHA1
    Dim sFile As String: sFile = GetSelectedFilePath()
    If Len(sFile) <> 0 Then
        If FileExists(sFile) Then ClipboardSetText GetFileSHA1(sFile, , True, False)
    End If
End Sub

Private Sub mnuResultCopyFileHashSHA256_Click() ' Context menu => Copy Hash SHA256
    Dim sFile As String: sFile = GetSelectedFilePath()
    If Len(sFile) <> 0 Then
        If FileExists(sFile) Then ClipboardSetText GetFileSHA256(sFile, , True, False)
    End If
End Sub

Private Sub mnuResultCopyValue_Click() ' Context menu => Copy => Value
    'There is no specific field to hold the value for multipurpose.
    'So, we're just using guessed parsing to retrieve the very end of the line
    Dim pos&
    Dim result As SCAN_RESULT
    result = GetSelected_OrCheckedItemResult()
    pos = InStr(1, result.HitLineW, "=")
    If pos <> 0 Then
        ClipboardSetText LTrim$(mid$(result.HitLineW, pos + 1))
    Else
        pos = InStr(1, result.HitLineW, ":\")
        If pos <> 0 Then
            ClipboardSetText mid$(result.HitLineW, pos - 1)
        Else
            pos = InStrRev(result.HitLineW, "-")
            If pos <> 0 Then
                ClipboardSetText LTrim$(mid$(result.HitLineW, pos + 1))
            End If
        End If
    End If
End Sub

Private Sub mnuResultVTHash_Click() 'Context menu => VirusTotal => Scan by Hash
    
    Dim result As SCAN_RESULT
    Dim sSha256 As String
    Dim bURL_Based As Boolean
    Dim sURL As String
    
    result = GetSelected_OrCheckedItemResult()
    
    If StrBeginWith(result.HitLineW, "O1 - Hosts:") Then
        sURL = GetStringToken(result.HitLineW, 5)
        If Not StrBeginWith(sURL, "http:") And Not StrBeginWith(sURL, "https:") Then
            sURL = "https://" & sURL & "/"
            sSha256 = CalcSha256(sURL)
        End If
        bURL_Based = True
    Else
        If AryPtr(result.File) Then
            If FileExists(result.File(0).Path) Then
                sSha256 = GetFileSHA256(result.File(0).Path, , True)
            End If
        Else
            If StrBeginWith(result.HitLineW, "O17") Then
                sSha256 = CalcSha256("http://" & result.Custom(0).Name & "/") 'don't touch http!
                bURL_Based = True
            End If
        End If
    End If
    
    If Len(sSha256) <> 0 Then
        sSha256 = LCase$(sSha256)
        sURL = Caes_Decode("iwywB://NPR.UJUZZ]ZaP].Xff") 'https://www.virustotal.com
        If bURL_Based Then
            sURL = sURL & "/gui/url/" & sSha256 & "/detection"
        Else
            sURL = sURL & "/gui/file/" & sSha256 & "/detection"
        End If
        OpenURL sURL
    End If
    
End Sub

Private Sub mnuResultVTSubmit_Click() 'Context menu => VirusTotal => Submit with 'AutoRuns'
    
    Dim result As SCAN_RESULT
    
    result = GetSelected_OrCheckedItemResult()
    
    If AryPtr(result.File) Then
        If DownloadAuto_runs() Then
            Call AR_CheckFile(result.File(0).Path)
        End If
    End If
    
End Sub

Private Sub mnuResultSearch_Click()       'Context menu => Search on Google
    Dim sItem$, pos&
    sItem = lstResults.List(lstResults.ListIndex)
    pos = InStr(sItem, ":")
    If pos > 0 Then
        sItem = mid$(sItem, pos + 1)
    End If
    pos = InStr(sItem, " (size: ")
    If pos > 0 Then
        sItem = Left$(sItem, pos - 1)
    End If
    OpenURL "https://www.google.com/search?ie=UTF-8&q=" & URLEncode(sItem)
End Sub

Private Sub mnuResultReScan_Click()       'Context menu => ReScan
    cmdScan.Caption = Translate(11): cmdScan.Tag = "1"
    'cmdScan_Click
    tmrRunScan.Enabled = True
End Sub

Private Sub mnuSaveReport_Click()         'Context menu => Save report...
    Call HJT_SaveReport
End Sub

'test stuff - BUTTON: enum tasks to CSV
Private Sub cmdTaskScheduler_Click()
    Call EnumTasksVista(True)
End Sub

Private Sub chkHelp_Click(Index As Integer)
    Static LastIdx As Long

    Dim i As Long, j As Long
    Dim sText As String
    Dim sSeparator$
    Dim aSect() As Variant
    
    TextBox_SetMargin txtHelp, 20, 20
    
    frmMain.pictLogo.Visible = False
    lblInfo(0).Visible = False
    lblInfo(1).Visible = False
    
    If bSwitchingTabs Then Exit Sub
    If frmMain.cmdHidden.Visible And frmMain.cmdHidden.Enabled Then
        frmMain.cmdHidden.SetFocus
    End If
    bSwitchingTabs = True
    
    chkHelp(Index).Value = 1
    
    For i = 0 To chkHelp.Count - 1
        If Index <> i Then
            chkHelp(i).Value = 0
            'chkHelp(i).Enabled = True
            chkHelp(i).ForeColor = vbBlack
        Else
            'chkHelp(i).Enabled = False
            chkHelp(i).ForeColor = vbBlue
        End If
    Next
    
    Select Case Index
    
    Case 0: 'Sections
        NotifyChangeFrame FRAME_ALIAS_HELP_SECTIONS
    
        aSect = Array("R0", "R1", "R2", "R3", "R4", "F0", "F1", "F2", "F3", "B", "O1", "O2", "O3", "O4", "O5", "O6", "O7", "O8", "O9", "O10", _
            "O11", "O12", "O13", "O14", "O15", "O16", "O17", "O18", "O19", "O20", "O21", "O22", "O23", "O24", "O25", "O26", "O27")
        
        sText = Translate(31) & vbCrLf & vbCrLf & Translate(490)
        sSeparator = String$(100, "-")
        
        For i = 0 To UBound(aSect)
            j = GetTranslationIndex_HelpSection(CStr(aSect(i)))

            sText = sText & vbCrLf & sSeparator & vbCrLf & FindLine(aSect(i) & " -", Translate(31)) & vbCrLf & sSeparator & vbCrLf & _
                Replace$(Translate(j), "\\p", vbNullString) & vbCrLf
        Next
        
        TextBox_SetUnlimitSize txtHelp, Len(sText)
        'txtHelp.Text = sText
        SendMessage txtHelp.hWnd, WM_SETTEXT, 0&, ByVal StrPtr(sText)
        
    Case 1: 'Keys
        NotifyChangeFrame FRAME_ALIAS_HELP_KEYS
        txtHelp.Text = Translate(32)
    
    Case 2: 'Purpose, Donations
        NotifyChangeFrame FRAME_ALIAS_HELP_PURPOSE
        txtHelp.Text = Translate(33) & TranslateNative(34)
    
    Case 3: 'History (Version history)
        NotifyChangeFrame FRAME_ALIAS_HELP_HISTORY
        TextBox_SetUnlimitSize txtHelp, Len(g_VersionHistory)
        'txtHelp.Text = g_VersionHistory 'VB6's .Text property doesn't support over limit!
        SendMessage txtHelp.hWnd, WM_SETTEXT, 0&, ByVal StrPtr(g_VersionHistory)
    End Select
    
    bSwitchingTabs = False
    LastIdx = Index
End Sub

Private Sub TextBox_SetUnlimitSize(txt As VBCCR17.TextBoxW, Optional iMaxSize As Long)
    
    Dim iSize As Long
    
    If iMaxSize <> 0 Then iSize = iMaxSize + 2
    
    SendMessage txt.hWnd, EM_LIMITTEXT, iSize, ByVal 0&
End Sub

Private Sub TextBox_SetMargin(txt As VBCCR17.TextBoxW, left_margin As Long, right_margin As Long)
    
    SendMessage txt.hWnd, EM_SETMARGINS, EC_LEFTMARGIN Or EC_RIGHTMARGIN, ByVal (right_margin * &H10000 + left_margin)
    
    ' Reset the text to make the right margin work
    Dim s As String
    s = txt.Text
    txt.Text = vbNullString
    txt.Text = s
End Sub

Function FindLine(sPartialText As String, sFullText As String) As String
    Dim arr() As String, i&
    arr = Split(sFullText, vbCrLf)
    If AryItems(arr) Then
        For i = 0 To UBound(arr)
            If InStr(1, arr(i), sPartialText, 1) <> 0 Then FindLine = arr(i): Exit For
        Next
    End If
End Function

Private Sub cmdProcessManager_Click() 'Misc Tools -> Process Manager
    frmProcMan.Show
End Sub

'Scan area frame
Private Sub chkLogProcesses_Click() ' Scan Area => Processes
    bLogProcesses = (chkLogProcesses.Value = 1)
    RegSaveHJT "LogProcesses", Abs(CLng(bLogProcesses))
End Sub

Private Sub chkAdvLogEnvVars_Click() ' Scan Area => Environment Variables
    bLogEnvVars = (chkAdvLogEnvVars.Value = 1)
    RegSaveHJT "LogEnvVars", Abs(CLng(bLogEnvVars))
End Sub

Private Sub chkAdditionalScan_Click() ' Scan Area => Additional Scan
    bAdditional = (chkAdditionalScan.Value = 1)
    RegSaveHJT "LogAdditional", Abs(CLng(bAdditional))
End Sub

'Backup & Fix frame
Private Sub chkBackup_Click() ' Fix & Backup => Make backups before fixing items
    bMakeBackup = (chkBackup.Value = 1)
    RegSaveHJT "MakeBackup", Abs(CLng(bMakeBackup))
End Sub

Private Sub chkConfirm_Click() ' Fix & Backup => Confirm fixing & ignoring of items (safe mode)
    bConfirm = (chkConfirm.Value = 1)
    RegSaveHJT "Confirm", Abs(CLng(bConfirm))
End Sub

Private Sub chkAutoMark_Click() 'Fix & Backup => Mark everything found for fixing after scan (DANGEROUS !!!)
    Dim sMsg$
    If chkAutoMark.Value = 0 Then
        bAutoSelect = False
        Exit Sub
    ElseIf RegReadHJT("SeenAutoMarkWarning", "0") = "1" Then
        bAutoSelect = True
        Exit Sub
    End If
    
    sMsg = Translate(57)
'    sMsg = "Are you sure you want to enable this option?" & vbCrLf & _
'           "HiJackThis is not a 'click & fix' program. " & _
'           "Because it targets *general* hijacking methods, " & _
'           "false positives are a frequent occurrence." & vbCrLf & _
'           "If you enable this option, you might disable " & _
'           "programs or drivers you need. However, it is " & _
'           "highly unlikely you will break your system " & _
'           "beyond repair. So you should only enable this " & _
'           "option if you know what you're doing!"
    
    If MsgBoxW(sMsg, vbExclamation + vbYesNo) = vbYes Then
        RegSaveHJT "SeenAutoMarkWarning", "1"
        bAutoSelect = True
        Exit Sub
    Else
        chkAutoMark.Value = 0
    End If
End Sub

'Scan options frame
Private Sub chkIgnoreMicrosoft_Click() ' Scan options => Hide Microsoft entries
    bHideMicrosoft = chkIgnoreMicrosoft.Value
    RegSaveHJT "HideMicrosoft", Abs(CLng(bHideMicrosoft))
End Sub

Private Sub chkIgnoreAll_Click() ' Scan options => Ignore ALL Whitelists
    bIgnoreAllWhitelists = chkIgnoreAll.Value
    RegSaveHJT "IgnoreAllWhiteList", Abs(CLng(bIgnoreAllWhitelists))
End Sub

Private Sub chkDoCheckSum_Click() ' Scan options => Calculate Checksum
    g_bCheckSum = (chkDoCheckSum.Value = 1)
    RegSaveHJT "CalcMD5", Abs(CLng(g_bCheckSum))
    cmbHashType.Enabled = g_bCheckSum
End Sub

Private Sub cmbHashType_Click() ' Scan options => Hash type
    Select Case cmbHashType.Text
    Case "SHA256":  g_eUseHashType = HASH_TYPE_SHA256
    Case "SHA1":    g_eUseHashType = HASH_TYPE_SHA1
    Case "MD5":     g_eUseHashType = HASH_TYPE_MD5
    End Select
End Sub

Private Sub chkConfigStartupScan_Click() ' Scan options => Add HiJackThis to startup
    If gNotUserClick Then gNotUserClick = False: Exit Sub
    If chkConfigStartupScan.Value = 1 Then
        InstallAutorunHJT
    Else
        RemoveAutorunHJT
    End If
End Sub

'Interface frame
'Private Sub chkSkipIntroFrame_Click()
'    RegSaveHJT "SkipIntroFrame", CStr(chkSkipIntroFrame.Value)
'    chkSkipIntroFrameSettings.Value = chkSkipIntroFrame.Value
'End Sub

Private Sub chkSkipIntroFrameSettings_Click() ' Interface => Do not show main menu at startup
    'RegSaveHJT "SkipIntroFrame", CStr(chkSkipIntroFrame.Value) 'should be commented !
End Sub

Private Sub chkSkipErrorMsg_Click() ' Interface => Do not show error messages
    bSkipErrorMsg = (chkSkipErrorMsg.Value = 1)
    RegSaveHJT "SkipErrorMsg", Abs(CLng(bSkipErrorMsg))
End Sub

Private Sub chkConfigMinimizeToTray_Click() ' Interface => Minimize program to system tray when clicking _ button
    bMinToTray = (chkConfigMinimizeToTray.Value = 1)
    RegSaveHJT "MinToTray", Abs(CLng(bMinToTray))
End Sub

'======= UPDATES and PROXY controls

Private Sub cmdCheckUpdate_Click() 'Misc Tools -> Check for update online
    cmdCheckUpdate.Enabled = False
    CheckForUpdate False, bUpdateSilently, bUpdateToTest
    cmdCheckUpdate.Enabled = True
    If g_NeedTerminate Then Unload Me
End Sub

Private Sub chkCheckUpdatesOnStart_Click()
    RegSaveHJT "CheckForUpdates", IIf(chkCheckUpdatesOnStart.Value, 1, 0)
End Sub

Private Sub chkUpdateToTest_Click()
    RegSaveHJT "UpdateToTest", IIf(chkUpdateToTest.Value, 1, 0)
    bUpdateToTest = chkUpdateToTest.Value
End Sub

Private Sub chkUpdateSilently_Click()
    RegSaveHJT "UpdateSilently", IIf(chkUpdateSilently.Value, 1, 0)
    bUpdateSilently = chkUpdateSilently.Value
End Sub

'Some Control Enabled/disabled staff for beautify
Private Sub chkUpdateUseProxyAuth_Click()
    ProxyCtlUpd
    RegSaveHJT "ProxyUseAuth", IIf(chkUpdateUseProxyAuth.Value, 1, 0)
End Sub

Private Sub OptProxyDirect_Click()
    ProxyCtlUpd
    RegSaveHJT "ProxyType", 0
End Sub

Private Sub optProxyIE_Click()
    ProxyCtlUpd
    RegSaveHJT "ProxyType", 1
End Sub

Private Sub optProxyManual_Click()
    ProxyCtlUpd
    RegSaveHJT "ProxyType", 2
End Sub

Private Sub chkSocks4_Click()
    RegSaveHJT "ProxySocksVer", IIf(chkSocks4.Value, 4, 0)
End Sub

Private Sub ProxyCtlUpd()
    lblUpdateServer.Enabled = optProxyManual.Value
    lblUpdatePort.Enabled = optProxyManual.Value
    txtUpdateProxyHost.Enabled = optProxyManual.Value
    txtUpdateProxyPort.Enabled = optProxyManual.Value
    chkSocks4.Enabled = optProxyManual.Value
    chkUpdateUseProxyAuth.Enabled = optProxyManual.Value Or optProxyIE.Value
    
    lblUpdateLogin.Enabled = (optProxyManual.Value Or optProxyIE.Value) And chkUpdateUseProxyAuth.Value
    lblUpdatePass.Enabled = (optProxyManual.Value Or optProxyIE.Value) And chkUpdateUseProxyAuth.Value
    txtUpdateProxyLogin.Enabled = (optProxyManual.Value Or optProxyIE.Value) And chkUpdateUseProxyAuth.Value
    txtUpdateProxyPass.Enabled = (optProxyManual.Value Or optProxyIE.Value) And chkUpdateUseProxyAuth.Value
End Sub

'======== FONT controls

Private Sub cmbFont_Click() ' Interface => Font
    If cmbFontSize.ListCount <> 0 Then
        SetFontByUserSettings
    End If
End Sub

Private Sub cmbFontSize_Click() ' Interface => Size
    SetFontByUserSettings
End Sub

Private Sub chkFontBold_Click() ' Interface => Bold
    SetFontByUserSettings
End Sub

Private Sub SetFontByUserSettings()
    On Error GoTo ErrorHandler:
    
    If bAutoLogSilent Then Exit Sub 'speed optimization
    
    Dim frm As Form
    If cmbFont.ListIndex <> -1 Then
        g_FontName = cmbFont.List(cmbFont.ListIndex)
        If Len(g_FontName) = 0 Then
            g_FontName = "Automatic"
        End If
    End If
    If cmbFontSize.ListIndex <> -1 Then
        g_FontSize = cmbFontSize.List(cmbFontSize.ListIndex)
    End If
    If g_FontSize = "0" Then g_FontSize = "8"
    g_bFontBold = (chkFontBold.Value = vbChecked)
    
    For Each frm In Forms
        SetAllFontCharset frm, g_FontName, g_FontSize, g_bFontBold
        'SetMenuFont Frm.hwnd, g_FontName, g_FontSize
    Next
    
    RegSaveHJT "FontName", g_FontName
    RegSaveHJT "FontSize", g_FontSize
    RegSaveHJT "FontBold", CInt(g_bFontBold)
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SetFontByUserSettings", g_FontName, g_FontSize
    If inIDE Then Stop: Resume Next
End Sub

'Use new Font on result lists (and input windows) only ?
Private Sub chkFontWholeInterface_Click() ' Interface => Apply selected font on whole interface
    RegSaveHJT "FontOnInterface", CStr(Abs(chkFontWholeInterface.Value))
    g_FontOnInterface = chkFontWholeInterface.Value
    SetFontByUserSettings
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ProcessHotkey KeyCode, Me
End Sub

Private Sub lblMD5_Click() ' click on hash progressbar's description => to open the web-page
    Dim sURL As String
    sURL = lblMD5.Tag
    If StrBeginWith(sURL, "http") Then
        lblMD5.Tag = vbNullString
        OpenURL sURL
        If Not g_bScanInProgress Or Not g_bCheckSum Then
            CloseHashProgressbar
        End If
        If Not g_bScanInProgress Then
            frmMain.lblStatus.Visible = False
            lblInfo(1).Visible = True
        End If
        g_bVTScanned = False
    End If
End Sub

'Main menu - Fixing
Private Sub cmdFixing_Click()
    '// TODO
End Sub
