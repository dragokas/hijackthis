VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   7380
   ClientLeft      =   4365
   ClientTop       =   1500
   ClientWidth     =   8775
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   480
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6960
      Top             =   480
   End
   Begin VB.Frame fraOther 
      Caption         =   "Other stuff"
      Height          =   1455
      Left            =   6000
      TabIndex        =   31
      Top             =   5880
      Width           =   2775
      Begin VB.CommandButton cmdSaveDef 
         Caption         =   "Add checked to ignorelist"
         Enabled         =   0   'False
         Height          =   450
         Left            =   240
         TabIndex        =   6
         Top             =   850
         Width           =   2295
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Config"
         Height          =   450
         Left            =   1440
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   450
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraSubmit 
      Height          =   1455
      Left            =   3000
      TabIndex        =   55
      Top             =   5880
      Width           =   2895
      Begin VB.CommandButton cmdAnalyze 
         Caption         =   "Analyze report"
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         TabIndex        =   56
         Top             =   195
         Width           =   1935
      End
      Begin VB.CommandButton cmdMainMenu 
         Caption         =   "Main Menu"
         Height          =   495
         Left            =   720
         TabIndex        =   58
         Top             =   825
         Width           =   1455
      End
   End
   Begin VB.Frame fraScan 
      Caption         =   "Scan && fix stuff"
      Height          =   1455
      Left            =   120
      TabIndex        =   30
      Top             =   5880
      Width           =   2775
      Begin VB.CommandButton CmdHidden2 
         Caption         =   "Focus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   240
         TabIndex        =   88
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Info on selected item..."
         Height          =   450
         Left            =   240
         TabIndex        =   3
         Top             =   850
         Width           =   2340
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmdFix 
         Caption         =   "Fix checked"
         Enabled         =   0   'False
         Height          =   450
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.PictureBox pictLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   7440
      Picture         =   "frmMain.frx":4B2A
      ScaleHeight     =   975
      ScaleWidth      =   1335
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   -20
      Width           =   1335
   End
   Begin VB.CommandButton cmdHidden 
      Default         =   -1  'True
      Height          =   195
      Left            =   24960
      TabIndex        =   87
      Top             =   14760
      Width           =   75
   End
   Begin VB.PictureBox picPaypal 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   6240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmMain.frx":9180
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   68
      Top             =   -450
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame fraHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkHelp 
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "Purpose"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "Keys"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "Sections"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtHelp 
         Height          =   3375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   600
         Width           =   5895
      End
   End
   Begin VB.ListBox lstResults 
      Height          =   1755
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   6135
   End
   Begin VB.TextBox txtNothing 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "No suspicious items found!"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame fraConfig 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CheckBox chkConfigTabs 
         Caption         =   "Misc Tools"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Index           =   3
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Width           =   1455
      End
      Begin VB.CheckBox chkConfigTabs 
         Caption         =   "Backups"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Index           =   2
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1335
      End
      Begin VB.CheckBox chkConfigTabs 
         Caption         =   "Ignorelist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Index           =   1
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   300
         Width           =   1335
      End
      Begin VB.CheckBox chkConfigTabs 
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   440
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Caption         =   "fraConfigBackup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CheckBox chkShowSRP 
            Caption         =   "Show System Restore Points"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   3960
            Width           =   6375
         End
         Begin VB.CommandButton cmdConfigBackupCreateSRP 
            Caption         =   "Create restore point"
            Height          =   720
            Left            =   7440
            TabIndex        =   97
            Top             =   3600
            Width           =   990
         End
         Begin VB.CommandButton cmdConfigBackupCreateRegBackup 
            Caption         =   "Create registry backup"
            Height          =   720
            Left            =   7440
            TabIndex        =   96
            Top             =   2760
            Width           =   990
         End
         Begin VB.CommandButton cmdConfigBackupDeleteAll 
            Caption         =   "Delete all"
            Height          =   495
            Left            =   7440
            TabIndex        =   25
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigBackupDelete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   7440
            TabIndex        =   24
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigBackupRestore 
            Caption         =   "Restore"
            Height          =   495
            Left            =   7440
            TabIndex        =   20
            Top             =   720
            Width           =   975
         End
         Begin VB.ListBox lstBackups 
            Height          =   2385
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   19
            Top             =   720
            Width           =   7215
         End
         Begin VB.Line linSeperator 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   -720
            X2              =   6480
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label lblConfigInfo 
            Caption         =   $"frmMain.frx":91C6
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   42
            Top             =   0
            Width           =   8250
         End
      End
      Begin VB.Frame fraUninstMan 
         Caption         =   "Add/Remove Programs Manager"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.ListBox lstUninstMan 
            Height          =   3540
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   71
            Top             =   960
            Width           =   3855
         End
         Begin VB.CommandButton cmdUninstManUninstall 
            Caption         =   "Uninstall application"
            Height          =   425
            Left            =   4080
            TabIndex        =   86
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdUninstManSave 
            Caption         =   "Save list..."
            Height          =   425
            Left            =   5400
            TabIndex        =   82
            Top             =   3900
            Width           =   1455
         End
         Begin VB.TextBox txtUninstManCmd 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   1880
            Width           =   4095
         End
         Begin VB.TextBox txtUninstManName 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   1200
            Width           =   4095
         End
         Begin VB.CommandButton cmdUninstManRefresh 
            Caption         =   "Refresh list"
            Height          =   425
            Left            =   4080
            TabIndex        =   79
            Top             =   3900
            Width           =   1215
         End
         Begin VB.CommandButton cmdUninstManEdit 
            Caption         =   "Edit uninstall command"
            Height          =   425
            Left            =   6120
            TabIndex        =   78
            Top             =   2835
            Width           =   2055
         End
         Begin VB.CommandButton cmdUninstManBack 
            Caption         =   "Back"
            Height          =   425
            Left            =   6960
            TabIndex        =   76
            Top             =   3900
            Width           =   1215
         End
         Begin VB.CommandButton cmdUninstManDelete 
            Caption         =   "Delete this entry"
            Height          =   425
            Left            =   4080
            TabIndex        =   75
            Top             =   2835
            Width           =   1935
         End
         Begin VB.CommandButton cmdUninstManOpen 
            Caption         =   "Open Add/Remove Software list"
            Height          =   425
            Left            =   4080
            TabIndex        =   74
            Top             =   3360
            Width           =   4150
         End
         Begin VB.Label lblInfo 
            Caption         =   $"frmMain.frx":92AB
            Height          =   615
            Index           =   11
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   7935
         End
         Begin VB.Label lblInfo 
            Caption         =   "Uninstall command:"
            Height          =   255
            Index           =   10
            Left            =   4125
            TabIndex        =   73
            Top             =   1600
            Width           =   1455
         End
         Begin VB.Label lblInfo 
            Caption         =   "Name:"
            Height          =   255
            Index           =   8
            Left            =   4125
            TabIndex        =   72
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Height          =   9120
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   -3480
         Visible         =   0   'False
         Width           =   8055
         Begin VB.VScrollBar vscMiscTools 
            Height          =   4095
            LargeChange     =   20
            Left            =   7680
            Max             =   100
            SmallChange     =   20
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
         Begin VB.Frame fraMiscToolsScroll 
            BorderStyle     =   0  'None
            Height          =   11415
            Left            =   0
            TabIndex        =   54
            Top             =   -1920
            Width           =   7695
            Begin VB.Frame FraRemoveHJT 
               Caption         =   "Uninstall"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   855
               Left            =   120
               TabIndex        =   141
               Top             =   9600
               Width           =   7335
               Begin VB.CommandButton cmdUninstall 
                  Caption         =   "Uninstall HiJackThis"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   142
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.Label lblUninstallHJT 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remove all HiJackThis Registry entries, backups and quit"
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   2640
                  TabIndex        =   143
                  Top             =   400
                  Width           =   4065
               End
            End
            Begin VB.Frame FraPlugins 
               Caption         =   "Plugins"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   1455
               Left            =   120
               TabIndex        =   134
               Top             =   6240
               Width           =   7335
               Begin VB.CommandButton cmdLnkCleaner 
                  Caption         =   "ClearLNK"
                  Height          =   480
                  Left            =   120
                  TabIndex        =   136
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.CommandButton cmdLnkChecker 
                  Caption         =   "Check Browsers' LNK"
                  Height          =   480
                  Left            =   120
                  TabIndex        =   135
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.Label lblLnkCleanerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Clean and restore list of infected shortcuts (.LNK), found via Check Browsers' LNK plugin"
                  Height          =   615
                  Left            =   2520
                  TabIndex        =   138
                  Top             =   800
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblLnkCheckerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Identify infected shortcuts (.LNK) which cause unwanted advertisement in browsers"
                  Height          =   390
                  Left            =   2520
                  TabIndex        =   137
                  Top             =   230
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame FraSysTools 
               Caption         =   "System tools"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   4695
               Left            =   120
               TabIndex        =   117
               Top             =   1440
               Width           =   7335
               Begin VB.CommandButton cmdDigiSigChecker 
                  Caption         =   "Digital signature checker"
                  Height          =   480
                  Left            =   120
                  TabIndex        =   132
                  Top             =   4080
                  Width           =   2295
               End
               Begin VB.CommandButton cmdRegKeyUnlocker 
                  Caption         =   "Registry Key Unlocker"
                  Height          =   480
                  Left            =   120
                  TabIndex        =   129
                  Top             =   3480
                  Width           =   2295
               End
               Begin VB.CommandButton cmdARSMan 
                  Caption         =   "Uninstall Manager..."
                  Height          =   480
                  Left            =   120
                  TabIndex        =   128
                  Top             =   2880
                  Width           =   2295
               End
               Begin VB.CommandButton cmdADSSpy 
                  Caption         =   "ADS Spy..."
                  Height          =   360
                  Left            =   120
                  TabIndex        =   125
                  Top             =   2400
                  Width           =   2295
               End
               Begin VB.CommandButton cmdDeleteService 
                  Caption         =   "Delete a Windows service..."
                  Height          =   360
                  Left            =   120
                  TabIndex        =   124
                  Top             =   1920
                  Width           =   2295
               End
               Begin VB.CommandButton cmdDelOnReboot 
                  Caption         =   "Delete a file on reboot..."
                  Height          =   480
                  Left            =   120
                  TabIndex        =   121
                  Top             =   1320
                  Width           =   2295
               End
               Begin VB.CommandButton cmdHostsManager 
                  Caption         =   "Hosts file manager"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   120
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.CommandButton cmdProcessManager 
                  Caption         =   "Process manager"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   118
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.Label lblDigiSigCheckerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Verify authenticode digital signatures on the given list of files"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   133
                  Top             =   4200
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblRegKeyUnlockerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Reset permissions on the given registry keys list"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   131
                  Top             =   3600
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblARSManAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Managing items in the Add/Remove Software list"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   130
                  Top             =   3000
                  Width           =   4410
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblADSSpyAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Scan for hidden data streams"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   127
                  Top             =   2460
                  Width           =   4665
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblDeleteServiceAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete a Windows Service (O23). USE WITH CAUTION!"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   126
                  Top             =   1900
                  Width           =   4660
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblDelOnRebootAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "If a file cannot be removed from the disk, Windows can be setup to delete it when the system is restarted"
                  Height          =   390
                  Left            =   2520
                  TabIndex        =   123
                  Top             =   1320
                  Width           =   4695
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblHostsManagerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Editor for the 'hosts' file"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   122
                  Top             =   900
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblProcessManagerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Small process manager, working much like the Task Manager"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   119
                  Top             =   420
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame FraStartupList 
               Caption         =   "StartupList"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1335
               Left            =   120
               TabIndex        =   114
               Top             =   0
               Width           =   7335
               Begin VB.CommandButton cmdStartupList 
                  Caption         =   "StartupList scan"
                  Height          =   465
                  Left            =   120
                  TabIndex        =   115
                  Top             =   480
                  Width           =   2295
               End
               Begin VB.Label lblStartupListAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   $"frmMain.frx":937D
                  Height          =   915
                  Left            =   2520
                  TabIndex        =   116
                  Top             =   200
                  Width           =   4635
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame FraUpdateCheck 
               Caption         =   "Update check"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   1695
               Left            =   120
               TabIndex        =   109
               Top             =   7800
               Width           =   7335
               Begin VB.TextBox txtCheckUpdateProxy 
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   113
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2535
               End
               Begin VB.CheckBox chkCheckUpdatesOnStart 
                  Caption         =   "Check updates automatically on program start"
                  Height          =   375
                  Left            =   2520
                  TabIndex        =   111
                  Top             =   360
                  Width           =   4695
               End
               Begin VB.CommandButton cmdCheckUpdate 
                  Caption         =   "Check for update online"
                  Height          =   480
                  Left            =   240
                  TabIndex        =   110
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.Label lblVersionRaw 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "x.x.x.x"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   140
                  Top             =   1320
                  Width           =   540
               End
               Begin VB.Label lblVersion 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Current version is:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   139
                  Top             =   1320
                  Width           =   2055
               End
               Begin VB.Label lblUseProxy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Use this proxy server (host:port) :"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   112
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2490
               End
            End
            Begin VB.Frame FraTestStaff 
               Caption         =   "Testing staff"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               TabIndex        =   107
               Top             =   10560
               Visible         =   0   'False
               Width           =   7335
               Begin VB.CommandButton cmdTaskScheduler 
                  Caption         =   "Task Scheduler Log"
                  Height          =   345
                  Left            =   240
                  TabIndex        =   108
                  Top             =   360
                  Width           =   2055
               End
            End
         End
      End
      Begin VB.Frame fraHostsMan 
         Caption         =   "Hosts file manager"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CommandButton cmdHostsManOpen 
            Caption         =   "Open in Notepad"
            Height          =   425
            Left            =   3600
            TabIndex        =   52
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CommandButton cmdHostsManBack 
            Caption         =   "Back"
            Height          =   425
            Left            =   5160
            TabIndex        =   51
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton cmdHostsManToggle 
            Caption         =   "Toggle line(s)"
            Height          =   425
            Left            =   1800
            TabIndex        =   50
            Top             =   3240
            Width           =   1695
         End
         Begin VB.CommandButton cmdHostsManDel 
            Caption         =   "Delete line(s)"
            Height          =   425
            Left            =   120
            TabIndex        =   49
            Top             =   3240
            Width           =   1575
         End
         Begin VB.ListBox lstHostsMan 
            Height          =   2340
            IntegralHeight  =   0   'False
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   48
            Top             =   600
            Width           =   8175
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Note: changes to the hosts file take effect when you restart your browser."
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   53
            Top             =   3000
            Width           =   5415
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Hosts file located at: C:\WINDOWS\hosts"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   2985
         End
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Caption         =   "fraConfigIgnorelist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CommandButton cmdConfigIgnoreDelSel 
            Caption         =   "Delete"
            Height          =   495
            Left            =   7440
            TabIndex        =   22
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigIgnoreDelAll 
            Caption         =   "Delete all"
            Height          =   495
            Left            =   7440
            TabIndex        =   23
            Top             =   1320
            Width           =   975
         End
         Begin VB.ListBox lstIgnore 
            Height          =   2625
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   21
            Top             =   480
            Width           =   7215
         End
         Begin VB.Label lblConfigInfo 
            Caption         =   "The following items will be ignored when scanning: "
            Height          =   585
            Index           =   5
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   7140
         End
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Caption         =   "fraConfigMain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4250
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   8440
         Begin VB.VScrollBar vscSettings 
            Height          =   4160
            LargeChange     =   20
            Left            =   8040
            Max             =   100
            TabIndex        =   147
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Frame fraConfigTabsNested 
            BorderStyle     =   0  'None
            Height          =   7815
            Left            =   0
            TabIndex        =   144
            Top             =   0
            Width           =   8055
            Begin VB.Frame FraInterface 
               Caption         =   "Interface"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   0
               TabIndex        =   145
               Top             =   3120
               Width           =   7935
               Begin VB.CheckBox chkConfigMinimizeToTray 
                  Caption         =   "Minimize program to system tray when clicking _ button"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   90
                  Top             =   840
                  Width           =   6015
               End
               Begin VB.CheckBox chkSkipErrorMsg 
                  Caption         =   "Do not show error messages"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   89
                  Top             =   600
                  Width           =   4695
               End
               Begin VB.CheckBox chkSkipIntroFrameSettings 
                  Caption         =   "Do not show main menu at startup"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   69
                  Top             =   360
                  Width           =   6975
               End
            End
            Begin VB.Frame FraIncludeSections 
               Caption         =   "Scan area"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1575
               Left            =   0
               TabIndex        =   99
               Top             =   120
               Width           =   3975
               Begin VB.CheckBox chkAdditionalScan 
                  Caption         =   "Additional scan"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   103
                  ToolTipText     =   "Include specific sections, like O4 - RenameOperations, O21 - Column Hanlders / Context menu, O23 - Drivers e.t.c."
                  Top             =   1080
                  Width           =   3015
               End
               Begin VB.CheckBox chkAdvLogEnvVars 
                  Caption         =   "Environment variables"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   102
                  ToolTipText     =   "Include environment variables in logfile"
                  Top             =   720
                  Width           =   3015
               End
               Begin VB.CheckBox chkLogProcesses 
                  Caption         =   "Processes"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   101
                  ToolTipText     =   "Include list of running processes in logfiles"
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   3015
               End
            End
            Begin VB.Frame FraFixing 
               Caption         =   "Fix && Backup"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   0
               TabIndex        =   146
               Top             =   1800
               Width           =   7935
               Begin VB.TextBox txtDefStartPage 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   15
                  Top             =   1560
                  Width           =   5175
               End
               Begin VB.TextBox txtDefSearchPage 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   16
                  Top             =   1920
                  Width           =   5175
               End
               Begin VB.TextBox txtDefSearchAss 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   17
                  Top             =   2280
                  Width           =   5175
               End
               Begin VB.TextBox txtDefSearchCust 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   18
                  Top             =   2640
                  Width           =   5175
               End
               Begin VB.CheckBox chkConfirm 
                  Caption         =   "Confirm fixing && ignoring of items (safe mode)"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   14
                  Top             =   600
                  Value           =   1  'Checked
                  Width           =   7455
               End
               Begin VB.CheckBox chkBackup 
                  Caption         =   "Make backups before fixing items"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   13
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   7335
               End
               Begin VB.CheckBox chkAutoMark 
                  Caption         =   "Mark everything found for fixing after scan (DANGEROUS !!!)"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   12
                  Top             =   840
                  Width           =   7335
               End
               Begin VB.Label lblConfigInfo 
                  Caption         =   "Below URLs will be used when fixing hijacked/unwanted MSIE pages:"
                  Height          =   195
                  Index           =   3
                  Left            =   120
                  TabIndex        =   41
                  Top             =   1200
                  Width           =   7305
               End
               Begin VB.Label lblConfigInfo 
                  AutoSize        =   -1  'True
                  Caption         =   "Default Start Page:"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   39
                  Top             =   1560
                  Width           =   1395
               End
               Begin VB.Label lblConfigInfo 
                  AutoSize        =   -1  'True
                  Caption         =   "Default Search Page:"
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   38
                  Top             =   1920
                  Width           =   1530
               End
               Begin VB.Label lblConfigInfo 
                  AutoSize        =   -1  'True
                  Caption         =   "Default Search Assistant:"
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   37
                  Top             =   2280
                  Width           =   1830
               End
               Begin VB.Label lblConfigInfo 
                  AutoSize        =   -1  'True
                  Caption         =   "Default Search Customize:"
                  Height          =   195
                  Index           =   4
                  Left            =   120
                  TabIndex        =   36
                  Top             =   2640
                  Width           =   1905
               End
            End
            Begin VB.Frame fraScanOpt 
               Caption         =   "Scan options"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1575
               Left            =   4080
               TabIndex        =   100
               Top             =   120
               Width           =   3855
               Begin VB.CheckBox chkConfigStartupScan 
                  Caption         =   "Add HiJackThis to startup"
                  Height          =   270
                  Left            =   120
                  TabIndex        =   83
                  ToolTipText     =   "Run HiJackThis scan at Windows startup and show results (if only items are found)"
                  Top             =   1120
                  Width           =   3255
               End
               Begin VB.CheckBox chkDoMD5 
                  Caption         =   "Calculate MD5"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   106
                  ToolTipText     =   "Calculate checksum of files (MD5) is possible"
                  Top             =   900
                  Width           =   3015
               End
               Begin VB.CheckBox chkIgnoreAll 
                  Caption         =   "Ignore ALL Whitelists"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   105
                  ToolTipText     =   "Include in log any entries regardless whitelist"
                  Top             =   610
                  Width           =   3015
               End
               Begin VB.CheckBox chkIgnoreMicrosoft 
                  Caption         =   "Hide Microsoft entries"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   104
                  ToolTipText     =   "Do not include in log files and registry related to Microsoft"
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   3015
               End
            End
         End
      End
   End
   Begin VB.Frame fraN00b 
      Caption         =   "Main menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   120
      TabIndex        =   59
      Top             =   1080
      Visible         =   0   'False
      Width           =   8655
      Begin VB.ComboBox cboN00bLanguage 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdN00bScan 
         Caption         =   "Do a system scan only"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   62
         Top             =   1440
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bHJTQuickStart 
         Caption         =   "Online Guide"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   65
         Top             =   3960
         Width           =   3975
      End
      Begin VB.CheckBox chkSkipIntroFrame 
         Caption         =   "Do not show this menu after starting the program"
         Height          =   255
         Left            =   360
         TabIndex        =   67
         Top             =   5520
         Width           =   5535
      End
      Begin VB.CommandButton cmdN00bClose 
         Caption         =   "None of above, just start the program"
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   66
         Top             =   4560
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bTools 
         Caption         =   "Misc Tools"
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   64
         Top             =   3000
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bBackups 
         Caption         =   "List of Backups"
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   63
         Top             =   2400
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bLog 
         Caption         =   "Do a system scan and save a logfile"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   61
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Change language:"
         Height          =   195
         Index           =   9
         Left            =   6480
         TabIndex        =   84
         Top             =   360
         Width           =   1320
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   10
         X1              =   480
         X2              =   4200
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   8
         X1              =   480
         X2              =   4200
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "What would you like to do?"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   60
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label lblMD5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calculating MD5 checksum of [file]..."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   600
      TabIndex        =   44
      Top             =   600
      Visible         =   0   'False
      Width           =   5595
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   400
      TabIndex        =   45
      Top             =   330
      Visible         =   0   'False
      Width           =   465
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
      Width           =   5655
   End
   Begin VB.Shape shpMD5Background 
      BackStyle       =   1  'Opaque
      Height          =   120
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Shape shpMD5Progress 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   120
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":9445
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   40
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":951D
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   40
      Width           =   7455
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
         Begin VB.Menu mnuToolsUnlockAndDelFile 
            Caption         =   "Unlock && Delete File..."
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
      Caption         =   "Basic manual"
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
         Begin VB.Menu mnuHelpManualAddition 
            Caption         =   "Additions to manual"
         End
         Begin VB.Menu mnuHelpManualSections 
            Caption         =   "Sections' description"
         End
         Begin VB.Menu mnuHelpManualCmdKeys 
            Caption         =   "Command line keys"
         End
      End
      Begin VB.Menu mnuHelpSupport 
         Caption         =   "Support"
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
      Visible         =   0   'False
      Begin VB.Menu mnuResultFix 
         Caption         =   "Fix checked"
      End
      Begin VB.Menu mnuResultAddToIgnore 
         Caption         =   "Add to ignore list"
      End
      Begin VB.Menu mnuResultAddALLToIgnore 
         Caption         =   "Add ALL to ignore list"
      End
      Begin VB.Menu mnuResultDelim1 
         Caption         =   "-"
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' HJT Main form
'

' Call stack note:
'
' "Do a system scan and save log file" button calls:
'    -> cmdN00bLog_Click -> cmdScan_Click -> StartScan -> HJT_SaveReport -> CreateLogFile (process list)
'
' App key: HKLM\Software\TrendMicro\HiJackThisFork

Option Explicit

Private Type AppVersion
    Major       As Long
    Minor       As Long
    Revision    As Long
    Build       As Long
End Type

Private Type UnintallManagerData
    AppRegKey    As String
    DisplayName  As String
    UninstString As String
    KeyTime      As String
End Type

Private ControlsEvent() As New clsEvents
Private WithEvents FormSys As frmSysTray
Attribute FormSys.VB_VarHelpID = -1

'Private txtHelpHasFocus As Boolean
Private AppVersion      As AppVersion
Private UninstData()    As UnintallManagerData
Private sKeyUninstall() As String
Private bSwitchingTabs  As Boolean
Private bIsBeta         As Boolean
Private bIsAlpha        As Boolean
Private hMutex          As Long
Private lToolsHeight    As Long
Private bLoadDefaults   As Boolean

Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExW" (ByVal lpLibFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long

Public Sub Test()
    'If you need something to test after program started and initialized all required variables, please use this sub.
    
    'Debug.Print Reg.GetString(0, "HKLM\System\CurrentControlSet\Control\Session Manager", "PendingFileRenameOperations")
    
    'Debug.Print Reg.GetKeyVirtualType(HKCU, "Software\Policies\Microsoft\Internet Explorer\Control Panel")
    '2 - KEY_VIRTUAL_USUAL
    '4 - KEY_VIRTUAL_SHARED
    '8 - KEY_VIRTUAL_REDIRECTED
    
    'Stop
End Sub


Private Sub Form_Load()
    Static bInit As Boolean
    
    pvSetFormIcon Me
    
    g_HwndMain = Me.hwnd
    
    If bInit Then
        If Not bAutoLogSilent Then
            MsgBoxW "Critical error. Main form is initialized twice!"
            Unload Me
            Exit Sub
        End If
    Else
        bInit = True
        If gNoGUI Then Me.Hide
        FormStart_Stady1
        tmrStart.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    If (GetTickCount() - Perf.StartTime) / 1000 > Perf.MAX_TimeOut Then
        HJT_Shutdown
    End If
End Sub

Private Sub tmrStart_Timer()
    tmrStart.Enabled = False
    If Not gNoGUI Then Me.Show vbModeless
    FormStart_Stady2
End Sub

Private Sub FormStart_Stady1()
    
    On Error GoTo ErrorHandler:
    
    Dim ctl   As Control
    Dim Btn   As CommandButton
    Dim ChkB  As CheckBox
    Dim OptB  As OptionButton
    Dim Fra   As Frame
    Dim i     As Long
    Dim Salt  As String
    Dim Ver   As Variant
    
    AppendErrorLogCustom "frmMain.FormStart_Stady1 - Begin"

    bIsAlpha = False
    bIsBeta = True
    
    StartupListVer = "2.12"
    ADSspyVer = "1.13"
    ProcManVer = "2.06"
    
    g_HJT_Items_Count = 28 'R + F + O26 (for progressbar)

    DisableSubclassing = False
    If inIDE Then DisableSubclassing = True
    
    If bAutoLogSilent Then 'timeout timer
        If Perf.MAX_TimeOut <> 0 Then
            Timer1.Interval = 1000
            Timer1.Enabled = True
        End If
    End If
    
    ' Result -> sWinVersion (global)
    sWinVersion = GetWindowsVersion()   'to get bIsWin64 and so...
    
    AppVersion.Major = 2
    AppVersion.Minor = 6
    ForkVer = "?"
    
    If InStr(AppVerString, ".") <> 0 Then
        Ver = Split(AppVerString, ".")
        If UBound(Ver) = 3 Then
            AppVersion.Major = Ver(0)
            AppVersion.Minor = Ver(1)
            AppVersion.Revision = Ver(2)
            AppVersion.Build = Ver(3)
            ForkVer = AppVersion.Revision & "." & AppVersion.Build
        End If
    End If
    
    AppVer = "HiJackThis Fork " & IIf(bIsBeta, "(Beta) ", IIf(bIsAlpha, "(Alpha) ", vbNullString)) & _
        "by Alex Dragokas v." & AppVerString
    
    lblVersionRaw.Caption = AppVerString & IIf(bIsBeta, " (Beta)", IIf(bIsAlpha, " (Alpha)", vbNullString))
    
    Call PictureBoxRgn(pictLogo, RGB(255, 255, 255))
    
    'enable x64 redirection
    'ToggleWow64FSRedirection True ' -> moved to GetWindowsVersion()
    
    'List of privileges: https://msdn.microsoft.com/en-us/library/windows/desktop/bb530716(v=vs.85).aspx
    '                    https://msdn.microsoft.com/en-us/library/windows/desktop/ee695867(v=vs.85).aspx
    SetCurrentProcessPrivileges "SeDebugPrivilege"
    SetCurrentProcessPrivileges "SeBackupPrivilege"
    SetCurrentProcessPrivileges "SeRestorePrivilege"
    SetCurrentProcessPrivileges "SeTakeOwnershipPrivilege"
    SetCurrentProcessPrivileges "SeSecurityPrivilege"       'SACL
    'SetCurrentProcessPrivileges "SeAssignPrimaryTokenPrivilege" '(SYSTEM, LocalService или NetworkService) only
    SetCurrentProcessPrivileges "SeIncreaseQuotaPrivilege"  'CreateProcessWithTokenW
    SetCurrentProcessPrivileges "SeImpersonatePrivilege"    'CreateProcessWithTokenW
    
    InitVariables   'sWinDir, classes init. and so.
    
    SetCurrentDirectory StrPtr(AppPath())
    
    'FixLog = BuildPath(AppPath(), "\HJT_Fix.log")           'not used yet
    'If FileExists(FixLog) Then DeleteFileWEx StrPtr(FixLog)
    
    bPolymorph = (InStr(1, AppExeName(), "_poly", 1) <> 0) Or (StrComp(GetExtensionName(AppExeName(True)), ".pif", 1) = 0)
    
    If Not bPolymorph Then
        Me.Caption = AppVer
    End If
    
    'testing stuff
    If inIDE Or InStr(1, AppExeName(), "test", 1) <> 0 Or bDebugMode Then
        'Task scheduler jobs log on 'misc section'.
        Me.FraTestStaff.Visible = True
        'cmdTaskScheduler.Visible = True
        lToolsHeight = 0
        'added autoadjustment depending on the top of the most bottom frame
        lToolsHeight = 850 - (FraRemoveHJT.Top + FraRemoveHJT.Height - (9600 + 855))
    Else
        lToolsHeight = 850 - (FraTestStaff.Top - 10560)
    End If
    
    LoadLanguageList
    LoadResources
    
    lblMD5.Caption = ""
    
    ' if Win XP -> disable all window styles from buttons on frames
    If bIsWinXP Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "CommandButton" Then
                Set Btn = ctl
                SetWindowTheme Btn.hwnd, StrPtr(" "), StrPtr(" ")
            ElseIf TypeName(ctl) = "CheckBox" Then
                Set ChkB = ctl
                SetWindowTheme ChkB.hwnd, StrPtr(" "), StrPtr(" ")
            ElseIf TypeName(ctl) = "OptionButton" Then
                Set OptB = ctl
                SetWindowTheme OptB.hwnd, StrPtr(" "), StrPtr(" ")
            End If
        Next
        Set OptB = Nothing
        Set ChkB = Nothing
        Set Btn = Nothing
        Set ctl = Nothing
    End If
    ' disable visual bugs with .caption property of frames (XP+)
    If OSver.MajorMinor >= 5.1 Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                Set Fra = ctl
                'If Fra.Name = "fraHostsMan" Or Fra.Name = "fraUninstMan" Then
                    SetWindowTheme Fra.hwnd, StrPtr(" "), StrPtr(" ")
                'End If
            End If
        Next
        Set Fra = Nothing
    End If
    
    CenterForm Me
    
    ' Set common events for controls
    ReDim ControlsEvent(0)
    'Set ControlsEvent(0).FrmInArr = Me
    For Each ctl In Me.Controls
        i = i + 1
        ReDim Preserve ControlsEvent(0 To i)
        Select Case TypeName(ctl)
            Case "CommandButton"
                Set ControlsEvent(i).BtnInArr = ctl
            Case "TextBox"
                Set ControlsEvent(i).txtBoxInArr = ctl
            Case "ListBox"
                Set ControlsEvent(i).lstBoxInArr = ctl
            'Case "Label"
            '    'Set ControlsEvent(i).LblInArr = ctl
            Case "CheckBox"
                Set ChkB = ctl
                'CheckBoxes in array dosn't support this type of events
                If ChkB.Name <> "chkConfigTabs" And ChkB.Name <> "chkHelp" Then
                    Set ControlsEvent(i).chkBoxInArr = ctl
                End If
        End Select
    Next ctl
    
    SetAllFontCharset
    
    GetHosts
    GetBrowsersInfo
    
    Set Proc = New clsProcess
    UninstMan_Init
    
    'set encryption string
    Salt = Reg.GetDword(0, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "InstallDate")
    If Salt = "0" Then Salt = Reg.GetBinaryToString(0, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "DigitalProductId")
    sProgramVersion = "THOU SHALT NOT STEAL - " & Salt 'don't touch this, please !!!
    cryptInit
    Base64_Init
    
    If bDebugMode Then
        bDebugToFile = True ' /debug also initiate /bDebugToFile
        OpenDebugLogHandle
    End If
    
    LoadStuff 'regvals, filevals, safelspfiles, safeprotocols
    GetLSPCatalogNames
    LoadSettings ' must go after LoadStuff()
    
    If InStr(1, Command$(), "/Area:Process", 1) > 0 Then bLogProcesses = True
    If InStr(1, Command$(), "/Area:Environment", 1) > 0 Then bLogEnvVars = True
    If InStr(1, Command$(), "/Area:Additional", 1) > 0 Then bAdditional = True
    
    If InStr(1, Command$(), "/ihatewhitelists", 1) > 0 Then bIgnoreAllWhitelists = True: bHideMicrosoft = False 'must go after LoadSettings !!!
    If InStr(1, Command$(), "/default", 1) > 0 Then bLoadDefaults = True
    If bLoadDefaults Then
        bAutoSelect = False
        bConfirm = True
        bMakeBackup = True
        'bIgnoreSafeDomains = True
        bLogProcesses = True
        bLogEnvVars = False
        bAdditional = False
        bSkipErrorMsg = False
        bMinToTray = False
        bCheckForUpdates = False
        bHideMicrosoft = True
        bIgnoreAllWhitelists = False
        bMD5 = False
    End If
    If InStr(1, Command$(), "/skipIgnoreList", 1) > 0 Then
        IsOnIgnoreList "", EraseList:=True
    End If
    
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
    
    If Screen.Height >= 9000 Then
        Me.Height = CLng(RegReadHJT("WinHeight", "8000"))
        If Me.Height < 8000 Then Me.Height = 8000
    Else
        Me.Height = CLng(RegReadHJT("WinHeight", "6600"))
        If Me.Height < 6600 Then Me.Height = 6600
    End If
    Me.Width = CLng(RegReadHJT("WinWidth", "9000"))
    If Me.Width < 9000 Then Me.Width = 9000
    
'    If RegReadHJT("SkipIntroFrame", "0") = "0" Or bFirstRun Then
'        fraN00b.Visible = True
'        fraScan.Visible = False
'        fraOther.Visible = False
'        lstResults.Visible = False
'        fraSubmit.Visible = False
'        If Not bFirstRun Then chkSkipIntroFrame.Value = 0
'    Else
'        chkSkipIntroFrame.Value = 1
'        pictLogo.Visible = False
'    End If

    If RegReadHJT("SkipIntroFrame", "0") = "0" Or (ConvertVersionToNumber(RegReadHJT("Version", "")) < ConvertVersionToNumber("2.7.0.11")) Then
        fraN00b.Visible = True
        fraScan.Visible = False
        fraOther.Visible = False
        lstResults.Visible = False
        fraSubmit.Visible = False
        
        Call RegSaveHJT("SkipIntroFrame", "0")
        
    Else
        chkSkipIntroFrame.Value = 1
        pictLogo.Visible = False
    End If
    
    If Not bAutoLogSilent Then
        If Not CheckForReadOnlyMedia() Then
            g_NeedTerminate = True
        End If
    End If
    
    CheckDateFormat
    CheckForStartedFromTempDir
    
    If Not bIsWinNT Then cmdDeleteService.Enabled = False
    
    SetMenuIcons Me.hwnd
    
    AppendErrorLogCustom "frmMain.FormStart_Stady1 - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FormStart_Stady1"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub FormStart_Stady2()
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "frmMain.FormStart_Stady2 - Begin"
    
    Static bInit As Boolean
    
    If bInit Then
        Exit Sub
    Else
        bInit = True
    End If
    
    If g_NeedTerminate Then Unload Me: Exit Sub
    
    If InStr(1, Command$(), "/uninstall", 1) > 0 Then
        Me.Hide
        cmdUninstall_Click
        'Unload Me ' included in cmdUninstall_Click
        Exit Sub
    End If

    If InStr(1, Command$(), "/md5", 1) > 0 Then bMD5 = True
    If InStr(1, Command$(), "/deleteonreboot", 1) > 0 Then
        SilentDeleteOnReboot UnQuote(Command$())
        Unload Me
        Exit Sub
    End If
    
    If (Not inIDE) And (Not bPolymorph) Then
        Err.Clear
        hMutex = CreateMutex(0&, 1&, StrPtr("mutex_HiJackThis_Fork"))
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) And 0 = Len(Command$()) Then
            If Not bAutoLogSilent Then
                If MsgBoxW(Translate(2), vbExclamation Or vbYesNo, g_AppName) = vbNo Then Unload Me: Exit Sub
            End If
        End If
    End If
    
    #If DoCrash Then
        DoCrash
    #End If
    
    If bCheckForUpdates Then
        If Not bAutoLogSilent Then
            CheckForUpdate True
        End If
    End If
    
    If InStr(1, Command(), "/install", 1) <> 0 Then
        If InStr(1, Command(), "/autostart", 1) <> 0 Then
            InstallAutorunHJT True
        Else
            InstallHJT True, (InStr(1, Command(), "/noGUI", 1) <> 0)
        End If
        Unload Me
        Exit Sub
    End If
    
    If bDebugMode Or bDebugToFile Then
        'checking is EDS machanism working correclty
        Dim SignResult As SignResult_TYPE
    
        'check sign. of core dll
        SignVerify BuildPath(sWinDir, "system32\ntdll.dll"), SV_LightCheck Or SV_SelfTest Or SV_PreferInternalSign, SignResult
        Dbg "Fingerprint should be: CDD4EEAE6000AC7F40C3802C171E30148030C072"
        If SignResult.HashRootCert = "CDD4EEAE6000AC7F40C3802C171E30148030C072" Then
            Dbg "Fingerprint is mathed (OK)."
        Else
            Dbg "Fingerprint is NOT mathed (FAILED)."
        End If
        'check sign of self
        SignVerify AppPath(True), SV_LightCheck Or SV_SelfTest Or SV_PreferInternalSign, SignResult
        Dbg "Fingerprint should be: 05F1F2D5BA84CDD6866B37AB342969515E3D912E"
        If SignResult.HashRootCert = "05F1F2D5BA84CDD6866B37AB342969515E3D912E" Then
            Dbg "Fingerprint is mathed (OK)."
        Else
            Dbg "Fingerprint is NOT mathed (FAILED)."
        End If
    End If
    
    Test 'for all of my tests
    
    CheckAutoLog
    
    If (Not inIDE) And Command() = "" Then
        If Not CheckIntegrityHJT() Then
            'Warning! Integrity of HiJackThis program is corrupted. Perhaps, file is patched or infected by file virus.
            MsgBoxW TranslateNative(1023), vbExclamation
        End If
    End If
    
    AppendErrorLogCustom "frmMain.FormStart_Stady2 - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FormStart_Stady2"
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
'        If (Not bFirstRun) And (chkSkipIntroFrame.Value = 1) And cmdScan.Visible Then
'            cmdScan.SetFocus
'        End If
        If (chkSkipIntroFrame.Value = 1) Then
            If cmdScan.Visible And cmdScan.Enabled Then
                cmdScan.SetFocus
            End If
        Else
            If cmdN00bLog.Visible And cmdN00bLog.Enabled Then
                cmdN00bLog.SetFocus
            End If
        End If
    Else
        Me.Hide
    End If
    
    If bAutoLog Then
        'If bAutoLogSilent Then Me.WindowState = vbMinimized
        'If bAutoLogSilent Then Me.WindowState = vbMinimizedNoFocus
        cmdN00bClose_Click
        cmdScan_Click
        If Not bAutoLogSilent Then DoEvents
        If bAutoLogSilent Then Unload Me: Exit Sub
    End If
    
    If InStr(1, Command$(), "/StartupScan", 1) > 0 Then
        'Me.Show
        'DoEvents
        'Me.WindowState = vbMinimized
        cmdN00bClose_Click
        cmdScan_Click
        'DoEvents
        If lstResults.ListCount = 0 Then
            Unload Me: Exit Sub
        Else
            'Me.WindowState = vbNormal
            Me.Show
        End If
    End If
    
'    If InStr(1, command$(), "/SilentStartupList", 1) > 0 Then
'        bStartupList = True
'        cmdN00bTools_Click
'        Call chkConfigTabs_Click(3)
'        cmdStartupList_Click
'        Unload Me: End
'    End If
    
    If InStr(1, Command$(), "/StartupList", 1) > 0 Then
        bStartupListSilent = True
        cmdN00bTools_Click
        Call chkConfigTabs_Click(3)
        cmdStartupList_Click
    End If
    
    If InStr(1, Command$(), "/SysTray", 1) > 0 Then
        bMinToTray = True
        Me.WindowState = vbMinimized
        'Form_Resize
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
    Dim j           As Long
    Dim Columns()   As String
    Dim ID          As Long
    
    'Task Scheduler white list
    sBuf = StrConv(LoadResData(101, "CUSTOM"), vbUnicode, 1049)
    sBuf = Replace$(sBuf, vbCr, vbNullString)
    
    Lines = Split(sBuf, vbLf)
    ReDim g_TasksWL(UBound(Lines))
    
    For i = 1 To UBound(Lines)  'skip header
    
        If 0 <> Len(Lines(i)) Then
    
            Columns = SplitSafe(Lines(i), ";")
            '---------------------------
            'Columns (0) 'OSver
            'Columns (1) 'Dir\Name
            'Columns (2) 'RunObj
            'Columns (3) 'Args
            'Columns (4) 'Note      (not used)
            '---------------------------
        
            With g_TasksWL(i)
                .OSver = Val(Replace$(Columns(0), ",", "."))
                
                'select appropriate version from DB
                If .OSver = OSver.MajorMinor Then
                
                    .Path = UnScreenChar(CStr(Columns(1)))
                    If UBound(Columns) > 1 Then
                        .RunObj = EnvironW(UnScreenChar(CStr(Columns(2))))
                        If Not isCLSID(.RunObj) Then
                            .RunObj = FindOnPath(.RunObj) 'find full path for relative name
                        End If
                    End If
                    
                    If UBound(Columns) > 2 Then
                        .Args = UnScreenChar(EnvironW(CStr(Columns(3))))
                        'normalize Excel quotes
                        .Args = Replace$(.Args, """""", """")
                        .Args = Trim$(UnQuote(.Args))
                    End If
                    
                    'Dictonary 'oDict.TaskWL_ID':
                    'value -> (dir + name of task)
                    'data -> id to 'g_TasksWL' user type array
                    
                    If Not oDict.TaskWL_ID.Exists(.Path) Then
                        oDict.TaskWL_ID.Add .Path, i
                    Else 'append several lines with same paths
                        ID = oDict.TaskWL_ID(.Path)
                        g_TasksWL(ID).RunObj = g_TasksWL(ID).RunObj & "|" & .RunObj
                        g_TasksWL(ID).Args = g_TasksWL(ID).Args & "|" & .Args
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
        End If
    End If
    BackupFlush
    If g_WER_Disabled Then DisableWER bRevert:=True
    
    Dim Frm As Form
    ToggleWow64FSRedirection True
    If Not g_UninstallState Then
        SaveSettings
        If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
            RegSaveHJT "WinHeight", CStr(Me.Height)
            RegSaveHJT "WinWidth", CStr(Me.Width)
        End If
        RegSaveHJT "Version", AppVerString
    End If
    SubClassScroll False
    For Each Frm In Forms
        If Not (Frm Is Me) And Not (Frm.Name = "frmEULA") Then
            Unload Frm
            Set Frm = Nothing
        End If
    Next
    
    If (UnloadMode = 0 Or bmnuExit_Clicked) And isRanHJT_Scan Then End
    If hLibPcre2 <> 0 Then FreeLibrary hLibPcre2: hLibPcre2 = 0
    
    MenuReleaseIcons
    Set HE = Nothing
    'because can still be used by StartupList2
    'Set Reg = Nothing
    
    g_HwndMain = 0
End Sub

Public Sub ReleaseMutex()
    If hMutex <> 0 Then CloseHandle hMutex
End Sub

Private Sub Form_Terminate()
    Set frmStartupList2 = Nothing
    Set ErrLogCustomText = Nothing
    If Not inIDE Then
        If FileExists(BuildPath(AppPath(), "MSComCtl.ocx")) Then
            Proc.ProcessRun AppPath(True), "/release:" & GetCurrentProcessId(), , vbMinimizedNoFocus, True
        End If
    End If
    Set oDictFileExist = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pvDestroyFormIcon Me
    ReleaseMutex
    ISL_Dispatch
    Close
End Sub

Private Sub cmdADSSpy_Click() 'Misc Tools -> ADS Spy
    frmADSspy.Show
End Sub

Private Sub mnuHelpManualBasic_Click()  'Help -> User's manual -> Basic manual
    'cmdN00bHJTQuickStart_Click
    PopupMenu mnuBasicManual
End Sub

Private Sub mnuHelpManualAddition_Click()   'Help -> User's manual -> Additions to manual
    Dim doOpen As Boolean
    
    If bForceEN Or Not (IsRussianLangCode(OSver.LangSystemCode) Or IsRussianLangCode(OSver.LangDisplayCode) Or (g_CurrentLang = "Russian")) Then
        If vbYes = MsgBox("This manual is not available on your language. Open it on Russian anyway?", vbYesNo Or vbQuestion) Then
            doOpen = True
        End If
    Else
        doOpen = True
    End If
    
    If doOpen Then
        OpenURL "https://safezone.cc/threads/27470/"
    End If
End Sub

Private Sub mnuHelpManualCmdKeys_Click()   'Help -> User's manual -> Command line keys
    cmdN00bClose_Click
    ' Закрыть окно "Инструменты"
    If cmdConfig.Caption = Translate(19) Then cmdConfig_Click
    If cmdHelp.Caption = Translate(16) Then cmdHelp_Click
    fraHelp.Visible = True
    fraHelp.ZOrder 0
    chkHelp_Click 1
End Sub

Private Sub mnuHelpManualSections_Click()   'Help -> User's manual -> Sections' description
    cmdN00bClose_Click
    ' Закрыть окно "Инструменты"
    If cmdConfig.Caption = Translate(19) Then cmdConfig_Click
    If cmdHelp.Caption = Translate(16) Then cmdHelp_Click
    fraHelp.Visible = True
    fraHelp.ZOrder 0
    chkHelp_Click 0
End Sub

Private Sub mnuHelpSupport_Click()
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
    
    'old URL: 'http://sourceforge.net/p/hjt/support-requests/
    
'    Dim sLog$, i&, sProcessList$
'    Dim BeginTime   As Date
'    Dim FinishTime  As Date
'    Dim ElapsedTime As Long
'
'    BeginTime = Now
'
'    'Dim gProcess() As MY_PROC_ENTRY
'
'    If GetProcesses_Zw(gProcess) Then
'        For i = 0 To UBound(gProcess)
'            sProcessList = sProcessList & gProcess(i).Name & ";" & gProcess(i).Path & "|"
'        Next
'    End If
'
'    szLogData = sProcessList
'
'    For i = 0 To lstResults.ListCount
'        szLogData = szLogData & lstResults.List(i) & "|"
'    Next i
'
'    If IsOnline Then
'        cmdAnalyze.Caption = Translate(500) '"Please Wait"
'
'        szLogData = ObfuscateData(szLogData)
'
'        Dim sThisVersion, szBuf As String
'        sThisVersion = App.Major & "." & App.Minor & "." & App.Revision
'        cmdAnalyze.Caption = Translate(521)  '"AnalyzeThis"
'        ShellExecute Me.hWnd, StrPtr("open"), StrPtr("http://sourceforge.net/p/hjt/support-requests/"), 0&, 0&, 1
'        Exit Sub
'    End If
'
'    ParseHTTPResponse szBuf
'
'    If Len(szSubmitUrl) > 7 Then
'        ShellExecute Me.hWnd, StrPtr("open"), StrPtr("http://sourceforge.net/p/hjt/support-requests/" & szResponse), 0&, 0&, 1
'        ParseHTTPResponse szResponse
'
'        cmdAnalyze.Enabled = True
'        FinishTime = Now
'        ElapsedTime = DateDiff("s", BeginTime, FinishTime)
'    Else
'        MsgBoxW Translate(501) '"Please go to http://sourceforge.net/p/hjt/support-requests/"
'    End If
'
'    cmdAnalyze.Caption = "AnalyzeThis"
End Sub

Function ObfuscateData(szDataIn As String) As String
    Dim szDataOut As String
    Dim szHexVal As String
    Dim chrCode As Long
    Dim i As Long
    
    szDataOut = "7"
    
    For i = 1 To Len(szDataIn)
        chrCode = Asc(Mid$(szDataIn, i, 1))
        szHexVal = Hex$(chrCode)
        szDataOut = szDataOut & StrReverse(szHexVal)
    Next i
    ObfuscateData = szDataOut
End Function

Private Sub cmdARSMan_Click() 'Misc Tools -> Uninstall Manager
    fraConfigTabs(3).Visible = False
    SubClassScroll False
    fraUninstMan.Visible = True
    cmdUninstManRefresh_Click
End Sub

Private Sub cmdDigiSigChecker_Click() 'Misc Tools -> Digital signature checker
    frmCheckDigiSign.Show vbModeless
End Sub

Private Sub cmdLnkChecker_Click() 'Misc Tools -> Check Browsers' LNK
    mnuToolsShortcutsChecker_Click
End Sub

Private Sub cmdLnkCleaner_Click() 'Misc Tools -> ClearLNK
    mnuToolsShortcutsFixer_Click
End Sub

Private Sub cmdRegKeyUnlocker_Click() 'Mic Tools -> Registry Keys unlocker
    frmUnlockRegKey.Show vbModeless
End Sub

Private Sub cmdDeleteService_Click() 'Misc Tools -> Delete service ...
    If Not bIsWinNT Then Exit Sub
    Dim sServiceName$, sDisplayName$, sFile$, sCompany$, sTmp$, sDllPath$
    Dim Result As SCAN_RESULT
    
    sServiceName = InputBox(Translate(113), Translate(114))
'    sServiceName = InputBox("Enter the exact service name as it appears " & _
'                            "in the scan results, or the short name " & _
'                            "between brackets if that is listed." & vbCrLf & _
'                            "The service needs to be stopped and disabled." & vbCrLf & _
'                            "Services that belong to Microsoft, Symantec " & _
'                            "and several others are system-critical and cannot be deleted." & vbCrLf & vbCrLf & _
'                            "WARNING! When the service is deleted, it " & _
'                            "cannot be restored!", "Delete a Windows NT Service")
    If Len(sServiceName) = 0 Then Exit Sub
    
    If Not IsServiceExists(sServiceName) Then
        sTmp = GetServiceNameByDisplayName(sServiceName)
        If sTmp <> 0 Then
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
    
    sCompany = GetFilePropCompany(IIf(sDllPath <> "", sDllPath, sFile))
    If sCompany = vbNullString Then sCompany = Translate(502) '"Unknown owner" '"?"
    
    If Not FileExists(sFile) Then sFile = sFile & " (" & Translate(503) & ")"  '" (file missing)"
    
    If MsgBoxW(Translate(117) & vbCrLf & _
              Translate(505) & ": " & sServiceName & vbCrLf & _
              Translate(506) & ": " & sDisplayName & vbCrLf & _
              Translate(507) & ": " & sFile & IIf(sDllPath <> "", " -> " & sDllPath, "") & vbCrLf & _
              Translate(508) & ": " & sCompany & vbCrLf & vbCrLf & _
              Translate(118), vbYesNo Or vbDefaultButton2 Or vbExclamation) = vbYes Then
'    If msgboxW("The following service was found:" & vbCrLf & _
'              "Short name: " & sServiceName & vbCrLf & _
'              "Full name: " & sDisplayName & vbCrLf & _
'              "File: " & sFile & vbCrLf & _
'              "Owner: " & sCompany & vbCrLf & vbCrLf & _
'              "Are you absolutely sure you want to delete this service?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
        
        With Result
            .Section = "O23"
            .HitLineW = "O23 - Service: " & sServiceName & " (" & sDisplayName & ")"
            AddServiceToFix .Service, DELETE_SERVICE, sServiceName
            .CureType = SERVICE_BASED
        End With
        
        FixServiceHandler Result
        
    End If
End Sub

Private Sub cmdDelOnReboot_Click() 'Misc Tools -> Delete on reboot ...
    Dim sFilename$
    'Enter file name:, Delete on Reboot
    sFilename = InputBox(Translate(1950), Translate(1951))
    If StrPtr(sFilename) = 0 Then Exit Sub
    'sFileName = CmnDlgOpenFile(Translate(509), Translate(1003) & " (*.*)|*.*|" & Translate(511) & " (*.dll)|*.dll|" & Translate(512) & " (*.exe)|*.exe")
    'sFileName = CmnDlgOpenFile("Enter file to delete on reboot...", "All files (*.*)|*.*|DLL libraries (*.dll)|*.dll|Program files (*.exe)|*.exe")
    If sFilename = vbNullString Then Exit Sub
    DeleteFileOnReboot sFilename, True, True
End Sub

Private Sub cmdHostsManager_Click() 'Misc Tools -> 'Hosts' file manager
    fraConfigTabs(3).Visible = False
    SubClassScroll False
    fraHostsMan.Visible = True
    ListHostsFile lstHostsMan, lblConfigInfo(14)
End Sub

Private Sub cmdHostsManBack_Click()
    fraHostsMan.Visible = False
    fraConfigTabs(3).Visible = True
    SubClassScroll True
End Sub

Private Sub cmdHostsManDel_Click()
    If lstHostsMan.ListIndex <> -1 And lstHostsMan.ListCount > 0 Then
        HostsDeleteLine lstHostsMan
        ListHostsFile lstHostsMan, lblConfigInfo(14)
    End If
End Sub

Private Sub cmdHostsManOpen_Click()
    'ShellExecute Me.hwnd, "open", sWinDir & "\notepad.exe", sHostsFile, vbNullString, 1
    'Shell "rundll32.exe shell32.dll,ShellExec_RunDLL " & """" & sHostsFile & """", vbNormalFocus
    
    Dim sTxtProg As String
    sTxtProg = Reg.GetDefaultProgram(".txt")
    Shell sTxtProg & " " & """" & sHostsFile & """", vbNormalFocus
End Sub

Private Sub cmdHostsManToggle_Click()
    If lstHostsMan.ListIndex <> -1 And lstHostsMan.ListCount > 0 Then
        HostsToggleLine lstHostsMan
        ListHostsFile lstHostsMan, lblConfigInfo(14)
    End If
End Sub

Private Sub cmdMainMenu_Click()

    CloseProgressbar
    
    SubClassScroll False
    frmMain.pictLogo.Visible = True
    If cmdConfig.Caption = Translate(19) Then
    
        AppendErrorLogCustom "SaveSettings initiated by clicking 'Main menu'."
        SaveSettings
        
        'If cmdScan.Caption = "Scan" Then
        If cmdScan.Caption = Translate(11) Then
            lblInfo(0).Visible = True
        Else
            lblInfo(1).Visible = True
        End If
        
        'picPaypal.Visible = True
        fraConfig.Visible = False
        fraHostsMan.Visible = False
        fraUninstMan.Visible = False
        If chkConfigTabs(3).Value = 1 Then fraConfigTabs(3).Visible = True
        'cmdConfig.Caption = "Config..."
        cmdConfig.Caption = Translate(18)
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
    cmdScan.Caption = Translate(11) ' don't touch it !!!
    cmdHelp.Caption = Translate(16)
    lblInfo(0).Visible = True
    lblInfo(1).Visible = False
    chkSkipIntroFrame.Value = RegReadHJT("SkipIntroFrame", "0")
End Sub

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

Private Sub cmdN00bClose_Click()
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    lblInfo(0).Visible = False
    lblInfo(1).Visible = True
    If cmdScan.Visible And cmdScan.Enabled Then
        cmdScan.SetFocus
    End If
End Sub

Private Sub cmdN00bHJTQuickStart_Click()
    'ShellExecute Me.hWnd, "open", "http://tomcoyote.org/hjt/#Top", "", "", 1
    'szQSUrl = Translate(360) & "?hjtver=" & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    
    'szQSUrl = "https://www.bleepingcomputer.com/tutorials/how-to-use-hijackthis/"
    
    OpenURL "https://github.com/dragokas/hijackthis/wiki/HJT:-Tutorial", "https://safezone.cc/threads/25184/", True
End Sub

Private Sub cmdN00bLog_Click()
    
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    bAutoLog = True
    cmdScan_Click
End Sub

Private Sub cmdN00bScan_Click()
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    pictLogo.Visible = False
    cmdScan_Click
End Sub

Private Sub cmdN00bTools_Click()
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    
    'lstResults.Visible = True
    
    'If cmdConfig.Caption = Translate(18) Then cmdConfig_Click
    
    cmdConfig.Caption = Translate(18)
    cmdConfig_Click
    chkConfigTabs_Click 3
End Sub

Private Sub chkConfigTabs_Click(Index As Integer)

    On Error GoTo ErrorHandler:
    
    Static idxLastTab As Long
    Static IsInit As Boolean
    
    Dim i           As Long
    Dim iIgnoreNum  As Long
    Dim sIgnore     As String
    
    If bSwitchingTabs Then Exit Sub
    If frmMain.cmdHidden.Visible And frmMain.cmdHidden.Enabled Then
        frmMain.cmdHidden.SetFocus
    End If
    bSwitchingTabs = True
    
    If idxLastTab = 0 And IsInit Then
        UpdateIE_RegVals
    End If
    
    If Not IsInit Then IsInit = True
    
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
    fraUninstMan.Visible = False
    
    bSwitchingTabs = False
    fraConfig.Visible = True
    
    Select Case Index
    
    Case 0 'main settings
        SubClassScroll False 'unSubClass
        
    Case 1 'ignore list
        SubClassScroll False 'unSubClass
        
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
        SubClassScroll False 'unSubClass
        ListBackups
        
    Case 3 'Misc tools
        SubClassScroll True ' mouse scrolling support
        
    End Select
    
    idxLastTab = Index
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "chkConfigTabs_Click", "idx:" & Index
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdCheckUpdate_Click() 'Misc Tools -> Check for update online
    CheckForUpdate
End Sub

Private Sub cmdConfig_Click()
    On Error GoTo ErrorHandler:

    'сперва выйти из фрейма "Help"
    If cmdHelp.Caption = Translate(17) Then cmdHelp_Click
    
    SubClassScroll True
    
    CloseProgressbar
    
    'If cmdConfig.Caption = "Config..." Then
    If cmdConfig.Caption = Translate(18) Then   'Config
    
        pictLogo.Visible = False
        
        'chkSkipIntroFrameSettings.Value = CLng(RegReadHJT("SkipIntroFrame", "0"))
        
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        picPaypal.Visible = False
        lstResults.Visible = False
        cmdConfig.Caption = Translate(19)
        cmdSaveDef.Enabled = False
        fraScan.Enabled = False
        cmdScan.Enabled = False
        cmdFix.Enabled = False
        cmdInfo.Enabled = False
        txtNothing.Visible = False
        
        'fraConfigTabs(0).Visible = True
        'fraConfig.Visible = True
        'chkConfigTabs(0).Value = 1
        
        chkConfigTabs_Click 0
        
    Else                            'Back
        
        lblInfo(0).Visible = False 'msg of main menu
        lblInfo(1).Visible = True 'msg of scan window
        picPaypal.Visible = True
        lstResults.Visible = True
        fraHostsMan.Visible = False
        fraUninstMan.Visible = False
        If chkConfigTabs(3).Value = 1 Then fraConfigTabs(3).Visible = True
        cmdConfig.Caption = Translate(18)
        cmdSaveDef.Enabled = True
        cmdScan.Enabled = True
        cmdFix.Enabled = True
        cmdInfo.Enabled = True
        fraConfig.Visible = False
        fraScan.Enabled = True
        
        AppendErrorLogCustom "SaveSettings initiated by clicking 'Back' button."
        
        SaveSettings
    End If

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfig_Click"
    If inIDE Then Stop: Resume Next
End Sub

'Private Sub cmdConfig_Tab(Tab_idx As Long)
'    On Error GoTo ErrorHandler:
'
'
'    Exit Sub
'End Sub

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
    DeleteBackup "", True
    lstBackups.Clear
End Sub

Private Sub cmdConfigBackupDelete_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    If lstBackups.ListIndex = -1 Then Exit Sub
    If lstBackups.SelCount = 0 Then
        'First you have to mark a checkbox next to at least one item!
        MsgBox Translate(554), vbInformation
        Exit Sub
    End If
    If lstBackups.SelCount = 1 Then
        If MsgBoxW(Translate(84), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    '    If msgboxW("Are you sure you want to delete this backup?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        If MsgBoxW(Replace$(Translate(85), "[]", lstBackups.SelCount), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If msgboxW("Are you sure you want to delete these " & lstBackups.SelCount & " backups?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    For i = lstBackups.ListCount - 1 To 0 Step -1
        If lstBackups.Selected(i) Then
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
    If lstBackups.SelCount = 0 Then
        'First you have to mark a checkbox next to at least one item!
        MsgBox Translate(554), vbInformation
        Exit Sub
    End If
    If lstBackups.SelCount = 1 Then
        'exclude question for ABR / SRP backups (it has inividual message)
        BackupSplitLine lstBackups.List(GetListBoxSelectedItemID(lstBackups)), , , , sDecription
        If sDecription <> ABR_BACKUP_TITLE _
          And Not StrBeginWith(sDecription, SRP_BACKUP_TITLE) Then
            If MsgBoxW(Translate(86), vbQuestion + vbYesNo) = vbNo Then Exit Sub
            'If msgboxW("Restore this item?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    Else
        If MsgBoxW(Replace$(Translate(87), "[]", lstBackups.SelCount), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If msgboxW("Restore these " & lstBackups.SelCount & " items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    'cache selected lines (to account for the shifting of elements in the list)
    ReDim Preserve aLines(lstBackups.ListCount - 1)
    j = 0
    For i = 0 To lstBackups.ListCount - 1   'vice versa order (list is already grouped vice versa)
        'only marked with checkbox
        If lstBackups.Selected(i) Then
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
    cmdScan.Caption = Translate(11) 'Scan
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfigBackupRestore_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Function GetListBoxSelectedItemID(lst As ListBox) As Long
    Dim i&
    For i = 0 To lst.ListCount - 1
        If lstBackups.Selected(i) Then
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
    IsOnIgnoreList "", UpdateList:=True
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfigIgnoreDelAll_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdConfigBackupCreateRegBackup_Click()
    'Run ABR to backup FULL registry
    If ABR_CreateBackup(True) Then
        'Full registry backup is successfully created.
        MsgBoxW Translate(1567), vbInformation
    End If
End Sub

Private Sub cmdConfigBackupCreateSRP_Click()
    'Create System Restore Point
    Dim nSeqNum As Long
    nSeqNum = SRP_Create_API()
    If nSeqNum <> 0 And bShowSRP Then
        frmMain.lstBackups.AddItem _
            BackupConcatLine(0&, 0&, BackupFormatDate(Now()), SRP_BACKUP_TITLE & " - " & nSeqNum & " - " & "Restore Point by HiJackThis"), 0
    End If
    'Note: that actual restore point record will appear in the WMI list after ~ 15 sec.
End Sub

Private Sub cmdConfigIgnoreDelSel_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    If lstIgnore.ListIndex = -1 Then Exit Sub
    If lstIgnore.SelCount = 0 Then
        'First you have to mark a checkbox next to at least one item!
        MsgBox Translate(554), vbInformation
        Exit Sub
    End If
    For i = 0 To lstIgnore.ListCount - 1
        RegDelHJT "Ignore" & CStr(i + 1)
    Next i
    For i = lstIgnore.ListCount - 1 To 0 Step -1
        If lstIgnore.Selected(i) Then lstIgnore.RemoveItem i
    Next i
    RegSaveHJT "IgnoreNum", lstIgnore.ListCount
    For i = 0 To lstIgnore.ListCount - 1
        RegSaveHJT "Ignore" & CStr(i + 1), Crypt(lstIgnore.List(i))
    Next i
    IsOnIgnoreList "", UpdateList:=True
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
    
    dNow = Now()
    dMidNight = GetDateAtMidnight(dNow)
    
    dLastFix = CDate(RegReadHJT("DateLastFix", "0"))
    lNumFixes = CLng(RegReadHJT("FixesToday", "0"))
    
    If lNumFixes = 0 Then
        RegSaveHJT "FixesToday", CStr(1)
    ElseIf dLastFix < dMidNight Then
        RegSaveHJT "FixesToday", CStr(1)
    Else
        RegSaveHJT "FixesToday", CStr(lNumFixes + 1)
    End If
    
    RegSaveHJT "DateLastFix", CStr(dNow)
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "IncreaseNumberOfFixes"
    If inIDE Then Stop: Resume Next
End Sub

'// Fix Checked
Private Sub cmdFix_Click()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.cmdFix_Click - Begin"
    
    Dim i&, j&, sPrefix$, pos&, sItem$
    Dim bFlushDNS As Boolean
    Dim bO24Fixed As Boolean
    Dim bO14Fixed As Boolean
    Dim bRestartExplorer As Boolean
    
    Dim Result As SCAN_RESULT
    
    If lstResults.ListCount = 0 Then Exit Sub
    
    If lstResults.SelCount = 0 Then
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
    
    If (lstResults.ListCount = lstResults.SelCount) And (InStr(1, Command(), "/StartupScan", 1) = 0) Then
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
    
    If bConfirm Then
        'lstResults.ListIndex = -1
        If MsgBoxW(Replace$(Translate(346), "[]", lstResults.SelCount) & _
           IIf(bMakeBackup, ".", ", " & Translate(347)), vbQuestion + vbYesNo, g_AppName) = vbNo Then Exit Sub
'        If msgboxW("Fix " & lstResults.SelCount & _
'         " selected items? This will permanently " & _
'         "delete and/or repair what you selected" & _
'         IIf(bMakeBackup, ".", ", unless you enable backups."), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    IncreaseNumberOfFixes 'save number of fixes for today in registry
    
    IncreaseFixID 'to track same items
    
    SetProgressBar lstResults.SelCount + 1
    UpdateProgressBar "Backup"
    
    If bMakeBackup Then
        'Creating FULL registry backup
        ABR_CreateBackup False
    End If
    
    'shpBackground.Tag = lstResults.SelCount
    'shpProgress.Tag = "0"
    
    'shpProgress.Width = 15
    'picPaypal.Visible = False
    bRebootRequired = False
    bUpdatePolicyNeeded = False
    bShownBHOWarning = False
    bShownToolbarWarning = False
    bSeenHostsFileAccessDeniedWarning = False
    
    Call GetProcesses(gProcess)
    
    For j = 0 To 1
      '0 - do backup only
      '1 - do fix
      
      If j = 1 Then BackupFlush
    
      For i = 0 To lstResults.ListCount - 1
        If lstResults.Selected(i) = True Then
            lstResults.ListIndex = i
            
            sPrefix = ""
            sItem = lstResults.List(i)
            pos = InStr(sItem, "-")
            If pos <> 0 Then
                sPrefix = Trim$(Left$(sItem, pos - 1))
            End If
            
            If j = 0 Then
                AppendErrorLogCustom "Backup: " & sItem
            Else
                AppendErrorLogCustom "Fixing: " & sItem
            End If
            
            If GetScanResults(sItem, Result) Then 'map ANSI string to Unicode
            
              If j = 0 Then
                MakeBackup Result
              Else
                UpdateProgressBar sPrefix
            
                Select Case sPrefix ' RTrim$(Left$(lstResults.List(i), 3))
                Case "R0", "R1", "R2": FixRegItem sItem, Result
                Case "R3":             FixR3Item sItem, Result
                Case "R4":             FixR4Item sItem, Result
                Case "F0", "F1":       FixFileItem sItem, Result
                Case "F2", "F3":       FixFileItem sItem, Result
                'Case "N1", "N2", "N3", "N4": FixNetscapeMozilla sItem,Result
                Case "O1":             FixO1Item sItem, Result: bFlushDNS = True
                Case "O2":             FixO2Item sItem, Result
                Case "O3":             FixO3Item sItem, Result
                Case "O4":             FixO4Item sItem, Result
                Case "O5":             FixO5Item sItem, Result
                Case "O6":             FixO6Item sItem, Result
                Case "O7":             FixO7Item sItem, Result
                Case "O8":             FixO8Item sItem, Result
                Case "O9":             FixO9Item sItem, Result
                Case "O10":            FixLSP
                Case "O11":            FixO11Item sItem, Result
                Case "O12":            FixO12Item sItem, Result
                Case "O13":            FixO13Item sItem, Result
                Case "O14":            If Not bO14Fixed Then FixO14Item sItem, Result: bO14Fixed = True 'O14 fix uses only once
                Case "O15":            FixO15Item sItem, Result
                Case "O16":            FixO16Item sItem, Result
                Case "O17":            FixO17Item sItem, Result: bFlushDNS = True
                Case "O18":            FixO18Item sItem, Result
                Case "O19":            FixO19Item sItem, Result
                Case "O20":            FixO20Item sItem, Result
                Case "O21":            FixO21Item sItem, Result: bRestartExplorer = True
                Case "O22":            FixO22Item sItem, Result
                Case "O23":            FixO23Item sItem, Result: If StrBeginWith(sItem, "O23 - Driver") Then bRebootRequired = True
                Case "O24":            FixO24Item sItem, Result: bO24Fixed = True
                Case "O25":            FixO25Item sItem, Result
                Case "O26":            FixO26Item sItem, Result
                Case Else
                   ' msgboxW "Fixing of " & Rtrim$(left$(lstResults.List(i), 3)) & _
                           " is not implemented yet. Bug me about it at " & _
                           "www.merijn.org/contact.html, because I should have done it.", _
                           vbInformation, "bad coder - no donuts"
                           
                    'Fixing of [] is not implemented yet."
                    MsgBoxW Replace$(Translate(348), "[]", sPrefix, vbInformation)
                End Select
              End If
            End If
            
        End If
      Next
    Next
    
    BackupFlush
    
    If bRestartExplorer Then RestartExplorer
    If bFlushDNS Then FlushDNS
    If bUpdatePolicyNeeded Then UpdatePolicy
    If bO24Fixed Then FixO24Item_Post ' restart shell
    
    UpdateProgressBar "Finish"
    lstResults.Clear
    bScanExecuted = False
    cmdFix.Enabled = False
    cmdFix.FontBold = False
    cmdScan.Caption = Translate(11)
    'cmdScan.Caption = "Scan"
    cmdScan.FontBold = True
    'lblInfo(0).Visible = True
    'lblInfo(1).Visible = False
    'picPaypal.Visible = True
    
    If cmdScan.Visible Then
        cmdScan.Enabled = True
        cmdScan.SetFocus
    End If
    
    'if somewhere explorer.exe has been killed, but not launched
    If Not ProcessExist("explorer.exe", True) Then
        RestartExplorer
    End If
    
    If bRebootRequired Then RestartSystem: bRebootRequired = False
    
    'CloseProgressbar 'leave progressBar visible to ensure the user saw completion of cure
    
    If Not inIDE Then MessageBeep MB_ICONINFORMATION
    
    AppendErrorLogCustom "frmMain.cmdFix_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdFix_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdHelp_Click()
    'Back
    If cmdConfig.Caption = Translate(19) Then
        cmdConfig_Click
    End If

    'If cmdHelp.Caption = "Help" Then
    If cmdHelp.Caption = Translate(16) Then
        lblInfo(0).Visible = False
        picPaypal.Visible = False
        lstResults.Visible = False
        fraHelp.Visible = True
        'cmdHelp.Caption = "Back"
        cmdHelp.Caption = Translate(17)
        'cmdConfig.Enabled = False
        cmdSaveDef.Enabled = False
        cmdScan.Enabled = False
        cmdFix.Enabled = False
        txtNothing.Visible = False
        
        fraHelp.Visible = True
        fraHelp.ZOrder 0
        chkHelp_Click 0 'help on section
    Else
        'lblInfo(0).Visible = True
        lblInfo(0).Visible = False
        lblInfo(1).Visible = True
        
        picPaypal.Visible = True
        lstResults.Visible = True
        fraHelp.Visible = False
        'cmdHelp.Caption = "Info..."
        cmdHelp.Caption = Translate(16)
        'cmdConfig.Enabled = True
        cmdSaveDef.Enabled = True
        cmdScan.Enabled = True
        cmdFix.Enabled = True
    End If
End Sub

Private Sub cmdInfo_Click()
    If lstResults.Visible Then
        If lstResults.SelCount = 0 And lstResults.ListIndex = -1 Then
            'First you have to mark a checkbox next to at least one item!
            MsgBox Translate(554), vbInformation
            Exit Sub
        End If
        GetInfo lstResults.List(lstResults.ListIndex)
    ElseIf txtHelp.Visible Then
        GetInfo LTrim$(txtHelp.SelText)
    End If
End Sub

Private Sub cmdSaveDef_Click()
    On Error GoTo ErrorHandler:
    If lstResults.SelCount = 0 Then
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
    RegSaveHJT "IgnoreNum", CStr(i + lstResults.SelCount)
    j = i + 1
    For i = 0 To lstResults.ListCount - 1
        If lstResults.Selected(i) Then
            RegSaveHJT "Ignore" & CStr(j), Crypt(lstResults.List(i))
            j = j + 1
        End If
    Next i
    IsOnIgnoreList "", UpdateList:=True
    
    For i = lstResults.ListCount - 1 To 0 Step -1
        If lstResults.Selected(i) Then lstResults.RemoveItem i
    Next i
    If lstResults.ListCount = 0 Then
        txtNothing.Visible = True
        cmdFix.FontBold = False
        'cmdScan.Caption = "Scan"
        cmdScan.Caption = Translate(11)
        cmdScan.FontBold = True
        If cmdScan.Visible Then
            If cmdScan.Enabled Then
                cmdScan.SetFocus
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdSaveDef_Click"
    If inIDE Then Stop: Resume Next
End Sub


Private Sub cmdScan_Click()
    On Error GoTo ErrorHandler:
    Dim Idx&, i&
    
    AppendErrorLogCustom "frmMain.cmdScan_Click - Begin"
    
    If bAutoLogSilent Then
        LockInterface bAllowInfoButtons:=False, bDoUnlock:=False
'    Else
'        LockInterface bAllowInfoButtons:=True, bDoUnlock:=False
    End If
    
    If isRanHJT_Scan Then
        Exit Sub
    Else
        isRanHJT_Scan = True
    End If
    cmdScan.Enabled = False
    
    Idx = 0
    
    'If cmdScan.Caption = "Scan" Then
    
    If cmdScan.Caption = Translate(11) Then
    
        bScanExecuted = True
        
        If bAutoLogSilent Then
            Call SystemPriorityDowngrade(True)
        End If
        
        'first scan after rebooting ?
        bFirstRebootScan = ScanAfterReboot()
    
        ' Erase main W array of scan results
        ReInitScanResults
        
        Idx = 1
        
        'labels off -> moved to SetProgressBar
        'lblInfo(0).Visible = False
        'lblInfo(1).Visible = False
        'shpBackground.Visible = True
        'shpProgress.Visible = True
        
        'picPaypal.Visible = False
        
        lblMD5.Visible = True
        'If bMD5 = False Then lblStatus.Visible = True
        
        cmdAnalyze.Enabled = False
    
        Idx = 2
    
        ' Clear Error Log
        ErrReport = ""
        
        CheckIntegrityHJT
    
        ' *******************************************************************

        StartScan '<<<<<<<-------- Main scan routine
        
        If txtNothing.Visible Or Not bAutoLog Then UpdateProgressBar "Finish"
        
        Idx = 3
        
        SortSectionsOfResultList
        
        Idx = 4
        
        'add the horizontal scrollbar if needed
        AddHorizontalScrollBarToResults lstResults
        
        If frmMain.lstResults.ListCount > 0 And Not bAutoLogSilent Then
            If bAutoSelect Then
                For i = 0 To frmMain.lstResults.ListCount - 1
                    frmMain.lstResults.Selected(i) = True
                Next i
            End If
        End If
        
        If Not bAutoLog Then
            If frmMain.WindowState <> vbMinimized Then
                SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                SetForegroundWindow frmMain.hwnd
                SetActiveWindow frmMain.hwnd
                SetFocus2 frmMain.hwnd
            End If
        End If
        
        Idx = 5
        
        'if no mark "No suspicious items found!"
        If txtNothing.Visible = False Then
        
            'cmdScan.Caption = "Save log"
            cmdScan.Caption = Translate(12)
            cmdScan.FontBold = False
            If Not bMD5 Then
                cmdFix.Enabled = True
                cmdFix.FontBold = True
            Else
                cmdFix.Enabled = False
            End If
'        Else
'            bAutoLog = False
        End If
        
        Idx = 6
        
        If bAutoLog Then
            If Not bAutoLogSilent Then DoEvents
            HJT_SaveReport '<<<<<< ------ Saving report
        End If
        
        cmdScan.Enabled = True
        cmdAnalyze.Enabled = True
        
        CloseProgressbar
        
        If Not bAutoLog Then
            If cmdFix.Visible And cmdFix.Enabled Then
                cmdFix.SetFocus
            End If
        End If
        
        bAutoLog = False
        
        If bAutoLogSilent Then
            Call SystemPriorityDowngrade(False)
        End If
        
    Else    'Caption = Save...

        LockInterface bAllowInfoButtons:=True, bDoUnlock:=True
        
        Call HJT_SaveReport
        
        UpdateProgressBar "Finish"

        'cmdScan.Caption = "Scan"
        'cmdScan.Caption = Translate(11)
        
        cmdScan.Enabled = True
    End If
    
    'focus on 1-st element of list
    Me.lstResults.ListIndex = -1
    'If Me.lstResults.Visible Then Me.lstResults.SetFocus
    isRanHJT_Scan = False
    
    LockInterface bAllowInfoButtons:=True, bDoUnlock:=True
    
    AppendErrorLogCustom "frmMain.cmdScan_Click - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdScan_Click", "(" & cmdScan.Caption & ")" & " (index = " & Idx & ")"
    LockInterface bAllowInfoButtons:=True, bDoUnlock:=True
    cmdScan.Enabled = True
    isRanHJT_Scan = False
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdStartupList_Click() 'Misc Tools -> StartupList scan
    Dim sPathComCtl As String, Success As Boolean
    sPathComCtl = BuildPath(AppPath(), "MSComCtl.ocx")
    If Not FileExists(sPathComCtl) Then
        If UnpackResource(102, sPathComCtl) Then Success = True
    Else
        Success = True
    End If
    If Success Then
        On Error Resume Next
        bSL_Abort = False
        bSL_Terminate = False
        frmStartupList2.Show
    Else
        MsgBoxW "Cannot unpack " & sPathComCtl, vbCritical
    End If
End Sub

Private Sub cmdUninstall_Click() 'Misc Tools -> Uninstall HiJackThis
    On Error GoTo ErrorHandler:
    Dim HJT_Install_Path As String
    Dim HJT_Location As String
    HJT_Location = BuildPath(PF_32, "HiJackThis Fork\HiJackThis.exe")
    
    If StrComp(AppPath(True), HJT_Location, 1) = 0 Then
        'This will completely remove HiJackThis, including settings and backups. Continue?
        If MsgBoxW(Translate(154), vbQuestion Or vbYesNo) = vbNo Then Exit Sub
    Else
'    If msgboxW("This will remove HiJackThis' settings from the Registry " & _
'              "and exit. Note that you will have to delete the " & _
'              "HiJackThis.exe file manually." & vbCrLf & vbCrLf & _
'              "Continue with uninstall?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        If MsgBoxW(Translate(153), vbQuestion Or vbYesNo) = vbNo Then Exit Sub
    End If
    
    Reg.DelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe"
    Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", False
    Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", True
    Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThisFork"
    If Not Reg.KeyHasSubKeys(HKEY_LOCAL_MACHINE, "Software\TrendMicro", False) Then
        Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro", False
    End If
    If Not Reg.KeyHasSubKeys(HKEY_LOCAL_MACHINE, "Software\TrendMicro", True) Then
        Reg.DelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro", True
    End If
    Reg.DelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan", False
    Reg.DelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan", True
    CreateUninstallKey False
    DeleteBackup "", True
    ABR_RemoveBackupALL True
    SubClassScroll False
    RemoveHJTShortcuts
    
    RemoveAutorunHJT
    
    SetCurrentDirectory StrPtr(SysDisk)
    HJT_Install_Path = BuildPath(PF_32, "HiJackThis Fork")
    
    If FolderExists(HJT_Install_Path) Then
        If StrComp(AppPath(True), HJT_Install_Path & "\HiJackThis.exe", 1) = 0 Then
        'delayed removing of HJT installation folder via cmd.exe, if it is launched from there
          Proc.ProcessRun _
            Environ("ComSpec"), _
            "/v /c (cd\& for /L %+ in (1,1,10) do ((timeout /t 1|| ping 127.1 -n 2)& rd /s /q """ & HJT_Install_Path & """&& exit))", _
            SysDisk, vbHide, True
        Else
            DeleteFolderForce HJT_Install_Path
            RemoveDirectory StrPtr(HJT_Install_Path)
        End If
    End If
    
    Close
    g_UninstallState = True
    Unload Me
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdUninstall_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub Form_Resize()
    'On Error Resume Next
    
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
    '(main)
    txtDefSearchAss.Width = Me.ScaleWidth - 2640
    txtDefSearchCust.Width = Me.ScaleWidth - 2640
    txtDefSearchPage.Width = Me.ScaleWidth - 2640
    txtDefStartPage.Width = Me.ScaleWidth - 2640
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
    fraUninstMan.Width = Me.ScaleWidth - 480
    lstUninstMan.Width = Me.ScaleWidth - 4995
    lblInfo(8).Left = Me.ScaleWidth - 4770
    lblInfo(10).Left = Me.ScaleWidth - 4770
    txtUninstManName.Left = Me.ScaleWidth - 4750 '3210
    txtUninstManCmd.Left = Me.ScaleWidth - 4750 '3210
    cmdUninstManUninstall.Left = Me.ScaleWidth - 4770
    cmdUninstManDelete.Left = Me.ScaleWidth - 4770
    cmdUninstManEdit.Left = Me.ScaleWidth - 2610 - 60
    cmdUninstManOpen.Left = Me.ScaleWidth - 4770
    cmdUninstManSave.Left = Me.ScaleWidth - 3450
    cmdUninstManBack.Left = Me.ScaleWidth - 1770 - 60
    cmdUninstManRefresh.Left = Me.ScaleWidth - 4770
    
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
    lstBackups.Height = Me.ScaleHeight - 3800
    chkShowSRP.Top = lstBackups.Top + lstBackups.Height + 50 '+ chkShowSRP.Height
    '(misc)
    
    fraHostsMan.Height = Me.ScaleHeight - 2805
    lstHostsMan.Height = Me.ScaleHeight - 4035 - 240
    lblConfigInfo(15).Top = Me.ScaleHeight - 3300 - 300
    cmdHostsManDel.Top = Me.ScaleHeight - 3300
    cmdHostsManToggle.Top = Me.ScaleHeight - 3300
    cmdHostsManOpen.Top = Me.ScaleHeight - 3300
    cmdHostsManBack.Top = Me.ScaleHeight - 3300
    vscMiscTools.Height = fraConfigTabs(3).Height
    fraN00b.Height = Me.ScaleHeight - 1175
    
    'If Me.ScaleHeight > 7250 Then
    If Me.ScaleHeight > 6500 Then
        fraUninstMan.Height = Me.ScaleHeight - 2725 '2805
        lstUninstMan.Height = Me.ScaleHeight - 2725 - 1100 '3855 - 60
        'cmdUninstManRefresh.Top = Me.ScaleHeight - 3315 - 60
        'cmdUninstManSave.Top = Me.ScaleHeight - 3315 - 60
        'cmdUninstManBack.Top = Me.ScaleHeight - 3315 - 60
    Else
        'fraUninstMan.Height = Me.ScaleHeight - 1850
        'lstUninstMan.Height = Me.ScaleHeight - 1850 - 1100
    End If
    
    If Me.ScaleHeight > 7200 Then
        cmdUninstManRefresh.Top = fraUninstMan.Top + fraUninstMan.Height - cmdUninstManRefresh.Height - 1000
        cmdUninstManSave.Top = fraUninstMan.Top + fraUninstMan.Height - cmdUninstManSave.Height - 1000
        cmdUninstManBack.Top = fraUninstMan.Top + fraUninstMan.Height - cmdUninstManBack.Height - 1000
    End If
    
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
    AddHorizontalScrollBarToResults lstResults
End Sub

Private Sub LoadSettings()
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "frmMain.LoadSettings - Begin"
    
    Dim bUseOldKey As Boolean, sCurLang$, WinHeight&, WinWidth&
    
    bUseOldKey = (Not Reg.KeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThisFork")) And _
        Reg.KeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis")
    
    chkAutoMark.Value = CInt(RegReadHJT("AutoSelect", "0", bUseOldKey))
    chkConfirm.Value = CInt(RegReadHJT("Confirm", "1", bUseOldKey))
    chkBackup.Value = CInt(RegReadHJT("MakeBackup", "1", bUseOldKey))
    'chkIgnoreSafeDomains.Value = CInt(RegReadHJT("IgnoreSafe", "1", bUseOldKey))
    chkLogProcesses.Value = CInt(RegReadHJT("LogProcesses", "1", bUseOldKey))
    chkAdditionalScan.Value = CInt(RegReadHJT("LogAdditional", "0", bUseOldKey))
    chkSkipIntroFrameSettings.Value = CInt(RegReadHJT("SkipIntroFrame", "0", bUseOldKey))
    chkSkipIntroFrame.Value = CInt(RegReadHJT("SkipIntroFrame", "0", bUseOldKey))
    chkSkipErrorMsg.Value = CInt(RegReadHJT("SkipErrorMsg", "0", bUseOldKey))
    chkConfigMinimizeToTray.Value = CInt(RegReadHJT("MinToTray", "0", bUseOldKey))
    chkIgnoreMicrosoft.Value = CInt(RegReadHJT("HideMicrosoft", "1", bUseOldKey))
    chkIgnoreAll.Value = CInt(RegReadHJT("IgnoreAllWhiteList", "0", bUseOldKey))
    chkDoMD5.Value = CInt(RegReadHJT("CalcMD5", "0", bUseOldKey))
    chkAdvLogEnvVars.Value = CInt(RegReadHJT("LogEnvVars", "0", bUseOldKey))
    chkCheckUpdatesOnStart.Value = CInt(RegReadHJT("CheckForUpdates", "0", bUseOldKey))
    chkShowSRP.Value = CInt(RegReadHJT("ShowSRP", "0", bUseOldKey))
    
    bHideMicrosoft = chkIgnoreMicrosoft.Value   'global
    bIgnoreAllWhitelists = chkIgnoreAll.Value   'global
    
    sCurLang = RegReadHJT("LanguageFile", "English")
    WinHeight = CLng(RegReadHJT("WinHeight", "6600"))
    WinWidth = CLng(RegReadHJT("WinWidth", "9000"))
    
    bLogEnvVars = (CLng(RegReadHJT("LogEnvVars", "0")) = 1) 'global
    bMD5 = (CLng(RegReadHJT("CalcMD5", "0")) = 1)           'global
    
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
    
    LastVerLaunched = RegReadHJT("Version", "", bUseOldKey)
    If ConvertVersionToNumber(LastVerLaunched) < ConvertVersionToNumber("2.6.1.21") Then isEncodedVer = True
    
    Dim CryptVer As Long, iIgnoreNum As Long, i As Long
    
    CryptVer = Val(RegReadHJT("CryptVer", "1"))
    
    If CryptVer < 2 Then
        RegSaveHJT "CryptVer", 2
    End If
    
    If CryptVer = 1 Then 'need to reEncode
        
        iIgnoreNum = Val(RegReadHJT("IgnoreNum", "0", True))
        
        If iIgnoreNum > 0 Then
            ReDim aIgnoreList(iIgnoreNum) As String
            
            'saving in binary format (no Base64 need)
            For i = 1 To iIgnoreNum
                aIgnoreList(i) = CryptV1(RegReadHJT("Ignore" & i, vbNullString, True), doCrypt:=False)
            Next
            For i = 1 To iIgnoreNum
                RegSaveHJT "Ignore" & CStr(i), Crypt(aIgnoreList(i))
            Next i
        End If
    End If
    
    sData = RegReadHJT("DefStartPage", "", bUseOldKey)
    'StrBeginWith(sData, "http") - необходим для обратной совместимости со старыми версиями HJT, чтобы не было конфликта криптографического модуля
    If sData = "" Or StrBeginWith(sData, "http") Or isEncodedVer Then
        g_DEFSTARTPAGE = "http://www.msn.com/"
    Else
        If CryptVer = 1 Then
            g_DEFSTARTPAGE = CryptV1(sData, doCrypt:=False)
        ElseIf CryptVer = 2 Then
            g_DEFSTARTPAGE = DeCrypt(Decode64(sData))
        Else
            g_DEFSTARTPAGE = "http://www.msn.com/"
        End If
    End If
    txtDefStartPage.Text = g_DEFSTARTPAGE

    sData = RegReadHJT("DefSearchPage", "", bUseOldKey)
    If sData = "" Or StrBeginWith(sData, "http") Or isEncodedVer Then
        g_DEFSEARCHPAGE = "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch"
    Else
        If CryptVer = 1 Then
            g_DEFSEARCHPAGE = CryptV1(sData, doCrypt:=False)
        ElseIf CryptVer = 2 Then
            g_DEFSEARCHPAGE = DeCrypt(Decode64(sData))
        Else
            g_DEFSEARCHPAGE = "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch"
        End If
    End If
    txtDefSearchPage.Text = g_DEFSEARCHPAGE
    
    sData = RegReadHJT("DefSearchAss", "", bUseOldKey)
    If sData = "" Or StrBeginWith(sData, "http") Or isEncodedVer Then
        g_DEFSEARCHASS = "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm"
    Else
        If CryptVer = 1 Then
            g_DEFSEARCHASS = CryptV1(sData, doCrypt:=False)
        ElseIf CryptVer = 2 Then
            g_DEFSEARCHASS = DeCrypt(Decode64(sData))
        Else
            g_DEFSEARCHASS = "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm"
        End If
    End If
    txtDefSearchAss.Text = g_DEFSEARCHASS
    
    sData = RegReadHJT("DefSearchCust", "", bUseOldKey)
    If sData = "" Or StrBeginWith(sData, "http") Or isEncodedVer Then
        g_DEFSEARCHCUST = "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm"
    Else
        If CryptVer = 1 Then
            g_DEFSEARCHCUST = CryptV1(sData, doCrypt:=False)
        ElseIf CryptVer = 2 Then
            g_DEFSEARCHCUST = DeCrypt(Decode64(sData))
        Else
            g_DEFSEARCHCUST = "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm"
        End If
    End If
    txtDefSearchCust.Text = g_DEFSEARCHCUST
    
    bAutoSelect = IIf(chkAutoMark.Value = 1, True, False)
    bConfirm = IIf(chkConfirm.Value = 1, True, False)
    bMakeBackup = IIf(chkBackup.Value = 1, True, False)
'    bIgnoreSafeDomains = IIf(chkIgnoreSafeDomains.Value = 1, True, False)
'    If bIgnoreAllWhitelists Then
'        bIgnoreSafeDomains = False
'    End If
    bLogProcesses = IIf(chkLogProcesses.Value = 1, True, False)
    bAdditional = IIf(chkAdditionalScan.Value = 1, True, False)
    bSkipErrorMsg = IIf(chkSkipErrorMsg.Value = 1, True, False)
    bMinToTray = IIf(chkConfigMinimizeToTray.Value = 1, True, False)
    bCheckForUpdates = IIf(chkCheckUpdatesOnStart.Value = 1, True, False)
    
    UpdateIE_RegVals
    
    For i = 0 To UBound(sFileVals)
        If sFileVals(i) = vbNullString Then Exit For
        'sFileVals(i) = replace$(sFileVals(i), "$WINDIR", sWinDir)
        sFileVals(i) = EnvironW(sFileVals(i))
    Next i
    
    ' move registry settings from old key to new
    If bUseOldKey Then
        SaveSettings
        RegSaveHJT "LanguageFile", sCurLang
        RegSaveHJT "WinHeight", CStr(WinHeight)
        RegSaveHJT "WinWidth", CStr(WinWidth)
    End If
    
    IsOnIgnoreList "", UpdateList:=True
    
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
    cboN00bLanguage.AddItem "Russian"
    cboN00bLanguage.AddItem "Ukrainian"
    
    sFile = DirW$(BuildPath(AppPath(), "*.lng"), vbFile)
    
    Do While Len(sFile)
        If sFile <> "_Lang_EN.lng" And _
            sFile <> "_Lang_RU.lng" And _
            sFile <> "_Lang_UA.lng" Then
                cboN00bLanguage.AddItem sFile
        End If
        sFile = DirW$()
    Loop
    
    sCurLang = RegReadHJT("LanguageFile", "English")  'HJT settings
    If bForceRU Then sCurLang = "Russian"
    If bForceUA Then sCurLang = "Ukrainian"
    If bForceEN Then sCurLang = "English"
    
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
    'https://msdn.microsoft.com/en-us/library/windows/desktop/dd318693(v=vs.85).aspx
    
    sFile = cboN00bLanguage.List(cboN00bLanguage.ListIndex)
    
    If InStr(1, Command$(), "/default", 1) > 0 Then sFile = "English"
    
    If Len(sFile) = 0 Then Exit Sub
    If sFile = "English" Then
        'LoadDefaultLanguage
        LoadLanguage &H409, bForceEN
        g_CurrentLang = sFile
    ElseIf sFile = "Russian" Then
        LoadLanguage &H419, bForceRU
        g_CurrentLang = sFile
    ElseIf sFile = "Ukrainian" Then
        LoadLanguage &H422, bForceUA
        g_CurrentLang = "Russian"
    Else
        LoadLangFile sFile
        ReloadLanguageNative
        ReloadLanguage
    End If
    
    ' Do not save force mode state!
    If Not (bForceRU Or bForceEN Or bForceUA) Then RegSaveHJT "LanguageFile", sFile
    
    AppendErrorLogCustom "frmMain.cboN00bLanguage_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_cboN00bLanguage_Click"
    If inIDE Then Stop: Resume Next
End Sub

'
' ====== Uninstall manager  ======
'

Sub UninstMan_Init()
    ReDim sKeyUninstall(1) As String
    sKeyUninstall(0) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    sKeyUninstall(1) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    If bIsWin64 Then
        ReDim Preserve sKeyUninstall(2) As String
        sKeyUninstall(2) = "HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"   '// TODO: Replace Wow6432Node with Reg. flag + add /reg:64
    End If
End Sub

Private Sub lstUninstMan_Click()
    Dim sName$, sUninst$, ItemID&

    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    
    UninstRefreshData ItemID, sName, sUninst 'refresh data
    txtUninstManName.Text = sName
    txtUninstManCmd.Text = sUninst
End Sub

Sub UninstRefreshData(IndexOfList As Long, sDisplayName$, sUninstString$)
    On Error GoTo ErrorHandler:

    Dim ID&
    ID = lstUninstMan.ItemData(IndexOfList)
    With UninstData(ID)
        sDisplayName = Reg.GetString(0&, UninstData(ID).AppRegKey, "DisplayName")
        sUninstString = Reg.GetString(0&, UninstData(ID).AppRegKey, "UninstallString")
        .DisplayName = sDisplayName
        .UninstString = sUninstString
    End With
    ' delete item if no data in registry
    If Len(sDisplayName) = 0 And Len(sUninstString) = 0 Then
        txtUninstManName.Text = vbNullString
        txtUninstManCmd.Text = vbNullString
        lstUninstMan.RemoveItem (IndexOfList)
        If lstUninstMan.ListCount <> 0 Then lstUninstMan.ListIndex = IIf(IndexOfList = -1, 0, IndexOfList)
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.UninstRefreshData"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdUninstManBack_Click()
    fraUninstMan.Visible = False
    fraConfigTabs(3).Visible = True
    SubClassScroll True
End Sub

Private Sub cmdUninstManDelete_Click()
    On Error GoTo ErrorHandler:

    Dim sName$, sUninst$, ItemID&, ID&
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    UninstRefreshData ItemID, sName, sUninst 'refresh data
    
    If Len(sUninst) <> 0 Then
        If MsgBoxW(Translate(220) & vbCrLf & vbCrLf & sName, vbQuestion + vbYesNo) = vbYes Then
            If Reg.DelKey(0&, UninstData(ID).AppRegKey) Then
                txtUninstManName.Text = vbNullString
                txtUninstManCmd.Text = vbNullString
                lstUninstMan.RemoveItem (ItemID)
                If lstUninstMan.ListCount <> 0 Then lstUninstMan.ListIndex = IIf(ItemID = -1, 0, ItemID)
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManDelete_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdUninstManUninstall_Click()
    On Error GoTo ErrorHandler:

    Dim sName$, sUninst$, ItemID&, sApplication$, sArguments$
    
    If lstUninstMan.ListCount = 0 Then Exit Sub

    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    
    UninstRefreshData ItemID, sName, sUninst 'refresh data
    
    If Len(sUninst) <> 0 Then
        sApplication = FindOnPath(sUninst)
        
        If FileExists(sApplication) Then
            sArguments = ExtractArguments(sUninst)
            ShellExecute 0&, 0&, StrPtr(sApplication), StrPtr(sArguments), 0&, 1&
        End If
    End If
    
    'cmdUninstManRefresh_Click
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManUninstall_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdUninstManEdit_Click()
    On Error GoTo ErrorHandler:

    Dim S$, sName$, sUninst$, ItemID&, ID&
    
    If lstUninstMan.ListCount = 0 Then Exit Sub

    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    UninstRefreshData ItemID, sName, sUninst 'refresh data
    
    If Len(sName) = 0 And Len(sUninst) = 0 Then Exit Sub
    
    'Edit uninstall command
    S = InputBox(Translate(221) & ": '" & sName & ":", Translate(215), IIf(Len(sUninst) > 255, vbNullString, sUninst)) 'InputBox cannot hold more than 255 chars
    's = InputBox("Enter the new uninstall command for this program, '" & txtUninstManName.Text & ":", "Edit uninstall command", txtUninstManCmd.Text)
    
    If StrPtr(S) <> 0 And S <> sUninst And Len(S) <> 0 Then
        If Reg.SetStringVal(0&, UninstData(ID).AppRegKey, "UninstallString", S) Then
            MsgBoxW Translate(222), vbInformation
            'msgboxW "New uninstall string saved!", vbInformation
            txtUninstManCmd.Text = S
            UninstData(ID).UninstString = S
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManEdit_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdUninstManOpen_Click()
    ShellExecute 0&, StrPtr("open"), StrPtr("control.exe"), StrPtr("appwiz.cpl"), 0&, 1
End Sub

Private Sub cmdUninstManRefresh_Click()
    On Error GoTo ErrorHandler:

    Dim sItems$(), sName$, sUninst$, i&, j&, cnt&
    
    lstUninstMan.Clear
    Erase UninstData
    cnt = -1
    
    'lstUninstMan.Sorted must be False ' Do not enable this kind of sorting at all!!! Otherwise, virus will eat your computer :)
    
    For j = 0 To UBound(sKeyUninstall)
        sItems = Split(Reg.EnumSubKeys(0&, sKeyUninstall(j)), "|")
        If UBound(sItems) <> -1 Then
            For i = 0 To UBound(sItems)
                sName = Reg.GetString(0&, sKeyUninstall(j) & "\" & sItems(i), "DisplayName")
                sUninst = Reg.GetString(0&, sKeyUninstall(j) & "\" & sItems(i), "UninstallString")
                
                If Len(sName) <> 0 And Len(sUninst) <> 0 Then
                    cnt = cnt + 1
                    ReDim Preserve UninstData(cnt)
                    With UninstData(cnt)                                        'array
                        .DisplayName = sName
                        .UninstString = sUninst
                        .AppRegKey = sKeyUninstall(j) & "\" & sItems(i)
                        .KeyTime = ConvertDateToUSFormat(Reg.GetKeyTime(0&, .AppRegKey))
                    End With
                End If
            Next i
        End If
    Next j
    If cnt = -1 Then Exit Sub
    
    'Sorting user type array using bufer array of positions (c) Dragokas
    Dim pos() As Long, names() As String: ReDim pos(cnt), names(cnt)
    For i = 0 To cnt: pos(i) = i: names(i) = UninstData(i).DisplayName: Next 'key of sort is DisplayName
    QuickSortSpecial names, pos, 0, cnt
    
    For i = 0 To cnt
        lstUninstMan.AddItem UninstData(pos(i)).DisplayName
        lstUninstMan.ItemData(i) = pos(i)     'array marker
    Next
    
    If lstUninstMan.ListCount Then lstUninstMan.ListIndex = 0
    If lstUninstMan.Visible And lstUninstMan.Enabled Then
        lstUninstMan.SetFocus
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManRefresh_Click"
    If inIDE Then Stop: Resume Next
End Sub

Function ConvertDateToUSFormat(d As Date) As String 'DD.MM.YYYY HH:MM:SS -> YYYY/MM/DD HH:MM:SS (for sorting purposes)
    ConvertDateToUSFormat = _
    Right$("000" & Year(d), 4) & "/" & _
    Right$("0" & Month(d), 2) & "/" & _
    Right$("0" & Hour(d), 2) & ":" & _
    Right$("0" & Day(d), 2) & " " & _
    Right$("0" & Minute(d), 2) & ":" & _
    Right$("0" & Second(d), 2)
End Function

Private Sub cmdUninstManSave_Click()
    On Error GoTo ErrorHandler:

    Dim sList$, i&, sFile$, ff%, ID&, b() As Byte, sTmpFile$, buf As String
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    'sFile = CmnDlgSaveFile("Save Add/Remove Software list to disk...", "Text files (*.txt)|*.txt|All files (*.*)|*.*", "uninstall_list.txt")
    sFile = CmnDlgSaveFile(Translate(225), Translate(226) & " (*.txt)|*.txt|" & Translate(1003) & " (*.*)|*.*", "uninstall_list.txt")
    
    If Len(sFile) = 0 Then Exit Sub
    
    sList = ChrW$(-257)
    
    sList = sList & String$(55, "-") & vbCrLf
    sList = sList & Space$(20) & "Sort by Date" & vbCrLf
    sList = sList & String$(55, "-") & vbCrLf & vbCrLf
    
    ' Make positions array of sorting by .KeyTime property (registry key date).
    Dim cnt&: cnt = lstUninstMan.ListCount - 1
    Dim pos() As Long, names() As String: ReDim pos(cnt), names(cnt)
    For i = 0 To cnt: pos(i) = i: names(i) = UninstData(i).KeyTime: Next
    QuickSortSpecial names, pos, 0, cnt
    
    For i = cnt To 0 Step -1 'descending order
        With UninstData(pos(i))
            sList = sList & .KeyTime & vbTab & .DisplayName & vbCrLf
        End With
    Next
    
    sList = sList & vbCrLf & vbCrLf
    sList = sList & String$(55, "-") & vbCrLf
    sList = sList & Space$(20) & "Sort by Alphabet" & vbCrLf
    sList = sList & String$(55, "-") & vbCrLf & vbCrLf
    
    For i = 0 To lstUninstMan.ListCount - 1
        ID = lstUninstMan.ItemData(i)
        sList = sList & UninstData(ID).DisplayName & vbCrLf
    Next i
    
    sList = sList & vbCrLf & vbCrLf
    
    sList = sList & String$(55, "-") & vbCrLf
    sList = sList & Space$(20) & "Registry Snapshot" & vbCrLf
    sList = sList & String$(55, "-") & vbCrLf

    For i = 0 To UBound(sKeyUninstall)
        sTmpFile = BuildPath(AppPath(), "HJT_tmp_" & i & ".reg")
        
        If Proc.ProcessRun("reg.exe", "export """ & sKeyUninstall(i) & """ """ & sTmpFile & """ /y", , 0) Then
            If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 15000) Then     'if ExitCode = 0, 15 sec for timeout
                Proc.ProcessClose , , True
            End If
            ff = FreeFile()
            Open sTmpFile For Binary Access Read Shared As #ff
                buf = Space$(LOF(ff) - 2)   '-BOM
                Get #ff, 3, buf
            Close #ff
            sList = sList & vbCrLf & StrConv(buf, vbFromUnicode)
            DeleteFileWEx (StrPtr(sTmpFile))
        End If
    Next

    b = sList ' UTF-16
    
    If FileExists(sFile) Then DeleteFileWEx (StrPtr(sFile))
    
    ff = FreeFile()
    Open sFile For Binary Access Write As #ff
        Put #ff, , b()
    Close #ff

    ShellExecute 0&, StrPtr("open"), StrPtr("notepad.exe"), StrPtr(sFile), 0&, 1&
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManSave_Click"
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

Private Sub mnuToolsDelFileOnReboot_Click()     'Tools -> Delete File -> Delete on reboot...
    cmdDelOnReboot_Click
End Sub

Private Sub mnuToolsUnlockAndDelFile_Click()    'Tools -> Delete File -> Unlock & Delete file
    Dim sFilename$
    'Enter file name:, Unlock & Delete
    sFilename = InputBox(Translate(1952), Translate(1953))
    If StrPtr(sFilename) = 0 Then Exit Sub
    'sFileName = OpenFileDialog("Enter file to unlock access and delete...")
    'sFileName = CmnDlgOpenFile("Enter file to unlock access and delete...", Translate(1003) & " (*.*)|*.*|" & Translate(511) & " (*.dll)|*.dll|" & Translate(512) & " (*.exe)|*.exe")
    'sFileName = CmnDlgOpenFile("Enter file to unlock access and delete...", "All files (*.*)|*.*|DLL libraries (*.dll)|*.dll|Program files (*.exe)|*.exe")
    If 0 = Len(sFilename) Then Exit Sub
    sFilename = UnQuote(EnvironW(sFilename))
    If 0 = DeleteFileWEx(StrPtr(sFilename)) Then
        'Could not delete file. & vbcrlf & Possible, it is locked by another process.
        MsgBoxW Translate(1954)
    Else
        'File: [] deleted successfully.
        MsgBoxW Replace$(Translate(1955), "[]", sFilename)
    End If
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
    cmdN00bTools_Click
    cmdARSMan_Click
End Sub

Private Sub mnuToolsShortcutsChecker_Click()    'Tools -> Shortcuts -> Check Browsers' LNK
    'Download Check Browsers' LNK by Dragokas & regist
    'and ask to run
    DownloadUnzipAndRun "https://dragokas.com/tools/CheckBrowsersLNK.zip", "Check Browsers LNK.exe", False
End Sub
Private Sub mnuToolsShortcutsFixer_Click()      'Tools -> Shortcuts -> ClearLNK
    'Download ClearLNK by Dragokas
    'and ask to run
    DownloadUnzipAndRun "https://dragokas.com/tools/ClearLNK.zip", "ClearLNK.exe", False
End Sub

Private Sub mnuHelpManualEnglish_Click()
    Dim szQSUrl$: szQSUrl = "https://www.bleepingcomputer.com/tutorials/how-to-use-hijackthis/"
    ShellExecute Me.hwnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualRussian_Click()
    Dim szQSUrl$: szQSUrl = "https://safezone.cc/threads/25184/"
    ShellExecute Me.hwnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualFrench_Click()
    Dim szQSUrl$: szQSUrl = "https://www.bleepingcomputer.com/tutorials/comment-utiliser-hijackthis/"
    ShellExecute Me.hwnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualGerman_Click()
    Dim szQSUrl$: szQSUrl = "https://www.bleepingcomputer.com/tutorials/wie-hijackthis-genutzt-wird-um/"
    ShellExecute Me.hwnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualSpanish_Click()
    Dim szQSUrl$: szQSUrl = "https://www.bleepingcomputer.com/tutorials/como-usar-hijackthis/"
    ShellExecute Me.hwnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualPortuguese_Click()
    Dim szQSUrl$: szQSUrl = "https://www.linhadefensiva.org/2005/06/hijackthis-completo/"
    ShellExecute Me.hwnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualDutch_Click()
    Dim szQSUrl$: szQSUrl = "https://www.bleepingcomputer.com/tutorials/hoe-gebruik-je-hijackthis/"
    ShellExecute Me.hwnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub

Private Sub mnuHelpUpdate_Click()       'Help -> Download new version
    CheckForUpdate
End Sub

Private Sub mnuHelpAbout_Click()        'Help -> About HJT
    cmdN00bClose_Click
    ' Закрыть окно "Инструменты"
    If cmdConfig.Caption = Translate(19) Then cmdConfig_Click
    If cmdHelp.Caption = Translate(16) Then cmdHelp_Click
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
    bMD5 = (chkDoMD5.Value = 1)
    bCheckForUpdates = IIf(chkCheckUpdatesOnStart.Value = 1, True, False)
    
    RegSaveHJT "AutoSelect", CStr(Abs(CInt(bAutoSelect)))
    RegSaveHJT "Confirm", CStr(Abs(CInt(bConfirm)))
    RegSaveHJT "MakeBackup", CStr(Abs(CInt(bMakeBackup)))
    'RegSaveHJT "IgnoreSafe", CStr(Abs(CInt(bIgnoreSafeDomains)))
    RegSaveHJT "LogProcesses", CStr(Abs(CInt(bLogProcesses)))
    RegSaveHJT "LogAdditional", CStr(Abs(CInt(bAdditional)))
    RegSaveHJT "SkipIntroFrame", CStr(chkSkipIntroFrameSettings.Value)
    RegSaveHJT "SkipErrorMsg", CStr(Abs(CInt(bSkipErrorMsg)))
    RegSaveHJT "MinToTray", CStr(Abs(CInt(bMinToTray)))
    RegSaveHJT "DefStartPage", Encode64(Crypt(txtDefStartPage.Text))
    RegSaveHJT "DefSearchPage", Encode64(Crypt(txtDefSearchPage.Text))
    RegSaveHJT "DefSearchAss", Encode64(Crypt(txtDefSearchAss.Text))
    RegSaveHJT "DefSearchCust", Encode64(Crypt(txtDefSearchCust.Text))
    RegSaveHJT "LogEnvVars", Abs(CLng(bLogEnvVars))
    RegSaveHJT "CalcMD5", Abs(CLng(bMD5))
    RegSaveHJT "CheckForUpdates", CStr(Abs(CInt(bCheckForUpdates)))
    
    'Update global state
    UpdateIE_RegVals
    
    AppendErrorLogCustom "frmMain.SaveSettings - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SaveSettings"
    If inIDE Then Stop: Resume Next
End Sub

'Context menu in result list of scan:

Private Sub lstResults_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ErrorHandler:
    
    Const MAX_JUMP_LIST_ITEMS As Long = 10
    
    Dim Result As SCAN_RESULT
    Dim sItem As String
    Dim sPrefix As String
    Dim pos As Long
    Dim i As Long, j As Long
    Dim RegItems As Long
    Dim FileItems As Long
    Dim bExists As Boolean, bNoValue As Boolean
    Dim sIniFile As String, sFile As String
    Dim Idx As Long, XY As Long, XPix As Long, YPix As Long
    
    'select item by right click
    If Button = 2 Then
        XPix = X / Screen.TwipsPerPixelX
        YPix = Y / Screen.TwipsPerPixelY
        XY = YPix * 65536 + XPix
        Idx = SendMessage(lstResults.hwnd, LB_ITEMFROMPOINT, 0&, ByVal XY)
        If Idx >= 0 And Idx <= (lstResults.ListCount - 1) Then
            lstResults.ListIndex = Idx
        End If
    End If
    
    If Button = 2 And Not (isRanHJT_Scan And bAutoLogSilent) Then
        If lstResults.SelCount = 0 Then     'items not checked ?
            mnuResultFix.Enabled = False
            mnuResultAddToIgnore.Enabled = False
        Else
            mnuResultFix.Enabled = True
            mnuResultAddToIgnore.Enabled = True
        End If
        If lstResults.ListIndex = -1 Then   'item not selected ?
            mnuResultInfo.Enabled = False
            mnuResultSearch.Enabled = False
            mnuResultDelim1.Enabled = False
        Else
            mnuResultInfo.Enabled = True
            mnuResultSearch.Enabled = True
            mnuResultDelim1.Enabled = True
        End If
        If lstResults.ListCount = 0 Then    'no items
            mnuResultAddALLToIgnore.Enabled = False
        Else
            mnuResultAddALLToIgnore.Enabled = True
        End If
        
        'building the jump list
        mnuResultJump.Enabled = False
        
        sItem = GetSelected_OrCheckedItem()
        
        If sItem <> "" Then
            pos = InStr(sItem, "-")
            If pos <> 0 Then
                sPrefix = Trim$(Left$(sItem, pos - 1))
            End If
            If GetScanResults(sItem, Result) Then
                If AryPtr(Result.File) Or AryPtr(Result.Reg) Then
                    mnuResultJump.Enabled = True
                    
                    If (AryPtr(Result.File) And AryPtr(Result.Reg)) Then
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
                    
                    If AryPtr(Result.File) Then
                        For j = 0 To UBound(Result.File)
                            FileItems = FileItems + 1
                            If FileItems > MAX_JUMP_LIST_ITEMS Then Exit For
                            
                            bExists = FileExists(Result.File(j).Path)
                            mnuResultJumpFile(FileItems - 1).Caption = Result.File(j).Path & IIf(bExists, "", " (no file)")
                        Next
                    End If
                    
                    If AryPtr(Result.Reg) Then
                        For j = 0 To UBound(Result.Reg)
                            With Result.Reg(j)
                                If .IniFile <> "" Then
                                    FileItems = FileItems + 1
                                    If FileItems <= MAX_JUMP_LIST_ITEMS Then
                                        bExists = FileExists(.IniFile)
                                        mnuResultJumpFile(FileItems - 1).Caption = .IniFile & " => [" & .Key & "], " & .Param & IIf(bExists, "", " (no file)")
                                    End If
                                Else
                                    RegItems = RegItems + 1
                                    If RegItems <= MAX_JUMP_LIST_ITEMS Then
                                        bExists = Reg.KeyExists(.Hive, .Key, .Redirected)
                                        bNoValue = False
                                        If (.ActionType And BACKUP_KEY) Or (.ActionType And REMOVE_KEY) Then
                                        Else
                                            bNoValue = Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected)
                                        End If
                                        mnuResultJumpReg(RegItems - 1).Caption = _
                                          Reg.GetShortHiveName(Reg.GetHiveNameByHandle(.Hive)) & "\" & .Key & ", " & .Param & _
                                          IIf(.Redirected, " (x32)", "") & IIf(bExists, "", " (no key)") & IIf(bNoValue, " (no value)", "")
                                    End If
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
        End If
        PopupMenu mnuResultList
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.lstResults_MouseUp"
    If inIDE Then Stop: Resume Next
End Sub

'// TODO: Why is it not working ??? Who intercepts en event ?
'Private Sub lstResults_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then cmdFix_Click
'End Sub

Private Sub mnuResultJumpFile_Click(Index As Integer)   'Context => Jump to ... => File
    Dim sItem As String
    Dim sFile As String
    Dim sFolder As String
    Dim Result As SCAN_RESULT
    
    sItem = GetSelected_OrCheckedItem()
    
    If GetScanResults(sItem, Result) Then
        If AryPtr(Result.File) Then
            If UBound(Result.File) >= Index Then
                sFile = Result.File(Index).Path
                sFolder = GetParentDir(sFile)
                If FileExists(sFile) Then
                    OpenAndSelectFile sFile
                ElseIf FolderExists(sFolder) Then
                    OpenAndSelectFile sFolder
                End If
            End If
        ElseIf AryPtr(Result.Reg) Then
            If Result.Reg(0).IniFile <> "" Then
                sFile = Result.Reg(0).IniFile
                sFolder = GetParentDir(sFile)
                If FileExists(sFile) Then
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
    Dim Result As SCAN_RESULT
    
    sItem = GetSelected_OrCheckedItem()
    
    If GetScanResults(sItem, Result) Then
        If AryPtr(Result.Reg) Then
            If UBound(Result.Reg) >= Index Then
                With Result.Reg(Index)
                    Reg.Jump .Hive, .Key, .Param, .Redirected
                End With
            End If
        End If
    End If
End Sub

Private Function GetSelected_OrCheckedItem() As String
    Dim i As Long
    If lstResults.SelCount = 1 Then 'checkbox
        For i = 0 To lstResults.ListCount - 1
            If lstResults.Selected(i) = True Then
                GetSelected_OrCheckedItem = lstResults.List(i)
                Exit For
            End If
        Next
    ElseIf (lstResults.ListIndex <> -1) Then  'selection
        GetSelected_OrCheckedItem = lstResults.List(lstResults.ListIndex)
    End If
End Function

Private Sub mnuResultFix_Click()          'Fix checked
    cmdFix_Click
End Sub

Private Sub mnuResultInfo_Click()         'Info on selected
    cmdInfo_Click
End Sub

Private Sub mnuResultAddToIgnore_Click()  'Add to ignore list
    cmdSaveDef_Click
End Sub

Private Sub mnuResultAddALLToIgnore_Click()  'Add to ignore list
    Dim i As Long
    For i = 0 To lstResults.ListCount - 1
        lstResults.Selected(i) = True
    Next
    cmdSaveDef_Click
    If lstResults.ListCount > 0 Then
        For i = 0 To lstResults.ListCount - 1
            lstResults.Selected(i) = False
        Next
    End If
End Sub

Private Sub mnuResultSearch_Click()       'Search on Google
    Dim sItem$, sURL$, pos&
    sItem = lstResults.List(lstResults.ListIndex)
    pos = InStr(sItem, ":")
    If pos > 0 Then
        sItem = Mid$(sItem, pos + 1)
    End If
    OpenURL "https://www.google.com/?ie=UTF-8#q=" & URLEncode(sItem)
End Sub

Private Sub mnuResultReScan_Click()       'ReScan
    cmdScan.Caption = Translate(11)
    cmdScan_Click
End Sub

Private Sub mnuSaveReport_Click()         'Save report...
    Call HJT_SaveReport
End Sub

'test stuff - BUTTON: enum tasks to CSV
Private Sub cmdTaskScheduler_Click()
    Call EnumTasks2(True)
End Sub

Private Sub chkHelp_Click(Index As Integer)
    Static LastIdx As Long

    Dim i As Long, j As Long
    Dim sText As String
    Dim sSeparator$
    Dim aSect() As Variant
    
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
        aSect = Array("R0", "R1", "R2", "R3", "R4", "F0", "F1", "F2", "F3", "O1", "O2", "O3", "O4", "O5", "O6", "O7", "O8", "O9", "O10", _
            "O11", "O12", "O13", "O14", "O15", "O16", "O17", "O18", "O19", "O20", "O21", "O22", "O23", "O24", "O25", "O26")
        
        sText = Translate(31) & vbCrLf & vbCrLf & Translate(490)
        sSeparator = String$(100, "-")
        
        For i = 0 To UBound(aSect)
            Select Case aSect(i)
            Case ""
                Case "R0": j = 401
                Case "R1": j = 402
                Case "R2": j = 403
                Case "R3": j = 404
                Case "R4": j = 434
                Case "F0": j = 405
                Case "F1": j = 406
                Case "F2": j = 407
                Case "F3": j = 408
                Case "O1": j = 409
                Case "O2": j = 410
                Case "O3": j = 411
                Case "O4": j = 412
                Case "O5": j = 413
                Case "O6": j = 414
                Case "O7": j = 415
                Case "O8": j = 416
                Case "O9": j = 417
                Case "O10": j = 418
                Case "O11": j = 419
                Case "O12": j = 420
                Case "O13": j = 421
                Case "O14": j = 422
                Case "O15": j = 423
                Case "O16": j = 424
                Case "O17": j = 425
                Case "O18": j = 426
                Case "O19": j = 427
                Case "O20": j = 428
                Case "O21": j = 429
                Case "O22": j = 430
                Case "O23": j = 431
                Case "O24": j = 432
                Case "O25": j = 433
                Case "O26": j = 435
            End Select

            sText = sText & vbCrLf & sSeparator & vbCrLf & FindLine(aSect(i) & " -", Translate(31)) & vbCrLf & sSeparator & vbCrLf & Translate(j) & vbCrLf
        Next
        
        txtHelp.Text = sText
    
    Case 1: 'Keys
        txtHelp.Text = Translate(32)
    
    Case 2: 'Purpose
        txtHelp.Text = Translate(33) & TranslateNative(34)
    
    Case 3: 'History
        ' Version history:
        txtHelp.Text = g_VersionHistory
    End Select
    
    bSwitchingTabs = False
    LastIdx = Index
End Sub

Function FindLine(sPartialText As String, sFullText As String) As String
    Dim arr() As String, i&
    arr = Split(sFullText, vbCrLf)
    If IsArrDimmed(arr) Then
        For i = 0 To UBound(arr)
            If InStr(1, arr(i), sPartialText, 1) <> 0 Then FindLine = arr(i): Exit For
        Next
    End If
End Function

Private Sub cmdProcessManager_Click() 'Misc Tools -> Process Manager
    frmProcMan.Show
End Sub

'Scan area frame
Private Sub chkLogProcesses_Click()
    bLogProcesses = (chkLogProcesses.Value = 1)
    RegSaveHJT "LogProcesses", Abs(CLng(bLogProcesses))
End Sub

Private Sub chkAdvLogEnvVars_Click()
    bLogEnvVars = (chkAdvLogEnvVars.Value = 1)
    RegSaveHJT "LogEnvVars", Abs(CLng(bLogEnvVars))
End Sub

Private Sub chkAdditionalScan_Click()
    bAdditional = (chkAdditionalScan.Value = 1)
    RegSaveHJT "LogAdditional", Abs(CLng(bAdditional))
End Sub

'Backup & Fix frame
Private Sub chkBackup_Click()
    bMakeBackup = (chkBackup.Value = 1)
    RegSaveHJT "MakeBackup", Abs(CLng(bMakeBackup))
End Sub

Private Sub chkConfirm_Click()
    bConfirm = (chkConfirm.Value = 1)
    RegSaveHJT "Confirm", Abs(CLng(bConfirm))
End Sub

Private Sub chkAutoMark_Click()
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
Private Sub chkIgnoreMicrosoft_Click()
    bHideMicrosoft = chkIgnoreMicrosoft.Value
    If bHideMicrosoft Then
        If chkIgnoreAll.Value = 1 Then chkIgnoreAll.Value = 0
    End If
    RegSaveHJT "HideMicrosoft", Abs(CLng(bHideMicrosoft))
End Sub

Private Sub chkIgnoreAll_Click()
    bIgnoreAllWhitelists = chkIgnoreAll.Value
    If bIgnoreAllWhitelists Then
        If chkIgnoreMicrosoft.Value = 1 Then chkIgnoreMicrosoft.Value = 0
    End If
    RegSaveHJT "IgnoreAllWhiteList", Abs(CLng(bIgnoreAllWhitelists))
End Sub

Private Sub chkDoMD5_Click()
    bMD5 = (chkDoMD5.Value = 1)
    RegSaveHJT "CalcMD5", Abs(CLng(bMD5))
End Sub

Private Sub chkConfigStartupScan_Click()
    If gNotUserClick Then gNotUserClick = False: Exit Sub
    If chkConfigStartupScan.Value = 1 Then
        InstallAutorunHJT
    Else
        RemoveAutorunHJT
    End If
End Sub

'Interface frame
Private Sub chkSkipIntroFrame_Click()
    RegSaveHJT "SkipIntroFrame", CStr(chkSkipIntroFrame.Value)
    chkSkipIntroFrameSettings.Value = chkSkipIntroFrame.Value
End Sub

Private Sub chkSkipIntroFrameSettings_Click()
    'RegSaveHJT "SkipIntroFrame", CStr(chkSkipIntroFrame.Value)
End Sub

Private Sub chkSkipErrorMsg_Click()
    bSkipErrorMsg = (chkSkipErrorMsg.Value = 1)
    RegSaveHJT "SkipErrorMsg", Abs(CLng(bSkipErrorMsg))
End Sub

Private Sub chkConfigMinimizeToTray_Click()
    bMinToTray = (chkConfigMinimizeToTray.Value = 1)
    RegSaveHJT "MinToTray", Abs(CLng(bMinToTray))
End Sub
