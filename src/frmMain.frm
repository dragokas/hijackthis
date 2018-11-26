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
      Caption         =   "   Other stuff"
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
         Caption         =   "Settings"
         Height          =   450
         Left            =   1440
         TabIndex        =   5
         Tag             =   "0"
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   450
         Left            =   240
         TabIndex        =   4
         Tag             =   "0"
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraSubmit 
      Height          =   1455
      Left            =   3000
      TabIndex        =   55
      Top             =   5880
      Width           =   2885
      Begin VB.CommandButton cmdAnalyze 
         Caption         =   "Analyze report"
         Enabled         =   0   'False
         Height          =   450
         Left            =   480
         TabIndex        =   56
         Top             =   300
         Width           =   1935
      End
      Begin VB.CommandButton cmdMainMenu 
         Caption         =   "Main Menu"
         Height          =   450
         Left            =   720
         TabIndex        =   58
         Top             =   850
         Width           =   1455
      End
   End
   Begin VB.Frame fraScan 
      Caption         =   "   Scan && fix stuff"
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
         TabIndex        =   87
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
         Tag             =   "1"
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
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   -20
      Width           =   1335
   End
   Begin VB.CommandButton cmdHidden 
      Default         =   -1  'True
      Height          =   195
      Left            =   24960
      TabIndex        =   86
      Top             =   14760
      Width           =   75
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
         Height          =   555
         Index           =   3
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
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
         Height          =   555
         Index           =   2
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
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
         Height          =   555
         Index           =   1
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1455
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
         Height          =   555
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Height          =   9120
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   -4080
         Visible         =   0   'False
         Width           =   8055
         Begin VB.Frame fraMiscToolsScroll 
            BorderStyle     =   0  'None
            Height          =   12015
            Left            =   0
            TabIndex        =   54
            Top             =   2000
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
               TabIndex        =   139
               Top             =   10200
               Width           =   7335
               Begin VB.CommandButton cmdUninstall 
                  Caption         =   "Uninstall HiJackThis"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   140
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
                  TabIndex        =   73
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
               TabIndex        =   132
               Top             =   6240
               Width           =   7335
               Begin VB.CommandButton cmdLnkCleaner 
                  Caption         =   "ClearLNK"
                  Height          =   480
                  Left            =   120
                  TabIndex        =   134
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.CommandButton cmdLnkChecker 
                  Caption         =   "Check Browsers' LNK"
                  Height          =   480
                  Left            =   120
                  TabIndex        =   133
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.Label lblLnkCleanerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Clean and restore list of infected shortcuts (.LNK), found via Check Browsers' LNK plugin"
                  Height          =   615
                  Left            =   2520
                  TabIndex        =   136
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
                  TabIndex        =   135
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
               TabIndex        =   115
               Top             =   1440
               Width           =   7335
               Begin VB.CommandButton cmdDigiSigChecker 
                  Caption         =   "Digital signature checker"
                  Height          =   480
                  Left            =   120
                  TabIndex        =   130
                  Top             =   4080
                  Width           =   2295
               End
               Begin VB.CommandButton cmdRegKeyUnlocker 
                  Caption         =   "Registry Key Unlocker"
                  Height          =   480
                  Left            =   120
                  TabIndex        =   127
                  Top             =   3480
                  Width           =   2295
               End
               Begin VB.CommandButton cmdARSMan 
                  Caption         =   "Uninstall Manager..."
                  Height          =   480
                  Left            =   120
                  TabIndex        =   126
                  Top             =   2880
                  Width           =   2295
               End
               Begin VB.CommandButton cmdADSSpy 
                  Caption         =   "ADS Spy..."
                  Height          =   360
                  Left            =   120
                  TabIndex        =   123
                  Top             =   2400
                  Width           =   2295
               End
               Begin VB.CommandButton cmdDeleteService 
                  Caption         =   "Delete a Windows service..."
                  Height          =   360
                  Left            =   120
                  TabIndex        =   122
                  Top             =   1920
                  Width           =   2295
               End
               Begin VB.CommandButton cmdDelOnReboot 
                  Caption         =   "Delete a file on reboot..."
                  Height          =   480
                  Left            =   120
                  TabIndex        =   119
                  Top             =   1320
                  Width           =   2295
               End
               Begin VB.CommandButton cmdHostsManager 
                  Caption         =   "Hosts file manager"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   118
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.CommandButton cmdProcessManager 
                  Caption         =   "Process manager"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   116
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.Label lblDigiSigCheckerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Verify authenticode digital signatures on the given list of files"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   131
                  Top             =   4120
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblRegKeyUnlockerAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Reset permissions on the given registry keys list"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   129
                  Top             =   3540
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblARSManAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Managing items in the Add/Remove Software list"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   128
                  Top             =   2930
                  Width           =   4410
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblADSSpyAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Scan for hidden data streams"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   125
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
                  TabIndex        =   124
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
                  TabIndex        =   121
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
                  TabIndex        =   120
                  Top             =   900
                  Width           =   4650
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblProcessManagerAbout 
                  AutoSize        =   -1  'True
                  Caption         =   "Small process manager, working much like the Task Manager"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   117
                  Top             =   360
                  Width           =   4320
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
               TabIndex        =   112
               Top             =   0
               Width           =   7335
               Begin VB.CommandButton cmdStartupList 
                  Caption         =   "StartupList scan"
                  Height          =   465
                  Left            =   120
                  TabIndex        =   113
                  Top             =   480
                  Width           =   2295
               End
               Begin VB.Label lblStartupListAbout 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   $"frmMain.frx":9180
                  Height          =   1035
                  Left            =   2520
                  TabIndex        =   114
                  Top             =   0
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
               Height          =   2295
               Left            =   120
               TabIndex        =   108
               Top             =   7800
               Width           =   7335
               Begin VB.CheckBox chkUpdateSilently 
                  Caption         =   "Update in silent mode"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   148
                  Top             =   1080
                  Width           =   4695
               End
               Begin VB.CheckBox chkUpdateToTest 
                  Caption         =   "Update to test versions"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   147
                  Top             =   740
                  Width           =   4575
               End
               Begin VB.CheckBox chkCheckUpdatesOnStart 
                  Caption         =   "Check updates automatically on program startup"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   146
                  Top             =   390
                  Width           =   4695
               End
               Begin VB.OptionButton OptProxyDirect 
                  Caption         =   "Direct connection"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   74
                  Top             =   960
                  Width           =   2175
               End
               Begin VB.CheckBox chkSocks4 
                  Caption         =   "Socks4"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   1440
                  TabIndex        =   75
                  Top             =   1480
                  Width           =   855
               End
               Begin VB.OptionButton optProxyManual 
                  Caption         =   "Proxy"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   76
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.OptionButton optProxyIE 
                  Caption         =   "IE settings"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   77
                  Top             =   1200
                  Value           =   -1  'True
                  Width           =   2055
               End
               Begin VB.TextBox txtUpdateProxyPass 
                  Enabled         =   0   'False
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   5640
                  PasswordChar    =   "*"
                  TabIndex        =   78
                  Top             =   1800
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdateProxyLogin 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   79
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.CheckBox chkUpdateUseProxyAuth 
                  Caption         =   "Use authorization"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   80
                  Top             =   1800
                  Width           =   2175
               End
               Begin VB.TextBox txtUpdateProxyPort 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   81
                  Text            =   "8080"
                  Top             =   1440
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdateProxyHost 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   85
                  Text            =   "127.0.0.1"
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.CommandButton cmdCheckUpdate 
                  Caption         =   "Check for update online"
                  Height          =   480
                  Left            =   240
                  TabIndex        =   109
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.Label lblUpdatePass 
                  Caption         =   "Password"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   4800
                  TabIndex        =   7
                  Top             =   1820
                  Width           =   855
               End
               Begin VB.Label lblUpdateLogin 
                  Caption         =   "Login"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   2520
                  TabIndex        =   110
                  Top             =   1820
                  Width           =   615
               End
               Begin VB.Label lblUpdatePort 
                  Caption         =   "Port"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   4800
                  TabIndex        =   111
                  Top             =   1480
                  Width           =   615
               End
               Begin VB.Label lblUpdateServer 
                  Caption         =   "Server"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   2520
                  TabIndex        =   137
                  Top             =   1480
                  Width           =   615
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
               TabIndex        =   106
               Top             =   11160
               Visible         =   0   'False
               Width           =   7335
               Begin VB.CommandButton cmdTaskScheduler 
                  Caption         =   "Task Scheduler Log"
                  Height          =   345
                  Left            =   240
                  TabIndex        =   107
                  Top             =   360
                  Width           =   2055
               End
            End
         End
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
         Begin VB.Frame fraConfigTabsNested 
            BorderStyle     =   0  'None
            Height          =   7815
            Left            =   0
            TabIndex        =   70
            Top             =   -120
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
               Height          =   1800
               Left            =   0
               TabIndex        =   71
               Top             =   3120
               Width           =   7935
               Begin VB.CheckBox chkFontWholeInterface 
                  Caption         =   "Apply selected font on whole interface"
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   145
                  Top             =   1400
                  Width           =   4695
               End
               Begin VB.ComboBox cmbFontSize 
                  Height          =   315
                  Left            =   2280
                  Style           =   2  'Dropdown List
                  TabIndex        =   144
                  Top             =   1380
                  Width           =   735
               End
               Begin VB.ComboBox cmbFont 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   141
                  Top             =   1380
                  Width           =   2055
               End
               Begin VB.CheckBox chkConfigMinimizeToTray 
                  Caption         =   "Minimize program to system tray when clicking _ button"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   89
                  Top             =   840
                  Width           =   6015
               End
               Begin VB.CheckBox chkSkipErrorMsg 
                  Caption         =   "Do not show error messages"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   88
                  Top             =   600
                  Width           =   4695
               End
               Begin VB.CheckBox chkSkipIntroFrameSettings 
                  Caption         =   "Do not show main menu at startup"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   68
                  Top             =   360
                  Width           =   4575
               End
               Begin VB.Label lblFontSize 
                  Caption         =   "Size"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   143
                  Top             =   1140
                  Width           =   975
               End
               Begin VB.Label lblFont 
                  Caption         =   "Font"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   142
                  Top             =   1140
                  Width           =   1935
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
               TabIndex        =   98
               Top             =   120
               Width           =   3975
               Begin VB.CheckBox chkAdditionalScan 
                  Caption         =   "Additional scan"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   102
                  ToolTipText     =   "Include specific sections, like O4 - RenameOperations, O21 - Column Hanlders / Context menu, O23 - Drivers e.t.c."
                  Top             =   1080
                  Width           =   3015
               End
               Begin VB.CheckBox chkAdvLogEnvVars 
                  Caption         =   "Environment variables"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   101
                  ToolTipText     =   "Include environment variables in logfile"
                  Top             =   720
                  Width           =   3015
               End
               Begin VB.CheckBox chkLogProcesses 
                  Caption         =   "Processes"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   100
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
               TabIndex        =   72
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
               TabIndex        =   99
               Top             =   120
               Width           =   3855
               Begin VB.CheckBox chkConfigStartupScan 
                  Caption         =   "Add HiJackThis to startup"
                  Height          =   270
                  Left            =   120
                  TabIndex        =   82
                  ToolTipText     =   "Run HiJackThis scan at Windows startup and show results (if only items are found)"
                  Top             =   1120
                  Width           =   3255
               End
               Begin VB.CheckBox chkDoMD5 
                  Caption         =   "Calculate MD5"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   105
                  ToolTipText     =   "Calculate checksum of files (MD5) is possible"
                  Top             =   900
                  Width           =   3015
               End
               Begin VB.CheckBox chkIgnoreAll 
                  Caption         =   "Ignore ALL Whitelists"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   104
                  ToolTipText     =   "Include in log any entries regardless whitelist"
                  Top             =   610
                  Width           =   3015
               End
               Begin VB.CheckBox chkIgnoreMicrosoft 
                  Caption         =   "Hide Microsoft entries"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   103
                  ToolTipText     =   "Do not include in log files and registry related to Microsoft"
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   3015
               End
            End
         End
         Begin VB.VScrollBar vscSettings 
            Height          =   4160
            LargeChange     =   20
            Left            =   8040
            Max             =   100
            TabIndex        =   69
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
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
            TabIndex        =   97
            Top             =   3960
            Width           =   6375
         End
         Begin VB.CommandButton cmdConfigBackupCreateSRP 
            Caption         =   "Create restore point"
            Height          =   720
            Left            =   7440
            TabIndex        =   96
            Top             =   3600
            Width           =   990
         End
         Begin VB.CommandButton cmdConfigBackupCreateRegBackup 
            Caption         =   "Create registry backup"
            Height          =   720
            Left            =   7440
            TabIndex        =   95
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
            Caption         =   $"frmMain.frx":9248
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   42
            Top             =   0
            Width           =   8250
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
            Caption         =   "Remove"
            Height          =   495
            Left            =   7440
            TabIndex        =   22
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigIgnoreDelAll 
            Caption         =   "Clear all"
            Height          =   495
            Left            =   7440
            TabIndex        =   23
            Top             =   1080
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   138
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
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "No suspicious items found!"
      Top             =   1560
      Visible         =   0   'False
      Width           =   4695
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
      Width           =   8275
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
      Caption         =   $"frmMain.frx":932D
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   45
      Visible         =   0   'False
      Width           =   8500
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":9405
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
            Caption         =   "Unlock && Reset permissions..."
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
' App key: HKLM\Software\TrendMicro\HiJackThisFork

Option Explicit

Private Const HJT_ALPHA             As Boolean = False
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

Private JumpFileCache() As FIX_FILE
Private JumpRegCache()  As FIX_REG_KEY

Public Sub Test()
    'If you need something to test after program started and initialized all required variables, please use this sub.
    
    'Debug.Print Reg.GetString(0, "HKLM\System\CurrentControlSet\Control\Session Manager", "PendingFileRenameOperations")
    
    'Debug.Print Reg.GetKeyVirtualType(HKCU, "Software\Policies\Microsoft\Internet Explorer\Control Panel")
    '2 - KEY_VIRTUAL_USUAL
    '4 - KEY_VIRTUAL_SHARED
    '8 - KEY_VIRTUAL_REDIRECTED

    'Debug.Print IsMicrosoftDriverFile("C:\Windows\system32\DRIVERS\VMNET.SYS")
    
'    Dim aSect(), i&
'    aSect = Array("R0", "R1", "R2", "R3", "R4", "F0", "F1", "F2", "F3", "O1", "O2", "O3", "O4", "O5", "O6", "O7", "O8", "O9", "O10", _
'            "O11", "O12", "O13", "O14", "O15", "O16", "O17", "O18", "O19", "O20", "O21", "O22", "O23", "O24", "O25", "O26")
'
'    For i = 0 To UBound(aSect)
'        Call GetInfo(aSect(i) & " - " & String$(150, "X"))
'    Next
    
'    Dim hFile As Long
'    Dim i As Long, j As Long
'    Dim aSubKeys() As String
'    Dim CSKey As String
'
'    If OpenW(AppPath() & "\ControlSet.txt", FOR_OVERWRITE_CREATE, hFile) Then
'        PrintBOM hFile
'
'        PrintW hFile, "Enumerating method" & vbCrLf, True
'
'        For i = 1 To Reg.NtEnumSubKeysToArray(HKLM, "SYSTEM", aSubKeys)
'            KeyShowInfo HKLM, "SYSTEM\" & aSubKeys(i), hFile
'        Next
'
'        PrintW hFile, vbCrLf & "Direct query method" & vbCrLf, True
'
'        For j = 0 To 110
'            CSKey = IIf(j = 0, "System\CurrentControlSet", "System\ControlSet" & Format$(j, "000"))
'            KeyShowInfo HKLM, CSKey, hFile
'        Next
'
'        CloseW hFile
'    End If
    
'    Stop
End Sub

'Private Sub KeyShowInfo(hHive As Long, sSubKey As String, hFile As Long)
'    Dim sSymTarget As String
'    Dim sLog As String
'    Dim bExist As Boolean
'    Dim bExist2 As Boolean
'    Dim hKey As Long
'    Dim hKey2 As Long
'
'    bExist = False
'    If STATUS_SUCCESS = Reg.WrapNtOpenKeyEx(hHive, sSubKey, WRITE_OWNER, hKey, , False) Then
'        bExist = True
'        NtClose hKey
'    End If
'
'    bExist2 = False
'    If STATUS_SUCCESS = Reg.WrapNtOpenKeyEx(hHive, sSubKey, WRITE_OWNER, hKey, , True) Then
'        bExist2 = True
'        NtClose hKey2
'    End If
'
'    If bExist <> bExist2 Then
'        sLog = "Warning: Key exist (with flag KEY_WOW64_64KEY) returns " & bExist & ", when without flag it returns " & bExist2
'        PrintW hFile, sLog, True
'    End If
'
'    Call Reg.NtGetKeyVirtualType(hHive, sSubKey, True, sSymTarget)
'
'    sLog = "HKLM\" & sSubKey & " => " & IIf(bExist, " (exist, handle = " & hKey & ")", "(not exist, handle = " & hKey & ")") & _
'        IIf(Len(sSymTarget) <> 0, " => (symlink) " & sSymTarget, "")
'
'    PrintW hFile, sLog, True
'End Sub


' Tips on functions:

'1. Use AddWarning() to append text to the end of the log, before debugging info.

Private Sub Form_Load()
    Static bInit As Boolean
    
    pvSetFormIcon Me
    
    g_HwndMain = Me.hwnd
    
    If bInit Then
        If Not bAutoLogSilent Then
            MsgBoxW "Critical error. Main form is initialized twice!"
            End
        End If
    Else
        bInit = True
        If gNoGUI Then Me.Hide
        FormStart_Stady1
        If g_NeedTerminate Then
            Me.WindowState = vbMinimized
        End If
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
    
    Dim Ctl   As Control
    Dim Btn   As CommandButton
    Dim ChkB  As CheckBox
    Dim OptB  As OptionButton
    Dim Fra   As Frame
    Dim i     As Long
    Dim Salt  As String
    Dim Ver   As Variant
    Dim sCmdLine As String
    
    AppendErrorLogCustom "frmMain.FormStart_Stady1 - Begin"
    
    If HJT_ALPHA Then bIsAlpha = True
    If HJT_BETA Then bIsBeta = True
    
    StartupListVer = STARTUP_LIST_VERSION
    ADSspyVer = ADS_SPY_VERSION
    ProcManVer = PROC_MAN_VERSION
    UninstManVer = UNINST_MAN_VERSION
    
    g_HJT_Items_Count = 36 'R + F + O1-...-O26 + Subsections (for progressbar)

    DisableSubclassing = False
    If inIDE Then DisableSubclassing = True
    
    If bAutoLogSilent Then 'timeout timer
        If Perf.MAX_TimeOut <> 0 Then
            Timer1.Interval = 1000
            Timer1.Enabled = True
        End If
    End If
    
    If Not DisableSubclassing And Not bAutoLogSilent Then
        SubClassScroll True
        RegisterHotKey Me.hwnd, HOTKEY_ID_CTRL_A, MOD_CONTROL Or MOD_NOREPEAT, vbKeyF
        RegisterHotKey Me.hwnd, HOTKEY_ID_CTRL_F, MOD_CONTROL Or MOD_NOREPEAT, vbKeyF
    End If
    
    ' Result -> sWinVersion (global)
    sWinVersion = GetWindowsVersion()   'to get bIsWin64 and so...
          
    AppVerPlusName = "HiJackThis Fork " & IIf(bIsAlpha, "(Alpha) ", IIf(bIsBeta, "(Beta) ", vbNullString)) & _
        "by Alex Dragokas v." & AppVerString
    
    'Ver. on Misc tools window
    'lblVersionRaw.Caption = AppVerString & IIf(bIsAlpha, " (Alpha)", IIf(bIsBeta, " (Beta)", vbNullString))
    
    If Not bAutoLogSilent Then
        Call PictureBoxRgn(pictLogo, RGB(255, 255, 255))
    End If
    
    'enable x64 redirection
    'ToggleWow64FSRedirection True ' -> moved to GetWindowsVersion()
        
    InitVariables   'sWinDir, classes init. and so.
    
    SetCurrentDirectory StrPtr(AppPath())
    
    'FixLog = BuildPath(AppPath(), "\HJT_Fix.log")           'not used yet
    'If FileExists(FixLog) Then DeleteFileWEx StrPtr(FixLog)
    
    bPolymorph = (InStr(1, AppExeName(), "_poly", 1) <> 0) Or (StrComp(GetExtensionName(AppExeName(True)), ".pif", 1) = 0)
    
    If Not bPolymorph Then
        Me.Caption = AppVerPlusName
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
    
    LoadLanguageList
    LoadResources
    
    lblMD5.Caption = ""
    txtNothing.ZOrder 1
    
    ' if Win XP/2003 -> disable all window styles from buttons on frames
    If bIsWinXP Then
        For Each Ctl In Me.Controls
            If TypeName(Ctl) = "CommandButton" Then
                Set Btn = Ctl
                SetWindowTheme Btn.hwnd, StrPtr(" "), StrPtr(" ")
            ElseIf TypeName(Ctl) = "CheckBox" Then
                Set ChkB = Ctl
                SetWindowTheme ChkB.hwnd, StrPtr(" "), StrPtr(" ")
            ElseIf TypeName(Ctl) = "OptionButton" Then
                Set OptB = Ctl
                SetWindowTheme OptB.hwnd, StrPtr(" "), StrPtr(" ")
            End If
        Next
        Set OptB = Nothing
        Set ChkB = Nothing
        Set Btn = Nothing
        Set Ctl = Nothing
    End If
    ' disable visual bugs with .caption property of frames (XP+)
    If OSver.MajorMinor >= 5.1 Then
        For Each Ctl In Me.Controls
            If TypeName(Ctl) = "Frame" Then
                Set Fra = Ctl
                'If Fra.Name = "fraHostsMan" Or Fra.Name = "fraUninstMan" Then
                    SetWindowTheme Fra.hwnd, StrPtr(" "), StrPtr(" ")
                'End If
            End If
        Next
        Set Fra = Nothing
    End If
    
    'move frame with "AnalyzeThis" button to the left a little bit (Vista+)
    If OSver.IsWindowsVistaOrGreater Then
        fraSubmit.Left = fraSubmit.Left - 65
    End If
    
    If OSver.IsLocalSystemContext Then
        'block some tools to prevent damage to system or output of inaccurate data
        mnuFileInstallHJT.Enabled = False
        mnuToolsRegUnlockKey.Enabled = False
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
            Case "CommandButton"
                Set ControlsEvent(i).BtnInArr = Ctl
            Case "TextBox"
                Set ControlsEvent(i).txtBoxInArr = Ctl
            Case "ListBox"
                Set ControlsEvent(i).lstBoxInArr = Ctl
            'Case "Label"
            '    'Set ControlsEvent(i).LblInArr = ctl
            Case "CheckBox"
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
    
    'header of tracing log
    AppendErrorLogCustom vbCrLf & vbCrLf & "Logfile ( tracing ) of HiJackThis Fork v." & AppVerString & vbCrLf & vbCrLf & _
        "Command line: " & AppPath(True) & " " & Command() & vbCrLf & vbCrLf & MakeLogHeader()
    
    LoadStuff 'regvals, filevals, safelspfiles, safeprotocols
    GetLSPCatalogNames
    LoadSettings ' must go after LoadStuff()
    
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
    If InStr(1, Command$(), "ihatewhitelists", 1) > 0 Then bIgnoreAllWhitelists = True: bHideMicrosoft = False 'must go after LoadSettings !!!
    '/default
    If InStr(1, Command$(), "default", 1) > 0 Then bLoadDefaults = True
    If bLoadDefaults Then
        bAutoSelect = False
        bConfirm = True
        bMakeBackup = True
        bLogProcesses = True
        bLogModules = False
        bLogEnvVars = False
        bAdditional = False
        bSkipErrorMsg = False
        bMinToTray = False
        bCheckForUpdates = False
        bHideMicrosoft = True
        bIgnoreAllWhitelists = False
        bMD5 = False
    End If
    '/skipIgnoreList
    If InStr(1, Command$(), "skipIgnoreList", 1) > 0 Then
        bSkipIgnoreList = True
        IsOnIgnoreList "", EraseList:=True
    End If
    
    '/Area:xxx
    sCmdLine = Replace$(Command(), ":", "+")
    
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
        
        CenterForm Me
    End If
    
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
    
    If Not bAutoLogSilent Then
        If CheckForStartedFromTempDir() Then
            g_NeedTerminate = True
            Exit Sub
        End If
    End If
    
    If Not bIsWinNT Then cmdDeleteService.Enabled = False
    
    If Not bAutoLogSilent Then
        SetMenuIcons Me.hwnd
    End If
    
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
    Dim bSilentUninst As Boolean
    
    If bInit Then
        Exit Sub
    Else
        bInit = True
    End If
    
    '/silentuninstall
    bSilentUninst = (InStr(1, Command$(), "silentuninstall", 1) > 0)
    
    '/uninstall
    If (InStr(1, Command$(), "/uninstall", 1) > 0) Or (InStr(1, Command$(), "-uninstall", 1) > 0) Or bSilentUninst Then
        Me.Hide
        If Not HJT_Uninstall(bSilentUninst) Then
            g_ExitCodeProcess = 1
        End If
        Unload Me
        Exit Sub
    End If
    
    If g_NeedTerminate Then Unload Me: Exit Sub

    '/md5
    If InStr(1, Command$(), "/md5", 1) > 0 Or InStr(1, Command$(), "-md5", 1) > 0 Then bMD5 = True
    '/deleteonreboot
    If InStr(1, Command$(), "deleteonreboot", 1) > 0 Then
        SilentDeleteOnReboot UnQuote(Command$())
        Unload Me
        Exit Sub
    End If
    
    If (Not inIDE) And (Not bPolymorph) Then
        Err.Clear
        g_hMutex = CreateMutex(0&, 1&, StrPtr("mutex_HiJackThis_Fork"))
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) And 0 = Len(Command$()) Then
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
    
    '/install
    If InStr(1, Command(), "/install", 1) <> 0 Or _
      InStr(1, Command(), "-install", 1) <> 0 Then
    
        '/autostart
        If InStr(1, Command(), "autostart", 1) <> 0 Then
            InstallAutorunHJT True
        Else
            InstallHJT True, (InStr(1, Command(), "noGUI", 1) <> 0) '/noGUI
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
    If InStr(1, Command$(), "StartupList", 1) > 0 Then
        bStartupListSilent = True
        cmdN00bTools_Click
        Call chkConfigTabs_Click(3)
        cmdStartupList_Click
    End If
    
    '/SysTray
    If InStr(1, Command$(), "SysTray", 1) > 0 Then
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
                            If InStr(.RunObj, "\") = 0 Then
                                'find full path for relative name
                                'filename without full path can be used in database to do comparision by filename only (see: isInTasksWhiteList())
                                .RunObj = FindOnPath(.RunObj, True)
                            Else
                                'commented, because it doesn't matter: such check will be done further by O22 routine
                                'If Not FileExists(.RunObj) Then .RunObj = vbNullString
                            End If
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
                        
                        'additional check in case 'FindOnPath' didn't find executable
                        g_TasksWL(ID).RunObj = g_TasksWL(ID).RunObj & IIf(Len(g_TasksWL(ID).RunObj) = 0, "", "|") & .RunObj
                        g_TasksWL(ID).Args = g_TasksWL(ID).Args & IIf(Len(g_TasksWL(ID).Args) = 0, "", "|") & .Args
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
            g_ExitCodeProcess = 1067
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
    UnregisterHotKey Me.hwnd, HOTKEY_ID_CTRL_A
    UnregisterHotKey Me.hwnd, HOTKEY_ID_CTRL_F
    For Each Frm In Forms
        If Not (Frm Is Me) And Not (Frm.Name = "frmEULA") Then
            Unload Frm
            Set Frm = Nothing
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
    Set HE = Nothing
    
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
    If Not inIDE Then
        If FileExists(BuildPath(AppPath(), "MSComCtl.ocx")) Then
            Proc.ProcessRun AppPath(True), "/release:" & GetCurrentProcessId(), , vbMinimizedNoFocus, True
        End If
    End If
    Set oDictFileExist = Nothing
    
    If g_ExitCodeProcess <> 0 Then
        ExitProcess g_ExitCodeProcess
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim s$
    pvDestroyFormIcon Me
    ReleaseMutex
    ISL_Dispatch
    Close
    SetFontDefaults Nothing, True
    If g_hDebugLog <> 0 Then
        s = vbCrLf & "--" & vbCrLf & "Debug log closed because main form is terminated (!!!)"
        PutW g_hDebugLog, 1, StrPtr(s), LenB(s), True
        CloseHandle g_hDebugLog
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

Private Sub cmdRegKeyUnlocker_Click() 'Mic Tools -> Registry Keys unlocker
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
        
        With result
            .Section = "O23"
            .HitLineW = "O23 - Service: " & sServiceName & " (" & sDisplayName & ")"
            AddServiceToFix .service, DELETE_SERVICE, sServiceName
            
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
        Translate(1003) & " (*.*)|*.*|" & Translate(1956) & " (*.dll)|*.dll|" & Translate(1957) & " (*.exe)|*.exe", Me.hwnd)
    If Len(sFilename) = 0 Then Exit Sub
    
    DeleteFileOnReboot sFilename, True, True
End Sub

Private Sub cmdHostsManager_Click() 'Misc Tools -> 'Hosts' file manager
    fraConfigTabs(3).Visible = False
    'SubClassScroll False
    fraHostsMan.Visible = True
    ListHostsFile lstHostsMan, lblConfigInfo(14)
End Sub

Private Sub cmdHostsManBack_Click()
    fraHostsMan.Visible = False
    fraConfigTabs(3).Visible = True
    'SubClassScroll True
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

    txtNothing.Visible = False

    CloseProgressbar
    
    'SubClassScroll False
    frmMain.pictLogo.Visible = True
    'If cmdConfig.Caption = Translate(19) Then 'Report
    If cmdConfig.Tag = "1" Then 'Report
    
        AppendErrorLogCustom "SaveSettings initiated by clicking 'Main menu'."
        SaveSettings
        
        fraConfig.Visible = False
        fraHostsMan.Visible = False
        If chkConfigTabs(3).Value = 1 Then fraConfigTabs(3).Visible = True
        cmdConfig.Caption = Translate(18): cmdConfig.Tag = "0" 'Settings
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
    shpMD5Background.Visible = False
    chkSkipIntroFrame.Value = RegReadHJT("SkipIntroFrame", "0")
End Sub

'// List of Backups
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

'// None of above, just start the program
Private Sub cmdN00bClose_Click()
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
    'SubClassScroll True
End Sub

'// Online guide
Private Sub cmdN00bHJTQuickStart_Click()
    'ShellExecute Me.hWnd, "open", "http://tomcoyote.org/hjt/#Top", "", "", 1
    'szQSUrl = Translate(360) & "?hjtver=" & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    
    'szQSUrl = "https://www.bleepingcomputer.com/tutorials/how-to-use-hijackthis/"
    
    OpenURL "http://dragokas.com/tools/help/hjt_tutorial.html", "http://regist.safezone.cc/hijackthis_help/hijackthis.html", True
End Sub

'// Do a system scan and save a log file
Private Sub cmdN00bLog_Click()
    
    cmdScan.Caption = Translate(11) 'don't touch this!!!
    cmdScan.Tag = "1"
    
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    bAutoLog = True
    'SubClassScroll True
    cmdScan_Click
End Sub

'// Do a system scan only
Private Sub cmdN00bScan_Click()
    cmdScan.Caption = Translate(11) 'don't touch this!!!
    cmdScan.Tag = "1"
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    pictLogo.Visible = False
    'SubClassScroll True
    cmdScan_Click
End Sub

'// Misc Tools
Private Sub cmdN00bTools_Click()
    pictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    
    'lstResults.Visible = True
    
    'If cmdConfig.Caption = Translate(18) Then cmdConfig_Click
    
    cmdConfig.Caption = Translate(18): cmdConfig.Tag = "0"
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
    
    If idxLastTab = 0 And isInit Then
        UpdateIE_RegVals
    End If
    
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
        'SubClassScroll False 'unSubClass
        
    Case 1 'ignore list
        'SubClassScroll False 'unSubClass
        
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
        'SubClassScroll False 'unSubClass
        ListBackups
        
    Case 3 'Misc tools
        'SubClassScroll True ' mouse scrolling support
        
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
    
    'SubClassScroll True
    
    CloseProgressbar
    
    'If cmdConfig.Caption = Translate(18) Then   'Settings
    If cmdConfig.Tag = "0" Then
    
        pictLogo.Visible = False
        
        'chkSkipIntroFrameSettings.Value = CLng(RegReadHJT("SkipIntroFrame", "0"))
        
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        'picPaypal.Visible = False
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
        
        Call pvSetVisionForLabelResults '"Welcome to HJT" / or "Below are the results..."
        'picPaypal.Visible = True
        lstResults.Visible = True
        fraHostsMan.Visible = False
        If chkConfigTabs(3).Value = 1 Then fraConfigTabs(3).Visible = True
        cmdConfig.Caption = Translate(18): cmdConfig.Tag = "0"
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
    ListBackups
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
    
    If bRebootRequired Or ("1" = RegReadHJT("RebootRequired", "0")) Then
        'Cannot start restoring until the system will be rebooted!
        MsgBoxW Translate(1577), vbExclamation
        RestartSystem
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
    cmdScan.Tag = "1"
    
    If bRebootRequired Then RestartSystem
    
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
            BackupConcatLine(0&, 0&, BackupFormatDate(Now()), SRP_BACKUP_TITLE & " - " & nSeqNum & " - " & "Restore Point by HiJackThis"), 0
    End If
    cmdConfigBackupCreateSRP.Enabled = True
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
    Dim sTime       As String
    
    dNow = Now()
    dMidNight = GetDateAtMidnight(dNow)
    
    sTime = RegReadHJT("DateLastFix", "")
    
    If Len(sTime) <> 0 Then
        If StrBeginWith(sTime, "HJT:") Then
            sTime = Mid$(sTime, 6)
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
    
    Dim i&, j&, sPrefix$, pos&, sItem$
    Dim bFlushDNS As Boolean
    Dim bO24Fixed As Boolean
    Dim bO14Fixed As Boolean
    Dim bRestartExplorer As Boolean
    
    Dim result As SCAN_RESULT
    
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
    
    If lstResults.ListCount = 0 Then Exit Sub
    
    If (lstResults.ListCount = lstResults.SelCount) And (InStr(1, Command(), "StartupScan", 1) = 0) And (lstResults.SelCount > 5) Then '/startupscan
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
    
    LockInterface bAllowInfoButtons:=False, bDoUnlock:=False
    
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
    'bRebootRequired = False
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
            
            If GetScanResults(sItem, result) Then 'map ANSI string to Unicode
            
              If j = 0 Then
                MakeBackup result
              Else
                UpdateProgressBar sPrefix
                
                bRebootRequired = bRebootRequired Or result.Reboot
            
                Select Case sPrefix ' RTrim$(Left$(lstResults.List(i), 3))
                Case "R0", "R1", "R2": FixRegItem sItem, result
                Case "R3":             FixR3Item sItem, result
                Case "R4":             FixR4Item sItem, result
                Case "F0", "F1":       FixFileItem sItem, result
                Case "F2", "F3":       FixFileItem sItem, result
                'Case "N1", "N2", "N3", "N4": FixNetscapeMozilla sItem,Result
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
    Reg.FlushAll
    
    If bRestartExplorer Then RestartExplorer
    If bFlushDNS Then FlushDNS
    If bUpdatePolicyNeeded Then UpdatePolicy
    If bO24Fixed Then FixO24Item_Post ' restart shell
    
    UpdateProgressBar "Finish"
    lstResults.Clear
    
    bScanExecuted = False
       
    'if somewhere explorer.exe has been killed, but not launched
    If Not ProcessExist("explorer.exe", True) Then
        RestartExplorer
    End If
    
    If bRebootRequired Then
        RegSaveHJT "RebootRequired", 1
        RestartSystem ': bRebootRequired = False
    End If
    
    'CloseProgressbar 'leave progressBar visible to ensure the user saw completion of cure
    
    If Not inIDE Then MessageBeep MB_ICONINFORMATION
    
    LockInterface bAllowInfoButtons:=True, bDoUnlock:=True
    
    cmdFix.Enabled = False
    cmdFix.FontBold = False
    cmdScan.Caption = Translate(11)
    cmdScan.Tag = "1"
    'cmdScan.Caption = "Scan"
    cmdScan.FontBold = True
    
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
        cmdInfo.Enabled = False
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        'picPaypal.Visible = False
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
        Call pvSetVisionForLabelResults '"Welcome to HJT" / or "Below are the results..."
        cmdInfo.Enabled = True
        'picPaypal.Visible = True
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
        If lstResults.SelCount = 0 And lstResults.ListIndex = -1 Then
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
            'sync listbox records with the RAM
            RemoveFromScanResults lstResults.List(i)
        End If
    Next i
    IsOnIgnoreList "", UpdateList:=True
    
    For i = lstResults.ListCount - 1 To 0 Step -1
        If lstResults.Selected(i) Then lstResults.RemoveItem i
    Next i
    If lstResults.ListCount = 0 Then
        txtNothing.Visible = True
        txtNothing.ZOrder 0
        cmdFix.FontBold = False
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
    cmdFix.Enabled = False
    cmdSaveDef.Enabled = False
    
    If cmdScan.Caption = Translate(11) Then 'Scan
    
        bScanExecuted = True
        
        If bAutoLogSilent And Not bStartupScan Then
            Call SystemPriorityDowngrade(True)
        End If
        
        'first scan after rebooting ?
        bFirstRebootScan = ScanAfterReboot()
        
        ' Erase main W array of scan results
        ReInitScanResults
        
        cmdAnalyze.Enabled = False
    
        ' Clear Error Log
        ErrReport = ""
        
        CheckIntegrityHJT
        
        'pre-adding horizontal scrollbar
        If Not bAutoLogSilent Then
            SendMessage frmMain.lstResults.hwnd, LB_SETHORIZONTALEXTENT, 1500&, ByVal 0&
        End If
        
        ' *******************************************************************

        StartScan '<<<<<<<-------- Main scan routine
        
        If txtNothing.Visible Or Not bAutoLog Then UpdateProgressBar "Finish"
        
        lblMD5.Visible = False
        shpMD5Background.Visible = False
        
        SortSectionsOfResultList
        
        'add the horizontal scrollbar if needed
        If Not bAutoLogSilent Then
            AddHorizontalScrollBarToResults lstResults
        End If
        
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
        
        cmdScan.Caption = Translate(12)
        cmdScan.Tag = "2"
        cmdScan.FontBold = False
        
        If bAutoLog Then
            If Not bAutoLogSilent Then DoEvents
            HJT_SaveReport '<<<<<< ------ Saving report
        End If
        
        'we are on the results window frame?
        If lstResults.Visible Then
            cmdScan.Enabled = True
            cmdFix.Enabled = True
            cmdAnalyze.Enabled = True
            cmdSaveDef.Enabled = True
        End If
        
        CloseProgressbar True
        
        If Not bAutoLog Then
            If cmdFix.Visible And cmdFix.Enabled Then
                cmdFix.SetFocus
            End If
        End If
        
        bAutoLog = False
        
        If bAutoLogSilent And Not bStartupScan Then
            Call SystemPriorityDowngrade(False)
        End If
        
    Else    'Caption = Save...

        If bAutoLogSilent Then
            'LockInterface bAllowInfoButtons:=True, bDoUnlock:=True
            'LockMenu bDoUnlock:=True
        End If
        
        Call HJT_SaveReport
        
        UpdateProgressBar "Finish"
        
        cmdScan.Enabled = True
        cmdFix.Enabled = True
    End If
    
    'focus on 1-st element of list
    Me.lstResults.ListIndex = -1
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
    cmdFix.Enabled = True
    isRanHJT_Scan = False
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdStartupList_Click() 'Misc Tools -> StartupList scan
    RunStartupList False
End Sub

Private Sub RunStartupList(bModal As Boolean)
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
        '!!! vbModal is not working here !!!
        frmStartupList2.Show IIf(bModal, vbModal, vbModeless)
    Else
        MsgBoxW "Cannot unpack " & sPathComCtl, vbCritical
    End If
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
            
            If MsgBoxW(Translate(153), vbQuestion Or vbYesNo, "HiJackThis") = vbNo Then Exit Function
        End If
    End If
    
    KillOtherHJTInstances HJT_Location
    
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
            "/v /d /c (cd\& for /L %+ in (1,1,10) do ((timeout /t 1|| ping 127.1 -n 2)& rd /s /q """ & HJT_Install_Path & """&& exit))", _
            SysDisk, vbHide, True
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

Private Sub LoadSettings()
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "frmMain.LoadSettings - Begin"
    
    Dim bUseOldKey As Boolean, sCurLang$, WinHeight&, WinWidth&, lProxyType&, lSocksVer&
    
    bUseOldKey = (Not Reg.KeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThisFork")) And _
        Reg.KeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis")
    
    ' Scan area
    
    chkLogProcesses.Value = CInt(RegReadHJT("LogProcesses", "1", bUseOldKey))
    chkAdvLogEnvVars.Value = CInt(RegReadHJT("LogEnvVars", "0", bUseOldKey))
    chkAdditionalScan.Value = CInt(RegReadHJT("LogAdditional", "0", bUseOldKey))
    
    bLogProcesses = chkLogProcesses.Value
    bLogEnvVars = chkAdvLogEnvVars.Value
    bAdditional = chkAdditionalScan.Value

    ' Scan options
    
    chkIgnoreMicrosoft.Value = CInt(RegReadHJT("HideMicrosoft", "1", bUseOldKey))
    chkIgnoreAll.Value = CInt(RegReadHJT("IgnoreAllWhiteList", "0", bUseOldKey))
    chkDoMD5.Value = CInt(RegReadHJT("CalcMD5", "0", bUseOldKey))
    
    bHideMicrosoft = chkIgnoreMicrosoft.Value
    bIgnoreAllWhitelists = chkIgnoreAll.Value
    bMD5 = chkDoMD5.Value
    
    ' Fix & Backup
    
    chkBackup.Value = CInt(RegReadHJT("MakeBackup", "1", bUseOldKey))
    chkConfirm.Value = CInt(RegReadHJT("Confirm", "1", bUseOldKey))
    chkAutoMark.Value = CInt(RegReadHJT("AutoSelect", "0", bUseOldKey))
    
    bMakeBackup = chkBackup.Value
    bConfirm = chkConfirm.Value
    bAutoSelect = chkAutoMark.Value

    ' Interface
    
    chkSkipIntroFrameSettings.Value = CInt(RegReadHJT("SkipIntroFrame", "0", bUseOldKey))
    chkSkipIntroFrame.Value = CInt(RegReadHJT("SkipIntroFrame", "0", bUseOldKey))
    chkSkipErrorMsg.Value = CInt(RegReadHJT("SkipErrorMsg", "0", bUseOldKey))
    chkConfigMinimizeToTray.Value = CInt(RegReadHJT("MinToTray", "0", bUseOldKey))
    
    bSkipErrorMsg = chkSkipErrorMsg.Value
    bMinToTray = chkConfigMinimizeToTray.Value
    
    g_FontName = RegReadHJT("FontName", "Automatic")
    g_FontSize = RegReadHJT("FontSize", "Auto")
    chkFontWholeInterface.Value = CInt(RegReadHJT("FontOnInterface", "0"))
    
    sCurLang = RegReadHJT("LanguageFile", "English")
    WinHeight = CLng(RegReadHJT("WinHeight", "6600"))
    WinWidth = CLng(RegReadHJT("WinWidth", "9000"))
    
    ' Updates

    chkCheckUpdatesOnStart.Value = CInt(RegReadHJT("CheckForUpdates", "0", bUseOldKey))
    chkUpdateToTest.Value = CInt(RegReadHJT("UpdateToTest", "0"))
    chkUpdateSilently.Value = CInt(RegReadHJT("UpdateSilently", "0"))
    lProxyType = CInt(RegReadHJT("ProxyType", "1")) '0 - Direct, 1 - IE, 2 - Manual proxy
    OptProxyDirect.Value = Abs(lProxyType = 0)
    optProxyIE.Value = Abs(lProxyType = 1)
    optProxyManual.Value = Abs(lProxyType = 2)
    
    lSocksVer = CInt(RegReadHJT("ProxySocksVer", "0"))
    chkSocks4.Value = Abs(lSocksVer = 4)
    
    chkUpdateUseProxyAuth.Value = CInt(RegReadHJT("ProxyUseAuth", "0"))
    
    txtUpdateProxyHost.Text = RegReadHJT("ProxyServer", "")
    txtUpdateProxyPort.Text = RegReadHJT("ProxyPort", "")
    txtUpdateProxyLogin.Text = RegReadHJT("ProxyLogin", "")
    txtUpdateProxyPass.Text = DeCrypt(RegReadHJT("ProxyPass", ""))
    
    bCheckForUpdates = chkCheckUpdatesOnStart.Value
    bUpdateToTest = chkUpdateToTest.Value
    bUpdateSilently = chkUpdateSilently.Value

    ' Backup (restore point)
    
    chkShowSRP.Value = CInt(RegReadHJT("ShowSRP", "0", bUseOldKey))
    
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
    cboN00bLanguage.AddItem "French"
    cboN00bLanguage.AddItem "Russian"
    cboN00bLanguage.AddItem "Ukrainian"
    
    sFile = DirW$(BuildPath(AppPath(), "*.lng"), vbFile)
    
    Do While Len(sFile)
        If sFile <> "_Lang_EN.lng" And _
            sFile <> "_Lang_FR.lng" And _
            sFile <> "_Lang_RU.lng" And _
            sFile <> "_Lang_UA.lng" Then
                cboN00bLanguage.AddItem sFile
        End If
        sFile = DirW$()
    Loop
    
    sCurLang = RegReadHJT("LanguageFile", "English")  'HJT settings
    If bForceFR Then sCurLang = "French"
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
    ElseIf sFile = "French" Then
        LoadLanguage &H40C, bForceFR
        g_CurrentLang = sFile
    Else
        LoadLangFile sFile
        ReloadLanguageNative
        ReloadLanguage
    End If
    
    ' Do not save force mode state!
    If Not (bForceRU Or bForceEN Or bForceUA Or bForceFR) Then RegSaveHJT "LanguageFile", sFile
    
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

Private Sub mnuToolsUnlockAndDelFile_Click()    'Tools -> Delete File -> Unlock & Reset permissions...
    Dim sFilename$
    
'    'Enter file name:, Unlock & Delete
'    sFilename = InputBox(Translate(1952), Translate(1953))
'    If StrPtr(sFilename) = 0 Then Exit Sub
    
    'Unlock & Delete
    sFilename = OpenFileDialog(Translate(1953), Desktop, _
        Translate(1003) & " (*.*)|*.*|" & Translate(1956) & " (*.dll)|*.dll|" & Translate(1957) & " (*.exe)|*.exe", Me.hwnd)
    If 0 = Len(sFilename) Then Exit Sub
    
    'sFilename = UnQuote(EnvironW(sFilename))
    
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
    cmdARSMan_Click
End Sub

Private Sub mnuToolsShortcutsChecker_Click()    'Tools -> Shortcuts -> Check Browsers' LNK
    'Download Check Browsers' LNK by Dragokas & regist
    'and ask to run
    Dim sTool As String: sTool = BuildPath(AppPath, "Check Browsers LNK.exe")
    If FileExists(sTool) Then
        Proc.ProcessRun sTool, "", AppPath(False), 1, True
    Else
        DownloadUnzipAndRun "https://dragokas.com/tools/CheckBrowsersLNK.zip", "Check Browsers LNK.exe", False
    End If
End Sub
Private Sub mnuToolsShortcutsFixer_Click()      'Tools -> Shortcuts -> ClearLNK
    'Download ClearLNK by Dragokas
    'and ask to run
    Dim sTool As String: sTool = BuildPath(AppPath, "ClearLNK.exe")
    If FileExists(sTool) Then
        Proc.ProcessRun sTool, "", AppPath(False), 1, True
    Else
        DownloadUnzipAndRun "https://dragokas.com/tools/ClearLNK.zip", "ClearLNK.exe", False
    End If
End Sub

Private Sub mnuHelpManualEnglish_Click()
    Dim szQSUrl$: szQSUrl = "http://dragokas.com/tools/help/hjt_tutorial.html"
    ShellExecute Me.hwnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualRussian_Click()
    Dim szQSUrl$
    'szQSUrl = "https://safezone.cc/threads/25184/"
    szQSUrl = "https://regist.safezone.cc/hijackthis_help/hijackthis.html"
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
    RegSaveHJT "UpdateToTest", CStr(Abs(CInt(bUpdateToTest)))
    RegSaveHJT "UpdateSilently", CStr(Abs(CInt(bUpdateSilently)))
    RegSaveProxySettings
    
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
    
    Dim result As SCAN_RESULT
    Dim sItem As String
    Dim sPrefix As String
    Dim pos As Long
    Dim i As Long, j As Long
    Dim RegItems As Long
    Dim FileItems As Long
    'Dim sIniFile As String, sFile As String
    Dim idx As Long, XY As Long, XPix As Long, YPix As Long
    
    'select item by right click
    If Button = 2 Then
        XPix = X / Screen.TwipsPerPixelX
        YPix = Y / Screen.TwipsPerPixelY
        XY = YPix * 65536 + XPix
        idx = SendMessage(lstResults.hwnd, LB_ITEMFROMPOINT, 0&, ByVal XY)
        If idx >= 0 And idx <= (lstResults.ListCount - 1) Then
            lstResults.ListIndex = idx
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
        
        Erase JumpFileCache
        Erase JumpRegCache
        
        'building the jump list
        mnuResultJump.Enabled = False
        
        sItem = GetSelected_OrCheckedItem()
        
        If sItem <> "" Then
            pos = InStr(sItem, "-")
            If pos <> 0 Then
                sPrefix = Trim$(Left$(sItem, pos - 1))
            End If
            If GetScanResults(sItem, result) Then
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
Private Sub JumpListExtractFiles(aFixFile() As FIX_FILE, FileItems As Long)
    Dim bExists As Boolean
    Dim j As Long
    If AryPtr(aFixFile) Then
        For j = 0 To UBound(aFixFile)
            If FileItems >= MAX_JUMP_LIST_ITEMS Then Exit For
            FileItems = FileItems + 1
            
            bExists = FileExists(aFixFile(j).Path)
            mnuResultJumpFile(FileItems - 1).Caption = aFixFile(j).Path & IIf(bExists, "", " (no file)")
            
            If AryPtr(JumpFileCache) Then
                ReDim Preserve JumpFileCache(UBound(JumpFileCache) + 1)
            Else
                ReDim JumpFileCache(0)
            End If
            
            JumpFileCache(UBound(JumpFileCache)) = aFixFile(j)
        Next
    End If
End Sub

'RegItems++ (FileItems++)
'FIX_REG_KEY -> mnuResultJumpFile
'FIX_REG_KEY -> JumpRegCache()
'consider MAX_JUMP_LIST_ITEMS
Private Sub JumpListExtractRegistry(aFixReg() As FIX_REG_KEY, FileItems As Long, RegItems As Long)
    Dim bNoValue As Boolean
    Dim bExists As Boolean
    Dim j As Long
    If AryPtr(aFixReg) Then
        For j = 0 To UBound(aFixReg)
            With aFixReg(j)
                If .IniFile <> "" Then
                    
                    If FileItems < MAX_JUMP_LIST_ITEMS Then
                        FileItems = FileItems + 1
                        bExists = FileExists(.IniFile)
                        mnuResultJumpFile(FileItems - 1).Caption = .IniFile & " => [" & .Key & "], " & .Param & IIf(bExists, "", " (no file)")
                        
                        If AryPtr(JumpFileCache) Then
                            ReDim Preserve JumpFileCache(UBound(JumpFileCache) + 1)
                        Else
                            ReDim JumpFileCache(0)
                        End If
                        
                        JumpFileCache(UBound(JumpFileCache)).Path = .IniFile
                    End If
                Else
                    If RegItems < MAX_JUMP_LIST_ITEMS Then
                        RegItems = RegItems + 1
                        bExists = Reg.KeyExists(.Hive, .Key, .Redirected)
                        bNoValue = False
                        If (.ActionType And BACKUP_KEY) Or (.ActionType And REMOVE_KEY) Then
                        Else
                            bNoValue = Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected)
                        End If
                        mnuResultJumpReg(RegItems - 1).Caption = _
                          Reg.GetShortHiveName(Reg.GetHiveNameByHandle(.Hive)) & "\" & .Key & ", " & .Param & _
                          IIf(.Redirected, " (x32)", "") & IIf(bExists, "", " (no key)") & IIf(bNoValue, " (no value)", "")
                    
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

'// TODO: Why is it not working ??? Who intercepts en event ?
'Private Sub lstResults_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then cmdFix_Click
'End Sub

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
    ElseIf (lstResults.SelCount = 1) Or (lstResults.SelCount > 1 And lstResults.ListIndex = -1) Then 'checkbox
        For i = 0 To lstResults.ListCount - 1
            If lstResults.Selected(i) = True Then
                GetSelected_OrCheckedItem = lstResults.List(i)
                Exit For
            End If
        Next
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
    cmdScan.Tag = "1"
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

            sText = sText & vbCrLf & sSeparator & vbCrLf & FindLine(aSect(i) & " -", Translate(31)) & vbCrLf & sSeparator & vbCrLf & _
                Replace$(Translate(j), "\\p", "") & vbCrLf
        Next
        
        txtHelp.Text = sText
    
    Case 1: 'Keys
        txtHelp.Text = Translate(32)
    
    Case 2: 'Purpose, Donations
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
'    If bHideMicrosoft Then
'        If chkIgnoreAll.Value = 1 Then chkIgnoreAll.Value = 0
'    End If
    RegSaveHJT "HideMicrosoft", Abs(CLng(bHideMicrosoft))
End Sub

Private Sub chkIgnoreAll_Click()
    bIgnoreAllWhitelists = chkIgnoreAll.Value
    'If bIgnoreAllWhitelists Then
    '    If chkIgnoreMicrosoft.Value = 1 Then chkIgnoreMicrosoft.Value = 0
    'End If
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
    'RegSaveHJT "SkipIntroFrame", CStr(chkSkipIntroFrame.Value) 'should be commented
End Sub

Private Sub chkSkipErrorMsg_Click()
    bSkipErrorMsg = (chkSkipErrorMsg.Value = 1)
    RegSaveHJT "SkipErrorMsg", Abs(CLng(bSkipErrorMsg))
End Sub

Private Sub chkConfigMinimizeToTray_Click()
    bMinToTray = (chkConfigMinimizeToTray.Value = 1)
    RegSaveHJT "MinToTray", Abs(CLng(bMinToTray))
End Sub

'======= UPDATES and PROXY controls

Private Sub cmdCheckUpdate_Click() 'Misc Tools -> Check for update online
    CheckForUpdate False, bUpdateSilently, bUpdateToTest
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

Private Sub cmbFont_Click()
    If cmbFontSize.ListCount <> 0 Then
        SetFontByUserSettings
    End If
End Sub

Private Sub cmbFontSize_Click()
    SetFontByUserSettings
End Sub

Private Sub SetFontByUserSettings()
    On Error GoTo ErrorHandler:
    
    If bAutoLogSilent Then Exit Sub 'speed optimization
    
    Dim i As Long
    Dim Frm As Form
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
    
    For Each Frm In Forms
        SetAllFontCharset Frm, g_FontName, g_FontSize
        'SetMenuFont Frm.hwnd, g_FontName, g_FontSize
    Next
    
    RegSaveHJT "FontName", g_FontName
    RegSaveHJT "FontSize", g_FontSize
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SetFontByUserSettings", g_FontName, g_FontSize
    If inIDE Then Stop: Resume Next
End Sub

'Use new Font on result lists (and input windows) only ?
Private Sub chkFontWholeInterface_Click()
    RegSaveHJT "FontOnInterface", CStr(Abs(chkFontWholeInterface.Value))
    g_FontOnInterface = chkFontWholeInterface.Value
    SetFontByUserSettings
End Sub
