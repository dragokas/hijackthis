VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HiJackThis"
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
   Begin VB.CommandButton cmdHidden 
      Default         =   -1  'True
      Height          =   195
      Left            =   24960
      TabIndex        =   114
      Top             =   14760
      Width           =   75
   End
   Begin VB.PictureBox picPaypal 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   6240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmMain.frx":4B2A
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   86
      Top             =   -450
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame fraScan 
      Caption         =   "Scan && fix stuff"
      Height          =   1455
      Left            =   120
      TabIndex        =   31
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
         TabIndex        =   115
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Info on selected item..."
         Height          =   425
         Left            =   240
         TabIndex        =   3
         Top             =   880
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
         Height          =   425
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdFix 
         Caption         =   "Fix checked"
         Enabled         =   0   'False
         Height          =   425
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame fraSubmit 
      Height          =   1455
      Left            =   3000
      TabIndex        =   58
      Top             =   5880
      Width           =   2895
      Begin VB.CommandButton cmdAnalyze 
         Caption         =   "AnalyzeThis"
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         TabIndex        =   66
         Top             =   195
         Width           =   1935
      End
      Begin VB.CommandButton cmdMainMenu 
         Caption         =   "Main Menu"
         Height          =   495
         Left            =   720
         TabIndex        =   75
         Top             =   825
         Width           =   1455
      End
   End
   Begin VB.Frame fraOther 
      Caption         =   "Other stuff"
      Height          =   1455
      Left            =   6000
      TabIndex        =   32
      Top             =   5880
      Width           =   2775
      Begin VB.CommandButton cmdSaveDef 
         Caption         =   "Add checked to ignorelist"
         Enabled         =   0   'False
         Height          =   440
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Config..."
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Info..."
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
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
      TabIndex        =   28
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Width           =   1335
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1215
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkConfigTabs 
         Caption         =   "Main"
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
         Width           =   1215
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
         TabIndex        =   48
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CommandButton cmdHostsManOpen 
            Caption         =   "Open in Notepad"
            Height          =   425
            Left            =   3600
            TabIndex        =   54
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CommandButton cmdHostsManBack 
            Caption         =   "Back"
            Height          =   425
            Left            =   5160
            TabIndex        =   53
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton cmdHostsManToggle 
            Caption         =   "Toggle line(s)"
            Height          =   425
            Left            =   1800
            TabIndex        =   52
            Top             =   3240
            Width           =   1695
         End
         Begin VB.CommandButton cmdHostsManDel 
            Caption         =   "Delete line(s)"
            Height          =   425
            Left            =   120
            TabIndex        =   51
            Top             =   3240
            Width           =   1575
         End
         Begin VB.ListBox lstHostsMan 
            Height          =   2340
            IntegralHeight  =   0   'False
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   50
            Top             =   600
            Width           =   8175
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Note: changes to the hosts file take effect when you restart your browser."
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   55
            Top             =   3000
            Width           =   5415
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Hosts file located at: C:\WINDOWS\hosts"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   49
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
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CommandButton cmdConfigIgnoreDelSel 
            Caption         =   "Delete"
            Height          =   495
            Left            =   7440
            TabIndex        =   23
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigIgnoreDelAll 
            Caption         =   "Delete all"
            Height          =   495
            Left            =   7440
            TabIndex        =   24
            Top             =   1320
            Width           =   975
         End
         Begin VB.ListBox lstIgnore 
            Height          =   2625
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   480
            Width           =   7215
         End
         Begin VB.Label lblConfigInfo 
            Caption         =   "The following items will be ignored when scanning: "
            Height          =   585
            Index           =   5
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   7140
         End
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Height          =   9000
         Index           =   3
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Visible         =   0   'False
         Width           =   8295
         Begin VB.VScrollBar vscMiscTools 
            Height          =   4095
            LargeChange     =   20
            Left            =   7680
            Max             =   100
            SmallChange     =   20
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
         Begin VB.Frame fraMiscToolsScroll 
            BorderStyle     =   0  'None
            Height          =   10455
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Width           =   7455
            Begin VB.CommandButton cmdTaskScheduler 
               Caption         =   "Task Scheduler Log"
               Height          =   480
               Left            =   120
               TabIndex        =   119
               Top             =   9240
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.CheckBox chkIgnoreMicrosoft 
               Caption         =   "Ignore Microsoft files"
               Height          =   195
               Left            =   120
               TabIndex        =   118
               Top             =   5880
               Value           =   1  'Checked
               Width           =   3735
            End
            Begin VB.CheckBox chkIgnoreAll 
               Caption         =   "Ignore all Whitelists"
               Height          =   195
               Left            =   120
               TabIndex        =   117
               Top             =   6240
               Width           =   3735
            End
            Begin VB.CommandButton cmdUninstall 
               Caption         =   "Uninstall HiJackThis"
               Height          =   375
               Left            =   120
               TabIndex        =   111
               Top             =   6600
               Width           =   2295
            End
            Begin VB.CommandButton cmdARSMan 
               Caption         =   "Uninstall Manager..."
               Height          =   450
               Left            =   120
               TabIndex        =   93
               Top             =   4030
               Width           =   2295
            End
            Begin VB.CommandButton cmdDeleteService 
               Caption         =   "Delete a Windows service..."
               Height          =   375
               Left            =   120
               TabIndex        =   89
               Top             =   3000
               Width           =   2295
            End
            Begin VB.CheckBox chkAdvLogEnvVars 
               Caption         =   "Include environment variables in logfile"
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   5160
               Width           =   6015
            End
            Begin VB.CommandButton cmdADSSpy 
               Caption         =   "ADS Spy..."
               Height          =   375
               Left            =   120
               TabIndex        =   76
               Top             =   3480
               Width           =   2295
            End
            Begin VB.CommandButton cmdDelOnReboot 
               Caption         =   "Delete a file on reboot..."
               Height          =   450
               Left            =   120
               TabIndex        =   64
               Top             =   2400
               Width           =   2295
            End
            Begin VB.CommandButton cmdHostsManager 
               Caption         =   "Hosts file manager"
               Height          =   375
               Left            =   120
               TabIndex        =   63
               Top             =   1920
               Width           =   2295
            End
            Begin VB.CommandButton cmdProcessManager 
               Caption         =   "Process manager"
               Height          =   375
               Left            =   120
               TabIndex        =   62
               Top             =   1440
               Width           =   2295
            End
            Begin VB.TextBox txtCheckUpdateProxy 
               Height          =   285
               Left            =   2640
               TabIndex        =   60
               Top             =   8280
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.CommandButton cmdCheckUpdate 
               Caption         =   "Check for update online"
               Height          =   495
               Left            =   120
               TabIndex        =   59
               Top             =   7680
               Width           =   2295
            End
            Begin VB.CommandButton cmdStartupList 
               Caption         =   "StartupList scan"
               Height          =   495
               Left            =   120
               TabIndex        =   57
               Top             =   360
               Width           =   2295
            End
            Begin VB.CheckBox chkDoMD5 
               Caption         =   "Calculate MD5 of files if possible"
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   5520
               Width           =   6015
            End
            Begin VB.Label lblStartupListAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   $"frmMain.frx":4B70
               Height          =   1065
               Left            =   2520
               TabIndex        =   120
               Top             =   0
               Width           =   4800
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblUninstallHJT 
               Caption         =   "Remove all HiJackThis Registry entries, backups and quit"
               Height          =   255
               Left            =   2640
               TabIndex        =   112
               Top             =   6720
               Width           =   4335
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000010&
               X1              =   120
               X2              =   7320
               Y1              =   4680
               Y2              =   4680
            End
            Begin VB.Label lblInfo 
               Caption         =   "Open the integrated ADS Spy utility to scan for hidden data streams."
               Height          =   435
               Index           =   5
               Left            =   2520
               TabIndex        =   110
               Top             =   3540
               Width           =   3960
            End
            Begin VB.Label lblConfigInfo 
               Caption         =   "Opens a small editor for the 'hosts' file."
               Height          =   435
               Index           =   13
               Left            =   2520
               TabIndex        =   69
               Top             =   1960
               Width           =   4770
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "Testing stuff"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   120
               TabIndex        =   107
               Top             =   8880
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000010&
               Index           =   14
               X1              =   120
               X2              =   7320
               Y1              =   7200
               Y2              =   7200
            End
            Begin VB.Label lblInfo 
               Caption         =   "Open a utility to manage the items in the Add/Remove Software list."
               Height          =   495
               Index           =   7
               Left            =   2520
               TabIndex        =   94
               Top             =   4050
               Width           =   4095
            End
            Begin VB.Label lblInfo 
               Caption         =   "Delete a Windows Service (O23). USE WITH CAUTION! (WinNT4/2k/XP only)"
               Height          =   495
               Index           =   6
               Left            =   2520
               TabIndex        =   90
               Top             =   2980
               Width           =   4815
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000010&
               Index           =   6
               X1              =   120
               X2              =   7320
               Y1              =   8760
               Y2              =   8760
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   74
               Top             =   7320
               Width           =   1155
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "Advanced settings (these will not be saved)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   73
               Top             =   4800
               Width           =   3705
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   72
               Top             =   1100
               Width           =   1110
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000010&
               Index           =   0
               X1              =   120
               X2              =   7320
               Y1              =   1000
               Y2              =   1000
            End
            Begin VB.Label lblInfo 
               Caption         =   "If a file cannot be removed from memory, Windows can be setup to delete it when the system is restarted."
               Height          =   585
               Index           =   2
               Left            =   2520
               TabIndex        =   70
               Top             =   2400
               Width           =   4320
            End
            Begin VB.Label lblConfigInfo 
               Caption         =   "Opens a small process manager, working much like the Task Manager."
               Height          =   435
               Index           =   12
               Left            =   2520
               TabIndex        =   68
               Top             =   1500
               Width           =   4770
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "Use this proxy server (host:port) :"
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   67
               Top             =   8280
               Visible         =   0   'False
               Width           =   2490
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   65
               Top             =   0
               Width           =   2595
            End
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
         TabIndex        =   35
         Top             =   830
         Width           =   8440
         Begin VB.CheckBox chkConfigMinimizeToTray 
            Caption         =   "Hide program to system tray when clicking 'minimize' button"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   2300
            Width           =   6015
         End
         Begin VB.CheckBox chkSkipErrorMsg 
            Caption         =   "Do not show error messages"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   1600
            Width           =   4695
         End
         Begin VB.CheckBox chkConfigStartupScan 
            Caption         =   "Run HiJackThis scan at startup and show it when items are found"
            Height          =   385
            Left            =   120
            TabIndex        =   106
            Top             =   1880
            Width           =   8175
         End
         Begin VB.CheckBox chkShowN00bFrame 
            Caption         =   "Do not show intro frame at startup"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   1350
            Width           =   6975
         End
         Begin VB.CheckBox chkLogProcesses 
            Caption         =   "Include list of running processes in logfiles"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1080
            Width           =   7215
         End
         Begin VB.CheckBox chkAutoMark 
            Caption         =   "Mark everything found for fixing after scan"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   0
            Width           =   7455
         End
         Begin VB.TextBox txtDefStartPage 
            Height          =   285
            Left            =   2040
            TabIndex        =   16
            Top             =   3150
            Width           =   6375
         End
         Begin VB.TextBox txtDefSearchPage 
            Height          =   285
            Left            =   2040
            TabIndex        =   17
            Top             =   3480
            Width           =   6375
         End
         Begin VB.TextBox txtDefSearchAss 
            Height          =   285
            Left            =   2040
            TabIndex        =   18
            Top             =   3800
            Width           =   6375
         End
         Begin VB.TextBox txtDefSearchCust 
            Height          =   285
            Left            =   2040
            TabIndex        =   19
            Top             =   4120
            Width           =   6375
         End
         Begin VB.CheckBox chkConfirm 
            Caption         =   "Confirm fixing && ignoring of items (safe mode)"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   540
            Width           =   7455
         End
         Begin VB.CheckBox chkBackup 
            Caption         =   "Make backups before fixing items"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   270
            Width           =   7335
         End
         Begin VB.CheckBox chkIgnoreSafe 
            Caption         =   "Ignore non-standard but safe domains in IE (e.g. msn.com, microsoft.com)"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   810
            Width           =   8295
         End
         Begin VB.Label lblConfigInfo 
            Caption         =   "Below URLs will be used when fixing hijacked/unwanted MSIE pages:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   2800
            Width           =   8265
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Default Start Page:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   3180
            Width           =   1395
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Default Search Page:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   3500
            Width           =   1530
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Default Search Assistant:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   3810
            Width           =   1830
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Default Search Customize:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   37
            Top             =   4140
            Width           =   1905
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
         Height          =   3135
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CommandButton cmdConfigBackupDeleteAll 
            Caption         =   "Delete all"
            Height          =   495
            Left            =   7440
            TabIndex        =   26
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigBackupDelete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   7440
            TabIndex        =   25
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigBackupRestore 
            Caption         =   "Restore"
            Height          =   495
            Left            =   7440
            TabIndex        =   21
            Top             =   720
            Width           =   975
         End
         Begin VB.ListBox lstBackups 
            Height          =   2385
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   20
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
            Caption         =   $"frmMain.frx":4C38
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   43
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
         TabIndex        =   91
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CommandButton cmdUninstManUninstall 
            Caption         =   "Uninstall application"
            Height          =   425
            Left            =   4080
            TabIndex        =   113
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdUninstManSave 
            Caption         =   "Save list..."
            Height          =   425
            Left            =   5400
            TabIndex        =   105
            Top             =   3900
            Width           =   1455
         End
         Begin VB.TextBox txtUninstManCmd 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   1880
            Width           =   4095
         End
         Begin VB.TextBox txtUninstManName 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   1200
            Width           =   4095
         End
         Begin VB.CommandButton cmdUninstManRefresh 
            Caption         =   "Refresh list"
            Height          =   425
            Left            =   4080
            TabIndex        =   102
            Top             =   3900
            Width           =   1215
         End
         Begin VB.CommandButton cmdUninstManEdit 
            Caption         =   "Edit uninstall command"
            Height          =   425
            Left            =   6120
            TabIndex        =   101
            Top             =   2835
            Width           =   2055
         End
         Begin VB.CommandButton cmdUninstManBack 
            Caption         =   "Back"
            Height          =   425
            Left            =   6960
            TabIndex        =   99
            Top             =   3900
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdUninstManDelete 
            Caption         =   "Delete this entry"
            Height          =   425
            Left            =   4080
            TabIndex        =   98
            Top             =   2835
            Width           =   1935
         End
         Begin VB.CommandButton cmdUninstManOpen 
            Caption         =   "Open Add/Remove Software list"
            Height          =   425
            Left            =   4080
            TabIndex        =   97
            Top             =   3360
            Width           =   4150
         End
         Begin VB.ListBox lstUninstMan 
            Height          =   3540
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   92
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label lblInfo 
            Caption         =   $"frmMain.frx":4D1D
            Height          =   615
            Index           =   11
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   8175
         End
         Begin VB.Label lblInfo 
            Caption         =   "Uninstall command:"
            Height          =   255
            Index           =   10
            Left            =   4125
            TabIndex        =   96
            Top             =   1600
            Width           =   1455
         End
         Begin VB.Label lblInfo 
            Caption         =   "Name:"
            Height          =   255
            Index           =   8
            Left            =   4125
            TabIndex        =   95
            Top             =   960
            Width           =   1095
         End
      End
   End
   Begin VB.ListBox lstResults 
      Height          =   1755
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   6135
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
      TabIndex        =   77
      Top             =   1080
      Visible         =   0   'False
      Width           =   8655
      Begin VB.ComboBox cboN00bLanguage 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdN00bScan 
         Caption         =   "Do a system scan only"
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
         TabIndex        =   80
         Top             =   1440
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bHJTQuickStart 
         Caption         =   "Online Guide"
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
         TabIndex        =   83
         Top             =   3960
         Width           =   3975
      End
      Begin VB.CheckBox chkShowN00b 
         Caption         =   "Don't show this menu again"
         Height          =   255
         Left            =   360
         TabIndex        =   85
         Top             =   5520
         Width           =   5535
      End
      Begin VB.CommandButton cmdN00bClose 
         Caption         =   "None of above, just start the program"
         Height          =   495
         Left            =   360
         TabIndex        =   84
         Top             =   4800
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bTools 
         Caption         =   "Misc Tools"
         Height          =   495
         Left            =   360
         TabIndex        =   82
         Top             =   3000
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bBackups 
         Caption         =   "List of Backups"
         Height          =   495
         Left            =   360
         TabIndex        =   81
         Top             =   2400
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bLog 
         Caption         =   "Do a system scan and save a logfile"
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
         TabIndex        =   79
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Change language:"
         Height          =   195
         Index           =   9
         Left            =   6480
         TabIndex        =   108
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
         TabIndex        =   78
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
      TabIndex        =   29
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   240
         Width           =   1215
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   240
         Width           =   1215
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "Sections"
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
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
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
      TabIndex        =   33
      Text            =   "No suspicious items found!"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblMD5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calculating MD5 checksum of [file]..."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   600
      TabIndex        =   46
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
      TabIndex        =   47
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
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":4DEF
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   7455
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
      Caption         =   $"frmMain.frx":4EC7
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   8175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileSettings 
         Caption         =   "Settings"
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
      Begin VB.Menu mnuToolsProcMan 
         Caption         =   "Process Manager"
      End
      Begin VB.Menu mnuToolsHosts 
         Caption         =   "Hosts file Manager"
      End
      Begin VB.Menu mnuToolsDelFile 
         Caption         =   "DeleteFile"
         Begin VB.Menu mnuToolsUnlockAndDelFile 
            Caption         =   "Unlock && Delete File..."
         End
         Begin VB.Menu mnuToolsDelFileOnReboot 
            Caption         =   "Plan to Delete File on Reboot..."
         End
      End
      Begin VB.Menu mnuToolsDelServ 
         Caption         =   "Delete Service"
      End
      Begin VB.Menu mnuToolsUnlockRegKey 
         Caption         =   "Registry Key Unlocker"
      End
      Begin VB.Menu mnuToolsADSSpy 
         Caption         =   "Alternative Data Streams Spy"
      End
      Begin VB.Menu mnuToolsDigiSign 
         Caption         =   "Digital Signature checker"
      End
      Begin VB.Menu mnuToolsUninst 
         Caption         =   "Uninstall Manager"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "User's Manual"
         Begin VB.Menu mnuHelpManualEnglish 
            Caption         =   "English (outdated)"
         End
         Begin VB.Menu mnuHelpManualRussian 
            Caption         =   "Russian"
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
      Begin VB.Menu mnuHelpUpdate 
         Caption         =   "Download new version"
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
      Begin VB.Menu mnuResultDelim1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResultInfo 
         Caption         =   "Info on selected"
      End
      Begin VB.Menu mnuResultSearch 
         Caption         =   "Search on Google"
      End
      Begin VB.Menu mnuResultDelim2 
         Caption         =   "-"
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
Option Explicit

' Call stack note:
'
' "Do a system scan and save log file"
'    -> cmdN00bLog_Click -> cmdScan_Click -> SaveReport -> CreateLogFile (process list)
'

#Const DebugMode = False
#Const DebugToFile = False
#Const SilentAutoLog = False

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

Private Type MY_PROC_LOG
    ProcName    As String
    Number      As Long
    IsMicrosoft As Boolean
    EDS_issued  As String
End Type

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Private Declare Function SHRunDialog Lib "shell32.dll" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long
Private Declare Function PathFindOnPath Lib "Shlwapi.dll" Alias "PathFindOnPathW" (ByVal pszFile As Long, ppszOtherDirs As Any) As Boolean
Private Declare Function CreateMutex Lib "kernel32.dll" Alias "CreateMutexW" (ByVal lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As Long) As Long
Private Declare Function SetWindowTheme Lib "UxTheme.dll" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function MessageBeep Lib "user32.dll" (ByVal uType As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32.dll" (ByVal Handle As Long)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function ILCreateFromPath Lib "shell32.dll" Alias "ILCreateFromPathW" (ByVal pszPath As Long) As Long
Private Declare Function SHOpenFolderAndSelectItems Lib "shell32.dll" (ByVal pidlFolder As Long, ByVal cidl As Long, ByVal apidl As Long, ByVal dwFlags As Long) As Long
Private Declare Sub ILFree Lib "shell32.dll" (ByVal pIDL As Long)


Private Const INVALID_HANDLE_VALUE      As Long = -1&
Private Const ERROR_ALREADY_EXISTS      As Long = 183&
Private Const MB_ICONINFORMATION        As Long = &H40&

Private ControlsEvent() As New clsEvents
Private WithEvents FormSys As frmSysTray
Attribute FormSys.VB_VarHelpID = -1

Private txtHelpHasFocus As Boolean
Private AppVersion      As AppVersion
Private UninstData()    As UnintallManagerData
Private sKeyUninstall() As String
Private HitSorted()     As String
Private bSwitchingTabs  As Boolean
Private bIsBeta         As Boolean
Private bIsAlpha        As Boolean
Private szLogData       As String
Private hMutex          As Long
Private lToolsHeight    As Long



Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    
    Dim ctl   As Control
    Dim Btn   As CommandButton
    Dim ChkB  As CheckBox
    Dim OptB  As OptionButton
    Dim Fra   As Frame
    Dim i     As Long
    Dim Salt  As String
    Dim Ver   As Variant

    AppendErrorLogCustom "frmMain.Form_Load - Begin"
    

    bIsAlpha = True
    'bIsBeta = True
    
    StartupListVer = "2.11"
    ADSspyVer = "1.13"
    ProcManVer = "2.06"
    
    g_HJT_Items_Count = 27 'R + F + O25 (for progressbar)

    DisableSubclassing = False
    If inIDE Then DisableSubclassing = True

    #If DebugMode Then
        DebugMode = True
    #End If
    #If DebugToFile Then
        DebugToFile = True
    #End If
    #If SilentAutoLog Then
        bAutoLog = True: bAutoLogSilent = True
    #End If
    
    If InStr(1, Command(), "/autolog", 1) > 0 Then bAutoLog = True
    If InStr(1, Command(), "/silentautolog", 1) > 0 Then bAutoLog = True: bAutoLogSilent = True
    
    If bAutoLog Then Perf.StartTime = GetTickCount()
    
    sWinVersion = GetWindowsVersion()   'to get bIsWin64 and so...
    
    AppVerString = GetFilePropVersion(AppPath(True))
    
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
    
    AppVer = "HiJackThis Fork " & IIf(bIsBeta, "(Beta)", IIf(bIsAlpha, "(Alpha)", vbNullString)) & _
        " by Alex Dragokas v." & AppVerString
    
    
    'enable x64 redirection
    'ToggleWow64FSRedirection True ' -> moved to GetWindowsVersion()
    
    SetCurrentProcessPrivileges "SeDebugPrivilege"
    SetCurrentProcessPrivileges "SeBackupPrivilege"
    SetCurrentProcessPrivileges "SeRestorePrivilege"
    SetCurrentProcessPrivileges "SeTakeOwnershipPrivilege"
    SetCurrentProcessPrivileges "SeSecurityPrivilege"       'SACL
    'SetCurrentProcessPrivileges "SeAssignPrimaryTokenPrivilege"
    'SetCurrentProcessPrivileges "SeIncreaseQuotaPrivilege"
    
    InitVariables   'sWinDir, classes init. and so.
    
    bHideMicrosoft = True
    
    'FixLog = BuildPath(AppPath(), "\HJT_Fix.log")           'not used yet
    'If FileExists(FixLog) Then DeleteFileWEx StrPtr(FixLog)
    
    Me.Caption = AppVer
    
    'test stuff
    If inIDE Or InStr(1, AppExeName(), "test", 1) <> 0 Then
        'Task scheduler jobs log on 'misc section'.
        lblConfigInfo(22).Visible = True
        cmdTaskScheduler.Visible = True
        'Batch Verifier of Digitial signature
        'mnuToolsDigiSign.Visible = True
        lToolsHeight = 0
    Else
        lToolsHeight = 2200
    End If
    
    LoadLanguageSettings
    LoadLanguageList
    LoadResources
    
    lblMD5.Caption = ""
    
    ' if Win XP -> disable all window styles from buttons on frames
    If Not OSver.bIsVistaOrLater Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "CommandButton" Then
                Set Btn = ctl
                SetWindowTheme Btn.hWnd, StrPtr(" "), StrPtr(" ")
            ElseIf TypeName(ctl) = "CheckBox" Then
                Set ChkB = ctl
                SetWindowTheme ChkB.hWnd, StrPtr(" "), StrPtr(" ")
            ElseIf TypeName(ctl) = "OptionButton" Then
                Set OptB = ctl
                SetWindowTheme OptB.hWnd, StrPtr(" "), StrPtr(" ")
            End If
        Next
        Set OptB = Nothing
        Set ChkB = Nothing
        Set Btn = Nothing
        Set ctl = Nothing
    End If
    ' disable visual bugs with .caption property of frames
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Then
            Set Fra = ctl
            If Fra.Name = "fraHostsMan" Or Fra.Name = "fraUninstMan" Then
                SetWindowTheme Fra.hWnd, StrPtr(" "), StrPtr(" ")
            End If
        End If
    Next
    Set Fra = Nothing
    
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
            'Case "TextBox"
            '    Set ControlsEvent(i).txtBoxInArr = ctl
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
    Salt = RegGetDword(0, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "InstallDate")
    If Salt = "0" Then Salt = RegGetBinaryToString(0, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "DigitalProductId")
    sProgramVersion = "THOU SHALT NOT STEAL - " & Salt
    
    If InStr(1, Command(), "/uninstall", 1) > 0 Then
        Me.Hide
        cmdUninstall_Click
        Unload Me
        End
    End If

    If InStr(1, Command(), "/md5", 1) > 0 Then bMD5 = True
    If InStr(1, Command(), "/deleteonreboot", 1) > 0 Then
        SilentDeleteOnReboot UnQuote(Command())
        Unload Me
        End
    End If
    
    If InStr(1, Command(), "/ihatewhitelists", 1) > 0 Then bIgnoreAllWhitelists = True
    
    If Not inIDE Then
        Err.Clear
        hMutex = CreateMutex(0&, 1&, StrPtr("mutex_HiJackThis"))
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) And 0 = Len(Command()) Then
            If Not bAutoLogSilent Then
                If MsgBoxW(Translate(2), vbExclamation Or vbYesNo, "HiJackThis") = vbNo Then Unload Me: End
            End If
        End If
    End If
    
    If InStr(1, Command(), "/debug ", 1) <> 0 Or _
        StrEndWith(Command(), "/debug") Or _
        InStr(1, AppExeName(), "_debug", 1) <> 0 Then
        DebugMode = True
        DebugToFile = True ' /debug also initiate /debugtofile
        OpenDebugLogHandle
    End If
    
    fraConfig.Left = 120
    fraHelp.Left = 120
    fraConfig.Top = 120
    fraHelp.Top = 120
    fraMiscToolsScroll.Top = 0
    
    If Screen.Height >= 9000 Then
        Me.Height = CLng(RegReadHJT("WinHeight", "8000"))
        If Me.Height < 8000 Then Me.Height = 8000
    Else
        Me.Height = CLng(RegReadHJT("WinHeight", "6600"))
        If Me.Height < 6600 Then Me.Height = 6600
    End If
    Me.Width = CLng(RegReadHJT("WinWidth", "9000"))
    If Me.Width < 9000 Then Me.Width = 9000
    
    If RegReadHJT("ShowIntroFrame", "0") = "0" Then
        fraN00b.Visible = True
        fraScan.Visible = False
        fraOther.Visible = False
        lstResults.Visible = False
        fraSubmit.Visible = False
    Else
        'PictLogo.Visible = False
    End If
    If RegReadHJT("ShowIntroFrame", "0") = "0" Then
        chkShowN00b.value = 0
    Else
        chkShowN00b.value = 1
    End If
       
    LoadStuff 'regvals, filevals, safelspfiles, safeprotocols
    LoadSettings
    GetLSPCatalogNames
    CheckForReadOnlyMedia
    CheckDateFormat
    CheckForStartedFromTempDir
    'msgboxW "bIsUSADateFormat: " & bIsUSADateFormat
    
    If Not bIsWinNT Then cmdDeleteService.Enabled = False
    
    CheckAutoLog
    
    AppendErrorLogCustom "frmMain.Form_Load - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_Load"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckAutoLog()
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "frmMain.CheckAutoLog - Begin"

    DoEvents
    Me.Hide
    
    If InStr(1, Command(), "/noGUI", 1) = 0 Then
        DoEvents
        Me.Show
        DoEvents
        Me.Refresh
        DoEvents
    End If

    If bAutoLog Then
        'If bAutoLogSilent Then Me.WindowState = vbMinimized
        'If bAutoLogSilent Then Me.WindowState = vbMinimizedNoFocus
        cmdN00bClose_Click
        cmdScan_Click
        DoEvents
        If bAutoLogSilent Then Unload Me: End
    End If
    
    If InStr(1, Command(), "/StartupScan", 1) > 0 Then
        'Me.Show
        'DoEvents
        'Me.WindowState = vbMinimized
        cmdN00bClose_Click
        cmdScan_Click
        'DoEvents
        If lstResults.ListCount = 0 Then
            Unload Me: End
        Else
            'Me.WindowState = vbNormal
            Me.Show
        End If
    End If
    
'    If InStr(1, Command(), "/SilentStartupList", 1) > 0 Then
'        bStartupList = True
'        cmdN00bTools_Click
'        Call chkConfigTabs_Click(3)
'        cmdStartupList_Click
'        Unload Me: End
'    End If
    
    If InStr(1, Command(), "/StartupList", 1) > 0 Then
        bStartupListSilent = True
        cmdN00bTools_Click
        Call chkConfigTabs_Click(3)
        cmdStartupList_Click
    End If
    
    If InStr(1, Command(), "/SysTray", 1) > 0 Then
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
    
    Dim Lines() As String
    Dim sBuf As String
    Dim i As Long
    Dim Columns

    'Task Scheduler white list
    sBuf = StrConv(LoadResData(101, "CUSTOM"), vbUnicode, 1049)
    sBuf = Replace$(sBuf, vbCr, vbNullString)
    
    Lines = Split(sBuf, vbLf)
    ReDim g_TasksWL(UBound(Lines))
    
    For i = 1 To UBound(Lines)  'skip header
        Columns = Split(Lines(i), ";")
        '---------------------------
        'Columns (0) 'OSver
        'Columns (1) 'State     (not used)
        'Columns (2) 'Name
        'Columns (3) 'Dir
        'Columns (4) 'RunObj
        'Columns (5) 'Args
        'Columns (6) 'Note      (not used)
        'Columns (7) 'Error     (not used)
        '---------------------------
        
        If UBound(Columns) > 2 Then    ' protection: if last DB line is empty
            With g_TasksWL(i)
                .OSver = Val(Replace$(Columns(0), ",", "."))
                'select appropriate version from DB
                If .OSver = OSver.MajorMinor Then
                    .Name = UnScreenChar(CStr(Columns(2)))
                    .Directory = UnScreenChar(CStr(Columns(3)))
                    If UBound(Columns) > 3 Then
                        .RunObj = EnvironW(UnScreenChar(CStr(Columns(4))))
                    End If
                    If UBound(Columns) > 4 Then
                        .Args = UnScreenChar(CStr(Columns(5)))
                    End If
                    'Dictonary 'oDict.TaskWL_ID':
                    'value -> (dir + name of task)
                    'data -> id to 'g_TasksWL' user type array
            
                    If Not oDict.TaskWL_ID.Exists(.Directory & "\" & .Name) Then
                        oDict.TaskWL_ID.Add .Directory & "\" & .Name, i
                    Else
                        Debug.Print "Critical Database error: duplicate entry key: " & .Directory & "\" & .Name
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
    Dim frm As Form
    ToggleWow64FSRedirection True
    If Not g_UninstallState Then
        SaveAllSettings
        If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
            RegSaveHJT "WinHeight", CStr(Me.Height)
            RegSaveHJT "WinWidth", CStr(Me.Width)
        End If
        If Not inIDE Then RegSaveHJT "version", AppVerString
    End If
    'If bDebugMode Then Close #ffDebug
    SubClassScroll False
    For Each frm In Forms
        If Not (frm Is Me) Then
            Unload frm
            Set frm = Nothing
        End If
    Next
End Sub

Public Sub ReleaseMutex()
    If hMutex <> 0 Then CloseHandle hMutex
End Sub

Private Sub Form_Terminate()
    Set frmStartupList2 = Nothing
    
    If FileExists(BuildPath(AppPath(), "MSComCtl.ocx")) Then
        Proc.ProcessRun AppPath(True), "/release:" & GetCurrentProcessId(), , vbMinimizedNoFocus, True
    End If
    Set ErrLogCustomText = Nothing
    Set oDictFileExist = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseMutex
    ISL_Dispatch
    Close
End Sub

Private Sub chkAdvLogEnvVars_Click()
    bLogEnvVars = (chkAdvLogEnvVars.value = 1)
    RegSaveHJT "LogEnvVars", Abs(CLng(bLogEnvVars))
End Sub

Private Sub chkDoMD5_Click()
    bMD5 = (chkDoMD5.value = 1)
    RegSaveHJT "CalcMD5", Abs(CLng(bMD5))
End Sub

Private Sub chkShowN00b_Click()
    RegSaveHJT "ShowIntroFrame", CStr(chkShowN00b.value)
    chkShowN00bFrame.value = chkShowN00b.value
End Sub

Private Sub cmdADSSpy_Click()
    frmADSspy.Show
End Sub

Private Sub cmdAnalyze_Click()
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr("http://sourceforge.net/p/hjt/support-requests/"), 0&, 0&, 1
    
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

Private Sub cmdARSMan_Click()
    fraConfigTabs(3).Visible = False
    SubClassScroll False
    fraUninstMan.Visible = True
    cmdUninstManRefresh_Click
End Sub

Private Sub cmdDeleteService_Click()
    If Not bIsWinNT Then Exit Sub
    Dim sServiceName$, sWhiteList$, sDisplayName$, sFile$, sCompany$, J&
    
    sWhiteList = "Microsoft Corporation" '|" & _
                 '"Symantec Corporation|" & _
                 '"Trend Micro Inc.|" & _
                 '"Trend Micro Incorporated.|" & _
                 '"GRISOFT, s.r.o."
    
    sServiceName = InputBox(Translate(113), Translate(114))
'    sServiceName = InputBox("Enter the exact service name as it appears " & _
'                            "in the scan results, or the short name " & _
'                            "between brackets if that is listed." & vbCrLf & _
'                            "The service needs to be stopped and disabled." & vbCrLf & _
'                            "Services that belong to Microsoft, Symantec " & _
'                            "and several others are system-critical and cannot be deleted." & vbCrLf & vbCrLf & _
'                            "WARNING! When the service is deleted, it " & _
'                            "cannot be restored!", "Delete a Windows NT Service")
    If sServiceName = vbNullString Then Exit Sub
    
    If Not RegKeyExists(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName) Then
        MsgBoxW Replace$(Translate(115), "[]", sServiceName), vbExclamation
'        msgboxW "Service '" & sServiceName & "' was not found in the Registry." & vbCrLf & _
'               "Make sure you entered the name of the service correctly.", vbExclamation
        Exit Sub
    End If
    
    sFile = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "ImagePath")
    sDisplayName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "DisplayName")
    If sFile <> vbNullString Then
        'remove double quotes for long pathnames
        If Left$(sFile, 1) = """" Then sFile = Mid$(sFile, 2)
        If Right$(sFile, 1) = """" Then sFile = Left$(sFile, Len(sFile) - 1)
        
        'expand aliases
        sFile = Replace$(sFile, "%systemroot%", sWinDir, , 1, vbTextCompare)
        sFile = Replace$(sFile, "\systemroot", sWinDir, , 1, vbTextCompare)
        sFile = Replace$(sFile, "systemroot", sWinDir, , 1, vbTextCompare)
        
        'prefix windows folder if not specified
        If InStr(1, sFile, "system32\", vbTextCompare) = 1 Then
            sFile = sWinDir & "\" & sFile
        End If
        
        'remove parameters
        J = InStrRev(sFile, ".exe", , vbTextCompare) + 3
        If J < Len(sFile) And J > 3 Then sFile = Left$(sFile, J)
        
        'add .exe if not specified
        If InStr(1, sFile, ".exe", vbTextCompare) = 0 And _
           InStr(1, sFile, ".sys", vbTextCompare) = 0 Then
            If InStr(sFile, " ") > 0 Then
                sFile = Left$(sFile, InStr(sFile, " ") - 1)
                sFile = sFile & ".exe"
            End If
        End If
    Else
        sFile = "(no file)"
    End If
    
    sCompany = GetFilePropCompany(sFile)
    If sCompany = vbNullString Then sCompany = Translate(502) '"Unknown owner" '"?"
    
    If Not FileExists(sFile) Then sFile = sFile & " (" & Translate(503) & ")"  '" (file missing)"
    
    If InStr(1, sWhiteList, sCompany, vbTextCompare) > 0 Then
        MsgBoxW Translate(504), vbCritical  '"The service you entered is system-critical! It can't be deleted."
        Exit Sub
    End If
    
    If MsgBoxW(Translate(117) & vbCrLf & _
              Translate(505) & ": " & sServiceName & vbCrLf & _
              Translate(506) & ": " & sDisplayName & vbCrLf & _
              Translate(507) & ": " & sFile & vbCrLf & _
              Translate(508) & ": " & sCompany & vbCrLf & vbCrLf & _
              Translate(118), vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
'    If msgboxW("The following service was found:" & vbCrLf & _
'              "Short name: " & sServiceName & vbCrLf & _
'              "Full name: " & sDisplayName & vbCrLf & _
'              "File: " & sFile & vbCrLf & _
'              "Owner: " & sCompany & vbCrLf & vbCrLf & _
'              "Are you absolutely sure you want to delete this service?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
        
        DeleteNTService sServiceName
        
    End If
End Sub

Private Sub cmdDelOnReboot_Click()
    Dim sFileName$
    'Enter file name:, Delete on Reboot
    sFileName = InputBox(Translate(1950), Translate(1951))
    If StrPtr(sFileName) = 0 Then Exit Sub
    'sFileName = CmnDlgOpenFile(Translate(509), Translate(1003) & " (*.*)|*.*|" & Translate(511) & " (*.dll)|*.dll|" & Translate(512) & " (*.exe)|*.exe")
    'sFileName = CmnDlgOpenFile("Enter file to delete on reboot...", "All files (*.*)|*.*|DLL libraries (*.dll)|*.dll|Program files (*.exe)|*.exe")
    If sFileName = vbNullString Then Exit Sub
    DeleteFileOnReboot sFileName, True
End Sub

Private Sub cmdHostsManager_Click()
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
    sTxtProg = GetDefaultProgram(".txt")
    Shell sTxtProg & " " & """" & sHostsFile & """", vbNormalFocus
End Sub

Private Sub cmdHostsManToggle_Click()
    If lstHostsMan.ListIndex <> -1 And lstHostsMan.ListCount > 0 Then
        HostsToggleLine lstHostsMan
        ListHostsFile lstHostsMan, lblConfigInfo(14)
    End If
End Sub

Private Sub cmdMainMenu_Click()
    Dim hHive As Long
    
    CloseProgressbar
    
    SubClassScroll False
    'frmMain.PictLogo.Visible = True
    If cmdConfig.Caption = Translate(19) Then
    
        SaveAllSettings
        
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
        If chkConfigTabs(3).value = 1 Then fraConfigTabs(3).Visible = True
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
    chkShowN00b.value = RegReadHJT("ShowIntroFrame", "0")
End Sub

Private Sub cmdN00bBackups_Click()
    'PictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
    cmdConfig_Click
    chkConfigTabs_Click 2
End Sub

Private Sub cmdN00bClose_Click()
    'PictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    lblInfo(0).Visible = False
    lblInfo(1).Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
End Sub

Private Sub cmdN00bHJTQuickStart_Click()
'    fraN00b.Visible = False
'    fraScan.Visible = True
'    fraOther.Visible = True
'    fraSubmit.Visible = True
'    lstResults.Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
    'ShellExecute Me.hWnd, "open", "http://tomcoyote.org/hjt/#Top", "", "", 1
    Dim szQSUrl As String
    'szQSUrl = Translate(360) & "?hjtver=" & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    
    'If (IsRussianLangCode(OSVer.LangDisplayCode) Or bForceRU) And Not bForceEN Then
    If cboN00bLanguage.List(cboN00bLanguage.ListIndex) = "Russian" Then
        szQSUrl = "http://safezone.cc/threads/25184/"
    Else
        szQSUrl = "http://www.bleepingcomputer.com/tutorials/how-to-use-hijackthis/"
    End If
    
    'If True = IsOnline Then
        'ShellExecute Me.hwnd, "open", szQSUrl, "", "", 1
        ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
    'Else
    '    MsgBoxW Translate(513) '"No Internet Connection Available"
    'End If
End Sub

Private Sub cmdN00bLog_Click()
    
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    
    'PictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
    
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
    'PictLogo.Visible = False
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
    cmdScan_Click
End Sub

Private Sub cmdN00bTools_Click()
    'PictLogo.Visible = False
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
    'If cmdConfig.Caption = Translate(18) Then cmdConfig_Click
    
    cmdConfig.Caption = Translate(18)
    cmdConfig_Click
    chkConfigTabs_Click 3
End Sub

Private Sub chkAutoMark_Click()
    Dim sMsg$
    If chkAutoMark.value = 0 Then Exit Sub
    If RegReadHJT("SeenAutoMarkWarning", "0") = "1" Then Exit Sub
    
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
        Exit Sub
    Else
        chkAutoMark.value = Abs(chkAutoMark.value - 1)
    End If
End Sub

Private Sub chkConfigTabs_Click(index As Integer)
    If bSwitchingTabs Then Exit Sub
    frmMain.cmdHidden.SetFocus
    bSwitchingTabs = True
    
    chkConfigTabs(0).value = 0
    chkConfigTabs(1).value = 0
    chkConfigTabs(2).value = 0
    chkConfigTabs(3).value = 0
    chkConfigTabs(index).value = 1
    
    fraConfigTabs(0).Visible = False
    fraConfigTabs(1).Visible = False
    fraConfigTabs(2).Visible = False
    fraConfigTabs(3).Visible = False
    fraConfigTabs(index).Visible = True
    
    If index = 3 Then 'Misc tools
        SubClassScroll True ' mouse scrolling support
    Else
        SubClassScroll False 'unSubClass
    End If
    
    fraHostsMan.Visible = False
    fraUninstMan.Visible = False
    
    bSwitchingTabs = False
End Sub

Private Sub cmdCheckUpdate_Click()
    CheckForUpdate
End Sub

Private Sub cmdConfig_Click()
    Dim i&, iIgnoreNum&, sIgnore$
    On Error GoTo ErrorHandler:
    
    'сперва выйти из фрейма "Help"
    If cmdHelp.Caption = Translate(17) Then cmdHelp_Click
    
    CloseProgressbar
    
    SubClassScroll True
    
    'If cmdConfig.Caption = "Config..." Then
    If cmdConfig.Caption = Translate(18) Then
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        picPaypal.Visible = False
        lstResults.Visible = False
        
        'cmdConfig.Caption = "Back"
        cmdConfig.Caption = Translate(19)
        'cmdHelp.Enabled = False
        cmdSaveDef.Enabled = False
        fraScan.Enabled = False
        cmdScan.Enabled = False
        cmdFix.Enabled = False
        cmdInfo.Enabled = False
        chkShowN00bFrame.value = CLng(RegReadHJT("ShowIntroFrame", "0"))
        
        txtNothing.Visible = False
        
        For i = 0 To 50
            sRegVals(i) = Replace$(sRegVals(i), txtDefStartPage.Text, "$DEFSTARTPAGE")
            sRegVals(i) = Replace$(sRegVals(i), txtDefSearchPage.Text, "$DEFSEARCHPAGE")
            sRegVals(i) = Replace$(sRegVals(i), txtDefSearchAss.Text, "$DEFSEARCHASS")
            sRegVals(i) = Replace$(sRegVals(i), txtDefSearchCust.Text, "$DEFSEARCHCUST")
        Next i
        
        lstIgnore.Clear
        iIgnoreNum = CInt(RegReadHJT("IgnoreNum", "0"))
        If iIgnoreNum > 0 Then
            For i = 1 To iIgnoreNum
                sIgnore = Crypt(RegReadHJT("Ignore" & CStr(i), vbNullString), sProgramVersion, doCrypt:=False)
                If sIgnore <> vbNullString Then
                    lstIgnore.AddItem sIgnore
                Else
                    Exit For
                End If
            Next i
        End If
        ListBackups
        fraConfigTabs(0).Visible = True
        fraConfig.Visible = True
    Else
        
        SaveAllSettings
        
        'If cmdScan.Caption = "Scan" Then
        'If cmdScan.Caption = Translate(11) Then
            lblInfo(0).Visible = False 'msg of main menu
        'Else
            lblInfo(1).Visible = True 'msg of scan window
        'End If
        picPaypal.Visible = True
        lstResults.Visible = True
        
        fraHostsMan.Visible = False
        fraUninstMan.Visible = False
        If chkConfigTabs(3).value = 1 Then fraConfigTabs(3).Visible = True
        'cmdConfig.Caption = "Config..."
        cmdConfig.Caption = Translate(18)
        'cmdHelp.Enabled = True
        cmdSaveDef.Enabled = True
        
        cmdScan.Enabled = True
        cmdFix.Enabled = True
        cmdInfo.Enabled = True
        fraConfig.Visible = False
        fraScan.Enabled = True
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
    DeleteBackup vbNullString
    lstBackups.Clear
End Sub

Private Sub cmdConfigBackupDelete_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    If lstBackups.ListIndex = -1 Then Exit Sub
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

Private Sub cmdConfigBackupRestore_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    If lstBackups.SelCount = 0 Then Exit Sub
    If lstBackups.SelCount = 1 Then
        If MsgBoxW(Translate(86), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If msgboxW("Restore this item?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        If MsgBoxW(Replace$(Translate(87), "[]", lstBackups.SelCount), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If msgboxW("Restore these " & lstBackups.SelCount & " items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    For i = lstBackups.ListCount - 1 To 0 Step -1
        If lstBackups.Selected(i) Then
            RestoreBackup lstBackups.List(i)
            lstBackups.RemoveItem i
        End If
    Next i
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfigBackupRestore_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdConfigIgnoreDelAll_Click()
    On Error GoTo ErrorHandler:
    Dim i&
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

Private Sub cmdConfigIgnoreDelSel_Click()
    On Error GoTo ErrorHandler:
    Dim i&
    For i = 0 To lstIgnore.ListCount - 1
        RegDelHJT "Ignore" & CStr(i + 1)
    Next i
    For i = lstIgnore.ListCount - 1 To 0 Step -1
        If lstIgnore.Selected(i) Then lstIgnore.RemoveItem i
    Next i
    RegSaveHJT "IgnoreNum", lstIgnore.ListCount
    For i = 0 To lstIgnore.ListCount - 1
        RegSaveHJT "Ignore" & CStr(i + 1), Crypt(lstIgnore.List(i), sProgramVersion, doCrypt:=True)
    Next i
    IsOnIgnoreList "", UpdateList:=True
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdConfigIgnoreDelSel_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdFix_Click()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.cmdFix_Click - Begin"
    
    Dim i&, sPrefix$, bO14Fixed As Boolean, pos&
    
    If lstResults.ListCount = 0 Then Exit Sub
    
    If lstResults.SelCount = 0 Then
        If MsgBoxW(Translate(344), vbQuestion + vbYesNo) = vbNo Then
        'If msgboxW("Nothing selected! Continue?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        Else
            lstResults.Clear
            cmdFix.FontBold = False
            cmdFix.Enabled = False
            'cmdScan.Caption = "Scan"
            cmdScan.Caption = Translate(11)
            cmdScan.FontBold = True
            Exit Sub
        End If
    End If
    
    If lstResults.ListCount = lstResults.SelCount Then
        If MsgBoxW(Translate(345), vbExclamation + vbYesNo) = vbNo Then Exit Sub
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
           IIf(bMakeBackup, ".", ", " & Translate(347)), vbQuestion + vbYesNo, "HiJackThis") = vbNo Then Exit Sub
'        If msgboxW("Fix " & lstResults.SelCount & _
'         " selected items? This will permanently " & _
'         "delete and/or repair what you selected" & _
'         IIf(bMakeBackup, ".", ", unless you enable backups."), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    SetProgressBar lstResults.SelCount + 1
    UpdateProgressBar "Backup"
    
    If bMakeBackup Then
        For i = 0 To lstResults.ListCount - 1
            If lstResults.Selected(i) Then
                MakeBackup lstResults.List(i)
            End If
        Next i
    End If
    
    'shpBackground.Tag = lstResults.SelCount
    'shpProgress.Tag = "0"
    
    'shpProgress.Width = 15
    'picPaypal.Visible = False
    bRebootNeeded = False
    bUpdatePolicyNeeded = False
    bShownBHOWarning = False
    bShownToolbarWarning = False
    bSeenHostsFileAccessDeniedWarning = False
    
    Call GetProcesses_Zw(gProcess)
    
    For i = 0 To lstResults.ListCount - 1
        If lstResults.Selected(i) = True Then
            lstResults.ListIndex = i
            
            sPrefix = ""
            pos = InStr(lstResults.List(i), "-")
            If pos <> 0 Then
                sPrefix = Trim(Left$(lstResults.List(i), pos - 1))
            End If
            
            UpdateProgressBar sPrefix
            
            AppendErrorLogCustom "Fixing: " & lstResults.List(i)
            
            Select Case sPrefix ' RTrim$(Left$(lstResults.List(i), 3))
                Case "R0", "R1", "R2": FixRegItem lstResults.List(i)
                Case "R3":             FixRegistry3Item lstResults.List(i)
                Case "F0", "F1":       FixFileItem lstResults.List(i)
                Case "F2", "F3":       FixFileItem lstResults.List(i)
                'Case "N1", "N2", "N3", "N4": FixNetscapeMozilla lstResults.List(i)
                Case "O1":             FixO1Item lstResults.List(i)
                Case "O2":             FixO2Item lstResults.List(i)
                Case "O3":             FixO3Item lstResults.List(i)
                Case "O4":             FixO4Item lstResults.List(i)
                Case "O5":             FixO5Item lstResults.List(i)
                Case "O6":             FixO6Item lstResults.List(i)
                Case "O7":             FixO7Item lstResults.List(i)
                Case "O8":             FixO8Item lstResults.List(i)
                Case "O9":             FixO9Item lstResults.List(i)
                Case "O10":            FixLSP
                Case "O11":            FixO11Item lstResults.List(i)
                Case "O12":            FixO12Item lstResults.List(i)
                Case "O13":            FixO13Item lstResults.List(i)
                Case "O14":            If Not bO14Fixed Then FixO14Item lstResults.List(i): bO14Fixed = True 'O14 fix uses only once
                Case "O15":            FixO15Item lstResults.List(i)
                Case "O16":            FixO16Item lstResults.List(i)
                Case "O17":            FixO17Item lstResults.List(i)
                Case "O18":            FixO18Item lstResults.List(i)
                Case "O19":            FixO19Item lstResults.List(i)
                Case "O20":            FixO20Item lstResults.List(i)
                Case "O21":            FixO21Item lstResults.List(i)
                Case "O22":            FixO22Item lstResults.List(i)
                Case "O23":            FixO23Item lstResults.List(i)
                Case "O24":            FixO24Item lstResults.List(i)
                Case "O25":            FixO25Item lstResults.List(i)
                Case Else
                   ' msgboxW "Fixing of " & Rtrim$(left$(lstResults.List(i), 3)) & _
                           " is not implemented yet. Bug me about it at " & _
                           "www.merijn.org/contact.html, because I should have done it.", _
                           vbInformation, "bad coder - no donuts"
                           
                    'Fixing of [] is not implemented yet."
                    MsgBoxW Replace$(Translate(348), "[]", sPrefix, vbInformation)
            End Select
            
        End If
    Next i
    UpdateProgressBar "Finish"
    lstResults.Clear
    cmdFix.Enabled = False
    cmdFix.FontBold = False
    cmdScan.Caption = Translate(11)
    'cmdScan.Caption = "Scan"
    cmdScan.FontBold = True
    'lblInfo(0).Visible = True
    'lblInfo(1).Visible = False
    'picPaypal.Visible = True
    On Error Resume Next
    cmdScan.SetFocus
    
    If bRebootNeeded Then RestartSystem: Exit Sub
    If bUpdatePolicyNeeded Then UpdatePolicy
    
    'CloseProgressbar 'leave progressBar visible to ensure the user saw completion of cure
    
    If Not inIDE Then MessageBeep MB_ICONINFORMATION
    Exit Sub
    
    AppendErrorLogCustom "frmMain.cmdFix_Click - End"
    
ErrorHandler:
    ErrorMsg Err, "cmdFix_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdHelp_Click()
    'Back
    If cmdConfig.Caption = Translate(19) Then
        cmdConfig_Click
    End If

    'If cmdHelp.Caption = "Info..." Then
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
    Else
        lblInfo(0).Visible = True
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
        GetInfo lstResults.List(lstResults.ListIndex)
    ElseIf txtHelp.Visible Then
        GetInfo LTrim$(txtHelp.SelText)
    End If
End Sub

Private Sub cmdSaveDef_Click()
    On Error GoTo ErrorHandler:
    If lstResults.SelCount = 0 Then Exit Sub
    If bConfirm Then
        If MsgBoxW(Translate(25), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'        If msgboxW("This will set HiJackThis to ignore the " & _
'                  "checked items, unless they change. Cont" & _
'                  "inue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    Dim i&, J&
    i = CInt(RegReadHJT("IgnoreNum", "0"))
    RegSaveHJT "IgnoreNum", CStr(i + lstResults.SelCount)
    J = i + 1
    For i = 0 To lstResults.ListCount - 1
        If lstResults.Selected(i) Then
            RegSaveHJT "Ignore" & CStr(J), Crypt(lstResults.List(i), sProgramVersion, doCrypt:=True)
            J = J + 1
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
        On Error Resume Next
        cmdScan.SetFocus
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdSaveDef_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdScan_Click()
    On Error GoTo ErrorHandler:
    Static isRan As Boolean
    Dim idx&
    
    AppendErrorLogCustom "frmMain.cmdScan_Click - Begin"
    
    If isRan Then
        Exit Sub
    Else
        isRan = True
    End If
    cmdScan.Enabled = False
    
    idx = 0
    
    'If cmdScan.Caption = "Scan" Then
    
    If cmdScan.Caption = Translate(11) Then
    
        ' Erase main W array of scan results
        ReInitScanResults
        
        idx = 1
        
        'labels off -> moved to SetProgressBar
        'lblInfo(0).Visible = False
        'lblInfo(1).Visible = False
        'shpBackground.Visible = True
        'shpProgress.Visible = True
        
        'picPaypal.Visible = False
        
        lblMD5.Visible = True
        'If bMD5 = False Then lblStatus.Visible = True
        
        cmdAnalyze.Enabled = False
    
        idx = 2
    
        ' Clear Error Log
        ErrReport = ""
    
        ' *******************************************************************

        StartScan '<<<<<<<-------- Main scan routine
        
        If txtNothing.Visible Or Not bAutoLog Then UpdateProgressBar "Finish"
        
        idx = 3
        
        SortSectionsOfResultList
        
        idx = 4
        
        'add the horizontal scrollbar if needed
        AddHorizontalScrollBarToResults lstResults
        
        idx = 5
        
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
        Else
            bAutoLog = False
        End If
        
        idx = 6
        
        If bAutoLog Then DoEvents: SaveReport
        bAutoLog = False
        
        cmdScan.Enabled = True
        cmdAnalyze.Enabled = True
        
        CloseProgressbar
        
    Else    'Caption = Save...

        Call SaveReport
        
        UpdateProgressBar "Finish"

        'cmdScan.Caption = "Scan"
        cmdScan.Caption = Translate(11)
        
        cmdScan.Enabled = True
    End If
    
    'focus on 1-st element of list
    Me.lstResults.ListIndex = -1
    If Me.lstResults.Visible Then Me.lstResults.SetFocus
    isRan = False
    
    AppendErrorLogCustom "frmMain.cmdScan_Click - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdScan_Click", "(" & cmdScan.Caption & ")" & " (index = " & idx & ")"
    cmdScan.Enabled = True
    isRan = False
    If inIDE Then Stop: Resume Next
End Sub

Private Sub SaveReport()
    On Error GoTo ErrorHandler:
    Dim ffLog As Long
    Dim idx&

    AppendErrorLogCustom "frmMain.SaveReport - Begin"

    Dim sLogFile$, i&, errN&
        
        idx = 7
        
        If bAutoLog Then
            sLogFile = BuildPath(AppPath(), "HiJackThis.log")
        Else
            bGlobalDontFocusListBox = True
            'sLogFile = CmnDlgSaveFile("Save logfile...", "Log files (*.log)|*.log|All files (*.*)|*.*", "HiJackThis.log")
            sLogFile = CmnDlgSaveFile(Translate(1001), Translate(1002) & " (*.log)|*.log|" & Translate(1003) & " (*.*)|*.*", "HiJackThis.log")
            bGlobalDontFocusListBox = False
        End If
        
        idx = 8
        
        If 0 <> Len(sLogFile) Then
            
            idx = 11
            
            Dim b() As Byte
            
            b = CreateLogFile()
            
            idx = 12
            
            If FileExists(sLogFile) Then DeleteFileWEx (StrPtr(sLogFile))
            
            idx = 13
            
            If Not OpenW(sLogFile, FOR_OVERWRITE_CREATE, ffLog) Then

                If Not bAutoLogSilent Then
                    'try another name

                    sLogFile = sLogFile & "_2.log"

                    Call OpenW(sLogFile, FOR_OVERWRITE_CREATE, ffLog)

                End If
            End If

            If ffLog <= 0 Then
                If Not bAutoLogSilent Then
'                   msgboxW "Write access was denied to the " & _
'                       "location you specified. Try a " & _
'                       "different location please.", vbExclamation
                    MsgBoxW Translate(26), vbExclamation
                End If
                Exit Sub
            End If

            PutW ffLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=False
            
            CloseW ffLog
            
            idx = 14
            
            If (Not bAutoLogSilent) Or inIDE Then
                If ShellExecute(Me.hWnd, StrPtr("open"), StrPtr(sLogFile), 0&, 0&, 1) <= 32 Then
                    'system doesn't know what .log is
                    If FileExists(sWinDir & "\notepad.exe") Then
                        ShellExecute Me.hWnd, StrPtr("open"), StrPtr(sWinDir & "\notepad.exe"), StrPtr(sLogFile), 0&, 1
                    Else
                        If FileExists(sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\notepad.exe") Then
                            ShellExecute Me.hWnd, StrPtr("open"), StrPtr(sWinDir & IIf(bIsWinNT, "\sytem32", "\system") & "\notepad.exe"), StrPtr(sLogFile), 0&, 1
                        Else
                            'MsgBoxW Replace$(Translate(27), "[]", sLogFile), vbInformation
    '                        msgboxW "The logfile has been saved to " & sLogFile & "." & vbCrLf & _
    '                               "You can open it in a text editor like Notepad.", vbInformation
                        
                            Dim hRet As Long
                            Dim pIDL As Long
            
                            If OSver.MajorMinor >= 5.1 Then
            
                                pIDL = ILCreateFromPath(StrPtr(sLogFile))
    
                                If pIDL <> 0 Then
                                    hRet = SHOpenFolderAndSelectItems(pIDL, 0, 0, 0)
    
                                    ILFree pIDL
                                End If
                            End If
            
                            If pIDL = 0 Or hRet <> S_OK Then
                                'альтернативно
                                Shell "explorer.exe /select," & """" & sLogFile & """", vbNormalFocus   ' открываем папку с логом
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    AppendErrorLogCustom "frmMain.SaveReport - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdScan_SaveReport", "(" & cmdScan.Caption & ")" & " (index = " & idx & ")"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdStartupList_Click()
    Dim sPathComCtl As String, Success As Boolean
    sPathComCtl = BuildPath(AppPath(), "MSComCtl.ocx")
    If Not FileExists(sPathComCtl) Then
        If UnpackResource(102, sPathComCtl) Then Success = True
    Else
        Success = True
    End If
    If Success Then
        frmStartupList2.Show
    Else
        MsgBoxW "Cannot unpack " & sPathComCtl, vbCritical
    End If
End Sub

Private Sub cmdUninstall_Click()
    On Error GoTo ErrorHandler:
    If MsgBoxW(Translate(153), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'    If msgboxW("This will remove HiJackThis' settings from the Registry " & _
'              "and exit. Note that you will have to delete the " & _
'              "HiJackThis.exe file manually." & vbCrLf & vbCrLf & _
'              "Continue with uninstall?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe"
    RegDelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", False
    RegDelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", True
    If Not RegKeyHasSubKeys(HKEY_LOCAL_MACHINE, "Software\TrendMicro", False) Then
        RegDelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro", False
    End If
    If Not RegKeyHasSubKeys(HKEY_LOCAL_MACHINE, "Software\TrendMicro", True) Then
        RegDelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro", True
    End If
    RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan", False
    RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan", True
    CreateUninstallKey False
    DeleteBackup vbNullString
    Close
    g_UninstallState = True
    Unload Me
    End
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdUninstall_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
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
    lstResults.Height = Me.ScaleHeight - 2415
    fraScan.Top = Me.ScaleHeight - 1530
    fraOther.Top = Me.ScaleHeight - 1530
    fraSubmit.Top = Me.ScaleHeight - 1530
    txtNothing.Top = lstResults.Top + (lstResults.Height - txtNothing.Height) / 2
    ' - help -
    fraHelp.Height = Me.ScaleHeight - 1755
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
    lstBackups.Height = Me.ScaleHeight - 3615
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
    On Error Resume Next
    
    AppendErrorLogCustom "frmMain.LoadSettings - Begin"
    
    chkAutoMark.value = CInt(RegReadHJT("AutoSelect", "0"))
    chkConfirm.value = CInt(RegReadHJT("Confirm", "1"))
    chkBackup.value = CInt(RegReadHJT("MakeBackup", "1"))
    chkIgnoreSafe.value = CInt(RegReadHJT("IgnoreSafe", "1"))
    chkLogProcesses.value = CInt(RegReadHJT("LogProcesses", "1"))
    chkShowN00bFrame.value = CInt(RegReadHJT("ShowIntroFrame", "1"))
    chkShowN00b.value = CInt(RegReadHJT("ShowIntroFrame", "1"))
    chkSkipErrorMsg.value = CInt(RegReadHJT("SkipErrorMsg", "0"))
    chkConfigMinimizeToTray.value = CInt(RegReadHJT("MinToTray", "0"))
    chkIgnoreMicrosoft.value = CInt(RegReadHJT("HideMicrosoft", "1"))
    chkIgnoreAll.value = CInt(RegReadHJT("IgnoreAllWhiteList", "0"))
    chkDoMD5.value = CInt(RegReadHJT("CalcMD5", "0"))
    chkAdvLogEnvVars.value = CInt(RegReadHJT("LogEnvVars", "0"))
    
    bHideMicrosoft = chkIgnoreMicrosoft.value
    bIgnoreAllWhitelists = chkIgnoreAll.value
    bMD5 = chkDoMD5.value
    bLogEnvVars = chkAdvLogEnvVars.value
    
    If RegValueExists(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan") Then
        chkConfigStartupScan.value = 1
    Else
        chkConfigStartupScan.value = 0
    End If
    
    If bIgnoreAllWhitelists Then chkIgnoreSafe.value = 0
    
    Dim sData$, LastVerLaunched$, isEncodedVer As Boolean
    
    LastVerLaunched = RegReadHJT("version", "")
    If ConvertVersionToNumber(LastVerLaunched) < ConvertVersionToNumber("2.6.1.21") Then isEncodedVer = True
    
    sData = RegReadHJT("DefStartPage", "")
    'StrBeginWith(sData, "http") - необходим для обратной совместимости со старыми версиями HJT, чтобы не было конфликта криптографического модуля
    If sData = "" Or StrBeginWith(sData, "http") Or isEncodedVer Then
        g_DEFSTARTPAGE = "http://www.msn.com/"
    Else
        g_DEFSTARTPAGE = Crypt(sData, sProgramVersion, doCrypt:=False)
    End If
    txtDefStartPage.Text = g_DEFSTARTPAGE

    sData = RegReadHJT("DefSearchPage", "")
    If sData = "" Or StrBeginWith(sData, "http") Or isEncodedVer Then
        g_DEFSEARCHPAGE = "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch"
    Else
        g_DEFSEARCHPAGE = Crypt(sData, sProgramVersion, doCrypt:=False)
    End If
    txtDefSearchPage.Text = g_DEFSEARCHPAGE
    
    sData = RegReadHJT("DefSearchAss", "")
    If sData = "" Or StrBeginWith(sData, "http") Or isEncodedVer Then
        g_DEFSEARCHASS = "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm"
    Else
        g_DEFSEARCHASS = Crypt(sData, sProgramVersion, doCrypt:=False)
    End If
    txtDefSearchAss.Text = g_DEFSEARCHASS
    
    sData = RegReadHJT("DefSearchCust", "")
    If sData = "" Or StrBeginWith(sData, "http") Or isEncodedVer Then
        g_DEFSEARCHCUST = "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm"
    Else
        g_DEFSEARCHCUST = Crypt(sData, sProgramVersion, doCrypt:=False)
    End If
    txtDefSearchCust.Text = g_DEFSEARCHCUST
    
    bAutoSelect = IIf(chkAutoMark.value = 1, True, False)
    bConfirm = IIf(chkConfirm.value = 1, True, False)
    bMakeBackup = IIf(chkBackup.value = 1, True, False)
    bIgnoreSafe = IIf(chkIgnoreSafe.value = 1, True, False)
    bLogProcesses = IIf(chkLogProcesses.value = 1, True, False)
    bSkipErrorMsg = IIf(chkSkipErrorMsg.value = 1, True, False)
    bMinToTray = IIf(chkConfigMinimizeToTray.value = 1, True, False)
    
    Dim i&
    On Error GoTo ErrorHandler:
    
    For i = 0 To UBound(sRegVals)
        If sRegVals(i) = vbNullString Then Exit For
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSTARTPAGE", txtDefStartPage.Text)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHPAGE", txtDefSearchPage.Text)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHASS", txtDefSearchAss.Text)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHCUST", txtDefSearchCust.Text)
        
        'sRegVals(i) = replace$(sRegVals(i), "$WINSYSDIR", sWinSysDir)
        'sRegVals(i) = replace$(sRegVals(i), "$WINDIR", sWinDir)
        sRegVals(i) = EnvironW(sRegVals(i))
    Next i
    For i = 0 To UBound(sFileVals)
        If sFileVals(i) = vbNullString Then Exit For
        'sFileVals(i) = replace$(sFileVals(i), "$WINDIR", sWinDir)
        sFileVals(i) = EnvironW(sFileVals(i))
    Next i
        
    If Not RegKeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis") Then
        'first use, show moron warning
        'msgboxW Translate(3)
'        msgboxW "Warning!" & vbCrLf & vbCrLf & _
'               "Since HiJackThis targets browser hijacking methods " & _
'               "instead of actual browser hijackers, entries may " & _
'               "appear in the scan list that are not hijackers. " & _
'               "Be careful what you delete, some system utilities " & _
'               "can cause problems if disabled." & vbCrLf & _
'               "For best results, ask spyware experts for help and " & _
'               "show them your scan log. They will advise you what " & _
'               "to fix and what to keep." & vbCrLf & vbCrLf & _
'               "Some adware-supported programs may cease to " & _
'               "function if the associated adware is removed.", vbExclamation
        
        RegCreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis"
    Else
        If RegGetString(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis", "WinWidth") = vbNullString Then
            'clear all previous settings
            RegDelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis"
            RegCreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HiJackThis"
        End If
    End If
    ''If Not RegKeyExists(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe") Then
    ''    RegCreateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe"
    ''    RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe", "", AppPath() & IIf(right$(AppPath(), 1) = "\", "", "\") & "HiJackThis.exe"
    ''    RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe", "Path", AppPath()
    ''Else
        
        '2.0.7 - Routine has been disabled
        'If RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe", vbNullString) _
        '   <> AppPath() & IIf(Right$(AppPath(), 1) = "\", vbNullString, "\") & "HiJackThis.exe" Then
        '    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe", vbNullString, AppPath() & IIf(Right$(AppPath(), 1) = "\", vbNullString, "\") & "HiJackThis.exe"
        'End If
        'If RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe", "Path") _
        '   <> AppPath() Then
        '    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HiJackThis.exe", "Path", AppPath()
        'End If
    ''End If
    ''CreateUninstallKey True
    
    AppendErrorLogCustom "frmMain.LoadSettings - End"
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "frmMain_LoadSettings"
    If inIDE Then Stop: Resume Next
End Sub

Private Function CreateLogFile() As String
    Dim sLog$, i&, sProcessList$
    Dim hSnap&, uProcess As PROCESSENTRY32, sDummy$ '9x
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024) 'NT
    Dim Col As New Collection, Cnt&, sPrefix$, sPrefixLast$
    
    On Error GoTo MakeLog:
    
    AppendErrorLogCustom "frmMain.CreateLogFile - Begin"
    
    If Not bLogProcesses Then GoTo MakeLog
    
    'UpdateProgressBar "ProcList"
    'DoEvents
    
    If Not bIsWinNT Then
        hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = INVALID_HANDLE_VALUE Then
            'sProcessList = "(Unable to list running processes)" & vbCrLf
            sProcessList = "(" & Translate(28) & ")" & vbCrLf
            GoTo MakeLog
        End If
        
        uProcess.dwSize = Len(uProcess)
        If Process32First(hSnap, uProcess) = 0 Then
            'sProcessList = "(Unable to list running processes)" & vbCrLf
            sProcessList = "(" & Translate(28) & ")" & vbCrLf
            CloseHandle hSnap: hSnap = 0
            GoTo MakeLog
        End If
        
        On Error Resume Next
        
        Do
            sDummy = TrimNull(uProcess.szExeFile)
            
            Col.Add 1&, sDummy              ' item - count, key - name of process

            If Err.Number <> 0& Then        ' if the same process
                Cnt = Col.Item(sDummy)      ' replacing item of collection
                Col.Remove (sDummy)
                Col.Add Cnt + 1&, sDummy    ' increase count of identical processes
            End If
            
        Loop Until Process32Next(hSnap, uProcess) = 0
        
        On Error GoTo MakeLog:
        
        CloseHandle hSnap: hSnap = 0
        sProcessList = sProcessList & vbCrLf
    Else
        
        Dim Process() As MY_PROC_ENTRY
        
        lNumProcesses = GetProcesses_Zw(Process)
        
        On Error Resume Next
        
        If lNumProcesses Then
        
            For i = 0 To UBound(Process)
        
                sProcessName = Process(i).Path
                
                'Debug.Print Process(i).Path
                
                If Len(Process(i).Path) = 0 Then
                    If Not ((StrComp(Process(i).Name, "System Idle Process", 1) = 0 And Process(i).PID = 0) _
                        Or (StrComp(Process(i).Name, "System", 1) = 0 And Process(i).PID = 4) _
                        Or (StrComp(Process(i).Name, "Memory Compression", 1) = 0)) Then
                          sProcessName = Process(i).Name '& " (cannot get Process Path)"
                    End If
                End If
                
                If Len(sProcessName) <> 0 Then
                    Col.Add 1&, sProcessName              ' item - count, key - name of process

                    If Err.Number <> 0& Then              ' if the same process
                        Cnt = Col.Item(sProcessName)      ' replacing item of collection
                        Col.Remove (sProcessName)
                        Col.Add Cnt + 1&, sProcessName    ' increase count of identical processes
                        Err.Clear
                    End If
                End If
            Next
        End If
        
    End If
    
    'sProcessList = "Running processes:" & vbCrLf
    sProcessList = Translate(29) & ":" & vbCrLf
    
    sProcessList = sProcessList & "Number | Path" & vbCrLf
    
    ' Sort using positions array method (Key - Process Path).
    ReDim ProcLog(Col.Count) As MY_PROC_LOG
    ReDim apos(Col.Count), names(Col.Count)
    
    Dim SignResult  As SignResult_TYPE
    
    For i = 1& To Col.Count
        With ProcLog(i)
            .ProcName = GetCollectionKey(i, Col)
            .Number = Col(i)
            
' I temporarily disable EDS checking
'            SignVerify .ProcName, 0&, SignResult
'            If SignResult.isLegit Then
'                .EDS_issued = SignResult.SubjectName
'            End If
'
'            If Not bIgnoreAllWhitelists Then
'                UpdateProgressBar "ProcList", .ProcName
'                .IsMicrosoft = (IsMicrosoftCertHash(SignResult.HashRootCert) And SignResult.isLegit)  'validate EDS
'            End If
            
            names(i) = IIf(.IsMicrosoft, "(Microsoft) ", IIf(.EDS_issued <> "", "(" & .EDS_issued & ") ", " (not signed)")) & .ProcName
            apos(i) = i
        End With
    Next
    QuickSortSpecial names, apos, 0, Col.Count
    
    For i = 1& To UBound(apos)
        With ProcLog(apos(i))
            'sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & IIf(.IsMicrosoft, "(Microsoft) ", "") & .ProcName & vbCrLf
            'sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & names(i) & vbCrLf
            If .IsMicrosoft Then
                sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & "(Microsoft) " & .ProcName & vbCrLf
            Else
                'sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & .ProcName & IIf(.EDS_issued <> "", " (" & .EDS_issued & ")", " (not signed)") & vbCrLf
                sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & .ProcName & vbCrLf
            End If
        End With
    Next
    
    sProcessList = sProcessList & vbCrLf
    
    '------------------------------
MakeLog:

    If Err.Number Then
        sProcessList = "(" & Translate(28) & " (error#" & Err.Number & "))" & vbCrLf
        If Not bAutoLogSilent Then MsgBoxW Err.Description
    End If
    
    On Error GoTo ErrorHandler:
    
    'UpdateProgressBar "Finish"
    'DoEvents
    
    sLog = ChrW$(-257) & "Logfile of " & AppVer & vbCrLf & vbCrLf ' + BOM UTF-16 LE
    
    Dim TimeCreated As String
    
    TimeCreated = Right$("0" & Day(Now), 2) & "." & Right$("0" & Month(Now), 2) & "." & Year(Now) & " - " & _
        Right$("0" & Hour(Now), 2) & ":" & Right$("0" & Minute(Now), 2)
    
    sLog = sLog & "Platform:  " & OSInfo.Bitness & " " & OSInfo.OSName & " (" & OSInfo.Edition & "), " & _
            OSInfo.Major & "." & OSInfo.Minor & "." & OSInfo.Build & IIf(OSInfo.ReleaseId <> 0, " (ReleaseId: " & OSInfo.ReleaseId & ")", "") & ", " & _
            "Service Pack: " & OSInfo.SPVer & "" & IIf(OSInfo.IsSafeBoot, " (Safe Boot)", "") & vbCrLf '& vbTab & vbTab & _
            '"(SM=" & OSInfo.SuiteMask & ", PT=" & OSInfo.ProductType & ")" & vbCrLf
    sLog = sLog & "Time:      " & TimeCreated & vbCrLf
    sLog = sLog & "Language:  " & "OS: " & OSInfo.LangSystemNameFull & " (" & "0x" & Hex(OSInfo.LangSystemCode) & "). " & _
            "Display: " & OSInfo.LangDisplayNameFull & " (" & "0x" & Hex(OSInfo.LangDisplayCode) & "). " & _
            "Non-Unicode: " & OSInfo.LangNonUnicodeNameFull & " (" & "0x" & Hex(OSInfo.LangNonUnicodeCode) & ")" & vbCrLf
    
    If OSInfo.MajorMinor >= 6 Then
        sLog = sLog & "Elevated:  " & IIf(OSInfo.IsElevated, "Yes", "No") & vbCrLf  '& vbTab & "IL: " & Osinfo.GetIntegrityLevel & vbCrLf
   End If
    
    sLog = sLog & "Ran by:    " & GetUser() & vbTab & "(group: " & OSInfo.UserType & ") on " & GetComputer() & vbCrLf & vbCrLf
    
    
    Dim tmp$
    With BROWSERS   'MY_BROWSERS (look at modUtils.GetBrowsersInfo())
        tmp = .Opera.Version
        If Len(tmp) Then sLog = sLog & "Opera:   " & tmp & vbCrLf
        tmp = .Chrome.Version
        If Len(tmp) Then sLog = sLog & "Chrome:  " & tmp & vbCrLf
        tmp = .Firefox.Version
        If Len(tmp) Then sLog = sLog & "Firefox: " & tmp & vbCrLf
        tmp = .Edge.Version
        If Len(tmp) Then sLog = sLog & "Edge:    " & tmp & vbCrLf
        tmp = .IE.Version
        If Len(tmp) Then sLog = sLog & "Internet Explorer: " & tmp & vbCrLf
    End With
   
    sLog = sLog & vbCrLf & "Boot mode: " & OSver.BootMode & vbCrLf
    
    '// TODO: improve it (Get environment block)
    If bLogEnvVars Then
        sLog = sLog & "Windows folder: " & sWinDir & vbCrLf & _
                      "System folder: " & sWinSysDir & vbCrLf & _
                      "Hosts file: " & sHostsFile & vbCrLf
    End If
    
    sLog = sLog & vbCrLf & sProcessList
    
    ' Adding empty lines beetween sections (cancelled)
    For i = 0 To UBound(HitSorted)
        'sPrefix = RTrim(Splitsafe(HitSorted(i), "-")(0))
        'If sPrefixLast <> "" And sPrefixLast <> sPrefix Then sLog = sLog & vbCrLf
        'sPrefixLast = sPrefix
        sLog = sLog & HitSorted(i) & vbCrLf
    Next
    
    Dim IgnoreCnt&
    IgnoreCnt = RegReadHJT("IgnoreNum", "0")
    If IgnoreCnt <> 0 Then
        sLog = sLog & vbCrLf & vbCrLf & "Warning: Ignore list contains " & IgnoreCnt & " items." & vbCrLf
    End If
    
    'Append by Error Log
    If 0 <> Len(ErrReport) Then
        sLog = sLog & vbCrLf & vbCrLf & "Debug information:" & vbCrLf & ErrReport & vbCrLf
        '& vbCrLf & "CmdLine: " & AppPath(True) & " " & Command()
    End If
    
    If 0 <> ErrLogCustomText.length Then
        sLog = sLog & vbCrLf & vbCrLf & "Trace information:" & vbCrLf & ErrLogCustomText.ToString & vbCrLf
    End If
    
    If bAutoLog Then Perf.EndTime = GetTickCount()
    sLog = sLog & vbCrLf & "--" & vbCrLf & "End of file - " & "Time spent: " & ((Perf.EndTime - Perf.StartTime) \ 1000) & " sec. - "
    
    Dim Size_1 As Long
    Dim Size_2 As Long
    Dim Size_3 As Long
    
    Size_1 = 2& * (Len(sLog) + Len(" bytes, CRC32: FFFFFFFF. Sign:   "))  'Вычисление размера лога (в байтах)
    Size_2 = Size_1 + 2& * Len(CStr(Size_1))                              'с учетом самого числа "кол-во байт"
    Size_3 = Size_2 - 2& * Len(CStr(Size_1)) + 2& * Len(CStr(Size_2))     'пересчет, если число байт увеличилось на 1 разряд
    
    sLog = sLog & CStr(Size_3) & " bytes, CRC32: FFFFFFFF. Sign: "
    
    Dim ForwCRC As Long
    Dim b()     As Byte
    
    b() = sLog                                                          'считаем CRC лога
    ForwCRC = CalcArrayCRCLong(b()) Xor -1
    
    Dim CorrBytes$: CorrBytes = RecoverCRC(ForwCRC, &HFFFFFFFF)         'считаем байты корректировки
    
    ReDim Preserve b(UBound(b) + 4)                                     'добавляем их в конец массива
    b(UBound(b) - 3) = Asc(Mid$(CorrBytes, 1, 1))
    b(UBound(b) - 2) = Asc(Mid$(CorrBytes, 2, 1))
    b(UBound(b) - 1) = Asc(Mid$(CorrBytes, 3, 1))
    b(UBound(b) - 0) = Asc(Mid$(CorrBytes, 4, 1))
    
    CreateLogFile = b()
    
    If hSnap Then CloseHandle hSnap
    If hProc Then CloseHandle hProc
    
    AppendErrorLogCustom "frmMain.CreateLogFile - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "frmMain_CreateLogFile"
    If inIDE Then Stop: Resume Next
End Function

Private Sub SortSectionsOfResultList()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.SortSectionsOfResultList - Begin"
    
    ' Special sort procedure
    ' ---------------------------------
    Dim Hit() As String
    Dim HitSection() As String
    Dim SectSorted() As String
    Dim SectNames() As String
    Dim SectName As String
    Dim nHit As Long
    Dim nSect As Long
    Dim nItemsSect As Long
    Dim nItemsHit As Long
    Dim pos As Long
    Dim i As Long
    
    If lstResults.ListCount = 0 Then
        ReDim HitSorted(0): HitSorted(0) = ""
        Exit Sub
    End If
    
    ReDim Hit(lstResults.ListCount - 1)
    ReDim HitSorted(UBound(Hit))        'var. is global (frmMain module) !
    ReDim SectNames(41)
    
    SectNames(0) = "R0"
    SectNames(1) = "R1"
    SectNames(2) = "R2"
    SectNames(3) = "R3"
    SectNames(4) = "F0"
    SectNames(5) = "F1"
    SectNames(6) = "F2"
    SectNames(7) = "F3"
    SectNames(8) = "N1"
    SectNames(9) = "N2"
    SectNames(10) = "N3"
    SectNames(11) = "N4"
    For i = 1 To 30
        SectNames(11 + i) = "O" & i
    Next
    
    For i = 0 To lstResults.ListCount - 1
        Hit(i) = lstResults.List(i)
    Next i
    
    nItemsHit = 0
    
    For nSect = 0 To UBound(SectNames)
        nItemsSect = 0
        For nHit = 0 To UBound(Hit)
            If 0 <> Len(Hit(nHit)) Then
                pos = InStr(Hit(nHit), "-")
                If pos = 0 Then
                    If Not bAutoLog Then
                        MsgBoxW "Warning! Wrong format of hit line. Must include dash after the name of the section!"
                    End If
                Else
                    SectName = Trim$(Left$(Hit(nHit), pos - 1))
                    ' Сортирую посекционно
                    If SectName = SectNames(nSect) Then
                        ' Создаю временный массив этой секции для сортировки
                        nItemsSect = nItemsSect + 1
                        ReDim Preserve SectSorted(nItemsSect)
                        SectSorted(nItemsSect) = Hit(nHit)
                        Hit(nHit) = vbNullString
                    End If
                End If
            End If
        Next
        ' Сборка секции завершена.
        If 0 <> nItemsSect Then
            ' Начало сортировки секции
            ' O1 не сортируем (hosts)
            If SectNames(nSect) <> "O1" Then
                QuickSort SectSorted, 0, UBound(SectSorted)
            End If
            For i = 0 To UBound(SectSorted)
                If 0 <> Len(SectSorted(i)) Then
                    'Переносим отсортированную секцию в общий массив
                    HitSorted(nItemsHit) = SectSorted(i)
                    nItemsHit = nItemsHit + 1
                End If
            Next
        End If
    Next
    ' Проверяем, не осталось ли неотсортированных элементов
    For nHit = 0 To UBound(Hit)
        If 0 <> Len(Hit(nHit)) Then
            HitSorted(nItemsHit) = Hit(nHit)
            nItemsHit = nItemsHit + 1
        End If
    Next
    ' Rearrange listbox data accorting to sorted list of sections
    lstResults.Clear
    For i = 0 To UBound(HitSorted)
        lstResults.AddItem HitSorted(i)
    Next
    ' --------------------------------- Sorting complete
    
    Perf.EndTime = GetTickCount()
    
    AppendErrorLogCustom "frmMain.SortSectionsOfResultList - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_SortSectionsOfResultList"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub mnuToolsDigiSign_Click()
    frmCheckDigiSign.Show vbModeless
End Sub

Private Sub mnuToolsUnlockRegKey_Click()
    frmUnlockRegKey.Show vbModeless
End Sub

Private Sub vscMiscTools_Change()
    'lToolsHeight = 2200 ' lower this value if you would like more space inside scroll of last config tab
    fraMiscToolsScroll.Top = -vscMiscTools.value * (fraMiscToolsScroll.Height - (fraConfigTabs(3).Height + lToolsHeight)) / 100
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
    
    sFile = DirW$(BuildPath(AppPath(), "*.lng"), vbFile)
    
    Do While Len(sFile)
        If sFile <> "_Lang_EN.lng" And _
            sFile <> "_Lang_RU.lng" Then
                cboN00bLanguage.AddItem sFile
        End If
        sFile = DirW$()
    Loop
    
    sCurLang = RegReadHJT("LanguageFile")   'HJT settings
    If bForceRU Then sCurLang = "Russian"
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
    
    sFile = cboN00bLanguage.List(cboN00bLanguage.ListIndex)
    If Len(sFile) = 0 Then Exit Sub
    If sFile = "English" Then
        'LoadDefaultLanguage
        LoadLanguage &H409, bForceEN
    ElseIf sFile = "Russian" Then
        LoadLanguage &H419, bForceRU
    Else
        LoadLangFile sFile
        ReloadLanguageNative
        ReloadLanguage
    End If
    
    ' Do not save force mode state!
    If Not (bForceRU Or bForceEN) Then RegSaveHJT "LanguageFile", sFile
    
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
        sDisplayName = RegGetString(0&, UninstData(ID).AppRegKey, "DisplayName")
        sUninstString = RegGetString(0&, UninstData(ID).AppRegKey, "UninstallString")
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
            If RegDelKey(0&, UninstData(ID).AppRegKey) Then
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
        sApplication = Space$(MAX_PATH)
        LSet sApplication = ExtractFilename(sUninst) & vbNullChar
        If PathFindOnPath(StrPtr(sApplication), 0&) Or FileExists(sApplication) Then
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

    Dim s$, sName$, sUninst$, i&, ItemID&, ID&
    
    If lstUninstMan.ListCount = 0 Then Exit Sub

    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    UninstRefreshData ItemID, sName, sUninst 'refresh data
    
    If Len(sName) = 0 And Len(sUninst) = 0 Then Exit Sub
    
    'Edit uninstall command
    s = InputBox(Translate(221) & ": '" & sName & ":", Translate(215), IIf(Len(sUninst) > 255, vbNullString, sUninst)) 'InputBox cannot hold more than 255 chars
    's = InputBox("Enter the new uninstall command for this program, '" & txtUninstManName.Text & ":", "Edit uninstall command", txtUninstManCmd.Text)
    
    If StrPtr(s) <> 0 And s <> sUninst And Len(s) <> 0 Then
        If RegSetStringVal(0&, UninstData(ID).AppRegKey, "UninstallString", s) Then
            MsgBoxW Translate(222), vbInformation
            'msgboxW "New uninstall string saved!", vbInformation
            txtUninstManCmd.Text = s
            UninstData(ID).UninstString = s
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

    Dim sItems$(), sName$, sUninst$, i&, J&, Cnt&
    
    lstUninstMan.Clear
    Erase UninstData
    Cnt = -1
    
    'lstUninstMan.Sorted must be False ' Do not enable this kind of sorting at all!!! Otherwise, virus will eat your computer :)
    
    For J = 0 To UBound(sKeyUninstall)
        sItems = Split(RegEnumSubKeys(0&, sKeyUninstall(J)), "|")
        If UBound(sItems) <> -1 Then
            For i = 0 To UBound(sItems)
                sName = RegGetString(0&, sKeyUninstall(J) & "\" & sItems(i), "DisplayName")
                sUninst = RegGetString(0&, sKeyUninstall(J) & "\" & sItems(i), "UninstallString")
                
                If Len(sName) <> 0 And Len(sUninst) <> 0 Then
                    Cnt = Cnt + 1
                    ReDim Preserve UninstData(Cnt)
                    With UninstData(Cnt)                                        'array
                        .DisplayName = sName
                        .UninstString = sUninst
                        .AppRegKey = sKeyUninstall(J) & "\" & sItems(i)
                        .KeyTime = ConvertDateToUSFormat(GetRegKeyTime(0&, .AppRegKey))
                    End With
                End If
            Next i
        End If
    Next J
    If Cnt = -1 Then Exit Sub
    
    'Sorting user type array using bufer array of positions (c) Dragokas
    Dim pos(), names(): ReDim pos(Cnt), names(Cnt)
    For i = 0 To Cnt: pos(i) = i: names(i) = UninstData(i).DisplayName: Next 'key of sort is DisplayName
    QuickSortSpecial names, pos, 0, Cnt
    
    For i = 0 To Cnt
        lstUninstMan.AddItem UninstData(pos(i)).DisplayName
        lstUninstMan.ItemData(i) = pos(i)     'array marker
    Next
    
    If lstUninstMan.ListCount Then lstUninstMan.ListIndex = 0
    lstUninstMan.SetFocus
    
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

    Dim sList$, i&, sUninst$, sFile$, ff%, ID&, b() As Byte, sTmpFile$, lret&, buf As String
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    'sFile = CmnDlgSaveFile("Save Add/Remove Software list to disk...", "Text files (*.txt)|*.txt|All files (*.*)|*.*", "uninstall_list.txt")
    sFile = CmnDlgSaveFile(Translate(225), Translate(226) & " (*.txt)|*.txt|" & Translate(1003) & " (*.*)|*.*", "uninstall_list.txt")
    
    If Len(sFile) = 0 Then Exit Sub
    
    sList = ChrW$(-257)
    
    sList = sList & String$(55, "-") & vbCrLf
    sList = sList & Space$(20) & "Sort by Date" & vbCrLf
    sList = sList & String$(55, "-") & vbCrLf & vbCrLf
    
    ' Make positions array of sorting by .KeyTime property (registry key date).
    Dim Cnt&: Cnt = lstUninstMan.ListCount - 1
    Dim pos(), names(): ReDim pos(Cnt), names(Cnt)
    For i = 0 To Cnt: pos(i) = i: names(i) = UninstData(i).KeyTime: Next
    QuickSortSpecial names, pos, 0, Cnt
    
    For i = Cnt To 0 Step -1 'descending order
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

' Сортировка по Хоару. На вход - массив j(), на выходе массив k() с индексами массива j в отсортированном порядке + отсортированный массив.
' Алгоритм особо отлично подходит для сортировки User type arrays по любому из полей.
Private Sub QuickSortSpecial(J, k, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler:
    Dim i As Long, l As Long, M As String, wsp As String
    i = low: l = high: M = J((i + l) \ 2)
    Do Until i > l: Do While J(i) < M: i = i + 1: Loop: Do While J(l) > M: l = l - 1: Loop
        If (i <= l) Then wsp = J(i): J(i) = J(l): J(l) = wsp: wsp = k(i): k(i) = k(l): k(l) = wsp: i = i + 1: l = l - 1
    Loop
    If low < l Then QuickSortSpecial J, k, low, l
    If i < high Then QuickSortSpecial J, k, i, high
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "QuickSortSpecial"
    If inIDE Then Stop: Resume Next
End Sub


' =========== Menus ==============
'
'

Private Sub mnuFileSettings_Click()     'File -> Settings
    cmdN00bTools_Click
    Call chkConfigTabs_Click(0)
End Sub

Private Sub mnuFileUninstHJT_Click()    'File -> Uninstall HJT
    cmdUninstall_Click
End Sub

Private Sub mnuFileExit_Click()         'File -> Exit
    Unload Me
End Sub

Private Sub mnuToolsADSSpy_Click()      'Tools -> ADS Spy
    'cmdN00bTools_Click
    cmdADSSpy_Click
End Sub

Private Sub mnuToolsDelFileOnReboot_Click()     'Tools -> Delete File -> Delete on reboot...
    cmdDelOnReboot_Click
End Sub

Private Sub mnuToolsUnlockAndDelFile_Click()    'Tools -> Delete File -> Unlock & Delete file
    Dim sFileName$
    'Enter file name:, Unlock & Delete
    sFileName = InputBox(Translate(1952), Translate(1953))
    If StrPtr(sFileName) = 0 Then Exit Sub
    'sFileName = OpenFileDialog("Enter file to unlock access and delete...")
    'sFileName = CmnDlgOpenFile("Enter file to unlock access and delete...", Translate(1003) & " (*.*)|*.*|" & Translate(511) & " (*.dll)|*.dll|" & Translate(512) & " (*.exe)|*.exe")
    'sFileName = CmnDlgOpenFile("Enter file to unlock access and delete...", "All files (*.*)|*.*|DLL libraries (*.dll)|*.dll|Program files (*.exe)|*.exe")
    If 0 = Len(sFileName) Then Exit Sub
    If 0 = DeleteFileWEx(StrPtr(sFileName)) Then
        'Could not delete file. & vbcrlf & Possible, it is locked by another process.
        MsgBoxW Translate(1954)
    Else
        'File: [] deleted successfully.
        MsgBoxW Replace$(Translate(1955), "[]", sFileName)
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

Private Sub mnuHelpManualEnglish_Click()
    Dim szQSUrl$: szQSUrl = "http://www.bleepingcomputer.com/tutorials/how-to-use-hijackthis/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualRussian_Click()
    Dim szQSUrl$: szQSUrl = "http://safezone.cc/threads/25184/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualFrench_Click()
    Dim szQSUrl$: szQSUrl = "http://www.bleepingcomputer.com/tutorials/comment-utiliser-hijackthis/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualGerman_Click()
    Dim szQSUrl$: szQSUrl = "http://www.bleepingcomputer.com/tutorials/wie-hijackthis-genutzt-wird-um/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualSpanish_Click()
    Dim szQSUrl$: szQSUrl = "http://www.bleepingcomputer.com/tutorials/como-usar-hijackthis/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualPortuguese_Click()
    Dim szQSUrl$: szQSUrl = "http://www.linhadefensiva.org/2005/06/hijackthis-completo/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub
Private Sub mnuHelpManualDutch_Click()
    Dim szQSUrl$: szQSUrl = "http://www.bleepingcomputer.com/tutorials/hoe-gebruik-je-hijackthis/"
    ShellExecute Me.hWnd, StrPtr("open"), StrPtr(szQSUrl), 0&, 0&, 1
End Sub

Private Sub mnuHelpUpdate_Click()       'Help -> Download new version
    CheckForUpdate
End Sub

Private Sub mnuHelpAbout_Click()        'Help -> About HJT
    cmdN00bClose_Click
    ' Закрыть окно "Инструменты"
    If cmdConfig.Caption = Translate(19) Then cmdConfig_Click
    If cmdHelp.Caption = Translate(16) Then cmdHelp_Click
    chkHelp(2).value = 1
End Sub

' --------------------------------------

Private Sub txtHelp_LostFocus()
    txtHelpHasFocus = False
End Sub
Private Sub txtHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not txtHelpHasFocus Then txtHelpHasFocus = True: txtHelp.SetFocus
End Sub

Private Sub chkConfigStartupScan_Click()
    '// TODO:
    Dim hHive As Long
        If chkConfigStartupScan.value = 1 Then
            'Sorry! Not implemented yet.
            MsgBoxW Translate(65)
            Exit Sub
            RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan", """" & AppPath(True) & """" & " /startupscan"
        Else
            RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis startup scan"
        End If
End Sub

Sub SaveAllSettings()
    Dim i&, iIgnoreNum&, sIgnore$
    
    AppendErrorLogCustom "frmMain.SaveAllSettings - Begin"
    
    bAutoSelect = IIf(chkAutoMark.value = 1, True, False)
    bConfirm = IIf(chkConfirm.value = 1, True, False)
    bMakeBackup = IIf(chkBackup.value = 1, True, False)
    bIgnoreSafe = IIf(chkIgnoreSafe.value = 1, True, False)
    bLogProcesses = IIf(chkLogProcesses.value = 1, True, False)
    bSkipErrorMsg = IIf(chkSkipErrorMsg.value = 1, True, False)
    bMinToTray = IIf(chkConfigMinimizeToTray.value = 1, True, False)
    
'    RegDelHJT "IgnoreNum"
'    For i = 1 To 99
'        RegDelHJT "Ignore" & CStr(i)
'    Next i
'    RegSaveHJT "IgnoreNum", CStr(lstIgnore.ListCount)
'    For i = 0 To lstIgnore.ListCount - 1
'        RegSaveHJT "Ignore" & CStr(i + 1), Crypt(lstIgnore.List(i), sProgramVersion, doCrypt:=True)
'    Next i
    
    RegSaveHJT "AutoSelect", CStr(Abs(CInt(bAutoSelect)))
    RegSaveHJT "Confirm", CStr(Abs(CInt(bConfirm)))
    RegSaveHJT "MakeBackup", CStr(Abs(CInt(bMakeBackup)))
    RegSaveHJT "IgnoreSafe", CStr(Abs(CInt(bIgnoreSafe)))
    RegSaveHJT "LogProcesses", CStr(Abs(CInt(bLogProcesses)))
    RegSaveHJT "ShowIntroFrame", CStr(chkShowN00bFrame.value)
    RegSaveHJT "SkipErrorMsg", CStr(Abs(CInt(bSkipErrorMsg)))
    RegSaveHJT "MinToTray", CStr(Abs(CInt(bMinToTray)))
    RegSaveHJT "DefStartPage", Crypt(txtDefStartPage.Text, sProgramVersion, doCrypt:=True)
    RegSaveHJT "DefSearchPage", Crypt(txtDefSearchPage.Text, sProgramVersion, doCrypt:=True)
    RegSaveHJT "DefSearchAss", Crypt(txtDefSearchAss.Text, sProgramVersion, doCrypt:=True)
    RegSaveHJT "DefSearchCust", Crypt(txtDefSearchCust.Text, sProgramVersion, doCrypt:=True)
    
    'Update global state
    g_DEFSTARTPAGE = txtDefStartPage.Text
    g_DEFSEARCHPAGE = txtDefSearchPage.Text
    g_DEFSEARCHASS = txtDefSearchAss.Text
    g_DEFSEARCHCUST = txtDefSearchCust.Text
    For i = 0 To UBound(sRegVals)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSTARTPAGE", g_DEFSTARTPAGE)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHPAGE", g_DEFSEARCHPAGE)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHASS", g_DEFSEARCHASS)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHCUST", g_DEFSEARCHCUST)
    Next i
    
    AppendErrorLogCustom "frmMain.SaveAllSettings - End"
End Sub

Private Sub chkIgnoreAll_Click()
    bIgnoreAllWhitelists = chkIgnoreAll.value
    If bIgnoreAllWhitelists Then
        If chkIgnoreMicrosoft.value = 1 Then chkIgnoreMicrosoft.value = 0
    End If
    RegSaveHJT "IgnoreAllWhiteList", Abs(CLng(bIgnoreAllWhitelists))
End Sub

Private Sub chkIgnoreMicrosoft_Click()
    bHideMicrosoft = chkIgnoreMicrosoft.value
    If bHideMicrosoft Then
        If chkIgnoreAll.value = 1 Then chkIgnoreAll.value = 0
    End If
    RegSaveHJT "HideMicrosoft", Abs(CLng(bHideMicrosoft))
End Sub

'Context menu in result list of scan:

Private Sub lstResults_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lstResults.SelCount = 0 Then     'items not checked ?
            mnuResultFix.Visible = False
            mnuResultAddToIgnore.Visible = False
        Else
            mnuResultFix.Visible = True
            mnuResultAddToIgnore.Visible = True
        End If
        If lstResults.ListIndex = -1 Then   'item not selected ?
            mnuResultInfo.Visible = False
            mnuResultSearch.Visible = False
            mnuResultDelim1.Visible = False
        Else
            mnuResultInfo.Visible = True
            mnuResultSearch.Visible = True
            mnuResultDelim1.Visible = True
        End If
        PopupMenu mnuResultList
    End If
End Sub

Private Sub mnuResultFix_Click()          'Fix checked
    cmdFix_Click
End Sub

Private Sub mnuResultInfo_Click()         'Info on selected
    cmdInfo_Click
End Sub

Private Sub mnuResultAddToIgnore_Click()  'Add to ignore list
    cmdSaveDef_Click
End Sub

Private Sub mnuResultSearch_Click()       'Search on Google
    Dim sItem$, sURL$, pos&
    sItem = lstResults.List(lstResults.ListIndex)
    pos = InStr(sItem, ":")
    If pos > 0 Then
        sItem = Mid$(sItem, pos + 1)
    End If
    sURL = "https://www.google.com/?ie=UTF-8#q=" & URLEncode(sItem)
    ShellExecute 0&, StrPtr("open"), StrPtr(sURL), 0&, 0&, vbNormalFocus
End Sub

Private Sub mnuResultReScan_Click()       'ReScan
    cmdScan.Caption = Translate(11)
    cmdScan_Click
End Sub

'test stuff - BUTTON: enum tasks to CSV
Private Sub cmdTaskScheduler_Click()
    Call EnumTasks(True)
End Sub

Private Sub chkConfigMinimizeToTray_Click()
    bMinToTray = chkConfigMinimizeToTray.value
End Sub

Private Sub chkHelp_Click(index As Integer)
    Dim i As Long, J As Long
    Dim sText As String
    Dim sSeparator$
    Dim aSect()
    
    If bSwitchingTabs Then Exit Sub
    frmMain.cmdHidden.SetFocus
    bSwitchingTabs = True
    
    For i = 0 To chkHelp.Count - 1
        If index <> i Then
            chkHelp(i).value = 0
            chkHelp(i).Enabled = True
        Else
            chkHelp(i).Enabled = False
        End If
    Next
    
    Select Case index
    'Sections
    Case 0:
        aSect = Array("R0", "R1", "R2", "R3", "F0", "F1", "F2", "F3", "O1", "O2", "O3", "O4", "O5", "O6", "O7", "O8", "O9", "O10", _
            "O11", "O12", "O13", "O14", "O15", "O16", "O17", "O18", "O19", "O20", "O21", "O22", "O23", "O24", "O25")
        
        sText = Translate(31) & vbCrLf & vbCrLf & Translate(490)
        sSeparator = String$(100, "-")
        
        For i = 0 To UBound(aSect)
            J = 401 + i
            sText = sText & vbCrLf & sSeparator & vbCrLf & FindLine(aSect(i) & " -", Translate(31)) & vbCrLf & sSeparator & vbCrLf & Translate(J) & vbCrLf
        Next
        
        txtHelp.Text = sText
    'Keys
    Case 1: txtHelp.Text = Translate(32)
    'Purpose
    Case 2: txtHelp.Text = Translate(33)
    'History
    Case 3:
        ' Updates
        ' ------------
        ' You can find list of recent updates at:
        txtHelp.Text = Translate(1300) & " http://safezone.cc/threads/24988/"
    End Select
    
    bSwitchingTabs = False
End Sub

Function FindLine(sPartialText As String, sFullText As String) As String
    Dim arr, i&
    arr = Split(sFullText, vbCrLf)
    If IsArrDimmed(arr) Then
        For i = 0 To UBound(arr)
            If InStr(1, arr(i), sPartialText, 1) <> 0 Then FindLine = arr(i): Exit For
        Next
    End If
End Function

Private Sub cmdProcessManager_Click()
    frmProcMan.Show
End Sub

