VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "coolwebsearch"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
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
         Height          =   375
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
   Begin VB.Frame fraScan 
      Caption         =   "Scan && fix stuff"
      Height          =   1455
      Left            =   120
      TabIndex        =   31
      Top             =   5880
      Width           =   2775
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Info on selected item..."
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2340
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdFix 
         Caption         =   "Fix checked"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1135
      End
   End
   Begin VB.PictureBox picPaypal 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   6240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmMain.frx":1A7FA
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   112
      Top             =   210
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.TextBox txtNothing 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
   Begin VB.Frame fraHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      Begin VB.TextBox txtHelp 
         Height          =   3735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   5895
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
   Begin VB.Frame fraSubmit 
      Height          =   1455
      Left            =   3000
      TabIndex        =   68
      Top             =   5880
      Width           =   2895
      Begin VB.CommandButton cmdAnalyze 
         Caption         =   "AnalyzeThis"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   78
         Top             =   195
         Width           =   1935
      End
      Begin VB.CommandButton cmdMainMenu 
         Caption         =   "Main Menu"
         Height          =   375
         Left            =   720
         TabIndex        =   89
         Top             =   945
         Width           =   1455
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkConfigTabs 
         Caption         =   "Backups"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkConfigTabs 
         Caption         =   "Ignorelist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkConfigTabs 
         Caption         =   "Main"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Height          =   4935
         Index           =   3
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Visible         =   0   'False
         Width           =   6135
         Begin VB.VScrollBar vscMiscTools 
            Height          =   4695
            LargeChange     =   100
            Left            =   5760
            Max             =   100
            SmallChange     =   20
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
         Begin VB.Frame fraMiscToolsScroll 
            BorderStyle     =   0  'None
            Height          =   9735
            Left            =   0
            TabIndex        =   66
            Top             =   -2040
            Width           =   5655
            Begin VB.CommandButton cmdLangLoad 
               Caption         =   "Load this file"
               Height          =   375
               Left            =   4080
               TabIndex        =   140
               Top             =   6240
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CommandButton cmdLangReset 
               Caption         =   "Reset to default"
               Height          =   375
               Left            =   4080
               TabIndex        =   139
               Top             =   6720
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.FileListBox filLanguage 
               Height          =   1065
               Left            =   120
               Pattern         =   "*.lng;*.LNG"
               TabIndex        =   138
               Top             =   6120
               Visible         =   0   'False
               Width           =   3855
            End
            Begin VB.CommandButton cmdARSMan 
               Caption         =   "Open Uninstall Manager..."
               Height          =   375
               Left            =   120
               TabIndex        =   123
               Top             =   4080
               Width           =   2295
            End
            Begin VB.CommandButton cmdDeleteService 
               Caption         =   "Delete an NT service..."
               Height          =   375
               Left            =   120
               TabIndex        =   118
               Top             =   3000
               Width           =   2295
            End
            Begin VB.CheckBox chkAdvLogEnvVars 
               Caption         =   "Include environment variables in logfile"
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   5400
               Width           =   5055
            End
            Begin VB.CommandButton cmdADSSpy 
               Caption         =   "Open ADS Spy..."
               Height          =   375
               Left            =   120
               TabIndex        =   90
               Top             =   3480
               Width           =   2295
            End
            Begin VB.CommandButton cmdDelOnReboot 
               Caption         =   "Delete a file on reboot..."
               Height          =   375
               Left            =   120
               TabIndex        =   76
               Top             =   2400
               Width           =   2295
            End
            Begin VB.CommandButton cmdHostsManager 
               Caption         =   "Open hosts file manager"
               Height          =   375
               Left            =   120
               TabIndex        =   75
               Top             =   1920
               Width           =   2295
            End
            Begin VB.CommandButton cmdProcessManager 
               Caption         =   "Open process manager"
               Height          =   375
               Left            =   120
               TabIndex        =   74
               Top             =   1440
               Width           =   2295
            End
            Begin VB.CheckBox chkStartupListComplete 
               Caption         =   "List empty sections (complete)"
               Height          =   255
               Left            =   2760
               TabIndex        =   72
               Top             =   600
               Width           =   2535
            End
            Begin VB.CheckBox chkStartupListFull 
               Caption         =   "List also minor sections (full)"
               Height          =   255
               Left            =   2760
               TabIndex        =   71
               Top             =   360
               Width           =   2415
            End
            Begin VB.TextBox txtCheckUpdateProxy 
               Height          =   285
               Left            =   2640
               TabIndex        =   70
               Top             =   8400
               Width           =   2895
            End
            Begin VB.CommandButton cmdCheckUpdate 
               Caption         =   "Check for update online"
               Height          =   375
               Left            =   120
               TabIndex        =   69
               Top             =   7920
               Width           =   2295
            End
            Begin VB.CommandButton cmdStartupList 
               Caption         =   "Generate StartupList log"
               Height          =   375
               Left            =   120
               TabIndex        =   67
               Top             =   360
               Width           =   2295
            End
            Begin VB.CheckBox chkDoMD5 
               Caption         =   "Calculate MD5 of files if possible"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   5040
               Width           =   5055
            End
            Begin VB.Label lblInfo 
               Caption         =   "Open the integrated ADS Spy utility to scan for hidden data streams."
               Height          =   435
               Index           =   3
               Left            =   2520
               TabIndex        =   91
               Top             =   3480
               Width           =   3000
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "Language files"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   120
               TabIndex        =   137
               Top             =   5880
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000014&
               Index           =   15
               X1              =   120
               X2              =   5520
               Y1              =   5775
               Y2              =   5775
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000010&
               Index           =   14
               X1              =   120
               X2              =   5520
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Label lblInfo 
               Caption         =   "Open a utility to manage the items in the Add/Remove Software list."
               Height          =   405
               Index           =   7
               Left            =   2520
               TabIndex        =   124
               Top             =   4080
               Width           =   3015
            End
            Begin VB.Label lblInfo 
               Caption         =   "Delete a Windows NT Service (O23). USE WITH CAUTION! (WinNT4/2k/XP only)"
               Height          =   495
               Index           =   6
               Left            =   2520
               TabIndex        =   119
               Top             =   3000
               Width           =   3015
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000014&
               Index           =   7
               X1              =   120
               X2              =   5520
               Y1              =   8895
               Y2              =   8895
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "Update check"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   720
               TabIndex        =   88
               Top             =   7920
               Width           =   1155
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000010&
               Index           =   6
               X1              =   120
               X2              =   5520
               Y1              =   8880
               Y2              =   8880
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "Update check"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   87
               Top             =   7560
               Width           =   1155
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "Advanced settings (these will not be saved)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   86
               Top             =   4680
               Width           =   3705
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "System tools"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   85
               Top             =   1080
               Width           =   1110
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000010&
               Index           =   5
               X1              =   120
               X2              =   5520
               Y1              =   4560
               Y2              =   4560
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000014&
               Index           =   4
               X1              =   120
               X2              =   5520
               Y1              =   4575
               Y2              =   4575
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000014&
               Index           =   3
               X1              =   120
               X2              =   5520
               Y1              =   7455
               Y2              =   7455
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000010&
               Index           =   2
               X1              =   120
               X2              =   5520
               Y1              =   7440
               Y2              =   7440
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000014&
               Index           =   1
               X1              =   120
               X2              =   5520
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line linSeperator 
               BorderColor     =   &H80000010&
               Index           =   0
               X1              =   120
               X2              =   5520
               Y1              =   945
               Y2              =   945
            End
            Begin VB.Label lblInfo 
               Caption         =   "If a file cannot be removed from memory, Windows can be setup to delete it when the system is restarted."
               Height          =   585
               Index           =   2
               Left            =   2520
               TabIndex        =   83
               Top             =   2340
               Width           =   3000
            End
            Begin VB.Label lblConfigInfo 
               Caption         =   "Opens a small editor for the 'hosts' file."
               Height          =   435
               Index           =   13
               Left            =   2520
               TabIndex        =   82
               Top             =   1920
               Width           =   2970
            End
            Begin VB.Label lblConfigInfo 
               Caption         =   "Opens a small process manager, working much like the Task Manager."
               Height          =   435
               Index           =   12
               Left            =   2520
               TabIndex        =   81
               Top             =   1440
               Width           =   3090
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "Use this proxy server (host:port) :"
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   80
               Top             =   8430
               Width           =   2490
            End
            Begin VB.Label lblConfigInfo 
               Caption         =   "Phones home to www.spywareinfo.com to see if a newer HijackThis version exists."
               Height          =   435
               Index           =   10
               Left            =   2520
               TabIndex        =   79
               Top             =   7920
               Width           =   3090
            End
            Begin VB.Label lblConfigInfo 
               AutoSize        =   -1  'True
               Caption         =   "StartupList (integrated: v1.52)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   77
               Top             =   0
               Width           =   2595
            End
         End
         Begin VB.Image imgMiscToolsUp 
            Height          =   120
            Left            =   5805
            Picture         =   "frmMain.frx":1A840
            ToolTipText     =   "Click to scroll"
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Image imgMiscToolsUp2 
            Height          =   120
            Left            =   5805
            Picture         =   "frmMain.frx":1A8B1
            Top             =   240
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Image imgMiscToolsDown 
            Height          =   120
            Left            =   5805
            Picture         =   "frmMain.frx":1A922
            ToolTipText     =   "Click to scroll"
            Top             =   4560
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Image imgMiscToolsDown2 
            Height          =   120
            Left            =   5805
            Picture         =   "frmMain.frx":1A992
            Top             =   4320
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Image imgMiscToolsUp1 
            Height          =   120
            Left            =   5805
            Picture         =   "frmMain.frx":1AA02
            Top             =   120
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Image imgMiscToolsDown1 
            Height          =   120
            Left            =   5805
            Picture         =   "frmMain.frx":1AA73
            Top             =   4440
            Visible         =   0   'False
            Width           =   150
         End
      End
      Begin VB.Frame fraProcessManager 
         Caption         =   "Itty Bitty Process Manager"
         Height          =   3855
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.ListBox lstProcManDLLs 
            Height          =   660
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   63
            Top             =   2520
            Visible         =   0   'False
            Width           =   8175
         End
         Begin VB.CheckBox chkProcManShowDLLs 
            Alignment       =   1  'Right Justify
            Caption         =   "Show DLLs"
            Height          =   255
            Left            =   4440
            TabIndex        =   62
            Top             =   330
            Width           =   1215
         End
         Begin VB.ListBox lstProcessManager 
            Height          =   1185
            IntegralHeight  =   0   'False
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   53
            Top             =   600
            Width           =   8175
         End
         Begin VB.CommandButton cmdProcManKill 
            Caption         =   "Kill process"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmdProcManRun 
            Caption         =   "Run..."
            Height          =   375
            Left            =   2760
            TabIndex        =   51
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmdProcManBack 
            Caption         =   "Back"
            Height          =   375
            Left            =   4440
            TabIndex        =   50
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmdProcManRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   1440
            TabIndex        =   49
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Loaded DLL libraries by selected process:"
            Height          =   195
            Index           =   21
            Left            =   240
            TabIndex        =   120
            Top             =   1920
            Width           =   2955
         End
         Begin VB.Image imgProcManCopy 
            Height          =   240
            Left            =   3720
            Picture         =   "frmMain.frx":1AAE3
            ToolTipText     =   "Copy list to clipboard"
            Top             =   330
            Width           =   240
         End
         Begin VB.Image imgProcManSave 
            Height          =   240
            Left            =   4080
            Picture         =   "frmMain.frx":1AC2D
            ToolTipText     =   "Save list to file.."
            Top             =   330
            Width           =   240
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Running processes:"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label lblProcManDblClick 
            Caption         =   "Double-click a file to view its properties."
            Height          =   390
            Left            =   5760
            TabIndex        =   65
            Top             =   3330
            Width           =   1935
         End
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Caption         =   "fraConfigBackup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
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
            Height          =   375
            Left            =   7440
            TabIndex        =   26
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigBackupDelete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   7440
            TabIndex        =   25
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigBackupRestore 
            Caption         =   "Restore"
            Height          =   375
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
         Begin VB.Label lblConfigInfo 
            Caption         =   $"frmMain.frx":1AFB6
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   43
            Top             =   0
            Width           =   5490
         End
      End
      Begin VB.Frame fraADSSpy 
         Caption         =   "ADS Spy"
         Height          =   3615
         Left            =   120
         TabIndex        =   92
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.Frame fraADSSpyStatus 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   2880
            Width           =   8055
            Begin VB.Label lblADSSpyStatus 
               AutoSize        =   -1  'True
               Caption         =   "Ready."
               Height          =   195
               Left            =   0
               TabIndex        =   117
               Top             =   0
               Width           =   525
            End
         End
         Begin VB.ListBox lstADSSpyResults 
            Height          =   1620
            IntegralHeight  =   0   'False
            ItemData        =   "frmMain.frx":1B09B
            Left            =   120
            List            =   "frmMain.frx":1B09D
            Style           =   1  'Checkbox
            TabIndex        =   102
            Top             =   1200
            Width           =   8175
         End
         Begin VB.CommandButton cmdADSSpyBack 
            Caption         =   "Back"
            Height          =   375
            Left            =   5160
            TabIndex        =   101
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdADSSpySaveLog 
            Caption         =   "Save log..."
            Height          =   375
            Left            =   1440
            TabIndex        =   100
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CheckBox chkADSSpyCalcMD5 
            Caption         =   "Calculate MD5 checksum of streams"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   840
            Width           =   3255
         End
         Begin VB.CheckBox chkADSSpyIgnoreSystem 
            Caption         =   "Ignore safe system info streams"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   600
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox chkADSSpyQuick 
            Caption         =   "Quick scan (Windows base folder only)"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   360
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CommandButton cmdADSSpyRemove 
            Caption         =   "Remove selected"
            Height          =   375
            Left            =   3000
            TabIndex        =   96
            Top             =   3120
            Width           =   1695
         End
         Begin VB.CommandButton cmdADSSpyScan 
            Caption         =   "Scan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdADSSpyWhatsThis 
            Caption         =   "What's this?"
            Height          =   375
            Left            =   3720
            TabIndex        =   94
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdADSSpyHelp 
            Caption         =   "Help"
            Height          =   375
            Left            =   5160
            TabIndex        =   93
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Caption         =   "fraConfigIgnorelist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
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
         Begin VB.ListBox lstIgnore 
            Height          =   2385
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   720
            Width           =   7215
         End
         Begin VB.CommandButton cmdConfigIgnoreDelSel 
            Caption         =   "Delete"
            Height          =   375
            Left            =   7440
            TabIndex        =   23
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdConfigIgnoreDelAll 
            Caption         =   "Delete all"
            Height          =   375
            Left            =   7440
            TabIndex        =   24
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblConfigInfo 
            Caption         =   $"frmMain.frx":1B09F
            Height          =   585
            Index           =   5
            Left            =   120
            TabIndex        =   41
            Top             =   0
            Width           =   5700
         End
      End
      Begin VB.Frame fraConfigTabs 
         BorderStyle     =   0  'None
         Caption         =   "fraConfigMain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   8415
         Begin VB.CheckBox chkConfigStartupScan 
            Caption         =   "Run HijackThis scan at startup and show it when items are found"
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   1620
            Width           =   7335
         End
         Begin VB.CheckBox chkShowN00bFrame 
            Caption         =   "Show intro frame at startup"
            Height          =   255
            Left            =   120
            TabIndex        =   114
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
         Begin VB.CheckBox chkIgnoreSafe 
            Caption         =   "Ignore non-standard but safe domains in IE (e.g. msn.com, microsoft.com)"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   810
            Width           =   7335
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
            Top             =   2280
            Width           =   6375
         End
         Begin VB.TextBox txtDefSearchPage 
            Height          =   285
            Left            =   2040
            TabIndex        =   17
            Top             =   2640
            Width           =   6375
         End
         Begin VB.TextBox txtDefSearchAss 
            Height          =   285
            Left            =   2040
            TabIndex        =   18
            Top             =   3000
            Width           =   6375
         End
         Begin VB.TextBox txtDefSearchCust 
            Height          =   285
            Left            =   2040
            TabIndex        =   19
            Top             =   3360
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
         Begin VB.Label lblConfigInfo 
            Caption         =   "Below URLs will be used when fixing hijacked/unwanted MSIE pages:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   2010
            Width           =   5025
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Default Start Page:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   2280
            Width           =   1395
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Default Search Page:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   1530
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Default Search Assistant:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   3000
            Width           =   1830
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Default Search Customize:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   37
            Top             =   3360
            Width           =   1905
         End
      End
      Begin VB.Frame fraUninstMan 
         Caption         =   "Add/Remove Programs Manager"
         Height          =   3855
         Left            =   120
         TabIndex        =   121
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CommandButton cmdUninstManSave 
            Caption         =   "Save list..."
            Height          =   375
            Left            =   5400
            TabIndex        =   135
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtUninstManCmd 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   134
            Top             =   1410
            Width           =   2535
         End
         Begin VB.TextBox txtUninstManName 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   133
            Top             =   1050
            Width           =   2535
         End
         Begin VB.CommandButton cmdUninstManRefresh 
            Caption         =   "Refresh list"
            Height          =   375
            Left            =   4080
            TabIndex        =   132
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmdUninstManEdit 
            Caption         =   "Edit uninstall command"
            Height          =   375
            Left            =   6120
            TabIndex        =   131
            Top             =   1920
            Width           =   2055
         End
         Begin VB.CommandButton cmdUninstManBack 
            Caption         =   "Back"
            Height          =   375
            Left            =   6960
            TabIndex        =   129
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmdUninstManDelete 
            Caption         =   "Delete this entry"
            Height          =   375
            Left            =   4080
            TabIndex        =   128
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CommandButton cmdUninstManOpen 
            Caption         =   "Open Add/Remove Software list"
            Height          =   375
            Left            =   4080
            TabIndex        =   127
            Top             =   2400
            Width           =   4155
         End
         Begin VB.ListBox lstUninstMan 
            Height          =   2820
            IntegralHeight  =   0   'False
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   122
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label lblInfo 
            Caption         =   $"frmMain.frx":1B177
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   130
            Top             =   360
            Width           =   8175
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Caption         =   "Uninstall command:"
            Height          =   255
            Index           =   10
            Left            =   4080
            TabIndex        =   126
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Caption         =   "Name:"
            Height          =   255
            Index           =   8
            Left            =   4080
            TabIndex        =   125
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.Frame fraHostsMan 
         Caption         =   "Hosts file manager"
         Height          =   3735
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CommandButton cmdHostsManOpen 
            Caption         =   "Open in Notepad"
            Height          =   375
            Left            =   2760
            TabIndex        =   61
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CommandButton cmdHostsManBack 
            Caption         =   "Back"
            Height          =   375
            Left            =   4440
            TabIndex        =   60
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton cmdHostsManToggle 
            Caption         =   "Toggle line(s)"
            Height          =   375
            Left            =   1440
            TabIndex        =   59
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton cmdHostsManDel 
            Caption         =   "Delete line(s)"
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   3240
            Width           =   1215
         End
         Begin VB.ListBox lstHostsMan 
            Height          =   2340
            IntegralHeight  =   0   'False
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   57
            Top             =   600
            Width           =   8175
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Note: changes to the hosts file take effect when you restart your browser."
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   64
            Top             =   3000
            Width           =   5415
         End
         Begin VB.Label lblConfigInfo 
            AutoSize        =   -1  'True
            Caption         =   "Hosts file located at: C:\WINDOWS\hosts"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   2985
         End
      End
   End
   Begin VB.Frame fraN00b 
      Caption         =   "New users quickstart"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   103
      Top             =   840
      Visible         =   0   'False
      Width           =   8655
      Begin VB.ComboBox cboN00bLanguage 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   4380
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdN00bScan 
         Caption         =   "Do a system scan only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   106
         Top             =   1440
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bHJTQuickStart 
         Caption         =   "Open online HijackThis QuickStart"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   109
         Top             =   3720
         Width           =   3975
      End
      Begin VB.CheckBox chkShowN00b 
         Caption         =   "Don't show this frame again when I start HijackThis"
         Height          =   255
         Left            =   360
         TabIndex        =   111
         Top             =   5520
         Width           =   5535
      End
      Begin VB.CommandButton cmdN00bClose 
         Caption         =   "None of the above, just start the program"
         Height          =   495
         Left            =   360
         TabIndex        =   110
         Top             =   4800
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bTools 
         Caption         =   "Open the Misc Tools section"
         Height          =   495
         Left            =   360
         TabIndex        =   108
         Top             =   2880
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bBackups 
         Caption         =   "View the list of backups"
         Height          =   495
         Left            =   360
         TabIndex        =   107
         Top             =   2280
         Width           =   3975
      End
      Begin VB.CommandButton cmdN00bLog 
         Caption         =   "Do a system scan and save a logfile"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   105
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Change language:"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   141
         Top             =   4440
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000014&
         Index           =   11
         X1              =   480
         X2              =   4200
         Y1              =   3555
         Y2              =   3555
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   10
         X1              =   480
         X2              =   4200
         Y1              =   3540
         Y2              =   3540
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000014&
         Index           =   9
         X1              =   480
         X2              =   4200
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   8
         X1              =   480
         X2              =   4200
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Courtesy of TomCoyote.org"
         Height          =   195
         Index           =   5
         Left            =   4560
         TabIndex        =   113
         Top             =   3840
         Width           =   2025
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "What would you like to do?"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   104
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   47
      Top             =   330
      Visible         =   0   'False
      Width           =   45
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
   Begin VB.Shape shpMD5Background 
      BackStyle       =   1  'Opaque
      Height          =   120
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label lblMD5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calculating MD5 checksum of [file]..."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   480
      TabIndex        =   46
      Top             =   360
      Visible         =   0   'False
      Width           =   5595
   End
   Begin VB.Shape shpProgress 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape shpBackground 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":1B249
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":1B321
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   6015
   End
   Begin VB.Menu mnuADSSpy 
      Caption         =   "ADSSpy popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuADSSpySelAll 
         Caption         =   "Select all"
      End
      Begin VB.Menu mnuADSSpySelNone 
         Caption         =   "Select none"
      End
      Begin VB.Menu mnuADSSpySelInv 
         Caption         =   "Invert selection"
      End
      Begin VB.Menu mnuADSSpyStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuADSSpySave 
         Caption         =   "Save results to disk..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Make the "results" list have a horizontal scrollbar
'DECLARATIONS: begin
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
'DECLARATIONS: end

Private bSwitchingTabs As Boolean
Private bIsBeta As Boolean
Private Const sKeyUninstall As String = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
Private szLogData As String
Public bEULAArgee As Boolean

Private Sub cboN00bLanguage_Click()
    Dim sFile$
    sFile = cboN00bLanguage.List(cboN00bLanguage.ListIndex)
    If sFile = vbNullString Then Exit Sub
    If sFile = "(Default)" Then
        LoadDefaultLanguage
    Else
        LoadLanguageFile sFile, Not Me.Visible
    End If
    RegSave "LanguageFile", sFile
End Sub

Private Sub chkAdvLogEnvVars_Click()
    If chkAdvLogEnvVars.Value = 1 Then
        bLogEnvVars = True
    Else
        bLogEnvVars = False
    End If
End Sub

Private Sub chkDoMD5_Click()
    If chkDoMD5.Value = 1 Then
        bMD5 = True
    Else
        bMD5 = False
    End If
End Sub

Private Sub chkProcManShowDLLs_Click()
    lstProcManDLLs.Visible = CBool(chkProcManShowDLLs.Value)
    'lblConfigInfo(21).Caption = "Loaded DLL libraries by selected process: (" & lstProcManDLLs.ListCount & ")"
    lblConfigInfo(21).Caption = Translate(178) & " (" & lstProcManDLLs.ListCount & ")"
    On Error Resume Next
    'lstProcessManager.ListIndex = 0
    lstProcessManager_Click
    lstProcessManager.SetFocus
    Form_Resize
End Sub

Private Sub chkShowN00b_Click()
RegSave "ShowIntroFrame", CStr(chkShowN00b.Value)
End Sub

Private Sub cmdADSSpy_Click()
    fraConfigTabs(3).Visible = False
    lstADSSpyResults.Clear
    chkADSSpyQuick.Value = 1
    chkADSSpyIgnoreSystem.Value = 1
    chkADSSpyCalcMD5.Value = 0
    fraADSSpy.Visible = True
    'lblADSSpyStatus.Caption = "Ready."
    lblADSSpyStatus.Caption = Translate(200)
    If cmdADSSpyScan.Enabled = True Then modADSSpy.CheckIfSystemIsNTFS
End Sub

Private Sub cmdADSSpyBack_Click()
    If cmdADSSpyScan.Caption = "Abort" Then cmdADSSpyScan_Click
    fraADSSpy.Visible = False
    fraConfigTabs(3).Visible = True
End Sub

Private Sub cmdADSSpyHelp_Click()
    MsgBox Translate(201), vbInformation
'    MsgBox "Using ADS Spy is very easy: just click 'Scan', wait until the " & _
'           "scan completes, then select the ADS streams you want to " & _
'           "remove and click 'Remove selected'. If you are unsure which " & _
'           "streams to remove, ask someone for help. Don't delete streams " & _
'           "if you don't know what they are!" & vbCrLf & vbCrLf & _
'           "The three checkboxes are:" & vbCrLf & vbCrLf & _
'           "Quick Scan: only scans the Windows folder. So far all known malware that " & _
'           "uses ADS to hide itself, hides in the Windows folder. Unchecking " & _
'           "this will make ADS Spy scan the entire system (i.e. all drives)." & vbCrLf & vbCrLf & _
'           "Ignore safe system info streams: Windows, Internet Explorer and a few antivirus " & _
'           "programs use ADS to store metadata for certain folders and files. " & _
'           "These streams can safely be ignored, they are harmless." & vbCrLf & vbCrLf & _
'           "Calculate MD5 checksums of streams: For antispyware program " & _
'           "development or antivirus analysis only." & vbCrLf & vbCrLf & _
'           "Note: the default settings of above three checkboxes should " & _
'           "be fine for most people. There's no need to change any " & _
'           "of them unless you are a developer or anti-malware expert.", vbInformation
End Sub

Private Sub cmdADSSpyRemove_Click()
    ADSSpyRemove lstADSSpyResults
End Sub

Private Sub cmdADSSpySaveLog_Click()
    If lstADSSpyResults.ListCount = 0 Then Exit Sub
    Dim sLogFile$, sLog$, i%
    sLogFile = CmnDlgSaveFile("Save ADS Spy log...", "Text files (*.txt)|*.txt|All files (*.*)|*.*", "adsspy.txt")
    If sLogFile = vbNullString Then Exit Sub
    For i = 0 To lstADSSpyResults.ListCount - 1
        sLog = sLog & lstADSSpyResults.List(i) & vbCrLf
    Next i
    Open sLogFile For Output As #1
        Print #1, sLog
    Close #1
    ShellExecute Me.hwnd, "open", sWinDir & "\notepad.exe", sLogFile, vbNullString, 1
End Sub

Private Sub cmdADSSpyScan_Click()
    'If cmdADSSpyScan.Caption = "Abort" Then
    If cmdADSSpyScan.Caption = Translate(202) Then
        bADSSpyAbortScanNow = True
        Exit Sub
    End If
    
    bADSSpyAbortScanNow = False
    'cmdADSSpyScan.Caption = "Abort"
    cmdADSSpyScan.Caption = Translate(202)
    lstADSSpyResults.Clear
    ADSSpyScan CBool(chkADSSpyQuick.Value), CBool(chkADSSpyIgnoreSystem.Value), CBool(chkADSSpyCalcMD5)
    'cmdADSSpyScan.Caption = "Scan"
    cmdADSSpyScan.Caption = Translate(196)
    If bADSSpyAbortScanNow Then
        'lblADSSpyStatus.Caption = "Scan aborted!"
        lblADSSpyStatus.Caption = Translate(203)
    Else
        'lblADSSpyStatus.Caption = "Scan complete."
        lblADSSpyStatus.Caption = Translate(204)
    End If
End Sub

Private Sub cmdADSSpyWhatsThis_Click()
    MsgBox Translate(205), vbInformation
'    MsgBox "Alternate Data Streams (ADSs) are pieces of info hidden as metadata on files. They are " & _
'           "not visible in Explorer and the size they take up is not reported by Windows. " & _
'           "Recent browser hijackers started hiding their files inside ADSs, and very few anti-malware " & _
'           "scanners detect this (yet)." & vbCrLf & _
'           "Use ADS Spy to find and remove these streams." & vbCrLf & vbCrLf & _
'           "Note: this app also displays legitimate " & _
'           "ADS streams. Do not delete streams if you are not completely sure they are malicious!", vbInformation
End Sub

Private Sub cmdAnalyze_Click()
    
    Dim sLog$, i%, sProcessList$
    Dim hSnap&, uProcess As PROCESSENTRY32, sDummy$ '9x
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024) 'NT
    
    Dim BeginTime As Date
    Dim FinishTime As Date
    Dim ElapsedTime As Long
    BeginTime = Now
         
    If Not bIsWinNT Then
        hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap < 1 Then
            'sProcessList = "(Unable to list running processes)" & vbCrLf
            'sProcessList = "(" & Translate(28) & ")" & vbCrLf
          
        End If
        
        uProcess.dwSize = Len(uProcess)
        If ProcessFirst(hSnap, uProcess) = 0 Then
            'sProcessList = "(Unable to list running processes)" & vbCrLf
           ' sProcessList = "(" & Translate(28) & ")" & vbCrLf
          
        End If
        
        'sProcessList = "Running processes:" & vbCrLf
        'sProcessList = Translate(29) & ":" & vbCrLf
        Do
            sDummy = Left(uProcess.szExeFile, InStr(uProcess.szExeFile, Chr(0)) - 1)
            sProcessList = sProcessList & "P01 - " & sDummy & "|"
        Loop Until ProcessNext(hSnap, uProcess) = 0
        CloseHandle hSnap
        sProcessList = sProcessList & "|"
    Else
        On Error Resume Next
        If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
            'sProcessList = "(" & Translate(28) & ")" & vbCrLf
            'sProcessList = "(Unable to list running processes)" & vbCrLf
           
        End If
             
        'sProcessList = "Running processes:" & vbCrLf
        'sProcessList = Translate(29) & ":" & vbCrLf
        lNumProcesses = lNeeded / 4
        For i = 1 To lNumProcesses
            hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(i))
            If hProc <> 0 Then
                lNeeded = 0
                sProcessName = String(260, 0)
                If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                    GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                    sProcessName = Left(sProcessName, InStr(sProcessName, Chr(0)) - 1)
                    
                    'do some spell-checking
                    If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                    If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                    sProcessName = Replace(sProcessName, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
                    sProcessName = Replace(sProcessName, "SYSTEMROOT", sWinDir, , , vbTextCompare)
                    
                    sProcessList = sProcessList & "P01 - " & sProcessName & "|"
                End If
                CloseHandle hProc
            End If
        Next i
        sProcessList = sProcessList & "|"
    End If
    
       
    Dim q As Integer
    szLogData = sProcessList
    For q = 0 To lstResults.ListCount
        szLogData = szLogData & lstResults.List(q) & "|"
    Next q
    cmdAnalyze.Enabled = False
    
    If True = IsOnline Then
             
        cmdAnalyze.Caption = "Please Wait"
        
        szLogData = ObfuscateData(szLogData)
        
        Dim sThisVersion, szBuf As String
        sThisVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
        cmdAnalyze.Caption = "AnalyzeThis"
        ShellExecute Me.hwnd, "open", "http://sourceforge.net/p/hjt/support-requests/", "", "", 1
        Exit Sub
        End If
        
        ParseHTTPResponse szBuf
        If 7 < Len(szSubmitUrl) Then
            ShellExecute Me.hwnd, "open", "http://sourceforge.net/p/hjt/support-requests/" & szResponse, "", "", 1
            ParseHTTPResponse szResponse
            
            cmdAnalyze.Enabled = True
            FinishTime = Now
            ElapsedTime = DateDiff("s", BeginTime, FinishTime)
        Else: MsgBox "Please go to http://sourceforge.net/p/hjt/support-requests/"
        End If
    
    cmdAnalyze.Caption = "AnalyzeThis"

End Sub
Function ObfuscateData(szDataIn As String) As String

Dim szDataOut As String
Dim szHexVal As String
Dim chrCode As Integer
Dim i As Long
szDataOut = "7"
For i = 1 To Len(szDataIn)
    chrCode = Asc(Mid(szDataIn, i, 1))
    szHexVal = Hex$(chrCode)
    szDataOut = szDataOut & StrReverse(szHexVal)
Next i
ObfuscateData = szDataOut
Exit Function
End Function

Private Sub cmdARSMan_Click()
    fraConfigTabs(3).Visible = False
    fraUninstMan.Visible = True
    cmdUninstManRefresh_Click
End Sub

Private Sub cmdDeleteService_Click()
    If Not bIsWinNT Then Exit Sub
    Dim sServiceName$, sWhiteList$, sDisplayName$, sFile$, sCompany$, j%
    sWhiteList = "Microsoft Corporation|" & _
                 "Symantec Corporation|" & _
                 "Trend Micro Inc.|" & _
                 "Trend Micro Incorporated.|" & _
                 "GRISOFT, s.r.o."
    
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
        MsgBox Replace(Translate(115), "[]", sServiceName), vbExclamation
'        MsgBox "Service '" & sServiceName & "' was not found in the Registry." & vbCrLf & _
'               "Make sure you entered the name of the service correctly.", vbExclamation
        Exit Sub
    End If
    
    If RegGetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "Start") <> 4 Then
        MsgBox Replace(Translate(116), "[]", sServiceName), vbCritical
'        MsgBox "The service '" & sServiceName & "' is enabled and/or running. Disable it first, " & _
'               "using HijackThis itself (from the scan results) or the Services.msc window.", vbCritical
        Exit Sub
    End If
    
    sFile = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "ImagePath")
    sDisplayName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServiceName, "DisplayName")
    If sFile <> vbNullString Then
        'remove double quotes for long pathnames
        If Left(sFile, 1) = """" Then sFile = Mid(sFile, 2)
        If Right(sFile, 1) = """" Then sFile = Left(sFile, Len(sFile) - 1)
        
        'expand aliases
        sFile = Replace(sFile, "%systemroot%", sWinDir, , , vbTextCompare)
        sFile = Replace(sFile, "\systemroot", sWinDir, , , vbTextCompare)
        sFile = Replace(sFile, "systemroot", sWinDir, , , vbTextCompare)
        
        'prefix windows folder if not specified
        If InStr(1, sFile, "system32\", vbTextCompare) = 1 Then
            sFile = sWinDir & "\" & sFile
        End If
        
        'remove parameters
        j = InStrRev(sFile, ".exe", , vbTextCompare) + 3
        If j < Len(sFile) And j > 3 Then sFile = Left(sFile, j)
        
        'add .exe if not specified
        If InStr(1, sFile, ".exe", vbTextCompare) = 0 And _
           InStr(1, sFile, ".sys", vbTextCompare) = 0 Then
            If InStr(sFile, " ") > 0 Then
                sFile = Left(sFile, InStr(sFile, " ") - 1)
                sFile = sFile & ".exe"
            End If
        End If
    Else
        sFile = "(no file)"
    End If
    
    sCompany = GetFilePropCompany(sFile)
    If sCompany = vbNullString Then sCompany = "Unknown owner" '"?"
    
    If Not FileExists(sFile) Then sFile = sFile & " (file missing)"
    
    If InStr(1, sWhiteList, sCompany, vbTextCompare) > 0 Then
        MsgBox "The service you entered is system-critical! " & _
               "It can't be deleted.", vbCritical
        Exit Sub
    End If
    
    If MsgBox(Translate(117) & vbCrLf & _
              "Short name: " & sServiceName & vbCrLf & _
              "Full name: " & sDisplayName & vbCrLf & _
              "File: " & sFile & vbCrLf & _
              "Owner: " & sCompany & vbCrLf & vbCrLf & _
              Translate(118), vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
'    If MsgBox("The following service was found:" & vbCrLf & _
'              "Short name: " & sServiceName & vbCrLf & _
'              "Full name: " & sDisplayName & vbCrLf & _
'              "File: " & sFile & vbCrLf & _
'              "Owner: " & sCompany & vbCrLf & vbCrLf & _
'              "Are you absolutely sure you want to delete this service?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
        
        DeleteNTService sServiceName
        
    End If
End Sub

Private Sub cmdDelOnReboot_Click()
    Dim sFilename$
    sFilename = CmnDlgOpenFile("Enter file to delete on reboot...", "All files (*.*)|*.*|DLL libraries (*.dll)|*.dll|Program files (*.exe)|*.exe")
    If sFilename = vbNullString Then Exit Sub
    DeleteFileOnReboot sFilename, True
End Sub

Private Sub cmdHostsManager_Click()
    fraConfigTabs(3).Visible = False
    fraHostsMan.Visible = True
    ListHostsFile lstHostsMan, lblConfigInfo(14)
End Sub

Private Sub cmdHostsManBack_Click()
    fraHostsMan.Visible = False
    fraConfigTabs(3).Visible = True
End Sub

Private Sub cmdHostsManDel_Click()
    If lstHostsMan.ListIndex <> -1 And lstHostsMan.ListCount > 0 Then
        HostsDeleteLine lstHostsMan
        ListHostsFile lstHostsMan, lblConfigInfo(14)
    End If
End Sub

Private Sub cmdHostsManOpen_Click()
    ShellExecute Me.hwnd, "open", sWinDir & "\notepad.exe", sHostsFile, vbNullString, 1
End Sub

Private Sub cmdHostsManToggle_Click()
    If lstHostsMan.ListIndex <> -1 And lstHostsMan.ListCount > 0 Then
        HostsToggleLine lstHostsMan
        ListHostsFile lstHostsMan, lblConfigInfo(14)
    End If
End Sub

Private Sub cmdLangLoad_Click()
    LoadLanguageFile filLanguage.List(filLanguage.ListIndex)
    RegSave "LanguageFile", filLanguage.List(filLanguage.ListIndex)
End Sub

Private Sub cmdLangReset_Click()
    LoadDefaultLanguage
    RegDel "LanguageFile"
End Sub

Private Sub cmdMainMenu_Click()
    If cmdConfig.Caption = Translate(19) Then
        Dim i%, iIgnoreNum%, sIgnore$
        bAutoSelect = IIf(chkAutoMark.Value = 1, True, False)
        bConfirm = IIf(chkConfirm.Value = 1, True, False)
        bMakeBackup = IIf(chkBackup.Value = 1, True, False)
        bIgnoreSafe = IIf(chkIgnoreSafe.Value = 1, True, False)
        bLogProcesses = IIf(chkLogProcesses.Value = 1, True, False)
        
        
        For i = 0 To UBound(sRegVals)
            If sRegVals(i) = vbNullString Then Exit For
            sRegVals(i) = Crypt(sRegVals(i), sProgramVersion)
        Next i
        For i = 0 To UBound(sRegVals)
            If sRegVals(i) = vbNullString Then Exit For
            sRegVals(i) = Replace(sRegVals(i), "$DEFSTARTPAGE", txtDefStartPage.Text)
            sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHPAGE", txtDefSearchPage.Text)
            sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHASS", txtDefSearchAss.Text)
            sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHCUST", txtDefSearchCust.Text)
        Next i
        For i = 0 To UBound(sRegVals)
            If sRegVals(i) = vbNullString Then Exit For
            sRegVals(i) = Crypt(sRegVals(i), sProgramVersion, True)
        Next i
        
        RegDel "IgnoreNum"
        For i = 1 To 99
            RegDel "Ignore" & CStr(i)
        Next i
        RegSave "IgnoreNum", CStr(lstIgnore.ListCount)
        For i = 0 To lstIgnore.ListCount - 1
            RegSave "Ignore" & CStr(i + 1), lstIgnore.List(i)
        Next i
        
        RegSave "AutoSelect", CStr(Abs(CInt(bAutoSelect)))
        RegSave "Confirm", CStr(Abs(CInt(bConfirm)))
        RegSave "MakeBackup", CStr(Abs(CInt(bMakeBackup)))
        RegSave "IgnoreSafe", CStr(Abs(CInt(bIgnoreSafe)))
        RegSave "LogProcesses", CStr(Abs(CInt(bLogProcesses)))
        RegSave "ShowIntroFrame", CStr(chkShowN00bFrame.Value)
        
        If chkConfigStartupScan.Value = 1 Then
            RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HijackThis startup scan", App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "HijackThis.exe /startupscan"
        Else
            RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HijackThis startup scan"
        End If
        
        RegSave "DefStartPage", txtDefStartPage.Text
        RegSave "DefSearchPage", txtDefSearchPage.Text
        RegSave "DefSearchAss", txtDefSearchAss.Text
        RegSave "DefSearchCust", txtDefSearchCust.Text
        
        'If cmdScan.Caption = "Scan" Then
        If cmdScan.Caption = Translate(11) Then
            lblInfo(0).Visible = True
        Else
            lblInfo(1).Visible = True
        End If
        picPaypal.Visible = True
        fraConfig.Visible = False
        fraProcessManager.Visible = False
        fraHostsMan.Visible = False
        fraUninstMan.Visible = False
        fraADSSpy.Visible = False
        If chkConfigTabs(3).Value = 1 Then fraConfigTabs(3).Visible = True
        'cmdConfig.Caption = "Config..."
        cmdConfig.Caption = Translate(18)
        cmdHelp.Enabled = True
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
    cmdScan.Caption = Translate(11)
    cmdHelp.Caption = Translate(16)
    lblInfo(0).Visible = True
    lblInfo(1).Visible = False
    chkShowN00b.Value = RegRead("ShowIntroFrame", "0")
End Sub

Private Sub cmdN00bBackups_Click()
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
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
End Sub

Private Sub cmdN00bHJTQuickStart_Click()
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
    'ShellExecute Me.hWnd, "open", "http://tomcoyote.org/hjt/#Top", "", "", 1
    Dim szQSUrl As String
    szQSUrl = Translate(360) & "?hjtver=" & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    If True = IsOnline Then
        ShellExecute Me.hwnd, "open", szQSUrl, "", "", 1
    Else
        MsgBox "No Internet Connection Available"
    End If
End Sub

Private Sub cmdN00bLog_Click()
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
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
    cmdScan_Click
End Sub

Private Sub cmdN00bTools_Click()
    fraN00b.Visible = False
    fraScan.Visible = True
    fraOther.Visible = True
    fraSubmit.Visible = True
    lstResults.Visible = True
    'If chkShowN00b.Value Then RegSave "ShowIntroFrame", "0"
    cmdConfig_Click
    chkConfigTabs_Click 3
End Sub

Private Sub cmdProcessManager_Click()
    fraConfigTabs(3).Visible = False
    fraProcessManager.Visible = True
    cmdProcManRefresh_Click
End Sub

Private Sub cmdProcManBack_Click()
    fraProcessManager.Visible = False
    fraConfigTabs(3).Visible = True
End Sub

Private Sub cmdProcManKill_Click()
    Dim sMsg$, i%, s$
    sMsg = Translate(179) & vbCrLf
    'sMsg = "Are you sure you want to close the selected processes?" & vbCrLf
    For i = 0 To lstProcessManager.ListCount - 1
        If lstProcessManager.Selected(i) Then
            sMsg = sMsg & Mid(lstProcessManager.List(i), InStr(lstProcessManager.List(i), vbTab) + 1) & vbCrLf
        End If
    Next i
    'sMsg = sMsg & vbCrLf & "Any unsaved data in it will be lost."
    sMsg = sMsg & vbCrLf & Translate(180)
    If MsgBox(sMsg, vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    'pause selected processes
    For i = 0 To lstProcessManager.ListCount - 1
        If lstProcessManager.Selected(i) Then
            s = lstProcessManager.List(i)
            s = Left(s, InStr(s, vbTab) - 1)
            PauseProcess CLng(s)
        End If
    Next i
    For i = 0 To lstProcessManager.ListCount - 1
        If lstProcessManager.Selected(i) Then
            s = lstProcessManager.List(i)
            s = Left(s, InStr(s, vbTab) - 1)
            If Not bIsWinNT Then
                KillProcess CLng(s)
            Else
                KillProcessNT CLng(s)
            End If
        End If
    Next i
    Sleep 1000
    'resume any processes still alive
    For i = 0 To lstProcessManager.ListCount - 1
        If lstProcessManager.Selected(i) Then
            s = lstProcessManager.List(i)
            s = Left(s, InStr(s, vbTab) - 1)
            PauseProcess CLng(s), False
        End If
    Next i
    
    cmdProcManRefresh_Click
End Sub

Private Sub cmdProcManRefresh_Click()
    Dim s$
    lstProcessManager.Clear
    If Not bIsWinNT Then
        RefreshProcessList lstProcessManager
    Else
        RefreshProcessListNT lstProcessManager
        lstProcessManager.ListIndex = 0
        If lstProcManDLLs.Visible Then
            s = lstProcessManager.List(lstProcessManager.ListIndex)
            s = Left(s, InStr(s, vbTab) - 1)
            If Not bIsWinNT Then
                RefreshDLLList CLng(s), lstProcManDLLs
            Else
                RefreshDLLListNT CLng(s), lstProcManDLLs
            End If
        End If
    End If
    lblConfigInfo(8).Caption = Translate(171) & " (" & lstProcessManager.ListCount & ")"
    lblConfigInfo(21).Caption = Translate(178) & " (" & lstProcManDLLs.ListCount & ")"
    'lblConfigInfo(8).Caption = "Running processes: (" & lstProcessManager.ListCount & ")"
    'lblConfigInfo(21).Caption = "Loaded DLL libraries by selected process: (" & lstProcManDLLs.ListCount & ")"
    
End Sub

Private Sub cmdProcManRun_Click()
    If Not bIsWinNT Then
        SHRunDialog Me.hwnd, 0, 0, Translate(181), Translate(182), 0
        'SHRunDialog Me.hWnd, 0, 0, "Run", "Type the name of a program, folder, document or Internet resource, and Windows will open it for you.", 0
    Else
        SHRunDialog Me.hwnd, 0, 0, StrConv(Translate(181), vbUnicode), StrConv(Translate(182), vbUnicode), 0
        'SHRunDialog Me.hWnd, 0, 0, StrConv("Run", vbUnicode), StrConv("Type the name of a program, folder, document or Internet resource, and Windows will open it for you.", vbUnicode), 0
    End If
    Sleep 1000
    cmdProcManRefresh_Click
End Sub

Private Sub cmdUninstManBack_Click()
    fraUninstMan.Visible = False
    fraConfigTabs(3).Visible = True
End Sub

Private Sub cmdUninstManDelete_Click()
    Dim sItems$(), i&, j&, sName$, sUninst$
    If lstUninstMan.ListCount = 0 Then Exit Sub
    sName = txtUninstManName.Text
    sUninst = txtUninstManCmd.Text
    j = lstUninstMan.ListIndex
    If MsgBox(Translate(220) & vbCrLf & vbCrLf & sName, vbQuestion + vbYesNo) = vbYes Then
        sItems = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sKeyUninstall), "|")
        If UBound(sItems) <> -1 Then
            For i = 0 To UBound(sItems)
                If sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & sItems(i), "DisplayName") And _
                   sUninst = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & sItems(i), "UninstallString") Then
                    RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & sItems(i)
                    Exit For
                End If
            Next i
        End If
    End If
    cmdUninstManRefresh_Click
    On Error Resume Next
    lstUninstMan.ListIndex = j - 1
    lstUninstMan.ListIndex = j
End Sub

Private Sub cmdUninstManEdit_Click()
    Dim s$, sItems$(), i&
    If lstUninstMan.ListCount = 0 Then Exit Sub
    s = InputBox(Translate(221) & ", '" & txtUninstManName.Text & ":", "Edit uninstall command", txtUninstManCmd.Text)
    's = InputBox("Enter the new uninstall command for this program, '" & txtUninstManName.Text & ":", "Edit uninstall command", txtUninstManCmd.Text)
    If s = vbNullString Then Exit Sub
    
    sItems = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sKeyUninstall), "|")
    If UBound(sItems) = -1 Then Exit Sub
    For i = 0 To UBound(sItems)
        If txtUninstManName.Text = RegGetString(HKEY_LOCAL_MACHINE, sKeyUninstall & "\" & sItems(i), "DisplayName") And _
           txtUninstManCmd.Text = RegGetString(HKEY_LOCAL_MACHINE, sKeyUninstall & "\" & sItems(i), "UninstallString") Then
            RegSetStringVal HKEY_LOCAL_MACHINE, sKeyUninstall & "\" & sItems(i), "UninstallString", s
            MsgBox Translate(222), vbInformation
            'MsgBox "New uninstall string saved!", vbInformation
            Exit For
        End If
    Next i
End Sub

Private Sub cmdUninstManOpen_Click()
    ShellExecute 0, "open", "control.exe", "appwiz.cpl", vbNullString, 1
End Sub

Private Sub cmdUninstManRefresh_Click()
    Dim sItems$(), sName$, sUninst$, i&
    lstUninstMan.Clear
    sItems = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sKeyUninstall), "|")
    If UBound(sItems) = -1 Then Exit Sub
    For i = 0 To UBound(sItems)
        sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & sItems(i), "DisplayName")
        sUninst = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & sItems(i), "UninstallString")
        If sName <> vbNullString And sUninst <> vbNullString Then
            lstUninstMan.AddItem sName
        End If
    Next i
    On Error Resume Next
    lstUninstMan.ListIndex = 0
    lstUninstMan.SetFocus
End Sub

Private Sub cmdUninstManSave_Click()
    Dim sList$, i&, sUninst$, sFile$
    If lstUninstMan.ListCount = 0 Then Exit Sub
    sFile = CmnDlgSaveFile("Save Add/Remove Software list to disk...", "Text files (*.txt)|*.txt|All files (*.*)|*.*", "uninstall_list.txt")
    If sFile = vbNullString Then Exit Sub
    For i = 0 To lstUninstMan.ListCount - 1
        sList = sList & lstUninstMan.List(i) & vbCrLf
    Next i
    
    Open sFile For Output As #1
        Print #1, sList
    Close #1
    ShellExecute 0, "open", "notepad.exe", sFile, vbNullString, 1
End Sub

Private Sub Form_Load()
    
    LoadDefaultLanguage
    On Error GoTo Error:
    'bIsBeta = True
    bIsBeta = False
    
    Me.Caption = "Trend Micro HijackThis - v" & App.Major & "." & App.Minor & "." & App.Revision & IIf(bIsBeta, " (BETA)", vbNullString)
    SetAllFontCharset
    
    'Only load 64 bit support if OS supports it.
    If True = IsProcedureAvail("IsWow64Process", "kernel32") Then
        ToggleWow64FSRedirection False
    End If
    
    GetHostsAndWinDir
    
    If App.PrevInstance Then
        'MsgBox "HijackThis is already running.", vbExclamation
        MsgBox Translate(2), vbExclamation
        End
    End If
    
    If InStr(1, Command$, "/uninstall") > 0 Then
        Me.Hide
        cmdUninstall_Click
        End
    End If
    If InStr(1, Command$, "/complete") > 0 Then frmMain.chkStartupListComplete.Value = 1
    If InStr(1, Command$, "/full") > 0 Then frmMain.chkStartupListFull.Value = 1
    'If InStr(1, Command$, "/md5") > 0 Then bMD5 = True
    If InStr(1, Command$, "/deleteonreboot") > 0 Then
        SilentDeleteOnReboot Command$
        End
    End If
    
    If InStr(1, Command$, "/ihatewhitelists") > 0 Then bIgnoreAllWhitelists = True
    
    'set encryption string - THOU SHALT NOT STEAL
     sProgramVersion = Chr(&H54) & Chr(&H48) & Chr(&H4F) & _
      Chr(&H55) & Chr(&H20) & Chr(&H53) & Chr(&H48) & _
      Chr(&H41) & Chr(&H4C) & Chr(&H54) & Chr(&H20) & _
      Chr(&H4E) & Chr(&H4F) & Chr(&H54) & Chr(&H20) & _
      Chr(&H53) & Chr(&H54) & Chr(&H45) & Chr(&H41) & _
      Chr(&H4C)
    
    If Command$ = "/debug" Then
        bDebugMode = True
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "debug.log") <> vbNullString Then
            DeleteFile App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "debug.log"
            Open App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "debug.log" For Output As #3
        Else
            Open App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "debug.log" For Append As #3
        End If
    End If
    
    fraConfig.Left = 120
    fraHelp.Left = 120
    fraConfig.Top = 120
    fraHelp.Top = 120
    fraMiscToolsScroll.Top = 0
    filLanguage.Path = App.Path
    
    If Screen.Height >= 9000 Then
        Me.Height = CLng(RegRead("WinHeight", "8000"))
        If Me.Height < 8000 Then Me.Height = 8000
    Else
        Me.Height = CLng(RegRead("WinHeight", "6600"))
        If Me.Height < 6600 Then Me.Height = 6600
    End If
    Me.Width = CLng(RegRead("WinWidth", "9000"))
    If Me.Width < 9000 Then Me.Width = 9000
    
    If RegRead("ShowIntroFrame", "0") = "0" Then
        LoadLanguageList
        fraN00b.Visible = True
        fraScan.Visible = False
        fraOther.Visible = False
        lstResults.Visible = False
        fraSubmit.Visible = False
    End If
    If RegRead("ShowIntroFrame", "0") = "0" Then
    chkShowN00b.Value = 0
    Else
    chkShowN00b.Value = 1
    End If
    
    If RegRead("LanguageFile") = vbNullString Then
        LoadDefaultLanguage
    Else
        LoadLanguageFile RegRead("LanguageFile"), True
    End If
    
    LoadStuff 'regvals, filevals, safelspfiles,safeprotocols
    LoadSettings
    GetLSPCatalogNames
    CheckForReadOnlyMedia
    CheckDateFormat
    CheckForStartedFromTempDir
    SetListBoxColumns lstProcessManager
    SetListBoxColumns lstProcManDLLs
    'MsgBox "bIsUSADateFormat: " & bIsUSADateFormat
    iItems = UBound(sRegVals) + 1 + UBound(sFileVals) + _
             1 + 23 'regvals+r3+filevals+ns/moz+other
    
    'If Not bIsWinNT Then cmdDelOnReboot.Enabled = False
    If Not bIsWinNT Then cmdDeleteService.Enabled = False
    
    If InStr(Command$, "/autolog") > 0 Or _
       InStr(Command$, "/silentautolog") > 0 Then
        bAutoLog = True
        If InStr(Command$, "/silentautolog") > 0 Then bAutoLogSilent = True
        On Error Resume Next
        If bAutoLogSilent Then Me.WindowState = vbMinimized
        If bAutoLogSilent Then Me.WindowState = vbMinimizedNoFocus
        On Error GoTo Error:
        cmdN00bClose_Click
        cmdScan_Click
        DoEvents
        If bAutoLogSilent Then End
    End If
    
    If InStr(1, Command$, "/startupscan") > 0 Then
        Me.Show
        DoEvents
        Me.WindowState = vbMinimized
        cmdN00bClose_Click
        cmdScan_Click
        DoEvents
        If lstResults.ListCount = 0 Then
            End
        Else
            Me.WindowState = vbNormal
        End If
    End If
    
    txtHelp.Text = vbCrLf & _
     "* Trend Micro HijackThis v" & App.Major & "." & App.Minor & "." & App.Revision & " *" & vbCrLf & _
     vbCrLf & vbCrLf & "See bottom for version history." & vbCrLf & vbCrLf

    'maybe I'll put this in later
'    If Translate(361) <> vbNullString Then
'        txtHelp.Text = txtHelp.Text & Translate(361) & vbCrLf & vbCrLf
'    End If

    txtHelp.Text = txtHelp.Text & "The different sections of hijacking " & _
     "possibilities have been separated into the following groups." & vbCrLf & _
     "You can get more detailed information about an item " & _
     "by selecting it from the list of found items OR " & _
     "highlighting the relevant line below, and clicking " & _
     "'Info on selected item'." & vbCrLf & vbCrLf & _
     " R - Registry, StartPage/SearchPage changes" & vbCrLf & _
     "    R0 - Changed registry value" & vbCrLf & _
     "    R1 - Created registry value" & vbCrLf & _
     "    R2 - Created registry key" & vbCrLf & _
     "    R3 - Created extra registry value where only one should be" & vbCrLf & _
     " F - IniFiles, autoloading entries" & vbCrLf & _
     "    F0 - Changed inifile value" & vbCrLf & _
     "    F1 - Created inifile value" & vbCrLf & _
     "    F2 - Changed inifile value, mapped to Registry" & vbCrLf & _
     "    F3 - Created inifile value, mapped to Registry" & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     " N - Netscape/Mozilla StartPage/SearchPage changes" & vbCrLf & _
     "    N1 - Change in prefs.js of Netscape 4.x" & vbCrLf & _
     "    N2 - Change in prefs.js of Netscape 6" & vbCrLf & _
     "    N3 - Change in prefs.js of Netscape 7" & vbCrLf & _
     "    N4 - Change in prefs.js of Mozilla" & vbCrLf & _
     " O - Other, several sections which represent:" & vbCrLf & _
     "    O1 - Hijack of auto.search.msn.com with Hosts file" & vbCrLf & _
     "    O2 - Enumeration of existing MSIE BHO's" & vbCrLf & _
     "    O3 - Enumeration of existing MSIE toolbars" & vbCrLf & _
     "    O4 - Enumeration of suspicious autoloading Registry entries" & vbCrLf & _
     "    O5 - Blocking of loading Internet Options in Control Panel" & vbCrLf & _
     "    O6 - Disabling of 'Internet Options' Main tab with Policies" & vbCrLf & _
     "    O7 - Disabling of Regedit with Policies" & vbCrLf & _
     "    O8 - Extra MSIE context menu items" & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "    O9 - Extra 'Tools' menuitems and buttons" & vbCrLf & _
     "    O10 - Breaking of Internet access by New.Net or WebHancer" & vbCrLf & _
     "    O11 - Extra options in MSIE 'Advanced' settings tab" & vbCrLf & _
     "    O12 - MSIE plugins for file extensions or MIME types" & vbCrLf & _
     "    O13 - Hijack of default URL prefixes" & vbCrLf & _
     "    O14 - Changing of IERESET.INF" & vbCrLf & _
     "    O15 - Trusted Zone Autoadd" & vbCrLf & _
     "    O16 - Download Program Files item" & vbCrLf & _
     "    O17 - Domain hijack" & vbCrLf & _
     "    O18 - Enumeration of existing protocols and filters" & vbCrLf & _
     "    O19 - User stylesheet hijack" & vbCrLf & _
     "    O20 - AppInit_DLLs autorun Registry value, Winlogon Notify Registry keys" & vbCrLf & _
     "    O21 - ShellServiceObjectDelayLoad (SSODL) autorun Registry key" & vbCrLf & _
     "    O22 - SharedTaskScheduler autorun Registry key" & vbCrLf & _
     "    O23 - Enumeration of NT Services" & vbCrLf & _
     "    O24 - Enumeration of ActiveX Desktop Components" & vbCrLf & vbCrLf
     
    txtHelp.Text = txtHelp.Text & _
     "Command-line parameters:" & vbCrLf & _
     "* /autolog - automatically scan the system, save a logfile and open it" & vbCrLf & _
     "* /ihatewhitelists - ignore all internal whitelists" & vbCrLf & _
     "* /uninstall - remove all HijackThis Registry entries, backups and quit" & vbCrLf & _
     "* /silentautuolog - the same as /autolog, except with no required user intervention" & vbCrLf
          
    txtHelp.Text = txtHelp.Text & vbCrLf & "* Version history *" & vbCrLf & vbCrLf & _
     "[v2.0.5 Beta]" & vbCrLf & _
     "* Fixed No internet connection available when pressing the button Analyze This" & vbCrLf & _
     "* Fixed the link of update website, now send you to sourceforge.net projects" & vbCrLf & _
     "* Fixed left-right scrollbar when in safe mode or low screen resolution" & vbCrLf & _
     "[v2.0.4]" & vbCrLf & _
     "* Fixed parser issues on winlogon notify" & vbCrLf & _
     "* Fixed issues to handle certain environment variables" & vbCrLf & _
     "* Rename HJT generates complete scan log" & vbCrLf & _
     "[v2.00.0]" & vbCrLf & _
     "* AnalyzeThis added for log file statistics" & vbCrLf & _
     "* Recognizes Windows Vista and IE7" & vbCrLf & _
     "* Fixed a few bugs in the O23 method" & vbCrLf & _
     "* Fixed a bug in the O22 method (SharedTaskScheduler)" & vbCrLf & _
     "* Did a few tweaks on the log format" & vbCrLf & _
     "* Fixed and improved ADS Spy" & vbCrLf & _
     "* Improved Itty Bitty Procman (processes are frozen before they are killed)" & vbCrLf & _
     "* Added listing of O4 autoruns from other users" & vbCrLf & _
     "* Added listing of the Policies Run items in O4 method, used by SmitFraud trojan" & vbCrLf & _
     "* Added /silentautolog parameter for system admins" & vbCrLf & _
     "* Added /deleteonreboot [file] parameter for system admins" & vbCrLf & _
     "* Added O24 - ActiveX Desktop Components enumeration" & vbCrLf
     '"* Added multilanguage support" & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "* Added Enhanced Security Confirguration (ESC) Zones to O15 Trusted Sites check" & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.99.1]" & vbCrLf & _
     "* Added Winlogon Notify keys to O20 listing" & vbCrLf & _
     "* Fixed crashing bug on certain Win2000 and WinXP systems at O23 listing" & vbCrLf & _
     "* Fixed lots and lots of 'unexpected error' bugs" & vbCrLf & _
     "* Fixed lots of inproper functioning bugs (i.e. stuff that didn't work)" & vbCrLf & _
     "* Added 'Delete NT Service' function in Misc Tools section" & vbCrLf & _
     "* Added ProtocolDefaults to O15 listing" & vbCrLf & _
     "* Fixed MD5 hashing not working" & vbCrLf & _
     "* Fixed 'ISTSVC' autorun entries with garbage data not being fixed" & vbCrLf & _
     "* Fixed HijackThis uninstall entry not being updated/created on new versions" & vbCrLf & _
     "* Added Uninstall Manager in Misc Tools to manage 'Add/Remove Software' list" & vbCrLf & _
     "* Added option to scan the system at startup, then show results or quit if nothing found" & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.99]" & vbCrLf & _
     " * Added O23 (NT Services) in light of newer trojans" & vbCrLf & _
     " * Integrated ADS Spy into Misc Tools section" & vbCrLf & _
     " * Added 'Action taken' to info in 'More info on this item'" & vbCrLf & _
     "[v1.98]" & vbCrLf & _
     " * Definitive support for Japanese/Chinese/Korean systems" & vbCrLf & _
     " * Added O20 (AppInit_DLLs) in light of newer trojans" & vbCrLf & _
     " * Added O21 (ShellServiceObjectDelayLoad, SSODL) in light of newer trojans" & vbCrLf & _
     " * Added O22 (SharedTaskScheduler) in light of newer trojans" & vbCrLf & _
     " * Backups of fixed items are now saved in separate folder" & vbCrLf & _
     " * HijackThis now checks if it was started from a temp folder" & vbCrLf & _
     " * Added a small process manager (Misc Tools section)" & vbCrLf & _
     "[v1.96]" & vbCrLf & _
     " * Lots of bugfixes and small enhancements! Among others:" & vbCrLf & _
     " * Fix for Japanese IE toolbars" & vbCrLf & _
     " * Fix for searchwww.com fake CLSID trick in IE toolbars and BHO's" & vbCrLf & _
     " * Attributes on Hosts file will now be restored when scanning/fixing/restoring it." & vbCrLf & _
     " * Added several files to the LSP whitelist" & vbCrLf & _
     " * Fixed some issues with incorrectly re-encrypting data, making R0/R1 go undetected until a restart" & vbCrLf & _
     " * All sites in the Trusted Zone are now shown, with the exception of those on the nonstandard but safe domain list" & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.95]" & vbCrLf & _
     " * Added a new regval to check for from Whazit hijack (Start Page_bak)." & vbCrLf & _
     " * Excluded IE logo change tweak from toolbar detection (BrandBitmap and SmBrandBitmap)." & vbCrLf & _
     " * New in logfile: Running processes at time of scan." & vbCrLf & _
     " * Checkmarks for running StartupList with /full and /complete in HijackThis UI." & vbCrLf & _
     " * New O19 method to check for Datanotary hijack of user stylesheet." & vbCrLf & _
     " * Google.com IP added to whitelist for Hosts file check." & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.94]" & vbCrLf & _
     " * Fixed a bug in the Check for Updates function that could cause corrupt downloads on certain systems." & vbCrLf & _
     " * Fixed a bug in enumeration of toolbars (Lop toolbars are now listed!)." & vbCrLf & _
     " * Added imon.dll, drwhook.dll and wspirda.dll to LSP safelist." & vbCrLf & _
     " * Fixed a bug where DPF could not be deleted." & vbCrLf & _
     " * Fixed a stupid bug in enumeration of autostarting shortcuts." & vbCrLf & _
     " * Fixed info on Netscape 6/7 and Mozilla saying '%shitbrowser%' (oops)." & vbCrLf & _
     " * Fixed bug where logfile would not auto-open on systems that don't have .log filetype registered." & vbCrLf & _
     " * Added support for backing up F0 and F1 items (d'oh!)." & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.93]" & vbCrLf & _
     " * Added mclsp.dll (McAfee), WPS.DLL (Sygate Firewall), zklspr.dll (Zero Knowledge) and mxavlsp.dll (OnTrack) to LSP safelist." & vbCrLf & _
     " * Fixed a bug in LSP routine for Win95. " & vbCrLf & _
     " * Made taborder nicer." & vbCrLf & _
     " * Fixed a bug in backup/restore of IE plugins." & vbCrLf & _
     " * Added UltimateSearch hijack in O17 method (I think). " & vbCrLf & _
     " * Fixed a bug with detecting/removing BHO's disabled by BHODemon." & vbCrLf & _
     " * Also fixed a bug in StartupList (now version 1.52.1)." & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.92]" & vbCrLf & _
     " * Fixed two stupid bugs in backup restore function. " & vbCrLf & _
     " * Added DiamondCS file to LSP files safelist." & vbCrLf & _
     " * Added a few more items to the protocol safelist." & vbCrLf & _
     " * Log is now opened immediately after saving. " & vbCrLf & _
     " * Removed rd.yahoo.com from NSBSD list (spammers are starting to use this, no doubt spyware authors will follow)." & vbCrLf & _
     " * Updated integrated StartupList to v1.52." & vbCrLf & _
     " * In light of SpywareNuker/BPS Spyware Remover, any strings relevant to reverse-engineers are now encrypted." & vbCrLf & _
     " * Rudimentary proxy support for the Check for Updates function." & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.91]" & vbCrLf & _
     " * Added rd.yahoo.com to the Nonstandard But Safe Domains list. " & vbCrLf & _
     " * Added 8 new protocols to the protocol check safelist, as well as showing the file that handles the protocol in the log (O18)." & vbCrLf & _
     " * Added listing of programs/links in Startup folders (O4)." & vbCrLf & _
     " * Fixed 'Check for Update' not detecting new versions." & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.9]" & vbCrLf & _
     " * Added check for Lop.com 'Domain' hijack (O17)." & vbCrLf & _
     " * Bugfix in URLSearchHook (R3) fix." & vbCrLf & _
     " * Improved O1 (Hosts file) check." & vbCrLf & _
     " * Rewrote code to delete BHO's, fixing a really nasty bug with orphaned BHO keys." & vbCrLf & _
     " * Added AutoConfigURL and proxyserver checks (R1)." & vbCrLf & _
     " * IE Extensions (Button/Tools menuitem) in HKEY_CURRENT_USER are now also detected." & vbCrLf & _
     " * Added check for extra protocols (O18)." & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.81]" & vbCrLf & _
     " * Added 'ignore non-standard but safe domains' option." & vbCrLf & _
     " * Improved Winsock LSP hijackers detection." & vbCrLf & _
     " * Integrated StartupList updated to v1.4." & vbCrLf & _
     "[v1.8]" & vbCrLf & _
     " * Fixed a few bugs." & vbCrLf & _
     " * Adds detecting of free.aol.com in Trusted Zone." & vbCrLf & _
     " * Adds checking of URLSearchHooks key, which should have only one value." & vbCrLf & _
     " * Adds listing/deleting of Download Program Files." & vbCrLf & _
     " * Integrated StartupList into the new 'Misc Tools' section of the Config screen!" & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.71]" & vbCrLf & _
     " * Improves detecting of O6." & vbCrLf & _
     " * Some internal changes/improvements." & vbCrLf & _
     "[v1.7]" & vbCrLf & _
     " * Adds backup function! Yay!" & vbCrLf & _
     " * Added check for default URL prefix" & vbCrLf & _
     " * Added check for changing of IERESET.INF" & vbCrLf & _
     " * Added check for changing of Netscape/Mozilla homepage and default search engine." & vbCrLf & _
     "[v1.61]" & vbCrLf & _
     " * Fixes Runtime Error when Hosts file is empty." & vbCrLf & _
     "[v1.6]" & vbCrLf & _
     " * Added enumerating of MSIE plugins" & vbCrLf & _
     " * Added check for extra options in 'Advanced' tab of 'Internet Options'." & vbCrLf
    txtHelp.Text = txtHelp.Text & _
     "[v1.5]" & vbCrLf & _
     " * Adds 'Uninstall & Exit' and 'Check for update online' functions. " & vbCrLf & _
     " * Expands enumeration of autoloading Registry entries (now also scans for .vbs, .js, .dll, rundll32 and service)" & vbCrLf & _
     "[v1.4]" & vbCrLf & _
     " * Adds repairing of broken Internet access (aka Winsock or LSP fix) by New.Net/WebHancer" & vbCrLf & _
     " * A few bugfixes/enhancements" & vbCrLf & _
     "[v1.3]" & vbCrLf & _
     " * Adds detecting of extra MSIE context menu items" & vbCrLf & _
     " * Added detecting of extra 'Tools' menu items and extra buttons" & vbCrLf & _
     " * Added 'Confirm deleting/ignoring items' checkbox" & vbCrLf & _
     "[v1.2]" & vbCrLf & _
     " * Adds 'Ignorelist' and 'Info' functions" & vbCrLf & _
     "[v1.1]" & vbCrLf & _
     " * Supports BHO's, some default URL changes" & vbCrLf & _
     "[v1.0]" & vbCrLf & _
     " * Original release" & vbCrLf & vbCrLf & _
     "A good thing to do after version updates is clear " & _
     "your Ignore list and re-add them, as the format of " & _
     "detected items sometimes changes." & vbCrLf & vbCrLf
    
    Exit Sub
    
Error:
    ErrorMsg "frmMain_Load", Err.Number, Err.Description
End Sub

Private Sub chkAutoMark_Click()
    Dim sMsg$
    If chkAutoMark.Value = 0 Then Exit Sub
    If RegRead("SeenAutoMarkWarning", "0") = "1" Then Exit Sub
    
    sMsg = Translate(57)
'    sMsg = "Are you sure you want to enable this option?" & vbCrLf & _
'           "HijackThis is not a 'click & fix' program. " & _
'           "Because it targets *general* hijacking methods, " & _
'           "false positives are a frequent occurrence." & vbCrLf & _
'           "If you enable this option, you might disable " & _
'           "programs or drivers you need. However, it is " & _
'           "highly unlikely you will break your system " & _
'           "beyond repair. So you should only enable this " & _
'           "option if you know what you're doing!"
           
    If MsgBox(sMsg, vbExclamation + vbYesNo) = vbYes Then
        RegSave "SeenAutoMarkWarning", "1"
        Exit Sub
    Else
        chkAutoMark.Value = Abs(chkAutoMark.Value - 1)
    End If
End Sub

Private Sub chkConfigTabs_Click(Index As Integer)
    If bSwitchingTabs Then Exit Sub
    bSwitchingTabs = True
    
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
    
    fraProcessManager.Visible = False
    fraHostsMan.Visible = False
    fraADSSpy.Visible = False
    fraUninstMan.Visible = False
    
    bSwitchingTabs = False
End Sub

Private Sub cmdCheckUpdate_Click()
    CheckForUpdate
End Sub

Private Sub cmdConfig_Click()
    Dim i%, iIgnoreNum%, sIgnore$
    On Error GoTo Error:
    'If cmdConfig.Caption = "Config..." Then
    If cmdConfig.Caption = Translate(18) Then
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        picPaypal.Visible = False
        lstResults.Visible = False
        fraConfig.Visible = True
        'cmdConfig.Caption = "Back"
        cmdConfig.Caption = Translate(19)
        cmdHelp.Enabled = False
        cmdSaveDef.Enabled = False
        fraScan.Enabled = False
        cmdScan.Enabled = False
        cmdFix.Enabled = False
        cmdInfo.Enabled = False
        chkShowN00bFrame.Value = RegRead("ShowIntroFrame", "0")
        
        txtNothing.Visible = False
        
        For i = 0 To UBound(sRegVals)
            If sRegVals(i) = vbNullString Then Exit For
            sRegVals(i) = Crypt(sRegVals(i), sProgramVersion)
        Next i
        For i = 0 To 50
            sRegVals(i) = Replace(sRegVals(i), txtDefStartPage.Text, "$DEFSTARTPAGE")
            sRegVals(i) = Replace(sRegVals(i), txtDefSearchPage.Text, "$DEFSEARCHPAGE")
            sRegVals(i) = Replace(sRegVals(i), txtDefSearchAss.Text, "$DEFSEARCHASS")
            sRegVals(i) = Replace(sRegVals(i), txtDefSearchCust.Text, "$DEFSEARCHCUST")
        Next i
        For i = 0 To UBound(sRegVals)
            If sRegVals(i) = vbNullString Then Exit For
            sRegVals(i) = Crypt(sRegVals(i), sProgramVersion, True)
        Next i
        
        lstIgnore.Clear
        iIgnoreNum = CInt(RegRead("IgnoreNum", "0"))
        If iIgnoreNum > 0 Then
            For i = 1 To iIgnoreNum
                sIgnore = RegRead("Ignore" & CStr(i), "")
                If sIgnore <> vbNullString Then
                    lstIgnore.AddItem sIgnore
                Else
                    Exit For
                End If
            Next i
        End If
        ListBackups
    Else
        bAutoSelect = IIf(chkAutoMark.Value = 1, True, False)
        bConfirm = IIf(chkConfirm.Value = 1, True, False)
        bMakeBackup = IIf(chkBackup.Value = 1, True, False)
        bIgnoreSafe = IIf(chkIgnoreSafe.Value = 1, True, False)
        bLogProcesses = IIf(chkLogProcesses.Value = 1, True, False)
        
        For i = 0 To UBound(sRegVals)
            If sRegVals(i) = vbNullString Then Exit For
            sRegVals(i) = Crypt(sRegVals(i), sProgramVersion)
        Next i
        For i = 0 To UBound(sRegVals)
            If sRegVals(i) = vbNullString Then Exit For
            sRegVals(i) = Replace(sRegVals(i), "$DEFSTARTPAGE", txtDefStartPage.Text)
            sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHPAGE", txtDefSearchPage.Text)
            sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHASS", txtDefSearchAss.Text)
            sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHCUST", txtDefSearchCust.Text)
        Next i
        For i = 0 To UBound(sRegVals)
            If sRegVals(i) = vbNullString Then Exit For
            sRegVals(i) = Crypt(sRegVals(i), sProgramVersion, True)
        Next i
        
        RegDel "IgnoreNum"
        For i = 1 To 99
            RegDel "Ignore" & CStr(i)
        Next i
        RegSave "IgnoreNum", CStr(lstIgnore.ListCount)
        For i = 0 To lstIgnore.ListCount - 1
            RegSave "Ignore" & CStr(i + 1), lstIgnore.List(i)
        Next i
        
        RegSave "AutoSelect", CStr(Abs(CInt(bAutoSelect)))
        RegSave "Confirm", CStr(Abs(CInt(bConfirm)))
        RegSave "MakeBackup", CStr(Abs(CInt(bMakeBackup)))
        RegSave "IgnoreSafe", CStr(Abs(CInt(bIgnoreSafe)))
        RegSave "LogProcesses", CStr(Abs(CInt(bLogProcesses)))
        RegSave "ShowIntroFrame", CStr(chkShowN00bFrame.Value)
        
        If chkConfigStartupScan.Value = 1 Then
            RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HijackThis startup scan", App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "HijackThis.exe /startupscan"
        Else
            RegDelVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HijackThis startup scan"
        End If
        
        RegSave "DefStartPage", txtDefStartPage.Text
        RegSave "DefSearchPage", txtDefSearchPage.Text
        RegSave "DefSearchAss", txtDefSearchAss.Text
        RegSave "DefSearchCust", txtDefSearchCust.Text
        
        'If cmdScan.Caption = "Scan" Then
        If cmdScan.Caption = Translate(11) Then
            lblInfo(0).Visible = True
        Else
            lblInfo(1).Visible = True
        End If
        picPaypal.Visible = True
        lstResults.Visible = True
        fraConfig.Visible = False
        fraProcessManager.Visible = False
        fraHostsMan.Visible = False
        fraUninstMan.Visible = False
        fraADSSpy.Visible = False
        If chkConfigTabs(3).Value = 1 Then fraConfigTabs(3).Visible = True
        'cmdConfig.Caption = "Config..."
        cmdConfig.Caption = Translate(18)
        cmdHelp.Enabled = True
        cmdSaveDef.Enabled = True
        fraScan.Enabled = True
        cmdScan.Enabled = True
        cmdFix.Enabled = True
        cmdInfo.Enabled = True
    End If
    Exit Sub
    
Error:
    ErrorMsg "cmdConfig_Click", Err.Number, Err.Description
End Sub

Private Sub cmdConfigBackupDeleteAll_Click()
    If lstBackups.ListCount = 0 Then Exit Sub
    'If MsgBox("Are you sure you want to delete ALL backups?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    If MsgBox(Translate(88), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'    If MsgBox("Delete all backups? Are you sure? I mean, " & _
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
    On Error Resume Next
    Dim i%
    If lstBackups.ListIndex = -1 Then Exit Sub
    If lstBackups.SelCount = 1 Then
        If MsgBox(Translate(84), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    '    If MsgBox("Are you sure you want to delete this backup?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        If MsgBox(Replace(Translate(85), "[]", lstBackups.SelCount), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If MsgBox("Are you sure you want to delete these " & lstBackups.SelCount & " backups?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    For i = lstBackups.ListCount - 1 To 0 Step -1
        If lstBackups.Selected(i) Then
            DeleteBackup lstBackups.List(i)
            lstBackups.RemoveItem i
        End If
    Next i
End Sub

Private Sub cmdConfigBackupRestore_Click()
    On Error Resume Next
    Dim i%
    If lstBackups.SelCount = 0 Then Exit Sub
    If lstBackups.SelCount = 1 Then
        If MsgBox(Translate(86), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If MsgBox("Restore this item?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        If MsgBox(Replace(Translate(87), "[]", lstBackups.SelCount), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If MsgBox("Restore these " & lstBackups.SelCount & " items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    For i = lstBackups.ListCount - 1 To 0 Step -1
        If lstBackups.Selected(i) Then
            RestoreBackup lstBackups.List(i)
            lstBackups.RemoveItem i
        End If
    Next i
End Sub

Private Sub cmdConfigIgnoreDelAll_Click()
    lstIgnore.Clear
End Sub

Private Sub cmdConfigIgnoreDelSel_Click()
    Dim i%
    On Error Resume Next
    For i = lstIgnore.ListCount - 1 To 0 Step -1
        If lstIgnore.Selected(i) Then lstIgnore.RemoveItem i
    Next i
End Sub

Private Sub cmdFix_Click()
    Dim i%
    On Error GoTo Error:
    If lstResults.ListCount = 0 Then Exit Sub
    If lstResults.SelCount = 0 Then
        If MsgBox(Translate(21), vbQuestion + vbYesNo) = vbNo Then
        'If MsgBox("Nothing selected! Continue?", vbQuestion + vbYesNo) = vbNo Then
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
        If MsgBox(Translate(22), vbExclamation + vbYesNo) = vbNo Then Exit Sub
'        If MsgBox("You selected to fix everything HijackThis has found. " & _
'                  "This could mean items important to your system " & _
'                  "will be deleted and the full functionality of your " & _
'                  "system will degrade." & vbCrLf & vbCrLf & _
'                  "If you aren't sure how to use HijackThis, you should " & _
'                  "ask for help, not blindly fix things. The SpywareInfo " & _
'                  "forums will gladly help you with your log." & vbCrLf & vbCrLf & _
'                  "Are you sure you want to fix all items in your scan " & _
'                  "results?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    End If
    
    If bConfirm Then
        lstResults.ListIndex = -1
        If MsgBox(Replace(Translate(23), "[]", lstResults.SelCount) & _
           IIf(bMakeBackup, ".", ", " & Translate(24)), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'        If MsgBox("Fix " & lstResults.SelCount & _
'         " selected items? This will permanently " & _
'         "delete and/or repair what you selected" & _
'         IIf(bMakeBackup, ".", ", unless you enable backups."), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    If bMakeBackup Then
        For i = 0 To lstResults.ListCount - 1
            If lstResults.Selected(i) Then
                MakeBackup lstResults.List(i)
            End If
        Next i
    End If
    
    shpBackground.Tag = lstResults.SelCount
    shpProgress.Tag = "0"
    shpProgress.Width = 15
    lblInfo(1).Visible = False
    picPaypal.Visible = False
    shpBackground.Visible = True
    shpProgress.Visible = True
    bRebootNeeded = False
    bShownBHOWarning = False
    bShownToolbarWarning = False
    bSeenHostsFileAccessDeniedWarning = False
    
    For i = 0 To lstResults.ListCount - 1
        If lstResults.Selected(i) = True Then
            lstResults.ListIndex = i
            Select Case RTrim(Left(lstResults.List(i), 3))
                Case "R0", "R1", "R2": FixRegItem lstResults.List(i)
                Case "R3":             FixRegistry3Item lstResults.List(i)
                Case "F0", "F1":       FixFileItem lstResults.List(i)
                Case "F2", "F3":       FixFileItem lstResults.List(i)
                Case "N1", "N2", "N3", "N4": FixNetscapeMozilla lstResults.List(i)
                Case "O1":             FixOther1Item lstResults.List(i)
                Case "O2":             FixOther2Item lstResults.List(i)
                Case "O3":             FixOther3Item lstResults.List(i)
                Case "O4":             FixOther4Item lstResults.List(i)
                Case "O5":             FixOther5Item lstResults.List(i)
                Case "O6":             FixOther6Item lstResults.List(i)
                Case "O7":             FixOther7Item lstResults.List(i)
                Case "O8":             FixOther8Item lstResults.List(i)
                Case "O9":             FixOther9Item lstResults.List(i)
                Case "O10":            FixLSP
                Case "O11":            FixOther11Item lstResults.List(i)
                Case "O12":            FixOther12Item lstResults.List(i)
                Case "O13":            FixOther13Item lstResults.List(i)
                Case "O14":            FixOther14Item lstResults.List(i)
                Case "O15":            FixOther15Item lstResults.List(i)
                Case "O16":            FixOther16Item lstResults.List(i)
                Case "O17":            FixOther17Item lstResults.List(i)
                Case "O18":            FixOther18Item lstResults.List(i)
                Case "O19":            FixOther19Item lstResults.List(i)
                Case "O20":            FixOther20Item lstResults.List(i)
                Case "O21":            FixOther21Item lstResults.List(i)
                Case "O22":            FixOther22Item lstResults.List(i)
                Case "O23":            FixOther23Item lstResults.List(i)
                Case "O24":            FixOther24Item lstResults.List(i)
                Case Else
                   ' MsgBox "Fixing of " & RTrim(Left(lstResults.List(i), 3)) & _
                           " is not implemented yet. Bug me about it at " & _
                           "www.merijn.org/contact.html, because I should have done it.", _
                           vbInformation, "bad coder - no donuts"
                    MsgBox "Fixing of " & RTrim(Left(lstResults.List(i), 3)) & _
                           " is not implemented yet.", _
                           vbInformation
            End Select
            UpdateProgressBar
        End If
    Next i
    lstResults.Clear
    cmdFix.Enabled = False
    cmdFix.FontBold = False
    cmdScan.Caption = Translate(11)
    'cmdScan.Caption = "Scan"
    cmdScan.FontBold = True
    shpBackground.Visible = False
    shpProgress.Visible = False
    lblInfo(0).Visible = True
    lblInfo(1).Visible = False
    picPaypal.Visible = True
    On Error Resume Next
    cmdScan.SetFocus
    
    If bRebootNeeded = True Then RestartSystem
    Exit Sub
    
Error:
    ErrorMsg "cmdFix_Click", Err.Number, Err.Description & " (" & lstResults.ListCount & " items in results list)"
End Sub

Private Sub cmdHelp_Click()
    'If cmdHelp.Caption = "Info..." Then
    If cmdHelp.Caption = Translate(16) Then
        lblInfo(0).Visible = False
        picPaypal.Visible = False
        lstResults.Visible = False
        fraHelp.Visible = True
        'cmdHelp.Caption = "Back"
        cmdHelp.Caption = Translate(17)
        cmdConfig.Enabled = False
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
        cmdConfig.Enabled = True
        cmdSaveDef.Enabled = True
        cmdScan.Enabled = True
        cmdFix.Enabled = True
    End If
End Sub

Private Sub cmdInfo_Click()
    If lstResults.Visible Then
        GetInfo lstResults.List(lstResults.ListIndex)
    ElseIf txtHelp.Visible Then
        GetInfo LTrim(txtHelp.SelText)
    End If
End Sub

Private Sub cmdSaveDef_Click()
    On Error GoTo Error:
    If lstResults.SelCount = 0 Then Exit Sub
    If bConfirm Then
        If MsgBox(Translate(25), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'        If MsgBox("This will set HijackThis to ignore the " & _
'                  "checked items, unless they change. Cont" & _
'                  "inue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    Dim i%, j%
    i = CInt(RegRead("IgnoreNum", "0"))
    RegSave "IgnoreNum", CStr(i + lstResults.SelCount)
    j = i + 1
    For i = 0 To lstResults.ListCount - 1
        If lstResults.Selected(i) Then
            RegSave "Ignore" & CStr(j), lstResults.List(i)
            j = j + 1
        End If
    Next i
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
    
Error:
    ErrorMsg "cmdSaveDef_Click", Err.Number, Err.Description
End Sub

Private Sub AddHorizontalScrollBarToResults()
'Adds a horizontal scrollbar to the results display if it is needed.

        'add horizontal scrollbar (after the scan)
        Dim x As Long
        Dim listLength As Integer
        With lstResults
        For listLength = 0 To .ListCount - 1
        If lstResults.Width < TextWidth(.List(listLength)) And x < TextWidth(.List(listLength)) Then
            x = TextWidth(.List(listLength))
        End If
        Next
        End With
        If ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX + 50  ' if twips change to pixels (+50 to account for the width of the vertical scrollbar
        SendMessageByNum lstResults.hwnd, LB_SETHORIZONTALEXTENT, x, 0
        'end add horizontal scrollbar (after the scan)
End Sub

Private Sub cmdScan_Click()
    On Error GoTo Error:
    'If cmdScan.Caption = "Scan" Then
    If cmdScan.Caption = Translate(11) Then
        lblInfo(0).Visible = False
        lblInfo(1).Visible = False
        picPaypal.Visible = False
        shpBackground.Visible = True
        shpProgress.Width = 30 '0.5 * shpBackground.Width
        shpProgress.Visible = True
        'lblMD5.Visible = True
        If bMD5 = False Then lblStatus.Visible = True
        
        cmdScan.Enabled = False
        cmdAnalyze.Enabled = False
    
        StartScan
        
        'add the horizontal scrollbar if needed
        AddHorizontalScrollBarToResults
        
        cmdScan.Enabled = True
        cmdAnalyze.Enabled = True
        
        lblStatus.Visible = False
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
        
        If bAutoLog Then cmdScan_Click
        bAutoLog = False
    Else
        Dim sLogFile$, i%
        If bAutoLog Then
            sLogFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "hijackthis.log"
        Else
            sLogFile = CmnDlgSaveFile("Save logfile...", "Log files (*.log)|*.log|All files (*.*)|*.*", "hijackthis.log")
        End If
        If sLogFile <> vbNullString Then
            
            'check for read-only access on location first
            On Error Resume Next
            Open sLogFile For Output As #1
                Print #1, "."
            Close #1
            If Err Then
                If Err.Number = 52 Then
                    'path/file access error -> readonly
                    MsgBox Translate(26), vbExclamation
'                    MsgBox "Write access was denied to the " & _
'                           "location you specified. Try a " & _
'                           "different location please.", vbExclamation
                    Exit Sub
                End If
            End If
            On Error GoTo Error:
            
            Open sLogFile For Output As #1
                Print #1, CreateLogFile()
            Close #1
            
            If Not bAutoLogSilent Then
                If ShellExecute(Me.hwnd, "open", sLogFile, vbNullString, vbNullString, 1) <= 32 Then
                    'system doesn't know what .log is
                    If FileExists(sWinDir & "\notepad.exe") Then
                        ShellExecute Me.hwnd, "open", sWinDir & "\notepad.exe", sLogFile, vbNullString, 1
                    Else
                        If FileExists(sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\notepad.exe") Then
                            ShellExecute Me.hwnd, "open", sWinDir & IIf(bIsWinNT, "\sytem32", "\system") & "\notepad.exe", sLogFile, vbNullString, 1
                        Else
                            MsgBox Replace(Translate(27), "[]", sLogFile), vbInformation
    '                        MsgBox "The logfile has been saved to " & sLogFile & "." & vbCrLf & _
    '                               "You can open it in a text editor like Notepad.", vbInformation
                        End If
                    End If
                End If
            End If
        End If
        'cmdScan.Caption = "Scan"
        cmdScan.Caption = Translate(11)
    End If
    Exit Sub
    
Error:
    ErrorMsg "cmdScan_Click (" & cmdScan.Caption & ")", Err.Number, Err.Description
End Sub

Private Sub cmdStartupList_Click()
    bStartupListFull = IIf(chkStartupListFull.Value = 1, True, False)
    bStartupListComplete = IIf(chkStartupListComplete.Value = 1, True, False)
    modStartupList.Main
End Sub

Private Sub cmdUninstall_Click()
    On Error Resume Next
    If MsgBox(Translate(153), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'    If MsgBox("This will remove HijackThis' settings from the Registry " & _
'              "and exit. Note that you will have to delete the " & _
'              "HijackThis.exe file manually." & vbCrLf & vbCrLf & _
'              "Continue with uninstall?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe"
    RegDelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis"
    If Not RegKeyHasSubKeys(HKEY_LOCAL_MACHINE, "Software\TrendMicro") Then
        RegDelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro"
    End If
    CreateUninstallKey False
    DeleteBackup vbNullString
    Close
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
        RegSave "WinHeight", CStr(Me.Height)
        RegSave "WinWidth", CStr(Me.Width)
    End If
    If bDebugMode Then Close
    ToggleWow64FSRedirection True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleHeight < 5800 Then Exit Sub
    If Me.ScaleWidth < 6560 Then Exit Sub
    
    '== width ==
    ' - main -
    lstResults.Width = Me.ScaleWidth - 240
    shpBackground.Width = Me.ScaleWidth - 240
    shpMD5Background.Width = Me.ScaleWidth - 240
    lblMD5.Width = Me.ScaleWidth - 240
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
    fraProcessManager.Width = Me.ScaleWidth - 480
    lstProcessManager.Width = Me.ScaleWidth - 720
    lstProcManDLLs.Width = Me.ScaleWidth - 720
    fraHostsMan.Width = Me.ScaleWidth - 480
    lstHostsMan.Width = Me.ScaleWidth - 720
    chkProcManShowDLLs.Left = Me.ScaleWidth - 1815
    imgProcManSave.Left = Me.ScaleWidth - 2295
    imgProcManCopy.Left = Me.ScaleWidth - 2295 - 360
    fraADSSpy.Width = Me.ScaleWidth - 480
    lstADSSpyResults.Width = Me.ScaleWidth - 720
    fraADSSpyStatus.Width = Me.ScaleWidth - 720
    fraN00b.Width = Me.ScaleWidth - 195
    fraUninstMan.Width = Me.ScaleWidth - 480
    lstUninstMan.Width = Me.ScaleWidth - 4995
    lblInfo(8).Left = Me.ScaleWidth - 4770
    lblInfo(10).Left = Me.ScaleWidth - 4770
    txtUninstManName.Left = Me.ScaleWidth - 3210
    txtUninstManCmd.Left = Me.ScaleWidth - 3210
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
    txtHelp.Height = Me.ScaleHeight - 2115
    ' - config -
    fraConfig.Height = Me.ScaleHeight - 1755
    fraConfigTabs(0).Height = Me.ScaleHeight - 2805
    fraConfigTabs(1).Height = Me.ScaleHeight - 2805
    fraConfigTabs(2).Height = Me.ScaleHeight - 2805
    fraConfigTabs(3).Height = Me.ScaleHeight - 2805
    '(main)
    '(ignorelist)
    lstIgnore.Height = Me.ScaleHeight - 3615
    '(backups)
    lstBackups.Height = Me.ScaleHeight - 3615
    '(misc)
    fraProcessManager.Height = Me.ScaleHeight - 2805
    If chkProcManShowDLLs.Value = 0 Then
        lstProcessManager.Height = Me.ScaleHeight - 4035
    Else
        lstProcessManager.Height = (Me.ScaleHeight - 4035) / 2 - 120
        lblConfigInfo(21).Top = (Me.ScaleHeight - 4035) / 2 + 600 - 105
        lstProcManDLLs.Top = (Me.ScaleHeight - 4035) / 2 + 600 + 120
        lstProcManDLLs.Height = Me.ScaleHeight - 4035 - (Me.ScaleHeight - 4035) / 2 - 120
    End If
    cmdProcManKill.Top = Me.ScaleHeight - 3300
    cmdProcManRefresh.Top = Me.ScaleHeight - 3300
    cmdProcManRun.Top = Me.ScaleHeight - 3300
    cmdProcManBack.Top = Me.ScaleHeight - 3300
    lblProcManDblClick.Top = Me.ScaleHeight - 3300
    fraHostsMan.Height = Me.ScaleHeight - 2805
    lstHostsMan.Height = Me.ScaleHeight - 4035 - 240
    lblConfigInfo(15).Top = Me.ScaleHeight - 3300 - 300
    cmdHostsManDel.Top = Me.ScaleHeight - 3300
    cmdHostsManToggle.Top = Me.ScaleHeight - 3300
    cmdHostsManOpen.Top = Me.ScaleHeight - 3300
    cmdHostsManBack.Top = Me.ScaleHeight - 3300
    vscMiscTools.Height = fraConfigTabs(3).Height
    cmdADSSpyScan.Top = Me.ScaleHeight - 3315
    cmdADSSpySaveLog.Top = Me.ScaleHeight - 3315
    cmdADSSpyRemove.Top = Me.ScaleHeight - 3315
    cmdADSSpyBack.Top = Me.ScaleHeight - 3315
    lstADSSpyResults.Height = Me.ScaleHeight - 4875
    fraADSSpyStatus.Top = Me.ScaleHeight - 3585
    fraADSSpy.Height = Me.ScaleHeight - 2805
    fraN00b.Height = Me.ScaleHeight - 900
    fraUninstMan.Height = Me.ScaleHeight - 2805
    lstUninstMan.Height = Me.ScaleHeight - 3855 - 60
    cmdUninstManRefresh.Top = Me.ScaleHeight - 3315 - 60
    cmdUninstManSave.Top = Me.ScaleHeight - 3315 - 60
    cmdUninstManBack.Top = Me.ScaleHeight - 3315 - 60
    
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
    AddHorizontalScrollBarToResults
End Sub

Private Sub LoadSettings()
    On Error Resume Next
    chkAutoMark.Value = CInt(RegRead("AutoSelect", "0"))
    chkConfirm.Value = CInt(RegRead("Confirm", "1"))
    chkBackup.Value = CInt(RegRead("MakeBackup", "1"))
    chkIgnoreSafe.Value = CInt(RegRead("IgnoreSafe", "1"))
    chkLogProcesses.Value = CInt(RegRead("LogProcesses", "1"))
    chkShowN00bFrame.Value = CInt(RegRead("ShowIntroFrame", "1"))
    chkShowN00b.Value = CInt(RegRead("ShowIntroFrame", "1"))
    
    If RegValueExists(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "HijackThis startup scan") Then
        chkConfigStartupScan.Value = 1
    Else
        chkConfigStartupScan.Value = 0
    End If
    
    If bIgnoreAllWhitelists Then chkIgnoreSafe.Value = 0
    
    txtDefStartPage.Text = RegRead("DefStartPage", "http://www.msn.com/")
    txtDefSearchPage.Text = RegRead("DefSearchPage", "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch")
    txtDefSearchAss.Text = RegRead("DefSearchAss", "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm")
    txtDefSearchCust.Text = RegRead("DefSearchCust", "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm")
    
    bAutoSelect = IIf(chkAutoMark.Value = 1, True, False)
    bConfirm = IIf(chkConfirm.Value = 1, True, False)
    bMakeBackup = IIf(chkBackup.Value = 1, True, False)
    bIgnoreSafe = IIf(chkIgnoreSafe.Value = 1, True, False)
    bLogProcesses = IIf(chkLogProcesses.Value = 1, True, False)
    
    Dim i%
    On Error GoTo Error:
    'decrypt stuff
    For i = 0 To UBound(sRegVals)
        If sRegVals(i) = vbNullString Then Exit For
        sRegVals(i) = Crypt(sRegVals(i), sProgramVersion)
    Next i
    'For i = 0 To UBound(sFileVals)
    '    If sFileVals(i) = vbNullString Then Exit For
    '    sFileVals(i) = Crypt(sFileVals(i), sProgramVersion)
    'Next i
    
    For i = 0 To UBound(sRegVals)
        If sRegVals(i) = vbNullString Then Exit For
        sRegVals(i) = Replace(sRegVals(i), "$DEFSTARTPAGE", txtDefStartPage.Text)
        sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHPAGE", txtDefSearchPage.Text)
        sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHASS", txtDefSearchAss.Text)
        sRegVals(i) = Replace(sRegVals(i), "$DEFSEARCHCUST", txtDefSearchCust.Text)
        
        sRegVals(i) = Replace(sRegVals(i), "$WINSYSDIR", sWinSysDir)
    Next i
    For i = 0 To UBound(sFileVals)
        If sFileVals(i) = vbNullString Then Exit For
        sFileVals(i) = Replace(sFileVals(i), "$WINDIR", sWinDir)
    Next i
    
    're-encrypt stuff
    For i = 0 To UBound(sRegVals)
        If sRegVals(i) = vbNullString Then Exit For
        sRegVals(i) = Crypt(sRegVals(i), sProgramVersion, True)
    Next i
    'For i = 0 To UBound(sFileVals)
    '    If sFileVals(i) = vbNullString Then Exit For
    '    sFileVals(i) = Crypt(sFileVals(i), sProgramVersion, True)
    'Next i
        
    If Not RegKeyExists(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis") Then
        'first use, show moron warning
        'MsgBox Translate(3)
'        MsgBox "Warning!" & vbCrLf & vbCrLf & _
'               "Since HijackThis targets browser hijacking methods " & _
'               "instead of actual browser hijackers, entries may " & _
'               "appear in the scan list that are not hijackers. " & _
'               "Be careful what you delete, some system utilities " & _
'               "can cause problems if disabled." & vbCrLf & _
'               "For best results, ask spyware experts for help and " & _
'               "show them your scan log. They will advise you what " & _
'               "to fix and what to keep." & vbCrLf & vbCrLf & _
'               "Some adware-supported programs may cease to " & _
'               "function if the associated adware is removed.", vbExclamation
        
        RegCreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis"
    Else
        If RegGetString(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis", "WinWidth") = vbNullString Then
            'clear all previous settings
            RegDelKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis"
            RegCreateKey HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis"
        End If
    End If
    ''If Not RegKeyExists(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe") Then
    ''    RegCreateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe"
    ''    RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe", "", App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "hijackthis.exe"
    ''    RegSetStringVal HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe", "Path", App.Path
    ''Else
        If RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe", "") _
           <> App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "hijackthis.exe" Then
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe", "", App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "hijackthis.exe"
        End If
        If RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe", "Path") _
           <> App.Path Then
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\HijackThis.exe", "Path", App.Path
        End If
    ''End If
    ''CreateUninstallKey True
    Exit Sub
    
Error:
    ErrorMsg "frmMain_LoadSettings", Err.Number, Err.Description
End Sub

Private Function CreateLogFile$()
    Dim sLog$, i%, sProcessList$
    Dim hSnap&, uProcess As PROCESSENTRY32, sDummy$ '9x
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024) 'NT
    On Error GoTo MakeLog:
    If Not bLogProcesses Then GoTo MakeLog
        
    If Not bIsWinNT Then
        hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap < 1 Then
            'sProcessList = "(Unable to list running processes)" & vbCrLf
            sProcessList = "(" & Translate(28) & ")" & vbCrLf
            GoTo MakeLog
        End If
        
        uProcess.dwSize = Len(uProcess)
        If ProcessFirst(hSnap, uProcess) = 0 Then
            'sProcessList = "(Unable to list running processes)" & vbCrLf
            sProcessList = "(" & Translate(28) & ")" & vbCrLf
            GoTo MakeLog
        End If
        
        'sProcessList = "Running processes:" & vbCrLf
        sProcessList = Translate(29) & ":" & vbCrLf
        Do
            sDummy = Left(uProcess.szExeFile, InStr(uProcess.szExeFile, Chr(0)) - 1)
            sProcessList = sProcessList & sDummy & vbCrLf
        Loop Until ProcessNext(hSnap, uProcess) = 0
        CloseHandle hSnap
        sProcessList = sProcessList & vbCrLf
    Else
        On Error Resume Next
        If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
            sProcessList = "(" & Translate(28) & ")" & vbCrLf
            'sProcessList = "(Unable to list running processes)" & vbCrLf
            GoTo MakeLog
        End If
        On Error GoTo MakeLog:
        
        'sProcessList = "Running processes:" & vbCrLf
        sProcessList = Translate(29) & ":" & vbCrLf
        lNumProcesses = lNeeded / 4
        For i = 1 To lNumProcesses
            hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(i))
            If hProc <> 0 Then
                lNeeded = 0
                sProcessName = String(260, 0)
                If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                    GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                    sProcessName = Left(sProcessName, InStr(sProcessName, Chr(0)) - 1)
                    
                    'do some spell-checking
                    If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                    If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                    sProcessName = Replace(sProcessName, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
                    sProcessName = Replace(sProcessName, "SYSTEMROOT", sWinDir, , , vbTextCompare)
                    
                    sProcessList = sProcessList & sProcessName & vbCrLf
                End If
                CloseHandle hProc
            End If
        Next i
        sProcessList = sProcessList & vbCrLf
    End If
    
    '------------------------------
MakeLog:
    If Err Then
        sProcessList = "(" & Translate(28) & " (error#" & Err.Number & "))" & vbCrLf
        MsgBox Err.Description
    End If
    sLog = "Logfile of Trend Micro HijackThis v" & App.Major & "." & App.Minor & "." & App.Revision & IIf(bIsBeta, " (BETA)", vbNullString) & vbCrLf
    sLog = sLog & "Scan saved at " & Format(Time, "Long Time") & ", on " & Format(Date, "Short Date") & vbCrLf
    sLog = sLog & "Platform: " & GetWindowsVersion & vbCrLf
    'sLog = sLog & "Logged on as " & GetUser(bIsWinNT) & " to " & GetComputer & IIf(bIsWinNT, " (user is " & GetUserType & ")", vbNullString) & vbCrLf
    sLog = sLog & "MSIE: " & GetMSIEVersion & vbCrLf
    'GetChromeVersion
    sLog = sLog & GetChromeVersion() & vbCrLf
    sLog = sLog & GetChromeVersion64() & vbCrLf
    sLog = sLog & GetFirefoxVersion() & vbCrLf
   
    'sLog = sLog & "Spybot S&D version: " & GetSpybotVersion & vbCrLf
    'sLog = sLog & "Ad-Aware version: " & GetAdAwareVersion & vbCrLf
    sLog = sLog & "Boot mode: " & GetBootMode & vbCrLf

    If bLogEnvVars Then
        sLog = sLog & "Windows folder: " & sWinDir & vbCrLf & _
                      "System folder: " & sWinSysDir & vbCrLf & _
                      "Hosts file: " & sHostsFile & vbCrLf
    End If
    
    sLog = sLog & vbCrLf & sProcessList
    
    For i = 0 To lstResults.ListCount - 1
        sLog = sLog & lstResults.List(i) & vbCrLf
    Next i
    sLog = sLog & vbCrLf & "--" & vbCrLf & "End of file - xXxXx bytes"
    sLog = Replace(sLog, "xXxXx", Len(sLog), , , vbTextCompare)
    
    CreateLogFile = sLog
End Function

Private Sub imgMiscToolsDown_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    imgMiscToolsDown.Picture = imgMiscToolsDown2.Picture
End Sub

Private Sub imgMiscToolsDown_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    imgMiscToolsDown.Picture = imgMiscToolsDown1.Picture
    fraMiscToolsScroll.Top = fraConfigTabs(3).Height - fraMiscToolsScroll.Height
End Sub

Private Sub imgMiscToolsUp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    imgMiscToolsUp.Picture = imgMiscToolsUp2.Picture
    fraMiscToolsScroll.Top = 0
End Sub

Private Sub imgMiscToolsUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    imgMiscToolsUp.Picture = imgMiscToolsUp1.Picture
End Sub

Private Sub imgProcManCopy_Click()
    If chkProcManShowDLLs.Value = 1 Then
        CopyProcessList lstProcessManager, lstProcManDLLs, True
    Else
        CopyProcessList lstProcessManager, lstProcManDLLs, False
    End If
End Sub

Private Sub imgProcManSave_Click()
    If chkProcManShowDLLs.Value = 1 Then
        SaveProcessList lstProcessManager, lstProcManDLLs, True
    Else
        SaveProcessList lstProcessManager, lstProcManDLLs, False
    End If
End Sub

Private Sub lstADSSpyResults_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuADSSpy
End Sub

Private Sub lstProcessManager_Click()
    If lstProcManDLLs.Visible = False Then Exit Sub
    Dim s$
    s = lstProcessManager.List(lstProcessManager.ListIndex)
    s = Left(s, InStr(s, vbTab) - 1)
    If Not bIsWinNT Then
        RefreshDLLList CLng(s), lstProcManDLLs
    Else
        RefreshDLLListNT CLng(s), lstProcManDLLs
    End If
    lblConfigInfo(21).Caption = Translate(178) & " (" & lstProcManDLLs.ListCount & ")"
    'lblConfigInfo(21).Caption = "Loaded DLL libraries by selected process: (" & lstProcManDLLs.ListCount & ")"
End Sub

Private Sub lstProcessManager_DblClick()
    Dim s$
    s = lstProcessManager.List(lstProcessManager.ListIndex)
    s = Mid(s, InStr(s, vbTab) + 1)
    ShowFileProperties s
End Sub

Private Sub lstProcManDLLs_DblClick()
    Dim s$
    s = lstProcManDLLs.List(lstProcManDLLs.ListIndex)
    s = Mid(s, InStr(s, vbTab) + 1)
    ShowFileProperties s
End Sub

Private Sub lstUninstMan_Click()
    Dim sName$, sUninst$, sItems$(), i&
    sName = lstUninstMan.List(lstUninstMan.ListIndex)
    sItems = Split(RegEnumSubkeys(HKEY_LOCAL_MACHINE, sKeyUninstall), "|")
    If UBound(sItems) = -1 Then Exit Sub
    For i = 0 To UBound(sItems)
        If sName = RegGetString(HKEY_LOCAL_MACHINE, sKeyUninstall & "\" & sItems(i), "DisplayName") Then
            sUninst = RegGetString(HKEY_LOCAL_MACHINE, sKeyUninstall & "\" & sItems(i), "UninstallString")
            Exit For
        End If
    Next i
    
    txtUninstManName.Text = sName
    txtUninstManCmd.Text = sUninst
End Sub

Private Sub mnuADSSpySave_Click()
    cmdADSSpySaveLog_Click
End Sub

Private Sub mnuADSSpySelAll_Click()
    Dim i%
    If lstADSSpyResults.ListCount = 0 Then Exit Sub
    For i = 0 To lstADSSpyResults.ListCount - 1
        lstADSSpyResults.Selected(i) = True
    Next i
End Sub

Private Sub mnuADSSpySelInv_Click()
    Dim i%
    If lstADSSpyResults.ListCount = 0 Then Exit Sub
    For i = 0 To lstADSSpyResults.ListCount - 1
        lstADSSpyResults.Selected(i) = Not lstADSSpyResults.Selected(i)
    Next i
End Sub

Private Sub mnuADSSpySelNone_Click()
    Dim i%
    If lstADSSpyResults.ListCount = 0 Then Exit Sub
    For i = 0 To lstADSSpyResults.ListCount - 1
        lstADSSpyResults.Selected(i) = False
    Next i
End Sub

Private Sub picPaypal_Click()
    'ShellExecute Me.hwnd, "open", "http://www.merijn.org/donate.html", "", "", 1
End Sub

Private Sub vscMiscTools_Change()
    fraMiscToolsScroll.Top = -vscMiscTools.Value * (fraMiscToolsScroll.Height - fraConfigTabs(3).Height) / 100
    DoEvents
End Sub

Private Sub vscMiscTools_Scroll()
    Call vscMiscTools_Change
End Sub

Private Sub LoadLanguageList()
    Dim sFile$, sCurLang$, i%
    sCurLang = RegRead("LanguageFile")
    cboN00bLanguage.AddItem "(Default)"
    sFile = Dir(App.Path & "\*.lng")
    Do Until sFile = vbNullString
        If sFile = sCurLang Then i = cboN00bLanguage.ListCount
        cboN00bLanguage.AddItem sFile
        sFile = Dir
    Loop
    cboN00bLanguage.ListIndex = i
End Sub

