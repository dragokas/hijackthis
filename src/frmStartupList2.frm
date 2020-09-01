VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartupList2 
   Caption         =   "StartupList 2"
   ClientHeight    =   4812
   ClientLeft      =   168
   ClientTop       =   636
   ClientWidth     =   8844
   Icon            =   "frmStartupList2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4812
   ScaleWidth      =   8844
   Tag             =   "DesktopComponents"
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4092
      ScaleWidth      =   8808
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   8805
      Begin VB.Frame fraSave 
         Height          =   4095
         Left            =   0
         TabIndex        =   76
         Top             =   0
         Width           =   8775
         Begin VB.CommandButton cmdSaveCancel 
            Cancel          =   -1  'True
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5640
            TabIndex        =   64
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveOK 
            Caption         =   "OK"
            Default         =   -1  'True
            Height          =   375
            Left            =   7320
            TabIndex        =   65
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Frame fraSections 
            BorderStyle     =   0  'None
            Height          =   2535
            Left            =   120
            TabIndex        =   79
            Top             =   960
            Width           =   8175
            Begin VB.Frame fraScroller 
               BorderStyle     =   0  'None
               Height          =   8295
               Left            =   0
               TabIndex        =   80
               Top             =   -5760
               Width           =   8100
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Drivers32 libraries"
                  Height          =   255
                  Index           =   30
                  Left            =   4200
                  TabIndex        =   82
                  Tag             =   "Drivers32"
                  Top             =   1680
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionDisabled 
                  Caption         =   "Stopped/disabled services"
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   81
                  Tag             =   "StoppedServices"
                  Top             =   7800
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionHijack 
                  Caption         =   "Internet Explorer URLs"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   21
                  Tag             =   "IEURLs"
                  Top             =   5520
                  Width           =   2535
               End
               Begin VB.CheckBox chkSectionHardware 
                  Caption         =   "Items for other hardware cfgs"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   63
                  Tag             =   "Hardware"
                  Top             =   7920
                  Width           =   3900
               End
               Begin VB.CheckBox chkSectionUsers 
                  Caption         =   "Items for other users"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4200
                  TabIndex        =   62
                  Tag             =   "Users"
                  Top             =   7680
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionDisabled 
                  Caption         =   "Windows XP Security Center"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   30
                  Tag             =   "XPSecurity"
                  Top             =   8040
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionDisabled 
                  Caption         =   "Msconfig XP disabled items"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   29
                  Tag             =   "msconfigxp"
                  Top             =   7560
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionDisabled 
                  Caption         =   "Msconfig 9x/ME disabled items"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   28
                  Tag             =   "msconfig9x"
                  Top             =   7320
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionDisabled 
                  Caption         =   "Zones"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   27
                  Tag             =   "Zones"
                  Top             =   7080
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionDisabled 
                  Caption         =   "ActiveX kill bits"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   26
                  Tag             =   "Killbits"
                  Top             =   6840
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionDisabled 
                  Caption         =   "Hosts file items"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   25
                  Tag             =   "HostsFile"
                  Top             =   6600
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionDisabled 
                  Caption         =   "Disabled items, protection"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   24
                  Tag             =   "Disabled"
                  Top             =   6360
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionHijack 
                  Caption         =   "Hosts file path"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   23
                  Tag             =   "HostsFilePath"
                  Top             =   6000
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionHijack 
                  Caption         =   "Default URL prefixes"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   22
                  Tag             =   "URLPrefix"
                  Top             =   5760
                  Width           =   2535
               End
               Begin VB.CheckBox chkSectionHijack 
                  Caption         =   "Reset web settings URLs"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   20
                  Tag             =   "ResetWebSettings"
                  Top             =   5280
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionHijack 
                  Caption         =   "Hijack points"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   19
                  Tag             =   "Hijack"
                  Top             =   5040
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "ActiveX autoruns"
                  Height          =   255
                  Index           =   9
                  Left            =   120
                  TabIndex        =   16
                  Tag             =   "ActiveX"
                  Top             =   4080
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "ActiveX objects (DPFs)"
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   15
                  Tag             =   "DPFs"
                  Top             =   3840
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "IE bands"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   14
                  Tag             =   "IEBands"
                  Top             =   3600
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "IE menu extensions"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   13
                  Tag             =   "IEMenuExt"
                  Top             =   3360
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "IE Bars"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   12
                  Tag             =   "IEExplBars"
                  Top             =   3120
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "IE Buttons / Tools"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   11
                  Tag             =   "IEExtensions"
                  Top             =   2880
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "IE Toolbars"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   10
                  Tag             =   "IEToolbars"
                  Top             =   2640
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "Browser Helper Objects"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   9
                  Tag             =   "BHOs"
                  Top             =   2400
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "Internet Explorer items"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   8
                  Tag             =   "MSIE"
                  Top             =   2160
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "Desktop Components"
                  Height          =   255
                  Index           =   10
                  Left            =   120
                  TabIndex        =   17
                  Tag             =   "DesktopComponents"
                  Top             =   4320
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionMSIE 
                  Caption         =   "URLSearchHooks"
                  Height          =   255
                  Index           =   8
                  Left            =   120
                  TabIndex        =   18
                  Tag             =   "URLSearchHooks"
                  Top             =   4560
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "3rd party autostarts"
                  Height          =   255
                  Index           =   29
                  Left            =   4200
                  TabIndex        =   61
                  Tag             =   "3rdPartyApps"
                  Top             =   7200
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Utility Manager autostarts"
                  Height          =   255
                  Index           =   28
                  Left            =   4200
                  TabIndex        =   56
                  Tag             =   "UtilityManager"
                  Top             =   6000
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Command processor autostart"
                  Height          =   255
                  Index           =   27
                  Left            =   4200
                  TabIndex        =   35
                  Tag             =   "CmdProcAutorun"
                  Top             =   960
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Winsock LSPs"
                  Height          =   255
                  Index           =   26
                  Left            =   4200
                  TabIndex        =   59
                  Tag             =   "WinsockLSP"
                  Top             =   6720
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Application paths"
                  Height          =   255
                  Index           =   25
                  Left            =   4200
                  TabIndex        =   32
                  Tag             =   "AppPaths"
                  Top             =   240
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Security Providers"
                  Height          =   255
                  Index           =   24
                  Left            =   4200
                  TabIndex        =   49
                  Tag             =   "SecurityProviders"
                  Top             =   4560
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "MPRServices"
                  Height          =   255
                  Index           =   23
                  Left            =   4200
                  TabIndex        =   41
                  Tag             =   "MPRServices"
                  Top             =   2640
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "SharedTaskScheduler"
                  Height          =   255
                  Index           =   22
                  Left            =   4200
                  TabIndex        =   51
                  Tag             =   "SharedTaskScheduler"
                  Top             =   5040
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "ShellServiceObjectDelayLoad"
                  Height          =   255
                  Index           =   21
                  Left            =   4200
                  TabIndex        =   55
                  Tag             =   "ShellServiceObjectDelayLoad"
                  Top             =   5760
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "WOW"
                  Height          =   255
                  Index           =   20
                  Left            =   4200
                  TabIndex        =   60
                  Tag             =   "WOW"
                  Top             =   6960
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Protocol handers/filters"
                  Height          =   255
                  Index           =   19
                  Left            =   4200
                  TabIndex        =   45
                  Tag             =   "Protocols"
                  Top             =   3600
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Registry 'Run' subkeys"
                  Height          =   255
                  Index           =   18
                  Left            =   4200
                  TabIndex        =   48
                  Tag             =   "RunExRegkeys"
                  Top             =   4320
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Registry 'Run' keys"
                  Height          =   255
                  Index           =   17
                  Left            =   4200
                  TabIndex        =   47
                  Tag             =   "RunRegkeys"
                  Top             =   4080
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Approved shell extensions"
                  Height          =   255
                  Index           =   16
                  Left            =   4200
                  TabIndex        =   33
                  Tag             =   "ShellExts"
                  Top             =   480
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "ShellExecuteHooks"
                  Height          =   255
                  Index           =   15
                  Left            =   4200
                  TabIndex        =   53
                  Tag             =   "ShellExecuteHooks"
                  Top             =   5520
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "ColumnHandlers"
                  Height          =   255
                  Index           =   14
                  Left            =   4200
                  TabIndex        =   34
                  Tag             =   "ColumnHandlers"
                  Top             =   720
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "ContextMenuHandlers"
                  Height          =   255
                  Index           =   13
                  Left            =   4200
                  TabIndex        =   36
                  Tag             =   "ContextMenuHandlers"
                  Top             =   1200
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "ImageFileExecution"
                  Height          =   255
                  Index           =   12
                  Left            =   4200
                  TabIndex        =   38
                  Tag             =   "ImageFileExecution"
                  Top             =   1920
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Policies"
                  Height          =   255
                  Index           =   11
                  Left            =   4200
                  TabIndex        =   43
                  Tag             =   "Policies"
                  Top             =   3120
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "LSA packages"
                  Height          =   255
                  Index           =   10
                  Left            =   4200
                  TabIndex        =   39
                  Tag             =   "LsaPackages"
                  Top             =   2160
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Winlogon autoruns"
                  Height          =   255
                  Index           =   9
                  Left            =   4200
                  TabIndex        =   57
                  Tag             =   "WinLogonAutoruns"
                  Top             =   6240
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Print monitors"
                  Height          =   255
                  Index           =   8
                  Left            =   4200
                  TabIndex        =   44
                  Tag             =   "PrintMonitors"
                  Top             =   3360
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Driver filters"
                  Height          =   255
                  Index           =   7
                  Left            =   4200
                  TabIndex        =   37
                  Tag             =   "DriverFilters"
                  Top             =   1440
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Services"
                  Height          =   255
                  Index           =   6
                  Left            =   4200
                  TabIndex        =   50
                  Tag             =   "Services"
                  Top             =   4800
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Shell commands"
                  Height          =   255
                  Index           =   5
                  Left            =   4200
                  TabIndex        =   52
                  Tag             =   "ShellCommands"
                  Top             =   5280
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "On-reboot actions"
                  Height          =   255
                  Index           =   4
                  Left            =   4200
                  TabIndex        =   42
                  Tag             =   "OnRebootActions"
                  Top             =   2880
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "WinNT script policies"
                  Height          =   255
                  Index           =   3
                  Left            =   4200
                  TabIndex        =   58
                  Tag             =   "ScriptPolicies"
                  Top             =   6480
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Mountpoints"
                  Height          =   255
                  Index           =   2
                  Left            =   4200
                  TabIndex        =   40
                  Tag             =   "MountPoints"
                  Top             =   2400
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Registry-mapped .ini files"
                  Height          =   255
                  Index           =   1
                  Left            =   4200
                  TabIndex        =   46
                  Tag             =   "IniMapping"
                  Top             =   3840
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionRegistry 
                  Caption         =   "Registry items"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   4080
                  TabIndex        =   31
                  Tag             =   "Registry"
                  Top             =   0
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionFiles 
                  Caption         =   "Loaded/autoloading files"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   0
                  Tag             =   "Files"
                  Top             =   0
                  Width           =   3975
               End
               Begin VB.CheckBox chkSectionFiles 
                  Caption         =   "Explorer.exe clones"
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   7
                  Tag             =   "ExplorerClones"
                  Top             =   1680
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionFiles 
                  Caption         =   "Autostarting batch files"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   6
                  Tag             =   "BatFiles"
                  Top             =   1440
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionFiles 
                  Caption         =   "Autorun.inf files"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   5
                  Tag             =   "AutorunInfs"
                  Top             =   1200
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionFiles 
                  Caption         =   ".Ini file values"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   4
                  Tag             =   "IniFiles"
                  Top             =   960
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionFiles 
                  Caption         =   "Task Scheduler jobs"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   3
                  Tag             =   "TaskSchedulerJobs"
                  Top             =   720
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionFiles 
                  Caption         =   "Autostart folders"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   2
                  Tag             =   "AutoStartFolders"
                  Top             =   480
                  Width           =   3800
               End
               Begin VB.CheckBox chkSectionFiles 
                  Caption         =   "Running processes"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   1
                  Tag             =   "RunningProcesses"
                  Top             =   240
                  Width           =   3800
               End
            End
         End
         Begin VB.VScrollBar scrSaveSections 
            Height          =   2535
            LargeChange     =   1000
            Left            =   8280
            SmallChange     =   1000
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblInfo 
            Caption         =   $"frmStartupList2.frx":0442
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   7215
         End
      End
   End
   Begin VB.TextBox txtWarning 
      BackColor       =   &H8000000F&
      Height          =   1095
      Left            =   510
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   74
      Top             =   3240
      Visible         =   0   'False
      Width           =   6645
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh (F5)"
      Height          =   375
      Left            =   5760
      TabIndex        =   73
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picWarning 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   15
      Picture         =   "frmStartupList2.frx":0512
      ScaleHeight     =   492
      ScaleWidth      =   492
      TabIndex        =   72
      ToolTipText     =   "Click icon to close the warning box"
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort (Esc)"
      Height          =   495
      Left            =   5760
      TabIndex        =   71
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picHelp 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   15
      Picture         =   "frmStartupList2.frx":0DDC
      ScaleHeight     =   492
      ScaleWidth      =   492
      TabIndex        =   70
      Top             =   3255
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H8000000F&
      Height          =   1095
      Left            =   510
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   69
      Top             =   3240
      Visible         =   0   'False
      Width           =   6645
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   68
      Top             =   4560
      Width           =   8844
      _ExtentX        =   15600
      _ExtentY        =   445
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgbStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   67
      Top             =   4320
      Width           =   7215
      _ExtentX        =   12721
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   6480
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":16A6
            Key             =   "system"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":1C40
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":1FDA
            Key             =   "registry"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":2134
            Key             =   "memory"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":26CE
            Key             =   "exe"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":2828
            Key             =   "bat"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":2982
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":2ADC
            Key             =   "ini"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":2C36
            Key             =   "reg"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":2D90
            Key             =   "text"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":2EEA
            Key             =   "internet"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":3484
            Key             =   "msie"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":3A1E
            Key             =   "onreboot"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":3FB8
            Key             =   "run"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":4552
            Key             =   "help"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":46AC
            Key             =   "explorer"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":4806
            Key             =   "winlogon"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":4DA0
            Key             =   "cmd"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":533A
            Key             =   "good"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":58D4
            Key             =   "bad"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":5E6E
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":6408
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":69A2
            Key             =   "attn"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":6F3C
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":7096
            Key             =   "hardware"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":71F0
            Key             =   "icq"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":7595
            Key             =   "user"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":792F
            Key             =   "users"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":7CC9
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":7E23
            Key             =   "policy"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":81BD
            Key             =   "script"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":8557
            Key             =   "lsp"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":8AF1
            Key             =   "xpsec"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":908B
            Key             =   "mirc"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":91E5
            Key             =   "firewall"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":E9D7
            Key             =   "defender"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":EF71
            Key             =   "wintrust1"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":F50B
            Key             =   "wintrust2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":FAA5
            Key             =   "wintrust3"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartupList2.frx":1003F
            Key             =   "wintrust4"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   4575
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   7095
      _ExtentX        =   12510
      _ExtentY        =   8065
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwTriage 
      Height          =   3975
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11028
      _ExtentY        =   7006
      _Version        =   393217
      Indentation     =   0
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save as..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileCopy 
         Caption         =   "&Copy to clipboard"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileTriage 
         Caption         =   "Submit to &Triage!"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileTriageClose 
         Caption         =   "Close Triage &report"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileVerify 
         Caption         =   "Verify all file signatures"
      End
      Begin VB.Menu mnuFileStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuFind 
      Caption         =   "Fin&d"
      Begin VB.Menu mnuFindFind 
         Caption         =   "F&ind..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewExpand 
         Caption         =   "&Expand all"
      End
      Begin VB.Menu mnuViewCollapse 
         Caption         =   "&Collapse all"
      End
      Begin VB.Menu mnuViewStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsShowEmpty 
         Caption         =   "Show &empty sections"
      End
      Begin VB.Menu mnuOptionsShowCLSID 
         Caption         =   "Show &CLSIDs"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsShowCmts 
         Caption         =   "Show co&mments in .bat files"
      End
      Begin VB.Menu mnuOptionsShowPrivacy 
         Caption         =   "Show &privacy-related data"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsShowUsers 
         Caption         =   "Show other &users"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsShowHardware 
         Caption         =   "Show other h&ardware configurations"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsShowLargeHosts 
         Caption         =   "Show large hosts file (>1000 lines)"
      End
      Begin VB.Menu mnuOptionsShowLargeZones 
         Caption         =   "Show large Zones (>1000 domains)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpShow 
         Caption         =   "&Show help text"
         Checked         =   -1  'True
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWarning 
         Caption         =   "Show &warning log"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuHelpStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupShowFile 
         Caption         =   "Show this file"
      End
      Begin VB.Menu mnuPopupShowProp 
         Caption         =   "Show this file's properties"
      End
      Begin VB.Menu mnuPopupNotepad 
         Caption         =   "Send file to Notepad"
      End
      Begin VB.Menu mnuPopupFilenameCopy 
         Caption         =   "Copy filename to clipboard"
      End
      Begin VB.Menu mnuPopupStr4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupVerifyFile 
         Caption         =   "Verify file signature"
      End
      Begin VB.Menu mnuPopupFileRunScanner 
         Caption         =   "Lookup file on RunScanner.net..."
      End
      Begin VB.Menu mnuPopupCLSIDRunScanner 
         Caption         =   "Lookup CLSID on RunScanner.net"
      End
      Begin VB.Menu mnuPopupFileGoogle 
         Caption         =   "Lookup file on Google...."
      End
      Begin VB.Menu mnuPopupCLSIDGoogle 
         Caption         =   "Lookup CLSID on Google..."
      End
      Begin VB.Menu mnuPopupStr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupRegJump 
         Caption         =   "Regedit jump"
      End
      Begin VB.Menu mnuPopupRegkeyCopy 
         Caption         =   "Copy Registry key name to clipboard"
      End
      Begin VB.Menu mnuPopupStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupCopyNode 
         Caption         =   "Copy node text to clipboard"
      End
      Begin VB.Menu mnuPopupCopyPath 
         Caption         =   "Copy node path and text to clipboard"
      End
      Begin VB.Menu mnuPopupCopyTree 
         Caption         =   "Copy node and all subnodes to clipboard"
      End
      Begin VB.Menu mnuPopupSaveTree 
         Caption         =   "Save node and all subnodes as..."
      End
   End
End
Attribute VB_Name = "frmStartupList2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmStartupList2.frm]

'
' StartupList by Merijn Bellekom
'

' Fork by Dragokas

' frmStartupList form:
' --------------------
' Added clone: syswow64\explorer.exe
' Fixed: LSP Enum crash
' Fixed: GetNSProviderFile
' Added translation support, added Russian language

' modSartupList module:
' ---------------------
' WinTrustVerifyChildNodes. Fixed error with empty node
' istrusted.dll replaced by internal digital signature checking
' list of process replaced by function NtQuerySystemInformation

' Check 'frmMain.frm' to change version number

' v.2.12
' Improved Services ImagePath & DisplayName parsing
' 'desktop.ini' is whitelisted for autorun folders

' v.2.13
' Added error handling and tracing to every function

Option Explicit
'TODO
'* schermpje bij log save met secties selectie
'  - werkt nog niet voor secties bij andere users/hardware
'V nieuw item! HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Drivers32
'V Lookup file on Google in contextmenu
'V right-click op node: nodeisvalidfile moet ook met folders overweg kunnen
'V stukje voor geerts dll vast schrijven zover mogelijk
'V EnumHostsFilePaths for other hw configs
'V help beschikbaar bij log save secties
'V iereset.inf werkt niet
'V naamgeving in log en treeview gelijktrekken met log save secties
'V sectie volgorde in log en treeview aanpassen aan log save secties
'? DNS servers bij Hijack points
'? runscanner secties uitzoeken
'V fixed bug when 'find next' after refresh
'V replaced (most common) 8.3 filename occurrences
'V HKLM\System\CurrentControlSet\Control\Lsa\Authentication Packages
'V HKLM\System\CurrentControlSet\Control\Lsa\Security Packages
'V HKLM\System\CurrentControlSet\Control\Lsa\Notification Packages
'V fixed refresh/abort buttons not hiding/showing when refreshing
'V fixed no help text when enabling help right after scan
'V made lookup link to RunScanner.net from Geert Moernaut
'V optimized code
'V added all verbs to EnumShellCommands, added HKCU classes, HKUS classes
'V add bAbort to everything new
'V dingen als @xpsp2res.dll,-22019 kunnen omzetten naar strings
'V Desktop Components
'  HKCU\Software\Microsoft\Internet Explorer\Desktop\Components
'V windows xp firewall exception list
'  HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\AuthorizedApplications\List
'  HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List
'V SecurityProviders dlls
'  HKLM\System\CurrentControlSet\Control\SecurityProviders
'V autorun MountPoints (wtf zijn die CLSIDs? cd/dvd burners?)
'  HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints (Win 9x, Windows 2000)
'  HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2 (Windows XP)
'V App Paths hijack
'  HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\
'  == above: http://gladiator-antivirus.com/forum/index.php?showtopic=24610
'V Tasks: %windir%\system32\Tasks (Windows Vista)
'V %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Startup (Windows Vista)
'V %USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup (Windows Vista)
'  == above 3: http://www.silentrunners.org/sr_launchpoints.html ==
'-------- v2.02 ----------
'V appinit_dlls zit in Windows key, niet Winlogon
'V path voor log bij /autosave kunnen aangeven
'-------- v2.01 ----------
'V ICQ/mIRC help text
'V save node tree to disk
'V mirc voor 3rd party autoruns
'V winnt4 process enum werkt niet
'V help text voor: wininit.bak
'V root zones (my computer/local intranet) voor other users leeg
'V services wtf
'V zone 0 in EnumZones
'V view warning log
'V refresh knoppie e.d.
'V meer info voor error
'V progress voor large hosts/zones
'V mnuPopupCopyTree
'V skippen van grote hostsfile & zones, cmdline arguments
'V node right-click wat beter
'------- v2.00 --------
'V Checken of alle stuff te zien is bij bShowEmpty
'V Users Software \ DisabledEnums \ Zones is leeg, moet weg (win98)
'V Abort knopje tijdens scan/save log?
'V HKLM\SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal en Network (services)
'V HKLM\SYSTEM\CurrentControlSet\Control\SafeBoot,AlternateShell
'V VxD services voor andere hardware cfgs
'V Als wmi niet werkt -> geen usernames maar SIDs
'V Windows versions in modHelp.bas voor sections
'V Printer monitors
'V EnumXPSecurity voor andere users
'V EnumPolicies voor andere users
'X EnumZones: ZoneMap\Domains root value
'V Windows XP Security Center stuff:
'  SOFTWARE\Microsoft\Security Center
'  Software\Microsoft\Windows NT\CurrentVersion\systemrestore
'V fix bug in EnumZones when ZoneMap key is missing (HKCU/HKLM/HKUS)
'V fix Win2003 Small Biz Server being recognized as WinXP 64-bit (wtf?)
'V duplicate process/module entries in win9x
'V disable contextmenu items
'V dll modules loaded by running processes?
'V use marcin's code for regedit jump
'V registry jump - werkt soms niet ?
'V policies subkeys?
'V Help texts
'* Triage

Private Declare Function RegOpenKeyEx Lib "Advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyEx Lib "Advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "Advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "Advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long

'Private Const HKEY_CLASSES_ROOT = &H80000000
'Private Const HKEY_CURRENT_USER = &H80000001
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const HKEY_USERS = &H80000003

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7

Private NUM_OF_SECTIONS As Long
Private lCountedNodes& 'for GetStartupListReport()

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ProcessHotkey KeyCode, Me
End Sub

Private Sub chkSectionDisabled_Click(Index As Integer)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "chkSectionDisabled_Click - Begin"

    If chkSectionDisabled(0).Tag = "stop" Then Exit Sub
    If chkSectionDisabled(Index).Enabled = False Then Exit Sub
    Dim objCheck As CheckBox
    If Index = 0 Then
        For Each objCheck In chkSectionDisabled
            If objCheck.Index > 0 And chkSectionDisabled(objCheck.Index).Enabled Then
                chkSectionDisabled(0).Tag = "stop"
                chkSectionDisabled(objCheck.Index).Value = chkSectionDisabled(0).Value
                chkSectionDisabled(0).Tag = vbNullString
            End If
        Next objCheck
    Else
        chkSectionDisabled(0).Tag = "stop"
        chkSectionDisabled(0).Value = 0
        chkSectionDisabled(0).Tag = vbNullString
    End If
    If txtHelp.Visible Then
        If Index = 0 Then
            txtHelp.Text = GetHelpText("Disabled")
        Else
            txtHelp.Text = GetHelpText(chkSectionDisabled(Index).Tag)
        End If
    End If
    
    AppendErrorLogCustom "chkSectionDisabled_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "chkSectionDisabled_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub chkSectionFiles_Click(Index As Integer)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "chkSectionFiles_Click - Begin"

    If chkSectionFiles(0).Tag = "stop" Then Exit Sub
    If chkSectionFiles(Index).Enabled = False Then Exit Sub
    Dim objCheck As CheckBox
    If Index = 0 Then
        For Each objCheck In chkSectionFiles
            If objCheck.Index > 0 And chkSectionFiles(objCheck.Index).Enabled Then
                chkSectionFiles(0).Tag = "stop"
                chkSectionFiles(objCheck.Index).Value = chkSectionFiles(0).Value
                chkSectionFiles(0).Tag = vbNullString
            End If
        Next objCheck
    Else
        chkSectionFiles(0).Tag = "stop"
        chkSectionFiles(0).Value = 0
        chkSectionFiles(0).Tag = vbNullString
    End If
    If txtHelp.Visible Then
        If Index = 0 Then
            txtHelp.Text = GetHelpText("Files")
        Else
            txtHelp.Text = GetHelpText(chkSectionFiles(Index).Tag)
        End If
    End If
    AppendErrorLogCustom "chkSectionFiles_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "chkSectionFiles_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub chkSectionHardware_Click()
    If txtHelp.Visible Then txtHelp.Text = GetHelpText(chkSectionHardware.Tag)
End Sub

Private Sub chkSectionHijack_Click(Index As Integer)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "chkSectionHijack_Click - Begin"

    If chkSectionHijack(0).Tag = "stop" Then Exit Sub
    If chkSectionHijack(Index).Enabled = False Then Exit Sub
    Dim objCheck As CheckBox
    If Index = 0 Then
        For Each objCheck In chkSectionHijack
            If objCheck.Index > 0 And chkSectionHijack(objCheck.Index).Enabled Then
                chkSectionHijack(0).Tag = "stop"
                chkSectionHijack(objCheck.Index).Value = chkSectionHijack(0).Value
                chkSectionHijack(0).Tag = vbNullString
            End If
        Next objCheck
    Else
        chkSectionHijack(0).Tag = "stop"
        chkSectionHijack(0).Value = 0
        chkSectionHijack(0).Tag = vbNullString
    End If
    If txtHelp.Visible Then
        If Index = 0 Then
            txtHelp.Text = GetHelpText("Hijack")
        Else
            txtHelp.Text = GetHelpText(chkSectionHijack(Index).Tag)
        End If
    End If
    
    AppendErrorLogCustom "chkSectionHijack_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "chkSectionHijack_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub chkSectionMSIE_Click(Index As Integer)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "chkSectionMSIE_Click - Begin"

    If chkSectionMSIE(0).Tag = "stop" Then Exit Sub
    If chkSectionMSIE(Index).Enabled = False Then Exit Sub
    Dim objCheck As CheckBox
    If Index = 0 Then
        For Each objCheck In chkSectionMSIE
            If objCheck.Index > 0 And chkSectionMSIE(objCheck.Index).Enabled Then
                chkSectionMSIE(0).Tag = "stop"
                chkSectionMSIE(objCheck.Index).Value = chkSectionMSIE(0).Value
                chkSectionMSIE(0).Tag = vbNullString
            End If
        Next objCheck
    Else
        chkSectionMSIE(0).Tag = "stop"
        chkSectionMSIE(0).Value = 0
        chkSectionMSIE(0).Tag = vbNullString
    End If
    If txtHelp.Visible Then
        If Index = 0 Then
            txtHelp.Text = GetHelpText("MSIE")
        Else
            txtHelp.Text = GetHelpText(chkSectionMSIE(Index).Tag)
        End If
    End If
    
    AppendErrorLogCustom "chkSectionMSIE_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "chkSectionMSIE_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub chkSectionRegistry_Click(Index As Integer)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "chkSectionRegistry_Click - Begin"

    If chkSectionRegistry(0).Tag = "stop" Then Exit Sub
    If chkSectionRegistry(Index).Enabled = False Then Exit Sub
    Dim objCheck As CheckBox
    If Index = 0 Then
        For Each objCheck In chkSectionRegistry
            If objCheck.Index > 0 And chkSectionRegistry(objCheck.Index).Enabled Then
                chkSectionRegistry(0).Tag = "stop"
                chkSectionRegistry(objCheck.Index).Value = chkSectionRegistry(0).Value
                chkSectionRegistry(0).Tag = vbNullString
            End If
        Next objCheck
    Else
        chkSectionRegistry(0).Tag = "stop"
        chkSectionRegistry(0).Value = 0
        chkSectionRegistry(0).Tag = vbNullString
    End If
    If txtHelp.Visible Then
        If Index = 0 Then
            txtHelp.Text = GetHelpText("Registry")
        Else
            txtHelp.Text = GetHelpText(chkSectionRegistry(Index).Tag)
        End If
    End If
    AppendErrorLogCustom "chkSectionRegistry_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "chkSectionRegistry_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub chkSectionUsers_Click()
    If txtHelp.Visible Then txtHelp.Text = GetHelpText(chkSectionUsers.Tag)
End Sub

Private Sub cmdAbort_Click()
    'pgbStatus.Visible = Not pgbStatus.Visible
    'Form_Resize
    'Exit Sub
    
    bSL_Abort = True
    'cmdAbort.Enabled = False
    cmdAbort.Visible = False
    cmdRefresh.Visible = True
End Sub

Private Sub cmdRefresh_Click()
    mnuFindFind.Tag = vbNullString
    cmdRefresh.Visible = False
    GetAllEnums
    If bSL_Abort Then
        If bSL_Terminate Then
            bSL_Terminate = False
            Unload Me
        Else
            Status Translate(929): bSL_Abort = False
        End If
        Exit Sub
    End If
End Sub

Private Sub cmdSaveCancel_Click()
    picFrame.Visible = False
    tvwMain.Visible = True
End Sub

Private Sub cmdSaveOK_Click()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "cmdSaveOK_Click - Begin"

    Dim i%, L%, sTag$
    For i = 1 To chkSectionFiles.UBound
        If chkSectionFiles(i).Value = 1 Then
            sTag = chkSectionFiles(i).Tag
            tvwMain.Nodes(sTag).Tag = "1"
            For L = 0 To UBound(sUsernames)
                
            Next L
            For L = 1 To UBound(sHardwareCfgs)
                If NodeExists(sHardwareCfgs(L) & tvwMain.Nodes(sTag).Tag) Then
                    'tvwmain.Nodes(shard
                End If
            Next L
        Else
            If chkSectionFiles(i).Enabled Then
                tvwMain.Nodes(chkSectionFiles(i).Tag).Tag = "0"
            End If
        End If
    Next i
    For i = 1 To chkSectionMSIE.UBound
        If chkSectionMSIE(i).Value = 1 Then
            tvwMain.Nodes(chkSectionMSIE(i).Tag).Tag = "1"
        Else
            If chkSectionMSIE(i).Enabled Then
                tvwMain.Nodes(chkSectionMSIE(i).Tag).Tag = "0"
            End If
        End If
    Next i
    For i = 1 To chkSectionHijack.UBound
        If chkSectionHijack(i).Value = 1 Then
            tvwMain.Nodes(chkSectionHijack(i).Tag).Tag = "1"
        Else
            If chkSectionHijack(i).Enabled Then
                tvwMain.Nodes(chkSectionHijack(i).Tag).Tag = "0"
            End If
        End If
    Next i
    For i = 1 To chkSectionDisabled.UBound
        If chkSectionDisabled(i).Value = 1 Then
            tvwMain.Nodes(chkSectionDisabled(i).Tag).Tag = "1"
        Else
            If chkSectionDisabled(i).Enabled Then
                tvwMain.Nodes(chkSectionDisabled(i).Tag).Tag = "0"
            End If
        End If
    Next i
    For i = 1 To chkSectionRegistry.UBound
        If chkSectionRegistry(i).Value = 1 Then
            tvwMain.Nodes(chkSectionRegistry(i).Tag).Tag = "1"
        Else
            If chkSectionRegistry(i).Enabled Then
                tvwMain.Nodes(chkSectionRegistry(i).Tag).Tag = "0"
            End If
        End If
    Next i
    If chkSectionUsers.Value = 1 Then
        tvwMain.Nodes("Users").Tag = "1"
    Else
        If chkSectionUsers.Enabled Then
            tvwMain.Nodes("Users").Tag = "0"
        End If
    End If
    If chkSectionHardware.Value = 1 Then
        tvwMain.Nodes("Hardware").Tag = "1"
    Else
        If chkSectionHardware.Enabled Then
            tvwMain.Nodes("Hardware").Tag = "0"
        End If
    End If

    Dim sFile$, sLog$, hFile&
    '"Save file...", Text files, All files
    sFile = SaveFileDialog(Translate(900), AppPath(), "startuplist.txt", Translate(901) & " (*.txt)|*.txt|" & Translate(902) & " (*.*)|*.*", Me.hwnd)
    If sFile = vbNullString Then Exit Sub
    If Not (LCase$(Right$(sFile, 4)) = ".txt") Then sFile = sFile & ".txt"
    sLog = GetStartupListReport
    
    If OpenW(sFile, FOR_OVERWRITE_CREATE, hFile) Then
        PrintBOM hFile
        PutStringW hFile, , sLog
        CloseW hFile
    End If
    
    If bSL_Abort Then
        '"Generating of StartupList report was aborted!"
        Status Translate(903)
    Else
        If Err Then
            'The StartupList log could not be written to disk
            Status Translate(904) & ": " & Err.Description
        Else
            'The StartupList log has been written to disk
            Status Translate(905) & ". (" & Format$(Len(sLog) / 1024, "#,00") & " Kb)"
        End If
    End If
    picFrame.Visible = False
    tvwMain.Visible = True
    AppendErrorLogCustom "cmdSaveOK_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdSaveOK_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub LoadStrings()
    SEC_RUNNINGPROCESSES = Translate(2000)
    SEC_AUTOSTARTFOLDERS = Translate(2001)
    SEC_TASKSCHEDULER = Translate(2002)
    SEC_INIFILE = Translate(2003)
    SEC_AUTORUNINF = Translate(2004)
    SEC_BATFILES = Translate(2005)
    SEC_EXPLORERCLONES = Translate(2006)
    SEC_BHOS = Translate(2007)
    SEC_IETOOLBARS = Translate(2008)
    SEC_IEEXTENSIONS = Translate(2009)
    SEC_IEBARS = Translate(2010)
    SEC_IEMENUEXT = Translate(2011)
    SEC_IEBANDS = Translate(2012)
    SEC_DPFS = Translate(2013)
    SEC_ACTIVEX = Translate(2014)
    SEC_DESKTOPCOMPONENTS = Translate(2015)
    SEC_URLSEARCHHOOKS = Translate(2016)
    SEC_APPPATHS = Translate(2017)
    SEC_SHELLEXT = Translate(2018)
    SEC_COLUMNHANDLERS = Translate(2019)
    SEC_CMDPROC = Translate(2020)
    SEC_CONTEXTMENUHANDLERS = Translate(2021)
    SEC_DRIVERFILTERS = Translate(2022)
    SEC_DRIVERS32 = Translate(2023)
    SEC_IMAGEFILEEXECUTION = Translate(2024)
    SEC_LSAPACKAGES = Translate(2025)
    SEC_MOUNTPOINTS = Translate(2026)
    SEC_MPRSERVICES = Translate(2027)
    SEC_ONREBOOT = Translate(2028)
    SEC_POLICIES = Translate(2029)
    SEC_PRINTMONITORS = Translate(2030)
    SEC_PROTOCOLS = Translate(2031)
    SEC_INIMAPPING = Translate(2032)
    SEC_REGRUNKEYS = Translate(2033)
    SEC_REGRUNEXKEYS = Translate(2034)
    SEC_SECURITYPROVIDERS = Translate(2035)
    SEC_SERVICES = Translate(2036)
    SEC_SHAREDTASKSCHEDULER = Translate(2037)
    SEC_SHELLCOMMANDS = Translate(2038)
    SEC_SHELLEXECUTEHOOKS = Translate(2039)
    SEC_SSODL = Translate(2040)
    SEC_UTILMANAGER = Translate(2041)
    SEC_WINLOGON = Translate(2042)
    SEC_SCRIPTPOLICIES = Translate(2043)
    SEC_WINSOCKLSP = Translate(2044)
    SEC_WOW = Translate(2045)
    SEC_3RDPARTY = Translate(2046)
    SEC_RESETWEBSETTINGS = Translate(2047)
    SEC_IEURLS = Translate(2048)
    SEC_URLPREFIX = Translate(2049)
    SEC_HOSTSFILEPATH = Translate(2050)
    SEC_HOSTSFILE = Translate(2051)
    SEC_KILLBITS = Translate(2052)
    SEC_ZONES = Translate(2053)
    SEC_MSCONFIG9X = Translate(2054)
    SEC_MSCONFIGXP = Translate(2055)
    SEC_STOPPEDSERVICES = Translate(2056)
    SEC_XPSECURITY = Translate(2057)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "Form_Load - Begin"

    Dim hFile As Long, sPath As String
    
    If bSL_Abort Then Exit Sub
    
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    ReloadLanguage True
    LoadStrings

'    If App.PrevInstance Then
'        msgboxw "StartupList is already running.", vbExclamation
'        End
'    End If
    If IsRunningInIDE Or InStr(1, Command$, "/debug", 1) > 0 Then bDebug = True
    'If InStr(1, Command$, "/debug") > 0 Then bDebug = True
    
    'InitCommonControls
    
    '" - checking " & NUM_OF_SECTIONS & " autostart groups!"
    'Me.Caption = "StartupList v." & StartupListVer & " fork" & _
    '    Replace$(" - " & Translate(906), "[]", NUM_OF_SECTIONS)

    NUM_OF_SECTIONS = StartupList_UpdateCaption(Me)

    lEnumBufLen = 16400

    tvwMain.LineStyle = tvwRootLines
    tvwMain.LabelEdit = tvwManual
    tvwMain.ImageList = imlMain
    tvwTriage.LineStyle = tvwRootLines
    tvwTriage.LabelEdit = tvwManual
    tvwTriage.ImageList = imlMain
    bShowCLSIDs = True
    bShowPrivacy = True
    bShowUsers = True
    bShowHardware = True
    fraScroller.Top = 0
    
    If InStr(1, Command$, "/showempty", vbTextCompare) > 0 Then
        bShowEmpty = True
        mnuOptionsShowEmpty.Checked = True
    End If
    If InStr(1, Command$, "/noclsids", vbTextCompare) > 0 Then
        bShowCLSIDs = False
        mnuOptionsShowCLSID.Checked = False
    End If
    If InStr(1, Command$, "/noshowprivate", vbTextCompare) > 0 Then
        bShowPrivacy = False
        mnuOptionsShowPrivacy.Checked = False
    End If
    If InStr(1, Command$, "/showcmts", vbTextCompare) > 0 Then
        bShowCmts = True
        mnuOptionsShowCmts.Checked = True
    End If
    If InStr(1, Command$, "/nousers", vbTextCompare) > 0 Then
        bShowUsers = False
        mnuOptionsShowUsers.Checked = False
    End If
    If InStr(1, Command$, "/nohardware", vbTextCompare) > 0 Then
        bShowHardware = False
        mnuOptionsShowHardware.Checked = False
    End If
    If InStr(1, Command$, "/showlargehostsfile", vbTextCompare) > 0 Then
        bShowLargeHosts = True
        mnuOptionsShowLargeHosts.Checked = True
    End If
    If InStr(1, Command$, "/showlargezones", vbTextCompare) > 0 Then
        bShowLargeZones = True
        mnuOptionsShowLargeZones.Checked = True
    End If
    
    GetWindowsVersion
    
    If InStr(1, Command$, "/autosave", vbTextCompare) > 0 Then
        'get everything, save and exit
        bAutoSave = True
        Me.Hide
        If InStr(1, Command$, "/autosavepath:", vbTextCompare) > 0 Then
            'path to save logfile to
            sAutoSavePath = Mid$(Command$, InStr(1, Command$, "/autosavepath:", 1) + 14)
            If Left$(sAutoSavePath, 1) = """" Then
                'path enclosed in quotes, get what's between
                sAutoSavePath = Mid$(sAutoSavePath, 2)
                If InStr(sAutoSavePath, """") > 0 Then
                    sAutoSavePath = Left$(sAutoSavePath, InStr(sAutoSavePath, """") - 1)
                Else
                    'no closing quote
                    sAutoSavePath = vbNullString
                End If
            Else
                'path has no quotes, stop at first space
                If InStr(sAutoSavePath, " ") > 0 Then
                    sAutoSavePath = Left$(sAutoSavePath, InStr(sAutoSavePath, " ") - 1)
                End If
            End If
        End If
        'check if path exists
        sPath = GetParentDir(sAutoSavePath)
        If Not FolderExists(sPath) Then MkDirW sPath
    End If
    
    If Not bAutoSave And Not Me.WindowState = vbMinimized Then
    
        If Not LoadWindowPos(Me, SETTINGS_SECTION_STARTUPLIST) Then
    
            'center and size window
            If Screen.Width < 1024 * 15 Then
                Me.Width = Screen.Width * 0.6
                Me.Height = Screen.Height * 0.8
            Else
                Me.Width = 600 * 15
                Me.Height = 600 * 15
            End If
            Me.Left = (Screen.Width - Me.Width) \ 2
            Me.Top = (Screen.Height - Me.Height) \ 2
        End If
        
    End If
    
    If bShowUsers Or bShowPrivacy Then GetUsernames
    If bShowHardware Then GetHardwareCfgs
    
    LoadSectionNames
    GetAllEnums
    If bSL_Abort Then
        If bSL_Terminate Then
            bSL_Terminate = False
            Unload Me
        Else
            Status Translate(929): bSL_Abort = False
        End If
        Exit Sub
    End If
    
    If bAutoSave Then
        If OpenW(BuildPath(IIf(Len(sAutoSavePath) <> 0, sAutoSavePath, AppPath()), "startuplist.txt"), FOR_OVERWRITE_CREATE, hFile) Then
            PrintBOM hFile
            PutStringW hFile, , GetStartupListReport
            CloseW hFile
        End If
        '//TODO: check this
        Terminate_HJT
    End If
    
    mnuHelpShow_Click
    txtHelp.Text = Translate(600)
    
    AppendErrorLogCustom "Form_Load - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Form_Load"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub GetAllEnums()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetAllEnums - Begin"

    Dim lTicks&
    If bDebug Then lTicks = GetTickCount()
    tvwMain.Nodes.Clear
    'mnuFile.Enabled = False
    mnuFileSave.Enabled = False
    mnuFileCopy.Enabled = False
    mnuFileVerify.Enabled = False
    mnuViewRefresh.Enabled = False
    mnuOptions.Enabled = False
    mnuHelpShow.Checked = False
    txtHelp.Visible = False
    picHelp.Visible = False
    mnuHelpWarning.Checked = False
    txtWarning.Visible = False
    picWarning.Visible = False
    Form_Resize
    
    If txtWarning.Text <> vbNullString Then
        'Use the Options
        txtWarning.Text = Left$(txtWarning.Text, InStr(txtWarning.Text, Translate(907)) - 3)
    End If
    
    cmdAbort.Enabled = True
    cmdAbort.Visible = True
    bSL_Abort = False
    
'    If Not bDebug And Not IsRunningInIDE Then
'        On Error Resume Next
'    End If
    
    If Not bAutoSave Then Me.Show
    'Loading...
    Status Translate(909)
    pgbStatus.Max = NUM_OF_SECTIONS
    pgbStatus.Value = 0
    pgbStatus.Visible = True
    Form_Resize

    If bShowPrivacy Then
        '[*user*] on [*computer*]
        tvwMain.Nodes.Add , tvwFirst, "System", _
            Replace$(Replace$(Translate(926), "[*user*]", "'" & OSver.UserName & "'"), "[*computer*]", "'" & OSver.ComputerName & "'") & ", " & GetWindowsVersion, _
            "system", "system"
    Else
        tvwMain.Nodes.Add , tvwFirst, "System", GetWindowsVersion, "system", "system"
    End If
    tvwMain.Nodes("System").Expanded = True
    
    Dim i%, sName$
    If bShowUsers Then
        'Loading... usernames
        Status Translate(910)
        'Other users on this computer
        tvwMain.Nodes.Add , tvwFirst, "Users", Translate(927), "system"
        tvwMain.Nodes("Users").Expanded = True
        For i = 0 To UBound(sUsernames)
            sName = MapSIDToUsername(sUsernames(i))
            If sName <> OSver.UserName And sName <> vbNullString Then
                If bShowPrivacy Then
                    tvwMain.Nodes.Add "Users", tvwChild, "Users" & sUsernames(i), sName, "user"
                Else
                    tvwMain.Nodes.Add "Users", tvwChild, "Users" & sUsernames(i), sUsernames(i), "user"
                End If
            End If
        Next i
    End If
    If bShowHardware Then
        'Loading... hardware configurations
        Status Translate(911)
        'Other hardware configurations
        tvwMain.Nodes.Add , tvwFirst, "Hardware", Translate(928), "system"
        tvwMain.Nodes("Hardware").Expanded = True
        For i = 1 To UBound(sHardwareCfgs)
            sName = MapControlSetToHardwareCfg(sHardwareCfgs(i))
            tvwMain.Nodes.Add "Hardware", tvwChild, "Hardware" & sHardwareCfgs(i), sName, "system"
        Next i
    End If
    pgbStatus.Value = 1
    
    AppendErrorLogCustom "EnumProcesses"
    
    'running processes
    Status Translate(909) & " " & SEC_RUNNINGPROCESSES
    DoTicks tvwMain
    EnumProcesses
    DoTicks tvwMain, "RunningProcesses"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumAutoStartFolders"
    
    'startup folders in start menu etc
    Status Translate(909) & " " & SEC_AUTOSTARTFOLDERS
    DoTicks tvwMain
    EnumAutoStartFolders
    DoTicks tvwMain, "AutoStartFolders"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumJobs"
    
    'Task Scheduler jobs
    Status Translate(909) & " " & SEC_TASKSCHEDULER
    DoTicks tvwMain
    EnumJobs
    DoTicks tvwMain, "TaskSchedulerJobs"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumIniFiles"
    
    'autoload entries from ini files, shell
    Status Translate(909) & " " & SEC_INIFILE
    DoTicks tvwMain
    EnumIniFiles
    DoTicks tvwMain, "IniFiles"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumAutorunInf"
    
    'autorun.inf op alle schijven
    Status Translate(909) & " " & SEC_AUTORUNINF
    DoTicks tvwMain
    EnumAutorunInf
    DoTicks tvwMain, "AutorunInfs"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumBatAutostartFiles"
    
    'autoexec.bat, winstart.bat, dosstart.bat
    Status Translate(909) & " " & SEC_BATFILES
    DoTicks tvwMain
    EnumBatAutostartFiles
    DoTicks tvwMain, "BatFiles"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumExplorerClones"
    
    'Explorer clones
    Status Translate(909) & " " & SEC_EXPLORERCLONES
    DoTicks tvwMain
    EnumExplorerClones
    DoTicks tvwMain, "ExplorerClones"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumBHOs"
    
    'Browser Helper Objects
    Status Translate(909) & " " & SEC_BHOS
    DoTicks tvwMain
    EnumBHOs
    DoTicks tvwMain, "BHOs"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumIEToolbars"
    
    'IE Toolbars
    Status Translate(909) & " " & SEC_IETOOLBARS
    DoTicks tvwMain
    EnumIEToolbars
    DoTicks tvwMain, "IEToolbars"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumIEExtensions"
    
    'IE Extensions
    Status Translate(909) & " " & SEC_IEEXTENSIONS
    DoTicks tvwMain
    EnumIEExtensions
    DoTicks tvwMain, "IEExtensions"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumIEExplBars"
    
    'IE Explorer Bars
    Status Translate(909) & " " & SEC_IEBARS
    DoTicks tvwMain
    EnumIEExplBars
    DoTicks tvwMain, "IEExplBars"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumIEMenuExt"
    
    'IE MenuExt
    Status Translate(909) & " " & SEC_IEMENUEXT
    DoTicks tvwMain
    EnumIEMenuExt
    DoTicks tvwMain, "IEMenuExt"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumIEBands"
    
    'IE bands
    Status Translate(909) & " " & SEC_IEBANDS
    DoTicks tvwMain
    EnumIEBands
    DoTicks tvwMain, "IEBands"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumDPFs"
    
    'Downloaded Program Files
    Status Translate(909) & " " & SEC_DPFS
    DoTicks tvwMain
    EnumDPFs
    DoTicks tvwMain, "DPFs"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumActiveXAutoruns"
    
    'ActiveSetup\StubPath autoruns
    Status Translate(909) & " " & SEC_ACTIVEX
    DoTicks tvwMain
    EnumActiveXAutoruns
    DoTicks tvwMain, "ActiveX"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumDesktopComponents"
    
    'Desktop Components
    Status Translate(909) & " " & SEC_DESKTOPCOMPONENTS
    DoTicks tvwMain
    EnumDesktopComponents
    DoTicks tvwMain, "DesktopComponents"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumURLSearchHooks"
    
    'HK..\..\IE\URLSearchHooks
    Status Translate(909) & " " & SEC_URLSEARCHHOOKS
    DoTicks tvwMain
    EnumURLSearchHooks
    DoTicks tvwMain, "URLSearchHooks"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumIniMappingKeys"
    
    'ini file values, mapped to the registry in NT
    Status Translate(909) & " " & SEC_INIMAPPING
    DoTicks tvwMain
    EnumIniMappingKeys
    DoTicks tvwMain, "IniMapping"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumMountPoints"
    
    'MountPoints
    Status Translate(909) & " " & SEC_MOUNTPOINTS
    DoTicks tvwMain
    EnumMountPoints
    DoTicks tvwMain, "MountPoints"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumNTScripts"
    
    'NT scripts
    Status Translate(909) & " " & SEC_SCRIPTPOLICIES
    DoTicks tvwMain
    EnumNTScripts
    DoTicks tvwMain, "ScriptPolicies"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumOnRebootActions"
    
    'wininit.ini/.bak, PendingFileRenameOperations
    Status Translate(909) & " " & SEC_ONREBOOT
    DoTicks tvwMain
    EnumOnRebootActions
    DoTicks tvwMain, "OnRebootActions"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumShellCommands"
    
    'shell commands for .exe, bat, com, pif, etc
    Status Translate(909) & " " & SEC_SHELLCOMMANDS
    DoTicks tvwMain
    EnumShellCommands
    DoTicks tvwMain, "ShellCommands"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumServices"
    
    'NT Services + 9x device drivers
    Status Translate(909) & " " & SEC_SERVICES
    DoTicks tvwMain
    EnumServices
    DoTicks tvwMain, "Services"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumDriverFilters"
    
    'Driver filters
    Status Translate(909) & " " & SEC_DRIVERFILTERS
    DoTicks tvwMain
    EnumDriverFilters
    DoTicks tvwMain, "DriverFilters"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "RegEnumDrivers32"
    
    'Drivers32
    Status Translate(909) & " " & SEC_DRIVERS32
    DoTicks tvwMain
    RegEnumDrivers32 tvwMain
    DoTicks tvwMain, "Drivers32"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumPrintMonitors"
    
    'Print Monitors
    Status Translate(909) & " " & SEC_PRINTMONITORS
    DoTicks tvwMain
    EnumPrintMonitors
    DoTicks tvwMain, "PrintMonitors"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumWinLogonAutoruns"
    
    'Winlogon autoruns
    Status Translate(909) & " " & SEC_WINLOGON
    DoTicks tvwMain
    EnumWinLogonAutoruns
    DoTicks tvwMain, "WinLogonAutoruns"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumLSAPackages"
    
    'LSA packages (security, notification, authentication)
    Status Translate(909) & " " & SEC_LSAPACKAGES
    DoTicks tvwMain
    EnumLSAPackages
    DoTicks tvwMain, "LsaPackages"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumPolicies"
    
    'policies
    Status Translate(909) & " " & SEC_POLICIES
    DoTicks tvwMain
    EnumPolicies
    DoTicks tvwMain, "Policies"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumImageFileExecution"
    
    'Image File Execution
    Status Translate(909) & " " & SEC_IMAGEFILEEXECUTION
    DoTicks tvwMain
    EnumImageFileExecution
    DoTicks tvwMain, "ImageFileExecution"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumContextMenuHandlers"
    
    'HKCR\*\shellex\ContextMenuHandlers
    Status Translate(909) & " " & SEC_CONTEXTMENUHANDLERS
    DoTicks tvwMain
    EnumContextMenuHandlers
    DoTicks tvwMain, "ContextMenuHandlers"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumColumnHandlers"
    
    'HKCR\*\shellex\ColumnHandlers
    Status Translate(909) & " " & SEC_COLUMNHANDLERS
    DoTicks tvwMain
    EnumColumnHandlers
    DoTicks tvwMain, "ColumnHandlers"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumShellExecuteHooks"
    
    'HKLM\..\explorer\ShellExecuteHooks
    Status Translate(909) & " " & SEC_SHELLEXECUTEHOOKS
    DoTicks tvwMain
    EnumShellExecuteHooks
    DoTicks tvwMain, "ShellExecuteHooks"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumShellExtensions"
    
    'HKLM\..\Shell Extensions\Approved
    Status Translate(909) & " " & SEC_SHELLEXT
    DoTicks tvwMain
    EnumShellExtensions
    DoTicks tvwMain, "ShellExts"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumRunRegKeys"
    
    'all Run/RunOnce/RunServices/etc regkeys
    Status Translate(909) & " " & SEC_REGRUNKEYS
    DoTicks tvwMain
    EnumRunRegKeys
    DoTicks tvwMain, "RunRegkeys"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumRunExRegKeys"
    
    'same, RunEx keys
    Status Translate(909) & " " & SEC_REGRUNEXKEYS
    DoTicks tvwMain
    EnumRunExRegKeys
    DoTicks tvwMain, "RunExRegkeys"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumProtocols"
    
    'HKCR\Protocols\Filter + \Handler
    Status Translate(909) & " " & SEC_PROTOCOLS
    DoTicks tvwMain
    EnumProtocols
    DoTicks tvwMain, "Protocols"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumAccUtilManager"
    
    'Accessibility\Utility Manager autoruns
    Status Translate(909) & " " & SEC_UTILMANAGER
    DoTicks tvwMain
    EnumAccUtilManager
    DoTicks tvwMain, "UtilityManager"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumWOW"
    
    'WOW
    Status Translate(909) & " " & SEC_WOW
    DoTicks tvwMain
    EnumWOW
    DoTicks tvwMain, "WOW"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumSSODL"
    
    'SSODL
    Status Translate(909) & " " & SEC_SSODL
    DoTicks tvwMain
    EnumSSODL
    DoTicks tvwMain, "ShellServiceObjectDelayLoad"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumSharedTaskScheduler"
    
    'STS
    Status Translate(909) & " " & SEC_SHAREDTASKSCHEDULER
    DoTicks tvwMain
    EnumSharedTaskScheduler
    DoTicks tvwMain, "SharedTaskScheduler"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumMPRServices"
    
    'MPRServices
    Status Translate(909) & " " & SEC_MPRSERVICES
    DoTicks tvwMain
    EnumMPRServices
    DoTicks tvwMain, "MPRServices"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumSecurityProviders"
    
    'SecurityProviders
    Status Translate(909) & " " & SEC_SECURITYPROVIDERS
    DoTicks tvwMain
    EnumSecurityProviders
    DoTicks tvwMain, "SecurityProviders"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumAppPaths"
    
    'App Paths
    Status Translate(909) & " " & SEC_APPPATHS
    DoTicks tvwMain
    EnumAppPaths
    DoTicks tvwMain, "AppPaths"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumCmdProcessorAutorun"
    
    'Command Processor Autostart
    Status Translate(909) & " " & SEC_CMDPROC
    DoTicks tvwMain
    EnumCmdProcessorAutorun
    DoTicks tvwMain, "CmdProcAutorun"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumLSP"
    
    'Winsock LSP
    Status Translate(909) & " " & SEC_WINSOCKLSP
    DoTicks tvwMain
    EnumLSP
    DoTicks tvwMain, "WinsockLSP"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "Enum3rdPartyAutostarts"
    
    '3rd party autostarts, e.g. icq
    Status Translate(909) & " " & SEC_3RDPARTY
    DoTicks tvwMain
    Enum3rdPartyAutostarts
    DoTicks tvwMain, "3rdPartyApps"
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumHijack"
    
    'Hijack points
    EnumHijack
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "EnumDisabled"
    
    'Disabled stuff
    EnumDisabled
    UpdateProgressBar
    If bSL_Abort Then Exit Sub
    
    AppendErrorLogCustom "Removing empty users/hardware nodes"
    
    '-----------------------------------------
    'remove empty users/hardware nodes
    If bShowUsers Then
        For i = 0 To UBound(sUsernames)
            If NodeExists("Users" & sUsernames(i)) Then
                If tvwMain.Nodes("Users" & sUsernames(i)).Children = 0 And Not bShowEmpty Then
                    tvwMain.Nodes.Remove "Users" & sUsernames(i)
                End If
            End If
        Next i
    End If
    If bShowHardware Then
        For i = 1 To UBound(sHardwareCfgs)
            If NodeExists("Hardware" & sHardwareCfgs(i)) Then
                If tvwMain.Nodes("Hardware" & sHardwareCfgs(i)).Children = 0 And Not bShowEmpty Then
                    tvwMain.Nodes.Remove "Hardware" & sHardwareCfgs(i)
                End If
            End If
        Next i
    End If
    
    tvwMain.Nodes("System").Expanded = True
    'Status "Ready."
    Status Translate(974)
    UpdateProgressBar
    
    pgbStatus.Visible = False
    Form_Resize
    'mnuFile.Enabled = True
    mnuFileSave.Enabled = True
    mnuFileCopy.Enabled = True
    mnuFileVerify.Enabled = True
    mnuViewRefresh.Enabled = True
    mnuOptions.Enabled = True
    cmdAbort.Visible = False
    
    '"Use the Options menu to " & _
    '"override skipped items." & vbCrLf & _
    '"Click Help > Show warnings to close this message."
    If picWarning.Visible Then
        txtWarning.Text = txtWarning.Text & vbCrLf & Translate(908)
    End If
    
    'Aborted!
    If bSL_Abort Then Exit Sub
    
    If bDebug Then
        tvwMain.Nodes.Add "System", tvwChild, "SystemTicks", " Time: " & GetTickCount - lTicks & " ms", "clock"
    End If
    
    AppendErrorLogCustom "GetAllEnums - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetAllEnums"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    tvwMain.Width = Me.ScaleWidth
    'tvwTriage.Width = Me.ScaleWidth
    cmdAbort.Left = Me.ScaleWidth - 1440 - 120
    cmdRefresh.Left = Me.ScaleWidth - 1440 - 120
    txtHelp.Width = Me.ScaleWidth - 510
    txtWarning.Width = Me.ScaleWidth - 510
    pgbStatus.Width = Me.ScaleWidth
    picFrame.Width = Me.ScaleWidth
    fraSave.Width = picFrame.Width
    
    If txtHelp.Visible Or txtWarning.Visible Then
        picFrame.Height = Me.ScaleHeight - stbStatus.Height - 1125
        If pgbStatus.Visible Then
            pgbStatus.Top = Me.ScaleHeight - 525
            txtHelp.Top = Me.ScaleHeight - 1650
            txtWarning.Top = Me.ScaleHeight - 1650
            picHelp.Top = Me.ScaleHeight - 1635
            picWarning.Top = Me.ScaleHeight - 1635
            tvwMain.Height = Me.ScaleHeight - 1710
            'tvwTriage.Height = Me.ScaleHeight - 1710
            cmdAbort.Top = Me.ScaleHeight - 2295
            cmdRefresh.Top = Me.ScaleHeight - 2295
        Else
            txtHelp.Top = Me.ScaleHeight - 1365
            txtWarning.Top = Me.ScaleHeight - 1365
            picHelp.Top = Me.ScaleHeight - 1365
            picWarning.Top = Me.ScaleHeight - 1365
            tvwMain.Height = Me.ScaleHeight - 1425
            'tvwTriage.Height = Me.ScaleHeight - 1425
            cmdAbort.Top = Me.ScaleHeight - 1995
            cmdRefresh.Top = Me.ScaleHeight - 1995
        End If
    Else
        picFrame.Height = Me.ScaleHeight - stbStatus.Height - 30
        If pgbStatus.Visible Then
            pgbStatus.Top = Me.ScaleHeight - 525
            tvwMain.Height = Me.ScaleHeight - 555
            'tvwTriage.Height = Me.ScaleHeight - 555
            cmdAbort.Top = Me.ScaleHeight - 1200
            cmdRefresh.Top = Me.ScaleHeight - 1200
        Else
            tvwMain.Height = Me.ScaleHeight - 300
            'tvwTriage.Height = Me.ScaleHeight - 300
            cmdAbort.Top = Me.ScaleHeight - 900
            cmdRefresh.Top = Me.ScaleHeight - 900
        End If
    End If
    cmdSaveCancel.Top = picFrame.Height - cmdSaveCancel.Height - 120
    cmdSaveOK.Top = picFrame.Height - cmdSaveOK.Height - 120
    fraSections.Height = picFrame.Height - 1500 - 60
    fraSave.Height = picFrame.Height
    scrSaveSections.Height = fraSections.Height
    scrSaveSections.Max = fraScroller.Height - fraSections.Height
    scrSaveSections.Visible = IIf(scrSaveSections.Max > 0, True, False)
End Sub

Private Sub mnuFileCopy_Click()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "mnuFileCopy_Click - Begin"
    Dim sLog$
    If tvwMain.Visible Then
        mnuFileCopy.Enabled = False
        sLog = GetStartupListReport
        If Not bSL_Abort Then
            ClipboardSetText sLog
            'The StartupList report has been copied to your clipboard.
            Status Translate(930) & " (" & Format$(Len(sLog) / 1024, "#,00") & " Kb)"
        Else
            bSL_Abort = False
            'Generating of StartupList report was aborted!
            Status Translate(931)
        End If
        mnuFileCopy.Enabled = True
    Else
        sLog = GetTriageReport
        'Clipboard.Clear
        'ClipboardSetText sLog
        'Status "The Triage report has been copied to your clipboard."
    End If
    AppendErrorLogCustom "mnuFileCopy_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "mnuFileCopy_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "mnuFileSave_Click - Begin"
    
    tvwMain.Visible = False
    picFrame.Visible = True
    scrSaveSections.Value = 0
    
    Dim i%, sAllChecked As Boolean
    i = 1
    sAllChecked = True
    chkSectionFiles(0).Value = 1
    On Error Resume Next
    Do
        If chkSectionFiles(i).Caption <> vbNullString Then
            If NodeExists(chkSectionFiles(i).Tag) Then
                chkSectionFiles(i).Enabled = True
                chkSectionFiles(i).Value = 1
            Else
                chkSectionFiles(i).Enabled = False
                chkSectionFiles(i).Value = 0
                sAllChecked = False
            End If
        End If
        i = i + 1
    Loop Until Err
    If sAllChecked Then chkSectionFiles(0).Value = 1
    
    i = 1
    sAllChecked = True
    chkSectionMSIE(0).Value = 1
    Err.Clear
    Do
        If chkSectionMSIE(i).Caption <> vbNullString Then
            If NodeExists(chkSectionMSIE(i).Tag) Then
                chkSectionMSIE(i).Enabled = True
                chkSectionMSIE(i).Value = 1
            Else
                chkSectionMSIE(i).Enabled = False
                chkSectionMSIE(i).Value = 0
                sAllChecked = False
            End If
        End If
        i = i + 1
    Loop Until Err
    If sAllChecked Then chkSectionMSIE(0).Value = 1
    
    i = 1
    sAllChecked = True
    chkSectionHijack(0).Value = 1
    Err.Clear
    Do
        If chkSectionHijack(i).Caption <> vbNullString Then
            If NodeExists(chkSectionHijack(i).Tag) Then
                chkSectionHijack(i).Enabled = True
                chkSectionHijack(i).Value = 1
            Else
                chkSectionHijack(i).Enabled = False
                chkSectionHijack(i).Value = 0
                sAllChecked = False
            End If
        End If
        i = i + 1
    Loop Until Err
    If sAllChecked Then chkSectionHijack(0).Value = 1
    
    i = 1
    sAllChecked = True
    chkSectionDisabled(0).Value = 1
    Err.Clear
    Do
        If chkSectionDisabled(i).Caption <> vbNullString Then
            If NodeExists(chkSectionDisabled(i).Tag) Then
                chkSectionDisabled(i).Enabled = True
                chkSectionDisabled(i).Value = 1
            Else
                chkSectionDisabled(i).Enabled = False
                chkSectionDisabled(i).Value = 0
                sAllChecked = False
            End If
        End If
        i = i + 1
    Loop Until Err
    If sAllChecked Then chkSectionDisabled(0).Value = 1
    
    i = 1
    sAllChecked = True
    chkSectionRegistry(0).Value = 1
    Err.Clear
    Do
        If chkSectionRegistry(i).Caption <> vbNullString Then
            If NodeExists(chkSectionRegistry(i).Tag) Then
                chkSectionRegistry(i).Enabled = True
                chkSectionRegistry(i).Value = 1
            Else
                chkSectionRegistry(i).Enabled = False
                chkSectionRegistry(i).Value = 0
                sAllChecked = False
            End If
        End If
        i = i + 1
    Loop Until Err
    If sAllChecked Then chkSectionRegistry(0).Value = 1
    
    If NodeExists("Users") Then
        chkSectionUsers.Value = 1
        chkSectionUsers.Enabled = True
    Else
        chkSectionUsers.Value = 0
        chkSectionUsers.Enabled = False
    End If
    If NodeExists("Hardware") Then
        chkSectionHardware.Value = 1
        chkSectionHardware.Enabled = True
    Else
        chkSectionHardware.Value = 0
        chkSectionHardware.Enabled = False
    End If
    AppendErrorLogCustom "mnuFileSave_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "mnuFileSave_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub mnuFileTriage_Click()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "mnuFileTriage_Click - Begin"
    Dim i&, j&, Section As Node, Subsection As Node
    Dim sName$, sFile$, sCLSID$, sDummy$()
    Dim sMsg$
    'This will send your StartupList report to the XBlock " & _
    '"Triage server for live analysis.
    sMsg = Translate(932)
    If mnuOptionsShowCLSID.Checked = False Then
        '"To get a more accurate Triage " & _
               "report, it is recommended to turn on class IDs before sending " & _
               "(this option is available from the Options menu). It is not " & _
               "enabled now." & vbCrLf & vbCrLf & _
               "Enable class IDs and continue?"
        sMsg = sMsg & vbCrLf & Translate(932)
    Else
        'Continue?
        sMsg = sMsg & vbCrLf & vbCrLf & Translate(933)
    End If
    If MsgBoxW(sMsg, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If mnuOptionsShowCLSID.Checked = False Then mnuOptionsShowCLSID_Click
    DoEvents
    'Creating Triage report to send...
    Status Translate(934)
    tvwTriage.Nodes.Clear
    tvwTriage.Nodes.Add , tvwFirst, "Triage", "Triage by XBlock.com - www.spywareguide.com", "system"
        
'    EnumProcesses
    Set Section = tvwMain.Nodes("RunningProcesses")
    If Section.Children > 0 Then
        'Running processes
        tvwTriage.Nodes.Add "Triage", tvwChild, "RunningProcesses", Translate(935), "memory"
        For i = Section.Index + 1 To Section.Children + Section.Index
            sFile = tvwMain.Nodes(i).Text
            AddTriageObj "RunningProcesses" & i - 1 - Section.Index, "Process", sFile
        Next i
    End If
'    EnumAutoStartFolders
    '[not useful, unless .exe?]
'    EnumIniFiles
    '[submit .exe MD5]
'    EnumIniMappingKeys
    '[submit .exe MD5]
'    EnumRunRegKeys
    Set Section = tvwMain.Nodes("RunRegkeys")
    If Section.Children > 0 Then
        'Registry 'Run' keys
        tvwTriage.Nodes.Add "Triage", tvwChild, "RunRegKeys", Translate(936), "registry"
        For i = Section.Index + 1 To Section.Children + Section.Index
            Set Subsection = tvwMain.Nodes(i)
            If Subsection.Children > 0 Then
                tvwTriage.Nodes.Add "RunRegKeys", tvwChild, Subsection.Key, Subsection.Text, "registry"
                For j = Subsection.Index + 1 To Subsection.Children + Subsection.Index
                    If InStr(tvwMain.Nodes(j).Text, " = ") > 0 Then
                        sName = Left$(tvwMain.Nodes(j).Text, InStr(tvwMain.Nodes(j).Text, " = ") - 1)
                        sFile = Mid$(tvwMain.Nodes(j).Text, InStr(tvwMain.Nodes(j).Text, " = ") + 3)
                        AddTriageObj tvwMain.Nodes(j).Key, "Registry value", sFile
                    End If
                Next j
            End If
        Next i
    End If
'    EnumRunExRegKeys
'    EnumPolicyAutoruns
'    EnumBatAutostartFiles
'    EnumOnRebootActions
'    EnumShellCommands
'    EnumUserProgramAutostarts
'    EnumActiveXAutoruns
    Set Section = tvwMain.Nodes("ActiveX")
    If Section.Children > 0 Then
        'ActiveX Autoruns
        tvwTriage.Nodes.Add "Triage", tvwChild, "ActiveX", Translate(937), "msie"
        For i = Section.Index + 1 To Section.Children + Section.Index
            sDummy = Split(tvwMain.Nodes(i).Text, " - ")
            If UBound(sDummy) = 1 Or UBound(sDummy) = 2 Then
                sName = sDummy(0)
                sCLSID = sDummy(1)
                If InStr(sCLSID, "{") > 0 And InStr(sCLSID, "}") > 0 Then
                    sCLSID = Mid$(sName, InStr(sCLSID, "{"))
                    sCLSID = Mid$(sCLSID, 1, InStr(sCLSID, "}") + 1)
                End If
                sFile = sDummy(1)
                sFile = GuessFullpathFromAutorun(sFile)
                AddTriageObj "ActiveX" & (i - 1 - Section.Index), "ActiveX Object", sFile, sCLSID
            End If
        Next i
    End If
    
'    EnumProtocols
'    EnumExplorerClones
    Set Section = tvwMain.Nodes("ExplorerClones")
    If Section.Children > 0 Then
        'Explorer clones
        tvwTriage.Nodes.Add "Triage", tvwChild, "ExplorerClones", Translate(938), "explorer"
        For i = Section.Index + 1 To Section.Children + Section.Index
            sFile = tvwMain.Nodes(i).Text
            AddTriageObj tvwMain.Nodes(i).Key, "File", sFile
        Next i
    End If
'    EnumServices
'    EnumLSP
'    EnumWinLogonAutoruns
'    EnumNTScripts
'    EnumBHOs
    Set Section = tvwMain.Nodes("BHOs")
    If Section.Children > 0 Then
        'Browser Helper Objects
        tvwTriage.Nodes.Add "Triage", tvwChild, "BHOs", Translate(939), "msie"
        For i = Section.Index + 1 To Section.Children + Section.Index
            sDummy = Split(tvwMain.Nodes(i).Text, " = ")
            If UBound(sDummy) = 2 Then
                'sName = sDummy(0)
                sCLSID = sDummy(1)
                sFile = sDummy(2)
            Else
                'sName = sDummy(0)
                sCLSID = vbNullString
                sFile = sDummy(1)
            End If
            
            AddTriageObj "BHO" & (i - 1 - Section.Index), "BHO", sFile, sCLSID
        Next i
    End If
'    EnumImageFileExecution
'    EnumContextMenuHandlers
'    EnumShellExecuteHooks
'    EnumAccUtilManager
'    EnumJobs
'    EnumWOW
'    EnumCmdProcessorAutorun
'    EnumSSODL
    Set Section = tvwMain.Nodes("ShellServiceObjectDelayLoad")
    If Section.Children > 0 Then
        'ShellServiceObjectDelayLoad
        tvwTriage.Nodes.Add "Triage", tvwChild, "ShellServiceObjectDelayLoad", Translate(940), "registry"
        For i = Section.Index + 1 To Section.Children + Section.Index
            sDummy = Split(tvwMain.Nodes(i).Text, " = ")
            If UBound(sDummy) = 2 Then
                sName = sDummy(0)
                sCLSID = sDummy(1)
                sFile = sDummy(2)
            Else
                sName = sDummy(0)
                sFile = sDummy(1)
            End If
            AddTriageObj tvwMain.Nodes(i).Key, "DLL", sFile, sCLSID
        Next i
    End If
'    EnumSharedTaskScheduler
    Set Section = tvwMain.Nodes("SharedTaskScheduler")
    If Section.Children > 0 Then
        'SharedTaskScheduler
        tvwTriage.Nodes.Add "Triage", tvwChild, "SharedTaskScheduler", Translate(941), "registry"
        For i = Section.Index + 1 To Section.Children + Section.Index
            sDummy = Split(tvwMain.Nodes(i).Text, " = ")
            If UBound(sDummy) = 2 Then
                sName = sDummy(0)
                sCLSID = sDummy(1)
                sFile = sDummy(2)
            Else
                sName = sDummy(0)
                sFile = sDummy(1)
            End If
            AddTriageObj tvwMain.Nodes(i).Key, "DLL", sFile, sCLSID
        Next i
    End If
'    EnumMPRServices
    
    '--------------------------------------------------------------------------
    Set Section = Nothing
    'TRIAGERESULT|[id]|[1/2/3]|[descr]|[url]
    'OK
    Dim sTriage$(), SID$, sRet$, sDesc$, sURL$, sParent$
    'Sending Triage report...
    Status Translate(942)
    'tvwTriage.Text = replace$(GetTriage, vbLf, vbCrLf)
    sDummy = Split(GetTriage, vbLf)
    
    For i = 0 To UBound(sDummy)
        sTriage = Split(sDummy(i), "|")
        If UBound(sTriage) > 0 Then
            If sTriage(0) = "TRIAGERESULT" Then
                SID = sTriage(1)
                sRet = sTriage(2)
                sDesc = sTriage(3)
                sURL = sTriage(4)
                sParent = tvwMain.Nodes(SID).Parent.Key
                Select Case sRet
                    Case 1 'unknown
                        tvwTriage.Nodes.Add sParent, tvwChild, SID, tvwMain.Nodes(SID).Text, "unknown"
                    Case 2 'good
                        tvwTriage.Nodes.Add sParent, tvwChild, SID, tvwMain.Nodes(SID).Text, "good"
                        tvwTriage.Nodes.Add SID, tvwChild, SID & "info", sDesc & IIf(sURL <> vbNullString, " (" & sURL & ")", vbNullString), "good"
                    Case 3 'bad
                        tvwTriage.Nodes.Add sParent, tvwChild, SID, tvwMain.Nodes(SID).Text, "bad"
                        tvwTriage.Nodes.Add SID, tvwChild, SID & "info", sDesc & IIf(sURL <> vbNullString, " (" & sURL & ")", vbNullString), "bad"
                End Select
            End If
        End If
    Next i
    tvwMain.Visible = False
    tvwTriage.Visible = True
    tvwTriage.Nodes("Triage").Expanded = True
    mnuFileTriageClose.Enabled = True
    'Ready.
    Status Translate(943)
    AppendErrorLogCustom "mnuFileTriage_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "mnuFileTriage_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub mnuFileTriageClose_Click()
    tvwTriage.Visible = False
    tvwTriage.Nodes.Clear
    tvwMain.Visible = True
    mnuFileTriageClose.Enabled = False
End Sub


Private Sub mnuFileVerify_Click()
    bSL_Abort = False
    cmdAbort.Visible = True
    WinTrustVerifyChildNodes "System"
    If NodeExists("Users") Then WinTrustVerifyChildNodes "Users"
    If NodeExists("Hardware") Then WinTrustVerifyChildNodes "Hardware"
    If bSL_Abort Then
        'Verification aborted.
        Status Translate(944)
    Else
        'Verification done.
        Status Translate(945)
    End If
    bSL_Abort = False
End Sub

Private Sub mnuFindFind_Click()
    Dim sFind$
    sFind = mnuFindFind.Tag
    'Enter a filename, word or phrase to look for:, "Find..."
    sFind = InputBox(Translate(946), Translate(947), sFind)
    If sFind = vbNullString Then Exit Sub
    
    mnuFindFind.Tag = sFind
    tvwMain.SelectedItem = tvwMain.Nodes("System")
    mnuFindNext_Click
End Sub

Private Sub mnuFindNext_Click()
    Dim sFind$
    sFind = mnuFindFind.Tag
    If sFind = vbNullString Then
        mnuFindFind_Click
        Exit Sub
    End If
    
    Dim iFirst&, i&
    iFirst = tvwMain.SelectedItem.Index + 1
    For i = iFirst To tvwMain.Nodes.Count
        If InStr(1, tvwMain.Nodes(i).Text, sFind, vbTextCompare) > 0 Then
            tvwMain.SelectedItem = tvwMain.Nodes(i)
            Exit For
        End If
    Next i
    If i = tvwMain.Nodes.Count + 1 Then
        'No further hits beyond this point.
        MsgBoxW Translate(955), vbInformation
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    Dim sMsg$
    '"StartupList" & vbCrLf & _
           "Written by Merijn Bellekom - http://www.merijn.org/" & vbCrLf & vbCrLf & _
           "Based on StartupList v1, TonyKlein's Collection of Autostart " & _
           "Locations and Andrew Aronoff's SilentRunners" & vbCrLf & vbCrLf & _
           "Thanks also to:" & vbCrLf & _
           "Mosaic1, Philip Sloss, Gkweb, Dmitry Sokolov, Oleg Lembievskiy" & vbCrLf & vbCrLf & _
           "Note: StartupList does not and cannot change anything on the system." & _
           vbCrLf & vbCrLf & "If you find this program useful, please donate!" & _
           vbCrLf & "http://www.merijn.org/donate.php"
    sMsg = Translate(948)
    MsgBoxW sMsg, vbInformation
End Sub

Private Sub mnuHelpShow_Click()
    mnuHelpShow.Checked = Not mnuHelpShow.Checked
    txtHelp.Visible = mnuHelpShow.Checked
    picHelp.Visible = mnuHelpShow.Checked
    
    mnuHelpWarning.Checked = False
    picWarning.Visible = False
    txtWarning.Visible = False
    Form_Resize
    On Error Resume Next
    If tvwMain.Visible And tvwMain.Enabled Then
        tvwMain.SetFocus
        tvwMain.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub mnuHelpWarning_Click()
    mnuHelpWarning.Checked = Not mnuHelpWarning.Checked
    txtWarning.Visible = mnuHelpWarning.Checked
    picWarning.Visible = mnuHelpWarning.Checked
    
    mnuHelpShow.Checked = False
    picHelp.Visible = False
    txtHelp.Visible = False
    Form_Resize
    On Error Resume Next
    If tvwMain.Visible And tvwMain.Enabled Then
        tvwMain.SetFocus
    End If
    tvwMain_MouseUp 1, 0, 0, 0
End Sub

Private Sub mnuOptionsShowHardware_Click()
    mnuOptionsShowHardware.Checked = Not mnuOptionsShowHardware.Checked
    bShowHardware = CBool(mnuOptionsShowHardware.Checked)
    cmdRefresh.Visible = True
End Sub

Private Sub mnuOptionsShowLargeHosts_Click()
    bShowLargeHosts = Not bShowLargeHosts
    mnuOptionsShowLargeHosts.Checked = Not mnuOptionsShowLargeHosts.Checked
    cmdRefresh.Visible = True
End Sub

Private Sub mnuOptionsShowLargeZones_Click()
    bShowLargeZones = Not bShowLargeZones
    mnuOptionsShowLargeZones.Checked = Not mnuOptionsShowLargeZones.Checked
    cmdRefresh.Visible = True
End Sub

Private Sub mnuOptionsShowPrivacy_Click()
    mnuOptionsShowPrivacy.Checked = Not mnuOptionsShowPrivacy.Checked
    bShowPrivacy = CBool(mnuOptionsShowPrivacy.Checked)
    cmdRefresh.Visible = True
End Sub

Private Sub mnuOptionsShowUsers_Click()
    mnuOptionsShowUsers.Checked = Not mnuOptionsShowUsers.Checked
    bShowUsers = CBool(mnuOptionsShowUsers.Checked)
    cmdRefresh.Visible = True
End Sub

Private Sub mnuPopupCopyNode_Click()
    ClipboardSetText tvwMain.SelectedItem.Text
    'Node text copied to clipboard.
    Status Translate(949)
End Sub

Private Sub mnuPopupCopyPath_Click()
    'This Computer
    ClipboardSetText Replace$(tvwMain.SelectedItem.FullPath, tvwMain.Nodes("System").Text, Translate(956))
    'Node path & text copied to clipboard.
    Status Translate(950)
End Sub

Private Sub mnuPopupCopyTree_Click()
    Dim sReport$
    pgbStatus.Visible = True
    pgbStatus.Value = 0
    pgbStatus.Max = tvwMain.Nodes.Count
    Form_Resize
    lCountedNodes = 1
    sReport = GetNodeChildren(tvwMain.SelectedItem.Key, 4)
    pgbStatus.Visible = False
    Form_Resize
    
    '" partial report" & vbCrLf & _
              "Root node was '" & tvwMain.SelectedItem.Text & "'" & vbCrLf & _
              "Full path to root node: "
    sReport = Me.Caption & Replace$(Translate(954), "[]", "'" & tvwMain.SelectedItem.Text & "'") & _
        " " & Replace$(tvwMain.SelectedItem.FullPath, tvwMain.Nodes("System").Text, Translate(956)) & vbCrLf & _
              sReport
    ClipboardSetText sReport
    'Node tree copied to clipboard.
    Status Translate(951)
End Sub

Private Sub mnuPopupFilenameCopy_Click()
    If tvwMain.SelectedItem.Tag <> vbNullString Then
        ClipboardSetText tvwMain.SelectedItem.Tag
        'Filename was copied to the clipboard.
        Status Translate(952)
    End If
End Sub

Private Sub mnuPopupNotepad_Click()
    SendToNotepad tvwMain.SelectedItem.Tag
End Sub

Private Sub mnuPopupRegJump_Click()
    If InStr(1, tvwMain.SelectedItem.Tag, "HKEY_") <> 1 Then
        'selected item is not a regkey but a file - climb up in the
        'tree until we find a regkey
        Dim MyNode As Node
        Set MyNode = tvwMain.SelectedItem
        Do Until MyNode = tvwMain.Nodes("System") Or _
                 MyNode = tvwMain.Nodes("Users") Or _
                 MyNode = tvwMain.Nodes("Hardware")
            Set MyNode = MyNode.Parent
            If InStr(1, MyNode.Tag, "HKEY_") = 1 Then
                Reg.Jump 0, MyNode.Tag
                Exit Sub
            End If
        Loop
    Else
        Reg.Jump 0, tvwMain.SelectedItem.Tag
    End If
End Sub

Private Sub mnuPopupRegkeyCopy_Click()
    If InStr(1, tvwMain.SelectedItem.Tag, "HKEY_") <> 1 Then
        'selected item is not a regkey but a file - climb up in the
        'tree until we find a regkey
        Dim MyNode As Node
        Set MyNode = tvwMain.SelectedItem
        Do Until MyNode = tvwMain.Nodes("System") Or _
                 MyNode = tvwMain.Nodes("Users") Or _
                 MyNode = tvwMain.Nodes("Hardware")
            Set MyNode = MyNode.Parent
            If InStr(1, MyNode.Tag, "HKEY_") = 1 Then
                ClipboardSetText MyNode.Tag
                'Registry key name was copied to the clipboard.
                Status Translate(953)
                Exit Sub
            End If
        Loop
    Else
        ClipboardSetText tvwMain.SelectedItem.Tag
        'Registry key name was copied to the clipboard.
        Status Translate(953)
    End If
End Sub

Private Sub mnuPopupFileRunScanner_Click()
    RunScannerGetMD5 tvwMain.SelectedItem.Tag, tvwMain.SelectedItem.Key
End Sub

Private Sub mnuPopupFileGoogle_Click()
    If Trim$(tvwMain.SelectedItem.Tag) <> vbNullString Then
        'ShellRun "https://www.google.com/search?q=" & Mid$(tvwMain.SelectedItem.Tag, InStrRev(tvwMain.SelectedItem.Tag, "\") + 1)
        OpenURL "https://www.google.com/search?q=" & Mid$(tvwMain.SelectedItem.Tag, InStrRev(tvwMain.SelectedItem.Tag, "\") + 1)
    End If
End Sub

Private Sub mnuPopupCLSIDRunScanner_Click()
    Dim sCLSID$
    If InStr(tvwMain.SelectedItem.Text, "{") > 0 And InStr(tvwMain.SelectedItem.Text, "}") > 0 Then
        sCLSID = Mid$(tvwMain.SelectedItem.Text, InStr(tvwMain.SelectedItem.Text, "{"))
        sCLSID = Left$(sCLSID, InStr(sCLSID, "}"))
    ElseIf InStr(tvwMain.SelectedItem.Tag, "{") > 0 And InStr(tvwMain.SelectedItem.Tag, "}") > 0 Then
        sCLSID = Mid$(tvwMain.SelectedItem.Tag, InStr(tvwMain.SelectedItem.Tag, "{"))
        sCLSID = Left$(sCLSID, InStr(sCLSID, "}") + 1)
    End If
    If isCLSID(sCLSID) Then RunScannerGetCLSID sCLSID, tvwMain.SelectedItem.Key
End Sub

Private Sub mnuPopupCLSIDGoogle_Click()
    Dim sCLSID$
    If InStr(tvwMain.SelectedItem.Text, "{") > 0 And InStr(tvwMain.SelectedItem.Text, "}") > 0 Then
        sCLSID = Mid$(tvwMain.SelectedItem.Text, InStr(tvwMain.SelectedItem.Text, "{"))
        sCLSID = Left$(sCLSID, InStr(sCLSID, "}"))
    ElseIf InStr(tvwMain.SelectedItem.Tag, "{") > 0 And InStr(tvwMain.SelectedItem.Tag, "}") > 0 Then
        sCLSID = Mid$(tvwMain.SelectedItem.Tag, InStr(tvwMain.SelectedItem.Tag, "{"))
        sCLSID = Left$(sCLSID, InStr(sCLSID, "}") + 1)
    End If
    If isCLSID(sCLSID) Then
        'ShellRun "https://www.google.com/search?q=" & sCLSID
        OpenURL "https://www.google.com/search?q=" & sCLSID
    End If
End Sub

Private Sub mnuPopupSaveTree_Click()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "mnuPopupSaveTree_Click - Begin"
    
    Dim sReport$, sFile$, ff%
    pgbStatus.Visible = True
    pgbStatus.Value = 0
    pgbStatus.Max = tvwMain.Nodes.Count
    Form_Resize
    lCountedNodes = 1
    sReport = GetNodeChildren(tvwMain.SelectedItem.Key, 4)
    pgbStatus.Visible = False
    Form_Resize
    'Enter filename to save node tree to..., Text files, All files
    sFile = SaveFileDialog(Translate(957), AppPath(), "node.txt", Translate(901) & " (*.txt)|*.txt|" & Translate(902) & " (*.*)|*.*", Me.hwnd)
    If sFile = vbNullString Then Exit Sub
    
    '" partial report" & vbCrLf & _
              "Root node was '" & tvwMain.SelectedItem.Text & "'" & vbCrLf & _
              "Full path to root node: "
    sReport = Me.Caption & Replace$(Translate(954), "[]", "'" & tvwMain.SelectedItem.Text & "'") & " " & _
        Replace$(tvwMain.SelectedItem.FullPath, tvwMain.Nodes("System").Text, Translate(956)) & vbCrLf & _
        sReport
    
    On Error Resume Next
    ff = FreeFile()
    Open sFile For Output As #ff
        Print #ff, sReport
    Close #ff
    
    If Err.Number = 0 Then
        'Node tree saved to disk as
        Status Translate(958) & " " & sFile
    Else
        'Failed to save tree to disk, error
        Status Translate(959) & ": " & Err.Description & " (" & Translate(960) & " " & Err.Number & ")"
    End If
    
    AppendErrorLogCustom "mnuPopupSaveTree_Click - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "mnuPopupSaveTree_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub mnuPopupShowFile_Click()
    ShowFile tvwMain.SelectedItem.Tag
End Sub

Private Sub mnuPopupShowProp_Click()
    ShowFileProperties tvwMain.SelectedItem.Tag, Me.hwnd
End Sub

Private Sub mnuPopupVerifyFile_Click()
    bSL_Abort = False
    WinTrustVerifyNode tvwMain.SelectedItem.Key
    'Done.
    Status Translate(961)
End Sub

Private Sub mnuViewCollapse_Click()
    Dim Node As Node
    tvwMain.Visible = False
    For Each Node In tvwMain.Nodes
        Node.Expanded = False
    Next Node
    tvwMain.Nodes("System").Expanded = True
    tvwMain.Nodes("System").EnsureVisible
    tvwMain.Visible = True
End Sub

Private Sub mnuOptionsShowCLSID_Click()
    mnuOptionsShowCLSID.Checked = Not mnuOptionsShowCLSID.Checked
    bShowCLSIDs = CBool(mnuOptionsShowCLSID.Checked)
    cmdRefresh.Visible = True
End Sub

Private Sub mnuOptionsShowCmts_Click()
    mnuOptionsShowCmts.Checked = Not mnuOptionsShowCmts.Checked
    bShowCmts = CBool(mnuOptionsShowCmts.Checked)
    cmdRefresh.Visible = True
End Sub

Private Sub mnuOptionsShowEmpty_Click()
    mnuOptionsShowEmpty.Checked = Not mnuOptionsShowEmpty.Checked
    bShowEmpty = CBool(mnuOptionsShowEmpty.Checked)
    cmdRefresh.Visible = True
End Sub

Private Sub mnuViewExpand_Click()
    Dim Node As Node
    tvwMain.Visible = False
    For Each Node In tvwMain.Nodes
        Node.Expanded = True
    Next Node
    tvwMain.Nodes("System").EnsureVisible
    tvwMain.Visible = True
End Sub

Private Sub mnuViewRefresh_Click()
    cmdRefresh.Visible = False
    cmdAbort.Visible = True
    GetAllEnums
    If bSL_Abort Then
        If bSL_Terminate Then
            bSL_Terminate = False
            Unload Me
        Else
            Status Translate(929): bSL_Abort = False
        End If
        Exit Sub
    End If
End Sub

Private Function GetStartupListReport$()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetStartupListReport - Begin"

    Dim sLog$
    'Generating report...
    Status Translate(962)
    'bSL_Abort = False
    cmdAbort.Enabled = True
    cmdAbort.Visible = True
    Form_Resize
    
    sLog = "StartupList report, " & Date & ", " & time & vbCrLf & _
            "StartupList version " & App.Major & "." & Format$(App.Minor, "00") & _
            "." & App.Revision & vbCrLf & _
            "Started from: " & AppPath(True) & vbCrLf & _
            "Detected: " & GetWindowsVersion & vbCrLf
    If bShowPrivacy Then
        sLog = sLog & "Logged on as '" & OSver.UserName & "' to '" & OSver.ComputerName & "'" & vbCrLf
    End If
    
    If Not bShowEmpty And bShowCLSIDs And Not bShowCmts And _
       bShowUsers And bShowHardware And Not bAutoSave Then
       '* Using default options (see end of log for possible options)
        sLog = sLog & Translate(963) & vbCrLf
    End If
    '* Showing empty sections
    If bShowEmpty Then sLog = sLog & Translate(964) & vbCrLf
    '* Hiding CLSIDs
    If Not bShowCLSIDs Then sLog = sLog & Translate(965) & vbCrLf
    '* Showing comments in ini/bat files
    If bShowCmts Then sLog = sLog & Translate(966) & vbCrLf
    '* Hiding entries from other users
    If Not bShowUsers Then sLog = sLog & Translate(967) & vbCrLf
    '* Hiding entries from other hardware configurations
    If Not bShowHardware Then sLog = sLog & Translate(968) & vbCrLf
    '* Automatically saving a report and quitting
    If bAutoSave Then sLog = sLog & Translate(969) & vbCrLf
    
    sLog = sLog & String$(50, "=") & vbCrLf
    
    pgbStatus.Visible = True
    pgbStatus.Value = 0
    pgbStatus.Max = tvwMain.Nodes.Count
    Form_Resize
    lCountedNodes = 1
    sLog = sLog & GetNodeChildren(tvwMain.Nodes("System").Key, 2)
    If bSL_Abort Then Exit Function
    If bShowUsers Then sLog = sLog & GetNodeChildren(tvwMain.Nodes("Users").Key, 1)
    If bSL_Abort Then Exit Function
    If bShowHardware Then sLog = sLog & GetNodeChildren(tvwMain.Nodes("Hardware").Key, 1)
    pgbStatus.Visible = False
    Form_Resize
    If bSL_Abort Then Exit Function
    
    If InStr(sLog, vbCrLf & vbCrLf & vbCrLf) > 0 Then
        sLog = Replace$(sLog, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    End If
    If InStr(sLog, String$(50, "=") & vbCrLf & vbCrLf & String$(20, "-")) Then
        sLog = Replace$(sLog, String$(50, "=") & vbCrLf & vbCrLf & String$(20, "-"), String$(50, "="))
    End If
    
    'Commandline options:" & vbCrLf & _
            "   /showempty      - Show empty sections" & vbCrLf & _
            "   /showcmts       - Show comments in .bat files" & vbCrLf & _
            "   /noshowclsids   - Hide class IDs" & vbCrLf & _
            "   /noshowprivate  - Hide usernames and computer name" & vbCrLf & _
            "   /noshowusers    - Hide entries from other users" & vbCrLf & _
            "   /noshowhardware - Hide entries from other hardware configurations" & vbCrLf & _
            "   /showlargehosts - Show hosts file even when more than 1000 lines are in it" & vbCrLf & _
            "   /showlargezones - Show Zones even when more than 1000 domains are in them" & vbCrLf & _
            "   /autosave       - Run hidden, automatically save a report and quit" & vbCrLf & _
            "   /autosavepath:  - Specify where to save log, when using /autosave." & vbCrLf & _
            "                     Use surrounding quotes for paths with spaces."
    sLog = sLog & String$(50, "-") & vbCrLf & _
            "End of report, xXxXxXx bytes" & vbCrLf & vbCrLf & _
            Translate(970)
            
    sLog = Replace$(sLog, "xXxXxXx", Format$(Len(sLog), "###,###,###"))
    cmdAbort.Visible = False
    Form_Resize
    GetStartupListReport = sLog
    
    AppendErrorLogCustom "GetStartupListReport - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetStartupListReport"
    If inIDE Then Stop: Resume Next
End Function

Private Function GetNodeChildren$(sKey$, iLevel%)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetNodeChildren - Begin"

    Dim s$, t$, u$, nodFirst As Node, nodCurrent As Node
    If bSL_Abort Then Exit Function
    If Trim$(sKey) = vbNullString Then Exit Function
    If tvwMain.Nodes(sKey).Children = 0 Then Exit Function
    If bDebug Then Status sKey
    Set nodFirst = tvwMain.Nodes(sKey).Child
    If Not IsSectionChecked(sKey) Then Exit Function
    Set nodCurrent = nodFirst
    Do
        Select Case iLevel
            Case 1:
                t = String$(50, "=") & vbCrLf & "= "
                u = " =" & vbCrLf & String$(50, "=")
            Case 2:
                t = String$(20, "-") & vbCrLf & vbCrLf
                u = ":" & vbCrLf
            Case 3:
                If nodCurrent.Children > 0 Then
                    t = vbCrLf & "["
                    u = "]"
                Else
                    t = vbNullString
                    u = vbNullString
                End If
            Case 4:
                If nodCurrent.Children > 0 Then
                    t = "* "
                    u = " *"
                Else
                    t = vbNullString
                    u = vbNullString
                End If
            Case 5:
                If nodCurrent.Children > 0 Then
                    t = "- "
                    u = vbNullString
                End If
            Case Else:
                t = vbNullString
                u = vbNullString
        End Select
        If iLevel <> 1 Then
            s = s & vbCrLf & t & nodCurrent.Text & u
        Else
            s = s & vbCrLf & t & nodCurrent.Parent.Text & ": " & nodCurrent.Text & u
        End If
        If nodCurrent.Children > 0 Then s = s & GetNodeChildren(nodCurrent.Key, iLevel + 1)
        If nodCurrent = nodFirst.LastSibling Then Exit Do
        Set nodCurrent = nodCurrent.Next
        lCountedNodes = lCountedNodes + 1
        If lCountedNodes Mod 100 = 0 And lCountedNodes <= pgbStatus.Max Then
            pgbStatus.Value = lCountedNodes
            DoEvents
        End If
        If bSL_Abort Then Exit Function
    Loop
    GetNodeChildren = s & vbCrLf
    AppendErrorLogCustom "GetNodeChildren - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetNodeChildren"
    If inIDE Then Stop: Resume Next
End Function

Private Sub UpdateProgressBar()
    On Error Resume Next
    If bSL_Abort Or bSL_Terminate Then Exit Sub
    If bDebug Then
        If pgbStatus.Value = pgbStatus.Max Then MsgBoxW "UpdateProgressBar: at max!"
    End If
    pgbStatus.Value = pgbStatus.Value + 1
    DoEvents
End Sub

Private Function GetTriageReport() As String
    'not done yet :/
    MsgBoxW Translate(971), vbInformation
End Function

Private Sub EnumProcesses()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumProcesses - Begin"

    Dim sProcessList$(), sModuleList$(), i&, j&, sProc$, lPID&
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "RunningProcesses", SEC_RUNNINGPROCESSES, "memory", "memory"
    sProcessList = Split(GetRunningProcesses, "|")
    For i = 0 To UBound(sProcessList)
        sProc = Mid$(sProcessList(i), InStr(sProcessList(i), "=") + 1)
        lPID = CLng(Left$(sProcessList(i), InStr(sProcessList(i), "=") - 1))
        tvwMain.Nodes.Add "RunningProcesses", tvwChild, "RunningProcesses" & i, GetLongFilename(sProc), "exe", "exe"
        tvwMain.Nodes("RunningProcesses" & i).Tag = GetLongFilename(sProc)
        sModuleList = Split(GetLoadedModules(lPID, sProc), "|")
        For j = 0 To UBound(sModuleList)
            tvwMain.Nodes.Add "RunningProcesses" & i, tvwChild, "RunningProcesses" & i & "." & j, GetLongFilename(sModuleList(j)), "dll"
            tvwMain.Nodes("RunningProcesses" & i & "." & j).Tag = GetLongFilename(sModuleList(j))
        Next j
        If tvwMain.Nodes("RunningProcesses" & i).Children > 0 Then
            tvwMain.Nodes("RunningProcesses" & i).Text = tvwMain.Nodes("RunningProcesses" & i).Text & " (" & tvwMain.Nodes("RunningProcesses" & i).Children & ")"
            tvwMain.Nodes("RunningProcesses" & i).Sorted = True
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("RunningProcesses").Children > 0 Then
        tvwMain.Nodes("RunningProcesses").Text = tvwMain.Nodes("RunningProcesses").Text & " (" & tvwMain.Nodes("RunningProcesses").Children & ")"
        tvwMain.Nodes("RunningProcesses").Sorted = True
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove ("RunningProcesses")
        End If
    End If
    '----------------------------------------------------------------
    'system-wide
    
    AppendErrorLogCustom "EnumProcesses - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, ""
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumAutoStartFolders()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumAutoStartFolders - Begin"
    
    Dim sFolders$(), i&, j&, sFiles$()
    Dim sName$, sDir$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "AutoStartFolders", SEC_AUTOSTARTFOLDERS, "folder", "folder"
    
    ReDim sFolders(12)
    sFolders(0) = "Startup|" & Reg.GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders", "Startup")
    sFolders(1) = "AltStartup|" & Reg.GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders", "AltStartup")
    sFolders(2) = "User Startup|" & Reg.GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\explorer\User Shell Folders", "Startup")
    sFolders(3) = "User AltStartup|" & Reg.GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\explorer\User Shell Folders", "AltStartup")
    sFolders(4) = "Common Startup|" & Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders", "Common Startup")
    sFolders(5) = "Common AltStartup|" & Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders", "Common AltStartup")
    sFolders(6) = "User Common Startup|" & Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\User Shell Folders", "Common Startup")
    sFolders(7) = "User Common AltStartup|" & Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\User Shell Folders", "Common AltStartup")
    sFolders(8) = "IOSUBSYS folder|" & BuildPath(sSysDir, "IOSUBSYS")
    sFolders(9) = "VMM32 folder|" & BuildPath(sSysDir, "vmm32")
    sFolders(10) = "Windows Vista common Startup|%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Startup"
    sFolders(11) = "Windows Vista roaming profile Startup|%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    sFolders(12) = "Windows Vista roaming profile Startup 2|%USERPROFILE%\Application Data\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    
    For i = 0 To UBound(sFolders)
        sName = Left$(sFolders(i), InStr(sFolders(i), "|") - 1)
        sDir = Mid$(sFolders(i), InStr(sFolders(i), "|") + 1)
        sDir = ExpandEnvironmentVars(sDir)
        
        If sDir <> vbNullString Then
            sFiles = Split(EnumFiles(sDir), "|")
            tvwMain.Nodes.Add "AutoStartFolders", tvwChild, "AutoStartFolders" & sName, sName, "folder", "folder"
            tvwMain.Nodes("AutoStartFolders" & sName).Tag = sDir
            For j = 0 To UBound(sFiles)
                If StrComp(sFiles(j), "desktop.ini", 1) <> 0 Then
                    tvwMain.Nodes.Add "AutoStartFolders" & sName, tvwChild, "AutoStartFolders" & sName & j, sFiles(j), "dll", "dll"
                    tvwMain.Nodes("AutoStartFolders" & sName & j).Tag = BuildPath(sDir, sFiles(j))
                End If
            Next j
            If tvwMain.Nodes("AutoStartFolders" & sName).Children > 0 Then
                tvwMain.Nodes("AutoStartFolders" & sName).Text = tvwMain.Nodes("AutoStartFolders" & sName).Text & " (" & tvwMain.Nodes("AutoStartFolders" & sName).Children & ")"
            Else
                If Not bShowEmpty Then
                    tvwMain.Nodes.Remove ("AutoStartFolders" & sName)
                End If
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("AutoStartFolders").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "AutoStartFolders"
    End If
    
    If Not bShowUsers Then Exit Sub
    '--------------------------------------------------------------
    ReDim sFolders(3)
    Dim sUsername$, k&

    For k = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(k))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(k), tvwChild, sUsernames(k) & "AutoStartFolders", SEC_AUTOSTARTFOLDERS, "folder"
            sFolders(0) = "Startup|" & Reg.GetString(HKEY_USERS, sUsernames(k) & "\Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders", "Startup")
            sFolders(1) = "AltStartup|" & Reg.GetString(HKEY_USERS, sUsernames(k) & "\Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders", "AltStartup")
            sFolders(2) = "User Startup|" & Reg.GetString(HKEY_USERS, sUsernames(k) & "\Software\Microsoft\Windows\CurrentVersion\explorer\User Shell Folders", "Startup")
            sFolders(3) = "User AltStartup|" & Reg.GetString(HKEY_USERS, sUsernames(k) & "\Software\Microsoft\Windows\CurrentVersion\explorer\User Shell Folders", "AltStartup")
            
            For i = 0 To UBound(sFolders)
                sName = Left$(sFolders(i), InStr(sFolders(i), "|") - 1)
                sDir = Mid$(sFolders(i), InStr(sFolders(i), "|") + 1)
                sDir = ExpandEnvironmentVars(sDir)
                If sDir <> vbNullString Then
                    sFiles = Split(EnumFiles(sDir), "|")
                    tvwMain.Nodes.Add sUsernames(k) & "AutoStartFolders", tvwChild, sUsernames(k) & "AutoStartFolders" & sName, sName, "folder", "folder"
                    tvwMain.Nodes(sUsernames(k) & "AutoStartFolders" & sName).Tag = sDir
                    For j = 0 To UBound(sFiles)
                        If StrComp(sFiles(j), "desktop.ini", 1) <> 0 Then
                            tvwMain.Nodes.Add sUsernames(k) & "AutoStartFolders" & sName, tvwChild, sUsernames(k) & "AutoStartFolders" & sName & j, sFiles(j), "dll", "dll"
                            tvwMain.Nodes(sUsernames(k) & "AutoStartFolders" & sName & j).Tag = BuildPath(sDir, sFiles(j))
                        End If
                    Next j
                    If tvwMain.Nodes(sUsernames(k) & "AutoStartFolders" & sName).Children = 0 And Not bShowEmpty Then
                        tvwMain.Nodes.Remove (sUsernames(k) & "AutoStartFolders" & sName)
                    End If
                End If
                If bSL_Abort Then Exit Sub
            Next i
            
            If tvwMain.Nodes(sUsernames(k) & "AutoStartFolders").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(k) & "AutoStartFolders"
            End If
        End If
    Next k
    AppendErrorLogCustom "EnumAutoStartFolders - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumAutoStartFolders"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumRunRegKeys()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumRunRegKeys - Begin"

    Dim sKeys$(), sNames$(), i&, j&, sValues$()
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "RunRegkeys", SEC_REGRUNKEYS, "registry", "registry"
    
    ReDim sNames(9)
    sNames(0) = "Run"
    sNames(1) = "RunOnce"
    sNames(2) = "RunServices"
    sNames(3) = "RunServicesOnce"
    sNames(4) = "RunOnceEx"
    sNames(5) = "NT Run"
    sNames(6) = "NT RunOnce"
    sNames(7) = "NT RunServices"
    sNames(8) = "NT RunServicesOnce"
    sNames(9) = "NT RunOnceEx"
    
    ReDim sKeys(9)
    sKeys(0) = "Software\Microsoft\Windows\CurrentVersion\Run"
    sKeys(1) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
    sKeys(2) = "Software\Microsoft\Windows\CurrentVersion\RunServices"
    sKeys(3) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    sKeys(4) = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
    sKeys(5) = "Software\Microsoft\Windows NT\CurrentVersion\Run"
    sKeys(6) = "Software\Microsoft\Windows NT\CurrentVersion\RunOnce"
    sKeys(7) = "Software\Microsoft\Windows NT\CurrentVersion\RunServices"
    sKeys(8) = "Software\Microsoft\Windows NT\CurrentVersion\RunServicesOnce"
    sKeys(9) = "Software\Microsoft\Windows NT\CurrentVersion\RunOnceEx"
    
    For i = 0 To UBound(sKeys)
        sValues = Split(RegEnumValues(HKEY_CURRENT_USER, sKeys(i)), "|")
        tvwMain.Nodes.Add "RunRegkeys", tvwChild, "RunRegkeysUser" & i, "User " & sNames(i), "registry", "registry"
        tvwMain.Nodes("RunRegkeysUser" & i).Tag = "HKEY_CURRENT_USER\" & sKeys(i)
        For j = 0 To UBound(sValues)
            tvwMain.Nodes.Add "RunRegkeysUser" & i, tvwChild, "RunRegkeysUser" & i & "Val" & j, sValues(j), "reg", "reg"
            tvwMain.Nodes("RunRegkeysUser" & i & "Val" & j).Tag = GuessFullpathFromAutorun(Mid$(sValues(j), InStr(sValues(j), " = ") + 3))
        Next j
        tvwMain.Nodes("RunRegkeysUser" & i).Sorted = True
        If tvwMain.Nodes("RunRegkeysUser" & i).Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove ("RunRegkeysUser" & i)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    For i = 0 To UBound(sKeys)
        sValues = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sKeys(i)), "|")
        tvwMain.Nodes.Add "RunRegkeys", tvwChild, "RunRegkeysSystem" & i, "System " & sNames(i), "registry", "registry"
        tvwMain.Nodes("RunRegkeysSystem" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKeys(i)
        For j = 0 To UBound(sValues)
            tvwMain.Nodes.Add "RunRegkeysSystem" & i, tvwChild, "RunRegkeysSystem" & i & "Val" & j, sValues(j), "reg", "reg"
            tvwMain.Nodes("RunRegkeysSystem" & i & "Val" & j).Tag = GuessFullpathFromAutorun(Mid$(sValues(j), InStr(sValues(j), " = ") + 3))
        Next j
        tvwMain.Nodes("RunRegkeysSystem" & i).Sorted = True
        If tvwMain.Nodes("RunRegkeysSystem" & i).Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove ("RunRegkeysSystem" & i)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("RunRegkeys").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "RunRegkeys"
    End If
    
    If Not bShowUsers Then Exit Sub
    '-------------------------------------------------------------------
    Dim sUsername$, k&
    For k = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(k))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(k), tvwChild, sUsernames(k) & "RunRegkeys", SEC_REGRUNKEYS, "registry"
        
            For i = 0 To UBound(sKeys)
                sValues = Split(RegEnumValues(HKEY_USERS, sUsernames(k) & "\" & sKeys(i)), "|")
                tvwMain.Nodes.Add sUsernames(k) & "RunRegkeys", tvwChild, sUsernames(k) & "RunRegkeysUser" & i, "User " & sNames(i), "registry", "registry"
                tvwMain.Nodes(sUsernames(k) & "RunRegkeysUser" & i).Tag = "HKEY_USERS\" & sUsernames(k) & "\" & sKeys(i)
                For j = 0 To UBound(sValues)
                    tvwMain.Nodes.Add sUsernames(k) & "RunRegkeysUser" & i, tvwChild, sUsernames(k) & "RunRegkeysUser" & i & "Val" & j, sValues(j), "reg", "reg"
                    tvwMain.Nodes(sUsernames(k) & "RunRegkeysUser" & i & "Val" & j).Tag = GuessFullpathFromAutorun(Mid$(sValues(j), InStr(sValues(j), " = ") + 3))
                Next j
                tvwMain.Nodes(sUsernames(k) & "RunRegkeysUser" & i).Sorted = True
                If tvwMain.Nodes(sUsernames(k) & "RunRegkeysUser" & i).Children = 0 And Not bShowEmpty Then
                    tvwMain.Nodes.Remove (sUsernames(k) & "RunRegkeysUser" & i)
                End If
            Next i
    
            If tvwMain.Nodes(sUsernames(k) & "RunRegkeys").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(k) & "RunRegkeys"
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next k
    AppendErrorLogCustom "EnumRunRegKeys - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumRunRegKeys"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumRunExRegKeys()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumRunExRegKeys - Begin"
    
    Dim sKeys$(), sNames$(), i&
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "RunExRegkeys", SEC_REGRUNEXKEYS, "registry", "registry"

    ReDim sNames(9)
    sNames(0) = "Run"
    sNames(1) = "RunOnce"
    sNames(2) = "RunOnceEx"
    sNames(3) = "RunServicesOnce"
    sNames(4) = "RunServicesOnceEx"
    sNames(5) = "NT Run"
    sNames(6) = "NT RunOnce"
    sNames(7) = "NT RunOnceEx"
    sNames(8) = "NT RunServicesOnce"
    sNames(9) = "NT RunServicesOnceEx"
    
    ReDim sKeys(9)
    sKeys(0) = "Software\Microsoft\Windows\CurrentVersion\Run"
    sKeys(1) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
    sKeys(2) = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
    sKeys(3) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    sKeys(4) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnceEx"
    sKeys(5) = "Software\Microsoft\Windows NT\CurrentVersion\Run"
    sKeys(6) = "Software\Microsoft\Windows NT\CurrentVersion\RunOnce"
    sKeys(7) = "Software\Microsoft\Windows NT\CurrentVersion\RunOnceEx"
    sKeys(8) = "Software\Microsoft\Windows NT\CurrentVersion\RunServicesOnce"
    sKeys(9) = "Software\Microsoft\Windows NT\CurrentVersion\RunServicesOnceEx"

    Dim sSubkeys$(), sVals$(), j&, k&
    For i = 0 To UBound(sKeys)
        sSubkeys = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sKeys(i)), "|")
        tvwMain.Nodes.Add "RunExRegkeys", tvwChild, "RunEx" & i & "User", "User " & sNames(i), "registry", "registry"
        tvwMain.Nodes("RunEx" & i & "User").Tag = "HKEY_CURRENT_USER\" & sKeys(i)
        For j = 0 To UBound(sSubkeys)
            tvwMain.Nodes.Add "RunEx" & i & "User", tvwChild, "RunEx" & i & "User.sub" & j, sSubkeys(j), "registry", "registry"
            tvwMain.Nodes("RunEx" & i & "User.sub" & j).Tag = "HKEY_CURRENT_USER\" & sKeys(i) & "\" & sSubkeys(j)
            sVals = Split(RegEnumValues(HKEY_CURRENT_USER, sKeys(i) & "\" & sSubkeys(j), True), Chr$(0))
            For k = 0 To UBound(sVals)
                tvwMain.Nodes.Add "RunEx" & i & "User.sub" & j, tvwChild, "RunEx" & i & "User.sub" & j & "val" & k, sVals(k), "reg", "reg"
                tvwMain.Nodes("RunEx" & i & "User.sub" & j & "val" & k).Tag = GuessFullpathFromAutorun(Mid$(sVals(k), InStr(sVals(k), " = ") + 3))
            Next k
            If tvwMain.Nodes("RunEx" & i & "User.sub" & j).Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove "RunEx" & i & "User.sub" & j
            End If
        Next j
        tvwMain.Nodes("RunEx" & i & "User").Sorted = True
        If tvwMain.Nodes("RunEx" & i & "User").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove ("RunEx" & i & "User")
        End If
        If bSL_Abort Then Exit Sub
    Next i
    For i = 0 To UBound(sKeys)
        sSubkeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKeys(i)), "|")
        tvwMain.Nodes.Add "RunExRegkeys", tvwChild, "RunEx" & i & "System", "System " & sNames(i), "registry", "registry"
        tvwMain.Nodes("RunEx" & i & "System").Tag = "HKEY_LOCAL_MACHINE\" & sKeys(i)
        For j = 0 To UBound(sSubkeys)
            tvwMain.Nodes.Add "RunEx" & i & "System", tvwChild, "RunEx" & i & "System.sub" & j, sSubkeys(j), "registry", "registry"
            tvwMain.Nodes("RunEx" & i & "System.sub" & j).Tag = "HKEY_LOCAL_MACHINE\" & sKeys(i) & "\" & sSubkeys(j)
            sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sKeys(i) & "\" & sSubkeys(j), True), Chr$(0))
            For k = 0 To UBound(sVals)
                tvwMain.Nodes.Add "RunEx" & i & "System.sub" & j, tvwChild, "RunEx" & i & "System.sub" & j & "val" & k, sVals(k), "reg", "reg"
                tvwMain.Nodes("RunEx" & i & "System.sub" & j & "val" & k).Tag = GuessFullpathFromAutorun(Mid$(sVals(k), InStr(sVals(k), " = ") + 3))
            Next k
            If tvwMain.Nodes("RunEx" & i & "System.sub" & j).Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove "RunEx" & i & "System.sub" & j
            End If
        Next j
        tvwMain.Nodes("RunEx" & i & "System").Sorted = True
        If tvwMain.Nodes("RunEx" & i & "System").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove ("RunEx" & i & "System")
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("RunExRegkeys").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "RunExRegkeys"
    End If
    
    If Not bShowUsers Then Exit Sub
    '--------------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "RunExRegkeys", SEC_REGRUNEXKEYS, "registry"
            
            For i = 0 To UBound(sKeys)
                sSubkeys = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\" & sKeys(i)), "|")
                tvwMain.Nodes.Add sUsernames(L) & "RunExRegkeys", tvwChild, sUsernames(L) & "RunEx" & i & "User", "User " & sNames(i), "registry", "registry"
                tvwMain.Nodes(sUsernames(L) & "RunEx" & i & "User").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKeys(i)
                For j = 0 To UBound(sSubkeys)
                    tvwMain.Nodes.Add sUsernames(L) & "RunEx" & i & "User", tvwChild, sUsernames(L) & "RunEx" & i & "User.sub" & j, sSubkeys(j), "registry", "registry"
                    tvwMain.Nodes(sUsernames(L) & "RunEx" & i & "User.sub" & j).Tag = "HKEY_CURRENT_USER\" & sUsernames(L) & "\" & sKeys(i) & "\" & sSubkeys(j)
                    sVals = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\" & sKeys(i) & "\" & sSubkeys(j), True), Chr$(0))
                    For k = 0 To UBound(sVals)
                        tvwMain.Nodes.Add sUsernames(L) & "RunEx" & i & "User.sub" & j, tvwChild, sUsernames(L) & "RunEx" & i & "User.sub" & j & "val" & k, sVals(k), "reg", "reg"
                        tvwMain.Nodes(sUsernames(L) & "RunEx" & i & "User.sub" & j & "val" & k).Tag = GuessFullpathFromAutorun(Mid$(sVals(k), InStr(sVals(k), " = ") + 3))
                    Next k
                Next j
                tvwMain.Nodes(sUsernames(L) & "RunEx" & i & "User").Sorted = True
                If tvwMain.Nodes(sUsernames(L) & "RunEx" & i & "User").Children = 0 And Not bShowEmpty Then
                    tvwMain.Nodes.Remove (sUsernames(L) & "RunEx" & i & "User")
                End If
            Next i
            
            If tvwMain.Nodes(sUsernames(L) & "RunExRegkeys").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove (sUsernames(L) & "RunExRegkeys")
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next L
    AppendErrorLogCustom "EnumRunExRegKeys - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumRunExRegKeys"
    If inIDE Then Stop: Resume Next
End Sub

'Private Sub EnumPolicyAutoruns()
'    Dim sPolicies$(), sNames$(), i&
'
'    tvwMain.Nodes.Add "System", tvwChild, "Policy", "Policies autoruns", "registry", "registry"
'
'    ReDim sNames(1)
'    sNames(0) = "Explorer Run"
'    sNames(1) = "Shell"
'    ReDim sPolicies(1)
'    sPolicies(0) = "Software\Microsoft\Windows\CurrentVersion\policies\Explorer\Run"
'    sPolicies(1) = "Software\Microsoft\Windows\CurrentVersion\policies\System"
'
'    Dim sVals$(), j&
'    For i = 0 To UBound(sPolicies)
'        sVals = Split(RegEnumValues(HKEY_CURRENT_USER, sPolicies(i)), "|")
'        tvwMain.Nodes.Add "Policy", tvwChild, "PolicyUser" & i, "User policy " & sNames(i), "registry", "registry"
'        tvwMain.Nodes("PolicyUser" & i).Tag = "HKEY_CURRENT_USER\" & sPolicies(i)
'        For j = 0 To UBound(sVals)
'            If InStr(sVals(j), " = ") <> Len(sVals(j)) - 2 Then
'                tvwMain.Nodes.Add "PolicyUser" & i, tvwChild, "PolicyUser" & i & "sub" & j, sVals(j), "reg", "reg"
'            End If
'        Next j
'        If tvwMain.Nodes("PolicyUser" & i).Children = 0 And Not bShowEmpty Then
'            tvwMain.Nodes.Remove ("PolicyUser" & i)
'        End If
'
'        sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sPolicies(i)), "|")
'        tvwMain.Nodes.Add "Policy", tvwChild, "PolicySystem" & i, "System policy " & sNames(i), "registry", "registry"
'        tvwMain.Nodes("PolicySystem" & i).Tag = "HKEY_LOCAL_MACHINE\" & sPolicies(i)
'        For j = 0 To UBound(sVals)
'            If InStr(sVals(j), " = ") <> Len(sVals(j)) - 2 Then
'                tvwMain.Nodes.Add "PolicySystem" & i, tvwChild, "PolicySystem" & i & "sub" & j, sVals(j), "reg", "reg"
'            End If
'        Next j
'        If tvwMain.Nodes("PolicySystem" & i).Children = 0 And Not bShowEmpty Then
'            tvwMain.Nodes.Remove ("PolicySystem" & i)
'        End If
'    Next i
'
'    If tvwMain.Nodes("Policy").Children = 0 And Not bShowEmpty Then
'        tvwMain.Nodes.Remove "Policy"
'    End If
'
'    If Not bShowUsers Then Exit Sub
'    '--------------------------------------------------------------------
'    Dim sUsername$, l&
'    For l = 0 To UBound(sUsernames)
'        sUsername = MapSIDToUsername(sUsernames(l))
'        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
'            tvwMain.Nodes.Add "Users" & sUsernames(l), tvwChild, sUsernames(l) & "Policy", "Policies autoruns", "registry", "registry"
'
'            For i = 0 To UBound(sPolicies)
'                sVals = Split(RegEnumValues(HKEY_USERS, sUsernames(l) & "\" & sPolicies(i)), "|")
'                tvwMain.Nodes.Add sUsernames(l) & "Policy", tvwChild, sUsernames(l) & "PolicyUser" & i, "User policy " & sNames(i), "registry", "registry"
'                tvwMain.Nodes(sUsernames(l) & "PolicyUser" & i).Tag = "HKEY_USERS\" & sUsernames(l) & "\" & sPolicies(i)
'                For j = 0 To UBound(sVals)
'                    If InStr(sVals(j), " = ") <> Len(sVals(j)) - 2 Then
'                        tvwMain.Nodes.Add sUsernames(l) & "PolicyUser" & i, tvwChild, sUsernames(l) & "Policy" & i & "sub" & j, sVals(j), "reg", "reg"
'                    End If
'                Next j
'                If tvwMain.Nodes(sUsernames(l) & "PolicyUser" & i).Children = 0 And Not bShowEmpty Then
'                    tvwMain.Nodes.Remove (sUsernames(l) & "PolicyUser" & i)
'                End If
'            Next i
'
'            If tvwMain.Nodes(sUsernames(l) & "Policy").Children = 0 And Not bShowEmpty Then
'                tvwMain.Nodes.Remove (sUsernames(l) & "Policy")
'            End If
'        End If
'    Next l
'
'End Sub

Private Sub EnumBatAutostartFiles()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumBatAutostartFiles - Begin"

    Dim sBats$(), i&
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "BatFiles", SEC_BATFILES, "bat", "bat"
    
    ReDim sBats(5)
    sBats(0) = BuildPath(sWinDir, "winstart.bat")
    sBats(1) = BuildPath(sWinDir, "dosstart.bat")
    sBats(2) = Left$(sWinDir, 3) & "autoexec.bat"
    sBats(3) = Left$(sWinDir, 3) & "config.sys"
    sBats(4) = BuildPath(sSysDir, "autoexec.nt")
    sBats(5) = BuildPath(sSysDir, "config.nt")
    
    Dim sFile$, sContent$(), j&
    For i = 0 To UBound(sBats)
        sFile = Mid$(sBats(i), InStrRev(sBats(i), "\") + 1)
        tvwMain.Nodes.Add "BatFiles", tvwChild, "BatFiles" & sFile, sFile, "bat", "bat"
        tvwMain.Nodes("BatFiles" & sFile).Tag = sBats(i)
        sContent = Split(InputFile(sBats(i)), vbCrLf)
        For j = 0 To UBound(sContent)
            If Trim$(sContent(j)) <> vbNullString Then
                If bShowCmts Or Not ( _
                   InStr(1, LTrim$(sContent(j)), "rem", vbTextCompare) > 0 Or _
                   InStr(1, LTrim$(sContent(j)), "::", vbTextCompare) > 0) Then
                    
                    If InStr(sContent(j), vbTab) > 0 Then
                        sContent(j) = Replace$(sContent(j), vbTab, " ")
                    End If
                    tvwMain.Nodes.Add "BatFiles" & sFile, tvwChild, "BatFiles" & sFile & j, sContent(j), "text", "text"
                End If
            End If
        Next j
        If tvwMain.Nodes("BatFiles" & sFile).Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove "BatFiles" & sFile
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("BatFiles").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "BatFiles"
    End If
    
    '--------------------------------------------------------------------
    'nothing for other users - this is system-wide
    AppendErrorLogCustom "EnumBatAutostartFiles - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumBatAutostartFiles"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumAutorunInf()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumAutorunInf - Begin"
    
    Dim sDrives$(), i&, j&, sContent$()
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "AutorunInfs", SEC_AUTORUNINF, "ini"
    sDrives = Split(GetLocalDisks(), " ")
    For i = 0 To UBound(sDrives)
        If FileExists(sDrives(i) & ":\autorun.inf") Then
            tvwMain.Nodes.Add "AutorunInfs", tvwChild, "AutorunInfs" & sDrives(i), sDrives(i) & ":\", "drive"
            tvwMain.Nodes("AutorunInfs" & sDrives(i)).Tag = sDrives(i) & ":\autorun.inf"
            sContent = Split(InputFile(sDrives(i) & ":\autorun.inf"), vbCrLf)
            For j = 0 To UBound(sContent)
                If InStr(1, Trim$(sContent(j)), "open=", vbTextCompare) = 1 Then
                    tvwMain.Nodes.Add "AutorunInfs" & sDrives(i), tvwChild, "AutorunInfs" & sDrives(i) & j, sContent(j), "text"
                End If
                If InStr(1, Trim$(sContent(j)), "shellexecute", vbTextCompare) = 1 Then
                    tvwMain.Nodes.Add "AutorunInfs" & sDrives(i), tvwChild, "AutorunInfs" & sDrives(i) & j, sContent(j), "text"
                End If
            Next j
            If tvwMain.Nodes("AutorunInfs" & sDrives(i)).Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove "AutorunInfs" & sDrives(i)
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("AutorunInfs").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "AutorunInfs"
    End If
    '------------------------------------
    'nothing, this is system-wide
    AppendErrorLogCustom "EnumAutorunInf - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumAutorunInf"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumOnRebootActions()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumOnRebootActions - Begin"
    
    Dim sWininitIni$, sWininitBak$, i&, sContent$()
    Dim sSessionMan$, sBootEx$, sPFRO$
    If bSL_Abort Then Exit Sub
    sSessionMan = "System\CurrentControlSet\Control\Session Manager"
    sWininitIni = sWinDir & "\wininit.ini"
    sWininitBak = sWinDir & "\wininit.bak"
    
    tvwMain.Nodes.Add "System", tvwChild, "OnRebootActions", SEC_ONREBOOT, "onreboot", "onreboot"
    
    sContent = Split(InputFile(sWininitIni), vbCrLf)
    tvwMain.Nodes.Add "OnRebootActions", tvwChild, "OnRebootActionsWininit.ini", "Wininit.ini", "ini", "ini"
    tvwMain.Nodes("OnRebootActionsWininit.ini").Tag = sWininitIni
    For i = 0 To UBound(sContent)
        If Trim$(sContent(i)) <> vbNullString Then
            If InStr(sContent(i), vbTab) > 0 Then
                sContent(i) = Replace$(sContent(i), vbTab, " ")
            End If
            tvwMain.Nodes.Add "OnRebootActionsWininit.ini", tvwChild, "OnRebootActionsWininit.ini" & i, sContent(i), "text", "text"
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("OnRebootActionsWininit.ini").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "OnRebootActionsWininit.ini"
    End If
    
    sContent = Split(InputFile(sWininitBak), vbCrLf)
    tvwMain.Nodes.Add "OnRebootActions", tvwChild, "OnRebootActionsWininit.bak", "Wininit.bak", "ini", "ini"
    tvwMain.Nodes("OnRebootActionsWininit.bak").Tag = sWininitBak
    For i = 0 To UBound(sContent)
        If Trim$(sContent(i)) <> vbNullString Then
            If InStr(sContent(i), vbTab) > 0 Then
                sContent(i) = Replace$(sContent(i), vbTab, " ")
            End If
            tvwMain.Nodes.Add "OnRebootActionsWininit.bak", tvwChild, "OnRebootActionsWininit.bak" & i, sContent(i), "text", "text"
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("OnRebootActionsWininit.bak").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "OnRebootActionsWininit.bak"
    End If
    
    sBootEx = Reg.GetString(HKEY_LOCAL_MACHINE, sSessionMan, "BootExecute")
    If sBootEx <> vbNullString Then
        tvwMain.Nodes.Add "OnRebootActions", tvwChild, "OnRebootActionsBootExecute", "BootExecute = " & sBootEx, "exe", "exe"
        tvwMain.Nodes("OnRebootActionsBootExecute").Tag = "HKEY_LOCAL_MACHINE\" & sSessionMan
    End If
    
    sPFRO = Reg.GetString(HKEY_LOCAL_MACHINE, sSessionMan, "PendingFileRenameOperations", False)
    sContent = Split(sPFRO, Chr$(0))
    If UBound(sContent) > -1 Then
        tvwMain.Nodes.Add "OnRebootActions", tvwChild, "OnRebootActionsPendingFileRenameOperations", "PendingFileRenameOperations", "reg", "reg"
        tvwMain.Nodes("OnRebootActionsPendingFileRenameOperations").Tag = "HKEY_LOCAL_MACHINE\" & sSessionMan
        For i = 0 To UBound(sContent) Step 2
            If i + 1 <= UBound(sContent) Then
                If sContent(i + 1) = vbNullString Then sContent(i + 1) = "NULL"
                If InStr(sContent(i), "!\??\") = 1 Then sContent(i) = Mid$(sContent(i), 6)
                If InStr(sContent(i), "\??\") = 1 Then sContent(i) = Mid$(sContent(i), 5)
                If InStr(sContent(i + 1), "\??\") = 1 Then sContent(i + 1) = Mid$(sContent(i + 1), 5)
                tvwMain.Nodes.Add "OnRebootActionsPendingFileRenameOperations", tvwChild, "OnRebootActionsPendingFileRenameOperations" & i, sContent(i) & " -> " & sContent(i + 1), "reg", "reg"
            End If
            If bSL_Abort Then Exit Sub
        Next i
    End If
    
    If tvwMain.Nodes("OnRebootActions").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "OnRebootActions"
    End If
    
    If Not bShowHardware Then Exit Sub
    '-------------------------------------------------------------------------
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sSessionMan = "System\" & sHardwareCfgs(L) & "\Control\Session Manager"
        
        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "OnRebootActions", SEC_ONREBOOT, "onreboot", "onreboot"
        sBootEx = Reg.GetString(HKEY_LOCAL_MACHINE, sSessionMan, "BootExecute")
        If sBootEx <> vbNullString Then
            tvwMain.Nodes.Add sHardwareCfgs(L) & "OnRebootActions", tvwChild, sHardwareCfgs(L) & "OnRebootActionsBootExecute", "BootExecute = " & sBootEx, "exe", "exe"
            tvwMain.Nodes(sHardwareCfgs(L) & "OnRebootActionsBootExecute").Tag = "HKEY_LOCAL_MACHINE\" & sSessionMan
        End If
        
        sPFRO = Reg.GetString(HKEY_LOCAL_MACHINE, sSessionMan, "PendingFileRenameOperations", False)
        sContent = Split(sPFRO, Chr$(0))
        If UBound(sContent) > -1 Then
            tvwMain.Nodes.Add sHardwareCfgs(L) & "OnRebootActions", tvwChild, sHardwareCfgs(L) & "OnRebootActionsPendingFileRenameOperations", "PendingFileRenameOperations", "reg", "reg"
            tvwMain.Nodes(sHardwareCfgs(L) & "OnRebootActionsPendingFileRenameOperations").Tag = "HKEY_LOCAL_MACHINE\" & sSessionMan
            For i = 0 To UBound(sContent) Step 2
                If i + 1 <= UBound(sContent) Then
                    If sContent(i + 1) = vbNullString Then sContent(i + 1) = "NULL"
                    If InStr(sContent(i), "\??\") = 1 Then sContent(i) = Mid$(sContent(i), 5)
                    If InStr(sContent(i + 1), "\??\") = 1 Then sContent(i + 1) = Mid$(sContent(i + 1), 5)
                    tvwMain.Nodes.Add sHardwareCfgs(L) & "OnRebootActionsPendingFileRenameOperations", tvwChild, sHardwareCfgs(L) & "OnRebootActionsPendingFileRenameOperations" & i, sContent(i) & " -> " & sContent(i + 1), "reg", "reg"
                End If
                If bSL_Abort Then Exit Sub
            Next i
        End If
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "OnRebootActions").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove sHardwareCfgs(L) & "OnRebootActions"
        End If
    Next L
    AppendErrorLogCustom "EnumOnRebootActions - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumOnRebootActions"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumIniFiles()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumIniFiles - Begin"

    Dim sIniStuff$(), i&, j&, sDummy$()
    Dim sFile$, sSection$, sVal$, sData$
    If bSL_Abort Then Exit Sub
    ReDim sIniStuff(9)
    '0/1 at end means can the line occur multiple times
    sIniStuff(0) = "win.ini|windows|load|0"
    sIniStuff(1) = "win.ini|windows|run|0"
    sIniStuff(2) = "system.ini|boot|shell|0"
    sIniStuff(3) = "system.ini|boot|SCRNSAVE.EXE|0"
    sIniStuff(4) = "system.ini|boot|drivers|0"
    sIniStuff(5) = "system.ini|386Enh|device|1"
    sIniStuff(6) = "system.ini|386Enh|mouse|1"
    sIniStuff(7) = "system.ini|386Enh|keyboard|1"
    sIniStuff(8) = "system.ini|386Enh|display|1"
    sIniStuff(9) = "system.ini|386Enh|ebios|1"
    
    tvwMain.Nodes.Add "System", tvwChild, "IniFiles", SEC_INIFILE, "ini", "ini"
    tvwMain.Nodes.Add "IniFiles", tvwChild, "IniFilessystem.ini", "system.ini", "ini", "ini"
    tvwMain.Nodes.Add "IniFiles", tvwChild, "IniFileswin.ini", "win.ini", "ini", "ini"
    tvwMain.Nodes("IniFilessystem.ini").Tag = GuessFullpathFromAutorun("system.ini")
    tvwMain.Nodes("IniFileswin.ini").Tag = GuessFullpathFromAutorun("win.ini")
    
    For i = 0 To UBound(sIniStuff)
        sDummy = Split(sIniStuff(i), "|")
        sFile = sDummy(0)
        sSection = sDummy(1)
        sVal = sDummy(2)
        If sDummy(3) = "0" Then
            sData = IniGetString(sFile, sSection, sVal)
            If sData <> vbNullString Or bShowEmpty Then
                tvwMain.Nodes.Add "IniFiles" & sFile, tvwChild, "IniFiles" & sFile & i, sVal & " = " & sData, "ini", "ini"
            End If
        Else
            sData = IniGetString(sFile, sSection, sVal, , , True)
            sDummy = Split(sData, "|")
            For j = 0 To UBound(sDummy)
                tvwMain.Nodes.Add "IniFiles" & sFile, tvwChild, "IniFiles" & sFile & i & sVal & j, sDummy(j), "ini"
            Next j
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("IniFileswin.ini").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "IniFileswin.ini"
    End If
    If tvwMain.Nodes("IniFilessystem.ini").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "IniFilessystem.ini"
    End If
    If tvwMain.Nodes("IniFiles").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "IniFiles"
    End If
    '----------------------------------------------------------------
    'system-wide
    AppendErrorLogCustom "EnumIniFiles - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumIniFiles"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumIniMappingKeys()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumIniMappingKeys - Begin"
    
    Dim sIniMapping$(), sNames$(), i&
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "IniMapping", SEC_INIMAPPING, "ini", "ini"
    
    ReDim sNames(17)
    sNames(0) = "System NT shell"
    
    sNames(1) = "System NT WinLogon load"
    sNames(2) = "System NT WinLogon run"
    sNames(3) = "User NT WinLogon load"
    sNames(4) = "User NT WinLogon run"
    sNames(5) = "System WinLogon load"
    sNames(6) = "System WinLogon run"
    sNames(7) = "User WinLogon load"
    sNames(8) = "User WinLogon run"
    
    sNames(9) = "System NT Windows load"
    sNames(10) = "System NT Windows run"
    sNames(11) = "User NT Windows load"
    sNames(12) = "User NT Windows run"
    sNames(13) = "System Windows load"
    sNames(14) = "System Windows run"
    sNames(15) = "User Windows load"
    sNames(16) = "User Windows run"
    
    sNames(17) = "User screensaver"
    
    ReDim sIniMapping(17)
    sIniMapping(0) = "HKLM|Software\Microsoft\Windows NT\CurrentVersion\WinLogon|shell"
    
    sIniMapping(1) = "HKLM|Software\Microsoft\Windows NT\CurrentVersion\WinLogon|load"
    sIniMapping(2) = "HKLM|Software\Microsoft\Windows NT\CurrentVersion\WinLogon|run"
    sIniMapping(3) = "HKCU|Software\Microsoft\Windows NT\CurrentVersion\WinLogon|load"
    sIniMapping(4) = "HKCU|Software\Microsoft\Windows NT\CurrentVersion\WinLogon|run"
    sIniMapping(5) = "HKLM|Software\Microsoft\Windows\CurrentVersion\WinLogon|load"
    sIniMapping(6) = "HKLM|Software\Microsoft\Windows\CurrentVersion\WinLogon|run"
    sIniMapping(7) = "HKCU|Software\Microsoft\Windows\CurrentVersion\WinLogon|load"
    sIniMapping(8) = "HKCU|Software\Microsoft\Windows\CurrentVersion\WinLogon|run"
    
    sIniMapping(9) = "HKLM|Software\Microsoft\Windows NT\CurrentVersion\Windows|load"
    sIniMapping(10) = "HKLM|Software\Microsoft\Windows NT\CurrentVersion\Windows|run"
    sIniMapping(11) = "HKCU|Software\Microsoft\Windows NT\CurrentVersion\Windows|load"
    sIniMapping(12) = "HKCU|Software\Microsoft\Windows NT\CurrentVersion\Windows|run"
    sIniMapping(13) = "HKLM|Software\Microsoft\Windows\CurrentVersion\Windows|load"
    sIniMapping(14) = "HKLM|Software\Microsoft\Windows\CurrentVersion\Windows|run"
    sIniMapping(15) = "HKCU|Software\Microsoft\Windows\CurrentVersion\Windows|load"
    sIniMapping(16) = "HKCU|Software\Microsoft\Windows\CurrentVersion\Windows|run"
    
    sIniMapping(17) = "HKCU|Control Panel\Desktop|SCRNSAVE.EXE"
    
    Dim lHive&, sKey$, sVal$, sData$
    For i = 0 To UBound(sIniMapping)
        Select Case Left$(sIniMapping(i), 4)
            Case "HKCU": lHive = HKEY_CURRENT_USER
            Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        End Select
        sVal = Mid$(sIniMapping(i), InStrRev(sIniMapping(i), "|") + 1)
        sKey = Mid$(sIniMapping(i), 6)
        sKey = Mid$(sKey, 1, InStrRev(sKey, "|") - 1)
        sData = ExpandEnvironmentVars(Reg.GetString(lHive, sKey, sVal))
        If sData <> vbNullString Or bShowEmpty Then
            tvwMain.Nodes.Add "IniMapping", tvwChild, "IniMapping" & i, sNames(i) & " = " & sData, "reg", "reg"
            'tvwMain.Nodes("IniMapping" & i).Tag = GuessFullpathFromAutorun(sData)
            Select Case lHive
                Case HKEY_CURRENT_USER:  tvwMain.Nodes("IniMapping" & i).Tag = "HKEY_CURRENT_USER\" & sKey
                Case HKEY_LOCAL_MACHINE: tvwMain.Nodes("IniMapping" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKey
            End Select
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("IniMapping").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "IniMapping"
    End If
    
    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "IniMapping", SEC_INIMAPPING, "ini"

            For i = 0 To UBound(sIniMapping)
                If Left$(sIniMapping(i), 4) = "HKCU" Then
                    sVal = Mid$(sIniMapping(i), InStrRev(sIniMapping(i), "|") + 1)
                    sKey = Mid$(sIniMapping(i), 6)
                    sKey = Mid$(sKey, 1, InStrRev(sKey, "|") - 1)
                    sData = ExpandEnvironmentVars(Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sKey, sVal))
                    If sData <> vbNullString Or bShowEmpty Then
                        tvwMain.Nodes.Add sUsernames(L) & "IniMapping", tvwChild, sUsernames(L) & "IniMapping" & i, sNames(i) & " = " & sData, "reg", "reg"
                        'tvwMain.Nodes(sUsernames(l) & "IniMapping" & i).Tag = GuessFullpathFromAutorun(sData)
                        tvwMain.Nodes(sUsernames(L) & "IniMapping" & i).Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKey
                    End If
                End If
                If bSL_Abort Then Exit Sub
            Next i
            
            If tvwMain.Nodes(sUsernames(L) & "IniMapping").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove (sUsernames(L) & "IniMapping")
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumIniMappingKeys - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumIniMappingKeys"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumShellCommands()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumShellCommands - Begin"

    Dim sTypes$(), i&
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "ShellCommands", SEC_SHELLCOMMANDS, "run", "run"
    tvwMain.Nodes.Add "ShellCommands", tvwChild, "ShellCommandsSystem", "All users", "users"
    tvwMain.Nodes.Add "ShellCommands", tvwChild, "ShellCommandsUser", "This user", "user"
    tvwMain.Nodes("ShellCommandsSystem").Tag = "HKEY_CLASSES_ROOT"
    tvwMain.Nodes("ShellCommandsUser").Tag = "HKEY_CURRENT_USER\Software\Classes"
    
    ReDim sTypes(13)
    sTypes(0) = "exe"
    sTypes(1) = "com"
    sTypes(2) = "bat"
    sTypes(3) = "pif"
    sTypes(4) = "hta"
    sTypes(5) = "vbs"
    sTypes(6) = "vbe"
    sTypes(7) = "js"
    sTypes(8) = "jse"
    sTypes(9) = "wsh"
    sTypes(10) = "wsf"
    sTypes(11) = "scr"
    sTypes(12) = "txt"
    sTypes(13) = "cmd"
    
    Dim sName$, sDesc$, sCmd$
    Dim sVerbs$(), j&
    For i = 0 To UBound(sTypes)
        If Reg.KeyExists(HKEY_CLASSES_ROOT, "." & sTypes(i)) Then
            sName = Reg.GetString(HKEY_CLASSES_ROOT, "." & sTypes(i), vbNullString)
            sDesc = Reg.GetString(HKEY_CLASSES_ROOT, sName, vbNullString)
            
            sVerbs = Split(Reg.EnumSubKeys(HKEY_CLASSES_ROOT, sName & "\shell"), "|")
            For j = 0 To UBound(sVerbs)
                If sDesc = vbNullString Then sDesc = "(no description)"
                'command
                sCmd = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, sName & "\shell\" & sVerbs(j) & "\command", vbNullString))
                sCmd = GetLongFilename(sCmd)
                If Trim$(sCmd) <> vbNullString Or bShowEmpty Then
                    tvwMain.Nodes.Add "ShellCommandsSystem", tvwChild, "ShellCommandsSystem" & sTypes(i) & j, "." & sTypes(i) & " '" & sVerbs(j) & "' - " & sDesc & " - " & sCmd, "exe"
                    tvwMain.Nodes("ShellCommandsSystem" & sTypes(i) & j).Tag = "HKEY_CLASSES_ROOT\" & sName & "\shell\" & sVerbs(j) & "\command"
                End If
                'ddeexec
                sCmd = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, sName & "\shell\" & sVerbs(j) & "\ddeexec", vbNullString))
                sCmd = GetLongFilename(sCmd)
                If Trim$(sCmd) <> vbNullString Or bShowEmpty Then
                    tvwMain.Nodes.Add "ShellCommandsSystem", tvwChild, "ShellCommandsSystem" & sTypes(i) & j & "dde", "." & sTypes(i) & " '" & sVerbs(j) & "' - " & sDesc & " - " & sCmd, "exe"
                    tvwMain.Nodes("ShellCommandsSystem" & sTypes(i) & j & "dde").Tag = "HKEY_CLASSES_ROOT\" & sName & "\shell\" & sVerbs(j) & "\ddeexec"
                End If
            Next j
        Else
            If bShowEmpty Then tvwMain.Nodes.Add "ShellCommandsSystem", tvwChild, "ShellCommandsSystem" & sTypes(i), "." & sTypes(i), "exe"
            If bShowEmpty Then tvwMain.Nodes.Add "ShellCommandsSystem", tvwChild, "ShellCommandsSystem" & sTypes(i) & "dde", "." & sTypes(i), "exe"
        End If
        If bSL_Abort Then Exit Sub
    Next i
    tvwMain.Nodes("ShellCommandsSystem").Sorted = True
    If tvwMain.Nodes("ShellCommandsSystem").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ShellCommandsSystem"
    End If
    
    '2.03 - seems there's something going on for the users as well
    For i = 0 To UBound(sTypes)
        If Reg.KeyExists(HKEY_CURRENT_USER, "Software\Classes\." & sTypes(i)) Then
            sName = Reg.GetString(HKEY_CURRENT_USER, "Software\Classes\." & sTypes(i), vbNullString)
            sDesc = Reg.GetString(HKEY_CURRENT_USER, "Software\Classes\" & sName, vbNullString)
            If sDesc = vbNullString Then sDesc = "(no description)"
            
            sVerbs = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, "Software\Classes\" & sName & "\shell"), "|")
            For j = 0 To UBound(sVerbs)
                'command
                sCmd = ExpandEnvironmentVars(Reg.GetString(HKEY_CURRENT_USER, "Software\Classes\" & sName & "\shell\" & sVerbs(j) & "\command", vbNullString))
                sCmd = GetLongFilename(sCmd)
                If Trim$(sCmd) <> vbNullString Or bShowEmpty Then
                    tvwMain.Nodes.Add "ShellCommandsUser", tvwChild, "ShellCommandsUser" & sTypes(i) & j, "." & sTypes(i) & " '" & sVerbs(j) & "' - " & sDesc & " - " & sCmd, "exe"
                    tvwMain.Nodes("ShellCommandsUser" & sTypes(i) & j).Tag = "HKEY_CURRENT_USER\Software\Classes\" & sName & "\shell\" & sVerbs(j) & "\command"
                End If
                'ddeexec
                sCmd = ExpandEnvironmentVars(Reg.GetString(HKEY_CURRENT_USER, "Software\Classes\" & sName & "\shell\" & sVerbs(j) & "\ddeexec", vbNullString))
                sCmd = GetLongFilename(sCmd)
                If Trim$(sCmd) <> vbNullString Or bShowEmpty Then
                    tvwMain.Nodes.Add "ShellCommandsUser", tvwChild, "ShellCommandsUser" & sTypes(i) & j & "dde", "." & sTypes(i) & " '" & sVerbs(j) & "' - " & sDesc & " - " & sCmd, "exe"
                    tvwMain.Nodes("ShellCommandsUser" & sTypes(i) & j & "dde").Tag = "HKEY_CURRENT_USER\Software\Classes\" & sName & "\shell\" & sVerbs(j) & "\ddeexec"
                End If
            Next j
        Else
            If bShowEmpty Then tvwMain.Nodes.Add "ShellCommandsUser", tvwChild, "ShellCommandsUser" & sTypes(i), "." & sTypes(i), "exe"
            If bShowEmpty Then tvwMain.Nodes.Add "ShellCommandsUser", tvwChild, "ShellCommandsUser" & sTypes(i) & "dde", "." & sTypes(i), "exe"
        End If
        If bSL_Abort Then Exit Sub
    Next i
    tvwMain.Nodes("ShellCommandsUser").Sorted = True
    If tvwMain.Nodes("ShellCommandsUser").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ShellCommandsUser"
    End If
    
    
    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "ShellCommandsUser", SEC_SHELLCOMMANDS, "run"
        
            For i = 0 To UBound(sTypes)
                If Reg.KeyExists(HKEY_USERS, sUsernames(L) & "\Software\Classes\." & sTypes(i)) Then
                    sName = Reg.GetString(HKEY_USERS, sUsernames(L) & "\Software\Classes\." & sTypes(i), vbNullString)
                    sDesc = Reg.GetString(HKEY_USERS, sUsernames(L) & "\Software\Classes\" & sName, vbNullString)
                    If sDesc = vbNullString Then sDesc = "(no description)"
                    
                    sVerbs = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\Software\Classes\" & sName & "\shell"), "|")
                    For j = 0 To UBound(sVerbs)
                        'command
                        sCmd = ExpandEnvironmentVars(Reg.GetString(HKEY_USERS, sUsernames(L) & "\Software\Classes\" & sName & "\shell\" & sVerbs(j) & "\command", vbNullString))
                        sCmd = GetLongFilename(sCmd)
                        If Trim$(sCmd) <> vbNullString Or bShowEmpty Then
                            tvwMain.Nodes.Add sUsernames(L) & "ShellCommandsUser", tvwChild, sUsernames(L) & "ShellCommandsUser" & sTypes(i) & j, "." & sTypes(i) & " '" & sVerbs(j) & "' - " & sDesc & " - " & sCmd, "exe"
                            tvwMain.Nodes(sUsernames(L) & "ShellCommandsUser" & sTypes(i) & j).Tag = "HKEY_USERS\" & sUsernames(L) & "\Software\Classes\" & sName & "\shell\" & sVerbs(j) & "\command"
                        End If
                        'ddeexec
                        sCmd = ExpandEnvironmentVars(Reg.GetString(HKEY_USERS, sUsernames(L) & "\Software\Classes\" & sName & "\shell\" & sVerbs(j) & "\ddeexec", vbNullString))
                        sCmd = GetLongFilename(sCmd)
                        If Trim$(sCmd) <> vbNullString Or bShowEmpty Then
                            tvwMain.Nodes.Add sUsernames(L) & "ShellCommandsUser", tvwChild, sUsernames(L) & "ShellCommandsUser" & sTypes(i) & j & "dde", "." & sTypes(i) & " '" & sVerbs(j) & "' - " & sDesc & " - " & sCmd, "exe"
                            tvwMain.Nodes(sUsernames(L) & "ShellCommandsUser" & sTypes(i) & j & "dde").Tag = "HKEY_USERS\" & sUsernames(L) & "\Software\Classes\" & sName & "\shell\" & sVerbs(j) & "\ddeexec"
                        End If
                    Next j
                Else
                    If bShowEmpty Then tvwMain.Nodes.Add sUsernames(L) & "ShellCommandsUser", tvwChild, sUsernames(L) & "ShellCommandsUser" & sTypes(i), "." & sTypes(i), "exe"
                    If bShowEmpty Then tvwMain.Nodes.Add sUsernames(L) & "ShellCommandsUser", tvwChild, sUsernames(L) & "ShellCommandsUser" & sTypes(i) & "dde", "." & sTypes(i), "exe"
                End If
                If bSL_Abort Then Exit Sub
            Next i
            tvwMain.Nodes(sUsernames(L) & "ShellCommandsUser").Sorted = True
            If tvwMain.Nodes(sUsernames(L) & "ShellCommandsUser").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "ShellCommandsUser"
            End If
                    
        End If
    Next L
    
    AppendErrorLogCustom "EnumShellCommands - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumShellCommands"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub Enum3rdPartyAutostarts()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "Enum3rdPartyAutostarts - Begin"
    
    If bSL_Abort Then Exit Sub
    Dim sUsername$, L&
    
    tvwMain.Nodes.Add "System", tvwChild, "3rdPartyApps", SEC_3RDPARTY, "help", "help"
    If bShowUsers Then
        For L = 0 To UBound(sUsernames)
            sUsername = MapSIDToUsername(sUsernames(L))
            If sUsername <> OSver.UserName And sUsername <> vbNullString Then
                tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "3rdPartyApps", SEC_3RDPARTY, "help"
            End If
        Next L
    End If
    
    
    'ICQ
    EnumICQAgentAutostarts

    'mIRC
    EnumMircAutostarts


    If tvwMain.Nodes("3rdPartyApps").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "3rdPartyApps"
    End If
    If bShowUsers Then
        For L = 0 To UBound(sUsernames)
            sUsername = MapSIDToUsername(sUsernames(L))
            If sUsername <> OSver.UserName And sUsername <> vbNullString Then
                If tvwMain.Nodes(sUsernames(L) & "3rdPartyApps").Children = 0 And Not bShowEmpty Then
                    tvwMain.Nodes.Remove sUsernames(L) & "3rdPartyApps"
                End If
            End If
        Next L
    End If
    
    AppendErrorLogCustom "Enum3rdPartyAutostarts - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Enum3rdPartyAutostarts"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumICQAgentAutostarts()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumICQAgentAutostarts - Begin"

    Dim sICQ$, sKeys$(), i&, sPath$, sFile$
    sICQ = "Software\Mirabilis\ICQ\Agent\Apps"
    sKeys = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sICQ), "|")
    
    tvwMain.Nodes.Add "3rdPartyApps", tvwChild, "ICQ", "ICQ", "icq"
    tvwMain.Nodes("ICQ").Tag = "HKEY_CURRENT_USER\" & sICQ
    For i = 0 To UBound(sKeys)
        tvwMain.Nodes.Add "ICQ", tvwChild, "ICQ" & i, sKeys(i), "reg", "reg"
        tvwMain.Nodes("ICQ" & i).Tag = "HKEY_CURRENT_USER\" & sICQ & "\" & sKeys(i)
        sPath = Reg.GetString(HKEY_CURRENT_USER, sICQ & "\" & sKeys(i), "Path")
        sFile = Reg.GetString(HKEY_CURRENT_USER, sICQ & "\" & sKeys(i), "Startup")
        If sFile <> vbNullString Then
            If sPath <> vbNullString Then sFile = BuildPath(sPath, sFile)
            tvwMain.Nodes.Add "ICQ" & i, tvwChild, "ICQ" & i & "app", sFile, "exe", "exe"
            tvwMain.Nodes("ICQ" & i & "app").Tag = GuessFullpathFromAutorun(sFile)
        Else
            tvwMain.Nodes.Remove "ICQ" & i
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ICQ").Children > 0 Then
        tvwMain.Nodes("ICQ").Text = tvwMain.Nodes("ICQ").Text & " (" & tvwMain.Nodes("ICQ").Children & ")"
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "ICQ"
        End If
    End If
    
    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            sKeys = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\" & sICQ), "|")
                
            tvwMain.Nodes.Add sUsernames(L) & "3rdPartyApps", tvwChild, sUsernames(L) & "ICQ", "ICQ", "exe", "exe"
            For i = 0 To UBound(sKeys)
                tvwMain.Nodes.Add sUsernames(L) & "ICQ", tvwChild, sUsernames(L) & "ICQ" & i, sKeys(i), "reg", "reg"
                sPath = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sICQ & "\" & sKeys(i), "Path")
                sFile = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sICQ & "\" & sKeys(i), "Startup")
                If sFile <> vbNullString Then
                    If sPath <> vbNullString Then sFile = BuildPath(sPath, sFile)
                    tvwMain.Nodes.Add sUsernames(L) & "ICQ" & i, tvwChild, sUsernames(L) & "ICQ" & i & "app", sFile, "exe", "exe"
                Else
                    tvwMain.Nodes.Remove sUsernames(L) & "ICQ" & i
                End If
                If bSL_Abort Then Exit Sub
            Next i
            If tvwMain.Nodes(sUsernames(L) & "ICQ").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "ICQ").Text = tvwMain.Nodes(sUsernames(L) & "ICQ").Text & " (" & tvwMain.Nodes(sUsernames(L) & "ICQ").Children & ")"
            Else
                If Not bShowEmpty Then
                    tvwMain.Nodes.Remove sUsernames(L) & "ICQ"
                End If
            End If
        End If
    Next L
    
    AppendErrorLogCustom "EnumICQAgentAutostarts - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumICQAgentAutostarts"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumMircAutostarts()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumMircAutostarts - Begin"
    
    If bSL_Abort Then Exit Sub
    'mirc autostarts:
    '* mirc.ini [rfiles] remote.ini
    '* mirc.ini [afiles] aliases.ini
    '* perform.ini
    
    Dim sMircPath$
    tvwMain.Nodes.Add "3rdPartyApps", tvwChild, "mIRC", "mIRC", "mirc"
    
    If Not Reg.KeyExists(HKEY_CURRENT_USER, "Software\mIRC") Then
        If Not bShowEmpty Then tvwMain.Nodes.Remove "mIRC"
        Exit Sub
    End If
    'mirc is installed! try to find mIRC path
    
    'from uninstall key
    sMircPath = Reg.GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\mIRC", "UninstallString")
    If sMircPath <> vbNullString Then
        sMircPath = Left$(sMircPath, InStrRev(sMircPath, "\") - 1)
        If Mid$(sMircPath, 1, 1) = """" Then sMircPath = Mid$(sMircPath, 2)
    Else
    
        'from irc protocol key
        sMircPath = Reg.GetString(HKEY_CLASSES_ROOT, "irc\Shell\open\command", vbNullString)
        If sMircPath <> vbNullString Then
            sMircPath = Left$(sMircPath, InStrRev(sMircPath, "\") - 1)
            If Mid$(sMircPath, 1, 1) = """" Then sMircPath = Mid$(sMircPath, 2)
        Else
            
            'from .chat file extension
            sMircPath = Reg.GetString(HKEY_CLASSES_ROOT, "ChatFile\Shell\open\command", vbNullString)
            If sMircPath <> vbNullString Then
                sMircPath = Left$(sMircPath, InStrRev(sMircPath, "\") - 1)
                If Mid$(sMircPath, 1, 1) = """" Then sMircPath = Mid$(sMircPath, 2)
            Else
            
                'guess it!
                If FileExists("C:\mirc") Then sMircPath = "C:\mirc"
                If FileExists("C:\Program Files\mirc") Then sMircPath = "C:\Program Files\mirc"
                If FileExists("D:\mirc") Then sMircPath = "D:\mirc"
                If FileExists("D:\Program Files\mirc") Then sMircPath = "D:\Program Files\mirc"
            End If
        End If
    End If
    If sMircPath = vbNullString Then
        If Not bShowEmpty Then tvwMain.Nodes.Remove "mIRC"
        Exit Sub
    End If
    '===============================
    
    Dim sIni$, i&, j&, sRemote$(), sAliases$()
    ReDim sRemote(0)
    ReDim sAliases(0)
    'get remote/aliases file(s) from mirc.ini
    If FileExists(BuildPath(sMircPath, "mirc.ini")) Then
        tvwMain.Nodes.Add "mIRC", tvwChild, "mIRCmirc.ini", "mirc.ini", "ini"
        tvwMain.Nodes("mIRCmirc.ini").Tag = BuildPath(sMircPath, "mirc.ini")
        
        For i = 0 To 99
            sIni = IniGetString(BuildPath(sMircPath, "mirc.ini"), "rfiles", "n" & i)
            If sIni <> vbNullString Then
                tvwMain.Nodes.Add "mIRCmirc.ini", tvwChild, "mIRCrfiles" & i, "Remote: " & sIni, "text"
                If InStr(sIni, "\") = 0 Then sIni = BuildPath(sMircPath, sIni)
                tvwMain.Nodes("mIRCrfiles" & i).Tag = sIni
                ReDim Preserve sRemote(UBound(sRemote) + 1)
                sRemote(UBound(sRemote)) = sIni
            End If
        Next i
        For i = 0 To 99
            sIni = IniGetString(BuildPath(sMircPath, "mirc.ini"), "afiles", "n" & i)
            If sIni <> vbNullString Then
                tvwMain.Nodes.Add "mIRCmirc.ini", tvwChild, "mIRCafiles" & i, "Aliases: " & sIni, "text"
                If InStr(sIni, "\") = 0 Then sIni = BuildPath(sMircPath, sIni)
                tvwMain.Nodes("mIRCafiles" & i).Tag = sIni
                ReDim Preserve sAliases(UBound(sAliases) + 1)
                sAliases(UBound(sAliases)) = sIni
            End If
        Next i
    End If
    
    'get perform.ini
    If FileExists(BuildPath(sMircPath, "perform.ini")) Then
        tvwMain.Nodes.Add "mIRC", tvwChild, "mIRCperform.ini", "perform.ini", "ini"
        tvwMain.Nodes("mIRCperform.ini").Tag = BuildPath(sMircPath, "perform.ini")
        
        For i = 0 To 99
            sIni = IniGetString(BuildPath(sMircPath, "perform.ini"), "perform", "n" & i)
            If sIni <> vbNullString Then
                tvwMain.Nodes.Add "mIRCperform.ini", tvwChild, "mIRCperform.ini" & i, sIni, "text"
            End If
        Next i
    End If
    
    'get all remotes
    Dim sContent$()
    For i = 1 To UBound(sRemote)
        If FileExists(sRemote(i)) Then
            sContent = Split(InputFile(sRemote(i)), vbCrLf)
            For j = 0 To UBound(sContent)
                If Trim$(sContent(j)) <> vbNullString Then
                    tvwMain.Nodes.Add "mIRCrfiles" & i - 1, tvwChild, "mIRCrfiles" & i - 1 & "sub" & j, sContent(j), "text"
                End If
            Next j
            'tvwMain.Nodes("mIRCrfiles" & i - 1).Sorted = True
            tvwMain.Nodes("mIRCrfiles" & i - 1).Text = tvwMain.Nodes("mIRCrfiles" & i - 1).Text & " (" & tvwMain.Nodes("mIRCrfiles" & i - 1).Children & ")"
        End If
    Next i
    
    'get all aliases
    For i = 1 To UBound(sAliases)
        If FileExists(sAliases(i)) Then
            sContent = Split(InputFile(sAliases(i)), vbCrLf)
            For j = 0 To UBound(sContent)
                If Trim$(sContent(j)) <> vbNullString Then
                    tvwMain.Nodes.Add "mIRCafiles" & i - 1, tvwChild, "mIRCafiles" & i - 1 & "sub" & j, sContent(j), "text"
                End If
            Next j
            'tvwMain.Nodes("mIRCafiles" & i - 1).Sorted = True
            tvwMain.Nodes("mIRCafiles" & i - 1).Text = tvwMain.Nodes("mIRCafiles" & i - 1).Text & " (" & tvwMain.Nodes("mIRCafiles" & i - 1).Children & ")"
        End If
    Next i
    
    If tvwMain.Nodes("mIRC").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "mIRC"
    End If
    
    AppendErrorLogCustom "EnumMircAutostarts - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumMircAutostarts"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumActiveXAutoruns()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumActiveXAutoruns - Begin"
    
    Dim sAXKey$, sKeys$(), i&, sStubPath$, sName$
    If bSL_Abort Then Exit Sub
    sAXKey = "Software\Microsoft\Active Setup\Installed Components"
    tvwMain.Nodes.Add "System", tvwChild, "ActiveX", SEC_ACTIVEX, "msie", "msie"
    tvwMain.Nodes("ActiveX").Tag = "HKEY_LOCAL_MACHINE\" & sAXKey
        
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sAXKey), "|")
    For i = 0 To UBound(sKeys)
        sStubPath = ExpandEnvironmentVars(Reg.GetString(HKEY_LOCAL_MACHINE, sAXKey & "\" & sKeys(i), "StubPath"))
        If sStubPath <> vbNullString Then
            sName = Reg.GetString(HKEY_LOCAL_MACHINE, sAXKey & "\" & sKeys(i), "ComponentID")
            If sName = vbNullString Then sName = "(no name)"
            If InStr(sKeys(i), "{") > 0 Then
                sKeys(i) = Mid$(sKeys(i), InStr(sKeys(i), "{"))
                sKeys(i) = Mid$(sKeys(i), 1, InStr(sKeys(i), "}"))
            End If
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "ActiveX", tvwChild, "ActiveX" & i, sName & " - " & sStubPath, "reg", "reg"
            Else
                tvwMain.Nodes.Add "ActiveX", tvwChild, "ActiveX" & i, sName & " - " & sKeys(i) & " - " & sStubPath, "reg", "reg"
            End If
            tvwMain.Nodes("ActiveX" & i).Tag = "HKEY_LOCAL_MACHINE\" & sAXKey & "\" & sKeys(i)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ActiveX").Children > 0 Then
        tvwMain.Nodes("ActiveX").Text = tvwMain.Nodes("ActiveX").Text & " (" & tvwMain.Nodes("ActiveX").Children & ")"
        tvwMain.Nodes("ActiveX").Sorted = True
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "ActiveX"
        End If
    End If
    
    '----------------------------------------------------------------
    'no per-user stuff, this is system-wide
    AppendErrorLogCustom "EnumActiveXAutoruns - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumActiveXAutoruns"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumDPFs()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumDPFs - Begin"
    Dim sKey$
    If bSL_Abort Then Exit Sub
    sKey = "Software\Microsoft\Code Store Database\Distribution Units"
    Dim sKeys$(), i&, sName$, sFile$, sCodebase$
    tvwMain.Nodes.Add "System", tvwChild, "DPFs", SEC_DPFS, "msie"
    tvwMain.Nodes("DPFs").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey), "|")
    If UBound(sKeys) > -1 Then
        For i = 0 To UBound(sKeys)
            sCodebase = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i) & "\DownloadInformation", "CODEBASE")
            sName = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), vbNullString)
            If sName = vbNullString Then
                sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i), vbNullString)
                If sName = vbNullString Then sName = "(no name)"
            End If
            sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i) & "\InprocServer32", vbNullString))
            If sFile = vbNullString Then sFile = "(no file)"
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "DPFs", tvwChild, "DPFs" & i, sName & " - " & sFile & " - " & sCodebase, "reg"
            Else
                tvwMain.Nodes.Add "DPFs", tvwChild, "DPFs" & i, sName & " - " & sKeys(i) & " - " & sFile & " - " & sCodebase, "reg"
            End If
            tvwMain.Nodes("DPFs" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKey & "\" & sKeys(i)
            If bSL_Abort Then Exit Sub
        Next i
    End If
    If tvwMain.Nodes("DPFs").Children > 0 Then
        tvwMain.Nodes("DPFs").Text = tvwMain.Nodes("DPFs").Text & " (" & tvwMain.Nodes("DPFs").Children & ")"
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "DPFs"
        End If
    End If
    
    '----------------------------------------------------------------
    'no per-user stuff, this is system-wide
    AppendErrorLogCustom "EnumDPFs - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumDPFs"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumProtocols()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumProtocols - Begin"
    
    Dim i&, sKeys$(), sCLSID$, sFile$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "Protocols", SEC_PROTOCOLS, "registry", "registry"
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_CLASSES_ROOT, "Protocols\Filter"), "|")
    If UBound(sKeys) > -1 Then
        tvwMain.Nodes.Add "Protocols", tvwChild, "ProtocolsFilter", "Pluggable MIME filters", "registry", "registry"
        tvwMain.Nodes("ProtocolsFilter").Tag = "HKEY_CLASSES_ROOT\Protocols\Filters"
        For i = 0 To UBound(sKeys)
            sCLSID = Reg.GetString(HKEY_CLASSES_ROOT, "Protocols\Filter\" & sKeys(i), "CLSID")
            sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
            sFile = GetLongFilename(sFile)
            If sFile <> vbNullString Then
                If Not bShowCLSIDs Then
                    tvwMain.Nodes.Add "ProtocolsFilter", tvwChild, "ProtocolsFilter" & i, sKeys(i) & " = " & sFile, "reg", "reg"
                Else
                    tvwMain.Nodes.Add "ProtocolsFilter", tvwChild, "ProtocolsFilter" & i, sKeys(i) & " = " & sCLSID & " = " & sFile, "reg"
                End If
                tvwMain.Nodes("ProtocolsFilter" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
            If bSL_Abort Then Exit Sub
        Next i
        tvwMain.Nodes("ProtocolsFilter").Text = tvwMain.Nodes("ProtocolsFilter").Text & " (" & tvwMain.Nodes("ProtocolsFilter").Children & ")"
    End If
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_CLASSES_ROOT, "Protocols\Handler"), "|")
    If UBound(sKeys) > -1 Then
        tvwMain.Nodes.Add "Protocols", tvwChild, "ProtocolsHandler", "Protocol handlers", "registry", "registry"
        tvwMain.Nodes("ProtocolsHandler").Tag = "HKEY_CLASSES_ROOT\Protocols\Handler"
        For i = 0 To UBound(sKeys)
            sCLSID = Reg.GetString(HKEY_CLASSES_ROOT, "Protocols\Handler\" & sKeys(i), "CLSID")
            sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
            sFile = GetLongFilename(sFile)
            If sFile <> vbNullString Then
                If Not bShowCLSIDs Then
                    tvwMain.Nodes.Add "ProtocolsHandler", tvwChild, "ProtocolsHandler" & i, sKeys(i) & " = " & sFile, "reg", "reg"
                Else
                    tvwMain.Nodes.Add "ProtocolsHandler", tvwChild, "ProtocolsHandler" & i, sKeys(i) & " = " & sCLSID & " = " & sFile, "reg", "reg"
                End If
                tvwMain.Nodes("ProtocolsHandler" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
            If bSL_Abort Then Exit Sub
        Next i
        tvwMain.Nodes("ProtocolsHandler").Text = tvwMain.Nodes("ProtocolsHandler").Text & " (" & tvwMain.Nodes("ProtocolsHandler").Children & ")"
    End If
    
    If tvwMain.Nodes("Protocols").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "Protocols"
    End If
    
    '----------------------------------------------------------------
    'no per-user stuff, this is system-wide
    AppendErrorLogCustom "EnumProtocols - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumProtocols"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumExplorerClones()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumExplorerClones - Begin"
    
    Dim sExplorers$(), i&
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "ExplorerClones", SEC_EXPLORERCLONES, "explorer", "explorer"
        
    ReDim sExplorers(7)
    sExplorers(0) = BuildPath(sWinDir, "explorer.exe")
    sExplorers(1) = BuildPath(Left$(sWinDir, 3), "explorer.exe")
    sExplorers(2) = BuildPath(sWinDir, "system\explorer.exe")
    sExplorers(3) = BuildPath(sWinDir, "system32\explorer.exe")
    sExplorers(3) = BuildPath(sWinDir, "syswow64\explorer.exe")
    sExplorers(4) = BuildPath(sWinDir, "command\explorer.exe")
    sExplorers(5) = BuildPath(sWinDir, "fonts\explorer.exe")
    sExplorers(6) = BuildPath(sWinDir, "explorer\explorer.exe")
    sExplorers(7) = BuildPath(sSysDir, "wbem\explorer.exe")
    
    For i = 0 To UBound(sExplorers)
        If FileExists(sExplorers(i)) Then
            tvwMain.Nodes.Add "ExplorerClones", tvwChild, "ExplorerClones" & i, sExplorers(i), "explorer", "explorer"
            tvwMain.Nodes("ExplorerClones" & i).Tag = sExplorers(i)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("ExplorerClones").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ExplorerClones"
    End If
    
    '----------------------------------------------------------------
    'no per-user stuff, this is system-wide
    AppendErrorLogCustom "EnumExplorerClones - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumExplorerClones"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumServices()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumServices - Begin"
    Dim sKey$, sKeys$(), i&, sDisplayName$, sFile$, lStart&, lType&, sType$, sSafeBoot$
    Dim sBuf$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "Services", SEC_SERVICES, "exe", "exe"
    sKey = "System\CurrentControlSet\Services"
    sSafeBoot = "System\CurrentControlSet\Control\SafeBoot"
    
    'normal Windows NT services
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey), "|")
    tvwMain.Nodes.Add "Services", tvwChild, "NTServices", "NT Services", "exe", "exe"
    tvwMain.Nodes("NTServices").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    For i = 0 To UBound(sKeys)
        sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "DisplayName")
        If sDisplayName = vbNullString Then
            sDisplayName = sKeys(i)
        Else
            If Left$(sDisplayName, 1) = "@" Then                    'extract string resource from file
                sBuf = GetStringFromBinary(, , sDisplayName)
                If 0 <> Len(sBuf) Then sDisplayName = sBuf
            End If
        End If
        
        lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "Start")
        lType = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "Type")
        
        'sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "ImagePath"))
        sFile = GetServiceImagePath(sKeys(i))
        
        If lStart = 2 And sDisplayName <> vbNullString And sFile <> vbNullString And lType >= 16 Then
            tvwMain.Nodes.Add "NTServices", tvwChild, "NTServices" & i, sDisplayName & " = " & sFile, "exe", "exe"
            tvwMain.Nodes("NTServices" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("NTServices").Children > 0 Then
        tvwMain.Nodes("NTServices").Text = tvwMain.Nodes("NTServices").Text & " (" & tvwMain.Nodes("NTServices").Children & ")"
        tvwMain.Nodes("NTServices").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "NTServices"
    End If
    
    'Windows 9x vxd services
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey & "\VxD"), "|")
    tvwMain.Nodes.Add "Services", tvwChild, "VxDServices", "VxD Services", "exe", "exe"
    tvwMain.Nodes("VxDServices").Tag = "HKEY_LOCAL_MACHINE\" & sKey & "\VxD"
    For i = 0 To UBound(sKeys)
        sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\VxD\" & sKeys(i), "DisplayName")
        If sDisplayName = vbNullString Then sDisplayName = sKeys(i)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\VxD\" & sKeys(i), "StaticVxD")
        If sDisplayName <> vbNullString And sFile <> vbNullString Then
            tvwMain.Nodes.Add "VxDServices", tvwChild, "VxDServices" & i, sDisplayName & " = " & sFile, "exe", "exe"
            tvwMain.Nodes("VxDServices" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("VxDServices").Children > 0 Then
        tvwMain.Nodes("VxDServices").Text = tvwMain.Nodes("VxDServices").Text & " (" & tvwMain.Nodes("VxDServices").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "VxDServices"
    End If
    
    'SafeBoot services: Minimal
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sSafeBoot & "\Minimal"), "|")
    tvwMain.Nodes.Add "Services", tvwChild, "SafeBootMinimal", "SafeBoot services (Minimal boot)", "exe"
    tvwMain.Nodes("SafeBootMinimal").Tag = "HKEY_LOCAL_MACHINE\" & sSafeBoot & "\Minimal"
    For i = 0 To UBound(sKeys)
        sType = Reg.GetString(HKEY_LOCAL_MACHINE, sSafeBoot & "\Minimal\" & sKeys(i), vbNullString)
        If Trim$(sType) <> vbNullString Then
            If Not NodeExists("SafeBootMinimal" & Replace$(sType, " ", vbNullString)) Then
                tvwMain.Nodes.Add "SafeBootMinimal", tvwChild, "SafeBootMinimal" & Replace$(sType, " ", vbNullString), sType, "exe"
            End If
            tvwMain.Nodes.Add "SafeBootMinimal" & Replace$(sType, " ", vbNullString), tvwChild, "SafeBootMinimal" & Replace$(sType, " ", vbNullString) & i, sKeys(i), "dll"
            If isCLSID(sKeys(i)) Then
                tvwMain.Nodes("SafeBootMinimal" & Replace$(sType, " ", vbNullString) & i).Tag = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Class\" & sKeys(i)
            Else
                sFile = sKeys(i)
                If InStr(sFile, ".") = Len(sFile) - 3 Then
                    sFile = sSysDir & "\drivers\" & sFile
                    If Not FileExists(sFile) Then
                        sFile = GuessFullpathFromAutorun(sKeys(i))
                    End If
                End If
                tvwMain.Nodes("SafeBootMinimal" & Replace$(sType, " ", vbNullString) & i).Tag = sFile
            End If
        End If
    Next i
    If tvwMain.Nodes("SafeBootMinimal").Children > 0 Then
        'tvwMain.Nodes("SafeBootMinimal").Text = tvwMain.Nodes("SafeBootMinimal").Text & " (" & tvwMain.Nodes("SafeBootMinimal").Children & ")"
        tvwMain.Nodes("SafeBootMinimal").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "SafeBootMinimal"
    End If
    
    'SafeBoot services: Network
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sSafeBoot & "\Network"), "|")
    tvwMain.Nodes.Add "Services", tvwChild, "SafeBootNetwork", "SafeBoot services (Minimal boot + network support)", "exe"
    tvwMain.Nodes("SafeBootNetwork").Tag = "HKEY_LOCAL_MACHINE\" & sSafeBoot & "\Network"
    For i = 0 To UBound(sKeys)
        sType = Reg.GetString(HKEY_LOCAL_MACHINE, sSafeBoot & "\Network\" & sKeys(i), vbNullString)
        If Trim$(sType) <> vbNullString Then
            If Not NodeExists("SafeBootNetwork" & Replace$(sType, " ", vbNullString)) Then
                tvwMain.Nodes.Add "SafeBootNetwork", tvwChild, "SafeBootNetwork" & Replace$(sType, " ", vbNullString), sType, "exe"
            End If
            tvwMain.Nodes.Add "SafeBootNetwork" & Replace$(sType, " ", vbNullString), tvwChild, "SafeBootNetwork" & Replace$(sType, " ", vbNullString) & i, sKeys(i), "dll"
            If isCLSID(sKeys(i)) Then
                tvwMain.Nodes("SafeBootNetwork" & Replace$(sType, " ", vbNullString) & i).Tag = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Class\" & sKeys(i)
            Else
                sFile = sKeys(i)
                If InStr(sFile, ".") = Len(sFile) - 3 Then
                    sFile = sSysDir & "\drivers\" & sFile
                    If Not FileExists(sFile) Then
                        sFile = GuessFullpathFromAutorun(sKeys(i))
                    End If
                End If
                tvwMain.Nodes("SafeBootNetwork" & Replace$(sType, " ", vbNullString) & i).Tag = sFile
            End If
        End If
    Next i
    If tvwMain.Nodes("SafeBootNetwork").Children > 0 Then
        'tvwMain.Nodes("SafeBootNetwork").Text = tvwMain.Nodes("SafeBootNetwork").Text & " (" & tvwMain.Nodes("SafeBootNetwork").Children & ")"
        tvwMain.Nodes("SafeBootNetwork").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "SafeBootNetwork"
    End If
    
    'SafeBoot: AlternateShell
    Dim sAltShell$, lEnable&
    sAltShell = Reg.GetString(HKEY_LOCAL_MACHINE, sSafeBoot, "AlternateShell")
    lEnable = Reg.GetDword(HKEY_LOCAL_MACHINE, sSafeBoot & "\Options", "UseAlternateShell")
    If sAltShell <> vbNullString Or bShowEmpty Then
        tvwMain.Nodes.Add "Services", tvwChild, "SafeBootAltShell", "SafeBoot: Alternate shell", "registry"
        tvwMain.Nodes("SafeBootAltShell").Tag = "HKEY_LOCAL_MACHINE\" & sSafeBoot
        tvwMain.Nodes.Add "SafeBootAltShell", tvwChild, "SafeBootAltShell0", sAltShell & IIf(lEnable = 0, " (not enabled)", " (enabled)"), "explorer"
        tvwMain.Nodes("SafeBootAltShell0").Tag = GuessFullpathFromAutorun(sAltShell)
    End If
    
    If tvwMain.Nodes("Services").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "Services"
    End If
    
    If Not bShowHardware Then Exit Sub
    '----------------------------------------------------------------
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sKey = "System\" & sHardwareCfgs(L) & "\Services"
        sSafeBoot = "System\" & sHardwareCfgs(L) & "\Control\SafeBoot"
    
        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "Services", SEC_SERVICES, "exe", "exe"
        
        sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey), "|")
        tvwMain.Nodes.Add sHardwareCfgs(L) & "Services", tvwChild, sHardwareCfgs(L) & "NTServices", "NT Services", "exe", "exe"
        tvwMain.Nodes(sHardwareCfgs(L) & "NTServices").Tag = "HKEY_LOCAL_MACHINE\" & sKey
        For i = 0 To UBound(sKeys)
            sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "DisplayName")
            If sDisplayName = vbNullString Then
                sDisplayName = sKeys(i)
            Else
                If Left$(sDisplayName, 1) = "@" Then                    'extract string resource from file
                    sBuf = GetStringFromBinary(, , sDisplayName)
                    If 0 <> Len(sBuf) Then sDisplayName = sBuf
                End If
            End If
            lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "Start")
            lType = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "Type")
            sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "ImagePath")
            sFile = CleanServiceFileName(sFile, sKeys(i), sKey)
            
            If lStart = 2 And sDisplayName <> vbNullString And sFile <> vbNullString And lType >= 16 Then
                If InStr(1, sFile, "%systemroot%", vbTextCompare) > 0 Then
                    sFile = Replace$(sFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
                End If
                tvwMain.Nodes.Add sHardwareCfgs(L) & "NTServices", tvwChild, sHardwareCfgs(L) & "NTServices" & i, sDisplayName & " = " & sFile, "exe", "exe"
                tvwMain.Nodes(sHardwareCfgs(L) & "NTServices" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "NTServices").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "NTServices").Text = tvwMain.Nodes(sHardwareCfgs(L) & "NTServices").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "NTServices").Children & ")"
            tvwMain.Nodes(sHardwareCfgs(L) & "NTServices").Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "NTServices"
        End If
    
        'Windows 9x vxd services
        sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey & "\VxD"), "|")
        tvwMain.Nodes.Add sHardwareCfgs(L) & "Services", tvwChild, sHardwareCfgs(L) & "VxDServices", "VxD Services", "exe", "exe"
        tvwMain.Nodes(sHardwareCfgs(L) & "VxDServices").Tag = "HKEY_LOCAL_MACHINE\" & sKey & "\VxD"
        For i = 0 To UBound(sKeys)
            sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\VxD\" & sKeys(i), "DisplayName")
            If sDisplayName = vbNullString Then sDisplayName = sKeys(i)
            sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\VxD\" & sKeys(i), "StaticVxD")
            If sDisplayName <> vbNullString And sFile <> vbNullString Then
                tvwMain.Nodes.Add sHardwareCfgs(L) & "VxDServices", tvwChild, sHardwareCfgs(L) & "VxDServices" & i, sDisplayName & " = " & sFile, "exe", "exe"
                tvwMain.Nodes(sHardwareCfgs(L) & "VxDServices" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "VxDServices").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "VxDServices").Text = tvwMain.Nodes(sHardwareCfgs(L) & "VxDServices").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "VxDServices").Children & ")"
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "VxDServices"
        End If
        
        'SafeBoot services: Minimal
        sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sSafeBoot & "\Minimal"), "|")
        tvwMain.Nodes.Add sHardwareCfgs(L) & "Services", tvwChild, sHardwareCfgs(L) & "SafeBootMinimal", "SafeBoot services (Minimal boot)", "exe"
        tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootMinimal").Tag = "HKEY_LOCAL_MACHINE\" & sSafeBoot & "\Minimal"
        For i = 0 To UBound(sKeys)
            sType = Reg.GetString(HKEY_LOCAL_MACHINE, sSafeBoot & "\Minimal\" & sKeys(i), vbNullString)
            If Trim$(sType) <> vbNullString Then
                If Not NodeExists(sHardwareCfgs(L) & "SafeBootMinimal" & Replace$(sType, " ", vbNullString)) Then
                    tvwMain.Nodes.Add sHardwareCfgs(L) & "SafeBootMinimal", tvwChild, sHardwareCfgs(L) & "SafeBootMinimal" & Replace$(sType, " ", vbNullString), sType, "exe"
                End If
                tvwMain.Nodes.Add sHardwareCfgs(L) & "SafeBootMinimal" & Replace$(sType, " ", vbNullString), tvwChild, sHardwareCfgs(L) & "SafeBootMinimal" & Replace$(sType, " ", vbNullString) & i, sKeys(i), "dll"
                If isCLSID(sKeys(i)) Then
                    tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootMinimal" & Replace$(sType, " ", vbNullString) & i).Tag = "HKEY_LOCAL_MACHINE\System\" & sHardwareCfgs(L) & "\Control\Class\" & sKeys(i)
                Else
                    sFile = sKeys(i)
                    If InStr(sFile, ".") = Len(sFile) - 3 Then
                        sFile = sSysDir & "\drivers\" & sFile
                        If Not FileExists(sFile) Then
                            sFile = GuessFullpathFromAutorun(sKeys(i))
                        End If
                    End If
                    tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootMinimal" & Replace$(sType, " ", vbNullString) & i).Tag = sFile
                End If
            End If
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootMinimal").Children > 0 Then
            'tvwMain.Nodes("SafeBootMinimal").Text = tvwMain.Nodes("SafeBootMinimal").Text & " (" & tvwMain.Nodes("SafeBootMinimal").Children & ")"
            tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootMinimal").Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "SafeBootMinimal"
        End If
        
        'SafeBoot services: Network
        sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sSafeBoot & "\Network"), "|")
        tvwMain.Nodes.Add sHardwareCfgs(L) & "Services", tvwChild, sHardwareCfgs(L) & "SafeBootNetwork", "SafeBoot services (Minimal boot + network support)", "exe"
        tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootNetwork").Tag = "HKEY_LOCAL_MACHINE\" & sSafeBoot & "\Network"
        For i = 0 To UBound(sKeys)
            sType = Reg.GetString(HKEY_LOCAL_MACHINE, sSafeBoot & "\Network\" & sKeys(i), vbNullString)
            If Trim$(sType) <> vbNullString Then
                If Not NodeExists(sHardwareCfgs(L) & "SafeBootNetwork" & Replace$(sType, " ", vbNullString)) Then
                    tvwMain.Nodes.Add sHardwareCfgs(L) & "SafeBootNetwork", tvwChild, sHardwareCfgs(L) & "SafeBootNetwork" & Replace$(sType, " ", vbNullString), sType, "exe"
                End If
                tvwMain.Nodes.Add sHardwareCfgs(L) & "SafeBootNetwork" & Replace$(sType, " ", vbNullString), tvwChild, sHardwareCfgs(L) & "SafeBootNetwork" & Replace$(sType, " ", vbNullString) & i, sKeys(i), "dll"
                If isCLSID(sKeys(i)) Then
                    tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootNetwork" & Replace$(sType, " ", vbNullString) & i).Tag = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Class\" & sKeys(i)
                Else
                    sFile = sKeys(i)
                    If InStr(sFile, ".") = Len(sFile) - 3 Then
                        sFile = sSysDir & "\drivers\" & sFile
                        If Not FileExists(sFile) Then
                            sFile = GuessFullpathFromAutorun(sKeys(i))
                        End If
                    End If
                    tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootNetwork" & Replace$(sType, " ", vbNullString) & i).Tag = sFile
                End If
            End If
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootNetwork").Children > 0 Then
            'tvwMain.Nodes("SafeBootNetwork").Text = tvwMain.Nodes("SafeBootNetwork").Text & " (" & tvwMain.Nodes("SafeBootNetwork").Children & ")"
            tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootNetwork").Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "SafeBootNetwork"
        End If
        
        'SafeBoot: AlternateShell
        sAltShell = Reg.GetString(HKEY_LOCAL_MACHINE, sSafeBoot, "AlternateShell")
        lEnable = Reg.GetDword(HKEY_LOCAL_MACHINE, sSafeBoot & "\Options", "UseAlternateShell")
        If sAltShell <> vbNullString Or bShowEmpty Then
            tvwMain.Nodes.Add sHardwareCfgs(L) & "Services", tvwChild, sHardwareCfgs(L) & "SafeBootAltShell", "SafeBoot: Alternate shell", "registry"
            tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootAltShell").Tag = "HKEY_LOCAL_MACHINE\" & sSafeBoot
            tvwMain.Nodes.Add sHardwareCfgs(L) & "SafeBootAltShell", tvwChild, sHardwareCfgs(L) & "SafeBootAltShell0", sAltShell & IIf(lEnable = 0, " (not enabled)", " (enabled)"), "explorer"
            tvwMain.Nodes(sHardwareCfgs(L) & "SafeBootAltShell0").Tag = GuessFullpathFromAutorun(sAltShell)
        End If
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "Services").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove sHardwareCfgs(L) & "Services"
        End If
    Next L
    AppendErrorLogCustom "EnumServices - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumServices"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumLSP()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumLSP - Begin"
    
    'Winsock LSP entries
    Dim sLSPKey$
    If bSL_Abort Then Exit Sub
    sLSPKey = "System\CurrentControlSet\Services\Winsock2"
    tvwMain.Nodes.Add "System", tvwChild, "WinsockLSP", SEC_WINSOCKLSP, "lsp"
    If Not Reg.KeyExists(HKEY_LOCAL_MACHINE, sLSPKey) Then
        'winsock2 not installed (win95 /wo winsock2 update)
        If bShowEmpty Then
            'Winsock 2 not installed
            tvwMain.Nodes.Add "WinsockLSP", tvwChild, "WinsockLSPWin95", Translate(972), "internet"
        End If
        Exit Sub
    End If
    
    Dim sWinsock$(), i&, sFile$
    sWinsock = Split(EnumWinsockProtocol, "|")
    tvwMain.Nodes.Add "WinsockLSP", tvwChild, "WinsockLSPProtocols", "Protocols", "internet"
    tvwMain.Nodes("WinsockLSPProtocols").Tag = "HKEY_LOCAL_MACHINE\" & sLSPKey & "\Parameters\Protocol_Catalog9\Catalog_Entries"
    For i = 0 To UBound(sWinsock)
        sWinsock(i) = ExpandEnvironmentVars(sWinsock(i))
        tvwMain.Nodes.Add "WinsockLSPProtocols", tvwChild, "WinsockLSPProtocols" & i, sWinsock(i), "internet"
        sFile = Mid$(sWinsock(i), InStrRev(sWinsock(i), " - ") + 3)
        tvwMain.Nodes("WinsockLSPProtocols" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("WinsockLSPProtocols").Children > 0 Then
        tvwMain.Nodes("WinsockLSPProtocols").Text = tvwMain.Nodes("WinsockLSPProtocols").Text & " (" & tvwMain.Nodes("WinsockLSPProtocols").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "WinsockLSPProtocols"
    End If
    sWinsock = Split(EnumWinsockNameSpace, "|")
    tvwMain.Nodes.Add "WinsockLSP", tvwChild, "WinsockLSPNamespaces", "Namespace Providers", "internet"
    tvwMain.Nodes("WinsockLSPNamespaces").Tag = "HKEY_LOCAL_MACHINE\" & sLSPKey & "\Parameters\NameSpace_Catalog5\Catalog_Entries"
    For i = 0 To UBound(sWinsock)
        sWinsock(i) = ExpandEnvironmentVars(sWinsock(i))
        tvwMain.Nodes.Add "WinsockLSPNamespaces", tvwChild, "WinsockLSPNamespaces" & i, sWinsock(i), "internet"
        sFile = Mid$(sWinsock(i), InStrRev(sWinsock(i), " - ") + 3)
        tvwMain.Nodes("WinsockLSPNamespaces" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("WinsockLSPNamespaces").Children > 0 Then
        tvwMain.Nodes("WinsockLSPNamespaces").Text = tvwMain.Nodes("WinsockLSPNamespaces").Text & " (" & tvwMain.Nodes("WinsockLSPNamespaces").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "WinsockLSPNamespaces"
    End If
    If tvwMain.Nodes("WinsockLSP").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "WinsockLSP"
    End If
    
    '----------------------------------------------------------------
    'other controlsets would be nice, but the APIs can only read the
    'active one :/
    AppendErrorLogCustom "EnumLSP - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumLSP"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumWinLogonAutoruns()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumWinLogonAutoruns - Begin"
    
    'Winlogon\Notify,GinaDLL,GPExtensions,UserInit,
    'AppInit_DLLs,System,VmApplet,TaskMan,and a shitload more
    
    Dim sKeys$(), i&, sWinLogon$, sWindows$
    Dim Stady As Long
    If bSL_Abort Then Exit Sub
    sWinLogon = "Software\Microsoft\Windows NT\CurrentVersion\WinLogon"
    sWindows = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
    
    Stady = 1
    
    Dim sFile$, sName$
    tvwMain.Nodes.Add "System", tvwChild, "WinLogonAutoruns", SEC_WINLOGON, "winlogon", "winlogon"
    
    Stady = 2
    
    Dim sValsL$(), sValsW$()
    ReDim sValsL(4) 'sWinLogon
    sValsL(0) = "UserInit"
    sValsL(1) = "System"
    sValsL(2) = "VmApplet"
    sValsL(3) = "TaskMan"
    sValsL(4) = "UIHost"
    ReDim sValsW(0) 'sWindows
    sValsW(0) = "AppInit_DLLs"
    
    For i = 0 To UBound(sValsL)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sWinLogon, sValsL(i))
        If sFile <> vbNullString Or bShowEmpty Then
            tvwMain.Nodes.Add "WinLogonAutoruns", tvwChild, "WinLogonL" & i, sValsL(i) & " = " & sFile, "reg", "reg"
            tvwMain.Nodes("WinLogonL" & i).Tag = "HKEY_LOCAL_MACHINE\" & sWinLogon
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    Stady = 3
    
    For i = 0 To UBound(sValsW)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sWindows, sValsW(i))
        If sFile <> vbNullString Or bShowEmpty Then
            tvwMain.Nodes.Add "WinLogonAutoruns", tvwChild, "WinLogonW" & i, sValsW(i) & " = " & sFile, "reg", "reg"
            tvwMain.Nodes("WinLogonW" & i).Tag = "HKEY_LOCAL_MACHINE\" & sWindows
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    Stady = 4
    
    'Winlogon\Notify
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sWinLogon & "\Notify"), "|")
    tvwMain.Nodes.Add "WinLogonAutoruns", tvwChild, "WinLogonNotify", "Notify", "registry", "registry"
    tvwMain.Nodes("WinLogonNotify").Tag = "HKEY_LOCAL_MACHINE\" & sWinLogon & "\Notify"
    
    Stady = 5
    
    For i = 0 To UBound(sKeys)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sWinLogon & "\Notify\" & sKeys(i), "DllName")
        If sFile <> vbNullString Then
            tvwMain.Nodes.Add "WinLogonNotify", tvwChild, "WinLogonNotify" & i, sKeys(i) & " = " & sFile, "dll", "dll"
            tvwMain.Nodes("WinLogonNotify" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    Stady = 6
    
    If tvwMain.Nodes("WinLogonNotify").Children > 0 Then
        tvwMain.Nodes("WinLogonNotify").Text = tvwMain.Nodes("WinLogonNotify").Text & " (" & tvwMain.Nodes("WinLogonNotify").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "WinLogonNotify"
    End If

    Stady = 7

    'GinaDLL
    sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sWinLogon, "GinaDLL")
    If sFile <> vbNullString Or bShowEmpty Then
        tvwMain.Nodes.Add "WinLogonAutoruns", tvwChild, "WinLogonGinaDLL", "GinaDLL = " & sFile, "dll", "dll"
        tvwMain.Nodes("WinLogonGinaDLL").Tag = "HKEY_LOCAL_MACHINE\" & sWinLogon
    End If
    
    Stady = 8
    
    'GPExtensions
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sWinLogon & "\GPExtensions"), "|")
    tvwMain.Nodes.Add "WinLogonAutoruns", tvwChild, "WinLogonGPExtensions", "Group policy extensions", "registry", "registry"
    tvwMain.Nodes("WinLogonGPExtensions").Tag = "HKEY_LOCAL_MACHINE\" & sWinLogon & "\GPExtensions"
    
    Stady = 9
    
    For i = 0 To UBound(sKeys)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sWinLogon & "\GPExtensions\" & sKeys(i), "DllName")
        sName = Reg.GetString(HKEY_LOCAL_MACHINE, sWinLogon & "\GPExtensions\" & sKeys(i), vbNullString)
        If sFile <> vbNullString Then
            tvwMain.Nodes.Add "WinLogonGPExtensions", tvwChild, "WinLogonGPExtensions" & i, sName & " = " & sFile, "dll", "dll"
            tvwMain.Nodes("WinLogonGPExtensions" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    Stady = 10
    
    If tvwMain.Nodes("WinLogonGPExtensions").Children > 0 Then
        tvwMain.Nodes("WinLogonGPExtensions").Text = tvwMain.Nodes("WinLogonGPExtensions").Text & " (" & tvwMain.Nodes("WinLogonGPExtensions").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "WinLogonGPExtensions"
    End If
    
    Stady = 11
    
    If tvwMain.Nodes("WinLogonAutoruns").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "WinLogonAutoruns"
    End If

    '----------------------------------------------------------------
    'system-wide
    AppendErrorLogCustom "EnumWinLogonAutoruns - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumWinLogonAutoruns", "Stady: " & Stady
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumBHOs()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumBHOs - Begin"
    
    Dim sBHO$, sName$, sFile$, sKeys$(), i&
    If bSL_Abort Then Exit Sub
    sBHO = "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects"
    tvwMain.Nodes.Add "System", tvwChild, "BHOs", SEC_BHOS, "msie", "msie"
    tvwMain.Nodes("BHOs").Tag = "HKEY_LOCAL_MACHINE\" & sBHO
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sBHO), "|")
    For i = 0 To UBound(sKeys)
        sName = Reg.GetString(HKEY_LOCAL_MACHINE, sBHO & "\" & sKeys(i), vbNullString)
        If sName = vbNullString Then
            sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i), vbNullString)
        End If
        If sName = vbNullString Then sName = "(no name)"
        sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i) & "\InprocServer32", vbNullString)
        sFile = GetLongFilename(sFile)
        If Not bShowCLSIDs Then
            tvwMain.Nodes.Add "BHOs", tvwChild, "BHO" & i, sName & " = " & sFile, "dll", "dll"
        Else
            tvwMain.Nodes.Add "BHOs", tvwChild, "BHO" & i, sName & " = " & sKeys(i) & " = " & sFile, "dll", "dll"
        End If
        tvwMain.Nodes("BHO" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("BHOs").Children > 0 Then
        tvwMain.Nodes("BHOs").Text = tvwMain.Nodes("BHOs").Text & " (" & tvwMain.Nodes("BHOs").Children & ")"
        tvwMain.Nodes("BHOs").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "BHOs"
    End If
    
    '----------------------------------------------------------------
    'system-wide
    AppendErrorLogCustom "EnumBHOs - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumBHOs"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumSSODL()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumSSODL - Begin"
    
    Dim sSSODL$
    If bSL_Abort Then Exit Sub
    sSSODL = "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
    tvwMain.Nodes.Add "System", tvwChild, "ShellServiceObjectDelayLoad", SEC_SSODL, "registry", "registry"
    
    tvwMain.Nodes.Add "ShellServiceObjectDelayLoad", tvwChild, "ShellServiceObjectDelayLoadSystem", "All users", "users"
    tvwMain.Nodes("ShellServiceObjectDelayLoadSystem").Tag = "HKEY_LOCAL_MACHINE\" & sSSODL
    
    Dim sVals$(), i&, sCLSID$, sFile$
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sSSODL), "|")
    For i = 0 To UBound(sVals)
        sCLSID = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
        sVals(i) = Mid$(sVals(i), 1, InStr(sVals(i), " = ") - 1)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
        If sFile <> vbNullString Then
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "ShellServiceObjectDelayLoadSystem", tvwChild, "SSODLSystem" & i, sVals(i) & " = " & sFile, "dll", "dll"
            Else
                tvwMain.Nodes.Add "ShellServiceObjectDelayLoadSystem", tvwChild, "SSODLSystem" & i, sVals(i) & " = " & sCLSID & " = " & sFile, "dll", "dll"
            End If
            tvwMain.Nodes("SSODLSystem" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i

    If tvwMain.Nodes("ShellServiceObjectDelayLoadSystem").Children > 0 Then
        tvwMain.Nodes("ShellServiceObjectDelayLoadSystem").Text = tvwMain.Nodes("ShellServiceObjectDelayLoadSystem").Text & " (" & tvwMain.Nodes("ShellServiceObjectDelayLoadSystem").Children & ")"
        tvwMain.Nodes("ShellServiceObjectDelayLoadSystem").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "ShellServiceObjectDelayLoadSystem"
    End If
    
    tvwMain.Nodes.Add "ShellServiceObjectDelayLoad", tvwChild, "ShellServiceObjectDelayLoadUser", "This user", "user"
    tvwMain.Nodes("ShellServiceObjectDelayLoadUser").Tag = "HKEY_CURRENT_USER\" & sSSODL
    
    sVals = Split(RegEnumValues(HKEY_CURRENT_USER, sSSODL), "|")
    For i = 0 To UBound(sVals)
        sCLSID = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
        sVals(i) = Mid$(sVals(i), 1, InStr(sVals(i), " = ") - 1)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
        If sFile <> vbNullString Then
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "ShellServiceObjectDelayLoadUser", tvwChild, "SSODLUser" & i, sVals(i) & " = " & sFile, "dll", "dll"
            Else
                tvwMain.Nodes.Add "ShellServiceObjectDelayLoadUser", tvwChild, "SSODLUser" & i, sVals(i) & " = " & sCLSID & " = " & sFile, "dll", "dll"
            End If
            tvwMain.Nodes("SSODLUser" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ShellServiceObjectDelayLoadUser").Children > 0 Then
        tvwMain.Nodes("ShellServiceObjectDelayLoadUser").Text = tvwMain.Nodes("ShellServiceObjectDelayLoadUser").Text & " (" & tvwMain.Nodes("ShellServiceObjectDelayLoadUser").Children & ")"
        tvwMain.Nodes("ShellServiceObjectDelayLoadUser").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "ShellServiceObjectDelayLoadUser"
    End If
    
    If tvwMain.Nodes("ShellServiceObjectDelayLoad").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ShellServiceObjectDelayLoad"
    End If

    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "ShellServiceObjectDelayLoadUser", SEC_SSODL, "registry"
            tvwMain.Nodes(sUsernames(L) & "ShellServiceObjectDelayLoadUser").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sSSODL
    
            sVals = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\" & sSSODL), "|")
            For i = 0 To UBound(sVals)
                sCLSID = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
                sVals(i) = Mid$(sVals(i), 1, InStr(sVals(i), " = ") - 1)
                sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
                If sFile <> vbNullString Then
                    If Not bShowCLSIDs Then
                        tvwMain.Nodes.Add sUsernames(L) & "ShellServiceObjectDelayLoadUser", tvwChild, sUsernames(L) & "SSODL" & i, sVals(i) & " = " & sFile, "dll", "dll"
                    Else
                        tvwMain.Nodes.Add sUsernames(L) & "ShellServiceObjectDelayLoadUser", tvwChild, sUsernames(L) & "SSODL" & i, sVals(i) & " = " & sCLSID & " = " & sFile, "dll", "dll"
                    End If
                    tvwMain.Nodes(sUsernames(L) & "SSODL" & i).Tag = GuessFullpathFromAutorun(sFile)
                End If
                If bSL_Abort Then Exit Sub
            Next i
            If tvwMain.Nodes(sUsernames(L) & "ShellServiceObjectDelayLoadUser").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "ShellServiceObjectDelayLoadUser").Text = tvwMain.Nodes(sUsernames(L) & "ShellServiceObjectDelayLoadUser").Text & " (" & tvwMain.Nodes(sUsernames(L) & "ShellServiceObjectDelayLoadUser").Children & ")"
                tvwMain.Nodes(sUsernames(L) & "ShellServiceObjectDelayLoadUser").Sorted = True
            Else
                If Not bShowEmpty Then tvwMain.Nodes.Remove sUsernames(L) & "ShellServiceObjectDelayLoadUser"
            End If
        End If
    Next L
    
    AppendErrorLogCustom "EnumSSODL - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumSSODL"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumSharedTaskScheduler()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumSharedTaskScheduler - Begin"
    
    Dim sSTS$
    If bSL_Abort Then Exit Sub
    sSTS = "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler"
    Dim sVals$(), i&, sName$, sFile$
    tvwMain.Nodes.Add "System", tvwChild, "SharedTaskScheduler", SEC_SHAREDTASKSCHEDULER, "registry", "registry"
    tvwMain.Nodes("SharedTaskScheduler").Tag = "HKEY_LOCAL_MACHINE\" & sSTS
    
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sSTS), "|")
    For i = 0 To UBound(sVals)
        sName = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
        sVals(i) = Mid$(sVals(i), 1, InStr(sVals(i), " = ") - 1)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sVals(i) & "\InprocServer32", vbNullString))
        If sFile <> vbNullString Then
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "SharedTaskScheduler", tvwChild, "SharedTaskScheduler" & i, sName & " = " & sFile, "dll", "dll"
            Else
                tvwMain.Nodes.Add "SharedTaskScheduler", tvwChild, "SharedTaskScheduler" & i, sName & " = " & sVals(i) & " = " & sFile, "dll", "dll"
            End If
            tvwMain.Nodes("SharedTaskScheduler" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("SharedTaskScheduler").Children > 0 Then
        tvwMain.Nodes("SharedTaskScheduler").Text = tvwMain.Nodes("SharedTaskScheduler").Text & " (" & tvwMain.Nodes("SharedTaskScheduler").Children & ")"
        tvwMain.Nodes("SharedTaskScheduler").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "SharedTaskScheduler"
    End If
    
    '----------------------------------------------------------------
    'system-wide
    
    AppendErrorLogCustom "EnumSharedTaskScheduler - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumSharedTaskScheduler"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumMPRServices()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumMPRServices - Begin"
    
    Dim sMPR$
    If bSL_Abort Then Exit Sub
    sMPR = "System\CurrentControlSet\Control\MPRServices"
    Dim sKeys$(), i&, sFile$ ', sEntryPoint$
    tvwMain.Nodes.Add "System", tvwChild, "MPRServices", SEC_MPRSERVICES, "registry", "registry"
    tvwMain.Nodes("MPRServices").Tag = "HKEY_LOCAL_MACHINE\" & sMPR
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sMPR), "|")
    For i = 0 To UBound(sKeys)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sMPR & "\" & sKeys(i), "DllName")
        'sEntryPoint = Reg.GetString(HKEY_LOCAL_MACHINE, sMPR & "\" & sKeys(i), "EntryPoint")
        If sFile <> vbNullString Then
            tvwMain.Nodes.Add "MPRServices", tvwChild, "MPRServices" & i, sKeys(i) & " = " & sFile, "dll", "dll"
            tvwMain.Nodes("MPRServices" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("MPRServices").Children > 0 Then
        tvwMain.Nodes("MPRServices").Text = tvwMain.Nodes("MPRServices").Text & " (" & tvwMain.Nodes("MPRServices").Children & ")"
        tvwMain.Nodes("MPRServices").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "MPRServices"
    End If
    
    If Not bShowHardware Then Exit Sub
    '----------------------------------------------------------------
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sMPR = "System\" & sHardwareCfgs(L) & "\Control\MPRServices"
        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "MPRServices", SEC_MPRSERVICES, "registry", "registry"
        
        sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sMPR), "|")
        For i = 0 To UBound(sKeys)
            sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sMPR & "\" & sKeys(i), "DllName")
            'sEntryPoint = Reg.GetString(HKEY_LOCAL_MACHINE, sMPR & "\" & sKeys(i), "EntryPoint")
            If sFile <> vbNullString Then
                tvwMain.Nodes.Add sHardwareCfgs(L) & "MPRServices", tvwChild, sHardwareCfgs(L) & "MPRServices" & i, sKeys(i) & " = " & sFile, "dll", "dll"
                tvwMain.Nodes(sHardwareCfgs(L) & "MPRServices" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "MPRServices").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "MPRServices").Text = tvwMain.Nodes(sHardwareCfgs(L) & "MPRServices").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "MPRServices").Children & ")"
            tvwMain.Nodes(sHardwareCfgs(L) & "MPRServices").Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "MPRServices"
        End If
    Next L
    AppendErrorLogCustom "EnumMPRServices - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumMPRServices"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumWOW()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumWOW - Begin"
    
    Dim sCmd$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "WOW", SEC_WOW, "registry", "registry"
    
    Dim sVals$(), i&, sWOW$, sSessionMan$, sKnownDlls$()
    sWOW = "System\CurrentControlSet\Control\WOW"
    sSessionMan = "System\CurrentControlSet\Control\Session Manager"
    tvwMain.Nodes("WOW").Tag = "HKEY_LOCAL_MACHINE\" & sWOW
    ReDim sVals(1)
    sVals(0) = "cmdline"
    sVals(1) = "wowcmdline"
    
    For i = 0 To UBound(sVals)
        sCmd = ExpandEnvironmentVars(Reg.GetString(HKEY_LOCAL_MACHINE, sWOW, sVals(i)))
        If sCmd <> vbNullString Or bShowEmpty Then
            tvwMain.Nodes.Add "WOW", tvwChild, "WOW" & i, sVals(i) & " = " & sCmd, "exe", "exe"
            tvwMain.Nodes("WOW" & i).Tag = GuessFullpathFromAutorun(sCmd)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    sKnownDlls = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sWOW, "KnownDlls"), " ")
    tvwMain.Nodes.Add "WOW", tvwChild, "WOWKnownDlls", "KnownDlls (16-bit)", "reg"
    tvwMain.Nodes("WOWKnownDlls").Tag = "HKEY_LOCAL_MACHINE\" & sWOW
    For i = 0 To UBound(sKnownDlls)
        tvwMain.Nodes.Add "WOWKnownDlls", tvwChild, "WOWKnownDlls" & i, sKnownDlls(i), "dll"
        tvwMain.Nodes("WOWKnownDlls" & i).Tag = GuessFullpathFromAutorun(sKnownDlls(i))
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("WOWKnownDlls").Children > 0 Then
        tvwMain.Nodes("WOWKnownDlls").Text = tvwMain.Nodes("WOWKnownDlls").Text & " (" & tvwMain.Nodes("WOWKnownDlls").Children & ")"
        tvwMain.Nodes("WOWKnownDlls").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "WOWKnownDlls"
    End If
    
    sKnownDlls = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sSessionMan & "\KnownDlls"), "|")
    tvwMain.Nodes.Add "WOW", tvwChild, "WOWKnownDlls32b", "KnownDlls (32-bit)", "reg"
    tvwMain.Nodes("WOWKnownDlls32b").Tag = "HKEY_LOCAL_MACHINE\" & sSessionMan & "\KnownDlls"
    For i = 0 To UBound(sKnownDlls)
        sCmd = Mid$(sKnownDlls(i), InStr(sKnownDlls(i), " = ") + 3)
        tvwMain.Nodes.Add "WOWKnownDlls32b", tvwChild, "WOWKnownDlls32b" & i, sCmd, "dll"
        tvwMain.Nodes("WOWKnownDlls32b" & i).Tag = GuessFullpathFromAutorun(sCmd)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("WOWKnownDlls32b").Children > 0 Then
        tvwMain.Nodes("WOWKnownDlls32b").Text = tvwMain.Nodes("WOWKnownDlls32b").Text & " (" & tvwMain.Nodes("WOWKnownDlls32b").Children & ")"
        tvwMain.Nodes("WOWKnownDlls32b").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "WOWKnownDlls32b"
    End If
    
    Dim sEFKD$, sContent$()
    sEFKD = Reg.GetString(HKEY_LOCAL_MACHINE, sSessionMan, "ExcludeFromKnownDlls", False)
    sContent = Split(sEFKD, Chr$(0))
    tvwMain.Nodes.Add "WOW", tvwChild, "ExcludeFromKnownDlls", "ExcludeFromKnownDlls", "reg"
    tvwMain.Nodes("ExcludeFromKnownDlls").Tag = "HKEY_LOCAL_MACHINE\" & sSessionMan
    For i = 0 To UBound(sContent)
        If Trim$(sContent(i)) <> vbNullString Then
            tvwMain.Nodes.Add "ExcludeFromKnownDlls", tvwChild, "ExcludeFromKnownDlls" & i, sContent(i), "dll"
            tvwMain.Nodes("ExcludeFromKnownDlls" & i).Tag = GuessFullpathFromAutorun(sContent(i))
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ExcludeFromKnownDlls").Children > 0 Then
        tvwMain.Nodes("ExcludeFromKnownDlls").Text = tvwMain.Nodes("ExcludeFromKnownDlls").Text & " (" & tvwMain.Nodes("ExcludeFromKnownDlls").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "ExcludeFromKnownDlls"
    End If
    
    If tvwMain.Nodes("WOW").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "WOW"
    End If

    If Not bShowHardware Then Exit Sub
    '----------------------------------------------------------------
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sWOW = "System\" & sHardwareCfgs(L) & "\Control\WOW"
        sSessionMan = "System\" & sHardwareCfgs(L) & "\Control\Session Manager"
        
        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "WOW", SEC_WOW, "registry", "registry"
        tvwMain.Nodes(sHardwareCfgs(L) & "WOW").Tag = "HKEY_LOCAL_MACHINE\" & sWOW
        
        For i = 0 To UBound(sVals)
            sCmd = ExpandEnvironmentVars(Reg.GetString(HKEY_LOCAL_MACHINE, sWOW, sVals(i)))
            If sCmd <> vbNullString Or bShowEmpty Then
                tvwMain.Nodes.Add sHardwareCfgs(L) & "WOW", tvwChild, sHardwareCfgs(L) & "WOW" & i, sVals(i) & " = " & sCmd, "exe", "exe"
                tvwMain.Nodes(sHardwareCfgs(L) & "WOW" & i).Tag = GuessFullpathFromAutorun(sCmd)
            End If
            If bSL_Abort Then Exit Sub
        Next i
        
        sKnownDlls = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sWOW, "KnownDlls"), " ")
        tvwMain.Nodes.Add sHardwareCfgs(L) & "WOW", tvwChild, sHardwareCfgs(L) & "WOWKnownDlls", "KnownDlls (16-bit)", "reg"
        tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls").Tag = "HKEY_LOCAL_MACHINE\" & sWOW
        For i = 0 To UBound(sKnownDlls)
            tvwMain.Nodes.Add sHardwareCfgs(L) & "WOWKnownDlls", tvwChild, sHardwareCfgs(L) & "WOWKnownDlls" & i, sKnownDlls(i), "dll"
            tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls" & i).Tag = GuessFullpathFromAutorun(sKnownDlls(i))
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls").Text = tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls").Children & ")"
            tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls").Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "WOWKnownDlls"
        End If
        
        sKnownDlls = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sSessionMan & "\KnownDlls"), "|")
        tvwMain.Nodes.Add sHardwareCfgs(L) & "WOW", tvwChild, sHardwareCfgs(L) & "WOWKnownDlls32b", "KnownDlls (32-bit)", "reg"
        tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls32b").Tag = "HKEY_LOCAL_MACHINE\" & sSessionMan & "\KnownDlls"
        For i = 0 To UBound(sKnownDlls)
            sCmd = Mid$(sKnownDlls(i), InStr(sKnownDlls(i), " = ") + 3)
            tvwMain.Nodes.Add sHardwareCfgs(L) & "WOWKnownDlls32b", tvwChild, sHardwareCfgs(L) & "WOWKnownDlls32b" & i, sCmd, "dll"
            tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls32b" & i).Tag = GuessFullpathFromAutorun(sCmd)
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls32b").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls32b").Text = tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls32b").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls32b").Children & ")"
            tvwMain.Nodes(sHardwareCfgs(L) & "WOWKnownDlls32b").Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "WOWKnownDlls32b"
        End If
        
        sEFKD = Reg.GetString(HKEY_LOCAL_MACHINE, sSessionMan, "ExcludeFromKnownDlls", False)
        sContent = Split(sEFKD, Chr$(0))
        tvwMain.Nodes.Add sHardwareCfgs(L) & "WOW", tvwChild, sHardwareCfgs(L) & "ExcludeFromKnownDlls", "ExcludeFromKnownDlls", "reg"
        tvwMain.Nodes(sHardwareCfgs(L) & "ExcludeFromKnownDlls").Tag = "HKEY_LOCAL_MACHINE\" & sSessionMan
        For i = 0 To UBound(sContent)
            If Trim$(sContent(i)) <> vbNullString Then
                tvwMain.Nodes.Add sHardwareCfgs(L) & "ExcludeFromKnownDlls", tvwChild, sHardwareCfgs(L) & "ExcludeFromKnownDlls" & i, sContent(i), "dll"
                tvwMain.Nodes(sHardwareCfgs(L) & "ExcludeFromKnownDlls" & i).Tag = GuessFullpathFromAutorun(sContent(i))
            End If
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "ExcludeFromKnownDlls").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "ExcludeFromKnownDlls").Text = tvwMain.Nodes(sHardwareCfgs(L) & "ExcludeFromKnownDlls").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "ExcludeFromKnownDlls").Children & ")"
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "ExcludeFromKnownDlls"
        End If
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "WOW").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove sHardwareCfgs(L) & "WOW"
        End If
    Next L
    AppendErrorLogCustom "EnumWOW - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumWOW"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumCmdProcessorAutorun()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumCmdProcessorAutorun - Begin"
    
    Dim sCmd$, sCmdKey$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "CmdProcAutorun", SEC_CMDPROC, "cmd", "cmd"
    sCmdKey = "Software\Microsoft\Command Processor"
    
    sCmd = Reg.GetString(HKEY_CURRENT_USER, sCmdKey, "AutoRun")
    If sCmd <> vbNullString Or bShowEmpty Then
        tvwMain.Nodes.Add "CmdProcAutorun", tvwChild, "CmdProcAutorunUser", "User autorun = " & sCmd, "exe", "exe"
        tvwMain.Nodes("CmdProcAutorunUser").Tag = "HKEY_CURRENT_USER\" & sCmdKey
    End If
    sCmd = Reg.GetString(HKEY_LOCAL_MACHINE, sCmdKey, "AutoRun")
    If sCmd <> vbNullString Or bShowEmpty Then
        tvwMain.Nodes.Add "CmdProcAutorun", tvwChild, "CmdProcAutorunSystem", "System autorun = " & sCmd, "exe", "exe"
        tvwMain.Nodes("CmdProcAutorunSystem").Tag = "HKEY_LOCAL_MACHINE\" & sCmdKey
    End If
    
    If tvwMain.Nodes("CmdProcAutorun").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "CmdProcAutorun"
    End If

    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "CmdProcAutorun", SEC_CMDPROC, "cmd"
    
            sCmd = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sCmdKey, "AutoRun")
            If sCmd <> vbNullString Or bShowEmpty Then
                tvwMain.Nodes.Add sUsernames(L) & "CmdProcAutorun", tvwChild, sUsernames(L) & "CmdProcAutorunUser", "User autorun = " & sCmd, "exe", "exe"
                tvwMain.Nodes(sUsernames(L) & "CmdProcAutorunUser").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sCmdKey
            End If
            
            If tvwMain.Nodes(sUsernames(L) & "CmdProcAutorun").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "CmdProcAutorun"
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumCmdProcessorAutorun - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumCmdProcessorAutorun"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumImageFileExecution()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumImageFileExecution - Begin"

    Dim sKeys$(), i&, sIFE$, sFile$
    If bSL_Abort Then Exit Sub
    sIFE = "Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
    tvwMain.Nodes.Add "System", tvwChild, "ImageFileExecution", SEC_IMAGEFILEEXECUTION, "explorer", "explorer"
    tvwMain.Nodes("ImageFileExecution").Tag = "HKEY_LOCAL_MACHINE\" & sIFE
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sIFE), "|")
    For i = 0 To UBound(sKeys)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sIFE & "\" & sKeys(i), "Debugger")
        If sFile <> vbNullString Then
            tvwMain.Nodes.Add "ImageFileExecution", tvwChild, "ImageFileExecution" & i, sKeys(i) & " = " & sFile, "exe", "exe"
            tvwMain.Nodes("ImageFileExecution" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ImageFileExecution").Children > 0 Then
        tvwMain.Nodes("ImageFileExecution").Text = tvwMain.Nodes("ImageFileExecution").Text & " (" & tvwMain.Nodes("ImageFileExecution").Children & ")"
        tvwMain.Nodes("ImageFileExecution").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "ImageFileExecution"
    End If
    
    '----------------------------------------------------------------
    'no per-user, this is system-wide
    AppendErrorLogCustom "EnumImageFileExecution - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumImageFileExecution"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumContextMenuHandlers()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumContextMenuHandlers - Begin"
    
    Dim sKeys$(), sObjects$(), i&, j&, sCLSID$, sFile$, sDummy$, sName$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "ContextMenuHandlers", SEC_CONTEXTMENUHANDLERS, "explorer", "explorer"
    
    ReDim sObjects(10)
    sObjects(0) = "*"
    sObjects(1) = "Drive"
    sObjects(2) = "Folder"
    sObjects(3) = "CompressedFolder"
    sObjects(4) = "Directory"
    sObjects(5) = "Directory\Background"
    sObjects(6) = "file"
    sObjects(7) = "ChannelShortcut"
    sObjects(8) = "InternetShortcut"
    sObjects(9) = "Printer"
    sObjects(10) = "AllFileSystemObjects"
    
    For j = 0 To UBound(sObjects)
        tvwMain.Nodes.Add "ContextMenuHandlers", tvwChild, "ContextMenuHandlers" & j, sObjects(j), "reg"
        tvwMain.Nodes("ContextMenuHandlers" & j).Tag = "HKEY_CLASSES_ROOT\" & sObjects(j) & "\shellex\ContextMenuHandlers"
        
        sKeys = Split(Reg.EnumSubKeys(HKEY_CLASSES_ROOT, sObjects(j) & "\shellex\ContextMenuHandlers"), "|")
        For i = 0 To UBound(sKeys)
            sCLSID = Reg.GetString(HKEY_CLASSES_ROOT, sObjects(j) & "\shellex\ContextMenuHandlers\" & sKeys(i), vbNullString)
            sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString)
            sFile = GetLongFilename(sFile)
            sName = sKeys(i)
            If sName = vbNullString Or InStr(sName, "{") = 1 Then
                sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i), vbNullString)
                If sName = vbNullString Then sName = "(no name)"
            End If
            
            'retarded 'start menu pin' uses name and clsid wrong way around
            If sFile = vbNullString Then
                sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i) & "\InprocServer32", vbNullString)
                sFile = GetLongFilename(sFile)
                sDummy = sKeys(i)
                sKeys(i) = sCLSID
                sCLSID = sDummy
                sDummy = vbNullString
            End If
            
            If sFile <> vbNullString Then
                sFile = ExpandEnvironmentVars(sFile)
                sFile = GetLongFilename(sFile)
                If Not bShowCLSIDs Then
                    tvwMain.Nodes.Add "ContextMenuHandlers" & j, tvwChild, "ContextMenuHandlers" & j & "sub" & i, sName & " = " & sFile, "dll", "dll"
                Else
                    tvwMain.Nodes.Add "ContextMenuHandlers" & j, tvwChild, "ContextMenuHandlers" & j & "sub" & i, sName & " = " & sCLSID & " = " & sFile, "dll", "dll"
                End If
                tvwMain.Nodes("ContextMenuHandlers" & j & "sub" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes("ContextMenuHandlers" & j).Children > 0 Then
            tvwMain.Nodes("ContextMenuHandlers" & j).Text = tvwMain.Nodes("ContextMenuHandlers" & j).Text & " (" & tvwMain.Nodes("ContextMenuHandlers" & j).Children & ")"
            tvwMain.Nodes("ContextMenuHandlers" & j).Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove "ContextMenuHandlers" & j
        End If
    Next j
    
    '----------------------------------------------------------------
    'no per-user, this is system-wide
    AppendErrorLogCustom "EnumContextMenuHandlers - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumContextMenuHandlers"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumColumnHandlers()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumColumnHandlers - Begin"

    'HKCR\Folder\shellex\ColumnHandlers\*
    Dim sKeys$(), sTheKey$, i&, sCLSID$, sName$, sFile$
    If bSL_Abort Then Exit Sub
    sTheKey = "Folder\shellex\ColumnHandlers"
    tvwMain.Nodes.Add "System", tvwChild, "ColumnHandlers", SEC_COLUMNHANDLERS, "explorer"
    tvwMain.Nodes("ColumnHandlers").Tag = "HKEY_CLASSES_ROOT\" & sTheKey
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_CLASSES_ROOT, sTheKey), "|")
    For i = 0 To UBound(sKeys)
        sCLSID = sKeys(i)
        'the name is blank, but try it anyway
        sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
        If sName = vbNullString Then sName = "(no name)"
        If sFile = vbNullString Then sFile = "(no file)"
        
        If bShowCLSIDs Then
            tvwMain.Nodes.Add "ColumnHandlers", tvwChild, "ColumnHandlers" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
        Else
            tvwMain.Nodes.Add "ColumnHandlers", tvwChild, "ColumnHandlers" & i, sName & " - " & sFile, "dll"
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ColumnHandlers").Children > 0 Then
        tvwMain.Nodes("ColumnHandlers").Text = tvwMain.Nodes("ColumnHandlers").Text & " (" & tvwMain.Nodes("ColumnHandlers").Children & ")"
        tvwMain.Nodes("ColumnHandlers").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "ColumnHandlers"
    End If
    '----------------------------------------------------------------
    'system-wide
    AppendErrorLogCustom "EnumColumnHandlers - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumColumnHandlers"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumShellExecuteHooks()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumShellExecuteHooks - Begin"
    Dim sSEH$
    If bSL_Abort Then Exit Sub
    sSEH = "Software\Microsoft\Windows\CurrentVersion\explorer\ShellExecuteHooks"
    Dim sVals$(), i&, sName$, sFile$
    tvwMain.Nodes.Add "System", tvwChild, "ShellExecuteHooks", SEC_SHELLEXECUTEHOOKS, "explorer", "explorer"
    tvwMain.Nodes("ShellExecuteHooks").Tag = "HKEY_LOCAL_MACHINE\" & sSEH
    
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sSEH), "|")
    For i = 0 To UBound(sVals)
        If Right$(sVals(i), 3) <> " = " Then
            sName = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
        End If
        sVals(i) = Mid$(sVals(i), 1, InStr(sVals(i), " = ") - 1)
        If sName = vbNullString Then
            sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sVals(i), vbNullString)
            If sName = vbNullString Then sName = "(no name)"
        End If
        sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sVals(i) & "\InprocServer32", vbNullString)
        sFile = GetLongFilename(sFile)
        If sFile <> vbNullString Then
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "ShellExecuteHooks", tvwChild, "ShellExecuteHooks" & i, sName & " = " & sFile, "dll", "dll"
            Else
                tvwMain.Nodes.Add "ShellExecuteHooks", tvwChild, "ShellExecuteHooks" & i, sName & " = " & sVals(i) & " = " & sFile, "dll", "dll"
            End If
            tvwMain.Nodes("ShellExecuteHooks" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ShellExecuteHooks").Children > 0 Then
        tvwMain.Nodes("ShellExecuteHooks").Text = tvwMain.Nodes("ShellExecuteHooks").Text & " (" & tvwMain.Nodes("ShellExecuteHooks").Children & ")"
        tvwMain.Nodes("ShellExecuteHooks").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "ShellExecuteHooks"
    End If
    
    '----------------------------------------------------------------
    'no per-user, this is system-wide
    AppendErrorLogCustom "EnumShellExecuteHooks - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumShellExecuteHooks"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumShellExtensions()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumShellExtensions - Begin"
    'HKLM\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved
    Dim sKey$, sVals$(), sKeys$(), i&, sName$, sFile$
    If bSL_Abort Then Exit Sub
    sKey = "Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved"
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sKey), "|")
    tvwMain.Nodes.Add "System", tvwChild, "ShellExts", SEC_SHELLEXT, "explorer"
    
    tvwMain.Nodes.Add "ShellExts", tvwChild, "ShellExtsSystem", "All users", "users"
    tvwMain.Nodes("ShellExtsSystem").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    For i = 0 To UBound(sVals)
        sName = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
        sVals(i) = Left$(sVals(i), InStr(sVals(i), " = ") - 1)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sVals(i) & "\InprocServer32", vbNullString))
        sFile = GetLongFilename(sFile)
        If Not bShowCLSIDs Then
            tvwMain.Nodes.Add "ShellExtsSystem", tvwChild, "ShellExtsSystem" & i, sName & " - " & sFile, "reg"
        Else
            tvwMain.Nodes.Add "ShellExtsSystem", tvwChild, "ShellExtsSystem" & i, sName & " - " & sVals(i) & " - " & sFile, "reg"
        End If
        tvwMain.Nodes("ShellExtsSystem" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ShellExtsSystem").Children > 0 Then
        tvwMain.Nodes("ShellExtsSystem").Text = tvwMain.Nodes("ShellExtsSystem").Text & " (" & tvwMain.Nodes("ShellExtsSystem").Children & ")"
        tvwMain.Nodes("ShellExtsSystem").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "ShellExtsSystem"
    End If
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sKey), "|")
    tvwMain.Nodes.Add "ShellExts", tvwChild, "ShellExtsUser", "This user", "user"
    tvwMain.Nodes("ShellExtsUser").Tag = "HKEY_CURRENT_USER\" & sKey
    For i = 0 To UBound(sKeys)
        sName = Reg.GetString(HKEY_CURRENT_USER, sKey & "\" & sKeys(i), vbNullString)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i) & "\InprocServer32", vbNullString))
        sFile = GetLongFilename(sFile)
        If Not bShowCLSIDs Then
            tvwMain.Nodes.Add "ShellExtsUser", tvwChild, "ShellExtsUser" & i, sName & " - " & sFile, "reg"
        Else
            tvwMain.Nodes.Add "ShellExtsUser", tvwChild, "ShellExtsUser" & i, sName & " - " & sKeys(i) & " - " & sFile, "reg"
        End If
        tvwMain.Nodes("ShellExtsUser" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("ShellExtsUser").Children > 0 Then
        tvwMain.Nodes("ShellExtsUser").Text = tvwMain.Nodes("ShellExtsUser").Text & " (" & tvwMain.Nodes("ShellExtsUser").Children & ")"
        tvwMain.Nodes("ShellExtsUser").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "ShellExtsUser"
    End If
    
    If tvwMain.Nodes("ShellExts").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ShellExts"
    End If
    
    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "ShellExtsUser", SEC_SHELLEXT, "explorer"
    
            sKeys = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\" & sKey), "|")
            tvwMain.Nodes(sUsernames(L) & "ShellExtsUser").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKey
            For i = 0 To UBound(sKeys)
                sName = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sKey & "\" & sKeys(i), vbNullString)
                sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i) & "\InprocServer32", vbNullString))
                sFile = GetLongFilename(sFile)
                If Not bShowCLSIDs Then
                    tvwMain.Nodes.Add sUsernames(L) & "ShellExtsUser", tvwChild, sUsernames(L) & "ShellExtsUser" & i, sName & " - " & sFile, "reg"
                Else
                    tvwMain.Nodes.Add sUsernames(L) & "ShellExtsUser", tvwChild, sUsernames(L) & "ShellExtsUser" & i, sName & " - " & sKeys(i) & " - " & sFile, "reg"
                End If
                tvwMain.Nodes(sUsernames(L) & "ShellExtsUser" & i).Tag = GuessFullpathFromAutorun(sFile)
                If bSL_Abort Then Exit Sub
            Next i
            If tvwMain.Nodes(sUsernames(L) & "ShellExtsUser").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "ShellExtsUser").Text = tvwMain.Nodes(sUsernames(L) & "ShellExtsUser").Text & " (" & tvwMain.Nodes(sUsernames(L) & "ShellExtsUser").Children & ")"
                tvwMain.Nodes(sUsernames(L) & "ShellExtsUser").Sorted = True
            Else
                If Not bShowEmpty Then tvwMain.Nodes.Remove sUsernames(L) & "ShellExtsUser"
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumShellExtensions - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumShellExtensions"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumURLSearchHooks()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumURLSearchHooks - Begin"
    Dim sKey$
    If bSL_Abort Then Exit Sub
    sKey = "Software\Microsoft\Internet Explorer\URLSearchHooks"
    Dim sVals$(), i&, sCLSID$, sName$, sFile$
    tvwMain.Nodes.Add "System", tvwChild, "URLSearchHooks", SEC_URLSEARCHHOOKS, "msie"
    
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sKey), "|")
    tvwMain.Nodes.Add "URLSearchHooks", tvwChild, "URLSearchHooksSystem", "All users", "users"
    tvwMain.Nodes("URLSearchHooksSystem").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    For i = 0 To UBound(sVals)
        sCLSID = sVals(i)
        sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
        If sFile <> vbNullString Then
            If sName = vbNullString Then sName = "(no name)"
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "URLSearchHooksSystem", tvwChild, "URLSearchHooksSystem" & i, sName & " - " & sFile, "reg"
            Else
                tvwMain.Nodes.Add "URLSearchHooksSystem", tvwChild, "URLSearchHooksSystem" & i, sName & " - " & sCLSID & " - " & sFile, "reg"
            End If
            tvwMain.Nodes("URLSearchHooksSystem" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("URLSearchHooksSystem").Children > 0 Then
        tvwMain.Nodes("URLSearchHooksSystem").Text = tvwMain.Nodes("URLSearchHooksSystem").Text & " (" & tvwMain.Nodes("URLSearchHooksSystem").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "URLSearchHooksSystem"
    End If
    
    sVals = Split(RegEnumValues(HKEY_CURRENT_USER, sKey), "|")
    tvwMain.Nodes.Add "URLSearchHooks", tvwChild, "URLSearchHooksUser", "This user", "user"
    tvwMain.Nodes("URLSearchHooksUser").Tag = "HKEY_CURRENT_USER\" & sKey
    For i = 0 To UBound(sVals)
        sCLSID = Left$(sVals(i), InStr(sVals(i), " = ") - 1)
        sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
        If sFile <> vbNullString Then
            If sName = vbNullString Then sName = "(no name)"
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "URLSearchHooksUser", tvwChild, "URLSearchHooksUser" & i, sName & " - " & sFile, "reg"
            Else
                tvwMain.Nodes.Add "URLSearchHooksUser", tvwChild, "URLSearchHooksUser" & i, sName & " - " & sCLSID & " - " & sFile, "reg"
            End If
            tvwMain.Nodes("URLSearchHooksUser" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("URLSearchHooksUser").Children > 0 Then
        tvwMain.Nodes("URLSearchHooksUser").Text = tvwMain.Nodes("URLSearchHooksUser").Text & " (" & tvwMain.Nodes("URLSearchHooksUser").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "URLSearchHooksUser"
    End If
        
    If tvwMain.Nodes("URLSearchHooks").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "URLSearchHooks"
    End If

    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "URLSearchHooksUser", SEC_URLSEARCHHOOKS, "msie"
            
            sVals = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\" & sKey), "|")
            'tvwMain.Nodes.Add sUsernames(l) & "URLSearchHooks", tvwChild, sUsernames(l) & "URLSearchHooksUser", "This user", "user"
            tvwMain.Nodes(sUsernames(L) & "URLSearchHooksUser").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKey
            For i = 0 To UBound(sVals)
                sCLSID = Left$(sVals(i), InStr(sVals(i), " = ") - 1)
                sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
                sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
                If sFile <> vbNullString Then
                    If sName = vbNullString Then sName = "(no name)"
                    If Not bShowCLSIDs Then
                        tvwMain.Nodes.Add sUsernames(L) & "URLSearchHooksUser", tvwChild, sUsernames(L) & "URLSearchHooksUser" & i, sName & " - " & sFile, "reg"
                    Else
                        tvwMain.Nodes.Add sUsernames(L) & "URLSearchHooksUser", tvwChild, sUsernames(L) & "URLSearchHooksUser" & i, sName & " - " & sCLSID & " - " & sFile, "reg"
                    End If
                    tvwMain.Nodes(sUsernames(L) & "URLSearchHooksUser" & i).Tag = GuessFullpathFromAutorun(sFile)
                End If
                If bSL_Abort Then Exit Sub
            Next i
            If tvwMain.Nodes(sUsernames(L) & "URLSearchHooksUser").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "URLSearchHooksUser").Text = tvwMain.Nodes(sUsernames(L) & "URLSearchHooksUser").Text & " (" & tvwMain.Nodes(sUsernames(L) & "URLSearchHooksUser").Children & ")"
            Else
                If Not bShowEmpty Then tvwMain.Nodes.Remove sUsernames(L) & "URLSearchHooksUser"
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumURLSearchHooks - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumURLSearchHooks"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumKillBits()
    RegEnumKillBits tvwMain
End Sub

Private Sub EnumAccUtilManager()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumAccUtilManager - Begin"
    Dim sAUM$
    If bSL_Abort Then Exit Sub
    sAUM = "Software\Microsoft\Windows NT\CurrentVersion\Accessibility\Utility Manager"
    Dim sKeys$(), i&, lStart&, sFile$
    tvwMain.Nodes.Add "System", tvwChild, "UtilityManager", SEC_UTILMANAGER, "registry", "registry"
    tvwMain.Nodes("UtilityManager").Tag = "HKEY_LOCAL_MACHINE\" & sAUM
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sAUM), "|")
    For i = 0 To UBound(sKeys)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sAUM & "\" & sKeys(i), "Application Path")
        lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, sAUM & "\" & sKeys(i), "Start with Windows")
        If sFile <> vbNullString And lStart = 1 Then
            tvwMain.Nodes.Add "UtilityManager", tvwChild, "UtilityManager" & i, sKeys(i) & " = " & sFile, "exe", "exe"
            tvwMain.Nodes("UtilityManager" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("UtilityManager").Children > 0 Then
        tvwMain.Nodes("UtilityManager").Text = tvwMain.Nodes("UtilityManager").Text & " (" & tvwMain.Nodes("UtilityManager").Children & ")"
        tvwMain.Nodes("UtilityManager").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "UtilityManager"
    End If
    
    '----------------------------------------------------------------
    'nothing - this is system-wide
    AppendErrorLogCustom "EnumAccUtilManager - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumAccUtilManager"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumJobs()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumJobs - Begin"
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "TaskScheduler", SEC_TASKSCHEDULER, "folder"
    tvwMain.Nodes.Add "TaskScheduler", tvwChild, "TaskSchedulerJobs", "Jobs", "folder"
    tvwMain.Nodes.Add "TaskScheduler", tvwChild, "TaskSchedulerJobsSystem", "Jobs (system32 folder)", "folder"
    tvwMain.Nodes("TaskSchedulerJobs").Tag = sWinDir & "\Tasks"
    tvwMain.Nodes("TaskSchedulerJobsSystem").Tag = sSysDir & "\Tasks"
    
    Dim sFiles$(), i&
    sFiles = Split(EnumFiles(sWinDir & "\Tasks"), "|")
    For i = 0 To UBound(sFiles)
        If InStr(1, sFiles(i), ".job", vbTextCompare) = Len(sFiles(i)) - 3 Then
            tvwMain.Nodes.Add "TaskSchedulerJobs", tvwChild, "TaskSchedulerJobs" & i, sFiles(i), "bat", "bat"
            tvwMain.Nodes("TaskSchedulerJobs" & i).Tag = sWinDir & "\Tasks\" & sFiles(i)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    sFiles = Split(EnumFiles(sSysDir & "\Tasks"), "|")
    For i = 0 To UBound(sFiles)
        If Len(sFiles(i)) > 3 And InStr(1, sFiles(i), ".job", vbTextCompare) = Len(sFiles(i)) - 3 Then
            tvwMain.Nodes.Add "TaskSchedulerJobsSystem", tvwChild, "TaskSchedulerJobsSystem" & i, sFiles(i), "bat", "bat"
            tvwMain.Nodes("TaskSchedulerJobsSystem" & i).Tag = sSysDir & "\Tasks\" & sFiles(i)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("TaskSchedulerJobs").Children > 0 Then
        tvwMain.Nodes("TaskSchedulerJobs").Text = tvwMain.Nodes("TaskSchedulerJobs").Text & " (" & tvwMain.Nodes("TaskSchedulerJobs").Children & ")"
        tvwMain.Nodes("TaskSchedulerJobs").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "TaskSchedulerJobs"
    End If
    If tvwMain.Nodes("TaskSchedulerJobsSystem").Children > 0 Then
        tvwMain.Nodes("TaskSchedulerJobsSystem").Text = tvwMain.Nodes("TaskSchedulerJobsSystem").Text & " (" & tvwMain.Nodes("TaskSchedulerJobsSystem").Children & ")"
        tvwMain.Nodes("TaskSchedulerJobsSystem").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "TaskSchedulerJobsSystem"
    End If
    If tvwMain.Nodes("TaskScheduler").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "TaskScheduler"
    End If
    
    
    '----------------------------------------------------------------
    'nothing - this is system-wide
    AppendErrorLogCustom "EnumJobs - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumJobs"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumNTScripts()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumNTScripts - Begin"
    
    Dim sScripts$, sContents$(), i&
    If bSL_Abort Then Exit Sub
    sScripts = "Software\Policies\Microsoft\Windows\System\Scripts"
    tvwMain.Nodes.Add "System", tvwChild, "ScriptPolicies", SEC_SCRIPTPOLICIES, "script"
    
    Dim sLogon$, sLogOff$, sStartup$, sShutdown$
    sLogon = BuildPath(Reg.GetString(HKEY_CURRENT_USER, sScripts, "Logon"), "scripts.ini")
    sLogOff = BuildPath(Reg.GetString(HKEY_CURRENT_USER, sScripts, "Logoff"), "scripts.ini")
    sStartup = BuildPath(Reg.GetString(HKEY_LOCAL_MACHINE, sScripts, "Startup"), "scripts.ini")
    sShutdown = BuildPath(Reg.GetString(HKEY_LOCAL_MACHINE, sScripts, "Shutdown"), "scripts.ini")
    
    If sLogon = sStartup Then sLogon = vbNullString
    If sLogOff = sShutdown Then sLogOff = vbNullString
    
    If FileExists(sLogon) Then
        sContents = Split(InputFile(sLogon), vbCrLf)
        If UBound(sContents) > -1 Or bShowEmpty Then
            tvwMain.Nodes.Add "ScriptPolicies", tvwChild, "ScriptPoliciesLogon", "User logon script", "ini", "ini"
            For i = 0 To UBound(sContents)
                If Trim$(sContents(i)) <> vbNullString Then
                    tvwMain.Nodes.Add "ScriptPoliciesLogon", tvwChild, "ScriptPoliciesLogon" & i, sContents(i), "text", "text"
                End If
                If bSL_Abort Then Exit Sub
            Next i
        End If
    End If
    If FileExists(sLogOff) Then
        sContents = Split(InputFile(sLogOff), vbCrLf)
        If UBound(sContents) > -1 Or bShowEmpty Then
            tvwMain.Nodes.Add "ScriptPolicies", tvwChild, "ScriptPoliciesLogoff", "User logon script", "ini", "ini"
            For i = 0 To UBound(sContents)
                If Trim$(sContents(i)) <> vbNullString Then
                    tvwMain.Nodes.Add "ScriptPoliciesLogoff", tvwChild, "ScriptPoliciesLogoff" & i, sContents(i), "text", "text"
                End If
                If bSL_Abort Then Exit Sub
            Next i
        End If
    End If
    If FileExists(sStartup) Then
        sContents = Split(InputFile(sStartup), vbCrLf)
        If UBound(sContents) > -1 Or bShowEmpty Then
            tvwMain.Nodes.Add "ScriptPolicies", tvwChild, "ScriptPoliciesStartup", "User logon script", "ini", "ini"
            For i = 0 To UBound(sContents)
                If Trim$(sContents(i)) <> vbNullString Then
                    tvwMain.Nodes.Add "ScriptPoliciesStartup", tvwChild, "ScriptPoliciesStartup" & i, sContents(i), "text", "text"
                End If
                If bSL_Abort Then Exit Sub
            Next i
        End If
    End If
    If FileExists(sShutdown) Then
        sContents = Split(InputFile(sShutdown), vbCrLf)
        If UBound(sContents) > -1 Or bShowEmpty Then
            tvwMain.Nodes.Add "ScriptPolicies", tvwChild, "ScriptPoliciesShutdown", "User logon script", "ini", "ini"
            For i = 0 To UBound(sContents)
                If Trim$(sContents(i)) <> vbNullString Then
                    tvwMain.Nodes.Add "ScriptPoliciesShutdown", tvwChild, "ScriptPoliciesShutdown" & i, sContents(i), "text", "text"
                End If
                If bSL_Abort Then Exit Sub
            Next i
        End If
    End If
    
    If tvwMain.Nodes("ScriptPolicies").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ScriptPolicies"
    End If
    
    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "ScriptPolicies", SEC_SCRIPTPOLICIES, "ini", "ini"
            
            sLogon = BuildPath(Reg.GetString(HKEY_CURRENT_USER, sScripts, "Logon"), "scripts.ini")
            sLogOff = BuildPath(Reg.GetString(HKEY_CURRENT_USER, sScripts, "Logoff"), "scripts.ini")
            
            If sLogon = sStartup Then sLogon = vbNullString
            If sLogOff = sShutdown Then sLogOff = vbNullString
            
            If FileExists(sLogon) Then
                sContents = Split(InputFile(sLogon), vbCrLf)
                If UBound(sContents) > -1 Or bShowEmpty Then
                    tvwMain.Nodes.Add sUsernames(L) & "ScriptPolicies", tvwChild, sUsernames(L) & "ScriptPoliciesLogon", "User logon script", "ini", "ini"
                    For i = 0 To UBound(sContents)
                        If Trim$(sContents(i)) <> vbNullString Then
                            tvwMain.Nodes.Add sUsernames(L) & "ScriptPoliciesLogon", tvwChild, sUsernames(L) & "ScriptPoliciesLogon" & i, sContents(i), "text", "text"
                        End If
                        If bSL_Abort Then Exit Sub
                    Next i
                End If
            End If
            If FileExists(sLogOff) Then
                sContents = Split(InputFile(sLogOff), vbCrLf)
                If UBound(sContents) > -1 Or bShowEmpty Then
                    tvwMain.Nodes.Add sUsernames(L) & "ScriptPolicies", tvwChild, sUsernames(L) & "ScriptPoliciesLogoff", "User logon script", "ini", "ini"
                    For i = 0 To UBound(sContents)
                        If Trim$(sContents(i)) <> vbNullString Then
                            tvwMain.Nodes.Add sUsernames(L) & "ScriptPoliciesLogoff", tvwChild, sUsernames(L) & "ScriptPoliciesLogoff" & i, sContents(i), "text", "text"
                        End If
                        If bSL_Abort Then Exit Sub
                    Next i
                End If
            End If
            
            If tvwMain.Nodes(sUsernames(L) & "ScriptPolicies").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "ScriptPolicies"
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumNTScripts - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumNTScripts"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumDisabled()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumDisabled - Begin"
    
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "DisabledEnums", "Protection & disabled items", "bad"
    
    Dim sUsername$, L&
    If bShowUsers Then
        For L = 0 To UBound(sUsernames)
            sUsername = MapSIDToUsername(sUsernames(L))
            If sUsername <> OSver.UserName And sUsername <> vbNullString Then
                tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "DisabledEnums", "Protection & disabled items", "bad"
            End If
        Next L
    End If
    If bShowHardware Then
        For L = 1 To UBound(sHardwareCfgs)
            tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "DisabledEnums", "Protection & disabled items", "bad"
        Next L
    End If
    
    '* hosts file
    'Loading... Hosts file
    Status Translate(912)
    DoTicks tvwMain
    EnumHostsFile
    DoTicks tvwMain, "HostsFile"
    UpdateProgressBar
    
    '* killbits
    DoTicks tvwMain
    EnumKillBits
    DoTicks tvwMain, "Killbits"
    UpdateProgressBar
    
    '* restricted sites
    'Status "Loading..."
    DoTicks tvwMain
    EnumZones
    DoTicks tvwMain, "Zones"
    UpdateProgressBar
    
    '* msconfig 9x autoruns
    'Loading... Msconfig 9x/ME disabled items
    Status Translate(913)
    DoTicks tvwMain
    EnumDisabledMsconfig9x
    DoTicks tvwMain, "msconfig9x"
    UpdateProgressBar
    
    '* msconfig xp autoruns
    'Loading... Msconfig XP disabled items
    Status Translate(914)
    DoTicks tvwMain
    EnumDisabledMsconfigXP
    DoTicks tvwMain, "msconfigxp"
    
    'Stopped/Disabled NT Services
    'Loading... Stopped/disabled Services
    Status Translate(915)
    DoTicks tvwMain
    EnumStoppedServices
    DoTicks tvwMain, "StoppedServices"
    
    'XP Security Center
    'Loading... Windows XP Security Center settings
    Status Translate(916)
    DoTicks tvwMain
    EnumXPSecurity
    DoTicks tvwMain, "XPSecurity"

    If bShowUsers Then
        For L = 0 To UBound(sUsernames)
            sUsername = MapSIDToUsername(sUsernames(L))
            If sUsername <> OSver.UserName Then
                If tvwMain.Nodes(sUsernames(L) & "DisabledEnums").Children = 0 And Not bShowEmpty Then
                    tvwMain.Nodes.Remove sUsernames(L) & "DisabledEnums"
                End If
            End If
        Next L
    End If
    If bShowHardware Then
        For L = 1 To UBound(sHardwareCfgs)
            If tvwMain.Nodes(sHardwareCfgs(L) & "DisabledEnums").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sHardwareCfgs(L) & "DisabledEnums"
            End If
        Next L
    End If
    AppendErrorLogCustom "EnumDisabled - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumDisabled"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumDisabledMsconfig9x()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumDisabledMsconfig9x - Begin"
    
    Dim sNames$(), sKeys$(), i&, j&, sValues$()
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "DisabledEnums", tvwChild, "msconfig9x", SEC_MSCONFIG9X, "registry"
    
    ReDim sNames(1)
    sNames(0) = "Run"
    sNames(1) = "RunServices"
    
    ReDim sKeys(1)
    sKeys(0) = "Software\Microsoft\Windows\CurrentVersion\Run-"
    sKeys(1) = "Software\Microsoft\Windows\CurrentVersion\RunServices-"
    
    For i = 0 To UBound(sKeys)
        sValues = Split(RegEnumValues(HKEY_CURRENT_USER, sKeys(i)), "|")
        tvwMain.Nodes.Add "msconfig9x", tvwChild, "msconfig9xUser" & i, "User " & sNames(i), "registry", "registry"
        tvwMain.Nodes("msconfig9xUser" & i).Tag = "HKEY_CURRENT_USER\" & sKeys(i)
        For j = 0 To UBound(sValues)
            tvwMain.Nodes.Add "msconfig9xUser" & i, tvwChild, "msconfig9xUser" & i & "Val" & j, sValues(j), "reg", "reg"
        Next j
        If tvwMain.Nodes("msconfig9xUser" & i).Children > 0 Then
            tvwMain.Nodes("msconfig9xUser" & i).Text = tvwMain.Nodes("msconfig9xUser" & i).Text & " (" & tvwMain.Nodes("msconfig9xUser" & i).Children & ")"
            tvwMain.Nodes("msconfig9xUser" & i).Sorted = True
        Else
            If Not bShowEmpty Then
                tvwMain.Nodes.Remove ("msconfig9xUser" & i)
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next i
    For i = 0 To UBound(sKeys)
        sValues = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sKeys(i)), "|")
        tvwMain.Nodes.Add "msconfig9x", tvwChild, "msconfig9xSystem" & i, "System " & sNames(i), "registry", "registry"
        tvwMain.Nodes("msconfig9xSystem" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKeys(i)
        For j = 0 To UBound(sValues)
            tvwMain.Nodes.Add "msconfig9xSystem" & i, tvwChild, "msconfig9xSystem" & i & "Val" & j, sValues(j), "reg", "reg"
        Next j
        If tvwMain.Nodes("msconfig9xSystem" & i).Children > 0 Then
            tvwMain.Nodes("msconfig9xSystem" & i).Text = tvwMain.Nodes("msconfig9xSystem" & i).Text & " (" & tvwMain.Nodes("msconfig9xSystem" & i).Children & ")"
            tvwMain.Nodes("msconfig9xSystem" & i).Sorted = True
        Else
            If Not bShowEmpty Then
                tvwMain.Nodes.Remove ("msconfig9xSystem" & i)
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("msconfig9x").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "msconfig9x"
    End If

    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add sUsernames(L) & "DisabledEnums", tvwChild, sUsernames(L) & "msconfig9x", SEC_MSCONFIG9X, "registry"
            
            For i = 0 To UBound(sKeys)
                sValues = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\" & sKeys(i)), "|")
                tvwMain.Nodes.Add sUsernames(L) & "msconfig9x", tvwChild, sUsernames(L) & "msconfig9xUser" & i, "User " & sNames(i), "registry", "registry"
                tvwMain.Nodes(sUsernames(L) & "msconfig9xUser" & i).Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKeys(i)
                For j = 0 To UBound(sValues)
                    tvwMain.Nodes.Add sUsernames(L) & "msconfig9xUser" & i, tvwChild, sUsernames(L) & "msconfig9xUser" & i & "Val" & j, sValues(j), "reg", "reg"
                Next j
                If tvwMain.Nodes(sUsernames(L) & "msconfig9xUser" & i).Children > 0 Then
                    tvwMain.Nodes(sUsernames(L) & "msconfig9xUser" & i).Text = tvwMain.Nodes(sUsernames(L) & "msconfig9xUser" & i).Text & " (" & tvwMain.Nodes(sUsernames(L) & "msconfig9xUser" & i).Children & ")"
                    tvwMain.Nodes(sUsernames(L) & "msconfig9xUser" & i).Sorted = True
                Else
                    If Not bShowEmpty Then
                        tvwMain.Nodes.Remove (sUsernames(L) & "msconfig9xUser" & i)
                    End If
                End If
                If bSL_Abort Then Exit Sub
            Next i
            
            If tvwMain.Nodes(sUsernames(L) & "msconfig9x").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "msconfig9x"
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumDisabledMsconfig9x - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumDisabledMsconfig9x"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumDisabledMsconfigXP()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumDisabledMsconfigXP - Begin"
    
    'HKCU\Run, HKLM\Run
    'HKLM\Software\Microsoft\Shared Tools\MSConfig\startupreg
    Dim sKey$
    If bSL_Abort Then Exit Sub
    sKey = "Software\Microsoft\Shared Tools\MSConfig\startupreg"
    Dim sKeys$(), sSubkeys$(), j&, i&, sFile$
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey), "|")
    tvwMain.Nodes.Add "DisabledEnums", tvwChild, "msconfigxp", SEC_MSCONFIGXP, "registry"
    tvwMain.Nodes("msconfigxp").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    For i = 0 To UBound(sKeys)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "command")
        If sFile <> vbNullString Then
            tvwMain.Nodes.Add "msconfigxp", tvwChild, "msconfigxp" & i, sKeys(i) & " = " & sFile, "reg"
            tvwMain.Nodes("msconfigxp" & i).Tag = GuessFullpathFromAutorun(sFile)
        Else
            'get subkeys and get file there
            sSubkeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i)), "|")
            If UBound(sSubkeys) > -1 Then
                For j = 0 To UBound(sSubkeys)
                    sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i) & "\" & sSubkeys(j), "command")
                    If sFile <> vbNullString Then
                        tvwMain.Nodes.Add "msconfigxp", tvwChild, "msconfigxp" & i & "s" & j, sSubkeys(j) & " = " & sFile, "reg"
                        tvwMain.Nodes("msconfigxp" & i & "s" & j).Tag = GuessFullpathFromAutorun(sFile)
                    End If
                Next j
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("msconfigxp").Children > 0 Then
        tvwMain.Nodes("msconfigxp").Text = tvwMain.Nodes("msconfigxp").Text & " (" & tvwMain.Nodes("msconfigxp").Children & ")"
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "msconfigxp"
        End If
    End If
    
    '----------------------------------------------------------------
    'nothing - this is system-wide
    AppendErrorLogCustom "EnumDisabledMsconfigXP - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumDisabledMsconfigXP"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumHostsFile()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumHostsFile - Begin"
    
    Dim sHostsFile$
    If bSL_Abort Then Exit Sub
    If bIsWinNT Then
        sHostsFile = BuildPath(sSysDir, "drivers\etc\hosts")
    Else
        sHostsFile = BuildPath(sWinDir, "hosts")
    End If
    If Not FileExists(sHostsFile) Then
        If bShowEmpty Then tvwMain.Nodes.Add "DisabledEnums", tvwChild, "HostsFile", SEC_HOSTSFILE, "text"
        Exit Sub
    End If
    
    Dim sContents$(), i&, sIP$, sHost$, j&
    sContents = Split(InputFile(sHostsFile), vbCrLf)
    If UBound(sContents) > 1000 And Not bShowLargeHosts Then
        'big hosts file - skip it
        frmStartupList2.ShowError "Skipping hosts file, because it is over 1000 lines long. (file is " & UBound(sContents) & " lines, totalling " & Int(FileLen(sHostsFile) / 1024) & " kb)"
        Exit Sub
    End If
    If UBound(sContents) > -1 Then
        If InStr(sContents(0), vbCr) > 0 Then
            sContents = Split(Join(sContents, vbCr), vbCr)
        End If
        If InStr(sContents(0), vbLf) > 0 Then
            sContents = Split(Join(sContents, vbLf), vbLf)
        End If
        
        tvwMain.Nodes.Add "DisabledEnums", tvwChild, "HostsFile", "Hosts file", "text"
        tvwMain.Nodes("HostsFile").Tag = sHostsFile
        For i = 0 To UBound(sContents)
            sContents(i) = Replace$(sContents(i), vbTab, " ")
            If InStr(sContents(i), "#") > 0 Then
                sContents(i) = Left$(sContents(i), InStr(sContents(i), "#") - 1)
            End If
            If Trim$(sContents(i)) <> vbNullString Then
                If InStr(sContents(i), " ") > 1 Then
                    sIP = Trim$(Left$(sContents(i), InStr(sContents(i), " ") - 1))
                    sHost = Trim$(Mid$(sContents(i), InStr(sContents(i), " ") + 1))
                    If Not NodeExists("HostsFile" & sIP) Then
                        tvwMain.Nodes.Add "HostsFile", tvwChild, "HostsFile" & sIP, sIP, "internet"
                    End If
                    tvwMain.Nodes.Add "HostsFile" & sIP, tvwChild, "HostsFile" & sIP & j, sHost, "internet"
                    j = j + 1
                End If
            End If
            If bShowLargeHosts And i Mod 100 = 0 Then
                'Loading... Hosts file
                Status Translate(912) & " (" & Int(CLng(i) * 100 / UBound(sContents)) & "%, " & i & " lines)"
            End If
            If bSL_Abort Then Exit Sub
        Next i
        
        If tvwMain.Nodes("HostsFile").Children > 0 Then
            tvwMain.Nodes("HostsFile").Text = tvwMain.Nodes("HostsFile").Text & " (" & j & ")"
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove "HostsFile"
        End If
    End If
    
    '----------------------------------------------------------------
    'nothing - this is system-wide
    AppendErrorLogCustom "EnumHostsFile - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumHostsFile"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumZones()
    RegEnumZones tvwMain
End Sub

Private Sub EnumIEToolbars()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumIEToolbars - Begin"
    
    Dim sKey$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "IEToolbars", SEC_IETOOLBARS, "msie"
    sKey = "Software\Microsoft\Internet Explorer\Toolbar"
    
    tvwMain.Nodes.Add "IEToolbars", tvwChild, "IEToolbarsSystem", "All users", "users"
    tvwMain.Nodes("IEToolbarsSystem").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    Dim sVals$(), i&, sCLSID$, sName$, sFile$
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sKey), "|")
    For i = 0 To UBound(sVals)
        If InStr(sVals(i), " (binary)") = Len(sVals(i)) - 8 Then
            sVals(i) = Left$(sVals(i), Len(sVals(i)) - 9)
        End If
        sCLSID = Mid$(sVals(i), 1, InStr(sVals(i), " = ") - 1)
        sName = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
        If sName = vbNullString Then
            sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
            If sName = vbNullString Then sName = "(no name)"
        End If
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
        If sFile = vbNullString Then sFile = "(no file)"
        If Not bShowCLSIDs Then
            tvwMain.Nodes.Add "IEToolbarsSystem", tvwChild, "IEToolbarsSystem" & i, sName & " - " & sFile, "dll"
        Else
            tvwMain.Nodes.Add "IEToolbarsSystem", tvwChild, "IEToolbarsSystem" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
        End If
        tvwMain.Nodes("IEToolbarsSystem" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("IEToolbarsSystem").Children > 0 Then
        tvwMain.Nodes("IEToolbarsSystem").Text = tvwMain.Nodes("IEToolbarsSystem").Text & " (" & tvwMain.Nodes("IEToolbarsSystem").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "IEToolbarsSystem"
    End If
    
    tvwMain.Nodes.Add "IEToolbars", tvwChild, "IEToolbarsUser", "This user", "user"
    tvwMain.Nodes.Add "IEToolbarsUser", tvwChild, "IEToolbarsUserShell", "ShellBrowser", "registry"
    tvwMain.Nodes("IEToolbarsUserShell").Tag = "HKEY_CURRENT_USER\" & sKey & "\ShellBrowser"
    sVals = Split(RegEnumValues(HKEY_CURRENT_USER, sKey & "\ShellBrowser", , False), "|")
    For i = 0 To UBound(sVals)
        If InStr(sVals(i), " (binary)") = Len(sVals(i)) - 8 Then
            sVals(i) = Left$(sVals(i), Len(sVals(i)) - 9)
        End If
        sCLSID = sVals(i)
        If sCLSID <> "ITBarLayout" Then
            sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
            If sName = vbNullString Then sName = "(no name)"
            sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
            If sFile = vbNullString Then sFile = "(no file)"
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "IEToolbarsUserShell", tvwChild, "IEToolbarsUserShell" & i, sName & " - " & sFile, "dll"
            Else
                tvwMain.Nodes.Add "IEToolbarsUserShell", tvwChild, "IEToolbarsUserShell" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
            End If
            tvwMain.Nodes("IEToolbarsUserShell" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("IEToolbarsUserShell").Children > 0 Then
        tvwMain.Nodes("IEToolbarsUserShell").Text = tvwMain.Nodes("IEToolbarsUserShell").Text & " (" & tvwMain.Nodes("IEToolbarsUserShell").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "IEToolbarsUserShell"
    End If
    
    tvwMain.Nodes.Add "IEToolbarsUser", tvwChild, "IEToolbarsUserWeb", "WebBrowser", "registry"
    tvwMain.Nodes("IEToolbarsUserWeb").Tag = "HKEY_CURRENT_USER\" & sKey & "\WebBrowser"
    sVals = Split(RegEnumValues(HKEY_CURRENT_USER, sKey & "\WebBrowser", , False), "|")
    For i = 0 To UBound(sVals)
        If InStr(sVals(i), " (binary)") = Len(sVals(i)) - 8 Then
            sVals(i) = Left$(sVals(i), Len(sVals(i)) - 9)
        End If
        sCLSID = sVals(i)
        If InStr(sCLSID, "{") = 1 Then
            sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
            If sName = vbNullString Then sName = "(no name)"
            sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
            If sFile = vbNullString Then sFile = "(no file)"
            If Not bShowCLSIDs Then
                tvwMain.Nodes.Add "IEToolbarsUserWeb", tvwChild, "IEToolbarsUserWeb" & i, sName & " - " & sFile, "dll"
            Else
                tvwMain.Nodes.Add "IEToolbarsUserWeb", tvwChild, "IEToolbarsUserWeb" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
            End If
            tvwMain.Nodes("IEToolbarsUserWeb" & i).Tag = GuessFullpathFromAutorun(sFile)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("IEToolbarsUserWeb").Children > 0 Then
        tvwMain.Nodes("IEToolbarsUserWeb").Text = tvwMain.Nodes("IEToolbarsUserWeb").Text & " (" & tvwMain.Nodes("IEToolbarsUserWeb").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "IEToolbarsUserWeb"
    End If
    If tvwMain.Nodes("IEToolbarsUser").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "IEToolbarsUser"
    End If
    
    If tvwMain.Nodes("IEToolbars").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "IEToolbars"
    End If

    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "IEToolbars", SEC_IETOOLBARS, "msie"
            tvwMain.Nodes.Add sUsernames(L) & "IEToolbars", tvwChild, sUsernames(L) & "IEToolbarsUserShell", "ShellBrowser", "registry"
            tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserShell").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKey & "\ShellBrowser"
            sVals = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\" & sKey & "\ShellBrowser", , False), "|")
            For i = 0 To UBound(sVals)
                If InStr(sVals(i), " (binary)") = Len(sVals(i)) - 8 Then
                    sVals(i) = Left$(sVals(i), Len(sVals(i)) - 9)
                End If
                sCLSID = sVals(i)
                If sCLSID <> "ITBarLayout" Then
                    sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
                    If sName = vbNullString Then sName = "(no name)"
                    sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
                    If sFile = vbNullString Then sFile = "(no file)"
                    If Not bShowCLSIDs Then
                        tvwMain.Nodes.Add sUsernames(L) & "IEToolbarsUserShell", tvwChild, sUsernames(L) & "IEToolbarsUserShell" & i, sName & " - " & sFile, "dll"
                    Else
                        tvwMain.Nodes.Add sUsernames(L) & "IEToolbarsUserShell", tvwChild, sUsernames(L) & "IEToolbarsUserShell" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
                    End If
                    tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserShell" & i).Tag = GuessFullpathFromAutorun(sFile)
                End If
                If bSL_Abort Then Exit Sub
            Next i
            If tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserShell").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserShell").Text = tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserShell").Text & " (" & tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserShell").Children & ")"
            Else
                If Not bShowEmpty Then tvwMain.Nodes.Remove sUsernames(L) & "IEToolbarsUserShell"
            End If
            
            tvwMain.Nodes.Add sUsernames(L) & "IEToolbars", tvwChild, sUsernames(L) & "IEToolbarsUserWeb", "WebBrowser", "registry"
            tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserWeb").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKey & "\WebBrowser"
            sVals = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\" & sKey & "\WebBrowser", , False), "|")
            For i = 0 To UBound(sVals)
                If InStr(sVals(i), " (binary)") = Len(sVals(i)) - 8 Then
                    sVals(i) = Left$(sVals(i), Len(sVals(i)) - 9)
                End If
                sCLSID = sVals(i)
                If InStr(sCLSID, "{") = 1 Then
                    sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
                    If sName = vbNullString Then sName = "(no name)"
                    sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
                    If sFile = vbNullString Then sFile = "(no file)"
                    If Not bShowCLSIDs Then
                        tvwMain.Nodes.Add sUsernames(L) & "IEToolbarsUserWeb", tvwChild, sUsernames(L) & "IEToolbarsUserWeb" & i, sName & " - " & sFile, "dll"
                    Else
                        tvwMain.Nodes.Add sUsernames(L) & "IEToolbarsUserWeb", tvwChild, sUsernames(L) & "IEToolbarsUserWeb" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
                    End If
                    tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserWeb" & i).Tag = GuessFullpathFromAutorun(sFile)
                End If
                If bSL_Abort Then Exit Sub
            Next i
            If tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserWeb").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserWeb").Text = tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserWeb").Text & " (" & tvwMain.Nodes(sUsernames(L) & "IEToolbarsUserWeb").Children & ")"
            Else
                If Not bShowEmpty Then tvwMain.Nodes.Remove sUsernames(L) & "IEToolbarsUserWeb"
            End If
            
            If tvwMain.Nodes(sUsernames(L) & "IEToolbars").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "IEToolbars"
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumIEToolbars - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumIEToolbars"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumIEExtensions()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumIEExtensions - Begin"
    
    Dim sKey$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "IEExtensions", SEC_IEEXTENSIONS, "msie"
    sKey = "Software\Microsoft\Internet Explorer\Extensions"
    tvwMain.Nodes("IEExtensions").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    Dim sKeys$(), i&, sCLSID$, sName$, sFile$
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey), "|")
    For i = 0 To UBound(sKeys)
        sName = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "ButtonText")
        If sName = vbNullString Then
            sName = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "MenuText")
        End If
        'get file from insane amount of possible locations
        'Exec > Script > BandCLSID > CLSIDExtension > CLSIDExtension\TreatAs > CLSID
        sCLSID = sKeys(i)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "Exec")
        If sFile = vbNullString Then
            sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "Script")
            If sFile = vbNullString Then
                'break out the clsid's
                sCLSID = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "BandCLSID")
                sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString)
                If sFile = vbNullString Then
                    sCLSID = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "CLSIDExtension")
                    sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString)
                    If sFile = vbNullString Then
                        sCLSID = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\TreatAs", vbNullString)
                        sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString)
                        If sFile = vbNullString Then
                            sCLSID = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "CLSID")
                            sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString)
                        End If
                    End If
                End If
            End If
        End If
        
        sFile = GetLongFilename(sFile)
        If Not bShowCLSIDs Then
            tvwMain.Nodes.Add "IEExtensions", tvwChild, "IEExtensions" & i, sName & " - " & sFile, "dll"
        Else
            tvwMain.Nodes.Add "IEExtensions", tvwChild, "IEExtensions" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
        End If
        tvwMain.Nodes("IEExtensions" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("IEExtensions").Children > 0 Then
        tvwMain.Nodes("IEExtensions").Text = tvwMain.Nodes("IEExtensions").Text & " (" & tvwMain.Nodes("IEExtensions").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "IEExtensions"
    End If
    
    '----------------------------------------------------------------
    'nothing - this is system-wide
    AppendErrorLogCustom "EnumIEExtensions - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumIEExtensions"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumIEExplBars()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumIEExplBars - Begin"
    Dim sKey$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "IEExplBars", SEC_IEBARS, "msie"
    sKey = "Software\Microsoft\Internet Explorer\Explorer Bars"
    tvwMain.Nodes("IEExplBars").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    Dim sKeys$(), i&, sName$, sFile$
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey), "|")
    For i = 0 To UBound(sKeys)
        sName = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i), vbNullString)
        If sName = vbNullString Then sName = "(no name)"
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sKeys(i) & "\InprocServer32", vbNullString))
        If sFile = vbNullString Then sFile = "(no file)"
        If Not bShowCLSIDs Then
            tvwMain.Nodes.Add "IEExplBars", tvwChild, "IEExplBars" & i, sName & " - " & sFile, "dll"
        Else
            tvwMain.Nodes.Add "IEExplBars", tvwChild, "IEExplBars" & i, sName & " - " & sKeys(i) & " - " & sFile, "dll"
        End If
        tvwMain.Nodes("IEExplBars" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("IEExplBars").Children > 0 Then
        tvwMain.Nodes("IEExplBars").Text = tvwMain.Nodes("IEExplBars").Text & " (" & tvwMain.Nodes("IEExplBars").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "IEExplBars"
    End If

    '----------------------------------------------------------------
    'nothing - this is system-wide
    AppendErrorLogCustom "EnumIEExplBars - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumIEExplBars"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumIEMenuExt()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumIEMenuExt - Begin"
    Dim sKey$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "IEMenuExt", SEC_IEMENUEXT, "msie"
    sKey = "Software\Microsoft\Internet Explorer\MenuExt"
    Dim sKeys$(), i&, sFile$
    sKeys = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sKey), "|")
    tvwMain.Nodes.Add "IEMenuExt", tvwChild, "IEMenuExtUser", "This user", "user"
    tvwMain.Nodes("IEMenuExtUser").Tag = "HKEY_CURRENT_USER\" & sKey
    For i = 0 To UBound(sKeys)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_CURRENT_USER, sKey & "\" & sKeys(i), vbNullString))
        sFile = GetLongFilename(sFile)
        tvwMain.Nodes.Add "IEMenuExtUser", tvwChild, "IEMenuExtUser" & i, sKeys(i) & " - " & sFile, "exe"
        tvwMain.Nodes("IEMenuExtUser" & i).Tag = "HKEY_CURRENT_USER\" & sKey & "\" & sKeys(i)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("IEMenuExtUser").Children > 0 Then
        tvwMain.Nodes("IEMenuExtUser").Text = tvwMain.Nodes("IEMenuExtUser").Text & " (" & tvwMain.Nodes("IEMenuExtUser").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "IEMenuExtUser"
    End If
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey), "|")
    tvwMain.Nodes.Add "IEMenuExt", tvwChild, "IEMenuExtSystem", "All users", "users"
    tvwMain.Nodes("IEMenuExtSystem").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    For i = 0 To UBound(sKeys)
        sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), vbNullString))
        sFile = GetLongFilename(sFile)
        tvwMain.Nodes.Add "IEMenuExtSystem", tvwChild, "IEMenuExtSystem" & i, sKeys(i) & " - " & sFile, "exe"
        tvwMain.Nodes("IEMenuExtSystem" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKey & "\" & sKeys(i)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("IEMenuExtSystem").Children > 0 Then
        tvwMain.Nodes("IEMenuExtSystem").Text = tvwMain.Nodes("IEMenuExtSystem").Text & " (" & tvwMain.Nodes("IEMenuExtSystem").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "IEMenuExtSystem"
    End If
    
    If tvwMain.Nodes("IEMenuExt").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "IEMenuExt"
    End If

    If Not bShowUsers Then Exit Sub
    '-----------------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "IEMenuExt", SEC_IEMENUEXT, "msie"
            tvwMain.Nodes(sUsernames(L) & "IEMenuExt").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKey
            
            sKeys = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\" & sKey), "|")
            For i = 0 To UBound(sKeys)
                sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sKey & "\" & sKeys(i), vbNullString))
                sFile = GetLongFilename(sFile)
                tvwMain.Nodes.Add sUsernames(L) & "IEMenuExt", tvwChild, sUsernames(L) & "IEMenuExtUser" & i, sKeys(i) & " - " & sFile, "exe"
                If bSL_Abort Then Exit Sub
            Next i
            If tvwMain.Nodes(sUsernames(L) & "IEMenuExt").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "IEMenuExt").Text = tvwMain.Nodes(sUsernames(L) & "IEMenuExt").Text & " (" & tvwMain.Nodes(sUsernames(L) & "IEMenuExt").Children & ")"
            Else
                If Not bShowEmpty Then
                    tvwMain.Nodes.Remove sUsernames(L) & "IEMenuExt"
                End If
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumIEMenuExt - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumIEMenuExt"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumIEBands()
    RegEnumIEBands tvwMain
End Sub

Private Sub EnumHijack()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumHijack - Begin"
    
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "Hijack", "Hijack points", "attn"
    
    Dim sUsername$, L&
    If bShowUsers Then
        For L = 0 To UBound(sUsernames)
            sUsername = MapSIDToUsername(sUsernames(L))
            If sUsername <> OSver.UserName And sUsername <> vbNullString Then
                tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "Hijack", "Hijack points", "attn"
            End If
        Next L
    End If
    If bShowHardware Then
        For L = 1 To UBound(sHardwareCfgs)
            tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "Hijack", "Hijack points", "attn"
        Next L
    End If
    
    'to list:
    '* IERESET.INF
    'Loading... Reset web settings URLs
    Status Translate(917)
    DoTicks tvwMain
    EnumResetWebSettings
    DoTicks tvwMain, "ResetWebSettings"
    UpdateProgressBar
    
    '* IE URLs
    'Loading... IE URLs
    Status Translate(918)
    DoTicks tvwMain
    EnumIEURLs
    DoTicks tvwMain, "IEURLs"
    UpdateProgressBar
    
    '* DefaultPrefix / Prefixes
    'Loading... Default URL prefixes
    Status Translate(919)
    DoTicks tvwMain
    EnumDefaultPrefix
    DoTicks tvwMain, "URLPrefix"
    UpdateProgressBar
    
'    '* Policy restrictions
'    Status "Loading... Policy restrictions"
'    DoTicks tvwMain
'    EnumPolicyRestrictions
'    DoTicks tvwMain, "PolicyRestrictions"
'    UpdateProgressBar

    '* hosts file databasepath
    'Loading... Hosts file path
    Status Translate(920)
    DoTicks tvwMain
    EnumHostsFilePath
    DoTicks tvwMain, "HostsFilePath"

    If bShowUsers Then
        For L = 0 To UBound(sUsernames)
            sUsername = MapSIDToUsername(sUsernames(L))
            If sUsername <> OSver.UserName And sUsername <> vbNullString Then
                If tvwMain.Nodes(sUsernames(L) & "Hijack").Children = 0 And Not bShowEmpty Then
                    tvwMain.Nodes.Remove sUsernames(L) & "Hijack"
                End If
            End If
        Next L
    End If
    If bShowHardware Then
        For L = 1 To UBound(sHardwareCfgs)
            If tvwMain.Nodes(sHardwareCfgs(L) & "Hijack").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sHardwareCfgs(L) & "Hijack"
            End If
        Next L
    End If
    AppendErrorLogCustom "EnumHijack - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumHijack"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumDefaultPrefix()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumDefaultPrefix - Begin"
    
    Dim sKey$, sPrefix$(), i%, sName$, sData$
    If bSL_Abort Then Exit Sub
    sKey = "Software\Microsoft\Windows\CurrentVersion\URL"
    tvwMain.Nodes.Add "Hijack", tvwChild, "URLPrefix", SEC_URLPREFIX, "msie"
    tvwMain.Nodes("URLPrefix").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    
    ReDim sPrefix(0)
    sPrefix(0) = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\DefaultPrefix", vbNullString)
    If sPrefix(0) <> vbNullString Or bShowEmpty Then
        tvwMain.Nodes.Add "URLPrefix", tvwChild, "URLDefaultPrefix", "default = " & sPrefix(0), "reg"
        tvwMain.Nodes("URLDefaultPrefix").Tag = "HKEY_LOCAL_MACHINE\" & sKey & "\DefaultPrefix"
    End If
    
    sPrefix = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sKey & "\Prefixes"), "|")
    For i = 0 To UBound(sPrefix)
        sName = Left$(sPrefix(i), InStr(sPrefix(i), " = ") - 1)
        sData = Mid$(sPrefix(i), InStr(sPrefix(i), " = ") + 3)
        tvwMain.Nodes.Add "URLPrefix", tvwChild, "URLPrefix" & i, sPrefix(i), "reg"
        tvwMain.Nodes("URLPrefix" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKey & "\Prefixes"
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("URLPrefix").Children > 0 Then
        tvwMain.Nodes("URLPrefix").Sorted = True
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "URLPrefix"
        End If
    End If
    '----------------------------------------------------------------
    'system-wide
    AppendErrorLogCustom "EnumDefaultPrefix - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumDefaultPrefix"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumResetWebSettings()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumResetWebSettings - Begin"
    'we need the settings for the following values from IERESET.INF:
    '* SearchAssistant
    '* CustomizeSearch
    '* START_PAGE_URL
    '* SEARCH_PAGE_URL
    '* MS_START_PAGE_URL
    Dim sContents$
    Dim sInf$(), i&, sSA$, sCS$, sSTPU$, sSEPU$, sMSPU$
    sSA = "SearchAssistant = "
    sCS = "CustomizeSearch = "
    sSTPU = "START_PAGE_URL = "
    sMSPU = "MS_START_PAGE_URL = "
    sSEPU = "SEARCH_PAGE_URL = "
    If bSL_Abort Then Exit Sub
    sContents = InputFile(BuildPath(sWinDir, "inf\iereset.inf"))
    'it's not always unicode format when in winNT!
    If InStr(sContents, Chr$(255)) = 1 Then
    'If bIsWinNT And Not bIsWinNT4 Then
        sInf = Split(StrConv(sContents, vbFromUnicode), vbCrLf)
    Else
        sInf = Split(sContents, vbCrLf)
    End If
    For i = 0 To UBound(sInf)
        If InStr(sInf(i), "SearchAssistant") > 0 Then
            sSA = Mid$(sInf(i), InStr(sInf(i), "http://"))
            sSA = "SearchAssistant = " & Left$(sSA, Len(sSA) - 1)
        End If
        If InStr(sInf(i), "CustomizeSearch") > 0 Then
            sCS = Mid$(sInf(i), InStr(sInf(i), "http://"))
            sCS = "CustomizeSearch = " & Left$(sCS, Len(sCS) - 1)
        End If
        If InStr(sInf(i), "START_PAGE_URL") = 1 And InStr(sInf(i), "MS_START_PAGE_URL") = 0 Then
            sSTPU = Mid$(sInf(i), InStr(sInf(i), "http://"))
            sSTPU = "START_PAGE_URL = " & Left$(sSTPU, Len(sSTPU) - 1)
        End If
        If InStr(sInf(i), "MS_START_PAGE_URL") = 1 Then
            sSEPU = Mid$(sInf(i), InStr(sInf(i), "http://"))
            sSEPU = "MS_START_PAGE_URL = " & Left$(sSEPU, Len(sSEPU) - 1)
        End If
        If InStr(sInf(i), "SEARCH_PAGE_URL") = 1 Then
            sMSPU = Mid$(sInf(i), InStr(sInf(i), "http://"))
            sMSPU = "SEARCH_PAGE_URL = " & Left$(sMSPU, Len(sMSPU) - 1)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    tvwMain.Nodes.Add "Hijack", tvwChild, "ResetWebSettings", SEC_RESETWEBSETTINGS, "ini"
    tvwMain.Nodes("ResetWebSettings").Tag = BuildPath(sWinDir, "inf\iereset.inf")
    If sSA <> vbNullString Or bShowEmpty Then tvwMain.Nodes.Add "ResetWebSettings", tvwChild, "ResetWebSettings0", sSA, "text"
    If sCS <> vbNullString Or bShowEmpty Then tvwMain.Nodes.Add "ResetWebSettings", tvwChild, "ResetWebSettings1", sCS, "text"
    If sSTPU <> vbNullString Or bShowEmpty Then tvwMain.Nodes.Add "ResetWebSettings", tvwChild, "ResetWebSettings2", sSTPU, "text"
    If sSEPU <> vbNullString Or bShowEmpty Then tvwMain.Nodes.Add "ResetWebSettings", tvwChild, "ResetWebSettings3", sSEPU, "text"
    If sMSPU <> vbNullString Or bShowEmpty Then tvwMain.Nodes.Add "ResetWebSettings", tvwChild, "ResetWebSettings4", sMSPU, "text"
    If tvwMain.Nodes("ResetWebSettings").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ResetWebSettings"
    End If
    '----------------------------------------------------------------
    'system-wide
    AppendErrorLogCustom "EnumResetWebSettings - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumResetWebSettings"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumHostsFilePath()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumHostsFilePath - Begin"
    
    Dim sDatabasePath$, sKey$
    If bSL_Abort Then Exit Sub
    sKey = "System\CurrentControlSet\Services\Tcpip\Parameters"
    tvwMain.Nodes.Add "Hijack", tvwChild, "HostsFilePath", SEC_HOSTSFILEPATH, "registry"
    tvwMain.Nodes("HostsFilePath").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    sDatabasePath = ExpandEnvironmentVars(Reg.GetString(HKEY_LOCAL_MACHINE, sKey, "DatabasePath"))
    If sDatabasePath <> vbNullString Or bShowEmpty Then
        tvwMain.Nodes.Add "HostsFilePath", tvwChild, "HostsFilePath0", "DatabasePath = " & BuildPath(sDatabasePath, "hosts"), "text"
        tvwMain.Nodes("HostsFilePath0").Tag = BuildPath(sDatabasePath, "hosts")
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "HostsFilePath"
        End If
    End If
    
    If Not bShowHardware Then Exit Sub
    '----------------------------------------------------------------
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sKey = "System\" & sHardwareCfgs(L) & "\Services\Tcpip\Parameters"

        tvwMain.Nodes.Add sHardwareCfgs(L) & "Hijack", tvwChild, sHardwareCfgs(L) & "HostsFilePath", SEC_HOSTSFILEPATH, "registry"
        tvwMain.Nodes(sHardwareCfgs(L) & "HostsFilePath").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    
        sDatabasePath = ExpandEnvironmentVars(Reg.GetString(HKEY_LOCAL_MACHINE, sKey, "DatabasePath"))
        If sDatabasePath <> vbNullString Or bShowEmpty Then
            tvwMain.Nodes.Add sHardwareCfgs(L) & "HostsFilePath", tvwChild, sHardwareCfgs(L) & "HostsFilePath0", "DatabasePath = " & BuildPath(sDatabasePath, "hosts"), "text"
            tvwMain.Nodes(sHardwareCfgs(L) & "HostsFilePath0").Tag = BuildPath(sDatabasePath, "hosts")
        Else
            If Not bShowEmpty Then
                tvwMain.Nodes.Remove sHardwareCfgs(L) & "HostsFilePath"
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumHostsFilePath - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumHostsFilePath"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumIEURLs()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumIEURLs - Begin"
    
    Dim sKeys$(), sVals$(), i&, j&, sVal$
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "Hijack", tvwChild, "IEURLs", SEC_IEURLS, "msie"
    tvwMain.Nodes.Add "IEURLs", tvwChild, "IEURLsUser", "This user", "user"
    tvwMain.Nodes.Add "IEURLs", tvwChild, "IEURLsSystem", "All users", "users"
    ReDim sKeys(6)
    sKeys(0) = "Software\Microsoft\Internet Explorer"
    sKeys(1) = "Software\Microsoft\Internet Explorer\Main"
    sKeys(2) = "Software\Microsoft\Internet Explorer\Search"
    sKeys(3) = "Software\Microsoft\Internet Explorer\SearchURL"
    sKeys(4) = "Software\Microsoft\Internet Explorer\Desktop\General"
    sKeys(5) = "Software\Microsoft\Internet Explorer\SafeMode\Desktop"
    sKeys(6) = "Software\Microsoft\Internet Explorer\AboutURLs"
    
    ReDim sVals(24)
    sVals(0) = vbNullString
    sVals(1) = "Default_Page_Url"
    sVals(2) = "Default_Search_Url"
    sVals(3) = "SearchAssistant"
    sVals(4) = "CustomizeSearch"
    sVals(5) = "Search"
    sVals(6) = "Search Bar"
    sVals(7) = "Search Page"
    sVals(8) = "Start Page"
    sVals(9) = "SearchURL"
    sVals(10) = "www"
    sVals(11) = "Startpagina"
    sVals(12) = "First Home Page"
    sVals(13) = "Local Page"
    sVals(14) = "Start Page_bak"
    sVals(15) = "HomeOldSP"
    sVals(16) = "Window Title"
    sVals(17) = "Wallpaper"
    sVals(18) = "BackupWallpaper"
    sVals(19) = "blank"
    sVals(20) = "DesktopItemNavigationFailure"
    sVals(21) = "NavigationCanceled"
    sVals(22) = "NavigationFailure"
    sVals(23) = "OfflineInformation"
    sVals(24) = "PostNotCached"
    
    For i = 0 To UBound(sKeys)
        tvwMain.Nodes.Add "IEURLsSystem", tvwChild, "IEURLsSystem" & i, Mid$(sKeys(i), InStr(sKeys(i), "Internet")), "registry"
        tvwMain.Nodes("IEURLsSystem" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKeys(i)
        For j = 0 To UBound(sVals)
            sVal = Reg.GetString(HKEY_LOCAL_MACHINE, sKeys(i), sVals(j))
            If sVal <> vbNullString Or bShowEmpty Then
                If sVals(j) = vbNullString Then
                    tvwMain.Nodes.Add "IEURLsSystem" & i, tvwChild, "IEURLsSystem" & i & "sub" & j, "(Default) = " & sVal, "reg"
                Else
                    tvwMain.Nodes.Add "IEURLsSystem" & i, tvwChild, "IEURLsSystem" & i & "sub" & j, sVals(j) & " = " & sVal, "reg"
                End If
            End If
        Next j
        If tvwMain.Nodes("IEURLsSystem" & i).Children > 0 Then
            tvwMain.Nodes("IEURLsSystem" & i).Text = tvwMain.Nodes("IEURLsSystem" & i).Text & " (" & tvwMain.Nodes("IEURLsSystem" & i).Children & ")"
            tvwMain.Nodes("IEURLsSystem" & i).Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove "IEURLsSystem" & i
        End If
        If bSL_Abort Then Exit Sub
    Next i

    For i = 0 To UBound(sKeys)
        tvwMain.Nodes.Add "IEURLsUser", tvwChild, "IEURLsUser" & i, Mid$(sKeys(i), InStr(sKeys(i), "Internet")), "registry"
        tvwMain.Nodes("IEURLsUser" & i).Tag = "HKEY_CURRENT_USER\" & sKeys(i)
        For j = 0 To UBound(sVals)
            sVal = ExpandEnvironmentVars(Reg.GetString(HKEY_CURRENT_USER, sKeys(i), sVals(j)))
            If sVal <> vbNullString Or bShowEmpty Then
                If sVals(j) = vbNullString Then
                    tvwMain.Nodes.Add "IEURLsUser" & i, tvwChild, "IEURLsUser" & i & "sub" & j, "(Default) = " & sVal, "reg"
                Else
                    tvwMain.Nodes.Add "IEURLsUser" & i, tvwChild, "IEURLsUser" & i & "sub" & j, sVals(j) & " = " & sVal, "reg"
                End If
            End If
        Next j
        If tvwMain.Nodes("IEURLsUser" & i).Children > 0 Then
            tvwMain.Nodes("IEURLsUser" & i).Text = tvwMain.Nodes("IEURLsUser" & i).Text & " (" & tvwMain.Nodes("IEURLsUser" & i).Children & ")"
            tvwMain.Nodes("IEURLsUser" & i).Sorted = True
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove "IEURLsUser" & i
        End If
        If bSL_Abort Then Exit Sub
    Next i

    If Not bShowUsers Then Exit Sub
    '-----------------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add sUsernames(L) & "Hijack", tvwChild, sUsernames(L) & "IEURLsUser", SEC_IEURLS, "msie"
            
            For i = 0 To UBound(sKeys)
                tvwMain.Nodes.Add sUsernames(L) & "IEURLsUser", tvwChild, sUsernames(L) & "IEURLsUser" & i, Mid$(sKeys(i), InStr(sKeys(i), "Internet")), "registry"
                tvwMain.Nodes(sUsernames(L) & "IEURLsUser" & i).Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sKeys(i)
                For j = 0 To UBound(sVals)
                    sVal = ExpandEnvironmentVars(Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sKeys(i), sVals(j)))
                    If sVal <> vbNullString Or bShowEmpty Then
                        If sVals(j) = vbNullString Then
                            tvwMain.Nodes.Add sUsernames(L) & "IEURLsUser" & i, tvwChild, sUsernames(L) & "IEURLsUser" & i & "sub" & j, "(Default) = " & sVal, "reg"
                        Else
                            tvwMain.Nodes.Add sUsernames(L) & "IEURLsUser" & i, tvwChild, sUsernames(L) & "IEURLsUser" & i & "sub" & j, sVals(j) & " = " & sVal, "reg"
                        End If
                    End If
                Next j
                If tvwMain.Nodes(sUsernames(L) & "IEURLsUser" & i).Children > 0 Then
                    tvwMain.Nodes(sUsernames(L) & "IEURLsUser" & i).Text = tvwMain.Nodes(sUsernames(L) & "IEURLsUser" & i).Text & " (" & tvwMain.Nodes(sUsernames(L) & "IEURLsUser" & i).Children & ")"
                    tvwMain.Nodes(sUsernames(L) & "IEURLsUser" & i).Sorted = True
                Else
                    If Not bShowEmpty Then tvwMain.Nodes.Remove sUsernames(L) & "IEURLsUser" & i
                End If
                If bSL_Abort Then Exit Sub
            Next i
            
            If tvwMain.Nodes(sUsernames(L) & "IEURLsUser").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "IEURLsUser"
            End If
        End If
    Next L
    AppendErrorLogCustom "EnumIEURLs - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumIEURLs"
    If inIDE Then Stop: Resume Next
End Sub

'Private Sub EnumPolicyRestrictions()
'    Dim sKeys$(), sNames$(), sVals$(), i&, j&
'    tvwMain.Nodes.Add "Hijack", tvwChild, "PolicyRestrictions", "Policy restrictions", "registry"
'    tvwMain.Nodes.Add "PolicyRestrictions", tvwChild, "PolicyRestrictionsSystem", "All users", "users"
'    tvwMain.Nodes.Add "PolicyRestrictions", tvwChild, "PolicyRestrictionsUser", "This user", "user"
'    ReDim sKeys(10)
'    sKeys(0) = "Software\Microsoft\Windows\CurrentVersion\Policies"
'    sKeys(1) = "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop"
'    sKeys(2) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
'    sKeys(3) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
'    sKeys(4) = "Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate"
'    sKeys(5) = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
'    sKeys(6) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
'    sKeys(7) = "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions"
'    sKeys(8) = "Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions"
'    sKeys(9) = "Software\Policies\Microsoft\Conferencing"
'    sKeys(10) = "Software\Policies\Microsoft\Windows NT\SystemRestore"
'    ReDim sNames(10)
'    sNames(0) = "General policies"
'    sNames(1) = "Active Desktop"
'    sNames(2) = "Explorer"
'    sNames(3) = "System"
'    sNames(4) = "WindowsUpdate"
'    sNames(5) = "Internet Explorer"
'    sNames(6) = "Internet Explorer Control Panel applet"
'    sNames(7) = "Internet Explorer toolbars"
'    sNames(8) = "Internet Explorer synchronize"
'    sNames(9) = "Microsoft Netmeeting"
'    sNames(10) = "System Restore"
'
'    For i = 0 To UBound(sKeys)
'        tvwMain.Nodes.Add "PolicyRestrictionsSystem", tvwChild, "PolicyRestrictionsSystem" & i, sNames(i), "registry"
'        tvwMain.Nodes("PolicyRestrictionsSystem" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKeys(i)
'        sVals = Split(RegEnumDwordValues(HKEY_LOCAL_MACHINE, sKeys(i)), "|")
'        For j = 0 To UBound(sVals)
'            tvwMain.Nodes.Add "PolicyRestrictionsSystem" & i, tvwChild, "PolicyRestrictionsSystem" & i & "sub" & j, sVals(j), "reg"
'        Next j
'        If tvwMain.Nodes("PolicyRestrictionsSystem" & i).Children<> 0Then
'            tvwMain.Nodes("PolicyRestrictionsSystem" & i).Text = tvwMain.Nodes("PolicyRestrictionsSystem" & i).Text & " (" & tvwMain.Nodes("PolicyRestrictionsSystem" & i).Children & ")"
'        Else
'            If Not bShowEmpty Then
'                tvwMain.Nodes.Remove "PolicyRestrictionsSystem" & i
'            End If
'        End If
'    Next i
'    For i = 0 To UBound(sKeys)
'        tvwMain.Nodes.Add "PolicyRestrictionsUser", tvwChild, "PolicyRestrictionsUser" & i, sNames(i), "registry"
'        tvwMain.Nodes("PolicyRestrictionsUser" & i).Tag = "HKEY_CURRENT_USER\" & sKeys(i)
'        sVals = Split(RegEnumDwordValues(HKEY_CURRENT_USER, sKeys(i)), "|")
'        For j = 0 To UBound(sVals)
'            tvwMain.Nodes.Add "PolicyRestrictionsUser" & i, tvwChild, "PolicyRestrictionsUser" & i & "sub" & j, sVals(j), "reg"
'        Next j
'        If tvwMain.Nodes("PolicyRestrictionsUser" & i).Children<> 0Then
'            tvwMain.Nodes("PolicyRestrictionsUser" & i).Text = tvwMain.Nodes("PolicyRestrictionsUser" & i).Text & " (" & tvwMain.Nodes("PolicyRestrictionsUser" & i).Children & ")"
'        Else
'            If Not bShowEmpty Then
'                tvwMain.Nodes.Remove "PolicyRestrictionsUser" & i
'            End If
'        End If
'    Next i
'
'    If Not bShowUsers Then Exit Sub
'    '-----------------------------------------------------------------------
'    Dim sUsername$, l&
'    For l = 0 To UBound(sUsernames)
'        sUsername = MapSIDToUsername(sUsernames(l))
'        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
'            tvwMain.Nodes.Add sUsernames(l) & "Hijack", tvwChild, sUsernames(l) & "PolicyRestrictionsUser", "Policy restrictions", "msie"
'
'            For i = 0 To UBound(sKeys)
'                tvwMain.Nodes.Add sUsernames(l) & "PolicyRestrictionsUser", tvwChild, sUsernames(l) & "PolicyRestrictionsUser" & i, sNames(i), "registry"
'                tvwMain.Nodes(sUsernames(l) & "PolicyRestrictionsUser" & i).Tag = "HKEY_USERS\" & sUsernames(l) & "\" & sKeys(i)
'                sVals = Split(RegEnumDwordValues(HKEY_USERS, sUsernames(l) & "\" & sKeys(i)), "|")
'                For j = 0 To UBound(sVals)
'                    tvwMain.Nodes.Add sUsernames(l) & "PolicyRestrictionsUser" & i, tvwChild, sUsernames(l) & "PolicyRestrictionsUser" & i & "sub" & j, sVals(j), "reg"
'                Next j
'                If tvwMain.Nodes(sUsernames(l) & "PolicyRestrictionsUser" & i).Children<> 0Then
'                    tvwMain.Nodes(sUsernames(l) & "PolicyRestrictionsUser" & i).Text = tvwMain.Nodes(sUsernames(l) & "PolicyRestrictionsUser" & i).Text & " (" & tvwMain.Nodes(sUsernames(l) & "PolicyRestrictionsUser" & i).Children & ")"
'                Else
'                    If Not bShowEmpty Then
'                        tvwMain.Nodes.Remove sUsernames(l) & "PolicyRestrictionsUser" & i
'                    End If
'                End If
'            Next i
'
'            If tvwMain.Nodes(sUsernames(l) & "PolicyRestrictionsUser").Children = 0 And Not bShowEmpty Then
'                tvwMain.Nodes.Remove sUsernames(l) & "PolicyRestrictionsUser"
'            End If
'        End If
'    Next l
'End Sub

Private Sub EnumDriverFilters()
    RegEnumDriverFilters tvwMain
End Sub

Private Sub EnumStoppedServices()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumStoppedServices - Begin"
    
    Dim sKey$, sKeys$(), i&, sDisplayName$, lStart&, lType&, sFile$
    If bSL_Abort Then Exit Sub
    sKey = "System\CurrentControlSet\Services"
    tvwMain.Nodes.Add "DisabledEnums", tvwChild, "StoppedServices", SEC_STOPPEDSERVICES, "exe"
    tvwMain.Nodes.Add "StoppedServices", tvwChild, "StoppedOnlyServices", "Stopped", "exe"
    tvwMain.Nodes.Add "StoppedServices", tvwChild, "DisabledServices", "Stopped & disabled", "exe"
    tvwMain.Nodes("StoppedServices").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    tvwMain.Nodes("StoppedOnlyServices").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    tvwMain.Nodes("DisabledServices").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sKey), "|")
    For i = 0 To UBound(sKeys)
        sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "DisplayName")
        If sDisplayName = vbNullString Then
            sDisplayName = sKeys(i)
        End If
        lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "Start")
        lType = Reg.GetDword(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "Type")
        'sFile = ExpandEnvironmentVars(Reg.GetString(HKEY_LOCAL_MACHINE, sKey & "\" & sKeys(i), "ImagePath"))
        sFile = GetServiceImagePath(sKeys(i))
        If (lStart = 3 Or lStart = 4) And sDisplayName <> vbNullString And sFile <> vbNullString And lType >= 16 Then
            If lStart = 4 Then
                tvwMain.Nodes.Add "DisabledServices", tvwChild, "StoppedServices" & i, sDisplayName & " = " & sFile, "exe", "exe"
                tvwMain.Nodes("StoppedServices" & i).Tag = GuessFullpathFromAutorun(sFile)
            Else
                tvwMain.Nodes.Add "StoppedOnlyServices", tvwChild, "StoppedServices" & i, sDisplayName & " = " & sFile, "exe", "exe"
                tvwMain.Nodes("StoppedServices" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("StoppedOnlyServices").Children > 0 Then
        tvwMain.Nodes("StoppedOnlyServices").Text = tvwMain.Nodes("StoppedOnlyServices").Text & " (" & tvwMain.Nodes("StoppedOnlyServices").Children & ")"
        tvwMain.Nodes("StoppedOnlyServices").Sorted = True
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "StoppedOnlyServices"
    End If
    
    If tvwMain.Nodes("DisabledServices").Children > 0 Then
        tvwMain.Nodes("DisabledServices").Text = tvwMain.Nodes("DisabledServices").Text & " (" & tvwMain.Nodes("DisabledServices").Children & ")"
        tvwMain.Nodes("DisabledServices").Sorted = True
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "DisabledServices"
        End If
    End If
    
    If tvwMain.Nodes("StoppedServices").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "StoppedServices"
    End If
    
    AppendErrorLogCustom "EnumStoppedServices - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumStoppedServices"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumPolicies()
    RegEnumPolicies tvwMain
End Sub

Private Sub EnumXPSecurity()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumXPSecurity - Begin"
    
'  SOFTWARE\Microsoft\Security Center
'  Software\Microsoft\Windows NT\CurrentVersion\systemrestore
'  HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\AuthorizedApplications\List
'  HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List

'  Software\Microsoft\Windows Defender,DisableAntiSpyware !

    Dim sVals$(), i&
    If bSL_Abort Then Exit Sub
    tvwMain.Nodes.Add "DisabledEnums", tvwChild, "XPSecurity", SEC_XPSECURITY, "internet"
    
    tvwMain.Nodes.Add "XPSecurity", tvwChild, "XPSecurityCenter", "Security Center", "xpsec"
    tvwMain.Nodes.Add "XPSecurity", tvwChild, "XPFirewall", "Windows Firewall exceptions", "xpsec"
    tvwMain.Nodes.Add "XPSecurity", tvwChild, "XPSecurityRestore", "System Restore", "drive"
    tvwMain.Nodes.Add "XPSecurity", tvwChild, "WindowsDefender", "Windows Defender", "defender"
    
    sVals = Split(RegEnumValues(HKEY_CURRENT_USER, "Software\Microsoft\Security Center", , , False), "|")
    If UBound(sVals) > -1 Or bShowEmpty Then
        tvwMain.Nodes.Add "XPSecurityCenter", tvwChild, "XPSecurityCenterUser", "This user", "user"
        tvwMain.Nodes("XPSecurityCenterUser").Tag = "HKEY_CURRENT_USER\Software\Microsoft\Security Center"
        For i = 0 To UBound(sVals)
            tvwMain.Nodes.Add "XPSecurityCenterUser", tvwChild, "XPSecurityCenterUser" & i, sVals(i), "reg"
            If bSL_Abort Then Exit Sub
        Next i
    End If
    
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, "Software\Microsoft\Security Center", , , False), "|")
    If UBound(sVals) > -1 Or bShowEmpty Then
        tvwMain.Nodes.Add "XPSecurityCenter", tvwChild, "XPSecurityCenterSystem", "All users", "users"
        tvwMain.Nodes("XPSecurityCenterSystem").Tag = "HKEY_LOCAL_MACHINE\Software\Microsoft\Security Center"
        For i = 0 To UBound(sVals)
            tvwMain.Nodes.Add "XPSecurityCenterSystem", tvwChild, "XPSecurityCenterSystem" & i, sVals(i), "reg"
            If bSL_Abort Then Exit Sub
        Next i
    End If
    
    '------------------------------------
    
    Dim sFirewallKeyD$, sFirewallKeyS$
    Dim sFile$, sPort$, sProtocol$, sScope$, bEnabled As Boolean, sName$
    sFirewallKeyD = "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile"
    sFirewallKeyS = "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile"
    
    tvwMain.Nodes.Add "XPFirewall", tvwChild, "XPFirewallDomain", "Network domain mode", "lsp"
    tvwMain.Nodes.Add "XPFirewall", tvwChild, "XPFirewallStandard", "Standalone mode", "system"
    tvwMain.Nodes("XPFirewallDomain").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyD
    tvwMain.Nodes("XPFirewallStandard").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyS
    
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sFirewallKeyD & "\AuthorizedApplications\List"), "|")
    If UBound(sVals) > -1 Or bShowEmpty Then
        tvwMain.Nodes.Add "XPFirewallDomain", tvwChild, "XPFirewallDomainApps", "Applications", "exe"
        tvwMain.Nodes("XPFirewallDomainApps").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyD & "\AuthorizedApplications\List"
        For i = 0 To UBound(sVals)
            sVals(i) = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
            sFile = Left$(sVals(i), InStr(3, sVals(i), ":") - 1)
            sFile = ExpandEnvironmentVars(sFile)
            sScope = Mid$(sVals(i), InStr(3, sVals(i), ":") + 1)
            bEnabled = IIf(InStr(1, sScope, ":Enabled:", vbTextCompare) > 0, True, False)
            sName = Mid$(sScope, InStr(sScope, ":") + 1)
            sName = Mid$(sName, InStr(sName, ":") + 1)
            If InStr(sName, "@") = 1 Then
                sName = Mid$(sName, 2)
                sName = GetStringResFromDLL(sSysDir & "\" & Left$(sName, InStr(sName, ",") - 1), Mid$(sName, InStr(sName, ",") + 1))
            End If
            sScope = Left$(sScope, InStr(sScope, ":") - 1)
            If sScope = "*" Then sScope = "any computer"
            
            tvwMain.Nodes.Add "XPFirewallDomainApps", tvwChild, "XPFirewallDomainApps" & i, sName & " - " & sScope & " (" & IIf(bEnabled, "enabled)", "disabled)"), "firewall"
            tvwMain.Nodes("XPFirewallDomainApps" & i).Tag = sFile
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes("XPFirewallDomainApps").Children > 0 Then
            tvwMain.Nodes("XPFirewallDomainApps").Text = tvwMain.Nodes("XPFirewallDomainApps").Text & " (" & tvwMain.Nodes("XPFirewallDomainApps").Children & ")"
        End If
    End If
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sFirewallKeyD & "\GloballyOpenPorts\List"), "|")
    If UBound(sVals) > -1 Or bShowEmpty Then
        tvwMain.Nodes.Add "XPFirewallDomain", tvwChild, "XPFirewallDomainPorts", "Ports", "internet"
        tvwMain.Nodes("XPFirewallDomainPorts").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyD & "\GloballyOpenPorts\List"
        For i = 0 To UBound(sVals)
            sVals(i) = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
            sPort = Left$(sVals(i), InStr(3, sVals(i), ":") - 1)
            sProtocol = Mid$(sVals(i), InStr(3, sVals(i), ":") + 1)
            sScope = Mid$(sProtocol, InStr(sProtocol, ":") + 1)
            bEnabled = IIf(InStr(1, sScope, ":Enabled:", vbTextCompare) > 0, True, False)
            sName = Mid$(sScope, InStr(sScope, ":") + 1)
            sName = Mid$(sName, InStr(sName, ":") + 1)
            If InStr(sName, "@") = 1 Then
                sName = Mid$(sName, 2)
                sName = GetStringResFromDLL(sSysDir & "\" & Left$(sName, InStr(sName, ",") - 1), Mid$(sName, InStr(sName, ",") + 1))
            End If
            sProtocol = Left$(sProtocol, InStr(sProtocol, ":") - 1)
            sScope = Left$(sScope, InStr(sScope, ":") - 1)
            If sScope = "*" Then sScope = "any computer"
        
            tvwMain.Nodes.Add "XPFirewallDomainPorts", tvwChild, "XPFirewallDomainPorts" & i, sName & " - " & sProtocol & " port " & sPort & " - " & sScope & " (" & IIf(bEnabled, "enabled)", "disabled)"), "firewall"
            tvwMain.Nodes("XPFirewallDomainPorts" & i).Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyD & "\GloballyOpenPorts\List"
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes("XPFirewallDomainPorts").Children > 0 Then
            tvwMain.Nodes("XPFirewallDomainPorts").Text = tvwMain.Nodes("XPFirewallDomainPorts").Text & " (" & tvwMain.Nodes("XPFirewallDomainPorts").Children & ")"
        End If
    End If
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sFirewallKeyS & "\AuthorizedApplications\List"), "|")
    If UBound(sVals) > -1 Or bShowEmpty Then
        tvwMain.Nodes.Add "XPFirewallStandard", tvwChild, "XPFirewallStandardApps", "Applications", "exe"
        tvwMain.Nodes("XPFirewallStandardApps").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyS & "\AuthorizedApplications\List"
        For i = 0 To UBound(sVals)
            sVals(i) = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
            sFile = Left$(sVals(i), InStr(3, sVals(i), ":") - 1)
            sFile = ExpandEnvironmentVars(sFile)
            sScope = Mid$(sVals(i), InStr(3, sVals(i), ":") + 1)
            bEnabled = IIf(InStr(1, sScope, ":Enabled:", vbTextCompare) > 0, True, False)
            sName = Mid$(sScope, InStr(sScope, ":") + 1)
            sName = Mid$(sName, InStr(sName, ":") + 1)
            If InStr(sName, "@") = 1 Then
                sName = Mid$(sName, 2)
                sName = GetStringResFromDLL(sSysDir & "\" & Left$(sName, InStr(sName, ",") - 1), Mid$(sName, InStr(sName, ",") + 1))
            End If
            sScope = Left$(sScope, InStr(sScope, ":") - 1)
            If sScope = "*" Then sScope = "any computer"
                        
            tvwMain.Nodes.Add "XPFirewallStandardApps", tvwChild, "XPFirewallStandardApps" & i, sName & " - " & sScope & " (" & IIf(bEnabled, "enabled)", "disabled)"), "firewall"
            tvwMain.Nodes("XPFirewallStandardApps" & i).Tag = sFile
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes("XPFirewallStandardApps").Children > 0 Then
            tvwMain.Nodes("XPFirewallStandardApps").Text = tvwMain.Nodes("XPFirewallStandardApps").Text & " (" & tvwMain.Nodes("XPFirewallStandardApps").Children & ")"
        End If
    End If
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sFirewallKeyS & "\GloballyOpenPorts\List"), "|")
    If UBound(sVals) > -1 Or bShowEmpty Then
        tvwMain.Nodes.Add "XPFirewallStandard", tvwChild, "XPFirewallStandardPorts", "Ports", "internet"
        tvwMain.Nodes("XPFirewallStandardPorts").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyS & "\GloballyOpenPorts\List"
        For i = 0 To UBound(sVals)
            sVals(i) = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
            sPort = Left$(sVals(i), InStr(3, sVals(i), ":") - 1)
            sProtocol = Mid$(sVals(i), InStr(3, sVals(i), ":") + 1)
            sScope = Mid$(sProtocol, InStr(sProtocol, ":") + 1)
            bEnabled = IIf(InStr(1, sScope, ":Enabled:", vbTextCompare) > 0, True, False)
            sName = Mid$(sScope, InStr(sScope, ":") + 1)
            sName = Mid$(sName, InStr(sName, ":") + 1)
            If InStr(sName, "@") = 1 Then
                sName = Mid$(sName, 2)
                sName = GetStringResFromDLL(sSysDir & "\" & Left$(sName, InStr(sName, ",") - 1), Mid$(sName, InStr(sName, ",") + 1))
            End If
            sProtocol = Left$(sProtocol, InStr(sProtocol, ":") - 1)
            sScope = Left$(sScope, InStr(sScope, ":") - 1)
            If sScope = "*" Then sScope = "any computer"
        
            tvwMain.Nodes.Add "XPFirewallStandardPorts", tvwChild, "XPFirewallStandardPorts" & i, sName & " - " & sProtocol & " port " & sPort & " - " & sScope & " (" & IIf(bEnabled, "enabled)", "disabled)"), "firewall"
            tvwMain.Nodes("XPFirewallStandardPorts" & i).Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyS & "\GloballyOpenPorts\List"
            If bSL_Abort Then Exit Sub
        Next i
        If tvwMain.Nodes("XPFirewallStandardPorts").Children > 0 Then
            tvwMain.Nodes("XPFirewallStandardPorts").Text = tvwMain.Nodes("XPFirewallStandardPorts").Text & " (" & tvwMain.Nodes("XPFirewallStandardPorts").Children & ")"
        End If
    End If
    '------------------------------------
    
    sVals = Split(RegEnumValues(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\systemrestore", , , False), "|")
    If UBound(sVals) > -1 Or bShowEmpty Then
        tvwMain.Nodes.Add "XPSecurityRestore", tvwChild, "XPSecurityRestoreUser", "This user", "user"
        tvwMain.Nodes("XPSecurityRestoreUser").Tag = "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\SystemRestore"
        For i = 0 To UBound(sVals)
            tvwMain.Nodes.Add "XPSecurityRestoreUser", tvwChild, "XPSecurityRestoreUser" & i, sVals(i), "reg"
            If bSL_Abort Then Exit Sub
        Next i
    End If
    
    sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\systemrestore", , , False), "|")
    If UBound(sVals) > -1 Or bShowEmpty Then
        tvwMain.Nodes.Add "XPSecurityRestore", tvwChild, "XPSecurityRestoreSystem", "All users", "users"
        tvwMain.Nodes("XPSecurityRestoreSystem").Tag = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\SystemRestore"
        For i = 0 To UBound(sVals)
            tvwMain.Nodes.Add "XPSecurityRestoreSystem", tvwChild, "XPSecurityRestoreSystem" & i, sVals(i), "reg"
            If bSL_Abort Then Exit Sub
        Next i
    End If
    
    If tvwMain.Nodes("XPSecurityCenter").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "XPSecurityCenter"
    End If
    If tvwMain.Nodes("XPSecurityRestore").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "XPSecurityRestore"
    End If
    
    If tvwMain.Nodes("XPSecurity").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "XPSecurity"
    End If

    If bShowUsers Then
    '-----------------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add sUsernames(L) & "DisabledEnums", tvwChild, sUsernames(L) & "XPSecurity", SEC_XPSECURITY, "internet"
            
            tvwMain.Nodes.Add sUsernames(L) & "XPSecurity", tvwChild, sUsernames(L) & "XPSecurityCenter", "Security Center", "xpsec"
            tvwMain.Nodes.Add sUsernames(L) & "XPSecurity", tvwChild, sUsernames(L) & "XPSecurityRestore", "System Restore", "drive"
            tvwMain.Nodes(sUsernames(L) & "XPSecurityCenter").Tag = "HKEY_USERS\" & sUsernames(L) & "\Software\Microsoft\Security Center"
            tvwMain.Nodes(sUsernames(L) & "XPSecurityRestore").Tag = "HKEY_USERS\" & sUsernames(L) & "\Software\Microsoft\Windows NT\CurrentVersion\SystemRestore"
                        
            sVals = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\Software\Microsoft\Security Center", , , False), "|")
            If UBound(sVals) > -1 Or bShowEmpty Then
                For i = 0 To UBound(sVals)
                    tvwMain.Nodes.Add sUsernames(L) & "XPSecurityCenter", tvwChild, sUsernames(L) & "XPSecurityCenter" & i, sVals(i), "reg"
                Next i
            End If
            sVals = Split(RegEnumValues(HKEY_USERS, sUsernames(L) & "\Software\Microsoft\Windows NT\CurrentVersion\systemrestore", , , False), "|")
            If UBound(sVals) > -1 Or bShowEmpty Then
                For i = 0 To UBound(sVals)
                    tvwMain.Nodes.Add sUsernames(L) & "XPSecurityRestore", tvwChild, sUsernames(L) & "XPSecurityRestore" & i, sVals(i), "reg"
                Next i
            End If
            
            If tvwMain.Nodes(sUsernames(L) & "XPSecurityCenter").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "XPSecurityCenter"
            End If
            If tvwMain.Nodes(sUsernames(L) & "XPSecurityRestore").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "XPSecurityRestore"
            End If
            
            If tvwMain.Nodes(sUsernames(L) & "XPSecurity").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(L) & "XPSecurity"
            End If
        End If
        If bSL_Abort Then Exit Sub
    Next L
    
    End If
    
    If Not bShowHardware Then Exit Sub
    '-----------------------------------------------------------------------
    For L = 1 To UBound(sHardwareCfgs)
        tvwMain.Nodes.Add sHardwareCfgs(L) & "DisabledEnums", tvwChild, sHardwareCfgs(L) & "XPSecurity", SEC_XPSECURITY, "internet"
        tvwMain.Nodes.Add sHardwareCfgs(L) & "XPSecurity", tvwChild, sHardwareCfgs(L) & "XPFirewall", "Windows Firewall exceptions", "xpsec"
    
        sFirewallKeyD = "SYSTEM\" & sHardwareCfgs(L) & "\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile"
        sFirewallKeyS = "SYSTEM\" & sHardwareCfgs(L) & "\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile"
    
        tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewall", tvwChild, sHardwareCfgs(L) & "XPFirewallDomain", "Network domain mode", "lsp"
        tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewall", tvwChild, sHardwareCfgs(L) & "XPFirewallStandard", "Standalone mode", "system"
        tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomain").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyD
        tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandard").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyS
    
        sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sFirewallKeyD & "\AuthorizedApplications\List"), "|")
        If UBound(sVals) > -1 Or bShowEmpty Then
            tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewallDomain", tvwChild, sHardwareCfgs(L) & "XPFirewallDomainApps", "Applications", "exe"
            tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainApps").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyD & "\AuthorizedApplications\List"
            For i = 0 To UBound(sVals)
                sVals(i) = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
                sFile = Left$(sVals(i), InStr(3, sVals(i), ":") - 1)
                sFile = ExpandEnvironmentVars(sFile)
                sScope = Mid$(sVals(i), InStr(3, sVals(i), ":") + 1)
                bEnabled = IIf(InStr(1, sScope, ":Enabled:", vbTextCompare) > 0, True, False)
                sName = Mid$(sScope, InStr(sScope, ":") + 1)
                sName = Mid$(sName, InStr(sName, ":") + 1)
                If InStr(sName, "@") = 1 Then
                    sName = Mid$(sName, 2)
                    sName = GetStringResFromDLL(sSysDir & "\" & Left$(sName, InStr(sName, ",") - 1), Mid$(sName, InStr(sName, ",") + 1))
                End If
                sScope = Left$(sScope, InStr(sScope, ":") - 1)
                If sScope = "*" Then sScope = "any computer"
                
                tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewallDomainApps", tvwChild, sHardwareCfgs(L) & "XPFirewallDomainApps" & i, sName & " - " & sScope & " (" & IIf(bEnabled, "enabled)", "disabled)"), "firewall"
                tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainApps" & i).Tag = sFile
            Next i
            If tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainApps").Children > 0 Then
                tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainApps").Text = tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainApps").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainApps").Children & ")"
            End If
        End If
        sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sFirewallKeyD & "\GloballyOpenPorts\List"), "|")
        If UBound(sVals) > -1 Or bShowEmpty Then
            tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewallDomain", tvwChild, sHardwareCfgs(L) & "XPFirewallDomainPorts", "Ports", "internet"
            tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainPorts").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyD & "\GloballyOpenPorts\List"
            For i = 0 To UBound(sVals)
                sVals(i) = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
                sPort = Left$(sVals(i), InStr(3, sVals(i), ":") - 1)
                sProtocol = Mid$(sVals(i), InStr(3, sVals(i), ":") + 1)
                sScope = Mid$(sProtocol, InStr(sProtocol, ":") + 1)
                bEnabled = IIf(InStr(1, sScope, ":Enabled:", vbTextCompare) > 0, True, False)
                sName = Mid$(sScope, InStr(sScope, ":") + 1)
                sName = Mid$(sName, InStr(sName, ":") + 1)
                If InStr(sName, "@") = 1 Then
                    sName = Mid$(sName, 2)
                    sName = GetStringResFromDLL(sSysDir & "\" & Left$(sName, InStr(sName, ",") - 1), Mid$(sName, InStr(sName, ",") + 1))
                End If
                sProtocol = Left$(sProtocol, InStr(sProtocol, ":") - 1)
                sScope = Left$(sScope, InStr(sScope, ":") - 1)
                If sScope = "*" Then sScope = "any computer"
            
                tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewallDomainPorts", tvwChild, sHardwareCfgs(L) & "XPFirewallDomainPorts" & i, sName & " - " & sProtocol & " port " & sPort & " - " & sScope & " (" & IIf(bEnabled, "enabled)", "disabled)"), "firewall"
                tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainPorts" & i).Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyD & "\GloballyOpenPorts\List"
            Next i
            If tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainPorts").Children > 0 Then
                tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainPorts").Text = tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainPorts").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomainPorts").Children & ")"
            End If
        End If
        sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sFirewallKeyS & "\AuthorizedApplications\List"), "|")
        If UBound(sVals) > -1 Or bShowEmpty Then
            tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewallStandard", tvwChild, sHardwareCfgs(L) & "XPFirewallStandardApps", "Applications", "exe"
            tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardApps").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyS & "\AuthorizedApplications\List"
            For i = 0 To UBound(sVals)
                sVals(i) = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
                sFile = Left$(sVals(i), InStr(3, sVals(i), ":") - 1)
                sFile = ExpandEnvironmentVars(sFile)
                sScope = Mid$(sVals(i), InStr(3, sVals(i), ":") + 1)
                bEnabled = IIf(InStr(1, sScope, ":Enabled:", vbTextCompare) > 0, True, False)
                sName = Mid$(sScope, InStr(sScope, ":") + 1)
                sName = Mid$(sName, InStr(sName, ":") + 1)
                If InStr(sName, "@") = 1 Then
                    sName = Mid$(sName, 2)
                    sName = GetStringResFromDLL(sSysDir & "\" & Left$(sName, InStr(sName, ",") - 1), Mid$(sName, InStr(sName, ",") + 1))
                End If
                sScope = Left$(sScope, InStr(sScope, ":") - 1)
                If sScope = "*" Then sScope = "any computer"
                            
                tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewallStandardApps", tvwChild, sHardwareCfgs(L) & "XPFirewallStandardApps" & i, sName & " - " & sScope & " (" & IIf(bEnabled, "enabled)", "disabled)"), "firewall"
                tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardApps" & i).Tag = sFile
            Next i
            If tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardApps").Children > 0 Then
                tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardApps").Text = tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardApps").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardApps").Children & ")"
            End If
        End If
        sVals = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sFirewallKeyS & "\GloballyOpenPorts\List"), "|")
        If UBound(sVals) > -1 Or bShowEmpty Then
            tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewallStandard", tvwChild, sHardwareCfgs(L) & "XPFirewallStandardPorts", "Ports", "internet"
            tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardPorts").Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyS & "\GloballyOpenPorts\List"
            For i = 0 To UBound(sVals)
                sVals(i) = Mid$(sVals(i), InStr(sVals(i), " = ") + 3)
                sPort = Left$(sVals(i), InStr(3, sVals(i), ":") - 1)
                sProtocol = Mid$(sVals(i), InStr(3, sVals(i), ":") + 1)
                sScope = Mid$(sProtocol, InStr(sProtocol, ":") + 1)
                bEnabled = IIf(InStr(1, sScope, ":Enabled:", vbTextCompare) > 0, True, False)
                sName = Mid$(sScope, InStr(sScope, ":") + 1)
                sName = Mid$(sName, InStr(sName, ":") + 1)
                If InStr(sName, "@") = 1 Then
                    sName = Mid$(sName, 2)
                    sName = GetStringResFromDLL(sSysDir & "\" & Left$(sName, InStr(sName, ",") - 1), Mid$(sName, InStr(sName, ",") + 1))
                End If
                sProtocol = Left$(sProtocol, InStr(sProtocol, ":") - 1)
                sScope = Left$(sScope, InStr(sScope, ":") - 1)
                If sScope = "*" Then sScope = "any computer"
            
                tvwMain.Nodes.Add sHardwareCfgs(L) & "XPFirewallStandardPorts", tvwChild, sHardwareCfgs(L) & "XPFirewallStandardPorts" & i, sName & " - " & sProtocol & " port " & sPort & " - " & sScope & " (" & IIf(bEnabled, "enabled)", "disabled)"), "firewall"
                tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardPorts" & i).Tag = "HKEY_LOCAL_MACHINE\" & sFirewallKeyS & "\GloballyOpenPorts\List"
            Next i
            If tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardPorts").Children > 0 Then
                tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardPorts").Text = tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardPorts").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandardPorts").Children & ")"
            End If
        End If
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallDomain").Children = 0 Then
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "XPFirewallDomain"
        End If
        If tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewallStandard").Children = 0 Then
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "XPFirewallStandard"
        End If
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "XPFirewall").Children = 0 Then
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "XPFirewall"
        End If
        If tvwMain.Nodes(sHardwareCfgs(L) & "XPSecurity").Children = 0 Then
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "XPSecurity"
        End If
    Next L
    
    '----------------
    Dim sWDKey$, bWDDisable As Boolean
    sWDKey = "Software\Microsoft\Windows Defender"
    tvwMain.Nodes("WindowsDefender").Tag = "HKEY_LOCAL_MACHINE\" & sWDKey
    bWDDisable = CBool(Reg.GetDword(HKEY_LOCAL_MACHINE, sWDKey, "DisableAntiSpyware"))
    If bWDDisable Then
        tvwMain.Nodes.Add "WindowsDefender", tvwChild, "WindowsDefenderDisabled", "DisableAntiSpyware = 1", "reg"
        tvwMain.Nodes("WindowsDefenderDisabled").Tag = "HKEY_LOCAL_MACHINE\" & sWDKey
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "WindowsDefender"
    End If
    'system-wide
    
    AppendErrorLogCustom "EnumXPSecurity - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumXPSecurity"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumPrintMonitors()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumPrintMonitors - Begin"
    
    'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Monitors
    Dim sKeys$(), i&, sMonitors$, sName$, sFile$
    If bSL_Abort Then Exit Sub
    sMonitors = "System\CurrentControlSet\Control\Print\Monitors"
    
    tvwMain.Nodes.Add "System", tvwChild, "PrintMonitors", SEC_PRINTMONITORS, "printer"
    tvwMain.Nodes("PrintMonitors").Tag = "HKEY_LOCAL_MACHINE\" & sMonitors

    sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sMonitors), "|")
    For i = 0 To UBound(sKeys)
        sName = sKeys(i)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sMonitors & "\" & sName, "Driver")
        If sName = vbNullString Then sName = "(no name)"
        If sFile = vbNullString Then sFile = "(no file)"
        tvwMain.Nodes.Add "PrintMonitors", tvwChild, "PrintMonitors" & i, sName & " - " & sFile, "dll"
        tvwMain.Nodes("PrintMonitors" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("PrintMonitors").Children > 0 Then
        tvwMain.Nodes("PrintMonitors").Text = tvwMain.Nodes("PrintMonitors").Text & " (" & tvwMain.Nodes("PrintMonitors").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "PrintMonitors"
    End If

    If Not bShowHardware Then Exit Sub
    '----------------------------------------------------------------
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sMonitors = "System\" & sHardwareCfgs(L) & "\Control\Print\Monitors"

        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "PrintMonitors", SEC_PRINTMONITORS, "printer"
        tvwMain.Nodes(sHardwareCfgs(L) & "PrintMonitors").Tag = "HKEY_LOCAL_MACHINE\" & sMonitors
    
        sKeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sMonitors), "|")
        For i = 0 To UBound(sKeys)
            sName = sKeys(i)
            sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sMonitors & "\" & sName, "Driver")
            If sName = vbNullString Then sName = "(no name)"
            If sFile = vbNullString Then sFile = "(no file)"
            tvwMain.Nodes.Add sHardwareCfgs(L) & "PrintMonitors", tvwChild, sHardwareCfgs(L) & "PrintMonitors" & i, sName & " - " & sFile, "dll"
            tvwMain.Nodes(sHardwareCfgs(L) & "PrintMonitors" & i).Tag = GuessFullpathFromAutorun(sFile)
        Next i
        If tvwMain.Nodes(sHardwareCfgs(L) & "PrintMonitors").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "PrintMonitors").Text = tvwMain.Nodes(sHardwareCfgs(L) & "PrintMonitors").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "PrintMonitors").Children & ")"
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "PrintMonitors"
        End If
    Next L
    
    AppendErrorLogCustom "EnumPrintMonitors - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumPrintMonitors"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumSecurityProviders()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumSecurityProviders - Begin"
    
    '  HKLM\System\CurrentControlSet\Control\SecurityProviders
    Dim sSecProv$(), i&, sProviders$, sFile$
    If bSL_Abort Then Exit Sub
    sProviders = "System\CurrentControlSet\Control\SecurityProviders"
    
    tvwMain.Nodes.Add "System", tvwChild, "SecurityProviders", SEC_SECURITYPROVIDERS, "registry"
    tvwMain.Nodes("SecurityProviders").Tag = "HKEY_LOCAL_MACHINE\" & sProviders
    
    sSecProv = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sProviders, "SecurityProviders"), ",")
    For i = 0 To UBound(sSecProv)
        sFile = Trim$(sSecProv(i))
        
        tvwMain.Nodes.Add "SecurityProviders", tvwChild, "SecurityProviders" & i, sFile, "dll"
        tvwMain.Nodes("SecurityProviders" & i).Tag = GuessFullpathFromAutorun(sFile)
        If bSL_Abort Then Exit Sub
    Next i

    If tvwMain.Nodes("SecurityProviders").Children > 0 Then
        tvwMain.Nodes("SecurityProviders").Text = tvwMain.Nodes("SecurityProviders").Text & " (" & tvwMain.Nodes("SecurityProviders").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "SecurityProviders"
    End If

    If Not bShowHardware Then Exit Sub
    '----------------------------------------------------------------
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sProviders = "System\" & sHardwareCfgs(L) & "\Control\SecurityProviders"

        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "SecurityProviders", SEC_SECURITYPROVIDERS, "registry"
        tvwMain.Nodes(sHardwareCfgs(L) & "SecurityProviders").Tag = "HKEY_LOCAL_MACHINE\" & sProviders
        
        sSecProv = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sProviders, "SecurityProviders"), ",")
        For i = 0 To UBound(sSecProv)
            sFile = Trim$(sSecProv(i))
            
            tvwMain.Nodes.Add sHardwareCfgs(L) & "SecurityProviders", tvwChild, sHardwareCfgs(L) & "SecurityProviders" & i, sFile, "dll"
            tvwMain.Nodes(sHardwareCfgs(L) & "SecurityProviders" & i).Tag = GuessFullpathFromAutorun(sFile)
            If bSL_Abort Then Exit Sub
        Next i
    
        If tvwMain.Nodes(sHardwareCfgs(L) & "SecurityProviders").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "SecurityProviders").Text = tvwMain.Nodes(sHardwareCfgs(L) & "SecurityProviders").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "SecurityProviders").Children & ")"
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "SecurityProviders"
        End If
    Next L
    
    AppendErrorLogCustom "EnumSecurityProviders - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumSecurityProviders"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumDesktopComponents()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumDesktopComponents - Begin"
    
    'HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Desktop\Components
    Dim sDC$, sComponents$(), i&
    Dim sName$, sSource$, sSubURL$
    If bSL_Abort Then Exit Sub
    sDC = "Software\Microsoft\Internet Explorer\Desktop\Components"
    
    tvwMain.Nodes.Add "System", tvwChild, "DesktopComponents", SEC_DESKTOPCOMPONENTS, "msie"
    tvwMain.Nodes("DesktopComponents").Tag = "HKEY_CURRENT_USER\" & sDC
    
    sComponents = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sDC), "|")
    For i = 0 To UBound(sComponents)
        sName = Reg.GetString(HKEY_CURRENT_USER, sDC & "\" & sComponents(i), "FriendlyName")
        sSource = Reg.GetString(HKEY_CURRENT_USER, sDC & "\" & sComponents(i), "Source")
        sSubURL = Reg.GetString(HKEY_CURRENT_USER, sDC & "\" & sComponents(i), "SubscribedURL")
        
        tvwMain.Nodes.Add "DesktopComponents", tvwChild, "DesktopComponents" & i, sName & " - " & sSource & " - " & sSubURL, "reg"
        tvwMain.Nodes("DesktopComponents" & i).Tag = "HKEY_CURRENT_USER\" & sDC & "\" & sComponents(i)
        If bSL_Abort Then Exit Sub
    Next i

    If tvwMain.Nodes("DesktopComponents").Children > 0 Then
        tvwMain.Nodes("DesktopComponents").Text = tvwMain.Nodes("DesktopComponents").Text & " (" & tvwMain.Nodes("DesktopComponents").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "DesktopComponents"
    End If

    If Not bShowUsers Then Exit Sub
    '-----------------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "DesktopComponents", SEC_DESKTOPCOMPONENTS, "msie"
            tvwMain.Nodes(sUsernames(L) & "DesktopComponents").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sDC
            
            sComponents = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\" & sDC), "|")
            For i = 0 To UBound(sComponents)
                sName = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sDC & "\" & sComponents(i), "FriendlyName")
                sSource = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sDC & "\" & sComponents(i), "Source")
                sSubURL = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sDC & "\" & sComponents(i), "SubscribedURL")
                
                tvwMain.Nodes.Add sUsernames(L) & "DesktopComponents", tvwChild, sUsernames(L) & "DesktopComponents" & i, sName & " - " & sSource & " - " & sSubURL, "reg"
                tvwMain.Nodes(sUsernames(L) & "DesktopComponents" & i).Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sDC & "\" & sComponents(i)
                If bSL_Abort Then Exit Sub
            Next i
        
            If tvwMain.Nodes(sUsernames(L) & "DesktopComponents").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "DesktopComponents").Text = tvwMain.Nodes(sUsernames(L) & "DesktopComponents").Text & " (" & tvwMain.Nodes(sUsernames(L) & "DesktopComponents").Children & ")"
            Else
                If Not bShowEmpty Then tvwMain.Nodes.Remove sUsernames(L) & "DesktopComponents"
            End If
        End If
    Next L
    
    AppendErrorLogCustom "EnumDesktopComponents - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumDesktopComponents"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumAppPaths()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumAppPaths - Begin"

    'HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths
    Dim sAPKey$, sApps$(), i&, sExe$
    If bSL_Abort Then Exit Sub
    sAPKey = "Software\Microsoft\Windows\CurrentVersion\App Paths"
    
    tvwMain.Nodes.Add "System", tvwChild, "AppPaths", SEC_APPPATHS, "registry"
    tvwMain.Nodes("AppPaths").Tag = "HKEY_LOCAL_MACHINE\" & sAPKey
    
    sApps = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sAPKey), "|")
    For i = 0 To UBound(sApps)
        sExe = Reg.GetString(HKEY_LOCAL_MACHINE, sAPKey & "\" & sApps(i), vbNullString)
        sExe = ExpandEnvironmentVars(sExe)
        sExe = GetLongFilename(sExe)
        tvwMain.Nodes.Add "AppPaths", tvwChild, "AppPaths" & i, sApps(i) & " - " & sExe, "exe"
        tvwMain.Nodes("AppPaths" & i).Tag = sExe
        If bSL_Abort Then Exit Sub
    Next i
    If tvwMain.Nodes("AppPaths").Children > 0 Then
        tvwMain.Nodes("AppPaths").Text = tvwMain.Nodes("AppPaths").Text & " (" & tvwMain.Nodes("AppPaths").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "AppPaths"
    End If
    '------------------------------------
    'nothing, this is system-wide
    
    AppendErrorLogCustom "EnumAppPaths - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumAppPaths"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumMountPoints()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumMountPoints - Begin"
    
    'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints (win9x/2000)
    'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2 (winxp)
    Dim sMPKey$, sMPKey2$, sKeys$(), i&, sCmd$
    If bSL_Abort Then Exit Sub
    sMPKey = "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints"
    sMPKey2 = "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
    
    tvwMain.Nodes.Add "System", tvwChild, "MountPoints", SEC_MOUNTPOINTS, "drive"
    tvwMain.Nodes("MountPoints").Tag = "HKEY_CURRENT_USER\" & sMPKey2
    
    sKeys = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sMPKey), "|")
    For i = 0 To UBound(sKeys)
        sCmd = Reg.GetString(HKEY_CURRENT_USER, sMPKey & "\" & sKeys(i) & "\shell\Autoplay\command", vbNullString)
        If sCmd <> vbNullString Then
            tvwMain.Nodes.Add "MountPoints", tvwChild, "MountPoints" & i, sKeys(i) & " - " & sCmd, "reg"
            tvwMain.Nodes("MountPoints" & i).Tag = GuessFullpathFromAutorun(sCmd)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    sKeys = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sMPKey2), "|")
    For i = 0 To UBound(sKeys)
        sCmd = Reg.GetString(HKEY_CURRENT_USER, sMPKey2 & "\" & sKeys(i) & "\shell\Autoplay\command", vbNullString)
        If sCmd <> vbNullString Then
            tvwMain.Nodes.Add "MountPoints", tvwChild, "MountPoints2" & i, sKeys(i) & " - " & sCmd, "reg"
            tvwMain.Nodes("MountPoints2" & i).Tag = GuessFullpathFromAutorun(sCmd)
        End If
        If bSL_Abort Then Exit Sub
    Next i
    
    If tvwMain.Nodes("MountPoints").Children > 0 Then
        tvwMain.Nodes("MountPoints").Text = tvwMain.Nodes("MountPoints").Text & " (" & tvwMain.Nodes("MountPoints").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "MountPoints"
    End If
    
    If Not bShowUsers Then Exit Sub
    '-----------------------------------------------------------------------
    Dim sUsername$, L&
    For L = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(L))
        If sUsername <> OSver.UserName And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(L), tvwChild, sUsernames(L) & "MountPoints", SEC_MOUNTPOINTS, "drive"
            tvwMain.Nodes(sUsernames(L) & "MountPoints").Tag = "HKEY_USERS\" & sUsernames(L) & "\" & sMPKey2
            
            
            sKeys = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\" & sMPKey), "|")
            For i = 0 To UBound(sKeys)
                sCmd = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sMPKey & "\" & sKeys(i) & "\shell\Autoplay\command", vbNullString)
                If sCmd <> vbNullString Then
                    tvwMain.Nodes.Add sUsernames(L) & "MountPoints", tvwChild, sUsernames(L) & "MountPoints" & i, sKeys(i) & " - " & sCmd, "reg"
                    tvwMain.Nodes(sUsernames(L) & "MountPoints" & i).Tag = GuessFullpathFromAutorun(sCmd)
                End If
                If bSL_Abort Then Exit Sub
            Next i
            sKeys = Split(Reg.EnumSubKeys(HKEY_USERS, sUsernames(L) & "\" & sMPKey2), "|")
            For i = 0 To UBound(sKeys)
                sCmd = Reg.GetString(HKEY_USERS, sUsernames(L) & "\" & sMPKey2 & "\" & sKeys(i) & "\shell\Autoplay\command", vbNullString)
                If sCmd <> vbNullString Then
                    tvwMain.Nodes.Add sUsernames(L) & "MountPoints", tvwChild, sUsernames(L) & "MountPoints2" & i, sKeys(i) & " - " & sCmd, "reg"
                    tvwMain.Nodes(sUsernames(L) & "MountPoints2" & i).Tag = GuessFullpathFromAutorun(sCmd)
                End If
                If bSL_Abort Then Exit Sub
            Next i
            
            
            If tvwMain.Nodes(sUsernames(L) & "MountPoints").Children > 0 Then
                tvwMain.Nodes(sUsernames(L) & "MountPoints").Text = tvwMain.Nodes(sUsernames(L) & "MountPoints").Text & " (" & tvwMain.Nodes(sUsernames(L) & "MountPoints").Children & ")"
            Else
                If Not bShowEmpty Then tvwMain.Nodes.Remove sUsernames(L) & "MountPoints"
            End If
        End If
    Next L
    
    AppendErrorLogCustom "EnumMountPoints - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumMountPoints"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumLSAPackages()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "EnumLSAPackages - Begin"
    
    'HKLM\SYSTEM\CurrentControlSet\Control\Lsa
    'values: Authentication Packages, Notification Packages, Security Packages
    Dim sAuthPgs$(), sNotiPgs$(), sSecuPgs$(), i&, sRegKey$
    Dim Stady As Long
    sRegKey = "System\CurrentControlSet\Control\Lsa"
    
    tvwMain.Nodes.Add "System", tvwChild, "LsaPackages", SEC_LSAPACKAGES, "winlogon"
    tvwMain.Nodes("LsaPackages").Tag = "HKEY_LOCAL_MACHINE\" & sRegKey
    
    Stady = 1
    
    sAuthPgs = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sRegKey, "Authentication Packages", False), Chr$(0))
    sNotiPgs = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sRegKey, "Notification Packages", False), Chr$(0))
    sSecuPgs = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sRegKey, "Security Packages", False), Chr$(0))
    
    Stady = 2
    
    tvwMain.Nodes.Add "LsaPackages", tvwChild, "LsaPackagesAuth", "Authentication Packages", "policy"
    tvwMain.Nodes("LsaPackagesAuth").Tag = "HKEY_LOCAL_MACHINE\" & sRegKey
    
    Stady = 3
    
    For i = 0 To UBound(sAuthPgs)
        If Trim$(sAuthPgs(i)) <> vbNullString Then
            tvwMain.Nodes.Add "LsaPackagesAuth", tvwChild, "LsaPackagesAuth" & i, sAuthPgs(i) & ".dll", "dll"
            tvwMain.Nodes("LsaPackagesAuth" & i).Tag = GuessFullpathFromAutorun(sAuthPgs(i) & ".dll")
        End If
    Next i
    
    Stady = 4
    
    tvwMain.Nodes.Add "LsaPackages", tvwChild, "LsaPackagesNoti", "Notification Packages", "policy"
    tvwMain.Nodes("LsaPackagesNoti").Tag = "HKEY_LOCAL_MACHINE\" & sRegKey
    
    Stady = 5
    
    For i = 0 To UBound(sNotiPgs)
        If Trim$(sNotiPgs(i)) <> vbNullString Then
            tvwMain.Nodes.Add "LsaPackagesNoti", tvwChild, "LsaPackagesNoti" & i, sNotiPgs(i) & ".dll", "dll"
            tvwMain.Nodes("LsaPackagesNoti" & i).Tag = GuessFullpathFromAutorun(sNotiPgs(i) & ".dll")
        End If
    Next i
    
    Stady = 6
    
    tvwMain.Nodes.Add "LsaPackages", tvwChild, "LsaPackagesSecu", "Security Packages", "policy"
    tvwMain.Nodes("LsaPackagesSecu").Tag = "HKEY_LOCAL_MACHINE\" & sRegKey
    
    Stady = 7
    
    For i = 0 To UBound(sSecuPgs)
        If Trim$(sSecuPgs(i)) <> vbNullString Then
            tvwMain.Nodes.Add "LsaPackagesSecu", tvwChild, "LsaPackagesSecu" & i, sSecuPgs(i) & ".dll", "dll"
            tvwMain.Nodes("LsaPackagesSecu" & i).Tag = GuessFullpathFromAutorun(sSecuPgs(i) & ".dll")
        End If
    Next i
    
    Stady = 8
    
    If tvwMain.Nodes("LsaPackagesAuth").Children > 0 Then
        tvwMain.Nodes("LsaPackagesAuth").Text = tvwMain.Nodes("LsaPackagesAuth").Text & " (" & tvwMain.Nodes("LsaPackagesAuth").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "LsaPackagesAuth"
    End If
    
    Stady = 9
    
    If tvwMain.Nodes("LsaPackagesNoti").Children > 0 Then
        tvwMain.Nodes("LsaPackagesNoti").Text = tvwMain.Nodes("LsaPackagesNoti").Text & " (" & tvwMain.Nodes("LsaPackagesNoti").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "LsaPackagesNoti"
    End If
    
    Stady = 10
    
    If tvwMain.Nodes("LsaPackagesSecu").Children > 0 Then
        tvwMain.Nodes("LsaPackagesSecu").Text = tvwMain.Nodes("LsaPackagesSecu").Text & " (" & tvwMain.Nodes("LsaPackagesSecu").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "LsaPackagesSecu"
    End If
    
    Stady = 11
    
    If tvwMain.Nodes("LsaPackages").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "LsaPackages"
    End If
    
    If Not bShowHardware Then Exit Sub
    '----------------------------------------------------------------
    
    Stady = 12
    
    Dim L&
    For L = 1 To UBound(sHardwareCfgs)
        sRegKey = "System\" & sHardwareCfgs(L) & "\Control\Lsa"

        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(L), tvwChild, sHardwareCfgs(L) & "LsaPackages", SEC_LSAPACKAGES, "winlogon"
        tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackages").Tag = "HKEY_LOCAL_MACHINE\" & sRegKey
        
        Stady = 13
        
        sAuthPgs = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sRegKey, "Authentication Packages", False), Chr$(0))
        sNotiPgs = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sRegKey, "Notification Packages", False), Chr$(0))
        sSecuPgs = Split(Reg.GetString(HKEY_LOCAL_MACHINE, sRegKey, "Security Packages", False), Chr$(0))
        
        Stady = 14
        
        tvwMain.Nodes.Add sHardwareCfgs(L) & "LsaPackages", tvwChild, sHardwareCfgs(L) & "LsaPackagesAuth", "Authentication Packages", "policy"
        tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesAuth").Tag = "HKEY_LOCAL_MACHINE\" & sRegKey
        
        Stady = 15
        
        For i = 0 To UBound(sAuthPgs)
            If Trim$(sAuthPgs(i)) <> vbNullString Then
                tvwMain.Nodes.Add sHardwareCfgs(L) & "LsaPackagesAuth", tvwChild, sHardwareCfgs(L) & "LsaPackagesAuth" & i, sAuthPgs(i) & ".dll", "dll"
                tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesAuth" & i).Tag = GuessFullpathFromAutorun(sAuthPgs(i) & ".dll")
            End If
        Next i
        
        Stady = 16
        
        tvwMain.Nodes.Add sHardwareCfgs(L) & "LsaPackages", tvwChild, sHardwareCfgs(L) & "LsaPackagesNoti", "Notification Packages", "policy"
        tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesNoti").Tag = "HKEY_LOCAL_MACHINE\" & sRegKey
        
        Stady = 17
        
        For i = 0 To UBound(sNotiPgs)
            If Trim$(sNotiPgs(i)) <> vbNullString Then
                tvwMain.Nodes.Add sHardwareCfgs(L) & "LsaPackagesNoti", tvwChild, sHardwareCfgs(L) & "LsaPackagesNoti" & i, sNotiPgs(i) & ".dll", "dll"
                tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesNoti" & i).Tag = GuessFullpathFromAutorun(sNotiPgs(i) & ".dll")
            End If
        Next i
        
        Stady = 18
        
        tvwMain.Nodes.Add sHardwareCfgs(L) & "LsaPackages", tvwChild, sHardwareCfgs(L) & "LsaPackagesSecu", "Security Packages", "policy"
        tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesSecu").Tag = "HKEY_LOCAL_MACHINE\" & sRegKey
        
        Stady = 19
        
        For i = 0 To UBound(sSecuPgs)
            If Trim$(sSecuPgs(i)) <> vbNullString Then
                tvwMain.Nodes.Add sHardwareCfgs(L) & "LsaPackagesSecu", tvwChild, sHardwareCfgs(L) & "LsaPackagesSecu" & i, sSecuPgs(i) & ".dll", "dll"
                tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesSecu" & i).Tag = GuessFullpathFromAutorun(sSecuPgs(i) & ".dll")
            End If
        Next i
        
        Stady = 20
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesAuth").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesAuth").Text = tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesAuth").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesAuth").Children & ")"
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "LsaPackagesAuth"
        End If
        
        Stady = 21
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesNoti").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesNoti").Text = tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesNoti").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesNoti").Children & ")"
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "LsaPackagesNoti"
        End If
        
        Stady = 22
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesSecu").Children > 0 Then
            tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesSecu").Text = tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesSecu").Text & " (" & tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackagesSecu").Children & ")"
        Else
            If Not bShowEmpty Then tvwMain.Nodes.Remove sHardwareCfgs(L) & "LsaPackagesSecu"
        End If
        
        Stady = 23
        
        If tvwMain.Nodes(sHardwareCfgs(L) & "LsaPackages").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove sHardwareCfgs(L) & "LsaPackages"
        End If
    Next L
    
    AppendErrorLogCustom "EnumLSAPackages - End", "Stady: " & Stady, "Iteration: " & i
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "EnumLSAPackages"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub scrSaveSections_Change()
    fraScroller.Top = -scrSaveSections.Value
End Sub

Private Sub scrSaveSections_Scroll()
    scrSaveSections_Change
End Sub

Private Sub tvwMain_KeyDown(KeyCode As Integer, Shift As Integer)
    'moved this from KeyUp to KeyDown to prevent closing a window above SL2
    'with Esc closing SL2 as well when you release the key
    If KeyCode = 27 Then cmdAbort_Click
    If KeyCode = 118 Then Unload Me 'End 'F7
End Sub

Private Sub tvwMain_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Status tvwMain.SelectedItem.Tag
    If mnuHelpShow.Checked Then txtHelp.Text = GetHelpText(tvwMain.SelectedItem.Key)
End Sub

Private Sub tvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Status tvwMain.SelectedItem.Tag
    If mnuHelpShow.Checked Then txtHelp.Text = GetHelpText(tvwMain.SelectedItem.Key)
    If Button = 2 Then
        'node was right-clicked
        'not a valid regkey? disable reg crap
        If Not NodeIsValidRegkey(tvwMain.SelectedItem) Then
            mnuPopupStr2.Visible = False
            mnuPopupRegJump.Visible = False
            mnuPopupRegkeyCopy.Visible = False
        End If
        'not a valid file ? disable file stuff
        If Not NodeIsValidFile(tvwMain.SelectedItem) Then
            mnuPopupShowFile.Visible = False
            mnuPopupShowProp.Visible = False
            mnuPopupFilenameCopy.Visible = False
            mnuPopupNotepad.Visible = False
            mnuPopupVerifyFile.Visible = False
            mnuPopupFileRunScanner.Visible = False
            mnuPopupFileGoogle.Visible = False
            mnuPopupStr3.Visible = False
            mnuPopupStr4.Visible = False
        End If
        'neither? remove divider as well
        If Not NodeIsValidFile(tvwMain.SelectedItem) And _
           Not NodeIsValidRegkey(tvwMain.SelectedItem) Then
            mnuPopupStr2.Visible = False
            mnuPopupStr3.Visible = False
            mnuPopupStr4.Visible = False
        End If
        
        'check if a CLSID is in there somewhere
        If (InStr(tvwMain.SelectedItem.Text, "{") > 0 And _
           InStr(tvwMain.SelectedItem.Text, "}") > 0) Or _
           (InStr(tvwMain.SelectedItem.Tag, "{") > 0 And _
           InStr(tvwMain.SelectedItem.Tag, "}") > 0) Then
            mnuPopupCLSIDRunScanner.Visible = True
            mnuPopupCLSIDGoogle.Visible = True
        Else
            mnuPopupCLSIDRunScanner.Visible = False
            mnuPopupCLSIDGoogle.Visible = False
        End If
        
        'show the popup menu
        PopupMenu mnuPopup
        
        're-enable all disabled stuff
        mnuPopupShowFile.Visible = True
        mnuPopupShowProp.Visible = True
        mnuPopupFilenameCopy.Visible = True
        mnuPopupNotepad.Visible = True
        mnuPopupVerifyFile.Visible = True
        mnuPopupFileRunScanner.Visible = True
        mnuPopupCLSIDRunScanner.Visible = True
        mnuPopupFileGoogle.Visible = True
        mnuPopupCLSIDGoogle.Visible = True
        mnuPopupRegJump.Visible = True
        mnuPopupRegkeyCopy.Visible = True
        mnuPopupStr2.Visible = True
        mnuPopupStr3.Visible = True
        mnuPopupStr4.Visible = True
    End If
End Sub

Public Sub ShowError(sMsg$)
    If Not picWarning.Visible Then
        mnuHelpWarning.Checked = True
        txtWarning.Visible = True
        picWarning.Visible = True
        
        mnuHelpShow.Checked = False
        picHelp.Visible = False
        txtHelp.Visible = False
        Form_Resize
    End If
    txtWarning.Text = txtWarning.Text & "[" & Format$(time, "Hh:Mm:Ss") & "] " & sMsg & vbCrLf
End Sub

Public Function IsSectionChecked(sKey$) As Boolean
    'node must exist
    If Not NodeExists(sKey) Then Exit Function
    
    'tag not set: it's not a section, so do it
    'tag set to 1: do the section
    'tag set to 0: skip the section
    Select Case tvwMain.Nodes(sKey).Tag
        Case "1": IsSectionChecked = True
        Case "0": IsSectionChecked = False
        Case Else: IsSectionChecked = True
    End Select
End Function

Private Sub LoadSectionNames()
    chkSectionFiles(1).Caption = SEC_RUNNINGPROCESSES
    chkSectionFiles(2).Caption = SEC_AUTOSTARTFOLDERS
    chkSectionFiles(3).Caption = SEC_TASKSCHEDULER
    chkSectionFiles(4).Caption = SEC_INIFILE
    chkSectionFiles(5).Caption = SEC_AUTORUNINF
    chkSectionFiles(6).Caption = SEC_BATFILES
    chkSectionFiles(7).Caption = SEC_EXPLORERCLONES
    
    chkSectionMSIE(1).Caption = SEC_BHOS
    chkSectionMSIE(2).Caption = SEC_IETOOLBARS
    chkSectionMSIE(3).Caption = SEC_IEEXTENSIONS
    chkSectionMSIE(4).Caption = SEC_IEBARS
    chkSectionMSIE(5).Caption = SEC_IEMENUEXT
    chkSectionMSIE(6).Caption = SEC_IEBANDS
    chkSectionMSIE(7).Caption = SEC_DPFS
    chkSectionMSIE(8).Caption = SEC_URLSEARCHHOOKS
    chkSectionMSIE(9).Caption = SEC_ACTIVEX
    chkSectionMSIE(10).Caption = SEC_DESKTOPCOMPONENTS
    
    chkSectionHijack(1).Caption = SEC_RESETWEBSETTINGS
    chkSectionHijack(2).Caption = SEC_IEURLS
    chkSectionHijack(3).Caption = SEC_URLPREFIX
    chkSectionHijack(4).Caption = SEC_HOSTSFILEPATH
    
    chkSectionDisabled(1).Caption = SEC_HOSTSFILE
    chkSectionDisabled(2).Caption = SEC_KILLBITS
    chkSectionDisabled(3).Caption = SEC_ZONES
    chkSectionDisabled(4).Caption = SEC_MSCONFIG9X
    chkSectionDisabled(5).Caption = SEC_MSCONFIGXP
    chkSectionDisabled(6).Caption = SEC_XPSECURITY
    chkSectionDisabled(7).Caption = SEC_STOPPEDSERVICES
    
    chkSectionRegistry(1).Caption = SEC_INIMAPPING
    chkSectionRegistry(2).Caption = SEC_MOUNTPOINTS
    chkSectionRegistry(3).Caption = SEC_SCRIPTPOLICIES
    chkSectionRegistry(4).Caption = SEC_ONREBOOT
    chkSectionRegistry(5).Caption = SEC_SHELLCOMMANDS
    chkSectionRegistry(6).Caption = SEC_SERVICES
    chkSectionRegistry(7).Caption = SEC_DRIVERFILTERS
    chkSectionRegistry(8).Caption = SEC_PRINTMONITORS
    chkSectionRegistry(9).Caption = SEC_WINLOGON
    chkSectionRegistry(10).Caption = SEC_LSAPACKAGES
    chkSectionRegistry(11).Caption = SEC_POLICIES
    chkSectionRegistry(12).Caption = SEC_IMAGEFILEEXECUTION
    chkSectionRegistry(13).Caption = SEC_CONTEXTMENUHANDLERS
    chkSectionRegistry(14).Caption = SEC_COLUMNHANDLERS
    chkSectionRegistry(15).Caption = SEC_SHELLEXECUTEHOOKS
    chkSectionRegistry(16).Caption = SEC_SHELLEXT
    chkSectionRegistry(17).Caption = SEC_REGRUNKEYS
    chkSectionRegistry(18).Caption = SEC_REGRUNEXKEYS
    chkSectionRegistry(19).Caption = SEC_PROTOCOLS
    chkSectionRegistry(20).Caption = SEC_WOW
    chkSectionRegistry(21).Caption = SEC_SSODL
    chkSectionRegistry(22).Caption = SEC_SHAREDTASKSCHEDULER
    chkSectionRegistry(23).Caption = SEC_MPRSERVICES
    chkSectionRegistry(24).Caption = SEC_SECURITYPROVIDERS
    chkSectionRegistry(25).Caption = SEC_APPPATHS
    chkSectionRegistry(26).Caption = SEC_WINSOCKLSP
    chkSectionRegistry(27).Caption = SEC_CMDPROC
    chkSectionRegistry(28).Caption = SEC_UTILMANAGER
    chkSectionRegistry(29).Caption = SEC_3RDPARTY
    chkSectionRegistry(30).Caption = SEC_DRIVERS32
End Sub

Private Function GetString(lHive&, sKey$, sVal$, Optional bTrimNull As Boolean = True) As String
    On Error GoTo ErrorHandler:
    Dim hKey&, uData() As Byte, lDataLen&, sData$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        RegQueryValueEx hKey, sVal, 0, 0, ByVal 0, lDataLen
        ReDim uData(lDataLen)
        If RegQueryValueEx(hKey, sVal, 0, 0, uData(0), lDataLen) = 0 Then
            If bTrimNull Then
                sData = StrConv(uData, vbUnicode)
                sData = TrimNull(sData)
            Else
                If lDataLen > 2 Then
                    ReDim Preserve uData(lDataLen - 2)
                    sData = StrConv(uData, vbUnicode)
                End If
            End If
            GetString = sData
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetString"
    If inIDE Then Stop: Resume Next
End Function

Private Function GetDword&(lHive$, sKey$, sVal$)
    On Error GoTo ErrorHandler:
    Dim hKey&, lData&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        If RegQueryValueEx(hKey, sVal, 0, 0, lData, 4) = 0 Then
            GetDword = lData
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetDword"
    If inIDE Then Stop: Resume Next
End Function

Private Function KeyExists(lHive&, sKey$) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        KeyExists = True
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "KeyExists"
    If inIDE Then Stop: Resume Next
End Function

Private Function RegValExists(lHive&, sKey$, sVal$) As Boolean
    On Error GoTo ErrorHandler:
    Dim hKey&, lDataLen&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        If RegQueryValueEx(hKey, sVal, 0, 0, ByVal 0, lDataLen) = 0 Then
            RegValExists = True
        End If
        RegCloseKey hKey
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegValExists"
    If inIDE Then Stop: Resume Next
End Function

Private Function EnumSubKeys$(lHive&, sKey$)
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, sName$, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(MAX_PATH, 0)
        Do Until RegEnumKeyEx(hKey, i, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0
            sName = TrimNull(sName)
            sList = sList & "|" & sName
            i = i + 1
            sName = String$(MAX_PATH, 0)
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then EnumSubKeys = Mid$(sList, 2)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "EnumSubKeys"
    If inIDE Then Stop: Resume Next
End Function

Private Function RegEnumValues$(lHive&, sKey$, Optional bNullSep As Boolean = False, Optional bIgnoreBinaries As Boolean = True, Optional bIgnoreDwords As Boolean = True)
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, sName$, uData() As Byte, lDataLen&
    Dim lType&, sData$, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(lEnumBufLen, 0)
        ReDim uData(32768)
        lDataLen = UBound(uData)
        Do Until RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0
            
            sName = TrimNull(sName)
            If sName = vbNullString Then sName = "@"
            
            If lType = REG_SZ Then
                ReDim Preserve uData(lDataLen)
                sData = TrimNull(StrConv(uData, vbUnicode))
                If bNullSep Then
                    sList = sList & Chr$(0) & sName & " = " & sData
                Else
                    sList = sList & "|" & sName & " = " & sData
                End If
            End If
            
            If lType = REG_BINARY And Not bIgnoreBinaries Then
                sList = sList & "|" & sName & " (binary)"
            End If
            
            If lType = REG_DWORD And Not bIgnoreDwords Then
                'look at me! I'm haxxoring word values from binary!
                'sData = "dword: " & Hex$(uData(0)) & "." & Hex$(uData(1)) & "." & Hex$(uData(2)) & "." & Hex$(uData(3))
                'sData = "dword: " & Val("&H" & Hex$(uData(3)) & Hex$(uData(2)) & Hex$(uData(1)) & Hex$(uData(0)))
                sData = "dword: " & CStr(16 ^ 6 * uData(3) + 16 ^ 4 * uData(2) + 16 ^ 2 * uData(1) + uData(0))
                sList = sList & "|" & sName & " = " & sData
            End If
            sName = String$(lEnumBufLen, 0)
            ReDim uData(32768)
            lDataLen = UBound(uData)
            i = i + 1
            
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then RegEnumValues = Mid$(sList, 2)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegEnumValues"
    If inIDE Then Stop: Resume Next
End Function

Private Function RegEnumDwordValues$(lHive&, sKey$)
    On Error GoTo ErrorHandler:
    Dim hKey&, i&, sName$, uData() As Byte, lDataLen&
    Dim lType&, lData&, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(lEnumBufLen, 0)
        ReDim uData(32768)
        lDataLen = UBound(uData)
        Do Until RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0
            If lType = REG_DWORD And lDataLen = 4 Then
                sName = TrimNull(sName)
                If sName = vbNullString Then sName = "@"
                ReDim Preserve uData(4)
                CopyMemory lData, uData(0), 4
                sList = sList & "|" & sName & " = " & CStr(lData)
            End If
            sName = String$(lEnumBufLen, 0)
            ReDim uData(32768)
            lDataLen = UBound(uData)
            i = i + 1
        
            If bSL_Abort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then RegEnumDwordValues = Mid$(sList, 2)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "RegEnumDwordValues"
    If inIDE Then Stop: Resume Next
End Function

Private Sub Form_Activate()
    'bSL_Abort = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bSL_Abort = True
    bSL_Terminate = True
    'if launched from /tool:StartupList cmdline
    SaveWindowPos Me, SETTINGS_SECTION_STARTUPLIST
    If g_bStartupListTerminateOnExit Then Unload frmMain
End Sub
