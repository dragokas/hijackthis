VERSION 5.00
Begin VB.Form frmSecurity 
   Caption         =   "Security Fixes & Protection"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Performance && system cleaning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3840
      TabIndex        =   21
      Top             =   1080
      Width           =   3615
      Begin VB.CheckBox ChkSetServiceDelayed 
         Caption         =   "Set delayed state for some services"
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   2040
         Width           =   3135
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Clean temp files and backups with MS CleanMgr"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox ChkEnablePageFile 
         Caption         =   "Enable PageFile"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox ChkEnableReadyBoot 
         Caption         =   "Enable ReadyBoot"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox ChkEnableSuperfetch 
         Caption         =   "Enable Superfetch"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Fra3 
      Caption         =   "Basic security"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   3615
      Begin VB.CheckBox Check12 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkSetUAC 
         Caption         =   "Set UAC level at maximum"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Fra2 
      Caption         =   "Anti-Spy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3840
      TabIndex        =   7
      Top             =   3840
      Width           =   3615
      Begin VB.CheckBox chkDisableWPAD 
         Caption         =   "Disable WPAD protocol"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CheckBox ChkSpyDomains 
         Caption         =   "Disable spying domains"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   3135
      End
      Begin VB.CheckBox ChkSpyTasks 
         Caption         =   "Disable spying tasks"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox ChkGetWin10 
         Caption         =   "Disable ""Get Windows 10"""
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox ChkSpyUpdates 
         Caption         =   "Disable spying updates"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox ChkTelemetry 
         Caption         =   "Disable keylogger and telemetry"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CheckBox chkDisablePerfMeasure 
         Caption         =   "Disable performance assessment tools"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   3135
      End
   End
   Begin VB.Frame frm1 
      Caption         =   "Anti-ransom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   3615
      Begin VB.CheckBox chkDisableDDE 
         Caption         =   "Disable DDE (MsOffice)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CheckBox chkCheck6 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CheckBox chkCheck5 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CheckBox chkDisableSMBv1 
         Caption         =   "Disable SMBv1 protocol"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox chkDisableWScript 
         Caption         =   "Disable WScript Host scripts"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox chkDisablePoweShell 
         Caption         =   "Disable PoweShell scripts"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkInstallNullSector 
         Caption         =   "Install null sector protection driver"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Label lblLabel3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Security Fixes && Protection center"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   7365
   End
   Begin VB.Label lblLabel2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSecurity.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   30
      Top             =   480
      Width           =   7215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabel1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please, carefully read tooltips under each RED checkbox and do not change anything unless you really understand what you do."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   120
      TabIndex        =   29
      Top             =   6840
      Width           =   7395
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Options"
      Begin VB.Menu mnuOptReset 
         Caption         =   "Reset to defaults"
      End
      Begin VB.Menu mnuOptRecom 
         Caption         =   "Set recommended protection"
      End
      Begin VB.Menu mnuOptMax 
         Caption         =   "Set maximum protection"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Security Fixes & Protection module by Alex Dragokas

'
' Work in progress ...
'


' *--------------------------------------------------------------------------------------------*
'                                     Basic security
' *--------------------------------------------------------------------------------------------*

Private Sub chkSetUAC_Click()
    '
    
End Sub



' *--------------------------------------------------------------------------------------------*
'                                     Anti-Ransomeware
' *--------------------------------------------------------------------------------------------*

Private Sub chkDisablePoweShell_Click()
    '
    'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\PowerShell => ExecutionPolicy
    'gpedit.msc => Конфиг. компьютера => Административные шаблоны => Компоненты Windows => Windows PowerShell
    
End Sub

Private Sub chkDisableSMBv1_Click()
    'like EternalBlue
    'https://technet.microsoft.com/en-us/library/security/ms17-010.aspx
    '
    'sc.exe config lanmanworkstation depend=bowser/mrxsmb20/nsi
    'только нужно предварительно проверить, какие там уже стоят службы в зависимостях, иначе, может получится, что мы добавили новую зависимость
    'к службе, которой либо не существует, либо она отключена.

    'sc.exe config mrxsmb10 start=disabled
    
    'dism /online /norestart /disable-feature /featurename:SMB1Protocol
End Sub

Private Sub chkDisableWScript_Click()
    '
    
    'set REG_DWORD to 0
    'HKEY_CURRENT_USER\Software\Microsoft\Windows Script Host\Settings => Enabled
    'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows Script Host\Settings => Enabled
    
    'or maybe replace by Dragokas' Anti-Shell (?)
    
End Sub

Private Sub chkInstallNullSector_Click()
    '
    
    ' MBRFilter.inf -> verb "Install"
End Sub

Private Sub chkDisableDDE_Click()
    'https://technet.microsoft.com/en-us/library/security/4053440?f=255&MSPPError=-2147217396
    
    'Disabling automatic links update for Microsoft Office products
    
    ' --- Microsoft Excel ---
    
    'Office 2007 +
    '[HKEY_CURRENT_USER\Software\Microsoft\Office\<version>\Excel\Security]
    'WorkbookLinkWarnings(DWORD) = 2
    
    'Office 2007 -> 12.0
    'Office 2010 -> 14.0
    'Office 2013 -> 15.0
    'Office 2016 -> 16.0
    
    ' --- Microsoft Outlook ---
    
    'Office 2010 +
    '[HKEY_CURRENT_USER\Software\Microsoft\Office\<version>\Word\Options\WordMail]
    'DontUpdateLinks(DWORD) = 1
    
    'Office 2007
    '[HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Options\vpref]
    'fNoCalclinksOnopen_90_1(DWORD) = 1
    
    'Office 2010 -> 14.0
    'Office 2013 -> 15.0
    'Office 2016 -> 16.0
    
    ' --- Microsoft Word / Microsoft Publisher ---
    
    'Office 2010 +
    '[HKEY_CURRENT_USER\Software\Microsoft\Office\<version>\Word\Options]
    'DontUpdateLinks(DWORD) = 1
    
    'Office 2007
    '[HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Options\vpref]
    'fNoCalclinksOnopen_90_1(DWORD) = 1
    
    '---------------------
    'Note: protection is included in Windows Defender for Windows 10 Fall Creator Update (version 1709) -> as Attack Surface Reduction (ASR) technology.
    
End Sub


' *--------------------------------------------------------------------------------------------*
'                         Performance & System cleaning
' *--------------------------------------------------------------------------------------------*

Private Sub ChkEnablePageFile_Click()
    ' Check pagefile state
End Sub

Private Sub ChkEnableReadyBoot_Click()
    '
    ' Enable ReadyBoot: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\WMI\Autologger\ReadyBoot
    '
End Sub

Private Sub ChkEnableSuperfetch_Click()
    ' Enable Superfetch service
    
End Sub

Private Sub ChkSetServiceDelayed_Click()
    ' Put some services on delayed (automatic) start.
    '
End Sub



' *--------------------------------------------------------------------------------------------*
'                                     Anti-Spy
' *--------------------------------------------------------------------------------------------*

Private Sub ChkSpyDomains_Click()
    '
End Sub

Private Sub ChkSpyTasks_Click()
    '
End Sub

Private Sub ChkGetWin10_Click()
    '
End Sub

Private Sub ChkSpyUpdates_Click()
    '
End Sub

Private Sub ChkTelemetry_Click()
    '
End Sub

Private Sub chkDisablePerfMeasure_Click()
    ' Disable task for starting WinSAT.exe

End Sub

Private Sub chkDisableWPAD_Click()
    '
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Wpad
    'WpadOverride = 1
    '
    
End Sub


'------------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
End Sub



'--------------------------------------------------------------------------------------------------
'
' M E N U
'

Private Sub mnuHelpAbout_Click()    'Help -> About
    '
End Sub

Private Sub mnuOptReset_Click()     'Options -> Reset to defaults
    '
End Sub

Private Sub mnuOptRecom_Click()     'Options -> Set recommended protection
    '
End Sub

Private Sub mnuOptMax_Click()       'Options -> Set maximum protection
    '
End Sub


' // TODOs:
'
'For performance
'
' Suggest Windows ADK diagnostics (?) + series of rebooting (?)

