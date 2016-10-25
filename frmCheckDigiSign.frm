VERSION 5.00
Begin VB.Form frmCheckDigiSign 
   Caption         =   "Digital signature checker"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255
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
   ScaleHeight     =   4680
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraReportFormat 
      Caption         =   "Report format:"
      Height          =   1335
      Left            =   5400
      TabIndex        =   9
      Top             =   2640
      Width           =   3735
      Begin VB.OptionButton OptCSV 
         Caption         =   "CSV (Full log in ANSI)"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton optPlainText 
         Caption         =   "Plain Text (Short log in Unicode)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   3135
      End
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filter"
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   5055
      Begin VB.OptionButton OptExtension 
         Caption         =   "by extension"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptAllFiles 
         Caption         =   "All Files"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkIncludeSys 
         Caption         =   "Include files in Windows\System32 (SysWOW64) folder"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   4815
      End
      Begin VB.CheckBox chkRecur 
         Caption         =   "Recursively (include subfolders)"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtExtensions 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Text            =   ".exe;.dll;.sys"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Close"
      Height          =   480
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   480
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 / 1 - File - (in folder)"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape shpFore 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   4320
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   4320
      Top             =   4200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblThisTool 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This tool will create a detail report about digital signature of files/folders you specify below:"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   6555
   End
End
Attribute VB_Name = "frmCheckDigiSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Run on Ctrl + F5 only !!!


Private Declare Function SfcIsFileProtected Lib "Sfc.dll" (ByVal RpcHandle As Long, ByVal ProtFileName As Long) As Long


Dim isRan As Boolean

Private Sub cmdGo_Click()
    On Error GoTo ErrorHandler:

    Dim sPathes         As String
    Dim aPathes, vPath
    Dim bRecursively    As Boolean
    Dim bListSystemPath As Boolean
    Dim ReportPath      As String
    Dim arrFiles()      As String
    Dim arrTmp()        As String
    Dim i               As Long
    Dim SignResult      As SignResult_TYPE
    Dim SignResultTemp  As SignResult_TYPE
    Dim hFile           As Long
    Dim pos             As Long
    Dim bCSV            As Boolean
    Dim bPlainText      As Boolean
    Dim oDictFiles      As clsTrickHashTable
    Dim OriginPath      As String
    Dim SFCFiles()      As String
    
    Static IsInit       As Boolean
    Static oDictSFC     As clsTrickHashTable
    
    '// TODO: add checkbox 'Revocation checking' (warn. about: require internet connection)
    'Add issued to (name / email)
    'Add TimeStamp valid from / until (look at CERT_INFO structure)
    'Add date certificate added to store (look at CERT_DATE_STAMP_PROP_ID flag of CertGetCertificateContextProperty)
    'Add catalogue phisical location (look at CryptCATAdminResolveCatalogPath)
    
    If isRan Then Exit Sub
    
    ReDim arrFiles(0)
    Set oDictFiles = New clsTrickHashTable
    Set oDictSFC = New clsTrickHashTable
    oDictFiles.CompareMode = TextCompare
    oDictSFC.CompareMode = TextCompare
    
    'Get options
    bRecursively = (chkRecur.Value = 1)
    bListSystemPath = (chkIncludeSys.Value = 1) 'System32 / SysWow64
    bCSV = OptCSV.Value               'CSV (in ANSI)
    bPlainText = optPlainText.Value   'Plain (in Unicode)
    
    sPathes = Text1.Text
    
    If sPathes = "" And Not bListSystemPath Then
        'You should enter at least one path to file or folder!
        MsgBoxW Translate(1859), vbExclamation
        Exit Sub
    End If
    
    sPathes = Replace$(sPathes, vbCr, "")
    aPathes = Split(sPathes, vbLf)
    
    ReportPath = BuildPath(AppPath(), "DigiSign") & IIf(bCSV, ".csv", ".log")
    
    If FileExists(ReportPath) Then Call DeleteFileW(StrPtr(ReportPath))
    
    'Searching files / folders
    
    For Each vPath In aPathes
    
        vPath = UnQuote(Trim$(vPath))
        vPath = Replace$(vPath, "\\", "\")
        If Mid$(vPath, 2, 1) <> ":" Then
            'try to remove some remnants from beginning of line
            pos = InStr(vPath, ":\")
            If pos > 1 Then
                vPath = Mid$(vPath, pos - 1)
            End If
        End If
        
        If Not oDictFiles.Exists(vPath) Then
        
            oDictFiles.Add vPath, 0
        
            If FileExists(CStr(vPath)) Then
                ReDim Preserve arrFiles(UBound(arrFiles) + 1)
                arrFiles(UBound(arrFiles)) = vPath
            ElseIf FolderExists(CStr(vPath)) Then
                arrTmp = ListFiles(CStr(vPath), IIf(OptAllFiles.Value = 1, "", txtExtensions.Text), bRecursively)
                Call ConcatArrays(arrFiles, arrTmp)
                DoEvents
            Else
                If InStr(vPath, " ") <> 0 Then  'dirty path? - contain arguments? - Try to remove.
                    vPath = Left$(vPath, InStr(vPath, " ") - 1)
                ElseIf InStr(vPath, "/") <> 0 Then
                    OriginPath = vPath
                    vPath = Left$(vPath, InStr(vPath, "/") - 1)     'res://C:\PROGRA~1\MICROS~3\Office15\EXCEL.EXE/3000
                End If
            
                If Not oDictFiles.Exists(vPath) Then
                    oDictFiles.Add vPath, 0
                    
                    If FileExists(CStr(vPath)) Then
                        ReDim Preserve arrFiles(UBound(arrFiles) + 1)
                        arrFiles(UBound(arrFiles)) = vPath
                    Else
                        If InStr(OriginPath, "/") <> 0 Then
                            vPath = Replace$(OriginPath, "/", "\")      'path altered by / chars instead of \
                            vPath = Replace$(vPath, "\\", "\")
                            If Not oDictFiles.Exists(vPath) Then
                                oDictFiles.Add vPath, 0
                                If FileExists(CStr(vPath)) Then
                                    ReDim Preserve arrFiles(UBound(arrFiles) + 1)
                                    arrFiles(UBound(arrFiles)) = vPath
                                End If
                            End If
                        End If
                    End If
                End If
                'specified path is not exist
            End If
        End If
    Next
    
    If bListSystemPath Then
        'default - .exe;.dll;.sys
        arrTmp = ListFiles(sWinSysDir, IIf(OptAllFiles.Value = 1, "", txtExtensions.Text), bRecursively)
        DoEvents
        Call ConcatArrays(arrFiles, arrTmp)
        
        If bIsWin64 Then
            arrTmp = ListFiles(sWinSysDirWow64, IIf(OptAllFiles.Value = 1, "", txtExtensions.Text), bRecursively)
            DoEvents
            Call ConcatArrays(arrFiles, arrTmp)
        End If
    End If
    
    ' Checking digital signature
    
    'No files found.
    If UBound(arrFiles) = 0 Then MsgBoxW Translate(1860): Exit Sub
    
    'Abort
    CmdExit.Caption = Translate(1861)
    
    Dim bWHQL As Boolean
    Dim bWPF As Boolean
    Dim bIsDriver As Boolean
    Dim IsMicrosoftFile As Boolean
    Dim sLogLine As String
    Dim bHeader As Boolean
    Dim bData() As Byte
    Dim bPE_EXE As Boolean
    
    lblStatus.Caption = ""
    lblStatus.Visible = True
    shpBack.Visible = True
    
    If Not IsInit Then
        IsInit = True
        If OSver.bIsVistaOrLater Then
            SFCFiles = SFCList_Vista()
        Else
            SFCFiles = SFCList_XP()
        End If
        For i = 0 To UBound(SFCFiles)
            If Not oDictSFC.Exists(SFCFiles(i)) Then oDictSFC.Add SFCFiles(i), 0
        Next
    End If
    
    shpFore.Visible = True
    
    OpenW ReportPath, FOR_OVERWRITE_CREATE, hFile
    
    isRan = True
    cmdGo.Enabled = False
    
    For i = 1 To UBound(arrFiles)
        DoEvents
        If isRan = False Then
            CloseW hFile
            cmdGo.Enabled = True
            lblStatus.Caption = ""
            lblStatus.Visible = False
            shpBack.Visible = False
            shpFore.Visible = False
            Exit Sub
        End If
        bIsDriver = False
        bWHQL = False
        
        'ProgressBar
        lblStatus = i & " / " & UBound(arrFiles) & " - " & GetFileNameAndExt(arrFiles(i)) & " - " & GetPathName(arrFiles(i))
        shpFore.Width = CLng(shpBack.Width / UBound(arrFiles) * i)
        
        'bWPF = inArray(arrFiles(i), SFCFiles, , , vbTextCompare)
        bWPF = oDictSFC.Exists(arrFiles(i))
        If Not bWPF Then bWPF = SfcIsFileProtected(0&, StrPtr(arrFiles(i)))
        
        'bPE_EXE = isPE_EXE(arrFiles(i))
        'maybe replace by GetBinaryType API ?
        
        'If bPE_EXE Then
            SignVerify arrFiles(i), 0&, SignResult
            IsMicrosoftFile = IsMicrosoftCertHash(SignResult.RootCertHash) And SignResult.isLegit
        
            If StrComp(GetExtensionName(arrFiles(i)), ".sys", 1) = 0 Then
                SignVerify arrFiles(i), SV_isDriver, SignResultTemp
                bWHQL = SignResultTemp.isLegit
                bIsDriver = True
            End If
        'End If
        
        ' Log Header
        If Not bHeader Then
            bHeader = True
            If bCSV Then
                sLogLine = "File;FullPath;Legitimate;WQHL;IsMicrosoft;WPF / SFC;Embedded Sign;Issuer;RootCertHash;ErrCode;ErrMsgShort;ErrMsgFull" & vbCrLf
            Else
                sLogLine = ChrW$(-257) & _
                           "FileName - Is legitimate - (Microsoft, WFP)" & vbCrLf & _
                           "-------------------------------------------" & vbCrLf
            End If
        End If
        
        With SignResult
            If bCSV Then
                sLogLine = sLogLine & _
                    GetFileNameAndExt(arrFiles(i)) & ";" & _
                    arrFiles(i) & ";" & _
                    IIf(.isLegit, "Yes", "No") & ";" & _
                    IIf(bWHQL, "Yes", "No") & ";" & _
                    IIf(IsMicrosoftFile, "Yes", "No") & ";" & _
                    IIf(bWPF, "Yes", "No") & ";" & _
                    IIf(.ReturnCode = -2146762496, "", IIf(Not .isCert, "Yes", "No")) & ";" & _
                    .Issuer & ";" & _
                    .RootCertHash & ";" & _
                    .ReturnCode & ";" & _
                    .ShortMessage & ";" & _
                    .FullMessage & _
                    vbCrLf
                    
                bData() = StrConv(sLogLine, vbFromUnicode)
            Else
                'IIf(bPE_EXE, "[not PE EXE] ", "")
                sLogLine = sLogLine & _
                    IIf(bIsDriver, IIf(bWHQL, "[OK] ", ""), IIf(.isLegit, "[OK] ", "")) & _
                    IIf(.ShortMessage = "TRUST_E_NOSIGNATURE: Not signed", "[NoSign] ", "") & _
                    IIf(IsMicrosoftFile, "[MS] ", "") & _
                    arrFiles(i) & " - " & _
                    IIf(bIsDriver, IIf(bWHQL, "legit.", ""), IIf(.isLegit, "legit.", "")) & _
                    IIf(IsMicrosoftFile, " (Microsoft)", "") & _
                    IIf(bWPF, " (protected)", "") & _
                    vbCrLf
                    
                bData() = sLogLine
            End If
        End With
        
        'PutW hFile, 1&, VarPtr(bData(0)), UBound(bData) + 1, doAppend:=True
        'sLogLine = ""
    Next
    
    'Sort
    Dim Lines
    Lines = Split(sLogLine, vbCrLf)
    QuickSort Lines, 0, UBound(Lines)
    sLogLine = Join(Lines, vbCrLf)
    
    PutW hFile, 1&, VarPtr(bData(0)), UBound(bData) + 1, doAppend:=True
    
    lblStatus.Visible = False
    shpFore.Visible = False
    shpBack.Visible = False
    
    CloseW hFile
    
    isRan = False
    cmdGo.Enabled = True
    
    If FileExists(ReportPath) Then
        Shell "rundll32.exe shell32.dll,ShellExec_RunDLL " & """" & ReportPath & """", vbNormalFocus
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg err, "frmCheckDigiSign.cmdGo_Click"
    ToggleWow64FSRedirection True
    If inIDE Then
        Stop: Resume Next
    Else
        isRan = False
        cmdGo.Enabled = True
        CloseW hFile
    End If
End Sub

Private Sub CmdExit_Click()
    If isRan Then
        isRan = False
        ToggleWow64FSRedirection True
        'Unload Me
        'Close
        CmdExit.Caption = Translate(1858)
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    ReloadLanguage
    CenterForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        If isRan Then
            isRan = False
            ToggleWow64FSRedirection True
        Else
            Cancel = True
            Me.Hide
        End If
    End If
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Or Me.WindowState = vbMaximized Then Exit Sub
  
    Dim TopLevel1&, TopLevel2&, TopLevel3&
    
    If Me.Width < 8250 Then Me.Width = 8250
    If Me.Height < 3430 Then Me.Height = 3430
    
    Text1.Width = Me.Width - 630
    Text1.Height = Me.Height - 2850
    
    TopLevel1 = Me.Height - 1300
    TopLevel2 = TopLevel1 - 480
    TopLevel3 = TopLevel2 - 360
    
    cmdGo.Top = TopLevel1
    CmdExit.Top = TopLevel1
    
    shpBack.Top = TopLevel1 + 120
    shpFore.Top = TopLevel1 + 120
    lblStatus.Top = TopLevel1 + 210
    
    'chkIncludeSys.Top = TopLevel2
    'optPlainText.Top = TopLevel3 + 230
    
    'chkRecur.Top = TopLevel3
    'lblFormat.Top = TopLevel3
    'OptCSV.Top = TopLevel3 + 230 + 360
End Sub
