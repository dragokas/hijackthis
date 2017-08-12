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
         Width           =   3495
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
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OptAllFiles 
         Caption         =   "All Files"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
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
         Left            =   3000
         TabIndex        =   6
         Text            =   ".exe;.dll;.sys"
         Top             =   240
         Width           =   1935
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
Private Declare Function SetWindowTheme Lib "UxTheme.dll" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long

Private Const CERT_E_UNTRUSTEDROOT          As Long = &H800B0109

Dim isRan As Boolean

Private Sub cmdGo_Click()
    On Error GoTo ErrorHandler:

    Dim sPathes         As String
    Dim aPathes, vPath
    Dim bRecursively    As Boolean
    Dim bListSystemPath As Boolean
    Dim ReportPath      As String
    Dim arrTmp()        As String
    Dim i               As Long
    Dim SignResult      As SignResult_TYPE
    Dim SignResultTemp  As SignResult_TYPE
    Dim hFile           As Long
    Dim sFile           As String
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
    
    Set oDictFiles = New clsTrickHashTable
    Set oDictSFC = New clsTrickHashTable
    oDictFiles.CompareMode = TextCompare
    oDictSFC.CompareMode = TextCompare
    
    'Get options
    bRecursively = (chkRecur.value = 1)
    bListSystemPath = (chkIncludeSys.value = 1) 'System32 / SysWow64
    bCSV = OptCSV.value               'CSV (in ANSI)
    bPlainText = optPlainText.value   'Plain (in Unicode)
    
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
    
        'If InStr(1, vPath, "Wudfh", 1) <> 0 Then Stop
    
        vPath = Trim$(vPath)
        If Left$(vPath, 1) = """" Then
            pos = InStr(2, vPath, """")
            If pos <> 0 Then
                vPath = Mid$(vPath, 2, pos - 2)
            Else
                vPath = Mid$(vPath, 2)
            End If
        End If
        
        vPath = Replace$(vPath, "\\", "\")
        If Mid$(vPath, 2, 1) <> ":" Then
            'try to remove some remnants from beginning of line
            pos = InStr(vPath, ":\")
            If pos > 1 Then
                vPath = Mid$(vPath, pos - 1)
            End If
        End If
        
        If FileExists(CStr(vPath)) Then
            If Not oDictFiles.Exists(vPath) Then oDictFiles.Add vPath, 0
        ElseIf FolderExists(CStr(vPath)) Then
            arrTmp = ListFiles(CStr(vPath), IIf(OptAllFiles.value = 1, "", txtExtensions.Text), bRecursively)
            CopyArrayToDictionary arrTmp, oDictFiles
            DoEvents
        Else
            If InStr(vPath, " ") <> 0 Then  'dirty path? - contain arguments? - Try to remove.
                vPath = Left$(vPath, InStr(vPath, " ") - 1)
            ElseIf InStr(vPath, "/") <> 0 Then
                OriginPath = vPath
                vPath = RTrim(Left$(vPath, InStr(vPath, "/") - 1))     'res://C:\PROGRA~1\MICROS~3\Office15\EXCEL.EXE/3000
            End If

            If FileExists(CStr(vPath)) Then
                If Not oDictFiles.Exists(vPath) Then oDictFiles.Add vPath, 0
            Else
                If InStr(OriginPath, "/") <> 0 Then
                    vPath = Replace$(OriginPath, "/", "\")      'path altered by / chars instead of \
                    vPath = Replace$(vPath, "\\", "\")
                            
                    If FileExists(CStr(vPath)) Then
                        If Not oDictFiles.Exists(vPath) Then oDictFiles.Add vPath, 0
                    End If
                End If
            End If
        End If
    Next
    
    If bListSystemPath Then
        'default - .exe;.dll;.sys
        arrTmp = ListFiles(sWinSysDir, IIf(OptAllFiles.value = 1, "", txtExtensions.Text), bRecursively)
        DoEvents
        CopyArrayToDictionary arrTmp, oDictFiles
        
        If bIsWin64 Then
            arrTmp = ListFiles(sWinSysDirWow64, IIf(OptAllFiles.value = 1, "", txtExtensions.Text), bRecursively)
            DoEvents
            CopyArrayToDictionary arrTmp, oDictFiles
        End If
    End If
    
    ' Checking digital signature
    
    'No files found.
    If oDictFiles.Count = 0 Then MsgBoxW Translate(1860): Exit Sub
    
    'Abort
    cmdExit.Caption = Translate(1861)
    
    Dim bWHQL As Boolean
    Dim bWPF As Boolean
    Dim bIsDriver As Boolean
    Dim IsMicrosoftFile As Boolean
    Dim aLogLine() As String
    Dim sLogLine As String
    Dim bData() As Byte
    Dim bPE_EXE As Boolean
    Dim bSkipErrorOld As Boolean
    
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
    
    bSkipErrorOld = bSkipErrorMsg
    bSkipErrorMsg = True
    ErrReport = ""
    
    ReDim aLogLine(oDictFiles.Count - 1)
    
    For i = 0 To oDictFiles.Count - 1
    
        sFile = oDictFiles.Keys(i)
    
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
        lblStatus = (i + 1) & " / " & oDictFiles.Count & " - " & GetFileNameAndExt(sFile) & " - " & GetPathName(sFile)
        If i Mod 10 = 0 Then
            shpFore.Width = CLng((shpBack.Width / oDictFiles.Count) * i)
        End If
        
        'bWPF = inArray(arrFiles(i), SFCFiles, , , vbTextCompare)
        bWPF = oDictSFC.Exists(sFile)
        If Not bWPF Then bWPF = SfcIsFileProtected(0&, StrPtr(sFile))
        
        'bPE_EXE = isPE_EXE(arrFiles(i))
        'maybe replace by GetBinaryType API ?
        
        'If bPE_EXE Then
            If StrComp(GetExtensionName(sFile), ".sys", 1) = 0 Then
                SignVerify sFile, SV_isDriver, SignResult
                bWHQL = SignResult.isLegit
                bIsDriver = True
                If CERT_E_UNTRUSTEDROOT = SignResult.ReturnCode Then
                    SignVerify sFile, SV_CacheDoNotLoad, SignResult
                End If
            Else
                SignVerify sFile, 0&, SignResult
            End If
            
            IsMicrosoftFile = IsMicrosoftCertHash(SignResult.HashRootCert) 'And SignResult.isLegit
        'End If
        
        With SignResult
            If bCSV Then
                aLogLine(i) = _
                    GetFileNameAndExt(sFile) & ";" & _
                    sFile & ";" & _
                    IIf(.isLegit, "Yes", "No") & ";" & _
                    IIf(bWHQL, "Yes", "No") & ";" & _
                    IIf(IsMicrosoftFile, "Yes", "No") & ";" & _
                    IIf(bWPF, "Yes", "No") & ";" & _
                    IIf(.ReturnCode = -2146762496, "", IIf(Not .isSignedByCert, "Yes", "No")) & ";" & _
                    .Issuer & ";" & _
                    .HashRootCert & ";" & _
                    .ReturnCode & ";" & _
                    .ShortMessage & ";" & _
                    .FullMessage
            Else
                'IIf(bPE_EXE, "[not PE EXE] ", "")
                aLogLine(i) = _
                    IIf(bIsDriver, IIf(bWHQL, "[OK] ", ""), IIf(.isLegit, "[OK] ", "")) & _
                    IIf(.ShortMessage = "TRUST_E_NOSIGNATURE: Not signed", "[NoSign] ", IIf(.ShortMessage = "Legit signature.", "", "[" & .ShortMessage & "] ")) & _
                    IIf(IsMicrosoftFile, "[MS] ", "") & _
                    sFile & " - " & _
                    IIf(bIsDriver, IIf(bWHQL, "legit.", ""), IIf(.isLegit, "legit.", "")) & _
                    IIf(IsMicrosoftFile, " (Microsoft)", "") & _
                    IIf(bWPF, " (protected)", "")
            End If
        End With
    Next
    
    QuickSort aLogLine, 0, UBound(aLogLine)
    sLogLine = Join(aLogLine, vbCrLf) & IIf(Len(ErrReport) <> 0, vbCrLf & vbCrLf & "There are some errors while verification:" & vbCrLf & ErrReport, "")
    
    If bCSV Then
        sLogLine = "File;FullPath;Legitimate;WQHL;IsMicrosoft;WPF / SFC;Embedded Sign;Issuer;RootCertHash;ErrCode;ErrMsgShort;ErrMsgFull" & vbCrLf & _
            sLogLine
    Else
        sLogLine = ChrW$(-257) & _
            "FileName - Is legitimate - (Microsoft, WFP)" & vbCrLf & _
            "-------------------------------------------" & vbCrLf & _
            sLogLine
    End If
    
    If bCSV Then
        bData() = StrConv(sLogLine, vbFromUnicode)
    Else
        bData() = sLogLine
    End If
    
    bSkipErrorMsg = bSkipErrorOld
    
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
    ErrorMsg Err, "frmCheckDigiSign.cmdGo_Click"
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
    If Not inIDE Then
        isRan = False
        cmdGo.Enabled = True
        CloseW hFile
    End If
End Sub

Sub CopyArrayToDictionary(arr() As String, oDict As Object)
    If Not IsArrDimmed(arr) Then Exit Sub
    Dim i As Long
    For i = 0 To UBound(arr)
        If Not oDict.Exists(arr(i)) Then
            oDict.Add arr(i), 0
        End If
    Next
End Sub

Private Sub CmdExit_Click()
    If isRan Then
        isRan = False
        ToggleWow64FSRedirection True
        'Unload Me
        'Close
        cmdExit.Caption = Translate(1858)
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dim OptB As OptionButton
    Dim ctl As Control
    
    ReloadLanguage
    CenterForm Me
    Me.Icon = frmMain.Icon
    
    ' if Win XP -> disable all window styles from option buttons
    If bIsWinXP Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "OptionButton" Then
                Set OptB = ctl
                SetWindowTheme OptB.hWnd, StrPtr(" "), StrPtr(" ")
            End If
        Next
        Set OptB = Nothing
    End If
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
  
    Dim TopLevel1&, TopLevel2&
    
    If Me.Width < 8350 Then Me.Width = 8350
    If Me.Height < 3650 Then Me.Height = 3650
    
    Text1.Width = Me.Width - 630
    Text1.Height = Me.Height - 3500
    
    TopLevel1 = Me.Height - 1300
    TopLevel2 = TopLevel1 - 1440
    
    cmdGo.Top = TopLevel1
    cmdExit.Top = TopLevel1
    
    shpBack.Top = TopLevel1 + 120
    shpFore.Top = TopLevel1 + 120
    lblStatus.Top = TopLevel1 + 210
    
    fraFilter.Top = TopLevel2
    fraReportFormat.Top = TopLevel2
    
    'chkIncludeSys.Top = TopLevel2
    'optPlainText.Top = TopLevel3 + 230
    
    'chkRecur.Top = TopLevel3
    'lblFormat.Top = TopLevel3
    'OptCSV.Top = TopLevel3 + 230 + 360
End Sub

