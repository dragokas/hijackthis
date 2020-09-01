VERSION 5.00
Begin VB.Form frmCheckDigiSign 
   Caption         =   "Digital signature checker"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9252
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheckDigiSign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   9252
   Begin VB.CommandButton cmdSelectFile 
      Caption         =   "Add file(s) ..."
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
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
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optPlainText 
         Caption         =   "Plain Text (Short log in Unicode)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
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
         Left            =   1680
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
         Width           =   1455
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
         Value           =   1  'Checked
         Width           =   4815
      End
      Begin VB.TextBox txtExtensions 
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Text            =   ".exe;.dll;.sys"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
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
   Begin VB.TextBox txtPaths 
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
      Top             =   120
      Width           =   6555
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCheckDigiSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmCheckDigiSign.frm]

'
' Digital Signature checker by Alex Dragokas
'

Option Explicit

Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function SfcIsFileProtected Lib "Sfc.dll" (ByVal RpcHandle As Long, ByVal ProtFileName As Long) As Long
Private Declare Function SetWindowTheme Lib "UxTheme.dll" (ByVal hwnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long

Private Const CERT_E_UNTRUSTEDROOT          As Long = &H800B0109
Private Const TRUST_E_NOSIGNATURE           As Long = &H800B0100
Private Const CRYPT_E_BAD_MSG               As Long = &H8009200D

Dim isRan As Boolean

Private Sub cmdGo_Click()
    On Error GoTo ErrorHandler:

    Dim sPathes         As String
    Dim aPathes()       As String
    Dim vPath, vKey
    Dim bRecursively    As Boolean
    Dim bListSystemPath As Boolean
    Dim ReportPath      As String
    Dim arrTmp()        As String
    Dim i               As Long
    Dim SignResult      As SignResult_TYPE
    Dim LastSignResult  As SignResult_TYPE
    Dim hFile           As Long
    Dim sFile           As String
    Dim pos             As Long
    Dim bCSV            As Boolean
    Dim bPlainText      As Boolean
    Dim oDictFiles      As Object
    Dim OriginPath      As String
    Dim SFCFiles()      As String
    Dim sTmp            As String
    Dim cnt             As Long
    Dim hResult         As Long
    
    Static isInit       As Boolean
    Static oDictSFC     As Object
    
    '// TODO: add checkbox 'Revocation checking' (warn. about: require internet connection)
    'Add date certificate added to store (look at CERT_DATE_STAMP_PROP_ID flag of CertGetCertificateContextProperty)
    
    If isRan Then Exit Sub
    
    Set oDictFiles = New clsTrickHashTable  'CreateObject("Scripting.Dictionary")
    Set oDictSFC = New clsTrickHashTable  'CreateObject("Scripting.Dictionary")
    oDictFiles.CompareMode = vbTextCompare
    oDictSFC.CompareMode = vbTextCompare
    
    'Get options
    bRecursively = (chkRecur.Value = 1)
    bListSystemPath = (chkIncludeSys.Value = 1) 'System32 / SysWow64
    bCSV = OptCSV.Value               'CSV (in ANSI)
    bPlainText = optPlainText.Value   'Plain (in Unicode)
    
    sPathes = txtPaths.Text
    
    If sPathes = "" And Not bListSystemPath Then
        'You should enter at least one path to file or folder!
        MsgBoxW Translate(1859), vbExclamation
        Exit Sub
    End If
    
    sPathes = Replace$(sPathes, vbCr, "")
    aPathes = Split(sPathes, vbLf)
    
    ReportPath = BuildPath(App.Path(), "DigiSign") & IIf(bCSV, ".csv", ".log")
    
    If FileExists(ReportPath) Then Call DeleteFileW(StrPtr(ReportPath))
    
    'Searching files / folders
    'Enumerating files. Please wait...
    lblStatus.ForeColor = vbBlack
    lblStatus.Caption = Translate(1867)
    lblStatus.Visible = True
    DoEvents
    
    txtPaths.Enabled = False
    
    For Each vPath In aPathes
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
            arrTmp = ListFiles(CStr(vPath), IIf(OptAllFiles.Value, "", txtExtensions.Text), bRecursively)
            CopyArrayToDictionary arrTmp, oDictFiles
            DoEvents
        Else
            If InStr(vPath, " ") <> 0 Then  'dirty path? - contain arguments? - Try to remove.
                vPath = Left$(vPath, InStr(vPath, " ") - 1)
            ElseIf InStr(vPath, "/") <> 0 Then
                OriginPath = vPath
                vPath = RTrim$(Left$(vPath, InStr(vPath, "/") - 1))     'res://C:\PROGRA~1\MICROS~3\Office15\EXCEL.EXE/3000
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
        arrTmp = ListFiles(sWinDir, IIf(OptAllFiles.Value, "", txtExtensions.Text), bRecursively)
        DoEvents
        CopyArrayToDictionary arrTmp, oDictFiles
    End If
    
    lblStatus.ForeColor = vbYellow
    lblStatus.Caption = ""
    lblStatus.Visible = False
    
    ' Checking digital signature
    
    'No files found.
    If oDictFiles.Count = 0 Then
        MsgBoxW Translate(1860)
        txtPaths.Enabled = True
        Exit Sub
    End If
    
    'Abort
    cmdExit.Caption = Translate(1861)
    
    Dim bWHQL As Boolean
    Dim bWPF As Boolean
    Dim bIsDriver As Boolean
    Dim aLogLine() As String
    Dim sLogLine As String
    Dim bData() As Byte
    Dim bPE_File As Boolean
    Dim AddFlags As Long
    
    If oDictFiles.Count > 100 Then
        AddFlags = SV_EnableHashPrecache
        
        DoEvents
        'Precaching security catalogues ...
        lblStatus.ForeColor = vbBlack
        lblStatus.Visible = True
        lblStatus.Caption = Translate(1871)
        Me.Refresh
        SignVerify "", SV_EnableHashPrecache, SignResult
        lblStatus.ForeColor = vbYellow
    End If
    
    lblStatus.Caption = ""
    lblStatus.Visible = True
    shpBack.Visible = True
    
    If Not isInit Then
        isInit = True
        If bIsWinVistaAndNewer Then
            SFCFiles = SFCList_Vista()
        Else
            SFCFiles = SFCList_XP()
        End If
        For i = 0 To UBound(SFCFiles)
            If Not oDictSFC.Exists(SFCFiles(i)) Then oDictSFC.Add SFCFiles(i), 0
        Next
    End If
    
    shpFore.Visible = True
    
    isRan = True
    cmdGo.Enabled = False
    
    ErrReport = ""
    
    ReDim aLogLine(oDictFiles.Count - 1)
    
    i = 0
    For Each vKey In oDictFiles.Keys
        
        sFile = vKey
        DoEvents
        
        'досрочное прерывание работы программы
        If isRan = False Then
            CloseW hFile, True
            cmdGo.Enabled = True
            lblStatus.Caption = ""
            lblStatus.Visible = False
            shpBack.Visible = False
            shpFore.Visible = False
            txtPaths.Enabled = True
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
        
        bPE_File = isPE(sFile)
        
        'If bPE_File Then
            If StrComp(GetExtensionName(sFile), ".sys", 1) = 0 Then
            
                'Signature of driver can consist of both:
                ' - signature in catalogue (3-d party + MS)
                ' - internal signature (3-d party + MS)
                
                'So to check for WHQL and for legit 3d-party signature, you need to:
                '1) check by catalogue first by passing SV_isDriver flag, so SignVerify will use DRIVER_ACTION_VERIFY provider and return result in .IsWHQL,
                '   if found legit Microsoft signature
            
                'check WHQL mainly by the catalog
                SignVerify sFile, SV_isDriver Or SV_CheckEmbeddedPresence Or AddFlags, SignResult
                'save the state
                bWHQL = SignResult.isWHQL
'                bWHQL = SignResult.isLegit

                bIsDriver = True

                '2) check 3d-party signature with forcing WINTRUST_ACTION_GENERIC_VERIFY_V2 policy because in case driver has no corresponding
                '   Microsoft signature, WinVerifyTrust + DRIVER_ACTION_VERIFY will return CERT_E_UNTRUSTEDROOT

                'Since Microsoft SignTool also using WINTRUST_ACTION_GENERIC_VERIFY_V2, I commented this code and revoked SV_DefaultVerifyPolicy flag
'                'next we are checking for legit in usual way
'                SignVerify sFile, SV_isDriver Or SV_PreferInternalSign Or SV_DefaultVerifyPolicy Or SV_CacheDoNotLoad Or AddFlags, SignResult
'                If bWHQL Then
'                    SignResult.isWHQL = True
'                End If

                'If previous check identified Microsoft signature and internal signature is not verified yet, we need to check it for 3d-party publisher
                If SignResult.isMicrosoftSign And SignResult.isSignedByCert And SignResult.IsEmbedded Then
                    LastSignResult = SignResult
                    hResult = SignVerify(sFile, SV_isDriver Or SV_PreferInternalSign Or SV_CacheDoNotLoad Or AddFlags, SignResult)
                    SignResult.isWHQL = True
                    'For some reason "termdd.sys" has broken internal signature in XP
                    If hResult = CRYPT_E_BAD_MSG Then
                        SignResult = LastSignResult
                    End If
                End If
                
                'SignVerify returns CERT_E_UNTRUSTEDROOT if we check 3d-party signature that has no corresponding Microsoft signature
'                If CERT_E_UNTRUSTEDROOT = SignResult.ReturnCode Then
'                    bWHQL = False
'                    SignVerify sFile, SV_isDriver Or SV_CacheDoNotLoad Or AddFlags, SignResult
'                End If
            Else
                SignVerify sFile, SV_CheckEmbeddedPresence Or AddFlags, SignResult
            End If
        'End If
        
        With SignResult
            If Not bCSV Then
                aLogLine(i) = _
                    IIf(bPE_File, "", "[not PE File] ") & _
                    IIf(bIsDriver, IIf(bWHQL, "[OK] ", ""), IIf(.isLegit, "[OK] ", "")) & _
                    IIf(.ShortMessage = "TRUST_E_NOSIGNATURE: Not signed", "[NoSign] ", IIf(.ShortMessage = "Legit signature.", "", "[" & .ShortMessage & "] ")) & _
                    IIf(.isMicrosoftSign, "[MS] ", "") & _
                    sFile & " - " & _
                    IIf(bIsDriver, IIf(bWHQL, "legit.", IIf(.isLegit, "legit, but not WHQL", "")), IIf(.isLegit, "legit.", "")) & _
                    IIf(.isMicrosoftSign, " (Microsoft)", "") & _
                    IIf(bWPF, " (protected)", "")
            Else
                aLogLine(i) = GetFileNameAndExt(sFile)   'File
                aLogLine(i) = aLogLine(i) & ";" & sFile  'FullPath
                aLogLine(i) = aLogLine(i) & ";" & IIf(.isLegit, "Legit.", "no") 'Legitimate?
                aLogLine(i) = aLogLine(i) & ";" & IIf(bWHQL, "WHQL", "no")  'WQHL
                aLogLine(i) = aLogLine(i) & ";" & IIf(.isMicrosoftSign, "Microsoft", "no")  'IsMicrosoft
                aLogLine(i) = aLogLine(i) & ";" & IIf(bWPF, "protected", "no")  'WPF / SFC
                aLogLine(i) = aLogLine(i) & ";" & IIf(bPE_File, "PE", "no") 'PE
                aLogLine(i) = aLogLine(i) & ";" & .Issuer
                aLogLine(i) = aLogLine(i) & ";" & .SubjectName
                aLogLine(i) = aLogLine(i) & ";" & .SubjectEmail
                aLogLine(i) = aLogLine(i) & ";" & IIf(.ReturnCode = TRUST_E_NOSIGNATURE, "", IIf(.isSignedByCert, "Certificate", "Internal")) 'Embedded Sign?
                aLogLine(i) = aLogLine(i) & ";" & IIf(.IsEmbedded, "yes", "no")
                aLogLine(i) = aLogLine(i) & ";" & .CatalogPath
                aLogLine(i) = aLogLine(i) & ";" & .HashRootCert
                aLogLine(i) = aLogLine(i) & ";" & .HashFileCode
                aLogLine(i) = aLogLine(i) & ";" & .AlgorithmCertHash
                aLogLine(i) = aLogLine(i) & ";" & .AlgorithmSignDigest
                aLogLine(i) = aLogLine(i) & ";" & .ReturnCode
                aLogLine(i) = aLogLine(i) & ";" & .ShortMessage
                aLogLine(i) = aLogLine(i) & ";" & .FullMessage
                aLogLine(i) = aLogLine(i) & ";" & IIf(.DateTimeStamp = #12:00:00 AM#, "", Format$(.DateTimeStamp, "yyyy\/MM\/dd HH:nn:ss"))
            End If
        End With

        i = i + 1
    Next
    
    'Preparing a report. Please wait...
    lblStatus.ForeColor = vbBlack
    lblStatus.Caption = Translate(1868)
    DoEvents
    
    QuickSort aLogLine, 0, UBound(aLogLine)
    sLogLine = Join(aLogLine, vbCrLf) & IIf(Len(ErrReport) <> 0, vbCrLf & vbCrLf & "There are some errors while verification:" & vbCrLf & ErrReport, "")
    
    If bCSV Then
        sTmp = "File name"
        sTmp = sTmp & ";" & "Full path"
        sTmp = sTmp & ";" & "Legitimate?"
        sTmp = sTmp & ";" & "WQHL"
        sTmp = sTmp & ";" & "Microsoft signature?"
        sTmp = sTmp & ";" & "WPF / SFC"
        sTmp = sTmp & ";" & "is PE"
        sTmp = sTmp & ";" & "Issuer"
        sTmp = sTmp & ";" & "Signer name"
        sTmp = sTmp & ";" & "Signer email"
        sTmp = sTmp & ";" & "Signature location"
        sTmp = sTmp & ";" & "Has internal signature?"
        sTmp = sTmp & ";" & "Catalog path"
        sTmp = sTmp & ";" & "Hash of root certificate"
        sTmp = sTmp & ";" & "PE hash"
        sTmp = sTmp & ";" & "Algorithm of certificate hash"
        sTmp = sTmp & ";" & "Algorithm of signature digest"
        sTmp = sTmp & ";" & "Result code"
        sTmp = sTmp & ";" & "Result message (short)"
        sTmp = sTmp & ";" & "Result message (full)"
        sTmp = sTmp & ";" & "Time Stamp"
        
        sLogLine = sTmp & vbCrLf & sLogLine
    Else
        sLogLine = ChrW$(-257) & "Logfile of Digital Signature Checker (HJT v." & AppVerString & ")" & vbCrLf & vbCrLf & _
            MakeLogHeader() & vbCrLf & _
            "Is legitimate | FileName | Is Microsoft | Is WFP" & vbCrLf & _
            "------------------------------------------------" & vbCrLf & _
            sLogLine & vbCrLf & vbCrLf & _
            "--" & vbCrLf & "End of file"
    End If
    
    If bCSV Then
        bData() = StrConv(sLogLine, vbFromUnicode)
    Else
        bData() = sLogLine
    End If
    
    lblStatus.Visible = False
    shpFore.Visible = False
    shpBack.Visible = False
    
    isRan = False
    cmdGo.Enabled = True
    txtPaths.Enabled = True
    cmdExit.Caption = Translate(1858)
    
    If OpenW(ReportPath, FOR_OVERWRITE_CREATE, hFile, g_FileBackupFlag) Then
        PutW hFile, 1&, VarPtr(bData(0)), UBound(bData) + 1, doAppend:=True
        CloseW hFile, True
    Else
        If hFile <= 0 Then
            'Cannot open report file. Write access is restricted by another program.
            MsgBoxW Translate(1869) & vbCrLf & vbCrLf & ReportPath
            Exit Sub
        End If
    End If
    
    'rundll32.exe shell32.dll,ShellExec_RunDLL / or notepad
    OpenLogFile ReportPath
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmCheckDigiSign.cmdGo_Click"
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
    If Not inIDE Then
        isRan = False
        cmdGo.Enabled = True
    End If
    CloseW hFile, True
End Sub

Sub CopyArrayToDictionary(arr() As String, oDict As Object)
    If 0 = AryItems(arr) Then Exit Sub
    Dim i As Long
    For i = 0 To UBound(arr)
        If Not oDict.Exists(arr(i)) Then
            oDict.Add arr(i), 0
        End If
    Next
End Sub

Private Sub cmdExit_Click()
    If isRan Then
        isRan = False
        ToggleWow64FSRedirection True
        cmdExit.Caption = Translate(1858)
    Else
        Me.Hide
        'Unload Me
    End If
End Sub

Private Sub cmdSelectFile_Click()
    Dim aFile() As String
    Dim i As Long
    Dim sExt As String
    sExt = "*.exe;*.msi;*.dll;*.sys;*.ocx"
    'PE; All files
    For i = 1 To OpenFileDialog_Multi(aFile, Translate(122), Desktop, "PE (" & sExt & ")|" & sExt & "|" & Translate(1003) & " (*.*)|*.*", Me.hwnd)
        txtPaths.Text = txtPaths.Text & vbCrLf & aFile(i)
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cmdExit_Click
    ProcessHotkey KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim OptB As OptionButton
    Dim Ctl As Control
    
    LoadWindowPos Me, SETTINGS_SECTION_SIGNCHECKER
    
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    Call ReloadLanguage(True)
    
    ' if Win XP -> disable all window styles from option buttons
    If bIsWinXP Then
        For Each Ctl In Me.Controls
            If TypeName(Ctl) = "OptionButton" Then
                Set OptB = Ctl
                SetWindowTheme OptB.hwnd, StrPtr(" "), StrPtr(" ")
            End If
        Next
        Set OptB = Nothing
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveWindowPos Me, SETTINGS_SECTION_SIGNCHECKER
    
    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        If isRan Then
            isRan = False
            ToggleWow64FSRedirection True
            Cancel = True
            Me.Hide
        Else
            Cancel = True
            Me.Hide
        End If
    End If
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    
    Dim TopLevel1&, TopLevel2&
    
    If Me.Width < 8350 Then Me.Width = 8350
    If Me.Height < 3650 Then Me.Height = 3650
    
    txtPaths.Width = Me.Width - 630
    txtPaths.Height = Me.Height - 3500
    
    TopLevel1 = Me.Height - 1300
    TopLevel2 = TopLevel1 - 1440
    
    cmdGo.Top = TopLevel1
    cmdExit.Top = TopLevel1
    
    shpBack.Top = TopLevel1 + 120
    shpFore.Top = TopLevel1 + 120
    lblStatus.Top = TopLevel1 + 210
    shpBack.Width = Me.Width - 4680
    
    fraFilter.Top = TopLevel2
    fraReportFormat.Top = TopLevel2
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cmdExit_Click
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AddObjToList Data
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AddObjToList Data
End Sub

Private Sub AddObjToList(Data As DataObject)
    Const vbCFFiles As Long = 15&
    Dim vObj
    If Data.GetFormat(vbCFFiles) Then
        For Each vObj In Data.Files
            txtPaths.Text = txtPaths.Text & IIf(Right$(txtPaths.Text, 2) <> vbCrLf And Len(txtPaths.Text) > 0, vbCrLf, "") & CStr(vObj)
        Next
    End If
End Sub


