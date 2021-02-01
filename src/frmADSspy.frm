VERSION 5.00
Begin VB.Form frmADSspy 
   Caption         =   "ADS Spy [] fork - written by Merijn"
   ClientHeight    =   6768
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmADSspy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6768
   ScaleWidth      =   8340
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   120
      ScaleHeight     =   204
      ScaleWidth      =   6684
      TabIndex        =   17
      Top             =   6480
      Width           =   6735
   End
   Begin VB.Frame fraMain 
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   8175
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save log..."
         Height          =   375
         Left            =   2220
         TabIndex        =   19
         Tag             =   "1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   360
         Left            =   7080
         TabIndex        =   18
         Top             =   1680
         Width           =   990
      End
      Begin VB.CommandButton cmdScanFolder 
         Caption         =   "Choose..."
         Height          =   285
         Left            =   6840
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtScanFolder 
         Height          =   285
         Left            =   2880
         TabIndex        =   15
         Top             =   720
         Width           =   3855
      End
      Begin VB.OptionButton optScanLocation 
         Caption         =   "Scan only this folder:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton optScanLocation 
         Caption         =   "Full scan (all NTFS, ReFS drives)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   7215
      End
      Begin VB.OptionButton optScanLocation 
         Caption         =   "Quick scan (Windows base folder only)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   7215
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove selected streams"
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CheckBox chkIgnoreEncryptable 
         Caption         =   "Ignore safe system info data streams"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Value           =   1  'Checked
         Width           =   7935
      End
      Begin VB.CheckBox chkCalcMD5 
         Caption         =   "Calculate MD5 checksums of streams' contents"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   7935
      End
   End
   Begin VB.ListBox lstADSFound 
      Height          =   1605
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   4080
      Width           =   6735
   End
   Begin VB.TextBox txtUselessBlabber 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmADSspy.frx":1CFA
      Top             =   120
      Width           =   8055
   End
   Begin VB.CommandButton cmdViewSave 
      Caption         =   "Save to disk..."
      Height          =   470
      Left            =   1440
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdViewBack 
      Caption         =   "Back"
      Height          =   470
      Left            =   5160
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdViewEdit 
      Caption         =   "Edit in Wordpad"
      Height          =   470
      Left            =   3240
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdViewCopy 
      Caption         =   "Copy"
      Height          =   470
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtADSContent 
      BackColor       =   &H8000000F&
      Height          =   1815
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSelAll 
         Caption         =   "Select all"
      End
      Begin VB.Menu mnuPopupSelNone 
         Caption         =   "Select none"
      End
      Begin VB.Menu mnuPopupSelInvert 
         Caption         =   "Invert selection"
      End
      Begin VB.Menu mnuPopupStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupView 
         Caption         =   "View stream contents"
      End
      Begin VB.Menu mnuPopupShowFile 
         Caption         =   "Show file in Explorer"
      End
      Begin VB.Menu mnuPopupStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupSave 
         Caption         =   "Save scan results to disk..."
      End
   End
End
Attribute VB_Name = "frmADSspy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmADSspy.frm]

'
'ADS Spy by Merijn Bellekom
'

'Check 'frmMain.frm' to change version number

'Forked by Dragokas
'
'v.1.14
'Added ReFS support
'Added button "Save log..."
'
'v.1.13
'Partially added unicode support
'Fixed infinite loop in symlinks.
'Added x64 OS bitness support
'Added view stream contents on double click
'Fixed bug when 'cannot open file' to read stream contents
'WhiteList updated

Option Explicit

Private Type FILE_STREAM_INFORMATION
    NextEntryOffset As Long
    StreamNameLength As Long
    StreamSize As Long
    StreamSizeHi As Long
    StreamAllocationSize As Long
    StreamAllocationSizeHi As Long
    StreamName(259) As Byte
End Type

Private Type IO_STATUS_BLOCK
    IoStatus As Long
    Information As Long
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    lpszFileName(MAX_PATH - 1) As Integer
    lpszAlternate(13) As Integer
End Type

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
'Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function GetLogicalDrives Lib "kernel32.dll" () As Long
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function NtQueryInformationFile Lib "ntdll.dll" (ByVal FileHandle As Long, IoStatusBlock_Out As IO_STATUS_BLOCK, lpFileInformation_Out As Long, ByVal Length As Long, ByVal FileInformationClass As FILE_INFORMATION_CLASS) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
'Private Declare Function SHFileExists Lib "shell32.dll" Alias "#45" (ByVal szPath As String) As Long
'Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "Advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "Advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function CryptAcquireContext Lib "Advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "Advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "Advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "Advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData Lib "Advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "Advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderW" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListW" (ByVal pidList As Long, ByVal lpBuffer As Long) As Long
Private Declare Function SetWindowTheme Lib "UxTheme.dll" (ByVal hwnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStrDest As Long, ByVal lpStrSrc As Long) As Long

Private Enum FILE_INFORMATION_CLASS
    FileDirectoryInformation = 1
    FileFullDirectoryInformation   ' // 2
    FileBothDirectoryInformation   ' // 3
    FileBasicInformation           ' // 4  wdm
    FileStandardInformation        ' // 5  wdm
    FileInternalInformation        ' // 6
    FileEaInformation              ' // 74
    FileAccessInformation          ' // 8
    FileNameInformation            ' // 9
    FileRenameInformation          ' // 10
    FileLinkInformation            ' // 11
    FileNamesInformation           ' // 12
    FileDispositionInformation     ' // 13
    FilePositionInformation        ' // 14 wdm
    FileFullEaInformation          ' // 15
    FileModeInformation            ' // 16
    FileAlignmentInformation       ' // 17
    FileAllInformation             ' // 18
    FileAllocationInformation      ' // 19
    FileEndOfFileInformation       ' // 20 wdm
    FileAlternateNameInformation   ' // 21
    FileStreamInformation          ' // 22
    FilePipeInformation            ' // 23
    FilePipeLocalInformation       ' // 24
    FilePipeRemoteInformation      ' // 25
    FileMailslotQueryInformation   ' // 26
    FileMailslotSetInformation     ' // 27
    FileCompressionInformation     ' // 28
    FileObjectIdInformation        ' // 29
    FileCompletionInformation      ' // 30
    FileMoveClusterInformation     ' // 31
    FileQuotaInformation           ' // 32
    FileReparsePointInformation    ' // 33
    FileNetworkOpenInformation     ' // 34
    FileAttributeTagInformation    ' // 35
    FileTrackingInformation        ' // 36
    FileMaximumInformation
End Enum

Private Const BIF_RETURNONLYFSDIRS = 1

Private Const DRIVE_UNKNOWN = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_CLASS_HASH As Long = 32768

Private Const HP_HASHVAL As Long = 2
Private Const HP_HASHSIZE As Long = 4

Private Const CRYPT_VERIFYCONTEXT = &HF0000000

Private Const PROV_RSA_FULL As Long = 1
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

'Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_QUERY_VALUE = &H1

Private Const FILE_NAMED_STREAMS As Long = &H40000
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF
Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400&

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private bQuickScan As Boolean, bScanFolder As Boolean, bAbortScanNow As Boolean
Private bIgnoreEncryptable As Boolean, bCalcMD5 As Boolean, bQueryUnload As Boolean
Private sSafeStreams$()

Private Sub cmdExit_Click()
    bAbortScanNow = True
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    mnuPopupSave_Click
End Sub

Private Sub cmdScanFolder_Click()
    Dim sPath$, sNTFSDrives$(), i&
    'Select a folder to scan:
    sPath = BrowseForFolder(Translate(194))
    If sPath <> vbNullString Then
        sNTFSDrives = Split(GetNTFSDrives(), "|")
        For i = 0 To UBound(sNTFSDrives)
            If InStr(1, sPath, sNTFSDrives(i), vbTextCompare) = 1 Then
                'check if drive of given path is NTFS
                txtScanFolder.Text = sPath
                optScanLocation(2).Value = True
                Exit Sub
            End If
        Next i
        'The folder you have selected is not on a NTFS drive.
        MsgBoxW Translate(195), vbCritical
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then bAbortScanNow = True: Me.Hide
    ProcessHotkey KeyCode, Me
End Sub

Private Sub Form_Load()
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    ReloadLanguage True

    'ADS Spy v[] - written by Merijn
    Me.Caption = Replace$(Translate(190), "[]", ADSspyVer)
    
'    txtUselessBlabber.Text = "Alternate Data Streams (ADS) are pieces of " & _
'      "info hidden as metadata on files on NTFS drives. They are not visible in Explorer " & _
'      "and the size they take up is not reported by Windows. Recent " & _
'      "browser hijackers started using ADS to hide their files, and very " & _
'      "few anti-malware scanners detect this. Use ADS Spy to find and " & _
'      "remove these streams." & vbCrLf & _
'      "Note: this app can also display legitimate ADS streams. Don't " & _
'      "delete streams if you are not completely sure they are malicious!"
    txtUselessBlabber.Text = Translate(205)
    
    sWinDir = String$(MAX_PATH, 0)
    sWinDir = Left$(sWinDir, GetSystemWindowsDirectory(StrPtr(sWinDir), Len(sWinDir)))
    If InStr(Command$, "/debug") > 0 Then
    'If 1 Then
        Dim i%, s$()
        Me.Show
        'Listing drives...
        Status Translate(197), "1"
        s = Split(GetDrives, "|")
        lstADSFound.Clear
        'Enumerating system drives:
        lstADSFound.AddItem " " & Translate(204) & ":"
        For i = 0 To UBound(s)
            lstADSFound.AddItem s(i)
        Next i
    Else
        CheckIfSystemIsNTFS
        If InStr(1, GetNTFSDrives, Left$(sWinDir, 2), vbTextCompare) = 0 Then
            'chkQuickScan.Value = 0
            optScanLocation(0).Enabled = False
        End If
    End If
    
    ReDim sSafeStreams(8)
    sSafeStreams(0) = ":encryptable:$DATA"
    sSafeStreams(1) = ":SummaryInformation:$DATA"
    sSafeStreams(2) = ":DocumentSummaryInformation:$DATA"
    sSafeStreams(3) = ":{4c8cc155-6c1e-11d1-8e41-00c04fb9386d}:$DATA"
    sSafeStreams(4) = ":Zone.Identifier:$DATA"
    sSafeStreams(5) = ":Q30lsldxJoudresxAaaqpcawXc:$DATA"
    sSafeStreams(6) = ":Win32App_1:$DATA" 'Windows 10 "Storage Service"
    sSafeStreams(7) = ":favicon:$DATA"
    sSafeStreams(8) = ":OECustomProperty:$DATA"
    'Ready.
    Status Translate(209), "2"
    
    Dim OptB As OptionButton
    Dim Btn As CommandButton
    Dim Ctl As Control
    
    LoadWindowPos Me, SETTINGS_SECTION_ADSSPY
    
    ' if Win XP -> disable all window styles from option buttons
    If OSver.MajorMinor >= 5.1 And OSver.MajorMinor <= 5.2 Then
        For Each Ctl In Me.Controls
            If TypeName(Ctl) = "OptionButton" Then
                Set OptB = Ctl
                SetWindowTheme OptB.hwnd, StrPtr(" "), StrPtr(" ")
            ElseIf TypeName(Ctl) = "CommandButton" Then
                Set Btn = Ctl
                SetWindowTheme Btn.hwnd, StrPtr(" "), StrPtr(" ")
            End If
        Next
        Set OptB = Nothing
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bAbortScanNow = True
    DoEvents
    SaveWindowPos Me, SETTINGS_SECTION_ADSSPY
    'Unload Me
    Me.Hide
    If UnloadMode = 0 Then
        Cancel = True 'user click -> don't unload
    Else
        bQueryUnload = True
    End If
End Sub

Private Sub Form_Resize()
    Const BUTTON_BOTTOM_INDENT As Long = 900&
    If Me.WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    txtUselessBlabber.Width = Me.ScaleWidth - 255
    fraMain.Width = Me.ScaleWidth - 255
    lstADSFound.Width = Me.ScaleWidth - 255
    txtADSContent.Width = Me.ScaleWidth - 255
    picStatus.Width = Me.ScaleWidth - 255
    
    lstADSFound.Height = Me.ScaleHeight - 3690 - 480
    txtADSContent.Height = Me.ScaleHeight - 4170 - 880
    cmdViewCopy.Top = Me.ScaleHeight - BUTTON_BOTTOM_INDENT
    cmdViewSave.Top = Me.ScaleHeight - BUTTON_BOTTOM_INDENT
    cmdViewEdit.Top = Me.ScaleHeight - BUTTON_BOTTOM_INDENT
    cmdViewBack.Top = Me.ScaleHeight - BUTTON_BOTTOM_INDENT
    picStatus.Top = Me.ScaleHeight - 375
End Sub

Private Sub cmdViewBack_Click()
    txtADSContent.Visible = False
    cmdViewCopy.Visible = False
    cmdViewSave.Visible = False
    cmdViewEdit.Visible = False
    cmdViewBack.Visible = False
    lstADSFound.Visible = True
    cmdRemove.Enabled = True
    'Ready.
    Status Translate(209), "2"
End Sub

Private Sub cmdViewCopy_Click()
    ClipboardSetText txtADSContent.Text
    'The contents of the currently displayed ADS have been copied to the clipboard.
    Status Translate(2200), "3"
End Sub

Private Sub cmdViewEdit_Click()
    'get path to Wordpad.exe
    Dim hKey&, sWordpadPath$, sStream$
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\WORDPAD.EXE", 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        'Unable to open the ADS for editing: Wordpad not found.
        MsgBoxW Translate(2201), vbExclamation
        Exit Sub
    End If
    
    sWordpadPath = String$(MAX_PATH, 0)
    If RegQueryValueEx(hKey, "", 0, ByVal 0, ByVal sWordpadPath, Len(sWordpadPath)) <> 0 Then
        RegCloseKey hKey
        'Unable to open the ADS for editing: Wordpad not found.
        MsgBoxW Translate(2201), vbExclamation
        Exit Sub
    End If
    
    sWordpadPath = TrimNull(Replace$(sWordpadPath, "%ProgramFiles%", Environ$("ProgramFiles")))
    sStream = lstADSFound.List(lstADSFound.ListIndex)
    sStream = Replace$(sStream, " : ", ":")
    sStream = Left$(sStream, InStr(sStream, "  (") - 1)
    ShellExecute Me.hwnd, "open", sWordpadPath, """" & PathX64(sStream) & """", vbNullString, 1
    'Ready.
    Status Translate(209), "2"
End Sub

Private Sub cmdViewSave_Click()
    Dim sFilename$, sStream$  ', sStreamName$
    sStream = lstADSFound.List(lstADSFound.ListIndex)
    sStream = Replace$(sStream, " : ", ":")
    sStream = Left$(sStream, InStr(sStream, "  (") - 1)
    
    'sStreamName = mid$(sStream, InStrRev(sStream, ":") + 1)
    'If Asc(left$(sStreamName, 1)) = 5 Then sStreamName = mid$(sStreamName, 2)
        
    'Save stream contents to file...
    'Filters:
    'Binary files
    'Text files
    'All files
    sFilename = SaveFileDialog(Translate(2202), AppPath(), GetFileNameAndExt(sStream) & "." & GetStreamName(sStream) & ".bin", _
        Translate(2203) & " (*.bin)|*.bin|" & _
        Translate(2204) & " (*.txt)|*.txt|" & _
        Translate(2205) & " (*.*)|*.*", Me.hwnd)
    
    If Len(sFilename) <> 0 Then
        'Note: FileCopyW is not supported here
        '// TODO: add anicode support - need to use CreateFile / ReadFile / WriteFile
    
        FileCopy sStream, sFilename
        
        If FileExists(sFilename) Then
            'Stream contents saved to
            Status Translate(2206) & " " & sFilename & " (" & FileLen(sFilename) & " bytes).", "4"
        Else
            'An error occurred saving the stream to disk.
            Status Translate(2207), "5"
        End If
    End If
End Sub

Private Function GetStreamName(sPath As String) As String
    Dim pos As Long
    pos = InStr(4, sPath, ":")
    If pos Then
        GetStreamName = Mid$(sPath, pos + 1)
    End If
End Function

Private Sub cmdScan_Click()
    Dim lTicks&, sDrives$(), i&, sScanFolder$
    'Abort scan
    If cmdScan.Caption = Translate(2208) Then
        bAbortScanNow = True
        Exit Sub
    End If
    
    cmdViewBack_Click
    lstADSFound.Clear
    lTicks = GetTickCount()
    bQuickScan = IIf(optScanLocation(0).Value, True, False)
    bScanFolder = IIf(optScanLocation(2), True, False)
    sScanFolder = txtScanFolder.Text
    bIgnoreEncryptable = IIf(chkIgnoreEncryptable.Value = 1, True, False)
    bCalcMD5 = IIf(chkCalcMD5.Value = 1, True, False)
    
    'Abort scan
    cmdScan.Caption = Translate(2208)
    cmdScan.Tag = "2"
    If bQuickScan Then
        If InStr(1, GetNTFSDrives, Left$(sWinDir, 2), vbTextCompare) = 0 Then
            '"Unable to scan the Windows folder, because the volume " & _
                   "containing this folder does not use NTFS format." & vbCrLf & _
                   "Uncheck the 'quick scan' option to scan all drives " & _
                   "with NTFS format."
            MsgBoxW Translate(2209), vbExclamation
        Else
            EnumADSInAllFiles sWinDir
        End If
    Else
        If Not bScanFolder Then
            sDrives = Split(GetNTFSDrives, "|")     'check all drives
            For i = 0 To UBound(sDrives)
            
                If StrComp(sDrives(i), SysDisk & "\", 1) = 0 Then
                    ToggleWow64FSRedirection False
                End If

                EnumADSInFile sDrives(i), True
                If Not (bAbortScanNow) Then
                    EnumADSInAllFiles sDrives(i)
                End If
                
                If StrComp(sDrives(i), SysDisk & "\", 1) = 0 Then
                    ToggleWow64FSRedirection True
                End If
                
                If bAbortScanNow Then Exit For
            Next i
        Else
            If sScanFolder <> vbNullString And FolderExists(sScanFolder) Then   'particular folder
            
                If StrComp(sScanFolder, SysDisk, 1) = 0 Or _
                    StrComp(sScanFolder, SysDisk & "\", 1) = 0 Or _
                    StrBeginWith(sScanFolder, sWinDir) Then
                    
                    ToggleWow64FSRedirection False
                End If
            
                EnumADSInFile sScanFolder, True
                If Not (bAbortScanNow) Then
                    EnumADSInAllFiles sScanFolder
                End If
                
                ToggleWow64FSRedirection True
            End If
        End If
    End If
    
    If bQueryUnload Then Exit Sub
    
    'Scan the system for alternate data streams
    cmdScan.Caption = Translate(2210)
    cmdScan.Tag = "1"
    lTicks = GetTickCount() - lTicks
    If IsRunningInIDE() Then
        If bAbortScanNow Then
            'Scan ABORTED, found [] alternate data streams (ADS's) in [*] sec.
            Status Replace$(Replace$(Translate(2212), "[]", lstADSFound.ListCount), "[*]", Format$(lTicks / 1000, "##0.00#")), "6"
        Else
            'Scan complete, found [] alternate data streams (ADS's) in [*] sec.
            Status Replace$(Replace$(Translate(2211), "[]", lstADSFound.ListCount), "[*]", Format$(lTicks / 1000, "##0.00#")), "7"
        End If
    Else
        If bAbortScanNow Then
            'Scan ABORTED, found [] alternate data streams (ADS's).
            Status Replace$(Translate(2214), "[]", lstADSFound.ListCount), "8"
        Else
            'Scan complete, found [] alternate data streams (ADS's).
            Status Replace$(Translate(2213), "[]", lstADSFound.ListCount), "9"
        End If
    End If
    bAbortScanNow = False
End Sub

Private Sub cmdRemove_Click()
    If lstADSFound.ListCount = 0 Then Exit Sub
    If lstADSFound.SelCount = 0 Then lstADSFound.Clear: Exit Sub
    
    Dim i&, j&, k&, sStream$, sLockedStreams$
    k = lstADSFound.SelCount
    'Are you sure you want to remove the selected [] ADS's from your system? They will be deleted permanently!
    If MsgBoxW(Replace$(Translate(2215), "[]", k), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    'go from bottom of list to prevent .RemoveItem messing up the For loop
    'Removing selected streams...
    Status Translate(2216), "10"
    
    ToggleWow64FSRedirection False
    
    For i = lstADSFound.ListCount - 1 To 0 Step -1
        If lstADSFound.Selected(i) Then
            sStream = lstADSFound.List(i)
            sStream = Replace$(sStream, " : ", ":")
            sStream = Left$(sStream, InStr(sStream, "  (") - 1)
            If DeleteFile(StrPtr(sStream)) Then
                lstADSFound.RemoveItem i
            Else
                sLockedStreams = sLockedStreams & lstADSFound.List(i) & vbCrLf
            End If
            j = j + 1
        End If
        'Removing selected streams...
        Status Translate(2216) & " " & Int(CLng(j) / k * 100) & " %"
    Next i
    
    ToggleWow64FSRedirection True
    
    lstADSFound.Clear
    If sLockedStreams <> vbNullString Then
        'The following ADS streams could not be deleted. They may be locked by another program:
        MsgBoxW Translate(2217) & _
               vbCrLf & vbCrLf & Left$(sLockedStreams, Len(sLockedStreams) - 2), vbExclamation
    End If
    Status Translate(209), "2"
End Sub

Private Function GetDrives$()
    Dim lDrives&, i&, sDrive$, sDrives$, lDriveType&
    Dim lFlags&, sVolName$, lVolSN&, lMaxCompLen&, sVolFileSys$
    lDrives = GetLogicalDrives()
    For i = 0 To 26
        If (lDrives And 2 ^ i) Then
            sDrive = Chr$(65 + i) & ":\"
            lDriveType = GetDriveType(sDrive)
            sVolName = String$(MAX_PATH, 0)
            sVolFileSys = String$(MAX_PATH, 0)
            lFlags = 0
            GetVolumeInformation sDrive, sVolName, Len(sVolName), lVolSN, lMaxCompLen, lFlags, sVolFileSys, Len(sVolFileSys)
            
            Select Case lDriveType
                Case DRIVE_UNKNOWN:   sDrive = sDrive & " - " & Translate(2218)     'unknown type
                Case DRIVE_REMOVABLE: sDrive = sDrive & " - " & Translate(2219)     'removable
                Case DRIVE_FIXED:     sDrive = sDrive & " - " & Translate(2220)     'fixed
                Case DRIVE_REMOTE:    sDrive = sDrive & " - " & Translate(2221)     'network
                Case DRIVE_CDROM:     sDrive = sDrive & " - " & Translate(2222)     'CD-ROM
                Case DRIVE_RAMDISK:   sDrive = sDrive & " - " & Translate(2223)     'RAMdisk
            End Select
            sVolFileSys = TrimNull(sVolFileSys)
            If sVolFileSys = vbNullString Then sVolFileSys = "?"
            sDrive = sDrive & " - " & sVolFileSys
            'supports ADS
            If (lFlags And FILE_NAMED_STREAMS) Then sDrive = sDrive & " - " & Translate(2224)
            
            sDrives = sDrives & "|" & sDrive
        End If
    Next i
    If sDrives <> vbNullString Then GetDrives = Mid$(sDrives, 2)
End Function

Private Function GetNTFSDrives$()
    Dim lFlags&, sVolName$, lVolSN&, lMaxCompLen&, sVolFileSys$
    Dim lDrives&, i&, sDrive$, sNTFSDrives$, lDriveType&
    lDrives = GetLogicalDrives()
    For i = 0 To 26
        If (lDrives And 2 ^ i) Then
            sDrive = Chr$(65 + i) & ":\"
            lDriveType = GetDriveType(sDrive)
            If lDriveType = DRIVE_FIXED Or lDriveType = DRIVE_RAMDISK Or lDriveType = DRIVE_REMOVABLE Then
                sVolName = String$(MAX_PATH, 0)
                sVolFileSys = String$(MAX_PATH, 0)
                lFlags = 0
                GetVolumeInformation sDrive, sVolName, Len(sVolName), lVolSN, lMaxCompLen, lFlags, sVolFileSys, Len(sVolFileSys)
                If (lFlags And FILE_NAMED_STREAMS) Then
                    sNTFSDrives = sNTFSDrives & "|" & sDrive
                End If
'                If UCase$(TrimNull(sVolFileSys)) = "NTFS" Then
'                    sNTFSDrives = sNTFSDrives & "|" & sDrive
'                End If
            End If
        End If
    Next i
    If sNTFSDrives <> vbNullString Then GetNTFSDrives = Mid$(sNTFSDrives, 2)
End Function

Private Sub CheckIfSystemIsNTFS()
    If GetNTFSDrives() = vbNullString Then
        'Alternate Data Streams (ADS) are only possible on NTFS systems." & vbCrLf & _
               "Since there are no NTFS volumes on this system, ADS Spy will now exit.
        MsgBoxW Translate(2225), vbInformation
        Unload Me
    End If
End Sub

Private Sub EnumADSInAllFiles(sFolder$)
    Dim hFind As Long, uWFD As WIN32_FIND_DATA, sFilename$, lpSTR&, sPath$
    
    hFind = FindFirstFile(StrPtr(BuildPath(sFolder, "*.*")), uWFD)
    If hFind = INVALID_HANDLE_VALUE Then
        Status "FindFirstFile() failed", ""
        Exit Sub
    End If
    
    picStatus.Tag = ""
    Do
        lpSTR = VarPtr(uWFD.dwReserved1) + 4&
        sFilename = String$(lstrlen(lpSTR), 0)
        lstrcpy StrPtr(sFilename), lpSTR
    
        If (uWFD.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT) <> FILE_ATTRIBUTE_REPARSE_POINT Then
    
          If Not ((uWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 16) Then
            sPath = BuildPath(sFolder, sFilename)
            Status sPath
            EnumADSInFile sPath
          Else
            If sFilename <> "." And sFilename <> ".." And Not bQuickScan Then
                sPath = BuildPath(sFolder, sFilename)
                EnumADSInFile sPath, True
                EnumADSInAllFiles sPath
            End If
          End If
        End If
        If bAbortScanNow Then Exit Do
    Loop Until FindNextFile(hFind, uWFD) = 0
    FindClose hFind
End Sub

Private Sub EnumADSInFile(sFilePath$, Optional bIsFolder As Boolean = False)

    Dim hFile&, uIOSB As IO_STATUS_BLOCK, uFSI As FILE_STREAM_INFORMATION
    Dim uBuffer() As Byte, lStreamInfo&, lBufferLen&, sStreamName$
    If bAbortScanNow Then Exit Sub

    If bIsFolder = False Then
        hFile = CreateFileW(StrPtr(sFilePath), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0, OPEN_EXISTING, 0, 0)
    Else
        hFile = CreateFileW(StrPtr(sFilePath), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0, OPEN_EXISTING, g_FileBackupFlag, 0)
    End If
    If hFile = -1 Then Exit Sub
    
    lBufferLen = 96 * CLng(4096)
    ReDim uBuffer(1 To lBufferLen)
    If 0 = NtQueryInformationFile(hFile, uIOSB, ByVal VarPtr(uBuffer(1)), lBufferLen, FileStreamInformation) Then
    
      lStreamInfo = VarPtr(uBuffer(1))
      Do
        CopyMemory ByVal VarPtr(uFSI.NextEntryOffset), ByVal lStreamInfo, 24
        CopyMemory ByVal VarPtr(uFSI.StreamName(0)), ByVal lStreamInfo + 24, uFSI.StreamNameLength
        sStreamName = Left$(uFSI.StreamName, uFSI.StreamNameLength / 2)
        If sStreamName <> vbNullString And _
           sStreamName <> "::$DATA" Then
           If InStr(Join(sSafeStreams, "|"), sStreamName) = 0 _
               Or Not bIgnoreEncryptable Then
                sStreamName = Mid$(sStreamName, 2)
                sStreamName = Left$(sStreamName, InStr(sStreamName, ":") - 1)
                If bCalcMD5 Then
                    lstADSFound.AddItem sFilePath & " : " & sStreamName & "  (" & uFSI.StreamSize & " bytes, CheckSum: " & GetFileCheckSum(sFilePath & ":" & sStreamName, uFSI.StreamSize, True) & ")"
                Else
                    lstADSFound.AddItem sFilePath & " : " & sStreamName & "  (" & uFSI.StreamSize & " bytes)"
                End If
            End If
        End If
        If uFSI.NextEntryOffset > 0 Then
            lStreamInfo = lStreamInfo + uFSI.NextEntryOffset
        Else
            Exit Do
        End If
        If bAbortScanNow Then Exit Do
      Loop
    End If
    CloseHandle hFile
End Sub

'Private Function FileExists(sFile$) As Boolean
'    FileExists = CBool(SHFileExists(StrConv(sFile, vbUnicode)))
'End Function

Private Sub Status(s$, Optional sTag As String)
    If Not (bQueryUnload) Then
        picStatus.Cls
        picStatus.Print " " & s
        picStatus.Tag = sTag
    End If
    DoEvents
End Sub

Private Function TrimNull$(s$)
    If InStr(s, Chr$(0)) > 0 Then
        TrimNull = Left$(s, InStr(s, Chr$(0)) - 1)
    Else
        TrimNull = s
    End If
End Function

Private Sub lstADSFound_DblClick()
    mnuPopupView_Click
End Sub

Private Sub lstADSFound_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then bAbortScanNow = True: Me.Hide
End Sub

Private Sub lstADSFound_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And lstADSFound.ListCount > 0 Then PopupMenu mnuPopup
End Sub

Private Sub mnuPopupSave_Click()
    If lstADSFound.ListCount = 0 Then Exit Sub
    
    Dim sFilename$, i&, hFile As Long
    Dim sList As clsStringBuilder
    
    'Save scan results to disk...
    sFilename = SaveFileDialog(Translate(203), AppPath(), "Streams.txt", _
        Translate(2204) & " (*.txt)|*.txt|" & _
        Translate(2205) & " (*.*)|*.*", Me.hwnd)
        
    If Len(sFilename) = 0 Then Exit Sub
    
    Set sList = New clsStringBuilder
    
    'Header
    sList.Append ChrW$(-257)
    sList.AppendLine "Logfile of ADS Spy v." & ADSspyVer & " (HJT Fork v." & AppVerString & ")"
    sList.AppendLine
    sList.Append MakeLogHeader()
    sList.AppendLine

    For i = 0 To lstADSFound.ListCount - 1
        sList.AppendLine lstADSFound.List(i)
    Next i
    
    sList.Append "--" & vbCrLf & "End of file"
    
    If OpenW(sFilename, FOR_OVERWRITE_CREATE, hFile, g_FileBackupFlag) Then
        PutW hFile, 1, StrPtr(sList.ToString), sList.Length * 2
        CloseW hFile
    End If
        
    'Scan results saved to
    Status Translate(2226) & " " & sFilename & ".", "11"
    
    OpenLogFile sFilename
    
    Set sList = Nothing
End Sub

Private Sub mnuPopupSelAll_Click()
    Dim i%
    If lstADSFound.SelCount = lstADSFound.ListCount Then Exit Sub
    For i = 0 To lstADSFound.ListCount - 1
        lstADSFound.Selected(i) = True
    Next i
    lstADSFound.ListIndex = 0
End Sub

Private Sub mnuPopupSelInvert_Click()
    Dim i%
    For i = 0 To lstADSFound.ListCount - 1
        lstADSFound.Selected(i) = Not lstADSFound.Selected(i)
    Next i
    lstADSFound.ListIndex = 0
End Sub

Private Sub mnuPopupSelNone_Click()
    Dim i%
    If lstADSFound.SelCount = 0 Then Exit Sub
    For i = 0 To lstADSFound.ListCount - 1
        lstADSFound.Selected(i) = False
    Next i
    lstADSFound.ListIndex = 0
End Sub

Private Sub mnuPopupView_Click()
    On Error GoTo ErrorHandler:

    cmdRemove.Enabled = False
    lstADSFound.Visible = False
    txtADSContent.Visible = True
    cmdViewCopy.Visible = True
    cmdViewSave.Visible = True
    cmdViewEdit.Visible = True
    cmdViewBack.Visible = True
    
    Dim sStream$, lSize&, sStreamContents$, ff%, hFile As Long
    'Reading stream...
    Status Translate(2227), "12"
    txtADSContent.Text = vbNullString
    sStream = lstADSFound.List(lstADSFound.ListIndex)
    sStream = Replace$(sStream, " : ", ":")
    lSize = Val(Mid$(sStream, InStr(sStream, "  (") + 3))
    sStream = Left$(sStream, InStr(sStream, "  (") - 1)
    
    If FileExists(sStream) Then
        
        ff = FreeFile()
        ToggleWow64FSRedirection False, sStream
        OpenW sStream, FOR_READ, hFile, g_FileBackupFlag
        If hFile > 0 Then
            sStreamContents = String$(LOFW(hFile), 0)
            GetW hFile, 1, sStreamContents
            
            If StrContainsBinary(sStreamContents) Then
                'The stream '[]' contains binary data and will most likely not display properly or completely.
                MsgBoxW Replace$(Translate(2228), "[]", sStream), vbExclamation
            End If
            CloseW hFile
        End If
        ToggleWow64FSRedirection True
        txtADSContent.Text = sStreamContents
    End If
    'Contents of [].
    Status Replace$(Translate(2229), "[]", sStream), "13"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "mnuPopupView_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub mnuPopupShowFile_Click()
    Dim sStream As String
    
    sStream = lstADSFound.List(lstADSFound.ListIndex)
    sStream = RTrim$(Left$(sStream, InStr(sStream, " :") - 1))
    
    If FileExists(sStream) Or FolderExists(sStream) Then
        Shell "explorer.exe /select," & """" & sStream & """", vbNormalFocus
    End If
End Sub

Function StrContainsBinary(s$) As Boolean
    Dim i&
    For i = 0 To 31
      If i <> 9 And i <> 10 And i <> 13 Then 'cr,lf,tab
        If InStr(s, Chr$(i)) <> 0 Then StrContainsBinary = True: Exit For
      End If
    Next
End Function

Private Function IsRunningInIDE() As Boolean
    IsRunningInIDE = (App.LogMode = 0)
End Function

Private Function BrowseForFolder$(sPrompt$)
    Dim uBI As BrowseInfo, lIDList&, sPath$
    With uBI
        .hWndOwner = Me.hwnd
        .lpszTitle = StrPtr(sPrompt)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lIDList = SHBrowseForFolder(uBI)
    If lIDList > 0 Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lIDList, StrPtr(sPath)
        CoTaskMemFree lIDList
        BrowseForFolder = TrimNull(sPath)
    End If
End Function
