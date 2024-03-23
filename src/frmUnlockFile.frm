VERSION 5.00
Object = "{317589D1-37C8-47D9-B5B0-1C995741F353}#1.0#0"; "VBCCR17.OCX"
Begin VB.Form frmUnlockFile 
   Caption         =   "Files Unlocker"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmUnlockFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8445
   Begin VBCCR17.FrameW FrameGo 
      Height          =   615
      Left            =   1080
      TabIndex        =   12
      Top             =   4200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1085
      BorderStyle     =   0
      Caption         =   "FrameW2"
      Begin VBCCR17.CommandButtonW cmdGo 
         Height          =   495
         Left            =   5640
         TabIndex        =   3
         Top             =   120
         Width           =   1575
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   12648384
         Caption         =   "Go"
      End
      Begin VBCCR17.CheckBoxW chkRecur 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   873
         Value           =   1
         Caption         =   "Recursively (including files and subfolders)"
      End
   End
   Begin VBCCR17.FrameW FramePerm 
      Height          =   1695
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2990
      Begin VBCCR17.OptionButtonW optPermCustom 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Caption         =   "Custom SDDL:"
      End
      Begin VBCCR17.OptionButtonW optPermDefault 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         Value           =   -1  'True
         Caption         =   "Default permissions"
      End
      Begin VBCCR17.TextBoxW txtSDDL 
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1085
         Enabled         =   0   'False
         MultiLine       =   -1  'True
         ScrollBars      =   1
      End
      Begin VBCCR17.CommandButtonW cmdPickSDDL 
         Height          =   330
         Left            =   4080
         TabIndex        =   10
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         Caption         =   "Pick from folder..."
      End
   End
   Begin VBCCR17.CommandButtonW cmdAddFile 
      Height          =   492
      Left            =   6720
      TabIndex        =   6
      Top             =   600
      Width           =   1572
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Add File(s)..."
   End
   Begin VBCCR17.CommandButtonW cmdAddFolder 
      Height          =   492
      Left            =   6720
      TabIndex        =   5
      Top             =   1200
      Width           =   1572
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Add Folder(s)..."
   End
   Begin VBCCR17.CommandButtonW cmdJump 
      Height          =   456
      Left            =   6720
      TabIndex        =   4
      Top             =   1920
      Width           =   1572
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Open in Explorer"
   End
   Begin VBCCR17.TextBoxW txtInput 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6372
      _ExtentX        =   0
      _ExtentY        =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3
   End
   Begin VBCCR17.LabelW lblWhatToDo 
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6132
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Enter file(s) and folder(s) to unlock and reset access:"
   End
End
Attribute VB_Name = "frmUnlockFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmUnlockFile.frm]

'
' Files unlocker by Alex Dragokas
'

Option Explicit

Private sList As clsStringBuilder

Private Sub cmdAddFile_Click()
    Dim aFile() As String
    Static LastLocation As String
    Dim i As Long
    For i = 1 To OpenFileDialog_Multi(aFile, Translate(1003), IIf(FolderExists(LastLocation), LastLocation, Desktop), Translate(1003) & " (*.*)|*.*", Me.hWnd)
        If i = 1 Then
            LastLocation = GetParentDir(aFile(i))
        End If
        txtInput.Text = txtInput.Text & IIf(Len(txtInput.Text) = 0, "", vbCrLf) & aFile(i)
    Next
End Sub

Private Sub cmdAddFolder_Click()
    Dim aFolder() As String
    Static LastLocation As String
    Dim i As Long
    For i = 1 To OpenFolderDialog_Multi(aFolder, , IIf(FolderExists(LastLocation), LastLocation, Desktop), Me.hWnd)
        If i = 1 Then
            LastLocation = GetParentDir(aFolder(i))
        End If
        txtInput.Text = txtInput.Text & IIf(Len(txtInput.Text) = 0, "", vbCrLf) & aFolder(i)
    Next
End Sub

Private Sub cmdGo_Click()
    On Error GoTo ErrorHandler:

    Dim sFiles          As String
    Dim aFiles()        As String
    Dim aFolders()      As String
    Dim vFile           As Variant
    Dim Recursively     As Boolean
    Dim hFile           As Long
    Dim sLogPath        As String
    Dim TimeStarted     As String
    Dim TimeFinished    As String
    Dim bFolder         As Boolean
    Dim i               As Long
    Dim SDDL            As String
    
    If Me.optPermDefault Then
        SDDL = GetDefaultFileSDDL()
    Else
        SDDL = Me.txtSDDL.Text
        
        If Not IsValidSDDL(SDDL) Then
            'Invalid SDDL specified!
            MsgBox Translate(2416), vbExclamation
            Exit Sub
        End If
    End If
    
    sLogPath = BuildPath(AppPath(), "FixFile.log")
    
    sFiles = txtInput.Text
    
    If Len(sFiles) = 0 Then
        'You should enter at least one file or folder!
        MsgBoxW Translate(2405), vbExclamation
        Exit Sub
    End If
    
    TimeStarted = GetTime()
    
    Set sList = New clsStringBuilder
    sList.Append ChrW$(-257)
    sList.AppendLine "Logfile of Files Permission Unlocker (HJT+ v." & AppVerString & ")"
    sList.AppendLine
    sList.AppendLine MakeLogHeader()
    sList.AppendLine
    sList.AppendLine "Logging started at:      " & TimeStarted
    sList.AppendLine
    
    Recursively = (chkRecur.Value = 1)
    
    sFiles = Replace$(sFiles, vbCr, vbNullString)
    aFiles = Split(sFiles, vbLf)
    
    For Each vFile In aFiles
    
        vFile = Trim$(UnQuote(CStr(vFile)))
        
        If StrEndWith(CStr(vFile), "\") Then vFile = Left$(vFile, Len(vFile) - 1)
    
        If Len(vFile) <> 0 Then
            
            bFolder = FolderExists(vFile)
            
            If bFolder Or FileExists(vFile) Then
                
                Call UnlockMe(CStr(vFile), SDDL, Recursively)
                
' // TODO: log each file separately (only [Failed] events)
'
'                If Recursively And bFolder Then
'
'                    UnlockSubfolders CStr(vFile), True
'
'                End If
            Else
                sList.AppendLine "(not found)" & " - " & vFile
            End If
        End If
    Next
    
    sList.AppendLine
    TimeFinished = GetTime()
    sList.AppendLine "Logging finished at:     " & TimeFinished
    sList.AppendLine
    sList.Append "--" & vbCrLf & "End of file"
    
    If OpenW(sLogPath, FOR_OVERWRITE_CREATE, hFile, g_FileBackupFlag) Then
        PutW hFile, 1, StrPtr(sList.ToString), sList.Length * 2
        CloseW hFile, True
    End If
    
    OpenLogFile sLogPath
    
    Set sList = Nothing
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmUnlockFile.cmdGo_Click"
    If inIDE Then Stop: Resume Next
End Sub


'Private Sub UnlockSubfolders(path As String, Optional Recursively As Boolean = False)
'    On Error GoTo ErrorHandler
'
'    Dim SubPathName     As String
'    Dim PathName        As String
'    Dim hFind           As Long
'    Dim L               As Long
'    Dim lpSTR           As Long
'    Dim fd              As WIN32_FIND_DATA
'
'    Do
'        If hFind <> 0& Then
'            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: Exit Do
'        Else
'            hFind = FindFirstFile(StrPtr(path & "\*"), fd)
'            If hFind = INVALID_HANDLE_VALUE Then Exit Do
'        End If
'
'        L = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT
'        Do While L <> 0&
'            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0: Exit Do
'            L = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT
'        Loop
'
'        If hFind <> 0& Then
'            lpSTR = VarPtr(fd.dwReserved1) + 4&
'            PathName = Space$(lstrlen(lpSTR))
'            lstrcpy StrPtr(PathName), lpSTR
'
'            If fd.dwFileAttributes And vbDirectory Then
'                If PathName <> "." Then
'                    If PathName <> ".." Then
'                        SubPathName = path & "\" & PathName
'
'                        Call UnlockMe(SubPathName)
'
'                        If Recursively Then
'                            Call UnlockSubfolders(SubPathName, True)
'                        End If
'                    End If
'                End If
'            End If
'        End If
'
'    Loop While hFind
'
'    Exit Sub
'ErrorHandler:
'    ErrorMsg Err, "UnlockSubfolders", "Folder:", path
'    Resume Next
'End Sub

Private Sub UnlockMe(sObject As String, SDDL As String, bRecurse As Boolean)

    Dim SDDL_Before As String
    Dim SDDL_After As String
    Dim bSuccess As Boolean

    SDDL_Before = GetFileStringSD(sObject)
    
    bSuccess = SetFileStringSD(sObject, SDDL, bRecurse)
    
    SetFileAttributes StrPtr(sObject), FILE_ATTRIBUTE_ARCHIVE
    
    '[OK], [Fail]
    sList.AppendLine IIf(bSuccess, Translate(2406), Translate(2408)) & " - " & sObject
    
    SDDL_After = GetFileStringSD(sObject)
    
    sList.AppendLine "." & vbCrLf & "  " & SDDL_Before & vbCrLf & "=>" & SDDL_After & vbCrLf

End Sub

Private Function GetTime() As String
    GetTime = Right$("0" & Day(Now), 2) & "." & Right$("0" & Month(Now), 2) & "." & Year(Now) & " - " & _
            Right$("0" & Hour(Now), 2) & ":" & Right$("0" & Minute(Now), 2)
End Function

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdJump_Click()
    Dim sFiles As String
    Dim aFiles() As String
    
    sFiles = txtInput.Text
    
    If Len(sFiles) = 0 Then
        'You should enter at least one file or folder!
        MsgBoxW Translate(2405), vbExclamation
        Exit Sub
    End If
    
    sFiles = Replace$(sFiles, vbCr, vbNullString)
    aFiles = Split(TrimEx(sFiles, vbLf), vbLf)
    
    OpenAndSelectFile aFiles(0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
    ProcessHotkey KeyCode, Me
End Sub

Private Sub Form_Load()
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    ReloadLanguage True
    LoadWindowPos Me, SETTINGS_SECTION_FILEUNLOCKER
    SubClassTextbox Me.txtInput.hWnd, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    SaveWindowPos Me, SETTINGS_SECTION_FILEUNLOCKER

    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        Cancel = True
        Me.Hide
    Else
        SubClassTextbox Me.txtInput.hWnd, False
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Width < 8655 Then Me.Width = 8655
        If Me.Height < 5505 Then Me.Height = 5505
    End If
    
    txtInput.Height = Me.Height - 3690
    txtInput.Width = Me.Width - 2230
    
    FramePerm.Top = txtInput.Top + txtInput.Height + 50 'Me.Height - 1290
    FrameGo.Top = FramePerm.Top + FramePerm.Height + 50
    
    Me.cmdAddFile.Left = Me.Width - 1900
    Me.cmdAddFolder.Left = Me.Width - 1900
    Me.cmdJump.Left = Me.Width - 1900
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub

Private Sub cmdPickSDDL_Click()
    Static LastLocation As String
    Dim sFolder As String
    sFolder = OpenFolderDialog("", IIf(FolderExists(LastLocation), LastLocation, Desktop), Me.hWnd)
    If FolderExists(sFolder) Then
        txtSDDL.Text = GetFileStringSD(sFolder)
        If Len(txtSDDL.Text) <> 0 Then
            optPermCustom.Value = True
        End If
        LastLocation = sFolder
    End If
End Sub

Private Sub optPermCustom_Click()
    txtSDDL.Enabled = True
End Sub

Private Sub optPermDefault_Click()
    txtSDDL.Enabled = False
End Sub

