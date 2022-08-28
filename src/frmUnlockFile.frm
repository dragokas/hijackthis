VERSION 5.00
Begin VB.Form frmUnlockFile 
   Caption         =   "Files Unlocker"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8448
   Icon            =   "frmUnlockFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   8448
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Add File(s)..."
      Height          =   492
      Left            =   6720
      TabIndex        =   6
      Top             =   1200
      Width           =   1572
   End
   Begin VB.CommandButton cmdAddFolder 
      Caption         =   "Add Folder(s)..."
      Height          =   492
      Left            =   6720
      TabIndex        =   5
      Top             =   600
      Width           =   1572
   End
   Begin VB.CommandButton cmdJump 
      Caption         =   "Open in Explorer"
      Height          =   456
      Left            =   6720
      TabIndex        =   4
      Top             =   1920
      Width           =   1572
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox chkRecur 
      Caption         =   "Recursively (process files and all subfolders)"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.TextBox txtInput 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Width           =   6372
   End
   Begin VB.Label lblWhatToDo 
      Caption         =   "Enter file(s) and folder(s) to unlock and reset access:"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6132
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
    Dim i As Long
    For i = 1 To OpenFileDialog_Multi(aFile, Translate(1003), Desktop, Translate(1003) & " (*.*)|*.*", Me.hwnd)
        txtInput.Text = txtInput.Text & vbCrLf & aFile(i)
    Next
End Sub

Private Sub cmdAddFolder_Click()
    Dim aFile() As String
    Dim i As Long
    For i = 1 To OpenFolderDialog_Multi(aFile, , Desktop, Me.hwnd)
        txtInput.Text = txtInput.Text & vbCrLf & aFile(i)
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
    sList.AppendLine "Logfile of Files Unlocker (HJT v." & AppVerString & ")"
    sList.AppendLine
    sList.AppendLine MakeLogHeader()
    sList.AppendLine "Logging started at:      " & TimeStarted
    sList.AppendLine
    
    TimeStarted = GetTime()
    
    Recursively = (chkRecur.Value = 1)
    
    sFiles = Replace$(sFiles, vbCr, vbNullString)
    aFiles = Split(sFiles, vbLf)
    
    For Each vFile In aFiles
    
        vFile = UnQuote(Trim$(vFile))
    
        If Len(vFile) <> 0 Then
            
            bFolder = FolderExists(vFile)
            
            If bFolder Or FileExists(vFile) Then
                
                Call UnlockMe(CStr(vFile))
                
                If Recursively And bFolder Then

                    UnlockSubfolders CStr(vFile), True

                End If
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


Private Sub UnlockSubfolders(Path As String, Optional Recursively As Boolean = False)
    On Error GoTo ErrorHandler
    
    Dim SubPathName     As String
    Dim PathName        As String
    Dim hFind           As Long
    Dim L               As Long
    Dim lpSTR           As Long
    Dim fd              As WIN32_FIND_DATA
    
    Do
        If hFind <> 0& Then
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: Exit Do
        Else
            hFind = FindFirstFile(StrPtr(Path & "\*"), fd)
            If hFind = INVALID_HANDLE_VALUE Then Exit Do
        End If
        
        L = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT
        Do While L <> 0&
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0: Exit Do
            L = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT
        Loop
    
        If hFind <> 0& Then
            lpSTR = VarPtr(fd.dwReserved1) + 4&
            PathName = Space$(lstrlen(lpSTR))
            lstrcpy StrPtr(PathName), lpSTR
            
            If fd.dwFileAttributes And vbDirectory Then
                If PathName <> "." Then
                    If PathName <> ".." Then
                        SubPathName = Path & "\" & PathName
                        
                        Call UnlockMe(SubPathName)
                        
                        If Recursively Then
                            Call UnlockSubfolders(SubPathName, True)
                        End If
                    End If
                End If
            End If
        End If
        
    Loop While hFind
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "UnlockSubfolders", "Folder:", Path
    Resume Next
End Sub


Private Sub UnlockMe(sObject As String)

    Dim SDDL_Before As String
    Dim SDDL_After As String
    Dim bSuccess As Boolean

    SDDL_Before = GetFileStringSD(sObject)
    
    bSuccess = TryUnlock(sObject, False)
    
    '[OK], [Fail]
    sList.AppendLine IIf(bSuccess, Translate(2406), Translate(2408)) & " - " & sObject
    
    SDDL_After = GetFileStringSD(sObject)
    
    sList.AppendLine "." & vbCrLf & SDDL_Before & vbCrLf & "=>" & vbCrLf & SDDL_After & vbCrLf

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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    SaveWindowPos Me, SETTINGS_SECTION_FILEUNLOCKER

    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Width < 7860 Then Me.Width = 7860
        If Me.Height < 2570 Then Me.Height = 2570
    End If
    txtInput.Width = Me.Width - 2230
    txtInput.Height = Me.Height - 2010
    chkRecur.Top = Me.Height - 1300
    cmdGo.Top = Me.Height - 1300
    Me.cmdAddFile.Left = Me.Width - 1900
    Me.cmdAddFolder.Left = Me.Width - 1900
    Me.cmdJump.Left = Me.Width - 1900
    Me.cmdJump.Visible = Me.ScaleHeight > 2200
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub
