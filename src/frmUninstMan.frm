VERSION 5.00
Begin VB.Form frmUninstMan 
   Caption         =   "Remove Programs Manager"
   ClientHeight    =   6756
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   11316
   Icon            =   "frmUninstMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6756
   ScaleWidth      =   11316
   Begin VB.Frame fraUninstMan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmdWebSiteOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   9960
         TabIndex        =   23
         Top             =   2300
         Width           =   1095
      End
      Begin VB.TextBox txtWebSite 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Frame fraFilter 
         Caption         =   "Filter"
         Height          =   1815
         Left            =   6000
         TabIndex        =   14
         Top             =   4200
         Width           =   5000
         Begin VB.CheckBox chkFilterHKU 
            Caption         =   "HKU (other users)"
            Height          =   255
            Left            =   2600
            TabIndex        =   20
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkFilterHKCU 
            Caption         =   "HKCU (current user)"
            Height          =   255
            Left            =   2600
            TabIndex        =   19
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkFilterHKLM 
            Caption         =   "HKLM (all users)"
            Height          =   255
            Left            =   2600
            TabIndex        =   18
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkFilterNoUninstStr 
            Caption         =   "No Uninstall command"
            Height          =   255
            Left            =   140
            TabIndex        =   17
            Top             =   720
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.CheckBox chkFilterHidden 
            Caption         =   "Hidden"
            Height          =   255
            Left            =   140
            TabIndex        =   16
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkFilterCommon 
            Caption         =   "Common Software"
            Height          =   255
            Left            =   140
            TabIndex        =   15
            Top             =   360
            Value           =   1  'Checked
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdKeyJump 
         Caption         =   "Jump"
         Height          =   375
         Left            =   9960
         TabIndex        =   13
         Top             =   2900
         Width           =   1095
      End
      Begin VB.TextBox txtKey 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3000
         Width           =   3735
      End
      Begin VB.CommandButton cmdNameEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   9960
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdUninstStrEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   9960
         TabIndex        =   9
         Top             =   1700
         Width           =   1095
      End
      Begin VB.ListBox lstUninstMan 
         Height          =   5100
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   5775
      End
      Begin VB.CommandButton cmdUninstall 
         Caption         =   "Uninstall application"
         Height          =   425
         Left            =   6240
         TabIndex        =   2
         Top             =   3540
         Width           =   1935
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save list..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   4320
         TabIndex        =   3
         Top             =   6170
         Width           =   1575
      End
      Begin VB.TextBox txtUninstCmd 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh list"
         Height          =   425
         Left            =   2040
         TabIndex        =   6
         Top             =   6170
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete this entry"
         Height          =   425
         Left            =   8640
         TabIndex        =   7
         Top             =   3540
         Width           =   1935
      End
      Begin VB.CommandButton cmdOpenCP 
         Caption         =   "Open Control Panel ""Remove Software"""
         Height          =   425
         Left            =   6240
         TabIndex        =   8
         Top             =   6170
         Width           =   4515
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblUninstCmd 
         Caption         =   "Uninstall command"
         Height          =   255
         Left            =   6000
         TabIndex        =   25
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label lblAbout 
         Caption         =   $"frmUninstMan.frx":4072
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   10935
      End
      Begin VB.Label lblKey 
         Caption         =   "Key"
         Height          =   255
         Left            =   6000
         TabIndex        =   22
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label lblWebSite 
         Caption         =   "Web-site"
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   2160
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmUninstMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[frmUninstMan.frm]

'
' Uninstall Manager by Merijn Bellekom & Alex Dragokas
'

Option Explicit

Private Type UnintallManagerData
    AppRegHive      As ENUM_REG_HIVE
    AppRegKey       As String
    AppRegRedir     As Boolean
    DisplayName     As String
    UninstString    As String
    WebSite         As String
    KeyTime         As String
    Version         As String
    Publisher       As String
    Hidden          As Boolean
    User            As String
End Type

Private UninstData() As UnintallManagerData

Private Sub Form_Load()
    If OSver.MajorMinor >= 5.1 Then
        SetWindowTheme Me.fraFilter.hwnd, StrPtr(" "), StrPtr(" ")
    End If
    SetAllFontCharset Me, g_FontName, g_FontSize, g_bFontBold
    ReloadLanguage True
    LoadWindowPos Me, SETTINGS_SECTION_UNINSTMAN
    cmdRefresh_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveWindowPos Me, SETTINGS_SECTION_UNINSTMAN

    If UnloadMode = 0 Then 'initiated by user (clicking 'X')
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
    ProcessHotkey KeyCode, Me
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 10980 Then Me.Width = 10980
    If Me.Height < 2715 Then Me.Height = 2715
    
    fraUninstMan.Width = Me.ScaleWidth - 200
    lstUninstMan.Width = Me.ScaleWidth - 5500
    txtName.Left = Me.ScaleWidth - 5250
    txtUninstCmd.Left = txtName.Left
    Me.txtWebSite.Left = txtName.Left
    Me.txtKey.Left = txtName.Left
    Me.lblName.Left = txtName.Left
    Me.lblUninstCmd.Left = txtName.Left
    Me.lblWebSite.Left = txtName.Left
    Me.lblKey.Left = txtName.Left
    Me.cmdNameEdit.Left = txtName.Left + txtName.Width + 100
    Me.cmdUninstStrEdit.Left = Me.cmdNameEdit.Left
    Me.cmdWebSiteOpen.Left = Me.cmdNameEdit.Left
    Me.cmdKeyJump.Left = Me.cmdNameEdit.Left
    Me.fraFilter.Left = txtName.Left - 25
    lblAbout.Width = Me.fraUninstMan.Width - 240
    Me.fraUninstMan.Height = Me.ScaleHeight - 30
    lstUninstMan.Height = Me.fraUninstMan.Height - 1635
    cmdRefresh.Top = lstUninstMan.Top + lstUninstMan.Height + 110
    cmdSave.Top = cmdRefresh.Top
    cmdSave.Left = lstUninstMan.Left + lstUninstMan.Width - cmdSave.Width
    cmdRefresh.Left = lstUninstMan.Left + lstUninstMan.Width \ 2 - cmdRefresh.Width \ 2
    cmdUninstall.Left = txtName.Left + 240
    cmdDelete.Left = txtName.Left + 2640
    cmdOpenCP.Left = cmdUninstall.Left
    'cmdOpenCP.Top = fraFilter.Top + fraFilter.Height + 155
End Sub

'
' ====== Uninstall manager  ======
'

'click on list item
'
Private Sub lstUninstMan_Click()
    Dim ItemID&, ID&, sKey$, Blink As Boolean
    
    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    
    ID = lstUninstMan.ItemData(ItemID)
    With UninstData(ID)
        If Not Reg.KeyExists(.AppRegHive, .AppRegKey, .AppRegRedir) Then
            lstUninstMan.RemoveItem ItemID
            ClearTextboxes
        Else
            txtName.Text = .DisplayName
            txtUninstCmd.Text = .UninstString
            txtWebSite.Text = .WebSite
            sKey = Reg.GetShortHiveName(Reg.GetHiveNameByHandle(.AppRegHive)) & "\" & .AppRegKey
            If .AppRegRedir Then
                sKey = Replace(sKey, "SOFTWARE", "SOFTWARE\Wow6432Node", 1, 1, 1) 'to support WOW64 keys by regedit.exe export
            End If
            txtKey.Text = sKey
            If Len(.WebSite) <> 0 Then
                If isURL(.WebSite) Then
                    Blink = True
                ElseIf StrBeginWith(.WebSite, "file:///") Then
                    If FileExists(Replace$(Mid$(.WebSite, 9), "/", "\")) Then
                        Blink = True
                    End If
                ElseIf FileExists(.WebSite) Or FolderExists(.WebSite) Then
                    Blink = True
                End If
            End If
            
            If Blink Then
                cmdWebSiteOpen.Enabled = True
            Else
                cmdWebSiteOpen.Enabled = False
            End If
            If Len(.UninstString) <> 0 Then
                cmdUninstStrEdit.Enabled = True
            Else
                cmdUninstStrEdit.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub ClearTextboxes()
    txtName.Text = vbNullString
    txtUninstCmd.Text = vbNullString
    txtWebSite.Text = vbNullString
    txtKey.Text = vbNullString
End Sub

' delete registry entry only
'
Private Sub cmdDelete_Click()
    On Error GoTo ErrorHandler:

    Dim sName$, sUninst$, ItemID&, ID&
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    With UninstData(ID)
        'Are you sure you want to delete this item from the list?
        If MsgBoxW(Translate(1710) & vbCrLf & vbCrLf & .DisplayName, vbQuestion Or vbYesNo) = vbYes Then
            Reg.DelKey .AppRegHive, .AppRegKey, .AppRegRedir
            lstUninstMan.RemoveItem (ItemID)
            If lstUninstMan.ListCount = 0 Then
                ClearTextboxes
            Else
                lstUninstMan.ListIndex = IIf(ItemID = -1, 0, ItemID)
            End If
        End If
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManDelete_Click"
    If inIDE Then Stop: Resume Next
End Sub

' Uninstall application
'
Private Sub cmdUninstall_Click()
    On Error GoTo ErrorHandler:

    Dim sName$, sUninst$, ItemID&, sApplication$, sArguments$, ID&
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    With UninstData(ID)
        'if no uninstall string
        If Len(.UninstString) = 0 Then
            'MsgBox "No uninstall string"
            Exit Sub
        Else
            'if require uninstallation under certain user
            If Len(.User) <> 0 Then
                'You should be logged as user '[]' to do this action!
                MsgBox Replace$(Translate(1713), "[]", .User), vbExclamation
                Exit Sub
            Else
                sApplication = FindOnPath(.UninstString)
                
                If FileExists(sApplication) Then
                    sArguments = ExtractArguments(.UninstString)
            
                    Proc.ProcessRunUnelevated2 PathX64(sApplication), sArguments
                    
                    'ShellExecute 0&, 0&, StrPtr(sApplication), StrPtr(sArguments), 0&, 1&
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManUninstall_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdNameEdit_Click()
    On Error GoTo ErrorHandler:
    
    Dim s$, sName$, sUninst$, ItemID&, ID&
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    If cmdNameEdit.Caption = Translate(216) Then 'Edit
        cmdNameEdit.Caption = Translate(219)
        txtName.BackColor = &H80000005 'white
        txtName.Locked = False
    Else 'Save
        cmdNameEdit.Caption = Translate(216)
        txtName.BackColor = &H8000000F 'gray
        txtName.Locked = True
        With UninstData(ID)
            .DisplayName = txtName.Text
            Reg.SetStringVal .AppRegHive, .AppRegKey, "DisplayName", .DisplayName, .AppRegRedir
        End With
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManEdit_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdNameEdit.Caption = Translate(219) Then   'Save
            cmdNameEdit_Click
        End If
    End If
End Sub

Private Sub cmdUninstStrEdit_Click()
    On Error GoTo ErrorHandler:
    
    Dim s$, sName$, sUninst$, ItemID&, ID&
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    If cmdUninstStrEdit.Caption = Translate(216) Then 'Edit
        cmdUninstStrEdit.Caption = Translate(219)
        txtUninstCmd.BackColor = &H80000005 'white
        txtUninstCmd.Locked = False
    Else 'Save
        cmdUninstStrEdit.Caption = Translate(216)
        txtUninstCmd.BackColor = &H8000000F 'gray
        txtUninstCmd.Locked = True
        With UninstData(ID)
            .UninstString = txtUninstCmd.Text
            Reg.SetStringVal .AppRegHive, .AppRegKey, "UninstallString", .UninstString, .AppRegRedir
            Reg.DelVal .AppRegHive, .AppRegKey, "QuietUninstallString", .AppRegRedir
        End With
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManEdit_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub txtUninstCmd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdUninstStrEdit.Caption = Translate(219) Then   'Save
            cmdUninstStrEdit_Click
        End If
    End If
End Sub

Private Sub cmdWebSiteOpen_Click()
    On Error GoTo ErrorHandler:
    
    Dim ItemID&, ID&, sURL$, sFile$
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    sURL = UninstData(ID).WebSite
    
    If Len(sURL) <> 0 Then
        If isURL(sURL) Then
            OpenURL sURL
        ElseIf StrBeginWith(sURL, "file:///") Then
            sFile = Replace$(Mid$(sURL, 9), "/", "\")
            If FileExists(sFile) Then
                OpenAndSelectFile sFile
            End If
        ElseIf FileExists(sURL) Or FolderExists(sURL) Then
            OpenAndSelectFile Replace(sURL, "\\", "\")
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdWebSiteOpen_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdKeyJump_Click()
    On Error GoTo ErrorHandler:
    
    Dim ItemID&, ID&
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    ItemID = lstUninstMan.ListIndex
    If ItemID = -1 Then Exit Sub
    ID = lstUninstMan.ItemData(ItemID)
    
    With UninstData(ID)
        Reg.Jump .AppRegHive, .AppRegKey, , .AppRegRedir
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "cmdKeyJump_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub cmdOpenCP_Click()
    Proc.ProcessRunUnelevated2 "control.exe", "appwiz.cpl"
    'ShellExecute 0&, StrPtr("open"), StrPtr("control.exe"), StrPtr("appwiz.cpl"), 0&, 1
End Sub

Private Sub chkFilterCommon_Click()
    cmdRefresh_Click
End Sub

Private Sub chkFilterHidden_Click()
    cmdRefresh_Click
End Sub

Private Sub chkFilterHKCU_Click()
    cmdRefresh_Click
End Sub

Private Sub chkFilterHKLM_Click()
    cmdRefresh_Click
End Sub

Private Sub chkFilterHKU_Click()
    cmdRefresh_Click
End Sub

Private Sub chkFilterNoUninstStr_Click()
    cmdRefresh_Click
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler:

    Dim aItems() As String, sName$, sUninst$, i&, j&, cnt&, bHidden As Boolean, sURL$, sPublisher$, bComply As Boolean
    Dim sVer$, aVer(3) As Byte, sVerMajor$, sVerMinor$, lVerNum As Long
    Dim HiveFilter As HE_HIVE
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    lstUninstMan.Clear
    Erase UninstData
    cnt = -1
    
    'lstUninstMan.Sorted must be False ' Do not enable this kind of sorting at all!!! Otherwise, virus will eat your computer :)
    
    If chkFilterHKLM.Value = vbChecked Then HiveFilter = HiveFilter Or HE_HIVE_HKLM
    If chkFilterHKCU.Value = vbChecked Then HiveFilter = HiveFilter Or HE_HIVE_HKCU
    If chkFilterHKU.Value = vbChecked Then HiveFilter = HiveFilter Or HE_HIVE_HKU
    
    HE.Init HiveFilter, HE_SID_USER Or HE_SID_NO_VIRTUAL, HE_REDIR_BOTH
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Uninstall"
    
    Do While HE.MoveNext
        For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aItems, HE.Redirected, , False)
        
            sName = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "DisplayName", HE.Redirected)
            sUninst = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "UninstallString", HE.Redirected)
            If Len(sUninst) = 0 Then
                sUninst = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "QuietUninstallString", HE.Redirected)
            End If
            bHidden = (1 = Reg.GetDword(HE.Hive, HE.Key & "\" & aItems(i), "SystemComponent", HE.Redirected))
            sPublisher = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "Publisher", HE.Redirected)
            sURL = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "HelpLink", HE.Redirected)
            If Len(sURL) = 0 Then
                sURL = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "URLInfoAbout", HE.Redirected)
                If Len(sURL) = 0 Then
                    sURL = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "URLUpdateInfo", HE.Redirected)
                End If
            End If
            sVer = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "DisplayVersion", HE.Redirected)
            If Len(sVer) = 0 Then
                sVerMajor = Reg.GetDword(HE.Hive, HE.Key & "\" & aItems(i), "VersionMajor", HE.Redirected)
                sVerMinor = Reg.GetDword(HE.Hive, HE.Key & "\" & aItems(i), "VersionMinor", HE.Redirected)
                If Not (sVerMajor = 0 And sVerMinor = 0) Then
                    sVer = CStr(sVerMajor) & "." & CStr(sVerMinor)
                End If
                If Len(sVer) = 0 Then
                    sVer = Reg.GetString(HE.Hive, HE.Key & "\" & aItems(i), "Inno Setup: Setup Version", HE.Redirected)
                    If Len(sVer) = 0 Then
                        lVerNum = Reg.GetDword(HE.Hive, HE.Key & "\" & aItems(i), "Version", HE.Redirected)
                        If lVerNum <> 0 Then
                            GetMem4 lVerNum, aVer(0)
                            sVer = CStr(aVer(3)) & "." & CStr(aVer(2)) & "." & CStr(aVer(1)) & "." & CStr(aVer(0))
                        End If
                    End If
                End If
            End If
            
            bComply = True
            If chkFilterHidden.Value = vbUnchecked And bHidden Then bComply = False
            If chkFilterNoUninstStr.Value = vbUnchecked And Len(sUninst) = 0 Then bComply = False
            If chkFilterCommon.Value = vbUnchecked And (Not bHidden And Len(sUninst) <> 0) Then bComply = False
            
            If Len(sName) <> 0 And bComply Then
                cnt = cnt + 1
                ReDim Preserve UninstData(cnt)
                With UninstData(cnt)
                    .DisplayName = sName
                    .UninstString = sUninst
                    .AppRegHive = HE.Hive
                    .AppRegKey = HE.Key & "\" & aItems(i)
                    .AppRegRedir = HE.Redirected
                    .KeyTime = ConvertDateToUSFormat(Reg.GetKeyTime(HE.Hive, HE.Key & "\" & aItems(i), HE.Redirected))
                    .Hidden = bHidden
                    .WebSite = sURL
                    .Publisher = sPublisher
                    .Version = sVer
                    If .AppRegHive = HKU Then
                        .User = HE.UserName
                        If Len(.User) = 0 Then .User = "Unknown"
                    End If
                End With
            End If
        Next
    Loop
    If cnt = -1 Then Exit Sub
    
    'Sorting user type array using bufer array of positions (c) Dragokas
    Dim pos() As Variant, names() As String: ReDim pos(cnt), names(cnt)
    For i = 0 To cnt: pos(i) = i: names(i) = UninstData(i).DisplayName: Next 'key of sort is DisplayName
    QuickSortSpecial names, pos, 0, cnt
    
    For i = 0 To cnt
        With UninstData(pos(i))
            sName = .DisplayName
            If Len(.UninstString) = 0 Then sName = sName & " (No Uninstall command)"
            If .Hidden Then sName = sName & " (Hidden)"
            If Len(.User) <> 0 Then sName = sName & " (User: " & .User & ")"
        End With
        lstUninstMan.AddItem sName
        lstUninstMan.ItemData(i) = pos(i)     'array marker
    Next
    
    If lstUninstMan.ListCount Then lstUninstMan.ListIndex = 0
    If lstUninstMan.Visible And lstUninstMan.Enabled Then
        lstUninstMan.SetFocus
    End If
    
    'Set HE = Nothing
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManRefresh_Click"
    If inIDE Then Stop: Resume Next
End Sub

Private Function FormatLogString(ID As Long) As String
    
    Dim sLine As String
    Dim sKey As String
    
    With UninstData(ID)
        sLine = .DisplayName
        
        sKey = Reg.GetShortHiveName(Reg.GetHiveNameByHandle(.AppRegHive))
        If .AppRegRedir Then sKey = sKey & "-x32"
        If .AppRegHive = HKU Then
            sKey = sKey & "\" & GetRootPath(.AppRegKey)
        End If
        sKey = sKey & "\...\" & GetFileNameAndExt(.AppRegKey)
        
        sLine = sLine & " (" & sKey & ") (Version: " & .Version & " - " & .Publisher & ")"
        
        If Len(.UninstString) = 0 Then sLine = sLine & " (No Uninstall command)"
        If .Hidden Then sLine = sLine & " Hidden"
        If Len(.User) <> 0 Then sLine = sLine & " (User: " & .User & ")"
    End With

    FormatLogString = sLine
    
End Function

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler:
    
    Dim i&, sFile$, hFile&, ID&, bShowHeader As Boolean, Stady&, HE As clsHiveEnum
    Dim sList As clsStringBuilder
    
    Set HE = New clsHiveEnum
    Set sList = New clsStringBuilder
    
    Stady = 1
    
    If lstUninstMan.ListCount = 0 Then Exit Sub
    
    Stady = 2
    
    'sFile = SaveFileDialog("Save "Intalled Software list" to disk...", "Text files (*.txt)|*.txt|All files (*.*)|*.*", "uninstall_list.txt")
    sFile = SaveFileDialog(Translate(1711), AppPath(), "uninstall_list.txt", Translate(1712) & " (*.txt)|*.txt|" & Translate(1003) & " (*.*)|*.*")
    
    Stady = 3
    
    If Len(sFile) = 0 Then Exit Sub
    
    sList.Append ChrW$(-257)
    sList.AppendLine "Logfile of Uninstall manager v." & UninstManVer & " (HJT Fork v." & AppVerString & ")"
    sList.AppendLine
    sList.Append MakeLogHeader()
    
    Stady = 4
    
    'Log filters used
    If chkFilterCommon.Value = vbChecked And _
        chkFilterNoUninstStr.Value = vbChecked And _
        chkFilterHidden.Value = vbChecked And _
        chkFilterHKLM.Value = vbChecked And _
        chkFilterHKCU.Value = vbChecked And _
        chkFilterHKU.Value = vbChecked Then
        
        sList.AppendLine "Scan mode: Default"
    Else
        sList.AppendLine "Scan mode: Specific"
        sList.AppendLine IIf(chkFilterHKLM.Value = vbChecked, "{v}", "{-}") & " HKLM" & vbTab & _
            IIf(chkFilterCommon.Value = vbChecked, "{v}", "{-}") & " Common software"
            
        sList.AppendLine IIf(chkFilterHKCU.Value = vbChecked, "{v}", "{-}") & " HKCU" & vbTab & _
            IIf(chkFilterNoUninstStr.Value = vbChecked, "{v}", "{-}") & " No Uninstall string"
            
        sList.AppendLine IIf(chkFilterHKU.Value = vbChecked, "{v}", "{-}") & " HKU " & vbTab & _
            IIf(chkFilterHidden.Value = vbChecked, "{v}", "{-}") & " Hidden"
    End If
    
    sList.AppendLine ""
    sList.AppendLine String$(55, "-")
    sList.AppendLine Space$(20) & "Sort by Alphabet"
    sList.AppendLine String$(55, "-")
    sList.AppendLine
    
    Stady = 5
    
    For i = 0 To lstUninstMan.ListCount - 1
        ID = lstUninstMan.ItemData(i)
        sList.AppendLine FormatLogString(ID)
    Next i
    
    sList.AppendLine
    sList.AppendLine
    sList.AppendLine String$(55, "-")
    sList.AppendLine Space$(20) & "Sort by Date"
    sList.AppendLine String$(55, "-")
    sList.AppendLine
    
    Stady = 6
    
    ' Make positions array of sorting by .KeyTime property (registry key date).
    Dim cnt&: cnt = lstUninstMan.ListCount - 1
    Dim pos() As Variant, names() As String: ReDim pos(cnt), names(cnt)
    For i = 0 To cnt: pos(i) = i: names(i) = UninstData(i).KeyTime: Next
    
    Stady = 7
    QuickSortSpecial names, pos, 0, cnt
    
    Stady = 8
    
    For i = cnt To 0 Step -1 'descending order
        sList.AppendLine UninstData(pos(i)).KeyTime & vbTab & FormatLogString(CLng(pos(i)))
    Next
    
    Stady = 9
    
    sList.AppendLine
    sList.AppendLine
    sList.AppendLine String$(55, "-")
    sList.AppendLine Space$(11) & "Uninstall Key Registry Snapshot"
    sList.AppendLine String$(55, "-")
    sList.AppendLine
    
    Stady = 10
    
    Dim HiveFilter As HE_HIVE
    If chkFilterHKLM.Value = vbChecked Then HiveFilter = HiveFilter Or HE_HIVE_HKLM
    If chkFilterHKCU.Value = vbChecked Then HiveFilter = HiveFilter Or HE_HIVE_HKCU
    If chkFilterHKU.Value = vbChecked Then HiveFilter = HiveFilter Or HE_HIVE_HKU
    
    HE.Init HiveFilter, HE_SID_USER Or HE_SID_NO_VIRTUAL, HE_REDIR_BOTH
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Uninstall"
    
    bShowHeader = True
    Do While HE.MoveNext
        'walkaround for: ExportKeyToVariable via reg.exe doesn't show correct path to Wow6432Node keys
        If HE.Redirected Then
            sList.AppendLine Reg.ExportKeyToVariable(HE.Hive, "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", False, bShowHeader, True)
            sList.AppendLine
        Else
            sList.AppendLine Reg.ExportKeyToVariable(HE.Hive, HE.Key, False, bShowHeader, True)
            sList.AppendLine
        End If
        bShowHeader = False
    Loop
    
    sList.Append "--" & vbCrLf & "End of file"
    
    Stady = 11
    
    If FileExists(sFile) Then DeleteFileWEx (StrPtr(sFile))
    
    Stady = 12
    
    If OpenW(sFile, FOR_OVERWRITE_CREATE, hFile, g_FileBackupFlag) Then
        PutW hFile, 1, StrPtr(sList.ToString), sList.Length * 2
        CloseW hFile, True
    End If
    
    Stady = 13
    
    OpenLogFile sFile
    
    'Set HE = Nothing
    Set sList = Nothing
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain.cmdUninstManSave_Click", "Stady: " & Stady
    If inIDE Then Stop: Resume Next
End Sub



