Attribute VB_Name = "modHosts"
'[modHosts.bas]

'
' Hosts file module by Merijn Bellekom & Alex Dragokas
'

Option Explicit

Public Sub ListHostsFile(objList As VBCCR17.ListBoxW)
    On Error GoTo ErrorHandler:

    Dim sAttr$, iAttr&, aContent() As String, i&
    
    objList.Clear
    
    Dim objInfo As VBCCR17.LabelW
    Set objInfo = frmHostsMan.lblHostsTip1
    
    'objInfo.Caption = "Loading hosts file, please wait..."
    objInfo.Caption = Translate(279)

    If Not FileExists(g_HostsFile) Then
        'If MsgBoxW("Cannot find the hosts file." & vbCrLf & "Do you want to create a new, default hosts file?", vbExclamation + vbYesNo) = vbNo Then
        If MsgBoxW(Translate(280), vbExclamation Or vbYesNo) = vbNo Then
            'objInfo.Caption = "No hosts file found."
            objInfo.Caption = Translate(281)
            Exit Sub
        Else
            CreateDefaultHostsFile
        End If
    End If
    
    'Loading hosts file, please wait...
    objInfo.Caption = Translate(279)
    DoEvents
    iAttr = GetFileAttributes(StrPtr(g_HostsFile))
    If (iAttr And FILE_ATTRIBUTE_READONLY) Then sAttr = sAttr & "R"
    If (iAttr And FILE_ATTRIBUTE_ARCHIVE) Then sAttr = sAttr & "A"
    If (iAttr And FILE_ATTRIBUTE_HIDDEN) Then sAttr = sAttr & "H"
    If (iAttr And FILE_ATTRIBUTE_SYSTEM) Then sAttr = sAttr & "S"
    If (iAttr And FILE_ATTRIBUTE_COMPRESSED) Then sAttr = sAttr & "C"
    
    aContent = ReadHostsFileToArray()
    
    For i = 0 To UBound(aContent)
        objList.AddItem aContent(i)
    Next
    
    'objInfo.Caption = "Hosts file is located at " & sHostsFile & "Line: [], Attributes: []"
    objInfo.Caption = Translate(271) & " " & g_HostsFile & _
                      " (" & Translate(278) & " " & objList.ListCount & ", " & Translate(277) & " " & _
                      sAttr & ")"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ListHostsFile"
    If inIDE Then Stop: Resume Next
End Sub

Private Function ReadHostsFileToArray() As String()
    Dim sContents As String
    sContents = ReadFileContents(g_HostsFile, False)
    sContents = Replace$(sContents, vbCrLf, vbLf)
    ReadHostsFileToArray = SplitSafe(sContents, vbLf, True)
End Function

Private Function WriteHostsFileContents(sContents As String) As Boolean
    Dim hFile As Long
    If OpenW(g_HostsFile, FOR_OVERWRITE_CREATE, hFile) Then
        WriteHostsFileContents = PutString(hFile, 1, sContents, False)
        CloseW hFile
    End If
End Function

Public Function CreateDefaultHostsFile() As Boolean
    CreateDefaultHostsFile = WriteHostsFileContents(GetDefaultHostsContents())
End Function

Public Function HostsReset() As Boolean
    'Are you really sure to reset hosts file contents to defaults?
    If MsgBox(Translate(301), vbQuestion Or vbYesNo, "") = vbYes Then
        HostsReset = CreateDefaultHostsFile()
    End If
End Function

Public Function GetDefaultHostsContents() As String
    If OSver.MajorMinor < 6 Then
        GetDefaultHostsContents = LoadEncryptedResFile("database\hosts_xp", 105, Not inIDE)
    ElseIf OSver.MajorMinor < 6.1 Then
        GetDefaultHostsContents = LoadEncryptedResFile("database\hosts_vista", 106, Not inIDE)
    Else
        GetDefaultHostsContents = LoadEncryptedResFile("database\hosts_7-11", 107, Not inIDE)
    End If
End Function

Public Sub HostsDeleteLine(objList As VBCCR17.ListBoxW)
    On Error GoTo ErrorHandler:

    'delete the line in hosts file
    Dim iAttr&, i&
    
    iAttr = GetFileAttributes(StrPtr(g_HostsFile))
    
    If SetFileAttributes(StrPtr(g_HostsFile), vbArchive) = 0 Then
        'The hosts file is locked for reading and cannot be edited
        'Make sure you have privileges to modify the hosts file and no program is protecting it against changes.
        MsgBoxW Translate(282) & vbCrLf & Translate(284), vbCritical
        Exit Sub
    End If
    
    Dim sb As clsStringBuilder
    Set sb = New clsStringBuilder
    
    With objList
        For i = 0 To .ListCount - 1
            If Not .Selected(i) Then sb.AppendLine .List(i)
        Next i
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then .RemoveItem i
        Next i
    End With
    If sb.Length > 1 Then sb.Remove sb.Length - 1, 2 ' -CRLF
    
    If Not WriteHostsFileContents(sb.ToString()) Then
        'Unable to write the selected changes to your hosts file. Another program may be denying access to it, or your user account may have insufficient rights to access it.
        'Make sure you have privileges to modify the hosts file and no program is protecting it against changes.
        MsgBoxW Translate(303) & vbCrLf & Translate(284), vbCritical
        'revert changes
        ListHostsFile objList
    End If
    
    SetFileAttributes StrPtr(g_HostsFile), iAttr
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "HostsDeleteLine"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub HostsToggleLine(objList As VBCCR17.ListBoxW)
    On Error GoTo ErrorHandler:

    'enable/disable the line in hosts file
    Dim iAttr&, sLine As String, i&
    
    iAttr = GetFileAttributes(StrPtr(g_HostsFile))
    
    If SetFileAttributes(StrPtr(g_HostsFile), vbArchive) = 0 Then
        'The hosts file is locked for reading and cannot be edited
        'Make sure you have privileges to modify the hosts file and no program is protecting it against changes.
        MsgBoxW Translate(282) & vbCrLf & Translate(284), vbCritical
        Exit Sub
    End If
    
    Dim sb As clsStringBuilder
    Set sb = New clsStringBuilder
    
    With objList
        For i = 0 To .ListCount - 1
            sLine = .List(i)
            If .Selected(i) Then
                If Left$(LTrim$(sLine), 1) = "#" Then
                    sLine = mid$(sLine, 2)
                Else
                    sLine = "#" & sLine
                End If
            End If
            .List(i) = sLine
            sb.AppendLine sLine
        Next
    End With
    If sb.Length > 1 Then sb.Remove sb.Length - 1, 2 ' -CRLF
    
    If Not WriteHostsFileContents(sb.ToString()) Then
        'Unable to write the selected changes to your hosts file. Another program may be denying access to it, or your user account may have insufficient rights to access it.
        'Make sure you have privileges to modify the hosts file and no program is protecting it against changes.
        MsgBoxW Translate(303) & vbCrLf & Translate(284), vbCritical
        'revert changes
        ListHostsFile objList
    End If
    
    SetFileAttributes StrPtr(g_HostsFile), iAttr
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "HostsToggleLine"
    If inIDE Then Stop: Resume Next
End Sub
