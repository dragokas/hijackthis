Attribute VB_Name = "modHosts"
'[modHosts.bas]

'
' Hosts file module by Merijn Bellekom
'

Option Explicit

'Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
'Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long


Public Sub ListHostsFile(objList As ListBox, objInfo As Label)
    On Error GoTo ErrorHandler:

    'custom hosts file handling?
    Dim sAttr$, iAttr&, sDummy$, vContent As Variant, i&, ff%
    'On Error Resume Next
    'objInfo.Caption = "Loading hosts file, please wait..."
    objInfo.Caption = Translate(279)
    frmMain.cmdHostsManDel.Enabled = False
    frmMain.cmdHostsManToggle.Enabled = False
    DoEvents
    If Not FileExists(sHostsFile) Then
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
    frmMain.cmdHostsManDel.Enabled = False
    frmMain.cmdHostsManToggle.Enabled = False
    DoEvents
    iAttr = GetFileAttributes(StrPtr(sHostsFile))
    If (iAttr And 1) Then sAttr = sAttr & "R"
    If (iAttr And 32) Then sAttr = sAttr & "A"
    If (iAttr And 2) Then sAttr = sAttr & "H"
    If (iAttr And 4) Then sAttr = sAttr & "S"
    If (iAttr And 2048) Then sAttr = sAttr & "C"
    
    ff = FreeFile()
    Open sHostsFile For Binary As #ff
        sDummy = Input(FileLenW(sHostsFile), #ff)
    Close #ff
    vContent = Split(sDummy, vbCrLf)
    If UBound(vContent) = 0 And InStr(vContent(0), Chr$(10)) > 0 Then
        'unix style hosts file
        vContent = Split(sDummy, Chr$(10))
    End If
    
    objList.Clear
    For i = 0 To UBound(vContent)
        objList.AddItem CStr(vContent(i))
    Next i

    'objInfo.Caption = "Hosts file is located at " & sHostsFile & "Line: [], Attributes: []"
    objInfo.Caption = Translate(271) & " " & sHostsFile & _
                      " (" & Translate(278) & " " & objList.ListCount & ", " & Translate(277) & " " & _
                      sAttr & ")"
    frmMain.cmdHostsManDel.Enabled = True
    frmMain.cmdHostsManToggle.Enabled = True
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ListHostsFile"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CreateDefaultHostsFile()
    On Error GoTo ErrorHandler:
    
    Dim ff%
    Open sHostsFile For Output As #ff
        Print #ff, GetDefaultHostsContents()
    Close #ff
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CreateDefaultHostsFile"
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetDefaultHostsContents() As String
  Dim DefaultContent$
    
  If OSver.MajorMinor < 6 Then
    
    'XP
    DefaultContent = _
    "# Copyright (c) 1993-1999 Microsoft Corp." & vbCrLf & _
    "#" & vbCrLf & _
    "# This is a sample HOSTS file used by Microsoft TCP/IP for Windows." & vbCrLf & _
    "#" & vbCrLf & _
    "# This file contains the mappings of IP addresses to host names. Each" & vbCrLf & _
    "# entry should be kept on an individual line. The IP address should" & vbCrLf & _
    "# be placed in the first column followed by the corresponding host name." & vbCrLf & _
    "# The IP address and the host name should be separated by at least one" & vbCrLf & _
    "# space." & vbCrLf & _
    "#" & vbCrLf & _
    "# Additionally, comments (such as these) may be inserted on individual" & vbCrLf & _
    "# lines or following the machine name denoted by a '#' symbol." & vbCrLf & _
    "#" & vbCrLf & _
    "# For example:" & vbCrLf & _
    "#" & vbCrLf & _
    "#      102.54.94.97     rhino.acme.com          # source server" & vbCrLf & _
    "#       38.25.63.10     x.acme.com              # x client host" & vbCrLf & _
    "" & vbCrLf & _
    "127.0.0.1       localhost"
    
  ElseIf OSver.MajorMinor < 6.1 Then
    
    'Vista
    DefaultContent = _
    "# Copyright (c) 1993-2006 Microsoft Corp." & vbCrLf & _
    "#" & vbCrLf & _
    "# This is a sample HOSTS file used by Microsoft TCP/IP for Windows." & vbCrLf & _
    "#" & vbCrLf & _
    "# This file contains the mappings of IP addresses to host names. Each" & vbCrLf & _
    "# entry should be kept on an individual line. The IP address should" & vbCrLf & _
    "# be placed in the first column followed by the corresponding host name." & vbCrLf & _
    "# The IP address and the host name should be separated by at least one" & vbCrLf & _
    "# space." & vbCrLf & _
    "#" & vbCrLf & _
    "# Additionally, comments (such as these) may be inserted on individual" & vbCrLf & _
    "# lines or following the machine name denoted by a '#' symbol." & vbCrLf & _
    "#" & vbCrLf & _
    "# For example:" & vbCrLf & _
    "#" & vbCrLf & _
    "#      102.54.94.97     rhino.acme.com          # source server" & vbCrLf & _
    "#       38.25.63.10     x.acme.com              # x client host" & vbCrLf & _
    "" & vbCrLf & _
    "127.0.0.1       localhost" & vbCrLf & _
    "::1             localhost"
  
  Else
  
    '7 and higher (Win 10 checked)
    DefaultContent = _
    "# Copyright (c) 1993-2009 Microsoft Corp." & vbCrLf & _
    "#" & vbCrLf & _
    "# This is a sample HOSTS file used by Microsoft TCP/IP for Windows." & vbCrLf & _
    "#" & vbCrLf & _
    "# This file contains the mappings of IP addresses to host names. Each" & vbCrLf & _
    "# entry should be kept on an individual line. The IP address should" & vbCrLf & _
    "# be placed in the first column followed by the corresponding host name." & vbCrLf & _
    "# The IP address and the host name should be separated by at least one" & vbCrLf & _
    "# space." & vbCrLf & _
    "#" & vbCrLf & _
    "# Additionally, comments (such as these) may be inserted on individual" & vbCrLf & _
    "# lines or following the machine name denoted by a '#' symbol." & vbCrLf & _
    "#" & vbCrLf & _
    "# For example:" & vbCrLf & _
    "#" & vbCrLf & _
    "#      102.54.94.97     rhino.acme.com          # source server" & vbCrLf & _
    "#       38.25.63.10     x.acme.com              # x client host" & vbCrLf & _
    "" & vbCrLf & _
    "# localhost name resolution is handled within DNS itself." & vbCrLf & _
    "#   127.0.0.1       localhost" & vbCrLf & _
    "#   ::1             localhost"
    
  End If

  GetDefaultHostsContents = DefaultContent

End Function


Public Sub HostsDeleteLine(objList As ListBox)
    On Error GoTo ErrorHandler:

    'delete ith line in hosts file (zero-based)
    Dim iAttr&, sDummy$, vContent As Variant, i&, ff%
    
    iAttr = GetFileAttributes(StrPtr(sHostsFile))
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetFileAttributes StrPtr(sHostsFile), vbArchive
    If Err.Number Then
        'MsgBoxW "The hosts file is locked for reading and cannot be edited. " & vbCrLf & _
        '       "Make sure you have privileges to modify the hosts file and " & _
        '       "no program is protecting it against changes.", vbCritical
        MsgBoxW Translate(282), vbCritical
        Exit Sub
    End If
    
    ff = FreeFile()
    Open sHostsFile For Binary As #ff
        sDummy = Input(FileLenW(sHostsFile), #ff)
    Close #ff
    vContent = Split(sDummy, vbCrLf)
    If UBound(vContent) = 0 And InStr(vContent(0), Chr$(10)) > 0 Then
        'unix style hosts file
        vContent = Split(sDummy, Chr$(10))
    End If
    
    ff = FreeFile()
    Open sHostsFile For Output As #ff
        With objList
            For i = 0 To UBound(vContent) - 1
                If Not .Selected(i) Then Print #ff, vContent(i)
            Next i
        End With
    Close #ff
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "HostsDeleteLine"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub HostsToggleLine(objList As ListBox)
    On Error GoTo ErrorHandler:

    'enable/disable ith line in hosts file (zero-based)
    Dim iAttr&, sDummy$, vContent As Variant, i&, ff%
    
    iAttr = GetFileAttributes(StrPtr(sHostsFile))
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetFileAttributes StrPtr(sHostsFile), vbArchive
    If Err.Number Then
        '"The hosts file is locked for reading and cannot be edited. " & vbCrLf & _
               "Make sure you have privileges to modify the hosts file and " & _
               "no program is protecting it against changes."
        MsgBoxW Translate(282), vbCritical
        Exit Sub
    End If
    
    ff = FreeFile()
    Open sHostsFile For Binary As #ff
        sDummy = Input(FileLenW(sHostsFile), #ff)
    Close #ff
    vContent = Split(sDummy, vbCrLf)
    If UBound(vContent) = 0 And InStr(vContent(0), Chr$(10)) > 0 Then
        'unix style hosts file
        vContent = Split(sDummy, Chr$(10))
    End If
    
    With objList
        For i = 0 To UBound(vContent)
            If .Selected(i) Then
                If InStr(vContent(i), "#") = 1 Then
                    vContent(i) = Mid$(vContent(i), 2)
                Else
                    vContent(i) = "#" & vContent(i)
                End If
            End If
        Next i
    End With
    
    ff = FreeFile()
    Open sHostsFile For Output As #ff
        Print #ff, Join(vContent, vbCrLf)
    Close #ff
    SetFileAttributes StrPtr(sHostsFile), iAttr
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "HostsToggleLine"
    If inIDE Then Stop: Resume Next
End Sub
