Attribute VB_Name = "modHosts"
Option Explicit

Public Sub ListHostsFile(objList As ListBox, objInfo As Label)
    'custom hosts file handling?
    Dim sAttr$, iAttr%, sDummy$, vContent As Variant, i&
    On Error Resume Next
    objInfo.Caption = "Loading hosts file, please wait..."
    frmMain.cmdHostsManDel.Enabled = False
    frmMain.cmdHostsManToggle.Enabled = False
    DoEvents
    If Not FileExists(sHostsFile) Then
        If MsgBox("Cannot find the hosts file." & vbCrLf & _
                  "Do you want to create a new, default " & _
                  "hosts file?", vbExclamation + vbYesNo) = vbNo Then
            objInfo.Caption = "No hosts file found."
            Exit Sub
        Else
            CreateDefaultHostsFile
        End If
    End If
    
    objInfo.Caption = "Loading hosts file, please wait..."
    frmMain.cmdHostsManDel.Enabled = False
    frmMain.cmdHostsManToggle.Enabled = False
    DoEvents
    iAttr = GetAttr(sHostsFile)
    If (iAttr And 1) Then sAttr = sAttr & "R"
    If (iAttr And 32) Then sAttr = sAttr & "A"
    If (iAttr And 2) Then sAttr = sAttr & "H"
    If (iAttr And 4) Then sAttr = sAttr & "S"
    If (iAttr And 2048) Then sAttr = sAttr & "C"
    
    Open sHostsFile For Binary As #1
        sDummy = Input(FileLen(sHostsFile), #1)
    Close #1
    vContent = Split(sDummy, vbCrLf)
    If UBound(vContent) = 0 And InStr(vContent(0), Chr(10)) > 0 Then
        'unix style hosts file
        vContent = Split(sDummy, Chr(10))
    End If
    
    objList.Clear
    For i = 0 To UBound(vContent)
        objList.AddItem CStr(vContent(i))
    Next i

    'objInfo.Caption = "Hosts file is located at: " & sHostsFile &
    objInfo.Caption = Translate(271) & " " & sHostsFile & _
                      " (" & objList.ListCount & " lines, " & _
                      sAttr & ")"
    frmMain.cmdHostsManDel.Enabled = True
    frmMain.cmdHostsManToggle.Enabled = True
End Sub

Public Sub CreateDefaultHostsFile()
    On Error Resume Next
    Open sHostsFile For Output As #1
        Print #1, "# Copyright (c) 1993-2009 Microsoft Corp."
        Print #1, "#"
        Print #1, "# This is a sample HOSTS file used by Microsoft TCP/IP for Windows."
        Print #1, "#"
        Print #1, "# This file contains the mappings of IP addresses to host names. Each"
        Print #1, "# entry should be kept on an individual line. The IP address should"
        Print #1, "# be placed in the first column followed by the corresponding host name."
        Print #1, "# The IP address and the host name should be separated by at least one"
        Print #1, "# space."
        Print #1, "#"
        Print #1, "# Additionally, comments (such as these) may be inserted on individual"
        Print #1, "# lines or following the machine name denoted by a '#' symbol."
        Print #1, "#"
        Print #1, "# For example:"
        Print #1, "#"
        Print #1, "#      102.54.94.97     rhino.acme.com          # source server"
        Print #1, "#       38.25.63.10     x.acme.com              # x client host"
        Print #1,
        Print #1, "127.0.0.1       localhost"
        Print #1, "::1             localhost"
    Close #1
End Sub

Public Sub HostsDeleteLine(objList As ListBox)
    'delete ith line in hosts file (zero-based)
    Dim iAttr%, sDummy$, vContent As Variant, i&
    On Error Resume Next
    iAttr = GetAttr(sHostsFile)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetAttr sHostsFile, vbArchive
    If Err Then
        MsgBox "The hosts file is locked for reading and cannot be edited. " & vbCrLf & _
               "Make sure you have privileges to modify the hosts file and " & _
               "no program is protecting it against changes.", vbCritical
        Exit Sub
    End If
    
    Open sHostsFile For Binary As #1
        sDummy = Input(FileLen(sHostsFile), #1)
    Close #1
    vContent = Split(sDummy, vbCrLf)
    If UBound(vContent) = 0 And InStr(vContent(0), Chr(10)) > 0 Then
        'unix style hosts file
        vContent = Split(sDummy, Chr(10))
    End If
    
    Open sHostsFile For Output As #1
        With objList
            For i = 0 To UBound(vContent) - 1
                If Not .Selected(i) Then Print #1, vContent(i)
            Next i
        End With
    Close #1
End Sub

Public Sub HostsToggleLine(objList As ListBox)
    'enable/disable ith line in hosts file (zero-based)
    Dim iAttr%, sDummy$, vContent As Variant, i&
    On Error Resume Next
    iAttr = GetAttr(sHostsFile)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetAttr sHostsFile, vbArchive
    If Err Then
        MsgBox "The hosts file is locked for reading and cannot be edited. " & vbCrLf & _
               "Make sure you have privileges to modify the hosts file and " & _
               "no program is protecting it against changes.", vbCritical
        Exit Sub
    End If
    
    Open sHostsFile For Binary As #1
        sDummy = Input(FileLen(sHostsFile), #1)
    Close #1
    vContent = Split(sDummy, vbCrLf)
    If UBound(vContent) = 0 And InStr(vContent(0), Chr(10)) > 0 Then
        'unix style hosts file
        vContent = Split(sDummy, Chr(10))
    End If
    
    With objList
        For i = 0 To UBound(vContent)
            If .Selected(i) Then
                If InStr(vContent(i), "#") = 1 Then
                    vContent(i) = Mid(vContent(i), 2)
                Else
                    vContent(i) = "#" & vContent(i)
                End If
            End If
        Next i
    End With
    
    Open sHostsFile For Output As #1
        Print #1, Join(vContent, vbCrLf)
    Close #1
    SetAttr sHostsFile, iAttr
End Sub
