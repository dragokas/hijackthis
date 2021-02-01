Attribute VB_Name = "modVirusTotal"
Option Explicit

Public Function DownloadAutoruns() As Boolean
    
    On Error GoTo ErrorHandler:
    
    Const sURL As String = "https://download.sysinternals.com/files/Autoruns.zip"
    
    Dim bHasFile As Boolean
    Dim ToolsDir As String
    Dim sAutorunsExePath As String
    Dim bRequireDL As Boolean
    Dim ArcPath As String
    Dim UnpackPath As String
    Dim sAutorunsInZip As String
    
    ToolsDir = GetToolsDir()
    sAutorunsExePath = GetAutorunsPath()
    ArcPath = BuildPath(TempCU, "Autoruns.zip")
    UnpackPath = BuildPath(TempCU, "Autoruns")
    
    If Not FileExists(sAutorunsExePath) Then
        bRequireDL = True
    Else
        If bCheckForUpdates Then
            If DateDiff("d", GetFileDate(sAutorunsExePath, DATE_CREATED), Now()) > 30 Then ' 1 month elapsed
                bRequireDL = True
            End If
        End If
        
        If Not bRequireDL Then
            If OSver.IsWindowsVistaOrGreater Then
                If Not IsMicrosoftFile(sAutorunsExePath) Then
                    MsgBoxW "The following file didn't pass a signature verification:" & vbCrLf & sAutorunsExePath, vbExclamation
                    bRequireDL = True
                End If
            End If
        End If
    End If
    
    If Not bRequireDL Then
        DownloadAutoruns = True
        Exit Function
    End If
    
    If Not bUpdateSilently Then
        'I need to download SysInternals Autoruns to: [] Do you allow me?
        If MsgBoxW(Replace$(Translate(2350), "[]", sAutorunsExePath), vbYesNo Or vbQuestion) = vbNo Then
            Exit Function
        End If
    End If
    
    If Not FolderExists(GetParentDir(ArcPath)) Then 'if "%temp%" doesn't exist for some reason
        If Not MkDirW(GetParentDir(ArcPath)) Then
            MsgBoxW "Cannot create the folder:" & vbCrLf & GetParentDir(ArcPath), vbCritical
            Exit Function
        End If
    End If
    
    If DownloadFile(sURL, ArcPath, True) Then
        
        If FolderExists(UnpackPath) Then DeleteFolder UnpackPath
        
        If Not MkDirW(UnpackPath) Then
            MsgBoxW "Cannot create the folder:" & vbCrLf & UnpackPath, vbCritical
            Exit Function
        End If
        
        If Not UnpackZIP(ArcPath, UnpackPath) Then
            MsgBoxW "Cannot unpack the archive:" & vbCrLf & ArcPath, vbCritical
            Exit Function
        End If
        
        If Not FolderExists(ToolsDir) Then
            If Not MkDirW(ToolsDir) Then
                MsgBoxW "Cannot create the folder:" & vbCrLf & ToolsDir, vbCritical
                Exit Function
            End If
        End If
        
        DeleteFileWEx StrPtr(sAutorunsExePath), True, True
        
        If FileExists(sAutorunsExePath) Then
            MsgBoxW "Cannot remove the old file:" & vbCrLf & sAutorunsExePath, vbCritical
            Exit Function
        End If
        
        sAutorunsInZip = BuildPath(UnpackPath, IIf(OSver.IsWin64, "autorunsc64.exe", "autorunsc.exe"))
        
        If Not FileExists(sAutorunsInZip) Then
            MsgBoxW "Cannot find the file:" & vbCrLf & sAutorunsInZip, vbCritical
            Exit Function
        End If
        
        If Not FileCopyW(sAutorunsInZip, sAutorunsExePath) Then
            MsgBoxW "Cannot copy the file to:" & vbCrLf & sAutorunsExePath, vbCritical
            Exit Function
        End If
    Else
        MsgBoxW "Cannot download:" & vbCrLf & sURL, vbCritical
        Exit Function
    End If
    
    If OSver.IsWindowsVistaOrGreater Then
        If Not IsMicrosoftFile(sAutorunsExePath) Then
            MsgBoxW "The following file didn't pass a signature verification:" & vbCrLf & sAutorunsExePath, vbCritical
            Exit Function
        End If
    End If
    
    'Clear
    DeleteFolder UnpackPath
    DeleteFileWEx StrPtr(ArcPath)
    
    DownloadAutoruns = True
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DownloadAutoruns"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetToolsDir() As String
    
    Dim sToolsParentDir As String
    
    If IsInstalledHJT() Then
        sToolsParentDir = GetInstDir()
    Else
        sToolsParentDir = AppPath()
    End If
    
    GetToolsDir = BuildPath(sToolsParentDir, "tools\Scan")
    
End Function

Public Function GetAutorunsPath() As String
    
    Dim sAutorunsExe As String
    Dim sAutorunsExePath As String
    Dim ToolsDir As String
    
    ToolsDir = GetToolsDir()
    sAutorunsExe = IIf(OSver.IsWin64, "auto64.exe", "auto.exe")
    GetAutorunsPath = BuildPath(ToolsDir, sAutorunsExe)
End Function

' For 1 file only!
' For multiple files - need async / using other key (e.g. Run)
'
Public Function AR_CheckFile(sFile As String, Optional bSilent As Boolean) As Boolean
    
    On Error GoTo ErrorHandler:
    
    Dim AppInitBak As String
    Dim sAutorunsExePath As String
    Dim sSha256 As String
    
    sAutorunsExePath = GetAutorunsPath()
    
    'preserve the old value
    AppInitBak = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs")
    
    'temporarily substitute the new one for Autoruns
    Reg.SetStringVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", sFile
    
    Set Proc = New clsProcess
    If Proc.ProcessRun(sAutorunsExePath, "-accepteula -a d -vs -vt -nobanner -x", , vbHide, True) Then
        g_bVTScanInProgress = True
        g_bCalcHashInProgress = True
        If Not g_bScanInProgress Then
            frmMain.lblStatus.Visible = True
            frmMain.lblStatus.ForeColor = vbDarkRed
            frmMain.lblStatus.Font.Bold = True
            frmMain.lblStatus.Caption = "Virustotal: " & GetFileNameAndExt(sFile) & " - " & GetParentDir(sFile) & "\"
        End If
        ResumeHashProgressbar
        UpdateVTProgressbar False
        SetHashProgressBar 33
        Sleep 500
        SetHashProgressBar 66
        frmMain.tmrVTProgress.Enabled = True
        AR_CheckFile = True
    End If
    
    'restore the initial value
    Reg.SetStringVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", AppInitBak
    
    If Proc.pid = 0 Then
        If Not bSilent Then
            MsgBoxW "Error while submitting the file with 'AutoRuns':" & vbCrLf & sFile & vbCrLf & vbCrLf & "Code: " & Err.LastDllError
        End If
        Exit Function
    End If
    
    sSha256 = GetFileSHA256(sFile, , True)
    frmMain.lblMD5.Tag = "https://www.virustotal.com/gui/file/" & sSha256 & "/detection"
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "AR_CheckFile", sFile
    If inIDE Then Stop: Resume Next
End Function

Public Sub UpdateVTProgressbar(bFinished As Boolean)
    
    If bFinished Then
        SetHashProgressBar 100, Translate(2351) ' "Click here to get VT results"
        
        frmMain.lblMD5.Font.Bold = True
        frmMain.lblMD5.Font.Underline = True
    Else
        SetHashProgressBar 66, Translate(2352) ' "Uploading the file ..."
    End If
    
End Sub
