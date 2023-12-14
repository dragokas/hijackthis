Attribute VB_Name = "mMain"
Option Explicit

Public OSver As clsOSInfo

Public Sub Main()
    On Error GoTo ErrorHandler

    Dim Args()      As String
    Dim sTargetFile As String
    Dim sCmdLine    As String
    Dim sJson       As String
    Dim sURL        As String
    Dim oPrevBase   As Object
    Dim listPath    As Collection
    Dim listNewPath As Collection
    Dim sPath       As String
    Dim i           As Long
    Dim countExist  As Long
    Dim bExist      As Boolean
    
    sURL = "https://lolbas-project.github.io/api/lolbas.json"
    sJson = App.path & "\lolbas.json"
    
    Init
    Set oPrevBase = CreateObject("Scripting.Dictionary")
    
    sCmdLine = Command$()
    'sCmdLine = "..\..\database\LoLBin_Protect.txt"
    
    ParseCommandLine sCmdLine, Args()
    
    If UBound(Args) < 1 Then Using: ExitProcessVB 1

    sTargetFile = Args(1)
    
    If FileExists(sTargetFile) Then
        ReadFileToDictionary sTargetFile, oPrevBase
    End If
    
    WriteStdout "Downloading file: " & sURL
    
    If Not DownloadFile(sURL, sJson) Then
        WriteStderr "Failed to download the file!"
        ExitProcessVB 1
    End If
    
    ParseJsonFile_ToLoLbinList sJson, listPath
    
    Set listNewPath = New Collection
    
    If oPrevBase.Count = 0 Then
        'missing on lolbas-project.github.io
        listNewPath.Add "<SysRoot>\system32\svchost.exe"
        listNewPath.Add "<SysRoot>\System32\windowspowershell\v1.0\powershell.exe"
        listNewPath.Add "<PF64>\Windows Defender\MpCmdRun.exe"
    End If
    
    For i = 1 To listPath.Count
    
        sPath = listPath.Item(i)
        
        If IsPath(sPath) And _
          Not IsRandomPath(sPath) And _
          IsExecutablePath(sPath) Then
            
            sPath = NormalizePath(sPath)
            bExist = mFile.FileExists(sPath)
            sPath = Unexpand(sPath)
            
            If Not oPrevBase.Exists(sPath) Then
            
                listNewPath.Add sPath
                
                If bExist Then
                    countExist = countExist + 1
                    WriteStdout "{+} Found new entry: " & sPath
                Else
                    WriteStdout "{-} Found new entry: " & sPath
                End If
            End If
            
        End If
    Next
    
    If listNewPath.Count = 0 Then
        WriteStdout "Database is up to date."
        ExitProcessVB 0
    End If
    
    WriteStdout "{NOTIFY} " & countExist & " of " & listNewPath.Count & " new entries are present in your file system."
    
    Dim ch$: ch = ReadStdin("Do you want to update the database? (Y/n) ")
    If StrComp(ch, "Y", 1) <> 0 Then
        ExitProcessVB 0
    End If
    
    mFile.AppendFileWithCollection sTargetFile, listNewPath
    
    WriteStdout listNewPath.Count & " entries have been added to the database."
    
    Exit Sub
ErrorHandler:
    WriteStderr "Error #" & Err.Number & ". LastDll=" & Err.LastDllError & ". " & Err.Description
    ExitProcessVB 1
End Sub

Private Sub ParseJsonFile_ToLoLbinList(sFile As String, col As Collection)
    On Error GoTo ErrorHandler
    
    Dim JS As JsonBag
    Set JS = New JsonBag
    
    Set col = New Collection
    
    If JS.fromFile(sFile) Then
        Dim record As JsonBag
        Dim path As JsonBag
        For Each record In JS
            For Each path In record.ItemSafe("Full_Path")
                col.Add path.ItemSafe("Path")
            Next
        Next
    Else
        WriteStderr "Failed to read json contents!"
    End If
    
    Exit Sub
ErrorHandler:
    WriteStderr "Error parsing json file #" & Err.Number & ". LastDll=" & Err.LastDllError & ". " & Err.Description
    ExitProcessVB 1
End Sub

Private Sub Using()
    WriteStderr "LolBin database updater for HijackThis+"
    WriteStderr ""
    WriteStderr "Based on service: https://lolbas-project.github.io/"
    WriteStderr "Thanks to:"
    WriteStderr " - Oddvar Moe (@oddvarmoe)"
    WriteStderr " - Jimmy Bayne (@bohops)"
    WriteStderr " - Conor Richard (@xenosCR)"
    WriteStderr " - Chris 'Lopi' Spehn (@ConsciousHacker)"
    WriteStderr " - Liam (@liamsomerville)"
    WriteStderr " - Wietze (@Wietze)"
    WriteStderr " - Jose Hernandez (@_josehelps)"
    WriteStderr " - and everyone who contributed"
    WriteStderr ""
    WriteStderr "Using:"
    WriteStderr App.EXEName & ".exe [database path]"
End Sub

Private Function NormalizePath(sPath) As String
    NormalizePath = Replace$(sPath, "\\", "\")
End Function

Private Function IsRandomPath(sPath As String) As Boolean
    If InStr(1, sPath, "*") <> 0 Then
        IsRandomPath = True
    ElseIf InStr(1, sPath, "[") <> 0 Then
        IsRandomPath = True
    ElseIf InStr(1, sPath, "XXX", vbTextCompare) <> 0 Then
        IsRandomPath = True
    End If
End Function

Private Function IsPath(sPath As String) As Boolean
    IsPath = InStr(1, sPath, ":") <> 0
End Function

Private Function IsExecutablePath(sPath As String) As Boolean
    IsExecutablePath = StrEndWith(sPath, ".exe")
End Function
