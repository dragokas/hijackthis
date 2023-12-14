Attribute VB_Name = "mFile"
Option Explicit

Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Long
Private Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long

Private Const MAX_PATH = 260
Private Const FILE_SHARE_READ = &H1&
Private Const FILE_SHARE_WRITE = &H2&
Private Const FILE_SHARE_DELETE = 4&
Private Const OPEN_EXISTING = 3&
Private Const INVALID_FILE_ATTRIBUTES = -1&
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10&

Private lWow64Old               As Long

Public Function ReadFileAsBinary(sFileName As String, bin() As Byte) As Boolean
    On Error GoTo ErrorHandler:
    Dim hFile       As Long
    Dim iSize       As Long
    hFile = FreeFile()
    Open sFileName For Binary Access Read As #hFile
        ReDim bin(LOF(hFile) - 1)
        Get #hFile, , bin
    Close #hFile
    ReadFileAsBinary = True
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ReadFileAsBinary"
    If inIDE Then Stop: Resume Next
End Function

Public Function WriteFileBinary(sFileName As String, bin() As Byte) As Boolean
    On Error GoTo ErrorHandler:
    Dim hFile       As Long
    hFile = FreeFile()
    Open sFileName For Binary Access Write Lock Write As #hFile
        Put #hFile, , bin
    Close #hFile
    WriteFileBinary = True
    Exit Function
ErrorHandler:
    ErrorMsg Err, "WriteFileBinary"
    If inIDE Then Stop: Resume Next
End Function

Public Function ReadFileAsText(sFileName As String, out_Text As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim hFile       As Long
    Dim iSize       As Long
    hFile = FreeFile()
    Open sFileName For Binary Access Read As #hFile
        iSize = LOF(hFile)
        out_Text = String$(iSize, 0&)
        Get #hFile, , out_Text
    Close #hFile
    ReadFileAsText = True
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ReadFileAsText"
    If inIDE Then Stop: Resume Next
End Function

Public Function WriteFileText(sFileName As String, sText As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim hFile       As Long
    hFile = FreeFile()
    Open sFileName For Binary Access Write Lock Write As #hFile
        Put #hFile, , sText
    Close #hFile
    WriteFileText = True
    Exit Function
ErrorHandler:
    ErrorMsg Err, "WriteFileText"
    If inIDE Then Stop: Resume Next
End Function


Public Function FileLenW(path As String) As Currency
    Dim lr          As Long
    Dim hFile       As Long
    Dim FileSize    As Currency
    hFile = CreateFile(StrPtr(path), 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    If hFile > 0 Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then FileLenW = FileSize * 10000&
        End If
        CloseHandle hFile
    End If
End Function

Public Function FileExists(ByVal sFile As String, Optional bUseWow64 As Boolean, Optional bAllowNetwork As Boolean) As Boolean
    Dim lr As Long
    Dim bOldStatus As Boolean
    Dim Redirect As Boolean
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    lr = GetFileAttributes(StrPtr(sFile))
    If Err.LastDllError = 5 Or (lr <> INVALID_FILE_ATTRIBUTES And (0 = (lr And FILE_ATTRIBUTE_DIRECTORY))) Then
        FileExists = True
    End If
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
End Function

Public Function RemoveFile(sFile As String) As Boolean
    If Not FileExists(sFile) Then
        RemoveFile = True
    Else
        RemoveFile = (0 <> DeleteFileW(StrPtr(sFile)))
    End If
End Function

Public Sub ReadFileToDictionary(sFile As String, oDict As Object)
    On Error GoTo ErrorHandler:
    Dim ff%
    Dim sLine As String
    ff = FreeFile()
    Open sFile For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, sLine
            If Not oDict.Exists(sLine) Then
                oDict.Add sLine, vbNullString
            End If
        Loop
    Close #ff
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ReadFileToDictionary"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub AppendFileWithCollection(sFile As String, col As Collection)
    On Error GoTo ErrorHandler:
    Dim ff%
    Dim sLine As String
    ff = FreeFile()
    Open sFile For Append As #ff
        Dim i As Long
        For i = 1 To col.Count
            Print #ff, col.Item(i)
        Next
    Close #ff
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AppendFileWithCollection"
    If inIDE Then Stop: Resume Next
End Sub

Public Function ToggleWow64FSRedirection(bEnable As Boolean, Optional PathNecessity As String, Optional OldStatus As Boolean) As Boolean
    'Static lWow64Old        As Long    'Warning: do not use initialized variables for this API !
                                        'Static variables is not allowed !
                                        'lWow64Old is now declared globally
    'True - enable redirector
    'False - disable redirector

    'OldStatus: current state of redirection
    'True - redirector was enabled
    'False - redirector was disabled

    'Return value is:
    'true if success

    Static IsNotRedirected  As Boolean
    Static IsInit           As Boolean
    Static sWinSysDir       As String
    Dim lr                  As Long
    
    OldStatus = Not IsNotRedirected
    
    If Not OSver.IsWin64 Then Exit Function
    If Not IsInit Then
        IsInit = True
        sWinSysDir = String$(MAX_PATH, 0)
        lr = GetSystemWindowsDirectory(StrPtr(sWinSysDir), MAX_PATH)
        If lr Then
            sWinSysDir = Left$(sWinSysDir, lr) & "\System32"
        End If
    End If
    
    If Len(PathNecessity) <> 0 Then
        If StrComp(Left$(Replace(Replace(PathNecessity, "/", "\"), "\\", "\"), Len(sWinSysDir)), sWinSysDir, vbTextCompare) <> 0 Then Exit Function
    End If
    
    If bEnable Then
        If IsNotRedirected Then
            lr = Wow64RevertWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = False
        End If
    Else
        If Not IsNotRedirected Then
            lr = Wow64DisableWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = True
        End If
    End If
    
End Function

