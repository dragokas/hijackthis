Attribute VB_Name = "mFile"
Option Explicit

Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

Private Const FILE_SHARE_READ = &H1&
Private Const FILE_SHARE_WRITE = &H2&
Private Const FILE_SHARE_DELETE = 4&
Private Const OPEN_EXISTING = 3&
Private Const INVALID_FILE_ATTRIBUTES = -1&
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10&

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
    If inIde Then Stop: Resume Next
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
    If inIde Then Stop: Resume Next
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
    If inIde Then Stop: Resume Next
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
    If inIde Then Stop: Resume Next
End Function


Public Function FileLenW(Path As String) As Currency
    Dim lr          As Long
    Dim hFile       As Long
    Dim FileSize    As Currency
    hFile = CreateFile(StrPtr(Path), 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
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
    lr = GetFileAttributes(StrPtr(sFile))
    If Err.LastDllError = 5 Or (lr <> INVALID_FILE_ATTRIBUTES And (0 = (lr And FILE_ATTRIBUTE_DIRECTORY))) Then
        FileExists = True
    End If
End Function

Public Function RemoveFile(sFile As String) As Boolean
    If Not FileExists(sFile) Then
        RemoveFile = True
    Else
        RemoveFile = (0 <> DeleteFileW(StrPtr(sFile)))
    End If
End Function
