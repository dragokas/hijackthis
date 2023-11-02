Attribute VB_Name = "mMain"
Option Explicit

Public Sub Main()
    On Error GoTo ErrorHandler

    Dim Args()      As String
    Dim sFileName   As String
    Dim sTargetFile As String
    Dim sAction     As String
    Dim sType       As String
    Dim sCmdLine    As String
    Dim bSuccess    As Boolean
    
    Init
    
    sCmdLine = Command$()
    'sCmdLine = "encrypt text xx xx.cry"
    'sCmdLine = "decrypt text xx xx.cry"
    'sCmdLine = "encrypt binary xx xx.cry"
    'sCmdLine = "decrypt binary xx xx.cry"
    
    ParseCommandLine sCmdLine, Args()
    
    If UBound(Args) < 4 Then Using: ExitProcessVB 1
  
    sAction = Args(1)
    sType = Args(2)
    sFileName = Args(3)
    sTargetFile = Args(4)
    
    If Not FileExists(sFileName) Then
        WriteStderr "No such file: " & sFileName
        ExitProcessVB 1
    End If
    
    If FileLenW(sFileName) = 0 Then
        WriteStderr "File size is zero: " & sFileName
        ExitProcessVB 1
    End If
    
    If FileExists(sTargetFile) Then
        If Not mFile.RemoveFile(sTargetFile) Then
            WriteStderr "Cannot delete file: " & sTargetFile
            ExitProcessVB 1
        End If
    End If
    
    If StrComp(sAction, "encrypt", 1) = 0 Then
        If StrComp(sType, "binary", 1) = 0 Then
            bSuccess = EncryptBinary(sFileName, sTargetFile)
        ElseIf StrComp(sType, "text", 1) = 0 Then
            bSuccess = EncryptText(sFileName, sTargetFile)
        Else
            WriteStderr "Unknown argument: " & sType
            ExitProcessVB 1
        End If
    ElseIf StrComp(sAction, "decrypt", 1) = 0 Then
        If StrComp(sType, "binary", 1) = 0 Then
            bSuccess = DecryptBinary(sFileName, sTargetFile)
        ElseIf StrComp(sType, "text", 1) = 0 Then
            bSuccess = DecryptText(sFileName, sTargetFile)
        Else
            WriteStderr "Unknown argument: " & sType
            ExitProcessVB 1
        End If
    Else
        WriteStderr "Unknown argument: " & sAction
        ExitProcessVB 1
    End If
    
    If bSuccess Then
        WriteStdout sAction & "(" & sType & ") - " & sFileName & " - Success."
        ExitProcessVB 0
    Else
        WriteStderr sAction & "(" & sType & ") - " & sFileName & " - Failed!"
        ExitProcessVB 1
    End If
    
    Exit Sub
ErrorHandler:
    WriteStderr "Error #" & Err.Number & ". LastDll=" & Err.LastDllError & ". " & Err.Description
    ExitProcessVB 1
End Sub

Private Function EncryptBinary(sFileName As String, sDestinationFile As String) As Boolean
    Dim buf() As Byte
    If mFile.ReadFileAsBinary(sFileName, buf) Then
        Call Caesar_EncodeBin(buf)
        If mFile.WriteFileBinary(sDestinationFile, buf) Then EncryptBinary = True
    End If
End Function

Private Function DecryptBinary(sFileName As String, sDestinationFile As String) As Boolean
    Dim buf() As Byte
    If mFile.ReadFileAsBinary(sFileName, buf) Then
        Call Caesar_DecodeBin(buf)
        If mFile.WriteFileBinary(sDestinationFile, buf) Then DecryptBinary = True
    End If
End Function

Private Function EncryptText(sFileName As String, sDestinationFile As String) As Boolean
    Dim buf As String
    If mFile.ReadFileAsText(sFileName, buf) Then
        buf = Caesar_Encode(buf)
        If mFile.WriteFileText(sDestinationFile, buf) Then EncryptText = True
    End If
End Function

Private Function DecryptText(sFileName As String, sDestinationFile As String) As Boolean
    Dim buf As String
    If mFile.ReadFileAsText(sFileName, buf) Then
        buf = Caesar_Decode(buf)
        If mFile.WriteFileText(sDestinationFile, buf) Then DecryptText = True
    End If
End Function

Private Sub Using()
    WriteStderr "Resource files encryptor by caesar algorithm"
    WriteStderr ""
    WriteStderr "Using:"
    WriteStderr App.EXEName & ".exe [encrypt/decrypt] [binary/text] [source path] [destination path]"
End Sub
