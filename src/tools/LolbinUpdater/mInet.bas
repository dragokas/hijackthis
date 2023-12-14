Attribute VB_Name = "mInet"
Option Explicit

Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

Public Function DownloadFile(sURL As String, sTargetFile As String) As Boolean
    On Error GoTo ErrorHandler
    Dim oHttp As Object
    Set oHttp = CreateObject("Microsoft.XMLHTTP")
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    With oHttp
        .Open "GET", sURL, False
        .Send
        If .Status <> 200 Then
            WriteStderr "Download status: " & .Status
            Exit Function
        End If
        DeleteFileW StrPtr(sTargetFile)
        With oStream
            .Type = 1 ' binary
            .Open
            .Write oHttp.responseBody
            .SaveToFile sTargetFile, 2 ' overwrite
        End With
    End With
    DownloadFile = True
    Exit Function
ErrorHandler:
    WriteStderr "Error in DownloadFile #" & Err.Number & ". LastDll=" & Err.LastDllError & ". " & Err.Description
    ExitProcessVB 1
End Function
