Attribute VB_Name = "modFile"
Option Explicit

Private Const MAX_PATH As Long = 260&

Enum VB_FILE_ACCESS_MODE
    FOR_READ = 1
    FOR_READ_WRITE = 2
    FOR_OVERWRITE_CREATE = 4
End Enum

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long

Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long

Private Const INVALID_HANDLE_VALUE          As Long = -1&
Private Const ERROR_INSUFFICIENT_BUFFER     As Long = 122&
Private Const GENERIC_READ                  As Long = &H80000000
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const FILE_READ_ATTRIBUTES          As Long = &H80&
Private Const FILE_SHARE_READ               As Long = 1&
Private Const FILE_SHARE_WRITE              As Long = 2&
Private Const FILE_SHARE_DELETE             As Long = 4&
Private Const OPEN_EXISTING                 As Long = 3&
Private Const CREATE_ALWAYS                 As Long = 2&
Private Const INVALID_SET_FILE_POINTER      As Long = &HFFFFFFFF
Private Const FILE_BEGIN                    As Long = 0&
Private Const FILE_END                      As Long = 2&
Private Const NO_ERROR                      As Long = 0&

Function FileLenW(Optional Path As String, Optional hFileHandle As Long) As Currency
    Dim lr          As Long
    Dim hFile       As Long
    Dim FileSize    As Currency
    
    If hFileHandle = 0 Then
        hFile = CreateFile(StrPtr(Path), FILE_READ_ATTRIBUTES, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        hFile = hFileHandle
    End If

    If hFile > 0 Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then FileLenW = FileSize * 10000&
        End If
        If hFileHandle = 0 Then CloseHandle hFile
    End If
End Function


Public Function OpenW(Filename As String, Access As VB_FILE_ACCESS_MODE, retHandle As Long) As Boolean
    
    Dim FSize As Currency
    
    If Access And (FOR_READ Or FOR_READ_WRITE) Then
        If Not FileExists(Filename) Then
            retHandle = INVALID_HANDLE_VALUE
            Exit Function
        End If
    End If
    
    If Access = FOR_READ Then
        retHandle = CreateFile(StrPtr(Filename), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_OVERWRITE_CREATE Then
        retHandle = CreateFile(StrPtr(Filename), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_READ_WRITE Then
        retHandle = CreateFile(StrPtr(Filename), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        Debug.Print "Wrong access mode!"
    End If
    
    OpenW = ((INVALID_HANDLE_VALUE <> retHandle) And (retHandle <> 0))
    
End Function
                                                                  'do not change Variant type at all or you will die ^_^
Public Function GetW(hFile As Long, ByVal pos As Long, Optional vOut As Variant, Optional vOutPtr As Long, Optional cbToRead As Long) As Boolean
    Dim lBytesRead  As Long
    Dim lr          As Long
    Dim ptr         As Long
    Dim vType       As Long
    Dim UnknType    As Boolean
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If INVALID_SET_FILE_POINTER <> SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then
        If NO_ERROR = Err.LastDllError Then
            vType = VarType(vOut)
            
            If 0 <> cbToRead Then   'vbError = vType
                lr = ReadFile(hFile, vOutPtr, cbToRead, lBytesRead, 0&)
                
            ElseIf vbString = vType Then
                lr = ReadFile(hFile, StrPtr(vOut), Len(vOut), lBytesRead, 0&)
                If Err.LastDllError <> 0 Or lr = 0 Then Err.Raise 52, , "Cannot read file! Handle: " & hFile
                
                vOut = StrConv(vOut, vbUnicode)
                If Len(vOut) <> 0 Then vOut = Left$(vOut, Len(vOut) \ 2)
            Else
                'do a bit of magik :)
                memcpy ptr, ByVal VarPtr(vOut) + 8, 4& 'VT_BYREF
                Select Case vType
                Case vbByte
                    lr = ReadFile(hFile, ptr, 1&, lBytesRead, 0&)
                Case vbInteger
                    lr = ReadFile(hFile, ptr, 2&, lBytesRead, 0&)
                Case vbLong
                    lr = ReadFile(hFile, ptr, 4&, lBytesRead, 0&)
                Case vbCurrency
                    lr = ReadFile(hFile, ptr, 8&, lBytesRead, 0&)
                Case Else
                    UnknType = True
                    Debug.Print "Error! GetW for type #" & VarType(vOut) & " of buffer is not supported."
                    Err.Raise 52, , "Error! GetW for type #" & VarType(vOut) & " of buffer is not supported."
                End Select
            End If
            GetW = (0 <> lr)
            If 0 = lr And Not UnknType Then Debug.Print "Cannot read file!": Err.Raise 52, , "Cannot read file! Handle: " & hFile
        Else
            Debug.Print "Cannot set file pointer!": Err.Raise 52, , "Cannot set file pointer! Handle: " & hFile
        End If
    Else
        Debug.Print "Cannot set file pointer!": Err.Raise 52, , "Cannot set file pointer! Handle: " & hFile
    End If
End Function

Public Function PutW(hFile As Long, pos As Long, vInPtr As Long, cbToWrite As Long, Optional doAppend As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lBytesWrote  As Long
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If doAppend Then
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) Then Exit Function
    Else
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then Exit Function
    End If
    
    If NO_ERROR = Err.LastDllError Or (pos = 0 And Err.LastDllError = 183) Then
    
        If WriteFile(hFile, vInPtr, cbToWrite, lBytesWrote, 0&) Then
            PutW = True
        Else
            Debug.Print "PutW error: " & Err.LastDllError
        End If
    Else
        Debug.Print "SetFilePointer error: " & Err.LastDllError
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print Err, "modFile.PutW"
End Function

Public Sub CloseW(hFile As Long)    'закрывает файл, если его хендл не был расшарен
    If hFile <> 0 Then CloseHandle hFile
End Sub

Public Function FileExists(Path As String) As Boolean
    Dim l           As Long
    l = GetFileAttributes(StrPtr(Path))
    FileExists = Not CBool(l And vbDirectory) And (l <> INVALID_HANDLE_VALUE)
End Function

Public Function inIDE() As Boolean
    inIDE = (App.LogMode = 0)
End Function

