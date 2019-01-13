VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsStream.cls]

' // clsStream.cls - binary stream class
' // © Krivous Anatoly Anatolevich (The trick), 2016

' Forked by Dragokas

Option Explicit

Private Const GRANULARITY   As Long = &H400  ' // Glanularity of memory allocation

Private mBuffer()       As Byte ' // Local buffer
Private mSizeOfBuffer   As Long ' // Size of buffer
Private mSizeOfData     As Long ' // Size of data
Private mCurIndex       As Long ' // Current pointer

Private Type LARGE_INTEGER
    lowpart                     As Long
    highpart                    As Long
End Type

Private Declare Function RtlGetCompressionWorkSpaceSize Lib "ntdll" ( _
                        ByVal CompressionFormatAndEngine As Integer, _
                        ByRef CompressBufferWorkSpaceSize As Long, _
                        ByRef CompressFragmentWorkSpaceSize As Long) As Long

Private Declare Function RtlCompressBuffer Lib "ntdll" ( _
                        ByVal CompressionFormatAndEngine As Integer, _
                        ByRef UncompressedBuffer As Any, _
                        ByVal UncompressedBufferSize As Long, _
                        ByRef CompressedBuffer As Any, _
                        ByVal CompressedBufferSize As Long, _
                        ByVal UncompressedChunkSize As Long, _
                        ByRef FinalCompressedSize As Long, _
                        ByRef WorkSpace As Any) As Long

Private Declare Function IsBadReadPtr Lib "kernel32" ( _
                        ByRef lp As Any, _
                        ByVal ucb As Long) As Long

Private Declare Function IsBadWritePtr Lib "kernel32" ( _
                        ByRef lp As Any, _
                        ByVal ucb As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
                   Alias "RtlMoveMemory" ( _
                   ByRef Destination As Any, _
                   ByRef Source As Any, _
                   ByVal Length As Long)
                   
Private Declare Sub ZeroMemory Lib "kernel32" _
                   Alias "RtlZeroMemory" ( _
                   ByRef dest As Any, _
                   ByVal numBytes As Long)

Private Declare Function CreateFile Lib "kernel32" _
                        Alias "CreateFileW" ( _
                        ByVal lpFileName As Long, _
                        ByVal dwDesiredAccess As Long, _
                        ByVal dwShareMode As Long, _
                        ByRef lpSecurityAttributes As Any, _
                        ByVal dwCreationDisposition As Long, _
                        ByVal dwFlagsAndAttributes As Long, _
                        ByVal hTemplateFile As Long) As Long
                        
Private Declare Function WriteFile Lib "kernel32" ( _
                        ByVal hFile As Long, _
                        ByRef lpBuffer As Any, _
                        ByVal nNumberOfBytesToWrite As Long, _
                        ByRef lpNumberOfBytesWritten As Long, _
                        ByRef lpOverlapped As Any) As Long
                        
Private Declare Function ReadFile Lib "kernel32" ( _
                        ByVal hFile As Long, _
                        ByRef lpBuffer As Any, _
                        ByVal nNumberOfBytesToRead As Long, _
                        ByRef lpNumberOfBytesRead As Long, _
                        ByRef lpOverlapped As Any) As Long
                        
Private Declare Function CloseHandle Lib "kernel32" ( _
                        ByVal hObject As Long) As Long

Private Declare Function GetFileSizeEx Lib "kernel32" ( _
                        ByVal hFile As Long, _
                        ByRef lpFileSize As LARGE_INTEGER) As Long

Private Const INVALID_HANDLE_VALUE            As Long = -1
Private Const GENERIC_WRITE                   As Long = &H40000000
Private Const GENERIC_READ                    As Long = &H80000000
Private Const CREATE_ALWAYS                   As Long = 2
Private Const OPEN_EXISTING                   As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL           As Long = &H80
Private Const COMPRESSION_FORMAT_LZNT1        As Long = 2

' // Size of data
Public Property Get Size() As Long
    Size = mSizeOfData
End Property

' // Current pointer
Public Property Get BufferPointer() As Long
    BufferPointer = mCurIndex
End Property
Public Property Let BufferPointer( _
                    ByVal Value As Long)
    If Value < 0 Then
        Err.Raise 5
        Exit Property
    End If

    If Value <= mSizeOfBuffer Then
        mCurIndex = Value
    Else
        mCurIndex = mSizeOfBuffer
    End If
    
End Property

' // Compress stream
Public Sub CompressStream()
    Dim szWorkSpace     As Long
    Dim WorkSpace()     As Byte
    Dim output()        As Byte
    Dim outputSize      As Long
    Dim returnedSize    As Long
        
    If mSizeOfData = 0 Then
        Err.Raise 7
    End If
    
    If RtlGetCompressionWorkSpaceSize(COMPRESSION_FORMAT_LZNT1, szWorkSpace, 0) Then
        Err.Raise 7
    End If
    
    outputSize = mSizeOfData * 2
    ReDim WorkSpace(szWorkSpace - 1)
    ReDim output(outputSize - 1)
    
    If RtlCompressBuffer(COMPRESSION_FORMAT_LZNT1, _
                         mBuffer(0), mSizeOfData, _
                         output(0), outputSize, 4096, _
                         returnedSize, WorkSpace(0)) Then
        Err.Raise 7
    End If
    
    ReDim mBuffer(returnedSize - 1)
    
    CopyMemory mBuffer(0), output(0), returnedSize
    
    mSizeOfData = returnedSize
    
End Sub

' // Write data to stream
Public Sub WriteData( _
           ByVal pData As Long, _
           ByVal dwSize As Long)
           
    ' // Check data
    If IsBadReadPtr(ByVal pData, dwSize) Then
        Err.Raise 7
        Exit Sub
    End If
    
    ' // Check size of buffer
    If mCurIndex + dwSize >= mSizeOfBuffer Then
        
        If Not ReAllocMemory(mCurIndex + dwSize + 1) Then
            Err.Raise 7
            Exit Sub
        End If
        
    End If
    
    ' // Copy data to buffer
    CopyMemory mBuffer(mCurIndex), ByVal pData, dwSize
    
    mCurIndex = mCurIndex + dwSize
    
    If mCurIndex > mSizeOfData Then
        mSizeOfData = mCurIndex
    End If
    
End Sub

' // Write stream
Public Sub WriteStream( _
           ByVal oStream As clsStream)
    Dim oldStreamPointer    As Long
    Dim locBuffer()         As Byte
    
    On Error GoTo error_handler
    
    Err.Clear
    
    ' // Save old pointer
    oldStreamPointer = oStream.BufferPointer
        
    If oStream.Size > 0 Then

        oStream.BufferPointer = 0
        
        ' // Allocate memory
        ReDim locBuffer(oStream.Size - 1)
        
        oStream.ReadData VarPtr(locBuffer(0)), oStream.Size
        
        ' // Write to stream
        Me.WriteData VarPtr(locBuffer(0)), oStream.Size
        
    End If
    
error_handler:
    
    oStream.BufferPointer = oldStreamPointer
    
    ' // If an error occured throw exception
    If Err.Number Then Err.Raise Err.Number
    
End Sub

' // Read data from stream
Public Sub ReadData( _
           ByVal pData As Long, _
           ByVal dwSize As Long)
    Dim dwActualDataCount   As Long
    
    ' // Check memory
    If IsBadWritePtr(ByVal pData, dwSize) Then
        Err.Raise 7
        Exit Sub
    End If
    
    ' // Check size of buffer
    If mCurIndex + dwSize >= mSizeOfBuffer Then
        dwActualDataCount = mSizeOfBuffer - mCurIndex
    Else
        dwActualDataCount = dwSize
    End If
    
    If dwActualDataCount > 0 Then
        ' // Copy data
        CopyMemory ByVal pData, mBuffer(mCurIndex), dwActualDataCount
        pData = pData + dwActualDataCount
        
    End If
    
    If dwSize > dwActualDataCount Then
        ' // Fill with zero
        ZeroMemory ByVal pData, dwSize - dwActualDataCount
        
    End If
    
    mCurIndex = mCurIndex + dwSize
    
    If mCurIndex > mSizeOfBuffer Then
        mCurIndex = mSizeOfBuffer
    End If
    
End Sub

' // Realloc memory
Private Function ReAllocMemory( _
                 ByVal dwSize As Long) As Boolean
    Dim dwNewBufferSize As Long
    
    On Error GoTo error_handler
    
    If dwSize < 0 Then
        Exit Function
    End If
    
    If dwSize Mod GRANULARITY Then
        dwNewBufferSize = (dwSize \ GRANULARITY) * GRANULARITY + GRANULARITY
    Else
        dwNewBufferSize = dwSize
    End If
    
    If dwNewBufferSize = 0 Then
        ' // Erase buffer
        Erase mBuffer()
    Else
        ' // Try to realloc memory
        ReDim Preserve mBuffer(dwNewBufferSize - 1)
    End If
    
    mSizeOfBuffer = dwNewBufferSize
    ReAllocMemory = True
    
error_handler:
    
End Function

' // Load file to byte stream. Returns size of written data
Public Function LoadFileToStream( _
                 ByRef sFilename As String, _
                 ByVal oStream As clsStream) As Long
    Dim hFile       As Long
    Dim sizeOfFile  As LARGE_INTEGER
    Dim binData()   As Byte
    Dim dwRead      As Long
    Dim ret         As Long
    
    On Error GoTo error_handler
    
    Err.Clear
    
    ret = -1
    
    hFile = CreateFile(StrPtr(sFilename), GENERIC_READ, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, g_FileBackupFlag, 0)
        
    If hFile = INVALID_HANDLE_VALUE Then
        ErrorMsg Err, "LoadFileToStream", "CreateFile", sFilename
        If inIDE Then Stop
        GoTo error_handler:
    End If

    ' // Check size
    GetFileSizeEx hFile, sizeOfFile
    If sizeOfFile.highpart <> 0 Or sizeOfFile.lowpart > 100000000 Or sizeOfFile.lowpart <= 0 Then
        ErrorMsg Err, "LoadFileToStream", "GetFileSizeEx => File is too big: ", sFilename
        If inIDE Then Stop
        GoTo error_handler:
    End If
    
    ReDim binData(sizeOfFile.lowpart - 1)
    
    If ReadFile(hFile, binData(0), sizeOfFile.lowpart, dwRead, ByVal 0&) = 0 Then
        ErrorMsg Err, "LoadFileToStream", "ReadFile", sFilename
        If inIDE Then Stop
        GoTo error_handler:
    End If

    ' // Save to stream
    oStream.WriteData VarPtr(binData(0)), dwRead
    
    ret = dwRead
    
error_handler:
    
    ' // Clean up
    If hFile > 0 Then CloseHandle hFile
    
    LoadFileToStream = ret
    
    If Err.Number Then Err.Raise Err.Number
    
End Function


' // Save stream data to file
Public Function SaveStreamToFile( _
                 ByRef sFilename As String, _
                 ByVal oStream As clsStream) As Long
    Dim hFile       As Long
    Dim binData()   As Byte
    Dim dwWritten   As Long
    Dim ret         As Long
    
    On Error GoTo error_handler
    
    Err.Clear
    
    ret = -1
    
    hFile = CreateFile(StrPtr(sFilename), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE Or g_FileBackupFlag, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        ErrorMsg Err, "SaveStreamToFile", "CreateFile", sFilename
        If inIDE Then Stop
        GoTo error_handler:
    End If
    
    If oStream.Size > 0 Then
        
        oStream.BufferPointer = 0
        
        ' // Read data to local buffer
        ReDim binData(oStream.Size - 1)
        
        oStream.ReadData VarPtr(binData(0)), oStream.Size
        
        If WriteFile(hFile, binData(0), oStream.Size, dwWritten, ByVal 0&) = 0 Then
            'Err.Raise 3040
            ErrorMsg Err, "SaveStreamToFile", "WriteFile", sFilename
            If inIDE Then Stop
            GoTo error_handler:
        End If
        
    End If
    
    ret = dwWritten
    
error_handler:
    
    ' // Clean up
    If hFile > 0 Then CloseHandle hFile

    SaveStreamToFile = ret
    
    If Err.Number Then Err.Raise Err.Number
    
End Function

Private Sub Class_Terminate()
    Erase mBuffer
End Sub
