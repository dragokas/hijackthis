Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function ImageRemoveCertificate Lib "Imagehlp.dll" (ByVal FileHandle As Long, ByVal Index As Long) As Long
Private Declare Function ImageEnumerateCertificates Lib "Imagehlp.dll" (ByVal FileHandle As Long, ByVal TypeFilter As Integer, CertificateCount As Long, ByVal Indices As Long, ByVal IndexCount As Long) As Long
Private Declare Function MapFileAndCheckSum Lib "Imagehlp.dll" Alias "MapFileAndCheckSumW" (ByVal FileName As Long, HeaderSum As Long, CheckSum As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function OemToChar Lib "user32.dll" Alias "OemToCharA" (ByVal lpszScr As String, ByVal lpszDst As String) As Long
Private Declare Function CharToOem Lib "user32.dll" Alias "CharToOemA" (ByVal lpszScr As String, ByVal lpszDst As String) As Long

Const STD_OUTPUT_HANDLE         As Long = -11&
Const STD_ERROR_HANDLE          As Long = -12&

Const CHECKSUM_SUCCESS          As Long = 0&
Const CERT_SECTION_TYPE_ANY     As Integer = 255&

Const NO_ERROR                  As Long = 0&
Const INVALID_SET_FILE_POINTER  As Long = &HFFFFFFFF
Const FILE_BEGIN                As Long = 0&
Const FILE_CURRENT              As Long = 1&
Const FILE_END                  As Long = 2&
Const GENERIC_READ              As Long = &H80000000
Const GENERIC_WRITE             As Long = &H40000000
Const FILE_SHARE_READ           As Long = 1&
Const FILE_SHARE_WRITE          As Long = 2&
Const OPEN_EXISTING             As Long = 3&
Const INVALID_HANDLE_VALUE      As Long = -1&

Private cOut As Long
Private cErr As Long


Private Sub Main()
    On Error GoTo ErrorHandler
    
    ' &H3C -> PE_Header offset
    ' PE_Header offset + &H18 = Optional_PE_Header
    ' Optional_PE_Header + &H58 = PE EXE CheckSum
    
    ' PE_Header offset + &H78 (x32) or &H88 (x64) = Data_Directories offset
    ' Data_Directories offset + &H20 = SecurityDir
    ' SecurityDir (if no digital signature, must set VirtualAddress & Size to Zero)
    ' {
    '    DWORD   VirtualAddress; -> offset from the start of file to Digital signature Block
    '    DWORD   Size;
    ' } *PIMAGE_DATA_DIRECTORY;
    
    Dim i           As Long
    Dim hFile       As Long
    Dim lret        As Long
    Dim FileName    As String
    Dim CSum        As Long
    Dim HeaderSum   As Long
    Dim PE_offset   As Long
    Dim SignAddress As Long
    Dim cntCert     As Long
    Dim Indeces(127) As Long
    Dim ExitCode    As Long
    
    cOut = GetStdHandle(STD_OUTPUT_HANDLE)
    cErr = GetStdHandle(STD_ERROR_HANDLE)
    ExitCode = 1
  
    If Len(Command()) = 0 Then Using: ExitProcess 1
    
    FileName = UnQuote(Command())
    
    hFile = CreateFile(StrPtr(FileName), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    If hFile = INVALID_HANDLE_VALUE Then
        WriteC "Could not open file: " & FileName & ". Error: 0x" & Hex(Err.LastDllError), cErr
    Else
        WriteC "Begin digital signature removing: """ & FileName & """", cOut
        
        If Not IsSignPresent(hFile) Then
            WriteC "Error: no signatures!", cErr
            ExitCode = 0
        Else
            If Not CBool(ImageEnumerateCertificates(hFile, CERT_SECTION_TYPE_ANY, cntCert, VarPtr(Indeces(0)), UBound(Indeces) + 1)) Then
    
                WriteC "Error in ImageEnumerateCertificates: 0x" & Hex(Err.LastDllError), cErr
            Else
                ExitCode = 0
                
                For i = 0 To cntCert - 1
                    If Not CBool(ImageRemoveCertificate(hFile, Indeces(i))) Then
                        WriteC "Error in ImageRemoveCertificate: 0x" & Hex(Err.LastDllError), cErr
                        ExitCode = 1
                    End If
                Next
                
                lret = MapFileAndCheckSum(StrPtr(FileName), HeaderSum, CSum)
    
                If lret <> CHECKSUM_SUCCESS Then
                    WriteC "Error with calculation of PE CheckSum: Return value: " & lret & ". Error: 0x" & Hex(Err.LastDllError), cErr
                    ExitCode = 1
                Else
                    If CSum <> HeaderSum Then
                        WriteC "CheckSum is incorrect!", cErr
                        ExitCode = 1
                    End If
                End If
            End If
        End If
    End If
    
    CloseHandle hFile
    
    If 0 = ExitCode Then WriteC "Success." Else WriteC "Failed!"
    ExitProcess ExitCode
    Exit Sub
ErrorHandler:
    WriteC "Error #" & Err.Number & ". LastDll: 0x" & Hex(Err.LastDllError) & ". " & Err.Description, cErr
    ExitProcess 1
End Sub

Public Function IsSignPresent(hFile As Long) As Boolean
    ' 3Ch -> PE_Header offset
    ' PE_Header offset + 18h = Optional_PE_Header
    ' PE_Header offset + 78h (x86) or + 88h (x64) = Data_Directories offset
    ' Data_Directories offset + 20h = SecurityDir -> Address (dword), Size (dword) for digital signature.
    
    Const IMAGE_FILE_MACHINE_I386   As Integer = &H14C
    Const IMAGE_FILE_MACHINE_IA64   As Integer = &H200
    Const IMAGE_FILE_MACHINE_AMD64  As Integer = &H8664
    
    Dim PE_offset       As Long
    Dim SignAddress     As Long
    Dim DataDir_offset  As Long
    Dim DirSecur_offset As Long
    Dim Machine         As Integer
    Dim FSize           As Currency
    
    FSize = FileLenW(hFile)
    
    If FSize >= &H3C& + 6& Then
        GetW hFile, &H3C& + 1&, PE_offset
        GetW hFile, PE_offset + 4& + 1&, Machine
        
        Select Case Machine
            Case IMAGE_FILE_MACHINE_I386
                DataDir_offset = PE_offset + &H78&
            Case IMAGE_FILE_MACHINE_AMD64, IMAGE_FILE_MACHINE_IA64
                DataDir_offset = PE_offset + &H88&
            Case Else
                WriteC "Unknown architecture, not PE EXE or damaged image.", cErr
                Exit Function
        End Select
        
        DirSecur_offset = DataDir_offset + &H20&
        
        If FSize >= DirSecur_offset + 4& Then GetW hFile, DirSecur_offset + 1&, SignAddress
    End If
    
    IsSignPresent = (SignAddress <> 0)
End Function

                                                                  'do not change Variant type at all or you will die ^_^
Private Function GetW(hFile As Long, ByVal pos As Long, Optional vOut As Variant, Optional vOutPtr As Long, Optional cbToRead As Long) As Boolean
    Dim lBytesRead  As Long
    Dim lr          As Long
    Dim Ptr         As Long
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
                memcpy Ptr, ByVal VarPtr(vOut) + 8, 4& 'VT_BYREF
                Select Case vType
                Case vbByte
                    lr = ReadFile(hFile, Ptr, 1&, lBytesRead, 0&)
                Case vbInteger
                    lr = ReadFile(hFile, Ptr, 2&, lBytesRead, 0&)
                Case vbLong
                    lr = ReadFile(hFile, Ptr, 4&, lBytesRead, 0&)
                Case vbCurrency
                    lr = ReadFile(hFile, Ptr, 8&, lBytesRead, 0&)
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

Function FileLenW(hFile As Long) As Currency
    Dim lr          As Long
    Dim FileSize    As Currency

    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then FileLenW = FileSize * 10000&
        End If
    End If
End Function

Private Sub WriteC(ByVal txt As String, Optional cHandle As Long)
    Dim dwWritten As Long
    Debug.Print txt
    txt = txt & vbNewLine
    Call CharToOem(txt, txt)
    WriteFile IIf(cHandle = 0, cOut, cHandle), StrPtr(StrConv(txt, vbFromUnicode)), Len(txt), dwWritten, 0&
End Sub

Private Function DOS2Win(Str As String) As String
    If Len(Str) > 0 Then
        DOS2Win = String(Len(Str), 0&)
        OemToChar Str, DOS2Win
    End If
End Function

Private Function Win2Dos(Str As String) As String
    If Len(Str) > 0 Then
        Win2Dos = String(Len(Str), 0&)
        CharToOem Str, Win2Dos
    End If
End Function

Private Function UnQuote(sStr As String) As String
    If Left$(sStr, 1) = """" And Right$(sStr, 1) = """" And Len(sStr) > 1 Then
        UnQuote = Mid$(sStr, 2, Len(sStr) - 2)
    Else
        UnQuote = sStr
    End If
End Function

Sub Using()
    WriteC "Authenticode digital signature remover by Alex Dragokas"
    WriteC ""
    WriteC "Using:"
    WriteC ""
    WriteC "RemSign.exe [Path]"
End Sub
