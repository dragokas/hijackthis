Attribute VB_Name = "modBase"
Option Explicit

Private Type VERSION_NUMBER
    Major As Integer
    Minor As Integer
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress                  As Long
    Size                            As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic                           As Integer
    MajorLinkerVersion              As Byte
    MinorLinkerVersion              As Byte
    SizeOfCode                      As Long
    SizeOfInitializedData           As Long
    SizeOfUnitializedData           As Long
    AddressOfEntryPoint             As Long
    BaseOfCode                      As Long
    BaseOfData                      As Long
    ImageBase                       As Long
    SectionAlignment                As Long
    FileAlignment                   As Long
    MajorOperatingSystemVersion     As Integer
    MinorOperatingSystemVersion     As Integer
    MajorImageVersion               As Integer
    MinorImageVersion               As Integer
    MajorSubsystemVersion           As Integer
    MinorSubsystemVersion           As Integer
    W32VersionValue                 As Long
    SizeOfImage                     As Long
    SizeOfHeaders                   As Long
    CheckSum                        As Long
    SubSystem                       As Integer
    DllCharacteristics              As Integer
    SizeOfStackReserve              As Long
    SizeOfStackCommit               As Long
    SizeOfHeapReserve               As Long
    SizeOfHeapCommit                As Long
    LoaderFlags                     As Long
    NumberOfRvaAndSizes             As Long
    DataDirectory(15)               As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_FILE_HEADER
    Machine                         As Integer
    NumberOfSections                As Integer
    TimeDateStamp                   As Long
    PointerToSymbolTable            As Long
    NumberOfSymbols                 As Long
    SizeOfOptionalHeader            As Integer
    Characteristics                 As Integer
End Type

Private Type IMAGE_NT_HEADERS
    Signature                       As Long
    FileHeader                      As IMAGE_FILE_HEADER
    OptionalHeader                  As IMAGE_OPTIONAL_HEADER
End Type

Private Declare Function MapFileAndCheckSum Lib "Imagehlp.dll" Alias "MapFileAndCheckSumW" (ByVal Filename As Long, HeaderSum As Long, CheckSum As Long) As Long
Private Declare Function CharToOem Lib "user32.dll" Alias "CharToOemA" (ByVal lpszScr As String, ByVal lpszDst As String) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)

Const STD_OUTPUT_HANDLE As Long = -11&
Const CHECKSUM_SUCCESS As Long = 0&

Private ConsoleMode As Boolean


Public Sub Main()
    If App.LogMode <> 0 Then
        ConsoleMode = True
        If Len(Command()) = 0 Then
            WriteLog "Using: TSAwarePatch.exe [file]"
            Exit Sub
        End If
        TSPatch UnQuote(Trim(Command()))
    Else
        TSPatch "test.exe"
    End If
    ExitProcess 0
End Sub

Function TSPatch(sFile As String) As Boolean
    On Error GoTo ErrorHandler:
    
    Const IMAGE_FILE_MACHINE_I386   As Integer = &H14C
    Const IMAGE_FILE_MACHINE_AMD64  As Integer = &H8664
    
    Const IMAGE_DLLCHARACTERISTICS_NX_COMPAT As Integer = &H100 'DEP
    Const IMAGE_DLLCHARACTERISTICS_DYNAMIC_BASE As Integer = &H40 'ASLR
    Const IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE As Integer = &H8000 'TS Aware
    
    Dim NtHdr           As IMAGE_NT_HEADERS
    Dim CheckSum_offset As Long
    Dim DllCharacteristics_offset As Long
    Dim Subsystem_offset As Long
    Dim OSVersion_offset As Long
    Dim NT_Head_offset  As Long
    Dim Flags           As Integer
    Dim Machine         As Integer
    Dim FSize           As Currency
    Dim hFile           As Long
    Dim lHeaderSum      As Long
    Dim lCheckSum       As Long
    Dim lret            As Long
    
    OpenW sFile, FOR_READ_WRITE, hFile
    
    If hFile <= 0 Then
        WriteLog "Could not open file. Error: " & Err.LastDllError
        ExitProcess 2
        Exit Function
    End If
    
    FSize = FileLenW(, hFile)
    
    If FSize >= &H3C& + 6& Then
        GetW hFile, &H3C& + 1&, NT_Head_offset
        GetW hFile, NT_Head_offset + 1, , VarPtr(NtHdr), LenB(NtHdr)
        
        Select Case NtHdr.FileHeader.Machine
            Case IMAGE_FILE_MACHINE_I386

            Case IMAGE_FILE_MACHINE_AMD64
                'WriteLog "x64 images are not supported."
                'ExitProcess 1

            Case Else
                WriteLog "Unknown architecture, not PE EXE or damaged image."
                ExitProcess 1
        End Select
        
        CheckSum_offset = NT_Head_offset + (VarPtr(NtHdr.OptionalHeader.CheckSum) - VarPtr(NtHdr))
        Subsystem_offset = NT_Head_offset + (VarPtr(NtHdr.OptionalHeader.MajorSubsystemVersion) - VarPtr(NtHdr))
        OSVersion_offset = NT_Head_offset + (VarPtr(NtHdr.OptionalHeader.MajorOperatingSystemVersion) - VarPtr(NtHdr))
        DllCharacteristics_offset = NT_Head_offset + (VarPtr(NtHdr.OptionalHeader.DllCharacteristics) - VarPtr(NtHdr))
        
        Flags = NtHdr.OptionalHeader.DllCharacteristics
        
        WriteLog "Current flags: 0x" & Hex(Flags)
        
        'Fix flags
        If Flags And (IMAGE_DLLCHARACTERISTICS_NX_COMPAT Or IMAGE_DLLCHARACTERISTICS_DYNAMIC_BASE Or IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE) Then
            WriteLog "No patch required."
        Else
            Flags = Flags Or IMAGE_DLLCHARACTERISTICS_NX_COMPAT
            Flags = Flags Or IMAGE_DLLCHARACTERISTICS_DYNAMIC_BASE
            Flags = Flags Or IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE
            
            PutW hFile, DllCharacteristics_offset + 1, VarPtr(Flags), 2
            
            Flags = 0
            GetW hFile, DllCharacteristics_offset + 1, Flags
            
            WriteLog "New flags: 0x" & Hex(Flags)
        End If
        
        Dim OSVersion As VERSION_NUMBER
        Dim SubSystem As VERSION_NUMBER
        OSVersion.Major = NtHdr.OptionalHeader.MajorOperatingSystemVersion
        OSVersion.Minor = NtHdr.OptionalHeader.MinorOperatingSystemVersion
        SubSystem.Major = NtHdr.OptionalHeader.MajorSubsystemVersion
        SubSystem.Minor = NtHdr.OptionalHeader.MinorSubsystemVersion
        
        'Fix OS version/Subsystem required to run the image if new version of linker is used
        WriteLog "PE OS Version: " & OSVersion.Major & "." & OSVersion.Minor
        WriteLog "PE Subsystem:  " & SubSystem.Major & "." & SubSystem.Minor
        
        If Not (OSVersion.Major = 4 And OSVersion.Minor = 0) Then
            OSVersion.Major = 4
            OSVersion.Minor = 0
            PutW hFile, OSVersion_offset + 1, VarPtr(OSVersion), 4
            GetW hFile, OSVersion_offset + 1, , VarPtr(OSVersion), 4&
            WriteLog "New PE OS Version: " & OSVersion.Major & "." & OSVersion.Minor
        End If
        If Not (SubSystem.Major = 4 And SubSystem.Minor = 0) Then
            SubSystem.Major = 4
            SubSystem.Minor = 0
            PutW hFile, Subsystem_offset + 1, VarPtr(SubSystem), 4
            GetW hFile, Subsystem_offset + 1, , VarPtr(SubSystem), 4&
            WriteLog "New PE Subsystem: " & SubSystem.Major & "." & SubSystem.Minor
        End If
        
        'correction of checksum
        lret = MapFileAndCheckSum(StrPtr(sFile), lHeaderSum, lCheckSum)
        WriteLog "CheckSum is: " & "0x" & Hex(lHeaderSum)
        
        If CHECKSUM_SUCCESS = lret Then
            If lHeaderSum <> lCheckSum Then
                PutW hFile, CheckSum_offset + 1, VarPtr(lCheckSum), 4
                WriteLog "New CheckSum is: " & "0x" & Hex(lCheckSum)
                
                Call MapFileAndCheckSum(StrPtr(sFile), lHeaderSum, lCheckSum)
                
                If lHeaderSum <> lCheckSum Then
                    WriteLog "Cannot correct PE checksum!"
                Else
                    WriteLog "CheckSum is correct."
                End If
            Else
                WriteLog "CheckSum is correct."
            End If
        Else
            WriteLog "Cannot calculate PE checksum. Error: " & lret
            ExitProcess 1
        End If
    End If
    
    CloseW hFile
    
    Exit Function
ErrorHandler:
    Debug.Print Err, "TSPatch"
    If 0 <> hFile Then CloseW hFile
    If inIDE Then Stop: Resume Next
End Function

Private Sub WriteLog(ByVal sText As String)
    If ConsoleMode Then
        Static cHandle As Long
        Dim dwWritten  As Long
        If cHandle = 0 Then cHandle = GetStdHandle(STD_OUTPUT_HANDLE)
        CharToOem sText, sText
        sText = sText & vbNewLine
        WriteFile cHandle, StrPtr(StrConv(sText, vbFromUnicode)), Len(sText), dwWritten, 0&
    Else
        Debug.Print sText
    End If
End Sub

Private Function UnQuote(sStr As String) As String
    If Left$(sStr, 1) = """" And Right$(sStr, 1) = """" And Len(sStr) > 1 Then
        UnQuote = Mid$(sStr, 2, Len(sStr) - 2)
    Else
        UnQuote = sStr
    End If
End Function

