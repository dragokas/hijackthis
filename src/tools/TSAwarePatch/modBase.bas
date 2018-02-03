Attribute VB_Name = "modBase"
Option Explicit

Private Declare Function CharToOem Lib "user32.dll" Alias "CharToOemA" (ByVal lpszScr As String, ByVal lpszDst As String) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)

Const STD_OUTPUT_HANDLE As Long = -11&

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
    
    Dim NT_Head_offset  As Long
    Dim DllCharacteristics_offset As Long
    Dim Flags           As Integer
    Dim Machine         As Integer
    Dim FSize           As Currency
    Dim hFile           As Long
    
    OpenW sFile, FOR_READ_WRITE, hFile
    
    If hFile <= 0 Then
        WriteLog "Could not open file. Error: " & Err.LastDllError
        ExitProcess 2
        Exit Function
    End If
    
    FSize = FileLenW(, hFile)
    
    If FSize >= &H3C& + 6& Then
        GetW hFile, &H3C& + 1&, NT_Head_offset
        GetW hFile, NT_Head_offset + 4& + 1&, Machine
        
        Select Case Machine
            Case IMAGE_FILE_MACHINE_I386
                DllCharacteristics_offset = NT_Head_offset + &H5E&
            Case IMAGE_FILE_MACHINE_AMD64
                DllCharacteristics_offset = NT_Head_offset + &H5E&
            Case Else
                WriteLog "Unknown architecture, not PE EXE or damaged image."
                ExitProcess 1
        End Select
        
        GetW hFile, DllCharacteristics_offset + 1, Flags
        
        WriteLog "Current flags: 0x" & Hex(Flags)
        
        Flags = Flags Or IMAGE_DLLCHARACTERISTICS_NX_COMPAT
        Flags = Flags Or IMAGE_DLLCHARACTERISTICS_DYNAMIC_BASE
        Flags = Flags Or IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE
        
        PutW hFile, DllCharacteristics_offset + 1, VarPtr(Flags), 2
        
        Flags = 0
        GetW hFile, DllCharacteristics_offset + 1, Flags
        
        WriteLog "New flags: 0x" & Hex(Flags)
        
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

