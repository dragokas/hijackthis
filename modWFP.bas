Attribute VB_Name = "modWFP"
'
' WPF Enumerator by Alex Dragokas
'

' Regards to:
' - SSTREGG for variant of C-code implementation on Vista+
' - The Trick for patch to call VB6 functions by pointer

Option Explicit

Const MAX_PATH As Long = 260&

Type PPROTECTED_FILE_INFO
    length As Long
    FileName As String * MAX_PATH
End Type

Type PROTECTED_FILE_DATA
    FileName As String * MAX_PATH
    FileNumber As Long
End Type

Type PPROTECT_FILE_ENTRY
    SourceFileName As Long  'pointer PWSTR
    FileName As Long        'pointer PWSTR
    InfName As Long         'pointer PWSTR
End Type

Private Declare Function GetMem4 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long
Private Declare Function VirtualProtect Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Sub EbGetExecutingProj Lib "vba6.dll" (hProject As Long)
Private Declare Function TipGetFunctionId Lib "vba6.dll" (ByVal hProj As Long, ByVal bstrName As Long, ByRef bstrId As Long) As Long
Private Declare Function TipGetLpfnOfFunctionId Lib "vba6.dll" (ByVal hProject As Long, ByVal bstrId As Long, ByRef lpAddress As Long) As Long
Private Declare Sub SysFreeString Lib "oleaut32.dll" (ByVal lpbstr As Long)
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcAddressByOrd Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long
'Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

'Private Declare Function SfcGetNextProtectedFile Lib "sfc_os.dll" (ByVal RpcHandle As Long, ProtFileData As PROTECTED_FILE_DATA) As Long

Private Const PAGE_EXECUTE_READWRITE = &H40&
Private Const ERROR_INSUFFICIENT_BUFFER = &H7A&
Private Const ERROR_NO_MORE_FILES = &H12&

Dim SystemRoot      As String


' Прототипы функций, вызов которых перенаправляется по адресу Addr
Private Function BeginFileMapEnumeration(ByVal Addr As Long, ByVal Reserved0 As Long, ByVal Reserved1 As Long, Handle As Long) As Long: End Function
Private Function CloseFileMapEnumeration(ByVal Addr As Long, ByVal Handle As Long) As Long: End Function
Private Function GetNextFileMapContent(ByVal Addr As Long, ByVal Reserved As Long, ByVal SfcHandle As Long, ByVal size As Long, ProtectedInfo As PPROTECTED_FILE_INFO, dwNeeded As Long) As Long: End Function
Private Function SfcGetNextProtectedFile(ByVal Addr As Long, ByVal RpcHandle As Long, ProtFileData As PROTECTED_FILE_DATA) As Long: End Function
Private Function SfcGetFiles(ByVal Addr As Long, ProtFileData As PPROTECT_FILE_ENTRY, FileCount As Long) As Long: End Function


Private Sub InitVars()
    Dim ret             As Long
    SystemRoot = Space(MAX_PATH)
    ret = GetWindowsDirectory(StrPtr(SystemRoot), Len(SystemRoot))
    SystemRoot = Left$(SystemRoot, ret) & "\"
End Sub

Public Function SFCList_XP() As String()
    On Error GoTo ErrorHandler:

    Dim ret                     As Long
    Dim hSfc_Lib                As Long
    Dim SfcGetNextProtFileAddr  As Long
    Dim pfd                     As PROTECTED_FILE_DATA
    Dim SFCList()               As String
    Dim i                       As Long
    
    hSfc_Lib = LoadLibrary(StrPtr("sfc.dll"))
    If hSfc_Lib = 0 Then Debug.Print "Error! Cannot get sfc.dll library handle.": Exit Function
    
    PatchFunc "SfcGetNextProtectedFile", AddressOf SfcGetNextProtectedFile
    
    SfcGetNextProtFileAddr = GetProcAddress(hSfc_Lib, "SfcGetNextProtectedFile")
    If SfcGetNextProtFileAddr = 0 Then Debug.Print "Error: cannot get SfcGetNextProtectedFile function address!": FreeLibrary hSfc_Lib: Exit Function
    
    ReDim SFCList(300)
    
    Do
        ' by Patch
        ret = SfcGetNextProtectedFile(SfcGetNextProtFileAddr, 0&, pfd)
        
        If ret Then
            If UBound(SFCList) < i Then ReDim Preserve SFCList(i + 100)
            SFCList(i) = TrimChar0(pfd.FileName)
        End If
        
        ' by Declare
        'ret = SfcGetNextProtectedFile(0&, pfd)
        'If ret Then Print #ff, StrConv(pfd.FileName, vbFromUnicode)
    
        i = i + 1
    Loop While ret

    If i = 0 Then
        ReDim SFCList(0)
    Else
        ReDim Preserve SFCList(i - 1)
    End If
    SFCList_XP = SFCList
    
    FreeLibrary hSfc_Lib
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modWFP.SFCList_XP"
    If inIDE Then Stop: Resume Next
End Function

Public Function SFCList_XP_0() As String()  ' with SFCFILES.dll
    On Error GoTo ErrorHandler:

    Dim ret                     As Long
    Dim hSfcFil_Lib             As Long
    Dim SfcGetFilesAddr         As Long
    Dim FileCount               As Long
    Dim index                   As Long
    Dim strAdr                  As Long
    Dim strLen                  As Long
    Dim FileName                As String
    Dim SFCList()               As String
    Dim pfe                     As PPROTECT_FILE_ENTRY
    
    InitVars
    
    hSfcFil_Lib = LoadLibrary(StrPtr("sfcfiles.dll"))
    If hSfcFil_Lib = 0 Then Debug.Print "Error! Cannot get sfcfiles.dll library handle.": Exit Function
    
    PatchFunc "SfcGetFiles", AddressOf SfcGetFiles
    
    SfcGetFilesAddr = GetProcAddress(hSfcFil_Lib, "SfcGetFiles")
    If SfcGetFilesAddr = 0 Then Debug.Print "Error: cannot get SfcGetFiles function address!": FreeLibrary hSfcFil_Lib: Exit Function
    
    ret = SfcGetFiles(SfcGetFilesAddr, pfe, FileCount)

    'Debug.Print "FileName=" & pfe.FileName
    'Debug.Print "InfName=" & pfe.InfName
    'Debug.Print "SourceFileName=" & pfe.SourceFileName
        
    If pfe.SourceFileName = 0 Then Debug.Print "Error! Can't get a pointer to FileNames with SfcGetFiles function !": FreeLibrary hSfcFil_Lib: Exit Function
            
    ReDim SFCList(FileCount - 1)
    
    For index = 0 To FileCount - 1
        GetMem4 ByVal pfe.SourceFileName + 4 + index * 12, strAdr
        strLen = lstrlen(strAdr)
        FileName = Space(strLen)
        lstrcpyn StrPtr(FileName), strAdr, strLen + 1
        SFCList(index) = EnvironW(FileName) 'Replace$(FileName, "%systemroot%\", SystemRoot, , , 1)
    Next
    GlobalFree pfe.SourceFileName

    SFCList_XP_0 = SFCList
    
    FreeLibrary hSfcFil_Lib
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modWFP.SFCList_XP_0"
    If inIDE Then Stop: Resume Next
End Function


Function TrimChar0(sText As String) As String
    Dim pos   As Long
    pos = InStr(sText, vbNullChar)
    If pos Then TrimChar0 = Left$(sText, pos - 1) Else TrimChar0 = sText
End Function

Public Function SFCList_Vista() As String()
    On Error GoTo ErrorHandler:

    Dim dwNeeded         As Long
    Dim dwBufferSize     As Long
    Dim pData            As PPROTECTED_FILE_INFO
    Dim hSfc_os_Lib      As Long
    Dim hSFC             As Long
    Dim ret              As Long
    Dim BeginFileMapAddr As Long
    Dim GetNextFileAddr  As Long
    Dim CloseFileMapAddr As Long
    Dim SFCList()        As String
    Dim i                As Long
    
    InitVars
    
    hSfc_os_Lib = LoadLibrary(StrPtr("sfc_os.dll"))
    If hSfc_os_Lib = 0 Then Debug.Print "Error! Cannot get sfc_os.dll library handle.": Exit Function
    
    PatchFunc "BeginFileMapEnumeration", AddressOf BeginFileMapEnumeration
    PatchFunc "CloseFileMapEnumeration", AddressOf CloseFileMapEnumeration
    PatchFunc "GetNextFileMapContent", AddressOf GetNextFileMapContent

    BeginFileMapAddr = GetProcAddress(hSfc_os_Lib, "BeginFileMapEnumeration")
    If BeginFileMapAddr = 0 Then Debug.Print "Error: cannot get BeginFileMapEnumeration function address!": FreeLibrary hSfc_os_Lib: Exit Function
    
    ret = BeginFileMapEnumeration(BeginFileMapAddr, 0&, 0&, hSFC)
    If hSFC = 0 Then Debug.Print "Error! Cannot get handle of first element of BeginFileMapEnumeration.": FreeLibrary hSfc_os_Lib: Exit Function
    
    dwBufferSize = Len(pData)
    
    GetNextFileAddr = GetProcAddress(hSfc_os_Lib, "GetNextFileMapContent")
    
    ReDim SFCList(300)
    
    Do
        ret = GetNextFileMapContent(GetNextFileAddr, 0&, hSFC, dwBufferSize, pData, dwNeeded)
    
        Select Case Err.LastDllError ' <--- Does not working here !!!
        
            Case 0
                If UBound(SFCList) < i Then ReDim Preserve SFCList(i + 100)
                SFCList(i) = Replace$(Left$(pData.FileName, pData.length \ 2), "\SystemRoot\", SystemRoot, 1, 1, 1)
                i = i + 1
        
            Case ERROR_NO_MORE_FILES Or (pData.length = 0)
                Exit Do
        
            Case ERROR_INSUFFICIENT_BUFFER Or (dwNeeded > dwBufferSize)
                Debug.Print "ERROR_INSUFFICIENT_BUFFER"
    
        End Select

        If pData.length = 0 Then Exit Do

    Loop
    
    CloseFileMapAddr = GetProcAddress(hSfc_os_Lib, "CloseFileMapEnumeration")
    If CloseFileMapAddr = 0 Then
        Debug.Print "Error: cannot get CloseFileMapEnumeration function address!"
    Else
        CloseFileMapEnumeration CloseFileMapAddr, hSFC
    End If
    
    If i = 0 Then
        ReDim SFCList(0)
    Else
        ReDim Preserve SFCList(i - 1)
    End If
    SFCList_Vista = SFCList
    
    FreeLibrary hSfc_os_Lib
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modWFP.SFCList_Vista"
    If inIDE Then Stop: Resume Next
End Function
 
' Пропатчивание функции
Private Sub PatchFunc(FuncName As String, ByVal Addr As Long)
    Dim lpAddr As Long, hProj As Long, SID As Long, inIDE As Boolean
 
    Debug.Assert MakeTrue(inIDE)
 
    ' Получаем адрес функции
    If inIDE Then
        EbGetExecutingProj hProj
        TipGetFunctionId hProj, StrPtr(FuncName), SID
        TipGetLpfnOfFunctionId hProj, SID, lpAddr
        SysFreeString SID
    Else
        lpAddr = GetAddr(Addr)
        VirtualProtect lpAddr, 8, PAGE_EXECUTE_READWRITE, 0
    End If
 
    ' Записываем вставку
    ' Запускать только по Ctrl+F5!!
    ' pop eax
    ' pop ecx
    ' push eax
    ' jmp ecx
 
    GetMem4 &HFF505958, ByVal lpAddr
    GetMem4 &HE1, ByVal lpAddr + 4
End Sub
Private Function GetAddr(ByVal Addr As Long) As Long
    GetAddr = Addr
End Function
Public Function MakeTrue(ByRef bvar As Boolean) As Boolean
    bvar = True: MakeTrue = True
End Function
