Attribute VB_Name = "modMain"
Option Explicit

Private Type VerInfo
    Major As Byte
    Minor As Byte
    Build As Byte
    Revision As Byte
    aVerInfoData() As Byte
    ResName As Long
    ResLanguage As Integer
End Type

' PE EXE Version Patcher by Alex Dragokas

Private Declare Function EnumResourceTypes Lib "kernel32.dll" Alias "EnumResourceTypesW" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesW" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesW" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExW" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Integer) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hModule As Long, ByVal hResInfo As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function LoadResource Lib "kernel32.dll" (ByVal hModule As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32.dll" (ByVal hResData As Long) As Long
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceW" (ByVal pFileName As Long, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceW" (ByVal hUpdate As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceW" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim ResTypes    As New Collection
Dim ResNames    As Collection
Dim ResLang     As Collection
Dim ResSize     As Collection
Dim ResData     As Collection

Public Ver      As VerInfo


Sub Using(Optional AddToMsg As String)
    MsgBox IIf(Len(AddToMsg) <> 0, AddToMsg & vbCrLf & vbCrLf, "") & _
        "Using:" & vbCrLf & App.EXEName & ".exe [Path to exe] [Version]" & vbCrLf & vbCrLf & _
            "Example: " & App.EXEName & " c:\temp\my.exe 2.1.1.5", vbExclamation, "Version Patcher"
End Sub


Public Sub Main()

    Dim hModule As Long
    Dim i       As Long
    Dim j       As Long
    Dim k       As Long
    Dim ff      As Integer
    Dim LogPath As String
    Dim argv()  As String
    Dim aVer
    
    Dim PathToResourceFile As String
    
    '// TODO: make port of https://www.codeproject.com/Articles/6317/Updating-version-information-at-run-time
    
    
    
    
    
    
    
    
    
    If Not ParseCommandLine(Command(), argv) Then
        Using
        End
    End If

    If UBound(argv) < 1 Then
        Using "Not enough arguments."
        End
    End If

    aVer = Split(argv(1), ".")

    If UBound(aVer) < 3 Then
        Using "Version is specified in unknown format."
        End
    End If

    PathToResourceFile = argv(0)

    If 0 = Len(Dir$(PathToResourceFile)) Then
        Using "Specified file is not exist."
        End
    End If

    Ver.Major = aVer(0)
    Ver.Minor = aVer(1)
    Ver.Build = aVer(2)
    Ver.Revision = aVer(3)

'    LogPath = App.Path & "\EnumResReport.txt"
'
'    ff = FreeFile()
'    Open LogPath For Output As #ff

    'GetManifestLangCodeFromFile = -1
    
    hModule = LoadLibrary(PathToResourceFile)
    
    If hModule Then
        If EnumResourceTypes(hModule, AddressOf EnumResTypeProcW, 0&) Then
        
            For i = 1 To ResTypes.Count
                Debug.Print "Type: " & ResTypes(i)
                Print #ff, "Type: " & ResTypes(i)
                
                Set ResNames = New Collection
                If EnumResourceNames(hModule, GetVarRef(ResTypes(i)), AddressOf EnumResNameProcW, 0&) Then
                
                    For j = 1 To ResNames.Count
                        Debug.Print "  Name: " & ResNames(j)
                        Print #ff, "  Name: " & ResNames(j)
                        
                        Set ResLang = New Collection
                        Set ResSize = New Collection
                        Set ResData = New Collection
                        
                        If EnumResourceLanguages(hModule, GetVarRef(ResTypes(i)), GetVarRef(ResNames(j)), AddressOf EnumResLangProcW, 0&) Then
                        
                            For k = 1 To ResLang.Count
                        
                                Debug.Print "    Lang: " & ResLang(k) & "  -  " & ResSize(k) & " bytes"
                                Print #ff, "    Lang: " & ResLang(k) & "  -  " & ResSize(k) & " bytes"
                                Print #ff, ResData(k)
                                
                                'If ResTypes(i) = 24 And ResNames(j) = 1 Then GetManifestLangCodeFromFile = ResLang(k): Exit Sub
                        
                            Next
                        Else
                            MsgBox "Error #" & Err.LastDllError & " with enumeration of Resource Languages."
                        End If
                    Next
                Else
                    MsgBox "Error #" & Err.LastDllError & " with enumeration of Resource Names."
                End If
            Next
        Else
            MsgBox "Error #" & Err.LastDllError & " with enumeration of Resource Types."
        End If
        FreeLibrary hModule
    Else
        MsgBox "Cannot load file: " & PathToResourceFile
    End If
    
    'Close #ff
    
    UpdateVersionInfo Ver.aVerInfoData, PathToResourceFile, Ver.ResName, Ver.ResLanguage
    
    'If 0 <> Len(Dir$(LogPath)) Then Shell "rundll32.exe shell32,ShellExec_RunDLL " & """" & LogPath & """"
    
End Sub

Function EnumResLangProcW(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByVal lpszLanguage As Integer, ByVal lParam As Long) As Long
    
    Dim hResInfo    As Long
    Dim lSize       As Long
    Dim hLoadedRes  As Long
    Dim lpResLock   As Long
    Dim aResData()  As Byte
    Dim sResData    As String
    
    hResInfo = FindResourceEx(hModule, GetVarRef(lpszType), GetVarRef(lpszName), lpszLanguage)
    
    If hResInfo Then lSize = SizeofResource(hModule, hResInfo)
    
    ResLang.Add lpszLanguage
    ResSize.Add lSize

    ' Get Data from resource

    If 0 <> lSize Then
        hLoadedRes = LoadResource(hModule, hResInfo)
    
        If 0 <> hLoadedRes Then
            lpResLock = LockResource(hLoadedRes)
        
            If 0 <> lpResLock Then
                ReDim aResData(0 To lSize - 1)
                
                memcpy ByVal VarPtr(aResData(0)), ByVal lpResLock, lSize
                
                sResData = StrConv(StrConv(aResData, vbUnicode), vbFromUnicode)
                
                If lpszType = 16 Then
                
                    Ver.aVerInfoData = aResData
                    Ver.ResLanguage = lpszLanguage
                    Ver.ResName = lpszName

                End If
                
            End If
        End If
    End If
    
    ResData.Add sResData
    
    EnumResLangProcW = True
End Function

Function UpdateVersionInfo(aResData() As Byte, FileName As String, lpszName As Long, lpszLanguage As Integer)

    Dim hUpdate     As Long
    Dim ret         As Long
    Dim sResData    As String
    Dim sVer        As String
    Dim pos         As Long

    With Ver
        aResData(50) = .Major
        aResData(48) = .Minor
        aResData(54) = .Build
        aResData(52) = .Revision
                    
        aResData(58) = .Major
        aResData(56) = .Minor
        aResData(62) = .Build
        aResData(60) = .Revision
        
        sVer = .Major & "." & .Minor & "." & .Build & "." & .Revision
    End With
    
    sResData = StrConv(StrConv(aResData, vbUnicode), vbFromUnicode)
    
    pos = InStr(sResData, "ProductVersion")
    If pos <> 0 Then
        Mid(sResData, pos) = "ProductVersion" & vbNullChar & sVer & String$(9 - Len(sVer), vbNullChar)
    End If
    pos = InStr(sResData, "FileVersion")
    If pos <> 0 Then
        Mid(sResData, pos) = "FileVersion" & vbNullChar & vbNullChar & sVer & String$(9 - Len(sVer), vbNullChar)
    End If
    
    aResData() = sResData

    hUpdate = BeginUpdateResource(StrPtr(FileName), False)
    
    If 0 <> hUpdate Then
        ret = UpdateResource(hUpdate, 16, lpszName, lpszLanguage, aResData(0), UBound(aResData) + 1)
        
        If ret <> 0 And 0 = Err.LastDllError Then
            EndUpdateResource hUpdate, False
        
            If 0 <> Err.LastDllError Then
                Debug.Print "Ошибка EndUpdateResource= " & Err.LastDllError
            End If
        Else
            Debug.Print "Ошибка UpdateResource= " & Err.LastDllError
        End If
    Else
        Debug.Print "Ошибка BeginUpdateResource= " & Err.LastDllError
    End If
    
End Function

Function EnumResNameProcW(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByVal lParam As Long) As Long
    Dim sName   As String
    Dim lName   As Long
    Dim lLen    As Long

    If (lpszName And &HFFFF0000) = 0& Then  'ID
        lName = lpszName And &HFFFF&
        ResNames.Add lName
    Else                                    'String
        lLen = lstrlen(lpszName)
        If lLen > 0 Then
            sName = Space$(lLen)
            lstrcpyn StrPtr(sName), lpszName, lLen + 1
            ResNames.Add sName
        End If
   End If
   
   EnumResNameProcW = True
End Function

Function EnumResTypeProcW(ByVal hModule As Long, ByVal lpszType As Long, ByVal lParam As Long) As Boolean

    Dim sType As String, lLen As Long

    If IS_INTRESOURCE(lpszType) Then    'ID
        ResTypes.Add lpszType
    Else                                'String
        lLen = lstrlen(lpszType)
        If lLen > 0 Then
            sType = Space$(lLen)
            lstrcpyn StrPtr(sType), lpszType, lLen + 1
            ResTypes.Add sType
        End If
    End If

    EnumResTypeProcW = True
End Function

Public Function IS_INTRESOURCE(i As Long) As Boolean
    IS_INTRESOURCE = (i <= &HFFFF&)
End Function

Function GetVarRef(vVar As Variant) As Long
    If VarType(vVar) = vbLong Then GetVarRef = vVar Else GetVarRef = StrPtr(vVar)
End Function

Function ParseCommandLine(Line As String, argv() As String) As Boolean
  Dim Lex, nL&, nA&, Unit$, argc&, St$
  St = Line
  If Len(St) > 0 Then ParseCommandLine = True Else Exit Function
  Lex = Split(St) 'Разбиваем по пробелам для лексического анализа
  ReDim argv(0 To UBound(Lex)) As String 'Определяем выходной массив до максимально возможного числа параметров
  'argv(0) = App.Path & "\" & App.EXEName & ".exe"
  If Len(St) <> 0 Then
    Do While nL <= UBound(Lex)
      Unit = Lex(nL) 'Записысаем смысловую единицу
      If Len(Unit) <> 0 Then 'Защита от двойных пробелов между параметрами
        If Left$(Lex(nL), 1) = """" Then 'Если попалась кавычка
          Do Until Right$(Lex(nL), 1) = """" 'Пока не найдем кавычке пару
            nL = nL + 1
            If nL > UBound(Lex) Then Exit Do 'Если не дождались завершающей кавычки, а лимит превышен
            Unit = Unit & " " & Lex(nL)
          Loop
          Unit = Replace$(Unit, """", "") 'Удаляем кавычки
        End If
        argv(nA) = Unit
        nA = nA + 1 'Счетчик выходных аргументов
      End If
      nL = nL + 1 'Счетчик лексических единиц
    Loop
  End If
  If nA = 0 Then
    Erase argv
  Else
    ReDim Preserve argv(0 To nA - 1)
  End If
  argc = nA
End Function


