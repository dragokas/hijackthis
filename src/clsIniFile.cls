VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsIniFile.cls]

Option Explicit

'IniFile class by Alex Dragokas

'Mainly intended for caching data beetween read-write operations
'Support UTF-16 LE ini-files format
'Provide wide range of methods
'Doesn't support reading / saving commentary in ini file

#Const UseHashtable = True 'true, if you are using HashTable class by The Trick, instead of Scripting.Dictionary

Private mb_Loaded       As Boolean
Private mb_UnsavedData  As Boolean
Private mb_IsUnicode    As Boolean
Private ms_IniFile      As String
Private ml_Codepage     As Long
Private mo_Sect         As Object
Private mb_CompareMode  As VbCompareMethod

'// To begin working with the class
'
'@sFile         - Input file name
'@id_Codepage   - Input codepage
'
'Note: supported codepages:
' - 1251 (ANSI)
' - 1200 (UTF-16LE)
' - 0 (auto-detect)
'
Public Sub InitFile(sFile As String, id_Codepage As Long)
    On Error GoTo ErrorHandler:
    Dim sIniFile As String
    sIniFile = sFile
    If Not FileExists(sFile) Then
        'relative path -> search on %windir%, %windir%\System32
        If InStr(sFile, "\") = 0 Then
            If FileExists(sWinSysDir & "\" & sFile) Then
                sIniFile = sWinSysDir & "\" & sFile
            ElseIf FileExists(sWinDir & "\" & sFile) Then
                sIniFile = sWinDir & "\" & sFile
            End If
        End If
    End If
    'if another file is already loaded
    If (ms_IniFile <> "") Then
        'if the same file
        If StrComp(ms_IniFile, sIniFile, 1) = 0 And ml_Codepage = id_Codepage Then
            'do nothing
        Else
        'if another file or codepage
            'flush data to the last opened file
            Flush
            'reinit dictionary
            ReleaseDictionary
            InitDictionary
            
            ms_IniFile = sIniFile
            InitCodepage id_Codepage
            LoadFile
        End If
    Else 'first init
        ms_IniFile = sIniFile
        InitCodepage id_Codepage
        LoadFile
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_InitFile", sFile, id_Codepage
    If inIDE Then Stop: Resume Next
End Sub

Private Sub InitCodepage(id_Codepage As Long)
    On Error GoTo ErrorHandler:
    If id_Codepage <> 1251 And id_Codepage <> 1200 And id_Codepage <> 0 Then
        MsgBoxW "Incorrect codepage!"
        Exit Sub
    End If
    Dim CP As Long
    ml_Codepage = id_Codepage
    If ml_Codepage = 0 Then ' auto-detect
        CP = FileGetTypeBOM(ms_IniFile)
        If CP = 1200 Then ml_Codepage = 1200
    End If
    mb_IsUnicode = IIf(ml_Codepage = 1200, True, False)
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_InitCodepage", id_Codepage
    If inIDE Then Stop: Resume Next
End Sub

'// Method of comparing the strings
Public Property Get CompareMethod() As VbCompareMethod
    CompareMethod = mb_CompareMode
End Property

Public Property Let CompareMethod(ByVal CompareMode As VbCompareMethod)
    mb_CompareMode = CompareMode
End Property

Private Sub LoadFile()
    On Error GoTo ErrorHandler:

    Dim aContents() As String
    Dim i As Long, j As Long, pos As Long
    Dim sSection As String
    Dim sParam As String
    Dim sData As String
    
    If Not FileExists(ms_IniFile) Then Exit Sub
    
    aContents = ReadFileToArray(ms_IniFile, mb_IsUnicode)
    
    If AryItems(aContents) Then
        Do
            aContents(i) = Trim$(aContents(i))
            
            If Left$(aContents(i), 1) = "[" Then
                'begin new section
                sSection = Mid$(aContents(i), 2, Len(aContents(i)) - 2)
                
                If Not mo_Sect.Exists(sSection) Then
                    #If UseHashtable Then
                        mo_Sect.Add sSection, New clsTrickHashTable
                    #Else
                        mo_Sect.Add sSection, CreateObject("Scripting.Dictionary")
                    #End If
                    
                    With mo_Sect(sSection)
                        .CompareMode = mb_CompareMode
                        For j = i + 1 To UBound(aContents)
                            If Left$(aContents(j), 1) = "[" Then Exit For
                            pos = InStr(aContents(j), "=")
                            If pos <> 0 And Not (Left$(aContents(j), 1) = ";") Then
                                sParam = Trim$(Left$(aContents(j), pos - 1))
                                sData = Trim$(Mid$(aContents(j), pos + 1))
                                If Not .Exists(sParam) Then .Add sParam, sData
                            End If
                        Next
                        i = j - 1
                    End With
                End If
            End If
            i = i + 1
        Loop Until i > UBound(aContents)
    End If
    mb_Loaded = True
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_LoadFile", ms_IniFile
    If inIDE Then Stop: Resume Next
End Sub

Private Sub Class_Initialize()
    mb_CompareMode = vbTextCompare
    InitDictionary
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler:
    Flush
    ReleaseDictionary
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_Class_Terminate"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub InitDictionary()
    On Error GoTo ErrorHandler:
    #If UseHashtable Then
        Set mo_Sect = New clsTrickHashTable
    #Else
        Set mo_Sect = CreateObject("Scripting.Dictionary")
    #End If
    mo_Sect.CompareMode = mb_CompareMode
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_InitDictionary"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub ReleaseDictionary()
    On Error GoTo ErrorHandler:
    Dim oKey As Variant
    If Not (mo_Sect Is Nothing) Then
        For Each oKey In mo_Sect.Keys
            Set mo_Sect(CStr(oKey)) = Nothing
        Next
    End If
    Set mo_Sect = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_ReleaseDictionary"
    If inIDE Then Stop: Resume Next
End Sub

'// Update phisical ini file from the cache
Public Function Flush() As Boolean
    On Error GoTo ErrorHandler:

    Dim sb As clsStringBuilder
    Dim sDir As String, sIniFile As String
    Dim oKey As Variant
    Dim oKey2 As Variant
    
    If mb_UnsavedData Then
        
        sIniFile = ms_IniFile
    
        sDir = GetParentDir(sIniFile)
        
        If Not FolderExists(sDir) Then
            If Not MkDirW(sDir) Then
                TryUnlock sDir
                If Not MkDirW(sDir) Then
                    If Not bAutoLogSilent Then
                        'Could not create folder '[]'. Please verify that write access is allowed to this location.
                        MsgBoxW Replace$(Translate(1022), "[]", sDir), vbCritical
                    End If
                    Exit Function
                End If
            End If
        End If
        
        If Not CheckAccessWrite(sIniFile) Then
            TryUnlock sIniFile
            If Not CheckAccessWrite(sIniFile) Then
                If Not bAutoLogSilent Then
                    'The value '[*]' could not be written to the settings file '[**]'. Please verify that write access is allowed to that file.
                    MsgBoxW Replace$(Replace$(Translate(1008), "'[*]' ", ""), "[**]", sIniFile), vbCritical
                End If
                Exit Function
            End If
        End If
        
        Set sb = New clsStringBuilder
        Dim oDict As Object
        If Not (mo_Sect Is Nothing) Then
            For Each oKey In mo_Sect.Keys 'each section
                sb.Append "[" & oKey & "]" & vbCrLf
                Set oDict = mo_Sect(CStr(oKey))
                If Not (oDict Is Nothing) Then
                    For Each oKey2 In oDict.Keys 'each parameter
                        sb.Append oKey2 & "=" & oDict(CStr(oKey2)) & vbCrLf
                    Next
                    Set oDict = Nothing
                End If
                sb.Append vbCrLf
            Next
        End If
        Flush = WriteDataToFile(ms_IniFile, sb.ToString, mb_IsUnicode, True)
        
        mb_UnsavedData = False
        Set sb = Nothing
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_Flush"
    Set sb = Nothing
    If inIDE Then Stop: Resume Next
End Function

'// Read perameter data by its name + section name, where it is stored
Public Function ReadParam(sSection As Variant, sParameter As Variant, Optional vDefault As Variant = vbNullString) As String
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Exists(CStr(sParameter)) Then
            ReadParam = oDict(CStr(sParameter))
        Else
            ReadParam = vDefault
        End If
        Set oDict = Nothing
    Else
        ReadParam = vDefault
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_ReadParam", sSection, sParameter, vDefault
    If inIDE Then Stop: Resume Next
End Function

'// Write (or replace) parameter data
Public Sub WriteParam(sSection As Variant, sParameter As Variant, vData As Variant)
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Exists(CStr(sParameter)) Then
            If oDict(CStr(sParameter)) = vData Then
                Set oDict = Nothing
                Exit Sub
            Else
                oDict(CStr(sParameter)) = vData
            End If
        Else
            oDict.Add CStr(sParameter), vData
        End If
        Set oDict = Nothing
    Else
        #If UseHashtable Then
            mo_Sect.Add CStr(sSection), New clsTrickHashTable
        #Else
            mo_Sect.Add CStr(sSection), CreateObject("Scripting.Dictionary")
        #End If
        mo_Sect(CStr(sSection)).CompareMode = mb_CompareMode
        mo_Sect(CStr(sSection)).Add CStr(sParameter), vData
    End If
    mb_UnsavedData = True
    '// TODO: temporarily
    'Flush
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_ReadParam", sSection, sParameter, vData
    If inIDE Then Stop: Resume Next
End Sub

'// Get the list of all sections to array
Public Function GetSections() As Variant()
    On Error GoTo ErrorHandler:
    Dim bExist As Boolean
    If Not (mo_Sect Is Nothing) Then
        If mo_Sect.Count <> 0 Then
            GetSections = mo_Sect.Keys
            bExist = True
        End If
    End If
    If Not bExist Then GetSections = Array()
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_GetSections"
    If inIDE Then Stop: Resume Next
End Function

'// Get the list of all parameter names to array by section name, where there are stored
Public Function GetParamNames(sSection As Variant) As Variant()
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    Dim bExist As Boolean
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Count <> 0 Then
            GetParamNames = oDict.Keys
            bExist = True
        End If
        Set oDict = Nothing
    End If
    If Not bExist Then GetParamNames = Array()
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_GetParamNames", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Get the list of all data of parameters to array by section name, where there are stored
Public Function GetParamValues(sSection As Variant) As Variant()
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    Dim bExist As Boolean
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Count <> 0 Then
            GetParamValues = oDict.Items
            bExist = True
        End If
        Set oDict = Nothing
    End If
    If Not bExist Then GetParamValues = Array()
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_GetParamValues", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Check if section name is exist
Public Function ExistSection(sSection As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If Not (mo_Sect Is Nothing) Then
        ExistSection = mo_Sect.Exists(CStr(sSection))
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_ExistSection", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Check if parameter name in particular section is exist
Public Function ExistParam(sSection As Variant, sParameter As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If ExistSection(sSection) Then
        ExistParam = mo_Sect(CStr(sSection)).Exists(CStr(sParameter))
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_ExistParam", sSection, sParameter
    If inIDE Then Stop: Resume Next
End Function

'// Check if concrete data is exist in section
Public Function ExistData(sSection As Variant, sData As Variant) As Boolean
    On Error GoTo ErrorHandler:
    Dim aNames() As Variant
    Dim i As Long
    aNames = GetParamNames(sSection)
    If AryItems(aNames) Then
        For i = 0 To UBound(aNames)
            If StrComp(CStr(sData), ReadParam(sSection, aNames(i)), mb_CompareMode) = 0 Then
                ExistData = True
                Exit For
            End If
        Next
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_ExistData", sSection, sData
    If inIDE Then Stop: Resume Next
End Function

'// Get first parameter name by data stored there in particular section
Public Function GetParamNameByData(sSection As Variant, sData As Variant) As String
    On Error GoTo ErrorHandler:
    Dim aNames() As Variant
    Dim i As Long
    aNames = GetParamNames(sSection)
    If AryItems(aNames) Then
        For i = 0 To UBound(aNames)
            If StrComp(CStr(sData), ReadParam(sSection, aNames(i)), mb_CompareMode) = 0 Then
                GetParamNameByData = aNames(i)
                Exit For
            End If
        Next
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_GetParamNameByData", sSection, sData
    If inIDE Then Stop: Resume Next
End Function

'// Create empty section, if it is not exist yet
Public Function CreateSection(sSection As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If Not ExistSection(sSection) Then
        #If UseHashtable Then
            mo_Sect.Add CStr(sSection), New clsTrickHashTable
        #Else
            mo_Sect.Add CStr(sSection), CreateObject("Scripting.Dictionary")
        #End If
        mo_Sect(CStr(sSection)).CompareMode = mb_CompareMode
        CreateSection = True
        mb_UnsavedData = True
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_CreateSection", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Remove ALL sections
Public Sub RemoveSectionsAll()
    On Error GoTo ErrorHandler:
    ReleaseDictionary
    InitDictionary
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_RemoveSectionsAll"
    If inIDE Then Stop: Resume Next
End Sub

'// Remove concrete section with all parameters in it
Public Function RemoveSection(sSection As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If ExistSection(sSection) Then
        mo_Sect.Remove CStr(sSection)
        RemoveSection = True
        mb_UnsavedData = True
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_RemoveSection", sSection
    If inIDE Then Stop: Resume Next
End Function

'// Remove parameter in particular section
Public Function RemoveParam(sSection As Variant, sParameter As Variant) As Boolean
    On Error GoTo ErrorHandler:
    If ExistParam(sSection, sParameter) Then
        mo_Sect(CStr(sSection)).Remove CStr(sParameter)
        RemoveParam = True
        mb_UnsavedData = True
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsIniFile_RemoveParam", sSection, sParameter
    If inIDE Then Stop: Resume Next
End Function

'// Rename parameter
Public Sub RenameParam(sSection As Variant, sParameter As Variant, sNewParam As Variant)
    On Error GoTo ErrorHandler:
    Dim oDict As Object
    If ExistSection(sSection) Then
        Set oDict = mo_Sect(CStr(sSection))
        If oDict.Exists(CStr(sParameter)) Then
            Me.WriteParam sSection, sNewParam, oDict(CStr(sParameter))
            oDict.Remove CStr(sParameter)
        End If
        Set oDict = Nothing
    End If
    mb_UnsavedData = True
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsIniFile_RenameParam", sSection, sParameter, sNewParam
    If inIDE Then Stop: Resume Next
End Sub

'// Get currently loaded file name
Public Property Get FileName() As String
    FileName = ms_IniFile
End Property

'// Get number of sections
Public Property Get CountSections() As Long
    On Error GoTo ErrorHandler:
    If Not (mo_Sect Is Nothing) Then
        CountSections = mo_Sect.Count
    End If
    Exit Property
ErrorHandler:
    ErrorMsg Err, "clsIniFile_CountSections"
    If inIDE Then Stop: Resume Next
End Property

'// Get number of parameters within section
Public Property Get CountParams(sSection As Variant) As Long
    On Error GoTo ErrorHandler:
    If Not (mo_Sect Is Nothing) Then
        If ExistSection(sSection) Then
            CountParams = mo_Sect(CStr(sSection)).Count
        End If
    End If
    Exit Property
ErrorHandler:
    ErrorMsg Err, "clsIniFile_CountParams", sSection
    If inIDE Then Stop: Resume Next
End Property
