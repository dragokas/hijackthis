Attribute VB_Name = "modRegistry"
Option Explicit

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long

Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal lRootKey As Long, ByVal szKeyToDelete As String) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
'Public Const HKEY_PERFORMANCE_DATA = &H80000004
'Public Const HKEY_CURRENT_CONFIG = &H80000005
'Public Const HKEY_DYN_DATA = &H80000006

Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const SYNCHRONIZE = &H100000
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const REG_OPTION_NON_VOLATILE = 0

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Private Const REG_EXPAND_SZ = 2

Public lEnumBufSize& 'for RegEnumValue()

Public Sub RegCreateKey(lHive&, sKey$)
    Dim hKey&
    RegCreateKeyEx lHive, sKey, 0, vbNullString, 0, KEY_CREATE_SUB_KEY, ByVal 0, hKey, 0
    RegCloseKey hKey
End Sub

Public Function RegGetString$(lHive&, sKey$, sValue$)
    Dim hKey&, lRet&, uData() As Byte, sData$, lDataLen&, i%
    On Error GoTo Error:
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        Exit Function
    End If
    ReDim uData(260)
    lDataLen = 260
    lRet = RegQueryValueEx(hKey, sValue, 0, REG_SZ, uData(0), lDataLen)
    If lRet <> 0 Then
        If lDataLen > 260 And lRet = 234 Then
            'not enough room in string,
            'so enlarge buffer
            ReDim uData(lDataLen)
            lDataLen = UBound(uData)
            RegQueryValueEx hKey, sValue, 0, REG_SZ, uData(0), lDataLen
        Else
            RegCloseKey hKey
            Exit Function
        End If
    End If
    RegCloseKey hKey
    'with help from 'Adult' in Japan
    sData = StrConv(uData, vbUnicode)
    sData = Left(sData, InStr(sData, Chr(0)) - 1)
    'sData = ""
    'For i = 0 To lDataLen
    '    If uData(i) = 0 Then Exit For
    '    sData = sData & Chr(uData(i))
    'Next i
    RegGetString = sData
    Exit Function
    
Error:
    RegCloseKey hKey
    ErrorMsg "modRegistry_RegGetString", Err.Number, Err.Description, "sKey=" & sKey & ",sValue=" & sValue
End Function

Public Function RegGetDword&(lHive&, sKey$, sValue$)
    Dim hKey&, lData&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        Exit Function
    End If
    RegQueryValueEx hKey, sValue, 0, REG_DWORD, lData, 4
    RegCloseKey hKey
    RegGetDword = lData
End Function

Public Sub RegSetStringVal(lHive&, sKey$, sValue$, sData$)
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_SET_VALUE, hKey) <> 0 Then
        Exit Sub
    End If
    RegSetValueEx hKey, sValue, 0, REG_SZ, ByVal sData, Len(sData) + 1
    RegCloseKey hKey
End Sub

Public Sub RegSetDwordVal(lHive&, sKey$, sValue$, lData&)
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_SET_VALUE, hKey) <> 0 Then
        Exit Sub
    End If
    RegSetValueEx hKey, sValue, 0, REG_DWORD, lData, 4
    RegCloseKey hKey
End Sub

Public Sub RegDelVal(lHive&, sKey$, sValue$)
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_WRITE, hKey) <> 0 Then
        Exit Sub
    End If
    RegDeleteValue hKey, sValue
    RegCloseKey hKey
End Sub

Public Sub RegDelKey(lHive&, sKey$)
    'RegDeleteKey lHive, sKey
    SHDeleteKey lHive, sKey
End Sub

Public Function RegKeyExists(lHive&, sKey$) As Boolean
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        RegKeyExists = False
    Else
        RegKeyExists = True
        RegCloseKey hKey
    End If
End Function

Public Function RegValueExists(lHive&, sKey$, sValue$) As Boolean
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        RegValueExists = False
        Exit Function
    Else
        If RegQueryValueEx(hKey, sValue, 0, 0, ByVal 0, 0) <> 0 Then
            RegValueExists = False
        Else
            RegValueExists = True
        End If
        RegCloseKey hKey
    End If
End Function

Public Sub RegSave(sName$, sData$)
    RegSetStringVal HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis", sName, sData
End Sub

Public Function RegRead$(sName$, Optional sDefault$)
    RegRead = RegGetString(HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis", sName)
    If RegRead = "" Then RegRead = sDefault
End Function

Public Sub RegDel(sName$)
    RegDelVal HKEY_LOCAL_MACHINE, "Software\TrendMicro\HijackThis", sName
End Sub

Public Function RegKeyHasSubKeys(lHive&, sKey$) As Boolean
    Dim hKey&, sDummy$
    RegKeyHasSubKeys = False
    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sDummy = String(255, 0)
        If RegEnumKeyEx(hKey, 0, sDummy, 255, 0, vbNullString, ByVal 0, ByVal 0) = 0 Then
            RegKeyHasSubKeys = True
        End If
        RegCloseKey hKey
    End If
End Function

Public Sub RegDelSubKeys(lHive&, sKey$)
    'sub is no longer used, superseded with SHDeleteKey
    Dim hKey&, i&, sName$, sSubKeys$()
    On Error GoTo Error:
    ReDim sSubKeys(0)
    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sName = String(255, 0)
        If RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then
            'no subkeys
            RegCloseKey hKey
            Exit Sub
        End If
        Do
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
            ReDim Preserve sSubKeys(UBound(sSubKeys) + 1)
            sSubKeys(UBound(sSubKeys)) = sName
            
            sName = String(255, 0)
            i = i + 1
        Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
        RegCloseKey hKey
    End If
    
    For i = 1 To UBound(sSubKeys)
        If RegKeyHasSubKeys(lHive, sKey & "\" & sSubKeys(i)) Then
            RegDelSubKeys lHive, sKey & "\" & sSubKeys(i)
        End If
        RegDelKey lHive, sKey & "\" & sSubKeys(i)
    Next i
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg "modRegistry_RegDelSubKeys", Err.Number, Err.Description, "sKey=" & sKey
End Sub

Public Function RegGetFirstSubKey$(lHive&, sKey$)
    Dim hKey&, sName$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sName = String(255, 0)
        If RegEnumKeyEx(hKey, 0, sName, 255, 0, vbNullString, 0, ByVal 0) = 0 Then
            RegGetFirstSubKey = Left(sName, InStr(sName, Chr(0)) - 1)
        Else
            RegGetFirstSubKey = vbNullString
        End If
        RegCloseKey hKey
    Else
        RegGetFirstSubKey = vbNullString
    End If
End Function

Public Function FileExists(sFile$) As Boolean
    On Error Resume Next
    Dim sDummy$
    sDummy = Replace(sFile, "\\", "\")
    If bIsWinNT Then
        'FileExists = IIf(Dir(sDummy, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, True, False)
        FileExists = IIf(SHFileExists(StrConv(sDummy, vbUnicode)) = 1, True, False)
    Else
        FileExists = IIf(SHFileExists(sDummy) = 1, True, False)
    End If
End Function

Public Function FolderExists(sFolder$) As Boolean
    On Error Resume Next
    Dim sDummy$
    If InStr(sFolder, "\\") = 1 Then Exit Function 'network path
    If InStr(sFolder, ".zip") > 0 Then Exit Function 'running from zip in XP
    sDummy = Replace(sFolder, "\\", "\")
    If bIsWinNT Then
        'If Right(sDummy, 1) = "\" Then
        '    FolderExists = IIf(Dir(sDummy & "*.*", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, True, False)
        'Else
        '    FolderExists = IIf(Dir(sDummy & "\*.*", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, True, False)
        'End If
        FolderExists = IIf(SHFileExists(StrConv(sDummy, vbUnicode)) = 1, True, False)
    Else
        FolderExists = IIf(SHFileExists(sDummy) = 1, True, False)
    End If
End Function

Public Function GetFirstSubFolder$(sFolder$)
    Dim sBla$
    'this sub caused some stupid error msgs
    '(#52 Bad file or number and
    '#5 Invalid procedure call or argument)
    'which were caused by two Dir cmds interfering
    'commented out folderexists line to fix
    'On Error GoTo Error:
    On Error Resume Next
    sBla = Dir(sFolder & "\", vbDirectory + vbHidden + vbReadOnly + vbSystem)
    If sBla = vbNullString Then Exit Function
    Do
        If sBla <> "." And sBla <> ".." Then
            'If FolderExists(sFolder & "\" & sBla) Then
                If (GetAttr(sFolder & "\" & sBla) And vbDirectory) Then
                    GetFirstSubFolder = sBla
                    Exit Function
                End If
            'End If
        End If
        sBla = Dir
    Loop Until sBla = vbNullString
    Exit Function
    
Error:
    ErrorMsg "modRegistry_GetFirstSubFolder", Err.Number, Err.Description, "sFolder=" & sFolder
End Function

Public Function GetFirstSubKey$(lHive&, sKey$)
    Dim hKey&, sName$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        Exit Function
    End If
    
    sName = String(255, 0)
    If RegEnumKeyEx(hKey, 0, sName, 255, 0, vbNullString, 0, ByVal 0) = 0 Then
        sName = Left(sName, InStr(sName, Chr(0)) - 1)
        GetFirstSubKey = sName
    End If
    RegCloseKey hKey
End Function

Public Function RegKeyHasValues(lHive&, sKey$) As Boolean
    Dim hKey&, sName$, uData() As Byte
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        sName = String(lEnumBufSize, 0)
        ReDim uData(lEnumBufSize)
        If RegEnumValue(hKey, 0, sName, Len(sName), 0, ByVal 0, uData(0), UBound(uData)) = 0 Then
            RegKeyHasValues = True
        Else
            RegKeyHasValues = False
        End If
        RegCloseKey hKey
    Else
        RegKeyHasValues = False
    End If
End Function

Public Function RegEnumSubkeys$(lHive, sKey$)
    Dim hKey&, i&, sName$, sDummy$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        'key doesn't exist
        Exit Function
    End If
    
    sName = String(260, 0)
    If RegEnumKeyEx(hKey, 0, sName, 260, 0, vbNullString, 0, ByVal 0) <> 0 Then
        'key doesn't have subkeys
        RegCloseKey hKey
        Exit Function
    End If
    
    Do
        sName = Left(sName, InStr(sName, Chr(0)) - 1)
        
        sDummy = sDummy & sName & "|"
        
        sName = String(260, 0)
        i = i + 1
    Loop Until RegEnumKeyEx(hKey, i, sName, 260, 0, vbNullString, 0, ByVal 0) <> 0
    RegCloseKey hKey
    
    RegEnumSubkeys = Left(sDummy, Len(sDummy) - 1)
End Function

Public Sub RegSetExpandStringVal(lHive&, sKey$, sValue$, sData$)
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_SET_VALUE, hKey) <> 0 Then
        Exit Sub
    End If
    RegSetValueEx hKey, sValue, 0, REG_EXPAND_SZ, ByVal sData, Len(sData) + 1
    RegCloseKey hKey
End Sub

Public Function IniGetString$(sFile$, sSection$, sValue$)
    Dim sIniFile$, i%, iSect%, vContents As Variant, sData$, iAttr%
    On Error GoTo Error:
    If Not FileExists(sFile) Then
        If FileExists(sWinSysDir & "\" & sFile) Then
            sIniFile = sWinSysDir & "\" & sFile
        End If
        If FileExists(sWinDir & "\" & sFile) Then
            sIniFile = sWinDir & "\" & sFile
        End If
        If Not FileExists(sIniFile) Then Exit Function
    Else
        sIniFile = sFile
    End If
    If FileLen(sIniFile) = 0 Then Exit Function
    
    On Error Resume Next
    iAttr = GetAttr(sIniFile)
    If (iAttr And 2048) Then iAttr = iAttr - 2048 'compression flag
    SetAttr sIniFile, vbNormal
    If Err Then Exit Function
    On Error GoTo Error:
    
    Open sIniFile For Binary As #1
        vContents = Split(Input(FileLen(sIniFile), #1), vbCrLf)
    Close #1
    SetAttr sIniFile, iAttr
    
    If UBound(vContents) = -1 Then Exit Function 'file is empty
    
    For i = 0 To UBound(vContents)
        If Trim(vContents(i)) <> vbNullString Then
            If InStr(1, vContents(i), "[" & sSection & "]", vbTextCompare) = 1 Then
                'found the correct section
                iSect = i
                Exit For
            End If
        End If
    Next i
    If i = UBound(vContents) + 1 Then Exit Function 'section not found
    
    For i = iSect + 1 To UBound(vContents)
        If InStr(vContents(i), "[") = 1 Then Exit For 'value not found
        
        If InStr(1, vContents(i), sValue & "=", vbTextCompare) = 1 Then
            'found the value!
            IniGetString = Mid(vContents(i), InStr(vContents(i), "=") + 1)
            Exit Function
        End If
    Next i
    Exit Function
    
Error:
    Close
    ErrorMsg "modRegistry_IniGetString", Err.Number, Err.Description, "sFile=" & sFile & ", sSection=" & sSection & ", sValue=" & sValue
End Function

Public Sub IniSetString(sFile$, sSection$, sValue$, sData$)
    Dim sIniFile$, i%, iSect%, vContents As Variant, sNewData$, iAttr%
    On Error GoTo Error:
    If Not FileExists(sFile) Then
        If FileExists(sWinSysDir & "\" & sFile) Then
            sIniFile = sWinSysDir & "\" & sFile
        End If
        If FileExists(sWinDir & "\" & sFile) Then
            sIniFile = sWinDir & "\" & sFile
        End If
        If Not FileExists(sIniFile) Then Exit Sub
    Else
        sIniFile = sFile
    End If
    
    On Error Resume Next
    iAttr = GetAttr(sIniFile)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetAttr sIniFile, vbNormal
    If Err Then
        MsgBox "The value '" & sValue & "' could not be written to " & _
               "the settings file '" & sIniFile & "'. Please verify " & _
               "that write access is allowed to that file.", vbCritical
        Exit Sub
    End If
    On Error GoTo Error:
    
    Open sIniFile For Binary As #1
        vContents = Split(Input(FileLen(sIniFile), #1), vbCrLf)
    Close #1
    SetAttr sIniFile, iAttr
    
    If UBound(vContents) = -1 Then Exit Sub 'file is empty
    
    For i = 0 To UBound(vContents)
        If Trim(vContents(i)) <> vbNullString Then
            If InStr(1, vContents(i), "[" & sSection & "]", vbTextCompare) = 1 Then
                'found the correct section
                iSect = i
                Exit For
            End If
        End If
    Next i
    If i = UBound(vContents) + 1 Then Exit Sub 'section not found
    
    For i = iSect + 1 To UBound(vContents)
        If InStr(vContents(i), "[") = 1 Then Exit For 'value not found
        
        If InStr(1, vContents(i), sValue & "=", vbTextCompare) = 1 Then
            'found the value!
            sNewData = Left(vContents(i), InStr(vContents(i), "=")) & sData
            vContents(i) = sNewData
            'input new data, replace file
            SetAttr sIniFile, vbNormal
            DeleteFile sIniFile
            Open sIniFile For Output As #1
                Print #1, Join(vContents, vbCrLf)
            Close #1
            SetAttr sIniFile, iAttr
            Exit Sub
        End If
    Next i
    Exit Sub
    
Error:
    Close
    ErrorMsg "modRegistry_IniSetString", Err.Number, Err.Description, "sFile=" & sFile & ", sSection=" & sSection & ", sValue=" & sValue & ", sData=" & sData
End Sub

Public Sub CreateUninstallKey(bCreate As Boolean)
    If bNoWriteAccess Then bCreate = False
    If bCreate Then
        If RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis", "DisplayName") <> _
                   "HijackThis " & App.Major & "." & App.Minor & "." & App.Revision Then
            RegCreateKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis"
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis", "DisplayName", "HijackThis " & App.Major & "." & App.Minor & "." & App.Revision
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis", "UninstallString", """" & App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "HijackThis.exe"" /uninstall"
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis", "DisplayIcon", App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "HijackThis.exe"
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis", "DisplayVersion", App.Major & "." & App.Minor & "." & App.Revision
            RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis", "Publisher", "TrendMicro"
            'RegSetStringVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis", "URLInfoAbout", "http://www.spywareinfo.com/~merijn/"
        End If
    Else
        RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HijackThis"
        RegDelKey HKEY_CURRENT_USER, "Software\Microsoft\Installer\Products\8A9C1670A3F861244B7A7BFAFB422AA4"
    End If
End Sub
