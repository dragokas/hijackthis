Attribute VB_Name = "modUtils"
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

Public Function NormalizePath$(sFile$)
    
    Dim sBegin$, sValue$, sNext$
    Dim EnvVar As String
    Dim RealEnvVar As String
    
    If False Then
    Dim EnvRegExp As RegExp
    Dim ObjMatch As Match
    Dim ObjMatches As MatchCollection
    'Dim EnvVar As String
    
    Set EnvRegExp = New RegExp
    EnvRegExp.Pattern = "%[\w_-]+%"
    EnvRegExp.IgnoreCase = True
    EnvRegExp.Global = True
    
    If EnvRegExp.Test(sFile) = True Then
        Set ObjMatches = EnvRegExp.Execute(sFile)
        For Each ObjMatch In ObjMatches
            EnvVar = Replace(ObjMatch.Value, "%", "", , , vbTextCompare)
            If Len(Environ$(EnvVar)) > 0 Then
                sFile = Replace(sFile, ObjMatch.Value, Environ$(EnvVar), , , vbTextCompare)
            End If
        Next
    End If
    End If
    
'If False Then
    sBegin = 1
    Do
        sValue = InStr(sBegin, sFile, "%", vbTextCompare)
        If sValue = 0 Or sValue = Len(sFile) Or sBegin > Len(sFile) Then
            Exit Do
        End If
            
        sBegin = sValue + 1
        sNext = InStr(sBegin + 1, sFile, "%", vbTextCompare)
        If sNext = 0 Or sNext > Len(sFile) Or sBegin > Len(sFile) Then
            Exit Do
        End If
        
        EnvVar = Mid(sFile, sValue, sNext - sValue + 1)
        RealEnvVar = Mid(sFile, sValue + 1, sNext - sValue - 1)
        
        If Len(Environ$(RealEnvVar)) > 0 Then
            sFile = Replace(sFile, EnvVar, Environ$(RealEnvVar), sValue, sNext - sValue + 1, vbTextCompare)
            sBegin = sNext + 1 + Len(Environ$(RealEnvVar)) - Len(EnvVar)
        Else
            sBegin = sNext + 1
        End If
        
    Loop While True
    'End If
    NormalizePath = sFile
End Function

Public Function GetChromeVersion$()
    Dim sItems$(), sName$, sVer$, ChromeVer$
    Dim i&
    
    sItems = Split(RegEnumSubkeys(HKEY_CURRENT_USER, "Software\Google\Update\Clients"), "|")
    If UBound(sItems) <> -1 Then
        For i = 0 To UBound(sItems)
            sName = RegGetString(HKEY_CURRENT_USER, "Software\Google\Update\Clients\" & sItems(i), "name")
            sVer = RegGetString(HKEY_CURRENT_USER, "Software\Google\Update\Clients\" & sItems(i), "pv")
            If sName Like "*Chrome*" And sVer <> vbNullString Then
                ChromeVer = "CHROME: " & sVer
                Exit For
            End If
        Next i
    End If

    GetChromeVersion = ChromeVer
End Function

Public Function GetFirefoxVersion$()
    Dim sVer$, FirefoxVer$
    Dim i&
    
    sVer = RegGetString(HKEY_LOCAL_MACHINE, "Software\Mozilla\Mozilla Firefox", "CurrentVersion")
    If sVer <> vbNullString Then
        FirefoxVer = "FIREFOX: " & sVer
    End If

    GetFirefoxVersion = FirefoxVer
End Function
