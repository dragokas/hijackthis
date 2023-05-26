Attribute VB_Name = "VisualStyles"
Option Explicit
Public Declare Function ActivateVisualStyles Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
Public Declare Function RemoveVisualStyles Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long
Private Type DLLVERSIONINFO
cbSize As Long
dwMajor As Long
dwMinor As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type
Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function IsAppThemed Lib "uxtheme" () As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Long
Private Declare Function GetThemeAppProperties Lib "uxtheme" () As Long
Private Const STAP_ALLOW_CONTROLS As Long = (1 * (2 ^ 1))
Private Const S_OK As Long = &H0

Public Function EnabledVisualStyles() As Boolean
If GetComCtlVersion() >= 6 Then
    If IsThemeActive() <> 0 Then
        If IsAppThemed() <> 0 Then
            EnabledVisualStyles = True
        ElseIf (GetThemeAppProperties() And STAP_ALLOW_CONTROLS) <> 0 Then
            EnabledVisualStyles = True
        End If
    End If
End If
End Function

Public Function GetComCtlVersion() As Long
Static Done As Boolean, Value As Long
If Done = False Then
    Dim Version As DLLVERSIONINFO
    On Error Resume Next
    Version.cbSize = LenB(Version)
    If DllGetVersion(Version) = S_OK Then Value = Version.dwMajor
    Done = True
End If
GetComCtlVersion = Value
End Function
