Attribute VB_Name = "mDebug"
Option Explicit

Public inIDE   As Boolean

Public Sub Init()
    Debug.Assert MakeTrue(inIDE)
    InitConsole
    Set OSver = New clsOSInfo
End Sub

Public Function MakeTrue(Value As Boolean) As Boolean
    Value = True
    MakeTrue = True
End Function

Public Sub ErrorMsg(Err As ErrObject, sMsg As String)
    Debug.Print "Error: " & Err.Number & ". LastDllErr: " & Err.LastDllError & ". " & sMsg
End Sub

