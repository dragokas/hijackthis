Attribute VB_Name = "mDebug"
Option Explicit

Public inIde   As Boolean

Public Sub Init()
    Debug.Assert MakeTrue(inIde)
    InitConsole
End Sub

Public Function MakeTrue(Value As Boolean) As Boolean
    Value = True
    MakeTrue = True
End Function

Public Sub ErrorMsg(Err As ErrObject, sMsg As String)
    Debug.Print "Error: " & Err.Number & ". LastDllErr: " & Err.LastDllError & ". " & sMsg
End Sub

