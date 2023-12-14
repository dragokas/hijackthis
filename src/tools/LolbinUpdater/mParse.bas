Attribute VB_Name = "mParse"
Option Explicit

Public Function StrBeginWith(Text As String, BeginPart As String) As Boolean
    StrBeginWith = (StrComp(Left$(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function

Public Function StrEndWith(Text As String, LastPart As String) As Boolean
    StrEndWith = (StrComp(Right$(Text, Len(LastPart)), LastPart, 1) = 0)
End Function

Public Function Caes_Decode(encoded As String, Optional initial_seed As Long = 1, Optional stepping As Long = 2) As String
    On Error GoTo ErrorHandler:
    Dim seed As Long
    Dim i As Long, Code As Long
    seed = initial_seed
    Caes_Decode = String$(Len(encoded), 0&)
    For i = 1 To Len(encoded)
        Code = Asc(Mid(encoded, i, 1))
        If Code >= Asc("0") And Code <= Asc("9") Then
            Code = Code - seed
            Do While Code < Asc("0"): Code = Code + Asc("9") - Asc("0") + 1: Loop
        ElseIf Code >= Asc("A") And Code <= Asc("z") Then
            Code = Code - seed
            Do While Code < Asc("A"): Code = Code + Asc("z") - Asc("A") + 1: Loop
        End If
        Mid$(Caes_Decode, i) = Chr$(Code)
        seed = seed + stepping
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Caes_Decode"
    If inIDE Then Stop: Resume Next
End Function
