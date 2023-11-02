Attribute VB_Name = "mCaesar"
Option Explicit

Public Sub Caesar_EncodeBin(original() As Byte, Optional initial_seed As Long = 1, Optional stepping As Long = 2, Optional skip As Long = 3)
    On Error GoTo ErrorHandler:
    Dim seed As Long
    Dim encoded As String
    Dim i As Long, Code As Long
    seed = initial_seed
    For i = 0 To UBound(original) Step skip
        Code = original(i)
        Code = Code + seed
        If Code >= 256 Then Code = Code Mod 256
        original(i) = Code
        seed = seed + stepping
        If seed >= 256 Then seed = seed Mod 256
    Next
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Caesar_EncodeBin"
    If inIde Then Stop: Resume Next
End Sub

Public Sub Caesar_DecodeBin(original() As Byte, Optional initial_seed As Long = 1, Optional stepping As Long = 2, Optional skip As Long = 3)
    On Error GoTo ErrorHandler:
    Dim seed As Long
    Dim i As Long, Code As Long
    seed = initial_seed
    Dim decoded As String
    For i = 0 To UBound(original) Step skip
        Code = original(i)
        Code = Code - seed
        If Code < 0 Then Code = (Code Mod 256) + 256
        original(i) = Code
        seed = seed + stepping
        If seed >= 256 Then seed = seed Mod 256
    Next
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Caesar_DecodeBin"
    If inIde Then Stop: Resume Next
End Sub

'Useless for large text
Public Function Caesar_Encode(original As String, Optional initial_seed As Long = 1, Optional stepping As Long = 2) As String
    On Error GoTo ErrorHandler:
    Dim seed As Long
    Dim i As Long, Code As Long
    seed = initial_seed
    Caesar_Encode = String$(Len(original), 0&)
    For i = 1 To Len(original)
        Code = Asc(Mid(original, i, 1))
        If Code >= Asc("0") And Code <= Asc("9") Then
            Code = Code + seed
            Do While Code > Asc("9"): Code = Code - Asc("9") + Asc("0") - 1: Loop
        ElseIf (Code >= Asc("A") And Code <= Asc("z")) Then
            Code = Code + seed
            Do While Code > Asc("z"): Code = Code - Asc("z") + Asc("A") - 1: Loop
        End If
        Mid$(Caesar_Encode, i) = Chr$(Code)
        seed = seed + stepping
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Caesar_Encode"
    If inIde Then Stop: Resume Next
End Function

'Useless for large text
Public Function Caesar_Decode(encoded As String, Optional initial_seed As Long = 1, Optional stepping As Long = 2) As String
    On Error GoTo ErrorHandler:
    Dim seed As Long
    Dim i As Long, Code As Long
    seed = initial_seed
    Caesar_Decode = String$(Len(encoded), 0&)
    For i = 1 To Len(encoded)
        Code = Asc(Mid(encoded, i, 1))
        If Code >= Asc("0") And Code <= Asc("9") Then
            Code = Code - seed
            Do While Code < Asc("0"): Code = Code + Asc("9") - Asc("0") + 1: Loop
        ElseIf Code >= Asc("A") And Code <= Asc("z") Then
            Code = Code - seed
            Do While Code < Asc("A"): Code = Code + Asc("z") - Asc("A") + 1: Loop
        End If
        Mid$(Caesar_Decode, i) = Chr$(Code)
        seed = seed + stepping
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Caesar_Decode"
    If inIde Then Stop: Resume Next
End Function

