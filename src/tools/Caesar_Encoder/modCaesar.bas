Attribute VB_Name = "modCaesar"
Option Explicit

Private Const CAESAR_SEED = 1
Private Const CAESAR_STEP = 2

Public Function Caesar_Encode(original As String) As String
    On Error GoTo ErrorHandler:
    Dim seed As Long
    Dim encoded As String
    Dim i As Long, Code As Long
    seed = CAESAR_SEED
    For i = 1 To Len(original)
        Code = Asc(Mid(original, i, 1))
        If Code >= Asc("0") And Code <= Asc("9") Then
            Code = Code + seed
            Do While Code > Asc("9"): Code = Code - Asc("9") + Asc("0") - 1: Loop
        ElseIf (Code >= Asc("A") And Code <= Asc("z")) Then
            Code = Code + seed
            Do While Code > Asc("z"): Code = Code - Asc("z") + Asc("A") - 1: Loop
        End If
        encoded = encoded & Chr$(Code)
        seed = seed + CAESAR_STEP
    Next
    Caesar_Encode = encoded
    Exit Function
ErrorHandler:
    MsgBox "Error in: " & "Caesar_Encode" & ". " & Err.Description
End Function

Public Function Caesar_Decode(encoded As String) As String
    On Error GoTo ErrorHandler:
    Dim seed As Long
    Dim i As Long, Code As Long
    seed = CAESAR_SEED
    Dim decoded As String
    For i = 1 To Len(encoded)
        Code = Asc(Mid(encoded, i, 1))
        If Code >= Asc("0") And Code <= Asc("9") Then
            Code = Code - seed
            Do While Code < Asc("0"): Code = Code + Asc("9") - Asc("0") + 1: Loop
        ElseIf Code >= Asc("A") And Code <= Asc("z") Then
            Code = Code - seed
            Do While Code < Asc("A"): Code = Code + Asc("z") - Asc("A") + 1: Loop
        End If
        decoded = decoded & Chr$(Code)
        seed = seed + CAESAR_STEP
    Next
    Caesar_Decode = decoded
    Exit Function
ErrorHandler:
    MsgBox "Error in: " & "Caesar_Decode" & ". " & Err.Description
End Function
