Attribute VB_Name = "modEncrypt"
Public Function Crypt(ByRef sMsg As String, ByRef sPhrase As String, Optional ByVal bEnOrDec As Boolean) As String
    Dim i As Long, j As Long, iChar As Integer, sChar As String

    j = 1&

    For i = 1& To Len(sMsg)
        sChar = Mid$(sMsg, i, 1&)

        If bEnOrDec Then 'Encrypt
            sChar = Chr$(Asc(sChar) + Asc(Mid$(sPhrase, j, 1&)))
            If Asc(sChar) > 126 Then
               'Make sure encrypted char is within normal range (space to ~)
                sChar = Chr$(Asc(sChar) - 94)
            End If
        Else 'Decrypt
            iChar = Asc(sChar) - Asc(Mid$(sPhrase, j, 1&))
            If iChar < 32 Then
               'Make sure decrypted char is within normal range (space to ~)
                sChar = Chr$(iChar + 94)
            ElseIf Asc(sChar) < 192 Then
               'Old encrypter doesn't encrypt chars above 126 :(
                sChar = Chr$(iChar)
            End If
        End If

        Crypt = Crypt & sChar
        If j <= Len(sPhrase) Then j = j + 1& Else j = 1&
    Next i
End Function
