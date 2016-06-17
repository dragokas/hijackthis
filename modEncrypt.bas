Attribute VB_Name = "modEncrypt"
Option Explicit

Public Function Crypt$(sMsg$, sPhrase$, Optional doCrypt As Boolean = False)  'if Crypt = False then we do decryption
    Dim sEncryptionPhrase$
    On Error GoTo ErrorHandler:
    'if one error happens, don't screw up everything following
    
    'like, NOT!
    sEncryptionPhrase = "FUCK YOU SPYWARENUKER AND BPS SPYWARE REMOVER!"
    
    'bEnOrDec = false -> decrypt
    'bEnOrDec = true -> encrypt
    
    Dim i&, J&, sChar$, iChar&, sOut$
    J = 1
    For i = 1 To Len(sMsg)
        sChar = Mid$(sMsg, i, 1)
        If doCrypt Then
            'encrypt
            sChar = Chr(Asc(sChar) + Asc(Mid$(sPhrase, J, 1)))
            If iChar > 255 Then Exit Function 'Wrong Pass phrase
            If Asc(sChar) > 126 Then
                'make sure encrypted char is within
                'normal range (space to ~)
                sChar = Chr(Asc(sChar) - 94)
            End If
        Else
            'decrypt
            iChar = Asc(sChar) - Asc(Mid$(sPhrase, J, 1))
            If iChar < -94 Then Exit Function 'Wrong Pass phrase
            If iChar < 32 Then
                'make sure decrypted char is within
                'normal range (space to ~)
                sChar = Chr(iChar + 94)
            Else
                'old encrypter doesn't encrypt chars above 126 :(
                If Asc(sChar) < 192 Then
                    sChar = Chr(iChar)
                End If
            End If
        End If
        sOut = sOut & sChar
        J = J + 1
        If J > Len(sPhrase) Then J = 1
    Next i
    Crypt = sOut
    Exit Function
ErrorHandler:
    ErrorMsg err, "Crypt", sMsg
    If inIDE Then Stop: Resume Next
End Function
