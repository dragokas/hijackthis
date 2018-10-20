Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim sFile   As String
    Dim ff      As Integer
    Dim Size    As Long
    Dim Rest    As Long
    
    sFile = Command()
    
    If Len(sFile) = 0 Then MsgBox "Использование: " & App.EXEName & " " & "file.txt": End
    
    ff = FreeFile()
    
    sFile = UnQuote(sFile)
    
    Open sFile For Binary Access Read Write As #ff
        Size = LOF(ff)
        Rest = 4 - (Size Mod 4)
        If Rest <> 0 And Rest <> 4 Then
            Put #ff, Size + 1, String$(Rest, " ")
        End If
    Close #ff
    
End Sub

Function UnQuote(Str As String) As String   ' Убрать обрамление кавычками
    Dim s As String: s = Str
    Do While Left$(s, 1&) = """"
        s = Mid$(s, 2&)
    Loop
    Do While Right$(s, 1&) = """"
        s = Left$(s, Len(s) - 1&)
    Loop
    UnQuote = s
End Function
