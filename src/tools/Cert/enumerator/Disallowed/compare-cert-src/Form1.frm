VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oAll As Object

Private Sub Form_Load()
    
    Set oAll = CreateObject("Scripting.Dictionary")
    
    'get all hashes
    Dim ff%, s$
    ff = FreeFile()
    Open App.Path & "\Hashes.csv" For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, s
        SaveHash s
    Loop
    Close #ff
    
    Open App.Path & "\hjt.txt" For Input As #ff
    
    Do While Not EOF(ff)
        Line Input #ff, s
        RemoveHash s
    Loop
    Close #ff
    
    Dim key As Variant
    Open App.Path & "\Hashes_new.txt" For Output As #ff
    For Each key In oAll.keys
        Print #ff, "        .Add """ & oAll(key) & """, """ & key & """"
    Next
    Close #ff
    
    Set oAll = Nothing
    
    Unload Me
End Sub

Sub RemoveHash(s$)
    Dim sName As String
    Dim sHash As String
    Dim arr
    If Len(s) = 0 Then Exit Sub
    
    arr = Split(s, """")
    If UBound(arr) >= 3 Then
        sName = arr(1)
        sHash = arr(3)
        If oAll.Exists(sHash) Then
            oAll.Remove sHash
            Debug.Print "[REMOVED] " & sHash & " - " & sName
        End If
    End If
End Sub

Sub SaveHash(s$)
    If Len(s) = 0 Then Exit Sub
    Dim pos&
    Dim sName As String
    Dim sHash As String
    pos = InStr(s, ";")
    If pos <> 0 Then
        sName = Left$(s, pos - 1)
        If sName <> "Certificate name" Then
            s = Mid$(s, pos + 1)
            pos = InStr(s, ";")
            If pos <> 0 Then
                sHash = Left$(s, pos - 1)
                If Not oAll.Exists(sHash) Then
                    oAll.Add sHash, sName
                    Debug.Print "[ADDED] " & sHash & " - " & sName
                End If
            End If
        End If
    End If
End Sub
