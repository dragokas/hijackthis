VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MapFileAndCheckSum Lib "Imagehlp.dll" Alias "MapFileAndCheckSumW" (ByVal Filename As Long, HeaderSum As Long, CheckSum As Long) As Long

Private Sub Form_Load()
    Dim ff          As Integer
    Dim Filename    As String
    Dim CSum        As Long
    Dim HeaderSum   As Long
    Dim lr          As Long
    
    ff = FreeFile()
    
    Filename = "h:\_AVZ\Наши разработки\Check Browsers LNK\Check Browsers LNK.exe"
    
    lr = MapFileAndCheckSum(StrPtr(Filename), HeaderSum, CSum)
    
    Debug.Print CSum
    Debug.Print "ret=" & ret
    
    
    Exit Sub
    
    Open Filename For Binary Access Read Write As #ff
    
    Close #ff

    

End Sub
