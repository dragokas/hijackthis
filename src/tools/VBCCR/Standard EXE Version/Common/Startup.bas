Attribute VB_Name = "Startup"
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If

Sub Main()
If App.PrevInstance = True And InIDE() = False Then
    Dim hWnd As LongPtr
    hWnd = FindWindow(StrPtr("ThunderRT6FormDC"), StrPtr("ComCtls Demo"))
    If hWnd <> NULL_PTR Then
        Const SW_RESTORE As Long = 9
        ShowWindow hWnd, SW_RESTORE
        SetForegroundWindow hWnd
        AppActivate "ComCtls Demo"
    End If
Else
    Call InitVisualStylesFixes
    MainForm.Show vbModeless
End If
End Sub
