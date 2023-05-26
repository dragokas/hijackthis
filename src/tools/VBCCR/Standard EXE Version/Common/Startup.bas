Attribute VB_Name = "Startup"
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Sub Main()
If App.PrevInstance = True And InIDE() = False Then
    Dim hWnd As Long
    hWnd = FindWindow(StrPtr("ThunderRT6FormDC"), StrPtr("ComCtls Demo"))
    If hWnd <> 0 Then
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
