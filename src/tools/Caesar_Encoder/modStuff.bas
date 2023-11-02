Attribute VB_Name = "modStuff"
Option Explicit

Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32.dll" () As Long
Public Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Public Declare Function EmptyClipboard Lib "user32.dll" () As Long
Public Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Public Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal iMaxLength As Long) As Long

Public Const HWND_TOPMOST As Long = -1&
Public Const HWND_NOTOPMOST As Long = -2&
Public Const SWP_NOMOVE As Long = 2&
Public Const SWP_NOSIZE As Long = 1&
Public Const CF_UNICODETEXT    As Long = 13&
Public Const GMEM_MOVEABLE     As Long = &H2&
Public Const CF_LOCALE         As Long = 16

Public Sub SetWindowAlwaysOnTop(hWnd As Long, Enable As Boolean)
    SetWindowPos hWnd, IIf(Enable, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function ClipboardSetText(sText As String) As Boolean
    Dim LangNonUnicodeCode As Long
    LangNonUnicodeCode = GetSystemDefaultLCID Mod &H10000
    
    Dim hMem As Long
    Dim ptr As Long
    If OpenClipboard(Form1.hWnd) Then
        EmptyClipboard
        If Len(sText) <> 0 Then
            hMem = GlobalAlloc(GMEM_MOVEABLE, 4)
            If hMem <> 0 Then
                ptr = GlobalLock(hMem)
                If ptr <> 0 Then
                    GetMem4 LangNonUnicodeCode, ByVal ptr
                    GlobalUnlock hMem
                    If SetClipboardData(CF_LOCALE, hMem) = 0 Then
                        GlobalFree hMem
                    End If
                End If
            End If
            hMem = GlobalAlloc(GMEM_MOVEABLE, LenB(sText) + 2)
            If hMem <> 0 Then
                ptr = GlobalLock(hMem)
                If ptr <> 0 Then
                    lstrcpyn ByVal ptr, ByVal StrPtr(sText), LenB(sText)
                    GlobalUnlock hMem
                    ClipboardSetText = SetClipboardData(CF_UNICODETEXT, hMem)
                    If Not ClipboardSetText Then
                        GlobalFree hMem
                    End If
                End If
            End If
        End If
        CloseClipboard
    End If
End Function

Public Function ClipboardGetText() As String
    Dim hMem As Long
    Dim ptr  As Long
    Dim Size As Long
    Dim txt  As String
    If OpenClipboard(Form1.hWnd) Then
        hMem = GetClipboardData(CF_UNICODETEXT)
        If hMem Then
            Size = GlobalSize(hMem)
            If Size Then
                txt = Space$(Size \ 2 - 1)
                ptr = GlobalLock(hMem)
                lstrcpyn ByVal StrPtr(txt), ByVal ptr, Size
                GlobalUnlock hMem
                ClipboardGetText = txt
            End If
        End If
        CloseClipboard
    End If
End Function
