Attribute VB_Name = "mConsole"
Option Explicit

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleW" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleW" (ByVal hConsoleInput As Long, lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function CommandLineToArgvW Lib "Shell32.dll" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

Private m_hStdOut    As Long
Private m_hStdErr    As Long

Public Sub InitConsole()
    m_hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    m_hStdErr = GetStdHandle(STD_ERROR_HANDLE)
End Sub


Public Sub WriteStdout(ByVal txt As String)
    If inIde Then
        Debug.Print "[out] " & txt
    Else
        txt = txt & vbNewLine
        WriteConsole m_hStdOut, ByVal StrPtr(txt), Len(txt), 0, ByVal 0&
    End If
End Sub

Public Sub WriteStderr(ByVal txt As String)
    If inIde Then
        Debug.Print "[err] " & txt
    Else
        txt = txt & vbNewLine
        WriteConsole m_hStdErr, ByVal StrPtr(txt), Len(txt), 0, ByVal 0&
    End If
End Sub

Public Function ParseCommandLine(Line As String, out() As String) As Boolean
    Dim ptr     As Long
    Dim Count   As Long
    Dim index   As Long
    Dim strLen  As Long
    Dim strAdr  As Long
    If Len(Line) = 0 Then ReDim out(0): Exit Function
    ptr = CommandLineToArgvW(StrPtr(Line), Count)
    If Count < 1& Then Exit Function
    ReDim out(Count)
    out(0) = App.Path & "\" & App.EXEName & ".exe"
    If Len(Line) = 0 Then Exit Function
    For index = 0& To Count - 1&
        GetMem4 ByVal ptr + index * 4&, strAdr
        strLen = lstrlen(strAdr)
        out(index + 1) = Space(strLen)
        lstrcpyn StrPtr(out(index + 1)), strAdr, strLen + 1&
    Next
    GlobalFree ptr
    ParseCommandLine = True
End Function

Public Sub ExitProcessVB(iExitcode As Long)
    If inIde Then
        End
    Else
        ExitProcess iExitcode
    End If
End Sub
