Attribute VB_Name = "mEnv"
Option Explicit

Private Const MAX_PATH As Long = 260

Private SysDisk         As String
Private PF_32           As String
Private PF_64           As String
Private PF_32_Common    As String
Private PF_64_Common    As String

Private OSver           As clsOSInfo

Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long
Private Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long


Private Sub Init()
    Static bInit As Boolean
    If bInit Then Exit Sub
    bInit = True
    
    Dim lr As Long
    Dim sWinDir As String
    
    Set OSver = New clsOSInfo
    
    sWinDir = String$(MAX_PATH, 0)
    lr = GetSystemWindowsDirectory(StrPtr(sWinDir), MAX_PATH)
    If lr Then
        sWinDir = Left$(sWinDir, lr)
        SysDisk = Left$(SysDisk, 2)
    End If
    
    If OSver.IsWin64 Then
        If OSver.MajorMinor >= 6.1 Then     'Win 7 and later
            PF_64 = EnvironW("%ProgramW6432%")
        Else
            PF_64 = SysDisk & "\Program Files"
        End If
        PF_32 = EnvironW("%ProgramFiles%", True)
    Else
        PF_32 = EnvironW("%ProgramFiles%")
        PF_64 = PF_32
    End If
    
    PF_32_Common = PF_32 & "\Common Files"
    PF_64_Common = PF_64 & "\Common Files"
    
End Sub

Public Function Unexpand(ByVal sPath) As String
    If InStr(1, sPath, "%") <> 0 Then
        sPath = EnvironW(sPath)
    End If
        
    sPath = Replace$(sPath, "C:\Windows", "<SysRoot>", , , vbTextCompare)
    sPath = Replace$(sPath, "c:\Program Files (x86)", "<PF32>", , , vbTextCompare)
    sPath = Replace$(sPath, "C:\Program Files", "<PF64>", , , vbTextCompare)
    sPath = Replace$(sPath, "C:\Users\user\AppData\Local", "<LocalAppData>", , , vbTextCompare)
    sPath = Replace$(sPath, "C:\ProgramData", "<AllUsersProfile>", , , vbTextCompare)
    sPath = Replace$(sPath, "", "", vbTextCompare)
    sPath = Replace$(sPath, "", "", vbTextCompare)
    sPath = Replace$(sPath, "", "", vbTextCompare)
    
    Unexpand = sPath
End Function

Public Function EnvironW(ByVal SrcEnv As String, Optional UseRedir As Boolean) As String
    Dim lr As Long
    Dim buf As String
    
    Init
    
    If Len(SrcEnv) = 0 Then Exit Function
    If InStr(SrcEnv, "%") = 0 Then
        EnvironW = SrcEnv
    Else
        'redirector correction
        If OSver.IsWow64 Then
            If Not UseRedir Then
                If InStr(1, SrcEnv, "%PROGRAMFILES%", 1) <> 0 Then
                    SrcEnv = Replace$(SrcEnv, "%PROGRAMFILES%", PF_64, 1, 1, 1)
                End If
                If InStr(1, SrcEnv, "%COMMONPROGRAMFILES%", 1) <> 0 Then
                    SrcEnv = Replace$(SrcEnv, "%COMMONPROGRAMFILES%", PF_64_Common, 1, 1, 1)
                End If
            End If
        End If
        buf = String$(MAX_PATH, vbNullChar)
        lr = ExpandEnvironmentStrings(StrPtr(SrcEnv), StrPtr(buf), MAX_PATH + 1)
        
        If lr > MAX_PATH Then
            buf = String$(lr, vbNullChar)
            lr = ExpandEnvironmentStrings(StrPtr(SrcEnv), StrPtr(buf), lr + 1)
        End If
        
        If lr Then
            EnvironW = Left$(buf, lr - 1)
        Else
            EnvironW = SrcEnv
        End If
        
        If InStr(EnvironW, "%") <> 0 Then
            If OSver.MajorMinor <= 6 Then
                If InStr(1, EnvironW, "%ProgramW6432%", 1) <> 0 Then
                    EnvironW = Replace$(EnvironW, "%ProgramW6432%", SysDisk & "\Program Files", 1, -1, 1)
                End If
            End If
        End If
    End If

End Function
