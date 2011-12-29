Attribute VB_Name = "modProcMan"
Option Explicit
Public Declare Function CreateToolhelpSnapshot Lib "Kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "Kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "Kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Module32First Lib "Kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Public Declare Function Module32Next Lib "Kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Thread32First Lib "Kernel32" (ByVal hSnapshot As Long, uThread As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "Kernel32" (ByVal hSnapshot As Long, ByRef ThreadEntry As THREADENTRY32) As Long
Public Declare Function TerminateProcess Lib "Kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Private Declare Function SuspendThread Lib "Kernel32" (ByVal hThread As Long) As Long
Private Declare Function ResumeThread Lib "Kernel32" (ByVal hThread As Long) As Long
Private Declare Function OpenThread Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "Kernel32" () As Long

Public Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function lstrcpy Lib "Kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal Length As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type

Private Type THREADENTRY32
    dwSize As Long
    dwRefCount As Long
    th32ThreadID As Long
    th32ProcessID As Long
    dwBasePriority As Long
    dwCurrentPriority As Long
    dwFlags As Long
End Type

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPTHREAD = &H4
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Private Const THREAD_SUSPEND_RESUME = &H2

Private Const LB_SETTABSTOPS = &H192

Public Sub RefreshProcessList(objList As ListBox)
    Dim hSnap&, uPE32 As PROCESSENTRY32, i&
    Dim sExeFile$, hProcess&

    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
    
    uPE32.dwSize = Len(uPE32)
    If ProcessFirst(hSnap, uPE32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    objList.Clear
    Do
        sExeFile = TrimNull(uPE32.szExeFile)
        objList.AddItem uPE32.th32ProcessID & vbTab & sExeFile
    Loop Until ProcessNext(hSnap, uPE32) = 0
    CloseHandle hSnap
End Sub

Public Sub RefreshProcessListNT(objList As ListBox)
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i%
    On Error Resume Next

    If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
        'no PSAPI.DLL file or wrong version
        Exit Sub
    End If
    
    objList.Clear
    lNumProcesses = lNeeded / 4
    For i = 1 To lNumProcesses
        'hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(i))
        If hProc <> 0 Then
            'Openprocess can return 0 but we ignore this since
            'system processes are somehow protected, further
            'processes CAN be opened.... silly windows
        
            lNeeded = 0
            sProcessName = String(260, 0)
            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                sProcessName = TrimNull(sProcessName)
                If sProcessName <> vbNullString Then
                    If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                    If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                    If InStr(1, sProcessName, "%Systemroot%", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "%Systemroot%", sWinDir, , , vbTextCompare)
                    If InStr(1, sProcessName, "Systemroot", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)
                    
                    objList.AddItem lProcesses(i) & vbTab & sProcessName
                End If
            End If
        End If
        CloseHandle hProc
    Next i
End Sub

Public Sub KillProcess(lPID&)
    Dim hProcess&
    If lPID = 0 Then Exit Sub
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProcess = 0 Then
        MsgBox "The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows.", vbCritical
    Else
        If TerminateProcess(hProcess, 0) = 0 Then
            MsgBox "The selected process could not be killed." & _
                   " It may be protected by Windows.", vbCritical
        Else
            CloseHandle hProcess
            DoEvents
        End If
    End If
End Sub

Public Sub KillProcessNT(lPID&)
    Dim hProc&
    On Error Resume Next
    If lPID = 0 Then Exit Sub
    hProc = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProc <> 0 Then
        If TerminateProcess(hProc, 0) = 0 Then
            MsgBox "The selected process could not be killed." & _
                   " It may be protected by Windows.", vbCritical
        Else
            CloseHandle hProc
            DoEvents
        End If
    Else
        MsgBox "The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows." & vbCrLf & vbCrLf & _
               "This process might be a service, which you can " & _
               "stop from the Services applet in Admin Tools." & vbCrLf & _
               "(To load this window, click Start, Run and enter 'services.msc')", vbCritical
    End If
End Sub

Public Sub RefreshDLLListNT(lPID&, objList As ListBox)
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i%
    Dim sModuleName$, j&
    On Error Resume Next
    objList.Clear
    
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lPID)
    If hProc <> 0 Then
        lNeeded = 0
        If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
            For j = 2 To 1024
                If lModules(j) = 0 Then Exit For
                sModuleName = String(260, 0)
                GetModuleFileNameExA hProc, lModules(j), sModuleName, Len(sModuleName)
                sModuleName = TrimNull(sModuleName)
                If sModuleName <> vbNullString And _
                   sModuleName <> "?" Then
                    objList.AddItem sModuleName
                End If
            Next j
        End If
        CloseHandle hProc
    End If
End Sub

Public Sub RefreshDLLList(lPID&, objList As ListBox)
    Dim hSnap&, uME32 As MODULEENTRY32
    Dim sDllFile$
    objList.Clear
    If lPID = 0 Then Exit Sub
    
    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPMODULE, lPID)
    uME32.dwSize = Len(uME32)
    If Module32First(hSnap, uME32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    Do
        sDllFile = TrimNull(uME32.szExePath)
        objList.AddItem sDllFile
    Loop Until Module32Next(hSnap, uME32) = 0
    CloseHandle hSnap
End Sub

Public Sub PauseProcess(lPID&, Optional bPauseOrResume As Boolean = True)
    Dim hSnap&, uTE32 As THREADENTRY32, hThread&
    If Not bIsWinNT And Not bIsWinME Then Exit Sub
    If lPID = GetCurrentProcessId Then Exit Sub
    
    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPTHREAD, lPID)
    If hSnap = -1 Then Exit Sub
    
    uTE32.dwSize = Len(uTE32)
    If Thread32First(hSnap, uTE32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    Do
        If uTE32.th32ProcessID = lPID Then
            hThread = OpenThread(THREAD_SUSPEND_RESUME, False, uTE32.th32ThreadID)
            If bPauseOrResume Then
                SuspendThread hThread
            Else
                ResumeThread hThread
            End If
            CloseHandle hThread
        End If
    Loop Until Thread32Next(hSnap, uTE32) = 0
    CloseHandle hSnap
End Sub

Public Sub SaveProcessList(objProcess As ListBox, objDLL As ListBox, Optional bDoDLLs As Boolean = False)
    Dim sFilename$, i%, sProcess$, sModule$
    sFilename = CmnDlgSaveFile("Save process list to file..", "Text files (*.txt)|*.txt|All files (*.*)|*.*", "processlist.txt")
    If sFilename = vbNullString Then Exit Sub
    
    On Error Resume Next
    Open sFilename For Output As #1
        Print #1, "Process list saved on " & Format(Time, "Long Time") & ", on " & Format(Date, "Short Date")
        Print #1, "Platform: " & GetWindowsVersion & vbCrLf
        Print #1, "[pid]" & vbTab & "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]"
        For i = 0 To objProcess.ListCount - 1
            sProcess = objProcess.List(i)
            Print #1, sProcess & vbTab & vbTab & _
                      GetFilePropVersion(Mid(sProcess, InStr(sProcess, vbTab) + 1)) & vbTab & _
                      GetFilePropCompany(Mid(sProcess, InStr(sProcess, vbTab) + 1))
        Next i
    
        If bDoDLLs Then
            sProcess = objProcess.List(objProcess.ListIndex)
            sProcess = Mid(sProcess, InStr(sProcess, vbTab) + 1)
            Print #1, vbCrLf & vbCrLf & "DLLs loaded by process " & sProcess & ":" & vbCrLf
            Print #1, "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]"
            For i = 0 To objDLL.ListCount - 1
                sModule = objDLL.List(i)
                Print #1, sModule & vbTab & vbTab & GetFilePropVersion(sModule) & vbTab & GetFilePropCompany(sModule)
            Next i
        End If
    
    Close #1
    
    ShellExecute 0, "open", sFilename, vbNullString, vbNullString, 1
End Sub

Public Sub KillProcessByFile(sPath$)
    Dim hSnap&, uPE32 As PROCESSENTRY32, i&
    Dim sExeFile$, hProcess&
    'Note: this sub is silent - it displays no errors !
    If sPath = vbNullString Then Exit Sub
    If bIsWinNT Then
        KillProcessNTByFile sPath
        Exit Sub
    End If
    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
    
    uPE32.dwSize = Len(uPE32)
    If ProcessFirst(hSnap, uPE32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    Do
        sExeFile = TrimNull(uPE32.szExeFile)
        If InStr(1, sExeFile, sPath, vbTextCompare) > 0 Then
            CloseHandle hSnap
    
            'found the process!
            PauseProcess uPE32.th32ProcessID
            hProcess = OpenProcess(PROCESS_TERMINATE, 0, uPE32.th32ProcessID)
            If hProcess <> 0 Then
                If TerminateProcess(hProcess, 0) <> 0 Then
                    CloseHandle hProcess
                    DoEvents
                End If
            End If
            Exit Do
        End If
    Loop Until ProcessNext(hSnap, uPE32) = 0
    CloseHandle hSnap
End Sub

Public Sub KillProcessNTByFile(sPath$)
    'Note: this sub is silent - it displays no errors!
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i%
    On Error Resume Next
    If sPath = vbNullString Then Exit Sub
    If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
        'no PSAPI.DLL file or wrong version
        Exit Sub
    End If

    lNumProcesses = lNeeded / 4
    For i = 1 To lNumProcesses
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
        If hProc <> 0 Then
            'Openprocess can return 0 but we ignore this since
            'system processes are somehow protected, further
            'processes CAN be opened.... silly windows

            lNeeded = 0
            sProcessName = String(260, 0)
            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                sProcessName = TrimNull(sProcessName)
                If sProcessName <> vbNullString Then
                    If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                    If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                    If InStr(1, sProcessName, "%Systemroot%", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "%Systemroot%", sWinDir, , , vbTextCompare)
                    If InStr(1, sProcessName, "Systemroot", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)

                    If InStr(1, sProcessName, sPath, vbTextCompare) > 0 Then

                        'found the process!
                        PauseProcess lProcesses(i)
                        If TerminateProcess(hProc, 0) <> 0 Then
                            CloseHandle hProc
                            DoEvents
                        End If

                        Exit Sub
                    End If
                End If
            End If
        End If
        CloseHandle hProc
    Next i
End Sub

Public Sub CopyProcessList(objProcess As ListBox, objDLL As ListBox, Optional bDoDLLs As Boolean = False)
    Dim i%, sList$, sProcess$, sModule$
    
    On Error Resume Next
    sList = "Process list saved on " & Format(Time, "Long Time") & ", on " & Format(Date, "Short Date") & vbCrLf & _
            "Platform: " & GetWindowsVersion & vbCrLf & vbCrLf & _
            "[pid]" & vbTab & "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]" & vbCrLf
    For i = 0 To objProcess.ListCount - 1
        sProcess = objProcess.List(i)
        sList = sList & sProcess & vbTab & vbTab & _
                GetFilePropVersion(Mid(sProcess, InStr(sProcess, vbTab) + 1)) & vbTab & _
                GetFilePropCompany(Mid(sProcess, InStr(sProcess, vbTab) + 1)) & vbCrLf
    Next i
    
    If bDoDLLs Then
        sProcess = objProcess.List(objProcess.ListIndex)
        sProcess = Mid(sProcess, InStr(sProcess, vbTab) + 1)
        sList = sList & vbCrLf & vbCrLf & "DLLs loaded by process " & sProcess & ":" & vbCrLf & vbCrLf & _
                "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]" & vbCrLf
        For i = 0 To objDLL.ListCount - 1
            sModule = objDLL.List(i)
            sList = sList & sModule & vbTab & vbTab & GetFilePropVersion(sModule) & vbTab & GetFilePropCompany(sModule) & vbCrLf
        Next i
    End If
    
    Clipboard.Clear
    Clipboard.SetText sList
    If bDoDLLs Then
        MsgBox "The process list and dll list have been copied to your clipboard.", vbInformation
    Else
        MsgBox "The process list has been copied to your clipboard.", vbInformation
    End If
End Sub

Public Function GetFilePropVersion$(sFilename$)
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, uVFFI As VS_FIXEDFILEINFO, sVersion$
    If Not FileExists(sFilename) Then Exit Function
    
    lDataLen = GetFileVersionInfoSize(sFilename, ByVal 0)
    If lDataLen = 0 Then Exit Function
        
    ReDim uBuf(0 To lDataLen - 1)
    GetFileVersionInfo sFilename, 0, lDataLen, uBuf(0)
    VerQueryValue uBuf(0), "\", hData, lDataLen
    CopyMemory uVFFI, ByVal hData, Len(uVFFI)
    
    With uVFFI
        sVersion = .dwFileVersionMSh & "." & _
                   .dwFileVersionMSl & "." & _
                   .dwFileVersionLSh & "." & _
                   .dwFileVersionLSl
    End With
    GetFilePropVersion = sVersion
    DoEvents
End Function

Public Function GetFilePropCompany$(sFilename$)
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$
    If Not FileExists(sFilename) Then Exit Function
    
    lDataLen = GetFileVersionInfoSize(sFilename, ByVal 0)
    If lDataLen = 0 Then Exit Function
        
    ReDim uBuf(0 To lDataLen - 1)
    GetFileVersionInfo sFilename, 0, lDataLen, uBuf(0)
    VerQueryValue uBuf(0), "\VarFileInfo\Translation", hData, lDataLen
    If lDataLen = 0 Then Exit Function
    
    CopyMemory uCodePage(0), ByVal hData, 4
    sCodePage = Format(Hex(uCodePage(1)), "00") & _
                Format(Hex(uCodePage(0)), "00") & _
                Format(Hex(uCodePage(3)), "00") & _
                Format(Hex(uCodePage(2)), "00")
        
    'get CompanyName string
    If VerQueryValue(uBuf(0), "\StringFileInfo\" & sCodePage & "\CompanyName", hData, lDataLen) = 0 Then Exit Function
    sCompanyName = String(lDataLen, 0)
    lstrcpy sCompanyName, hData
    GetFilePropCompany = TrimNull(sCompanyName)
    DoEvents
End Function

Public Sub SetListBoxColumns(objListBox As ListBox)
    Dim lTabStop&(1)
    On Error GoTo 0:
    lTabStop(0) = 70
    lTabStop(1) = 0
    SendMessage objListBox.hwnd, LB_SETTABSTOPS, UBound(lTabStop), lTabStop(0)
End Sub
