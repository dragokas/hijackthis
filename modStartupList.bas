Attribute VB_Name = "modStartupList"
Option Explicit

Private Const MAX_PATH      As Long = 260&
Private Const MAX_PATH_W    As Long = 32767&

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst32 Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext32 Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetProductInfo Lib "kernel32.dll" (ByVal dwOSMajorVersion As Long, ByVal dwOSMinorVersion As Long, ByVal dwSpMajorVersion As Long, ByVal dwSpMinorVersion As Long, pdwReturnedProductType As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As Any, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function CheckTokenMembership Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal SidToCheck As Long, IsMember As Long) As Long
Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSid As Long)

Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, ByVal lpFilePart As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal length As Long)

Private Type SID_IDENTIFIER_AUTHORITY
    value(0 To 5) As Byte
End Type

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

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    lpszFileName(MAX_PATH) As Integer
    lpszAlternate(14) As Integer
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(255) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type PROCESSENTRY32
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

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_QUERY_LIMITED_INFORMATION = &H1000
Private Const PROCESS_VM_READ = 16

Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const TH32CS_SNAPPROCESS As Long = 2&

Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&
Private Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF
Private Const INVALID_FILE_ATTRIBUTES   As Long = -1&

Private sReport$
Private bVerbose    As Boolean, sVerbose$
Private bComplete   As Boolean
Private bForceWin9x As Boolean
Private bForceWinNT As Boolean
Private bForceAll   As Boolean
Private bFull       As Boolean
Private bHTML       As Boolean
Private sSLVersion$

Private bNoWriteAccess      As Boolean

Public bStartupListFull     As Boolean
Public bStartupListComplete As Boolean

Public ffDebug  As Integer
Public LogDebug As String


Sub Main()
    Dim sMsg$, sPath$, lTicks&, ff%
    'msgboxW "debug: Test"
    'Test sub
    'EnumNTScripts
    'End
    
    On Error GoTo ErrorHandler:
    sPath = AppPath() & IIf(Right$(AppPath(), 1) = "\", vbNullString, "\")
        
    ' CHANGE THIS OR DIE
    '=======================
    sSLVersion = "1.52.3" 'Changed by Dragokas (Now I will not die, haha :)
    '=======================
    ' FORGET AND YOU GET NO DONUTS!!!!!!!!
    
    On Error Resume Next
    ff = FreeFile()
    Open sPath & "test.tmp" For Output As #ff
        Print #ff, "."
    Close #ff
    Kill sPath & "test.tmp"
    If err.Number Then
        'may not have write access
        MsgBoxW "For some reason, write access was denied to " & _
        "StartupList. The log will not be written to disk, " & _
        "but copied to the clipboard instead." & vbCrLf & vbCrLf & _
        "To see it, open Notepad and hit Ctrl-V when the " & _
        "'finished' message pops up.", vbExclamation
        bNoWriteAccess = True
    End If
    On Error GoTo ErrorHandler:
        
    If Len(Command$) > 0 Then
        If InStr(Command$, "/verbose") > 0 Then bVerbose = True
        If InStr(Command$, "/complete") > 0 Then bComplete = True
        If InStr(Command$, "/force9x") > 0 Then bForceWin9x = True
        If InStr(Command$, "/forcent") > 0 Then bForceWinNT = True
        If InStr(Command$, "/forceall") > 0 Then bForceAll = True
        If InStr(Command$, "/html") > 0 Then bHTML = True
        If InStr(Command$, "/full") > 0 Then bFull = True
        If bForceAll Then
            bForceWin9x = False
            bForceWinNT = False
        End If
    End If
    
    ''If InStr(1, AppEXEName(), "HiJackThis", vbTextCompare) > 0 Then
    If Len(AppExeName()) > 0 Then
        'running from built-in SL version in HiJackThis
        'get parameters
        bFull = bStartupListFull
        bComplete = bStartupListComplete
    End If
    
    'msgboxW "== debug ==", vbExclamation
    'bVerbose = True
    'bComplete = True
    'bForceAll = True
    'bHTML = True
    'bFull = True
    
    sMsg = "This will create a list of all startup entries "
    sMsg = sMsg & "in the Registry and various Windows files" & vbCrLf
    sMsg = sMsg & "and display them in Notepad. "
    sMsg = sMsg & "The process may take up to a few seconds on "
    sMsg = sMsg & "slow systems." & vbCrLf
    sMsg = sMsg & "It will in no way alter anything on your system." & vbCrLf & vbCrLf
    sMsg = sMsg & "Do you want to continue?"
    If MsgBoxW(sMsg, vbQuestion + vbYesNo, "StartupList") = vbNo Then Exit Sub
    
    lTicks = GetTickCount()
    
    'header
    sReport = "StartupList report, " & CStr(Date) & ", " & CStr(Time) & vbCrLf
    sReport = sReport & "StartupList version: " & sSLVersion & vbCrLf
    sReport = sReport & "Started from : " & AppPath(True) & vbCrLf
    sReport = sReport & "Detected: " & sWinVersion & vbCrLf
    sReport = sReport & "Detected: " & BROWSERS.IE.Version & vbCrLf
    
    sReport = sReport & IIf(Command$ = vbNullString, "* Using default options" & vbCrLf, vbNullString)
    sReport = sReport & IIf(bVerbose, "* Using verbose mode" & vbCrLf, vbNullString)
    sReport = sReport & IIf(bComplete, "* Including empty and uninteresting sections" & vbCrLf, vbNullString)
    sReport = sReport & IIf(bForceWin9x, "* Forcing include of Win9x-only sections" & vbCrLf, vbNullString)
    sReport = sReport & IIf(bForceWinNT, "* Forcing include of WinNT-only sections" & vbCrLf, vbNullString)
    sReport = sReport & IIf(bForceAll, "* Forcing include of all possible sections" & vbCrLf, vbNullString)
    sReport = sReport & IIf(bFull, "* Showing rarely important sections" & vbCrLf, vbNullString)
    sReport = sReport & String$(50, "=") & vbCrLf & vbCrLf
    
    '====================
    '=== checks below ===
    '====================
    
    ListRunningProcesses    '1
    CheckAutoStartFolders   '2
    CheckWinNTUserInit      '3
    EnumKeys HKEY_LOCAL_MACHINE, "Run", False, 4
    EnumKeys HKEY_LOCAL_MACHINE, "RunOnce", False, 5
    EnumKeys HKEY_LOCAL_MACHINE, "RunOnceEx", False, 6
    EnumKeys HKEY_LOCAL_MACHINE, "RunServices", False, 7
    EnumKeys HKEY_LOCAL_MACHINE, "RunServicesOnce", False, 8
    EnumKeys HKEY_CURRENT_USER, "Run", False, 9
    EnumKeys HKEY_CURRENT_USER, "RunOnce", False, 10
    EnumKeys HKEY_CURRENT_USER, "RunOnceEx", False, 11
    EnumKeys HKEY_CURRENT_USER, "RunServices", False, 12
    EnumKeys HKEY_CURRENT_USER, "RunServicesOnce", False, 13
    EnumKeys HKEY_LOCAL_MACHINE, "Run", True, 14
    EnumKeys HKEY_CURRENT_USER, "Run", True, 15
    EnumExKeys HKEY_LOCAL_MACHINE, "Run", False, 16
    EnumExKeys HKEY_LOCAL_MACHINE, "RunOnce", False, 17
    EnumExKeys HKEY_LOCAL_MACHINE, "RunOnceEx", False, 18
    EnumExKeys HKEY_LOCAL_MACHINE, "RunServices", False, 19
    EnumExKeys HKEY_LOCAL_MACHINE, "RunServicesOnce", False, 20
    EnumExKeys HKEY_CURRENT_USER, "Run", False, 21
    EnumExKeys HKEY_CURRENT_USER, "RunOnce", False, 22
    EnumExKeys HKEY_CURRENT_USER, "RunOnceEx", False, 23
    EnumExKeys HKEY_CURRENT_USER, "RunServices", False, 24
    EnumExKeys HKEY_CURRENT_USER, "RunServicesOnce", False, 25
    EnumExKeys HKEY_LOCAL_MACHINE, "Run", True, 26
    EnumExKeys HKEY_CURRENT_USER, "Run", True, 27
    CheckClasses ".exe", 28
    CheckClasses ".com", 29
    CheckClasses ".bat", 30
    CheckClasses ".pif", 31
    CheckClasses ".scr", 32
    CheckClasses ".hta", 33
    CheckClasses ".txt", 34
    EnumStubPaths               '35
    EnumICQAgentProgs           '36
    CheckWinINI                 '37
    CheckSystemINI              '38
    CheckExplorer               '39
    EnumWininit "wininit.ini"   '40
    EnumWininit "wininit.bak"   '41
    EnumBAT "c:\autoexec.bat"   '42
    EnumBAT "c:\config.sys"     '43
    EnumBAT sWinDir & "\winstart.bat"   '44
    EnumBAT sWinDir & "\dosstart.bat"   '45
    CheckSuperHiddenExt         '46
    CheckRegedit                '47
    EnumBHOs                    '48
    EnumJOBs                    '49
    EnumDPF                     '50
    EnumLSP                     '51
    EnumServices                '52
    EnumNTScripts               '53
    EnumSSODelayLoad            '54
    EnumKeys HKEY_CURRENT_USER, "policies\Explorer\Run", False, 55
    EnumKeys HKEY_LOCAL_MACHINE, "policies\Explorer\Run", False, 56
    
    lTicks = GetTickCount() - lTicks
    
    'footer
    sReport = sReport & "End of report, xXxXx bytes" & vbCrLf
    sReport = sReport & "Report generated in " & Format(lTicks / 1000, "0.000") & " seconds" & vbCrLf & vbCrLf
    
    sReport = sReport & "Command line options:" & vbCrLf
    sReport = sReport & "   /verbose  - to add additional info on each section" & vbCrLf
    sReport = sReport & "   /complete - to include empty sections and unsuspicious data" & vbCrLf
    sReport = sReport & "   /full     - to include several rarely-important sections" & vbCrLf
    sReport = sReport & "   /force9x  - to include Win9x-only startups even if running on WinNT" & vbCrLf
    sReport = sReport & "   /forcent  - to include WinNT-only startups even if running on Win9x" & vbCrLf
    sReport = sReport & "   /forceall - to include all Win9x and WinNT startups, regardless of platform" & vbCrLf
    
    sReport = Replace$(sReport, "xXxXx", Format(Len(sReport), "###,###,###"))
    sReport = Left$(sReport, Len(sReport) - 2)
    
    'add/remove HTML tags etc
    HTMLize
    
    If bHTML Then
        If Not bNoWriteAccess Then
            ff = FreeFile()
            Open sPath & "startuplist.html" For Output As #ff
                Print #ff, sReport
            Close #ff
            ShellExecute 0, "open", "startuplist.html", vbNullString, AppPath(), 1
        Else
            Clipboard.Clear
            Clipboard.SetText sReport
            MsgBoxW "StartupList has finished generating your logfile!" & vbCrLf & _
                   vbCrLf & "To see it, open Notepad and hit Ctrl-V (paste).", vbInformation
        End If
    Else
        If Not bNoWriteAccess Then
            ff = FreeFile()
            Open sPath & "startuplist.txt" For Output As #ff
                Print #ff, sReport
            Close #ff
            ShellExecute 0, "open", "notepad.exe", sPath & "startuplist.txt", vbNullString, 1
        Else
            Clipboard.Clear
            Clipboard.SetText sReport
            MsgBoxW "StartupList has finished generating your logfile!" & vbCrLf & _
                   vbCrLf & "To see it, open Notepad and hit Ctrl-V (paste).", vbInformation
        End If
    End If
    Exit Sub
    
ErrorHandler:
    Close #ff
    ErrorMsg err, "Main", Command$
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumKeys(lRootKey&, sAutorunKey$, bNT As Boolean, iSection&)
    Dim lEnumBufSize As Long
    
    lEnumBufSize = 32767&

    If bNT Then
        'sub applies to NT only
        If Not bIsWinNT Then
            'this is not NT,override?
            If Not (bForceAll Or bForceWinNT) Then Exit Sub
        End If
    End If
    
    Dim sResult$, bInteresting As Boolean
    Dim hKey&, sValue$, lType&, i&, J&, sData$
    On Error GoTo ErrorHandler:
    
    sResult = "[tag" & CStr(iSection) & "]Autorun entries from Registry:[/tag" & CStr(iSection) & "]" & vbCrLf
    sResult = sResult & IIf(lRootKey = HKEY_LOCAL_MACHINE, "HKLM", "HKCU")
    
    'not using this since WinNT just uses \Windows\ instead
    'of previously assumed \Windows NT\
    If Not bNT Then
        sResult = sResult & "\Software\Microsoft\Windows\CurrentVersion\"
    Else
        sResult = sResult & "\Software\Microsoft\Windows NT\CurrentVersion\"
    End If
    sResult = sResult & sAutorunKey & vbCrLf & vbCrLf
    
    If Not bNT Then
        If RegOpenKeyEx(lRootKey, "Software\Microsoft\Windows\CurrentVersion\" & sAutorunKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
            sResult = sResult & "*Registry key not found*" & vbCrLf
            GoTo EndOfSub
        End If
    Else
        If RegOpenKeyEx(lRootKey, "Software\Microsoft\Windows NT\CurrentVersion\" & sAutorunKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
            sResult = sResult & "*Registry key not found*" & vbCrLf
            GoTo EndOfSub
        End If
    End If
    i = 0
    sData = String$(lEnumBufSize, 0)
    sValue = String$(lEnumBufSize, 0)
    If RegEnumValue(hKey, i, sValue, Len(sValue), 0, lType, ByVal sData, Len(sData)) <> 0 Then
        sResult = sResult & "*No values found*" & vbCrLf
        GoTo EndOfSub
    End If
    Do
        If lType = REG_SZ Then
            bInteresting = True
            sData = TrimNull(sData)
            sValue = Left$(sValue, InStr(sValue, vbNullChar) - 1)
            If sValue = vbNullString Then sValue = "(Default)"
            sResult = sResult & sValue & " = " & sData & vbCrLf
        End If
        sData = String$(lEnumBufSize, 0)
        sValue = String$(lEnumBufSize, 0)
        i = i + 1
    Loop Until RegEnumValue(hKey, i, sValue, Len(sValue), 0, lType, ByVal sData, Len(sData)) <> 0

EndOfSub:
    RegCloseKey hKey
    If bVerbose Then
        sVerbose = "This lists programs that run Registry keys marked by Windows as" & vbCrLf
        sVerbose = sVerbose & "'Autostart key'. To the left are values that are used to clarify what" & vbCrLf
        sVerbose = sVerbose & "program they belong to, to the right the program file that is started." & vbCrLf
        If Right$(sAutorunKey, 4) = "Once" Then
            sVerbose = sVerbose & "The values in the 'RunOnce', 'RunOnceEx' and 'RunServicesOnce' keys" & vbCrLf
            sVerbose = sVerbose & "are run once and then deleted by Windows." & vbCrLf
        End If
        sResult = sResult & vbCrLf & sVerbose
    End If
    sResult = sResult & vbCrLf
    sResult = sResult & String$(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    RegCloseKey hKey
    Dim sRoot$
    Select Case lRootKey
        Case HKEY_CURRENT_USER: sRoot = "HKLCU"
        Case HKEY_LOCAL_MACHINE: sRoot = "HKLM"
    End Select
    ErrorMsg err, "EnumKeys", sRoot, sAutorunKey, bNT
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckClasses(sSubKey$, iSection&)
    'sub applies to all windows versions
    
    Dim sResult$
    sSubKey = UCase$(sSubKey)
    sResult = sResult & "[tag" & iSection & "]File association entry for " & sSubKey & ":[/tag" & iSection & "]" & vbCrLf
    
    Dim hKey&, i&, sData$, bInteresting As Boolean
    On Error GoTo ErrorHandler:
    
    sData = RegGetString(HKEY_CLASSES_ROOT, sSubKey, vbNullString)
    If IsRegVal404(sData) Then
        sResult = sResult & sData & vbCrLf
        GoTo EndOfSub:
    End If
    If sData = vbNullString Then Exit Sub
        
    sData = sData & "\shell\open\command"
    sResult = sResult & "HKEY_CLASSES_ROOT\" & sData & vbCrLf & vbCrLf
    
    sData = RegGetString(HKEY_CLASSES_ROOT, sData, vbNullString)
    If IsRegVal404(sData) Then
        sResult = sResult & sData & vbCrLf
        GoTo EndOfSub
    End If
    
    Select Case sSubKey
        Case ".EXE", ".COM", ".BAT", ".PIF"
            If LCase$(sData) <> """%1"" %*" Then bInteresting = True
        Case ".SCR":
            If LCase$(sData) <> """%1"" /s" And _
               LCase$(sData) <> """%1"" /s ""%3""" Then bInteresting = True
        Case ".HTA":
            If LCase$(sData) <> LCase$(sWinDir) & "\system" & IIf(bIsWin9x, vbNullString, "32") & "\mshta.exe ""%1"" %*" Then bInteresting = True
        Case ".TXT":
            If LCase$(sData) <> sWinDir & "\notepad.exe %1" And _
               LCase$(sData) <> "%systemroot%\system32\notepad.exe %1" Then bInteresting = True
        Case Else
            'MsgBoxW "jackass coder  - no donuts"
    End Select
    
    sResult = sResult & "(Default) = " & sData & vbCrLf
    
EndOfSub:
    RegCloseKey hKey
    If bVerbose Then
        sVerbose = "This Registry value determines how Windows runs files (in this case" & vbCrLf
        sVerbose = sVerbose & sSubKey & " files). If this file is executable, it should read ""%1"" %*." & vbCrLf
        sVerbose = sVerbose & "(""%1"" /S for screensavers, .SCR files.) If it needs to be opened" & vbCrLf
        sVerbose = sVerbose & "with some other program, it should read program.exe ""%1"" %*." & vbCrLf
        sVerbose = sVerbose & "File types that are executable are .EXE, .COM, .PIF, .BAT, .SCR." & vbCrLf
        sVerbose = sVerbose & "File types that are not executable are types like .DOC, .LNK, .BMP," & vbCrLf
        sVerbose = sVerbose & ".JPEG, .SHS, .VBS, .HTA etc."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If
    If bInteresting Or bComplete Then
        sReport = sReport & sResult & vbCrLf & String$(50, "-") & vbCrLf & vbCrLf
    End If
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "CheckClasses", sSubKey
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckWinINI()
    'sub applies to all versions

    Dim sResult$, bInteresting As Boolean
    On Error GoTo ErrorHandler:
    
    sResult = sResult & "[tag36]Load/Run keys from " & sWinDir & "\WIN.INI:[/tag36]"
    sResult = sResult & vbCrLf & vbCrLf
    
    Dim sRet$
    sRet = IniGetString$("win.ini", "windows", "load", False)
    'If Not IsRegVal404(sRet) Then
    If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
        bInteresting = True
    End If
    sResult = sResult & "load=" & sRet & vbCrLf
    
    sRet = IniGetString$("win.ini", "windows", "run", False)
    'If Not IsRegVal404(sRet) Then
    If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
        bInteresting = True
    End If
    sResult = sResult & "run=" & sRet & vbCrLf
    
    'nt only: inifile mapping of win.ini
    If bIsWinNT Or bForceWinNT Or bForceAll Then
        sResult = sResult & vbCrLf & "Load/Run keys from Registry:" & vbCrLf & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "load")
        'sRet = INIGetString$("win.ini", "windows", "load", True)
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\WinLogon: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "run")
        'sRet = INIGetString$("win.ini", "windows", "run", True)
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\WinLogon: run=" & sRet & vbCrLf
        
        'the below 6 probably don't work, but it's just in case
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinLogon", "load")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows\CurrentVersion\WinLogon: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinLogon", "run")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows\CurrentVersion\WinLogon: run=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "load")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows NT\CurrentVersion\WinLogon: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "run")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows NT\CurrentVersion\WinLogon: run=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\WinLogon", "load")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows\CurrentVersion\WinLogon: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\WinLogon", "run")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows\CurrentVersion\WinLogon: run=" & sRet & vbCrLf
        
        'this is a new one
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "load")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows NT\CurrentVersion\Windows: load=" & sRet & vbCrLf
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "run")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows NT\CurrentVersion\Windows: run=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "load")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\Windows: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "run")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\Windows: run=" & sRet & vbCrLf
        
        'this shouldn't really belong here, but anyway..
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs")
        sRet = Replace$(sRet, vbNullChar, "|")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\Windows: AppInit_DLLs=" & sRet & vbCrLf
        
    End If
    
    If bVerbose Then
        sVerbose = "These two entries in WIN.INI are leftover from Windows 3.x, which" & vbCrLf
        sVerbose = sVerbose & "used them as values denoting programs that should be started up" & vbCrLf
        sVerbose = sVerbose & "with Windows. Since Windows 95 and higher uses the Registry to" & vbCrLf
        sVerbose = sVerbose & "store locations of autostart folders, these two entries in WIN.INI" & vbCrLf
        sVerbose = sVerbose & "are redundant, and are rarely used."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If
    sResult = sResult & vbCrLf
    sResult = sResult & String$(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "CheckWinINI"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckSystemINI()
    'sub applies to all versions
    
    Dim sResult$, sRet$, bInteresting As Boolean
    sResult = sResult & "[tag37]Shell & screensaver key from " & sWinDir & "\SYSTEM.INI:[/tag37]"
    sResult = sResult & vbCrLf & vbCrLf
    
    On Error GoTo ErrorHandler:
    sRet = IniGetString$("system.ini", "boot", "shell", False)
    If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString = 0 Then
        bInteresting = True
    End If
    sResult = sResult & "Shell=" & sRet & vbCrLf
    
    sRet = IniGetString$("system.ini", "boot", "SCRNSAVE.EXE", False)
    If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString = 0 Then
        bInteresting = True
    End If
    sResult = sResult & "SCRNSAVE.EXE=" & sRet & vbCrLf
    
    sRet = IniGetString$("system.ini", "boot", "drivers", False)
    If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString = 0 Then
        bInteresting = True
    End If
    sResult = sResult & "drivers=" & sRet & vbCrLf
    
    If bIsWinNT Or bForceWinNT Or bForceAll Then
        sResult = sResult & vbCrLf & "Shell & screensaver key from Registry:" & vbCrLf & vbCrLf
        
        'screw this, I know where it is anyway
        'sRet = INIGetString$("system.ini", "boot", "shell", True)
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "shell")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "Shell=" & sRet & vbCrLf
        
        'sRet = INIGetString$("system.ini", "boot", "SCRNSAVE.EXE", True)
        sRet = RegGetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "SCRNSAVE.EXE=" & sRet & vbCrLf
        
        'doesn't appear in IniFileMapping ?
        sRet = IniGetString$("system.ini", "boot", "drivers", True)
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "drivers=" & sRet & vbCrLf
        
        'got an extra one from policies key!
        sResult = sResult & vbCrLf & "Policies Shell key:" & vbCrLf & vbCrLf
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "Shell")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Policies: Shell=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "Shell")
        If InStr(sRet, "not found*") = 0 And Trim$(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Policies: Shell=" & sRet & vbCrLf
        
    End If
    
    If bVerbose Then
        sVerbose = "The Shell key from SYSTEM.INI tells Windows what file handles" & vbCrLf
        sVerbose = sVerbose & "the Windows shell, i.e. creates the taskbar, desktop icons etc. If" & vbCrLf
        sVerbose = sVerbose & "programs are added to this line, they are all ran at startup." & vbCrLf
        sVerbose = sVerbose & "The SCRNSAVE.EXE line tells Windows what is the default screensaver" & vbCrLf
        sVerbose = sVerbose & "file. This is also a leftover from Windows 3.x and should not be used." & vbCrLf
        sVerbose = sVerbose & "(Since Windows 95 and higher stores this setting in the Registry.)" & vbCrLf
        sVerbose = sVerbose & "The 'drivers' line loads non-standard DLLs or programs." & vbCrLf
        sResult = sResult & vbCrLf & sVerbose
    End If
    sResult = sResult & vbCrLf
    sResult = sResult & String$(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "CheckSystemINI"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumBAT(sFile$)
    'sub applies to 9x only
    If Not bIsWin9x Then
        'this is not 9x, override?
        If Not (bForceAll Or bForceWin9x) Then Exit Sub
    End If
    'display config.sys and dosstart.bat
    'only when using /full

    Dim sResult$, bInteresting As Boolean, ff%
    On Error GoTo ErrorHandler:
    If sFile = "c:\autoexec.bat" Then
        sResult = sResult & "[tag41]" & UCase$(sFile) & " listing:[/tag41]" & vbCrLf & vbCrLf
    ElseIf sFile = "c:\config.sys" And bFull Then
        sResult = sResult & "[tag42]" & UCase$(sFile) & " listing:[/tag42]" & vbCrLf & vbCrLf
    ElseIf sFile = sWinDir & "\winstart.bat" Then
        sResult = sResult & "[tag43]" & UCase$(sFile) & " listing:[/tag43]" & vbCrLf & vbCrLf
    ElseIf sFile = sWinDir & "\dosstart.bat" And bFull Then
        sResult = sResult & "[tag44]" & UCase$(sFile) & " listing:[/tag44]" & vbCrLf & vbCrLf
    Else
        Exit Sub
    End If
    
    Dim sLine$
    If Dir$(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) = vbNullString Then
        sResult = sResult & "*File not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    If FileLenW(sFile) = 0 Then
        sResult = sResult & "*File is empty*" & vbCrLf
        GoTo EndOfSub
    End If
    
    On Error Resume Next '<- NT generates stupid error here
    ff = FreeFile()
    Open sFile For Input As #ff
        Do
            Line Input #ff, sLine
            If Trim$(sLine) <> vbNullString Then
                If Left$(sLine, 1) = "@" Then sLine = Mid$(sLine, 2)
                If InStr(sLine, vbTab) > 0 Then sLine = Replace$(sLine, vbTab, " ")
                If UCase$(Trim$(sLine)) <> "REM" And _
                   (InStr(1, sLine, "REM ", vbTextCompare) <> 1 And _
                   (InStr(1, sLine, "ECHO ", vbTextCompare) <> 1) Or _
                    InStr(sLine, ">") > 0) Or _
                   bComplete Then
                    bInteresting = True
                    sResult = sResult & sLine & vbCrLf
                End If
            End If
        Loop Until EOF(ff)
    Close #ff
    On Error GoTo 0:
    
EndOfSub:
    If bVerbose Then
        If InStr(LCase$(sFile), "c:\autoexec.bat") > 0 Then
            sVerbose = "Autoexec.bat is the very first file to autostart when the computer" & vbCrLf
            sVerbose = sVerbose & "starts, it is a leftover from DOS and older Windows versions." & vbCrLf
            sVerbose = sVerbose & "Windows NT, Windows ME, Windows 2000 and Windows XP don't use this" & vbCrLf
            sVerbose = sVerbose & "file. It is generally used by virusscanners to scan files before" & vbCrLf
            sVerbose = sVerbose & "Windows starts."
            sResult = sResult & vbCrLf & sVerbose & vbCrLf
        ElseIf InStr(LCase$(sFile), "c:\config.sys") > 0 Then
            sVerbose = "Config.sys loads device drivers for DOS, and is rarely used in" & vbCrLf & _
                       "Windows versions newer than Windows 95. Originally it loaded" & vbCrLf & _
                       "drivers for legacy sound cards and such."
            sResult = sResult & vbCrLf & sVerbose & vbCrLf
        ElseIf InStr(LCase$(sFile), "winstart.bat") > 0 Then
            sVerbose = "Winstart.bat loads just before the Windows shell, and is used for" & vbCrLf
            sVerbose = sVerbose & "starting things like soundcard drivers, mouse drivers. Rarely used."
            sResult = sResult & vbCrLf & sVerbose & vbCrLf
        ElseIf InStr(LCase$(sFile), "dosstart.bat") > 0 Then
            sVerbose = "Dosstart.bat loads if you select 'MS-DOS Prompt' from the Startup" & vbCrLf
            sVerbose = sVerbose & "menu when the computer is starting, or if you select 'Restart in" & vbCrLf
            sVerbose = sVerbose & "MS-DOS Mode' from the Shutdown menu in Windows. Mostly used for" & vbCrLf
            sVerbose = sVerbose & "DOS-only drivers, like sound or mouse drivers."
            sResult = sResult & vbCrLf & sVerbose & vbCrLf
        End If
    End If
    sResult = sResult & vbCrLf
    sResult = sResult & String$(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    Close #ff
    ErrorMsg err, "EnumBAT", sFile
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumWininit(sIniFile$)
    'sub applies to 9x only
    If Not bIsWin9x Then
        'this is not 9x, override?
        If Not (bForceAll Or bForceWin9x) Then Exit Sub
    End If
    
    Dim sResult$, bInteresting As Boolean, ff%
    Dim lFileHandle&, uLastWritten As FILETIME
    Dim uLocalTime As FILETIME, uSystemTime As SYSTEMTIME
    On Error GoTo ErrorHandler:
    
    If sIniFile = "wininit.ini" Then
        sResult = sResult & "[tag39]" & sWinDir & "\" & UCase$(sIniFile) & " listing:[/tag39]" & vbCrLf
    ElseIf sIniFile = "wininit.bak" Then
        sResult = sResult & "[tag40]" & sWinDir & "\" & UCase$(sIniFile) & " listing:[/tag40]" & vbCrLf
    End If
    
    lFileHandle = CreateFile(sWinDir & "\" & sIniFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_EXISTING, 0, 0)
    If lFileHandle <> -1 Then
        GetFileTime lFileHandle, ByVal 0, ByVal 0, uLastWritten
        CloseHandle lFileHandle
        FileTimeToLocalFileTime uLastWritten, uLocalTime
        FileTimeToSystemTime uLocalTime, uSystemTime
        With uSystemTime
            sResult = sResult & "(Created " & _
                .wDay & "/" & .wMonth & "/" & .wYear & _
                ", " & .wHour & ":" & .wMinute & ":" & _
                .wSecond & ")" & vbCrLf & vbCrLf
        End With
    Else
        sResult = sResult & vbCrLf
    End If
    
    
    Dim sLine$
    If Dir$(sWinDir & "\" & sIniFile, vbArchive + vbHidden + vbReadOnly + vbSystem) = vbNullString Then
        sResult = sResult & "*File not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    bInteresting = True
    On Error Resume Next '<- NT makes some stupid error here
    ff = FreeFile()
    Open sWinDir & "\" & sIniFile For Input As #ff
        Do
            Line Input #ff, sLine
            If Trim$(sLine) <> vbNullString Then sResult = sResult & sLine & vbCrLf
        Loop Until EOF(ff)
    Close #ff
    On Error GoTo 0:
    
EndOfSub:
    If bVerbose Then
        sVerbose = "WININIT.INI is a settings file for WININIT.EXE, which updates files" & vbCrLf
        sVerbose = sVerbose & "at startup that are normally in use when Windows is running. It is" & vbCrLf
        sVerbose = sVerbose & "mostly used when installing programs or patches that need the" & vbCrLf
        sVerbose = sVerbose & "computer to be restarted to complete the install. After such a reboot," & vbCrLf
        sVerbose = sVerbose & "WININIT.INI is renamed to WININIT.BAK."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If
    sResult = sResult & vbCrLf & String$(50, "-") & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    Close #ff
    ErrorMsg err, "EnumWininit", sIniFile
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckAutoStartFolders()
    'sub applies to all windows versions
    
    Dim sResult$, SS$
    'Dim sDummy$, hKey&, uData() As Byte, i%, sData$
    On Error GoTo ErrorHandler:
    
    'checking all *8* possible folders now - 1.52+
    sResult = sResult & ListFiles(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup", "Shell folders Startup")
    SS = ListFiles(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AltStartup", "Shell folders AltStartup")
    sResult = sResult & SS
    sResult = sResult & ListFiles(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Startup", "User shell folders Startup")
    sResult = sResult & ListFiles(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AltStartup", "User shell folders AltStartup")
    sResult = sResult & ListFiles(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common Startup", "Shell folders Common Startup")
    sResult = sResult & ListFiles(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common AltStartup", "Shell folders Common AltStartup")
    sResult = sResult & ListFiles(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common Startup", "User shell folders Common Startup")
    sResult = sResult & ListFiles(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common AltStartup", "User shell folders Alternate Common Startup")
    
    '===============================================
    If bVerbose Then
        sVerbose = "This lists all programs or shortcuts in folders "
        sVerbose = sVerbose & "marked by Windows as" & vbCrLf & "'Autostart folder', "
        sVerbose = sVerbose & "which means any files within these folders "
        sVerbose = sVerbose & "are" & vbCrLf & "launched when Windows is started. "
        sVerbose = sVerbose & "The Windows standard is that only" & vbCrLf
        sVerbose = sVerbose & "shortcuts (*.lnk, *.pif) should be present "
        sVerbose = sVerbose & "in these folders." & vbCrLf
        sVerbose = sVerbose & "The location of these folders is set in the Registry." & vbCrLf & vbCrLf
        sResult = sResult & sVerbose
    End If
    sResult = sResult & String$(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If sResult <> sVerbose & String$(50, "-") & vbCrLf & vbCrLf Then
        sReport = sReport & "[tag2]Listing of startup folders:[/tag2]" & vbCrLf & vbCrLf
        sReport = sReport & sResult
    End If
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "CheckAutoStartFolders"
    If inIDE Then Stop: Resume Next
End Sub

Private Function ListFiles$(lRootKey&, sSubKey$, sValue$, sName$)
    'sub applies to all windows versions
    
    'ListFiles(HKEY_CURRENT_USER,
    '          "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
    '          "Startup",
    '          "Shell folders Startup")
    
    Dim sResult$, bInteresting As Boolean, sLastFile$
    Dim hKey&, sData$, i&, sFile$ ', uData() As Byte
    On Error GoTo ErrorHandler:
    
    sResult = sName & ":" & vbCrLf
    
    sData = RegGetString(lRootKey, sSubKey, sValue)
    If IsRegVal404(sData) Or Dir$(sData & "\NUL", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "NUL" Then
        sResult = sResult & "*Folder not found*" & vbCrLf & vbCrLf
        GoTo EndOfFunction
    End If
    
    If UCase$(Dir$(sData & "\NUL")) = "NUL" Then
        sResult = sResult & "[" & sData & "]" & vbCrLf
        sFile = Dir$(sData & "\*.*", vbArchive + vbHidden + vbReadOnly + vbSystem)
        If sFile = vbNullString Then
            sResult = sResult & "*No files*" & vbCrLf & vbCrLf
            GoTo EndOfFunction
        End If
        Do
            If LCase$(sFile) <> "desktop.ini" Then
                sResult = sResult & sFile & " -> " & GetFileFromShortcut(sData & "\" & sFile) & vbCrLf
                bInteresting = True
            End If
            sFile = Dir
        Loop Until sFile = vbNullString
        If Not bInteresting Then sResult = sResult & "*No files*" & vbCrLf
        sResult = sResult & vbCrLf
    Else
        sResult = sResult & "*Folder not found*" & vbCrLf & vbCrLf
        GoTo EndOfFunction:
    End If
    
EndOfFunction:
    If bInteresting Or bComplete Then ListFiles = sResult
    Exit Function
    
ErrorHandler:
    Dim sRoot
    Select Case lRootKey
        Case HKEY_CURRENT_USER: sRoot = "HKCU"
        Case HKEY_LOCAL_MACHINE: sRoot = "HKLM"
        Case Else: sRoot = "HK.."
    End Select
    ErrorMsg err, "ListFiles", sRoot, sSubKey, sValue, sName
    If inIDE Then Stop: Resume Next
End Function

Private Sub CheckNeverShowExt(sSubKey$, Optional bOverlay As Boolean = False)
    'sub applies to all windows versions
    
    Dim hKey&, i&, sData$
    On Error GoTo ErrorHandler:
    
    sData = RegGetString(HKEY_CLASSES_ROOT, sSubKey, vbNullString)
    If IsRegVal404(sData) Then
        sReport = sReport & sSubKey & ": " & sData & vbCrLf
        Exit Sub
    End If
    
    'open real key i.e. piffile
    If RegValueExists(HKEY_CLASSES_ROOT, sData, "NeverShowExt") Then
        sReport = sReport & sSubKey & ": HIDDEN!"
    Else
        sReport = sReport & sSubKey & ": not hidden"
    End If
    
    If bOverlay Then
        If RegValueExists(HKEY_CLASSES_ROOT, sData, "IsShortCut") Then
            sReport = sReport & " (arrow overlay: yes)"
        Else
            sReport = sReport & " (arrow overlay: NO!)"
        End If
    End If
    sReport = sReport & vbCrLf
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "CheckNeverShowExt", sSubKey, bOverlay
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckExplorer()
    'sub applies to all windows versions
    '(9x always, NT if improperly configured)
    'only display when using /full switch
    If Not bFull Then Exit Sub
    
    sReport = sReport & "[tag38]Checking for EXPLORER.EXE instances:[/tag38]" & vbCrLf & vbCrLf
    
    On Error Resume Next
    sReport = sReport & sWinDir & "\Explorer.exe: " & IIf(Dir$(sWinDir & "\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, "PRESENT!", "not present") & vbCrLf & vbCrLf
    sReport = sReport & "C:\Explorer.exe: " & IIf(Dir$("c:\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\Explorer\Explorer.exe: " & IIf(Dir$(sWinDir & "\explorer\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\System\Explorer.exe: " & IIf(Dir$(sWinDir & "\system\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\System32\Explorer.exe: " & IIf(Dir$(sWinDir & "\system\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\Command\Explorer.exe: " & IIf(Dir$(sWinDir & "\command\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\Fonts\Explorer.exe: " & IIf(Dir$(sWinDir & "\command\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, "PRESENT!", "not present") & vbCrLf
    On Error GoTo ErrorHandler:
    
    If bVerbose Then
        sVerbose = "Due to a bug in Windows 9x, it mistakenly uses C:\Explorer.exe and" & vbCrLf
        sVerbose = sVerbose & "other instances (if present) when searching for Explorer.exe." & vbCrLf
        sVerbose = sVerbose & "Explorer.exe should only exists in the Windows folder." & vbCrLf
        sVerbose = sVerbose & "Windows NT is vulnerable to this as well, but only if the " & vbCrLf
        sVerbose = sVerbose & "'Shell' Registry value from the previous section " & vbCrLf & _
                              "is just 'Explorer.exe' instead of the full path." & vbCrLf
        sVerbose = sVerbose & "Additionally, presence of \WINDOWS\Explorer\Explorer.exe indicates" & vbCrLf
        sVerbose = sVerbose & "infection with the W32@Trojan.Dlder virus."
        sReport = sReport & vbCrLf & sVerbose & vbCrLf
    End If
    sReport = sReport & vbCrLf & String$(50, "-") & vbCrLf & vbCrLf
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "CheckExplorer"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumStubPaths()
    'sub applies to all windows versions
    'only display when using /full switch
    If Not bFull Then Exit Sub
    
    Dim sResult$, bInteresting As Boolean
    On Error GoTo ErrorHandler:
    sResult = sResult & "[tag34]Enumerating Active Setup stub paths:[/tag34]" & vbCrLf
    sResult = sResult & "HKLM\Software\Microsoft\Active Setup\Installed Components" & vbCrLf
    sResult = sResult & "(* = disabled by HKCU twin)" & vbCrLf & vbCrLf
    Dim hKey&, hSubKey&, i&, J& ', uData() As Byte
    Dim sData$, sVal$, bHasHKCUTwin As Boolean
    
    'Open Active Setup registry key
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Active Setup\Installed Components", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        sResult = sResult & "*Registry key not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    i = 0
    sVal = String$(255, 0)
    'Start enumerating subkeys of 'Installed Components' key
    If RegEnumKey(hKey, i, sVal, 255) <> 0 Then
        sResult = sResult & "*No subkeys found*" & vbCrLf
        RegCloseKey hKey
        GoTo EndOfSub
    End If
    Do
        sVal = Left$(sVal, InStr(sVal, vbNullChar) - 1)
        'Try to open each enumerated subkey
        sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Active Setup\Installed Components\" & sVal, "StubPath")
        If sData <> vbNullString And Not IsRegVal404(sData) Then
            If LCase$(Left$(sData, 10)) <> "rundll.exe" And _
               LCase$(Left$(sData, 8)) <> "rundll32.exe" And _
               LCase$(Left$(sData, 9)) <> "rundll32 " And _
               InStr(sData, "LaunchINFSection") = 0 And _
               InStr(sData, "InstallHinfSection") = 0 Or _
               bComplete Then
                'check for HKCU twin, disabling the key
                If RegKeyExists(HKEY_CURRENT_USER, "Software\Microsoft\Active Setup\Installed Components\" & sVal) Then
                    bHasHKCUTwin = True
                Else
                    bHasHKCUTwin = False
                End If
                
                '...and report key name and stubpath value
                bInteresting = True
                sResult = sResult & "[" & sVal & "]" & IIf(bHasHKCUTwin, " *", vbNullString) & vbCrLf
                sResult = sResult & "StubPath = " & sData & vbCrLf & vbCrLf
            End If
        End If
        
        i = i + 1
        sVal = String$(255, 0)
    Loop Until RegEnumKey(hKey, i, sVal, 255) <> 0
    
EndOfSub:
    If bVerbose Then
        sVerbose = "Programs listed here are components of the Windows Setup that were" & vbCrLf
        sVerbose = sVerbose & "only ran when Windows started for the first time. To prevent them" & vbCrLf
        sVerbose = sVerbose & "from running multiple times, Windows checks for a key with the same" & vbCrLf
        sVerbose = sVerbose & "name at the HKCU root. If it's not found, the component at the HKLM" & vbCrLf
        sVerbose = sVerbose & "root is ran, and a matching key is created at the HKCU root so the" & vbCrLf
        sVerbose = sVerbose & "component is not ran again next time. Most entries involve either" & vbCrLf
        sVerbose = sVerbose & "RUNDLL.EXE or RUNDLL32.EXE, so a suspicious key is not hard to find." & vbCrLf
        sResult = sResult & sVerbose & vbCrLf
    End If
    sResult = sResult & String$(50, "-") & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    RegCloseKey hKey
    ErrorMsg err, "EnumStubPaths"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumICQAgentProgs()
    'sub applies to all windows versions
    
    Dim sResult$, bInteresting As Boolean
    On Error GoTo ErrorHandler:
    
    sResult = sResult & "[tag35]Enumerating ICQ Agent Autostart apps:[/tag35]" & vbCrLf
    sResult = sResult & "HKCU\Software\Mirabilis\ICQ\Agent\Apps" & vbCrLf & vbCrLf
    Dim hKey&, hSubKey&, sVal$, i&, J&, sData$ ', uData() As Byte
    
    'Open ICQ Agent key
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Mirabilis\ICQ\Agent\Apps", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        sResult = sResult & "*Registry key not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    'Start enumerating subkeys
    i = 0
    sVal = String$(255, 0)
    If RegEnumKey(hKey, i, sVal, 255) <> 0 Then
        sResult = sResult & "*No subkeys found*" & vbCrLf
        GoTo EndOfSub
    End If
    Do
        sVal = Left$(sVal, InStr(sVal, vbNullChar) - 1)
        'Try to open each enumerated subkey
        
        sData = RegGetString(HKEY_CURRENT_USER, "Software\Mirabilis\ICQ\Agent\Apps\" & sVal, "Path")
        If sData <> vbNullString Then bInteresting = True
        sResult = sResult & sData & IIf(Right$(sData, 1) = "\", vbNullString, "\")
        sData = RegGetString(HKEY_CURRENT_USER, "Software\Mirabilis\ICQ\Agent\Apps\" & sVal, "Startup")
        sResult = sResult & sData & vbCrLf
        
        i = i + 1
        sVal = String$(255, 0)
    Loop Until RegEnumKey(hKey, i, sVal, 255) <> 0
        
EndOfSub:
    RegCloseKey hKey
    If bVerbose Then
        sVerbose = "The chat program ICQ includes an ICQ Agent that can be configured to" & vbCrLf
        sVerbose = sVerbose & "launch one or multiple browsers when an Internet connection is" & vbCrLf
        sVerbose = sVerbose & "detected. To configure it, open the ICQ Preferences menu and check" & vbCrLf
        sVerbose = sVerbose & "under 'Connection' for a button labelled 'Edit Launch List'." & vbCrLf
        sResult = sResult & vbCrLf & sVerbose
    End If
    sResult = sResult & vbCrLf & String$(50, "-") & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    RegCloseKey hKey
    ErrorMsg err, "EnumICQAgentProgs"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumExKeys(lHive&, ByVal sKey$, bWinNT As Boolean, iNumber&)
    'sub applies to all windows versions
    
    Dim lEnumBufSize As Long
    
    lEnumBufSize = 32767&
    
    Dim sResult$, bInteresting As Boolean
    sKey = "Software\Microsoft\" & IIf(bWinNT, "Windows NT", "Windows") & "\CurrentVersion\" & sKey
    sResult = "[tag" & CStr(iNumber) & "]Autorun entries in Registry subkeys of:[/tag" & CStr(iNumber) & "]" & vbCrLf
    sResult = sResult & IIf(lHive = HKEY_LOCAL_MACHINE, "HKLM\", "HKCU\") & sKey & vbCrLf '& vbCrLf
    Dim hKey&, sVal$, i&, J&, k& ', uData() As Byte
    Dim sData$, hSubKey&, lType&
    On Error GoTo ErrorHandler:
    
    'Open RunOnceEx key
    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
    'If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        sResult = sResult & "*Registry key not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    'Start enumerating subkeys
    i = 0
    sVal = String$(255, 0)
    If RegEnumKey(hKey, i, sVal, 255) <> 0 Then
        sResult = sResult & "*No subkeys found*" & vbCrLf
        GoTo EndOfSub
    End If
    Do
        'Open each subkey...
        sVal = Left$(sVal, InStr(sVal, vbNullChar) - 1)
        If RegOpenKeyEx(lHive, sKey & "\" & sVal, 0, KEY_QUERY_VALUE, hSubKey) = 0 Then
            sResult = sResult & vbCrLf & "[" & sVal & "]" & vbCrLf
            'And enumerate values in it
            J = 0
            sVal = String$(lEnumBufSize, 0)
            sData = String$(lEnumBufSize, 0)
            If RegEnumValue(hSubKey, J, sVal, Len(sVal), 0, lType, ByVal sData, Len(sData)) <> 0 Then
                sResult = sResult & "*No values found*" & vbCrLf
            Else
                Do
                    bInteresting = True
                    sVal = Left$(sVal, InStr(sVal, vbNullChar) - 1)
                    sResult = sResult & sVal & " = "
                    sData = TrimNull(sData)
                    sResult = sResult & sData & vbCrLf
                    J = J + 1
                    sVal = String$(lEnumBufSize, 0)
                    sData = String$(lEnumBufSize, 0)
                Loop Until RegEnumValue(hSubKey, J, sVal, Len(sVal), 0, lType, ByVal sData, Len(sData)) <> 0
            End If
            sResult = sResult '& vbCrLf
            RegCloseKey hSubKey
        End If
        i = i + 1
        sVal = String$(255, 0)
    Loop Until RegEnumKey(hKey, i, sVal, 255) <> 0
    RegCloseKey hKey
    
EndOfSub:
    If bVerbose Then
        sVerbose = "This lists a special format of autorun Registry key, from which" & vbCrLf
        sVerbose = sVerbose & "both programs and functions within DLLs can be launched without" & vbCrLf
        sVerbose = sVerbose & "RUNDLL32.EXE. This autorun key is used very rarely."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If
    RegCloseKey hKey
    sResult = sResult & vbCrLf & String$(50, "-") & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    RegCloseKey hKey
    RegCloseKey hSubKey
    ErrorMsg err, "EnumRunOnceEx"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub ListRunningProcesses()
    'sub applies to all windows versions,
    'but uses different methods for WinNT/9x
    
    sReport = sReport & "[tag1]Running processes:[/tag1]" & vbCrLf & vbCrLf
    If (bIsWinNT Or bForceWinNT) And Not bForceAll Then GoTo WinNTMethod:
    
Win9xMethod:
    Dim hSnap&, uProcess As PROCESSENTRY32, sDummy$
    If bForceAll Then sReport = sReport & "[Using Win9x method]" & vbCrLf & vbCrLf
    On Error Resume Next
    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
    On Error GoTo ErrorHandler:
    If hSnap < 1 Then
        sReport = sReport & "*Unable to list processes*" & vbCrLf
        GoTo EndOfSub:
    End If
    
    uProcess.dwSize = Len(uProcess)
    
    If ProcessFirst32(hSnap, uProcess) = 0 Then
        sReport = sReport & "*No running processes found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    Do
        sDummy = Left$(uProcess.szExeFile, InStr(uProcess.szExeFile, vbNullChar) - 1)
        If Not bIsWinNT Then
            sReport = sReport & sDummy & vbCrLf
        Else
            If sDummy <> "[System Process]" And _
               sDummy <> "System" Then
                sReport = sReport & GetLongPath(sDummy) & vbCrLf
            End If
        End If
    Loop Until ProcessNext32(hSnap, uProcess) = 0
    CloseHandle hSnap
    If bForceAll Then
        sReport = sReport & vbCrLf & "[Using WinNT method]" & vbCrLf & vbCrLf
        GoTo WinNTMethod:
    End If
    GoTo EndOfSub:
    
WinNTMethod:
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i&
    On Error Resume Next
    If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
        sReport = sReport & "(PSAPI.DLL was not found or is " & _
                  "the wrong version. "
        If bForceAll Then
            sReport = sReport & ")" & vbCrLf
            GoTo EndOfSub
        End If
        sReport = sReport & "Using Win9x method instead.)" & vbCrLf & vbCrLf
        GoTo Win9xMethod:
    End If
    On Error GoTo ErrorHandler:
    
    lNumProcesses = lNeeded / 4
    For i = 1 To lNumProcesses
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0, lProcesses(i))
        If hProc <> 0 Then
            lNeeded = 0
            sProcessName = String$(260, 0)
            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                sProcessName = TrimNull(sProcessName)
                If Left$(sProcessName, 1) = "\" Then sProcessName = Mid$(sProcessName, 2)
                If Left$(sProcessName, 3) = "??\" Then sProcessName = Mid$(sProcessName, 4)
                If InStr(1, sProcessName, "%SYSTEMROOT%", vbTextCompare) > 0 Then sProcessName = Replace$(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)
                If InStr(1, sProcessName, "SYSTEMROOT", vbTextCompare) > 0 Then sProcessName = Replace$(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)
                
                sReport = sReport & sProcessName & vbCrLf
            End If
            CloseHandle hProc
        End If
    Next i
    
EndOfSub:
    If bVerbose Then
        sVerbose = "This lists all processes running in memory, "
        sVerbose = sVerbose & "which are all active" & vbCrLf & "programs "
        sVerbose = sVerbose & "and some non-exe system components." & vbCrLf
        '  uncomment when I have list of NT essential components
        'sVerbose = sVerbose  & "Essential processes include: "
        'sVerbose = sVerbose & "KERNEL32.DLL, MSGSRV32.EXE, MPREXE.EXE," & vbCrLf
        'sVerbose = sVerbose & "MMTASK.TSK, EXPLORER.EXE (only once), "
        'sVerbose = sVerbose & "DDHELP.EXE, RNAAPP.EXE," & vbCrLf & "TAPISRV.EXE  "
        'sVerbose = sVerbose & "and EM_EXEC.EXE." & vbCrLf
        sReport = sReport & vbCrLf & sVerbose
    End If
    sReport = sReport & vbCrLf & String$(50, "-") & vbCrLf & vbCrLf
    Exit Sub
    
ErrorHandler:
    CloseHandle hSnap
    ErrorMsg err, "ListRunningProcesses"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckWinNTUserInit()
    'sub applies to NT only
    If Not bIsWinNT Then
        'this is not NT,override?
        If Not (bForceAll Or bForceWinNT) Then Exit Sub
    End If
    
    Dim sDummy$, sResult$, bInteresting As Boolean
    Dim hKey&, sData$, lret&, i& ', uData() As Byte
    On Error GoTo ErrorHandler:
    
    'check HKLM\..\Windows NT\CurrenVersion\WinLogon,UserInit
    bInteresting = False
    sDummy = "[HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon]" & vbCrLf
    sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "UserInit")
    If IsRegVal404(sData) Then
        sDummy = sDummy & sData & vbCrLf & vbCrLf
    Else
        bInteresting = True
        sDummy = sDummy & "UserInit = " & sData & vbCrLf & vbCrLf
    End If
    If bInteresting Or bComplete Then sResult = sResult & sDummy
    'check \Windows\CurrentVer as well just in case
    bInteresting = False
    sDummy = "[HKLM\Software\Microsoft\Windows\CurrentVersion\Winlogon]" & vbCrLf
    sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Winlogon", "UserInit")
    If IsRegVal404(sData) Then
        sDummy = sDummy & sData & vbCrLf & vbCrLf
    Else
        bInteresting = True
        sDummy = sDummy & "UserInit = " & sData & vbCrLf & vbCrLf
    End If
    If bInteresting Or bComplete Then sResult = sResult & sDummy
    
    
    
    'check HKCU\..\Windows NT\CurrentVersion\WinLogon,UserInit
    bInteresting = False
    sDummy = "[HKCU\Software\Microsoft\Windows NT\CurrentVersion\Winlogon]" & vbCrLf
    sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "UserInit")
    If IsRegVal404(sData) Then
        sDummy = sDummy & sData & vbCrLf & vbCrLf
    Else
        bInteresting = True
        sDummy = sDummy & "UserInit = " & sDummy & vbCrLf & vbCrLf
    End If
    If bInteresting Or bComplete Then sResult = sResult & sDummy
    'check \Windows\CurrentVer as well, just in case
    bInteresting = False
    sDummy = "[HKCU\Software\Microsoft\Windows\CurrentVersion\Winlogon]" & vbCrLf
    sData = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Winlogon", "UserInit")
    If IsRegVal404(sData) Then
        sDummy = sDummy & sData & vbCrLf & vbCrLf
    Else
        bInteresting = True
        sDummy = sDummy & "UserInit = " & sDummy & vbCrLf & vbCrLf
    End If
    If bInteresting Or bComplete Then sResult = sResult & sDummy
    
    
    If bVerbose Then
        sVerbose = "These are Windows NT/2000/XP specific startup locations. They" & vbCrLf
        sVerbose = sVerbose & "execute when the user logs on to his workstation."
        sResult = sResult & sVerbose & vbCrLf & vbCrLf
    End If
    
EndOfSub:
    If sResult <> vbNullString And sResult <> sResult & sVerbose & vbCrLf & vbCrLf Then
        sReport = sReport & "[tag3]Checking Windows NT UserInit:[/tag3]" & vbCrLf & vbCrLf
        sReport = sReport & sResult & String$(50, "-") & vbCrLf & vbCrLf
    End If
    Exit Sub
    
ErrorHandler:
    RegCloseKey hKey
    ErrorMsg err, "CheckWinNTUserInit"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckSuperHiddenExt()
    'sub applies to all windows versions
    'only display when using /full switch
    If Not bFull Then Exit Sub
    
    On Error GoTo ErrorHandler:
    sReport = sReport & "[tag45]Checking for superhidden extensions:[/tag45]" & vbCrLf & vbCrLf
    CheckNeverShowExt ".lnk", True
    CheckNeverShowExt ".pif", True
    CheckNeverShowExt ".exe"
    CheckNeverShowExt ".com"
    CheckNeverShowExt ".bat"
    CheckNeverShowExt ".hta"
    CheckNeverShowExt ".scr"
    CheckNeverShowExt ".shs"
    CheckNeverShowExt ".shb"
    CheckNeverShowExt ".vbs"
    CheckNeverShowExt ".vbe"
    CheckNeverShowExt ".wsh"
    CheckNeverShowExt ".scf", True
    CheckNeverShowExt ".url", True
    CheckNeverShowExt ".js"
    CheckNeverShowExt ".jse"
    If bVerbose Then
        sVerbose = "Some file extensions are always hidden, like .lnk (shortcut) and" & vbCrLf
        sVerbose = sVerbose & ".pif (shortcut to MS-DOS program). The Life_Stages virus was a .shs" & vbCrLf
        sVerbose = sVerbose & "(Shell Scrap) file that had the extension hidden by default. This can" & vbCrLf
        sVerbose = sVerbose & "be a security risk when a virus with a double-extension filename is" & vbCrLf
        sVerbose = sVerbose & "on the loose, since the extension can be hidden even when 'Don't show" & vbCrLf
        sVerbose = sVerbose & "extensions for known filetypes' is turned off." & vbCrLf
        sVerbose = sVerbose & "The shortcut overlay acts as a reminder that the file " & _
                              "is just a shortcut." & vbCrLf & "If the shortcut overlay is removed, " & _
                              "the difference between a file and" & vbCrLf & "a shortcut is invisible." & vbCrLf
        sReport = sReport & vbCrLf & sVerbose
    End If
    sReport = sReport & vbCrLf & String$(50, "-") & vbCrLf
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "CheckSuperHiddenExt"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckRegedit()
    '35
    Dim sCmd$, sResult$, sRegedit$, bInteresting As Boolean
    'display section only with /full switch
    If Not bFull Then Exit Sub
    
    On Error GoTo ErrorHandler:
    sResult = "[tag46]Verifying REGEDIT.EXE integrity:[/tag46]" & vbCrLf & vbCrLf
    sRegedit = sWinDir & "\Regedit.exe"
    
    'check location of REGEDIT.EXE, should be WinDir
    If Dir$(sRegedit, vbArchive + vbReadOnly + vbSystem) <> vbNullString Then
        sResult = sResult & "- Regedit.exe found in " & sWinDir & vbCrLf
    Else
        sResult = sResult & "- Regedit.exe is MISSING!" & vbCrLf
        bInteresting = True
    End If
    
    'check .reg open command, should be regedit.exe "%1"
    sCmd = RegGetString(HKEY_CLASSES_ROOT, "regfile\shell\open\command", vbNullString)
    sCmd = Replace$(sCmd, """", vbNullString)
    'If LCase$(sCmd) = "regedit.exe ""%1""" Or _
    '   LCase$(sCmd) = LCase$(sRegedit) & """%1""" Or _
    '   LCase$(sCmd) = "%systemroot%\regedit.exe ""%1""" Then
    If Left$(LCase$(sCmd), 11) = "regedit.exe" And _
       InStr(1, sCmd, "%1", vbTextCompare) > 0 Then
        sResult = sResult & "- .reg open command is normal (" & sCmd & ")" & vbCrLf
    Else
        sResult = sResult & "- .reg open command is NOT normal! (" & sCmd & ")" & vbCrLf
        bInteresting = True
    End If
    
    'check regedit.exe file info
    Dim hData&, lDataLen& ', uVSFI As VS_FIXEDFILEINFO
    Dim uBuf() As Byte, uCodePage(0 To 3) As Byte, sCodePage$
    Dim sCompanyName$, sFileDescription$, sOriginalFilename$
    
    'get data length
    lDataLen = GetFileVersionInfoSize(sRegedit, ByVal 0)
    If lDataLen = 0 Then
        bInteresting = True
        sResult = sResult & "- Unable to retrieve file info on regedit.exe!" & vbCrLf
    Else
        ReDim uBuf(0 To lDataLen - 1)
        'get handle to file props
        GetFileVersionInfo sRegedit, 0, lDataLen, uBuf(0)
        
        'get codepage/language Hex string
        VerQueryValue uBuf(0), "\VarFileInfo\Translation", hData, lDataLen
            
            'convert to readable hex string
            CopyMemory uCodePage(0), ByVal hData, 4
            sCodePage = Format(Hex(uCodePage(1)), "00") & _
                        Format(Hex(uCodePage(0)), "00") & _
                        Format(Hex(uCodePage(3)), "00") & _
                        Format(Hex(uCodePage(2)), "00")
            
            'get CompanyName string
            'should be 'Microsoft Corporation'
            If VerQueryValue(uBuf(0), "\StringFileInfo\" & sCodePage & "\CompanyName", hData, lDataLen) = 0 Then
                bInteresting = True
                sResult = sResult & "- Regedit.exe has no CompanyName property! It is either missing or named something else." & vbCrLf
            Else
                sCompanyName = String$(lDataLen, 0)
                lstrcpy sCompanyName, hData
                sCompanyName = TrimNull(sCompanyName)
                If sCompanyName = "Microsoft Corporation" Then
                    sResult = sResult & "- Company name OK: '" & sCompanyName & "'" & vbCrLf
                Else
                    bInteresting = True
                    sResult = sResult & "- Company name NOT OK: '" & sCompanyName & "'" & vbCrLf
                End If
            End If
            
            'get OriginalFilename string
            'should be 'REGEDIT.EXE' (sic, in caps)
            If VerQueryValue(uBuf(0), "\StringFileInfo\" & sCodePage & "\OriginalFilename", hData, lDataLen) = 0 Then
                bInteresting = True
                sResult = sResult & "- Regedit.exe has no OriginalFilename property! It is either missing or named something else." & vbCrLf
            Else
                sOriginalFilename = String$(lDataLen, 0)
                lstrcpy sOriginalFilename, hData
                sOriginalFilename = TrimNull(sOriginalFilename)
                If sOriginalFilename = "REGEDIT.EXE" Then
                    sResult = sResult & "- Original filename OK: '" & sOriginalFilename & "'" & vbCrLf
                Else
                    bInteresting = True
                    sResult = sResult & "- Original filename NOT OK: '" & sOriginalFilename & "'" & vbCrLf
                End If
            End If
        
            'get FileDescription string
            If VerQueryValue(uBuf(0), "\StringFileInfo\" & sCodePage & "\FileDescription", hData, lDataLen) = 0 Then
                bInteresting = True
                sResult = sResult & "- Regedit.exe has no FileDescription property! It is either missing or named something else." & vbCrLf
            Else
                sFileDescription = String$(lDataLen, 0)
                lstrcpy sFileDescription, hData
                sFileDescription = TrimNull(sFileDescription)
                sResult = sResult & "- File description: '" & sFileDescription & "'" & vbCrLf
            End If
        End If
    
    If bInteresting = False Then
        sResult = sResult & vbCrLf & "Registry check passed" & vbCrLf
    Else
        sResult = sResult & vbCrLf & "Registry check failed!" & vbCrLf
    End If
    
    If bVerbose Then
        sVerbose = "Regedit.exe is the Windows Registry Editor. Without it, " & _
                   "you cannot" & vbCrLf & "access the Registry " & _
                  "or merge Registry scripts into the Registry." & vbCrLf & _
                  "Several viruses/trojans mess with this important " & _
                  "system file, e.g." & vbCrLf & "moving it somewhere else or " & _
                  "replacing it with a copy of the trojan." & vbCrLf & _
                  "Above checks will ensure that Regedit.exe is " & _
                  "in the correct place" & vbCrLf & "and that it really is " & _
                  "Regedit." & vbCrLf & _
                  "If you have ScriptSentry installed, the .reg command" & vbCrLf & _
                  "is altered and you fail the check. Don't worry" & vbCrLf & _
                  "about this."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If
    
    If bInteresting Or bComplete Then
        sReport = sReport & vbCrLf & sResult & vbCrLf & String$(50, "-") & vbCrLf
    End If
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "CheckRegedit"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumBHOs()
    '36
    Dim hKey&, sCLSID$, sName$, sFile$, i&, sResult$
    Dim bInteresting As Boolean, bDisabledBHODemon As Boolean
    On Error GoTo ErrorHandler:
    sResult = "[tag47]Enumerating Browser Helper Objects:[/tag47]" & vbCrLf & vbCrLf
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sCLSID = String$(255, 0)
        If RegEnumKey(hKey, i, sCLSID, 255) <> 0 Then
            'no BHO's
            RegCloseKey hKey
            sResult = sResult & "*No BHO's found*" & vbCrLf
            GoTo EndOfSub:
        End If
        bInteresting = True
        Do
            sCLSID = TrimNull(sCLSID)
            
            'get friendly name
            sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID, vbNullString)
            
            'get filename
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", vbNullString)
            
            'check for BHODemon-disabled BHO's
            bDisabledBHODemon = False
            If Len(sFile) > 18 Then
                If Right$(sFile, 18) = "__BHODemonDisabled" Then
                    sFile = Left$(sFile, Len(sFile) - 18)
                    bDisabledBHODemon = True
                End If
            End If
            
            sResult = sResult & IIf(sName <> vbNullString And Not IsRegVal404(sName), sName, "(no name)") & _
                      " - " & IIf(sFile <> vbNullString And Not IsRegVal404(sFile), sFile, "(no file)") & _
                      IIf(bDisabledBHODemon, " (disabled by BHODemon)", vbNullString) & _
                      IIf(FileExists(sFile) = False And (sFile <> vbNullString And Not IsRegVal404(sFile)), " (file missing)", vbNullString) & _
                      " - " & sCLSID & vbCrLf
            
            i = i + 1
            sCLSID = String$(255, 0)
        Loop Until RegEnumKey(hKey, i, sCLSID, 255) <> 0
        RegCloseKey hKey
    Else
        sResult = sResult & "*No BHO's found*" & vbCrLf
    End If
    
EndOfSub:
    If bVerbose Then
        sVerbose = "MSIE features Browser Helper Objects (BHO) that plug " & _
                   "into MSIE and" & vbCrLf & "can do virtually anything on your " & _
                   "system. Benevolant examples are" & vbCrLf & "the Google Toolbar " & _
                   "and the Acrobat Reader plugin. More often though, " & vbCrLf & _
                   "BHO's are installed by spyware and serve you to " & _
                   "a neverending flow" & vbCrLf & "of popups and ads as well as " & _
                   "tracking your browser habits, claiming" & vbCrLf & "they '" & _
                   "enhance your browsing experience'."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If
    
    If bInteresting Or bComplete Then
        sReport = sReport & vbCrLf & sResult & vbCrLf & String$(50, "-") & vbCrLf
    End If
    Exit Sub
    
ErrorHandler:
    RegCloseKey hKey
    ErrorMsg err, "EnumBHOs"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumJOBs()
    '37
    On Error GoTo ErrorHandler:
    Dim sResult$, bInteresting As Boolean, sFile$
    Dim sDummy$, vDummy As Variant, sBla, i&
    sResult = "[tag48]Enumerating Task Scheduler jobs:[/tag48]" & vbCrLf & vbCrLf
    If Dir$(sWinDir & "\Tasks\NUL") = vbNullString Then
        sResult = sResult & "*" & sWinDir & "\Tasks folder not found*" & vbCrLf
        GoTo EndOfSub:
    End If
    
    sFile = Dir$(sWinDir & "\Tasks\*.job", vbArchive + vbHidden + vbReadOnly + vbSystem)
    If sFile = vbNullString Then
        sResult = sResult & "*No jobs found*" & vbCrLf
        GoTo EndOfSub:
    End If
    bInteresting = True
    Do
        sResult = sResult & sFile & vbCrLf
        
        sFile = Dir
    Loop Until sFile = vbNullString
    
    
EndOfSub:
    If bVerbose Then
        sVerbose = "The Windows Task Scheduler can run programs " & _
                   "at a certain time," & vbCrLf & "automatically. Though very " & _
                   "unlikely, this can be exploited by" & vbCrLf & "making a job " & _
                   "that runs a virus or trojan."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If
    
    If bInteresting Or bComplete Then
        sReport = sReport & vbCrLf & sResult & vbCrLf & String$(50, "-") & vbCrLf
    End If
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "EnumJOBs"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumDPF()
    '38
    Dim hKey&, sName$, i&, sResult$, bInteresting As Boolean
    Dim sFriendlyName$, sCodeBase$, sOSD$, sINF$, sFile$
    Const sKeyDPF$ = "Software\Microsoft\Code Store Database\Distribution Units"
    On Error GoTo ErrorHandler:
    sResult = "[tag49]Enumerating Download Program Files:[/tag49]" & vbCrLf & vbCrLf
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyDPF, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        sResult = sResult & "*Registry key not found*" & vbCrLf
        Exit Sub
    End If
    
    sName = String$(255, 0)
    If RegEnumKey(hKey, i, sName, 255) <> 0 Then
        sResult = sResult & "*No subkeys found*" & vbCrLf
        Exit Sub
    End If
    
    Do
        sName = TrimNull(sName)
        If Left$(sName, 1) = "{" And Right$(sName, 1) = "}" Then
            'it's a CLSID, so get real name from HKCR\CLSID
            sFriendlyName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName, vbNullString)
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName & "\InProcServer32", vbNullString)
        End If
        If Not RegKeyExists(HKEY_LOCAL_MACHINE, sKeyDPF & "\" & sName & "\Contains\Java") _
           Or bComplete Then
        
            sCodeBase = RegGetString(HKEY_LOCAL_MACHINE, sKeyDPF & "\" & sName & "\DownloadInformation", "CODEBASE")
            sOSD = RegGetString(HKEY_LOCAL_MACHINE, sKeyDPF & "\" & sName & "\DownloadInformation", "OSD")
            'sINF = RegGetString(HKEY_LOCAL_MACHINE, sKeyDPF & "\" & sName & "\DownloadInformation", "INF")
            
            If (InStr(1, sCodeBase, "http://codecs.microsoft.com/", vbTextCompare) <> 1 And _
               InStr(1, sCodeBase, "http://java.sun.com/", vbTextCompare) <> 1) Or _
               bComplete Then
               
                bInteresting = True
                
                If Not IsRegVal404(sFriendlyName) And sFriendlyName <> vbNullString Then
                    sResult = sResult & "[" & sFriendlyName & "]" & vbCrLf
                Else
                    sResult = sResult & "[" & sName & "]" & vbCrLf
                End If
                sResult = sResult & IIf(sFile <> vbNullString And Not IsRegVal404(sFile), "InProcServer32 = " & sFile & vbCrLf, vbNullString) & _
                                IIf(sCodeBase <> vbNullString And Not IsRegVal404(sCodeBase), "CODEBASE = " & sCodeBase & vbCrLf, vbNullString) & _
                                IIf(sOSD <> vbNullString And Not IsRegVal404(sOSD), "OSD = " & sOSD & vbCrLf, vbNullString)
                                'IIf(sINF <> vbNullString And Not IsRegVal404(sINF), "INF = " & sINF & vbCrLf, vbNullString)
                sResult = sResult & vbCrLf
            End If
        End If
        i = i + 1
        sName = String$(255, 0)
        sFriendlyName = vbNullString
    Loop Until RegEnumKey(hKey, i, sName, 255) <> 0
    RegCloseKey hKey
    
    If bVerbose Then
        sVerbose = "The items in Download Program Files are programs " & _
                   "you downloaded and" & vbCrLf & "automatically installed " & _
                   "themselves in MSIE. Most of these are Java" & vbCrLf & "classes " & _
                   "Media Player codecs and the likes. Some items " & _
                   "are only" & vbCrLf & "visible from the Registry and may not " & _
                   "show up in the folder."
        sResult = sResult & sVerbose & vbCrLf & vbCrLf
    End If
    
    If (bInteresting Or bComplete) And _
       sResult <> "[tag38]Enumerating Download Program Files:[/tag38]" & vbCrLf & vbCrLf Then
        sReport = sReport & vbCrLf & sResult & String$(50, "-") & vbCrLf
    End If
    Exit Sub
    
ErrorHandler:
    RegCloseKey hKey
    ErrorMsg err, "EnumDPF"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumLSP()
    '39
    Dim i&, sKeyBase$, iNumNameSpace&, iNumProtocol&
    Dim sNamespace$, sProtocol$, sFile$, sSafeFiles$
    Dim sResult$, bInteresting As Boolean
    On Error GoTo ErrorHandler:
    
    sResult = "[tag50]Enumerating Winsock LSP files:[/tag50]" & vbCrLf & vbCrLf
    sSafeFiles = "rnr20.dll*" & _
                "mswsock.dll*" & _
                "winrnr.dll*" & _
                "msafd.dll*" & _
                "rsvpsp.dll*" & _
                "mswsosp.dll*" & _
                "nwprovau.dll*" & _
                "nwws2nds.dll*" & _
                "nwws2sap.dll*" & _
                "nwws2slp.dll*" & _
                "cslsp.dll*" & _
                "dcsws2.dll*" & _
                "wspwsp.dll*" & _
                "mclsp.dll*" & _
                "zklspr.dll*" & _
                "wps.dll*" & _
                "mxavlsp.dll*" & _
                "imon.dll*" & _
                "drwhook.dll*" & _
                "wspirda.dll*"
                
    sKeyBase = "System\CurrentControlSet\Services\WinSock2\Parameters"
    sNamespace = RegGetString(HKEY_LOCAL_MACHINE, sKeyBase, "Current_NameSpace_Catalog")
    sProtocol = RegGetString(HKEY_LOCAL_MACHINE, sKeyBase, "Current_Protocol_Catalog")
    iNumNameSpace = RegGetDword(HKEY_LOCAL_MACHINE, sKeyBase & "\" & sNamespace, "Num_Catalog_Entries")
    iNumProtocol = RegGetDword(HKEY_LOCAL_MACHINE, sKeyBase & "\" & sProtocol, "Num_Catalog_Entries")
    
    For i = 1 To iNumNameSpace
        sFile = RegGetString(HKEY_LOCAL_MACHINE, sKeyBase & "\" & sNamespace & "\Catalog_Entries\" & Format(i, "000000000000"), "LibraryPath")
        sFile = Replace$(sFile, "%systemroot%", sWinDir, 1, 1, vbTextCompare)
        If IsRegVal404(sFile) Then
            bInteresting = True
            sResult = sResult & "NameSpace #" & CStr(i) & " is MISSING" & vbCrLf
        ElseIf InStr(1, sSafeFiles, Mid$(sFile, InStrRev(sFile, "\") + 2), vbTextCompare) = 0 Or _
           bComplete Then
            bInteresting = True
            sResult = sResult & "NameSpace #" & CStr(i) & ": " & sFile & IIf(FileExists(sFile) = False, " (file MISSING)", vbNullString) & vbCrLf
        End If
    Next i
    
    For i = 1 To iNumProtocol
        sFile = RegGetFileFromBinary(HKEY_LOCAL_MACHINE, sKeyBase & "\" & sProtocol & "\Catalog_Entries\" & Format(i, "000000000000"), "PackedCatalogItem")
        sFile = Replace$(sFile, "%systemroot%", sWinDir, 1, 1, vbTextCompare)
        If IsRegVal404(sFile) Then
            bInteresting = True
            sResult = sResult & "Protocol #" & CStr(i) & " is MISSING" & vbCrLf
        ElseIf InStr(1, sSafeFiles, Mid$(sFile, InStrRev(sFile, "\") + 2), vbTextCompare) = 0 Or _
           bComplete = True Then
            bInteresting = True
            sResult = sResult & "Protocol #" & CStr(i) & ": " & sFile & IIf(FileExists(sFile) = False, " (file MISSING)", vbNullString) & vbCrLf
        End If
    Next i
    sResult = sResult & vbCrLf
    
    If bVerbose Then
        sVerbose = "The Windows Socket system (Winsock) connects your " & _
                   "system to the" & vbCrLf & "Internet. Part of this task is resolving " & _
                   "domain names (www.server.com)" & vbCrLf & "to IP addresses (12.23.34.45) " & _
                   "which is handler by several system" & vbCrLf & "files, called Layered " & _
                   "Service Providers (LSPs), which work as a" & vbCrLf & "chain: " & _
                   "if one LSP is gone, the chain is broken and Winsock " & _
                   "cannot" & vbCrLf & "resolve domain names - which means no program " & _
                   "on your system can" & vbCrLf & "access the Internet."
        sResult = sResult & sVerbose & vbCrLf & vbCrLf
    End If
    
    If bInteresting Or bComplete Then
        sReport = sReport & vbCrLf & sResult & String$(50, "-") & vbCrLf
    End If
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "EnumLSP"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumServices()
    'sub applies to all windows versions
    'only display when using /full switch
    If Not bFull Then Exit Sub
    
    'NT: HKLM\System\CurrentControlSet\Services\*
    '9x: HKLM\System\CurrentControlSet\Services\*
    'and HKLM\System\CurrentControlSet\Servces\VxD\*
    
    'enum subkeys of
    'HKLM\System\CurrentControlSet\Services
    'which have a value Start
    'Start 0 = hidden (system)
    'Start 1 = hidden (system)
    'Start 2 = automatic        <-- these are what we want
    'Start 3 = manual
    'Start 4 = disabled
    Dim hKey&, hKey2&, i&, sName$, sDummy$, bInteresting As Boolean
    Dim sServiceName$, sServicePath$, sServiceDesc$
    Dim sResult$, uStart() As Byte
    
    If bIsWinNT Or bForceAll Or bForceWinNT Then
        sResult = sResult & "[tag51]Enumerating Windows NT/2000/XP services[/tag51]" & vbCrLf & vbCrLf
        
        If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
            'failed to open key
            sResult = sResult & "*Registry key not found*" & vbCrLf
            GoTo EndOfSub
        End If
        
        sName = String$(255, 0)
        If RegEnumKey(hKey, 0, sName, 255) <> 0 Then
            'no subkeys
            sResult = sResult & "*No services found*" & vbCrLf
            RegCloseKey hKey
            GoTo EndOfSub
        End If
        
        Do
            sName = TrimNull(sName)
            ReDim uStart(0 To 3)
            RegOpenKeyEx HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, 0, KEY_QUERY_VALUE, hKey2
            If RegQueryValueEx(hKey2, "Start", 0, REG_BINARY, uStart(0), 4) = 0 Then
                sServiceName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DisplayName")
                sServicePath = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ImagePath")
                sServiceDesc = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Description")
                    
                'filter out invalid services
                If sServicePath <> vbNullString And Not IsRegVal404(sServicePath) Then
                    
                    'filter out uninteresting services
                    If bComplete Or (uStart(0) <> 0 And _
                                     uStart(0) <> 1 And _
                                     uStart(0) <> 3 And _
                                     uStart(0) <> 4) Then

                        bInteresting = True
                        sDummy = IIf(Not IsRegVal404(sServiceName), sServiceName, sName) & _
                                 ": " & sServicePath
                        
                        'sDummy = "[" & IIf(Not IsRegVal404(sServiceName), sServiceName, sName) & _
                        '         "]" & vbCrLf & sServicePath
                          
                        Select Case uStart(0)
                            Case 0, 1: sDummy = sDummy & " (system)"
                            Case 2: sDummy = sDummy & " (autostart)"
                            Case 3: sDummy = sDummy & " (manual start)"
                            Case 4: sDummy = sDummy & " (disabled)"
                        End Select
                        sDummy = sDummy & vbCrLf
                            
                        sResult = sResult & sDummy
                    End If
                End If
            End If
            RegCloseKey hKey2
            
            i = i + 1
            sName = String$(255, 0)
        Loop Until RegEnumKey(hKey, i, sName, 255) <> 0
        RegCloseKey hKey
        If Not bInteresting Then sResult = sResult & "*No services found*" & vbCrLf
        sResult = sResult & vbCrLf
    End If
    
    If bIsWin9x Or bForceAll Or bForceWin9x Then
        sResult = sResult & "Enumerating Win9x VxD services:" & vbCrLf & vbCrLf
        
        If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
            'failed to open key
            sResult = sResult & "*Registry key not found*" & vbCrLf
            GoTo EndOfSub
        End If
    
        sName = String$(255, 0)
        If RegEnumKey(hKey, 0, sName, 255) <> 0 Then
            'no subkeys
            sResult = sResult & "*No services found*" & vbCrLf
            RegCloseKey hKey
            GoTo EndOfSub
        End If
        
        Do
            sName = TrimNull(sName)
            ReDim uStart(0 To 3)
            RegOpenKeyEx HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD\" & sName, 0, KEY_QUERY_VALUE, hKey2
            If RegQueryValueEx(hKey2, "Start", 0, REG_BINARY, uStart(0), 4) = 0 Then
                sServiceName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD\" & sName, "DisplayName")
                sServicePath = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD\" & sName, "StaticVxD")
                
                If InStr(sServicePath, "*") = 0 Or bComplete Then
                    sResult = sResult & IIf(Not IsRegVal404(sServiceName), sServiceName, sName) & _
                              ": " & _
                              IIf(Not IsRegVal404(sServicePath), sServicePath, "(no file)") & vbCrLf
                End If
            End If
            RegCloseKey hKey2
        
            i = i + 1
            sName = String$(255, 0)
        Loop Until RegEnumKey(hKey, i, sName, 255) <> 0
        RegCloseKey hKey
        
    End If
    
EndOfSub:
    If bVerbose Then
        sVerbose = vbCrLf
        sVerbose = sVerbose & "Windows NT4/2000/XP launches several dozen of 'services' when" & vbCrLf
        sVerbose = sVerbose & "your system starts that range in importance from system-" & vbCrLf
        sVerbose = sVerbose & "critical (like RPCSS) to redundant (Remote Registry Editor)," & vbCrLf
        sVerbose = sVerbose & "or even dangerous (Universal Plug & Play). Though very little" & vbCrLf
        sVerbose = sVerbose & "malicious programs use this type of startup, it is included here" & vbCrLf
        sVerbose = sVerbose & "for completeness." & vbCrLf
        sVerbose = sVerbose & "Windows 9x/ME launches system-critical files in a similar way" & vbCrLf
        sVerbose = sVerbose & "at system startup, but unlike Windows NT services, the Windows 9x" & vbCrLf
        sVerbose = sVerbose & "VxD services are all important, and much less in number. Practically" & vbCrLf
        sVerbose = sVerbose & "the only non-Microsoft programs starting from here are software firewalls." & vbCrLf
        sResult = sResult & sVerbose '& vbCrLf
    End If
    
    sResult = vbCrLf & sResult & vbCrLf & String$(50, "-") & vbCrLf
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    ErrorMsg err, "EnumServices"
    RegCloseKey hKey
    RegCloseKey hKey2
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumNTScripts()
    'sub applies to NT only
    If Not bIsWinNT Then
        'this is not NT,override?
        If Not (bForceAll Or bForceWinNT) Then Exit Sub
    End If
    
    Dim sDummy$, sPath$, vPaths As Variant, i&, ff%
    Dim sResult$, bInteresting As Boolean
    On Error GoTo ErrorHandler:
    sResult = "[tag52]Enumerating Windows NT logon/logoff scripts:[/tag52]" & vbCrLf '& vbCrLf
    
    'get paths, probably
    'WinSysDir\GroupPolicy\[Machine|User]\Scripts
    'use intermediate string to filter dupes
    If Not RegKeyExists(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System\Scripts") Then
        'sResult = sResult & "*No logon/logoff scripts set to run*" & vbCrLf
    Else
        sDummy = RegGetString(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System\Scripts", "Logon")
        If IsRegVal404(sDummy) Then sDummy = vbNullString
        sPath = sDummy & "|"
        sDummy = RegGetString(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System\Scripts", "Logoff")
        If IsRegVal404(sDummy) Then sDummy = vbNullString
        If InStr(sPath, sDummy & "|") = 0 Then sPath = sPath & sDummy & "|"
    End If
    
    If Not RegKeyExists(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\System\Scripts") Then
        'sResult = sResult & "*No startup/shutdown scripts set to run*" & vbCrLf
    Else
        sDummy = RegGetString(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\System\Scripts", "Startup")
        If IsRegVal404(sDummy) Then sDummy = vbNullString
        If InStr(sPath, sDummy & "|") = 0 Then sPath = sPath & sDummy & "|"
        sDummy = RegGetString(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\System\Scripts", "Shutdown")
        If IsRegVal404(sDummy) Then sDummy = vbNullString
        If InStr(sPath, sDummy & "|") = 0 Then sPath = sPath & sDummy & "|"
    End If
    
    If sPath = vbNullString Then
        sResult = sResult & "*No scripts set to run*" & vbCrLf & vbCrLf
        GoTo NextCheck:
    End If
    
    sResult = sResult & vbCrLf
    vPaths = Split(sPath, "|")
    For i = 0 To UBound(vPaths)
        If vPaths(i) = vbNullString Then Exit For
        vPaths(i) = BuildPath(CStr(vPaths(i)), "scripts.ini")
        If FileExists(CStr(vPaths(i))) Then
            sResult = sResult & "[" & vPaths(i) & "]" & vbCrLf
            ff = FreeFile()
            Open vPaths(i) For Input As #ff
                Do
                    Line Input #ff, sDummy
                    If Trim$(sDummy) <> vbNullString And _
                        sDummy <> "€ю" Then
                        bInteresting = True
                        sResult = sResult & sDummy & vbCrLf
                    End If
                Loop Until EOF(ff)
            Close #ff
            sResult = sResult & vbCrLf
        End If
    Next i
    
    '============
    
NextCheck:
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\Session Manager", "BootExecute")
    If sDummy <> vbNullString Or bComplete Then
        sResult = sResult & "Windows NT checkdisk command:" & vbCrLf & "BootExecute = " & sDummy & vbCrLf & vbCrLf
    End If
    
    '============
    
    Dim hKey&, lDataLen&, uData() As Byte, sData$
    sResult = sResult & "Windows NT 'Wininit.ini':" & vbCrLf
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\Session Manager", 0, KEY_QUERY_VALUE, hKey) = 0 Then
        ReDim uData(1000)
        lDataLen = 1000
        If RegQueryValueEx(hKey, "PendingFileRenameOperations", 0, 0, uData(i), lDataLen) = 0 Then
            bInteresting = True
            sData = vbNullString
            For i = 0 To lDataLen
                If uData(i) = 0 Then uData(i) = Asc("|")
                sData = sData & Chr(uData(i))
            Next i
            sData = Replace$(sData, "||||", vbCrLf)
            sData = Replace$(sData, "\??\", vbNullString)
            sData = Replace$(sData, "|!", " => ")
            sData = Replace$(sData, "|" & vbCrLf, vbCrLf)
        Else
            sData = "*Registry value not found*"
        End If
        If bInteresting Or bComplete Then sResult = sResult & "PendingFileRenameOperations: " & sData & vbCrLf & vbCrLf
        RegCloseKey hKey
    End If
    
    If bVerbose Then
        sVerbose = vbNullString
        sVerbose = sVerbose & "Windows NT4/2000/XP can be setup to run scripts at user logon," & vbCrLf
        sVerbose = sVerbose & "logoff, and system startup or shutdown." & vbCrLf
        sVerbose = sVerbose & "These scripts can do virtually anything, from mapping a" & vbCrLf
        sVerbose = sVerbose & "network drive to starting a trojan horse virus. If scripts" & vbCrLf
        sVerbose = sVerbose & "are started on your system and you don't know what" & vbCrLf
        sVerbose = sVerbose & "they are, consider disabling them using the Group Policy" & vbCrLf
        sVerbose = sVerbose & "Editor (click Start, Run, type ""gpedit.msc"" and hit Enter)." & vbCrLf
        sResult = sResult & sVerbose & vbCrLf
    End If

    sResult = vbCrLf & sResult & String$(50, "-") & vbCrLf
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
ErrorHandler:
    RegCloseKey hKey
    ErrorMsg err, "EnumNTScripts"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub EnumSSODelayLoad()
    'sub applies to all windows versions
    
    'method 53
    
    Dim lEnumBufSize As Long
    
    lEnumBufSize = 32767&
    
    Dim hKey&, i&, sName$, sCLSID$, sFile$, sResult$
    Dim bInteresting As Boolean
    On Error GoTo ErrorHandler:
    sResult = "[tag53]Enumerating ShellServiceObjectDelayLoad items:[/tag53]" & vbCrLf & vbCrLf
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad", 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        'key doesn't exist
        sResult = sResult & "*Registry key not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    sName = String$(lEnumBufSize, 0)
    sCLSID = String$(lEnumBufSize, 0)
    If RegEnumValue(hKey, 0, sName, Len(sName), 0, ByVal 0, ByVal sCLSID, Len(sCLSID)) <> 0 Then
        'no values!
        RegCloseKey hKey
        sResult = sResult & "*No items found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    bInteresting = True
    Do
        sName = TrimNull(sName)
        sCLSID = TrimNull(sCLSID)
        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString)
        sFile = Replace$(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
        
        sResult = sResult & sName & ": " & sFile & vbCrLf
        
        sName = String$(lEnumBufSize, 0)
        sCLSID = String$(lEnumBufSize, 0)
        i = i + 1
    Loop Until RegEnumValue(hKey, i, sName, Len(sName), 0, ByVal 0, ByVal sCLSID, Len(sCLSID)) <> 0
    RegCloseKey hKey
    
EndOfSub:
    If bVerbose Then
        sVerbose = "This Registry key lists several system components " & _
        "are loaded at" & vbCrLf & "system startup. Not much is known about this key " & _
        "since it is" & vbCrLf & "virtually undocumented and only used by programs like " & _
        "the Volume" & vbCrLf & "Control, IE Webcheck and Power Management icons. " & _
        "However, a" & vbCrLf & "virus/trojan in the form of a DLL can also load " & _
        "from this key." & vbCrLf & "The Hitcap trojan is an example of this."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If

    sResult = vbCrLf & sResult & vbCrLf & String$(50, "-") & vbCrLf
    If bInteresting Or bComplete Then sReport = sReport & sResult
    
    Exit Sub
ErrorHandler:
    RegCloseKey hKey
    ErrorMsg err, "EnumOSSDelayLoad"
    If inIDE Then Stop: Resume Next
End Sub


Public Function GetWindowsVersion() As String    'Init by Form_load.     Result -> sWinVersion (global)
                                                 '                              -> bIsWin64 (global)
                                                 '                              -> bIsWin32 (global)
                                                 '                              -> bIsWinVistaOrLater (global)
    Const SM_SERVERR2               As Long = 89&
    Const VER_NT_WORKSTATION        As Long = 1&
    Const VER_SUITE_STORAGE_SERVER  As Long = &H2000&
    Const VER_SUITE_DATACENTER      As Long = &H80&
    Const VER_SUITE_PERSONAL        As Long = &H200&
    Const VER_SUITE_ENTERPRISE      As Long = 2&
    Const SM_CLEANBOOT              As Long = 67&

    Dim OVI As OSVERSIONINFOEX, sWinVer$, ProductType&, lret&, dec!
    On Error GoTo ErrorHandler:
    Dim lIsWin64&
    
    Set OSInfo = New clsOSInfo
    
    If IsProcedureAvail("IsWow64Process", "kernel32") Then
        IsWow64Process GetCurrentProcess(), lIsWin64
        bIsWin64 = lIsWin64
        bIsWin32 = Not bIsWin64
        OSVer.bIsWin64 = bIsWin64
        OSVer.Bitness = IIf(bIsWin64, "x64", "x32")
    End If
    
    'enable redirector
    If bIsWin64 Then ToggleWow64FSRedirection True
    
    lret = GetSystemMetrics(SM_CLEANBOOT)
    OSVer.bIsSafeBoot = (lret > 0)
    Select Case lret
        Case 1
            OSVer.BootMode = "Safe mode" '1 - Fail-safe boot
        Case 2
            OSVer.BootMode = "Safe mode with network support" '2 - Fail-safe with network boot
        Case Else
            OSVer.BootMode = "Normal" '0 - Normal boot
    End Select
    
    OSVer.bIsAdmin = CheckIsAdmin()
    
    OVI.dwOSVersionInfoSize = Len(OVI)
    GetVersionEx OVI
    
    OSVer.Major = OVI.dwMajorVersion
    OSVer.Minor = OVI.dwMinorVersion
    
    ' OS Major + Minor
    dec = OVI.dwMinorVersion
    If dec <> 0 Then Do: dec = dec / 10: Loop Until dec < 1
    OSVer.MajorMinor = OVI.dwMajorVersion + dec
    
    ' Service Pack Major + Minor
    dec = OVI.wServicePackMinor
    If dec <> 0 Then Do: dec = dec / 10: Loop Until dec < 1
    OSVer.SPVer = OVI.wServicePackMajor + dec
    
    OSVer.Build = OVI.dwBuildNumber
    
    With OVI
        bIsWinVistaOrLater = .dwMajorVersion >= 6
        OSVer.bIsVistaOrLater = bIsWinVistaOrLater
        
        Select Case .dwPlatformId
            Case 0: GetWindowsVersion = "Detected: Windows 3.x running Win32s": Exit Function
            Case 1: bIsWin9x = True: bIsWinNT = False
            Case 2: bIsWinNT = True: bIsWin9x = False
        End Select
        
        If bIsWin9x Then
            If .dwMajorVersion = 4 Then
                OSVer.Platform = "Win9x"
                Select Case .dwMinorVersion
                '4.0
                Case 0 'Windows 95 [A/B/C]
                    OSVer.OSName = "Windows 95"
                '4.10
                Case 10 'Windows 98 [Gold/SE]
                    OSVer.OSName = "Windows 98"
                    OSVer.Edition = IIf(OSVer.SPVer <> 0, "SE", "Gold")
                '4.90
                Case 90 'Windows Millennium Edition
                    OSVer.OSName = "Windows ME"
                    bIsWinME = True
                End Select
            End If
        ElseIf bIsWinNT Then
            OSVer.Platform = "WinNT"
            
            Select Case .dwMajorVersion
                '4.x
                Case 4 'Windows NT4
                    OSVer.OSName = "Windows NT"
                Case 5
                    Select Case .dwMinorVersion
                        '5.0
                        Case 0 'Windows 2000
                            OSVer.OSName = "Windows 2000"
                            
                            If .wProductType = VER_NT_WORKSTATION Then
                                OSVer.Edition = "Professional"
                            Else
                                If .wSuiteMask And VER_SUITE_DATACENTER Then
                                    OSVer.Edition = "Datacenter Server"
                                ElseIf .wSuiteMask And VER_SUITE_ENTERPRISE Then
                                    OSVer.Edition = "Advanced Server"
                                Else
                                    OSVer.Edition = "Server"
                                End If
                            End If
                        '5.1
                        Case 1 'Windows XP
                            OSVer.OSName = "Windows XP"
                            OSVer.Edition = IIf(.wSuiteMask = VER_SUITE_PERSONAL, "Home Edition", "Professional")

                        '5.2
                        Case 2 'Windows 2003 (or XP x64)
                                If GetSystemMetrics(SM_SERVERR2) Then
                                    OSVer.OSName = "Windows Server 2003 R2"
                                ElseIf .wSuiteMask And VER_SUITE_STORAGE_SERVER Then
                                    OSVer.OSName = "Windows Storage Server 2003"
                                ElseIf .wProductType = VER_NT_WORKSTATION And bIsWin64 Then
                                    OSVer.OSName = "Windows XP"
                                    OSVer.Edition = "Professional"
                                Else
                                    OSVer.OSName = "Windows Server 2003"
                                End If
                    End Select
                Case 6
                    Select Case .dwMinorVersion
                        '6.0
                        Case 0 'Windows Vista (or Server 2008)
                            OSVer.OSName = IIf(.wProductType = VER_NT_WORKSTATION, "Windows Vista", "Windows Server 2008")
                        '6.1
                        Case 1 'Windows 7 (or Server 2008 R2)
                            OSVer.OSName = IIf(.wProductType = VER_NT_WORKSTATION, "Windows 7", "Windows Server 2008 R2")
                        '6.2
                        Case 2 'Windows 8 (or Server 2012)
                            OSVer.OSName = IIf(.wProductType = VER_NT_WORKSTATION, "Windows 8", "Windows Server 2012")
                        '6.3
                        Case 3 'Windows 8.1 (or Server 2012 R2)
                            OSVer.OSName = IIf(.wProductType = VER_NT_WORKSTATION, "Windows 8.1", "Windows Server 2012 R2")
                        '6.4
                        Case 4 'Windows 10 Technical Preview
                            OSVer.OSName = IIf(.wProductType = VER_NT_WORKSTATION, "Windows 10", "Windows 10 Server")
                            OSVer.Edition = "Technical Preview"
                    End Select
                Case 10
                    Select Case .dwMinorVersion
                        '10.0
                        Case 0 'Windows 10
                            OSVer.OSName = IIf(.wProductType = VER_NT_WORKSTATION, "Windows 10", "Windows 10 Server")
                    End Select

            End Select
        End If
        
        If Len(OSVer.OSName) = 0 Then 'Unknown Windows -> looking to registry
            OSVer.OSName = modRegistry.GetRegData(0, "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
            If Len(OSVer.OSName) = 0 Then OSVer.OSName = "Unknown Windows " & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber
        End If
    
        If OSVer.Edition = vbNullString Then
            If OSVer.Major >= 6 Then
                If GetProductInfo(OVI.dwMajorVersion, OVI.dwMinorVersion, OVI.wServicePackMajor, OVI.wServicePackMinor, ProductType) Then
                    OSVer.Edition = GetProductName(ProductType)
                End If
            End If
        End If

        sWinVer = OSVer.OSName & " " & OSVer.Edition & " SP" & OSVer.SPVer & " " & _
            "(Windows " & OSVer.Platform & " " & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & ")"

    End With

    GetWindowsVersion = sWinVer
    Exit Function
    
ErrorHandler:
    ErrorMsg err, "GetWindowsVersion"
    If inIDE Then Stop: Resume Next
End Function

Function CheckIsAdmin() As Boolean
    On Error GoTo ErrorHandler:
    Dim uAuthNt As SID_IDENTIFIER_AUTHORITY
    Dim lpPSid  As Long
    Dim lr      As Long
    
    uAuthNt.value(5) = 5
    If AllocateAndInitializeSid(uAuthNt, 2, &H20&, &H220&, 0&, 0&, 0&, 0&, 0&, 0&, lpPSid) Then
        If CheckTokenMembership(0&, lpPSid, lr) Then CheckIsAdmin = (lr <> 0)
        Call FreeSid(lpPSid)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg err, "CheckIsAdmin"
    If inIDE Then Stop: Resume Next
End Function


Function GetProductName(ProductType As Long) As String
    On Error GoTo ErrorHandler:
    
    Dim ProductName As String
    
    Select Case ProductType
    Case &H6&
        ProductName = "Business"
    Case &H10&
        ProductName = "Business N"
    Case &H12&
        ProductName = "HPC Edition"
    Case &H40&
        ProductName = "Server Hyper Core V"
    Case &H65&
        ProductName = "" '"Windows 8" (Core)
    Case &H62&
        ProductName = "N" ' "Windows 8 N"
    Case &H63&
        ProductName = "China" ' "Windows 8 China"
    Case &H64&
        ProductName = "Single Language" ' "Windows 8 Single Language"
    Case &H50&
        ProductName = "Server Datacenter (EI)"
    Case &H8&
        ProductName = "Server Datacenter (FI)"
    Case &HC&
        ProductName = "Server Datacenter (CI)"
    Case &H27&
        ProductName = "Server Datacenter without Hyper-V (CI)"
    Case &H25&
        ProductName = "Server Datacenter without Hyper-V (FI)"
    Case &H4&
        ProductName = "Enterprise"
    Case &H46&
        ProductName = "Not supported"
    Case &H54&
        ProductName = "Enterprise N (EI)"
    Case &H1B&
        ProductName = "Enterprise N"
    Case &H48&
        ProductName = "Server Enterprise (EI)"
    Case &HA&
        ProductName = "Server Enterprise (FI)"
    Case &HE&
        ProductName = "Server Enterprise (CI)"
    Case &H29&
        ProductName = "Server Enterprise without Hyper-V (CI)"
    Case &HF&
        ProductName = "Server Enterprise for Itanium-based Systems"
    Case &H26&
        ProductName = "Server Enterprise without Hyper-V (FI)"
    Case &H3B&
        ProductName = "Windows Essential Server Solution Management"
    Case &H3C&
        ProductName = "Windows Essential Server Solution Additional"
    Case &H3D&
        ProductName = "Windows Essential Server Solution Management SVC"
    Case &H3E&
        ProductName = "Windows Essential Server Solution Additional SVC"
    Case &H2&
        ProductName = "Home Basic"
    Case &H43&
        ProductName = "Not supported"
    Case &H5&
        ProductName = "Home Basic N"
    Case &H3&
        ProductName = "Home Premium"
    Case &H44&
        ProductName = "Not supported"
    Case &H1A&
        ProductName = "Home Premium N"
    Case &H22&
        ProductName = "Windows Home Server 2011"
    Case &H13&
        ProductName = "Windows Storage Server 2008 R2 Essentials"
    Case &H2A&
        ProductName = "Microsoft Hyper-V Server"
    Case &H1E&
        ProductName = "Windows Essential Business Server Management Server"
    Case &H20&
        ProductName = "Windows Essential Business Server Messaging Server"
    Case &H1F&
        ProductName = "Windows Essential Business Server Security Server"
    Case &H4C&
        ProductName = "Windows MultiPoint Server Standard (FI)"
    Case &H4D&
        ProductName = "Windows MultiPoint Server Premium (FI)"
    Case &H30&
        ProductName = "Professional"
    Case &H45&
        ProductName = "Not supported"
    Case &H31&
        ProductName = "Professional N"
    Case &H67&
        ProductName = "Professional with Media Center"
    Case &H36&
        ProductName = "Server For SB Solutions EM"
    Case &H33&
        ProductName = "Server For SB Solutions"
    Case &H37&
        ProductName = "Server For SB Solutions EM"
    Case &H18&
        ProductName = "Windows Server 2008 for Windows Essential Server Solutions"
    Case &H23&
        ProductName = "Windows Server 2008 without Hyper-V for Windows Essential Server Solutions"
    Case &H21&
        ProductName = "Server Foundation"
    Case &H32&
        ProductName = "Windows Small Business Server 2011 Essentials"
    Case &H9&
        ProductName = "Windows Small Business Server"
    Case &H19&
        ProductName = "Small Business Server Premium"
    Case &H3F&
        ProductName = "Small Business Server Premium (CI)"
    Case &H38&
        ProductName = "Windows MultiPoint Server"
    Case &H4F&
        ProductName = "Server Standard (EI)"
    Case &H7&
        ProductName = "Server Standard"
    Case &HD&
        ProductName = "Server Standard (CI)"
    Case &H24&
        ProductName = "Server Standard without Hyper-V"
    Case &H28&
        ProductName = "Server Standard without Hyper-V (CI)"
    Case &H34&
        ProductName = "Server Solutions Premium"
    Case &H35&
        ProductName = "Server Solutions Premium (CI)"
    Case &HB&
        ProductName = "Starter"
    Case &H42&
        ProductName = "Not supported"
    Case &H2F&
        ProductName = "Starter N"
    Case &H17&
        ProductName = "Storage Server Enterprise"
    Case &H2E&
        ProductName = "Storage Server Enterprise (CI)"
    Case &H14&
        ProductName = "Storage Server Express"
    Case &H2B&
        ProductName = "Storage Server Express (CI)"
    Case &H60&
        ProductName = "Storage Server Standard (EI)"
    Case &H15&
        ProductName = "Storage Server Standard"
    Case &H2C&
        ProductName = "Storage Server Standard (CI)"
    Case &H5F&
        ProductName = "Storage Server Workgroup (EI)"
    Case &H16&
        ProductName = "Storage Server Workgroup"
    Case &H2D&
        ProductName = "Storage Server Workgroup (CI)"
    Case &H0&
        ProductName = "An unknown product"
    Case &H1&
        ProductName = "Ultimate"
    Case &H47&
        ProductName = "Not supported"
    Case &H1C&
        ProductName = "Ultimate N"
    Case &H11&
        ProductName = "Web Server (FI)"
    Case &H1D&
        ProductName = "Web Server (CI)"
    Case Else
        ProductName = "Unknown Edition"
    End Select

    GetProductName = ProductName
    Exit Function
ErrorHandler:
    ErrorMsg err, "GetProductName"
    If inIDE Then Stop: Resume Next
End Function

'»щет полный путь к файлу.
'‘ункци€ не гарантирует, что файл существует.
'–езультатом могут быть даже служебные символы, если таковые были поданы на вход.
'Ёто конечно странно, но пока что не трогаю это поведение функции, чтобы ничего не сломать вверх по стеку вызовов.

Public Function GetLongPath$(ByVal sFile$)
    'sub applies to NT only, checked in ListRunningProcesses()
    'attempt to find location of given file
    On Error GoTo ErrorHandler:
    'On Error Resume Next
    
    Dim pos&
    
    'evading parasites that put html or garbled data in
    'O4 autorun entries :P
    If InStr(sFile, "<") > 0 Or InStr(sFile, ">") > 0 Or _
       InStr(sFile, "|") > 0 Or InStr(sFile, "*") > 0 Or _
       InStr(sFile, "?") > 0 Then  'Or InStr(sFile, "/") > 0 Or InStr(sFile, ":") > 0 Then
        GetLongPath = sFile ' ???? //TODO
        Exit Function
    End If
    
    If InStr(sFile, "/") <> 0 Then sFile = Replace$(sFile, "/", "\")
    
    'Dim ProcPath$
    'ToggleWow64FSRedirection False
    'ProcPath = Space$(MAX_PATH)
    'LSet ProcPath = sFile & vbNullChar
    'If CBool(PathFindOnPath(StrPtr(ProcPath), 0&)) Then
    '    GetLongPath = ProcPath
    '    ToggleWow64FSRedirection True
    '    Exit Function
    'Else
    '    ToggleWow64FSRedirection True
    'End If
    
    If Left$(sFile, 1) = """" Then
        pos = InStr(2, sFile, """")
        If pos <> 0 Then
            sFile = Mid$(sFile, 2, pos - 2)
        Else
            sFile = Mid$(sFile, 2)
        End If
    End If
    
    GetLongPath = FindOnPath(sFile)
    If 0 <> Len(GetLongPath) Then Exit Function
    
    pos = InStrRev(sFile, ".exe", -1, 1)
    If 0 <> pos And pos <> Len(sFile) - 3 Then
        sFile = Left$(sFile, pos + 3)
        GetLongPath = FindOnPath(sFile)
        If 0 <> Len(GetLongPath) Then Exit Function
    End If
    
    'If sFile = "[System Process]" Or sFile = "System" Then
    '    GetLongPath = sFile
    '    Exit Function
    'End If
    
    If InStr(sFile, "\") > 0 Then
        'filename is already full path
        GetLongPath = sFile
        Exit Function
    End If
    
'    'check if file is self
'    If LCase$(sFile) = LCase$(AppExeName(True)) Then
'        GetLongPath = AppPath() & IIf(Right$(AppPath(), 1) = "\", vbNullString, "\") & sFile
'        Exit Function
'    End If
    
    Dim hKey, sData$, i&, sDummy$, sProgramFiles$
    'check App Paths regkey
    sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" & sFile, vbNullString)
    If Not IsRegVal404(sData) And sData <> vbNullString Then
        GetLongPath = sData
        Exit Function
    End If
    
    'check own folder
    If Dir$(BuildPath(AppPath(), sFile), vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
        GetLongPath = BuildPath(AppPath(), sFile)
        Exit Function
    End If
    
    'check windir
    If Dir$(sWinDir & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
        GetLongPath = sWinDir & "\" & sFile
        Exit Function
    End If
    
    'check windir\system
    If Dir$(sWinDir & "\system\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
        GetLongPath = sWinDir & "\system\" & sFile
        Exit Function
    End If
    
    'check windir\system32
    ToggleWow64FSRedirection False
    If Dir$(sWinDir & "\system32\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
        GetLongPath = sWinDir & "\system32\" & sFile
        ToggleWow64FSRedirection True
        Exit Function
    End If
    ToggleWow64FSRedirection True
    
    If OSVer.Bitness = "x64" Then
        If Dir$(sWinDir & "\syswow64\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = sWinDir & "\syswow64\" & sFile
            Exit Function
        End If
    End If
    
    If InStr(sFile, ".") > 0 Then
        'prog.exe -> prog
        sDummy = Left$(sFile, InStr(sFile, ".") - 1)
        sProgramFiles = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
        
        'check x:\program files\prog\prog.exe
        If Dir$(sProgramFiles & "\" & sDummy & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = sProgramFiles & "\" & sDummy & "\" & sFile
            Exit Function
        End If
        
        'check c:\prog\prog.exe
        If Dir$("C:\" & sDummy & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = "C:\" & sDummy & "\" & sFile
            Exit Function
        End If
        
        'check x:\program files\prog32\prog.exe
        If Dir$(sProgramFiles & "\" & sDummy & "32\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = sProgramFiles & "\" & sDummy & "32\" & sFile
            Exit Function
        End If
        If Dir$(sProgramFiles & "\" & sDummy & "16\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = sProgramFiles & "\" & sDummy & "16\" & sFile
            Exit Function
        End If
        
        'check c:\prog32\prog.exe
        If Dir$("C:\" & sDummy & "32\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = "C:\" & sDummy & "32\" & sFile
            Exit Function
        End If
        If Dir$("C:\" & sDummy & "16\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = "C:\" & sDummy & "16\" & sFile
            Exit Function
        End If
        
        If Right$(sDummy, 2) = "32" Or Right$(sDummy, 2) = "16" Then
            'asssuming sFile is prog32.exe,
            'check x:\program files\prog\prog32.exe
            sDummy = Left$(sDummy, Len(sDummy) - 2)
            If Dir$(sProgramFiles & "\" & sDummy & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
                GetLongPath = sProgramFiles & "\" & sDummy & "\" & sFile
                Exit Function
            End If
            
            'check c:\prog\prog32.exe
            If Dir$("C:\" & sDummy & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
                GetLongPath = "C:\" & sDummy & "\" & sFile
                Exit Function
            End If
        End If
    End If
    
    'can't find it!
    GetLongPath = "?:\?\" & sFile
    Exit Function
    
ErrorHandler:
    ErrorMsg err, "GetLongPath", sFile
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Function

'Public Function GetFileFromShortcut(ByVal sLink$) As String
'    Dim ff%, ext$, sTarget$, sArgument$
'    On Error GoTo ErrorHandler:
'
'    ext = LCase$(right$((sLink), 4))
'
'    If ext = ".lnk" Then
'        sTarget = GetMSILinkTarget(sLink)
'        If Len(sTarget) = 0 Or "INSTALLSTATE_UNKNOWN" = sTarget Then
'            GetTargetShellLinkW sLink, sTarget, sArgument
'        Else
'            GetTargetShellLinkW sLink, , sArgument
'        End If
'    ElseIf ext = ".pif" Then
'        ff = FreeFile()
'        sTarget = String$(62, Chr$(0))
'        sArgument = String$(63, Chr$(0))
'        Open sLink For Binary Access Read Shared As #ff
'        If LOF(ff) >= &H16E& Then
'            Get #ff, &H24& + 1&, sTarget
'            Get #ff, &HA5& + 1&, sArgument
'        End If
'        Close #ff
'        sTarget = TrimNull(sTarget)
'        sArgument = TrimNull(sArgument)
'    End If
'
'    GetFileFromShortcut = sTarget
'
'    Exit Function
'ErrorHandler:
'    ErrorMsg Err, "GetFileFromShortCut", sLink
'    If inIDE Then Stop: Resume Next
'End Function

Private Function TrimNull$(s$)
    On Error GoTo ErrorHandler:
    If Len(s) = 0 Then Exit Function
    TrimNull = Left$(s, lstrlen(StrPtr(s)))
    Exit Function
ErrorHandler:
    ErrorMsg err, "TrimNull"
    If inIDE Then Stop: Resume Next
End Function

Private Function RegGetString(lHive&, sKey$, sVal$)
    Dim hKey&, sData$
    On Error GoTo ErrorHandler:
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        sData = String$(255, 0)
        If RegQueryValueEx(hKey, sVal, 0, REG_SZ, ByVal sData, 255) = 0 Then
            RegGetString = TrimNull(sData)
        Else
            RegGetString = "*Registry value not found*"
        End If
        RegCloseKey hKey
    Else
        RegGetString = "*Registry key not found*"
    End If
    Exit Function
    
ErrorHandler:
    Dim sRoot
    RegCloseKey hKey
    Select Case lHive
        Case HKEY_CURRENT_USER: sRoot = "HKCU"
        Case HKEY_LOCAL_MACHINE: sRoot = "HKLM"
        Case Else: sRoot = "HK.."
    End Select
    ErrorMsg err, "RegGetString", sRoot, sKey, sVal
    If inIDE Then Stop: Resume Next
End Function

Private Function IsRegVal404(sData$) As Boolean
    If sData = "*Registry key not found*" Or _
       sData = "*Registry value not found*" Then
        IsRegVal404 = True
    Else
        IsRegVal404 = False
    End If
End Function

Private Function RegKeyExists(lHive&, sKey$) As Boolean
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        RegKeyExists = True
        RegCloseKey hKey
    Else
        RegKeyExists = False
    End If
End Function

Private Function RegValueExists(lHive&, sKey$, sValue$) As Boolean
    Dim hKey&, sData$
    On Error GoTo ErrorHandler:
    
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        sData = String$(255, 0)
        If RegQueryValueEx(hKey, sValue, 0, ByVal 0, ByVal sData, 255) = 0 Then
            RegValueExists = True
        Else
            RegValueExists = False
        End If
        RegCloseKey hKey
    Else
        RegValueExists = False
    End If
    Exit Function
    
ErrorHandler:
    Dim sRoot$
    RegCloseKey hKey
    Select Case lHive
        Case HKEY_CURRENT_USER: sRoot = "HKCU"
        Case HKEY_LOCAL_MACHINE: sRoot = "HKLM"
        Case Else: sRoot = "HK.."
    End Select
    ErrorMsg err, "RegValueExists", sRoot, sKey, sValue
    If inIDE Then Stop: Resume Next
End Function

Private Function RegGetDword&(lRoot&, sKey$, sValue$)
    Dim hKey&, lData&
    If RegOpenKeyEx(lRoot, sKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        Exit Function
    End If
    If RegQueryValueEx(hKey, sValue, 0, REG_DWORD, lData, 4) = 0 Then
        RegGetDword = lData
    End If
    RegCloseKey hKey
End Function

Private Function RegGetFileFromBinary$(lRoot&, sKey$, sValue$)
    Dim hKey&, i&, uData() As Byte, sFile$
    If RegOpenKeyEx(lRoot, sKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        'failed to open key
        Exit Function
    End If
    
    ReDim uData(1000)
    If RegQueryValueEx(hKey, sValue, 0, REG_BINARY, uData(0), 1001) <> 0 Then
        'failed to get value
        RegCloseKey hKey
        Exit Function
    End If
    
    sFile = vbNullString
    For i = 0 To 1000
        If uData(i) = 0 Then Exit For
        sFile = sFile & Chr(uData(i))
    Next i
    
    RegGetFileFromBinary = sFile
End Function

Private Function IniGetString$(sFile$, sSection$, sValue$, bNT As Boolean)
    Dim sIniFile$, sLine$, sDummy$, sRet$, ff%
    On Error GoTo ErrorHandler:
    
    If Not bNT Then
        'win9x method - check ini FILE
        
        If Dir$(sWinDir & "\" & sFile) <> vbNullString Then
            sIniFile = sWinDir & "\" & sFile
        Else
            If Dir$(sWinDir & "\system\" & sFile) <> vbNullString Then
                sIniFile = sWinDir & "\system\" & sFile
            Else
                If Dir$(sWinDir & "\system32\" & sFile) <> vbNullString Then
                    sIniFile = sWinDir & "\system32\" & sFile
                Else
                    IniGetString = "*INI file not found*"
                    Exit Function
                End If
            End If
        End If
        
        If FileLenW(sIniFile) = 0 Then
            IniGetString = "*INI section not found*"
            Exit Function
        End If
        
        ff = FreeFile()
        Open sIniFile For Input As #ff
            Do
                Line Input #ff, sLine
            Loop Until EOF(ff) Or LCase$(sLine) = LCase$("[" & sSection & "]")
            If EOF(ff) Or LCase$(sLine) <> LCase$("[" & sSection & "]") Then
                IniGetString = "*INI section not found*"
                Close #ff
                Exit Function
            End If
            'found the [section]
            Do
                Line Input #ff, sLine
                If Len(sLine) > Len(sValue) Then
                    If LCase$(Left$(sLine, Len(sValue))) = LCase$(sValue) Then
                        'found the setting=
                        IniGetString = Mid$(sLine, Len(sValue) + 2)
                        Exit Do
                    End If
                End If
            Loop Until EOF(ff) Or sLine = vbNullString
        Close #ff
    Else
        'winNT method - check ini REGISTRY MAPPING
        'HKLM\..\Windows NT\CurrentVersion\IniFileMapping
        
        'get actual regkey location from IniFileMapping key
        sDummy = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\IniFileMapping\" & sFile, sSection)
        If IsRegVal404(sDummy) Then
            sDummy = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\IniFileMapping\" & sFile & "\" & sSection, sValue)
            If IsRegVal404(sDummy) Then
                IniGetString = sDummy
                Exit Function
            End If
        End If
        If Left$(sDummy, 1) = "!" Then sDummy = Mid$(sDummy, 2)
        If Left$(sDummy, 1) = "#" Then sDummy = Mid$(sDummy, 2)
        
        'get actual setting
        Select Case Left$(sDummy, 3)
            Case "USR"
                sDummy = RegGetString(HKEY_CURRENT_USER, Mid$(sDummy, 5), sValue)
                sRet = "[HKCU\" & Mid$(sDummy, 5) & "]"
            Case "SYS"
                sDummy = RegGetString(HKEY_LOCAL_MACHINE, Mid$(sDummy, 5), sValue)
                sRet = "[HKLM\" & Mid$(sDummy, 5) & "]"
            Case Else
                IniGetString = vbNullString
                Exit Function
        End Select
        
        IniGetString = sRet & vbCrLf & Replace$(sDummy, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
    End If
    Exit Function

ErrorHandler:
    Close #ff
    ErrorMsg err, "INIGetString", sFile, sSection, sValue, bNT
    If inIDE Then Stop: Resume Next
End Function

Private Sub HTMLize()
    Dim sHTML$, sHeader$, sFooter$, sIndex$, i&
    Dim sReplace$(1), sSections$
    
    If bHTML Then
        'fix < and > occurrances not being HTML
        sReport = Replace$(sReport, "<", "&lt;")
        sReport = Replace$(sReport, ">", "&gt;")
        
        'replace -- and == bars with <HR> bars
        sReport = Replace$(sReport, String$(50, "-"), "<CENTER><HR WIDTH=""80%""></CENTER>")
        
        'expand [tag]s
        sSections = vbNullString
        For i = 1 To 99
            If InStr(sReport, "[tag" & CStr(i) & "]") > 0 Then
                sReplace$(0) = "<A NAME=""" & CStr(i) & """><B>"
                sReplace$(1) = "</B></A>"
                sReport = Replace$(sReport, "[tag" & CStr(i) & "]", sReplace$(0))
                sReport = Replace$(sReport, "[/tag" & CStr(i) & "]", sReplace$(1))
                
                sSections = sSections & "1"
            Else
                sSections = sSections & "0"
            End If
        Next i
    Else
        'remove [tag]s
        For i = 1 To 99
            If InStr(sReport, "[tag" & CStr(i) & "]") > 0 Then
                sReport = Replace$(sReport, "[tag" & CStr(i) & "]", vbNullString)
                sReport = Replace$(sReport, "[/tag" & CStr(i) & "]", vbNullString)
            End If
        Next i
    End If
    
    If bHTML Then
        'write header
        sHeader = "<HTML><HEAD>" & vbCrLf & _
                "<TITLE>StartupList report</TITLE>" & vbCrLf & _
                "<META NAME=""Generator"" CONTENT=""StartupList v" & App.Major & "." & App.Minor & """>" & vbCrLf & _
                "<META HTTP-EQUIV=""content-type"" CONTENT=""text/html; charset=ISO-8859-1"">" & _
                "<STYLE TYPE=""text/css""><!--" & vbCrLf & _
                "body{font-family: Fixedsys, Arial, monospace;margin-left:40px;margin-right:40px}" & vbCrLf & _
                "--></STYLE>" & vbCrLf & _
                "</HEAD>" & vbCrLf & _
                "<BODY><PRE><FONT FACE=""Fixedsys"">"
                sIndex = "</FONT></PRE><BLOCKQUOTE>Sections:<BLOCKQUOTE>" & vbCrLf
        If Mid$(sSections, 1, 1) = "1" Then sIndex = sIndex & "<A HREF=""#1"">Running processes</A><BR>" & vbCrLf
        If Mid$(sSections, 2, 1) = "1" Then sIndex = sIndex & "<A HREF=""#2"">Autostart folders</A><BR>" & vbCrLf
        If Mid$(sSections, 3, 1) = "1" Then sIndex = sIndex & "<A HREF=""#3"">Windows NT UserInit</A><BR>" & vbCrLf
        If Mid$(sSections, 4, 1) = "1" Then sIndex = sIndex & "<A HREF=""#4"">Autorun key HKLM\..\Run</A><BR>" & vbCrLf
        If Mid$(sSections, 5, 1) = "1" Then sIndex = sIndex & "<A HREF=""#5"">Autorun key HKLM\..\RunOnce</A><BR>" & vbCrLf
        If Mid$(sSections, 6, 1) = "1" Then sIndex = sIndex & "<A HREF=""#6"">Autorun key HKLM\..\RunOnceEx</A><BR>" & vbCrLf
        If Mid$(sSections, 7, 1) = "1" Then sIndex = sIndex & "<A HREF=""#7"">Autorun key HKLM\..\RunServices</A><BR>" & vbCrLf
        If Mid$(sSections, 8, 1) = "1" Then sIndex = sIndex & "<A HREF=""#8"">Autorun key HKLM\..\RunServicesOnce</A><BR>" & vbCrLf
        If Mid$(sSections, 9, 1) = "1" Then sIndex = sIndex & "<A HREF=""#9"">Autorun key HKCU\..\Run</A><BR>" & vbCrLf
        If Mid$(sSections, 10, 1) = "1" Then sIndex = sIndex & "<A HREF=""#10"">Autorun key HKCU\..\RunOnce</A><BR>" & vbCrLf
        If Mid$(sSections, 11, 1) = "1" Then sIndex = sIndex & "<A HREF=""#11"">Autorun key HKCU\..\RunOnceEx</A><BR>" & vbCrLf
        If Mid$(sSections, 12, 1) = "1" Then sIndex = sIndex & "<A HREF=""#12"">Autorun key HKCU\..\RunServices</A><BR>" & vbCrLf
        If Mid$(sSections, 13, 1) = "1" Then sIndex = sIndex & "<A HREF=""#13"">Autorun key HKCU\..\RunServicesOnce</A><BR>" & vbCrLf
        If Mid$(sSections, 14, 1) = "1" Then sIndex = sIndex & "<A HREF=""#14"">Autorun key HKLM\..\Run (NT)</A><BR>" & vbCrLf
        If Mid$(sSections, 15, 1) = "1" Then sIndex = sIndex & "<A HREF=""#15"">Autorun key HKCU\..\Run (NT)</A><BR>" & vbCrLf
        If Mid$(sSections, 16, 1) = "1" Then sIndex = sIndex & "<A HREF=""#16"">Autorun subkeys HKLM\..\Run\*</A><BR>" & vbCrLf
        If Mid$(sSections, 17, 1) = "1" Then sIndex = sIndex & "<A HREF=""#17"">Autorun subkeys HKLM\..\RunOnce\*</A><BR>" & vbCrLf
        If Mid$(sSections, 18, 1) = "1" Then sIndex = sIndex & "<A HREF=""#18"">Autorun subkeys HKLM\..\RunOnceEx\*</A><BR>" & vbCrLf
        If Mid$(sSections, 19, 1) = "1" Then sIndex = sIndex & "<A HREF=""#19"">Autorun subkeys HKLM\..\RunServices\*</A><BR>" & vbCrLf
        If Mid$(sSections, 20, 1) = "1" Then sIndex = sIndex & "<A HREF=""#20"">Autorun subkeys HKLM\..\RunServicesOnce\*</A><BR>" & vbCrLf
        If Mid$(sSections, 21, 1) = "1" Then sIndex = sIndex & "<A HREF=""#21"">Autorun subkeys HKCU\..\Run\*</A><BR>" & vbCrLf
        If Mid$(sSections, 22, 1) = "1" Then sIndex = sIndex & "<A HREF=""#22"">Autorun subkeys HKCU\..\RunOnce\*</A><BR>" & vbCrLf
        If Mid$(sSections, 23, 1) = "1" Then sIndex = sIndex & "<A HREF=""#23"">Autorun subkeys HKCU\..\RunOnceEx\*</A><BR>" & vbCrLf
        If Mid$(sSections, 24, 1) = "1" Then sIndex = sIndex & "<A HREF=""#24"">Autorun subkeys HKCU\..\RunServices\*</A><BR>" & vbCrLf
        If Mid$(sSections, 25, 1) = "1" Then sIndex = sIndex & "<A HREF=""#25"">Autorun subkeys HKCU\..\RunServicesOnce\*</A><BR>" & vbCrLf
        If Mid$(sSections, 26, 1) = "1" Then sIndex = sIndex & "<A HREF=""#26"">Autorun subkeys HKLM\..\Run\* (NT)</A><BR>" & vbCrLf
        If Mid$(sSections, 27, 1) = "1" Then sIndex = sIndex & "<A HREF=""#27"">Autorun subkeys HKCU\..\Run\* (NT)</A><BR>" & vbCrLf
        If Mid$(sSections, 28, 1) = "1" Then sIndex = sIndex & "<A HREF=""#28"">Class .EXE</A><BR>" & vbCrLf
        If Mid$(sSections, 29, 1) = "1" Then sIndex = sIndex & "<A HREF=""#29"">Class .COM</A><BR>" & vbCrLf
        If Mid$(sSections, 30, 1) = "1" Then sIndex = sIndex & "<A HREF=""#30"">Class .BAT</A><BR>" & vbCrLf
        If Mid$(sSections, 31, 1) = "1" Then sIndex = sIndex & "<A HREF=""#31"">Class .PIF</A><BR>" & vbCrLf
        If Mid$(sSections, 32, 1) = "1" Then sIndex = sIndex & "<A HREF=""#32"">Class .SCR</A><BR>" & vbCrLf
        If Mid$(sSections, 33, 1) = "1" Then sIndex = sIndex & "<A HREF=""#33"">Class .HTA</A><BR>" & vbCrLf
        If Mid$(sSections, 34, 1) = "1" Then sIndex = sIndex & "<A HREF=""#34"">Class .TXT</A><BR>" & vbCrLf
        
        If Mid$(sSections, 35, 1) = "1" Then sIndex = sIndex & "<A HREF=""#35"">Active Setup Stub Paths</A><BR>" & vbCrLf
        If Mid$(sSections, 36, 1) = "1" Then sIndex = sIndex & "<A HREF=""#36"">ICQ Agent</A><BR>" & vbCrLf
        If Mid$(sSections, 37, 1) = "1" Then sIndex = sIndex & "<A HREF=""#37"">Load/Run keys from WIN.INI</A><BR>" & vbCrLf
        If Mid$(sSections, 38, 1) = "1" Then sIndex = sIndex & "<A HREF=""#38"">Shell/SCRNSAVE.EXE keys from SYSTEM.INI</A><BR>" & vbCrLf
        If Mid$(sSections, 39, 1) = "1" Then sIndex = sIndex & "<A HREF=""#39"">Explorer check</A><BR>" & vbCrLf
        If Mid$(sSections, 40, 1) = "1" Then sIndex = sIndex & "<A HREF=""#40"">Wininit.ini</A><BR>" & vbCrLf
        If Mid$(sSections, 41, 1) = "1" Then sIndex = sIndex & "<A HREF=""#41"">Wininit.bak</A><BR>" & vbCrLf
        If Mid$(sSections, 42, 1) = "1" Then sIndex = sIndex & "<A HREF=""#42"">C:\Autoexec.bat</A><BR>" & vbCrLf
        If Mid$(sSections, 43, 1) = "1" Then sIndex = sIndex & "<A HREF=""#43"">C:\Config.sys</A><BR>" & vbCrLf
        If Mid$(sSections, 44, 1) = "1" Then sIndex = sIndex & "<A HREF=""#44"">" & sWinDir & "\Winstart.bat</A><BR>" & vbCrLf
        If Mid$(sSections, 45, 1) = "1" Then sIndex = sIndex & "<A HREF=""#45"">" & sWinDir & "\Dosstart.bat</A><BR>" & vbCrLf
        If Mid$(sSections, 46, 1) = "1" Then sIndex = sIndex & "<A HREF=""#46"">Superhidden extensions</A><BR>" & vbCrLf
        If Mid$(sSections, 47, 1) = "1" Then sIndex = sIndex & "<A HREF=""#47"">Regedit.exe check</A><BR>" & vbCrLf
        If Mid$(sSections, 48, 1) = "1" Then sIndex = sIndex & "<A HREF=""#48"">BHO list</A><BR>" & vbCrLf
        If Mid$(sSections, 49, 1) = "1" Then sIndex = sIndex & "<A HREF=""#49"">Task Scheduler jobs</A><BR>" & vbCrLf
        If Mid$(sSections, 50, 1) = "1" Then sIndex = sIndex & "<A HREF=""#50"">Download Program Files</A><BR>" & vbCrLf
        If Mid$(sSections, 51, 1) = "1" Then sIndex = sIndex & "<A HREF=""#51"">Winsock LSP files</A><BR>" & vbCrLf
        If Mid$(sSections, 52, 1) = "1" Then sIndex = sIndex & "<A HREF=""#52"">Windows NT Services</A><BR>" & vbCrLf
        If Mid$(sSections, 53, 1) = "1" Then sIndex = sIndex & "<A HREF=""#53"">Windows NT Logon/logoff scripts</A><BR>" & vbCrLf
        If Mid$(sSections, 54, 1) = "1" Then sIndex = sIndex & "<A HREF=""#54"">ShellServiceObjectDelayLoad regkey</A><BR>" & vbCrLf
        If Mid$(sSections, 55, 1) = "1" Then sIndex = sIndex & "<A HREF=""#55"">Policies Explorer Run key (global)</A><BR>" & vbCrLf
        If Mid$(sSections, 56, 1) = "1" Then sIndex = sIndex & "<A HREF=""#56"">Policies Explorer Run key (current user)</A><BR>" & vbCrLf
        
        If Mid$(sSections, 57, 1) = "1" Then sIndex = sIndex & "<A HREF=""#55""> </A><BR>" & vbCrLf
        
        sIndex = sIndex & "</BLOCKQUOTE></BLOCKQUOTE>" & vbCrLf
        'write footer
        sFooter = vbCrLf & vbCrLf & "</FONT></PRE></BODY></HTML>"
        
        sReport = Replace$(sReport, String$(50, "="), sIndex & "<CENTER><HR WIDTH=""80%"" SIZE=5""></CENTER>" & "<PRE><FONT FACE=""Fixedsys"">" & vbCrLf & vbCrLf)
        sReport = sHeader & sReport & sFooter
    End If
End Sub

Public Function BuildPath$(sPath$, sFile$)
    BuildPath = sPath & IIf(Right$(sPath, 1) = "\", vbNullString, "\") & sFile
End Function

' Some mistery going here. PathFindOnPath API function doen't work on this module at all (but the same function working on the modMain module pretty well) !!!

Public Function FindOnPath2(sAppName As String) As String
    Dim ProcPath$
    ToggleWow64FSRedirection False
    ProcPath = Space$(MAX_PATH)
    LSet ProcPath = sAppName & vbNullChar
    If CBool(PathFindOnPath(StrPtr(ProcPath), 0&)) Then
        FindOnPath2 = TrimNull(ProcPath)
        ToggleWow64FSRedirection True
    Else
        ToggleWow64FSRedirection True
    End If
End Function
