Attribute VB_Name = "modStartupList"
Option Explicit
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

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal Length As Long)

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

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
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
Private Const PROCESS_VM_READ = 16

Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const TH32CS_SNAPPROCESS As Long = 2&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const HKEY_USERS = &H80000003
'Private Const HKEY_PERFORMANCE_DATA = &H80000004
'Private Const HKEY_CURRENT_CONFIG = &H80000005
'Private Const HKEY_DYN_DATA = &H80000006

Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_QUERY_VALUE = &H1

Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4

Private sReport$
Public sWinDir$, sWinSysDir$
Private bVerbose As Boolean, sVerbose$
Private bComplete As Boolean
Private bForceWin9x As Boolean
Private bForceWinNT As Boolean
Private bForceAll As Boolean
Private bFull As Boolean
Private bHTML As Boolean
Private sSLVersion$

Private bNoWriteAccess As Boolean

Private bIsWin9x As Boolean, bIsWinNT As Boolean

Public bStartupListFull As Boolean
Public bStartupListComplete As Boolean

Sub Main()
    Dim sMsg$, sPath$, lTicks&
    'MsgBox "debug: Test"
    'Test sub
    'EnumNTScripts
    'End
    
    On Error GoTo Error:
    sPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
        
    ' CHANGE THIS OR DIE
    '=======================
    sSLVersion = "1.52.2"
    '=======================
    ' FORGET AND YOU GET NO DONUTS!!!!!!!!
    
    On Error Resume Next
    Open sPath & "test.tmp" For Output As #1
        Print #1, "."
    Close #1
    Kill sPath & "test.tmp"
    If Err Then
        'may not have write access
        MsgBox "For some reason, write access was denied to " & _
        "StartupList. The log will not be written to disk, " & _
        "but copied to the clipboard instead." & vbCrLf & vbCrLf & _
        "To see it, open Notepad and hit Ctrl-V when the " & _
        "'finished' message pops up.", vbExclamation
        bNoWriteAccess = True
    End If
    On Error GoTo Error:
    
    'If 1 Then
    If InStr(Command$, "/history") > 0 Then
        sMsg = "StartupList version history:" & vbCrLf & vbCrLf
        
        '1.52.2 - added .txt extension to checkclasses
        sMsg = sMsg & "v1.52" & vbCrLf & _
                      "* Fixed stupid 'Bad filename or number' error at startup (hopefully)" & vbCrLf & _
                      "* Fixed two bugs in function that reads settings from .ini files" & vbCrLf & _
                      "* Added two more files to LSP files safelist (MS Firewall and" & vbCrLf & _
                      "  DiamondCS)" & vbCrLf & _
                      "* Fixed not detecting modified Shell line in XP (among others, this" & vbCrLf & _
                      "  BIG bug affected two sections)" & vbCrLf & _
                      "* Added listing of values in ShellServiceObjectDelayLoad regkey" & vbCrLf & _
                      vbCrLf
        
        sMsg = sMsg & "v1.51" & vbCrLf & _
                      "* Added switch: /full, which will show some rarely important" & vbCrLf & _
                      "  sections that otherwise remain hidden:" & _
                      "  Stub Paths, Explorer Check, " & vbCrLf & _
                      "  Config.sys, Dosstart.bat, Superhidden Extensions, " & vbCrLf & _
                      "  Regedit.exe Check, WinNT Services, Win9x VxD Services" & vbCrLf & _
                      "* Lines in BAT files with both 'ECHO' and '>' are now shown" & vbCrLf & _
                      "* Windows NT Logon/logoff scripts are now listed (new section)" & vbCrLf & _
                      "* Rudimentary check for PendingFileRenameOperations in NT, located" & vbCrLf & _
                      "  in above section. Also moved BootExecute check to this section" & vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.5" & vbCrLf & _
                      "* Added more files to safe list of LSP files" & vbCrLf & _
                      "* REM/ECHO line in .bat files only listed with /complete switch" & vbCrLf & _
                      "* Check for Policies\System\Shell= at SYSTEM.INI check" & vbCrLf & _
                      "* Added enumeration of Windows NT/2000/XP services (only" & vbCrLf & "  with /full switch)" & vbCrLf & _
                      "* Also lists Windows 9x Vxd services (only with /full switch)" & vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.4" & vbCrLf & _
                      "* Added listing of Winsock LSP providers" & vbCrLf & _
                      "* Fixed a NT bug with Load key" & _
                      vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.35" & vbCrLf & _
                      "* Fixed a few items not appearing in NT/2000/XP." & vbCrLf & _
                      "* Made Regedit check even more supple." & _
                      vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.34" & vbCrLf & _
                      "* Added listing of drivers= line from system.ini" & vbCrLf & _
                      "* Some more sections are now hidden if nothing interesting is there" & vbCrLf & _
                      "* Enumeration of Stub Paths now shows disabled items" & vbCrLf & _
                      "* Fixed a few bugs" & vbCrLf & _
                      "* Workaround for Atguard 'From:' bug :)" & _
                      vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.33" & vbCrLf & _
                      "* Fixed some erroneous errors." & vbCrLf & _
                      "* Added listing of MSIE version." & vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.32" & vbCrLf & _
                      "* Fixed a few bugs. That's basically it. :)" & _
                      vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.31" & vbCrLf & _
                      "* Finally added alternative (and better) method for listing processes" & vbCrLf & _
                      "  in Windows NT/2000/XP (PSAPI.DLL needed for NT4)" & vbCrLf & _
                      "* Improved filename extracting from shortcuts - StartupList should" & vbCrLf & _
                      "  not be able to extract filenames with a 100% success rate" & vbCrLf & _
                      "* Creation date is now displayed for Wininit.ini and Wininit.bak" & vbCrLf & _
                      "* Added Regedit check" & vbCrLf & _
                      "* Added listing of BHO's" & vbCrLf & _
                      "* Added listing of Task Scheduler jobs" & vbCrLf & _
                      "* Added listing of 'Download Program Files' (aka ActiveX Objects)" & _
                      vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.3" & vbCrLf & _
                      "* Added /html parameter, for a report in HTML format" & vbCrLf & _
                      "* Lots of performance enhancements, more readble code (like you care :)" & vbCrLf & _
                      "* Also some small upgrades/tweaks" & vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.23" & vbCrLf & _
                      "* Now also lists WININIT.BAK (the last WININIT.INI)" & vbCrLf & vbCrLf

        sMsg = sMsg & "v1.22" & vbCrLf & _
                      "* Made System.ini check platform independant (was Win9x only)" & vbCrLf & _
                      "* The target file & path is now extracted from enumerated shortcuts" & vbCrLf & _
                      "* Fixed MAJOR bug - GetWindowsVersion wasn't remembered, WinNT was" & vbCrLf & _
                      "  assumed" & vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.21" & vbCrLf & _
                      "* Fixed some WinNT bugs" & vbCrLf & _
                      "* Slightly improved Explorer.exe check in WinNT" & _
                      vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.2" & vbCrLf & _
                      "* Added WinNT-only startups" & vbCrLf & _
                      "* Added Windows version check" & vbCrLf & _
                      "* Added command line parameters /verbose, /complete," & vbCrLf & _
                      "  /force9x, /forcent and /forceall" & _
                      vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.1" & vbCrLf & _
                      "* Added RunOnceEx listing" & _
                      vbCrLf & vbCrLf
        
        sMsg = sMsg & "v1.0" & vbCrLf & _
                      "* Initial release" & _
                      vbCrLf & vbCrLf
        
        'sMsg = sMsg & "Written by Merijn - http://www.merijn.org/"
        'MsgBox sMsg, vbInformation, "StartupList"
        If Not bNoWriteAccess Then
            Open sPath & "versions.txt" For Output As #1
                Print #1, sMsg
            Close #1
            ShellExecute 0, "open", "notepad.exe", sPath & "versions.txt", "", 1
        Else
            Clipboard.Clear
            Clipboard.SetText sMsg
            MsgBox "StartupList has finished generating your logfile!" & vbCrLf & _
                   vbCrLf & "To see it, open Notepad and hit Ctrl-V (paste).", vbInformation
        End If
        End
    End If
        
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
    
    ''If InStr(1, App.EXEName, "hijackthis", vbTextCompare) > 0 Then
    If Len(App.EXEName) > 0 Then
        'running from built-in SL version in HijackThis
        'get parameters
        bFull = bStartupListFull
        bComplete = bStartupListComplete
    End If
    
    'MsgBox "== debug ==", vbExclamation
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
    If MsgBox(sMsg, vbQuestion + vbYesNo, "StartupList") = vbNo Then Exit Sub
    
    lTicks = GetTickCount()
    sWinDir = String(255, 0)
    sWinDir = Left(sWinDir, GetWindowsDirectory(sWinDir, 255))
    sWinSysDir = sWinDir & "\" & IIf(bIsWinNT, "system32", "system")
    
    'header
    sReport = "StartupList report, " & CStr(Date) & ", " & CStr(Time) & vbCrLf
    sReport = sReport & "StartupList version: " & sSLVersion & vbCrLf
    sReport = sReport & "Started from : " & App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & ".EXE" & vbCrLf
    sReport = sReport & "Detected: " & GetWindowsVersion & vbCrLf
    sReport = sReport & "Detected: " & GetMSIEVersion & vbCrLf
    
    sReport = sReport & IIf(Command$ = "", "* Using default options" & vbCrLf, "")
    sReport = sReport & IIf(bVerbose, "* Using verbose mode" & vbCrLf, "")
    sReport = sReport & IIf(bComplete, "* Including empty and uninteresting sections" & vbCrLf, "")
    sReport = sReport & IIf(bForceWin9x, "* Forcing include of Win9x-only sections" & vbCrLf, "")
    sReport = sReport & IIf(bForceWinNT, "* Forcing include of WinNT-only sections" & vbCrLf, "")
    sReport = sReport & IIf(bForceAll, "* Forcing include of all possible sections" & vbCrLf, "")
    sReport = sReport & IIf(bFull, "* Showing rarely important sections" & vbCrLf, "")
    sReport = sReport & String(50, "=") & vbCrLf & vbCrLf
    
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
    sReport = sReport & "   /history  - to list version history only" & vbCrLf
    
    sReport = Replace(sReport, "xXxXx", Format(Len(sReport), "###,###,###"))
    sReport = Left(sReport, Len(sReport) - 2)
    
    'add/remove HTML tags etc
    HTMLize
    
    If bHTML Then
        If Not bNoWriteAccess Then
            Open sPath & "startuplist.html" For Output As #1
                Print #1, sReport
            Close #1
            ShellExecute 0, "open", "startuplist.html", vbNullString, App.Path, 1
        Else
            Clipboard.Clear
            Clipboard.SetText sReport
            MsgBox "StartupList has finished generating your logfile!" & vbCrLf & _
                   vbCrLf & "To see it, open Notepad and hit Ctrl-V (paste).", vbInformation
        End If
    Else
        If Not bNoWriteAccess Then
            Open sPath & "startuplist.txt" For Output As #1
                Print #1, sReport
            Close #1
            ShellExecute 0, "open", "notepad.exe", sPath & "startuplist.txt", "", 1
        Else
            Clipboard.Clear
            Clipboard.SetText sReport
            MsgBox "StartupList has finished generating your logfile!" & vbCrLf & _
                   vbCrLf & "To see it, open Notepad and hit Ctrl-V (paste).", vbInformation
        End If
    End If
    Exit Sub
    
Error:
    Close
    ErrorMsg Err.Number, Err.Description, "Main", Command$
End Sub

Private Sub EnumKeys(lRootKey&, sAutorunKey$, bNT As Boolean, iSection%)
    If bNT Then
        'sub applies to NT only
        If Not bIsWinNT Then
            'this is not NT,override?
            If Not (bForceAll Or bForceWinNT) Then Exit Sub
        End If
    End If
    
    Dim sResult$, bInteresting As Boolean
    Dim hKey&, sValue$, lType&, i%, j%, sData$
    On Error GoTo Error:
    
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
    sData = String(lEnumBufSize, 0)
    sValue = String(lEnumBufSize, 0)
    If RegEnumValue(hKey, i, sValue, Len(sValue), 0, lType, ByVal sData, Len(sData)) <> 0 Then
        sResult = sResult & "*No values found*" & vbCrLf
        GoTo EndOfSub
    End If
    Do
        If lType = REG_SZ Then
            bInteresting = True
            sData = TrimNull(sData)
            sValue = Left(sValue, InStr(sValue, Chr(0)) - 1)
            If sValue = vbNullString Then sValue = "(Default)"
            sResult = sResult & sValue & " = " & sData & vbCrLf
        End If
        sData = String(lEnumBufSize, 0)
        sValue = String(lEnumBufSize, 0)
        i = i + 1
    Loop Until RegEnumValue(hKey, i, sValue, Len(sValue), 0, lType, ByVal sData, Len(sData)) <> 0

EndOfSub:
    RegCloseKey hKey
    If bVerbose Then
        sVerbose = "This lists programs that run Registry keys marked by Windows as" & vbCrLf
        sVerbose = sVerbose & "'Autostart key'. To the left are values that are used to clarify what" & vbCrLf
        sVerbose = sVerbose & "program they belong to, to the right the program file that is started." & vbCrLf
        If Right(sAutorunKey, 4) = "Once" Then
            sVerbose = sVerbose & "The values in the 'RunOnce', 'RunOnceEx' and 'RunServicesOnce' keys" & vbCrLf
            sVerbose = sVerbose & "are run once and then deleted by Windows." & vbCrLf
        End If
        sResult = sResult & vbCrLf & sVerbose
    End If
    sResult = sResult & vbCrLf
    sResult = sResult & String(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    RegCloseKey hKey
    Dim sRoot
    Select Case lRootKey
        Case HKEY_CURRENT_USER: sRoot = "HKLCU"
        Case HKEY_LOCAL_MACHINE: sRoot = "HKLM"
    End Select
    ErrorMsg Err.Number, Err.Description, "EnumKeys", sRoot & ", " & sAutorunKey & ", " & bNT
End Sub

Private Sub CheckClasses(sSubKey$, iSection%)
    'sub applies to all windows versions
    
    Dim sResult$
    sSubKey = UCase(sSubKey)
    sResult = sResult & "[tag" & iSection & "]File association entry for " & sSubKey & ":[/tag" & iSection & "]" & vbCrLf
    
    Dim hKey&, i%, sData$, bInteresting As Boolean
    On Error GoTo Error:
    
    sData = RegGetString(HKEY_CLASSES_ROOT, sSubKey, "")
    If IsRegVal404(sData) Then
        sResult = sResult & sData & vbCrLf
        GoTo EndOfSub:
    End If
    If sData = vbNullString Then Exit Sub
        
    sData = sData & "\shell\open\command"
    sResult = sResult & "HKEY_CLASSES_ROOT\" & sData & vbCrLf & vbCrLf
    
    sData = RegGetString(HKEY_CLASSES_ROOT, sData, "")
    If IsRegVal404(sData) Then
        sResult = sResult & sData & vbCrLf
        GoTo EndOfSub
    End If
    
    Select Case sSubKey
        Case ".EXE", ".COM", ".BAT", ".PIF"
            If LCase(sData) <> """%1"" %*" Then bInteresting = True
        Case ".SCR":
            If LCase(sData) <> """%1"" /s" And _
               LCase(sData) <> """%1"" /s ""%3""" Then bInteresting = True
        Case ".HTA":
            If LCase(sData) <> LCase(sWinDir) & "\system" & IIf(bIsWin9x, "", "32") & "\mshta.exe ""%1"" %*" Then bInteresting = True
        Case ".TXT":
            If LCase(sData) <> sWinDir & "\notepad.exe %1" And _
               LCase(sData) <> "%systemroot%\system32\notepad.exe %1" Then bInteresting = True
        Case Else
            MsgBox "jackass coder  - no donuts"
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
        sReport = sReport & sResult & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
    End If
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "CheckClasses", sSubKey
End Sub

Private Sub CheckWinINI()
    'sub applies to all versions

    Dim sResult$, bInteresting As Boolean
    On Error GoTo Error:
    
    sResult = sResult & "[tag36]Load/Run keys from " & sWinDir & "\WIN.INI:[/tag36]"
    sResult = sResult & vbCrLf & vbCrLf
    
    Dim sRet$
    sRet = IniGetString("win.ini", "windows", "load", False)
    'If Not IsRegVal404(sRet) Then
    If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
        bInteresting = True
    End If
    sResult = sResult & "load=" & sRet & vbCrLf
    
    sRet = IniGetString("win.ini", "windows", "run", False)
    'If Not IsRegVal404(sRet) Then
    If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
        bInteresting = True
    End If
    sResult = sResult & "run=" & sRet & vbCrLf
    
    'nt only: inifile mapping of win.ini
    If bIsWinNT Or bForceWinNT Or bForceAll Then
        sResult = sResult & vbCrLf & "Load/Run keys from Registry:" & vbCrLf & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "load")
        'sRet = INIGetString("win.ini", "windows", "load", True)
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\WinLogon: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "run")
        'sRet = INIGetString("win.ini", "windows", "run", True)
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\WinLogon: run=" & sRet & vbCrLf
        
        'the below 6 probably don't work, but it's just in case
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinLogon", "load")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows\CurrentVersion\WinLogon: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinLogon", "run")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows\CurrentVersion\WinLogon: run=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "load")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows NT\CurrentVersion\WinLogon: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "run")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows NT\CurrentVersion\WinLogon: run=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\WinLogon", "load")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows\CurrentVersion\WinLogon: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\WinLogon", "run")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows\CurrentVersion\WinLogon: run=" & sRet & vbCrLf
        
        'this is a new one
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "load")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows NT\CurrentVersion\Windows: load=" & sRet & vbCrLf
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "run")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Windows NT\CurrentVersion\Windows: run=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "load")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\Windows: load=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "run")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKLM\..\Windows NT\CurrentVersion\Windows: run=" & sRet & vbCrLf
        
        'this shouldn't really belong here, but anyway..
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs")
        sRet = Replace(sRet, Chr(0), "|")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
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
    sResult = sResult & String(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "CheckWinINI"
End Sub

Private Sub CheckSystemINI()
    'sub applies to all versions
    
    Dim sResult$, sRet$, bInteresting As Boolean
    sResult = sResult & "[tag37]Shell & screensaver key from " & sWinDir & "\SYSTEM.INI:[/tag37]"
    sResult = sResult & vbCrLf & vbCrLf
    
    On Error GoTo Error:
    sRet = IniGetString("system.ini", "boot", "shell", False)
    If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString = 0 Then
        bInteresting = True
    End If
    sResult = sResult & "Shell=" & sRet & vbCrLf
    
    sRet = IniGetString("system.ini", "boot", "SCRNSAVE.EXE", False)
    If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString = 0 Then
        bInteresting = True
    End If
    sResult = sResult & "SCRNSAVE.EXE=" & sRet & vbCrLf
    
    sRet = IniGetString("system.ini", "boot", "drivers", False)
    If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString = 0 Then
        bInteresting = True
    End If
    sResult = sResult & "drivers=" & sRet & vbCrLf
    
    If bIsWinNT Or bForceWinNT Or bForceAll Then
        sResult = sResult & vbCrLf & "Shell & screensaver key from Registry:" & vbCrLf & vbCrLf
        
        'screw this, I know where it is anyway
        'sRet = INIGetString("system.ini", "boot", "shell", True)
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "shell")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "Shell=" & sRet & vbCrLf
        
        'sRet = INIGetString("system.ini", "boot", "SCRNSAVE.EXE", True)
        sRet = RegGetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "SCRNSAVE.EXE=" & sRet & vbCrLf
        
        'doesn't appear in IniFileMapping ?
        sRet = IniGetString("system.ini", "boot", "drivers", True)
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "drivers=" & sRet & vbCrLf
        
        'got an extra one from policies key!
        sResult = sResult & vbCrLf & "Policies Shell key:" & vbCrLf & vbCrLf
        sRet = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "Shell")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
            bInteresting = True
        End If
        sResult = sResult & "HKCU\..\Policies: Shell=" & sRet & vbCrLf
        
        sRet = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "Shell")
        If InStr(sRet, "not found*") = 0 And Trim(sRet) <> vbNullString Then
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
    sResult = sResult & String(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "CheckSystemINI"
End Sub

Private Sub EnumBAT(sFile$)
    'sub applies to 9x only
    If Not bIsWin9x Then
        'this is not 9x, override?
        If Not (bForceAll Or bForceWin9x) Then Exit Sub
    End If
    'display config.sys and dosstart.bat
    'only when using /full

    Dim sResult$, bInteresting As Boolean
    On Error GoTo Error:
    If sFile = "c:\autoexec.bat" Then
        sResult = sResult & "[tag41]" & UCase(sFile) & " listing:[/tag41]" & vbCrLf & vbCrLf
    ElseIf sFile = "c:\config.sys" And bFull Then
        sResult = sResult & "[tag42]" & UCase(sFile) & " listing:[/tag42]" & vbCrLf & vbCrLf
    ElseIf sFile = sWinDir & "\winstart.bat" Then
        sResult = sResult & "[tag43]" & UCase(sFile) & " listing:[/tag43]" & vbCrLf & vbCrLf
    ElseIf sFile = sWinDir & "\dosstart.bat" And bFull Then
        sResult = sResult & "[tag44]" & UCase(sFile) & " listing:[/tag44]" & vbCrLf & vbCrLf
    Else
        Exit Sub
    End If
    
    Dim sLine$
    If Dir(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) = "" Then
        sResult = sResult & "*File not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    If FileLen(sFile) = 0 Then
        sResult = sResult & "*File is empty*" & vbCrLf
        GoTo EndOfSub
    End If
    
    On Error Resume Next '<- NT generates stupid error here
    Open sFile For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> "" Then
                If Left(sLine, 1) = "@" Then sLine = Mid(sLine, 2)
                If InStr(sLine, vbTab) > 0 Then sLine = Replace(sLine, vbTab, " ")
                If UCase(Trim(sLine)) <> "REM" And _
                   (InStr(1, sLine, "REM ", vbTextCompare) <> 1 And _
                   (InStr(1, sLine, "ECHO ", vbTextCompare) <> 1) Or _
                    InStr(sLine, ">") > 0) Or _
                   bComplete Then
                    bInteresting = True
                    sResult = sResult & sLine & vbCrLf
                End If
            End If
        Loop Until EOF(1)
    Close #1
    On Error GoTo 0:
    
EndOfSub:
    If bVerbose Then
        If InStr(LCase(sFile), "c:\autoexec.bat") > 0 Then
            sVerbose = "Autoexec.bat is the very first file to autostart when the computer" & vbCrLf
            sVerbose = sVerbose & "starts, it is a leftover from DOS and older Windows versions." & vbCrLf
            sVerbose = sVerbose & "Windows NT, Windows ME, Windows 2000 and Windows XP don't use this" & vbCrLf
            sVerbose = sVerbose & "file. It is generally used by virusscanners to scan files before" & vbCrLf
            sVerbose = sVerbose & "Windows starts."
            sResult = sResult & vbCrLf & sVerbose & vbCrLf
        ElseIf InStr(LCase(sFile), "c:\config.sys") > 0 Then
            sVerbose = "Config.sys loads device drivers for DOS, and is rarely used in" & vbCrLf & _
                       "Windows versions newer than Windows 95. Originally it loaded" & vbCrLf & _
                       "drivers for legacy sound cards and such."
            sResult = sResult & vbCrLf & sVerbose & vbCrLf
        ElseIf InStr(LCase(sFile), "winstart.bat") > 0 Then
            sVerbose = "Winstart.bat loads just before the Windows shell, and is used for" & vbCrLf
            sVerbose = sVerbose & "starting things like soundcard drivers, mouse drivers. Rarely used."
            sResult = sResult & vbCrLf & sVerbose & vbCrLf
        ElseIf InStr(LCase(sFile), "dosstart.bat") > 0 Then
            sVerbose = "Dosstart.bat loads if you select 'MS-DOS Prompt' from the Startup" & vbCrLf
            sVerbose = sVerbose & "menu when the computer is starting, or if you select 'Restart in" & vbCrLf
            sVerbose = sVerbose & "MS-DOS Mode' from the Shutdown menu in Windows. Mostly used for" & vbCrLf
            sVerbose = sVerbose & "DOS-only drivers, like sound or mouse drivers."
            sResult = sResult & vbCrLf & sVerbose & vbCrLf
        End If
    End If
    sResult = sResult & vbCrLf
    sResult = sResult & String(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    Close
    ErrorMsg Err.Number, Err.Description, "EnumBAT", sFile
End Sub

Private Sub EnumWininit(sIniFile$)
    'sub applies to 9x only
    If Not bIsWin9x Then
        'this is not 9x, override?
        If Not (bForceAll Or bForceWin9x) Then Exit Sub
    End If
    
    Dim sResult$, bInteresting As Boolean
    Dim lFileHandle&, uLastWritten As FILETIME
    Dim uLocalTime As FILETIME, uSystemTime As SYSTEMTIME
    On Error GoTo Error:
    
    If sIniFile = "wininit.ini" Then
        sResult = sResult & "[tag39]" & sWinDir & "\" & UCase(sIniFile) & " listing:[/tag39]" & vbCrLf
    ElseIf sIniFile = "wininit.bak" Then
        sResult = sResult & "[tag40]" & sWinDir & "\" & UCase(sIniFile) & " listing:[/tag40]" & vbCrLf
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
    If Dir(sWinDir & "\" & sIniFile, vbArchive + vbHidden + vbReadOnly + vbSystem) = "" Then
        sResult = sResult & "*File not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    bInteresting = True
    On Error Resume Next '<- NT makes some stupid error here
    Open sWinDir & "\" & sIniFile For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> "" Then sResult = sResult & sLine & vbCrLf
        Loop Until EOF(1)
    Close #1
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
    sResult = sResult & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    Close
    ErrorMsg Err.Number, Err.Description, "EnumWininit", sIniFile
End Sub

Private Sub CheckAutoStartFolders()
    'sub applies to all windows versions
    
    Dim sResult$, ss$
    'Dim sDummy$, hKey&, uData() As Byte, i%, sData$
    On Error GoTo Error:
    
    'checking all *8* possible folders now - 1.52+
    sResult = sResult & ListFiles(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup", "Shell folders Startup")
    ss = ListFiles(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AltStartup", "Shell folders AltStartup")
    sResult = sResult & ss
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
    sResult = sResult & String(50, "-")
    sResult = sResult & vbCrLf & vbCrLf
    
    If sResult <> sVerbose & String(50, "-") & vbCrLf & vbCrLf Then
        sReport = sReport & "[tag2]Listing of startup folders:[/tag2]" & vbCrLf & vbCrLf
        sReport = sReport & sResult
    End If
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "CheckAutoStartFolders"
End Sub

Private Function ListFiles$(lRootKey&, sSubKey$, sValue$, sName$)
    'sub applies to all windows versions
    
    'ListFiles(HKEY_CURRENT_USER,
    '          "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
    '          "Startup",
    '          "Shell folders Startup")
    
    Dim sResult$, bInteresting As Boolean, sLastFile$
    Dim hKey&, sData$, i%, sFile$ ', uData() As Byte
    On Error GoTo Error:
    
    sResult = sName & ":" & vbCrLf
    
    sData = RegGetString(lRootKey, sSubKey, sValue)
    If IsRegVal404(sData) Or Dir(sData & "\NUL", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "NUL" Then
        sResult = sResult & "*Folder not found*" & vbCrLf & vbCrLf
        GoTo EndOfFunction
    End If
    
    If UCase(Dir(sData & "\NUL")) = "NUL" Then
        sResult = sResult & "[" & sData & "]" & vbCrLf
        sFile = Dir(sData & "\*.*", vbArchive + vbHidden + vbReadOnly + vbSystem)
        If sFile = vbNullString Then
            sResult = sResult & "*No files*" & vbCrLf & vbCrLf
            GoTo EndOfFunction
        End If
        Do
            If LCase(sFile) <> "desktop.ini" Then
                sResult = sResult & sFile & GetFileFromShortCut(sData & "\" & sFile) & vbCrLf
                bInteresting = True
            End If
            sFile = Dir
        Loop Until sFile = ""
        If Not bInteresting Then sResult = sResult & "*No files*" & vbCrLf
        sResult = sResult & vbCrLf
    Else
        sResult = sResult & "*Folder not found*" & vbCrLf & vbCrLf
        GoTo EndOfFunction:
    End If
    
EndOfFunction:
    If bInteresting Or bComplete Then ListFiles = sResult
    Exit Function
    
Error:
    Dim sRoot
    Select Case lRootKey
        Case HKEY_CURRENT_USER: sRoot = "HKCU"
        Case HKEY_LOCAL_MACHINE: sRoot = "HKLM"
        Case Else: sRoot = "HK.."
    End Select
    ErrorMsg Err.Number, Err.Description, "ListFiles", sRoot & ", " & sSubKey & ", " & sValue & ", " & sName
End Function

Private Sub CheckNeverShowExt(sSubKey$, Optional bOverlay As Boolean = False)
    'sub applies to all windows versions
    
    Dim hKey&, i%, sData$
    On Error GoTo Error:
    
    sData = RegGetString(HKEY_CLASSES_ROOT, sSubKey, "")
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
    
Error:
    ErrorMsg Err.Number, Err.Description, "CheckNeverShowExt", sSubKey & ", " & bOverlay
End Sub

Private Sub CheckExplorer()
    'sub applies to all windows versions
    '(9x always, NT if improperly configured)
    'only display when using /full switch
    If Not bFull Then Exit Sub
    
    sReport = sReport & "[tag38]Checking for EXPLORER.EXE instances:[/tag38]" & vbCrLf & vbCrLf
    
    On Error Resume Next
    sReport = sReport & sWinDir & "\Explorer.exe: " & IIf(Dir(sWinDir & "\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "", "PRESENT!", "not present") & vbCrLf & vbCrLf
    sReport = sReport & "C:\Explorer.exe: " & IIf(Dir("c:\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "", "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\Explorer\Explorer.exe: " & IIf(Dir(sWinDir & "\explorer\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "", "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\System\Explorer.exe: " & IIf(Dir(sWinDir & "\system\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "", "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\System32\Explorer.exe: " & IIf(Dir(sWinDir & "\system\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "", "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\Command\Explorer.exe: " & IIf(Dir(sWinDir & "\command\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "", "PRESENT!", "not present") & vbCrLf
    sReport = sReport & sWinDir & "\Fonts\Explorer.exe: " & IIf(Dir(sWinDir & "\command\explorer.exe", vbArchive + vbHidden + vbReadOnly + vbSystem) <> "", "PRESENT!", "not present") & vbCrLf
    On Error GoTo Error:
    
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
    sReport = sReport & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "CheckExplorer"
End Sub

Private Sub EnumStubPaths()
    'sub applies to all windows versions
    'only display when using /full switch
    If Not bFull Then Exit Sub
    
    Dim sResult$, bInteresting As Boolean
    On Error GoTo Error:
    sResult = sResult & "[tag34]Enumerating Active Setup stub paths:[/tag34]" & vbCrLf
    sResult = sResult & "HKLM\Software\Microsoft\Active Setup\Installed Components" & vbCrLf
    sResult = sResult & "(* = disabled by HKCU twin)" & vbCrLf & vbCrLf
    Dim hKey&, hSubKey&, i%, j% ', uData() As Byte
    Dim sData$, sVal$, bHasHKCUTwin As Boolean
    
    'Open Active Setup registry key
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Active Setup\Installed Components", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        sResult = sResult & "*Registry key not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    i = 0
    sVal = String(255, 0)
    'Start enumerating subkeys of 'Installed Components' key
    If RegEnumKey(hKey, i, sVal, 255) <> 0 Then
        sResult = sResult & "*No subkeys found*" & vbCrLf
        RegCloseKey hKey
        GoTo EndOfSub
    End If
    Do
        sVal = Left(sVal, InStr(sVal, Chr(0)) - 1)
        'Try to open each enumerated subkey
        sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Active Setup\Installed Components\" & sVal, "StubPath")
        If sData <> vbNullString And Not IsRegVal404(sData) Then
            If LCase(Left(sData, 10)) <> "rundll.exe" And _
               LCase(Left(sData, 8)) <> "rundll32.exe" And _
               LCase(Left(sData, 9)) <> "rundll32 " And _
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
                sResult = sResult & "[" & sVal & "]" & IIf(bHasHKCUTwin, " *", "") & vbCrLf
                sResult = sResult & "StubPath = " & sData & vbCrLf & vbCrLf
            End If
        End If
        
        i = i + 1
        sVal = String(255, 0)
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
    sResult = sResult & String(50, "-") & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg Err.Number, Err.Description, "EnumStubPaths"
End Sub

Private Sub EnumICQAgentProgs()
    'sub applies to all windows versions
    
    Dim sResult$, bInteresting As Boolean
    On Error GoTo Error:
    
    sResult = sResult & "[tag35]Enumerating ICQ Agent Autostart apps:[/tag35]" & vbCrLf
    sResult = sResult & "HKCU\Software\Mirabilis\ICQ\Agent\Apps" & vbCrLf & vbCrLf
    Dim hKey&, hSubKey&, sVal$, i%, j%, sData$ ', uData() As Byte
    
    'Open ICQ Agent key
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Mirabilis\ICQ\Agent\Apps", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        sResult = sResult & "*Registry key not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    'Start enumerating subkeys
    i = 0
    sVal = String(255, 0)
    If RegEnumKey(hKey, i, sVal, 255) <> 0 Then
        sResult = sResult & "*No subkeys found*" & vbCrLf
        GoTo EndOfSub
    End If
    Do
        sVal = Left(sVal, InStr(sVal, Chr(0)) - 1)
        'Try to open each enumerated subkey
        
        sData = RegGetString(HKEY_CURRENT_USER, "Software\Mirabilis\ICQ\Agent\Apps\" & sVal, "Path")
        If sData <> vbNullString Then bInteresting = True
        sResult = sResult & sData & IIf(Right(sData, 1) = "\", "", "\")
        sData = RegGetString(HKEY_CURRENT_USER, "Software\Mirabilis\ICQ\Agent\Apps\" & sVal, "Startup")
        sResult = sResult & sData & vbCrLf
        
        i = i + 1
        sVal = String(255, 0)
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
    sResult = sResult & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg Err.Number, Err.Description, "EnumICQAgentProgs"
End Sub

Private Sub EnumExKeys(lHive&, ByVal sKey$, bWinNT As Boolean, iNumber%)
    'sub applies to all windows versions
    
    Dim sResult$, bInteresting As Boolean
    sKey = "Software\Microsoft\" & IIf(bWinNT, "Windows NT", "Windows") & "\CurrentVersion\" & sKey
    sResult = "[tag" & CStr(iNumber) & "]Autorun entries in Registry subkeys of:[/tag" & CStr(iNumber) & "]" & vbCrLf
    sResult = sResult & IIf(lHive = HKEY_LOCAL_MACHINE, "HKLM\", "HKCU\") & sKey & vbCrLf '& vbCrLf
    Dim hKey&, sVal$, i%, j%, k% ', uData() As Byte
    Dim sData$, hSubKey&, lType&
    On Error GoTo Error:
    
    'Open RunOnceEx key
    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
    'If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx", 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        sResult = sResult & "*Registry key not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    'Start enumerating subkeys
    i = 0
    sVal = String(255, 0)
    If RegEnumKey(hKey, i, sVal, 255) <> 0 Then
        sResult = sResult & "*No subkeys found*" & vbCrLf
        GoTo EndOfSub
    End If
    Do
        'Open each subkey...
        sVal = Left(sVal, InStr(sVal, Chr(0)) - 1)
        If RegOpenKeyEx(lHive, sKey & "\" & sVal, 0, KEY_QUERY_VALUE, hSubKey) = 0 Then
            sResult = sResult & vbCrLf & "[" & sVal & "]" & vbCrLf
            'And enumerate values in it
            j = 0
            sVal = String(lEnumBufSize, 0)
            sData = String(lEnumBufSize, 0)
            If RegEnumValue(hSubKey, j, sVal, Len(sVal), 0, lType, ByVal sData, Len(sData)) <> 0 Then
                sResult = sResult & "*No values found*" & vbCrLf
            Else
                Do
                    bInteresting = True
                    sVal = Left(sVal, InStr(sVal, Chr(0)) - 1)
                    sResult = sResult & sVal & " = "
                    sData = TrimNull(sData)
                    sResult = sResult & sData & vbCrLf
                    j = j + 1
                    sVal = String(lEnumBufSize, 0)
                    sData = String(lEnumBufSize, 0)
                Loop Until RegEnumValue(hSubKey, j, sVal, Len(sVal), 0, lType, ByVal sData, Len(sData)) <> 0
            End If
            sResult = sResult '& vbCrLf
            RegCloseKey hSubKey
        End If
        i = i + 1
        sVal = String(255, 0)
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
    sResult = sResult & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
    
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    RegCloseKey hKey
    RegCloseKey hSubKey
    ErrorMsg Err.Number, Err.Description, "EnumRunOnceEx"
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
    On Error GoTo Error:
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
        sDummy = Left(uProcess.szExeFile, InStr(uProcess.szExeFile, Chr(0)) - 1)
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
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i%
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
    On Error GoTo Error:
    
    lNumProcesses = lNeeded / 4
    For i = 1 To lNumProcesses
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(i))
        If hProc <> 0 Then
            lNeeded = 0
            sProcessName = String(260, 0)
            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                sProcessName = TrimNull(sProcessName)
                If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                If InStr(1, sProcessName, "%SYSTEMROOT%", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)
                If InStr(1, sProcessName, "SYSTEMROOT", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)
                
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
    sReport = sReport & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
    Exit Sub
    
Error:
    CloseHandle hSnap
    ErrorMsg Err.Number, Err.Description, "ListRunningProcesses"
End Sub

Private Sub CheckWinNTUserInit()
    'sub applies to NT only
    If Not bIsWinNT Then
        'this is not NT,override?
        If Not (bForceAll Or bForceWinNT) Then Exit Sub
    End If
    
    Dim sDummy$, sResult$, bInteresting As Boolean
    Dim hKey&, sData$, lRet&, i% ', uData() As Byte
    On Error GoTo Error:
    
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
    If sResult <> "" And sResult <> sResult & sVerbose & vbCrLf & vbCrLf Then
        sReport = sReport & "[tag3]Checking Windows NT UserInit:[/tag3]" & vbCrLf & vbCrLf
        sReport = sReport & sResult & String(50, "-") & vbCrLf & vbCrLf
    End If
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg Err.Number, Err.Description, "CheckWinNTUserInit"
End Sub

Private Sub CheckSuperHiddenExt()
    'sub applies to all windows versions
    'only display when using /full switch
    If Not bFull Then Exit Sub
    
    On Error GoTo Error:
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
    sReport = sReport & vbCrLf & String(50, "-") & vbCrLf
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "CheckSuperHiddenExt"
End Sub

Private Sub CheckRegedit()
    '35
    Dim sCmd$, sResult$, sRegedit$, bInteresting As Boolean
    'display section only with /full switch
    If Not bFull Then Exit Sub
    
    On Error GoTo Error:
    sResult = "[tag46]Verifying REGEDIT.EXE integrity:[/tag46]" & vbCrLf & vbCrLf
    sRegedit = sWinDir & "\Regedit.exe"
    
    'check location of REGEDIT.EXE, should be WinDir
    If Dir(sRegedit, vbArchive + vbReadOnly + vbSystem) <> vbNullString Then
        sResult = sResult & "- Regedit.exe found in " & sWinDir & vbCrLf
    Else
        sResult = sResult & "- Regedit.exe is MISSING!" & vbCrLf
        bInteresting = True
    End If
    
    'check .reg open command, should be regedit.exe "%1"
    sCmd = RegGetString(HKEY_CLASSES_ROOT, "regfile\shell\open\command", "")
    sCmd = Replace(sCmd, """", "")
    'If LCase(sCmd) = "regedit.exe ""%1""" Or _
    '   LCase(sCmd) = LCase(sRegedit) & """%1""" Or _
    '   LCase(sCmd) = "%systemroot%\regedit.exe ""%1""" Then
    If Left(LCase(sCmd), 11) = "regedit.exe" And _
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
                sCompanyName = String(lDataLen, 0)
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
                sOriginalFilename = String(lDataLen, 0)
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
                sFileDescription = String(lDataLen, 0)
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
        sReport = sReport & vbCrLf & sResult & vbCrLf & String(50, "-") & vbCrLf
    End If
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "CheckRegedit"
End Sub

Private Sub EnumBHOs()
    '36
    Dim hKey&, sCLSID$, sName$, sFile$, i&, sResult$
    Dim bInteresting As Boolean, bDisabledBHODemon As Boolean
    On Error GoTo Error:
    sResult = "[tag47]Enumerating Browser Helper Objects:[/tag47]" & vbCrLf & vbCrLf
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects", 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
        sCLSID = String(255, 0)
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
            sName = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects\" & sCLSID, "")
            
            'get filename
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InProcServer32", "")
            
            'check for BHODemon-disabled BHO's
            bDisabledBHODemon = False
            If Len(sFile) > 18 Then
                If Right(sFile, 18) = "__BHODemonDisabled" Then
                    sFile = Left(sFile, Len(sFile) - 18)
                    bDisabledBHODemon = True
                End If
            End If
            
            sResult = sResult & IIf(sName <> vbNullString And Not IsRegVal404(sName), sName, "(no name)") & _
                      " - " & IIf(sFile <> vbNullString And Not IsRegVal404(sFile), sFile, "(no file)") & _
                      IIf(bDisabledBHODemon, " (disabled by BHODemon)", "") & _
                      IIf(FileExists(sFile) = False And (sFile <> vbNullString And Not IsRegVal404(sFile)), " (file missing)", "") & _
                      " - " & sCLSID & vbCrLf
            
            i = i + 1
            sCLSID = String(255, 0)
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
        sReport = sReport & vbCrLf & sResult & vbCrLf & String(50, "-") & vbCrLf
    End If
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg Err.Number, Err.Description, "EnumBHOs"
End Sub

Private Sub EnumJOBs()
    '37
    On Error GoTo Error:
    Dim sResult$, bInteresting As Boolean, sFile$
    Dim sDummy$, vDummy As Variant, sBla, i&
    sResult = "[tag48]Enumerating Task Scheduler jobs:[/tag48]" & vbCrLf & vbCrLf
    If Dir(sWinDir & "\Tasks\NUL") = vbNullString Then
        sResult = sResult & "*" & sWinDir & "\Tasks folder not found*" & vbCrLf
        GoTo EndOfSub:
    End If
    
    sFile = Dir(sWinDir & "\Tasks\*.job", vbArchive + vbHidden + vbReadOnly + vbSystem)
    If sFile = vbNullString Then
        sResult = sResult & "*No jobs found*" & vbCrLf
        GoTo EndOfSub:
    End If
    bInteresting = True
    Do
        sResult = sResult & sFile & vbCrLf
        
        sFile = Dir
    Loop Until sFile = ""
    
    
EndOfSub:
    If bVerbose Then
        sVerbose = "The Windows Task Scheduler can run programs " & _
                   "at a certain time," & vbCrLf & "automatically. Though very " & _
                   "unlikely, this can be exploited by" & vbCrLf & "making a job " & _
                   "that runs a virus or trojan."
        sResult = sResult & vbCrLf & sVerbose & vbCrLf
    End If
    
    If bInteresting Or bComplete Then
        sReport = sReport & vbCrLf & sResult & vbCrLf & String(50, "-") & vbCrLf
    End If
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "EnumJOBs"
End Sub

Private Sub EnumDPF()
    '38
    Dim hKey&, sName$, i&, sResult$, bInteresting As Boolean
    Dim sFriendlyName$, sCodeBase$, sOSD$, sINF$, sFile$
    Const sKeyDPF$ = "Software\Microsoft\Code Store Database\Distribution Units"
    On Error GoTo Error:
    sResult = "[tag49]Enumerating Download Program Files:[/tag49]" & vbCrLf & vbCrLf
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyDPF, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        sResult = sResult & "*Registry key not found*" & vbCrLf
        Exit Sub
    End If
    
    sName = String(255, 0)
    If RegEnumKey(hKey, i, sName, 255) <> 0 Then
        sResult = sResult & "*No subkeys found*" & vbCrLf
        Exit Sub
    End If
    
    Do
        sName = TrimNull(sName)
        If Left(sName, 1) = "{" And Right(sName, 1) = "}" Then
            'it's a CLSID, so get real name from HKCR\CLSID
            sFriendlyName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName, "")
            sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sName & "\InProcServer32", "")
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
                sResult = sResult & IIf(sFile <> vbNullString And Not IsRegVal404(sFile), "InProcServer32 = " & sFile & vbCrLf, "") & _
                                IIf(sCodeBase <> vbNullString And Not IsRegVal404(sCodeBase), "CODEBASE = " & sCodeBase & vbCrLf, "") & _
                                IIf(sOSD <> vbNullString And Not IsRegVal404(sOSD), "OSD = " & sOSD & vbCrLf, "")
                                'IIf(sINF <> vbNullString And Not IsRegVal404(sINF), "INF = " & sINF & vbCrLf, "")
                sResult = sResult & vbCrLf
            End If
        End If
        i = i + 1
        sName = String(255, 0)
        sFriendlyName = ""
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
        sReport = sReport & vbCrLf & sResult & String(50, "-") & vbCrLf
    End If
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg Err.Number, Err.Description, "EnumDPF"
End Sub

Private Sub EnumLSP()
    '39
    Dim i%, sKeyBase$, iNumNameSpace%, iNumProtocol%
    Dim sNameSpace$, sProtocol$, sFile$, sSafeFiles$
    Dim sResult$, bInteresting As Boolean
    On Error GoTo Error:
    
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
    sNameSpace = RegGetString(HKEY_LOCAL_MACHINE, sKeyBase, "Current_NameSpace_Catalog")
    sProtocol = RegGetString(HKEY_LOCAL_MACHINE, sKeyBase, "Current_Protocol_Catalog")
    iNumNameSpace = RegGetDword(HKEY_LOCAL_MACHINE, sKeyBase & "\" & sNameSpace, "Num_Catalog_Entries")
    iNumProtocol = RegGetDword(HKEY_LOCAL_MACHINE, sKeyBase & "\" & sProtocol, "Num_Catalog_Entries")
    
    For i = 1 To iNumNameSpace
        sFile = RegGetString(HKEY_LOCAL_MACHINE, sKeyBase & "\" & sNameSpace & "\Catalog_Entries\" & Format(i, "000000000000"), "LibraryPath")
        sFile = Replace(sFile, "%systemroot%", sWinDir, 1, 1, vbTextCompare)
        If IsRegVal404(sFile) Then
            bInteresting = True
            sResult = sResult & "NameSpace #" & CStr(i) & " is MISSING" & vbCrLf
        ElseIf InStr(1, sSafeFiles, Mid(sFile, InStrRev(sFile, "\") + 2), vbTextCompare) = 0 Or _
           bComplete Then
            bInteresting = True
            sResult = sResult & "NameSpace #" & CStr(i) & ": " & sFile & IIf(FileExists(sFile) = False, " (file MISSING)", "") & vbCrLf
        End If
    Next i
    
    For i = 1 To iNumProtocol
        sFile = RegGetFileFromBinary(HKEY_LOCAL_MACHINE, sKeyBase & "\" & sProtocol & "\Catalog_Entries\" & Format(i, "000000000000"), "PackedCatalogItem")
        sFile = Replace(sFile, "%systemroot%", sWinDir, 1, 1, vbTextCompare)
        If IsRegVal404(sFile) Then
            bInteresting = True
            sResult = sResult & "Protocol #" & CStr(i) & " is MISSING" & vbCrLf
        ElseIf InStr(1, sSafeFiles, Mid(sFile, InStrRev(sFile, "\") + 2), vbTextCompare) = 0 Or _
           bComplete = True Then
            bInteresting = True
            sResult = sResult & "Protocol #" & CStr(i) & ": " & sFile & IIf(FileExists(sFile) = False, " (file MISSING)", "") & vbCrLf
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
        sReport = sReport & vbCrLf & sResult & String(50, "-") & vbCrLf
    End If
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "EnumLSP"
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
        
        sName = String(255, 0)
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
            sName = String(255, 0)
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
    
        sName = String(255, 0)
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
            sName = String(255, 0)
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
    
    sResult = vbCrLf & sResult & vbCrLf & String(50, "-") & vbCrLf
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    ErrorMsg Err.Number, Err.Description, "EnumServices"
    RegCloseKey hKey
    RegCloseKey hKey2
End Sub

Private Sub EnumNTScripts()
    'sub applies to NT only
    If Not bIsWinNT Then
        'this is not NT,override?
        If Not (bForceAll Or bForceWinNT) Then Exit Sub
    End If
    
    Dim sDummy$, sPath$, vPaths As Variant, i%
    Dim sResult$, bInteresting As Boolean
    On Error GoTo Error:
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
            Open vPaths(i) For Input As #1
                Do
                    Line Input #1, sDummy
                    If Trim(sDummy) <> vbNullString And _
                        sDummy <> "ÿþ" Then
                        bInteresting = True
                        sResult = sResult & sDummy & vbCrLf
                    End If
                Loop Until EOF(1)
            Close #1
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
            sData = Replace(sData, "||||", vbCrLf)
            sData = Replace(sData, "\??\", "")
            sData = Replace(sData, "|!", " => ")
            sData = Replace(sData, "|" & vbCrLf, vbCrLf)
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

    sResult = vbCrLf & sResult & String(50, "-") & vbCrLf
    If bInteresting Or bComplete Then sReport = sReport & sResult
    Exit Sub
    
Error:
    RegCloseKey hKey
    ErrorMsg Err.Number, Err.Description, "EnumNTScripts"
End Sub

Private Sub EnumSSODelayLoad()
    'sub applies to all windows versions
    
    'method 53
    Dim hKey&, i&, sName$, sCLSID$, sFile$, sResult$
    Dim bInteresting As Boolean
    On Error GoTo Error:
    sResult = "[tag53]Enumerating ShellServiceObjectDelayLoad items:[/tag53]" & vbCrLf & vbCrLf
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad", 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        'key doesn't exist
        sResult = sResult & "*Registry key not found*" & vbCrLf
        GoTo EndOfSub
    End If
    
    sName = String(lEnumBufSize, 0)
    sCLSID = String(lEnumBufSize, 0)
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
        sFile = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", "")
        sFile = Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
        
        sResult = sResult & sName & ": " & sFile & vbCrLf
        
        sName = String(lEnumBufSize, 0)
        sCLSID = String(lEnumBufSize, 0)
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

    sResult = vbCrLf & sResult & vbCrLf & String(50, "-") & vbCrLf
    If bInteresting Or bComplete Then sReport = sReport & sResult
    
    Exit Sub
Error:
    RegCloseKey hKey
    ErrorMsg Err.Number, Err.Description, "EnumOSSDelayLoad"
End Sub

Private Sub ErrorMsg(iErrNum%, sErrDesc$, sSub$, Optional sArgs$)
Dim sMsg$
   'sMsg = "Unexpected error occurred!" & vbCrLf & _
           "Error #" & iErrNum & " (" & sErrDesc & _
           ") in Sub " & sSub & "(" & sArgs & ")." & vbCrLf & vbCrLf & _
           "Please send a report to www.merijn.org/contact.html, " & _
           "mentioning what you were doing, and what " & _
           "version of Windows you have." & vbCrLf & vbCrLf & _
           "This message has been copied to your clipboard."
    Clipboard.Clear
    Clipboard.SetText sMsg
    
    'MsgBox sMsg, vbExclamation
    sMsg = "Please help us improve HijackThis by reporting this error" & _
    vbCrLf & vbCrLf & "Click 'Yes' to submit" & _
    vbCrLf & vbCrLf & "Error Details: " & _
    vbCrLf & vbCrLf & "An unexpected error has occurred at procedure: " & _
    sSub$ & "(" & sArgs$ & ")" & vbCrLf & _
    "Error #" & CStr(iErrNum) & " - " & sErrDesc & _
     vbCrLf & vbCrLf & "Windows version: " & sWinVersion & vbCrLf & _
    "MSIE version: " & sMSIEVersion & vbCrLf & _
    "HijackThis version: " & App.Major & "." & App.Minor & "." & App.Revision
    
    Clipboard.Clear
    Clipboard.SetText sMsg
    If vbYes = MsgBox(sMsg, vbCritical + vbYesNo) Then
        Dim szParams As String
        Dim szCrashUrl As String
        szCrashUrl = "http://www.trendmicro.com/go/hjt/error/?"
        szParams = "function=" & sSub$
        szParams = szParams & "&params=" & sArgs$
        szParams = szParams & "&errorno=" & CStr(iErrNum)
        szParams = szParams & "&errortxt" & sErrDesc
        szParams = szParams & "&winver=" & sWinVersion
        szParams = szParams & "&iever=" & sMSIEVersion
        szParams = szParams & "&hjtver=" & App.Major & "." & App.Minor & "." & App.Revision
        szCrashUrl = szCrashUrl & URLEncode(szParams)
        
        If True = IsOnline Then
            ShellExecute 0&, "open", szCrashUrl, vbNullString, vbNullString, vbNormalFocus
        Else
            MsgBox "No Internet Connection Available"
        End If
    End If
End Sub

Public Function GetWindowsVersion$()
    Dim OVI As OSVERSIONINFO, sCSD$, sWinVer$, bla$
    On Error GoTo Error:
    
    OVI.szCSDVersion = String(128, 0)
    OVI.dwOSVersionInfoSize = Len(OVI)
    GetVersionEx OVI
    
    With OVI
        If .szCSDVersion <> "" Then
            sCSD = .szCSDVersion
            If InStr(sCSD, Chr(0)) > 0 Then sCSD = Left(sCSD, InStr(sCSD, Chr(0)) - 1)
            sCSD = Replace(sCSD, "ServicePack ", "SP", 1, 1, vbTextCompare)
            sCSD = Replace(sCSD, "Service Pack ", "SP", 1, 1, vbTextCompare)
            sCSD = Replace(sCSD, "ServicePack", "SP", 1, 1, vbTextCompare)
            sCSD = Replace(sCSD, "Service Pack", "SP", 1, 1, vbTextCompare)
            sCSD = Trim(sCSD)
        End If
        Select Case .dwPlatformId
            Case 0: GetWindowsVersion = "Detected: Windows 3.x running Win32s": Exit Function
            Case 1: bIsWin9x = True: bIsWinNT = False
            Case 2: bIsWinNT = True: bIsWin9x = False
        End Select
        If bIsWin9x Then
            If .dwMajorVersion <> 4 Then
                sWinVer = "Unknown Windows " & _
                        "(Win9x " & _
                        .dwMajorVersion & "." & _
                        Format(.dwMinorVersion, "00") & "." & _
                        Format(.dwBuildNumber And &HFFF, "0000") & _
                        sCSD & ")"
                GoTo EndOfFun
            End If
            Select Case .dwMinorVersion
                Case 0 'Windows 95 [A/B/C]
                    sWinVer = "Windows 95 " & _
                              sCSD & " " & _
                              "(Win9x 4.00." & _
                              Format(.dwBuildNumber And &HFFF, "0000") & ")"
                Case 10 'Windows 98 [Gold/SE]
                    sWinVer = "Windows 98 " & _
                              IIf(sCSD <> "", "SE ", "Gold ") & _
                              "(Win9x 4.10." & _
                              Format(.dwBuildNumber And &HFFF, "0000") & _
                              sCSD & ")"
                Case 90 'Windows Millennium Edition
                    sWinVer = "Windows ME " & _
                              "(Win9x 4.90." & _
                              Format(.dwBuildNumber And &HFFF, "0000") & _
                              sCSD & ")"
                Case Else 'WTF?
                    sWinVer = "Unknown Windows " & _
                            "(Win9x " & _
                            .dwMajorVersion & "." & _
                            Format(.dwMinorVersion, "00") & "." & _
                            Format(.dwBuildNumber And &HFFF, "0000") & _
                            sCSD & ")"
            End Select
        ElseIf bIsWinNT Then
            Select Case .dwMajorVersion
                Case 4 'Windows NT4
                    sWinVer = "Windows NT 4 " & _
                              sCSD & " " & _
                              "(WinNT 4." & _
                              Format(.dwMinorVersion, "00") & "." & _
                              Format(.dwBuildNumber And &HFFF, "0000") & ")"
                Case 5
                    Select Case .dwMinorVersion
                        Case 0 'Windows 2000
                            sWinVer = "Windows 2000 " & _
                                     sCSD & " " & _
                                     "(Windows 5." & .dwMinorVersion & "." & .dwBuildNumber & ")"
                                     '"(WinNT 5.00." & _
                                     'Format(.dwBuildNumber And &HFFF, "0000") & ")"
                                     
                        Case 1 'Windows XP
                            sWinVer = "Windows XP " & _
                                    sCSD & " " & _
                                    "(Windows 5." & .dwMinorVersion & "." & .dwBuildNumber & ")"
                                    '"(WinNT 5.01." & _
                                    'Format(.dwBuildNumber And &HFFF, "0000") & ")"
                                    
                        Case 2 'Windows 2003
                            sWinVer = "Windows 2003 " & _
                                    sCSD & " " & _
                                    "(Windows 5." & .dwMinorVersion & "." & .dwBuildNumber & ")"
                                    '"(WinNT 5.02." & _
                                    'Format(.dwBuildNumber And &HFFF, "0000") & ")"
                                    
                        Case Else 'WTF?
                            sWinVer = "Unknown Windows " & _
                                    "(WinNT " & _
                                    .dwMajorVersion & "." & _
                                    Format(.dwMinorVersion, "00") & "." & _
                                    Format(.dwBuildNumber And &HFFF, "0000") & _
                                    IIf(sCSD = "", "", " ") & sCSD & ")"
                    End Select
                Case 6
                    Select Case .dwMinorVersion
                        Case 0 'Windows Vista
                            sWinVer = "Windows Vista " & _
                                    sCSD & " " & _
                                    "(Windows 6." & .dwMinorVersion & "." & .dwBuildNumber & ")"
                                    
                        Case 1 'Windows 7
                            sWinVer = "Windows 7 " & _
                                    sCSD & " " & _
                                    "(Windows 6." & .dwMinorVersion & "." & .dwBuildNumber & ")"
                                                                        
                        Case 2 'Windows 8
                            sWinVer = "Windows 8 " & _
                                    sCSD & " " & _
                                    "(Windows 6." & .dwMinorVersion & "." & .dwBuildNumber & ")"
                                    
                        Case 3 'Windows 2008
                            sWinVer = "Windows 2008 " & _
                                    sCSD & " " & _
                                    "(Windows 6." & .dwMinorVersion & "." & .dwBuildNumber & ")"
                                    
                        Case 4 'Windows 2012
                            sWinVer = "Windows 2012 " & _
                                    sCSD & " " & _
                                    "(Windows 6." & .dwMinorVersion & "." & .dwBuildNumber & ")"
                        Case Else
                            sWinVer = "Unknown Windows " & _
                                    "(WinNT " & _
                                    .dwMajorVersion & "." & _
                                    Format(.dwMinorVersion, "00") & "." & _
                                    Format(.dwBuildNumber And &HFFF, "0000") & _
                                    IIf(sCSD = "", "", " ") & sCSD & ")"
                    End Select
                Case Else 'WTF?
                    sWinVer = "Unknown Windows " & _
                            "(WinNT " & _
                            .dwMajorVersion & "." & _
                            Format(.dwMinorVersion, "00") & "." & _
                            Format(.dwBuildNumber And &HFFF, "0000") & _
                            IIf(sCSD = "", "", " ") & sCSD & ")"
            End Select
        End If
    End With
EndOfFun:
    GetWindowsVersion = sWinVer
    Exit Function
    
Error:
    ErrorMsg Err.Number, Err.Description, "GetWindowsVersion"
End Function


Public Function GetMSIEVersion$()
    Dim sMSIEPath$, sMSIEVer$, sMSIEHotfixes$, sMSIEFriendlyVer$
    On Error GoTo Error:
    sMSIEPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE", "")
    If sMSIEPath = "" Then GoTo EndOfFun:
    If FileExists(sMSIEPath) = False Then GoTo EndOfFun:
    
    Dim hData&, lDataLen&, uBuf() As Byte, uVFFI As VS_FIXEDFILEINFO
    lDataLen = GetFileVersionInfoSize(sMSIEPath, ByVal 0)
    If lDataLen = 0 Then
        GoTo EndOfFun:
    End If
        
    ReDim uBuf(0 To lDataLen - 1)
    'get handle to file props
    GetFileVersionInfo sMSIEPath, 0, lDataLen, uBuf(0)
    VerQueryValue uBuf(0), "\", hData, lDataLen
    CopyMemory uVFFI, ByVal hData, Len(uVFFI)
    With uVFFI
        sMSIEVer = Format(.dwFileVersionMSh, "0") & "." & _
                   Format(.dwFileVersionMSl, "00") & "." & _
                   Format(.dwProductVersionLSh, "0000") & "." & _
                   Format(.dwProductVersionLSl, "0000")
    End With
    If sMSIEVer = "0.00.0000.0000" Then GoTo EndOfFun:
    
    sMSIEFriendlyVer = Left(sMSIEVer, 4)
    
    sMSIEHotfixes = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MinorVersion")
    If sMSIEHotfixes = vbNullString Then GoTo EndOfFun:
    If InStr(1, sMSIEHotfixes, "SP5", vbTextCompare) > 0 Then
        sMSIEFriendlyVer = sMSIEFriendlyVer & " SP5"
    Else
        If InStr(1, sMSIEHotfixes, "SP4", vbTextCompare) > 0 Then
            sMSIEFriendlyVer = sMSIEFriendlyVer & " SP4"
        Else
            If InStr(1, sMSIEHotfixes, "SP3", vbTextCompare) > 0 Then
                sMSIEFriendlyVer = sMSIEFriendlyVer & " SP3"
            Else
                
                If InStr(1, sMSIEHotfixes, "SP2", vbTextCompare) > 0 Then
                    sMSIEFriendlyVer = sMSIEFriendlyVer & " SP2"
                Else
                    If InStr(1, sMSIEHotfixes, "SP1", vbTextCompare) > 0 Then
                        sMSIEFriendlyVer = sMSIEFriendlyVer & " SP1"
                    End If
                End If
            End If
        End If
    End If
    
EndOfFun:
    If lDataLen > 0 And Left(sMSIEFriendlyVer, 1) <> "0" Then
        GetMSIEVersion = "Internet Explorer v" & sMSIEFriendlyVer & " (" & sMSIEVer & ")"
    Else
        GetMSIEVersion = "Unable to get Internet Explorer version!"
    End If
    Exit Function
    
Error:
    ErrorMsg Err.Number, Err.Description, "GetMSIEVersion"
End Function

Public Function GetLongPath$(sFile$)
    'sub applies to NT only, checked in ListRunningProcesses()
    'attempt to find location of given file
    'On Error GoTo Error:
    On Error Resume Next
    
    'evading parasites that put html or garbled data in
    'O4 autorun entries :P
    If InStr(sFile, "<") > 0 Or InStr(sFile, ">") > 0 Or _
       InStr(sFile, "|") > 0 Or InStr(sFile, "*") > 0 Or _
       InStr(sFile, "?") > 0 Or InStr(sFile, "/") > 0 Then
        GetLongPath = sFile
        Exit Function
    End If
    
    If sFile = "[System Process]" Or sFile = "System" Then
        GetLongPath = sFile
        Exit Function
    End If
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        GetLongPath = sFile
        Exit Function
    End If
    
    If InStr(sFile, "\") > 0 Then
        'filename is already full path
        GetLongPath = sFile
        Exit Function
    End If
    
    'check if file is self
    If LCase(sFile) = LCase(App.EXEName & ".exe") Then
        GetLongPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & sFile
        Exit Function
    End If
    
    Dim hKey, sData$, i%, sDummy$, sProgramFiles$
    'check App Paths regkey
    sData = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" & sFile, "")
    If Not IsRegVal404(sData) And sData <> vbNullString Then
        GetLongPath = sData
        Exit Function
    End If
    
    'check own folder
    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "" Then
        GetLongPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & sFile
        Exit Function
    End If
    
    'check windir
    If Dir(sWinDir & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "" Then
        GetLongPath = sWinDir & "\" & sFile
        Exit Function
    End If
    
    'check windir\system
    If Dir(sWinDir & "\system\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "" Then
        GetLongPath = sWinDir & "\system32\" & sFile
        Exit Function
    End If
    
    'check windir\system23
    If Dir(sWinDir & "\system32\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "" Then
        GetLongPath = sWinDir & "\system32\" & sFile
        Exit Function
    End If
    
    If InStr(sFile, ".") > 0 Then
        'prog.exe -> prog
        sDummy = Left(sFile, InStr(sFile, ".") - 1)
        sProgramFiles = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
        
        'check x:\program files\prog\prog.exe
        If Dir(sProgramFiles & "\" & sDummy & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "" Then
            GetLongPath = sProgramFiles & "\" & sDummy & "\" & sFile
            Exit Function
        End If
        
        'check c:\prog\prog.exe
        If Dir("C:\" & sDummy & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = "C:\" & sDummy & "\" & sFile
            Exit Function
        End If
        
        'check x:\program files\prog32\prog.exe
        If Dir(sProgramFiles & "\" & sDummy & "32\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = sProgramFiles & "\" & sDummy & "32\" & sFile
            Exit Function
        End If
        If Dir(sProgramFiles & "\" & sDummy & "16\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = sProgramFiles & "\" & sDummy & "16\" & sFile
            Exit Function
        End If
        
        'check c:\prog32\prog.exe
        If Dir("C:\" & sDummy & "32\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = "C:\" & sDummy & "32\" & sFile
            Exit Function
        End If
        If Dir("C:\" & sDummy & "16\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            GetLongPath = "C:\" & sDummy & "16\" & sFile
            Exit Function
        End If
        
        If Right(sDummy, 2) = "32" Or Right(sDummy, 2) = "16" Then
            'asssuming sFile is prog32.exe,
            'check x:\program files\prog\prog32.exe
            sDummy = Left(sDummy, Len(sDummy) - 2)
            If Dir(sProgramFiles & "\" & sDummy & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
                GetLongPath = sProgramFiles & "\" & sDummy & "\" & sFile
                Exit Function
            End If
            
            'check c:\prog\prog32.exe
            If Dir("C:\" & sDummy & "\" & sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
                GetLongPath = "C:\" & sDummy & "\" & sFile
                Exit Function
            End If
        End If
    End If
    
    'can't find it!
    GetLongPath = "?:\?\" & sFile
    Exit Function
    
Error:
    RegCloseKey hKey
    ErrorMsg Err.Number, Err.Description, "GetLongPath", sFile
End Function

Public Function GetFileFromShortCut$(ByVal sLink$)
    Dim sLnk$, vLnk As Variant, i&, sFile$
    On Error Resume Next
    If Right(LCase(sLink), 4) <> ".lnk" And _
       Right(LCase(sLink), 4) <> ".pif" Then Exit Function
    Open sLink For Input As #1
    Close #1
    If Err Then Exit Function
    On Error GoTo Error:
    
    'use binary access or input will be broken by Chr(0)
    Open sLink For Binary As #1
        sLnk = Input(FileLen(sLink), #1)
    Close #1
    
    'split into array by Chr(0) and check for path + file
    vLnk = Split(sLnk, Chr(0))
    For i = 0 To UBound(vLnk)
        If vLnk(i) <> vbNullString And StringIsAlphaNumeric(CStr(vLnk(i))) Then
            If InStr(vLnk(i), "\") > 0 And _
               InStr(vLnk(i), ".") > 0 And _
               InStr(vLnk(i), ":") > 0 And _
               InStr(vLnk(i), "\\") = 0 And _
               InStr(vLnk(i), "..") = 0 Then
                'found something!
                'most of the time only one REAL path + file
                'is in shortcut file, so this should work.
                sFile = CStr(vLnk(i))
                Exit For
            End If
        End If
    Next i
    
    'sometimes the path to the icon file is found when
    'the actual file path is not
    If Right(LCase(sFile), 4) = ".ico" Then sFile = ""
    
    If sFile = vbNullString Then
        'try alternate method, where drive and path have
        'been separated
        For i = 0 To UBound(vLnk)
            If vLnk(i) <> vbNullString And StringIsAlphaNumeric(CStr(vLnk(i))) Then
                If Right(vLnk(i), 2) = ":\" And _
                   Len(vLnk(i)) = 3 Then
                    sFile = vLnk(i)
                    Exit For
                End If
            End If
        Next i
        For i = 0 To UBound(vLnk)
            If vLnk(i) <> vbNullString And StringIsAlphaNumeric(CStr(vLnk(i))) Then
                If InStr(vLnk(i), "\") > 0 And _
                   InStr(vLnk(i), "\\") = 0 And _
                   InStr(vLnk(i), ".") > 0 Then
                    sFile = sFile & vLnk(i)
                    Exit For
                End If
            End If
        Next i
    End If
    
    'sometimes the path to the icon file is found when
    'the actual file path is not
    If Right(LCase(sFile), 4) = ".ico" Then sFile = ""
    
    'check for period at the right position
    If Left(Right(sFile, 4), 1) <> "." Then sFile = ""
    
    If sFile <> vbNullString Then
        GetFileFromShortCut = " = " & sFile
    Else
        GetFileFromShortCut = " = ?"
    End If
    Exit Function
    
Error:
    Close
    ErrorMsg Err.Number, Err.Description, "GetFileFromShortCut", sLink
End Function

Public Function StringIsAlphaNumeric(s$) As Boolean
    Dim i&
    On Error GoTo Error:
    
    StringIsAlphaNumeric = True
    For i = 1 To Len(s)
        Select Case Asc(Mid(s, i, 1))
            'Case 48 To 57 ' 0 - 9
            'Case 65 To 90 ' A - Z
            'Case 97 To 122 'a - z
            'Case Asc(""""), Asc(":"), Asc("\"), Asc("-")
            'Case Asc("."), Asc(" "), Asc("!"), Asc("_")
            'Case Asc("'"), Asc("("), Asc(")")
            Case 32 To 126
            Case Else
                StringIsAlphaNumeric = False
                Exit For
        End Select
    Next i
    Exit Function
    
Error:
    ErrorMsg Err.Number, Err.Description, "StringIsAlphaNumeric"
End Function

Private Function TrimNull$(s$, Optional bFullTrim As Boolean = False)
    On Error GoTo Error:
    If s = vbNullString Or InStr(s, Chr(0)) = 0 Then
        TrimNull = s
        Exit Function
    End If
    If bFullTrim Then
        Do Until Left(s, 1) <> Chr(0)
            s = Mid(s, 2)
        Loop
        Do Until Right(s, 1) <> Chr(0)
            s = Left(s, Len(s) - 1)
        Loop
    Else
        s = Left(s, InStr(s, Chr(0)) - 1)
    End If
    TrimNull = s
    Exit Function
    
Error:
    ErrorMsg Err.Number, Err.Description, "TrimNull"
End Function

Private Function RegGetString$(lHive&, sKey$, sVal$)
    Dim hKey&, sData$
    On Error GoTo Error:
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        sData = String(255, 0)
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
    
Error:
    Dim sRoot
    RegCloseKey hKey
    Select Case lHive
        Case HKEY_CURRENT_USER: sRoot = "HKCU"
        Case HKEY_LOCAL_MACHINE: sRoot = "HKLM"
        Case Else: sRoot = "HK.."
    End Select
    ErrorMsg Err.Number, Err.Description, "RegGetString", sRoot & ", " & sKey & ", " & sVal
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
    On Error GoTo Error:
    
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        sData = String(255, 0)
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
    
Error:
    Dim sRoot$
    RegCloseKey hKey
    Select Case lHive
        Case HKEY_CURRENT_USER: sRoot = "HKCU"
        Case HKEY_LOCAL_MACHINE: sRoot = "HKLM"
        Case Else: sRoot = "HK.."
    End Select
    ErrorMsg Err.Number, Err.Description, "RegValueExists", sRoot & ", " & sKey & ", " & sValue
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
    Dim hKey&, i%, uData() As Byte, sFile$
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

Private Function FileExists(sFile$) As Boolean
    On Error Resume Next
    If bIsWin9x Then
        FileExists = CBool(SHFileExists(sFile))
    Else
        FileExists = IIf(Dir(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, True, False)
    End If
End Function

Private Function IniGetString$(sFile$, sSection$, sValue$, bNT As Boolean)
    Dim sIniFile$, sLine$, sDummy$, sRet$
    On Error GoTo Error:
    
    If Not bNT Then
        'win9x method - check ini FILE
        
        If Dir(sWinDir & "\" & sFile) <> vbNullString Then
            sIniFile = sWinDir & "\" & sFile
        Else
            If Dir(sWinDir & "\system\" & sFile) <> vbNullString Then
                sIniFile = sWinDir & "\system\" & sFile
            Else
                If Dir(sWinDir & "\system32\" & sFile) <> vbNullString Then
                    sIniFile = sWinDir & "\system32\" & sFile
                Else
                    IniGetString = "*INI file not found*"
                    Exit Function
                End If
            End If
        End If
        
        If FileLen(sIniFile) = 0 Then
            IniGetString = "*INI section not found*"
            Exit Function
        End If
        Open sIniFile For Input As #1
            Do
                Line Input #1, sLine
            Loop Until EOF(1) Or LCase(sLine) = LCase("[" & sSection & "]")
            If EOF(1) Or LCase(sLine) <> LCase("[" & sSection & "]") Then
                IniGetString = "*INI section not found*"
                Close #1
                Exit Function
            End If
            'found the [section]
            Do
                Line Input #1, sLine
                If Len(sLine) > Len(sValue) Then
                    If LCase(Left(sLine, Len(sValue))) = LCase(sValue) Then
                        'found the setting=
                        IniGetString = Mid(sLine, Len(sValue) + 2)
                        Exit Do
                    End If
                End If
            Loop Until EOF(1) Or sLine = vbNullString
        Close #1
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
        If Left(sDummy, 1) = "!" Then sDummy = Mid(sDummy, 2)
        If Left(sDummy, 1) = "#" Then sDummy = Mid(sDummy, 2)
        
        'get actual setting
        Select Case Left(sDummy, 3)
            Case "USR"
                sDummy = RegGetString(HKEY_CURRENT_USER, Mid(sDummy, 5), sValue)
                sRet = "[HKCU\" & Mid(sDummy, 5) & "]"
            Case "SYS"
                sDummy = RegGetString(HKEY_LOCAL_MACHINE, Mid(sDummy, 5), sValue)
                sRet = "[HKLM\" & Mid(sDummy, 5) & "]"
            Case Else
                IniGetString = ""
                Exit Function
        End Select
        
        IniGetString = sRet & vbCrLf & Replace(sDummy, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
    End If
    Exit Function

Error:
    Close #1
    ErrorMsg Err.Number, Err.Description, "INIGetString", sFile & ", " & sSection & ", " & sValue & ", " & bNT
End Function

Private Sub HTMLize()
    Dim sHTML$, sHeader$, sFooter$, sIndex$, i%
    Dim sReplace$(1), sSections$
    
    If bHTML Then
        'fix < and > occurrances not being HTML
        sReport = Replace(sReport, "<", "&lt;")
        sReport = Replace(sReport, ">", "&gt;")
        
        'replace -- and == bars with <HR> bars
        sReport = Replace(sReport, String(50, "-"), "<CENTER><HR WIDTH=""80%""></CENTER>")
        
        'expand [tag]s
        sSections = ""
        For i = 1 To 99
            If InStr(sReport, "[tag" & CStr(i) & "]") > 0 Then
                sReplace(0) = "<A NAME=""" & CStr(i) & """><B>"
                sReplace(1) = "</B></A>"
                sReport = Replace(sReport, "[tag" & CStr(i) & "]", sReplace(0))
                sReport = Replace(sReport, "[/tag" & CStr(i) & "]", sReplace(1))
                
                sSections = sSections & "1"
            Else
                sSections = sSections & "0"
            End If
        Next i
    Else
        'remove [tag]s
        For i = 1 To 99
            If InStr(sReport, "[tag" & CStr(i) & "]") > 0 Then
                sReport = Replace(sReport, "[tag" & CStr(i) & "]", "")
                sReport = Replace(sReport, "[/tag" & CStr(i) & "]", "")
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
        If Mid(sSections, 1, 1) = "1" Then sIndex = sIndex & "<A HREF=""#1"">Running processes</A><BR>" & vbCrLf
        If Mid(sSections, 2, 1) = "1" Then sIndex = sIndex & "<A HREF=""#2"">Autostart folders</A><BR>" & vbCrLf
        If Mid(sSections, 3, 1) = "1" Then sIndex = sIndex & "<A HREF=""#3"">Windows NT UserInit</A><BR>" & vbCrLf
        If Mid(sSections, 4, 1) = "1" Then sIndex = sIndex & "<A HREF=""#4"">Autorun key HKLM\..\Run</A><BR>" & vbCrLf
        If Mid(sSections, 5, 1) = "1" Then sIndex = sIndex & "<A HREF=""#5"">Autorun key HKLM\..\RunOnce</A><BR>" & vbCrLf
        If Mid(sSections, 6, 1) = "1" Then sIndex = sIndex & "<A HREF=""#6"">Autorun key HKLM\..\RunOnceEx</A><BR>" & vbCrLf
        If Mid(sSections, 7, 1) = "1" Then sIndex = sIndex & "<A HREF=""#7"">Autorun key HKLM\..\RunServices</A><BR>" & vbCrLf
        If Mid(sSections, 8, 1) = "1" Then sIndex = sIndex & "<A HREF=""#8"">Autorun key HKLM\..\RunServicesOnce</A><BR>" & vbCrLf
        If Mid(sSections, 9, 1) = "1" Then sIndex = sIndex & "<A HREF=""#9"">Autorun key HKCU\..\Run</A><BR>" & vbCrLf
        If Mid(sSections, 10, 1) = "1" Then sIndex = sIndex & "<A HREF=""#10"">Autorun key HKCU\..\RunOnce</A><BR>" & vbCrLf
        If Mid(sSections, 11, 1) = "1" Then sIndex = sIndex & "<A HREF=""#11"">Autorun key HKCU\..\RunOnceEx</A><BR>" & vbCrLf
        If Mid(sSections, 12, 1) = "1" Then sIndex = sIndex & "<A HREF=""#12"">Autorun key HKCU\..\RunServices</A><BR>" & vbCrLf
        If Mid(sSections, 13, 1) = "1" Then sIndex = sIndex & "<A HREF=""#13"">Autorun key HKCU\..\RunServicesOnce</A><BR>" & vbCrLf
        If Mid(sSections, 14, 1) = "1" Then sIndex = sIndex & "<A HREF=""#14"">Autorun key HKLM\..\Run (NT)</A><BR>" & vbCrLf
        If Mid(sSections, 15, 1) = "1" Then sIndex = sIndex & "<A HREF=""#15"">Autorun key HKCU\..\Run (NT)</A><BR>" & vbCrLf
        If Mid(sSections, 16, 1) = "1" Then sIndex = sIndex & "<A HREF=""#16"">Autorun subkeys HKLM\..\Run\*</A><BR>" & vbCrLf
        If Mid(sSections, 17, 1) = "1" Then sIndex = sIndex & "<A HREF=""#17"">Autorun subkeys HKLM\..\RunOnce\*</A><BR>" & vbCrLf
        If Mid(sSections, 18, 1) = "1" Then sIndex = sIndex & "<A HREF=""#18"">Autorun subkeys HKLM\..\RunOnceEx\*</A><BR>" & vbCrLf
        If Mid(sSections, 19, 1) = "1" Then sIndex = sIndex & "<A HREF=""#19"">Autorun subkeys HKLM\..\RunServices\*</A><BR>" & vbCrLf
        If Mid(sSections, 20, 1) = "1" Then sIndex = sIndex & "<A HREF=""#20"">Autorun subkeys HKLM\..\RunServicesOnce\*</A><BR>" & vbCrLf
        If Mid(sSections, 21, 1) = "1" Then sIndex = sIndex & "<A HREF=""#21"">Autorun subkeys HKCU\..\Run\*</A><BR>" & vbCrLf
        If Mid(sSections, 22, 1) = "1" Then sIndex = sIndex & "<A HREF=""#22"">Autorun subkeys HKCU\..\RunOnce\*</A><BR>" & vbCrLf
        If Mid(sSections, 23, 1) = "1" Then sIndex = sIndex & "<A HREF=""#23"">Autorun subkeys HKCU\..\RunOnceEx\*</A><BR>" & vbCrLf
        If Mid(sSections, 24, 1) = "1" Then sIndex = sIndex & "<A HREF=""#24"">Autorun subkeys HKCU\..\RunServices\*</A><BR>" & vbCrLf
        If Mid(sSections, 25, 1) = "1" Then sIndex = sIndex & "<A HREF=""#25"">Autorun subkeys HKCU\..\RunServicesOnce\*</A><BR>" & vbCrLf
        If Mid(sSections, 26, 1) = "1" Then sIndex = sIndex & "<A HREF=""#26"">Autorun subkeys HKLM\..\Run\* (NT)</A><BR>" & vbCrLf
        If Mid(sSections, 27, 1) = "1" Then sIndex = sIndex & "<A HREF=""#27"">Autorun subkeys HKCU\..\Run\* (NT)</A><BR>" & vbCrLf
        If Mid(sSections, 28, 1) = "1" Then sIndex = sIndex & "<A HREF=""#28"">Class .EXE</A><BR>" & vbCrLf
        If Mid(sSections, 29, 1) = "1" Then sIndex = sIndex & "<A HREF=""#29"">Class .COM</A><BR>" & vbCrLf
        If Mid(sSections, 30, 1) = "1" Then sIndex = sIndex & "<A HREF=""#30"">Class .BAT</A><BR>" & vbCrLf
        If Mid(sSections, 31, 1) = "1" Then sIndex = sIndex & "<A HREF=""#31"">Class .PIF</A><BR>" & vbCrLf
        If Mid(sSections, 32, 1) = "1" Then sIndex = sIndex & "<A HREF=""#32"">Class .SCR</A><BR>" & vbCrLf
        If Mid(sSections, 33, 1) = "1" Then sIndex = sIndex & "<A HREF=""#33"">Class .HTA</A><BR>" & vbCrLf
        If Mid(sSections, 34, 1) = "1" Then sIndex = sIndex & "<A HREF=""#34"">Class .TXT</A><BR>" & vbCrLf
        
        If Mid(sSections, 35, 1) = "1" Then sIndex = sIndex & "<A HREF=""#35"">Active Setup Stub Paths</A><BR>" & vbCrLf
        If Mid(sSections, 36, 1) = "1" Then sIndex = sIndex & "<A HREF=""#36"">ICQ Agent</A><BR>" & vbCrLf
        If Mid(sSections, 37, 1) = "1" Then sIndex = sIndex & "<A HREF=""#37"">Load/Run keys from WIN.INI</A><BR>" & vbCrLf
        If Mid(sSections, 38, 1) = "1" Then sIndex = sIndex & "<A HREF=""#38"">Shell/SCRNSAVE.EXE keys from SYSTEM.INI</A><BR>" & vbCrLf
        If Mid(sSections, 39, 1) = "1" Then sIndex = sIndex & "<A HREF=""#39"">Explorer check</A><BR>" & vbCrLf
        If Mid(sSections, 40, 1) = "1" Then sIndex = sIndex & "<A HREF=""#40"">Wininit.ini</A><BR>" & vbCrLf
        If Mid(sSections, 41, 1) = "1" Then sIndex = sIndex & "<A HREF=""#41"">Wininit.bak</A><BR>" & vbCrLf
        If Mid(sSections, 42, 1) = "1" Then sIndex = sIndex & "<A HREF=""#42"">C:\Autoexec.bat</A><BR>" & vbCrLf
        If Mid(sSections, 43, 1) = "1" Then sIndex = sIndex & "<A HREF=""#43"">C:\Config.sys</A><BR>" & vbCrLf
        If Mid(sSections, 44, 1) = "1" Then sIndex = sIndex & "<A HREF=""#44"">" & sWinDir & "\Winstart.bat</A><BR>" & vbCrLf
        If Mid(sSections, 45, 1) = "1" Then sIndex = sIndex & "<A HREF=""#45"">" & sWinDir & "\Dosstart.bat</A><BR>" & vbCrLf
        If Mid(sSections, 46, 1) = "1" Then sIndex = sIndex & "<A HREF=""#46"">Superhidden extensions</A><BR>" & vbCrLf
        If Mid(sSections, 47, 1) = "1" Then sIndex = sIndex & "<A HREF=""#47"">Regedit.exe check</A><BR>" & vbCrLf
        If Mid(sSections, 48, 1) = "1" Then sIndex = sIndex & "<A HREF=""#48"">BHO list</A><BR>" & vbCrLf
        If Mid(sSections, 49, 1) = "1" Then sIndex = sIndex & "<A HREF=""#49"">Task Scheduler jobs</A><BR>" & vbCrLf
        If Mid(sSections, 50, 1) = "1" Then sIndex = sIndex & "<A HREF=""#50"">Download Program Files</A><BR>" & vbCrLf
        If Mid(sSections, 51, 1) = "1" Then sIndex = sIndex & "<A HREF=""#51"">Winsock LSP files</A><BR>" & vbCrLf
        If Mid(sSections, 52, 1) = "1" Then sIndex = sIndex & "<A HREF=""#52"">Windows NT Services</A><BR>" & vbCrLf
        If Mid(sSections, 53, 1) = "1" Then sIndex = sIndex & "<A HREF=""#53"">Windows NT Logon/logoff scripts</A><BR>" & vbCrLf
        If Mid(sSections, 54, 1) = "1" Then sIndex = sIndex & "<A HREF=""#54"">ShellServiceObjectDelayLoad regkey</A><BR>" & vbCrLf
        If Mid(sSections, 55, 1) = "1" Then sIndex = sIndex & "<A HREF=""#55"">Policies Explorer Run key (global)</A><BR>" & vbCrLf
        If Mid(sSections, 56, 1) = "1" Then sIndex = sIndex & "<A HREF=""#56"">Policies Explorer Run key (current user)</A><BR>" & vbCrLf
        
        If Mid(sSections, 57, 1) = "1" Then sIndex = sIndex & "<A HREF=""#55""> </A><BR>" & vbCrLf
        
        sIndex = sIndex & "</BLOCKQUOTE></BLOCKQUOTE>" & vbCrLf
        'write footer
        sFooter = vbCrLf & vbCrLf & "</FONT></PRE></BODY></HTML>"
        
        sReport = Replace(sReport, String(50, "="), sIndex & "<CENTER><HR WIDTH=""80%"" SIZE=5""></CENTER>" & "<PRE><FONT FACE=""Fixedsys"">" & vbCrLf & vbCrLf)
        sReport = sHeader & sReport & sFooter
    End If
End Sub

Public Function BuildPath$(sPath$, sFile$)
    BuildPath = sPath & IIf(Right(sPath, 1) = "\", vbNullString, "\") & sFile
End Function

