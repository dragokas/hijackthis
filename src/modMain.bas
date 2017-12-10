Attribute VB_Name = "modMain"
'
' Core check / Fix Engine
'
' (part 1: R0-R4 / F0-F1 / O1 - O24)
' (part 2: see modMain_2.bas)

' (c) Fork copyrights:
'
' R4 by Alex Dragokas
' O1 hosts.ics / DNSApi hijackers by Alex Dragokas
' O4 MSconfig and full rework by Alex Dragokas
' O7 IPSec / TroubleShoot by Alex Dragokas
' O17 DHCP DNS by Alex Dragokas
' O21 ShellIconOverlayIdentifiers by Alex Dragokas
' O22 Tasks (Vista+) by Alex Dragokas

'
' List of all sections:
'

'R0 - Changed Registry value (MSIE)
'R1 - Created Registry value
'R2 - Created Registry key
'R3 - Created extra value in regkey where only one should be
'R4 - IE SearchScopes, DefaultScope
'F0 - Changed inifile value (system.ini)
'F1 - Created inifile value (win.ini)
'N1 (removed in 2.0.7) - Changed NS4.x homepage
'N2 (removed in 2.0.7) - Changed NS6 homepage
'N3 (removed in 2.0.7) - Changed NS7 homepage/searchpage
'N4 (removed in 2.0.7) - Changed Moz homepage/searchpage
'O1 - Hosts / hosts.ics / DNSApi hijackers
'O2 - BHO
'O3 - IE Toolbar
'O4 - Reg. autorun entry / msconfig disabled items
'O5 - Control.ini IE Options block
'O6 - IE Policy: IE Options/Control Panel block
'O7 - Policies: Regedit block / IPSec; O7 - TroubleShoot: system settings, that lead to OS malfunction
'O8 - IE Context menuitem
'O9 - IE Tools menuitem/button
'O10 - Winsock hijack
'O11 - IE Advanced Options group
'O12 - IE Plugin
'O13 - IE DefaultPrefix hijack
'O14 - IERESET.INF hijack
'O15 - Trusted Zone autoadd
'O16 - Downloaded Program Files
'O17 - Domain hijacks / DHCP DNS
'O18 - Protocol & Filter enum
'O19 - User style sheet hijack
'O20 - AppInit_DLLs registry value + Winlogon Notify subkeys
'O21 - ShellServiceObjectDelayLoad / ShellIconOverlayIdentifiers enum
'O22 - SharedTaskScheduler enum
'O23 - Windows Services
'O24 - Active desktop components
'O25 - Windows Management Instrumentation (WMI) event consumers
'O26 - Image File Execution Options (IFEO)

'If you added new section, you also must:
' - add prefix to Backup module (2 times)
' - and prefix to Fix module
' - append progressbar max value - var. g_HJT_Items_Count
' - add procedure CheckOxxItem to 'StartScan'
' - check max.value in 'for' - function "SortSectionsOfResultList"
' - add translation strings: after # 31, 261, 435

'Next possible methods:
'* SearchAccurates 'URL' method in a InitPropertyBag (??)
'* HKLM\..\CurrentVersion\ModuleUsage
'* HKLM\..\CurrentVersion\Explorer\ShellExecuteHooks (eudora)
'* HKLM\..\Internet Explorer\SafeSites (searchaccurate)

Option Explicit

Public Enum ENUM_REG_HIVE_FIX
    HKCR_FIX = 1
    HKCU_FIX = 2
    HKLM_FIX = 4
    HKU_FIX = 8
End Enum
#If False Then
    Dim HKCR_FIX, HKCU_FIX, HKLM_FIX, HKU_FIX
#End If

Public Enum ENUM_REG_REDIRECTION
    REG_REDIRECTED = -1
    REG_NOTREDIRECTED = 0
    REG_REDIRECTION_BOTH = 1
    [_REG_REDIRECTION_NOT_DEFINED] = -2
End Enum
#If False Then
    Dim REG_REDIRECTED, REG_NOTREDIRECTED, REG_REDIRECTION_BOTH
#End If

Public Enum ENUM_REG_VALUE_TYPE_RESTORE
    REG_RESTORE_SAME = -1&
    REG_RESTORE_SZ = 1&
    REG_RESTORE_EXPAND_SZ = 2&
    'REG_RESTORE_BINARY = 3&
    REG_RESTORE_DWORD = 4&
    'REG_RESTORE_LINK = 6&
    REG_RESTORE_MULTI_SZ = 7&
End Enum
#If False Then
    Dim REG_RESTORE_SAME, REG_RESTORE_SZ, REG_RESTORE_EXPAND_SZ, REG_RESTORE_DWORD
#End If

Public Enum ENUM_CURE_BASED
    FILE_BASED = 1          ' if need to cure .File()
    REGISTRY_BASED = 2      ' if need to cure .Reg()
    INI_BASED = 4           ' if need to cure ini-file in .reg()
    PROCESS_BASED = 8       ' if need to kill/freeze a process
    SERVICE_BASED = 16      ' if need to delete/restore service .ServiceName
    CUSTOM_BASED = 32       ' individual rule, based on section name
End Enum

#If False Then
    Dim FILE_BASED, REGISTRY_BASED, INI_BASED, PROCESS_BASED, SERVICE_BASED, CUSTOM_BASED
#End If

Public Enum ENUM_REG_ACTION_BASED
    REMOVE_KEY = 1
    REMOVE_VALUE = 2
    RESTORE_VALUE = 4
    RESTORE_VALUE_INI = 8
    REMOVE_VALUE_INI = 16 'TODO
    REPLACE_VALUE = 32
    REMOVE_VALUE_IF_EMPTY = 64
    REMOVE_KEY_IF_NO_VALUES = 128
    TRIM_VALUE = 256
    BACKUP_KEY = 512
    BACKUP_VALUE = 1024
End Enum
#If False Then
    Dim REMOVE_KEY, REMOVE_VALUE, RESTORE_VALUE, RESTORE_VALUE_INI
#End If

Public Enum ENUM_FILE_ACTION_BASED
    REMOVE_FILE = 1
    REMOVE_FOLDER = 2  'not used yet
    RESTORE_FILE = 4   'not used yet
    RESTORE_FILE_SFC = 8
    UNREG_DLL = 16
    BACKUP_FILE = 32
End Enum
#If False Then
    Dim REMOVE_FILE, REMOVE_FOLDER, RESTORE_FILE
#End If

Public Enum ENUM_PROCESS_ACTION_BASED
    KILL_PROCESS = 1
    FREEZE_PROCESS = 2
End Enum
#If False Then
    Dim KILL_PROCESS, FREEZE_PROCESS
#End If

Public Enum ENUM_SERVICE_ACTION_BASED
    DELETE_SERVICE = 1
    RESTORE_SERVICE = 2 ' not yet implemented
End Enum
#If False Then
    Dim DELETE_SERVICE, RESTORE_SERVICE
#End If

Private Type O25_Info
    sScriptFile     As String
    '-------------------------
    sTimerClassName As String
    TimerID         As String
    '-------------------------
    ConsumerName    As String
    ConsumerNameSpace As String
    ConsumerPath    As String
    '-------------------------
    FilterName      As String
    FilterNameSpace As String
    FilterPath      As String
End Type

Public Type FIX_REG_KEY
    IniFile         As String
    Hive            As ENUM_REG_HIVE
    Key             As String
    Param           As String
    ParamType       As ENUM_REG_VALUE_TYPE_RESTORE
    DefaultData     As Variant
    Redirected      As Boolean  'is key under Wow64
    ActionType      As ENUM_REG_ACTION_BASED
    ReplaceDataWhat As String
    ReplaceDataInto As String
    TrimDelimiter   As String
End Type

Private Type FIX_FILE
    Path            As String
    GoodFile        As String
    ActionType      As ENUM_FILE_ACTION_BASED
End Type

Private Type FIX_PROCESS
    Path            As String
    ActionType      As ENUM_PROCESS_ACTION_BASED
End Type

Private Type FIX_SERVICE
    ImagePath       As String
    DllPath         As String
    ServiceName     As String
    ServiceDisplay  As String
    ActionType      As ENUM_SERVICE_ACTION_BASED
End Type

Public Type SCAN_RESULT
    HitLineW        As String
    HitLineA        As String
    Section         As String
    Alias           As String
    Reg()           As FIX_REG_KEY
    File()          As FIX_FILE
    Process()       As FIX_PROCESS
    Service()       As FIX_SERVICE
    CureType        As ENUM_CURE_BASED
    O25             As O25_Info
    NoNeedBackup    As Boolean          'if no backup required / possible
End Type

Type TYPE_PERFORMANCE
    StartTime       As Long ' time the program started its working
    EndTime         As Long ' time the program finished its working
    MAX_TimeOut     As Long ' maximum time (mm) allowed for program to run the scanning
End Type

Private Type TASK_WHITELIST_ENTRY
    OSver       As Single
    Path        As String
    RunObj      As String
    Args        As String
End Type

Private Type DICTIONARIES
    TaskWL_ID  As clsTrickHashTable
End Type

Private Type IPSEC_FILTER_RECORD    '36 bytes
    Mirrored       As Byte
    Unknown1(2)    As Byte
    IP1(3)         As Byte
    IPTypeFlag1    As Long
    IP2(3)         As Byte
    IPTypeFlag2    As Long
    Unknown2       As Long
    ProtocolType   As Byte
    Unknown3(2)    As Byte
    PortNum1       As Integer
    PortNum2       As Integer
    Unknown4       As Byte
    DynPacketType  As Byte
    Unknown5(1)    As Byte
End Type

Private Type MY_PROC_LOG
    ProcName    As String
    Number      As Long
    IsMicrosoft As Boolean
    EDS_issued  As String
End Type

Private Type CERTIFICATE_BLOB_PROPERTY
    PropertyID As Long
    Reserved As Long
    Length As Long
    Data() As Byte
End Type

Private Declare Sub OutputDebugStringA Lib "kernel32.dll" (ByVal lpOutputString As String)

Private HitSorted()     As String

Public gProcess()           As MY_PROC_ENTRY
Public g_TasksWL()          As TASK_WHITELIST_ENTRY
Public oDict                As DICTIONARIES

Public oDictFileExist       As clsTrickHashTable

Public Scan()   As SCAN_RESULT    '// Dragokas. Used instead of parsing lines from result screen (like it was in original HJT 2.0.5).
                                  '// User type structures of arrays is filled together with using of method frmMain.lstResults.AddItem
                                  '// It is much efficiently and have Unicode support (native vb6 ListBox is ANSI only).
                                  '// Result screen will be replaced with CommonControls unicode aware controls by Krool (vbforums.com) in the nearest update,
                                  '// as well as StartupList2 by Merijn that currently use separate Microsoft MSCOMCTL.OCX library file.

Public Perf     As TYPE_PERFORMANCE

Public OSver    As clsOSInfo
Public Proc     As clsProcess
Public cMath    As clsMath

Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal pszStrPtr As Long, ByVal Length As Long) As String


'it map ANSI scan result string from ListBox to Unicode string that is stored in memory (SCAN_RESULT structure)
Public Function GetScanResults(HitLineA As String, Result As SCAN_RESULT) As Boolean
    Dim i As Long
    For i = 1 To UBound(Scan)
        If HitLineA = Scan(i).HitLineA Then
            Result = Scan(i)
            GetScanResults = True
            Exit Function
        End If
    Next
    'Cannot find appropriate cure item for:, "Error"
    MsgBoxW Translate(592) & vbCrLf & HitLineA, vbCritical, Translate(591)
End Function

' it add Unicode SCAN_RESULT structure to shared array
Public Sub AddToScanResults(Result As SCAN_RESULT, Optional ByVal DoNotAddToListBox As Boolean, Optional DontClearResults As Boolean)
    Dim bFirstWarning As Boolean
    
    'LockWindowUpdate frmMain.lstResults.hwnd
    
    If bAutoLogSilent Then
        DoNotAddToListBox = True
    Else
        DoEvents
    End If
    If Not DoNotAddToListBox Then
        'checking if one of sections planned to be contains more then 50 entries -> block such attempt
        If Not SectionOutOfLimit(Result.Section, bFirstWarning) Then
            frmMain.lstResults.AddItem Result.HitLineW
            'select the last added line
            frmMain.lstResults.ListIndex = frmMain.lstResults.ListCount - 1
        Else
            If bFirstWarning Then
                frmMain.lstResults.AddItem Result.Section & " - Too many entries ( > 50 )"
                frmMain.lstResults.ListIndex = frmMain.lstResults.ListCount - 1
            End If
        End If
    End If
    ReDim Preserve Scan(UBound(Scan) + 1)
    'Unicode to ANSI mapping (dirty hack)
    Result.HitLineA = frmMain.lstResults.List(frmMain.lstResults.ListCount - 1)
    Scan(UBound(Scan)) = Result
    'Sleep 5
    'LockWindowUpdate False
    'Erase Result struct
    If Not DontClearResults Then
        EraseScanResults Result
    End If
End Sub

Public Sub EraseScanResults(Result As SCAN_RESULT)
    Dim EmptyResult As SCAN_RESULT
    Result = EmptyResult
End Sub

'// Increase number of sections +1 and returns TRUE, if total number > LIMIT
Private Function SectionOutOfLimit(p_Section As String, Optional bFirstWarning As Boolean, Optional bErase As Boolean) As Long
    Const LIMIT As Long = 50&
    
    Static Section As String
    Static Num As Long
    
    If bErase = True Then
        Section = ""
        Num = 0
        Exit Function
    End If
    
    If p_Section = Section Then
        Num = Num + 1
        If Num > LIMIT Then
            If Num = LIMIT + 1 Then
                bFirstWarning = True
            End If
            SectionOutOfLimit = True
        End If
    Else
        Section = p_Section
        Num = 1
    End If
End Function

Public Sub AddToScanResultsSimple(Section As String, HitLine As String, Optional DoNotAddToListBox As Boolean)
    Dim Result As SCAN_RESULT
    With Result
        .Section = Section
        .HitLineW = HitLine
    End With
    AddToScanResults Result, DoNotAddToListBox
End Sub

Public Sub GetHosts()
    If bIsWinNT Then
        sHostsFile = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DataBasePath")
        'sHostsFile = replace$(sHostsFile, "%SystemRoot%", sWinDir, , , vbTextCompare)
        sHostsFile = EnvironW(sHostsFile) & "\hosts"
    End If
End Sub

Public Sub UpdateIE_RegVals()
    On Error GoTo ErrorHandler:

    Dim i As Long
    
    ReLoadIE_RegVals
    
    g_DEFSTARTPAGE = frmMain.txtDefStartPage.Text
    g_DEFSEARCHPAGE = frmMain.txtDefSearchPage.Text
    g_DEFSEARCHASS = frmMain.txtDefSearchAss.Text
    g_DEFSEARCHCUST = frmMain.txtDefSearchCust.Text
    
    For i = 0 To UBound(sRegVals)
        If sRegVals(i) = vbNullString Then Exit For
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSTARTPAGE", g_DEFSTARTPAGE)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHPAGE", g_DEFSEARCHPAGE)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHASS", g_DEFSEARCHASS)
        sRegVals(i) = Replace$(sRegVals(i), "$DEFSEARCHCUST", g_DEFSEARCHCUST)
        
        'sRegVals(i) = replace$(sRegVals(i), "$WINSYSDIR", sWinSysDir)
        'sRegVals(i) = replace$(sRegVals(i), "$WINDIR", sWinDir)
        sRegVals(i) = EnvironW(sRegVals(i))
    Next i
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "UpdateIE_RegVals"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub ReLoadIE_RegVals()
    On Error GoTo ErrorHandler:

    '=== LOAD REGVALS ===
    'syntax:
    '  regkey,regvalue,resetdata,baddata
    '  |      |        |          |
    '  |      |        |          data that shouldn't be (never used)
    '  |      |        R0 - data to reset to
    '  |      R1 - value to check
    '  R2 - regkey to check
    '
    'when empty:
    'R0 - everything is considered bad (always used), change to resetdata
    'R1 - value being present is considered bad, delete value
    'R2 - key being present is considered bad, delete key (not used)
    
    Dim i As Long
    Dim colRegIE As Collection
    Set colRegIE = New Collection
    
    Dim Hive
    Dim Default_Page_URL$: Default_Page_URL = "http://go.microsoft.com/fwlink/p/?LinkId=255141"
    Dim Default_Search_URL$: Default_Search_URL = "http://go.microsoft.com/fwlink/?LinkId=54896"
    
    With colRegIE
    
        .Add "Software\Microsoft\Internet Explorer,Default_Page_URL," & Default_Page_URL & "|,"
        .Add "Software\Microsoft\Internet Explorer\Main,Default_Page_URL," & Default_Page_URL & "|http://www.msn.com|res://iesetup.dll/HardAdmin.htm|res://shdoclc.dll/softAdmin.htm|,"
        .Add "Software\Microsoft\Internet Explorer\Search,Default_Page_URL," & Default_Page_URL & "|,"
        
        .Add "Software\Microsoft\Internet Explorer,Default_Search_URL," & Default_Search_URL & "|,"
        .Add "Software\Microsoft\Internet Explorer\Main,Default_Search_URL," & Default_Search_URL & "|,"
        .Add "Software\Microsoft\Internet Explorer\Search,Default_Search_URL," & Default_Search_URL & "|,"
        
        .Add "Software\Microsoft\Internet Explorer,SearchAssistant,,"
        .Add "Software\Microsoft\Internet Explorer,CustomizeSearch,,"
        .Add "Software\Microsoft\Internet Explorer,Search,,"
        .Add "Software\Microsoft\Internet Explorer,Search Bar,,"
        .Add "Software\Microsoft\Internet Explorer,Search Page,,"
        .Add "Software\Microsoft\Internet Explorer,Start Page,,"
        .Add "Software\Microsoft\Internet Explorer,SearchURL,,"
        .Add "Software\Microsoft\Internet Explorer,(Default),,"
        .Add "Software\Microsoft\Internet Explorer,www,,"
        
        .Add "Software\Microsoft\Internet Explorer\Main,SearchAssistant,,"
        .Add "Software\Microsoft\Internet Explorer\Main,CustomizeSearch,,"
        .Add "Software\Microsoft\Internet Explorer\Main,Search Bar,http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm|Preserve|,"
        .Add "Software\Microsoft\Internet Explorer\Main,Search Page,http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch|www.bing.com|,"
        .Add "Software\Microsoft\Internet Explorer\Main,Start Page,$DEFSTARTPAGE|http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=msnhome|http://www.microsoft.com/isapi/redir.dll?prd={SUB_PRD}&clcid={SUB_CLSID}&pver={SUB_PVER}&ar=home|res://iesetup.dll/HardAdmin.htm|,"
        .Add "Software\Microsoft\Internet Explorer\Main,SearchURL,,"
        .Add "Software\Microsoft\Internet Explorer\Main,Start Page Redirect Cache,http://ru.msn.com/?ocid=iehp|,"
        
        .Add "Software\Microsoft\Internet Explorer\Search,SearchAssistant,$DEFSEARCHASS|,"
        .Add "Software\Microsoft\Internet Explorer\Search,CustomizeSearch,$DEFSEARCHCUST|,"
        .Add "Software\Microsoft\Internet Explorer\Search,(Default),,"
        
        .Add "Software\Microsoft\Internet Explorer\SearchURL,(Default),,"
        .Add "Software\Microsoft\Internet Explorer\SearchURL,SearchURL,,"
        
        .Add "Software\Microsoft\Internet Explorer\Main,Startpagina,,"
        .Add "Software\Microsoft\Internet Explorer\Main,First Home Page,|res://iesetup.dll/HardAdmin.htm,"
        .Add "Software\Microsoft\Internet Explorer\Main,Local Page,%SystemRoot%\System32\blank.htm|%SystemRoot%\SysWOW64\blank.htm|%11%\blank.htm|,"
        .Add "Software\Microsoft\Internet Explorer\Main,Start Page_bak,,"
        .Add "Software\Microsoft\Internet Explorer\Main,HomeOldSP,,"
        .Add "Software\Microsoft\Internet Explorer\Main,YAHOOSubst,,"
        .Add "Software\Microsoft\Internet Explorer\Main,Window Title,,"
        
        .Add "Software\Microsoft\Internet Explorer\Main,Extensions Off Page,about:NoAdd-ons|,"
        .Add "Software\Microsoft\Internet Explorer\Main,Security Risk Page,about:SecurityRisk|,"
        
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,blank,res://mshtml.dll/blank.htm|,"
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,DesktopItemNavigationFailure,res://ieframe.dll/navcancl.htm|res://shdoclc.dll/navcancl.htm|,"
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,InPrivate,res://ieframe.dll/inprivate.htm|res://ieframe.dll/inprivate_win7.htm|,"
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,NavigationCanceled,res://ieframe.dll/navcancl.htm|res://shdoclc.dll/navcancl.htm|,"
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,NavigationFailure,res://ieframe.dll/navcancl.htm|res://shdoclc.dll/navcancl.htm|,"
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,NoAdd-ons,res://ieframe.dll/noaddon.htm|,"
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,NoAdd-onsInfo,res://ieframe.dll/noaddoninfo.htm|,"
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,PostNotCached,res://ieframe.dll/repost.htm|res://mshtml.dll/repost.htm|,"
        .Add "Software\Microsoft\Internet Explorer\AboutURLs,SecurityRisk,res://ieframe.dll/securityatrisk.htm|,"

        .Add "Software\Microsoft\Internet Connection Wizard,ShellNext,|http://windowsupdate.microsoft.com/,"
        
        .Add "Software\Microsoft\Internet Explorer\Toolbar,LinksFolderName,Links|Ссылки|,"

        .Add "Software\Microsoft\Windows\CurrentVersion\Internet Settings,AutoConfigURL,,"
        .Add "Software\Microsoft\Windows\CurrentVersion\Internet Settings,ProxyServer,,"
        .Add "Software\Microsoft\Windows\CurrentVersion\Internet Settings,ProxyOverride,,"

        'Only short hive names permitted here !
        
        .Add "HKLM\System\CurrentControlSet\services\NlaSvc\Parameters\Internet\ManualProxies,(Default),,"
        .Add "HKLM\SOFTWARE\Clients\StartMenuInternet\IEXPLORE.EXE\shell\open\command,(Default)," & _
            IIf(bIsWin64, "%ProgramW6432%", "%ProgramFiles%") & "\Internet Explorer\iexplore.exe" & _
            IIf(bIsWin64, "|%ProgramFiles(x86)%\Internet Explorer\iexplore.exe", "") & _
            "|" & """" & "%ProgramFiles%\Internet Explorer\iexplore.exe" & """" & _
            IIf(OSver.MajorMinor <= 5, "|", "") & _
            "|iexplore.exe" & _
            ","
        
    End With
    ReDim sRegVals(colRegIE.Count - 1)
    For i = 1 To colRegIE.Count
        sRegVals(i - 1) = colRegIE.Item(i)
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ReLoadIE_RegVals"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub LoadStuff()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "LoadStuff - Begin"
    
    Dim i As Long
    
    '=== LOAD FILEVALS ===
    'syntax:
    ' inifile,section,value,resetdata,baddata
    ' |       |       |     |         |
    ' |       |       |     |         5) data that shouldn't be (never used)
    ' |       |       |     4) data to reset to
    ' |       |       |        (delete all if empty)
    ' |       |       3) value to check
    ' |       2) section to check
    ' 1) file to check
    
    Dim colFileVals As Collection
    Set colFileVals = New Collection
    
    'F0, F2 - if value modified
    'F1, F3 - if param. created
    
    With colFileVals
        .Add "system.ini;boot;Shell;explorer.exe;"        'F0 (boot;Shell)
        .Add "win.ini;windows;load;;"                     'F1 (windows;load)
        .Add "win.ini;windows;run;;"                      'F1 (windows;run)
        '\Software\Microsoft\Windows NT\CurrentVersion\WinLogon   'F2 (boot;Shell)
        .Add "REG:system.ini;boot;Shell;explorer.exe|%WINDIR%\explorer.exe;"
        '\Software\Microsoft\Windows NT\CurrentVersion\WinLogon   'F2 (boot;UserInit)
        .Add "REG:system.ini;boot;UserInit;%WINDIR%\System32\userinit.exe|userinit.exe;"
        '\Software\Microsoft\Windows NT\CurrentVersion\Windows    'F3 (windows;load)
        .Add "REG:win.ini;windows;load;;"
        '\Software\Microsoft\Windows NT\CurrentVersion\Windows    'F3 (windows;run)
        .Add "REG:win.ini;windows;run;;"
    End With
    ReDim sFileVals(colFileVals.Count - 1)
    For i = 1 To colFileVals.Count
        sFileVals(i - 1) = colFileVals.Item(i)
    Next

    '//TODO:
    '
    'What are ShellInfrastructure, VMApplet under winlogon ?
    'there are also 2 dll-s that may be interesting under \Windows (NaturalInputHandler, IconServiceLib)
    

    'R0, R1
    
    ReLoadIE_RegVals

    
    'R4 database
    
    'HKCU\Software\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Microsoft\Internet Explorer\SearchScopes
    
    Dim aHives() As String, sHive$, j&
    
    'GetBrowsersInfo
    
    GetHives aHives, bIncludeServiceSID:=False
    
    With cReg4vals
        '.Add "HKCU,DisplayName,Bing"
        '.Add "HKCU,FaviconURLFallback,http://www.bing.com/favicon.ico"
        
        If OSver.MajorMinor >= 6.3 Then 'Win 8.1
            '.Add "HKCU,FaviconURL,"
            '.Add "HKCU,NTLogoPath,"
            '.Add "HKCU,NTLogoURL,"
            '.Add "HKCU,NTSuggestionsURL,"
            '.Add "HKCU,NTTopResultURL,"
            '.Add "HKCU,NTURL,"
            .Add "HKCU,SuggestionsURL,"
            .Add "HKCU,TopResultURL,"
        
            .Add "HKCU,SuggestionsURLFallback,http://api.bing.com/qsml.aspx?query={searchTerms}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IE11SS&market={language}"
            .Add "HKCU,TopResultURLFallback,http://www.bing.com/search?q={searchTerms}&src=IE-TopResult&FORM=IE11TR"
            .Add "HKCU,URL,http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IE11SR"
            
        Else
        
            '.Add "HKCU,FaviconURL,http://www.bing.com/favicon.ico"
            '.Add "HKCU,NTLogoPath," & AppDataLocalLow & "\Microsoft\Internet Explorer\Services\"
            '.Add "HKCU,NTLogoURL,http://go.microsoft.com/fwlink/?LinkID=403856&language={language}&scale={scalelevel}&contrast={contrast}"
            '.Add "HKCU,NTSuggestionsURL,http://api.bing.com/qsml.aspx?query={searchTerms}&market={language}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IENTSS"
            '.Add "HKCU,NTTopResultURL,http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTTR"
            '.Add "HKCU,NTURL,http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTSR"
            
            .Add "HKCU,SuggestionsURL,http://api.bing.com/qsml.aspx?query={searchTerms}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IESS02&market={language}"
            .Add "HKCU,TopResultURL,http://www.bing.com/search?q={searchTerms}&src=IE-TopResult&FORM=IETR02"
            .Add "HKCU,SuggestionsURLFallback,http://api.bing.com/qsml.aspx?query={searchTerms}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IESS02&market={language}"
            .Add "HKCU,TopResultURLFallback,http://www.bing.com/search?q={searchTerms}&src=IE-TopResult&FORM=IETR02"
            .Add "HKCU,URL,http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IESR02"
        End If
        
        '.Add "HKLM,,Bing"
        '.Add "HKLM,DisplayName,@ieframe.dll,-12512|Bing"
        .Add "HKLM,URL,http://www.bing.com/search?q={searchTerms}&FORM=IE8SRC"
    End With
    
    For i = 1 To cReg4vals.Count  ' append HKU hive
        sHive = Left$(cReg4vals.Item(i), 4)
        If sHive = "HKCU" Then
            For j = 0 To UBound(aHives)
                If Left$(aHives(j), 3) = "HKU" Then
                    cReg4vals.Add Replace$(cReg4vals.Item(i), "HKCU", aHives(j), 1, 1)
                End If
            Next
        End If
    Next
    
    
    ' === LOAD NONSTANDARD-BUT-SAFE-DOMAINS LIST ===
    
    Dim colSafeRegDomains As Collection
    Set colSafeRegDomains = New Collection
    
    With colSafeRegDomains
        .Add "http://www.microsoft.com"
        .Add "http://home.microsoft.com"
        .Add "http://www.msn.com"
        .Add "http://search.msn.com"
        .Add "http://ie.search.msn.com"
        .Add "ie.search.msn.com"
        .Add "<local>"
        .Add "http://www.google.com"
        .Add "127.0.0.1;localhost"
        .Add "about:blank"
        .Add "http://go.microsoft.com/"
        .Add "www.microsoft.com/"
        .Add "microsoft.com"
        .Add "http://windowsupdate.com"
        .Add "http://runonce.msn.com"
        ' "iexplore"
        ' "http://www.aol.com"
    End With
    ReDim aSafeRegDomains(colSafeRegDomains.Count - 1)
    For i = 1 To colSafeRegDomains.Count
        aSafeRegDomains(i - 1) = colSafeRegDomains.Item(i)
    Next
    
    ' === LOAD LSP PROVIDERS SAFELIST ===
    'asterisk is used for filename separation, because.
    'did you ever see a filename with an asterisk?
    sSafeLSPFiles = "*A2antispamlsp.dll*Adlsp.dll*Agbfilt.dll*Antiyfilter.dll*Ao2lsp.dll*Aphish.dll*Asdns.dll*Aslsp.dll*Asnsp.dll*Avgfwafu.dll*Avsda.dll*Betsp.dll*Biolsp.dll*Bmi_lsp.dll*Caslsp.dll*Cavemlsp.dll*Cdnns.dll*Connwsp.dll*Cplsp.dll*Csesck32.dll*Cslsp.dll*Cssp.al*Ctxlsp.dll*Ctxnsp.dll*Cwhook.dll*Cwlsp.dll*Dcsws2.dll*Disksearchservicestub.dll*Drwebsp.dll*Drwhook.dll*Espsock2.dll*Farlsp.dll*Fbm.dll*Fbm_lsp.dll*Fortilsp.dll*Fslsp.dll*Fwcwsp.dll*Fwtunnellsp.dll*Gapsp.dll*Googledesktopnetwork1.dll*Hclsock5.dll*Iapplsp.dll*Iapp_lsp.dll*Ickgw32i.dll*Ictload.dll*Idmmbc.dll*Iga.dll*Imon.dll*Imslsp.dll*Inetcntrl.dll*Ippsp.dll*Ipsp.dll*Iss_clsp.dll*Iss_slsp.dll*Kvwsp.dll*Kvwspxp.dll*Lslsimon.dll*Lsp32.dll*" & _
        "Lspcs.dll*Mclsp.dll*Mdnsnsp.dll*Msafd.dll*Msniffer.dll*Mswsock.dll*Mswsosp.dll*Mwtsp.dll*Mxavlsp.dll*Napinsp.dll*Nblsp.dll*Ndpwsspr.dll*Netd.dll*Nihlsp.dll*Nlaapi.dll*Nl_lsp.dll*Nnsp.dll*Normanpf.dll*Nutafun4.dll*Nvappfilter.dll*Nwws2nds.dll*Nwws2sap.dll*Nwws2slp.dll*Odsp.dll*Pavlsp.dll*Pclsp.dll*Pctlsp.dll*Pfftsp.dll*Pgplsp.dll*Pidlsp.dll*Pnrpnsp.dll*Prifw.dll*Proxy.dll*Prplsf.dll*Pxlsp.dll*Rnr20.dll*Rsvpsp.dll*S5spi.dll*Samnsp.dll*Sarah.dll*Scopinet.dll*Skysocks.dll*Sliplsp.dll*Smnsp.dll*Spacklsp.dll*Spampallsp.dll*Spi.dll*Spidll.dll*Spishare.dll*Spsublsp.dll*Sselsp.dll*Stplayer.dll*Syspy.dll*Tasi.dll*Tasp.dll*Tcpspylsp.dll*Ua_lsp.dll*Ufilter.dll*Vblsp.dll*Vetredir.dll*Vlsp.dll*Vnsp.dll*" & _
        "Wglsp.dll*Whllsp.dll*Whlnsp.dll*Winrnr.dll*Wins4f.dll*Winsflt.dll*WinSysAM.dll*Wps.dll*Wshbth.dll*Wspirda.dll*Wspwsp.dll*Xfilter.dll*xfire_lsp.dll*Xnetlsp.dll*Ypclsp.dll*Zklspr.dll*_Easywall.dll*_Handywall.dll*vsocklib.dll*wlidnsp.dll*"
    
    ' === LOAD PROTOCOL SAFELIST === (O18)
    
    Dim colSafeProtocols As Collection
    Set colSafeProtocols = New Collection
    
    '//TODO: O18 - add file path checking to database
    
    With colSafeProtocols
        .Add "about|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "belarc|{6318E0AB-2E93-11D1-B8ED-00608CC9A71F}"
        .Add "BPC|{3A1096B3-9BFA-11D1-AE77-00C04FBBDEBC}"
        .Add "CDL|{3DD53D40-7B8B-11D0-B013-00AA0059CE02}"
        .Add "cdo|{CD00020A-8B95-11D1-82DB-00C04FB1625D}"
        .Add "copernicagentcache|{AAC34CFD-274D-4A9D-B0DC-C74C05A67E1D}"
        .Add "copernicagent|{A979B6BD-E40B-4A07-ABDD-A62C64A4EBF6}"
        .Add "dodots|{9446C008-3810-11D4-901D-00B0D04158D2}"
        .Add "DVD|{12D51199-0DB5-46FE-A120-47A3D7D937CC}"
        .Add "file|{79EAC9E7-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ftp|{79EAC9E3-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "gopher|{79EAC9E4-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "https|{79EAC9E5-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "http|{79EAC9E2-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ic32pp|{BBCA9F81-8F4F-11D2-90FF-0080C83D3571}"
        .Add "ipp|"
        .Add "its|{9D148291-B9C8-11D0-A4CC-0000F80149F6}"
        .Add "javascript|{3050F3B2-98B5-11CF-BB82-00AA00BDCE0B}" '|<SysRoot>\System32\mshtml.dll
        .Add "junomsg|{C4D10830-379D-11D4-9B2D-00C04F1579A5}"
        .Add "lid|{5C135180-9973-46D9-ABF4-148267CBB8BF}"
        .Add "local|{79EAC9E7-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "mailto|{3050F3DA-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "mctp|{D7B95390-B1C5-11D0-B111-0080C712FE82}"
        .Add "mhtml|{05300401-BCBC-11D0-85E3-00C04FD85AB4}"
        .Add "mk|{79EAC9E6-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ms-its50|{F8606A00-F5CF-11D1-B6BB-0000F80149F6}"
        .Add "ms-its51|{F6F1E82D-DE4D-11D2-875C-0000F8105754}"
        .Add "ms-itss|{0A9007C0-4076-11D3-8789-0000F8105754}"
        .Add "ms-its|{9D148291-B9C8-11D0-A4CC-0000F80149F6}"
        .Add "msdaipp|"
        .Add "mso-offdap|{3D9F03FA-7A94-11D3-BE81-0050048385D1}"
        .Add "ndwiat|{13F3EA8B-91D7-4F0A-AD76-D2853AC8BECE}"
        .Add "res|{3050F3BC-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "sysimage|{76E67A63-06E9-11D2-A840-006008059382}"
        .Add "tve-trigger|{CBD30859-AF45-11D2-B6D6-00C04FBBDE6E}"
        .Add "tv|{CBD30858-AF45-11D2-B6D6-00C04FBBDE6E}"
        .Add "vbscript|{3050F3B2-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "vnd.ms.radio|{3DA2AA3B-3D96-11D2-9BD2-204C4F4F5020}"
        .Add "wia|{13F3EA8B-91D7-4F0A-AD76-D2853AC8BECE}"
        .Add "mso-offdap11|{32505114-5902-49B2-880A-1F7738E5A384}"
        .Add "DirectDVD|{85A81A02-336B-43FF-998B-FE8E194FBA4D}"
        .Add "pcn|{D540F040-F3D9-11D0-95BE-00C04FD93CA5}"
        .Add "msencarta|{74D92DF3-6D9D-11D1-8B38-006097DBED7A}"
        .Add "msero|{B0D92A71-886B-453B-A649-1B91F93801E7}"
        .Add "msref|{74D92DF3-6D9D-11D1-8B38-006097DBED7A}"
        .Add "df2|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df3|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df4|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df5|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df23chat|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df5demo|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "ofpjoin|{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "saphtmlp|{D1F8BD1E-7967-11D2-B43A-006094B9EADB}"
        .Add "sapr3|{D1F8BD1E-7967-11D2-B43A-006094B9EADB}"
        .Add "lbxfile|{56831180-F115-11D2-B6AA-00104B2B9943}"
        .Add "lbxres|{24508F1B-9E94-40EE-9759-9AF5795ADF52}"
        .Add "cetihpz|{CF184AD3-CDCB-4168-A3F7-8E447D129300}"
        .Add "aim|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "shell|{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "asp|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "hsp|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-asp|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-hsp|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-zip|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "zip|{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "bega|{A57721C9-B905-49B3-8BCA-B99FBB8C627E}"
        .Add "bt2|{1730B77B-F429-498F-9B15-4514D83C8294}"
        .Add "cetihpz|{CF184AD3-CDCB-4168-A3F7-8E447D129300}"
        .Add "copernicdesktopsearch|{D9656C75-5090-45C3-B27E-436FBC7ACFA7}"
        .Add "crick|{B861500A-A326-11D3-A248-0080C8F7DE1E}"
        .Add "dadb|{82D6F09F-4AC2-11D3-8BD9-0080ADB8683C}"
        .Add "dialux|{8352FA4C-39C6-11D3-ADBA-00A0244FB1A2}"
        .Add "emistp|{0EFAEA2E-11C9-11D3-88E3-0000E867A001}"
        .Add "ezstor|{6344A3A0-96A7-11D4-88CC-000000000000}"
        .Add "flowto|{C7101FB0-28FB-11D5-883A-204C4F4F5020}"
        .Add "g7ps|{9EACF0FB-4FC7-436E-989B-3197142AD979}"
        .Add "intu-res|{9CE7D474-16F9-4889-9BB9-53E2008EAE8A}"
        .Add "iwd|{EA5F5649-A6C7-11D4-9E3C-0020AF0FFB56}"
        .Add "mavencache|{DB47FDC2-8C38-4413-9C78-D1A68BF24EED}"
        .Add "ms-help|{314111C7-A502-11D2-BBCA-00C04F8EC294}"
        .Add "msnim|{828030A1-22C1-4009-854F-8E305202313F}"
        .Add "myrm|{4D034FC3-013F-4B95-B544-44D49ABE3E76}"
        .Add "nbso|{DF700763-3EAD-4B64-9626-22BEEFF3EA47}"
        .Add "nim|{3D206AE2-3039-413B-B748-3ACC562EC22A}"
        .Add "OWC11.mso-offdap|{32505114-5902-49B2-880A-1F7738E5A384}"
        .Add "pcl|{182D0C85-206F-4103-B4FA-DCC1FB0A0A44}"
        .Add "pure-go|{4746C79A-2042-4332-8650-48966E44ABA8}"
        .Add "qrev|{9DE24BAC-FC3C-42C4-9FC4-76B3FAFDBD90}"
        .Add "rmh|{23C585BB-48FF-4865-8934-185F0A7EB84C}"
        .Add "SafeAuthenticate|{8125919B-9BE9-4213-A1D6-75188A22D21E}"
        .Add "sds|{79E0F14C-9C52-4218-89A7-7C4B0563D121}"
        .Add "siteadvisor|{3A5DC592-7723-4EAA-9EE6-AF4222BCF879}"
        .Add "smscrd|{FA3F5003-93D4-11D2-8E48-00A0C98BD8C3}"
        .Add "stibo|{FFAD3420-6D61-44F6-BA25-293F17152D79}"
        .Add "textwareilluminatorbase|{CE5CD329-1650-414A-8DB0-4CBF72FAED87}"
        .Add "widimg|{EE7C2AFF-5742-44FF-BD0E-E521B0D3C3BA}"
        .Add "wlmailhtml|{03C514A3-1EFB-4856-9F99-10D7BE1653C0}"
        .Add "x-atng|{7E8717B0-D862-11D5-8C9E-00010304F989}"
        .Add "x-excid|{9D6CC632-1337-4A33-9214-2DA092E776F4}"
        .Add "x-mem1|{C3719F83-7EF8-4BA0-89B0-3360C7AFB7CC}"
        .Add "x-mem3|{4F6D06DD-44AB-4F89-BF13-9027B505B15A}"
        .Add "ct|{774E529C-2458-48A2-8F57-3ED3105D8612}"
        .Add "cw|{774E529C-2458-48A2-8F57-3ED3105D8612}"
        .Add "eti|{3AAE7392-E7AA-11D2-969E-00105A088846}"
        .Add "livecall|{828030A1-22C1-4009-854F-8E305202313F}"
        .Add "tbauth|{14654CA6-5711-491D-B89A-58E571679951}"
        .Add "windows.tbauth|{14654CA6-5711-491D-B89A-58E571679951}"
    End With
    ReDim aSafeProtocols(colSafeProtocols.Count - 1)
    For i = 1 To colSafeProtocols.Count
        aSafeProtocols(i - 1) = colSafeProtocols.Item(i)
    Next
    sSafeProtocols = Join(aSafeProtocols, vbCrLf)
        
    ' === LOAD FILTER SAFELIST === (O18)
    
    Dim colSafeFilters As Collection
    Set colSafeFilters = New Collection
        
    With colSafeFilters
        .Add "application/octet-stream|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "application/x-complus|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "application/x-msdownload|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "Class Install Handler|{32B533BB-EDAE-11d0-BD5A-00AA00B92AF1}"
        .Add "deflate|{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "gzip|{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "lzdhtml|{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "text/webviewhtml|{733AC4CB-F1A4-11d0-B951-00A0C90312E1}"
        .Add "text/xml|{807553E5-5146-11D5-A672-00B0D022E945}"
        .Add "application/x-icq|{db40c160-09a1-11d3-baf2-000000000000}"
        'added in HJT 1.99.2 final
        .Add "application/msword|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd.ms-excel|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd.ms-powerpoint|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/x-microsoft-rpmsg-message|{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd-backup-octet-stream|{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "application/vnd-viewer|{CD4527E8-4FC7-48DB-9806-10537B501237}"
        .Add "application/x-bt2|{6E1DDCE8-76BC-4390-9488-806E8FB1AD77}"
        .Add "application/x-internet-signup|{A173B69A-1F9B-4823-9FDA-412F641E65D6}"
        .Add "text/html|{8D42AD12-D7A1-4797-BCB7-AD89E5FCE4F7}"
        .Add "text/html|{F79B2338-A6E7-46D4-9201-422AA6E74F43}"
        .Add "text/x-mrml|{C51721BE-858B-4A66-A8BF-D2882FF49820}"
        .Add "text/xml|{807563E5-5146-11D5-A672-00B0D022E945}"
        .Add "application/octet-stream|{F969FE8E-1937-45AD-AF42-8A4D11CBDC2A}"
        .Add "application/xhtml+xml|{32F66A26-7614-11D4-BD11-00104BD3F987}"
        .Add "text/xml|{32F66A26-7614-11D4-BD11-00104BD3F987}"
    End With
    ReDim aSafeFilters(colSafeFilters.Count - 1)
    For i = 1 To colSafeFilters.Count
        aSafeFilters(i - 1) = colSafeFilters.Item(i)
    Next
    sSafeFilters = Join(aSafeFilters, vbCrLf)


    'LOAD APPINIT_DLLS SAFELIST (O20)
    sSafeAppInit = "*aakah.dll*akdllnt.dll*ROUSRNT.DLL*ssohook*KATRACK.DLL*APITRAP.DLL*UmxSbxExw.dll*sockspy.dll*scorillont.dll*wbsys.dll*NVDESK32.DLL*hplun.dll*mfaphook.dll*PAVWAIT.DLL*OCMAPIHK.DLL*MsgPlusLoader.dll*IconCodecService.dll*wl_hook.dll*Google\GOOGLE~1\GOEC62~1.DLL*adialhk.dll*wmfhotfix.dll*interceptor.dll*qaphooks.dll*RMProcessLink.dll*msgrmate.dll*wxvault.dll*ctu33.dll*ati2evxx.dll*vsmvhk.dll*"
    
    'LOAD SSODL SAFELIST (O21)
    
    Dim colSafeSSODL As Collection
    Set colSafeSSODL = New Collection
        
    With colSafeSSODL
        .Add "{E6FB5E20-DE35-11CF-9C87-00AA005127ED}"  'WebCheck: C:\WINDOWS\System32\webcheck.dll (WinAll)
        .Add "{35CEC8A3-2BE6-11D2-8773-92E220524153}"  'SysTray: C:\WINDOWS\System32\stobject.dll (Win2k/XP)
        .Add "{7849596a-48ea-486e-8937-a2a3009f31a9}"  'PostBootReminder: C:\WINDOWS\system32\SHELL32.dll (WinXP)
        .Add "{fbeb8a05-beee-4442-804e-409d6c4515e9}"  'CDBurn: C:\WINDOWS\system32\SHELL32.dll (WinXP)
        .Add "{11566B38-955B-4549-930F-7B7482668782}"  'AUHook: C:\WINDOWS\SYSTEM\AUHOOK.DLL (WinME)
        .Add "{7007ACCF-3202-11D1-AAD2-00805FC1270E}"  'Network.ConnectionTray: C:\WINNT\system32\NETSHELL.dll (Win2k)
        .Add "{e57ce738-33e8-4c51-8354-bb4de9d215d1}"  'UPnPMonitor: C:\WINDOWS\SYSTEM\UPNPUI.DLL (WinME/XP)
        .Add "{BCBCD383-3E06-11D3-91A9-00C04F68105C}"  'AUHook: C:\WINDOWS\SYSTEM\AUHOOK.DLL (WinME)
        .Add "{F5DF91F9-15E9-416B-A7C3-7519B11ECBFC}"  '0aMCPClient: C:\Program Files\StarDock\MCPCore.dll
        .Add "{AAA288BA-9A4C-45B0-95D7-94D524869DB5}"  'WPDShServiceObj   WPDShServiceObj.dll Windows Portable Device Shell Service Object
        .Add "{1799460C-0BC8-4865-B9DF-4A36CD703FF0}" 'IconPackager Repair  iprepair.dll    Stardock\Object Desktop\ ThemeManager
        .Add "{6D972050-A934-44D7-AC67-7C9E0B264220}" 'EnhancedDialog   enhdlginit.dll  EnhancedDialog by Stardock
    End With
    'BE AWARE: SHELL32.dll - sometimes this file is patched (e.g. seen in Simplix)
    ReDim aSafeSSODL(colSafeSSODL.Count - 1)
    For i = 1 To colSafeSSODL.Count
        aSafeSSODL(i - 1) = colSafeSSODL.Item(i)
    Next
    
    'LOAD SIOI SAFELIST (O21)
    
    Dim colSafeSIOI As Collection
    Set colSafeSIOI = New Collection
        
    With colSafeSIOI
'        .Add "{D9144DCD-E998-4ECA-AB6A-DCD83CCBA16D}"  'EnhancedStorageShell: C:\Windows\system32\EhStorShell.dll (Win7)
'        .Add "{4E77131D-3629-431c-9818-C5679DC83E81}"  'Offline Files: C:\Windows\System32\cscui.dll (Win7)
'        .Add "{08244EE6-92F0-47f2-9FC9-929BAA2E7235}"  'SharingPrivate: C:\Windows\system32\ntshrui.dll (Win7)
'        .Add "{750fdf0e-2a26-11d1-a3ea-080036587f03}"  'Offline Files: C:\WINDOWS\System32\cscui.dll (WinXP)
'        .Add "{fbeb8a05-beee-4442-804e-409d6c4515e9}"  'CDBurn: C:\WINDOWS\system32\SHELL32.dll (WinXP)
'        .Add "{7849596a-48ea-486e-8937-a2a3009f31a9}"  'PostBootReminder: C:\WINDOWS\system32\SHELL32.dll (WinXP)
'        .Add "{0CA2640D-5B9C-4c59-A5FB-2DA61A7437CF}" 'StorageProviderError: C:\Windows\System32\shell32.dll, C:\Windows\SysWOW64\shell32.dll (Win 8.1)
'        .Add "{0A30F902-8398-4ee8-86F7-4CFB589F04D1}" 'StorageProviderSyncing: C:\Windows\System32\shell32.dll, C:\Windows\SysWOW64\shell32.dll (Win 8.1)

        .Add "<SysRoot>\system32\EhStorShell.dll"
        .Add "<SysRoot>\system32\cscui.dll"
        .Add "<SysRoot>\system32\ntshrui.dll"
        .Add "<SysRoot>\system32\SHELL32.dll"
        .Add "<SysRoot>\SysWOW64\shell32.dll"
        '.Add "<SysRoot>\system32\mscoree.dll" 'adware
    End With
    ReDim aSafeSIOI(colSafeSIOI.Count - 1)
    For i = 1 To colSafeSIOI.Count
        aSafeSIOI(i - 1) = Replace(colSafeSIOI.Item(i), "<SysRoot>", sWinDir, 1, -1, vbTextCompare)
    Next
    
    'LOAD ShellExecuteHooks (SEH) SAFELIST (O21)
    
    Dim colSafeSEH As Collection
    Set colSafeSEH = New Collection
        
    With colSafeSEH
        .Add "<SysRoot>\system32\shell32.dll"
    End With
    ReDim aSafeSEH(colSafeSEH.Count - 1)
    For i = 1 To colSafeSEH.Count
        aSafeSEH(i - 1) = Replace(colSafeSEH.Item(i), "<SysRoot>", sWinDir, 1, -1, vbTextCompare)
    Next
    
    
    'LOAD WINLOGON NOTIFY SAFELIST (O20)
    'second line added in HJT 1.99.2 final
    sSafeWinlogonNotify = "crypt32chain*cryptnet*cscdll*ScCertProp*Schedule*SensLogn*termsrv*wlballoon*igfxcui*AtiExtEvent*wzcnotif*" & _
                          "ActiveSync*atmgrtok*avldr*Caveo*ckpNotify*Command AntiVirus Download*ComPlusSetup*CwWLEvent*dimsntfy*DPWLN*EFS*FolderGuard*GoToMyPC*IfxWlxEN*igfxcui*IntelWireless*klogon*LBTServ*LBTWlgn*LMIinit*loginkey*MCPClient*MetaFrame*NavLogon*NetIdentity Notification*nwprovau*OdysseyClient*OPXPGina*PCANotify*pcsinst*PFW*PixVue*ppeclt*PRISMAPI.DLL*PRISMGNA.DLL*psfus*QConGina*RAinit*RegCompact*SABWinLogon*SDNotify*Sebring*STOPzilla*sunotify*SymcEventMonitors*T3Notify*TabBtnWL*Timbuktu Pro*tpfnf2*tpgwlnotify*tphotkey*VESWinlogon*WB*WBSrv*WgaLogon*wintask*WLogon*WRNotifier*Zboard*zsnotify*sclgntfy"
    
    sSafeIfeVerifier = "vrfcore.dll*vfbasics.dll*vfcompat.dll*vfluapriv.dll*vfprint.dll*vfnet.dll*vfntlmless.dll*vfnws.dll*vfcuzz.dll"
    
    'Loading Safe DNS list
    'https://www.comss.ru/list.php?c=securedns
    
    With colSafeDNS
        .Add "Google Public DNS", "8.8.8.8"
        .Add "Google Public DNS", "8.8.4.4"
        '2001:4860:4860::8888 IPv6
        '2001:4860:4860::8844 IPv6
        .Add "Verisign Public DNS", "64.6.64.6"
        .Add "Verisign Public DNS", "64.6.65.6"
        .Add "SkyDNS", "193.58.251.251"
        .Add "Cisco OpenDNS", "208.67.222.222"
        .Add "Cisco OpenDNS", "208.67.220.220"
        .Add "Norton ConnectSafe", "199.85.126.10"
        .Add "Norton ConnectSafe", "199.85.127.10"
        .Add "Norton ConnectSafe", "199.85.126.20"
        .Add "Norton ConnectSafe", "199.85.127.20"
        .Add "Norton ConnectSafe", "199.85.126.30"
        .Add "Norton ConnectSafe", "199.85.127.30"
        .Add "Norton ConnectSafe", "198.153.192.1"
        .Add "Norton ConnectSafe", "198.153.194.1"
        .Add "Norton ConnectSafe", "198.153.192.40"
        .Add "Norton ConnectSafe", "198.153.194.40"
        .Add "Norton ConnectSafe", "198.153.192.50"
        .Add "Norton ConnectSafe", "198.153.194.50"
        .Add "Norton ConnectSafe", "198.153.192.60"
        .Add "Norton ConnectSafe", "198.153.194.60"
        .Add "Adguard DNS", "176.103.130.130"
        .Add "Adguard DNS", "176.103.130.131"
        .Add "Yandex.DNS", "77.88.8.8"
        .Add "Yandex.DNS", "77.88.8.1"
        .Add "Yandex.DNS", "77.88.8.88"
        .Add "Yandex.DNS", "77.88.8.2"
        .Add "Yandex.DNS", "77.88.8.7"
        .Add "Yandex.DNS", "77.88.8.3"
        .Add "Comodo Secure DNS", "8.26.56.26"
        .Add "Comodo Secure DNS", "8.20.247.20"
        .Add "Verizon / Level 3 Communications", "209.244.0.3"
        .Add "Verizon / Level 3 Communications", "209.244.0.4"
        .Add "Verizon / Level 3 Communications", "4.2.2.1"
        .Add "Verizon / Level 3 Communications", "4.2.2.2"
        .Add "Verizon / Level 3 Communications", "4.2.2.3"
        .Add "Verizon / Level 3 Communications", "4.2.2.4"
        .Add "Verizon / Level 3 Communications", "4.2.2.5"
        .Add "Verizon / Level 3 Communications", "4.2.2.6"
        .Add "DNS.WATCH", "84.200.69.80"
        .Add "DNS.WATCH", "84.200.70.40"
        .Add "SafeDNS", "195.46.39.39"
        .Add "SafeDNS", "195.46.39.40"
        .Add "Dyn", "216.146.35.35"
        .Add "Dyn", "216.146.36.36"
        .Add "FreeDNS", "37.235.1.174"
        .Add "FreeDNS", "37.235.1.177"
        .Add "Alternate DNS", "198.101.242.72"
        .Add "Alternate DNS", "23.253.163.53"
        .Add "Rejector", "95.154.128.32"
        .Add "Rejector", "78.46.36.8"
        .Add "SmartViper Public DNS", "208.76.50.50"
        .Add "SmartViper Public DNS", "208.76.51.51"
        .Add "DNS Advantage / UltraDNS", "156.154.70.1"
        .Add "DNS Advantage / UltraDNS", "156.154.71.1"
        .Add "GreenTeamDNS", "81.218.119.11"
        .Add "GreenTeamDNS", "209.88.198.133"
        .Add "GTE", "192.76.85.133"
        .Add "GTE", "206.124.64.1"
    End With
    
    With colSafeCert
        .Add "Microsoft Enforced Licensing Registration Authority CA (SHA1)", "FA6660A94AB45F6A88C0D7874D89A863D74DEE97"
        .Add "DigiNotar Services 1024 CA", "F8A54E03AADC5692B850496A4C4630FFEAA29D83"
        .Add "login.yahoo.com", "D018B62DC518907247DF50925BB09ACF4A5CB3AD"
        .Add "login.live.com", "CEA586B2CE593EC7D939898337C57814708AB2BE"
        .Add "DigiNotar Root CA", "C060ED44CBD881BD0EF86C0BA287DDCF8167478C"
        .Add "DigiNotar Cyber CA", "B86E791620F759F17B8D25E38CA8BE32E7D5EAC2"
        .Add "DigiNotar PKIoverheid CA Overheid", "B533345D06F64516403C00DA03187D3BFEF59156"
        .Add "DigiNotar Cyber CA", "9845A431D51959CAF225322B4A4FE9F223CE6D15"
        .Add "Digisign Server ID - (Enrich)", "8E5BD50D6AE686D65252F843A9D4B96D197730AB"
        .Add "DigiNotar Root CA", "86E817C81A5CA672FE000F36F878C19518D6F844"
        .Add "login.yahoo.com", "80962AE4D6C5B442894E95A13E4A699E07D694CF"
        .Add "Microsoft Corporation", "7D7F4414CCEF168ADF6BF40753B5BECD78375931"
        .Add "mail.google.com", "6431723036FD26DEA502792FA595922493030F97"
        .Add "login.yahoo.com", "63FEAE960BAA91E343CE2BD8B71798C76BDB77D0"
        .Add "Microsoft Corporation", "637162CC59A3A1E25956FA5FA8F60D2E1C52EAC6"
        .Add "global trustee", "61793FCBFA4F9008309BBA5FF12D2CB29CD4151A"
        .Add "DigiNotar PKIoverheid CA Organisatie - G2", "5DE83EE82AC5090AEA9D6AC4E7A6E213F946E179"
        .Add "Digisign Server ID (Enrich)", "51C3247D60F356C7CA3BAF4C3F429DAC93EE7B74"
        .Add "login.skype.com", "471C949A8143DB5AD5CDF1C972864A2504FA23C9"
        .Add "DigiNotar Root CA G2", "43D9BCB568E039D073A74A71D8511F7476089CC3"
        .Add "DigiNotar PKIoverheid CA Overheid en Bedrijven", "40AA38731BD189F9CDB5B9DC35E2136F38777AF4"
        .Add "Microsoft Enforced Licensing Intermediate PCA", "3A850044D8A195CD401A680C012CB0A3B5F8DC08"
        .Add "DigiNotar Root CA", "367D4B3B4FCBBC0B767B2EC0CDB2A36EAB71A4EB"
        .Add "addons.mozilla.org", "305F8BD17AA2CBC483A4C41B19A39A0C75DA39D6"
        .Add "DigiNotar Cyber CA", "2B84BFBB34EE2EF949FE1CBE30AA026416EB2216"
        .Add "Microsoft Enforced Licensing Intermediate PCA", "2A83E9020591A55FC6DDAD3FB102794C52B24E70"
        .Add "www.google.com", "1916A2AF346D399F50313C393200F14140456616"
    End With
    
    With colBadCert
        .Add "Comodo", "03D22C9C66915D58C88912B64C1F984B8344EF09"
        .Add "F-Secure", "0F684EC1163281085C6AF20528878103ACEFCAAB"
        .Add "FRISK", "1667908C9E22EFBD0590E088715CC74BE4C60884"
        .Add "BitDefender", "18DEA4EFA93B06AE997D234411F3FD72A677EECE"
        .Add "GData", "2026D13756EB0DB753DF26CB3B7EEBE3E70BB2CF"
        .Add "Malwarebytes", "249BDA38A611CD746A132FA2AF995A2D3C941264"
        .Add "Symantec", "31AC96A6C17C425222C46D55C3CCA6BA12E54DAF"
        .Add "Trend Micro", "331E2046A1CCA7BFEF766724394BE6112B4CA3F7"
        .Add "Webroot", "3353EA609334A9F23A701B9159E30CB6C22D4C59"
        .Add "SUPERAntiSpyware", "373C33726722D3A5D1EDD1F1585D5D25B39BEA1A"
        .Add "Kaspersky", "3850EDD77CC74EC9F4829AE406BBF9C21E0DA87F"
        .Add "AVG", "3D496FA682E65FC122351EC29B55AB94F3BB03FC"
        .Add "PC Tools", "4243A03DB4C3C15149CEA8B38EEA1DA4F26BD159"
        .Add "K7 Computing", "42727E052C0C2E1B35AB53E1005FD9EDC9DE8F01"
        .Add "Doctor Web", "4420C99742DF11DD0795BC15B7B0ABF090DC84DF"
        .Add "Emsisoft", "4C0AF5719009B7C9D85C5EAEDFA3B7F090FE5FFF"
        .Add "Checkpoint Software", "5240AB5B05D11B37900AC7712A3C6AE42F377C8C"
        .Add "Emsisoft", "5DD3D41810F28B2A13E9A004E6412061E28FA48D"
        .Add "K7 Computing", "7457A3793086DBB58B3858D6476889E3311E550E"
        .Add "Bullguard", "76A9295EF4343E12DFC5FE05DC57227C1AB00D29"
        .Add "McAfee", "775B373B33B9D15B58BC02B184704332B97C3CAF"
        .Add "Comodo", "872CD334B7E7B3C3D1C6114CD6B221026D505EAB"
        .Add "McAfee", "88AD5DFE24126872B33175D1778687B642323ACF"
        .Add "Adaware", "9132E8B079D080E01D52631690BE18EBC2347C1E"
        .Add "Safer Networking", "982D98951CF3C0CA2A02814D474A976CBFF6BDB1"
        .Add "Webroot", "9A08641F7C5F2CCA0888388BE3E5DBDDAAA3B361"
        .Add "ThreatTrack Security", "9C43F665E690AB4D486D4717B456C5554D4BCEB5"
        .Add "CurioLab", "9E3F95577B37C74CA2F70C1E1859E798B7FC6B13"
        .Add "Avira", "A1F8DCB086E461E2ABB4B46ADCFA0B48C58B6E99"
        .Add "BullGuard", "A5341949ABE1407DD7BF7DFE75460D9608FBC309"
        .Add "ESET", "A59CC32724DD07A6FC33F7806945481A2D13CA2F"
        .Add "AVG", "AB7E760DA2485EA9EF5A6EEE7647748D4BA6B947"
        .Add "AVAST", "AD4C5429E10F4FF6C01840C20ABA344D7401209F"
        .Add "Symantec", "AD96BB64BA36379D2E354660780C2067B81DA2E0"
        .Add "Malwarebytes", "B8EBF0E696AF77F51C96DB4D044586E2F4F8FD84"
        .Add "Trend Micro", "CDC37C22FE9272D8F2610206AD397A45040326B8"
        .Add "Kaspersky", "D3F78D747E7C5D6D3AE8ABFDDA7522BFB4CBD598"
        .Add "ThreatTrack Security", "DB303C9B61282DE525DC754A535CA2D6A9BD3D87"
        .Add "AVAST", "DB77E5CFEC34459146748B667C97B185619251BA"
        .Add "Total Defense", "E22240E837B52E691C71DF248F12D27F96441C00"
        .Add "AVG Technologies", "E513EAB8610CFFD7C87E00BCA15C23AAB407FCEF"
        .Add "BitDefender", "ED841A61C0F76025598421BC1B00E24189E68D54"
        .Add "ESET", "F83099622B4A9F72CB5081F742164AD1B8D048C9"
        .Add "Panda", "FBB42F089AF2D570F2BF6F493D107A3255A9BB1A"
        .Add "Doctor Web", "FFFA650F2CB2ABC0D80527B524DD3F9FC172C138"
    End With
    
    AppendErrorLogCustom "LoadStuff - End"
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "modMain_LoadStuff"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub StartScan()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "StartScan - Begin"
    
    If bDebugToFile Then
        If hDebugLog = 0 Then OpenDebugLogHandle
    End If
    
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    
    bScanMode = True
    
    SetPriorityAllThreads GetCurrentProcess(), THREAD_PRIORITY_HIGHEST
    
    frmMain.txtNothing.Visible = False
    'frmMain.shpBackground.Tag = iItems
    SetProgressBar g_HJT_Items_Count   'R + F + O26
    
    Call GetProcesses(gProcess)
    
    Dim i&
    'load ignore list
    IsOnIgnoreList ""
    
    frmMain.lstResults.Clear
    
    'Registry
    
    UpdateProgressBar "R"
    For i = 0 To UBound(sRegVals)
        ProcessRuleReg sRegVals(i)
    Next i
    
    CheckR3Item
    CheckR4Item
    
    UpdateProgressBar "F"
    'File
    For i = 0 To UBound(sFileVals)
        If sFileVals(i) <> "" Then
            CheckFileItems sFileVals(i)
        End If
    Next i
    
    'Netscape/Mozilla stuff
    'CheckNetscapeMozilla        'N1-4
    
    
    'Other options
    UpdateProgressBar "O1"
    CheckO1Item
    CheckO1Item_ICS
    CheckO1Item_DNSApi
    UpdateProgressBar "O2"
    CheckO2Item
    UpdateProgressBar "O3"
    CheckO3Item
    UpdateProgressBar "O4"
    CheckO4Item
    UpdateProgressBar "O5"
    CheckO5Item
    UpdateProgressBar "O6"
    CheckO6Item
    UpdateProgressBar "O7"
    CheckO7Item
    UpdateProgressBar "O8"
    CheckO8Item
    UpdateProgressBar "O9"
    CheckO9Item
    UpdateProgressBar "O10"
    CheckO10Item
    UpdateProgressBar "O11"
    CheckO11Item
    UpdateProgressBar "O12"
    CheckO12Item
    UpdateProgressBar "O13"
    CheckO13Item
    UpdateProgressBar "O14"
    CheckO14Item
    UpdateProgressBar "O15"
    CheckO15Item
    UpdateProgressBar "O16"
    CheckO16Item
    UpdateProgressBar "O17"
    CheckO17Item
    UpdateProgressBar "O18"
    CheckO18Item
    UpdateProgressBar "O19"
    CheckO19Item
    UpdateProgressBar "O20"
    CheckO20Item
    UpdateProgressBar "O21"
    CheckO21Item
    UpdateProgressBar "O22"
    CheckO22Item
    UpdateProgressBar "O23"
    CheckO23Item
    'added in HJT 1.99.2: Desktop Components
    UpdateProgressBar "O24"
    CheckO24Item
    '2.0.7 - WMI Events
    UpdateProgressBar "O25"
    CheckO25Item
    UpdateProgressBar "O26"
    CheckO26Item
    UpdateProgressBar "ProcList"
    
    With frmMain
        .lblMD5.Visible = False
        '.lblInfo(1).Visible = True
        '.picPaypal.Visible = True
        If .lstResults.ListCount > 0 Then
            If bAutoSelect Then
                For i = 0 To .lstResults.ListCount - 1
                    .lstResults.Selected(i) = True
                Next i
            End If
            .txtNothing.Visible = False
            .cmdFix.Enabled = True
            .cmdSaveDef.Enabled = True
        Else
            .txtNothing.Visible = True
            .cmdFix.Enabled = False
            .cmdSaveDef.Enabled = False
        End If
    End With
    
    bScanMode = False
    SectionOutOfLimit "", bErase:=True
    
    Dim sEDS_Time   As String
    Dim OSData      As String
    
    If bDebugMode Or bDebugToFile Then
    
        If ObjPtr(OSver) <> 0 Then
                OSData = OSver.Bitness & " " & OSver.OSName & " (" & OSver.Edition & "), " & _
                    OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & ", " & _
                    "Service Pack: " & OSver.SPVer & "" & IIf(OSver.IsSafeBoot, " (Safe Boot)", "")
        End If
    
        sEDS_Time = vbCrLf & vbCrLf & "Logging is finished." & vbCrLf & vbCrLf & AppVer & vbCrLf & vbCrLf & OSData & vbCrLf & vbCrLf & _
                "Time spent: " & ((GetTickCount() - Perf.StartTime) \ 1000) & " sec." & vbCrLf & vbCrLf & _
                "Whole EDS function: " & Format$(tim(0).GetTime, "##0.000 sec.") & vbCrLf & _
                "CryptCATAdminAcquireContext: " & Format$(tim(1).GetTime, "##0.000 sec.") & vbCrLf & _
                "CryptCATAdminCalcHashFromFileHandle: " & Format$(tim(2).GetTime, "##0.000 sec.") & vbCrLf & _
                "CryptCATAdminEnumCatalogFromHash: " & Format$(tim(3).GetTime, "##0.000 sec.") & vbCrLf & _
                "WinVerifyTrust: " & Format$(tim(4).GetTime, "##0.000 sec.") & vbCrLf & _
                "GetSignerInfo: " & Format$(tim(5).GetTime, "##0.000 sec.") & vbCrLf & _
                "Release: " & Format$(tim(6).GetTime, "##0.000 sec.") & vbCrLf & _
                "CryptCATEnumerateMember: " & Format$(tim(7).GetTime, "##0.000 sec.") & vbCrLf & vbCrLf
        
        AppendErrorLogCustom sEDS_Time
    End If
    
    If bDebugToFile Then
        If hDebugLog <> 0 Then
            'Append Header to the end and close debug log file
            Dim b() As Byte
            b = sEDS_Time & vbCrLf & vbCrLf
            PutW hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
        End If
    End If
    
    SetPriorityAllThreads GetCurrentProcess(), THREAD_PRIORITY_NORMAL
    
    AppendErrorLogCustom "StartScan - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_StartScan"
    bScanMode = False
    SetPriorityAllThreads GetCurrentProcess(), THREAD_PRIORITY_NORMAL
    If inIDE Then Stop: Resume Next
End Sub

Public Sub SetProgressBar(lMaxTags As Long)
    
    'ProgressBar label settings
    frmMain.lblStatus.Visible = True
    frmMain.lblStatus.Caption = ""
    frmMain.lblStatus.ForeColor = &HFFFF&   'Yellow
    frmMain.lblStatus.ZOrder 0 'on top
    frmMain.lblStatus.Left = 400
    
    'results label -> off
    frmMain.lblInfo(0).Visible = False
    
    'Logo -> off
    frmMain.pictLogo.Visible = False
    
    'program description label -> off
    frmMain.lblInfo(1).Visible = False
    
    frmMain.shpBackground.Visible = True
    With frmMain.shpProgress
        .Tag = "0"
        .Visible = True
    End With
    frmMain.shpProgress.Width = 255 ' default
    frmMain.shpProgress.ZOrder 1
    frmMain.shpBackground.ZOrder 1
    g_ProgressMaxTags = lMaxTags
End Sub

Public Sub CloseProgressbar()
    frmMain.shpBackground.Visible = False
    'frmMain.lblInfo(0).Visible = True
    frmMain.lblInfo(1).Visible = True
    frmMain.shpProgress.Visible = False
    frmMain.lblStatus.Visible = False
    If Not TaskBar Is Nothing Then TaskBar.SetProgressState frmMain.hwnd, TBPF_NOPROGRESS
End Sub

Public Sub UpdateProgressBar(Section As String, Optional sAppendText As String)
    On Error GoTo ErrorHandler:
    
    Dim lTag As Long
    
    'If bAutoLogSilent Then Exit Sub
    
    With frmMain
    
        If Not IsNumeric(.shpProgress.Tag) Then .shpProgress.Tag = "0"
        lTag = .shpProgress.Tag
        If sAppendText = "" Then lTag = lTag + 1
        .shpProgress.Tag = lTag
        
        Select Case Section
            Case "R", "R0", "R1", "R2", "R3": .lblStatus.Caption = Translate(230) & "..."
            Case "F", "F1", "F2", "F3": .lblStatus.Caption = Translate(231) & "..."
            'Case 3: .lblStatus.Caption = Translate(232) & "..."
            Case "O1": .lblStatus.Caption = Translate(233) & "..."
            Case "O2": .lblStatus.Caption = Translate(234) & "..."
            Case "O3": .lblStatus.Caption = Translate(235) & "..."
            Case "O4": .lblStatus.Caption = Translate(236) & "..."
            Case "O5": .lblStatus.Caption = Translate(237) & "..."
            Case "O6": .lblStatus.Caption = Translate(238) & "..."
            Case "O7": .lblStatus.Caption = Translate(239) & "..."
            Case "O8": .lblStatus.Caption = Translate(240) & "..."
            Case "O9": .lblStatus.Caption = Translate(241) & "..."
            Case "O10": .lblStatus.Caption = Translate(242) & "..."
            Case "O11": .lblStatus.Caption = Translate(243) & "..."
            Case "O12": .lblStatus.Caption = Translate(244) & "..."
            Case "O13": .lblStatus.Caption = Translate(245) & "..."
            Case "O14": .lblStatus.Caption = Translate(246) & "..."
            Case "O15": .lblStatus.Caption = Translate(247) & "..."
            Case "O16": .lblStatus.Caption = Translate(248) & "..."
            Case "O17": .lblStatus.Caption = Translate(249) & "..."
            Case "O18": .lblStatus.Caption = Translate(250) & "..."
            Case "O19": .lblStatus.Caption = Translate(251) & "..."
            Case "O20": .lblStatus.Caption = Translate(252) & "..."
            Case "O21": .lblStatus.Caption = Translate(253) & "..."
            Case "O22": .lblStatus.Caption = Translate(254) & "..."
            Case "O23": .lblStatus.Caption = Translate(255) & "..."
            Case "O24": .lblStatus.Caption = Translate(257) & "..."
            Case "O25": .lblStatus.Caption = Translate(258) & "..."
            Case "O26": .lblStatus.Caption = Translate(261) & "..."
            
            Case "ProcList": .lblStatus.Caption = Translate(260) & "..."
            Case "Backup":   .lblStatus.Caption = Translate(259) & "...": .shpProgress.Width = 255
            Case "Report":   .lblStatus.Caption = Translate(262) & "..."
            Case "Finish":   .lblStatus.Caption = Translate(256): .shpProgress.Width = .shpBackground.Width + .shpBackground.Left - .shpProgress.Left
        End Select
        
        If Len(sAppendText) <> 0 Then .lblStatus.Caption = .lblStatus.Caption & " - " & sAppendText
        
        Select Case Section
            Case "ProcList": Exit Sub
            Case "Backup": Exit Sub
            Case "Finish": Exit Sub
        End Select
        
        If lTag > g_ProgressMaxTags Then lTag = g_ProgressMaxTags
        
        If g_ProgressMaxTags <> 0 Then
            .shpProgress.Width = .shpBackground.Width * (lTag / g_ProgressMaxTags)  'g_ProgressMaxTags = items to check or fix -1
            SetTaskBarProgressValue frmMain, (lTag / g_ProgressMaxTags)
        End If
        
        '.lblStatus.Refresh
        '.Refresh
    End With
    
    If Not bAutoLogSilent Then DoEvents
    
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "modMain_UpdateProgressBar", "shpProgress.Tag=", frmMain.shpProgress.Tag
    If inIDE Then Stop: Resume Next
End Sub


'CheckR0item
'CheckR1item
'CheckR2item
Private Sub ProcessRuleReg(ByVal sRule$)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ProcessRuleReg - Begin", "Rule: " & sRule
    
    Dim vRule As Variant, iMode&, bIsNSBSD As Boolean, Result As SCAN_RESULT
    Dim sHit$, sKey$, sParam$, sData$, sDefDataStrings$, Wow6432Redir As Boolean, UseWow
    Dim bProxyEnabled As Boolean, hHive As ENUM_REG_HIVE
    
    'Registry rule syntax:
    '[regkey],[regvalue],[infected data],[default data]
    '* [regkey]           = "" -> abort - no way man!
    ' * [regvaluStrings     = "" -> delete entire key
    '  * [default data]   = "" -> delete value
    '   * [infected data] = "" -> any value (other than default) is considered infected
    vRule = Split(sRule, ",")
    
    ' iMode = 0 -> check if value is infected
    ' iMode = 1 -> check if value is present
    ' iMode = 2 -> check if regkey is present
    If CStr(vRule(0)) = vbNullString Then Exit Sub
    If CStr(vRule(3)) = vbNullString Then iMode = 0
    If CStr(vRule(2)) = vbNullString Then iMode = 1
    If CStr(vRule(1)) = vbNullString Then iMode = 2
    
    sKey = vRule(0)
    sParam = vRule(1)
    If sParam = "(Default)" Then sParam = vbNullString
    sDefDataStrings = vRule(2)
       
    'Initialize hives enumerator
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HE.AddKey sKey
    
    Do While HE.MoveNext
    
        Wow6432Redir = HE.Redirected
        sKey = HE.Key
        hHive = HE.Hive
    
        Select Case iMode
        
        Case 0 'check for incorrect value
            sData = Reg.GetString(hHive, sKey, sParam, Wow6432Redir)
            sData = UnQuote(EnvironW(sData))
            
            If Not inArraySerialized(sData, sDefDataStrings, "|", , , 1) Then
                bIsNSBSD = False
                If bIgnoreSafeDomains And Not bIgnoreAllWhitelists Then bIsNSBSD = StrBeginWithArray(sData, aSafeRegDomains)
                If Not bIsNSBSD Then
                    If InStr(1, sData, "%2e", 1) > 0 Then sData = UnEscape(sData)
                    
                    sHit = IIf(bIsWin32, "R0 - ", IIf(Wow6432Redir, "R0-32 - ", "R0 - ")) & _
                        HE.KeyAndHive & "," & IIf(sParam = "", "(default)", sParam) & " = " & sData  'doSafeURLPrefix
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "R0"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, hHive, sKey, sParam, SplitSafe(sDefDataStrings, "|")(0), CLng(Wow6432Redir)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
            
        Case 1  'check for present value
            sData = Reg.GetString(hHive, sKey, sParam, Wow6432Redir)
            If 0 <> Len(sData) Then
                'check if domain is on safe list
                bIsNSBSD = False
                If bIgnoreSafeDomains And Not bIgnoreAllWhitelists Then bIsNSBSD = StrBeginWithArray(sData, aSafeRegDomains)
                'make hit
                If Not bIsNSBSD Then
                    If InStr(1, sData, "%2e", 1) > 0 Then sData = UnEscape(sData)
                    
                    If sParam = "ProxyServer" Then
                        bProxyEnabled = (Reg.GetDword(hHive, sKey, "ProxyEnable", Wow6432Redir) = 1)
                        
                        sHit = IIf(bIsWin32, "R1 - ", IIf(Wow6432Redir, "R1-32 - ", "R1 - ")) & _
                          HE.KeyAndHive & "," & IIf(sParam = "", "(default)", sParam) & IIf(sData <> "", " = " & sData, "") & IIf(bProxyEnabled, " (enabled)", " (disabled)")
                    Else
                        sHit = IIf(bIsWin32, "R1 - ", IIf(Wow6432Redir, "R1-32 - ", "R1 - ")) & _
                          HE.KeyAndHive & "," & IIf(sParam = "", "(default)", sParam) & IIf(sData <> "", " = " & sData, "") 'doSafeURLPrefix
                    End If
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "R1"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, hHive, sKey, sParam, , CLng(Wow6432Redir)
                            If sParam = "ProxyServer" Then
                                AddRegToFix .Reg, RESTORE_VALUE, hHive, sKey, "ProxyEnable", 0, CLng(Wow6432Redir), REG_RESTORE_DWORD
                            End If
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
            
        Case 2 'check if regkey is present
            If Reg.KeyExists(hHive, sKey, Wow6432Redir) Then
            
                sHit = IIf(bIsWin32, "R2 - ", IIf(Wow6432Redir, "R2-32 - ", "R2 - ")) & HE.KeyAndHive
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "R2"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, hHive, sKey, , , CLng(Wow6432Redir)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
            End If
        End Select
    Loop
    
    AppendErrorLogCustom "ProcessRuleReg - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_ProcessRuleReg", "sRule=", sRule
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixRegItem(sItem$, Result As SCAN_RESULT)
    'R0 - HKCU\Software\..\Main,Window Title
    'R1 - HKCU\Software\..\Main,Window Title=MSIE 5.01
    'R2 - HKCU\Software\..\Main
    FixRegistryHandler Result
End Sub


'CheckR3item
Public Sub CheckR3Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckR3Item - Begin"

    Dim sURLHook$, hKey&, i&, sName$, sHit$, sCLSID$, sFile$, Result As SCAN_RESULT, lRet&
    Dim bHookMising As Boolean, sDefHookDll$, sDefHookCLSID$
    
    sURLHook = "Software\Microsoft\Internet Explorer\URLSearchHooks"
    
    sDefHookCLSID = "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}"

    If OSver.MajorMinor >= 5.2 Then 'XP x64 +
        sDefHookDll = sWinSysDir & "\ieframe.dll"
    Else
        sDefHookDll = sWinSysDir & "\shdocvw.dll"
    End If
    
    HE.Init HE_HIVE_HKCU Or HE_HIVE_HKU, HE_SID_USER, HE_REDIR_NO_WOW
    HE.AddKey sURLHook
    
    Do While HE.MoveNext
        bHookMising = False
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_QUERY_VALUE, hKey) = 0 Then
            i = 0
            sCLSID = String$(MAX_VALUENAME, 0&)
            If RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&) <> 0 Then
                sHit = "R3 - " & HE.HiveNameAndSID & ": Default URLSearchHook is missing"
                bHookMising = True
                RegCloseKey hKey
            End If
        Else
            sHit = "R3 - " & HE.HiveNameAndSID & ": Default URLSearchHook is missing"
            bHookMising = True
        End If
    
        If bHookMising Then
            If Not IsOnIgnoreList(sHit) Then
                With Result
                    .Section = "R3"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sDefHookCLSID, vbNullString, , REG_RESTORE_SZ
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID, "", "Microsoft Url Search Hook"
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID & "\InProcServer32", "", sDefHookDll
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID & "\InProcServer32", "ThreadingModel", "Apartment"
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
        End If
    Loop
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HE.AddKey sURLHook
    
    Do While HE.MoveNext
        
        lRet = RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey)
        
        If lRet = 0 Then
        
          sCLSID = String$(MAX_VALUENAME, 0&)
          i = 0
          Do While 0 = RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&)
            
            sCLSID = TrimNull(sCLSID)
            
            GetFileByCLSID sCLSID, sFile, sName, HE.Redirected, HE.SharedKey
            
            If Not (sCLSID = sDefHookCLSID And StrComp(sFile, sDefHookDll, 1) = 0) Then
                
                sHit = IIf(bIsWin32, "R3 - ", IIf(HE.Redirected, "R3-32 - ", "R3 - ")) & HE.HiveNameAndSID & "\..\URLSearchHooks: " & _
                    sName & " - " & sCLSID & " - " & sFile
                    
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "R3"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
            
            i = i + 1
            sCLSID = String$(MAX_VALUENAME, 0&)
          Loop
          RegCloseKey hKey
        End If
    Loop
    
    AppendErrorLogCustom "CheckR3Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckR3Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixR3Item(sItem$, Result As SCAN_RESULT)
    'R3 - Shitty search hook - {00000000} - c:\windows\bho.dll"
    'R3 - Default URLSearchHook is missing
    
    FixRegistryHandler Result
End Sub

Public Sub CheckR4Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckR4Item - Begin"
    
    'http://ijustdoit.eu/changing-default-search-provider-in-internet-explorer-11-using-group-policies/
    
    'SearchScope
    'R4 - DefaultScope:
    'R4 - SearchScopes:
    
    Dim Result As SCAN_RESULT, sHit$, i&, j&, sDefScope$, sURL$, sProvider$, aScopes() As String, sHive$, vKey, aHives() As String
    Dim aKey() As String, aDes() As String, sValue$, sParam$, sDefData$, aData$(), sBuf As String, Param As Variant
    
    GetHives aHives, bIncludeServiceSID:=False
    
    'DefaultScope should NOT be here:
    'HKCU\Software\Policies\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Policies\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Microsoft\Internet Explorer\SearchScopes
    
    'DefaultScope should be here:
    'HKCU\Software\Microsoft\Internet Explorer\SearchScopes
     
    ReDim aKey((UBound(aHives) + 1) * 2 - 1)
    ReDim aDes((UBound(aHives) + 1) * 2 - 1)
    
    'preparing keys for checking and its descriptions
    For i = 0 To UBound(aHives)
        aKey(j) = aHives(i) & "\Software\Microsoft\Internet Explorer\SearchScopes"
        aDes(j) = aHives(i)
        j = j + 1
    Next
    For i = 0 To UBound(aHives)
        aKey(j) = aHives(i) & "\Software\Policies\Microsoft\Internet Explorer\SearchScopes"
        aDes(j) = aHives(i) & "\Software\Policies"
        j = j + 1
    Next
    
    'Checking if 'DefaultScope' for each user points to 'Bing'
    
    For i = 0 To UBound(aKey)
        sHive = Left$(aKey(i), 4)
        sDefScope = Reg.GetString(0&, aKey(i), "DefaultScope")
        
        If sDefScope <> "" Then
            
            For Each Param In Array("URL", "SuggestionsURL_JSON", "SuggestionsURL", "SuggestionsURLFallback", "TopResultURL", "TopResultURLFallback")
                sURL = Reg.GetString(0&, aKey(i) & "\" & sDefScope, CStr(Param))
                If sURL <> "" Then Exit For
            Next
            If sURL = "" Then sURL = "(no URL)"
            
            sProvider = Reg.GetString(0&, aKey(i) & "\" & sDefScope, "DisplayName")
            If sProvider = "" Then sProvider = "(no name)"
            If Left$(sProvider, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sProvider)
                If 0 <> Len(sBuf) Then sProvider = sBuf
            End If
            
            sHit = "R4 - " & aKey(i) & ": DefaultScope = " & sDefScope & " - " & sProvider & " - " & sURL
            
            If Not IsBingScopeKeyPara("URL", sURL) Then
              If Not IsOnIgnoreList(sHit) Then
                With Result
                    .Section = "R4"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, 0, aKey(i), "DefaultScope", "{0633EE93-D776-472f-A0FF-E1416B8B2E3A}"
                    .CureType = REGISTRY_BASED
                    AddRegToFix .Reg, BACKUP_KEY, 0, aKey(i) & "\" & sDefScope
                End With
                AddToScanResults Result
              End If
            End If
            
        End If
    Next
    
    'Checking consistency of default scope
    
    'HKCU\Software\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Microsoft\Internet Explorer\SearchScopes
    
    If OSver.MajorMinor >= 6 Then
      For i = 1 To cReg4vals.Count
        aData = Split(cReg4vals.Item(i), ",", 3)
        sHive = aData(0)
        sParam = aData(1)
        sDefData = aData(2)
        
            sValue = Reg.GetString(0&, sHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", sParam)
            
            'If Not inArraySerialized(sValue, sDefData, "|", , , vbTextCompare) And Not (sValue = "") Then
            If Not IsBingScopeKeyPara(sParam, sValue) And Len(sValue) > 0 Then

                sHit = "R4 - " & sHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}: " & sParam & " = " & sValue
                
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "R4"
                        .HitLineW = sHit
                        AddRegToFix .Reg, RESTORE_VALUE, 0, sHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", _
                          sParam, SplitSafe(sDefData, "|")(0)
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
      Next
    End If
    
    'Enum custom scopes
    '
    'HKCU\Software\Policies\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Policies\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Microsoft\Internet Explorer\SearchScopes
    'HKCU\Software\Microsoft\Internet Explorer\SearchScopes
    
    Dim sLastURL As String
    Dim sParams As String
    
    For i = 0 To UBound(aKey)
        sHive = Left$(aKey(i), 4)
        
        For j = 1 To Reg.EnumSubKeysToArray(0&, aKey(i), aScopes())
          If StrComp(aScopes(j), "{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", 1) <> 0 Then
            
            sProvider = Reg.GetString(0&, aKey(i) & "\" & aScopes(j), "DisplayName")
            If sProvider = "" Then sProvider = "(no name)"
            If Left$(sProvider, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sProvider)
                If 0 <> Len(sBuf) Then sProvider = sBuf
            End If
            
            sParams = ""
            sLastURL = ""
            
            For Each Param In Array("URL", "SuggestionsURL_JSON", "SuggestionsURL", "SuggestionsURLFallback", "TopResultURL", "TopResultURLFallback")
            
              sURL = Reg.GetString(0&, aKey(i) & "\" & aScopes(j), CStr(Param))
              
              If Len(sURL) <> 0 Or Reg.ValueExists(0&, aKey(i) & "\" & aScopes(j), CStr(Param)) Then
                
                sHit = "R4 - " & aKey(i) & "\" & aScopes(j) & " - " & sProvider & " - " '& sURL
                
                If Not IsBingScopeKeyPara("URL", sURL) Then
                  
                    With Result
                        .Section = "R4"
                        '.HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, 0, aKey(i) & "\" & aScopes(j)
                        'AddRegToFix .Reg, REMOVE_VALUE, 0, aKey(i) & "\" & aScopes(j), CStr(Param)
                        .CureType = REGISTRY_BASED
                    End With
                  
                    If sLastURL = "" Then 'first time?
                        sLastURL = sURL
                        sParams = CStr(Param)
                    Else
                        If sURL = sLastURL Then 'same URL ?
                            'Save several same URLs in one line by adding the list of param names at the end
                            sParams = sParams & IIf(sParams <> "", ",", "") & CStr(Param)
                        Else 'new URL ?
                            'save last result and flush
                            sHit = sHit & sLastURL & " (" & sParams & ")"
                            
                            If Not IsOnIgnoreList(sHit) Then
                                Result.HitLineW = sHit
                                AddToScanResults Result, , True
                            End If
                            
                            sLastURL = sURL
                            sParams = CStr(Param)
                        End If
                    End If
                End If
              End If
            Next
            
            If sParams <> "" Then
                sHit = "R4 - " & aKey(i) & "\" & aScopes(j) & " - " & sProvider & " - " & sLastURL & " (" & sParams & ")"
                
                If Not IsOnIgnoreList(sHit) Then
                    Result.HitLineW = sHit
                    AddToScanResults Result
                End If
            End If
            
          End If
        Next
    Next
    
    AppendErrorLogCustom "CheckR4Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckR4Item"
    If inIDE Then Stop: Resume Next
End Sub

Private Function IsBingScopeKeyPara(sRegParam As String, sURL As String) As Boolean
    If sURL = "" Then Exit Function
    
    'Is valid domain
    Dim pos As Long, sPrefix As String
    pos = InStr(sURL, "?")
    If pos = 0 Then Exit Function
    sPrefix = Left$(sURL, pos - 1)
    Select Case sPrefix
    Case "http://www.bing.com/search"
    Case "http://www.bing.com/as/api/qsml"
    Case "http://api.bing.com/qsml.aspx"
    Case "http://search.live.com/results.aspx"
    Case Else
        Exit Function
    End Select

    Dim aKey() As String, aVal() As String, i As Long
    Dim bSearchTermPresent As Boolean
    
    IsBingScopeKeyPara = True
    
    Call ParseKeysURL(sURL, aKey, aVal)
    
    Select Case UCase$(sRegParam)
    
        Case "URL", UCase$("SuggestionsURL"), UCase$("SuggestionsURLFallback"), UCase$("TopResultURL"), UCase$("TopResultURLFallback")
            If IsArrDimmed(aKey) Then
                For i = 0 To UBound(aKey)
                    Select Case LCase(aKey(i))
                    Case "q", "query"
                    '{searchTerms}
                        If aVal(i) = "{searchTerms}" Then bSearchTermPresent = True
                    
                    Case "src"
                    'IE-SearchBox
                    'IE11TR
                    '{referrer:source?}
                    'IE10TR
                    'src=ie9tr
                    'IE-TopResult 'for TopResultURL
                        If StrBeginWith(aVal(i), "IE") Then
                            If Len(aVal(i)) > 6 Then
                                If StrComp(aVal(i), "IE-SearchBox", 1) = 0 Then
                                ElseIf StrComp(aVal(i), "IE-TopResult", 1) = 0 Then
                                Else
                                    IsBingScopeKeyPara = False
                                End If
                            End If
                        ElseIf StrComp(aVal(i), "{referrer:source?}", 1) = 0 Then
                        Else
                            IsBingScopeKeyPara = False
                        End If
                    
                    Case "form"
                    'IE8SRC
                    'IE10SR
                    'IE11SR
                    'IESR02
                    'IESS02
                    'IE8SSC
                    'IE11SS
                    'IETR02 'for TopResultURL
                    'IE11TR 'for TopResultURL
                    'IE10TR 'for TopResultURL
                    'SKY2DF
                    'PRHPR1
                    'MSERBM
                    'IE8SRC
                    'HPNTDF
                    'MSSEDF
                    'APBTDF
                    'SK216DF
                        If Len(aVal(i)) > 7 Then IsBingScopeKeyPara = False
                    
                    Case "pc"
                    'HRTS
                    'MSERT1
                    'HPNTDF
                    'MSE1
                    'MAPB
                    'MAARJS
                    'MAMIJS;
                    'CMDTDFJS
                        If Len(aVal(i)) > 8 Then IsBingScopeKeyPara = False
                    
                    Case "maxwidth"
                    '{ie:maxWidth}
                        If StrComp(aVal(i), "{ie:maxWidth}", 1) <> 0 Then IsBingScopeKeyPara = False
                    
                    Case "rowheight"
                    '{ie:rowHeight}
                        If StrComp(aVal(i), "{ie:rowHeight}", 1) <> 0 Then IsBingScopeKeyPara = False
                    
                    Case "sectionheight"
                    '{ie:sectionHeight}
                        If StrComp(aVal(i), "{ie:sectionHeight}", 1) <> 0 Then IsBingScopeKeyPara = False
                    
                    Case "market"
                    '{language}
                        If StrComp(aVal(i), "{language}", 1) <> 0 Then IsBingScopeKeyPara = False
                    
                    Case ""
                        If Len(aVal(i)) > 0 Then IsBingScopeKeyPara = False
                        
                    Case Else
                        IsBingScopeKeyPara = False
                    End Select
                Next
            End If
        
        Case Else
            IsBingScopeKeyPara = False
    End Select
    
    If Not bSearchTermPresent Then IsBingScopeKeyPara = False
End Function

Public Sub FixR4Item(sItem$, Result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    
    Dim sFixHive As String, sParam As String, sHive As String, sDefData As String, j As Long, aData() As String, i&
    Dim Param As Variant, sURL As String
    
    FixRegistryHandler Result
    
'    If AryPtr(Result.Reg) <> 0 Then
'        For Each Param In Array("URL", "SuggestionsURL_JSON", "SuggestionsURL", "SuggestionsURLFallback", "TopResultURL", "TopResultURLFallback")
'
'            sURL = Reg.GetString(0&, Result.Reg(0).Key, CStr(Param))
'            If Len(sURL) <> 0 Then Exit For
'        Next
'    End If
'
'    'if all URL params was deleted we must delete parent key too
'    If Len(sURL) = 0 Then
'        If AryPtr(Result.Reg) <> 0 Then
'            BackupKey Result, 0&, Result.Reg(0).Key
'            Reg.DelKey 0&, Result.Reg(0).Key
'        End If
'    End If
    
    For i = 0 To UBound(Result.Reg)
      With Result.Reg(i)
        
        'if resetting DefaultScope, set defaults for all 'Bing' scope values
        If InStr(1, .Param, "DefaultScope", 1) <> 0 Then
            If .Hive <> 0 Then
                sFixHive = Reg.GetShortHiveName(Reg.GetHiveNameByHandle(.Hive))
            Else
                sFixHive = Reg.GetHiveName(.Key, bIncludeSID:=True)
            End If
            
            'moved to BACKUP_KEY verb
            'BackupKey Result, 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}"
            
            For j = 1 To cReg4vals.Count
                aData = Split(cReg4vals.Item(j), ",", 3)
                sHive = aData(0)
                sParam = aData(1)
                sDefData = SplitSafe(aData(2), "|")(0)
        
                If (sFixHive = "HKLM" And sHive = "HKLM") Or _
                    (sFixHive <> "HKLM" And sHive = "HKCU") Then
                
                    Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", sParam, sDefData
                End If
            Next
            
            Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes", "DefaultScope", "{0633EE93-D776-472f-A0FF-E1416B8B2E3A}"
            Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "DisplayName", "Bing"
            If sFixHive = "HKLM" Then
                Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "", "Bing"
            Else
                Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "FaviconURL", "http://www.bing.com/favicon.ico"
                Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "FaviconURLFallback", "http://www.bing.com/favicon.ico"
                Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTLogoPath", AppDataLocalLow & "\Microsoft\Internet Explorer\Services\"
                Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTLogoURL", "http://go.microsoft.com/fwlink/?LinkID=403856&language={language}&scale={scalelevel}&contrast={contrast}"
                Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTSuggestionsURL", "http://api.bing.com/qsml.aspx?query={searchTerms}&market={language}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IENTSS"
                Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTTopResultURL", "http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTTR"
                Reg.SetStringVal 0&, sFixHive & "\Software\Microsoft\Internet Explorer\SearchScopes\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTURL", "http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTSR"
            End If
        End If
        
      End With
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixR4Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckFileItems(ByVal sRule$)
    On Error GoTo ErrorHandler:
    
    Dim vRule As Variant, iMode&, sHit$, Result As SCAN_RESULT
    Dim sFile$, sSection$, sParam$, sData$, sLegitData$
    Dim sTmp$
    
    AppendErrorLogCustom "CheckFileItems - Begin", "Rule: " & sRule
    
    'IniFile rule syntax:
    '[inifile],[section],[value],[default data],[infected data]
    '* [inifile]          = "" -> abort
    ' * [section]         = "" -> abort
    '  * [value]          = "" -> abort
    '   * [default data]  = "" -> delete if found
    '    * [infected data]= "" -> fix if infected
    
    'decrypt rule
    'sRule = Crypt(sRule)
    
    'Checking white list rules
    '1-st token should contains .ini
    'total number of tokens should be 5 (0 to 4)
    
    vRule = Split(sRule, ";")
    If UBound(vRule) <> 4 Or InStr(CStr(vRule(0)), ".ini") = 0 Then
        If Not bAutoLogSilent Then
            MsgBoxW "CheckFileItems: Spelling error or decrypting error for: " & sRule
        End If
        Exit Sub
    End If
    
    '1,2,3 tokens should not be empty
    '4-th token is empty -> check if value is present     (F1)
    '4-th token is present -> check if value is infected  (F0)
    
    'File checking rules:
    '
    'example:
    '--------------
    '1. system.ini    (file)
    '2. boot          (section)
    '3. Shell         (parameter)
    '4. explorer.exe  (data / value)
    
    If CStr(vRule(0)) = vbNullString Then Exit Sub
    If CStr(vRule(1)) = vbNullString Then Exit Sub
    If CStr(vRule(2)) = vbNullString Then Exit Sub
    If CStr(vRule(4)) = vbNullString Then iMode = 0
    If CStr(vRule(3)) = vbNullString Then iMode = 1
    
    sFile = vRule(0)
    sSection = vRule(1)
    sParam = vRule(2)
    sLegitData = vRule(3)
    
    'Registry checking rules (prefix REG: on 1-st token)
    '
    'example:
    '1. REG:system.ini ()
    '2. boot           (section)
    '3. Shell          (parameter)
    '4. explorer.exe   (data / value)
    
    'if 4-th token is empty -> check if value is present, in the Registry      (F3)
    'if 4-th token is present -> check if value is infected, in the Registry   (F2)
    
'    ' adding char "," to each value 'UserInit'
'    If InStr(1, sLegitData, "UserInit", 1) <> 0 Then
'        arr = Split(sLegitData, "|")
'        For i = 0 To UBound(arr)
'            sTmp = sTmp & arr(i) & ",|"
'        Next
'        sTmp = Left$(sTmp, Len(sTmp) - 1)
'        sLegitData = sLegitData & "|" & sTmp
'    End If
    
    If Left$(sFile, 3) = "REG" Then
        'skip Win9x
        If Not bIsWinNT Then Exit Sub
        If CStr(vRule(4)) = vbNullString Then iMode = 2
        If CStr(vRule(3)) = vbNullString Then iMode = 3
    End If
    
    'iMode:
    ' F0 = check if value is infected (file)
    ' F1 = check if value is present (file)
    ' F2 = check if value is infected, in the Registry
    ' F3 = check if value is present, in the Registry
    
    Select Case iMode
        Case 0
            'F0 = check if value is infected (file)
            'sValue = String$(255, " ")
            'GetPrivateProfileString CStr(vRule(1)), CStr(vRule(2)), "", sValue, 255, CStr(vRule(0))
            'sValue = Rtrim$(sValue)
            
            If Not FileExists(sFile) Then
                sFile = FindOnPath(sFile, True)
            End If
            
            sData = IniGetString(sFile, sSection, sParam)
            sData = RTrimNull(sData)
            
            If Not inArraySerialized(sData, sLegitData, "|", , , vbTextCompare) Then
                If bIsWinNT And Trim$(sData) <> vbNullString Then
                    sHit = "F0 - " & sFile & ": " & sParam & "=" & sData
                    If Not IsOnIgnoreList(sHit) Then
                        If bMD5 Then sHit = sHit & GetFileMD5(sData)
                        With Result
                            .Section = "F0"
                            .HitLineW = sHit
                            'system.ini
                            AddIniToFix .Reg, RESTORE_VALUE_INI, sFile, "boot", "shell", SplitSafe(sLegitData, "|")(0) '"explorer.exe"
                            .CureType = INI_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
            
        Case 1
            'F1 = check if value is present (file)
            'sValue = String$(255, " ")
            'GetPrivateProfileString CStr(vRule(1)), CStr(vRule(2)), "", sValue, 255, CStr(vRule(0))
            'sValue = Rtrim$(sValue)
            
            If Not FileExists(sFile) Then
                sFile = FindOnPath(sFile, True)
            End If
            
            sData = IniGetString(sFile, sSection, sParam)
            sData = RTrimNull(sData)
            
            If Trim$(sData) <> vbNullString Then
                sHit = "F1 - " & sFile & ": " & sParam & "=" & sData
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sData)
                    With Result
                        .Section = "F1"
                        .HitLineW = sHit
                        'win.ini
                        AddIniToFix .Reg, RESTORE_VALUE_INI, sFile, "windows", sParam, "" 'param = 'load' or 'run'
                        .CureType = INI_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
            
        Case 2
            'F2 = check if value is infected, in the Registry
            'so far F2 is only reg:Shell and reg:UserInit
            
            HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
            HE.AddKey "Software\Microsoft\Windows NT\CurrentVersion\WinLogon"
            
            Do While HE.MoveNext
                
                sData = Reg.GetString(HE.Hive, HE.Key, sParam, HE.Redirected)
                sTmp = sData
                If Right$(sData, 1) = "," Then sTmp = Left$(sTmp, Len(sTmp) - 1)
                
                'Note: HKCU + empty values are allowed
                If Not inArraySerialized(sTmp, sLegitData, "|", , , vbTextCompare) And _
                  Not ((HE.Hive = HKCU Or HE.Hive = HKU) And sData = "") Then
            
                    'exclude no WOW64 value on Win10 for UserInit
                    If Not (HE.Redirected And OSver.MajorMinor >= 10 And sParam = "UserInit" And sData = "") Then
                
                        sHit = IIf(bIsWin32, "F2 - ", IIf(HE.Redirected, "F2-32 - ", "F2 - ")) & sFile & ": " & HE.HiveNameAndSID & "\..\" & sParam & "=" & sData
                        If Not IsOnIgnoreList(sHit) Then
                            If bMD5 Then sHit = sHit & GetFileMD5(sData)
                            With Result
                                .Section = "F2"
                                .HitLineW = sHit
                                AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sParam, SplitSafe(sLegitData, "|")(0), HE.Redirected
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults Result
                        End If
                    End If
                End If
            Loop
            
        Case 3
            'F3 = check if value is present, in the Registry
            'this is not really smart when more INIFile items get
            'added, but so far F3 is only reg:load and reg:run
        
            HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
            HE.AddKey "Software\Microsoft\Windows NT\CurrentVersion\Windows"
            
            Do While HE.MoveNext
            
                sData = Reg.GetString(HE.Hive, HE.Key, sParam, HE.Redirected)
                If 0 <> Len(sData) Then
                    sHit = IIf(bIsWin32, "F3 - ", IIf(HE.Redirected, "F3-32 - ", "F3 - ")) & sFile & ": " & HE.HiveNameAndSID & "\..\" & sParam & "=" & sData
                    If Not IsOnIgnoreList(sHit) Then
                        If bMD5 Then sHit = sHit & GetFileMD5(sData)
                        With Result
                            .Section = "F3"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            Loop
    End Select
    
    AppendErrorLogCustom "CheckFileItems - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_ProcessRuleIniFile", "sRule=", sRule
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixFileItem(sItem$, Result As SCAN_RESULT)
    'F0 - system.ini: Shell=c:\win98\explorer.exe openme.exe
    'F1 - win.ini: load=hpfsch
    'F2, F3 - registry

    'coding is easy if you cheat :)
    '(c) Dragokas: Cheaters will be punished ^_^
    
    FixRegistryHandler Result
End Sub

Private Sub CheckO1Item_DNSApi()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item_DNSApi - Begin"
    
    If OSver.MajorMinor <= 5 Then Exit Sub 'XP+ only
    
    Const MaxSize As Long = 5242880 ' 5 MB.
    
    Dim vFile As Variant, ff As Long, Size As Currency, p As Long, buf() As Byte, sHit As String, Result As SCAN_RESULT
    Dim bufExample() As Byte
    Dim bufExample_2() As Byte
    
    bufExample = StrConv(LCase$("\drivers\etc\hosts"), vbFromUnicode)
    bufExample_2 = StrConv(UCase$("\drivers\etc\hosts"), vbFromUnicode)
    
    ToggleWow64FSRedirection False
    
    For Each vFile In Array(sWinDir & "\system32\dnsapi.dll", sWinDir & "\syswow64\dnsapi.dll")
    
        If OSver.Bitness = "x32" And InStr(1, vFile, "syswow64", 1) <> 0 Then Exit For

        If OpenW(CStr(vFile), FOR_READ, ff) Then
            
            Size = LOFW(ff)
            
            If Size > MaxSize Then
                ErrorMsg Err, "modMain_CheckO1Item_DNSApi", "File is too big: " & vFile & " (Allowed: " & MaxSize & " byte max., current is: " & Size & "byte.)"
            ElseIf Size > 0 Then
                
                ReDim buf(Size - 1)
                
                If GetW(ff, 1, , VarPtr(buf(0)), CLng(Size)) Then
                
                    p = InArrSign_NoCase(buf, bufExample, bufExample_2)
                    
                    If p = -1 Then                      '//TODO: add isMicrosoftFile() ?
                        ' if signature not found
                        sHit = "O1 - DNSApi: File is patched - " & vFile
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With Result
                                .Section = "O1"
                                .HitLineW = sHit
                                AddFileToFix .File, RESTORE_FILE_SFC, CStr(vFile)
                                .CureType = FILE_BASED
                            End With
                            AddToScanResults Result
                        End If
                    End If
                End If
            End If
            CloseW ff
        End If
    Next
    ToggleWow64FSRedirection True
    AppendErrorLogCustom "CheckO1Item_DNSApi - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO1Item_DNSApi"
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub

Private Function InArrSign(ArrSrc() As Byte, ArrEx() As Byte) As Long
    Dim i As Long, j As Long, p As Long, Found As Boolean
    InArrSign = -1
    For i = 0 To UBound(ArrSrc) - UBound(ArrEx)
        p = i
        Found = True
        For j = 0 To UBound(ArrEx)
            If ArrSrc(p) <> ArrEx(j) Then Found = False: Exit For
            p = p + 1
        Next
        If Found Then InArrSign = p - UBound(ArrEx) - 1: Exit For
    Next
End Function

Private Function InArrSign_NoCase(ArrSrc() As Byte, ArrEx() As Byte, ArrEx_2() As Byte) As Long
    'ArrEx - all lcase
    'ArrEx_2 - all Ucase
    Dim i As Long, j As Long, p As Long, Found As Boolean
    InArrSign_NoCase = -1
    For i = 0 To UBound(ArrSrc) - UBound(ArrEx)
        p = i
        Found = True
        For j = 0 To UBound(ArrEx)
            If ArrSrc(p) <> ArrEx(j) And ArrSrc(p) <> ArrEx_2(j) Then Found = False: Exit For
            p = p + 1
        Next
        If Found Then InArrSign_NoCase = p - UBound(ArrEx) - 1: Exit For
    Next
End Function

Private Sub CheckO1Item_ICS()
    ' hosts.ics
    'https://support.microsoft.com/ru-ru/kb/309642
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item_ICS - Begin"
    
    Dim sHostsFileICS$, sHit$, sHostsFileICS_Default$
    Dim sLines$, sLine As Variant, NonDefaultPath As Boolean, cFileSize As Currency, hFile As Long
    Dim Result As SCAN_RESULT
    
    If bIsWin9x Then sHostsFileICS_Default = sWinDir & "\hosts.ics"
    If bIsWinNT Then sHostsFileICS_Default = sWinDir & "\System32\drivers\etc\hosts.ics"
    
    sHostsFileICS = sHostsFile & ".ics"
    
    If StrComp(sHostsFileICS, sHostsFileICS_Default) <> 0 Then
        NonDefaultPath = True
    End If
    
    If NonDefaultPath Then                              'Note: \System32\drivers\etc is not under Wow6432 redirection
        ToggleWow64FSRedirection False, sHostsFileICS
    End If
    
    cFileSize = FileLenW(sHostsFileICS)
    
    ' Size = 0 or just not exists
    If cFileSize = 0 Then
        ToggleWow64FSRedirection True
        
        If NonDefaultPath Then
            GoTo CheckHostsICS_Default:
        Else
            Exit Sub
        End If
    End If
    
    If OpenW(sHostsFileICS, FOR_READ, hFile) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
        ToggleWow64FSRedirection True
    Else
    
        sHit = "O1 - Unable to read Hosts.ICS file"
        
        If Not IsOnIgnoreList(sHit) Then
            With Result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFileICS
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults Result
        End If

        ToggleWow64FSRedirection True
        If NonDefaultPath Then
            GoTo CheckHostsICS_Default:
        Else
            Exit Sub
        End If
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    For Each sLine In Split(sLines, vbLf)
        sLine = Replace$(sLine, vbTab, " ")
        sLine = Replace$(sLine, vbCr, "")
        sLine = Trim$(sLine)
        
        If sLine <> vbNullString Then
            If Left$(sLine, 1) <> "#" Then
                Do
                    sLine = Replace$(sLine, "  ", " ")
                Loop Until InStr(sLine, "  ") = 0
                
                sHit = "O1 - Hosts.ICS: " & sLine
                'If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit

                '// TODO: сделать чтобы каждая строка бекапилась отдельно.
                'точнее она и так бекапится отдельно, но нужно чтобы модуль резервного копирования умел восстанавливать
                'не целиком файл, а отдельные строки.
                'при этом необходимость бекапить файл целиком отпадёт (т.е. вот эти строки ниже нужно будет удалить и вернуть AddToScanResultsSimple)
                
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, sHostsFileICS
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults Result
                End If

            End If
        End If
    Next
    
CheckHostsICS_Default:
    
    ToggleWow64FSRedirection True

    If Not NonDefaultPath Then Exit Sub
    
    cFileSize = FileLenW(sHostsFileICS_Default)
    
    ' Size = 0 or just not exists
    If cFileSize = 0 Then Exit Sub
    
    If OpenW(sHostsFileICS_Default, FOR_READ, hFile) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
    Else
        sHit = "O1 - Unable to read Hosts.ICS default file"
        
        If Not IsOnIgnoreList(sHit) Then
            With Result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFileICS_Default
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults Result
        End If
        
        Exit Sub
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    For Each sLine In Split(sLines, vbLf)
        sLine = Replace$(sLine, vbTab, " ")
        sLine = Replace$(sLine, vbCr, "")
        sLine = Trim$(sLine)
        
        If sLine <> vbNullString Then
            If Left$(sLine, 1) <> "#" Then
                Do
                    sLine = Replace$(sLine, "  ", " ")
                Loop Until InStr(sLine, "  ") = 0
                
                sHit = "O1 - Hosts.ICS default: " & sLine
                'If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
                
                '// TODO: сделать чтобы каждая строка бекапилась отдельно.
                'точнее она и так бекапится отдельно, но нужно чтобы модуль резервного копирования умел восстанавливать
                'не целиком файл, а отдельные строки.
                'при этом необходимость бекапить файл целиком отпадёт (т.е. вот эти строки ниже нужно будет удалить и вернуть AddToScanResultsSimple)
                
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, sHostsFileICS_Default
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults Result
                End If
                
            End If
        End If
    Next
    
    AppendErrorLogCustom "CheckO1Item_ICS - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO1Item_ICS"
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub


Private Sub CheckO1Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item - Begin"
    
    Dim sHit$, i&, ff%, HostsDefaultFile$, NonDefaultPath As Boolean
    Dim sLine As Variant, sLines$, cFileSize@
    Dim aHits() As String, j As Long, hFile As Long
    Dim HostsDefaultPath As String
    ReDim aHits(0)
    Dim Result As SCAN_RESULT
    
    '// TODO: Add UTF8.
    'http://serverfault.com/questions/452268/hosts-file-ignored-how-to-troubleshoot
    
    Dbg "1"
    
    GetHosts
    
    If bIsWin9x Then HostsDefaultFile = sWinDir & "\hosts"
    If bIsWinNT Then HostsDefaultFile = sWinDir & "\System32\drivers\etc\hosts"
    
    Dbg "2"
    
    If StrComp(sHostsFile, HostsDefaultFile) <> 0 Then
        'sHit = "O1 - Hosts file is located at: " & sHostsFile
        sHit = "O1 - " & Translate(271) & ": " & sHostsFile
        If Not IsOnIgnoreList(sHit) Then
            With Result
                .Section = "O1"
                .HitLineW = sHit
                HostsDefaultPath = EnvironUnexpand(GetParentDir(HostsDefaultFile))
                AddRegToFix .Reg, RESTORE_VALUE, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", _
                  HostsDefaultPath, , REG_RESTORE_EXPAND_SZ
                .CureType = REGISTRY_BASED Or CUSTOM_BASED
            End With
            AddToScanResults Result
        End If
        NonDefaultPath = True
    End If
    
    Dbg "3"
    
    'If Not FileExists(sHostsFile) Then Exit Sub
    
'    On Error Resume Next
'    iAttr = GetFileAttributes(StrPtr(sHostsFile))
'    If (iAttr And 2048) Then iAttr = iAttr - 2048
'
'    SetFileAttributes StrPtr(sHostsFile), vbNormal
'    SetFileAttributes StrPtr(sHostsFile), vbArchive
'
'    If Err.Number And Not inIDE And Not bAutoLogSilent Then  ' tired to see this warning from IDE
'        MsgBoxW replace$(Translate(300), "[]", sHostsFile), vbExclamation
''        msgboxW "For some reason your system denied write " & _
''        "access to the Hosts file." & vbCrLf & "If any hijacked domains " & _
''        "are in this file, HiJackThis may NOT be able to " & _
''        "fix this." & vbCrLf & vbCrLf & "If that happens, you need " & _
''        "to edit the file yourself. To do this, click " & _
''        "Start, Run and type:" & vbCrLf & vbCrLf & _
''        "   notepad """ & sHostsFile & """" & vbCrLf & vbCrLf & _
''        "and press Enter. Find the line(s) HiJackThis " & _
''        "reports and delete them. Save the file as " & _
''        """hosts."" (with quotes), and reboot.", vbExclamation
'    End If
'    SetFileAttributes StrPtr(sHostsFile), iAttr
    
    If NonDefaultPath Then                              'Note: \System32\drivers\etc is not under Wow6432 redirection
        ToggleWow64FSRedirection False, sHostsFile
    End If
    
    Dbg "4"
    
    cFileSize = FileLenW(sHostsFile)
    
    If cFileSize = 0 Then
        If NonDefaultPath Then
            'Check default path also
            GoTo CheckHostsDefault:
        Else
            sHit = "O1 - Hosts: Reset contents to default"

            If Not IsOnIgnoreList(sHit) Then
                With Result
                    .Section = "O1"
                    .HitLineW = sHit
                    AddFileToFix .File, BACKUP_FILE, sHostsFile
                    .CureType = CUSTOM_BASED
                End With
                AddToScanResults Result
            End If
            
            ToggleWow64FSRedirection True
            Exit Sub
        End If
    End If
    
    Dbg "5"
    
    If OpenW(sHostsFile, FOR_READ, hFile) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
        ToggleWow64FSRedirection True
    Else
    
        sHit = "O1 - Unable to read Hosts file"
        
        If Not IsOnIgnoreList(sHit) Then
            With Result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFile
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults Result
        End If
        
        ToggleWow64FSRedirection True
        If NonDefaultPath Then
            GoTo CheckHostsDefault:
        Else
            Exit Sub
        End If
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    i = 0
    
    Dbg "6"
    
    For Each sLine In Split(sLines, vbLf)
            
            'ignore all lines that start with loopback
            '(127.0.0.1), null (0.0.0.0) and private IPs
            '(192.168. / 10.)
            sLine = Replace$(sLine, vbTab, " ")
            sLine = Replace$(sLine, vbCr, "")
            sLine = Trim$(sLine)
            
            If sLine <> vbNullString Then
                'If InStr(sLine, "127.0.0.1") <> 1 And _
                '   InStr(sLine, "0.0.0.0") <> 1 And _
                '   InStr(sLine, "192.168.") <> 1 And _
                '   InStr(sLine, "10.") <> 1 And _
                '   InStr(sLine, "#") <> 1 And _
                '   Not (bIgnoreSafeDomains And InStr(sLine, "216.239.37.101") > 0) Or _
                '   bIgnoreAllWhitelists Then
                    '216.239.37.101 = google.com
                    
                '::1 - default for Vista
                If Left$(sLine, 1) <> "#" And _
                  StrComp(sLine, "127.0.0.1       localhost", 1) <> 0 And _
                  StrComp(sLine, "::1             localhost", 1) <> 0 And _
                  StrComp(sLine, "127.0.0.1 localhost", 1) <> 0 Then
                  
                    Do
                        sLine = Replace$(sLine, "  ", " ")
                    Loop Until InStr(sLine, "  ") = 0
                    
                    sHit = "O1 - Hosts: " & sLine
                    If Not IsOnIgnoreList(sHit) Then
                        'AddToScanResultsSimple "O1", sHit
                        If UBound(aHits) < i Then ReDim Preserve aHits(UBound(aHits) + 100)
                        aHits(i) = sHit
                        i = i + 1
                    End If
                    
'                    If i = 10 And Not NonDefaultPath And Not bResetOptAdded Then
'                        sHit = "O1 - Hosts: Reset contents to default"
'                        If Not IsOnIgnoreList(sHit) Then
'                            frmMain.lstResults.AddItem sHit, frmMain.lstResults.ListCount - 10
'                            AddToScanResultsSimple "O1", sHit, DoNotAddToListBox:=True
'                        End If
'                        bResetOptAdded = True
'                    End If
                    
                    'I don't plan to fix Hosts file on hijacked location for now.
                    
'                    If i > 100 Then
'                        If Not bAutoLogSilent Then
'                            MsgBoxW replace$(Translate(302), "[]", sHostsFile), vbExclamation
''                           msgboxW "You have an particularly large " & _
''                            "amount of hijacked domains. It's probably " & _
''                            "better to delete the file itself then to " & _
''                            "fix each item (and create a backup)." & vbCrLf & _
''                            vbCrLf & "If you see the same IP address in all " & _
''                            "the reported O1 items, consider deleting your " & _
''                            "Hosts file, which is located at " & sHostsFile & _
''                           ".", vbExclamation
'                        End If
'                        'Close #ff
'                        ToggleWow64FSRedirection True
'                        Exit For
'                    End If
                End If
            End If
        'Loop
    Next
    'Close #ff

    Dbg "7"

    If i > 0 Then
        If i >= 10 Then
            If Not NonDefaultPath Then
                sHit = "O1 - Hosts: Reset contents to default"
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, sHostsFile
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
        End If
'        'maximum 100 hosts entries
'        If i <= 100 Then
'            For j = 0 To i - 1
'                AddToScanResultsSimple "O1", aHits(j)
'            Next
'        Else
'            sHit = "O1 - Hosts: has " & i & " entries"
'        End If
        For j = 0 To i - 1
        
            'AddToScanResultsSimple "O1", aHits(j), IIf((j < 20) Or (j > i - 1 - 20), False, True)
        
            '// TODO: сделать чтобы каждая строка бекапилась отдельно.
            'точнее она и так бекапится отдельно, но нужно чтобы модуль резервного копирования умел восстанавливать
            'не целиком файл, а отдельные строки.
            'при этом необходимость бекапить файл целиком отпадёт (т.е. вот эти строки ниже нужно будет удалить и вернуть AddToScanResultsSimple)
        
            sHit = aHits(j)
            With Result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFile
                .CureType = CUSTOM_BASED
            End With
            'limit for first and last 20 entries only to view on results window
            AddToScanResults Result, IIf((j < 20) Or (j > i - 1 - 20), False, True)
        Next
    End If
    
    ReDim aHits(0)

CheckHostsDefault:
    'if Hosts was redirected -> checking records on default hosts also. ( Prefix "O1 - Hosts default: " )
    
    i = 0
    
    ToggleWow64FSRedirection True
    
    Dbg "8"
    
    If NonDefaultPath Then
        
        If FileExists(HostsDefaultFile) Then
            
            cFileSize = FileLenW(HostsDefaultFile)
            If cFileSize <> 0 Then

                Dbg "9"

                If OpenW(HostsDefaultFile, FOR_READ, hFile) Then
                    sLines = String$(cFileSize, vbNullChar)
                    GetW hFile, 1, sLines
                    CloseW hFile
                Else
                    sHit = "O1 - Unable to read Default Hosts file"

                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "O1"
                            .HitLineW = sHit
                            AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                            .CureType = CUSTOM_BASED
                        End With
                        AddToScanResults Result
                    End If
                    
                    Exit Sub
                End If
                
                Dbg "10"
                
                sLines = Replace$(sLines, vbCrLf, vbLf)

                For Each sLine In Split(sLines, vbLf)
                
                    sLine = Replace$(sLine, vbTab, " ")
                    sLine = Replace$(sLine, vbCr, "")
                    sLine = Trim$(sLine)
                    
                    If sLine <> vbNullString Then
                    
                        If Left$(sLine, 1) <> "#" And _
                          StrComp(sLine, "127.0.0.1       localhost", 1) <> 0 And _
                          StrComp(sLine, "::1             localhost", 1) <> 0 Then    '::1 - default for Vista
                            Do
                                sLine = Replace$(sLine, "  ", " ")
                            Loop Until InStr(sLine, "  ") = 0
                    
                            Dbg "11"
                    
                            sHit = "O1 - Hosts default: " & sLine
                            If Not IsOnIgnoreList(sHit) Then
                                'AddToScanResultsSimple "O1", sHit
                                If UBound(aHits) < i Then ReDim Preserve aHits(UBound(aHits) + 100)
                                aHits(i) = sHit
                                i = i + 1
                            End If
                    
'                            If i = 10 And Not bResetOptAdded Then
'                                sHit = "O1 - Hosts default: Reset contents to default"
'                                If Not IsOnIgnoreList(sHit) Then
'                                    frmMain.lstResults.AddItem sHit, frmMain.lstResults.ListCount - 10
'                                    AddToScanResultsSimple "O1", sHit, DoNotAddToListBox:=True
'                                End If
'                                bResetOptAdded = True
'                            End If
'
'                            If i > 100 Then
'                                If Not bAutoLogSilent Then
'                                    If vbYes = MsgBoxW(replace$(Translate(302), "[]", sHostsFile), vbExclamation Or vbYesNo) Then
'                                        Shell "explorer.exe /select," & """" & sHostsFile & """", vbNormalFocus
'                                    End If
'        '                           msgboxW "You have an particularly large " & _
'        '                            "amount of hijacked domains. It's probably " & _
'        '                            "better to delete the file itself then to " & _
'        '                            "fix each item (and create a backup)." & vbCrLf & _
'        '                            vbCrLf & "If you see the same IP address in all " & _
'        '                            "the reported O1 items, consider deleting your " & _
'        '                            "Hosts file, which is located at " & sHostsFile & _
'        '                           "." & vbcrlf & vbcrlf & "Would you like to open its folder now?", vbExclamation or vbyesno
'                                End If
'                                Exit Sub
'                            End If
                        End If
                    End If
                Next
            End If
        End If
        
        Dbg "12"
        
        If i > 0 Then
            If i >= 10 Then
                sHit = "O1 - Hosts default: Reset contents to default"
                    
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
'            'maximum 100 hosts entries
'            If i <= 100 Then
'                 For j = 0 To i - 1
'                    AddToScanResultsSimple "O1", aHits(j)
'                 Next
'            Else
'                sHit = "O1 - Hosts default: has " & i & " entries"
'            End If
            For j = 0 To i - 1
                
                'AddToScanResultsSimple "O1", aHits(j), IIf((j < 20) Or (j > i - 1 - 20), False, True)
                
                '// TODO: сделать чтобы каждая строка бекапилась отдельно.
                'точнее она и так бекапится отдельно, но нужно чтобы модуль резервного копирования умел восстанавливать
                'не целиком файл, а отдельные строки.
                'при этом необходимость бекапить файл целиком отпадёт (т.е. вот эти строки ниже нужно будет удалить и вернуть AddToScanResultsSimple)
            
                sHit = aHits(j)
                With Result
                    .Section = "O1"
                    .HitLineW = sHit
                    AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                    .CureType = CUSTOM_BASED
                End With
                'limit for first and last 20 entries only to view on results window
                AddToScanResults Result, IIf((j < 20) Or (j > i - 1 - 20), False, True)
            Next
        End If
    End If

    AppendErrorLogCustom "CheckO1Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO1Item"
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO1Item(sItem$, Result As SCAN_RESULT)
    'O1 - Hijack of auto.search.msn.com etc with Hosts file
    On Error GoTo ErrorHandler:
    Dim sLine As Variant, sHijacker$, i&, iAttr&, ff1%, ff2%, HostsDefaultPath$, sLines$, HostsDefaultFile$, cFileSize@, sHosts$
    Dim sHostsTemp$, bResetHosts As Boolean, aLines() As String, isICS As Boolean, SFC As String
    
    If InStr(1, sItem, "O1 - DNSApi:", 1) <> 0 Then
        FixFileHandler Result
        Exit Sub
    End If
    
    If bIsWin9x Then HostsDefaultPath = sWinDir
    If bIsWinNT Then HostsDefaultPath = "%SystemRoot%\System32\drivers\etc"
    
    HostsDefaultFile = EnvironW(HostsDefaultPath & "\" & "hosts")
    
    'If InStr(sItem, "Hosts file is located at") > 0 Then
    If InStr(sItem, Translate(271)) > 0 Then
        Reg.SetExpandStringVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", HostsDefaultPath
        GetHosts    'reload var. 'sHostsFile'
        Exit Sub
    End If
    
    If StrComp(sItem, "O1 - Hosts: Reset contents to default", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts default: Reset contents to default", 1) = 0 Then
        bResetHosts = True
    End If
    
    If StrBeginWith(sItem, "O1 - Hosts default: ") Or bResetHosts Then
        
        sHosts = HostsDefaultFile   'default hosts path
    Else
        sHosts = sHostsFile         'path that may be redirected
    End If
    
    If StrBeginWith(sItem, "O1 - Hosts.ICS: ") Then
        sHosts = sHostsFile & ".ics"
        isICS = True
    ElseIf StrBeginWith(sItem, "O1 - Hosts.ICS default: ") Then
        sHosts = HostsDefaultFile & ".ics"
        isICS = True
    End If
    
    sHostsTemp = TempCU & "\" & "hosts.new"
    If Not CheckAccessWrite(sHostsTemp, True) Then
        sHostsTemp = BuildPath(AppPath(), "hosts.new")
    End If
    
    If FileExists(sHostsTemp) Then
        DeleteFileWEx StrPtr(sHostsTemp)
    End If
    
    If StrComp(GetParentDir(sHosts), sWinDir & "\System32\drivers\etc\hosts", 1) <> 0 Then
        ToggleWow64FSRedirection False
    End If
    
    cFileSize = FileLenW(sHosts)
    
    If cFileSize = 0 Or bResetHosts Then
        'no reset for ICS for now
        If isICS Then GoTo Finalize
        '2.0.7. - Reset Hosts to its default contents
        ff2 = FreeFile()
        Open sHostsTemp For Output As #ff2
            Print #ff2, GetDefaultHostsContents()
        Close #ff2
        GoTo Replace
    End If
    
    'If Not StrBeginWith(sItem, "O1 - Hosts: ") Then Exit Sub
    
    'parse to server name
    ' Example: 127.0.0.1 my.dragokas.com -> var. 'sHijacker' = "my.dragokas.com"
    sHijacker = Mid$(sItem, InStr(sItem, ":") + 2)
    sHijacker = Trim$(sHijacker)
    If Not isICS Then
        If InStr(sHijacker, " ") > 0 Then
            Dim sTemp$
            sTemp = Mid$(sHijacker, InStr(sHijacker, " ") + 1)
            If 0 <> Len(sTemp) Then sHijacker = sTemp
        End If
    End If
    
    'Reset attributes (and save old one in var. 'iAttr')
    iAttr = GetFileAttributes(StrPtr(sHosts))
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    SetFileAttributes StrPtr(sHosts), vbNormal
    
    BackupFile Result, sHosts
    
    'read current hosts file
    ff1 = FreeFile()
    Open sHosts For Binary Access Read As #ff1
    sLines = String$(LOF(ff1), 0)
    Get #ff1, , sLines
    Close #ff1
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    'build new hosts file (exclude bad lines)
    ff2 = FreeFile()
    Open sHostsTemp For Output As #ff2
        aLines = Split(sLines, vbLf)
          For i = 0 To UBoundSafe(aLines)
            sLine = aLines(i)
            sLine = Replace$(sLine, vbTab, " ")
            sLine = Replace$(sLine, vbCr, "")
            Do
                sLine = Replace$(sLine, "  ", " ")
            Loop Until InStr(sLine, "  ") = 0
            If InStr(1, sLine, sHijacker, 1) <> 0 Then
                'don't write line to hosts file
            Else
                'skip last empty line
                If 0 <> Len(sLine) Or (0 = Len(sLine) And i < UBound(aLines)) Then Print #ff2, aLines(i)
            End If
          Next
    Close #ff2
    
Replace:
    If DeleteFileForce(sHosts) Then
        
        If StrComp(GetParentDir(sHosts), sWinDir & "\System32\drivers\etc\hosts", 1) <> 0 Then
            ToggleWow64FSRedirection False
        End If
        
        If 0 = MoveFile(StrPtr(sHostsTemp), StrPtr(sHosts)) Then
            If Err.LastDllError = 5 Then Err.Raise 70
        End If
        'Recover old one attrib.
        SetFileAttributes StrPtr(sHosts), iAttr
    Else
        Err.Raise 70
    End If
    
    
    '//TODO:
    'clear cache
    
    '1. Mozilla Firefox
    '%LocalAppData%\Mozilla\Firefox\Profiles\<Name>\cache2 -> rename to *.bak
    
    '2. Microsoft Internet Explorer
    
    '3. Google Chrome
    
    '4. Yandex Browser
    
    '5.1. Opera Presto
    
    '5.2. (Chromo) Opera
    
    '6. Edge
    '...

Finalize:
    ToggleWow64FSRedirection True
    
    AppendErrorLogCustom "FixO1Item - End"
    Exit Sub
    
ErrorHandler:
    If Err.Number = 70 And Not bSeenHostsFileAccessDeniedWarning Then
        'permission denied
        MsgBoxW Translate(303), vbExclamation
'        msgboxW "HiJackThis could not write the selected changes to your " & _
'               "hosts file. The probably cause is that some program is " & _
'               "denying access to it, or that your user account doesn't have " & _
'               "the rights to write to it.", vbExclamation
        bSeenHostsFileAccessDeniedWarning = True
    Else
        ErrorMsg Err, "modMain_FixO1Item", "sItem=", sItem
    End If
    Close #ff1, #ff2
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FlushDNS()
    On Error GoTo ErrorHandler:
    If GetServiceRunState("dnscache") <> SERVICE_RUNNING Then StartService "dnscache"

    If Proc.ProcessRun(BuildPath(sSysNativeDir, "ipconfig.exe"), "/flushdns", , vbHide) Then
        Proc.WaitForTerminate , , , 15000
    End If
    
    RestartService "dnscache"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FlushDNS"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckO2Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO2Item - Begin"
    
    Dim hKey&, i&, sName$, sCLSID$, lpcName&, sFile$, sHit$, BHO_key$, Result As SCAN_RESULT
    Dim sBuf$, sProgId$, sProgId_CLSID$, bSafe As Boolean
    
    Dim HEFixKey As clsHiveEnum
    Dim HEFixValue As clsHiveEnum
    
    Set HEFixKey = New clsHiveEnum
    Set HEFixValue = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HEFixKey.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HEFixValue.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    
    'key to check
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects"
    
    'keys to fix + \{CLSID} placeholder
    HEFixKey.AddKey "HKCR\CLSID", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer\Extension Compatibility", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\Stats", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\Settings", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer", "\ApprovedExtensionsMigration{CLSID}"
    
    'values to fix (value == {CLSID})
    HEFixValue.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID"
    HEFixValue.AddKey "Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration"
    
    Do While HE.MoveNext
   
            If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
        
                i = 0
                Do
                    sCLSID = String$(MAX_KEYNAME, vbNullChar)
                    lpcName = Len(sCLSID)
                    If RegEnumKeyExW(hKey, i, StrPtr(sCLSID), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
                    
                    sCLSID = Left$(sCLSID, lstrlen(StrPtr(sCLSID)))
                    
                    If sCLSID <> "" And Not StrBeginWith(sCLSID, "MSHist") Then
                        
                        BHO_key = HE.KeyAndHive & "\" & sCLSID
                        
                        If InStr(sCLSID, "}}") > 0 Then
                            'the new searchwww.com trick - use a double
                            '}} in the IE toolbar registration, reg the toolbar
                            'with only one } - IE ignores the double }}, but
                            'HT didn't. It does now!
                            sCLSID = Left$(sCLSID, Len(sCLSID) - 1)
                        End If
                        
                        'get filename from HKCR\CLSID\sName + BHO name
                        
                        'get bho name from BHO regkey
                        sName = Reg.GetString(0&, BHO_key, vbNullString, HE.Redirected)
                        If HE.SharedKey And sName = "" Then
                            sName = Reg.GetString(0&, BHO_key, vbNullString, Not HE.Redirected)
                        End If
                        'get BHO name from CLSID regkey
                        If sName = "" Then
                            GetFileByCLSID sCLSID, sFile, sName, HE.Redirected, HE.SharedKey
                        Else
                            GetFileByCLSID sCLSID, sFile, , HE.Redirected, HE.SharedKey
                        End If
                        
                        sProgId = Reg.GetString(HKEY_CLASSES_ROOT, "Clsid\" & sCLSID & "\ProgID", vbNullString, HE.Redirected)
                        If sProgId = "" And HE.SharedKey Then
                            sProgId = Reg.GetString(HKEY_CLASSES_ROOT, "Clsid\" & sCLSID & "\ProgID", vbNullString, Not HE.Redirected)
                        End If
                        
                        If sProgId <> "" Then
                            'safety check
                            sProgId_CLSID = Reg.GetString(HKEY_CLASSES_ROOT, sProgId & "\Clsid", vbNullString, HE.Redirected)
                            If sProgId_CLSID = "" And HE.SharedKey Then
                                sProgId_CLSID = Reg.GetString(HKEY_CLASSES_ROOT, sProgId & "\Clsid", vbNullString, Not HE.Redirected)
                            End If
                            
                            If sProgId_CLSID <> sCLSID Then
                                sProgId = ""
                            End If
                        End If
                        
                        sHit = IIf(bIsWin32, "O2", IIf(HE.Redirected, "O2-32", "O2")) & _
                            " - " & HE.HiveNameAndSID & "\..\BHO: " & sName & " - " & sCLSID & " - " & sFile
                        
                        bSafe = False
                        If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                            If IsMicrosoftFile(sFile) Then bSafe = True
                        End If

                        If Not IsOnIgnoreList(sHit) And (Not bSafe) Then
                            If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                            
                            With Result
                                .Section = "O2"
                                .HitLineW = sHit
                                
                                AddRegToFix .Reg, REMOVE_KEY, 0, BHO_key, , , IIf(HE.SharedKey, REG_REDIRECTION_BOTH, REG_NOTREDIRECTED)
                    
                                If 0 <> Len(sProgId) Then
                                    AddRegToFix .Reg, REMOVE_KEY, HKCR, sProgId, , , IIf(HE.SharedKey, REG_REDIRECTION_BOTH, REG_NOTREDIRECTED)
                                End If
                                
                                HEFixKey.Repeat
                                Do While HEFixKey.MoveNext
                                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, Replace$(HE.Key, "{CLSID}", sCLSID), , , HE.Redirected
                                Loop
                                
                                HEFixValue.Repeat
                                Do While HEFixValue.MoveNext
                                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                                Loop
                                
                                AddFileToFix .File, REMOVE_FILE Or UNREG_DLL, sFile
                                
                                .CureType = REGISTRY_BASED Or FILE_BASED
                            End With
                            AddToScanResults Result
                        End If
                    End If
                    i = i + 1
                Loop
                RegCloseKey hKey
            End If
    Loop
    
    Set HEFixKey = Nothing
    Set HEFixValue = Nothing
    
    AppendErrorLogCustom "CheckO2Item - End"
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO2Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO2Item(sItem$, Result As SCAN_RESULT)
    'O2 - Enumeration of existing MSIE BHO's
    'O2 - BHO: AcroIEHlprObj Class - {00000...000} - C:\PROGRAM FILES\ADOBE\ACROBAT 5.0\ACROBAT\ACTIVEX\ACROIEHELPER.OCX
    'O2 - BHO: ... (no file)
    'O2 - BHO: ... c:\bla.dll (file missing)
    
    On Error GoTo ErrorHandler:
    
    Dim bIE_Exist As Boolean
    
    bIE_Exist = ProcessExist("iexplore.exe", True)
    
    If Not bShownBHOWarning And bIE_Exist Then
        MsgBoxW Translate(310), vbExclamation
'        msgboxW "HiJackThis is about to remove a " & _
'               "BHO and the corresponding file from " & _
'               "your system. Close all Internet " & _
'               "Explorer windows AND all Windows " & _
'               "Explorer windows before continuing for " & _
'               "the best chance of success.", vbExclamation
        bShownBHOWarning = True
    End If
    
    If bIE_Exist Then
        If MsgBox(Translate(311), vbExclamation) = vbYes Then
            'Internet Explorer still run. Would you like HJT close IE forcibly?
            'WARNING: current browser session will be lost!
            Proc.ProcessClose ProcessName:="iexplore.exe", Async:=False, TimeOutMs:=1000, SendCloseMsg:=True
        End If
    End If
    
    '//TODO: Add:
    'HKLM\SOFTWARE\WOW6432NODE\MICROSOFT\INTERNET EXPLORER\LOW RIGHTS\ELEVATIONPOLICY\{CLSID}
    'HKLM\SOFTWARE\CLASSES\APPID\{Name}
    'HKLM\SOFTWARE\CLASSES\APPID\{GUID}
    'HKLM\SOFTWARE\WOW6432NODE\CLASSES\APPID\{Name}
    'HKLM\SOFTWARE\WOW6432NODE\CLASSES\APPID\{GUID}
    'HKLM\SOFTWARE\CLASSES\INTERFACE\{GUID}
    'HKLM\SOFTWARE\CLASSES\TYPELIB\{GUID}
    
    'file should go first bacause it can use reg. info for its dll unregistration.
    FixFileHandler Result
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO2Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO3Item()
    'HKLM\Software\Microsoft\Internet Explorer\Toolbar
    'HKLM\Software\Microsoft\Internet Explorer\Explorer Bars
  
    '//TODO:
    'Add handling of:
    'Locked value: http://www.tweaklibrary.com/windows/Software_Applications/Internet-Explorer/27/Unlock-the-Internet-Explorer-toolbars/11245/index.htm
    'Explorer, ShellBrowser and subkeys with ITBarLayout (ITBar7Layout, ITBar7Layout64) values: https://support.microsoft.com/en-us/help/555460
    'BackBitmapIE5 value (need ???): https://msdn.microsoft.com/en-us/library/aa753592(v=vs.85).aspx
    '
    'Detailed description: http://www.winblog.ru/admin/1147761976-ippon_170506_02.html
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO3Item - Begin"
    
    Dim hKey&, i&, sCLSID$, sName$, Result As SCAN_RESULT
    Dim sFile$, sHit$, SearchwwwTrick As Boolean, sBuf$, sProgId$, sProgId_CLSID$
    Dim bSafe As Boolean
    
    Dim HEFixKey As clsHiveEnum
    Dim HEFixValue As clsHiveEnum
    
    Set HEFixKey = New clsHiveEnum
    Set HEFixValue = New clsHiveEnum
    
    Dim aKeys(1) As String
    Dim aDescr(1) As String
    
    'keys to check
    aKeys(0) = "Software\Microsoft\Internet Explorer\Toolbar"
    aKeys(1) = "Software\Microsoft\Internet Explorer\Explorer Bars"
    
    aDescr(0) = "Toolbar"
    aDescr(1) = "Explorer Bars"
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HEFixKey.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HEFixValue.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    
    HE.AddKeys aKeys
    
    'keys to fix + placeholder
    HEFixKey.AddKey "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\Stats", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\Settings", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer", "\ApprovedExtensionsMigration{CLSID}"
    
    'values to fix (value == {CLSID})
    HEFixValue.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID"
    HEFixValue.AddKey "Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration"
    
    Do While HE.MoveNext
        
            If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
    
                i = 0
                Do
                    sCLSID = String$(MAX_VALUENAME, 0)
                    ReDim uData(MAX_VALUENAME)
        
                    'enumerate MSIE toolbars / Explorer Bars
                    If RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&) <> 0 Then Exit Do
                    sCLSID = TrimNull(sCLSID)
        
                    If InStr(sCLSID, "}}") > 0 Then
                        'the new searchwww.com trick - use a double
                        '}} in the IE toolbar registration, reg the toolbar
                        'with only one } - IE ignores the double }}, but
                        'HJT didn't. It does now!
            
                        sCLSID = Left$(sCLSID, Len(sCLSID) - 1)
                        SearchwwwTrick = True
                    Else
                        SearchwwwTrick = False
                    End If
        
                    'found one? then check corresponding HKCR key
                    GetFileByCLSID sCLSID, sFile, sName, HE.Redirected, HE.SharedKey
        
                    '   sCLSID <> "BrandBitmap" And _
                    '   sCLSID <> "SmBrandBitmap" And _
                    '   sCLSID <> "BackBitmap" And _
                    '   sCLSID <> "BackBitmapIE5" And _
                    '   sCLSID <> "OLE (Part 1 of 5)" And _
                    '   sCLSID <> "OLE (Part 2 of 5)" And _
                    '   sCLSID <> "OLE (Part 3 of 5)" And _
                    '   sCLSID <> "OLE (Part 4 of 5)" And _
                    '   sCLSID <> "OLE (Part 5 of 5)" Then
        
                    sProgId = Reg.GetString(HKEY_CLASSES_ROOT, "Clsid\" & sCLSID & "\ProgID", vbNullString, HE.Redirected)
                    If sProgId = "" And HE.SharedKey Then
                        sProgId = Reg.GetString(HKEY_CLASSES_ROOT, "Clsid\" & sCLSID & "\ProgID", vbNullString, Not HE.Redirected)
                    End If
        
                    If 0 <> Len(sProgId) Then
                        'safe check
                        sProgId_CLSID = Reg.GetString(HKEY_CLASSES_ROOT, sProgId & "\Clsid", vbNullString, HE.Redirected)
                        If sProgId_CLSID = "" And HE.SharedKey Then
                            sProgId_CLSID = Reg.GetString(HKEY_CLASSES_ROOT, sProgId & "\Clsid", vbNullString, Not HE.Redirected)
                        End If
                        
                        If sProgId_CLSID <> sCLSID Then
                            sProgId = ""
                        End If
                    End If
                    
                    bSafe = False
                    
                    If OSver.MajorMinor = 5 Then 'Win2k
                        If WhiteListed(sFile, sWinDir & "\system32\msdxm.ocx") Then bSafe = True
                    End If
                    
                    If 0 <> Len(sName) And InStr(sCLSID, "{") > 0 And Not bSafe Then
        
        '          If Not SearchwwwTrick Or _
        '            (SearchwwwTrick And (sCLSID <> "BrandBitmap" And sCLSID <> "SmBrandBitmap")) Then
        
                        sHit = IIf(bIsWin32, "O3", IIf(HE.Redirected, "O3-32", "O3")) & _
                            " - " & HE.HiveNameAndSID & "\..\" & aDescr(HE.KeyIndex) & ": " & sName & " - " & sCLSID & " - " & sFile
                        
                        If Not IsOnIgnoreList(sHit) Then
                            If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                            With Result
                                .Section = "O3"
                                .HitLineW = sHit
                                
                                AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                                AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\WebBrowser", sCLSID, , HE.Redirected
                                AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\ShellBrowser", sCLSID, , HE.Redirected
                                
                                If 0 <> Len(sProgId) Then
                                    AddRegToFix .Reg, REMOVE_VALUE, 0, "HKCR\" & sProgId, sCLSID, , IIf(HE.SharedKey, REG_REDIRECTION_BOTH, HE.Redirected)
                                End If
                                
                                HEFixKey.Repeat
                                Do While HEFixKey.MoveNext
                                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, Replace$(HE.Key, "{CLSID}", sCLSID), , , HE.Redirected
                                Loop
                                
                                HEFixValue.Repeat
                                Do While HEFixValue.MoveNext
                                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                                Loop
                                
                                AddFileToFix .File, REMOVE_FILE Or UNREG_DLL, sFile
                                
                                .CureType = REGISTRY_BASED Or FILE_BASED
                            End With
                            AddToScanResults Result
                        End If
'                     End If
                    End If
                    i = i + 1
                Loop
                RegCloseKey hKey
            End If
    Loop
    
    Set HEFixKey = Nothing
    Set HEFixValue = Nothing
    
    AppendErrorLogCustom "CheckO3Item - End"
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO3Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO3Item(sItem$, Result As SCAN_RESULT)
    'O3 - Enumeration of existing MSIE toolbars

    FixFileHandler Result
    FixRegistryHandler Result
End Sub


'returns array of SID strings, except of current user
Sub GetUserNamesAndSids(aSID() As String, aUser() As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetUserNamesAndSids - Begin"
    
    'get all users' SID and map it to the corresponding username
    'not all users visible in User Accounts screen have a SID though,
    'they probably get this when logging in for the first time

    Dim CurUserName$, i&, K&, sUsername$, aTmpSID() As String, aTmpUser() As String

    CurUserName = GetUser()
    
    aTmpSID = SplitSafe(Reg.EnumSubKeys(HKEY_USERS, vbNullString), "|")
    ReDim aTmpUser(UBound(aTmpSID))
    For i = 0 To UBound(aTmpSID)
        If (StrComp(aTmpSID(i), ".DEFAULT") = 0) Or ((aTmpSID(i) Like "S-#-#-#*") And Not StrEndWith(aTmpSID(i), "_Classes")) Then
            sUsername = MapSIDToUsername(aTmpSID(i))
            If 0 = Len(sUsername) Then sUsername = "?"
            If StrComp(sUsername, CurUserName, 1) <> 0 Then
                aTmpUser(i) = sUsername
            Else
                'filter current user key with HKCU
                aTmpSID(i) = ""
                aTmpUser(i) = ""
            End If
        Else
            aTmpSID(i) = ""
            aTmpUser(i) = ""
        End If
    Next i
    
    'compress array
    K = 0
    ReDim aSID(UBound(aTmpSID))
    ReDim aUser(UBound(aTmpSID))
    
    For i = 0 To UBound(aTmpSID)
        If 0 <> Len(aTmpSID(i)) Then
            aSID(K) = aTmpSID(i)
            aUser(K) = aTmpUser(i)
            K = K + 1
        End If
    Next
    If K > 0 Then
        ReDim Preserve aSID(K - 1)
        ReDim Preserve aUser(K - 1)
    Else
        ReDim Preserve aSID(0)
        ReDim Preserve aUser(0)
    End If
    
    AppendErrorLogCustom "GetUserNamesAndSids - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.GetUserNamesAndSids"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_RegRuns()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4_RegRuns - Begin"
    
    Const BAT_LENGTH_LIMIT As Long = 300&
    
    Dim aRegRuns() As String, aDes() As String, Result As SCAN_RESULT
    Dim i&, j&, sKey$, sData$, sHit$, sAlias$, sParam As String, sMD5$, aValue() As String
    Dim bData() As Byte, isDisabledWin8 As Boolean, isDisabledWinXP As Boolean, flagDisabled As Long, sKeyDisable As String
    Dim sFile$, sArgs$, sUser$, bSafe As Boolean, aLines() As String, sLine As String
    Dim bShowPendingDeleted As Boolean
    
    bShowPendingDeleted = False
    
    ReDim aRegRuns(1 To 9)
    ReDim aDes(UBound(aRegRuns))
    
    aRegRuns(1) = "Software\Microsoft\Windows\CurrentVersion\Run"
    aDes(1) = "Run"
    
    aRegRuns(2) = "Software\Microsoft\Windows\CurrentVersion\RunServices"
    aDes(2) = "RunServices"
    
    aRegRuns(3) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
    aDes(3) = "RunOnce"
    
    aRegRuns(4) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    aDes(4) = "RunServicesOnce"
    
    aRegRuns(5) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
    aDes(5) = "Policies\Explorer\Run"
    
    aRegRuns(6) = "Software\Microsoft\Windows\CurrentVersion\Run-"
    aDes(6) = "Run-"

    aRegRuns(7) = "Software\Microsoft\Windows\CurrentVersion\RunServices-"
    aDes(7) = "RunServices-"

    aRegRuns(8) = "Software\Microsoft\Windows\CurrentVersion\RunOnce-"
    aDes(8) = "RunOnce-"

    aRegRuns(9) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce-"
    aDes(9) = "RunServicesOnce-"
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HE.AddKeys aRegRuns
    
    Do While HE.MoveNext
        
        Erase aValue
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue(), HE.Redirected)
        
            isDisabledWin8 = False
                    
            isDisabledWinXP = (Right$(HE.Key, 1) = "-")    ' Run- e.t.c.
                    
            sData = Reg.GetData(HE.Hive, HE.Key, aValue(i), HE.Redirected)
            
            If OSver.MajorMinor >= 6.2 Then  ' Win 8+
                      
                If StrComp(HE.Key, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 1) = 0 Then
                    
                    'Param. name is always "Run" on x32 bit. OS.
                    sKeyDisable = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And HE.Redirected, "Run32", "Run")
                    
                    If Reg.ValueExists(HE.Hive, sKeyDisable, aValue(i)) Then
                            
                        ReDim bData(0)
                        bData() = Reg.GetBinary(HE.Hive, sKeyDisable, aValue(i))
            
                        If UBoundSafe(bData) >= 11 Then
                            
                            GetMem4 ByVal VarPtr(bData(0)), flagDisabled
                            
                            If flagDisabled <> 2 Then isDisabledWin8 = True
                        End If
                    End If
                End If
            End If
            
            If Len(sData) <> 0 And Not isDisabledWin8 Then
                
                'Example:
                '"O4 - HKLM\..\Run: "
                '"O4 - HKU\S-1-5-19\..\Run: "
                sAlias = IIf(bIsWin32, "O4", IIf(HE.Redirected, "O4-32", "O4")) & _
                    " - " & IIf(isDisabledWinXP, "(disabled) ", "") & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": "
                
                sHit = sAlias & "[" & aValue(i) & "]"
                
                sUser = ""
                If HE.IsSidUser Then
                    sUser = " (User '" & HE.UserName & "')"
                End If
                
                SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                
                sFile = FormatFileMissing(sFile)
                
                sHit = sHit & " " & ConcatFileArg(sFile, sArgs) & sUser
                bSafe = False
                
                If Not bIgnoreAllWhitelists And bHideMicrosoft Then
                    
                    '//TODO: narrow down to services' DIS only: S-1-5-19 + S-1-5-20 + 'UpdatusUser' (NVIDIA)
                    
                    'Note: For services only
                    If StrComp(sFile, PF_64 & "\Windows Sidebar\Sidebar.exe", 1) = 0 And sArgs = "/autoRun" Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinDir & "\System32\mctadmin.exe", 1) = 0 And Len(sArgs) = 0 Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinSysDirWow64 & "\OneDriveSetup.exe", 1) = 0 And sArgs = "/thfirstsetup" Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                    
                    If OSver.MajorMinor = 6 Then 'Vista/2008
                        If WhiteListed(sFile, sWinDir & "\system32\rundll32.exe") And sArgs = "oobefldr.dll,ShowWelcomeCenter" Then
                            If IsMicrosoftFile(sWinDir & "\system32\oobefldr.dll") Then bSafe = True
                        End If
                    End If
                    
                    If OSver.MajorMinor <= 5.2 Then 'Win2k/XP/2003
                        If WhiteListed(sFile, sWinDir & "\system32\CTFMON.EXE") And Len(sArgs) = 0 Then bSafe = True
                    End If
                    
                    If OSver.MajorMinor = 6 Then 'Vista/2008
                        If WhiteListed(sFile, PF_64 & "\Windows Sidebar\sidebar.exe") And (sArgs = "/autoRun" Or sArgs = "/detectMem") Then bSafe = True
                    End If
                    
                End If
                
                If (Not IsOnIgnoreList(sHit)) And (Not bSafe) Then
                    
                    If bMD5 Then sMD5 = GetFileMD5(sFile): sHit = sHit & sMD5
                    
                    With Result
                        .Section = "O4"
                        .HitLineW = sHit
                        .Alias = sAlias
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i), , HE.Redirected
                        AddFileToFix .File, REMOVE_FILE, sFile
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
        Next
    Loop
    
    'Certain param based checkings
    
    Dim aRegKey() As String
    ReDim aRegKey(1 To 4) As String                   'key
    ReDim aRegParam(1 To UBound(aRegKey)) As String   'param
    ReDim aDefData(1 To UBound(aRegKey)) As String    'data
    ReDim aDes(1 To UBound(aRegKey)) As String        'description
    
    aRegKey(1) = "Software\Microsoft\Command Processor"
    aRegParam(1) = "Autorun"
    aDefData(1) = ""
    aDes(1) = "Command Processor\Autorun"
    
    aRegKey(2) = "HKLM\SYSTEM\CurrentControlSet\Control\BootVerificationProgram"
    aRegParam(2) = "ImagePath"
    aDefData(2) = ""
    aDes(2) = "BootVerificationProgram"
    
    aRegKey(3) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(3) = "BootExecute"
    If OSver.MajorMinor = 5 Then 'Win2k
        aDefData(3) = "autocheck autochk *, DfsInit"
    Else
        aDefData(3) = "autocheck autochk *"
    End If
    aDes(3) = "BootExecute"
    
    aRegKey(4) = "HKLM\SYSTEM\CurrentControlSet\Control\SafeBoot"
    aRegParam(4) = "AlternateShell"
    aDefData(4) = "cmd.exe"
    aDes(4) = "AlternateShell (SafeBoot)"
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        
        sParam = aRegParam(HE.KeyIndex)
        
        sData = Reg.GetData(HE.Hive, HE.Key, sParam, HE.Redirected)
        
        If InStr(sData, vbNullChar) <> 0 Then 'if MULTI_SZ
            sData = Replace$(sData, vbNullChar, ", ")
        End If
        
        If sData <> aDefData(HE.KeyIndex) Then
            
            'HKLM\..\Command Processor\Autorun:
            sAlias = IIf(bIsWin32, "O4", IIf(HE.Redirected, "O4-32", "O4")) & " - " & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": "
            
            SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
            
            sFile = FormatFileMissing(sFile)
            
            sHit = sAlias & ConcatFileArg(sFile, sArgs)
            
            If Not IsOnIgnoreList(sHit) Then
                
                If bMD5 Then sMD5 = GetFileMD5(sFile): sHit = sHit & sMD5
                
                With Result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    If sParam = "BootExecute" Then
                        AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sParam, aDefData(HE.KeyIndex), HE.Redirected, REG_RESTORE_MULTI_SZ
                        .CureType = REGISTRY_BASED
                    Else
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                        AddFileToFix .File, REMOVE_FILE, sFile
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End If
                End With
                AddToScanResults Result
            End If
        End If
    Loop
    
    ReDim aRegKey(1 To 2) As String                   'key
    ReDim aRegParam(1 To UBound(aRegKey)) As String   'param
    ReDim aDes(1 To UBound(aRegKey)) As String        'description
    
    'https://technet.microsoft.com/en-us/library/cc960241.aspx
    'PendingFileRenameOperations
    'Shared
    
    aRegKey(1) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(1) = "PendingFileRenameOperations"
    aDes(1) = "FileRenameOperations"
    
    aRegKey(2) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(2) = "PendingFileRenameOperations2"
    aDes(2) = "FileRenameOperations2"
    
    HE.Init HE_HIVE_HKLM, , HE_REDIR_NO_WOW
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        
        sParam = aRegParam(HE.KeyIndex)
    
        sData = Reg.GetData(HE.Hive, HE.Key, sParam, HE.Redirected)
        
        If Len(sData) <> 0 Then
        
          'converting MULTI_SZ to [1] -> [2], [3] -> [4] ...
          aLines = SplitSafe(sData, vbNullChar)
        
          For j = 0 To UBound(aLines) Step 2
            sFile = NormalizePath(aLines(j))
            If j + 1 <= UBound(aLines) Then
                If aLines(j + 1) = "" Then
                    sArgs = "-> DELETE"
                Else
                    sArgs = "-> " & NormalizePath(aLines(j + 1))
                End If
            End If
            
            'HKLM\..\FileRenameOperations:
            sAlias = IIf(bIsWin32, "O4", IIf(HE.Redirected, "O4-32", "O4")) & " - " & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": "
            
            sFile = FormatFileMissing(sFile)
            
            sHit = sAlias & ConcatFileArg(sFile, sArgs)
            
            If Not IsOnIgnoreList(sHit) Then
            
              If sArgs <> "-> DELETE" Or (sArgs = "-> DELETE" And bShowPendingDeleted) Then

                If bMD5 Then sMD5 = GetFileMD5(sFile): sHit = sHit & sMD5

                With Result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                    AddFileToFix .File, REMOVE_FILE, sFile
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults Result
              End If
            End If
          Next
        End If
    Loop
    
    Dim aFiles() As String, aData() As String
    ReDim aFiles(5)
    
    'Win 9x
    aFiles(0) = BuildPath(sWinSysDir, "BatInit.bat")
    aFiles(1) = BuildPath(sWinDir, "WinStart.bat")
    aFiles(2) = BuildPath(sWinDir, "DosStart.bat")
    aFiles(3) = BuildPath(SysDisk, "AutoExec.bat")
    'Win NT
    aFiles(4) = BuildPath(sWinSysDir, "AutoExec.nt")
    aFiles(5) = BuildPath(sWinSysDir, "Config.nt")
    
    For i = 0 To UBound(aFiles)
        sFile = aFiles(i)
        If FileExists(sFile) Then
            
            sData = ReadFileContents(sFile, False)
            
            If Len(sData) <> 0 Then
                
                sData = Replace$(sData, vbCr, "")
                aData = Split(sData, vbLf)
                
                'exclude comments
                For j = 0 To UBound(aData)
                    If StrBeginWith(aData(j), "REM") Then
                        aData(j) = ""
                    ElseIf StrBeginWith(aData(j), "::") Then
                        aData(j) = ""
                    ElseIf Not bIgnoreAllWhitelists Then
                        'check whitelist
                        If StrEndWith(sFile, "AutoExec.nt") Then
                            If aData(j) = "@echo off" Then
                                aData(j) = ""
                            ElseIf aData(j) = "lh %SystemRoot%\system32\mscdexnt.exe" Then
                                aData(j) = ""
                            ElseIf aData(j) = "lh %SystemRoot%\system32\redir" Then
                                aData(j) = ""
                            ElseIf aData(j) = "lh %SystemRoot%\system32\dosx" Then
                                aData(j) = ""
                            ElseIf aData(j) = "SET BLASTER=A220 I5 D1 P330 T3" Then
                                aData(j) = ""
                            End If
                        ElseIf StrEndWith(sFile, "Config.nt") Then
                            If aData(j) = "dos=high, umb" Then
                                aData(j) = ""
                            ElseIf aData(j) = "device=%SystemRoot%\system32\himem.sys" Then
                                aData(j) = ""
                            ElseIf StrComp(aData(j), "Files=40", 1) = 0 Then
                                aData(j) = ""
                            End If
                        End If
                    End If
                    
                    If 0 <> Len(aData(j)) Then
                        If Len(aData(j)) > BAT_LENGTH_LIMIT Then
                            aData(j) = Left$(aData(j), BAT_LENGTH_LIMIT) & " ... (" & Len(aData(j)) - BAT_LENGTH_LIMIT & " more characters)"
                        End If
                        If i < 4 Then
                            sAlias = "O4 - Win9x BAT: "
                        Else
                            sAlias = "O4 - WinNT BAT: "
                        End If
                        sHit = sAlias & sFile & " => " & EscapeSpecialChars(aData(j)) & IIf(Len(aData(j)) = 0, " (0 bytes)", "")
                        
                        If Not IsOnIgnoreList(sHit) Then
                            If bMD5 Then sMD5 = GetFileMD5(sFile): sHit = sHit & sMD5
                            With Result
                                .Section = "O4"
                                .HitLineW = sHit
                                .Alias = sAlias
                                AddFileToFix .File, REMOVE_FILE, sFile
                                .CureType = FILE_BASED
                            End With
                            AddToScanResults Result
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    'RunOnceEx, RunServicesOnceEx
    'https://support.microsoft.com/en-us/kb/310593
    'http://www.oszone.net/2762
    '" DllFileName | FunctionName | CommandLineArguments "
    Dim aSubKey() As String

    ReDim aRegKey(1 To 2) As String                   'key
    ReDim aDes(1 To UBound(aRegKey)) As String        'description
    
    aRegKey(1) = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
    aDes(1) = "RunOnceEx"
    aRegKey(2) = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnceEx"
    aDes(2) = "RunServicesOnceEx"
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        If Reg.KeyHasSubKeys(HE.Hive, HE.Key, HE.Redirected) Then
            Erase aSubKey
            For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey(), HE.Redirected, , False)
                Erase aValue
                For j = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key & "\" & aSubKey(i), aValue(), HE.Redirected)
                    
                    sData = Reg.GetString(HE.Hive, HE.Key & "\" & aSubKey(i), aValue(j), HE.Redirected)
                    
                    sAlias = "O4 - " & aDes(HE.KeyIndex) & ": "
                    sHit = sAlias & HE.HiveNameAndSID & "\..\" & aSubKey(i) & " [" & aValue(j) & "] - " & sData
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "O4"
                            .HitLineW = sHit
                            .Alias = sAlias
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\" & aSubKey(i), aValue(j), , HE.Redirected
                            AddRegToFix .Reg, REMOVE_KEY_IF_NO_VALUES, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                Next
            Next
        End If
    Loop
    
'    'Autorun.inf
'    'http://journeyintoir.blogspot.com/2011/01/autoplay-and-autorun-exploit-artifacts.html
'    Dim aDrives() As String
'    Dim sAutorun As String
'    Dim aVerb() As String
'    Dim bOnce As Boolean
'
'    aVerb = Split("open|shellexecute|shell\open\command|shell\explore\command", "|")
'
'    ' Mapping scheme for "inf. verb" -> to "registry" :
'    '
'    ' icon                  -> _Autorun\Defaulticon
'    ' open                  -> shell\AutoRun\command
'    ' shellexecute          -> shell\AutoRun\command
'    ' shell\open\command    -> shell\open\command
'    ' shell\explore\command -> shell\explore\command
'
'    aDrives = GetDrives(DRIVE_BIT_FIXED Or DRIVE_BIT_REMOVABLE)
'
'    For i = 1 To UBound(aDrives)
'        sAutorun = BuildPath(aDrives(i), "autorun.inf")
'        If FileExists(sAutorun) Then
'
'            bOnce = False
'
'            For j = 0 To UBound(aVerb)
'
'                sFile = ""
'                sArgs = ""
'                sData = ReadIniA(sAutorun, "autorun", aVerb(j))
'
'                If Len(sData) <> 0 Then
'                    SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=False
'                    sFile = FormatFileMissing(sFile)
'
'                    sHit = "O4 - Autorun.inf: " & sAutorun & " - " & aVerb(j) & " - " & ConcatFileArg(sFile, sArgs)
'
'                    If Not IsOnIgnoreList(sHit) Then
'                        With Result
'                            .Section = "O4"
'                            .HitLineW = sHit
'                            AddFileToFix .File, REMOVE_FILE, sAutorun
'                            .CureType = FILE_BASED
'                        End With
'                        AddToScanResults Result
'                    End If
'
'                    bOnce = True
'                End If
'            Next
'
'            'if unknown data is inside autorun.inf
'            If Not bOnce Then
'
'                sHit = "O4 - Autorun.inf: " & sAutorun & " - " & "(unknown target)"
'
'                If Not IsOnIgnoreList(sHit) Then
'                    With Result
'                        .Section = "O4"
'                        .HitLineW = sHit
'                        AddFileToFix .File, REMOVE_FILE, sAutorun
'                        .CureType = FILE_BASED
'                    End With
'                    AddToScanResults Result
'                End If
'            End If
'        End If
'    Next
'
'    'MountPoints2
'    HE.Init HE_HIVE_ALL
'    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
'
'    aVerb = Split("shell\AutoRun\command|shell\open\command|shell\explore\command", "|")
'
'    Do While HE.MoveNext
'        erase aSubKey
'        For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey, HE.Redirected)
'            For j = 0 To UBound(aVerb)
'                sKey = HE.Key & "\" & aSubKey(i) & "\" & aVerb(j)
'
'                If Reg.KeyExists(HE.Hive, sKey, HE.Redirected) Then
'
'                    sData = Reg.GetString(HE.Hive, sKey, "", HE.Redirected)
'
'                    SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
'                    sFile = FormatFileMissing(sFile)
'
'                    sHit = IIf(HE.Redirected, "O4-32", "O4") & " - MountPoints2: " & HE.HiveNameAndSID & "\..\" & aSubKey(i) & "\" & aVerb(j) & " - " & ConcatFileArg(sFile, sArgs)
'
'                    If Not IsOnIgnoreList(sHit) Then
'                        With Result
'                            .Section = "O4"
'                            .HitLineW = sHit
'                            'remove MountPoints2\{CLSID}
'                            'or
'                            'remove MountPoints2\Letter
'                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
'                            .CureType = REGISTRY_BASED
'                        End With
'                        AddToScanResults Result
'                    End If
'                End If
'            Next
'        Next
'    Loop
    
    'ScreenSaver
    sFile = Reg.GetString(HKCU, "Control Panel\Desktop", "SCRNSAVE.EXE")
    If 0 <> Len(sFile) And sFile <> "(Нет)" Then
        bSafe = True
        sFile = FormatFileMissing(sFile)
        
        If FileMissing(sFile) Then
            bSafe = False
        Else
            If Not IsMicrosoftFile(sFile) Then bSafe = False
        End If
        
        If Not bSafe Then
            sHit = "O4 - ScreenSaver: " & sFile
        
            If Not IsOnIgnoreList(sHit) Then
                With Result
                    .Section = "O4"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HKCU, "Control Panel\Desktop", "SCRNSAVE.EXE"
                    AddFileToFix .File, REMOVE_FILE, sFile
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults Result
            End If
        End If
    End If
    
    AppendErrorLogCustom "CheckO4_RegRuns - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.CheckO4_RegRuns"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_MSConfig(aHives() As String, aUser() As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4_MSConfig - Begin"
    
    'HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupreg
    'HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupfolder
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder -> checked in CheckO4_AutostartFolder()
    
    Dim sHive$, i&, j&, sAlias$, sMD5$, Result As SCAN_RESULT
    Dim aSubKey$(), sDay$, sMonth$, sYear$, sKey$, sFile$, sTime$, sHit$, SourceHive$, dEpoch As Date, sArgs$, sUser$
    Dim Values$(), bData() As Byte, flagDisabled As Long, dDate As Date, UseWow As Variant, Wow6432Redir As Boolean, sTarget$, sData$
    
    dEpoch = #1/1/1601#
    
    If OSver.MajorMinor >= 6.2 Then ' Win 8+
    
        For i = 0 To UBound(aHives) 'HKLM, HKCU, HKU\SID()

            sHive = aHives(i)
            
            For Each UseWow In Array(False, True)
    
                Wow6432Redir = UseWow
  
                If (bIsWin32 And Wow6432Redir) _
                  Or bIsWin64 And Wow6432Redir And (sHive = "HKCU" Or StrBeginWith(sHive, "HKU\")) Then
                    Exit For
                End If
            
                Erase Values
                For j = 1 To Reg.EnumValuesToArray(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values())
            
                    flagDisabled = 2
                    ReDim bData(0)
                    
                    bData() = Reg.GetBinary(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values(j))
                    
                    If UBoundSafe(bData) >= 11 Then
                        GetMem4 ByVal VarPtr(bData(0)), flagDisabled
                    End If
                    
                    If IsArrDimmed(bData) And flagDisabled <> 2 Then   'is Disabled ?
                    
                        dDate = ConvertFileTimeToLocalDate(VarPtr(bData(4)))
                        
                        If Reg.ValueExists(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), Wow6432Redir) Then
                        
                            sData = Reg.GetString(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), Wow6432Redir)
                        
                            'if you change it, change fix appropriate !!!
                            sAlias = "O4 - " & sHive & "\..\StartupApproved\" & IIf(bIsWin64 And Wow6432Redir, "Run32", "Run") & ": "
            
                            sHit = sAlias & "[" & Values(j) & "] "
                            
                            If (dDate <> dEpoch) Then sHit = sHit & "(" & Format$(dDate, "yyyy\/mm\/dd") & ") "
                            
                            sUser = ""
                            If aUser(i) <> "" And StrBeginWith(sHive, "HKU\") Then
                                If (sHive <> "HKU\S-1-5-18" And _
                                    sHive <> "HKU\S-1-5-19" And _
                                    sHive <> "HKU\S-1-5-20") Then sUser = " (User '" & aUser(i) & "')"
                            End If
                            
                            SplitIntoPathAndArgs sData, sFile, sArgs, True
                            
                            sFile = FormatFileMissing(sFile)
                            
                            sHit = sHit & ConcatFileArg(sFile, sArgs) & sUser
                        
                            If Not IsOnIgnoreList(sHit) Then
                            
                                If bMD5 Then sMD5 = GetFileMD5(sFile): sHit = sHit & sMD5
                
                                With Result
                                    .Section = "O4"
                                    .HitLineW = sHit
                                    .Alias = sAlias
                                    AddRegToFix .Reg, REMOVE_VALUE, 0, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values(j), , False
                                    AddRegToFix .Reg, REMOVE_VALUE, 0, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), , CLng(Wow6432Redir)
                                    AddFileToFix .File, REMOVE_FILE, sFile
                                    .CureType = REGISTRY_BASED Or FILE_BASED
                                End With
                                AddToScanResults Result
                            End If
                        End If
                    End If
                Next
            Next
        Next
        
    Else
    
        sHive = "HKLM"
        sKey = sHive & "\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupreg"
        
        For i = 1 To Reg.EnumSubKeysToArray(0&, sKey, aSubKey())
        
            sData = Reg.GetData(0&, sKey & "\" & aSubKey(i), "command")
            
            sYear = Reg.GetData(0&, sKey & "\" & aSubKey(i), "YEAR")
            sMonth = Right$("0" & Reg.GetData(0&, sKey & "\" & aSubKey(i), "MONTH"), 2)
            sDay = Right$("0" & Reg.GetData(0&, sKey & "\" & aSubKey(i), "DAY"), 2)
            
            If Val(sYear) = 0 Or Val(sMonth) = 0 Or Val(sDay) = 0 Then
                sTime = Format$(Reg.GetKeyTime(0&, sKey & "\" & aSubKey(i)), "yyyy\/mm\/dd")
            Else
                sTime = sYear & "/" & sMonth & "/" & sDay
            End If
            
            SourceHive = Reg.GetData(0&, sKey & "\" & aSubKey(i), "hkey")
            If SourceHive <> "HKLM" And SourceHive <> "HKCU" Then SourceHive = ""
            
            'O4 - MSConfig\startupreg: [RtHDVCpl] C:\Program Files\Realtek\Audio\HDA\RAVCpl64.exe -s (HKLM) (2016/10/13)
            sAlias = "O4 - MSConfig\startupreg: "
            
            sHit = sAlias & "[" & aSubKey(i) & "] "

            SplitIntoPathAndArgs sData, sFile, sArgs, True
            
            sFile = FormatFileMissing(sFile)
            
            sHit = sHit & ConcatFileArg(sFile, sArgs)
            
            If SourceHive <> "" Then sHit = sHit & " (" & SourceHive & ")"
            sHit = sHit & " (" & sTime & ")"
            
            If Not IsOnIgnoreList(sHit) Then
                
                If bMD5 Then sMD5 = GetFileMD5(sFile): sHit = sHit & sMD5
                
                With Result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_KEY, 0, sKey & "\" & aSubKey(i), , , REG_NOTREDIRECTED
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
        Next
        
        'Startup folder items
        
        sKey = "HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupfolder"
        
        Erase aSubKey
        For i = 1 To Reg.EnumSubKeysToArray(0&, sKey, aSubKey())
        
            sFile = Reg.GetData(0&, sKey & "\" & aSubKey(i), "backup")
            
            sTime = Format$(Reg.GetKeyTime(0&, sKey & "\" & aSubKey(i)), "yyyy\/mm\/dd")
        
            sAlias = "O4 - MSConfig\startupfolder: "    'if you change it, change fix appropriate !!!
            
            If UCase$(GetExtensionName(aSubKey(i))) = ".LNK" Then
                'expand LNK, like:
                'C:^ProgramData^Microsoft^Windows^Start Menu^Programs^Startup^GIGABYTE OC_GURU.lnk - C:\Windows\pss\GIGABYTE OC_GURU.lnk.CommonStartup
            
                If FileExists(sFile) Then
                    sTarget = GetFileFromShortcut(sFile, sArgs, True)
                End If
            End If
            
            If 0 <> Len(sTarget) Then
                sHit = sAlias & aSubKey(i) & " - " & sTarget & IIf(sArgs <> "", " " & sArgs, "") & " (" & sTime & ")" & IIf(Not FileExists(sTarget), " (file missing)", "")
            Else
                sHit = sAlias & aSubKey(i) & " - " & sFile & " (" & sTime & ")" & IIf(sFile = "", " (no file)", IIf(Not FileExists(sFile), " (file missing)", ""))
            End If
            
            If Not IsOnIgnoreList(sHit) Then
                
                If bMD5 Then sMD5 = GetFileMD5(sFile): sHit = sHit & sMD5
                
                With Result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_KEY, 0&, sKey & "\" & aSubKey(i), , , REG_NOTREDIRECTED
                    AddFileToFix .File, REMOVE_FILE, sFile 'removing backup (.pss)
                    .CureType = FILE_BASED Or REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
        Next
    End If
    
    AppendErrorLogCustom "CheckO4_MSConfig - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.CheckO4_MSConfig"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_AutostartFolder(aSID() As String, aUser() As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4_AutostartFolder - Begin"

    Dim aRegKeys() As String, aParams() As String, aDes() As String, aDesConst() As String, Result As SCAN_RESULT
    Dim sAutostartFolder$(), sShortCut$, i&, K&, Wow6432Redir As Boolean, UseWow, sFolder$, sHit$, dEpoch As Date
    Dim FldCnt&, sKey$, sSID$, sFile$, sLinkPath$, sLinkExt$, sTarget$, Blink As Boolean, bPE_EXE As Boolean
    Dim bData() As Byte, isDisabled As Boolean, flagDisabled As Long, sKeyDisable As String, sHive As String, dDate As Date
    Dim StartupCU As String, aFiles() As String, sArguments As String, aUserNames() As String, aUserConst() As String, sUsername$
    
    ReDim aRegKeys(1 To 8)
    ReDim aParams(1 To UBound(aRegKeys))
    ReDim aDesConst(1 To UBound(aRegKeys))
    ReDim aUserConst(1 To UBound(aRegKeys))

    ReDim sAutostartFolder(100) ' HKCU + HKLM + Wow64 + HKU
    ReDim aDes(100)
    ReDim aUserNames(100)
    
    dEpoch = #1/1/1601#
    
    'aRegKeys  - Key
    'aParams   - Value
    'aDesConst - Description for HJT Section
    
    'HKLM (HKLM hives should go first)
    aRegKeys(1) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(1) = "Common Startup"
    aDesConst(1) = "Global Startup"
    'aUserConst(1) = "All users"
    
    aRegKeys(2) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(2) = "Common AltStartup"
    aDesConst(2) = "Global AltStartup"
    'aUserConst(2) = "All users"
    
    aRegKeys(3) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(3) = "Common Startup"
    aDesConst(3) = "Global User Startup"
    'aUserConst(3) = "All users"
    
    aRegKeys(4) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(4) = "Common AltStartup"
    aDesConst(4) = "Global User AltStartup"
    'aUserConst(4) = "All users"
    
    'HKCU
    aRegKeys(5) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(5) = "Startup"
    aDesConst(5) = "Startup"
    'aUserConst(5) = envCurUser
    
    aRegKeys(6) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(6) = "AltStartup"
    aDesConst(6) = "AltStartup"
    'aUserConst(6) = envCurUser
    
    aRegKeys(7) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(7) = "Startup"
    aDesConst(7) = "User Startup"
    'aUserConst(7) = envCurUser
    
    aRegKeys(8) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(8) = "AltStartup"
    aDesConst(8) = "User AltStartup"
    'aUserConst(8) = envCurUser
    
    
    FldCnt = 0
    
    ' Get folder pathes
    For K = 1 To UBound(aRegKeys)
    
        For Each UseWow In Array(False, True)
            
            Wow6432Redir = UseWow
        
            'skip HKCU Wow64
            If (bIsWin32 And Wow6432Redir) _
              Or bIsWin64 And Wow6432Redir And StrBeginWith(aRegKeys(K), "HKCU") Then Exit For
    
            FldCnt = FldCnt + 1
            sAutostartFolder(FldCnt) = Reg.GetString(0&, aRegKeys(K), aParams(K), Wow6432Redir)
            aDes(FldCnt) = aDesConst(K)
            aUserNames(FldCnt) = aUserConst(K)
            
            'save path of Startup for current user to substitute other user names
            If aParams(K) = "Startup" Then
                If Len(sAutostartFolder(FldCnt)) <> 0 Then
                    StartupCU = UnQuote(EnvironW(sAutostartFolder(FldCnt)))
                End If
            End If
        Next
    Next
    
    '+ HKU pathes
    For i = 0 To UBound(aSID)
        If Len(aSID(i)) <> 0 Then
            sSID = aSID(i)
            
            For K = 1 To UBound(aRegKeys)
            
                'only HKCU keys
                If StrBeginWith(aRegKeys(K), "HKCU") Then
                
                    ' Convert HKCU -> HKU
                    sKey = Replace$(aRegKeys(K), "HKCU\", "HKU\" & sSID)
                
                    FldCnt = FldCnt + 1
                    If UBound(sAutostartFolder) < FldCnt Then
                        ReDim Preserve sAutostartFolder(UBound(sAutostartFolder) + 100)
                        ReDim Preserve aDes(UBound(aDes) + 100)
                        ReDim Preserve aUserNames(UBound(aUserNames) + 100)
                    End If
            
                    sAutostartFolder(FldCnt) = Reg.GetString(0&, sKey, aParams(K))
                    aDes(FldCnt) = sSID & " " & aDesConst(K)
                    aUserNames(FldCnt) = aUser(i)
                End If
            Next
        End If
    Next
    
    ReDim Preserve sAutostartFolder(FldCnt)
    ReDim Preserve aDes(FldCnt)
    ReDim Preserve aUserNames(FldCnt)
    
    For K = 1 To UBound(sAutostartFolder)
        sAutostartFolder(K) = UnQuote(EnvironW(sAutostartFolder(K)))
    Next
    
    ' adding all similar folders in c:\users (in case user isn't logged - so HKU\SID willn't be exist for him, cos his hive is not mounted)
    
    For i = 1 To colProfiles.Count
        'not current user
        If StrComp(colProfiles(i), UserProfile, 1) <> 0 Then
            If Len(colProfiles(i)) <> 0 Then
                ReDim Preserve sAutostartFolder(UBound(sAutostartFolder) + 1)
                ReDim Preserve aDes(UBound(aDes) + 1)
                ReDim Preserve aUserNames(UBound(aUserNames) + 1)
                sAutostartFolder(UBound(sAutostartFolder)) = Replace$(StartupCU, UserProfile, colProfiles(i), 1, 1, 1)
                aDes(UBound(aDes)) = "Startup other users"
                aUserNames(UBound(aUserNames)) = "...\" & GetFileNameAndExt(colProfiles(i))
            End If
        End If
    Next
    
    DeleteDuplicatesInArray sAutostartFolder, vbTextCompare, DontCompress:=True
    
    For K = 1 To UBound(sAutostartFolder)
        
        sUsername = aUserNames(K)
        
        sFolder = sAutostartFolder(K)
        
        If 0 <> Len(sFolder) Then
          If FolderExists(sFolder) Then
            
            Erase aFiles
            aFiles = ListFiles(sFolder)
            
              For i = 0 To UBoundSafe(aFiles)
            
                sShortCut = GetFileNameAndExt(aFiles(i))

                If LCase$(sShortCut) <> "desktop.ini" Then

                  If Not FolderExists(sFolder & "\" & sShortCut) Then
                  
                    isDisabled = False
              
                    If OSver.MajorMinor >= 6.2 Then  ' Win 8+

                        If StrInParamArray(aDes(K), "Startup", "User Startup", "Global Startup", "Global User Startup") Then

                            Select Case aDes(K)
                                Case "Startup": sHive = "HKCU"
                                Case "User Startup": sHive = "HKCU"
                                Case "Global Startup": sHive = "HKLM"
                                Case "Global User Startup": sHive = "HKLM"
                            End Select

                            sKeyDisable = sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder"

                            If Reg.ValueExists(0&, sKeyDisable, sShortCut) Then

                                ReDim bData(0)
                                bData() = Reg.GetBinary(0&, sKeyDisable, sShortCut)
                                
                                If UBoundSafe(bData) >= 11 Then
                            
                                    GetMem4 ByVal VarPtr(bData(0)), flagDisabled

                                    If flagDisabled <> 2 Then
                        
                                        isDisabled = True
                                        dDate = ConvertFileTimeToLocalDate(VarPtr(bData(4)))
                                    End If
                                End If
                            End If
                        End If
                    End If
                  
                  
                    sFile = ""
                    Blink = False
                    bPE_EXE = False
                    
                    sLinkPath = sFolder & "\" & sShortCut
                    sLinkExt = UCase$(GetExtensionName(sShortCut))
                    
                    'Example:
                    '"O4 - Global User AltStartup: "
                    '"O4 - S-1-5-19 User AltStartup: "
                    If isDisabled Then
                        sHit = "O4 - " & sHive & "\..\StartupApproved\StartupFolder: " 'if you change it, change fix also !!!
                    Else
                        sHit = "O4 - " & aDes(K) & ": "
                    End If
                    
                    If StrInParamArray(sLinkExt, ".LNK", ".URL", ".WEBSITE", ".PIF") Then Blink = True
                    
                    If Not Blink Or sLinkExt = ".PIF" Then  'not a Shortcut ?
                        bPE_EXE = isPE(sLinkPath)       'PE EXE ?
                    End If
                    
                    sTarget = ""
                    
                    If Blink Then
                        sTarget = GetFileFromShortcut(sLinkPath, sArguments)
                            
                        sHit = sHit & sShortCut & "    ->    " & sTarget & IIf(Len(sArguments) <> 0, " " & sArguments, "") 'doSafeURLPrefix
                    Else
                        sHit = sHit & sShortCut & IIf(bPE_EXE, "    ->    (PE EXE)", "")
                    End If
                    
                    If sUsername <> "" Then sHit = sHit & " (Folder '" & sUsername & "')"
                    
                    If isDisabled Then sHit = sHit & IIf(dDate <> dEpoch, " (" & Format$(dDate, "yyyy\/mm\/dd") & ")", "")
                    
                    If Not IsOnIgnoreList(sHit) Then
                        
                        If bMD5 Then
                            If Not Blink Or bPE_EXE Then
                                sHit = sHit & GetFileMD5(sLinkPath)
                            Else
                                If 0 <> Len(sTarget) Then
                                    sHit = sHit & GetFileMD5(sTarget)
                                End If
                            End If
                        End If
                        
                        With Result
                          .Section = "O4"
                          .HitLineW = sHit
                          
                          If isDisabled Then
                            .Alias = sHive & "\..\StartupApproved\StartupFolder:"
                            AddRegToFix .Reg, REMOVE_VALUE, 0&, sKeyDisable, sShortCut, , REG_NOTREDIRECTED
                            AddFileToFix .File, REMOVE_FILE, sLinkPath
                            .CureType = FILE_BASED Or REGISTRY_BASED
                          Else
                            .Alias = aDes(K)
                            AddFileToFix .File, REMOVE_FILE, sLinkPath
                            AddProcessToFix .Process, KILL_PROCESS, sTarget
                            .CureType = FILE_BASED Or PROCESS_BASED
                          End If
                        End With
                        AddToScanResults Result
                    End If
                  End If
                End If
              Next
          End If
        End If
    Next
    
    AppendErrorLogCustom "CheckO4_AutostartFolder - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.CheckO4_AutostartFolder"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckO4Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4Item - Begin"
    
    'Alpha 1.0. // Dragokas. Reworked. Bugs fix. Deleted x64/x32 views shared keys.
    'Added support of msconfig disabled items. Unicode support.
    
    '2.6.1.25 [05.06.16] // Dragokas. Full revision, simplifying, merging CheckO4ItemX86, CheckO4ItemUsers to 1 func.
    
    ' look at keys affected by wow64 redirector
    ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa384253(v=vs.85).aspx
    ' http://safezone.cc/threads/27567/
    
    '2.7.0.18 [11.11.2017] // Dragokas. Removed "|" char. vulnerability.
    
    'Scanning routines
    
    CheckO4_RegRuns
    
    CheckO4_MSConfig gHives(), gUsers()
    
    CheckO4_AutostartFolder gSIDs(), gUsers()
    
    AppendErrorLogCustom "CheckO4Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO4Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FillUsers()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "FillUsers - Begin"
    
    Dim i&
    
    GetUserNamesAndSids gSIDs(), gUsers()
    
    ReDim gHives(UBound(gSIDs) + 2)  '+ HKLM, HKCU
    ReDim Preserve gUsers(UBound(gHives))
    
    'Convert SID -> to hive
    For i = 0 To UBound(gSIDs)
        gHives(i) = "HKU\" & gSIDs(i)
    Next
    'Add HKLM, HKCU
    gHives(UBound(gHives) - 1) = "HKLM"
    gUsers(UBound(gHives) - 1) = "All users"
    
    gHives(UBound(gHives)) = "HKCU"
    gUsers(UBound(gHives)) = GetUser()
    
    AppendErrorLogCustom "FillUsers - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FillUsers"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub GetHives(aHives() As String, Optional bIncludeServiceSID As Boolean)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetHives - Begin"
    
    Dim i&, j&, aSID() As String, aUser() As String
    
    GetUserNamesAndSids aSID(), aUser()
    
    ReDim aHives(UBound(aSID) + 2)  '+ HKLM, HKCU
    
    'Convert SID -> to hive
    For i = 0 To UBound(aSID)
        If bIncludeServiceSID Or (Not bIncludeServiceSID And aSID(i) <> "S-1-5-18" And aSID(i) <> "S-1-5-19" And aSID(i) <> "S-1-5-20") Then
            aHives(j) = "HKU\" & aSID(i)
            j = j + 1
        End If
    Next
    
    aHives(j) = "HKCU"
    j = j + 1
    aHives(j) = "HKLM"
    ReDim Preserve aHives(j)
    
    AppendErrorLogCustom "GetHives - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetHives"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO4Item(sItem$, Result As SCAN_RESULT)
    'O4 - Enumeration of autoloading Regedit entries
    'O4 - HKLM\..\Run: [blah] program.exe
    'O4 - Startup: bla.lnk = c:\bla.exe
    'O4 - HKU\S-1-5-19\..\Run: [blah] program.exe (Username 'Joe')
    'O4 - Startup: bla.exe
    'O4 - MSConfig:
    'O4 - \..\StartupApproved\StartupFolder:
    '...
    
    On Error GoTo ErrorHandler:
    
    Dim sFile$

    FixProcessHandler Result
    
    If InStr(sItem, "StartupApproved\StartupFolder") <> 0 Then
        
        sFile = Result.File(0).Path
        
        If FileExists(sFile) Then
            If DeleteFileForce(sFile) Then
                FixRegistryHandler Result 'remove registry value if only file successfully deleted (!!!)
            End If
        Else
            FixRegistryHandler Result
        End If
        
        Exit Sub
    End If
    
    FixFileHandler Result
    FixRegistryHandler Result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO4Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckO5Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO5Item - Begin"
    
    Dim sControlIni$, sDummy$, sHit$, Result As SCAN_RESULT
    
    sControlIni = sWinDir & "\control.ini"
    If DirW$(sControlIni) = vbNullString Then Exit Sub
    
    'sDummy = string(5, " ")
    'GetPrivateProfileString "don't load", "inetcpl.cpl", "", sDummy, 5, sControlIni
    'sDummy = RTrim$(sDummy)
    
    IniGetString sControlIni, "don't load", "inetcpl.cpl"
    sDummy = RTrimNull(sDummy)
    
    If sDummy <> vbNullString Then
        sHit = "O5 - control.ini: inetcpl.cpl=" & sDummy
        
        If Not IsOnIgnoreList(sHit) Then
            With Result
                .Section = "O5"
                .HitLineW = sHit
                AddIniToFix .Reg, RESTORE_VALUE_INI, "control.ini", "don't load", "inetcpl.cpl", vbNullString
                .CureType = INI_BASED
            End With
            AddToScanResults Result
        End If
    End If
    
    AppendErrorLogCustom "CheckO5Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO5Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO5Item(sItem$, Result As SCAN_RESULT)
    'O5 - Blocking of loading Internet Options in Control Panel
    'WritePrivateProfileString "don't load", "inetcpl.cpl", vbNullString, "control.ini"
    On Error GoTo ErrorHandler:
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO5Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckO6Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO6Item - Begin"
    
    'If there are sub folders called
    '"restrictions" and/or "control panel", delete them
    
    Dim sHit$, Key$(2), Des$(2), Result As SCAN_RESULT
    'keys 0,1,2 - are x6432 shared.
    
    Key(0) = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    Des(0) = "Software\Policies\Microsoft\Internet Explorer\Restrictions present"
    
    Key(1) = "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions"
    Des(1) = "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions present"
    
    Key(2) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
    Des(2) = "Software\Policies\Microsoft\Internet Explorer\Control Panel present"
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL, HE_REDIR_BOTH
    HE.AddKeys Key()
    
    Do While HE.MoveNext
        If Reg.KeyHasValues(HE.Hive, HE.Key, HE.Redirected) Then
            sHit = IIf(HE.Redirected, "O6-32", "O6") & " - IE Policy: " & HE.HiveNameAndSID & "\" & Des(HE.KeyIndex)
            If Not IsOnIgnoreList(sHit) Then
                With Result
                    .Section = "O6"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key, , , HE.Redirected
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
        End If
    Loop
    
    If Reg.GetDword(HKLM, "SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings", "Security_HKLM_only") = 1 Then
        sHit = "O6 - IE Policy: Internet Settings - Security_HKLM_only = 1"
        If Not IsOnIgnoreList(sHit) Then
            With Result
                .Section = "O6"
                .HitLineW = sHit
                AddRegToFix .Reg, REMOVE_VALUE, HKLM, "SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings", "Security_HKLM_only"
                .CureType = REGISTRY_BASED
            End With
            AddToScanResults Result
        End If
    End If
    
    AppendErrorLogCustom "CheckO6Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO6Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO6Item(sItem$, Result As SCAN_RESULT)
    'O6 - Disabling of Internet Options' Main tab with Policies
    FixRegistryHandler Result
End Sub

Private Sub CheckSystemProblems()
    On Error GoTo ErrorHandler:
    
    'Checking for present and correct type of parameters:
    'HKCU\Environment => temp, tmp
    '+HKU
    
    'Checking for present, correct type of parameters and correct value:
    'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment => temp, tmp ("%SystemRoot%\TEMP")
    
    Dim sData As String, sDataNonExpanded As String
    Dim vParam, sKeyFull As String, sHit As String, sDefValue As String, Result As SCAN_RESULT
    
    If OSver.MajorMinor = 5 Then 'Win2k
        HE.Init HE_HIVE_ALL, HE_SID_ALL And Not HE_SID_SERVICE, HE_REDIR_NO_WOW
    Else
        HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    End If
    
    Do While HE.MoveNext
        For Each vParam In Array("TEMP", "TMP")
            
            sHit = ""
            
            If HE.Hive = HKLM Then
                sKeyFull = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
            Else
                sKeyFull = HE.HiveNameAndSID & "\Environment"
            End If
            
            If Not Reg.ValueExists(0, sKeyFull, CStr(vParam)) Then
                sHit = "O7 - TroubleShoot: [EV] " & HE.HiveNameAndSID & "\..\" & "%" & vParam & "%" & " - (environment variable is not exist)"
            Else
                sData = Reg.GetString(0, sKeyFull, CStr(vParam))
                sDataNonExpanded = Reg.GetString(0, sKeyFull, CStr(vParam), , True)
                
                If InStr(sData, "%") <> 0 Then
                    sHit = "O7 - TroubleShoot: [EV] " & HE.HiveNameAndSID & "\..\" & "%" & vParam & "%" & " - " & sData & " (wrong type of parameter)"
                ElseIf sData = "" Then
                    sHit = "O7 - TroubleShoot: [EV] " & HE.HiveNameAndSID & "\..\" & "%" & vParam & "%" & " - (empty value)"
                End If
                
                sData = EnvironW(sData)
                
                If HE.Hive = HKLM Then
                    If StrComp(sData, SysDisk & "\TEMP", 1) <> 0 _
                      And StrComp(sData, sWinDir & "\TEMP", 1) <> 0 Then 'if wrong value
                        sHit = "O7 - TroubleShoot: [EV] " & HE.HiveNameAndSID & "\..\" & "%" & vParam & "%" & " - " & sData & " (environment value is altered)"
                    End If
                Else
                    If OSver.MajorMinor < 6 Then
                        If StrComp(sData, UserProfile & "\Local Settings\Temp", 1) <> 0 _
                          And StrComp(sData, SysDisk & "\TEMP", 1) <> 0 Then 'if wrong value
                            sHit = "O7 - TroubleShoot: [EV] " & HE.HiveNameAndSID & _
                              "\..\" & "%" & vParam & "%" & " - " & sData & " (environment value is altered)"
                        End If
                    Else
                        If StrComp(sData, LocalAppData & "\Temp", 1) <> 0 _
                          And StrComp(sData, SysDisk & "\TEMP", 1) <> 0 Then 'if wrong value
                            sHit = "O7 - TroubleShoot: [EV] " & HE.HiveNameAndSID & _
                              "\..\" & "%" & vParam & "%" & " - " & sData & " (environment value is altered)"
                        End If
                    End If
                End If
            End If
            
            If sHit <> "" Then
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O7"
                        .HitLineW = sHit
                        If HE.Hive = HKLM Then
                            sDefValue = "%SystemRoot%\TEMP"
                        Else
                            If OSver.MajorMinor < 6 Then
                                sDefValue = "%USERPROFILE%\Local Settings\Temp"
                            Else
                                sDefValue = "%USERPROFILE%\AppData\Local\Temp"
                            End If
                        End If
                        AddRegToFix .Reg, RESTORE_VALUE, 0, sKeyFull, CStr(vParam), sDefValue, REG_NOTREDIRECTED, REG_RESTORE_EXPAND_SZ
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
        Next
    Loop
    
    Dim cFreeSpace As Currency
    
    cFreeSpace = GetFreeDiscSpace(SysDisk, False)
    ' < 1 GB ?
    If (cFreeSpace < cMath.MBToInt64(1& * 1024)) And (cFreeSpace <> 0@) Then
        
        sHit = "O7 - TroubleShoot: [Disk] Free disk space on " & SysDisk & " is too low = " & (cFreeSpace / 1024& / 1024& * 10000& \ 1) & " MB."
        
        If Not IsOnIgnoreList(sHit) Then
            With Result
                .Section = "O7"
                .HitLineW = sHit
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults Result
        End If
    End If
    
    Dim sNetBiosName As String
    
    If GetCompName(ComputerNamePhysicalDnsHostname) = "" Then
    
        sNetBiosName = GetCompName(ComputerNameNetBIOS)
        sHit = "O7 - TroubleShoot: [Network] Computer name (hostname) is not set (should be: " & sNetBiosName & ")"
        
        If Not IsOnIgnoreList(sHit) Then
            With Result
                .Section = "O7"
                .HitLineW = sHit
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults Result
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblems"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckCertificatesEDS()
    On Error GoTo ErrorHandler:
    'Checking for untrusted code signing root certificates
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SystemCertificates\Disallowed\Certificates
    
    'infections examples:
    'https://blog.malwarebytes.com/cybercrime/2015/11/vonteera-adware-uses-certificates-to-disable-anti-malware/
    'https://www.bleepingcomputer.com/news/security/certlock-trojan-blocks-security-programs-by-disallowing-their-certificates/
    'https://www.securitylab.ru/news/486648.php
    
    'reverse 'blob'
    'https://namecoin.org/2017/05/27/reverse-engineering-cryptoapi-cert-blobs.html
    'https://itsme.home.xs4all.nl/projects/xda/smartphone-certificates.html
    'https://msdn.microsoft.com/en-us/library/windows/desktop/aa376079%28v=vs.85%29.aspx
    'https://msdn.microsoft.com/en-us/library/windows/desktop/aa376573%28v=vs.85%29.aspx
    'https://msdn.microsoft.com/en-us/library/cc232282.aspx
    
    Dim i&, aSubKey$(), Idx&, sTitle$, bSafe As Boolean, sHit$, Result As SCAN_RESULT
    Dim Blob() As Byte, CertHash As String, FriendlyName As String
    
    For i = 1 To Reg.EnumSubKeysToArray(HKLM, "SOFTWARE\Microsoft\SystemCertificates\Disallowed\Certificates", aSubKey())
        
        bSafe = True
        sTitle = ""
        
        Blob = Reg.GetBinary(HKLM, "SOFTWARE\Microsoft\SystemCertificates\Disallowed\Certificates\" & aSubKey(i), "Blob")
        
        If AryPtr(Blob) Then
            ParseCertBlob Blob, CertHash, FriendlyName
            
            'Debug.Print "(" & FriendlyName & ")"
            
            If CertHash = "" Then CertHash = aSubKey(i)
            
            Idx = GetCollectionIndexByKey(CertHash, colSafeCert)
            
            If Idx <> 0 Then
                'it's safe
                If Not bHideMicrosoft Or bIgnoreAllWhitelists Then
                    sTitle = GetCollectionKeyByIndex(Idx, colSafeCert)
                    bSafe = False
                End If
            Else
                bSafe = False
                Idx = GetCollectionIndexByKey(CertHash, colBadCert)
                
                If Idx <> 0 Then
                    sTitle = GetCollectionKeyByIndex(Idx, colBadCert) & " (Well-known cert.)"
                End If
            End If
            
            If Not bSafe Then
                If sTitle = "" Then sTitle = "Unknown"
                If FriendlyName <> "" Then sTitle = sTitle & " (" & FriendlyName & ")"
                If FriendlyName = "Fraudulent" Or FriendlyName = "Untrusted" Then sTitle = sTitle & " (HJT: possible, safe)"
                
                'Hash - 'Name, cert. issued to' (Name of cert.) (HJT rating, if possible)
                sHit = "O7 - Policy: [Untrusted Certificate] " & CertHash & " - " & sTitle
                
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\SystemCertificates\Disallowed\Certificates\" & aSubKey(i)
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblems"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub ParseCertBlob(Blob() As Byte, out_CertHash As String, out_FriendlyName As String)
    On Error GoTo ErrorHandler:
    
    'Thanks to Willem Jan Hengeveld
    'https://itsme.home.xs4all.nl/projects/xda/smartphone-certificates.html
    
    Const SHA1_HASH As Long = 3
    Const FRIENDLY_NAME As Long = 11 'Fraudulent
    
    Dim prop As CERTIFICATE_BLOB_PROPERTY
    Dim cStream As clsStream
    Set cStream = New clsStream
    
    'registry blob is an array of CERTIFICATE_BLOB_PROPERTY structures.
    
    cStream.WriteData VarPtr(Blob(0)), UBound(Blob) + 1
    cStream.BufferPointer = 0
    
    Do While cStream.BufferPointer < cStream.Size
        cStream.ReadData VarPtr(prop), 12
        If prop.Length > 0 Then
            ReDim prop.Data(prop.Length - 1)
            cStream.ReadData VarPtr(prop.Data(0)), prop.Length
            
'            Debug.Print "PropID: " & prop.PropertyID
'            Debug.Print "Length: " & prop.Length
'            Debug.Print "DataA:   " & Replace(StringFromPtrA(VarPtr(prop.Data(0))), vbNullChar, "-")
'            Debug.Print "DataW:   " & StringFromPtrW(VarPtr(prop.Data(0)))
'            Debug.Print "HexData: " & GetHexStringFromArray(prop.Data)
            'If prop.PropertyID = 32 Then Stop
            
            Select Case prop.PropertyID
            Case SHA1_HASH
                out_CertHash = GetHexStringFromArray(prop.Data)
            Case FRIENDLY_NAME
                out_FriendlyName = StringFromPtrW(VarPtr(prop.Data(0)))
            End Select
            
            If out_CertHash <> "" And out_FriendlyName <> "" Then Exit Do
        End If
    Loop
    
    Set cStream = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ParseCertBlob"
    If inIDE Then Stop: Resume Next
End Sub

Private Function GetHexStringFromArray(a() As Byte) As String
    Dim sHex As String
    Dim i As Long
    For i = 0 To UBound(a)
        sHex = sHex & Right$("0" & Hex(a(i)), 2)
    Next
    GetHexStringFromArray = sHex
End Function

Private Sub CheckO7Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO7Item - Begin"
    
    Dim lData&, sHit$, Result As SCAN_RESULT
    Dim i As Long
    
    'http://www.oszone.net/11424
    
    '//TODO:
    '%WinDir%\System32\GroupPolicyUsers"
    '%WinDir%\System32\GroupPolicy"
    'HKEY_CURRENT_USER\Software\Policies\Microsoft
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies

    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    
    Do While HE.MoveNext
        'key - x64 Shared
        lData = Reg.GetDword(HE.Hive, HE.Key, "DisableRegistryTools")
        If lData <> 0 Then
            sHit = "O7 - Policy: " & HE.HiveNameAndSID & "\..\" & "DisableRegistryTools = " & lData
            
            If Not IsOnIgnoreList(sHit) Then
                With Result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "DisableRegistryTools"
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
        End If
    Loop
    
    'Taskbar policies
    'см. Клименко Р. Тонкости реестра Windows Vista. Трюки и эффекты.
    Dim aValue() As String
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    aValue = Split("NoSetTaskbar|TaskbarLockAll|NoTrayItemsDisplay|NoChangeStartMenu|NoStartMenuMorePrograms|NoRun" & _
        "NoSMConfigurePrograms", "|")
    
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i))
            If lData <> 0 Then
                sHit = "O7 - Taskbar policy: " & HE.HiveNameAndSID & "\..\" & aValue(i) & " = " & lData
            
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
        Next
    Loop
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    
    aValue = Split("Start_ShowRun", "|")
    
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            If Reg.ValueExists(HE.Hive, HE.Key, aValue(i)) Then
                lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i))
                If lData = 0 Then
                    sHit = "O7 - Taskbar policy: " & HE.HiveNameAndSID & "\..\" & aValue(i) & " = " & lData
                
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "O7"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
        Next
    Loop
    
    'Control panel policies
    
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\don't load (by database)
    'HKCU\Control Panel\don't load
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => DisallowCpl = 1
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => RestrictCpl = 1
    'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => NoControlPanel = 1
    
    
    
    Call CheckCertificatesEDS 'Untrusted certificates
    
    Call CheckSystemProblems '%temp%, %tmp%, disk free space < 1 GB.
    
    'IPSec policy
    'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\
    'secpol.msc
    
    'ipsecPolicy{GUID}                  'example: 5d57bbac-8464-48b2-a731-9dd7e6f65c9f
    
    '\ipsecName                         -> Name of policy
    '\whenChanged                       -> Date in Unix format ( ConvertUnixTimeToLocalDate )
    '\ipsecNFAReference [REG_MULTI_SZ]  -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNFA{GUID_1} '96372b24-f2bf-4f50-a036-5897aac92f2f
                                                   'SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNFA{GUID_2} '8c676c64-306c-47db-ab50-e0108a1621dd
    
    'Note: One of these ipsecNFA{GUID} key may not contain 'ipsecFilterReference' parameter
    
    '\ipsecISAKMPReference              -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecISAKMPPolicy{GUID} '738d84c5-d070-4c6c-9468-12b171cfd10e
    
    '--------------------------------
    'ipsecNFA{GUID}
    
    '\ipsecNegotiationPolicyReference     -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNegotiationPolicy{GUID} '7c5a4ff0-ae4b-47aa-a2b6-9a72d2d6374c
    '\ipsecFilterReference [REG_MULTI_SZ] -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecFilter{GUID_1} 'c73baa5d-71a6-4533-bf7d-f640b1ff2eb8
    
    '--------------------------------
    'ipsecNegotiationPolicy{GUID}
    
    'Trojan example: https://www.trendmicro.com/vinfo/us/threat-encyclopedia/malware/troj_dloade.xn
    
    Dim KeyPolicy() As String, IPSecName$, KeyNFA() As String, KeyNegotiation() As String, dModify As Date, lModify As Long, IPSecID As String
    Dim KeyISAKMP As String, j As Long, KeyFilter() As String, K As Long, NegAction As String, NegType As String, bEnabled As Boolean, sActPolicy As String
    Dim bFilterData() As Byte, IP(1) As String, RuleAction As String, bMirror As Boolean, DataSerialized As String
    Dim Packet_Type(1) As String, M As Long, n As Long, PortNum(1) As Long, ProtocolType As String, idxBaseOffset As Long, IpFil As IPSEC_FILTER_RECORD, RecCnt As Byte
    Dim bRegexpInit As Boolean, oMatches As IRegExpMatchCollection, IPTypeFlag(1) As Long, b() As Byte, bAtLeastOneFilter As Boolean, bNoFilter As Boolean
    Dim sHitTrimmed$, bSafe As Boolean
    
    Erase KeyPolicy
    For i = 1 To Reg.EnumSubKeysToArray(0&, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local", KeyPolicy())
    
      If StrBeginWith(KeyPolicy(i), "ipsecPolicy{") Then
        
        'what policy is currently active?
        sActPolicy = Reg.GetString(0&, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local", "ActivePolicy")
        
        bEnabled = (StrComp(sActPolicy, "SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\" & KeyPolicy(i), 1) = 0)
        
        'add prefix
        KeyPolicy(i) = "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyPolicy(i)
        
        bMirror = False
        RuleAction = ""
        
        IPSecID = Mid$(KeyPolicy(i), InStrRev(KeyPolicy(i), "{"))
        
        IPSecName = Reg.GetString(0&, KeyPolicy(i), "ipsecName")
        
        lModify = Reg.GetDword(0&, KeyPolicy(i), "whenChanged")
        
        dModify = ConvertUnixTimeToLocalDate(lModify)
        
        KeyISAKMP = Reg.GetString(0&, KeyPolicy(i), "ipsecISAKMPReference")
        KeyISAKMP = MidFromCharRev(KeyISAKMP, "\")
        KeyISAKMP = IIf(KeyISAKMP = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyISAKMP)
        
        Erase KeyNFA
        Erase KeyFilter
        Erase KeyNegotiation
        Erase IP
        Erase Packet_Type: Packet_Type(0) = "Unknown": Packet_Type(1) = "Unknown"
        Erase PortNum
        RuleAction = ""
        ProtocolType = ""
        bMirror = False
        RuleAction = "Unknown"
        bNoFilter = False
        
        KeyNFA() = Reg.GetMultiSZ(0&, KeyPolicy(i), "ipsecNFAReference")
        '() -> ipsecNegotiationPolicy
        '() -> ipsecFilter (optional)
        
        If IsArrDimmed(KeyNFA) Then
            
          For j = 0 To UBound(KeyNFA)
            KeyNFA(j) = MidFromCharRev(KeyNFA(j), "\")
            KeyNFA(j) = IIf(KeyNFA(j) = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyNFA(j))
          Next
          
          ReDim KeyNegotiation(UBound(KeyNFA))
          
          For j = 0 To UBound(KeyNFA)
            KeyNegotiation(j) = Reg.GetString(0&, KeyNFA(j), "ipsecNegotiationPolicyReference")
            KeyNegotiation(j) = MidFromCharRev(KeyNegotiation(j), "\")
            KeyNegotiation(j) = IIf(KeyNegotiation(j) = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyNegotiation(j))
          Next
          
          For j = 0 To UBound(KeyNFA)
            
            NegType = Reg.GetString(0&, KeyNegotiation(j), "ipsecNegotiationPolicyType")
            NegAction = Reg.GetString(0&, KeyNegotiation(j), "ipsecNegotiationPolicyAction")
            
            'GUIDs: https://msdn.microsoft.com/en-us/library/cc232441.aspx
            
            If StrComp(NegType, "{62f49e10-6c37-11d1-864c-14a300000000}", 1) = 0 Then 'without last one "-" character (!)
                If StrComp(NegAction, "{8a171dd2-77e3-11d1-8659-a04f00000000}", 1) = 0 Then
                    RuleAction = "Allow"
                ElseIf StrComp(NegAction, "{3f91a819-7647-11d1-864d-d46a00000000}", 1) = 0 Then
                    RuleAction = "Block"
                ElseIf StrComp(NegAction, "{8a171dd3-77e3-11d1-8659-a04f00000000}", 1) = 0 Then
                    RuleAction = "Approve security"
                ElseIf StrComp(NegAction, "{3f91a81a-7647-11d1-864d-d46a00000000}", 1) = 0 Then
                    RuleAction = "Inbound pass-through"
                Else
                    RuleAction = "Unknown"
                End If
            ElseIf StrComp(NegType, "{62f49e13-6c37-11d1-864c-14a300000000}", 1) = 0 Then
                RuleAction = "Default response"
            Else
                RuleAction = "Unknown"
            End If
            
            Erase KeyFilter
            Erase IP
            Erase Packet_Type: Packet_Type(0) = "Unknown": Packet_Type(1) = "Unknown"
            Erase PortNum
            ProtocolType = ""
            bMirror = False
            
            KeyFilter() = Reg.GetMultiSZ(0&, KeyNFA(j), "ipsecFilterReference")
                        
            If Not IsArrDimmed(KeyFilter) Then
            
                bAtLeastOneFilter = False
                
                For M = 0 To UBound(KeyNFA)
                    If Reg.ValueExists(0&, KeyNFA(M), "ipsecFilterReference") Then
                        bAtLeastOneFilter = True
                        Exit For
                    End If
                Next
                
                If Not bAtLeastOneFilter Then
                    bNoFilter = True
                    GoSub AddItem
                End If
            Else
                
                For K = 0 To UBound(KeyFilter)
                    KeyFilter(K) = MidFromCharRev(KeyFilter(K), "\")
                    KeyFilter(K) = IIf(KeyFilter(K) = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyFilter(K))
                Next
                
                For K = 0 To UBound(KeyFilter)
                    
                    Erase IP
                    Erase Packet_Type: Packet_Type(0) = "Unknown": Packet_Type(1) = "Unknown"
                    Erase PortNum
                    ProtocolType = ""
                    bMirror = False
                    
                    bFilterData() = Reg.GetBinary(0&, KeyFilter(K), "ipsecData")
                    
                    If IsArrDimmed(bFilterData) Then

                      If Not bRegexpInit Then
                        bRegexpInit = True
                        Set oRegexp = New cRegExp
                        oRegexp.IgnoreCase = True
                        oRegexp.Global = True
                        oRegexp.Pattern = "(00|01)(000000)(........)(00000000|FFFFFFFF)(........)(00000000|FFFFFFFF)(00000000)(((06|11)000000........)|((00|01|06|08|11|14|16|1B|42|FF|..)00000000000000))00(00|01|02|03|04|81|82|83|84)0000"
                      End If
                    
                      Set oMatches = oRegexp.Execute(SerializeByteArray(bFilterData, ""))
    
                      For n = 0 To oMatches.Count - 1
                      
                        b = DeSerializeToByteArray(oMatches(n))
                        
                        memcpy IpFil, b(0), Len(IpFil)

                        '00,00,00,00,00,00,00,00 -> any IP
                        'xx,xx,xx,xx,ff,ff,ff,ff -> specified IP / subnet
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 0 -> my IP
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 1 or 0x81 -> DNS-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 2 or 0x82 -> WINS-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 3 or 0x83 -> DHCP-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 4 or 0x84 -> Gateway
                        '
                        '[0x4E] == 1 -> mirrored
                        '
                        '[0x66] -> port type
                        '[0x6A] -> port number (source) (2 bytes)
                        '[0x6C] -> port number (destination) (2 bytes)

                        bMirror = (IpFil.Mirrored = 1)
                        PortNum(0) = cMath.ShortIntToUShortInt(IpFil.PortNum1)
                        PortNum(1) = cMath.ShortIntToUShortInt(IpFil.PortNum2)
                        
                        Select Case IpFil.ProtocolType
                            Case 0: ProtocolType = "Any"
                            Case 6: ProtocolType = "TCP"
                            Case 17: ProtocolType = "UDP"
                            Case 1: ProtocolType = "ICMP"
                            Case 27: ProtocolType = "RDP"
                            Case 8: ProtocolType = "EGP"
                            Case 20: ProtocolType = "HMP"
                            Case 255: ProtocolType = "RAW"
                            Case 66: ProtocolType = "RVD"
                            Case 22: ProtocolType = "XNS-IDP"
                            Case Else: ProtocolType = "type: " & CLng(bFilterData(&H66))
                        End Select
                        
                        IP(0) = IpFil.IP1(0) & "." & IpFil.IP1(1) & "." & IpFil.IP1(2) & "." & IpFil.IP1(3)
                        IP(1) = IpFil.IP2(0) & "." & IpFil.IP2(1) & "." & IpFil.IP2(2) & "." & IpFil.IP2(3)

                        IPTypeFlag(0) = IpFil.IPTypeFlag1
                        IPTypeFlag(1) = IpFil.IPTypeFlag2

                        For M = 0 To 1
                        
                            If IPTypeFlag(M) = 0 Then       '00,00,00,00,00,00,00,00
                                If IP(M) = "0.0.0.0" Then Packet_Type(M) = "Any IP"
                            
                            ElseIf IPTypeFlag(M) = -1 Then  '00,00,00,00,ff,ff,ff,ff
                                If IP(M) = "0.0.0.0" Then
                            
                                    Select Case IpFil.DynPacketType
                                        Case 0: Packet_Type(M) = "my IP"
                                        Case &H81, 1: Packet_Type(M) = "DNS-servers"
                                        Case &H82, 2: Packet_Type(M) = "WINS-servers"
                                        Case &H83, 3: Packet_Type(M) = "DHCP-servers"
                                        Case &H84, 4: Packet_Type(M) = "Gateway"
                                        Case Else: Packet_Type(M) = "Unknown"
                                        '1,2,3,4 - Source packets
                                        '81,82,83,84 - Destination packets
                                    End Select
                                Else                            'xx,xx,xx,xx,ff,ff,ff,ff
                                    Packet_Type(M) = "IP"
                                End If
                            Else
                                Packet_Type(M) = "Unknown"
                            End If
                        
                            If IP(M) = "0.0.0.0" Then IP(M) = ""
                        Next
                        
                        GoSub AddItem
                        
                      Next
                      
                    Else
                        GoSub AddItem
                    End If
                Next

            End If
            
          Next
          
        Else
            GoSub AddItem
        End If
        
      End If
    Next
    
    AppendErrorLogCustom "CheckO7Item - End"
    Exit Sub
AddItem:
    'keys:
    'KeyPolicy(i) - 1
    'KeyISAKMP - 1
    'KeyNFA(j) - 0 to ...
    'KeyNegotiation - 1
    'KeyFilter(k) - 0 to ...
    
    'flags:
    'bEnabled - policy enabled ?
    'bMirror - true, if rule also applies to reverse direction: from destination to source
    
    'Other:
    'IPSecName - name of policy
    'IPSecID - identifier in registry
    'dModify - date last modified
    'RuleAction - action for filter
    'PortNum()
    'ProtocolType
    
    'example:
    'O7 - IPSec: (Enabled) IP_Policy_Name [yyyy/mm/dd] - {5d57bbac-8464-48b2-a731-9dd7e6f65c9f} - Source: My IP - Destination: 8.8.8.8 (Port 80 TCP) - (mirrored) Action: Block
    
    sHit = "O7 - IPSec: " & IPSecName & " " & _
        "[" & Format$(dModify, "yyyy\/mm\/dd") & "]" & " - " & IPSecID & " - " & _
        IIf(bNoFilter, "No rules ", _
        "Source: " & IIf(Packet_Type(0) = "IP", "IP: " & IP(0), Packet_Type(0)) & _
        IIf((ProtocolType = "TCP" Or ProtocolType = "UDP") And PortNum(0) <> 0, " (Port " & PortNum(0) & " " & ProtocolType & ")", "") & " - " & _
        "Destination: " & IIf(Packet_Type(1) = "IP", "IP: " & IP(1), Packet_Type(1)) & _
        IIf((ProtocolType = "TCP" Or ProtocolType = "UDP") And PortNum(1) <> 0, " (Port " & PortNum(1) & " " & ProtocolType & ")", "") & " " & _
        IIf(bMirror, "(mirrored) ", "")) & "- Action: " & RuleAction & IIf(bEnabled, "", " (disabled)")

    sHitTrimmed = Mid$(sHit, Len("O7 - IPSec: " & IPSecName & " " & _
        "[" & Format$(dModify, "yyyy\/mm\/dd") & "]" & " - ") + 1)
    
    bSafe = False
    
    'Whitelists
    If OSver.MajorMinor <= 5.2 Then 'Win2k / XP
        If sHitTrimmed = "{72385236-70fa-11d1-864c-14a300000000} - No rules - Action: Default response (disabled)" Then
            bSafe = True
        ElseIf sHitTrimmed = "{72385230-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Allow (disabled)" Then
            bSafe = True
        ElseIf sHitTrimmed = "{72385230-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Inbound pass-through (disabled)" Then
            bSafe = True
        ElseIf sHitTrimmed = "{7238523c-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Allow (disabled)" Then
            bSafe = True
        ElseIf sHitTrimmed = "{7238523c-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Inbound pass-through (disabled)" Then
            bSafe = True
        End If
    End If
       
    If Not bSafe Then
      If Not IsOnIgnoreList(sHit) Then
        With Result
            .Section = "O7"
            .HitLineW = sHit
            AddRegToFix .Reg, REMOVE_KEY, 0, KeyPolicy(i)
            If KeyISAKMP <> "" Then AddRegToFix .Reg, REMOVE_KEY, 0, KeyISAKMP
            If IsArrDimmed(KeyNFA) Then
                For M = 0 To UBound(KeyNFA)
                    If KeyNFA(M) <> "" Then
                        AddRegToFix .Reg, REMOVE_KEY, 0, KeyNFA(M)
                    End If
                    If KeyNegotiation(M) <> "" Then
                        AddRegToFix .Reg, REMOVE_KEY, 0, KeyNegotiation(M)
                    End If
                Next
            End If
            If IsArrDimmed(KeyFilter) Then
                For M = 0 To UBound(KeyFilter)
                    If KeyFilter(M) <> "" Then
                        AddRegToFix .Reg, REMOVE_KEY, 0, KeyFilter(M)
                    End If
                Next
            End If
            If bEnabled Then
                AddRegToFix .Reg, REMOVE_VALUE, 0, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local", "ActivePolicy"
            End If
            .CureType = REGISTRY_BASED
        End With
        AddToScanResults Result
      End If
    End If
    
    Return
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO7Item"
    If inIDE Then Stop: Resume Next
End Sub

'byte array -> to Hex String
Public Function SerializeByteArray(b() As Byte, Optional Delimiter As String = "") As String
    Dim i As Long
    'If isarrdimmed(b) Then
        For i = LBound(b) To UBound(b)
            SerializeByteArray = SerializeByteArray & Right$("0" & Hex$(b(i)), 2) & Delimiter
        Next
        
        If Len(Delimiter) <> 0 Then SerializeByteArray = Left$(SerializeByteArray, Len(SerializeByteArray) - Len(Delimiter))
    'End If
End Function

'Serialized Hex String of bytes -> byte array
Public Function DeSerializeToByteArray(S As String, Optional Delimiter As String = "") As Byte()
    On Error GoTo ErrorHandler:
    Dim i As Long
    Dim n As Long
    Dim b() As Byte
    Dim ArSize As Long
    If Len(S) = 0 Then Exit Function
    ArSize = (Len(S) + Len(Delimiter)) \ (2 + Len(Delimiter)) '2 chars on byte + add final delimiter
    ReDim b(ArSize - 1) As Byte
    For i = 1 To Len(S) Step 2 + Len(Delimiter)
        b(n) = CLng("&H" & Mid$(S, i, 2))
        n = n + 1
    Next
    DeSerializeToByteArray = b
    Exit Function
ErrorHandler:
    Debug.Print "Error in DeSerializeByteString"
End Function

Public Sub FixO7Item(sItem$, Result As SCAN_RESULT)
    'O7 - Disabling of Policies
    On Error GoTo ErrorHandler:
    
    If Result.CureType = CUSTOM_BASED Then
    
        If InStr(1, Result.HitLineW, "Free disk space", 1) <> 0 Then
            RunCleanMgr
            
        ElseIf InStr(1, Result.HitLineW, "Computer name (hostname) is not set", 1) <> 0 Then
            Dim sNetBiosName As String
            sNetBiosName = GetCompName(ComputerNameNetBIOS)
            If sNetBiosName = "" Then
                sNetBiosName = Environ("USERDOMAIN")
                If sNetBiosName = "" Then
                    sNetBiosName = "USER-PC"
                End If
                SetCompName ComputerNamePhysicalNetBIOS, sNetBiosName
            End If
            SetCompName ComputerNamePhysicalDnsHostname, sNetBiosName
            bRebootRequired = True
        End If
        
    Else
        FixRegistryHandler Result
        bUpdatePolicyNeeded = True
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO7Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RunCleanMgr()

    'https://winaero.com/blog/cleanmgr-exe-command-line-arguments-in-windows-10/

    Dim sRootKey As String
    Dim cKeys As Collection
    Dim nid As Long
    Dim i As Long
    Dim sKey As String
    Dim sParam As String
    Dim lData As Long
    Dim sCleanMgr As String
    
    'all, except:
    'Recycle Bin
    'Windows update report
    'System crash dump
    'Remote Desktop Cache Files
    'Windows EDS files (needed to Refresh or Reset PC on Win 8/10)
    
    '//TODO:
    'Добавь чистку папки c:\Windows\Installer от устаревших обновлений офиса.
    'На старых машинах там до 1 гб хлама. Чистильщик есть в моих сборках офиса - по флагу state реестра. => look Sources\Cleaner
    'Некоторые чистят папку c:\Windows\SoftwareDistribution\Download, но я не советую, т.к. будут проблемы
    'с ручным удалением обновлений из апплета "установка и удаление".
    'Кто-то чистит c:\Windows\winsxs\Backup, c:\Windows\winsxs\Temp,
    'но при их очистке возможны проблемы с откатом системы на ранние точки.
    
    Set cKeys = New Collection
    
    nid = 777
    sRootKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
    cKeys.Add 2, "Active Setup Temp Folders"
    cKeys.Add 2, "BranchCache" '8/10
    cKeys.Add 2, "Compress old files" 'XP
    cKeys.Add 0, "Content Indexer Cleaner"
    cKeys.Add 2, "Downloaded Program Files"
    cKeys.Add 2, "GameUpdateFiles"
    cKeys.Add 2, "Internet Cache Files"
    cKeys.Add 2, "Memory Dump Files"
    cKeys.Add 2, "Offline Pages Files"
    cKeys.Add 2, "Old ChkDsk Files"
    cKeys.Add 2, "Previous Installations"
    cKeys.Add 0, "Recycle Bin"
    cKeys.Add 0, "Remote Desktop Cache Files" 'XP
    cKeys.Add 2, "RetailDemo Offline Content" '8/10
    cKeys.Add 2, "Service Pack Cleanup"
    cKeys.Add 0, "Setup Log Files"
    cKeys.Add 0, "System error memory dump files"
    cKeys.Add 0, "System error minidump files"
    cKeys.Add 2, "Temporary Files"
    cKeys.Add 2, "Temporary Setup Files"
    cKeys.Add 2, "Thumbnail Cache"
    cKeys.Add 2, "Update Cleanup"
    cKeys.Add 2, "Windows Defender" '8/10
    cKeys.Add 2, "User file versions" '8/10
    cKeys.Add 2, "Upgrade Discarded Files"
    cKeys.Add 2, "WebClient and WebPublisher Cache" 'XP
    cKeys.Add 2, "Windows Error Reporting Archive Files"
    cKeys.Add 2, "Windows Error Reporting Queue Files"
    cKeys.Add 2, "Windows Error Reporting System Archive Files"
    cKeys.Add 2, "Windows Error Reporting System Queue Files"
    cKeys.Add 2, "Windows Error Reporting Temp Files" '8/10
    cKeys.Add 0, "Windows ESD installation files"
    cKeys.Add 0, "Windows Upgrade Log Files"

    sParam = "StateFlags" & Right$("000" & nid, 4)
    'set preset
    For i = 1 To cKeys.Count
        lData = CLng(cKeys(i))
        sKey = sRootKey & "\" & GetCollectionKeyByIndex(i, cKeys)
        
        If Reg.KeyExists(0, sKey) Then
            Call Reg.SetDwordVal(0, sKey, sParam, lData)
        End If
    Next
    'run cleaner

    sCleanMgr = sSysNativeDir & "\CleanMgr.exe"
    
    If Proc.ProcessRun(sCleanMgr, "/SAGERUN:" & nid) Then
        Do While Proc.IsRunned
            'Please, wait until Microsoft disk cleanup manager finish its work and press OK.
            MsgBoxW TranslateNative(351), vbInformation
        Loop
    End If
    'remove preset
    For i = 1 To cKeys.Count
        sKey = sRootKey & "\" & GetCollectionKeyByIndex(i, cKeys)
        
        If Reg.KeyExists(0, sKey) Then
            Call Reg.DelVal(0, sKey, sParam)
        End If
    Next
    Set cKeys = Nothing
End Sub

Public Sub CheckO8Item()
    'O8 - Extra context menu items
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt
    'HKLM\Software\Microsoft\Internet Explorer\MenuExt
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO8Item - Begin"
    
    Dim hKey&, i&, sName$, lpcName&, sFile$, sHit$, Result As SCAN_RESULT, pos&, bSafe As Boolean
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Internet Explorer\MenuExt"
    
    Do While HE.MoveNext
    
            If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, _
              KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
              
                i = 0
                sName = String$(MAX_KEYNAME, 0&)
                lpcName = Len(sName)
        
                Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
                    sName = RTrimNull(sName)
                    sFile = Reg.GetString(HE.Hive, HE.Key & "\" & sName, vbNullString, HE.Redirected)
            
                    If Len(sFile) = 0 Then
                        sFile = "(no file)"
                    Else
                        If InStr(1, sFile, "res://", vbTextCompare) = 1 Then
                            sFile = Mid$(sFile, 7)
                        End If
                
                        If InStr(1, sFile, "file://", vbTextCompare) = 1 Then
                            sFile = Mid$(sFile, 8)
                        End If
                        
                        pos = InStrRev(sFile, "/")
                        If pos <> 0 Then sFile = Left$(sFile, pos - 1)
                        
                        pos = InStrRev(sFile, "?")
                        If pos <> 0 Then sFile = Left$(sFile, pos - 1)
                        
                        sFile = FormatFileMissing(sFile)
                    End If
            
                    sHit = "O8 - " & HE.HiveNameAndSID & "\..\Extra context menu item: " & sName & " - " & sFile
                    
                    bSafe = False
                    If WhiteListed(sFile, "EXCEL.EXE", True) Then bSafe = True 'MS Office
                    If WhiteListed(sFile, "ONBttnIE.dll", True) Then bSafe = True 'MS Office
                    
                    If Not IsOnIgnoreList(sHit) And (Not bSafe) Then
                        If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                        With Result
                            .Section = "O8"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                    
                    sName = String$(MAX_KEYNAME, 0&)
                    lpcName = Len(sName)
                    i = i + 1
                Loop
                RegCloseKey hKey
            End If
        
    Loop
    
    AppendErrorLogCustom "CheckO8Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO8Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO8Item(sItem$, Result As SCAN_RESULT)
    'O8 - Extra context menu items
    'O8 - Extra context menu item: [name] - html file
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt
    
    On Error GoTo ErrorHandler:
    
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO8Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO9Item()
    'HKLM\Software\Microsoft\Internet Explorer\Extensions
    'HKCU\..\etc
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO9Item - Begin"
    
    Dim hKey&, i&, sData$, sCLSID$, sCLSID2$, lpcName&, sFile$, sHit$, sBuf$, Result As SCAN_RESULT
    Dim pos&, bSafe As Boolean
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Internet Explorer\Extensions"
    
    Do While HE.MoveNext
    
    'open root key
    If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
        i = 0
        sCLSID = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sCLSID)
        'start enum of root key subkeys (i.e., extensions)
        Do While RegEnumKeyExW(hKey, i, StrPtr(sCLSID), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
            sCLSID = TrimNull(sCLSID)
            If sCLSID = "CmdMapping" Then GoTo NextExt:
            
            'check for 'MenuText' or 'ButtonText'
            sData = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "ButtonText", HE.Redirected)
            
            'this clsid is mostly useless, always pointing to SHDOCVW.DLL
            'places to look for correct dll:
            '* Exec
            '* Script
            '* BandCLSID
            '* CLSIDExtension
            '* CLSIDExtension -> TreatAs CLSID
            '* CLSID
            '* ???
            '* actual CLSID of regkey (not used)
            sFile = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "Exec", HE.Redirected)
            If sFile = vbNullString Then
                sFile = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "Script", HE.Redirected)
                If sFile = vbNullString Then
                    sCLSID2 = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "BandCLSID", HE.Redirected)
                    sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, HE.Redirected)
                    If sFile = vbNullString Then
                        sCLSID2 = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "CLSIDExtension", HE.Redirected)
                        sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, HE.Redirected)
                        If sFile = vbNullString Then
                            sCLSID2 = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\TreatAs", vbNullString, HE.Redirected)
                            sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, HE.Redirected)
                            If sFile = vbNullString Then
                                sCLSID2 = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "CLSID", HE.Redirected)
                                sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, HE.Redirected)
                            End If
                        End If
                    End If
                End If
            End If
            
            If Len(sFile) = 0 Then
                sFile = "(no file)"
            Else
                'expand %systemroot% var
                'sFile = replace$(sFile, "%systemroot%", sWinDir, , , vbTextCompare)
                sFile = UnQuote(EnvironW(sFile))
                
                'strip stuff from res://[dll]/page.htm to just [dll]
                If InStr(1, sFile, "res://", vbTextCompare) = 1 And _
                   (LCase$(Right$(sFile, 4)) = ".htm" Or LCase$(Right$(sFile, 4)) = "html") Then
                    sFile = Mid$(sFile, 7)
                End If
                
                'remove other stupid prefixes
                If InStr(1, sFile, "file://", vbTextCompare) = 1 Then
                    sFile = Mid$(sFile, 8)
                End If
                
                pos = InStrRev(sFile, "/")
                If pos <> 0 Then sFile = Left$(sFile, pos - 1)
                
                pos = InStrRev(sFile, "?")
                If pos <> 0 Then sFile = Left$(sFile, pos - 1)
                
                If InStr(1, sFile, "http:", 1) <> 1 And _
                  InStr(1, sFile, "https:", 1) <> 1 Then
                    '8.3 -> Full
                    If FileExists(sFile) Then
                        sFile = GetLongPath(EnvironW(sFile))
                    Else
                        sFile = GetLongPath(EnvironW(sFile)) & " (file missing)"
                    End If
                End If
            End If
            
            bSafe = False
            If Not bIgnoreAllWhitelists And bHideMicrosoft Then
                If WhiteListed(sFile, PF_64 & "\Messenger\msmsgs.exe") Then bSafe = True
                
                If OSver.MajorMinor = 5 Then 'Win2k
                    If StrComp(sFile, sWinDir & "\web\related.htm", 1) = 0 Then bSafe = True
                End If
                If OSver.MajorMinor <= 5.2 Then 'win2k/xp/2003
                    If WhiteListed(sFile, sWinDir & "\Network Diagnostic\xpnetdiag.exe") Then bSafe = True
                End If
                If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                    If IsMicrosoftFile(sFile) Then bSafe = True
                End If
            End If
            
            If Not bSafe Then
            
              If sData = vbNullString Then sData = "(no name)"
              If Left$(sData, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sData)
                If 0 <> Len(sBuf) Then sData = sBuf
              End If
            
              'O9 - Extra button:
              'O9-32 - Extra button:
              sHit = IIf(bIsWin32, "O9", IIf(HE.Redirected, "O9-32", "O9")) & _
                " - Extra button: " & sData & " - " & HE.HiveNameAndSID & "\..\" & sCLSID & " - " & sFile
              
              If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                With Result
                    .Section = "O9"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sCLSID, , , HE.Redirected
                    AddRegToFix .Reg, REMOVE_VALUE, HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", sCLSID
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
              End If
            
              sData = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "MenuText", HE.Redirected)
            
              If Left$(sData, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sData)
                If 0 <> Len(sBuf) Then sData = sBuf
              End If

              bSafe = False
            
              If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                If OSver.MajorMinor = 5 Then 'Win2k
                  If StrComp(sFile, sWinDir & "\web\related.htm", 1) = 0 Then bSafe = True
                End If
                If OSver.MajorMinor <= 5.2 Then 'win2k/xp/2003
                    If WhiteListed(sFile, sWinDir & "\Network Diagnostic\xpnetdiag.exe") Then bSafe = True
                End If
                If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                    If IsMicrosoftFile(sFile) Then bSafe = True
                End If
              End If
            
              'don't show it again in case sdata=null
              If sData <> vbNullString And Not bSafe Then
                'O9 - Extra 'Tools' menuitem:
                'O9-32 - Extra 'Tools' menuitem:
                sHit = IIf(bIsWin32, "O9", IIf(HE.Redirected, "O9-32", "O9")) & _
                  " - Extra 'Tools' menuitem: " & sData & " - " & HE.HiveNameAndSID & "\..\" & sCLSID & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    With Result
                        .Section = "O9"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sCLSID, , , HE.Redirected
                        AddRegToFix .Reg, REMOVE_VALUE, HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", sCLSID
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
              End If
            End If
NextExt:
            sCLSID = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sCLSID)
            i = i + 1
        Loop
        RegCloseKey hKey
    End If
    Loop
    
    AppendErrorLogCustom "CheckO9Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO9Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO9Item(sItem$, Result As SCAN_RESULT)
    'O9 - Extra buttons/Tools menu items
    'O9 - Extra button: [name] - [CLSID] - [file] [(HKCU)]
    
    On Error GoTo ErrorHandler:

    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO9Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO10Item()
    CheckLSP
End Sub

Public Sub CheckO11Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO11Item - Begin"
    
    'HKLM\Software\Microsoft\Internet Explorer\AdvancedOptions
    Dim hKey&, i&, sSubKey$, sName$, lpcName&, sHit$, Result As SCAN_RESULT
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Internet Explorer\AdvancedOptions"
    
    Do While HE.MoveNext
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
        
            sSubKey = String$(MAX_KEYNAME, 0)
            lpcName = Len(sSubKey)
            i = 0
            Do While RegEnumKeyExW(hKey, i, StrPtr(sSubKey), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
                sSubKey = TrimNull(sSubKey)
                
                If InStr("JAVA_VM.JAVA_SUN.BROWSE.ACCESSIBILITY.SEARCHING." & _
                  "HTTP1.1.MULTIMEDIA.Multimedia.CRYPTO.PRINT." & _
                  "TOEGANKELIJKHEID.TABS.INTERNATIONAL*.ACCELERATED_GRAPHICS", sSubKey) = 0 Then
                  
                    sName = Reg.GetString(HE.Hive, HE.Key & "\" & sSubKey, "Text", HE.Redirected)
                  
                    If Len(sName) <> 0 Then
                        'O11 - Options group:
                        'O11-32 - Options group:
                        sHit = IIf(bIsWin32, "O11", IIf(HE.Redirected, "O11-32", "O11")) & _
                          " - " & HE.HiveNameAndSID & "\..\Options group: [" & sSubKey & "] " & sName
                
                        If bIgnoreAllWhitelists Or Not IsOnIgnoreList(sHit) Then
                            With Result
                                .Section = "O11"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sSubKey, , , HE.Redirected
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults Result
                        End If
                    End If
                End If
                sSubKey = String$(MAX_KEYNAME, 0&)
                lpcName = Len(sSubKey)
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    AppendErrorLogCustom "CheckO11Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO11Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO11Item(sItem$, Result As SCAN_RESULT)
    'O11 - Options group: [BLA] Blah"
    On Error GoTo ErrorHandler:
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO11Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO12Item()
    'HKLM\Software\Microsoft\Internet Explorer\Plugins\Extensions
    'HKLM\Software\Microsoft\Internet Explorer\Plugins\MIME
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO12Item - Begin"
    
    Dim hKey&, i&, sName$, sData$, sFile$, sArgs$, sHit$, lpcName&, Result As SCAN_RESULT
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Internet Explorer\Plugins\Extension"
    HE.AddKey "Software\Microsoft\Internet Explorer\Plugins\MIME"
    
    Do While HE.MoveNext
      
      If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
      
        sName = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sName)
        i = 0
        
        Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
            sName = TrimNull(sName)
            sData = Reg.GetString(HE.Hive, HE.Key & "\" & sName, "Location", HE.Redirected)
            
            SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
            sFile = FormatFileMissing(sFile)
            
            'O12 - Plugin
            'O12-32 - Plugin
            sHit = IIf(bIsWin32, "O12", IIf(HE.Redirected, "O12-32", "O12")) & " - " & _
              HE.HiveNameAndSID & "\..\Plugin for " & sName & ": " & ConcatFileArg(sFile, sArgs)
              
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                With Result
                    .Section = "O12"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName, , , HE.Redirected
                    AddFileToFix .File, REMOVE_FILE, sFile
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults Result
            End If
            
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            i = i + 1
        Loop
        RegCloseKey hKey
      End If
    Loop
    
    AppendErrorLogCustom "CheckO12Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO12Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO12Item(sItem$, Result As SCAN_RESULT)
    'O12 - Plugin for .ofb: C:\Win98\blah.dll
    'O12 - Plugin for text/blah: C:\Win98\blah.dll
    
    On Error GoTo ErrorHandler:
    
    If Not bShownToolbarWarning And ProcessExist("iexplore.exe", True) Then
        MsgBoxW Translate(330), vbExclamation
'        msgboxW "HiJackThis is about to remove a " & _
'               "plugin from " & _
'               "your system. Close all Internet " & _
'               "Explorer windows before continuing for " & _
'               "the best chance of success.", vbExclamation
        bShownToolbarWarning = True
    End If
    
    FixRegistryHandler Result
    FixFileHandler Result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO12Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO13Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO13Item - Begin"
    
    Dim sDummy$, sHit$, Result As SCAN_RESULT
    Dim aKey() As String, aVal() As String, aExa() As String, aDes() As String, i As Long
    
    ReDim aKey(6)
    ReDim aVal(UBound(aKey))
    ReDim aExa(UBound(aKey))
    ReDim aDes(UBound(aKey))
    
    aKey(0) = "DefaultPrefix"
    aVal(0) = ""
    aExa(0) = "http://"
    aDes(0) = "DefaultPrefix"
    
    aKey(1) = "Prefixes"
    aVal(1) = "www"
    aExa(1) = "http://"
    aDes(1) = "WWW Prefix"
    
    aKey(2) = "Prefixes"
    aVal(2) = "www."
    aExa(2) = ""
    aDes(2) = "WWW. Prefix"
    
    aKey(3) = "Prefixes"
    aVal(3) = "home"
    aExa(3) = "http://"
    aDes(3) = "Home Prefix"
    
    aKey(4) = "Prefixes"
    aVal(4) = "mosaic"
    aExa(4) = "http://"
    aDes(4) = "Mosaic Prefix"
    
    aKey(5) = "Prefixes"
    aVal(5) = "ftp"
    aExa(5) = "ftp://"
    aDes(5) = "FTP Prefix"
    
    aKey(6) = "Prefixes"
    aVal(6) = "gopher"
    aExa(6) = "gopher://|"
    aDes(6) = "Gopher Prefix"
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\URL"

    Do While HE.MoveNext
    
        For i = 0 To UBound(aKey)
        
            sDummy = Reg.GetString(HE.Hive, HE.Key & "\" & aKey(i), aVal(i), HE.Redirected)
            
            'exclude empty HKCU / HKU
            If Not (HE.Hive <> HKLM And sDummy = "") Then
            
                If Not inArraySerialized(sDummy, aExa(i), "|", , , vbBinaryCompare) Then
                    
                    sHit = IIf(bIsWin32, "O13", IIf(HE.Redirected, "O13-32", "O13")) & " - " & HE.HiveNameAndSID & "\..\" & aDes(i) & ": " & sDummy
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "O13"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key & "\" & aKey(i), aVal(i), aExa(i), HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckO13Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO13Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO13Item(sItem$, Result As SCAN_RESULT)
    'defaultprefix fix
    'O13 - DefaultPrefix: http://www.hijacker.com/redir.cgi?
    'O13 - [WWW/Home/Mosaic/FTP/Gopher] Prefix: ..
    
    On Error GoTo ErrorHandler:
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO13Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO14Item()
    'O14 - Reset Websettings check
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO14Item - Begin"
    
    Dim sLine$, sHit$, ff%
    Dim sStartPage$, sSearchPage$, sMsStartPage$
    Dim sSearchAssis$, sCustSearch$
    Dim sFile$, aLogStrings() As String, i&
    
    sFile = sWinDir & "\inf\iereset.inf"
    
    If Not FileExists(sFile) Then Exit Sub
    If FileLenW(sFile) = 0 Then Exit Sub
    
    Dim b() As Byte
    ReDim b(1)
    ff = FreeFile()
    Open sFile For Binary Access Read As #ff
    Get #ff, 1, b()
    Close #ff
    aLogStrings = ReadFileToArray(sFile, IIf(b(0) = &HFF& And b(1) = &HFE&, True, False))
    
    For i = 0 To UBound(aLogStrings)
        sLine = aLogStrings(i)
        
            If InStr(sLine, "SearchAssistant") > 0 Then
                sSearchAssis = Mid$(sLine, InStr(sLine, "http://"))
                sSearchAssis = Left$(sSearchAssis, Len(sSearchAssis) - 1)
            End If
            If InStr(sLine, "CustomizeSearch") > 0 Then
                sCustSearch = Mid$(sLine, InStr(sLine, "http://"))
                sCustSearch = Left$(sCustSearch, Len(sCustSearch) - 1)
            End If
            If InStr(sLine, "START_PAGE_URL=") = 1 And _
               InStr(sLine, "MS_START_PAGE_URL") = 0 Then
                sStartPage = Mid$(sLine, InStr(sLine, "=") + 1)
                sStartPage = UnQuote(sStartPage)
            End If
            If InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sSearchPage = Mid$(sLine, InStr(sLine, "=") + 1)
                sSearchPage = UnQuote(sSearchPage)
            End If
            If InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sMsStartPage = Mid$(sLine, InStr(sLine, "=") + 1)
                sMsStartPage = UnQuote(sMsStartPage)
            End If
    Next
    
    'SearchAssistant = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm
    If sSearchAssis <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm" And _
      sSearchAssis <> g_DEFSEARCHASS Then
        sHit = "O14 - IERESET.INF: SearchAssistant=" & sSearchAssis
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'CustomizeSearch = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm
    If sCustSearch <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm" And _
      sCustSearch <> g_DEFSEARCHCUST Then
        sHit = "O14 - IERESET.INF: CustomizeSearch=" & sCustSearch
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'SEARCH_PAGE_URL = http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch
    If sSearchPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch" And _
      sSearchPage <> g_DEFSEARCHPAGE Then
        sHit = "O14 - IERESET.INF: SEARCH_PAGE_URL=" & sSearchPage
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'START_PAGE_URL  = http://www.msn.com
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If sStartPage <> "http://www.msn.com" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome" And _
       sStartPage <> g_DEFSTARTPAGE Then
        sHit = "O14 - IERESET.INF: START_PAGE_URL=" & sStartPage
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'MS_START_PAGE_URL=http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '(=START_PAGE_URL) http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If sMsStartPage <> vbNullString Then
        If sMsStartPage <> "http://www.msn.com" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome" And _
           sMsStartPage <> g_DEFSTARTPAGE Then
            sHit = "O14 - IERESET.INF: MS_START_PAGE_URL=" & sMsStartPage
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
        End If
    End If
    
    AppendErrorLogCustom "CheckO14Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO14Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO14Item(sItem$, Result As SCAN_RESULT)
    'resetwebsettings fix
    'O14 - IERESET.INF: [item]=[URL]
    
    On Error GoTo ErrorHandler:
    'sItem - not used
    Dim sLine$, sFixedIeResetInf$, ff%
    Dim i&, aLogStrings() As String, sFile$, isUnicode As Boolean
    
    sFile = sWinDir & "\INF\iereset.inf"
    
    If Not FileExists(sFile) Then Exit Sub
    
    BackupFile Result, sFile
    
    ff = FreeFile()
    
    Dim b() As Byte
    ReDim b(1)
    ff = FreeFile()
    Open sFile For Binary Access Read As #ff
    Get #ff, 1, b()
    Close #ff
    If b(0) = &HFF& And b(1) = &HFE& Then isUnicode = True
    aLogStrings = ReadFileToArray(sFile, IIf(isUnicode, True, False))
    
    For i = 0 To UBound(aLogStrings)
        sLine = aLogStrings(i)

            If InStr(sLine, "SearchAssistant") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""SearchAssistant"",0,""" & _
                    IIf(g_DEFSEARCHASS <> "", g_DEFSEARCHASS, "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm") & """" & vbCrLf
            ElseIf InStr(sLine, "CustomizeSearch") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""CustomizeSearch"",0,""" & _
                    IIf(g_DEFSEARCHCUST <> "", g_DEFSEARCHCUST, "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm") & """" & vbCrLf
            ElseIf InStr(sLine, "START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "START_PAGE_URL=""" & _
                    IIf(g_DEFSTARTPAGE <> "", g_DEFSTARTPAGE, "http://www.msn.com") & """" & vbCrLf
            ElseIf InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "SEARCH_PAGE_URL=""" & _
                    IIf(g_DEFSEARCHPAGE <> "", g_DEFSEARCHPAGE, "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch") & """" & vbCrLf
            ElseIf InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "MS_START_PAGE_URL=""" & _
                    IIf(g_DEFSTARTPAGE <> "", g_DEFSTARTPAGE, "http://www.msn.com") & """" & vbCrLf
            Else
                sFixedIeResetInf = sFixedIeResetInf & sLine & vbCrLf
            End If
        
    Next
    sFixedIeResetInf = Left$(sFixedIeResetInf, Len(sFixedIeResetInf) - 2)   '-CrLf
    
    DeleteFileWEx (StrPtr(sFile))
    
    ff = FreeFile()
    
    If isUnicode Then
        b() = ChrW$(-257) & sFixedIeResetInf
        Open sFile For Binary Access Write As #ff
        Put #ff, , b()
    Else
        Open sFile For Output As #ff
        Print #ff, sFixedIeResetInf
    End If
    
    Close #ff
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO14Item", "sItem=", sItem
    Close #ff
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO15Item()
    'the value * / http / https denotes the protocol for which the rule is valid. It's:
    '2 for Trusted Zone;
    '4 for Restricted Zone.
    
    'Checks:
    '* ZoneMap\Domains          - trusted domains
    '* ZoneMap\Ranges           - trusted IPs and IP ranges
    '* ZoneMap\ProtocolDefaults - what zone rules does a protocol obey
    '* ZoneMap\EscDomains       - trusted domains for Enhanced Security Configuration
    '* ZoneMap\EscRanges        - trusted IPs and IP ranges for ESC
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO15Item - Begin"
    
    Dim sDomains$(), sSubDomains$(), vProtocol, sProtPrefix$
    Dim i&, j&, sHit$, sAlias$, sIPRange$, bSafe As Boolean, Result As SCAN_RESULT
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains"
    
    'enum all domains
    'value = http - only http on subdomain is trusted
    'value = https - only https on subdomain is trusted
    'value = * - both
    
    Do While HE.MoveNext
        sDomains = Split(Reg.EnumSubKeys(HE.Hive, HE.Key, HE.Redirected), "|")
        If UBound(sDomains) > -1 Then
            If StrEndWith(HE.Key, "EscDomains") Then
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - ESC Trusted Zone: "
            Else
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - Trusted Zone: "
            End If
            For i = 0 To UBound(sDomains)
                bSafe = False
                If bIgnoreSafeDomains And Not bIgnoreAllWhitelists Then
                    bSafe = StrBeginWithArray(sDomains(i), aSafeRegDomains)
                End If
                If Not bSafe Then
                    sSubDomains = Split(Reg.EnumSubKeys(HE.Hive, HE.Key & "\" & sDomains(i), HE.Redirected), "|")
                    If UBound(sSubDomains) <> -1 Then
                        'list any trusted subdomains for main domain
                        For j = 0 To UBound(sSubDomains)
                            For Each vProtocol In Array("*", "http", "https")
                                Select Case vProtocol
                                Case "*": sProtPrefix = "*."
                                Case "http": sProtPrefix = "http://"
                                Case "https": sProtPrefix = "https://"
                                End Select
                                If Reg.GetDword(HE.Hive, HE.Key & "\" & sDomains(i) & "\" & sSubDomains(j), CStr(vProtocol), HE.Redirected) = 2 Then
                                    sHit = sAlias & sProtPrefix & sSubDomains(j) & "." & sDomains(i)
                                    If Not IsOnIgnoreList(sHit) Then
                                        With Result
                                            .Section = "O15"
                                            .HitLineW = sHit
                                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i) & "\" & sSubDomains(j), , , HE.Redirected
                                            .CureType = REGISTRY_BASED
                                        End With
                                        AddToScanResults Result
                                    End If
                                End If
                            Next
                        Next
                    End If
                    'list main domain as well if that's trusted too (*grumble*)
                    For Each vProtocol In Array("*", "http", "https")
                        Select Case vProtocol
                        Case "*": sProtPrefix = "*."
                        Case "http": sProtPrefix = "http://"
                        Case "https": sProtPrefix = "https://"
                        End Select
                        If Reg.GetDword(HE.Hive, HE.Key & "\" & sDomains(i), CStr(vProtocol), HE.Redirected) = 2 Then
                            sHit = sAlias & HE.HiveNameAndSID & " - " & sProtPrefix & sDomains(i)
                            If Not IsOnIgnoreList(sHit) Then
                                With Result
                                    .Section = "O15"
                                    .HitLineW = sHit
                                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i), , , HE.Redirected
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults Result
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Loop
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscRanges"
    
    'enum all IP ranges
    Do While HE.MoveNext
        sDomains = Split(Reg.EnumSubKeys(HE.Hive, HE.Key, HE.Redirected), "|")
        If UBound(sDomains) > -1 Then
            If StrEndWith(HE.Key, "EscRanges") Then
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - ESC Trusted IP range: "
            Else
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - Trusted IP range: "
            End If
            For i = 0 To UBound(sDomains)
                sIPRange = Reg.GetString(HE.Hive, HE.Key & "\" & sDomains(i), ":Range", HE.Redirected)
                If Left$(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                    For Each vProtocol In Array("*", "http", "https")
                        Select Case vProtocol
                        Case "*": sProtPrefix = "*."
                        Case "http": sProtPrefix = "http://"
                        Case "https": sProtPrefix = "https://"
                        End Select
                        If Reg.GetDword(HE.Hive, HE.Key & "\" & sDomains(i), CStr(vProtocol), HE.Redirected) = 2 Then
                            sHit = sAlias & HE.HiveNameAndSID & " - " & sProtPrefix & sIPRange
                            If Not IsOnIgnoreList(sHit) Then
                                With Result
                                    .Section = "O15"
                                    .HitLineW = sHit
                                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i), , , HE.Redirected
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults Result
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Loop
    
    'check all ProtocolDefaults values
    
    Dim sZoneNames$(), sProtVals$(), lProtZoneDefs&(11), lProtZones&(11), LastIndex&
    sZoneNames = Split("My Computer|Intranet|Trusted|Internet|Restricted|Unknown", "|")
    sProtVals = Split("@ivt|file|ftp|http|https|shell|ldap|news|nntp|oecmd|snews|knownfolder", "|")
    
    lProtZoneDefs(0) = 1 '@ivt '2k+
    lProtZoneDefs(1) = 3 'file '2k+
    lProtZoneDefs(2) = 3 'ftp '2k+
    lProtZoneDefs(3) = 3 'http '2k+
    lProtZoneDefs(4) = 3 'https '2k+
    lProtZoneDefs(5) = 0 'shell 'XP+
    lProtZoneDefs(6) = 4 'ldap '(HKLM only) 'Vista+
    lProtZoneDefs(7) = 4 'news '(HKLM only) 'Vista+
    lProtZoneDefs(8) = 4 'nntp '(HKLM only) 'Vista+
    lProtZoneDefs(9) = 4 'oecmd '(HKLM only) 'Vista+
    lProtZoneDefs(10) = 4 'snews '(HKLM only) 'Vista+
    lProtZoneDefs(11) = 0 'knownfolder '7+
    
    If OSver.MajorMinor = 5 Then 'Win2k
        HE.Init HE_HIVE_ALL, HE_SID_USER
    ElseIf OSver.MajorMinor = 5.2 And OSver.IsServer Then 'Win 2003, 2003 R2
        HE.Init HE_HIVE_ALL, HE_SID_USER
    Else
        HE.Init HE_HIVE_ALL, HE_SID_DEFAULT Or HE_SID_USER
    End If
    
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\ProtocolDefaults"
    
    If OSver.IsWindows7OrGreater Then
        LastIndex = 11
    ElseIf OSver.IsWindowsVistaOrGreater Then
        LastIndex = 10
    Else
        LastIndex = 5
    End If
    
    Do While HE.MoveNext
        For i = 0 To LastIndex
            bSafe = False
            lProtZones(i) = Reg.GetDword(HE.Hive, HE.Key, sProtVals(i), HE.Redirected)
            
            If lProtZones(i) = 0 Then
                If Not Reg.ValueExists(HE.Hive, HE.Key, sProtVals(i), HE.Redirected) Then
                    If i >= 6 And i <= 10 And HE.Hive <> HKLM Then
                        bSafe = True
                    Else
                        lProtZones(i) = 5 'Unknown
                    End If
                End If
                
                If lProtZones(i) = 5 Then
                    If sProtVals(i) = "knownfolder" And OSver.MajorMinor = 6.1 Then bSafe = True
                End If
            End If
            
            If Not bSafe Then
                If lProtZones(i) < 0 Or lProtZones(i) > 5 Then lProtZones(i) = 5 'Unknown
                If lProtZones(i) <> lProtZoneDefs(i) Then 'check for legit
                    sHit = IIf(HE.Redirected, "O15-32", "O15") & " - ProtocolDefaults: " & HE.HiveNameAndSID & _
                        " - '" & sProtVals(i) & "' protocol is in " & sZoneNames(lProtZones(i)) & " Zone, should be " & sZoneNames(lProtZoneDefs(i)) & " Zone"
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "O15"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sProtVals(i), lProtZoneDefs(i), HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckO15Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO15Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO15Item(sItem$, Result As SCAN_RESULT)
'    'O15 - Trusted Zone: free.aol.com (HKLM)
'    'O15 - Trusted Zone: http://free.aol.com
'    'O15 - Trusted IP range: 66.66.66.66 (HKLM)
'    'O15 - Trusted IP range: http://66.66.66.*
'    'O15 - ESC Trusted Zone: free.aol.com (HKLM)
'    'O15 - ESC Trusted IP range: 66.66.66.66
'    'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)

    On Error GoTo ErrorHandler:
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO15Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

'Public Sub FixNetscapeMozilla(sItem$)
'    'N1 - Netscape 4: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'    'N2 - Netscape 6: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'    'N3 - Netscape 7: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'    'N4 - Mozilla: user_pref("browser.startup.homepage", "http://url"); (c:\..\prefs.js)
'    '               user_pref("browser.search.defaultengine", "http://url"); (c:\..\prefs.js)
'
'    Dim sPrefsJs$, sDummy$, ff1%, ff2%
'    On Error GoTo ErrorHandler:
'    sPrefsJs = Mid$(sItem, InStrRev(sItem, "(") + 1)
'    sPrefsJs = Left$(sPrefsJs, Len(sPrefsJs) - 1)
'    If FileExists(sPrefsJs) Then
'        ff1 = FreeFile()
'        Open sPrefsJs For Input As #ff1
'        ff2 = FreeFile()
'        Open sPrefsJs & ".new" For Output As #ff2
'            Do
'                Line Input #ff1, sDummy
'                If InStr(sDummy, "user_pref(""browser.startup.homepage"",") > 0 And _
'                   InStr(sItem, "user_pref(""browser.startup.homepage"",") > 0 Then
'                    Print #ff2, "user_pref(""browser.startup.homepage"", ""http://home.netscape.com/"");"
'                ElseIf InStr(sDummy, "user_pref(""browser.search.defaultengine"",") > 0 And _
'                   InStr(sItem, "user_pref(""browser.search.defaultengine"",") > 0 Then
'                    Print #ff2, "user_pref(""browser.search.defaultengine"", ""http://www.google.com/"");"
'                Else
'                    Print #ff2, sDummy
'                End If
'            Loop Until EOF(ff1)
'        Close #ff1
'        Close #ff2
'        deletefileWEx (StrPtr(sPrefsJs))
'        Name sPrefsJs & ".new" As sPrefsJs
'    End If
'    Exit Sub
'
'ErrorHandler:
'    Close #ff1
'    Close #ff2
'    ErrorMsg Err, "modMain_FixNetscapeMozilla", "sItem=", sItem
'    If inIDE Then Stop: Resume Next
'End Sub

Public Sub CheckO16Item()
    'O16 - Downloaded Program Files
    
    'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ActiveXCache
    'is location of actual %WINDIR%\DPF\ folder
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO16Item - Begin"
    
    Dim sName$, sFriendlyName$, sCodebase$, i&, hKey&, lpcName&, sHit$, Result As SCAN_RESULT
    Dim sOSD$, sInf$, sInProcServer32$
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Code Store Database\Distribution Units"
    
    Do While HE.MoveNext
    
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
    
            sName = String$(MAX_KEYNAME, 0)
            lpcName = Len(sName)
            i = 0
    
            Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
      
                sName = Left$(sName, InStr(sName, vbNullChar) - 1)
        
                sCodebase = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\DownloadInformation", "CODEBASE", HE.Redirected)
        
                If (InStr(sCodebase, "http://www.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://webresponse.one.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://rtc.webresponse.one.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://office.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://officeupdate.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://protect.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://dql.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://codecs.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://download.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://windowsupdate.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://v4.windowsupdate.microsoft.com") <> 1) _
                  Or bIgnoreAllWhitelists Then
           
                  'InStr(sCodeBase, "http://java.sun.com") <> 1 And _
                  'InStr(sCodeBase, "http://download.macromedia.com") <> 1 And _
                  'InStr(sCodeBase, "http://fpdownload.macromedia.com") <> 1 And _
                  'InStr(sCodeBase, "http://active.macromedia.com") <> 1 And _
                  'InStr(sCodeBase, "http://www.apple.com") <> 1 And _
                  'InStr(sCodeBase, "http://http://security.symantec.com") <> 1 And _
                  'InStr(sCodeBase, "http://download.yahoo.com") <> 1 And _
                  'InStr(sName, "Microsoft XML Parser") = 0 And _
                  'InStr(sName, "Java Classes") = 0 And _
                  'InStr(sName, "Classes for Java") = 0 And _
                  'InStr(sName, "Java Runtime Environment") = 0 Or _

                  'a DPF object can consist of:
                  '* DPF regkey           -> sDPFKey
                  '* CLSID regkey         -> CLSID\ & sName
                  '* OSD file             -> sOSD = Reg.GetString
                  '* INF file             -> sINF = Reg.GetString
                  '* InProcServer32 file  -> sIPS = Reg.GetString
                    
                    If Left$(sName, 1) = "{" And Right$(sName, 1) = "}" Then
                        Call GetFileByCLSID(sName, sInProcServer32, sFriendlyName, HE.Redirected, HE.SharedKey)
                    End If
                    
                    'not http ?
                    If Mid$(sCodebase, 2, 1) = ":" Then
                        sCodebase = FormatFileMissing(sCodebase)
                    End If
                    
                    ' "O16 - DPF: "
                    sHit = IIf(bIsWin32, "O16", IIf(HE.Redirected, "O16-32", "O16")) & " - DPF: " & HE.HiveNameAndSID & "\..\" & _
                      sName & " - " & sFriendlyName & " - " & sCodebase
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "O16"
                            .HitLineW = sHit
                            
                            sOSD = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\DownloadInformation", "OSD", HE.Redirected)
                            sInf = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\DownloadInformation", "INF", HE.Redirected)
                            
                            AddFileToFix .File, REMOVE_FILE Or UNREG_DLL, sInProcServer32
                            AddFileToFix .File, REMOVE_FILE, sOSD
                            AddFileToFix .File, REMOVE_FILE, sInf
                            
                            AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sName, , , REG_REDIRECTION_BOTH
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName, , , HE.Redirected
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
                
                i = i + 1
                sName = String$(MAX_KEYNAME, 0)
                lpcName = Len(sName)
            
            Loop
            
            RegCloseKey hKey
        End If
    Loop
    
    AppendErrorLogCustom "CheckO16Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO16Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO16Item(sItem$, Result As SCAN_RESULT)
    'O16 - DPF: {0000000} (shit toolbar) - http://bla.com/bla.dll
    'O16 - DPF: Plugin - http://bla.com/bla.dll
    
    On Error GoTo ErrorHandler:
    FixFileHandler Result
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO16Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO17Item()
    'check 'domain' and 'domainname' values in:
    'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters
    'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\*
    'HKLM\Software\Microsoft\Windows\CurrentVersion\Telephony
    'HKLM\System\CurrentControlSet\Services\VxD\MSTCP
    'and all values in other ControlSet's as well
    '
    'new one from UltimateSearch: value 'SearchList' in
    'HKLM\System\CurrentControlSet\Services\VxD\MSTCP
    '
    'just in case: NameServer as well, CoolWebSearch
    'maybe using this
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO17Item - Begin"
    
    Dim hKey&, i&, j&, sDomain$, sHit$, sParam$, vParam, CSKey$, n&, sData$, aNames() As String
    Dim UseWow, Wow6432Redir As Boolean, Result As SCAN_RESULT, Data() As String, sTrimChar As String
    Dim TcpIpNameServers() As String: ReDim TcpIpNameServers(0)
    ReDim sKeyDomain(0 To 1) As String
    Dim sProviderDNS As String
    Dim sSymTarget As String, ccsID As Long
    
'    'get target of CurrentControlSet
'    If Reg.IsKeySymLink(HKLM, "SYSTEM\CurrentControlSet", sSymTarget) Then
'        If IsNumeric(Right$(sSymTarget, 3)) Then
'            ccsID = CLng(Right$(sSymTarget, 3))
'        End If
'    End If

    ccsID = Reg.GetDword(HKLM, "SYSTEM\Select", "Current")
    
    'these keys are x64 shared
    sKeyDomain(0) = "Services\Tcpip\Parameters"
    sKeyDomain(1) = "Services\VxD\MSTCP"
    
    For j = 0 To 999    ' 0 - is CCS
    
        CSKey = IIf(j = 0, "System\CurrentControlSet", "System\ControlSet" & Format$(j, "000"))
        
        If j > 0 Then
            If Not Reg.KeyExists(HKEY_LOCAL_MACHINE, CSKey) Then Exit For
            If j = ccsID Then GoTo Continue
        End If
    
        For Each vParam In Array("Domain", "DomainName", "SearchList", "NameServer")
            sParam = vParam
            
            For n = 0 To UBound(sKeyDomain)
                'HKLM\System\CCS\Services\Tcpip\Parameters,Domain
                'HKLM\System\CCS\Services\Tcpip\Parameters,DomainName
                'HKLM\System\CCS\Services\VxD\MSTCP,Domain
                'HKLM\System\CCS\Services\VxD\MSTCP,DomainName
                'new one from UltimateSearch!
                'HKLM\System\CCS\Services\VxD\MSTCP,SearchList
                'HKLM\System\CCS\Services\VxD\MSTCP,SearchList
                'HKLM\System\CCS\Services\Tcpip\Parameters,SearchList
                'HKLM\System\CCS\Services\Tcpip\Parameters,NameServer
                sData = Reg.GetString(HKEY_LOCAL_MACHINE, CSKey & "\" & sKeyDomain(n), sParam)
                
                If Len(sData) <> 0 Then
                    sHit = "O17 - HKLM\" & IIf(j = 0, "System\CCS", CSKey) & "\" & sKeyDomain(n) & ": " & sParam & " = " & sData
                    
                    If sParam = "NameServer" Then
                        sProviderDNS = GetCollectionKeyByItemName(sData, colSafeDNS)
                        If sProviderDNS <> "" Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                    End If
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With Result
                            .Section = "O17"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, CSKey & "\" & sKeyDomain(n), sParam
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            Next
            
            'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\.. subkeys
            'HKLM\System\CS*\Services\Tcpip\Parameters\Interfaces\.. subkeys
            Erase aNames
            For n = 1 To Reg.EnumSubKeysToArray(HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces", aNames)
                
                sData = Reg.GetString(HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces\" & aNames(n), sParam)
                If sData <> vbNullString Then
                
                    ReDim Data(0)
                    Data(0) = sData
                    
                    If sParam = "NameServer" Then
                        
                        'Split lines like:
                        'O17 - HKLM\System\CCS\Services\Tcpip\..\{19B2C21E-CA09-48A1-9456-E4191BE91F00}: NameServer = 89.20.100.53 83.219.25.69
                        'O17 - HKLM\System\CCS\Services\Tcpip\..\{2A220B45-7A12-4A0B-92F0-00254794215A}: NameServer = 192.168.1.1,8.8.8.8
                        'into several separate
                        sData = Trim$(sData)
                        If InStr(sData, " ") <> 0 Then
                            Data = Split(sData)
                            sTrimChar = " "
                        ElseIf InStr(sData, ",") <> 0 Then
                            Data = Split(sData, ",")
                            sTrimChar = ","
                        End If
                        
                        For i = 0 To UBound(Data)
                            ReDim Preserve TcpIpNameServers(UBound(TcpIpNameServers) + 1)   'for using in filtering DNS DHCP later
                            TcpIpNameServers(UBound(TcpIpNameServers)) = Data(i)
                        Next
                    End If
                    
                    For i = 0 To UBound(Data)
                        
                        sHit = "O17 - HKLM\" & IIf(j = 0, "System\CCS", CSKey) & "\Services\Tcpip\..\" & aNames(n) & ": " & sParam & " = " & Data(i)
                        
                        If sParam = "NameServer" Then
                            sProviderDNS = GetCollectionKeyByItemName(CStr(Data(i)), colSafeDNS)
                            If sProviderDNS <> "" Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                        End If
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With Result
                                .Section = "O17"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                    HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces\" & aNames(n), sParam, _
                                    , , , CStr(Data(i)), "", sTrimChar
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults Result
                        End If
                    Next
                End If
            Next
        Next
Continue:
    Next
    
    Dim sTelephonyDomain$
    sTelephonyDomain = "Software\Microsoft\Windows\CurrentVersion\Telephony"
    
    For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If bIsWin32 And Wow6432Redir Then Exit For
    
        'HKLM\Software\MS\Windows\CurVer\Telephony,Domain
        'HKLM\Software\MS\Windows\CurVer\Telephony,DomainName
        For Each vParam In Array("Domain", "DomainName")
            sParam = vParam
            sDomain = Reg.GetString(HKEY_LOCAL_MACHINE, sTelephonyDomain, sParam, Wow6432Redir)
            If sDomain <> vbNullString Then
                'O17 - HKLM\Software\..\Telephony:
                sHit = IIf(bIsWin32, "O17", IIf(Wow6432Redir, "O17-32", "O17")) & " - HKLM\Software\..\Telephony: " & sParam & " = " & sDomain
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Section = "O17"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, sTelephonyDomain, sParam
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
        Next
    Next
    
    '------------------------------------------------------------
    
    Dim DNS() As String
    
    If GetDNS(DNS) Then
        For i = 0 To UBound(DNS)
            If Len(DNS(i)) <> 0 Then
                'If Not (DNS(i) = "192.168.0.1" Or DNS(i) = "192.168.1.1") Then
                    If Not inArray(DNS(i), TcpIpNameServers, , , vbTextCompare) Then
                        sHit = "O17 - DHCP DNS - " & i + 1 & ": " & DNS(i)
                        
                        sProviderDNS = GetCollectionKeyByItemName(DNS(i), colSafeDNS)
                        If sProviderDNS <> "" Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                        
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O17", sHit
                    End If
                'End If
            End If
        Next
    End If
    
    AppendErrorLogCustom "CheckO17Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO17Item"
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO17Item(sItem$, Result As SCAN_RESULT)
    'O17 - Domain hijack
    'O17 - HKLM\System\CCS\Services\VxD\MSTCP: Domain[Name] = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\Parameters: Domain[Name] = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\..\{0000}: Domain[Name] = blah
    '                  CS1
    '                  CS2
    '                  ...
    'O17 - HKLM\Software\..\Telephony: SearchList = blah
    'O17 - HKLM\System\CCS\Services\VxD\MSTCP: SearchList = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\Parameters: SearchList = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\..\{0000}: SearchList = blah
    '                  CS1
    '                  CS2
    '                  ...
    'ditto for NameServer
    
    On Error GoTo ErrorHandler:
    
    If StrBeginWith(sItem, "O17 - DHCP DNS:") Then
        'Cure for this object is not provided: []
        'You need to manually set the DNS address on the router, which is issued to you by provider.
        MsgBoxW Replace$(TranslateNative(349), "[]", sItem), vbExclamation
        Exit Sub
    End If
    
    FixRegistryHandler Result
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "modMain_FixO17Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO18Item()
    'enumerate everything in HKCR\Protocols\Handler
    'enumerate everything in HKCR\Protocols\Filters (section 2)
    'keys are x64 shared
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO18Item - Begin"
    
    Dim hKey&, i&, sName$, sCLSID$, sFile$, lpcName&, sHit$, Wow6432Redir As Boolean, Result As SCAN_RESULT
    Dim bShared As Boolean, vkt As KEY_VIRTUAL_TYPE, bSafe As Boolean
    
    Wow6432Redir = False
    
    vkt = Reg.GetKeyVirtualType(HKLM, "Software\Classes\Protocols\Handler")
    bShared = (vkt And KEY_VIRTUAL_SHARED)
    
    If RegOpenKeyExW(HKEY_CLASSES_ROOT, StrPtr("Protocols\Handler"), 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
      sName = String$(MAX_KEYNAME, 0&)
      lpcName = Len(sName)
      i = 0
      Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
        sName = TrimNull(sName)
        sCLSID = UCase$(Reg.GetString(HKEY_CLASSES_ROOT, "Protocols\Handler\" & sName, "CLSID", Wow6432Redir))
        
        If sCLSID <> "" Then
            Call GetFileByCLSID(sCLSID, sFile, , Wow6432Redir, bShared)
        End If
        If sCLSID = "" Then sCLSID = "(no CLSID)"
        If sFile = "" Then sFile = "(no file)"
        
        bSafe = False
        If bHideMicrosoft And Not bIgnoreAllWhitelists Then
            If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                If IsMicrosoftFile(sFile) Then bSafe = True
            End If
        End If
        
        If Not bSafe Then
          'for each protocol, check if name is on safe list
          If InStr(1, sSafeProtocols, sName, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
            sHit = "O18 - Protocol: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                With Result
                    .Section = "O18"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "Protocols\Handler\" & sName
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
          Else
            'and if so, check if CLSID is also on safe list
            '(no hijacker would hijack a protocol by
            'changing the CLSID to another safe one, right?)
            If sCLSID <> "(no CLSID)" Then
                If InStr(1, sSafeProtocols, sCLSID, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                
                    sHit = "O18 - Protocol hijack: " & sName & " - " & sCLSID
                    If Not IsOnIgnoreList(sHit) Then
                        If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                        
                        With Result
                            .Section = "O18"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, HKEY_CLASSES_ROOT, "Protocols\Handler\" & sName, "CLSID", O18_GetCLSIDByProtocol(sName)
                            If Not IsMicrosoftFile(sFile) Then
                                AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
                                If Not FileMissing(sFile) Then
                                    AddFileToFix .File, REMOVE_FILE, sFile
                                End If
                            End If
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
          End If
        End If
        '//TODO
        'checking file by CLSID
        
        sName = String$(MAX_KEYNAME, 0)
        lpcName = Len(sName)
        i = i + 1
      Loop
      RegCloseKey hKey
    End If
    
    '-------------------
    'Filters:
    
    hKey = 0
    sCLSID = vbNullString
    sFile = vbNullString
    
    vkt = Reg.GetKeyVirtualType(HKLM, "Software\Classes\Protocols\Filter")
    bShared = (vkt And KEY_VIRTUAL_SHARED)
    
    If RegOpenKeyExW(HKEY_CLASSES_ROOT, StrPtr("PROTOCOLS\Filter"), 0, KEY_ENUMERATE_SUB_KEYS, hKey) = 0 Then
      sName = String$(MAX_KEYNAME, 0&)
      lpcName = Len(sName)
      i = 0
      Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
        sName = TrimNull(sName)
        sCLSID = Reg.GetString(HKEY_CLASSES_ROOT, "PROTOCOLS\Filter\" & sName, "CLSID", Wow6432Redir)
        
        If sCLSID = "" Then
            sCLSID = "(no CLSID)"
            sFile = "(no file)"
        Else
            Call GetFileByCLSID(sCLSID, sFile, , Wow6432Redir, bShared)
        End If
        
        bSafe = False
        If bHideMicrosoft And Not bIgnoreAllWhitelists Then
            If InStr(1, sFile, "\Microsoft Shared\", 1) <> 0 Then
                If IsMicrosoftFile(sFile) Then bSafe = True
            End If
        End If
        
        If Not bSafe Then
          If InStr(1, sSafeFilters, sName, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
            'add to results list
            sHit = "O18 - Filter: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                With Result
                    .Section = "O18"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "Protocols\Filter\" & sName
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
          Else
            If sCLSID <> "(no CLSID)" Then
                If InStr(1, sSafeFilters, sCLSID, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                    'add to results list
                    sHit = "O18 - Filter hijack: " & sName & " - " & sCLSID & " - " & sFile
                    If Not IsOnIgnoreList(sHit) Then
                        If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    
                        With Result
                            .Section = "O18"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, HKEY_CLASSES_ROOT, "Protocols\Filter\" & sName, "CLSID", O18_GetCLSIDByFilter(sName)
                            If Not IsMicrosoftFile(sFile) Then
                                AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
                                If Not FileMissing(sFile) Then
                                    AddFileToFix .File, REMOVE_FILE, sFile
                                End If
                            End If
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End If
          End If
        End If
        
        sName = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sName)
        i = i + 1
      Loop
      RegCloseKey hKey
    End If
    
    AppendErrorLogCustom "CheckO18Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO18Item"
    If hKey <> 0 Then RegCloseKey hKey
End Sub

Public Function FileMissing(sFile$) As Boolean
    If Len(sFile) = 0 Then FileMissing = True: Exit Function
    If sFile = "(no file)" Then FileMissing = True: Exit Function
    If StrEndWith(sFile, "(file missing)") Then FileMissing = True: Exit Function
End Function

Private Function O18_GetCLSIDByProtocol(sProtocol$) As String
    Dim i&, sCLSID$
    For i = 0 To UBound(aSafeProtocols)
        'find CLSID for protocol name
        If InStr(1, aSafeProtocols(i), sProtocol) > 0 Then
            sCLSID = SplitSafe(aSafeProtocols(i), "|")(1)
            Exit For
        End If
    Next i
    O18_GetCLSIDByProtocol = sCLSID
End Function

Private Function O18_GetCLSIDByFilter(sFilter$) As String
    Dim i&, sCLSID$
    For i = 0 To UBound(aSafeFilters)
        'find CLSID for protocol name
        If InStr(1, aSafeFilters(i), sFilter) > 0 Then
            sCLSID = SplitSafe(aSafeFilters(i), "|")(1)
            Exit For
        End If
    Next i
    O18_GetCLSIDByFilter = sCLSID
End Function

Public Sub FixO18Item(sItem$, Result As SCAN_RESULT)
    'O18 - Protocol: cn
    'O18 - Filter: text/blah - {0} - c:\file.dll
    On Error GoTo ErrorHandler:
    FixRegistryHandler Result
    FixFileHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO18Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO19Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO19Item - Begin"
    
    'Software\Microsoft\Internet Explorer\Styles,Use My Stylesheet
    'Software\Microsoft\Internet Explorer\Styles,User Stylesheet
    
    Dim lUseMySS&, sUserSS$, sHit$, Result As SCAN_RESULT
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Internet Explorer\Styles"
    
    Do While HE.MoveNext
        lUseMySS = Reg.GetDword(HE.Hive, HE.Key, "Use My Stylesheet", HE.Redirected)
        sUserSS = Reg.GetString(HE.Hive, HE.Key, "User Stylesheet", HE.Redirected)
        
        sUserSS = FormatFileMissing(sUserSS)
        
        If lUseMySS = 1 And sUserSS <> vbNullString Then
            'O19 - User stylesheet (HKCU,HKLM):
            'O19-32 - User stylesheet (HKCU,HKLM):
            sHit = IIf(bIsWin32, "O19", IIf(HE.Redirected, "O19-32", "O19")) & " - " & HE.HiveNameAndSID & "\..\User stylesheet: " & sUserSS
            If Not IsOnIgnoreList(sHit) Then
                'md5 doesn't seem useful here
                'If bMD5 Then sHit = sHit & getfilemd5(sUserSS)
                With Result
                    .Section = "O19"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "Use My Stylesheet", , HE.Redirected
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "User Stylesheet", , HE.Redirected
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
        End If
    Loop
    
    AppendErrorLogCustom "CheckO19Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO19Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO19Item(sItem$, Result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    'O19 - User stylesheet: c:\file.css (file missing)
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO19Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO20Item()
    'AppInit_DLLs - https://support.microsoft.com/ru-ru/kb/197571
    
    'According to MSDN:
    ' - modules are delimited by spaces or commas
    ' - long file names are not permitted
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO20Item - Begin"
    
    'appinit_dlls + winlogon notify
    Dim sAppInit$, sFile$, sHit$, UseWow, Wow6432Redir As Boolean, Result As SCAN_RESULT
    Dim bEnabled As Boolean, bCodeSigned As Boolean
    
    For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If bIsWin32 And Wow6432Redir Then Exit For
    
        sAppInit = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
        
        If OSver.MajorMinor <= 5.2 Then 'XP/2003-
            bEnabled = True
        Else
            bEnabled = (1 = Reg.GetDword(HKEY_LOCAL_MACHINE, sAppInit, "LoadAppInit_DLLs", Wow6432Redir))
            bCodeSigned = (1 = Reg.GetDword(HKEY_LOCAL_MACHINE, sAppInit, "RequireSignedAppInit_DLLs", Wow6432Redir))
        End If
        
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sAppInit, "AppInit_DLLs", Wow6432Redir)
        If sFile <> vbNullString Then
            sFile = Replace$(sFile, vbNullChar, "|")                        '// TODO: !!!
            If InStr(1, sSafeAppInit, sFile, vbTextCompare) = 0 Or bIgnoreAllWhitelists Then
                'item is not on whitelist
                'O20 - AppInit_DLLs
                'O20-32 - AppInit_DLLs
                sHit = IIf(bIsWin32, "O20", IIf(Wow6432Redir, "O20-32", "O20")) & " - AppInit_DLLs: " & sFile & _
                  IIf(bCodeSigned, " (required code signed dll)", "") & _
                  IIf(Not bEnabled, " (disabled by registry)", "") & IIf(OSver.SecureBoot, " (disabled by SecureBoot)", "")
            
                If Not IsOnIgnoreList(sHit) Or bIgnoreAllWhitelists Then
                    With Result
                        .Section = "O20"
                        .HitLineW = sHit
                        AddRegToFix .Reg, RESTORE_VALUE, 0, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", "", CLng(Wow6432Redir), REG_RESTORE_SZ
                        'AddRegToFix .Reg, RESTORE_VALUE, 0, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Windows", "LoadAppInit_DLLs", 0, CLng(Wow6432Redir), REG_RESTORE_DWORD
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
        End If
        
        Dim sSubkeys$(), i&, sWinLogon$
        sWinLogon = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify"
        sSubkeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sWinLogon, Wow6432Redir), "|")
        If UBound(sSubkeys) <> -1 Then
            For i = 0 To UBound(sSubkeys)
                If InStr(1, "*" & sSafeWinlogonNotify & "*", "*" & sSubkeys(i) & "*", vbTextCompare) = 0 Then
                    sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sWinLogon & "\" & sSubkeys(i), "DllName", Wow6432Redir)
                    
                    sFile = FormatFileMissing(sFile)
                    
                    'O20 - Winlogon Notify:
                    'O20-32 - Winlogon Notify:
                    sHit = IIf(bIsWin32, "O20", IIf(Wow6432Redir, "O20-32", "O20")) & " - Winlogon Notify: " & sSubkeys(i) & " - " & sFile
                    If Not IsOnIgnoreList(sHit) Then
                        If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                        With Result
                            .Section = "O20"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify\" & sSubkeys(i), , , CLng(Wow6432Redir)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            Next i
        End If
    Next

    AppendErrorLogCustom "CheckO20Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO20Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO20Item(sItem$, Result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    
    'O20 - AppInit_DLLs: file.dll
    'O20 - Winlogon Notify: bladibla - c:\file.dll
    '
    '* clear appinit regval (don't delete it)
    '* kill regkey (for winlogon notify)
    
    FixRegistryHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO20Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO21Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO21Item - Begin"
    
    'Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad
    'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers
    'Software\Microsoft\Windows\CurrentVersion\explorer\ShellExecuteHooks
    
    '//TODO
    'Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved
    'HKCR\Folder\shellex\ColumnHandlers
    'HKCR\AllFilesystemObjects\shellex\ContextMenuHandlers
    
    Dim sSSODL$, sHit$, sFile$, bOnWhiteList As Boolean
    Dim hKey&, i&, sName$, lNameLen&, sCLSID$, lDataLen&, sValueName$
    Dim Result As SCAN_RESULT, bSafe As Boolean, bInList As Boolean
    
    sSSODL = "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
    
    'BE AWARE: SHELL32.dll - sometimes this file is patched
    '(e.g. seen after "Windown XP Update pack by Simplix" together with his certificate installed to trusted root storage)
    
    HE.Init HE_HIVE_HKLM
    HE.AddKey sSSODL
    
    Do While HE.MoveNext
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
        
            Do
                lNameLen = MAX_VALUENAME
                sValueName = String$(lNameLen, 0&)
                lDataLen = MAX_VALUENAME
                sCLSID = String$(lDataLen, 0&)
                
                If RegEnumValueW(hKey, i, StrPtr(sValueName), lNameLen, 0&, REG_SZ, StrPtr(sCLSID), lDataLen) <> 0 Then Exit Do
                
                sValueName = Left$(sValueName, lNameLen)
                sCLSID = TrimNull(sCLSID)
                
                Call GetFileByCLSID(sCLSID, sFile, sName, HE.Redirected, HE.SharedKey)
                
                sFile = FormatFileMissing(sFile)
                
                bSafe = False
                If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                    
                    bInList = inArray(sCLSID, aSafeSSODL, , , vbTextCompare)
                    If StrComp(GetFileName(sFile, True), "GROOVEEX.DLL", 1) = 0 Then bInList = True
                    
                    If bInList Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                End If
                
                If sName = "(no name)" Then sName = sValueName
                
                sHit = IIf(bIsWin32, "O21", IIf(HE.Redirected, "O21-32", "O21")) & " - ShellServiceObjectDelayLoad: " & sName & " - " & sCLSID & " - " & sFile
                
                'some shit leftover by Microsoft ^)
                If sName = "WebCheck" And sCLSID = "{E6FB5E20-DE35-11CF-9C87-00AA005127ED}" And sFile = "(no file)" Then bSafe = True
                
                If Not IsOnIgnoreList(sHit) And Not bSafe Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    With Result
                        .Section = "O21"
                        .HitLineW = sHit
                        If sCLSID <> "" Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sName, , HE.Redirected
                        If Not FileMissing(sFile) Then AddFileToFix .File, REMOVE_FILE, sFile
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End With
                    AddToScanResults Result
                End If
                
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    Dim aSubKey() As String
    Dim sSIOI As String
    
    sSIOI = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers"
    
    HE.Init HE_HIVE_HKLM
    HE.AddKey sSIOI
    
    Do While HE.MoveNext
        
        Erase aSubKey
        If Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey, HE.Redirected) > 0 Then
        
            For i = 1 To UBound(aSubKey)
            
                sName = aSubKey(i)
                sCLSID = Reg.GetString(HE.Hive, HE.Key & "\" & aSubKey(i), vbNullString, HE.Redirected)
                
                Call GetFileByCLSID(sCLSID, sFile, sName, HE.Redirected, HE.SharedKey)
                
                sFile = FormatFileMissing(sFile)
                
                bSafe = False
                If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                    
                    bInList = inArray(sFile, aSafeSIOI, , , vbTextCompare)
                    
                    If StrComp(GetFileName(sFile, True), "GROOVEEX.DLL", 1) = 0 Then bInList = True
                    If StrComp(GetFileName(sFile, True), "FileSyncShell.dll", 1) = 0 Then bInList = True
                    If StrComp(GetFileName(sFile, True), "FileSyncShell64.dll", 1) = 0 Then bInList = True
                    
                    If bInList Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                End If
                
                If sName = "(no name)" Then sName = aSubKey(i)
                
                sHit = IIf(bIsWin32, "O21", IIf(HE.Redirected, "O21-32", "O21")) & " - ShellIconOverlayIdentifiers: " & sName & " - " & sCLSID & " - " & sFile
                
                If Not IsOnIgnoreList(sHit) And Not bSafe Then
                    If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                    With Result
                        .Section = "O21"
                        .HitLineW = sHit
                        If sCLSID <> "" Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                        If Not FileMissing(sFile) Then AddFileToFix .File, REMOVE_FILE, sFile
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End With
                    AddToScanResults Result
                End If
            
            Next
        End If
    Loop
    
    'ShellExecuteHooks
    'See: http://blog.zemana.com/2016/06/youndoocom-using-shellexecutehooks-to.html
    
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => EnableShellExecuteHooks
    
    Dim bDisabled As Boolean
    Dim aValue() As String
    
    HE.Init HE_HIVE_HKLM
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks"
    
    Do While HE.MoveNext
        Erase aValue
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue, HE.Redirected)
            sCLSID = aValue(i)
            
            Call GetFileByCLSID(sCLSID, sFile, sName, HE.Redirected, HE.SharedKey)
            
            sFile = FormatFileMissing(sFile)

            bSafe = False
            If bHideMicrosoft And Not bIgnoreAllWhitelists Then
            
                bInList = inArray(sFile, aSafeSEH, , , vbTextCompare)
                If StrComp(GetFileName(sFile, True), "GROOVEEX.DLL", 1) = 0 Then bInList = True
                
                If bInList Then
                    If IsMicrosoftFile(sFile) Then bSafe = True
                End If
            End If
            
            If OSver.MajorMinor >= 6 Then 'XP/2003 has no policy
                bDisabled = Not (1 = Reg.GetDword(HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "EnableShellExecuteHooks"))
            End If
            
            sHit = IIf(HE.Redirected, "O21-32", "O21") & " - ShellExecuteHooks: " & sName & " - " & sCLSID & " - " & sFile & IIf(bDisabled, " (disabled)", "")
            If Not IsOnIgnoreList(sHit) And Not bSafe Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                With Result
                    .Section = "O21"
                    .HitLineW = sHit
                    If sCLSID <> "" Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                    If Not FileMissing(sFile) Then AddFileToFix .File, REMOVE_FILE, sFile
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults Result
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckO21Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO21Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO21Item(sItem$, Result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    
    'O21 - SSODL: webcheck - {000....000} - c:\file.dll (file missing)
    'actions to take:
    '* kill file
    '* kill regkey - ShellIconOverlayIdentifiers
    '* kill regparam - SSODL
    '* kill clsid regkey
    
    FixRegistryHandler Result
    FixFileHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO21Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO22Item()
    'ScheduledTask
    'XP    - HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler
    'Vista - HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO22Item - Begin"
    
    If OSver.IsWindowsVistaOrGreater Then
        EnumTasks2   '<--- New routine
        EnumJobs
        Exit Sub
    End If
    
    '//TODO: Add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\
    'for Windows Vista and higher.
    
    'Win XP / Server 2003
    
    Dim sSTS$, hKey&, i&, sCLSID$, lCLSIDLen&, lDataLen&
    Dim sFile$, sName$, sHit$, isSafe As Boolean
    Dim Wow6432Redir As Boolean, Result As SCAN_RESULT
    
    EnumJobs
    
    Wow6432Redir = False
    
    sSTS = "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler"
    If RegOpenKeyExW(HKEY_LOCAL_MACHINE, StrPtr(sSTS), 0, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not Wow6432Redir), hKey) <> 0 Then
        'regkey doesn't exist, or failed to open
        Exit Sub
    End If
    
    Do
        lCLSIDLen = MAX_VALUENAME
        sCLSID = String$(lCLSIDLen, 0&)
        lDataLen = MAX_VALUENAME
        sName = String$(lDataLen, 0&)
    
        If RegEnumValueW(hKey, i, StrPtr(sCLSID), lCLSIDLen, 0&, REG_SZ, StrPtr(sName), lDataLen) <> 0 Then Exit Do
    
        sCLSID = Left$(sCLSID, lCLSIDLen)
        sName = TrimNull(sName)
        If sName = vbNullString Then sName = "(no name)"
        sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString, Wow6432Redir)
        sFile = UnQuote(EnvironW(sFile))
        If sFile = vbNullString Then
            sFile = "(no file)"
        Else
            If Not FileExists(sFile) Then
                sFile = sFile & " (file missing)"
            End If
        End If
        
        'whitelist
        isSafe = isInTasksWhiteList(sCLSID & "\" & sName, sFile, "")
        
        If Not isSafe Then
            sHit = "O22 - ScheduledTask: " & sName & " - " & sCLSID & " - " & sFile
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                AddToScanResultsSimple "O22", sHit
            End If
            
            If Not IsOnIgnoreList(sHit) Then
                If bMD5 Then sHit = sHit & GetFileMD5(sFile)
                With Result
                    .Section = "O22"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler", sCLSID
                    AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults Result
            End If
            
        End If
        i = i + 1
    Loop
    RegCloseKey hKey
    
    AppendErrorLogCustom "CheckO22Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO22Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO22Item(sItem$, Result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    'O22 - ScheduledTask: blah - {000...000} - file.dll
    FixIt Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO22Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO23Item()
    'https://www.bleepingcomputer.com/tutorials/how-malware-hides-as-a-service/
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO23Item - Begin"
    
    'enum NT services
    Dim sServices$(), i&, j&, sName$, sDisplayName$, tmp$, Result As SCAN_RESULT
    Dim lStart&, lType&, sFile$, sHit$, sBuf$, IsCompositeCmd As Boolean
    Dim bHideDisabled As Boolean, Stady As Long, sServiceDll As String, sServiceDll_2 As String, bDllMissing As Boolean
    Dim ServState As SERVICE_STATE
    Dim argc As Long
    Dim argv() As String
    Dim isSafeMSCmdLine As Boolean
    Dim SignResult As SignResult_TYPE
    Dim FoundFile As String
    Dim IsMSCert As Boolean
    
    If Not bIsWinNT Then Exit Sub
    
    If Not bIgnoreAllWhitelists Then
        bHideDisabled = True
    End If
    
    sServices = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    Stady = 1: Dbg CStr(Stady)
    If UBound(sServices) = -1 Then Exit Sub
    
    Stady = 2: Dbg CStr(Stady)
    
    For i = 0 To UBound(sServices)
        
        sName = sServices(i)
        Dbg sName
        
        Stady = 3: Dbg CStr(Stady)
        lType = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Type")
        
        If lType < 16 Then 'Driver
            GoTo Continue
        End If
        
        Stady = 4: Dbg CStr(Stady)
        lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start")
        
        If (lStart = 4 And bHideDisabled) Then
            GoTo Continue
        End If
        
        UpdateProgressBar "O23", sName
        
        Stady = 5: Dbg CStr(Stady)
        sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DisplayName")
        Stady = 6: Dbg CStr(Stady)
        If Len(sDisplayName) = 0 Then
            sDisplayName = sName
        Else
            If Left$(sDisplayName, 1) = "@" Then                    'extract string resource from file
                Stady = 7: Dbg CStr(Stady)
                sBuf = GetStringFromBinary(, , sDisplayName)
                Stady = 8: Dbg CStr(Stady)
                If 0 <> Len(sBuf) Then sDisplayName = sBuf
            End If
        End If
        
        Stady = 9: Dbg CStr(Stady)
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ImagePath")
        sServiceDll = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName & "\Parameters", "ServiceDll")
        Stady = 10: Dbg CStr(Stady)
        
        bDllMissing = False
        
        'Checking Service Dll
        If Len(sServiceDll) <> 0 Then
            sServiceDll = EnvironW(UnQuote(sServiceDll))
            
            tmp = FindOnPath(sServiceDll)
            
            If Len(tmp) = 0 Then
                sServiceDll = sServiceDll & " (file missing)"
                bDllMissing = True
            Else
                sServiceDll = tmp
            End If
        End If
        
        If bDllMissing Then
            
            sServiceDll_2 = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ServiceDll")
            
            If Len(sServiceDll_2) <> 0 Then
                
                sServiceDll_2 = EnvironW(UnQuote(sServiceDll_2))
                
                tmp = FindOnPath(sServiceDll_2)
                
                If Len(tmp) <> 0 Then sServiceDll = tmp: bDllMissing = False
            End If
        End If
        
        'cleanup filename
        sFile = CleanServiceFileName(sFile, sName)
        
        '//// TODO: Check this !!!
        
        'https://technet.microsoft.com/en-us/library/cc959922.aspx
        'https://support.microsoft.com/en-us/kb/103000
        
        'Start
        '0 - Boot
        '1 - System
        '2 - Automatic
        '3 - Manual
        '4 - Disabled
        
        'Type
        '1 - Kernel device driver
        '2 - File System driver
        '4 - A set of arguments for an adapter
        '8 - File System driver service
        '16 - A Win32 program that runs in a process by itself. This type of Win32 service can be started by the service controller.
        '32 - A Win32 service that can share a process with other Win32 services
        '272 - A Win32 program that runs in a process by itself (like Type16) and that can interact with users.
        '288 - A Win32 program that shares a process and that can interact with users.
        
        ServState = GetServiceRunState(sName)
        
        If lType >= 16 Then
          If Not (lStart = 4 And bHideDisabled) Then
               
            Stady = 22: Dbg CStr(Stady)
            
            IsCompositeCmd = False
            isSafeMSCmdLine = False
            
            If Not FileExists(sFile) And sFile <> "" Then
            
                ' Дальше идут процедуры парсинга командной строки и проверки сертиката для каждого файла из этой цепочки
                ' Если любой файл из цепочки не проходит проверку, строка считается небезопасной
            
                Stady = 23: Dbg CStr(Stady)
            
                ParseCommandLine sFile, argc, argv
                
                '// TODO: добавить к FindOnPath папку, в которой находится основной запускаемый службой файл
                
                'если файл в составе коммандной строки, например: C:\WINDOWS\system32\svchost -k rpcss.exe
                
                If argc > 2 Then        ' 1 -> app exe self, 2 -> actual cmd, 3 -> arg
                
                  Stady = 24: Dbg CStr(Stady)
                
                  If Not FileExists(argv(1)) Then   ' если запускающий файл не существует -> ищем его
                    Stady = 25: Dbg CStr(Stady)
                    FoundFile = FindOnPath(argv(1))
                    argv(1) = FoundFile
                  Else
                    FoundFile = argv(1)
                  End If
                
                  Stady = 26: Dbg CStr(Stady)
                
                  ' если запускающий файл существует (иначе, нет смысла проверять остальные аргументы)
                  If 0 <> Len(FoundFile) Then
                    
                    'флаг о том, что служба запускает составную командную строку, в которой как минимум первый (запускающий файл) существует
                    IsCompositeCmd = True
                
                    isSafeMSCmdLine = True
                
                    Stady = 27: Dbg CStr(Stady)
                 
                    For j = 1 To UBound(argv) ' argv[1] -> запускающий файл в цепочке
                    
                        ' проверяем хеш корневого сертификата каждого из элементов командной строки, если он был найден по известным путям Path
                        
                        FoundFile = FindOnPath(argv(j))
                        
                        Stady = 28: Dbg CStr(Stady)
                        
                        If 0 <> Len(FoundFile) Then
                        
                            Stady = 29: Dbg CStr(Stady)
                        
                            If IsWinServiceFileName(FoundFile, sDisplayName) Then
                                SignVerify FoundFile, SV_LightCheck Or SV_PreferInternalSign, SignResult
                                IsMSCert = SignResult.isMicrosoftSign And SignResult.isLegit
                            Else
                                IsMSCert = False
                            End If
                            
                            If Not IsMSCert Then isSafeMSCmdLine = False: Exit For
                        End If
                    Next
                  End If
                End If
            
            End If
            
            Stady = 32: Dbg CStr(Stady)
            
            If 0 = Len(sFile) Then
                sFile = "(no file)"
            Else
                If (Not FileExists(sFile)) And (Not IsCompositeCmd) Then
                    sFile = sFile & " (file missing)"
                Else
                    If IsCompositeCmd Then
                        FoundFile = argv(1)
                    Else
                        FoundFile = sFile
                    End If
                    Stady = 33: Dbg CStr(Stady)
                    
                    'sCompany = GetFilePropCompany(FoundFile)
                    'If Len(sCompany) = 0 Then sCompany = "Unknown owner"
                    
                End If
            End If
            
            If Not IsCompositeCmd And sFile <> "(no file)" Then    'иначе, такая проверка уже выполнена ранее
                If IsWinServiceFileName(sFile, sDisplayName) Then
                    SignVerify sFile, SV_LightCheck Or SV_PreferInternalSign, SignResult
                Else
                    WipeSignResult SignResult
                End If
            End If
            
            'override by checkind EDS of service dll if original file is Microsoft (usually, svchost)
            If Len(sServiceDll) <> 0 And (Not bDllMissing) Then
                If IsWinServiceFileName(sServiceDll, sDisplayName) Then
                    SignVerify sServiceDll, SV_LightCheck Or SV_PreferInternalSign, SignResult
                Else
                    WipeSignResult SignResult
                End If
            End If
            
            With SignResult
                ' если корневой сертификат цепочки доверия принадлежит Майкрософт, то исключаем службу из лога
                
                If bDllMissing Or Not (.isMicrosoftSign And .isLegit And bHideMicrosoft) Then
                    Stady = 36: Dbg CStr(Stady)
                    If bMD5 Then
                        If sFile <> "(no file)" Then sFile = sFile & GetFileMD5(sFile)
                    End If
                    Stady = 37: Dbg CStr(Stady)
                    'sHit = "O23 - Service " & IIf(ServState <> SERVICE_STOPPED, "R", "S") & lStart & _
                    '    ": " & sDisplayName & " - (" & sName & ")" & " - " & sCompany & " - " & sFile
                    
                    'sHit = "O23 - Service " & IIf(ServState <> SERVICE_STOPPED, "R", "S") & lStart & _
                    '    ": " & sDisplayName & " - HKLM\..\" & sName & " - " & sFile
                    
                    sHit = "O23 - Service " & IIf(ServState <> SERVICE_STOPPED, "R", "S") & lStart & _
                        ": " & IIf(sDisplayName = sName, sName, sDisplayName & " - (" & sName & ")") & " - " & sFile
                    
                    If Len(sServiceDll) <> 0 Then
                        sHit = sHit & "; ""ServiceDll"" = " & sServiceDll
                    End If
                    
' I temporarily remove EDS name in log
'                    If .isLegit And 0 <> Len(.SubjectName) And Not bDllMissing Then
'                        sHit = sHit & " (" & .SubjectName & ")"
'                    Else
'                        sHit = sHit & " (not signed)"
'                    End If
                    
                    If Not IsOnIgnoreList(sHit) Then
                        Stady = 38: Dbg CStr(Stady)
                        With Result
                            .Section = "O23"
                            .HitLineW = sHit
                            AddServiceToFix .Service, DELETE_SERVICE, sName
                            .CureType = SERVICE_BASED
                        End With
                        AddToScanResults Result
                    End If
                End If
            End With
          End If
        End If
Continue:
    Next i

    '//TODO: Add checking device drivers
    
    '
    '
    'https://docs.microsoft.com/en-us/windows-hardware/drivers/install/using-setupapi-to-uninstall-devices-and-driver-packages
    'https://stackoverflow.com/questions/12756712/windows-device-uninstall-using-c
    '
    'Enum via SetupAPI or via NtQuerySystemInformation
    
'via NtQuerySystemInformation:
    
'    Const DRIVER_INFORMATION            As Long = 11
'    Const SYSTEM_MODULE_SIZE            As Long = 284
'    Const STATUS_INFO_LENGTH_MISMATCH   As Long = &HC0000004
    
'    Dim ret     As Long
'    Dim buf()   As Byte
'    Dim mdl     As SYSTEM_MODULE_INFORMATION
'    Dim Driver  As String
'
'    If NtQuerySystemInformation(DRIVER_INFORMATION, ByVal 0&, 0, ret) = STATUS_INFO_LENGTH_MISMATCH Then
'        ReDim buf(ret - 1)
'        If NtQuerySystemInformation(DRIVER_INFORMATION, buf(0), ret, ret) = STATUS_SUCCESS Then
'            mdl.ModulesCount = buf(0) Or (buf(1) * &H100&) Or (buf(2) * &H10000) Or (buf(3) * &H1000000)
'            If mdl.ModulesCount Then
'                ReDim mdl.Modules(mdl.ModulesCount - 1)
'                For ret = 0 To mdl.ModulesCount - 1
'                    memcpy mdl.Modules(ret), buf(ret * SYSTEM_MODULE_SIZE + 4), SYSTEM_MODULE_SIZE
'                    Driver = NormalizeDriverString(mdl.Modules(ret).Name)
'                    AddtoLog Driver  '"[" & Format(ret + 1, "000") & "] "
'                Next
'            End If
'        End If
'    End If

    AppendErrorLogCustom "CheckO23Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO23Item", "Service=", sDisplayName, "Stady=", Stady
    If inIDE Then Stop: Resume Next
End Sub

'Function NormalizeDriverString(ByVal Driver As String) As String
'    Dim WinDir         As String
'    Dim WinDirName     As String
'    Dim pos            As Long
'
'    WinDrivers = Environ("SystemRoot") & "\System32\Drivers"
'    WinDir = Environ("SystemRoot")
'    WinDirName = "\" & Mid(WinDir, 4) & "\"
'    ' чистка
'    pos = InStr(Driver, Chr(0))
'    If pos Then Driver = Left(Driver, pos - 1)
'    ' если нет абсолютного пути
'    If InStr(Driver, "\") = 0 Then
'        Driver = FindFileOnPathFolders(Driver)
'    End If
'    ' нормализация путей
'    If Left(Driver, 4) = "\??\" Then Driver = Mid(Driver, 5)
'    ' приведение пути к системной папке к единому виду "\SystemRoot\" (отключил)
'    If StrComp(Left(Driver, Len(WinDirName)), WinDirName, 1) = 0 Then
'        Driver = "\SystemRoot\" & Mid(Driver, Len(WinDirName) + 1)
'        'Driver = WinDir & "\" & Mid(Driver, Len(WinDirName) + 1)
'    End If
'    If StrComp(Left(Driver, Len(WinDir)), WinDir, 1) = 0 Then Driver = "\SystemRoot\" & Mid(Driver, Len(WinDir) + 2)
'    NormalizeDriverString = Driver
'End Function

Private Function IsWinServiceFileName(sFilePath As String, sServiceName As String) As Boolean
    
    On Error GoTo ErrorHandler:
    
    Static IsInit As Boolean
    Static oDictSRV As clsTrickHashTable
    Dim sCompany As String
    
    If Not IsInit Then
        Dim vKey, prefix$
        IsInit = True
        Set oDictSRV = New clsTrickHashTable
        
        With oDictSRV
            .CompareMode = TextCompare
            .Add "<PF32>\Common Files\Microsoft Shared\Phone Tools\CoreCon\11.0\bin\IpOverUsbSvc.exe", 0&
            .Add "<PF32>\Common Files\Microsoft Shared\Source Engine\OSE.exe", 0&
            .Add "<PF32>\Common Files\Microsoft Shared\VS7DEBUG\MDM.exe", 0&
            .Add "<PF32>\Windows Kits\8.1\App Certification Kit\fussvc.exe", 0&
            .Add "<PF32>\Skype\Updater\Updater.exe", 0&
            .Add "<PF64>\Common Files\Microsoft Shared\OfficeSoftwareProtectionPlatform\OSPPSVC.exe", 0&
            .Add "<PF64>\Microsoft SQL Server\90\Shared\sqlwriter.exe", 0&
            .Add "<PF64>\Windows Media Player\wmpnetwk.exe", 0&
            .Add "<SysRoot>\ehome\ehRecvr.exe", 0&
            .Add "<SysRoot>\ehome\ehsched.exe", 0&
            .Add "<SysRoot>\ehome\ehstart.dll", 0&
            .Add "<SysRoot>\Microsoft.NET\Framework64\v2.0.50727\mscorsvw.exe", 0&
            .Add "<SysRoot>\Microsoft.NET\Framework64\v3.0\Windows Communication Foundation\infocard.exe", 0&
            .Add "<SysRoot>\Microsoft.Net\Framework64\v3.0\WPF\PresentationFontCache.exe", 0&
            .Add "<SysRoot>\Microsoft.NET\Framework64\v4.0.30319\mscorsvw.exe", 0&
            .Add "<SysRoot>\Microsoft.NET\Framework\v2.0.50727\mscorsvw.exe", 0&
            .Add "<SysRoot>\Microsoft.NET\Framework\v4.0.30319\mscorsvw.exe", 0&
            .Add "<SysRoot>\PCHealth\HelpCtr\Binaries\pchsvc.dll", 0&
            .Add "<SysRoot>\servicing\TrustedInstaller.exe", 0&
            .Add "<SysRoot>\System32\advapi32.dll", 0&
            .Add "<SysRoot>\System32\aelupsvc.dll", 0&
            .Add "<SysRoot>\System32\AJRouter.dll", 0&
            .Add "<SysRoot>\System32\alg.exe", 0&
            .Add "<SysRoot>\System32\APHostService.dll", 0&
            .Add "<SysRoot>\System32\appidsvc.dll", 0&
            .Add "<SysRoot>\System32\appinfo.dll", 0&
            .Add "<SysRoot>\System32\appmgmts.dll", 0&
            .Add "<SysRoot>\System32\AppReadiness.dll", 0&
            .Add "<SysRoot>\System32\appxdeploymentserver.dll", 0&
            .Add "<SysRoot>\System32\AudioEndpointBuilder.dll", 0&
            .Add "<SysRoot>\System32\Audiosrv.dll", 0&
            .Add "<SysRoot>\System32\AxInstSV.dll", 0&
            .Add "<SysRoot>\System32\bdesvc.dll", 0&
            .Add "<SysRoot>\System32\bfe.dll", 0&
            .Add "<SysRoot>\System32\bisrv.dll", 0&
            .Add "<SysRoot>\System32\browser.dll", 0&
            .Add "<SysRoot>\System32\BthHFSrv.dll", 0&
            .Add "<SysRoot>\System32\bthserv.dll", 0&
            .Add "<SysRoot>\System32\CDPSvc.dll", 0&
            .Add "<SysRoot>\System32\CDPUserSvc.dll", 0&
            .Add "<SysRoot>\System32\certprop.dll", 0&
            .Add "<SysRoot>\System32\cisvc.exe", 0&
            .Add "<SysRoot>\System32\ClipSVC.dll", 0&
            .Add "<SysRoot>\System32\coremessaging.dll", 0&
            .Add "<SysRoot>\System32\cryptsvc.dll", 0&
            .Add "<SysRoot>\System32\cscsvc.dll", 0&
            .Add "<SysRoot>\System32\das.dll", 0&
            .Add "<SysRoot>\System32\dcpsvc.dll", 0&
            .Add "<SysRoot>\System32\defragsvc.dll", 0&
            .Add "<SysRoot>\System32\DeviceSetupManager.dll", 0&
            .Add "<SysRoot>\System32\DevQueryBroker.dll", 0&
            .Add "<SysRoot>\System32\DFSR.exe", 0&
            .Add "<SysRoot>\System32\dhcpcore.dll", 0&
            .Add "<SysRoot>\System32\dhcpcsvc.dll", 0&
            .Add "<SysRoot>\System32\DiagSvcs\DiagnosticsHub.StandardCollector.Service.exe", 0&
            .Add "<SysRoot>\System32\diagtrack.dll", 0&
            .Add "<SysRoot>\System32\dllhost.exe", 0&
            .Add "<SysRoot>\System32\dmadmin.exe", 0&
            .Add "<SysRoot>\System32\dmserver.dll", 0&
            .Add "<SysRoot>\System32\dmwappushsvc.dll", 0&
            .Add "<SysRoot>\System32\dnsrslvr.dll", 0&
            .Add "<SysRoot>\System32\dot3svc.dll", 0&
            .Add "<SysRoot>\System32\dps.dll", 0&
            .Add "<SysRoot>\System32\DsSvc.dll", 0&
            .Add "<SysRoot>\System32\eapsvc.dll", 0&
            .Add "<SysRoot>\System32\efssvc.dll", 0&
            .Add "<SysRoot>\System32\embeddedmodesvc.dll", 0&
            .Add "<SysRoot>\System32\emdmgmt.dll", 0&
            .Add "<SysRoot>\System32\EnterpriseAppMgmtSvc.dll", 0&
            .Add "<SysRoot>\System32\ersvc.dll", 0&
            .Add "<SysRoot>\System32\es.dll", 0&
            .Add "<SysRoot>\System32\fdPHost.dll", 0&
            .Add "<SysRoot>\System32\fdrespub.dll", 0&
            .Add "<SysRoot>\System32\fhsvc.dll", 0&
            .Add "<SysRoot>\System32\flightsettings.dll", 0&
            .Add "<SysRoot>\System32\FntCache.dll", 0&
            .Add "<SysRoot>\System32\FrameServer.dll", 0&
            .Add "<SysRoot>\System32\fxssvc.exe", 0&
            .Add "<SysRoot>\System32\GeofenceMonitorService.dll", 0&
            .Add "<SysRoot>\System32\gpsvc.dll", 0&
            .Add "<SysRoot>\System32\hidserv.dll", 0&
            .Add "<SysRoot>\System32\hvhostsvc.dll", 0&
            .Add "<SysRoot>\System32\icsvc.dll", 0&
            .Add "<SysRoot>\System32\icsvcext.dll", 0&
            .Add "<SysRoot>\System32\IEEtwCollector.exe", 0&
            .Add "<SysRoot>\System32\ikeext.dll", 0&
            .Add "<SysRoot>\System32\imapi.exe", 0&
            .Add "<SysRoot>\System32\ipbusenum.dll", 0&
            .Add "<SysRoot>\System32\iphlpsvc.dll", 0&
            .Add "<SysRoot>\System32\ipnathlp.dll", 0&
            .Add "<SysRoot>\System32\ipsecsvc.dll", 0&
            .Add "<SysRoot>\System32\irmon.dll", 0&
            .Add "<SysRoot>\System32\iscsiexe.dll", 0&
            .Add "<SysRoot>\System32\keyiso.dll", 0&
            .Add "<SysRoot>\System32\kmsvc.dll", 0&
            .Add "<SysRoot>\System32\lfsvc.dll", 0&
            .Add "<SysRoot>\System32\LicenseManagerSvc.dll", 0&
            .Add "<SysRoot>\System32\ListSvc.dll", 0&
            .Add "<SysRoot>\System32\lltdsvc.dll", 0&
            .Add "<SysRoot>\System32\lmhsvc.dll", 0&
            .Add "<SysRoot>\System32\locator.exe", 0&
            .Add "<SysRoot>\System32\lsass.exe", 0&
            .Add "<SysRoot>\System32\lsm.dll", 0&
            .Add "<SysRoot>\System32\MessagingService.dll", 0&
            .Add "<SysRoot>\System32\mmcss.dll", 0&
            .Add "<SysRoot>\System32\mnmsrvc.exe", 0&
            .Add "<SysRoot>\System32\moshost.dll", 0&
            .Add "<SysRoot>\System32\mpssvc.dll", 0&
            .Add "<SysRoot>\System32\msdtc.exe", 0&
            .Add "<SysRoot>\System32\msdtckrm.dll", 0&
            .Add "<SysRoot>\System32\msiexec.exe", 0&
            .Add "<SysRoot>\System32\mspmsnsv.dll", 0&
            .Add "<SysRoot>\System32\mswsock.dll", 0&
            .Add "<SysRoot>\System32\ncasvc.dll", 0&
            .Add "<SysRoot>\System32\ncbservice.dll", 0&
            .Add "<SysRoot>\System32\NcdAutoSetup.dll", 0&
            .Add "<SysRoot>\System32\netlogon.dll", 0&
            .Add "<SysRoot>\System32\netman.dll", 0&
            .Add "<SysRoot>\System32\netprofm.dll", 0&
            .Add "<SysRoot>\System32\netprofmsvc.dll", 0&
            .Add "<SysRoot>\System32\NetSetupSvc.dll", 0&
            .Add "<SysRoot>\System32\NgcCtnrSvc.dll", 0&
            .Add "<SysRoot>\System32\ngcsvc.dll", 0&
            .Add "<SysRoot>\System32\nlasvc.dll", 0&
            .Add "<SysRoot>\System32\nsisvc.dll", 0&
            .Add "<SysRoot>\System32\ntmssvc.dll", 0&
            .Add "<SysRoot>\System32\p2psvc.dll", 0&
            .Add "<SysRoot>\System32\pcasvc.dll", 0&
            .Add "<SysRoot>\System32\peerdistsvc.dll", 0&
            .Add "<SysRoot>\System32\PhoneService.dll", 0&
            .Add "<SysRoot>\System32\PimIndexMaintenance.dll", 0&
            .Add "<SysRoot>\System32\pla.dll", 0&
            .Add "<SysRoot>\System32\pnrpauto.dll", 0&
            .Add "<SysRoot>\System32\pnrpsvc.dll", 0&
            .Add "<SysRoot>\System32\profsvc.dll", 0&
            .Add "<SysRoot>\System32\provsvc.dll", 0&
            .Add "<SysRoot>\System32\qagentRT.dll", 0&
            .Add "<SysRoot>\System32\qmgr.dll", 0&
            .Add "<SysRoot>\System32\qwave.dll", 0&
            .Add "<SysRoot>\System32\rasauto.dll", 0&
            .Add "<SysRoot>\System32\rasmans.dll", 0&
            .Add "<SysRoot>\System32\RDXService.dll", 0&
            .Add "<SysRoot>\System32\regsvc.dll", 0&
            .Add "<SysRoot>\System32\RMapi.dll", 0&
            .Add "<SysRoot>\System32\RpcEpMap.dll", 0&
            .Add "<SysRoot>\System32\rpcss.dll", 0&
            .Add "<SysRoot>\System32\rsvp.exe", 0&
            .Add "<SysRoot>\System32\SCardSvr.dll", 0&
            .Add "<SysRoot>\System32\SCardSvr.exe", 0&
            .Add "<SysRoot>\System32\ScDeviceEnum.dll", 0&
            .Add "<SysRoot>\System32\schedsvc.dll", 0&
            .Add "<SysRoot>\System32\SDRSVC.dll", 0&
            .Add "<SysRoot>\System32\SearchIndexer.exe", 0&
            .Add "<SysRoot>\System32\seclogon.dll", 0&
            .Add "<SysRoot>\System32\sens.dll", 0&
            .Add "<SysRoot>\System32\SensorDataService.exe", 0&
            .Add "<SysRoot>\System32\SensorService.dll", 0&
            .Add "<SysRoot>\System32\sensrsvc.dll", 0&
            .Add "<SysRoot>\System32\services.exe", 0&
            .Add "<SysRoot>\System32\sessenv.dll", 0&
            .Add "<SysRoot>\System32\sessmgr.exe", 0&
            .Add "<SysRoot>\System32\shsvcs.dll", 0&
            .Add "<SysRoot>\System32\SLsvc.exe", 0&
            .Add "<SysRoot>\System32\SLUINotify.dll", 0&
            .Add "<SysRoot>\System32\smlogsvc.exe", 0&
            .Add "<SysRoot>\System32\smphost.dll", 0&
            .Add "<SysRoot>\System32\SmsRouterSvc.dll", 0&
            .Add "<SysRoot>\System32\snmptrap.exe", 0&
            .Add "<SysRoot>\System32\spool\drivers\x64\3\PrintConfig.dll", 0&
            .Add "<SysRoot>\System32\spoolsv.exe", 0&
            .Add "<SysRoot>\System32\sppsvc.exe", 0&
            .Add "<SysRoot>\System32\sppuinotify.dll", 0&
            .Add "<SysRoot>\System32\srsvc.dll", 0&
            .Add "<SysRoot>\System32\srvsvc.dll", 0&
            .Add "<SysRoot>\System32\ssdpsrv.dll", 0&
            .Add "<SysRoot>\System32\sstpsvc.dll", 0&
            .Add "<SysRoot>\System32\storsvc.dll", 0&
            .Add "<SysRoot>\System32\svchost.exe", 0&
            .Add "<SysRoot>\System32\svsvc.dll", 0&
            .Add "<SysRoot>\System32\swprv.dll", 0&
            .Add "<SysRoot>\System32\sysmain.dll", 0&
            .Add "<SysRoot>\System32\SystemEventsBrokerServer.dll", 0&
            .Add "<SysRoot>\System32\TabSvc.dll", 0&
            .Add "<SysRoot>\System32\tapisrv.dll", 0&
            .Add "<SysRoot>\System32\tbssvc.dll", 0&
            .Add "<SysRoot>\System32\termsrv.dll", 0&
            .Add "<SysRoot>\System32\tetheringservice.dll", 0&
            .Add "<SysRoot>\System32\themeservice.dll", 0&
            .Add "<SysRoot>\System32\TieringEngineService.exe", 0&
            .Add "<SysRoot>\System32\tileobjserver.dll", 0&
            .Add "<SysRoot>\System32\TimeBrokerServer.dll", 0&
            .Add "<SysRoot>\System32\trkwks.dll", 0&
            .Add "<SysRoot>\System32\UI0Detect.exe", 0&
            .Add "<SysRoot>\System32\umpnpmgr.dll", 0&
            .Add "<SysRoot>\System32\umpo.dll", 0&
            .Add "<SysRoot>\System32\umrdp.dll", 0&
            .Add "<SysRoot>\System32\unistore.dll", 0&
            .Add "<SysRoot>\System32\upnphost.dll", 0&
            .Add "<SysRoot>\System32\ups.exe", 0&
            .Add "<SysRoot>\System32\userdataservice.dll", 0&
            .Add "<SysRoot>\System32\usermgr.dll", 0&
            .Add "<SysRoot>\System32\usocore.dll", 0&
            .Add "<SysRoot>\System32\uxsms.dll", 0&
            .Add "<SysRoot>\System32\vaultsvc.dll", 0&
            .Add "<SysRoot>\System32\vds.exe", 0&
            .Add "<SysRoot>\System32\vssvc.exe", 0&
            .Add "<SysRoot>\System32\w32time.dll", 0&
            .Add "<SysRoot>\System32\w3ssl.dll", 0&
            .Add "<SysRoot>\System32\WalletService.dll", 0&
            .Add "<SysRoot>\System32\Wat\WatAdminSvc.exe", 0&
            .Add "<SysRoot>\System32\wbem\WmiApSrv.exe", 0&
            .Add "<SysRoot>\System32\wbem\WMIsvc.dll", 0&
            .Add "<SysRoot>\System32\wbengine.exe", 0&
            .Add "<SysRoot>\System32\wbiosrvc.dll", 0&
            .Add "<SysRoot>\System32\wcmsvc.dll", 0&
            .Add "<SysRoot>\System32\wcncsvc.dll", 0&
            .Add "<SysRoot>\System32\WcsPlugInService.dll", 0&
            .Add "<SysRoot>\System32\wdi.dll", 0&
            .Add "<SysRoot>\System32\webclnt.dll", 0&
            .Add "<SysRoot>\System32\wecsvc.dll", 0&
            .Add "<SysRoot>\System32\wephostsvc.dll", 0&
            .Add "<SysRoot>\System32\wercplsupport.dll", 0&
            .Add "<SysRoot>\System32\WerSvc.dll", 0&
            .Add "<SysRoot>\System32\wiarpc.dll", 0&
            .Add "<SysRoot>\System32\wiaservc.dll", 0&
            .Add "<SysRoot>\System32\Windows.Internal.Management.dll", 0&
            .Add "<SysRoot>\System32\windows.staterepository.dll", 0&
            .Add "<SysRoot>\System32\winhttp.dll", 0&
            .Add "<SysRoot>\System32\wkssvc.dll", 0&
            .Add "<SysRoot>\System32\wlansvc.dll", 0&
            .Add "<SysRoot>\System32\wlidsvc.dll", 0&
            .Add "<SysRoot>\System32\workfolderssvc.dll", 0&
            .Add "<SysRoot>\System32\wpcsvc.dll", 0&
            .Add "<SysRoot>\System32\wpdbusenum.dll", 0&
            .Add "<SysRoot>\System32\WpnService.dll", 0&
            .Add "<SysRoot>\System32\WpnUserService.dll", 0&
            .Add "<SysRoot>\System32\wscsvc.dll", 0&
            .Add "<SysRoot>\System32\WsmSvc.dll", 0&
            .Add "<SysRoot>\System32\WSService.dll", 0&
            .Add "<SysRoot>\System32\wuaueng.dll", 0&
            .Add "<SysRoot>\System32\wuauserv.dll", 0&
            .Add "<SysRoot>\System32\WUDFSvc.dll", 0&
            .Add "<SysRoot>\System32\wwansvc.dll", 0&
            .Add "<SysRoot>\System32\wzcsvc.dll", 0&
            .Add "<SysRoot>\System32\XblAuthManager.dll", 0&
            .Add "<SysRoot>\System32\XblGameSave.dll", 0&
            .Add "<SysRoot>\System32\XboxNetApiSvc.dll", 0&
            .Add "<SysRoot>\System32\xmlprov.dll", 0&
            .Add "<SysRoot>\SysWow64\perfhost.exe", 0&
            .Add "<SysRoot>\SysWow64\svchost.exe", 0&
            .Add "<SysRoot>\Microsoft.NET\Framework\v3.0\Windows Communication Foundation\infocard.exe", 0&
            .Add "<SysRoot>\Microsoft.Net\Framework\v3.0\WPF\PresentationFontCache.exe", 0&
            .Add "<SysRoot>\system32\lserver.exe", 0&
            .Add "<SysRoot>\system32\mprdim.dll", 0&
            .Add "<SysRoot>\system32\wdfmgr.exe", 0&
            .Add "<SysRoot>\system32\sacsvr.dll", 0&
            .Add "<SysRoot>\system32\RSoPProv.exe", 0&
            .Add "<SysRoot>\system32\Dfssvc.exe", 0&
            .Add "<SysRoot>\system32\ntfrs.exe", 0&
            
            For Each vKey In .Keys
                prefix = Left$(vKey, InStr(vKey, "\") - 1)
                Select Case prefix
                    Case "<SysRoot>"
                        .Add Replace$(vKey, prefix, sWinDir), 0&
                    Case "<PF64>"
                        .Add Replace$(vKey, prefix, PF_64), 0&
                    Case "<PF32>"
                        .Add Replace$(vKey, prefix, PF_32), 0&
                End Select
            Next
        End With
    End If
    
    IsWinServiceFileName = oDictSRV.Exists(sFilePath)
    
    If Not IsWinServiceFileName Then
    
        If Not (StrComp(sFilePath, PF_64 & "\Windows Defender\mpsvc.dll", 1) = 0) _
          And Not (StrComp(sFilePath, PF_64 & "\Windows Defender\NisSrv.exe", 1) = 0) _
          And Not (StrComp(sFilePath, PF_64 & "\Windows Defender\MsMpEng.exe", 1) = 0) _
          And Not (StrComp(sFilePath, PF_64 & "\Microsoft Security Client\MsMpEng.exe", 1) = 0) _
          And Not (StrComp(sFilePath, PF_64 & "\Microsoft Security Client\NisSrv.exe", 1) = 0) _
          And Not (StrComp(sFilePath, PF_64 & "\Windows Defender Advanced Threat Protection\MsSense.exe", 1) = 0) Then
        
            If Not IsSecurityProductName(sServiceName) Then
        
                sCompany = GetFilePropCompany(sFilePath)
                If InStr(1, sCompany, "Microsoft", 1) > 0 Or InStr(1, sCompany, "Корпорация Майкрософт", 1) > 0 Then
                    IsWinServiceFileName = True
                End If
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IsWinServiceFileName", "File: " & sFilePath
    If inIDE Then Stop: Resume Next
End Function

Function IsSecurityProductName(sProductName As String) As Boolean
    Static IsInit As Boolean
    Static AV() As String
    Dim i&
    
    If Not IsInit Then
        IsInit = True
        AddToArray AV, "security"
        AddToArray AV, "antivirus"
        AddToArray AV, "firewall"
        AddToArray AV, "protect"
        AddToArray AV, "Ad-aware"
        AddToArray AV, "Avast"
        AddToArray AV, "AVG"
        AddToArray AV, "Avira"
        AddToArray AV, "Baidu"
        AddToArray AV, "BitDefender"
        AddToArray AV, "Comodo"
        AddToArray AV, "DrWeb"
        AddToArray AV, "Emsisoft"
        AddToArray AV, "ESET"
        AddToArray AV, "F-Secure"
        AddToArray AV, "GData"
        AddToArray AV, "Hitman"
        AddToArray AV, "Kaspersky"
        AddToArray AV, "Malwarebytes"
        AddToArray AV, "McAfee"
        AddToArray AV, "Norton"
        AddToArray AV, "Panda"
        AddToArray AV, "Qihoo"
        AddToArray AV, "Symantec"
        AddToArray AV, "TrendMicro"
        AddToArray AV, "Vipre"
        AddToArray AV, "Zillya"
        AddToArray AV, "360"
    End If
    
    For i = 0 To UBound(AV)
        If InStr(1, sProductName, AV(i), 1) <> 0 Then IsSecurityProductName = True: Exit Function
    Next
End Function

Public Sub FixO23Item(sItem$, Result As SCAN_RESULT)
    'stop & disable & delete NT service
    'O23 - Service: <displayname> - <company> - <file>
    ' (file missing) or (filesize .., MD5 ..) can be appended
    If Not bIsWinNT Then Exit Sub
    
    On Error GoTo ErrorHandler:
    FixServiceHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO23Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO24Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO24Item - Begin"
    
    'activex desktop components
    Dim sDCKey$, sComponents$(), i&
    Dim sSource$, sSubscr$, sName$, sHit$, Wow64key As Boolean, Result As SCAN_RESULT
    
    Wow64key = False
    
    sDCKey = "Software\Microsoft\Internet Explorer\Desktop\Components"
    sComponents = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sDCKey, Wow64key), "|")
    
    For i = 0 To UBound(sComponents)
        If Reg.KeyExists(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), Wow64key) Then
            sSource = Reg.GetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "Source", Wow64key)
            sSubscr = Reg.GetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "SubscribedURL", Wow64key)
            sSubscr = UnQuote(EnvironW(sSubscr))
            sSubscr = GetLongPath(sSubscr)  ' 8.3 -> Full
            sName = Reg.GetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "FriendlyName", Wow64key)
            If sName = vbNullString Then sName = "(no name)"
            If Not (LCase$(sSource) = "about:home" And LCase$(sSubscr) = "about:home") And _
               Not (UCase$(sSource) = "131A6951-7F78-11D0-A979-00C04FD705A2" And UCase$(sSubscr) = "131A6951-7F78-11D0-A979-00C04FD705A2") Then
               
                sHit = "O24 - Desktop Component " & sComponents(i) & ": " & sName & " - " & IIf(sSource <> "", sSource, IIf(sSubscr <> "", sSubscr, "(no file)"))
                
                If Not IsOnIgnoreList(sHit) Then
                    With Result
                        .Alias = "O24"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), , , CLng(Wow64key)
                        If sSource <> "" Then AddFileToFix .File, REMOVE_FILE, sSource
                        If sSubscr <> "" Then AddFileToFix .File, REMOVE_FILE, sSubscr
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End With
                    AddToScanResults Result
                End If
            End If
        End If
    Next i
    
    AppendErrorLogCustom "CheckO24Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO24Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO24Item(sItem$, Result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    'delete the entire registry key
    'O24 - Desktop Component 1: Internet Explorer Channel Bar - 131A6951-7F78-11D0-A979-00C04FD705A2
    'O24 - Desktop Component 2: Security - %windir%\index.html
    FixRegistryHandler Result
    FixFileHandler Result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO23Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO24Item_Post()
    Const SPIF_UPDATEINIFILE As Long = 1&
    
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, 0&, SPIF_UPDATEINIFILE 'SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    Sleep 1000
    KillProcessByFile sWinDir & "\" & "explorer.exe", True
    Sleep 1000
    Proc.ProcessRun sWinDir & "\" & "explorer.exe", CloseHandles:=True
End Sub
    
Public Function IsOnIgnoreList(sHit$, Optional UpdateList As Boolean, Optional EraseList As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "IsOnIgnoreList - Begin", "Line: " & sHit
    
    Static IsInit As Boolean
    Static aIgnoreList() As String
    
    If EraseList Then
        ReDim aIgnoreList(0)
        Exit Function
    End If
    
    If IsInit And Not UpdateList Then
        If inArray(sHit, aIgnoreList) Then IsOnIgnoreList = True
    Else
        Dim iIgnoreNum&, i&
        
        IsInit = True
        ReDim aIgnoreList(0)
        
        iIgnoreNum = Val(RegReadHJT("IgnoreNum", "0"))
        If iIgnoreNum > 0 Then ReDim aIgnoreList(iIgnoreNum)
        
        For i = 1 To iIgnoreNum
            aIgnoreList(i) = DeCrypt(RegReadHJT("Ignore" & i, vbNullString))
        Next
    End If
    
    AppendErrorLogCustom "IsOnIgnoreList - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modMain_IsOnIgnoreList", sHit
    If inIDE Then Stop: Resume Next
End Function

Public Sub ErrorMsg(ErrObj As ErrObject, sProcedure$, ParamArray vCodeModule())
    Dim sMsg$, sParameters$, HRESULT$, HRESULT_LastDll$, sErrDesc$, iErrNum&, iErrLastDll&, i&
    Dim hwnd As Long, ptr As Long, hMem As Long
    Dim DateTime As String, curTime As Date, ErrText$
    Dim sErrHeader$
    
    'If iErrNum = 0 Then Exit Sub
    'sMsg = "An unexpected error has occurred at procedure: " & _
           sProcedure & "(" & sParameters & ")" & vbCrLf & _
           "Error #" & CStr(iErrNum) & " - " & sErrDesc & vbCrLf & vbCrLf & _
           "Please email me at www.merijn.org/contact.html, reporting the following:" & vbCrLf & _
           "* What you were trying to fix when the error occurred, if applicable" & vbCrLf & _
           "* How you can reproduce the error" & vbCrLf & _
           "* A complete HiJackThis scan log, if possible" & vbCrLf & vbCrLf & _
           "Windows version: " & sWinVersion & vbCrLf & _
           "MSIE version: " & sMSIEVersion & vbCrLf & _
           "HiJackThis version: " & App.Major & "." & App.Minor & "." & App.Revision & _
           vbCrLf & vbCrLf & "This message has been copied to your clipboard." & _
           vbCrLf & "Click OK to continue the rest of the scan."
    
    sErrDesc = ErrObj.Description
    iErrNum = ErrObj.Number
    iErrLastDll = ErrObj.LastDllError
    
    'If iErrNum = 0 Then Exit Sub
    
    If iErrNum <> 33333 And iErrNum <> 0 Then    'error defined by HJT
        HRESULT = ErrMessageText(CLng(iErrNum))
    End If
    
    If iErrLastDll <> 0 Then
        HRESULT_LastDll = ErrMessageText(iErrLastDll)
    End If
    
    For i = 0 To UBound(vCodeModule)
        sParameters = sParameters & vCodeModule(i) & " "
    Next
    
    If IsArrDimmed(TranslateNative) Then
        sErrHeader = TranslateNative(590)
    End If
    If 0 = Len(sErrHeader) Then
        If IsArrDimmed(Translate) Then
            sErrHeader = Translate(590)
        End If
    End If
    If 0 = Len(sErrHeader) Then
        ' Emergency mode (if translation module is not initialized yet)
        sErrHeader = "Please help us improve HiJackThis by reporting this error." & _
            vbCrLf & vbCrLf & "Error message has been copied to clipboard." & _
            vbCrLf & "Click 'Yes' to submit." & _
            vbCrLf & vbCrLf & "Error Details: " & _
            vbCrLf & vbCrLf & "An unexpected error has occurred at function: "
    End If
    
    Dim OSData As String
    
    If ObjPtr(OSver) <> 0 Then
        OSData = OSver.Bitness & " " & OSver.OSName & " (" & OSver.Edition & "), " & _
            OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & ", " & _
            "Service Pack: " & OSver.SPVer & "" & IIf(OSver.IsSafeBoot, " (Safe Boot)", "")
    End If
    
    sMsg = sErrHeader & " " & _
        sProcedure & vbCrLf & _
        "Error # " & iErrNum & IIf(iErrNum <> 0, " - " & sErrDesc, "") & _
        vbCrLf & "HRESULT: " & HRESULT & _
        vbCrLf & "LastDllError # " & iErrLastDll & IIf(iErrLastDll <> 0, " (" & HRESULT_LastDll & ")", "") & _
        vbCrLf & "Trace info: " & sParameters & _
        vbCrLf & vbCrLf & "Windows version: " & OSData & _
        vbCrLf & AppVer
    
    '"Windows version: " & sWinVersion & vbCrLf & vbCrLf & AppVer
    
    If Not bAutoLogSilent Then
    
      Clipboard.Clear
      Clipboard.SetText sMsg
      
      If OpenClipboard(hwnd) Then
        hMem = GlobalAlloc(GMEM_MOVEABLE, 4)
        If hMem <> 0 Then
            ptr = GlobalLock(hMem)
            If ptr <> 0 Then
                GetMem4 &H419, ByVal ptr
                GlobalUnlock hMem
                SetClipboardData CF_LOCALE, hMem
            End If
        End If
        hMem = GlobalAlloc(GMEM_MOVEABLE, LenB(sMsg))
        If hMem <> 0 Then
            ptr = GlobalLock(hMem)
            If ptr <> 0 Then
                lstrcpyn ByVal ptr, ByVal StrPtr(sMsg), LenB(sMsg)
                'CopyMemory ByVal ptr, ByVal StrPtr(sMsg), LenB(sMsg)
                GlobalUnlock hMem
                SetClipboardData CF_UNICODETEXT, hMem
            End If
        End If
        CloseClipboard
      End If
    End If
    
    ' Append error log
    
    curTime = Now()
    
    DateTime = Right$("0" & Day(curTime), 2) & _
        "." & Right$("0" & Month(curTime), 2) & _
        "." & Year(curTime) & _
        " " & Right$("0" & Hour(curTime), 2) & _
        ":" & Right$("0" & Minute(curTime), 2) & _
        ":" & Right$("0" & Second(curTime), 2)
    
    ErrText = " - " & sProcedure & " - #" & iErrNum
    If iErrNum <> 0 Then ErrText = ErrText & " (" & sErrDesc & ")" & IIf(Len(HRESULT) <> 0, " (" & HRESULT & ")", "")
    ErrText = ErrText & " LastDllError = " & iErrLastDll
    If iErrLastDll <> 0 Then ErrText = ErrText & " (" & HRESULT_LastDll & ")"
    If Len(sParameters) <> 0 Then ErrText = ErrText & " " & sParameters
    
    Debug.Print ErrText
    
    ErrReport = ErrReport & vbCrLf & _
        "- " & DateTime & ErrText
    
    AppendErrorLogCustom ">>> ERROR:" & vbCrLf & _
        "- " & DateTime & ErrText
    
    'If Not bAutoLogSilent Then
    
    If Not bAutoLog And Not bSkipErrorMsg Then
        frmError.Show vbModeless
        frmError.Label1.Caption = sMsg
        frmError.Hide
        frmError.Show vbModal
        
'        If vbYes = MsgBoxW(sMsg, vbCritical Or vbYesNo, Translate(591)) Then
'            Dim szParams As String
'            Dim szCrashUrl As String
'            szCrashUrl = "http://safezone.cc/threads/25222/" 'https://sourceforge.net/p/hjt/_list/tickets"
'
'            'szParams = "function=" & sProcedure
'            'szParams = szParams & "&params=" & sParameters
'            'szParams = szParams & "&errorno=" & iErrNum
'            'szParams = szParams & "&errorlastdll=" & iErrLastDll
'            'szParams = szParams & "&errortxt" & sErrDesc
'            'szParams = szParams & "&winver=" & sWinVersion
'            'szParams = szParams & "&hjtver=" & App.Major & "." & App.Minor & "." & App.Revision
'            'szCrashUrl = szCrashUrl & URLEncode(szParams)
'            If True = IsOnline Then
'                ShellExecute 0&, "open", szCrashUrl, vbNullString, vbNullString, vbNormalFocus
'            Else
'                'MsgBoxW "No Internet Connection Available"
'                MsgBoxW Translate(560)
'            End If
'        End If
    End If
    
    If inIDE Then Stop
End Sub

Public Sub AppendErrorLogNoErr(ErrObj As ErrObject, sProcedure As String, ParamArray CodeModule())
    'to append error log without displaying error message to user
    
    On Error Resume Next
    
    Dim i           As Long
    Dim DateTime    As String
    Dim ErrText     As String
    Dim sErrDesc    As String
    Dim iErrNum     As Long
    Dim iErrLastDll As Long
    Dim HRESULT     As String
    Dim HRESULT_LastDll As String
    Dim sParameters As String

    DateTime = Right$("0" & Day(Now), 2) & _
        "." & Right$("0" & Month(Now), 2) & _
        "." & Year(Now) & _
        " " & Right$("0" & Hour(Now), 2) & _
        ":" & Right$("0" & Minute(Now), 2) & _
        ":" & Right$("0" & Second(Now), 2)
    
    sErrDesc = ErrObj.Description
    iErrNum = ErrObj.Number
    iErrLastDll = ErrObj.LastDllError
    
    If iErrNum <> 33333 And iErrNum <> 0 Then    'error defined by HJT
        HRESULT = ErrMessageText(CLng(iErrNum))
    End If
    
    If iErrLastDll <> 0 Then
        HRESULT_LastDll = ErrMessageText(iErrLastDll)
    End If
    
    For i = 0 To UBound(CodeModule)
        sParameters = sParameters & CodeModule(i) & " "
    Next

    ErrText = " - " & sProcedure & " - #" & iErrNum
    If iErrNum <> 0 Then ErrText = ErrText & " (" & sErrDesc & ")" & IIf(Len(HRESULT) <> 0, " (" & HRESULT & ")", "")
    ErrText = ErrText & " LastDllError = " & iErrLastDll
    If iErrLastDll <> 0 Then ErrText = ErrText & " (" & HRESULT_LastDll & ")"
    If Len(sParameters) <> 0 Then ErrText = ErrText & " " & sParameters
    
    ErrReport = ErrReport & vbCrLf & _
        "- " & DateTime & ErrText
    
    AppendErrorLogCustom ">>> ERROR:" & vbCrLf & _
        "- " & DateTime & ErrText
End Sub

Public Function ErrMessageText(lCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
    
    Dim sRtrnMsg   As String
    Dim lRet        As Long

    sRtrnMsg = Space$(MAX_PATH)
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, lCode, 0&, StrPtr(sRtrnMsg), MAX_PATH, 0&)
    If lRet > 0 Then
        ErrMessageText = Left$(sRtrnMsg, lRet)
        ErrMessageText = Replace$(ErrMessageText, vbCrLf, vbNullString)
    End If
End Function

Public Sub CheckDateFormat()
    Dim sBuffer$, uST As SYSTEMTIME
    With uST
        .wDay = 10
        .wMonth = 11
        .wYear = 2003
    End With
    sBuffer = String$(255, 0)
    GetDateFormat 0&, 0&, uST, 0&, StrPtr(sBuffer), 255&
    sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    
    'last try with GetLocaleInfo didn't work on Win2k/XP
    If InStr(sBuffer, "10") < InStr(sBuffer, "11") Then
        bIsUSADateFormat = False
        'msgboxW "sBuffer = " & sBuffer & vbCrLf & "10 < 11, so bIsUSADateFormat False"
    Else
        bIsUSADateFormat = True
        'msgboxW sBuffer & vbCrLf & "10 !< 11, so bIsUSADateFormat True"
    End If
    
    'Dim lLndID&, sDateFormat$
    'lLndID = GetSystemDefaultLCID()
    'sDateFormat = String$(255, 0)
    'GetLocaleInfo lLndID, LOCALE_SSHORTDATE, sDateFormat, 255
    'sDateFormat = left$(sDateFormat, InStr(sDateFormat, vbnullchar) - 1)
    'If sDateFormat = vbNullString Then Exit Sub
    ''sDateFormat = "dd-MM-yy" or "M/d/yy"
    ''I hope this works - dunno what happens in
    ''yyyy-mm-dd or yyyy-dd-mm format
    'If InStr(1, sDateFormat, "d", vbTextCompare) < _
    '   InStr(1, sDateFormat, "m", vbTextCompare) Then
    '    bIsUSADateFormat = False
    'Else
    '    bIsUSADateFormat = True
    'End If
End Sub

Public Function UnEscape(ByVal StringToDecode As String) As String
    Dim i As Long
    Dim acode As Integer, lTmp As Long, HexChar As String

    On Error GoTo ErrorHandler

'    Set scr = CreateObject("MSScriptControl.ScriptControl")
'    scr.Language = "VBScript"
'    scr.Reset
'    Escape = scr.Eval("unescape(""" & s & """)")

    UnEscape = StringToDecode

    If InStr(UnEscape, "%") = 0 Then
         Exit Function
    End If
    For i = Len(UnEscape) To 1 Step -1
        acode = Asc(Mid$(UnEscape, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars

            Case 37
                ' Decode % value
                HexChar = UCase$(Mid$(UnEscape, i + 1, 2))
                If HexChar Like "[0123456789ABCDEF][0123456789ABCDEF]" Then
                    lTmp = CLng("&H" & HexChar)
                    UnEscape = Left$(UnEscape, i - 1) & Chr$(lTmp) & Mid$(UnEscape, i + 3)
                End If
        End Select
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "UnEscape", "string:", StringToDecode
End Function

Public Function HasSpecialCharacters(sName$) As Boolean
    'function checks for special characters in string,
    'like Chinese or Japanese.
    'Used in CheckO3Item (IE Toolbar)
    HasSpecialCharacters = False
    
    'function disabled because of proper DBCS support
    Exit Function
    
    If Len(sName) <> lstrlen(StrPtr(sName)) Then
        HasSpecialCharacters = True
        Exit Function
    End If
    
    If Len(sName) <> LenB(StrConv(sName, vbFromUnicode)) Then
        HasSpecialCharacters = True
        Exit Function
    End If
End Function

Public Sub CheckForReadOnlyMedia()
    Dim sMsg$, hFile As Long, sTempFile$
    
    AppendErrorLogCustom "CheckForReadOnlyMedia - Begin"

    '// TODO: replace by token privilages checking
    
    sTempFile = BuildPath(AppPath(), "~dummy.tmp")
    
    hFile = CreateFile(StrPtr(sTempFile), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
    
    If hFile <= 0 Then
    
    'If Err.Number Then     'Some strange error happens here, if we delete .Number property
        'damn, got no write access
        bNoWriteAccess = True
        sMsg = Translate(7)
'        sMsg = "It looks like you're running HiJackThis from " & _
'               "a read-only device like a CD or locked floppy disk." & _
'               "If you want to make backups of items you fix, " & _
'               "you must copy HiJackThis.exe to your hard disk " & _
'               "first, and run it from there." & vbCrLf & vbCrLf & _
'               "If you continue, you might get 'Path/File Access' "
        MsgBoxW sMsg, vbExclamation
    Else
        CloseW hFile
    End If
    DeleteFileWEx (StrPtr(sTempFile))
    
    AppendErrorLogCustom "CheckForReadOnlyMedia - End"
End Sub

Public Sub SetAllFontCharset()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "SetAllFontCharset - Begin"

    Dim ctl         As Control
    Dim ctlBtn      As CommandButton
    Dim ctlCheckBox As CheckBox
    Dim ctlTxtBox   As TextBox
    Dim ctlLstBox   As ListBox
    Dim CtlLbl      As Label

    For Each ctl In frmMain.Controls
        Select Case TypeName(ctl)
            Case "CommandButton"
                Set ctlBtn = ctl
                SetFontCharSet ctlBtn.Font
            Case "TextBox"
                Set ctlTxtBox = ctl
                SetFontCharSet ctlTxtBox.Font
            Case "ListBox"
                Set ctlLstBox = ctl
                SetFontCharSet ctlLstBox.Font
            Case "Label"
                Set CtlLbl = ctl
                SetFontCharSet CtlLbl.Font
            Case "CheckBox"
                Set ctlCheckBox = ctl
                If ctlCheckBox.Name <> "chkConfigTabs" Then
                    SetFontCharSet ctlCheckBox.Font
                End If
        End Select
    Next ctl

'    With frmMain
'        SetFontCharSet .txtCheckUpdateProxy.Font
'        SetFontCharSet .txtDefSearchAss.Font
'        SetFontCharSet .txtDefSearchCust.Font
'        SetFontCharSet .txtDefSearchPage.Font
'        SetFontCharSet .txtDefStartPage.Font
'        SetFontCharSet .txtHelp.Font
'        SetFontCharSet .txtNothing.Font
'
'        SetFontCharSet .lstBackups.Font
'        SetFontCharSet .lstIgnore.Font
'        SetFontCharSet .lstResults.Font
'    End With
    
    AppendErrorLogCustom "SetAllFontCharset - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_SetAllFontCharset"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub SetFontCharSet(objTxtboxFont As Object)
    On Error GoTo ErrorHandler:
    
    'A big thanks to 'Gun' and 'Adult', two Japanese users
    'who helped me greatly with this
    
    'https://msdn.microsoft.com/en-us/library/aa241713(v=vs.60).aspx
    
    Static IsInit As Boolean
    Static lLCID As Long
    Dim bNonUsCharset As Boolean
    
    bNonUsCharset = True
    
    If Not IsInit Then
        lLCID = GetUserDefaultLCID()
        IsInit = True
    End If
    
    Select Case lLCID
         Case &H404
            objTxtboxFont.Charset = CHINESEBIG5_CHARSET
            objTxtboxFont.Name = ChrW$(&H65B0) & ChrW$(&H7D30) & ChrW$(&H660E) & ChrW$(&H9AD4)   'New Ming-Li
         Case &H411
            objTxtboxFont.Charset = SHIFTJIS_CHARSET
            objTxtboxFont.Name = ChrW$(&HFF2D) & ChrW$(&HFF33) & ChrW$(&H20) & ChrW$(&HFF30) & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
         Case &H412
            objTxtboxFont.Charset = HANGEUL_CHARSET
            objTxtboxFont.Name = ChrW$(&HAD74) & ChrW$(&HB9BC)
         Case &H804
            objTxtboxFont.Charset = CHINESESIMPLIFIED_CHARSET
            objTxtboxFont.Name = ChrW$(&H5B8B) & ChrW$(&H4F53)
         Case Else
            objTxtboxFont.Charset = DEFAULT_CHARSET
            'objTxtboxFont.Name = ""
            bNonUsCharset = False
    End Select
    
    If bNonUsCharset Then objTxtboxFont.Size = 9
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_SetFontCharSet"
    If inIDE Then Stop: Resume Next
End Sub

Private Function TrimNull(S$) As String
    TrimNull = Left$(S, lstrlen(StrPtr(S)))
End Function

Public Sub CheckForStartedFromTempDir()
    'if user picks 'run from current location when downloading HiJackThis.exe,
    'or runs file directly from zip file, exe will be ran from temp folder,
    'meaning a reboot or cache clean could delete it, as well any backups
    'made. Also the user won't be able to find the exe anymore :P
    
    'fixed - 2.0.7
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckForStartedFromTempDir - Begin"
    
    Dim v1          As String
    Dim v2          As String
    Dim cnt         As Long
    Dim sBuffer     As String
    Dim RunFromTemp As Boolean
    Dim sMsg        As String
    
'    sMsg = "HiJackThis appears to have been started from a temporary " & _
'               "folder. Since temp folders tend to be be emptied regularly, " & _
'               "it's wise to copy HiJackThis.exe to a folder of its own, " & _
'               "for instance C:\Program Files\HiJackThis." & vbCrLf & _
'               "This way, any backups that will be made of fixed items " & _
'               "won't be lost." & vbCrLf & vbCrLf & _
'               "May I unpack HJT to desktop for you ?"
'               '"Please quit HiJackThis and copy it to a separate folder " & _
'               '"first before fixing any items."

    'Just too many words
    'User can be shocked and he will close this program immediately and forewer :)
    'l'll try this simple (just only this time):
    
    'Launch from the archive is forbidden !" & vbCrLf & vbCrLf & "May I unzip to desktop for you ?"
    sMsg = TranslateNative(8)
    
    ' проверка на запуск из архива
    If Len(TempCU) <> 0& Then
    
        If StrBeginWith(AppPath(), TempCU) Then RunFromTemp = True
        If Not RunFromTemp Then

            'fix, когда app.path раскрывается в стиле 8.3
            sBuffer = String$(MAX_PATH, vbNullChar)
            cnt = GetLongPathName(StrPtr(AppPath()), StrPtr(sBuffer), Len(sBuffer))
            If cnt Then
                v1 = Left$(sBuffer, cnt)
            Else
                v1 = AppPath()
            End If

            sBuffer = String$(MAX_PATH, vbNullChar)
            cnt = GetLongPathName(StrPtr(TempCU), StrPtr(sBuffer), Len(sBuffer))
            If cnt Then
                v2 = Left$(sBuffer, cnt)
            Else
                v2 = TempCU
            End If
            
            If Len(v1) <> 0 And Len(v2) <> 0 And StrBeginWith(v1, v2) Then RunFromTemp = True
        End If
        
        If RunFromTemp And (Command() = "") Then
            'msgboxW "Запуск из архива запрещен !" & vbCrLf & "Распаковать на рабочий стол для Вас ?", vbExclamation, AppName
            If MsgBoxW(sMsg, vbExclamation Or vbYesNo, g_AppName) = vbYes Then
                Dim NewFile As String
                NewFile = Desktop & "\" & AppExeName(True)
                If FileExists(NewFile) Then     ', Cache:=NO_CACHE
                    SetFileAttributes StrPtr(NewFile), GetFileAttributes(StrPtr(NewFile)) And Not FILE_ATTRIBUTE_READONLY
                    DeleteFileWEx StrPtr(NewFile)
                End If
                CopyFile StrPtr(AppPath(True)), StrPtr(NewFile), ByVal 0&
                If FileExists(NewFile) Then     ', Cache:=NO_CACHE
                    frmMain.ReleaseMutex
                    Proc.ProcessRun NewFile     ', "/twice"
                    Unload frmMain
                    End
                Else
                    'Could not unzip file to Desktop! Please, unzip it manually.
                    MsgBoxW Translate(1007), vbCritical
                    Unload frmMain
                    End
                End If
            Else
                Unload frmMain
                End
            End If
        End If
    End If
    
    AppendErrorLogCustom "CheckForStartedFromTempDir - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckForStartedFromTempDir"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RestartSystem(Optional sExtraPrompt$, Optional bSilent As Boolean)
    Dim OpSysSet As Object
    Dim OpSys As Object
    Dim lRet As Long
    
    'HiJackThis needs to restart the system.
    'Please, save your work and press 'OK'.
    If Not bSilent Then
        MsgBoxW IIf(Len(sExtraPrompt) <> 0, sExtraPrompt & vbCrLf & vbCrLf, "") & TranslateNative(350)
    End If
    
    SetCurrentProcessPrivileges "SeRemoteShutdownPrivilege"
    
    If bIsWinNT Then
        'SHRestartSystemMB frmMain.hWnd, StrConv(sExtraPrompt & IIf(sExtraPrompt <> vbNullString, vbCrLf & vbCrLf, vbNullString), vbUnicode), 2
        
        If OSver.IsWindowsVistaOrGreater Then
            lRet = ExitWindowsEx(EWX_REBOOT Or EWX_FORCEIFHUNG, SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION Or SHTDN_REASON_FLAG_PLANNED)
            'lRet = ExitWindowsEx(EWX_RESTARTAPPS Or EWX_FORCEIFHUNG, SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION Or SHTDN_REASON_FLAG_PLANNED)
        Else 'XP/2000
            lRet = ExitWindowsEx(EWX_REBOOT Or EWX_FORCEIFHUNG, SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION Or SHTDN_REASON_FLAG_PLANNED)
        End If
        
        If lRet = 0 Then 'if ExitWindowsEx somehow failed
            Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
            For Each OpSys In OpSysSet
                OpSys.Reboot
            Next
        End If
        
    Else
        SHRestartSystemMB frmMain.hwnd, sExtraPrompt, 0
    End If
End Sub

Public Function IsIPAddress(sIP$) As Boolean
    'IsIPAddress = IIf(inet_addr(sIP) <> -1, True, False)
    'can't really trust this API, sometimes it bails when the fourth
    'octet is >127
    Dim sOctets$()
    If InStr(sIP, ".") = 0 Then Exit Function
    sOctets = Split(sIP, ".")
    If UBound(sOctets) = 3 Then
        If IsNumeric(sOctets(0)) And _
           IsNumeric(sOctets(1)) And _
           IsNumeric(sOctets(2)) And _
           IsNumeric(sOctets(3)) Then
            If (sOctets(0) >= 0 And sOctets(0) <= 255) And _
               (sOctets(1) >= 0 And sOctets(1) <= 255) And _
               (sOctets(2) >= 0 And sOctets(2) <= 255) And _
               (sOctets(3) >= 0 And sOctets(3) <= 255) Then
                IsIPAddress = True
            End If
        End If
    End If
End Function

Public Function DomainHasDoubleTLD(sDomain$) As Boolean
    Dim sDoubleTLDs$(), i&
    sDoubleTLDs = Split(".co.uk|" & _
                        ".da.ru|" & _
                        ".h1.ru|" & _
                        ".me.uk|" & _
                        ".ss.ru|" & _
                        ".xu.pl", "|")
                        '".com.au|" & _
                        ".com.br|" & _
                        ".1gb.ru|" & _
                        ".biz.ua|" & _
                        ".jps.ru|" & _
                        ".psn.cn|" & _
                        ".spb.ru|" & _
                        'above stuff somehow isn't recognized by IE
                        'as a double TLD - it's not a bug, it's a feature!

    For i = 0 To UBound(sDoubleTLDs)
        If InStr(sDomain, sDoubleTLDs(i)) = Len(sDomain) - Len(sDoubleTLDs(i)) + 1 Then
            DomainHasDoubleTLD = True
            Exit Function
        End If
    Next i
End Function

Public Function GetUser() As String
    AppendErrorLogCustom "GetUser - Begin"
    Dim sUsername$
    sUsername = String$(MAX_PATH, vbNullChar)
    If 0 <> GetUserName(StrPtr(sUsername), MAX_PATH) Then
        sUsername = Left$(sUsername, lstrlen(StrPtr(sUsername)))
    End If
    GetUser = sUsername 'UCase$(sUserName)
    AppendErrorLogCustom "GetUser - End"
End Function

Public Function GetComputer() As String
    AppendErrorLogCustom "GetComputer - Begin"
    Dim sComputerName$
    sComputerName = String$(MAX_PATH, vbNullChar)
    If 0 <> GetComputerName(StrPtr(sComputerName), MAX_PATH) Then
        sComputerName = Left$(sComputerName, lstrlen(StrPtr(sComputerName)))
    End If
    GetComputer = sComputerName 'UCase$(sComputerName)
    AppendErrorLogCustom "GetComputer - End"
End Function

Public Function GetUserType$()
    'based on OpenProcessToken API example from API-Guide
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetUserType - Begin"
    
    Dim hProcessToken&
    Dim BufferSize&
    Dim psidAdmin&, psidPower&, psidUser&, psidGuest&
    Dim lResult&
    Dim i&
    Dim tpTokens As TOKEN_GROUPS
    Dim tpSidAuth As SID_IDENTIFIER_AUTHORITY
    
    If Not bIsWinNT Then
        GetUserType = "Administrator"
        Exit Function
    End If
    
    GetUserType = "unknown"
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
    
    ' Obtain current process token
    If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) Then
        Call OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken)
    End If
    If hProcessToken Then

        ' Determine the buffer size required
        Call GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) ' Determine required buffer size
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
            
            ' Retrieve your token information
            If GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize) <> 1 Then
                CloseHandle hProcessToken
                Exit Function
            End If
            
            ' Move it from memory into the token structure
            Call CopyMemory(tpTokens, InfoBuffer(0), Len(tpTokens))
            
            ' Retreive the builtin sid pointers
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_POWER_USERS, 0, 0, 0, 0, 0, 0, psidPower)
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_USERS, 0, 0, 0, 0, 0, 0, psidUser)
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_GUESTS, 0, 0, 0, 0, 0, 0, psidGuest)
            
            If IsValidSid(psidAdmin) And IsValidSid(psidPower) And _
               IsValidSid(psidUser) And IsValidSid(psidGuest) Then
                For i = 0 To tpTokens.GroupCount
                
                    ' Run through your token sid pointers
                    If IsValidSid(tpTokens.Groups(i).SID) Then
                    
                        ' Test for a match between the admin sid equalling your sid's
                        If EqualSid(tpTokens.Groups(i).SID, psidAdmin) Then
                            GetUserType = "Administrator"
                            Exit For
                        End If
                        If EqualSid(tpTokens.Groups(i).SID, psidPower) Then
                            GetUserType = "Power User"
                            Exit For
                        End If
                        If EqualSid(tpTokens.Groups(i).SID, psidUser) Then
                            GetUserType = "Limited User"
                            Exit For
                        End If
                        If EqualSid(tpTokens.Groups(i).SID, psidGuest) Then
                            GetUserType = "Guest"
                            Exit For
                        End If
                    End If
                Next
            End If
            If psidAdmin Then FreeSid psidAdmin
            If psidPower Then FreeSid psidPower
            If psidUser Then FreeSid psidUser
            If psidGuest Then FreeSid psidGuest
        End If
        CloseHandle hProcessToken
    End If
    
    AppendErrorLogCustom "GetUserType - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetUserType"
    If inIDE Then Stop: Resume Next
End Function

Public Function MapSIDToUsername(sSID As String) As String
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "MapSIDToUsername - Begin", "SID: " & sSID
    
'    Dim objWMI As Object, objSID As Object
'    Set objWMI = GetObject("winmgmts:{impersonationLevel=Impersonate}")
'    Set objSID = objWMI.Get("Win32_SID.SID='" & sSID & "'")
'    MapSIDToUsername = objSID.AccountName
'    Set objSID = Nothing
'    Set objWMI = Nothing
    
    '   PURPOSE: there are certain builtin accounts on Windows NT which do not have a mapped
    '   account name. LookupAccountSid will return the error ERROR_NONE_MAPPED.  This function
    '   generates SIDs for the following accounts that are not mapped:
    '    * ACCOUNT OPERATORS
    '    * SYSTEM OPERATORS
    '    * PRINTER OPERATORS
    '    * BACKUP OPERATORS
    '   the other SID it creates is a LOGON SID, it has a prefix of S-1-5-5.  a LOGON SID is a
    '   unique identifier for a user's logon session.
    
    Dim bufSid() As Byte
    Dim AccName As String
    Dim AccDomain As String
    Dim AccType As Long
    Dim ccAccName As Long
    Dim ccAccDomain As Long
    Dim vOtherName()
    Dim tpSidAuth As SID_IDENTIFIER_AUTHORITY
    Dim pSid(3) As Long
    Dim psidLogonSid As Long
    Dim psidCheck As Long
    Dim i As Long
    
    If UCase$(sSID) = ".DEFAULT" Then
        MapSIDToUsername = "Default user"
        Exit Function
    End If
    
    MapSIDToUsername = "unknown"
    
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
    
    vOtherName = Array("Account operators", "Server operators", "Printer operators", "Backup operators")
    
    bufSid = CreateBufferedSID(sSID)
    
    If IsArrDimmed(bufSid) Then
    
        AccName = String$(MAX_NAME, 0)
        AccDomain = String$(MAX_NAME, 0)
        ccAccName = Len(AccName)
        ccAccDomain = Len(AccDomain)
        psidCheck = VarPtr(bufSid(0))
    
        If 0 <> LookupAccountSid(0&, psidCheck, StrPtr(AccName), ccAccName, StrPtr(AccDomain), ccAccDomain, AccType) Then
        
            MapSIDToUsername = Left$(AccName, ccAccName)
            
        Else
        
            If Err.LastDllError = ERROR_NONE_MAPPED Then
            
                ' Create account operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ACCOUNT_OPS, 0, 0, 0, 0, 0, 0, pSid(0))

                ' Create system operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_SYSTEM_OPS, 0, 0, 0, 0, 0, 0, pSid(1))
        
                ' Create printer operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_PRINT_OPS, 0, 0, 0, 0, 0, 0, pSid(2))
        
                ' Create backup operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_BACKUP_OPS, 0, 0, 0, 0, 0, 0, pSid(3))

                ' Create a logon SID.
                Call AllocateAndInitializeSid(tpSidAuth, 2, 5, 0, 0, 0, 0, 0, 0, 0, psidLogonSid)
    
                '*psnu =  SidTypeAlias;

                If EqualPrefixSid(psidCheck, psidLogonSid) Then
                    MapSIDToUsername = "LOGON SID"
                Else
                    For i = 0 To 3
                        If EqualSid(psidCheck, pSid(i)) Then
                            MapSIDToUsername = vOtherName(i)
                            Exit For
                        End If
                    Next
                End If

                For i = 0 To 3
                    FreeSid pSid(i)
                Next
                FreeSid psidLogonSid
            End If
        End If
    End If
    
    AppendErrorLogCustom "MapSIDToUsername - End"
  Exit Function
ErrorHandler:
    ErrorMsg Err, "MapSIDToUsername", "SID: ", sSID
    If inIDE Then Stop: Resume Next
End Function

Public Sub SilentDeleteOnReboot(sCmd$)
    Dim sDummy$, sFileName$
    'sCmd is all command-line parameters, like this
    '/param1 /deleteonreboot c:\progra~1\bla\bla.exe /param3
    '/param1 /deleteonreboot "c:\program files\bla\bla.exe" /param3
    
    sDummy = Mid$(sCmd, InStr(sCmd, "/deleteonreboot") + Len("/deleteonreboot") + 1)
    If InStr(sDummy, """") = 1 Then
        'enclosed in quotes, chop off at next quote
        sFileName = Mid$(sDummy, 2)
        sFileName = Left$(sFileName, InStr(sFileName, """") - 1)
    Else
        'no quotes, chop off at next space if present
        If InStr(sDummy, " ") > 0 Then
            sFileName = Left$(sDummy, InStr(sDummy, " ") - 1)
        Else
            sFileName = sDummy
        End If
    End If
    DeleteFileOnReboot sFileName, True
End Sub

'Public Sub DeleteFileShell(ByVal sFile$)
'    If Not FileExists(sFile) Then Exit Sub
'    Dim uSFO As SHFILEOPSTRUCT
'    sFile = sFile & chr$(0)
'    With uSFO
'        .pFrom = StrPtr(sFile)
'        .wFunc = FO_DELETE
'        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI Or FOF_NOCONFIRMMKDIR
'    End With
'    SHFileOperation uSFO
'End Sub

Public Function IsProcedureAvail(ByVal ProcedureName As String, ByVal DllFilename As String) As Boolean
    AppendErrorLogCustom "IsProcedureAvail - Begin", "Function: " & ProcedureName, "Dll: " & DllFilename
    Dim hModule As Long, procAddr As Long
    hModule = LoadLibrary(StrPtr(DllFilename))
    If hModule Then
        procAddr = GetProcAddress(hModule, StrPtr(StrConv(ProcedureName, vbFromUnicode)))
        FreeLibrary hModule
    End If
    IsProcedureAvail = (procAddr <> 0)
    AppendErrorLogCustom "IsProcedureAvail - End"
End Function


Public Function CmnDlgSaveFile(sTitle$, sFilter$, Optional sDefFile$)
    Dim uOFN As OPENFILENAME, sFile$
    On Error GoTo ErrorHandler:
    
    Const OFN_ENABLESIZING As Long = &H800000
    
    sFile = String$(MAX_PATH, 0)
    LSet sFile = sDefFile
    With uOFN
        .lStructSize = Len(uOFN)
        If InStr(sFilter, "|") > 0 Then sFilter = Replace$(sFilter, "|", vbNullChar)
        If Right$(sFilter, 2) <> vbNullChar & vbNullChar Then sFilter = sFilter & vbNullChar & vbNullChar
        .lpstrFilter = StrPtr(sFilter)
        .lpstrFile = StrPtr(sFile)
        .lpstrTitle = StrPtr(sTitle)
        .nMaxFile = Len(sFile)
        .Flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_OVERWRITEPROMPT Or OFN_ENABLESIZING
    End With
    If GetSaveFileName(uOFN) = 0 Then Exit Function
    sFile = TrimNull(sFile)
    CmnDlgSaveFile = sFile
    Exit Function
    
ErrorHandler:
    ErrorMsg Err, "modMain_CmnDlgSaveFile", "sTitle=", sTitle, "sFilter=", sFilter, "sDefFile=", sDefFile
    If inIDE Then Stop: Resume Next
End Function

'Public Function CmnDlgOpenFile(sTitle$, sFilter$, Optional sDefFile$)
'    Dim uOFN As OPENFILENAME, sFile$
'    On Error GoTo ErrorHandler:
'
'    sFile = sDefFile & string(256 - Len(sDefFile), 0)
'    With uOFN
'        .lStructSize = Len(uOFN)
'        If InStr(sFilter, "|") > 0 Then sFilter = replace$(sFilter, "|", vbNullChar)
'        If Right$(sFilter, 2) <> vbNullChar & vbNullChar Then sFilter = sFilter & vbNullChar & vbNullChar
'        .lpstrFilter = sFilter
'        .lpstrFile = sFile
'        .lpstrTitle = sTitle
'        .nMaxFile = 256
'        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_PATHMUSTEXIST
'    End With
'    If GetOpenFileName(uOFN) = 0 Then Exit Function
'    sFile = TrimNull(uOFN.lpstrFile)
'    CmnDlgOpenFile = sFile
'    Exit Function
'
'ErrorHandler:
'    ErrorMsg err, "modMain_CmnDlgOpenFile", "sTitle=", sTitle, "sFilter=", sFilter, "sDefFile=", sDefFile
'    If inIDE Then Stop: Resume Next
'End Function

Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String = " ") As VbMsgBoxResult
    Dim hActiveWnd As Long, hMyWnd As Long, frm As Form
    If inIDE Then
        MsgBoxW = MsgBox(Prompt, Buttons, Title)
    Else
        hActiveWnd = GetForegroundWindow()
        For Each frm In Forms
            If frm.hwnd = hActiveWnd Then hMyWnd = hActiveWnd: Exit For
        Next
        MsgBoxW = MessageBox(IIf(hMyWnd <> 0, hMyWnd, g_HwndMain), StrPtr(Prompt), StrPtr(Title), ByVal Buttons)
    End If
End Function

Public Function UnQuote(str As String) As String   ' Trim quotes
    If Len(str) = 0 Then Exit Function
    If Left$(str, 1) = """" And Right$(str, 1) = """" Then
        UnQuote = Mid$(str, 2, Len(str) - 2)
    Else
        UnQuote = str
    End If
End Function

Public Sub ReInitScanResults()  'Global results structure will be cleaned

    'ReDim Scan.Globals(0)
    ReDim Scan(0)

End Sub

Public Sub InitVariables()

    On Error GoTo ErrorHandler:

    Const CSIDL_LOCAL_APPDATA       As Long = &H1C&
    Const CSIDL_COMMON_PROGRAMS     As Long = &H17&

    'SysDisk
    'sWinDir
    'sWinSysDir
    'sSysNativeDir
    'sSysDir (the same as sWinSysDir)
    'sWinSysDirWow64
    'PF_32
    'PF_64
    'AppData
    'LocalAppData
    'Desktop
    'UserProfile
    'AllUsersProfile
    'TempCU
    'envCurUser
    'ProgramData
    'StartMenuPrograms

    AppendErrorLogCustom "InitVariables - Begin"

    Const CSIDL_DESKTOP = 0&

    CRCinit

    'Init user type arrays of scan results
    ReInitScanResults
    
    Dim lr As Long, i As Long, nChars As Long
    
    ReDim tim(10)
    For i = 0 To UBound(tim)
        Set tim(i) = New clsTimer
        tim(i).Index = i
    Next
    
    SysDisk = Space$(MAX_PATH)
    lr = GetSystemWindowsDirectory(StrPtr(SysDisk), MAX_PATH)
    If lr Then
        sWinDir = Left$(SysDisk, lr)
        SysDisk = Left$(SysDisk, 2)
    Else
        sWinDir = EnvironW("%SystemRoot%")
        SysDisk = EnvironW("%SystemDrive%")
    End If
    sWinSysDir = sWinDir & "\" & IIf(bIsWinNT, "system32", "system")
    sSysDir = sWinSysDir
    sWinSysDirWow64 = sWinDir & "\SysWow64"
    
    If bIsWin64 And FolderExists(sWinDir & "\sysnative") And OSver.MajorMinor >= 6 Then
        sSysNativeDir = sWinDir & "\sysnative"
    Else
        sSysNativeDir = sWinDir & "\system32"
    End If
    
    If bIsWin64 Then
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
    
    UserProfile = GetSpecialFolderPath(CSIDL_PROFILE)
    If UserProfile = "" Then UserProfile = EnvironW("%UserProfile%")
    
    nChars = MAX_PATH
    AllUsersProfile = String$(nChars, 0)
    If GetAllUsersProfileDirectory(StrPtr(AllUsersProfile), nChars) Then
        AllUsersProfile = Left$(AllUsersProfile, nChars - 1)
    Else
        AllUsersProfile = EnvironW("%ALLUSERSPROFILE%")
    End If
    
    AppData = GetSpecialFolderPath(CSIDL_APPDATA)
    If AppData = "" Then AppData = EnvironW("%AppData%")
    
    LocalAppData = GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)
    If Len(LocalAppData) = 0 Then
        If OSver.IsWindowsVistaOrGreater Then
            LocalAppData = EnvironW("%LocalAppData%")
        Else
            LocalAppData = UserProfile & "\Local Settings\Application Data"
        End If
    End If
    
    If OSver.MajorMinor < 6 Then
        AppDataLocalLow = AppData
    Else
        AppDataLocalLow = GetKnownFolderPath("{A520A1A4-1780-4FF6-BD18-167343C5AF16}")
    End If
    
    StartMenuPrograms = GetSpecialFolderPath(CSIDL_COMMON_PROGRAMS)
    
    Desktop = GetSpecialFolderPath(CSIDL_DESKTOP)
    
    'TempCU = Environ$("temp") ' will return path in format 8.3 on XP
    TempCU = Reg.GetData(HKEY_CURRENT_USER, "Environment", "Temp")
    ' if REG_EXPAND_SZ is missing
    If InStr(TempCU, "%") <> 0 Then
        TempCU = EnvironW(TempCU)
    End If
    If Len(TempCU) = 0 Or InStr(TempCU, "%") <> 0 Then ' if there TEMP is not defined
        If OSver.IsWindowsVistaOrGreater Then
            TempCU = UserProfile & "\Local\Temp"
        Else
            TempCU = UserProfile & "\Local Settings\Temp"
        End If
    End If
    
    envCurUser = GetUser()
    'envCurUser = EnvironW("%UserName%")
    
    ProgramData = EnvironW("%ProgramData%")
    
    ' Shortcut interfaces initialization
    'IURL_Init
    ISL_Init
    
    Set oDict.TaskWL_ID = New clsTrickHashTable
    
    Set colProfiles = New Collection
    GetProfiles
    
    FillUsers
    
    Set cMath = New clsMath
    Set oRegexp = New cRegExp
    
    LIST_BACKUP_FILE = BuildPath(AppPath(), "Backups\List.ini")
    
    InitBackupIni
    
    If OSver.MajorMinor >= 6.1 Then
        Set TaskBar = New TaskbarLib.TaskbarList
    End If
    
    Set HE = New clsHiveEnum
    
    AppendErrorLogCustom "InitVariables - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "InitVariables"
    If inIDE Then Stop: Resume Next
End Sub

Public Function EnvironW(ByVal SrcEnv As String, Optional UseRedir As Boolean) As String
    Dim lr As Long
    Dim buf As String
    Static LastFile As String
    Static LastResult As String
    
    AppendErrorLogCustom "EnvironW - Begin", "SrcEnv: " & SrcEnv
    
    If Len(SrcEnv) = 0 Then Exit Function
    If InStr(SrcEnv, "%") = 0 Then
        EnvironW = SrcEnv
    Else
        If LastFile = SrcEnv Then
            EnvironW = LastResult
            Exit Function
        End If
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
    LastFile = SrcEnv
    LastResult = EnvironW
    
    AppendErrorLogCustom "EnvironW - End"
End Function

Public Function StrInParamArray(Stri As String, ParamArray vEtalon()) As Boolean
    Dim i As Long
    For i = 0 To UBound(vEtalon)
        If StrComp(Stri, vEtalon(i), 1) = 0 Then StrInParamArray = True: Exit For
    Next
End Function

' Возвращает true, если искомое значение найдено в одном из элементов массива (lB, uB ограничивает просматриваемый диапазон индексов)
Public Function inArray( _
    Stri As String, _
    MyArray() As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    If lB = -2147483647 Then lB = LBound(MyArray)   'some trick
    If uB = 2147483647 Then uB = UBound(MyArray)    'Thanks to Казанский :)
    Dim i As Long
    For i = lB To uB
        If StrComp(Stri, MyArray(i), CompareMethod) = 0 Then inArray = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "inArray"
    If inIDE Then Stop: Resume Next
End Function

'Note: Serialized array - it is a string which stores all items of array delimited by some character (default delimiter in HJT is '|' and '*' chars)
'Example 1: "string1*string2*string3"
'Example 2: "string1|string2|string3" and so.

'this function returns true, if any of items in serialized array has exact match with 'Stri' variable
'you can restrict search with LBound and UBound items only.
Public Function inArraySerialized( _
    Stri As String, _
    SerializedArray As String, _
    Delimiter As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    Dim MyArray() As String
    If 0 = Len(SerializedArray) Then
        If 0 = Len(Stri) Then inArraySerialized = True
        Exit Function
    End If
    MyArray = Split(SerializedArray, Delimiter)
    If lB = -2147483647 Or lB < LBound(MyArray) Then lB = LBound(MyArray)  'some trick
    If uB = 2147483647 Or uB > UBound(MyArray) Then uB = UBound(MyArray)  'Thanks to Казанский :)
    
    Dim i As Long
    For i = lB To uB
        If StrComp(Stri, MyArray(i), CompareMethod) = 0 Then inArraySerialized = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "inArraySerialized", "SerializedString: ", SerializedArray, "delim: ", Delimiter
    If inIDE Then Stop: Resume Next
End Function

'The same as Split(), except of proper error handling when source data is empty string and you assign result to variable defined as array.
'So, in case of empty string it return array with 0 items.
'Also: return type is 'string()' instead of 'variant()'
'
'Warning note: Do not use this function in For each statement !!! - use default Split() instead:
'Differences in behavior:
'Split() with empty string cause 'For each' to not execute any its cycles at all.
'Split() cause to execute 'For Each' for a 1 cycle with empty value.
Public Function SplitSafe(sComplexString As String, Optional Delimiter As String = " ") As String()
    If 0 = Len(sComplexString) Then
        ReDim arr(0) As String
        SplitSafe = arr
    Else
        SplitSafe = Split(sComplexString, Delimiter)
    End If
End Function

'get the first item of serilized array
Public Function SplitExGetFirst(sSerializedArray As String, Optional Delimiter As String = " ") As String
    SplitExGetFirst = SplitSafe(sSerializedArray, Delimiter)(0)
End Function

'get the last item of serialized array
Public Function SplitExGetLast(sSerializedArray As String, Optional Delimiter As String = " ") As String
    Dim Ret() As String
    Ret = SplitSafe(sSerializedArray, Delimiter)
    SplitExGetLast = Ret(UBound(Ret))
End Function

Private Sub DeleteDuplicatesInArray(arr() As String, CompareMethod As VbCompareMethod, Optional DontCompress As Boolean)
    On Error GoTo ErrorHandler:
    
    'DontCompress:
    'if true, do not move items:
    'function will return array with empty items in places where duplicate match were found
    'so, its structure will be similar to the source array
    
    'if false, returns new reconstructed array:
    'all subsequent array items are shifted to the item where duplicate was found.
    
    Dim i   As Long
    
    If DontCompress Then
        For i = LBound(arr) To UBound(arr)
            If inArray(arr(i), arr, i + 1, UBound(arr), CompareMethod) Then
                arr(i) = vbNullString
            End If
        Next
    Else
        Dim TmpArr() As String
        ReDim TmpArr(LBound(arr) To UBound(arr))
        Dim cnt As Long
        cnt = LBound(arr)
        For i = LBound(arr) To UBound(arr)
            If Not inArray(arr(i), arr, i + 1, UBound(arr), CompareMethod) Then
                TmpArr(cnt) = arr(i)
                cnt = cnt + 1
            End If
        Next
        ReDim Preserve TmpArr(LBound(TmpArr) To cnt - 1)
        arr = TmpArr
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "DeleteDuplicatesInArray"
    If inIDE Then Stop: Resume Next
End Sub

Public Function StrBeginWith(Text As String, BeginPart As String) As Boolean
    StrBeginWith = (StrComp(Left$(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function

Public Function StrEndWith(Text As String, LastPart As String) As Boolean
    StrEndWith = (StrComp(Right$(Text, Len(LastPart)), LastPart, 1) = 0)
End Function

Public Function StrEndWithParamArray(Text As String, ParamArray vLastPart()) As Boolean
    Dim i As Long
    For i = 0 To UBound(vLastPart)
        If Len(vLastPart(i)) <> 0 Then
            If StrComp(Right$(Text, Len(vLastPart(i))), vLastPart(i), 1) = 0 Then
                StrEndWithParamArray = True
                Exit For
            End If
        End If
    Next
End Function

Public Function StrBeginWithArray(Text As String, BeginPart() As String) As Boolean
    Dim i As Long
    For i = 0 To UBound(BeginPart)
        If Len(BeginPart(i)) <> 0 Then
            If StrComp(Left$(Text, Len(BeginPart(i))), BeginPart(i), 1) = 0 Then
                StrBeginWithArray = True
                Exit For
            End If
        End If
    Next
End Function

Public Sub CenterForm(myForm As Form) ' Центрирование формы на экране с учетом системных панелей
    On Error Resume Next
    Dim Left    As Long
    Dim Top     As Long
    Left = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN) / 2 - myForm.Width / 2
    Top = Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYFULLSCREEN) / 2 - myForm.Height / 2
    myForm.Move Left, Top
End Sub

Public Function ConvertVersionToNumber(sVersion As String) As Long  '"1.1.1.1" -> 1 number
    On Error GoTo ErrorHandler:
    Dim Ver() As String
    
    If 0 = Len(sVersion) Then Exit Function
    
    Ver = SplitSafe(sVersion, ".")
    If UBound(Ver) = 3 Then
        ConvertVersionToNumber = cMath.Shl(Val(Ver(0)), 24) + cMath.Shl(Val(Ver(1)), 16) + cMath.Shl(Val(Ver(2)), 8) + Val(Ver(3))
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ConvertVersionToNumber"
    If inIDE Then Stop: Resume Next
End Function

Public Sub UpdatePolicy(Optional noWait As Boolean)
    Dim GPUpdatePath$
    If bIsWin64 And FolderExists(sWinDir & "\sysnative") And OSver.MajorMinor >= 6 Then
        GPUpdatePath = sWinDir & "\sysnative\gpupdate.exe"
    Else
        GPUpdatePath = sWinDir & "\system32\gpupdate.exe"
    End If
    If Proc.ProcessRun(GPUpdatePath, "/force", , vbHide) Then
        If Not noWait Then
            Proc.WaitForTerminate , , , 15000
        End If
    End If
End Sub

Public Sub ConcatArrays(DestArray() As String, AddArray() As String)
    'Appends AddArray() to the end of DestArray.
    'DestArray() should be declared as dynamic
    
    'UnInitialized arrays are permitted
    'Warning: if both arrays is uninitialized - DestArray() will remain the same (with uninitialized state)
    
    On Error GoTo ErrorHandler
    
    Dim i&, Idx&
    
    If Not CBool(IsArrDimmed(AddArray)) Then Exit Sub
    If Not CBool(IsArrDimmed(DestArray)) Then
        Idx = -1
        ReDim DestArray(UBound(AddArray) - LBound(AddArray))
    Else
        Idx = UBound(DestArray)
        ReDim Preserve DestArray(UBound(DestArray) + (UBound(AddArray) - LBound(AddArray)) + 1)
    End If
    
    For i = LBound(AddArray) To UBound(AddArray)
        Idx = Idx + 1
        DestArray(Idx) = AddArray(i)
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Parser.ConcatArrays"
End Sub

Public Sub QuickSort(j() As String, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler:
    Dim i As Long, l As Long, M As String, wsp As String
    i = low: l = high: M = j((i + l) \ 2)
    Do Until i > l: Do While j(i) < M: i = i + 1: Loop: Do While j(l) > M: l = l - 1: Loop
        If (i <= l) Then wsp = j(i): j(i) = j(l): j(l) = wsp: i = i + 1: l = l - 1
    Loop
    If low < l Then QuickSort j, low, l
    If i < high Then QuickSort j, i, high
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "QuickSort"
    If inIDE Then Stop: Resume Next
End Sub

' exclude items from ArraySrc() that is not match 'Mask' and save to 'ArrayDest()'
' return value is a number of items in 'ArrayDest'
' if number of items is 0, ArrayDest() will have 1 empty item.
Public Function FilterArray(ArraySrc() As String, ArrayDest() As String, Mask As String) As Long
    On Error GoTo ErrorHandler:
    Dim i As Long, j As Long
    ReDim ArrayDest(LBound(ArraySrc) To UBound(ArraySrc))
    For i = LBound(ArraySrc) To UBound(ArraySrc)
        If ArraySrc(i) Like Mask Then
            j = j + 1
            ArrayDest(LBound(ArraySrc) + j - 1) = ArraySrc(i)
        End If
    Next
    If j = 0 Then
        ReDim ArrayDest(LBound(ArraySrc) To LBound(ArraySrc))
    Else
        ReDim Preserve ArrayDest(LBound(ArraySrc) To LBound(ArraySrc) + j - 1)
    End If
    FilterArray = j
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FilterArray"
    If inIDE Then Stop: Resume Next
End Function

'get a substring starting at the specified character (search begins with the end of the line)
Public Function MidFromCharRev(sText As String, Delimiter As String) As String
    On Error GoTo ErrorHandler:
    Dim iPos As Long
    If 0 <> Len(sText) Then
        iPos = InStrRev(sText, Delimiter)
        If iPos <> 0 Then
            MidFromCharRev = Mid$(sText, iPos + 1)
        Else
            MidFromCharRev = ""
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MidFromCharRev"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetCollectionKeyByIndex(ByVal Index As Long, Col As Collection) As String ' Thanks to 'The Trick' (А. Кривоус) for this code
    
    '//TODO: WARNING: this code can cause crash on XP !!!
    'Dragokas: Added IsBadReadPtr for assurance
    
    On Error GoTo ErrorHandler:
    Dim lpSTR As Long, ptr As Long, Key As String
    If Col Is Nothing Then Exit Function
    Select Case Index
    Case Is < 1, Is > Col.Count: Exit Function
    Case Else
        ptr = ObjPtr(Col)
        Do While Index
            If 0 = IsBadReadPtr(ptr + 24, 4) Then
                GetMem4 ByVal ptr + 24, ptr
            End If
            Index = Index - 1
        Loop
    End Select
    lpSTR = StrPtr(Key)
    
    If 0 = IsBadReadPtr(ptr + 16, 4) Then
        GetMem4 ByVal ptr + 16, ByVal VarPtr(Key)
        GetCollectionKeyByIndex = Key
        GetMem4 lpSTR, ByVal VarPtr(Key)
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetCollectionKey"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetCollectionIndexByItem(sItem As String, Col As Collection) As Long
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(Col.Item(i), sItem, 1) = 0 Then
            GetCollectionIndexByItem = i
            Exit For
        End If
    Next
End Function

Public Function GetCollectionKeyByItem(sItem As String, Col As Collection) As String
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(Col.Item(i), sItem, 1) = 0 Then
            GetCollectionKeyByItem = GetCollectionKeyByIndex(i, Col)
            Exit For
        End If
    Next
End Function

Public Function isCollectionKeyExists(Key As String, Col As Collection) As Boolean
    Dim i As Long
    For i = 1 To Col.Count
        If GetCollectionKeyByIndex(i, Col) = Key Then isCollectionKeyExists = True: Exit For
    Next
End Function

Public Function GetCollectionKeyByItemName(Key As String, Col As Collection) As String
    Dim i As Long
    For i = 1 To Col.Count
        If GetCollectionKeyByIndex(i, Col) = Key Then GetCollectionKeyByItemName = Col.Item(i)
    Next
End Function

Public Function GetCollectionIndexByKey(Key As String, Col As Collection) As Long
    Dim i As Long
    For i = 1 To Col.Count
        If GetCollectionKeyByIndex(i, Col) = Key Then GetCollectionIndexByKey = i
    Next
End Function

Public Sub GetProfiles()    'result -> in global variable 'colProfiles' (collection)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetProfiles - Begin"
    
    'include all folders inside <c:\users>
    'without 'Public'
    
    Dim ProfileListKey      As String
    Dim ProfilesDirectory   As String
    Dim ProfileSubKey()     As String
    Dim ProfilePath         As String
    Dim SubFolders()        As String
    'Dim UserProfile         As String
    Dim i                   As Long
    Dim lr                  As Long
    Dim Path                As String
    Dim objFolder           As Variant
    
    ProfileListKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    ProfilesDirectory = Reg.GetData(0&, ProfileListKey, "ProfilesDirectory")

    Erase ProfileSubKey
    If Reg.EnumSubKeysToArray(0&, ProfileListKey, ProfileSubKey()) > 0 Then
        For i = 1 To UBound(ProfileSubKey)
            If Not (ProfileSubKey(i) = "S-1-5-18" Or _
                    ProfileSubKey(i) = "S-1-5-19" Or _
                    ProfileSubKey(i) = "S-1-5-20") Then
                
                ProfilePath = Reg.GetData(0&, ProfileListKey & "\" & ProfileSubKey(i), "ProfileImagePath")
                
                If Len(ProfilePath) <> 0 Then
                    If FolderExists(ProfilePath) Then
                        If Not isCollectionKeyExists(ProfilePath, colProfiles) Then
                            On Error Resume Next
                            colProfiles.Add ProfilePath, ProfilePath
                            On Error GoTo ErrorHandler:
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    'UserProfile = EnvironW("%UserProfile%")
    
    'добавляю папки, которые находятся в подкаталоге (на 1 уровень ниже) профиля текущего пользователя
    
    If Len(UserProfile) <> 0 Then
        If FolderExists(UserProfile) Then
            Path = UserProfile
            lr = PathRemoveFileSpec(StrPtr(Path))   ' get Parent directory
            If lr Then Path = Left$(Path, lstrlen(StrPtr(Path)))

            SubFolders() = ListSubfolders(Path)

            If CBool(IsArrDimmed(SubFolders)) Then
                For Each objFolder In SubFolders()
                    If Len(objFolder) <> 0 And Not (StrEndWith(CStr(objFolder), "\Public") And OSver.MajorMinor >= 6) Then
                        If FolderExists(CStr(objFolder)) Then
                            If Not isCollectionKeyExists(CStr(objFolder), colProfiles) Then
                                On Error Resume Next
                                colProfiles.Add CStr(objFolder), CStr(objFolder)
                                On Error GoTo ErrorHandler:
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    AppendErrorLogCustom "GetProfiles - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetProfiles"
    If inIDE Then Stop: Resume Next
End Sub

Public Function UnpackResource(ResourceID As Long, DestinationPath As String) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "UnpackResource - Begin", "ID: " & ResourceID, "Destination: " & DestinationPath
    Dim ff      As Integer
    Dim b()     As Byte
    UnpackResource = True
    b = LoadResData(ResourceID, "CUSTOM")
    ff = FreeFile
    Open DestinationPath For Binary Access Write As #ff
        Put #ff, , b
    Close #ff
    AppendErrorLogCustom "UnpackResource - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "UnpackResource", "ID: " & ResourceID, "Destination path: " & DestinationPath
    UnpackResource = False
    If inIDE Then Stop: Resume Next
End Function

Public Sub Terminate_HJT()
    Unload frmMain
    End
End Sub

Public Sub AddHorizontalScrollBarToResults(lstControl As ListBox)
    'Adds a horizontal scrollbar to the results display if it is needed (after the scan)
    Dim x As Long, S$
    Dim listLength As Long
    With lstControl
        For listLength = 0 To .ListCount - 1
            S = Replace$(.List(listLength), vbTab, "12345678")
            If .Width < frmMain.TextWidth(S) + 1000 And x < frmMain.TextWidth(S) + 1000 Then
                x = frmMain.TextWidth(.List(listLength)) + 500
            End If
        Next
        If frmMain.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX + 50  ' if twips change to pixels (+50 to account for the width of the vertical scrollbar
        SendMessage .hwnd, LB_SETHORIZONTALEXTENT, x, ByVal 0&
    End With
End Sub

Public Function IsArrDimmed(vArray As Variant) As Boolean
    IsArrDimmed = (GetArrDims(vArray) > 0)
End Function

Public Function GetArrDims(vArray As Variant) As Integer
    Dim ppSA As Long
    Dim pSA As Long
    Dim vt As Long
    Dim sa As SAFEARRAY
    Const vbByRef As Integer = 16384

    If IsArray(vArray) Then
        GetMem4 ByVal VarPtr(vArray) + 8, ppSA      ' pV -> ppSA (pSA)
        If ppSA <> 0 Then
            GetMem2 vArray, vt
            If vt And vbByRef Then
                GetMem4 ByVal ppSA, pSA                 ' ppSA -> pSA
            Else
                pSA = ppSA
            End If
            If pSA <> 0 Then
                memcpy sa, ByVal pSA, LenB(sa)
                If sa.pvData <> 0 Then
                    GetArrDims = sa.cDims
                End If
            End If
        End If
    End If
End Function

Public Function UBoundSafe(vArray As Variant) As Long
    If GetArrDims(vArray) > 0 Then
        UBoundSafe = UBound(vArray)
    Else
        UBoundSafe = -2147483648#
    End If
End Function

' Преобразовать HTTP: -> HXXP:, HTTPS: -> HXXPS:, WWW -> VVV
Public Function doSafeURLPrefix(sURL As String) As String
    doSafeURLPrefix = Replace(Replace(Replace(sURL, "http:", "hxxp:", , , 1&), "www", "vvv", , , 1&), "https:", "hxxps:", , , 1&)
End Function

Public Sub Dbg(sMsg As String)
    If bDebugMode Then
        AppendErrorLogCustom sMsg
        'OutputDebugStringA sMsg ' -> because already is in AppendErrorLogCustom sub()
    End If
End Sub

Public Sub AppendErrorLogCustom(ParamArray CodeModule())    'trace info
    
    If Not (bDebugMode Or bDebugToFile) Then Exit Sub
    Static freq As Currency
    Static IsInit As Boolean
    
    Dim Other       As String
    Dim i           As Long
    For i = 0 To UBound(CodeModule)
        Other = Other & CodeModule(i) & " | "
    Next
    
    Dim tim1 As Currency
    If Not IsInit Then
        IsInit = True
        QueryPerformanceFrequency freq
    End If
    QueryPerformanceCounter tim1
    
    If bDebugToFile Then
        If hDebugLog <> 0 Then
            If InStr(Other, "modFile.PutW") = 0 Then 'prevent infinite loop
                Dim b() As Byte
                b = "- " & time & " - " & Format$(tim1 / freq, "##0.000") & " - " & Other & vbCrLf
                PutW hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
            End If
        End If
    End If
    
    If bDebugMode Then
    
        OutputDebugStringA Other

        ErrLogCustomText.Append (vbCrLf & "- " & time & " - " & Format$(tim1 / freq, "##0.000") & " - " & Other)
    
        'If DebugHeavy Then AddtoLog vbCrLf & "- " & time & " - " & Other
    End If
End Sub

Public Sub OpenDebugLogHandle()
    Dim sDebugLogFile$
    
    If hDebugLog <> 0 Then Exit Sub
    
    sDebugLogFile = BuildPath(AppPath(), "HiJackThis_debug.log")
    
    If FileExists(sDebugLogFile) Then DeleteFileWEx (StrPtr(sDebugLogFile)), , True
    
    On Error Resume Next
    OpenW sDebugLogFile, FOR_OVERWRITE_CREATE, hDebugLog
    
    If hDebugLog = 0 Then
        sDebugLogFile = BuildPath(AppPath(), "HiJackThis_debug_2.log")
                    
        Call OpenW(sDebugLogFile, FOR_OVERWRITE_CREATE, hDebugLog)
        
    End If
    
    Dim sCurTime$
    sCurTime = vbCrLf & vbCrLf & "Logging started at: " & Now() & vbCrLf & vbCrLf
    PutW hDebugLog, 1&, StrPtr(sCurTime), LenB(sCurTime), doAppend:=True
End Sub

Public Sub OpenLogHandle()
    Dim sLogFile$
    
    sLogFile = BuildPath(AppPath(), "HiJackThis.log")
    
    If FileExists(sLogFile) Then DeleteFileWEx (StrPtr(sLogFile)), , True
    
    On Error Resume Next
    OpenW sLogFile, FOR_OVERWRITE_CREATE, hLog
    
    If hLog = 0 Then
        sLogFile = BuildPath(AppPath(), "HiJackThis_2.log")
        
        Call OpenW(sLogFile, FOR_OVERWRITE_CREATE, hLog)
        
    End If
End Sub

Public Function StringFromPtrA(ByVal ptr As Long) As String
    If 0& <> ptr Then
        StringFromPtrA = SysAllocStringByteLen(ptr, lstrlenA(ptr))
    End If
End Function

Public Function StringFromPtrW(ByVal ptr As Long) As String
    Dim strSize As Long
    If 0 <> ptr Then
        strSize = lstrlen(ptr)
        If 0 <> strSize Then
            StringFromPtrW = String$(strSize, 0&)
            lstrcpyn StrPtr(StringFromPtrW), ptr, strSize + 1&
        End If
    End If
End Function

Public Sub AddToArray(ByRef uArray As Variant, sItem$)
    If Not IsArrDimmed(uArray) Then
        ReDim uArray(0)
        uArray(0) = sItem
    Else
        ReDim Preserve uArray(UBound(uArray) + 1)
        uArray(UBound(uArray)) = sItem
    End If
End Sub

Public Sub DoCrash()
    memcpy 0, ByVal 0, 4
End Sub

Public Sub ParseKeysURL(ByVal sURL As String, aKey() As String, aVal() As String)
    On Error GoTo ErrorHandler:
    'Example:
    'http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IE8SRC
    ' =>
    ' 1) Key(0) = q,    Val(0) = {searchTerms}
    ' 2) Key(1) = src,  Val(1) = IE-SearchBox
    ' 3) Key(2) = FORM, Val(2) = IE8SRC
    
    Dim pos As Long, aKeyPara() As String, aTmp() As String, i As Long
    
    Erase aKey
    Erase aVal
    
    pos = InStr(sURL, "?")
    If pos = 0 Or pos = Len(sURL) Then Exit Sub 'no '?' or last '?'
    sURL = Mid$(sURL, pos + 1)
    
    aKeyPara = Split(sURL, "&")
    
    ReDim aKey(UBound(aKeyPara))
    ReDim aVal(UBound(aKeyPara))
    
    For i = 0 To UBound(aKeyPara)
        aTmp = Split(aKeyPara(i), "=", 2) 'split keypara
        If UBound(aTmp) >= 0 Then 'not empty key ?
            aKey(i) = aTmp(0)
            If StrBeginWith(aKey(i), "amp;") Then 'remove some strange amp; in key that happen sometimes
                If aKey(i) <> "amp;q" And aKey(i) <> "amp;query" Then 'restrict this rule for 'q' and 'query' important keys
                    aKey(i) = Mid$(aKey(i), 5)
                End If
            End If
        End If
        If UBound(aTmp) > 0 Then 'not empty value ?
            aVal(i) = aTmp(1)
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ParseKeysURL", "sURL: " & sURL
    If inIDE Then Stop: Resume Next
End Sub

Public Sub LockInterface(bAllowInfoButtons As Boolean, bDoUnlock As Boolean)
    'Lock controls when scanning
    On Error GoTo ErrorHandler:
    
    Dim mnu As Menu
    Dim ctl As Control
    
    For Each ctl In frmMain.Controls
        If TypeName(ctl) = "Menu" Then
            Set mnu = ctl
            If InStr(1, mnu.Name, "delim", 1) = 0 Then
                mnu.Enabled = bDoUnlock
            End If
        End If
    Next
    Set mnu = Nothing
    Set ctl = Nothing
    
    With frmMain
        .cmdScan.Enabled = bDoUnlock
        .cmdFix.Enabled = bDoUnlock
        If Not bAllowInfoButtons Or bDoUnlock Then 'if allow pressing info..., analyze this
            .cmdInfo.Enabled = bDoUnlock
            .cmdAnalyze.Enabled = bDoUnlock
        End If
        .cmdMainMenu.Enabled = bDoUnlock
        .cmdHelp.Enabled = bDoUnlock
        .cmdConfig.Enabled = bDoUnlock
        .cmdSaveDef.Enabled = bDoUnlock
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "LockInterface", "bAllowInfoButtons: " & bAllowInfoButtons, "bDoUnlock: " & bDoUnlock
    If inIDE Then Stop: Resume Next
End Sub

Public Sub LockInterfaceMain(bDoUnlock As Boolean)
    With frmMain
        .cmdN00bLog.Enabled = bDoUnlock
        .cmdN00bScan.Enabled = bDoUnlock
        .cmdN00bBackups.Enabled = bDoUnlock
        .cmdN00bTools.Enabled = bDoUnlock
        .cmdN00bHJTQuickStart.Enabled = bDoUnlock
        .cmdN00bClose.Enabled = bDoUnlock
    End With
End Sub

Public Function TrimEx(ByVal sStr As String, sDelimiter As String) As String
    Do While Left$(sStr, 1) = sDelimiter And Len(sStr) <> 0
        sStr = Mid$(sStr, 2)
    Loop
    Do While Right$(sStr, 1) = sDelimiter And Len(sStr) <> 0
        sStr = Left$(sStr, Len(sStr) - 1)
    Loop
    TrimEx = sStr
End Function


Public Function CreateLogFile() As String
    Dim sLog As clsStringBuilder
    Dim i&, sProcessList$
    Dim lNumProcesses&
    Dim hProc&, sProcessName$
    Dim Col As New Collection, cnt&
    Dim sTmp$
    
    On Error GoTo MakeLog:
    
    AppendErrorLogCustom "frmMain.CreateLogFile - Begin"
    
    Set sLog = New clsStringBuilder 'speed optimization for huge logs
    
    If Not bLogProcesses Then GoTo MakeLog
    
    'UpdateProgressBar "ProcList"
    'DoEvents
            
        Dim Process() As MY_PROC_ENTRY
        
        lNumProcesses = GetProcesses(Process)
        
        On Error Resume Next
        
        If lNumProcesses Then
        
            For i = 0 To UBound(Process)
        
                sProcessName = Process(i).Path
                
                If Len(Process(i).Path) = 0 Then
                    If Not ((StrComp(Process(i).Name, "System Idle Process", 1) = 0 And Process(i).PID = 0) _
                        Or (StrComp(Process(i).Name, "System", 1) = 0 And Process(i).PID = 4) _
                        Or (StrComp(Process(i).Name, "Memory Compression", 1) = 0 And Process(i).PID = 0) _
                        Or (StrComp(Process(i).Name, "Secure System", 1) = 0) And Process(i).PID = 0) Then
                          sProcessName = Process(i).Name '& " (cannot get Process Path)"
                    End If
                End If
                
                If Len(sProcessName) <> 0 Then
                    Col.Add 1&, sProcessName              ' item - count, key - name of process

                    If Err.Number <> 0& Then              ' if the same process
                        cnt = Col.Item(sProcessName)      ' replacing item of collection
                        Col.Remove (sProcessName)
                        Col.Add cnt + 1&, sProcessName    ' increase count of identical processes
                        Err.Clear
                    End If
                End If
            Next
        End If
    
    'sProcessList = "Running processes:" & vbCrLf
    sProcessList = Translate(29) & ":" & vbCrLf
    
    'sProcessList = sProcessList & "Number | Path" & vbCrLf
    sProcessList = sProcessList & Translate(1020) & " | " & Translate(1021) & vbCrLf
    
    ' Sort using positions array method (Key - Process Path).
    ReDim ProcLog(Col.Count) As MY_PROC_LOG
    ReDim aPos(Col.Count) As Long, aNames(Col.Count) As String
    
    Dim SignResult  As SignResult_TYPE
    
    For i = 1& To Col.Count
        With ProcLog(i)
            .ProcName = GetCollectionKeyByIndex(i, Col)
            .Number = Col(i)
            
' I temporarily disable EDS checking
'            SignVerify .ProcName, 0&, SignResult
'            If SignResult.isLegit Then
'                .EDS_issued = SignResult.SubjectName
'            End If
'
'            If Not bIgnoreAllWhitelists Then
'                UpdateProgressBar "ProcList", .ProcName
'                .IsMicrosoft = (IsMicrosoftCertHash(SignResult.HashRootCert) And SignResult.isLegit)  'validate EDS
'            End If
            
            aNames(i) = IIf(.IsMicrosoft, "(Microsoft) ", IIf(.EDS_issued <> "", "(" & .EDS_issued & ") ", " (not signed)")) & .ProcName
            aPos(i) = i
        End With
    Next
    QuickSortSpecial aNames, aPos, 0, Col.Count
    
    For i = 1& To UBound(aPos)
        With ProcLog(aPos(i))
            'sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & IIf(.IsMicrosoft, "(Microsoft) ", "") & .ProcName & vbCrLf
            'sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & anames(i) & vbCrLf
            If .IsMicrosoft Then
                sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & "(Microsoft) " & .ProcName & vbCrLf
            Else
                'sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & .ProcName & IIf(.EDS_issued <> "", " (" & .EDS_issued & ")", " (not signed)") & vbCrLf
                sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & .ProcName & vbCrLf
            End If
        End With
    Next
    
    sProcessList = sProcessList & vbCrLf
    
    'show all PIDs in debug. mode
    If bDebug Or bDebugToFile Then
        If lNumProcesses Then
            sTmp = ""
            For i = 0 To UBound(Process)
                sTmp = sTmp & Process(i).PID & " | " & IIf(Len(Process(i).Path) <> 0, Process(i).Path, Process(i).Name) & vbCrLf
            Next
            AppendErrorLogCustom sTmp
            sTmp = ""
        End If
    End If
    
    '------------------------------
MakeLog:
    
    UpdateProgressBar "Report"
    
    If Err.Number Then
        sProcessList = "(" & Translate(28) & " (error#" & Err.Number & "))" & vbCrLf
        If Not bAutoLogSilent Then MsgBoxW Err.Description
    End If
    
    On Error GoTo ErrorHandler:
    
    'UpdateProgressBar "Finish"
    'DoEvents
    
    sLog.Append ChrW$(-257) & "Logfile of " & AppVer & vbCrLf & vbCrLf ' + BOM UTF-16 LE
    
    Dim TimeCreated As String
    Dim bSPOld As Boolean
    Dim sUTC As String
    
    'Service pack relevance checking
    
    Select Case OSver.MajorMinor
      
        Case 10
        
        Case 6.3
      
        Case 6.4 '10 Technical preview
            bSPOld = True
            
        Case 6.2 '8
            If Not OSver.IsServer Then bSPOld = True
            
        Case 6.1 '7 / Server 2008 R2
            If OSver.SPVer < 1 Then bSPOld = True
            
        Case 6 'Vista / Server 2008
            If OSver.SPVer < 2 Then bSPOld = True
            
        Case 5.2 'XP x64 / Server 2003 / Server 2003 R2
            If OSver.SPVer < 2 Then bSPOld = True
        
        Case 5.1 'XP
            If OSver.SPVer < 3 Then bSPOld = True
        
        Case 5 '2k / 2k Server
            If OSver.SPVer < 4 Then bSPOld = True
        
    End Select
    
    If GetTimeZone(sUTC) Then
        sUTC = "UTC" & sUTC
    Else
        sUTC = "UTC is unknown"
    End If
    
    TimeCreated = Right$("0" & Day(Now), 2) & "." & Right$("0" & Month(Now), 2) & "." & Year(Now) & " - " & _
        Right$("0" & Hour(Now), 2) & ":" & Right$("0" & Minute(Now), 2)
    
    sLog.Append "Platform:  " & OSver.Bitness & " " & OSver.OSName & " (" & OSver.Edition & "), " & _
            OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & _
            IIf(OSver.ReleaseId <> 0, " (ReleaseId: " & OSver.ReleaseId & ")", "") & ", " & _
            "Service Pack: " & OSver.SPVer & IIf(bSPOld, " <=== Attention! (outdated SP)", "") & _
            IIf(OSver.MajorMinor <> OSver.MajorMinorNTDLL And OSver.MajorMinorNTDLL <> 0, " (NTDLL.dll = " & OSver.NtDllVersion & ")", "") & _
            vbCrLf
    sLog.Append "Time:      " & TimeCreated & " (" & sUTC & ")," & vbTab & "Uptime: " & TrimSeconds(GetSystemUpTime()) & " h/m" & vbCrLf
    sLog.Append "Language:  " & "OS: " & OSver.LangSystemNameFull & " (" & "0x" & Hex$(OSver.LangSystemCode) & "). " & _
            "Display: " & OSver.LangDisplayNameFull & " (" & "0x" & Hex$(OSver.LangDisplayCode) & "). " & _
            "Non-Unicode: " & OSver.LangNonUnicodeNameFull & " (" & "0x" & Hex$(OSver.LangNonUnicodeCode) & ")" & vbCrLf
    
    If OSver.MajorMinor >= 6 Then
        sLog.Append "Elevated:  " & IIf(OSver.IsElevated, "Yes", "No") & vbCrLf  '& vbTab & "IL: " & OSver.GetIntegrityLevel & vbCrLf
    End If
    
    sLog.Append "Ran by:    " & GetUser() & vbTab & "(group: " & OSver.UserType & ") on " & GetComputer() & _
        ", FirstRun: " & IIf(bFirstRebootScan, "yes", "no") & vbCrLf & vbCrLf
    
    
    Dim tmp$
    With BROWSERS   'MY_BROWSERS (look at modUtils.GetBrowsersInfo())
        tmp = .Opera.Version
        If Len(tmp) Then sLog.Append "Opera:   " & tmp & vbCrLf
        tmp = .Chrome.Version
        If Len(tmp) Then sLog.Append "Chrome:  " & tmp & vbCrLf
        tmp = .Firefox.Version
        If Len(tmp) Then sLog.Append "Firefox: " & tmp & vbCrLf
        tmp = .Edge.Version
        If Len(tmp) Then sLog.Append "Edge:    " & tmp & vbCrLf
        tmp = .IE.Version
        If Len(tmp) Then sLog.Append "Internet Explorer: " & tmp & vbCrLf
                         sLog.Append "Default: " & .Default & vbCrLf
    End With
   
    sLog.Append vbCrLf & "Boot mode: " & OSver.SafeBootMode & vbCrLf
    
    '// TODO: improve it (Get environment block)
    If bLogEnvVars Then
        sLog.Append "Windows folder: " & sWinDir & vbCrLf & _
                      "System folder: " & sWinSysDir & vbCrLf & _
                      "Hosts file: " & sHostsFile & vbCrLf
    End If
    
    sLog.Append vbCrLf & sProcessList
    
    ' -----> MAIN Sections <------
    
    If IsArrDimmed(HitSorted) Then
      For i = 0 To UBound(HitSorted)
        ' Adding empty lines beetween sections (cancelled)
        'sPrefix = rtrim$(Splitsafe(HitSorted(i), "-")(0))
        'If sPrefixLast <> "" And sPrefixLast <> sPrefix Then sLog = sLog & vbCrLf
        'sPrefixLast = sPrefix
        sLog.Append HitSorted(i) & vbCrLf
      Next
    End If
    
    ' ----------------------------
    
    Dim IgnoreCnt&
    IgnoreCnt = RegReadHJT("IgnoreNum", "0")
    If IgnoreCnt <> 0 Then
        ' "Warning: Ignore list contains " & IgnoreCnt & " items." & vbCrLf
        sLog.Append vbCrLf & vbCrLf & Replace$(Translate(1011), "[]", IgnoreCnt) & vbCrLf
    End If
    If Not bScanExecuted Then
        ' "Warning: General scanning was not performed." & vbCrLf
        sLog.Append vbCrLf & vbCrLf & Translate(1012) & vbCrLf
    End If
    
    'Append by Error Log
    If 0 <> Len(ErrReport) Then
        sLog.Append vbCrLf & vbCrLf & "Debug information:" & vbCrLf & ErrReport & vbCrLf
        '& vbCrLf & "CmdLine: " & AppPath(True) & " " & command$()
    End If
    
    Dim b()     As Byte
    
    If bDebugToFile Then
        If hDebugLog <> 0 Then
            b() = vbCrLf & vbCrLf & "Contents of the main logfile:" & vbCrLf & vbCrLf & sLog.ToString & vbCrLf
            PutW hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
        End If
    End If
    
    If 0 <> ErrLogCustomText.Length Then
        sLog.Append vbCrLf & vbCrLf & "Trace information:" & vbCrLf & ErrLogCustomText.ToString & vbCrLf
    End If
    
    If bAutoLog Then Perf.EndTime = GetTickCount()
    sLog.Append vbCrLf & "--" & vbCrLf & "End of file - " & "Time spent: " & ((Perf.EndTime - Perf.StartTime) \ 1000) & " sec. - "
    
    If bDebugToFile Then
        If hDebugLog <> 0 Then
            b() = vbCrLf & "--" & vbCrLf & "End of file - " & "Time spent: " & ((Perf.EndTime - Perf.StartTime) \ 1000) & " sec."
            PutW hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
            CloseW hDebugLog: hDebugLog = 0
        End If
    End If
    
    Dim Size_1 As Long
    Dim Size_2 As Long
    Dim Size_3 As Long
    
    Size_1 = 2& * (sLog.Length + Len(" bytes, CRC32: FFFFFFFF. Sign:   "))   'Вычисление размера лога (в байтах)
    Size_2 = Size_1 + 2& * Len(CStr(Size_1))                                 'с учетом самого числа "кол-во байт"
    Size_3 = Size_2 - 2& * Len(CStr(Size_1)) + 2& * Len(CStr(Size_2))        'пересчет, если число байт увеличилось на 1 разряд
    
    sLog.Append CStr(Size_3) & " bytes, CRC32: FFFFFFFF. Sign: "
    
    Dim ForwCRC As Long
    
    b() = sLog.ToString                                                 'считаем CRC лога
    ForwCRC = CalcArrayCRCLong(b()) Xor -1
    
    Dim CorrBytes$: CorrBytes = RecoverCRC(ForwCRC, &HFFFFFFFF)         'считаем байты корректировки
    
    ReDim Preserve b(UBound(b) + 4)                                     'добавляем их в конец массива
    b(UBound(b) - 3) = Asc(Mid$(CorrBytes, 1, 1))
    b(UBound(b) - 2) = Asc(Mid$(CorrBytes, 2, 1))
    b(UBound(b) - 1) = Asc(Mid$(CorrBytes, 3, 1))
    b(UBound(b) - 0) = Asc(Mid$(CorrBytes, 4, 1))
    
    CreateLogFile = b()
    
    If hProc Then CloseHandle hProc
    
    Set sLog = Nothing
    
    AppendErrorLogCustom "frmMain.CreateLogFile - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "frmMain_CreateLogFile"
    Set sLog = Nothing
    If inIDE Then Stop: Resume Next
End Function


' Сортировка по Хоару. На вход - массив j(), на выходе массив k() с индексами массива j в отсортированном порядке + отсортированный массив.
' Алгоритм особо отлично подходит для сортировки User type arrays по любому из полей.
Public Sub QuickSortSpecial(j() As String, K() As Long, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler:
    Dim i As Long, l As Long, M As String, wsp As String
    i = low: l = high: M = j((i + l) \ 2)
    Do Until i > l: Do While j(i) < M: i = i + 1: Loop: Do While j(l) > M: l = l - 1: Loop
        If (i <= l) Then wsp = j(i): j(i) = j(l): j(l) = wsp: wsp = K(i): K(i) = K(l): K(l) = wsp: i = i + 1: l = l - 1
    Loop
    If low < l Then QuickSortSpecial j, K, low, l
    If i < high Then QuickSortSpecial j, K, i, high
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "QuickSortSpecial"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub SortSectionsOfResultList()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.SortSectionsOfResultList - Begin"
    
    ' -----> Sorting of items in results window (ANSI)
    
    Dim Hit() As String
    Dim i As Long
    'HitSorted() -> is a global array
    
    If frmMain.lstResults.ListCount <> 0 Then
    
        ReDim Hit(frmMain.lstResults.ListCount - 1)
        
        For i = 0 To frmMain.lstResults.ListCount - 1
            Hit(i) = frmMain.lstResults.List(i)
        Next i
        
        SortSectionsOfResultList_Ex Hit, HitSorted
        
        ' Rearrange listbox data accorting to sorted list of sections
        frmMain.lstResults.Clear
        For i = 0 To UBound(HitSorted)
            frmMain.lstResults.AddItem HitSorted(i)
        Next
        
    End If
    
    ' -----> Sorting of items in global array (Unicode)
    '
    ' Number of items can be different beetween log and results window,
    ' e.g. O1 - Hosts limited to ~ 20 items for results windows, when in the same time all items are included in the logfile.
    
    If AryPtr(Scan) = 0 Then
    
        ReDim HitSorted(0): HitSorted(0) = ""
    Else
        
        ReDim Hit(UBound(Scan))
        
        For i = 0 To UBound(Scan)
            Hit(i) = Scan(i).HitLineW
        Next i
        
        SortSectionsOfResultList_Ex Hit, HitSorted
    End If
    
    Perf.EndTime = GetTickCount()
    
    AppendErrorLogCustom "frmMain.SortSectionsOfResultList - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_SortSectionsOfResultList"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub SortSectionsOfResultList_Ex(aSrcArray() As String, aDstArray() As String)
    On Error GoTo ErrorHandler:

    ' Special sort procedure
    ' ---------------------------------

    Dim SectSorted() As String
    Dim SectNames() As String
    Dim SectName As String
    Dim nHit As Long
    Dim nSect As Long
    Dim nItemsSect As Long
    Dim nItemsHit As Long
    Dim pos As Long
    Dim i As Long

    ReDim SectNames(41) As String
    ReDim aDstArray(UBound(aSrcArray))
    
    SectNames(0) = "R0"
    SectNames(1) = "R1"
    SectNames(2) = "R2"
    SectNames(3) = "R3"
    SectNames(4) = "R4"
    SectNames(5) = "F0"
    SectNames(6) = "F1"
    SectNames(7) = "F2"
    SectNames(8) = "F3"
    For i = 1 To 30     ' <<<<<<<<<<< increase here in case you added new section !!!
        SectNames(8 + i) = "O" & i
    Next
    
    nItemsHit = 0
    
    For nSect = 0 To UBound(SectNames)
        nItemsSect = 0
        For nHit = 0 To UBound(aSrcArray)
            If 0 <> Len(aSrcArray(nHit)) Then
                pos = InStr(aSrcArray(nHit), "-")
                If pos = 0 Then
                    If Not bAutoLog Then
                        MsgBoxW "Warning! Wrong format of hit line. Must include dash after the name of the section!" & vbCrLf & "Line: " & aSrcArray(nHit)
                    End If
                Else
                    SectName = Trim$(Left$(aSrcArray(nHit), pos - 1))
                    ' Сортирую посекционно
                    If SectName = SectNames(nSect) Then
                        ' Создаю временный массив этой секции для сортировки
                        nItemsSect = nItemsSect + 1
                        ReDim Preserve SectSorted(nItemsSect)
                        SectSorted(nItemsSect) = aSrcArray(nHit)
                        aSrcArray(nHit) = vbNullString
                    End If
                End If
            End If
        Next
        ' Сборка секции завершена.
        If 0 <> nItemsSect Then
            ' Начало сортировки секции
            ' O1 не сортируем (hosts)
            If SectNames(nSect) <> "O1" Then
                QuickSort SectSorted, 0, UBound(SectSorted)
            End If
            For i = 0 To UBound(SectSorted)
                If 0 <> Len(SectSorted(i)) Then
                    'Переносим отсортированную секцию в общий массив
                    aDstArray(nItemsHit) = SectSorted(i)
                    nItemsHit = nItemsHit + 1
                End If
            Next
        End If
    Next
    ' Проверяем, не осталось ли неотсортированных элементов
    For nHit = 0 To UBound(aSrcArray)
        If 0 <> Len(aSrcArray(nHit)) Then
            aDstArray(nItemsHit) = aSrcArray(nHit)
            nItemsHit = nItemsHit + 1
        End If
    Next

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_SortSectionsOfResultListEx"
    If inIDE Then Stop: Resume Next
End Sub


' Append results array with new registry key
Public Sub AddRegToFix( _
    KeyArray() As FIX_REG_KEY, _
    ActionType As ENUM_REG_ACTION_BASED, _
    ByVal lHive As ENUM_REG_HIVE, _
    ByVal sKey As String, _
    Optional sParam As String = "", _
    Optional vDefaultData As Variant = "", _
    Optional eRedirected As ENUM_REG_REDIRECTION = REG_NOTREDIRECTED, _
    Optional ParamType As ENUM_REG_VALUE_TYPE_RESTORE = REG_RESTORE_SAME, _
    Optional ReplaceDataWhat As String = "", _
    Optional ReplaceDataInto As String = "", _
    Optional TrimDelimiter As String = "")
    
    On Error GoTo ErrorHandler
    
    Dim vHiveFix As Variant, eHiveFix As ENUM_REG_HIVE_FIX
    Dim vUseWow As Variant, Wow6432Redir As Boolean
    Dim lActualHive As ENUM_REG_HIVE
    Dim bNoItem As Boolean
    
    If Len(sKey) = 0 Then Exit Sub
    
    If lHive = 0 Then 'if hive handle defined by Key prefix -> ltrim prefix of sKey, and assign handle for lHive
        Call Reg.NormalizeKeyNameAndHiveHandle(lHive, sKey)
    End If
    
    If Not CBool(lHive And &H1000&) Then 'if not combined hive
        lHive = CombineHives(lHive)      'convert ENUM_REG_HIVE -> ENUM_REG_HIVE_FIX to be able to iterate
    End If
    
    For Each vHiveFix In Array(HKCR_FIX, HKCU_FIX, HKLM_FIX, HKU_FIX)
        
        eHiveFix = vHiveFix
        
        If lHive And eHiveFix Then
            
            For Each vUseWow In Array(False, True)
                
                Wow6432Redir = vUseWow
                
                If eRedirected = REG_REDIRECTION_BOTH _
                  Or ((eRedirected = REG_REDIRECTED) And (Wow6432Redir = True)) _
                  Or ((eRedirected = REG_NOTREDIRECTED) And (Wow6432Redir = False)) Then
    
                    lActualHive = ConvertHiveFixToHive(eHiveFix)
                    
                    bNoItem = False
                    
                    If (ActionType And BACKUP_KEY) Or (ActionType And REMOVE_KEY) Or (ActionType And REMOVE_KEY_IF_NO_VALUES) Then
                        If Not Reg.KeyExists(lActualHive, sKey, Wow6432Redir) Then bNoItem = True
                        
                    ElseIf (ActionType And BACKUP_VALUE) Or (ActionType And REMOVE_VALUE) _
                      Or (ActionType And REMOVE_VALUE_IF_EMPTY) Then
                        If Not Reg.ValueExists(lActualHive, sKey, sParam, Wow6432Redir) Then bNoItem = True
                    End If
                    
                    If Not bNoItem Then
                    
                        If AryPtr(KeyArray) Then
                            ReDim Preserve KeyArray(UBound(KeyArray) + 1)
                        Else
                            ReDim KeyArray(0)
                        End If
                        
                        With KeyArray(UBound(KeyArray))
                            .ActionType = ActionType
                            .Hive = lActualHive
                            .Key = sKey
                            .Param = sParam
                            .DefaultData = CStr(vDefaultData)
                            .Redirected = Wow6432Redir
                            .ParamType = ParamType
                            .ReplaceDataWhat = ReplaceDataWhat
                            .ReplaceDataInto = ReplaceDataInto
                            .TrimDelimiter = TrimDelimiter
                        End With
                    End If
                End If
            Next
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddRegToFix", ActionType, lHive, sKey, sParam, vDefaultData, eRedirected, ParamType, ReplaceDataWhat, ReplaceDataInto, TrimDelimiter
    If inIDE Then Stop: Resume Next
End Sub

Function CombineHives(ParamArray vHives() As Variant) As ENUM_REG_HIVE_FIX
    Dim vHive As Variant, lHive As ENUM_REG_HIVE_FIX
    
    For Each vHive In vHives
        If vHive = HKEY_CLASSES_ROOT Then lHive = lHive Or HKCR_FIX
        If vHive = HKEY_CURRENT_USER Then lHive = lHive Or HKCU_FIX
        If vHive = HKEY_LOCAL_MACHINE Then lHive = lHive Or HKLM_FIX
        If vHive = HKEY_USERS Then lHive = lHive Or HKU_FIX
    Next
    CombineHives = lHive Or &H1000&
End Function

Function ConvertHiveFixToHive(lHive As ENUM_REG_HIVE_FIX) As ENUM_REG_HIVE
    Select Case lHive
        Case HKCR_FIX: ConvertHiveFixToHive = HKEY_CLASSES_ROOT
        Case HKCU_FIX: ConvertHiveFixToHive = HKEY_CURRENT_USER
        Case HKLM_FIX: ConvertHiveFixToHive = HKEY_LOCAL_MACHINE
        Case HKU_FIX: ConvertHiveFixToHive = HKEY_USERS
    End Select
End Function

' Append results array with new ini record
Public Sub AddIniToFix( _
    KeyArray() As FIX_REG_KEY, _
    ActionType As ENUM_REG_ACTION_BASED, _
    ByVal sInitFile As String, _
    sSection As String, _
    Optional sParam As String = "", _
    Optional sDefaultData As Variant = "")
    
    On Error GoTo ErrorHandler
    
    If Len(sInitFile) = 0 Then Exit Sub
    
    If ActionType And REMOVE_VALUE_INI Then
        If Not FileExists(sInitFile) Then Exit Sub
    End If
    
    If AryPtr(KeyArray) Then
        ReDim Preserve KeyArray(UBound(KeyArray) + 1)
    Else
        ReDim KeyArray(0)
    End If
    
    With KeyArray(UBound(KeyArray))
        .ActionType = ActionType
        .IniFile = sInitFile
        .Key = sSection
        .Param = sParam
        .DefaultData = sDefaultData
    End With

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddIniToFix", ActionType, sInitFile, sSection, sParam, sDefaultData
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new ini record
Public Sub AddFileToFix( _
    FileArray() As FIX_FILE, _
    ActionType As ENUM_FILE_ACTION_BASED, _
    sFilePath As String, _
    Optional sArguments As String = "", _
    Optional sExpanded As String = "", _
    Optional sAutorun As String = "", _
    Optional sGoodFile As String = "")
    
    On Error GoTo ErrorHandler
    
    If Len(sFilePath) = 0 Then Exit Sub
    
    'if restoring is not required
    If Not CBool(ActionType And RESTORE_FILE) And Not CBool(ActionType And RESTORE_FILE_SFC) Then
    
        'if no file to remove
        If ActionType And REMOVE_FILE Then
            If Not FileExists(sFilePath) Then Exit Sub
        End If
    
        'if nothing to backup
        If ActionType And BACKUP_FILE Then
            If Not FileExists(sFilePath) Then Exit Sub
        End If
    
        'if nothing to unreg.
        If ActionType And UNREG_DLL Then
            If Not FileExists(sFilePath) Then Exit Sub
        End If
    
        'if no folder to remove
        If ActionType And REMOVE_FOLDER Then
            If Not FolderExists(sFilePath) Then Exit Sub
        End If
    End If
    
    If AryPtr(FileArray) Then
        ReDim Preserve FileArray(UBound(FileArray) + 1)
    Else
        ReDim FileArray(0)
    End If
    
    With FileArray(UBound(FileArray))
        .ActionType = ActionType
        .Path = sFilePath
        .GoodFile = sGoodFile
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddFileToFix", ActionType, sFilePath, sArguments, sExpanded, sAutorun, sGoodFile
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new process record
Public Sub AddProcessToFix( _
    ProcessArray() As FIX_PROCESS, _
    ActionType As ENUM_PROCESS_ACTION_BASED, _
    sFilePath As String)
    
    On Error GoTo ErrorHandler
    
    If Len(sFilePath) = 0 Then Exit Sub
    
    If AryPtr(ProcessArray) Then
        ReDim Preserve ProcessArray(UBound(ProcessArray) + 1)
    Else
        ReDim ProcessArray(0)
    End If
    
    With ProcessArray(UBound(ProcessArray))
        .ActionType = ActionType
        .Path = sFilePath
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddProcessToFix", ActionType, sFilePath
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new process record
Public Sub AddServiceToFix( _
    ServiceArray() As FIX_SERVICE, _
    ActionType As ENUM_SERVICE_ACTION_BASED, _
    sServiceName As String, _
    Optional sServiceDisplay As String = "", _
    Optional sImagePath As String = "", _
    Optional sDllPath As String = "")
    
    On Error GoTo ErrorHandler
    
    If Len(sServiceName) = 0 Then Exit Sub
    
    If AryPtr(ServiceArray) Then
        ReDim Preserve ServiceArray(UBound(ServiceArray) + 1)
    Else
        ReDim ServiceArray(0)
    End If
    
    With ServiceArray(UBound(ServiceArray))
        .ActionType = ActionType
        .ImagePath = sImagePath
        .DllPath = sDllPath
        .ServiceName = sServiceName
        .ServiceDisplay = sServiceDisplay
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddServiceToFix", ActionType, sServiceName, sServiceDisplay, sImagePath, sDllPath
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixIt(Result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    If Result.CureType And PROCESS_BASED Then FixProcessHandler Result
    If Result.CureType And FILE_BASED Then FixFileHandler Result
    If (Result.CureType And REGISTRY_BASED) Or (Result.CureType And INI_BASED) Then FixRegistryHandler Result
    If Result.CureType And SERVICE_BASED Then FixServiceHandler Result
    'If Result.CureType And CUSTOM_BASED Then
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixIt", Result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixProcessHandler(Result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    '// TODO: Add protection against system critical processes by full path name
    
    If Result.CureType And PROCESS_BASED Then
        If AryPtr(Result.Process) Then
            For i = 0 To UBound(Result.Process)
                With Result.Process(i)
                    Select Case .ActionType
                
                    Case FREEZE_PROCESS
                        PauseProcessByFile .Path
                
                    Case KILL_PROCESS
                        KillProcessByFile .Path
                        
                    End Select
                End With
            Next
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixProcessHandler", Result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixRegistryHandler(Result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim sData As String, i As Long
    Dim lType As REG_VALUE_TYPE
    
    'If Result.CureType And REGISTRY_BASED Then
        If AryPtr(Result.Reg) Then
            For i = 0 To UBound(Result.Reg)
                With Result.Reg(i)
                    Select Case .ActionType
                    
                    Case RESTORE_VALUE
                        
                        'if need to leave the same type of value
                        If .ParamType = REG_RESTORE_SAME Then
                            If Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then
                                .ParamType = MapRegValueTypeToRegRestoreType(Reg.GetValueType(.Hive, .Key, .Param, .Redirected))
                            Else
                                If .DefaultData <> "" Then
                                    If IsNumeric(.DefaultData) Then
                                        .ParamType = REG_RESTORE_DWORD
                                    Else
                                        .ParamType = REG_RESTORE_EXPAND_SZ
                                    End If
                                Else
                                    .ParamType = REG_RESTORE_EXPAND_SZ
                                End If
                            End If
                        End If
                        
                        Reg.DelVal .Hive, .Key, .Param, .Redirected
                        
                        Select Case .ParamType
                        
                        Case REG_RESTORE_SZ
                            Reg.SetStringVal .Hive, .Key, .Param, CStr(.DefaultData), .Redirected
                        
                        Case REG_RESTORE_EXPAND_SZ
                            Reg.SetExpandStringVal .Hive, .Key, .Param, CStr(.DefaultData), .Redirected
                        
                        'Case REG_RESTORE_BINARY
                        
                        Case REG_RESTORE_DWORD
                            Reg.SetDwordVal .Hive, .Key, .Param, CLng(.DefaultData), .Redirected
                        
                        'Case REG_RESTORE_LINK
                        
                        Case REG_RESTORE_MULTI_SZ
                            ReDim aData(0) As String
                            aData(0) = .DefaultData
                            Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                        
                        End Select
                
                    Case REMOVE_VALUE
                        Reg.DelVal .Hive, .Key, .Param, .Redirected
                
                    Case REMOVE_KEY
                        Reg.DelKey .Hive, .Key, .Redirected
                    
                    Case REMOVE_KEY_IF_NO_VALUES
                        If Not Reg.KeyHasValues(.Hive, .Key, .Redirected) Then
                            Reg.DelKey .Hive, .Key, .Redirected
                        End If
                    
                    Case RESTORE_VALUE_INI
                        IniSetString .IniFile, .Key, .Param, CStr(.DefaultData)
                        
                    Case REMOVE_VALUE_INI
                        '//TODO
                    
                    Case Else
                        If .ActionType And REPLACE_VALUE Then
                        
                            If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                        
                            sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                            
                            sData = Replace$(sData, .ReplaceDataWhat, .ReplaceDataInto, , , vbTextCompare)
                            
                            lType = Reg.GetValueType(.Hive, .Key, .Param, .Redirected)
                            
                            Select Case lType
                            
                            Case REG_SZ
                                Reg.SetStringVal .Hive, .Key, .Param, sData, .Redirected
                                    
                            Case REG_EXPAND_SZ
                                Reg.SetExpandStringVal .Hive, .Key, .Param, sData, .Redirected
                            
                            End Select
                        End If
                        
                        If .ActionType And TRIM_VALUE Then
                            
                            If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                            
                            sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                            
                            sData = TrimEx(sData, .TrimDelimiter)
                            
                            lType = Reg.GetValueType(.Hive, .Key, .Param, .Redirected)
                            
                            Select Case lType
                            
                            Case REG_SZ
                                Reg.SetStringVal .Hive, .Key, .Param, sData, .Redirected
                                    
                            Case REG_EXPAND_SZ
                                Reg.SetExpandStringVal .Hive, .Key, .Param, sData, .Redirected
                            
                            End Select
                        End If
                        
                        If .ActionType And REMOVE_VALUE_IF_EMPTY Then
                            
                            If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                            
                            sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                            If Len(sData) = 0 Then
                                Reg.DelVal .Hive, .Key, .Param, .Redirected
                            End If
                        End If
                    End Select
                End With
            Next
        End If
    'End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixRegistryHandler", Result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Function MapRegValueTypeToRegRestoreType(ordType As REG_VALUE_TYPE, Optional vData As Variant) As ENUM_REG_VALUE_TYPE_RESTORE
    On Error GoTo ErrorHandler
    
    Dim bRequiredDefault As Boolean
    Dim rType As ENUM_REG_VALUE_TYPE_RESTORE
    
    Select Case ordType
    
    Case REG_NONE
        bRequiredDefault = True
        
    Case REG_SZ
        rType = REG_RESTORE_SZ
    
    Case REG_EXPAND_SZ
        rType = REG_RESTORE_EXPAND_SZ
    
    Case REG_BINARY
        bRequiredDefault = True
        
    Case REG_DWORD
        rType = REG_RESTORE_DWORD
    
    Case REG_DWORDLittleEndian
        rType = REG_RESTORE_DWORD
    
    Case REG_DWORDBigEndian
        rType = REG_RESTORE_DWORD
    
    Case REG_LINK
        bRequiredDefault = True
        
    Case REG_MULTI_SZ
        bRequiredDefault = True '// TODO
        
    Case REG_ResourceList
        bRequiredDefault = True
        
    Case REG_FullResourceDescriptor
        bRequiredDefault = True
        
    Case REG_ResourceRequirementsList
        bRequiredDefault = True
        
    Case REG_QWORD        '// TODO
        bRequiredDefault = True
        
    Case REG_QWORD_LITTLE_ENDIAN
        bRequiredDefault = True '// TODO
    End Select
    
    If bRequiredDefault Then
        If IsMissing(vData) Then
            rType = REG_RESTORE_EXPAND_SZ 'default restore value type
        Else
            If IsNumeric(vData) Then
                rType = REG_RESTORE_DWORD
            Else
                rType = REG_RESTORE_EXPAND_SZ
            End If
        End If
    End If
    
    MapRegValueTypeToRegRestoreType = rType

    Exit Function
ErrorHandler:
    ErrorMsg Err, "MapRegValueTypeToRegRestoreType", ordType, vData
    If inIDE Then Stop: Resume Next
End Function

Public Sub FixFileHandler(Result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    If Result.CureType And FILE_BASED Then
        If AryPtr(Result.File) Then
            For i = 0 To UBound(Result.File)
                With Result.File(i)
                
                    If .ActionType And UNREG_DLL Then
                        If Not IsMicrosoftFile(.Path) Then
                            Reg.UnRegisterDll .Path
                        End If
                    End If
                    
                    If .ActionType And REMOVE_FILE Then
                        If FileExists(.Path) Then
                            DeleteFileWEx StrPtr(.Path)
                        End If
                    End If
                    
                    If .ActionType And REMOVE_FOLDER Then
                        If FolderExists(.Path) Then
                            DeleteFolderForce .Path
                        End If
                    End If
                    
                    If .ActionType And RESTORE_FILE Then
                        If FileExists(.GoodFile) Then
                            '// TODO: PendingFileOperation with replacing
                            If DeleteFileWEx(StrPtr(.Path), DisallowRemoveOnReboot:=True) Then
                                FileCopyW .GoodFile, .Path, True
                            End If
                        End If
                    End If

                    If .ActionType And RESTORE_FILE_SFC Then
                        SFC_RestoreFile .Path
                    End If

                End With
            Next
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixFileHandler", Result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Function SFC_RestoreFile(sHijacker As String, Optional bAsync As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    Dim SFC As String
    Dim sHashOld As String
    Dim sHashNew As String
    Dim bNoOldFile As Boolean
    If OSver.Bitness = "x64" And FolderExists(sWinDir & "\sysnative") Then 'Vista+
        SFC = EnvironW("%SystemRoot%") & "\sysnative\sfc.exe"
    Else
        SFC = EnvironW("%SystemRoot%") & "\System32\sfc.exe"
    End If
    If FileExists(SFC) Then
        bNoOldFile = Not FileExists(sHijacker)
        If Not bNoOldFile Then
            TryUnlock sHijacker
            sHashOld = GetFileSHA1(sHijacker, , True)
        End If
        If Proc.ProcessRun(SFC, "/SCANFILE=" & """" & sHijacker & """", , 0) Then
            If Not bAsync Then
                If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 15000) Then
                    Proc.ProcessClose , , True
                End If
                If FileExists(sHijacker) Then
                    sHashNew = GetFileSHA1(sHijacker, , True)
                    If (sHashOld <> sHashNew) Or bNoOldFile Then
                        SFC_RestoreFile = True
                    End If
                End If
            Else
                SFC_RestoreFile = True
            End If
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SFC_RestoreFile", sHijacker
    If inIDE Then Stop: Resume Next
End Function

Public Sub FixServiceHandler(Result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    If Result.CureType And SERVICE_BASED Then
        If AryPtr(Result.Service) Then
            For i = 0 To UBound(Result.Service)
                With Result.Service(i)
                    Select Case .ActionType
            
                    Case DELETE_SERVICE
                        SetServiceStartMode .ServiceName, SERVICE_MODE_DISABLED
                        StopService .ServiceName
                        SetServiceStartMode .ServiceName, SERVICE_MODE_DISABLED
                        DeleteNTService .ServiceName
            
                    Case RESTORE_SERVICE
                        '// TODO

                    End Select
                End With
            Next
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixServiceHandler", Result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Function CheckIntegrityHJT() As Boolean
    'Checking consistency of HiJackThis.exe
    Dim SignResult As SignResult_TYPE
    CheckIntegrityHJT = True
    If Not inIDE Then
        If OSver.IsWindows7OrGreater Then
            If Not (OSver.SPVer = 0 And (OSver.MajorMinor = 6.1)) Then
                'ensure EDS subsystem is working correctly
                SignVerify BuildPath(sWinDir, "system32\ntdll.dll"), SV_LightCheck Or SV_SelfTest Or SV_PreferInternalSign, SignResult
                If SignResult.HashRootCert = "CDD4EEAE6000AC7F40C3802C171E30148030C072" Then
                    SignVerify AppPath(True), SV_PreferInternalSign, SignResult
                    If SignResult.HashRootCert <> "05F1F2D5BA84CDD6866B37AB342969515E3D912E" Then
                        'not a developer machine ?
                        If Not (GetUser() = "Alex" And (GetDateAtMidnight(GetFileDate(AppPath(True), DATE_MODIFIED)) = GetDateAtMidnight(Now()))) Then
                            'Warning! Integrity of HiJackThis program is corrupted. Perhaps, file is patched or infected by file virus.
                            ErrReport = ErrReport & vbCrLf & Translate(1023) & vbCrLf
                            CheckIntegrityHJT = False
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function SetTaskBarProgressValue(frm As Form, ByVal Value As Single) As Boolean
    If Value < 0 Or Value > 1 Then Exit Function
    If Not (TaskBar Is Nothing) Then
        If Value = 0 Then
            TaskBar.SetProgressState frmMain.hwnd, TBPF_NOPROGRESS
        Else
            TaskBar.SetProgressValue frm.hwnd, CCur(Value * 10000), CCur(10000)
        End If
    End If
End Function

Public Function InstallHJT(Optional bAskToCreateDesktopShortcut As Boolean, Optional bSilent As Boolean) As Boolean
    Dim HJT_Location As String
    
    InstallHJT = True

    HJT_Location = BuildPath(PF_32, "HiJackThis Fork\HiJackThis.exe")
    
    MkDirW GetParentDir(HJT_Location)
    
    'Copy exe to Program Files dir
    If Not FileCopyW(AppPath(True), HJT_Location, True) Then
        'MsgBoxW "Error while installing HiJackThis to program files folder. Cannot copy. Error = " & Err.LastDllError, vbCritical
        MsgBoxW Translate(593) & " " & Err.LastDllError, vbCritical
        InstallHJT = False
        Exit Function
    End If
    
    'create Control panel -> 'Uninstall programs' entry
    CreateUninstallKey True, HJT_Location
    
    'Shortcuts in Start Menu
    InstallHJT = CreateHJTShortcuts(HJT_Location)
    
    If bAskToCreateDesktopShortcut Then
        If bSilent Then
            CreateHJTShortcutDesktop HJT_Location
        Else
            'Do you want to create shortcut in Desktop?
            If MsgBoxW(Translate(69), vbYesNo) = vbYes Then
                CreateHJTShortcutDesktop HJT_Location
            End If
        End If
    End If
End Function

Public Function InstallAutorunHJT(Optional bSilent As Boolean) As Boolean
    Dim JobCommand As String
    Dim HJT_Location As String
    Dim iExitCode As Long
    
    HJT_Location = BuildPath(PF_32, "HiJackThis Fork\HiJackThis.exe")
    
'    If MsgBox("This will install HJT to 'Program Files' folder and set task scheduler for automatically run HJT scan at system startup." & _
'        vbCrLf & vbCrLf & "Continue?" & vbCrLf & vbCrLf & "Note: it is recommended that you add all safe items to ignore list, so " & _
'        "the results window will appear at system startup if only new item will be found.", vbYesNo Or vbQuestion) = vbNo Then
    If Not bSilent Then
        If MsgBoxW(Translate(66), vbYesNo Or vbQuestion) = vbNo Then
            gNotUserClick = True
            frmMain.chkConfigStartupScan.Value = 0
            gNotUserClick = False
            Exit Function
        End If
    End If
    
    'check if 'Schedule' service is launched
    If Not RunScheduler_Service(True, Not bSilent, bSilent) Then
        Exit Function
    End If
    
    If InstallHJT(, (InStr(1, Command(), "/noGUI", 1) <> 0)) Then
    
        'delay after system startup for 1 min.
        JobCommand = "/create /tn ""HiJackThis Autostart Scan"" /SC ONSTART /DELAY 0001:00 /F /RL HIGHEST " & _
            "/tr ""\""" & HJT_Location & "\"" /startupscan"""
    
        If Proc.ProcessRun("schtasks.exe", JobCommand, , 0) Then
            iExitCode = Proc.WaitForTerminate(, , , 15000)     'if ExitCode = 0, 15 sec for timeout
            If ERROR_SUCCESS <> iExitCode Then
                Proc.ProcessClose , , True
                'MsgBoxW "Error while creating task. Error = " & iExitCode, vbCritical
                MsgBoxW Translate(594) & " " & iExitCode, vbCritical
            Else
                InstallAutorunHJT = True
            End If
        End If
    End If
End Function

Public Function RemoveAutorunHJT() As Boolean
    RemoveAutorunHJT = KillTask2("\HiJackThis Autostart Scan")
End Function

Public Sub OpenAndSelectFile(sFile As String)
    Dim hRet As Long
    Dim pIDL As Long
    
    If OSver.MajorMinor >= 5.1 Then '(XP+)
    
        pIDL = ILCreateFromPath(StrPtr(sFile))

        If pIDL <> 0 Then
            hRet = SHOpenFolderAndSelectItems(pIDL, 0, 0, 0)
    
            ILFree pIDL
        End If
    End If
    
    If pIDL = 0 Or hRet <> S_OK Then
        'alternate
        Shell "explorer.exe /select," & """" & sFile & """", vbNormalFocus   ' open folder with a log
    End If
End Sub

Public Function GetDateAtMidnight(dDate As Date) As Date
    GetDateAtMidnight = DateAdd("s", -Second(dDate), DateAdd("n", -Minute(dDate), DateAdd("h", -Hour(dDate), dDate)))
End Function

Public Sub HJT_SaveReport()
    On Error GoTo ErrorHandler:
    Dim Idx&
    
    AppendErrorLogCustom "HJT_SaveReport - Begin"

    Dim sLogFile$
        
        Idx = 7
        
        If bAutoLog Then
            sLogFile = BuildPath(AppPath(), "HiJackThis.log")
        Else
            bGlobalDontFocusListBox = True
            'sLogFile = CmnDlgSaveFile("Save logfile...", "Log files (*.log)|*.log|All files (*.*)|*.*", "HiJackThis.log")
            sLogFile = CmnDlgSaveFile(Translate(1001), Translate(1002) & " (*.log)|*.log|" & Translate(1003) & " (*.*)|*.*", "HiJackThis.log")
            bGlobalDontFocusListBox = False
        End If
        
        Idx = 8
        
        If 0 <> Len(sLogFile) Then
            
            Idx = 11
            
            Dim b() As Byte
            
            b = CreateLogFile() '<<<<<< ------- preparing all text for log file
            
            Idx = 12
            
            'If FileExists(sLogFile) Then DeleteFileWEx (StrPtr(sLogFile))
            
            Idx = 13
            
            If hLog <= 0 Then
                If Not OpenW(sLogFile, FOR_OVERWRITE_CREATE, hLog) Then
    
                    If Not bAutoLogSilent Then 'not via AutoLogger
                        'try another name
    
                        sLogFile = BuildPath(AppPath(), "HiJackThis_2.log")
    
                        Call OpenW(sLogFile, FOR_OVERWRITE_CREATE, hLog)
                    End If
                End If
            End If
            
            If hLog <= 0 Then
                If bAutoLogSilent Then 'via AutoLogger
                    Exit Sub
                Else
                
                    If bAutoLog Then ' if user clicked 1-st button (and HJT on ReadOnly media) => try another folder
                    
                        bGlobalDontFocusListBox = True
                        'sLogFile = CmnDlgSaveFile("Save logfile...", "Log files (*.log)|*.log|All files (*.*)|*.*", "HiJackThis.log")
                        sLogFile = CmnDlgSaveFile(Translate(1001), Translate(1002) & " (*.log)|*.log|" & Translate(1003) & " (*.*)|*.*", "HiJackThis.log")
                        bGlobalDontFocusListBox = False
                        
                        If 0 <> Len(sLogFile) Then
                            If Not OpenW(sLogFile, FOR_OVERWRITE_CREATE, hLog) Then    '2-nd try
                                MsgBoxW Translate(26), vbExclamation
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If
                        
                    Else 'if user already clicked button "Save report"
                    
'                   msgboxW "Write access was denied to the " & _
'                       "location you specified. Try a " & _
'                       "different location please.", vbExclamation
                        MsgBoxW Translate(26), vbExclamation
                        Exit Sub
                    End If
                End If
            End If

            PutW hLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
            
            CloseW hLog: hLog = 0
            
            Idx = 14
            
            If (Not bAutoLogSilent) Or inIDE Then
                If ShellExecute(g_HwndMain, StrPtr("open"), StrPtr(sLogFile), 0&, 0&, 1) <= 32 Then
                    'system doesn't know what .log is
                    If FileExists(sWinDir & "\notepad.exe") Then
                        ShellExecute g_HwndMain, StrPtr("open"), StrPtr(sWinDir & "\notepad.exe"), StrPtr(sLogFile), 0&, 1
                    Else
                        If FileExists(sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\notepad.exe") Then
                            ShellExecute g_HwndMain, StrPtr("open"), StrPtr(sWinDir & IIf(bIsWinNT, "\sytem32", "\system") & "\notepad.exe"), StrPtr(sLogFile), 0&, 1
                        Else
                            'MsgBoxW Replace$(Translate(27), "[]", sLogFile), vbInformation
    '                        msgboxW "The logfile has been saved to " & sLogFile & "." & vbCrLf & _
    '                               "You can open it in a text editor like Notepad.", vbInformation
                        
                            OpenAndSelectFile sLogFile
                        End If
                    End If
                End If
            End If
        End If
    
    AppendErrorLogCustom "HJT_SaveReport - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "HJT_SaveReport", "index = ", Idx
    If inIDE Then Stop: Resume Next
End Sub

Public Sub HJT_Shutdown()   ' emergency exits the program due to exceeding the timeout limit
    
    '!!! HiJackThis was shut down due to exceeding the maximum allowed timeout: [] sec. !!! Report file will be incomplete!
    'Please, restart the program manually (not via Autologger).
    ErrReport = ErrReport & vbCrLf & Replace$(Translate(1027), "[]", Perf.MAX_TimeOut)

    SortSectionsOfResultList
    HJT_SaveReport
    
    Unload frmMain
    If Not inIDE Then ExitProcess 1001&
    If inIDE Then Stop: Resume Next
End Sub

Public Function WhiteListed(sFile As String, sWhiteListedPath As String, Optional bCheckFileNameOnly As Boolean) As Boolean
    'to check matching the file with the specified name and verify it by EDS

    If bHideMicrosoft And Not bIgnoreAllWhitelists Then
        If bCheckFileNameOnly Then
            If StrComp(GetFileName(sFile, True), sWhiteListedPath, 1) = 0 Then
                If IsMicrosoftFile(sFile) Then WhiteListed = True
            End If
        Else
            If StrComp(sFile, sWhiteListedPath, 1) = 0 Then
                If IsMicrosoftFile(sFile) Then WhiteListed = True
            End If
        End If
    End If
End Function
